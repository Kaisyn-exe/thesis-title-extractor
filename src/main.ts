import "./styles.css";
import { getVersion } from "@tauri-apps/api/app";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";
import { openUrl } from "@tauri-apps/plugin-opener";
import {
  buildExcel, DEFAULT_EXCEL_OPTIONS, EXCEL_COLUMNS, exportCoverImages, exportCoverPdf,
  type ExcelOptions, type PaperRecord, type SortKey,
} from "./exporters";
import { extract, Method, needsCheck } from "./extractor";
import { basename, collectPdfs, dirname, isDir, joinPath, openPath, readFile, renameFile, revealPath, writeFile } from "./native";
import { ocrPage } from "./ocr";
import { loadPdf, renderCover } from "./pdf";
import { loadProgress, pickFields, sameFields, saveProgress, type EditField, type Fields } from "./progress";
import { DEFAULT_PATTERN, planRenames, TOKENS } from "./rename";
import { matchRoster, parseRoster, type Student } from "./roster";
import { checkForUpdate, RELEASES_URL } from "./update";

type Paper = PaperRecord & {
  id: number;
  state: "pending" | "done" | "error";
  size: number;
  mtime: number;
  /** 自动识别出的原始结果，用来判断老师改过什么、以及「放弃修改」 */
  auto?: Fields & { method: string };
  restored?: boolean;
};
type Filter = "all" | "check" | "done";
type View = "table" | "gallery";
type Opt = "recursive" | "ocr" | "autoUpdate";
type EditKey = EditField;
type Status = "pending" | "err" | "manual" | "warn" | "ok";

const $ = <T extends HTMLElement = HTMLElement>(sel: string) => document.querySelector(sel) as T;
const $$ = <T extends HTMLElement = HTMLElement>(sel: string) => [...document.querySelectorAll<T>(sel)];

const ui = {
  app: $("#app"),
  folder: $("#folder"),
  rows: $("#rows"),
  sheet: $("#sheet"),
  noMatch: $("#no-match"),
  cellName: $("#cell-name"),
  fx: $<HTMLInputElement>("#fx-input"),
  paneCover: $<HTMLButtonElement>("#pane-cover"),
  paneTitle: $("#pane-title"),
  paneNote: $("#pane-note"),
  paneKv: $("#pane-kv"),
  grid: $("#grid"),
  drawer: $("#drawer"),
  drawerPos: $("#drawer-pos"),
  drawerCover: $<HTMLButtonElement>("#drawer-cover"),
  drawerFile: $("#drawer-file"),
  drawerMeta: $("#drawer-meta"),
  drawerNote: $("#drawer-note"),
  search: $<HTMLInputElement>("#search"),
  cnt: { all: $("#cnt-all"), check: $("#cnt-check"), done: $("#cnt-done") },
  status: $("#status"),
  progress: $("#progress"),
  bar: $("#bar"),
  summary: $("#summary"),
  hint: $("#hint"),
  optionsPop: $("#options-pop"),
  themeLabel: $("#theme-label"),
  saved: $("#saved"),
  versionLabel: $("#version-label"),
  modal: $("#modal"),
  overlay: $("#drop-overlay"),
  toasts: $("#toasts"),
};

const state = {
  papers: [] as Paper[],
  selected: null as number | null,
  col: 2, // 表格中当前单元格所在列（见 COLS）
  filter: "all" as Filter,
  query: "",
  view: "table" as View,
  busy: false,
  baseFolder: "",
  sources: [] as string[],
  opts: { recursive: false, ocr: true, autoUpdate: true } as Record<Opt, boolean>,
  roster: null as { file: string; students: Student[] } | null,
};

/** 表格的列：第 0 列是行号 */
const COLS: { key: "no" | "status" | EditKey | "method" | "file"; letter: string; editable?: boolean }[] = [
  { key: "no", letter: "" },
  { key: "status", letter: "A" },
  { key: "title", letter: "B", editable: true },
  { key: "name", letter: "C", editable: true },
  { key: "studentId", letter: "D", editable: true },
  { key: "college", letter: "E", editable: true },
  { key: "major", letter: "F", editable: true },
  { key: "advisor", letter: "G", editable: true },
  { key: "advisorTitle", letter: "H", editable: true },
  { key: "method", letter: "I" },
  { key: "file", letter: "J" },
];

const esc = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));

const isCheck = (p: Paper) => p.state !== "pending" && needsCheck(p);
const isDone = (p: Paper) => p.state === "done" && !needsCheck(p);
const current = () => (state.selected === null ? null : state.papers[state.selected]);

function statusOf(p: Paper): Status {
  if (p.state === "pending") return "pending";
  if (p.state === "error") return "err";
  if (p.method === Method.Manual || p.method === Method.Confirmed) return "manual";
  return needsCheck(p) ? "warn" : "ok";
}

const STATUS_TEXT: Record<Status, string> = { pending: "识别中", err: "× 打不开", manual: "✎ 已确认", warn: "! 需核对", ok: "✓ 已识别" };
const BADGE: Record<Status, string> = {
  pending: `<span class="spin"></span>`,
  ok: `<svg viewBox="0 0 12 12"><path d="m2.5 6.3 2.3 2.2 4.7-5"/></svg>`,
  warn: `<svg viewBox="0 0 12 12"><path d="M6 2.5v4.3M6 9.4v.1"/></svg>`,
  manual: `<svg viewBox="0 0 12 12"><path d="m2.5 9.5.4-2 5-5 1.6 1.6-5 5Z"/></svg>`,
  err: `<svg viewBox="0 0 12 12"><path d="m3.5 3.5 5 5m0-5-5 5"/></svg>`,
};

// ---------------------------------------------------------------- 提示与状态栏

interface ToastAction { label: string; run: () => void; primary?: boolean }

function toast(message: string, kind: "ok" | "warn" | "error" = "ok", actions: ToastAction[] = []) {
  const el = document.createElement("div");
  el.className = `toast ${kind}`;
  el.innerHTML = `<div class="msg">${esc(message)}<div class="acts"></div></div><button class="x" title="关闭">×</button>`;
  const acts = el.querySelector(".acts")!;
  if (!actions.length) acts.remove();
  for (const a of actions) {
    const b = document.createElement("button");
    b.className = `btn ${a.primary ? "primary" : ""}`;
    b.textContent = a.label;
    b.onclick = () => { a.run(); el.remove(); };
    acts.append(b);
  }
  el.querySelector<HTMLButtonElement>(".x")!.onclick = () => el.remove();
  ui.toasts.append(el);
  if (kind !== "error") setTimeout(() => el.remove(), actions.length ? 14000 : 6000);
}

function setStatus(text: string, done?: number, total?: number) {
  ui.status.textContent = text;
  ui.progress.hidden = total === undefined;
  if (total) ui.bar.style.width = `${(100 * (done ?? 0)) / total}%`;
}

function updateCounts() {
  const n = state.papers.length;
  const check = state.papers.filter(isCheck).length;
  const done = state.papers.filter(isDone).length;
  ui.cnt.all.textContent = String(n);
  ui.cnt.check.textContent = String(check);
  ui.cnt.done.textContent = String(done);
  ui.summary.textContent = n ? `共 ${n} 篇　需核对 ${check}　已完成 ${done}` : "";
  ui.folder.textContent = n ? `— ${basename(state.baseFolder)}（${n} 篇）` : "";
}

function updateHint() {
  ui.hint.textContent = !state.papers.length
    ? "把论文文件夹拖进窗口即可开始"
    : state.view === "table"
      ? "双击单元格或按 F2 修改　·　Ctrl+Enter 确认并跳到下一篇"
      : "点击封面查看详情　·　方向键切换　·　Ctrl+Enter 确认并跳到下一篇";
}

function setBusy(busy: boolean) {
  state.busy = busy;
  for (const b of $$<HTMLButtonElement>("#open-folder, #pick, #export-excel, #export-pdf, #export-png, #roster-btn, #rename-btn")) {
    b.disabled = busy;
  }
  $<HTMLButtonElement>("#rescan").disabled = busy || !state.sources.length;
}

// ---------------------------------------------------------------- 窗口按钮（系统标题栏已隐藏）

const appWindow = getCurrentWindow();
$("#win-min").onclick = () => appWindow.minimize();
$("#win-max").onclick = () => appWindow.toggleMaximize();
$("#win-close").onclick = () => appWindow.close();

async function syncMaximized() {
  const max = await appWindow.isMaximized();
  ui.app.classList.toggle("maximized", max);
  $("#win-max").title = max ? "还原" : "最大化";
}
appWindow.onResized(syncMaximized);
syncMaximized();

// ---------------------------------------------------------------- 主题

function applyTheme(theme: "light" | "dark") {
  document.documentElement.dataset.theme = theme;
  ui.themeLabel.textContent = theme === "dark" ? "浅色模式" : "深色模式";
}

function savedTheme(): string | null {
  try {
    return localStorage.getItem("theme");
  } catch {
    return null;
  }
}

$("#theme-btn").onclick = () => {
  const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  applyTheme(next);
  try {
    localStorage.setItem("theme", next);
  } catch {}
};
// 老师没有手动选过时，跟随系统的深浅色设置
matchMedia("(prefers-color-scheme: dark)").addEventListener("change", (e) => {
  if (!savedTheme()) applyTheme(e.matches ? "dark" : "light");
});

// ---------------------------------------------------------------- 视图切换

function setView(view: View) {
  state.view = view;
  ui.app.dataset.view = view;
  for (const b of $$<HTMLButtonElement>(".view-btn")) b.setAttribute("aria-pressed", String(b.dataset.view === view));
  updateHint();
  render();
  scrollToSelected();
}
for (const b of $$<HTMLButtonElement>(".view-btn")) b.onclick = () => setView(b.dataset.view as View);

function visiblePapers() {
  const q = state.query.trim().toLowerCase();
  return state.papers.filter(
    (p) =>
      (state.filter === "all" || (state.filter === "check" ? isCheck(p) : isDone(p))) &&
      (!q || [p.title, p.name, p.studentId, p.advisor, p.file].some((v) => v.toLowerCase().includes(q))),
  );
}

/** 重新绘制当前视图（列表内容变了时调用） */
function render() {
  const list = visiblePapers();
  if (state.view === "table") renderTable(list);
  else renderGallery(list);
  ui.noMatch.hidden = list.length > 0 || !state.papers.length;
  updateCounts();
  showDetail();
}

function updatePaper(p: Paper) {
  if (state.view === "table") updateRow(p);
  else updateCard(p);
  updateCounts();
  if (p.id === state.selected) showDetail();
}

function setFilter(f: Filter) {
  state.filter = f;
  for (const b of $$<HTMLButtonElement>(".filters button")) b.classList.toggle("on", b.dataset.filter === f);
  const list = visiblePapers();
  if (!list.some((p) => p.id === state.selected)) state.selected = list[0]?.id ?? null;
  render();
  scrollToSelected();
}
for (const b of $$<HTMLButtonElement>(".filters button")) b.onclick = () => setFilter(b.dataset.filter as Filter);
ui.search.oninput = () => {
  state.query = ui.search.value;
  render();
};

function select(id: number | null) {
  state.selected = id;
  if (state.view === "table") {
    for (const tr of ui.rows.querySelectorAll("tr.sel")) tr.classList.remove("sel");
    for (const td of ui.rows.querySelectorAll("td.active")) td.classList.remove("active");
    const tr = id === null ? null : ui.rows.querySelector(`tr[data-id="${id}"]`);
    tr?.classList.add("sel");
    tr?.children[state.col]?.classList.add("active");
  } else {
    for (const c of ui.grid.querySelectorAll(".card.sel")) c.classList.remove("sel");
    if (id !== null) ui.grid.querySelector(`.card[data-id="${id}"]`)?.classList.add("sel");
  }
  scrollToSelected();
  showDetail();
}

function scrollToSelected() {
  if (state.selected === null) return;
  const el =
    state.view === "table"
      ? ui.rows.querySelector(`tr[data-id="${state.selected}"]`)
      : ui.grid.querySelector(`.card[data-id="${state.selected}"]`);
  el?.scrollIntoView({ block: "nearest" });
}

function move(step: number) {
  const list = visiblePapers();
  const i = list.findIndex((p) => p.id === state.selected);
  const next = list[Math.min(list.length - 1, Math.max(0, i + step))];
  if (next) select(next.id);
}

// ---------------------------------------------------------------- 表格视图

function cellText(p: Paper, key: (typeof COLS)[number]["key"]) {
  if (key === "status") {
    const st = statusOf(p);
    const text = st === "manual" && p.method === Method.Manual ? "✎ 已修改" : STATUS_TEXT[st];
    return `<span class="st ${st}">${st === "pending" ? `<span class="spin"></span>` : ""}${text}</span>`;
  }
  if (key === "title" && p.state === "pending") return "正在识别…";
  if (key === "no") return String(p.id + 1);
  return esc(String(p[key]));
}

function rowHtml(p: Paper) {
  const st = statusOf(p);
  const sel = p.id === state.selected;
  const cells = COLS.map((c, i) => {
    const cls = [
      c.key === "no" ? "n" : "",
      c.key === "title" ? (p.state === "pending" ? "t pending" : "t") : "",
      c.key === "file" ? "file" : "",
      sel && i === state.col ? "active" : "",
    ].join(" ");
    const tip = { title: p.title, file: p.path, method: p.method } as Record<string, string>;
    const title = c.key in tip ? ` title="${esc(tip[c.key])}"` : "";
    return `<td class="${cls}" data-col="${i}"${title}>${cellText(p, c.key)}</td>`;
  }).join("");
  return `<tr data-id="${p.id}" class="${st === "warn" || st === "err" ? "warn" : ""} ${sel ? "sel" : ""}">${cells}</tr>`;
}

function renderTable(list: Paper[]) {
  ui.rows.innerHTML = list.map(rowHtml).join("");
}

function updateRow(p: Paper) {
  const tr = ui.rows.querySelector(`tr[data-id="${p.id}"]`);
  if (tr && !tr.querySelector("td.editing")) tr.outerHTML = rowHtml(p);
}

ui.rows.addEventListener("mousedown", (e) => {
  const td = (e.target as HTMLElement).closest<HTMLTableCellElement>("td");
  if (!td || td.classList.contains("editing")) return;
  state.col = Math.max(1, Number(td.dataset.col));
  select(Number(td.parentElement!.dataset.id));
});
ui.rows.addEventListener("dblclick", (e) => {
  const td = (e.target as HTMLElement).closest<HTMLTableCellElement>("td");
  if (!td) return;
  const col = COLS[Number(td.dataset.col)];
  const p = state.papers[Number(td.parentElement!.dataset.id)];
  if (col.editable) startEdit();
  else if (col.key === "file") openPath(p.path);
});

/** 在当前单元格里直接编辑。initial 不为空时用它替换原内容（像 Excel 一样直接打字） */
function startEdit(initial?: string) {
  const p = current();
  const col = COLS[state.col];
  if (!p || !col.editable || p.state === "pending" || state.view !== "table") return;
  const key = col.key as EditKey;
  const td = ui.rows.querySelector<HTMLTableCellElement>(`tr[data-id="${p.id}"] td[data-col="${state.col}"]`);
  if (!td || td.classList.contains("editing")) return;
  const input = document.createElement("input");
  input.value = initial ?? p[key];
  td.classList.add("editing");
  td.replaceChildren(input);
  input.focus();
  if (initial === undefined) input.select();

  let finished = false;
  const finish = (commit: boolean, then?: () => void) => {
    if (finished) return;
    finished = true;
    if (commit) setField(p, key, input.value);
    td.classList.remove("editing");
    updateRow(p);
    updateCounts();
    showDetail();
    then?.();
  };
  input.onkeydown = (e) => {
    e.stopPropagation();
    if (e.key === "Enter" && e.ctrlKey) finish(true, confirmAndNext);
    else if (e.key === "Enter") finish(true, () => move(1));
    else if (e.key === "Tab") {
      e.preventDefault();
      finish(true, () => moveCol(e.shiftKey ? -1 : 1));
    } else if (e.key === "Escape") finish(false);
  };
  input.onblur = () => finish(true);
}

function moveCol(step: number) {
  state.col = Math.min(COLS.length - 1, Math.max(1, state.col + step));
  select(state.selected);
}

function setField(p: Paper, key: EditKey, raw: string) {
  const value = key === "title" ? raw.replace(/\s*\n\s*/g, "").trim() : raw.trim();
  if (value === p[key]) return;
  p[key] = value;
  if (key === "title") p.method = Method.Manual;
  scheduleSave();
}

// 编辑栏（fx）：显示并可修改当前单元格
ui.fx.addEventListener("input", () => {
  const p = current();
  const col = COLS[state.col];
  if (!p || !col.editable) return;
  setField(p, col.key as EditKey, ui.fx.value);
  updateRow(p);
  updateCounts();
  showNote(p);
});
ui.fx.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.ctrlKey) {
    e.preventDefault();
    ui.fx.blur();
    move(1);
  }
});

// ---------------------------------------------------------------- 整体预览（封面墙）

const thumbCache = new Map<string, string>(); // 路径 -> 缩略图 objectURL
const thumbQueue: string[] = [];
let thumbWorkers = 0;

function cardHtml(p: Paper) {
  const st = statusOf(p);
  const meta = [p.name, p.studentId].filter(Boolean).join(" · ") || p.file;
  const thumb = thumbCache.get(p.path);
  const img = p.state === "error" ? `<span class="ph">无法打开</span>` : thumb ? `<img src="${thumb}" alt="">` : `<span class="ph">…</span>`;
  return `<div class="card ${st === "warn" || st === "err" ? "warn" : ""} ${p.id === state.selected ? "sel" : ""}" data-id="${p.id}">
    <div class="thumb" data-path="${esc(p.path)}">${img}<span class="num">${String(p.id + 1).padStart(2, "0")}</span><span class="badge ${st}">${BADGE[st]}</span></div>
    <div class="ct">${esc(p.state === "pending" ? "正在识别…" : p.title || "（未识别出题目）")}</div>
    <div class="cm">${esc(meta)}</div></div>`;
}

const thumbObserver = new IntersectionObserver(
  (entries) => {
    for (const e of entries) {
      if (!e.isIntersecting) continue;
      thumbObserver.unobserve(e.target);
      const path = (e.target as HTMLElement).dataset.path!;
      if (!thumbCache.has(path) && !thumbQueue.includes(path)) thumbQueue.push(path);
    }
    pumpThumbs();
  },
  { root: ui.grid, rootMargin: "300px" },
);

/** 只渲染进入可视区域的封面，最多同时渲染 2 张 */
function pumpThumbs() {
  while (thumbWorkers < 2 && thumbQueue.length) {
    const path = thumbQueue.shift()!;
    thumbWorkers++;
    readFile(path)
      .then((data) => renderCover(data, 0.6))
      .then((blob) => {
        const url = URL.createObjectURL(blob);
        thumbCache.set(path, url);
        for (const t of ui.grid.querySelectorAll<HTMLElement>(".thumb")) {
          if (t.dataset.path === path) t.querySelector(".ph")?.replaceWith(Object.assign(new Image(), { src: url, alt: "" }));
        }
      })
      .catch(() => {})
      .finally(() => {
        thumbWorkers--;
        pumpThumbs();
      });
  }
}

/** 还没有缩略图的封面，进入可视区域时再渲染 */
function observeThumb(card: Element | null, p: Paper) {
  const thumb = card?.querySelector<HTMLElement>(".thumb");
  if (thumb && p.state !== "error" && !thumbCache.has(p.path)) thumbObserver.observe(thumb);
}

function renderGallery(list: Paper[]) {
  thumbObserver.disconnect();
  ui.grid.innerHTML = list.map(cardHtml).join("");
  for (const p of list) observeThumb(ui.grid.querySelector(`.card[data-id="${p.id}"]`), p);
}

function updateCard(p: Paper) {
  const card = ui.grid.querySelector(`.card[data-id="${p.id}"]`);
  if (!card) return;
  card.outerHTML = cardHtml(p);
  observeThumb(ui.grid.querySelector(`.card[data-id="${p.id}"]`), p);
}

ui.grid.addEventListener("click", (e) => {
  const card = (e.target as HTMLElement).closest<HTMLElement>(".card");
  if (card) select(Number(card.dataset.id));
});
ui.grid.addEventListener("dblclick", (e) => {
  const card = (e.target as HTMLElement).closest<HTMLElement>(".card");
  if (card) openPath(state.papers[Number(card.dataset.id)].path);
});
$("#drawer-close").onclick = () => select(null);

/** 封面墙每行有几张，用于上下方向键 */
function gridColumns() {
  const cards = ui.grid.querySelectorAll<HTMLElement>(".card");
  if (cards.length < 2) return 1;
  const top = cards[0].offsetTop;
  let n = 0;
  for (const c of cards) {
    if (c.offsetTop !== top) break;
    n++;
  }
  return n;
}

// ---------------------------------------------------------------- 详情（表格底部 / 封面墙右侧）

const coverCache = new Map<string, string>(); // 路径 -> 大封面 objectURL
let coverToken = 0;

async function coverUrl(path: string) {
  let url = coverCache.get(path);
  if (!url) {
    url = URL.createObjectURL(await renderCover(await readFile(path), 1.3));
    coverCache.set(path, url);
    if (coverCache.size > 30) {
      const [oldPath, oldUrl] = coverCache.entries().next().value!;
      URL.revokeObjectURL(oldUrl);
      coverCache.delete(oldPath);
    }
  }
  return url;
}

async function showCover(p: Paper | null) {
  const box = state.view === "table" ? ui.paneCover : ui.drawerCover;
  const token = ++coverToken;
  if (!p) {
    box.innerHTML = "";
    return;
  }
  if (p.state === "error") {
    box.innerHTML = `<span>无法显示封面</span>`;
    return;
  }
  if (!coverCache.has(p.path)) box.innerHTML = `<span>正在生成封面…</span>`;
  try {
    const url = await coverUrl(p.path);
    if (token === coverToken) box.innerHTML = `<img alt="论文封面" src="${url}">`;
  } catch {
    if (token === coverToken) box.innerHTML = `<span>无法显示封面</span>`;
  }
}

function noteFor(p: Paper): [string, string] {
  if (p.state === "pending") return ["pending", "正在识别这篇论文…"];
  if (p.state === "error") return ["err", `无法打开这个文件（${esc(p.method.replace("打开失败: ", ""))}），可以直接手动填写。`];
  switch (p.method) {
    case Method.Label:
      return ["", "✓ 已从封面的「题目」栏自动识别。"];
    case Method.Font:
      return ["", "✓ 已按页面上字号最大的标题自动识别。"];
    case Method.Layout:
      return ["", "✓ 已按封面版式自动识别（取栏目上方的标题）。"];
    case Method.OcrLabel:
    case Method.OcrFont:
    case Method.OcrLayout:
      return ["warn", "⚠ 扫描件，题目由 OCR 识别，可能有个别错字。请对照封面核对，无误后点「确认无误」。"];
    case Method.Meta:
      return ["warn", "⚠ 没能从页面中识别出题目，暂用 PDF 属性里的标题，请对照封面核对。"];
    case Method.FileName:
      return ["warn", "⚠ 没能从页面中识别出题目，暂用文件名代替，请对照封面填写正确的题目。"];
    case Method.Scan:
      return ocrRun.used
        ? ["warn", "⚠ 扫描件，OCR 也没能读出文字（页面可能是空白或太模糊），暂用文件名代替，请对照封面填写。"]
        : ["warn", "⚠ 扫描件，暂用文件名代替。可在「选项」里打开扫描件 OCR 后重新识别，或对照封面手动填写。"];
    case Method.Manual:
      return ["manual", "✎ 题目已人工修改。"];
    case Method.Confirmed:
      return ["manual", "✎ 已人工确认无误。"];
    default:
      return ["", esc(p.method)];
  }
}

function showNote(p: Paper) {
  const [cls, html] = noteFor(p);
  for (const el of [ui.paneNote, ui.drawerNote]) {
    el.className = `note ${cls}`;
    el.innerHTML = html;
  }
}

const drawerFields: [HTMLInputElement | HTMLTextAreaElement, EditKey][] = [
  [$("#d-title"), "title"],
  [$("#d-name"), "name"],
  [$("#d-sid"), "studentId"],
  [$("#d-adv"), "advisor"],
  [$("#d-college"), "college"],
  [$("#d-major"), "major"],
  [$("#d-rank"), "advisorTitle"],
];

function showDetail() {
  const p = current();
  const list = visiblePapers();
  const i = p ? list.findIndex((x) => x.id === p.id) : -1;
  const last = i < 0 || i >= list.length - 1;

  // 编辑栏
  const col = COLS[state.col];
  ui.cellName.textContent = p && i >= 0 ? `${col.letter}${p.id + 1}` : "";
  if (document.activeElement !== ui.fx) {
    ui.fx.value = p && col.key !== "no" && col.key !== "status" ? String(p[col.key as EditKey | "method" | "file"]) : "";
  }
  ui.fx.disabled = !p || !col.editable || p.state === "pending";

  // 确认 / 下一篇按钮
  const confirming = !!p && isCheck(p) && p.state === "done";
  for (const b of $$<HTMLButtonElement>(".next-btn")) {
    b.textContent = confirming ? (last ? "✓ 确认无误" : "✓ 确认无误，下一篇") : last ? "已是最后一篇" : "下一篇 →";
    b.disabled = !p || (!confirming && last);
  }
  $<HTMLButtonElement>("#drawer-prev").disabled = i <= 0;

  ui.drawer.hidden = !p || state.view !== "gallery";
  if (!p) {
    ui.paneTitle.textContent = "选择一篇论文查看封面";
    ui.paneNote.className = "note pending";
    ui.paneNote.textContent = "点击表格中的任意一行。";
    ui.paneKv.innerHTML = "";
    showCover(null);
    return;
  }

  showNote(p);
  ui.paneTitle.textContent = p.state === "pending" ? "正在识别…" : p.title || "（未识别出题目）";
  const kv = (k: string, v: string | number) => `<dt>${k}</dt><dd>${esc(String(v || "—"))}</dd>`;
  ui.paneKv.innerHTML =
    kv("学生", [p.name, p.studentId].filter(Boolean).join(" · ")) +
    kv("专业", [p.college, p.major].filter(Boolean).join(" · ")) +
    kv("导师", [p.advisor, p.advisorTitle].filter(Boolean).join(" · ")) +
    kv("文件", `${p.file}${p.pages ? ` · ${p.pages} 页` : ""}`);

  ui.drawerPos.textContent = i >= 0 ? `第 ${i + 1} 篇 / 共 ${list.length} 篇` : "";
  ui.drawerFile.textContent = basename(p.path);
  ui.drawerMeta.textContent = [p.pages ? `${p.pages} 页` : "", p.method].filter(Boolean).join(" · ");
  for (const [el, key] of drawerFields) {
    el.disabled = p.state === "pending";
    if (document.activeElement !== el) el.value = p[key];
  }
  showCover(p);
}

for (const [el, key] of drawerFields) {
  el.addEventListener("input", () => {
    const p = current();
    if (!p || p.state === "pending") return;
    setField(p, key, el.value);
    updateCard(p);
    updateCounts();
    showNote(p);
  });
}
$("#d-title").addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.ctrlKey) e.preventDefault(); // 题目不允许换行
});

/** 「确认无误，下一篇」/「下一篇」 */
function confirmAndNext() {
  const p = current();
  if (!p) return;
  const before = visiblePapers();
  const i = before.findIndex((x) => x.id === p.id);
  const nextId = before[i + 1]?.id ?? null;
  const confirming = isCheck(p) && p.state === "done";
  const fromCheckView = state.filter === "check";
  if (confirming) {
    p.method = Method.Confirmed;
    scheduleSave();
  }

  const allChecked = !state.papers.some(isCheck);
  if (allChecked && fromCheckView) {
    // 需核对的都处理完了：回到全部，停在刚确认的这篇
    state.filter = "all";
    for (const b of $$<HTMLButtonElement>(".filters button")) b.classList.toggle("on", b.dataset.filter === "all");
    state.selected = p.id;
  } else {
    // 在「需核对」里确认到末尾时，从头找剩下的
    const wrap = fromCheckView ? visiblePapers().find((x) => x.id !== p.id)?.id : undefined;
    state.selected = nextId ?? wrap ?? p.id;
  }
  render();
  scrollToSelected();
  if (allChecked && (confirming || fromCheckView)) {
    toast("全部核对完了，可以导出了。", "ok", [{ label: "导出 Excel", run: showExportDialog, primary: true }]);
  }
}

for (const b of $$<HTMLButtonElement>(".next-btn")) b.onclick = confirmAndNext;
$("#drawer-prev").onclick = () => move(-1);
ui.paneCover.onclick = ui.drawerCover.onclick = () => {
  const p = current();
  if (p) openPath(p.path);
};
for (const b of $$<HTMLButtonElement>("[data-act]")) {
  b.onclick = () => {
    const p = current();
    if (p) (b.dataset.act === "open" ? openPath : revealPath)(p.path);
  };
}

// ---------------------------------------------------------------- 键盘

document.addEventListener("keydown", (e) => {
  if (!ui.modal.hidden) return;
  const inField = e.target instanceof Element && e.target.matches("input, textarea");
  if (e.key === "Escape") {
    ui.optionsPop.hidden = true;
    if (inField) (e.target as HTMLElement).blur();
    else if (state.view === "gallery") select(null);
    return;
  }
  if (!state.papers.length) return;
  if (e.ctrlKey && e.key === "Enter") {
    e.preventDefault();
    confirmAndNext();
    return;
  }
  if (e.ctrlKey && e.key.toLowerCase() === "f") {
    e.preventDefault();
    ui.search.focus();
    return;
  }
  if (inField || e.ctrlKey || e.altKey || e.metaKey) return;

  if (state.view === "table") {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") move(e.key === "ArrowDown" ? 1 : -1);
    else if (e.key === "ArrowLeft" || e.key === "ArrowRight") moveCol(e.key === "ArrowRight" ? 1 : -1);
    else if (e.key === "Enter") move(1);
    else if (e.key === "Tab") moveCol(e.shiftKey ? -1 : 1);
    else if (e.key === "F2") startEdit();
    else if (e.key.length === 1 && COLS[state.col].editable) startEdit(e.key); // 选中单元格后直接打字即可替换内容
    else return;
    e.preventDefault();
  } else {
    const cols = gridColumns();
    const step = ({ ArrowLeft: -1, ArrowRight: 1, ArrowUp: -cols, ArrowDown: cols } as Record<string, number>)[e.key];
    if (step === undefined) return;
    e.preventDefault();
    move(step);
  }
});

// ---------------------------------------------------------------- 提取

function relative(path: string, base: string) {
  const b = base.replace(/[\\/]+$/, "");
  return path.toLowerCase().startsWith(b.toLowerCase()) ? path.slice(b.length).replace(/^[\\/]+/, "") : basename(path);
}

/** 本轮提取中 OCR 是否可用：第一次失败后就不再尝试，并只提示一次；used 表示确实运行过 */
const ocrRun = { enabled: false, warnedLang: false, used: false };

async function ocrForExtract(page: Parameters<typeof ocrPage>[0]) {
  if (!ocrRun.enabled) return [];
  try {
    ocrRun.used = true;
    return await ocrPage(page, (lang) => {
      if (!lang.startsWith("zh") && !ocrRun.warnedLang) {
        ocrRun.warnedLang = true;
        toast(
          `这台电脑没有安装中文 OCR（当前为 ${lang || "未知"}），扫描件的中文题目可能识别不准。\n` +
            "可在「设置 → 时间和语言 → 语言」中添加「中文(简体)」并勾选「光学字符识别」。",
          "warn",
        );
      }
    });
  } catch (e) {
    ocrRun.enabled = ocrRun.used = false;
    toast(`OCR 不可用，扫描件将改用文件名：${errText(e)}`, "warn");
    return [];
  }
}

async function extractOne(p: Paper) {
  try {
    const task = loadPdf(await readFile(p.path));
    try {
      const options = { ocr: state.opts.ocr ? ocrForExtract : undefined };
      Object.assign(p, await extract(await task.promise, p.path, options), { state: "done" });
    } finally {
      await task.destroy();
    }
  } catch (e) {
    p.state = "error";
    p.method = (e as Error)?.name === "PasswordException" ? "打开失败: 有密码" : "打开失败: 文件损坏";
  }
}

async function start(sources: string[]) {
  if (state.busy || !sources.length) return;
  let files: { path: string; size: number; mtime: number }[];
  try {
    files = await collectPdfs(sources, state.opts.recursive);
  } catch (e) {
    return toast(errText(e), "error");
  }
  if (!files.length) {
    const hint = state.opts.recursive ? "" : "\n如果论文放在子文件夹里，请在「选项」里打开「包含子文件夹」。";
    return toast(`这里没有找到 PDF 文件。${hint}`, "warn");
  }

  state.sources = sources;
  state.baseFolder = sources.length === 1 && (await isDir(sources[0])) ? sources[0] : dirname(files[0].path);
  state.papers = files.map(({ path, size, mtime }, id) => ({
    id, path, size, mtime, file: relative(path, state.baseFolder), state: "pending",
    title: "", method: "", name: "", studentId: "", advisor: "", advisorTitle: "", college: "", major: "", pages: 0,
  }));
  const saved = loadProgress(state.baseFolder);
  let restored = 0;
  ui.saved.textContent = "";
  state.selected = 0;
  state.col = 2;
  state.filter = "all";
  state.query = ui.search.value = "";
  for (const b of $$<HTMLButtonElement>(".filters button")) b.classList.toggle("on", b.dataset.filter === "all");
  ui.app.dataset.empty = "false";
  updateHint();
  render();
  setBusy(true);
  Object.assign(ocrRun, { enabled: state.opts.ocr, warnedLang: false, used: false });

  const queue = [...state.papers];
  let done = 0;
  const worker = async () => {
    for (let p = queue.shift(); p; p = queue.shift()) {
      await extractOne(p);
      p.auto = { ...pickFields(p), method: p.method };
      // 上次在这个文件夹里改过、确认过的内容（文件没被替换才套用）
      const s = saved[p.file];
      if (s && s.size === p.size && s.mtime === p.mtime) {
        Object.assign(p, pickFields(s));
        if (s.method === Method.Manual || s.method === Method.Confirmed) p.method = s.method;
        p.restored = true;
        restored++;
      }
      updatePaper(p);
      setStatus(`正在识别 ${++done} / ${files.length}`, done, files.length);
    }
  };
  await Promise.all([worker(), worker()]);
  setBusy(false);
  setStatus("就绪");
  showDetail();
  if (restored) {
    toast(`已恢复上次在这个文件夹里的修改（${restored} 篇）。`, "ok", [{ label: "放弃这些修改", run: discardRestored }]);
  }

  const check = state.papers.filter(isCheck);
  if (!check.length) {
    const msg = restored ? `全部 ${files.length} 篇都已识别或核对完毕。` : `全部 ${files.length} 篇都已自动识别。`;
    toast(msg, "ok", [{ label: "导出 Excel", run: showExportDialog, primary: true }]);
    return;
  }
  const ocr = check.filter((p) => p.method.startsWith("OCR")).length;
  const lines = [
    ocr ? `${ocr} 篇是扫描件，已用 OCR 识别，可能有个别错字` : "",
    check.length - ocr ? `${check.length - ocr} 篇没能识别出题目，需要对照封面填写` : "",
  ].filter(Boolean);
  toast(`识别完成，有 ${check.length} 篇需要核对：\n${lines.join("；\n")}。`, "warn", [
    { label: "开始逐篇核对", primary: true, run: () => setFilter("check") },
  ]);
}

async function pickFolder() {
  if (state.busy) return;
  const dir = await open({ directory: true, title: "选择存放论文 PDF 的文件夹" });
  if (typeof dir === "string") await start([dir]);
}

$("#open-folder").onclick = pickFolder;
$("#pick").onclick = pickFolder;
$("#rescan").onclick = () => start(state.sources);

getCurrentWebview().onDragDropEvent(async ({ payload }) => {
  if (payload.type === "enter" || payload.type === "over") ui.overlay.hidden = state.busy;
  else if (payload.type === "leave") ui.overlay.hidden = true;
  else if (payload.type === "drop") {
    ui.overlay.hidden = true;
    if (!state.busy && payload.paths.length) await start(payload.paths);
  }
});

// ---------------------------------------------------------------- 选项（记住老师的选择）

function loadOpts() {
  for (const k of Object.keys(state.opts) as Opt[]) {
    try {
      const saved = localStorage.getItem(`opt:${k}`);
      if (saved !== null) state.opts[k] = saved === "1";
    } catch {}
  }
  for (const box of $$<HTMLInputElement>("input[data-opt]")) box.checked = state.opts[box.dataset.opt as Opt];
}

for (const box of $$<HTMLInputElement>("input[data-opt]")) {
  box.onchange = () => {
    const k = box.dataset.opt as Opt;
    state.opts[k] = box.checked;
    try {
      localStorage.setItem(`opt:${k}`, box.checked ? "1" : "0");
    } catch {}
    if (state.sources.length) setStatus("选项已更改，点「重新识别」后生效");
  };
}

$("#options-btn").onclick = (e) => {
  e.stopPropagation();
  ui.optionsPop.hidden = !ui.optionsPop.hidden;
};
ui.optionsPop.onclick = (e) => e.stopPropagation();
document.addEventListener("click", () => (ui.optionsPop.hidden = true));

// ---------------------------------------------------------------- 导出

/** 保存对话框默认在论文所在文件夹，防止手滑选中某篇论文把原件覆盖掉 */
function isPaperFile(target: string) {
  const norm = (p: string) => p.replace(/\//g, "\\").toLowerCase();
  if (state.papers.some((p) => norm(p.path) === norm(target))) {
    toast("不能保存到论文原文件上，这会覆盖学生的论文。请换一个文件名。", "error");
    return true;
  }
  return false;
}

const openActions = (path: string): ToastAction[] => [
  { label: "打开", run: () => openPath(path), primary: true },
  { label: "在文件夹中显示", run: () => revealPath(path) },
];

async function exportExcel(opts: ExcelOptions) {
  const target = await save({
    title: "保存 Excel",
    defaultPath: joinPath(state.baseFolder, "论文题目汇总.xlsx"),
    filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
  });
  if (!target || isPaperFile(target)) return;
  try {
    await writeFile(target, await buildExcel(state.papers, opts));
    toast(`已导出 ${state.papers.length} 篇论文到 Excel。`, "ok", openActions(target));
  } catch (e) {
    toast(errText(e), "error");
  }
}

async function exportCovers(kind: "pdf" | "png") {
  const papers = state.papers.filter((p) => p.state === "done");
  let target: string | null;
  if (kind === "pdf") {
    target = await save({
      title: "保存封面合集",
      defaultPath: joinPath(state.baseFolder, "论文封面合集.pdf"),
      filters: [{ name: "PDF 文件", extensions: ["pdf"] }],
    });
  } else {
    const dir = await open({ directory: true, title: "选择封面图片的保存位置", defaultPath: state.baseFolder });
    target = typeof dir === "string" ? joinPath(dir, "论文封面") : null;
  }
  if (!target || isPaperFile(target)) return;

  setBusy(true);
  try {
    const progress = (d: number, t: number) => setStatus(`正在导出封面 ${d} / ${t}`, d, t);
    const res = kind === "pdf" ? await exportCoverPdf(papers, target, progress) : await exportCoverImages(papers, target, progress);
    const what = kind === "pdf" ? "合并为一个 PDF" : "保存为图片";
    toast(`已把 ${res.count} 篇论文的封面${what}。`, "ok", openActions(target));
    if (res.failed.length) toast(`以下文件导出失败：\n${res.failed.join("\n")}`, "warn");
  } catch (e) {
    toast(errText(e), "error");
  } finally {
    setBusy(false);
    setStatus("就绪");
  }
}

async function copyTitles() {
  const text = state.papers.map((p) => p.title).join("\n");
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const ta = Object.assign(document.createElement("textarea"), { value: text });
    document.body.append(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
  toast(`已复制 ${state.papers.length} 个题目，可以直接粘贴到 Word 或 Excel。`);
}

$("#export-excel").onclick = showExportDialog;
$("#export-pdf").onclick = () => exportCovers("pdf");
$("#export-png").onclick = () => exportCovers("png");
$("#copy-titles").onclick = copyTitles;

// ---------------------------------------------------------------- 对话框

interface DialogAction {
  label: string;
  primary?: boolean;
  id?: string;
  /** 返回 true 表示保持对话框打开 */
  run?: () => boolean | void | Promise<boolean | void>;
}

function openDialog(title: string, body: string, actions: DialogAction[], opts: { wide?: boolean; foot?: string } = {}) {
  const back = ui.modal;
  back.innerHTML = `<div class="modal ${opts.wide ? "wide" : ""}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
    <div class="modal-head"><h2>${esc(title)}</h2>
      <button class="icon-btn" data-close title="关闭"><svg viewBox="0 0 24 24"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>
    <div class="modal-body">${body}</div>
    <div class="modal-foot"><span class="grow">${opts.foot ?? ""}</span></div></div>`;
  const foot = back.querySelector(".modal-foot")!;
  const onKey = (e: KeyboardEvent) => {
    if (e.key === "Escape") {
      e.stopPropagation();
      close();
    }
  };
  const close = () => {
    back.hidden = true;
    back.innerHTML = "";
    document.removeEventListener("keydown", onKey, true);
  };
  for (const a of actions) {
    const b = document.createElement("button");
    b.className = `btn ${a.primary ? "primary" : ""}`;
    b.textContent = a.label;
    if (a.id) b.id = a.id;
    b.onclick = async () => {
      if ((await a.run?.()) !== true) close();
    };
    foot.append(b);
  }
  back.querySelector<HTMLButtonElement>("[data-close]")!.onclick = close;
  back.onmousedown = (e) => {
    if (e.target === back) close();
  };
  document.addEventListener("keydown", onKey, true);
  back.hidden = false;
  return { el: back.querySelector<HTMLElement>(".modal")!, foot: foot.querySelector<HTMLElement>(".grow")!, close };
}

async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const ta = Object.assign(document.createElement("textarea"), { value: text });
    document.body.append(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }
}

// ---------------------------------------------------------------- 自动保存核对进度

let saveTimer = 0;

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveNow, 400);
}

/** 只保存老师改过或确认过的论文；同一文件夹里没在本次打开的论文（比如只拖进来几个文件）保留原记录 */
function saveNow() {
  if (!state.baseFolder) return;
  const all = loadProgress(state.baseFolder);
  for (const p of state.papers) {
    delete all[p.file];
    if (!p.auto || p.state === "pending") continue;
    const manual = p.method === Method.Manual || p.method === Method.Confirmed;
    if (!manual && sameFields(p, p.auto)) continue;
    all[p.file] = { ...pickFields(p), size: p.size, mtime: p.mtime, method: p.method };
  }
  const ok = saveProgress(state.baseFolder, all);
  const time = new Date().toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" });
  ui.saved.textContent = ok ? `✓ 修改已自动保存 ${time}` : "⚠ 自动保存失败";
}

function discardRestored() {
  for (const p of state.papers) {
    if (!p.restored || !p.auto) continue;
    Object.assign(p, p.auto);
    p.restored = false;
  }
  saveNow();
  render();
  toast("已放弃上次的修改，恢复为自动识别的结果。");
}

// ---------------------------------------------------------------- 名单核对

function loadRoster(): typeof state.roster {
  try {
    const raw = localStorage.getItem("roster");
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

async function importRoster(): Promise<boolean> {
  const f = await open({
    title: "选择学生名单",
    filters: [{ name: "学生名单（Excel / CSV）", extensions: ["xlsx", "csv", "txt"] }],
  });
  if (typeof f !== "string") return false;
  try {
    const students = await parseRoster(await readFile(f), f);
    state.roster = { file: basename(f), students };
    try {
      localStorage.setItem("roster", JSON.stringify(state.roster));
    } catch {}
    return true;
  } catch (e) {
    toast(
      `读取名单失败：${errText(e)}\n请使用 .xlsx 或 .csv 文件，并包含「学号」或「姓名」列。旧版 .xls 请先在 Excel 里另存为 .xlsx。`,
      "error",
    );
    return false;
  }
}

const rosterResult = () =>
  state.roster ? matchRoster(state.roster.students, state.papers.filter((p) => p.state !== "pending")) : undefined;

async function showRoster() {
  if (!state.roster && !(await importRoster())) return;
  const roster = state.roster!;
  const r = rosterResult()!;
  type Tab = "missing" | "mismatch" | "extra" | "ok";
  let tab: Tab = r.missing.length ? "missing" : r.mismatched.length ? "mismatch" : r.extra.length ? "extra" : "ok";
  const counts: Record<Tab, number> = {
    missing: r.missing.length,
    mismatch: r.mismatched.length,
    extra: r.extra.length,
    ok: r.submitted.length,
  };
  const labels: Record<Tab, string> = { missing: "未交", mismatch: "信息不一致", extra: "不在名单", ok: "已交" };
  const body = `<p class="desc">名单：<b>${esc(roster.file)}</b>（${roster.students.length} 人）。按学号、姓名、文件名依次匹配论文。</p>
    <div class="stats">${(Object.keys(labels) as Tab[])
      .map((t) => `<button class="stat ${t}" data-tab="${t}"><b>${counts[t]}</b><span>${labels[t]}</span></button>`)
      .join("")}</div><div id="roster-list"></div>`;

  const dlg = openDialog(
    "名单核对",
    body,
    [
      {
        label: "更换名单",
        run: async () => {
          if (await importRoster()) {
            dlg.close();
            showRoster();
          }
          return true;
        },
      },
      {
        label: "清除名单",
        run: () => {
          state.roster = null;
          try {
            localStorage.removeItem("roster");
          } catch {}
          toast("已清除名单。");
        },
      },
      {
        label: "复制未交名单",
        run: async () => {
          await copyText(r.missing.map((s) => [s.id, s.name].filter(Boolean).join("\t")).join("\n"));
          toast(`已复制 ${r.missing.length} 位未交学生的学号和姓名。`);
          return true;
        },
      },
      { label: "关闭", primary: true },
    ],
    { wide: true, foot: "导出 Excel 时可以附带一张「名单核对」表" },
  );

  const list = dlg.el.querySelector<HTMLElement>("#roster-list")!;
  const paperRow = (id: number, studentId: string, name: string, title: string, note: string) =>
    `<tr data-paper="${id}" title="点击定位到这篇论文" style="cursor:pointer"><td>${esc(studentId)}</td><td>${esc(name)}</td><td class="wrap">${esc(title)}</td><td class="wrap">${esc(note)}</td></tr>`;
  const table = (heads: string[], widths: string[], rows: string) =>
    rows
      ? `<table class="mini"><colgroup>${widths.map((w) => `<col style="width:${w}">`).join("")}</colgroup>
         <thead><tr>${heads.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>${rows}</tbody></table>`
      : `<div class="empty-note">没有${labels[tab]}的情况</div>`;

  const draw = () => {
    for (const b of dlg.el.querySelectorAll<HTMLElement>(".stat")) b.classList.toggle("on", b.dataset.tab === tab);
    if (tab === "missing") {
      list.innerHTML = table(["学号", "姓名"], ["40%", "60%"], r.missing.map((s) => `<tr><td>${esc(s.id)}</td><td>${esc(s.name)}</td></tr>`).join(""));
    } else if (tab === "mismatch") {
      list.innerHTML = table(["名单学号", "名单姓名", "论文题目", "说明"], ["16%", "14%", "40%", "30%"],
        r.mismatched.map((m) => paperRow(m.paper.id, m.student.id, m.student.name, m.paper.title, m.reason)).join(""));
    } else if (tab === "extra") {
      list.innerHTML = table(["论文上的学号", "姓名", "论文题目", "文件名"], ["16%", "14%", "40%", "30%"],
        r.extra.map((p) => paperRow(p.id, p.studentId, p.name, p.title, p.file)).join(""));
    } else {
      list.innerHTML = table(["学号", "姓名", "论文题目", "说明"], ["16%", "14%", "50%", "20%"],
        r.submitted.flatMap((s) => s.papers.map((p) => paperRow(p.id, s.student.id, s.student.name, p.title,
          s.papers.length > 1 ? `交了 ${s.papers.length} 份` : ""))).join(""));
    }
  };
  for (const b of dlg.el.querySelectorAll<HTMLElement>(".stat")) {
    b.onclick = () => {
      tab = b.dataset.tab as Tab;
      draw();
    };
  }
  // 点论文行：关闭对话框并在主界面选中它
  list.onclick = (e) => {
    const tr = (e.target as HTMLElement).closest<HTMLElement>("tr[data-paper]");
    if (!tr) return;
    dlg.close();
    const id = Number(tr.dataset.paper);
    if (!visiblePapers().some((p) => p.id === id)) setFilter("all");
    select(id);
  };
  draw();
}

$("#roster-btn").onclick = showRoster;

// ---------------------------------------------------------------- 批量重命名

function showRename() {
  let pattern = DEFAULT_PATTERN;
  try {
    pattern = localStorage.getItem("rename:pattern") || DEFAULT_PATTERN;
  } catch {}
  const body = `<p class="desc">按模板修改 PDF 文件名。<b>不会覆盖任何已有文件</b>，改完后可以一键撤销。</p>
    <div class="pattern"><input id="rn-pattern" spellcheck="false" value="${esc(pattern)}" /></div>
    <div class="tokens">${Object.keys(TOKENS).map((t) => `<button data-token="${t}" title="插入${TOKENS[t].label}">${t}</button>`).join("")}</div>
    <label class="check-line"><input type="checkbox" id="rn-skip" checked /> 跳过还需要核对的论文（题目可能不准）</label>
    <table class="mini"><colgroup><col style="width:38%"><col style="width:4%"><col style="width:42%"><col style="width:16%"></colgroup>
      <thead><tr><th>原文件名</th><th></th><th>新文件名</th><th>情况</th></tr></thead><tbody id="rn-rows"></tbody></table>`;

  let plans: ReturnType<typeof planRenames> = [];
  const dlg = openDialog("批量重命名", body, [
    { label: "取消" },
    { label: "开始重命名", primary: true, id: "rn-go", run: () => doRename(plans.filter((pl) => pl.status === "ok")) },
  ], { wide: true });

  const input = dlg.el.querySelector<HTMLInputElement>("#rn-pattern")!;
  const skip = dlg.el.querySelector<HTMLInputElement>("#rn-skip")!;
  const rows = dlg.el.querySelector<HTMLElement>("#rn-rows")!;
  const go = dlg.el.querySelector<HTMLButtonElement>("#rn-go")!;
  const statusText: Record<string, string> = { ok: "将改名", same: "已符合", missing: "缺少信息", conflict: "名字冲突", skip: "跳过" };

  const update = () => {
    const papers = state.papers.filter((p) => p.state !== "pending");
    plans = planRenames(papers, input.value, (p) => skip.checked && isCheck(state.papers[p.id]));
    rows.innerHTML = plans
      .map((pl) => `<tr><td title="${esc(pl.from)}">${esc(basename(pl.from))}</td><td class="arrow">→</td>
        <td title="${esc(pl.newName)}">${pl.status === "ok" ? esc(pl.newName) : "<span class='arrow'>—</span>"}</td>
        <td title="${esc(pl.note)}"><span class="plan-st ${pl.status}">${statusText[pl.status]}</span> ${esc(pl.note)}</td></tr>`)
      .join("");
    const n = plans.filter((pl) => pl.status === "ok").length;
    const bad = plans.filter((pl) => pl.status === "conflict" || pl.status === "missing").length;
    go.textContent = n ? `重命名 ${n} 个文件` : "没有需要改名的文件";
    go.disabled = !n;
    dlg.foot.textContent = bad ? `${bad} 个文件缺少信息或名字冲突，不会改名` : "";
  };
  input.oninput = () => {
    try {
      localStorage.setItem("rename:pattern", input.value);
    } catch {}
    update();
  };
  skip.onchange = update;
  for (const b of dlg.el.querySelectorAll<HTMLButtonElement>("[data-token]")) {
    b.onclick = () => {
      const t = b.dataset.token!;
      const at = input.selectionStart ?? input.value.length;
      input.value = input.value.slice(0, at) + t + input.value.slice(input.selectionEnd ?? at);
      input.focus();
      input.setSelectionRange(at + t.length, at + t.length);
      input.dispatchEvent(new Event("input"));
    };
  }
  update();
}

/** 改路径后，封面缓存、文件名等跟着更新 */
function movePaper(p: Paper, to: string) {
  for (const cache of [thumbCache, coverCache]) {
    const url = cache.get(p.path);
    if (url) {
      cache.delete(p.path);
      cache.set(to, url);
    }
  }
  p.path = to;
  p.file = relative(to, state.baseFolder);
}

async function doRename(plans: ReturnType<typeof planRenames>) {
  const done: { id: number; from: string; to: string }[] = [];
  const failed: string[] = [];
  setBusy(true);
  for (const [i, pl] of plans.entries()) {
    setStatus(`正在重命名 ${i + 1} / ${plans.length}`, i + 1, plans.length);
    try {
      await renameFile(pl.from, pl.to);
      movePaper(state.papers[pl.paper.id], pl.to);
      done.push({ id: pl.paper.id, from: pl.from, to: pl.to });
    } catch (e) {
      failed.push(`${basename(pl.from)}：${errText(e)}`);
    }
  }
  setBusy(false);
  setStatus("就绪");
  saveNow();
  render();
  if (done.length) toast(`已重命名 ${done.length} 个文件。`, "ok", [{ label: "撤销", run: () => undoRename(done) }]);
  if (failed.length) toast(`以下文件没有改名：\n${failed.join("\n")}`, "warn");
}

async function undoRename(done: { id: number; from: string; to: string }[]) {
  const failed: string[] = [];
  for (const d of [...done].reverse()) {
    try {
      await renameFile(d.to, d.from);
      movePaper(state.papers[d.id], d.from);
    } catch (e) {
      failed.push(`${basename(d.to)}：${errText(e)}`);
    }
  }
  saveNow();
  render();
  toast(failed.length ? `部分文件没能改回：\n${failed.join("\n")}` : `已撤销，${done.length} 个文件恢复了原来的名字。`, failed.length ? "warn" : "ok");
}

$("#rename-btn").onclick = showRename;

// ---------------------------------------------------------------- 导出 Excel 设置

function loadExcelOpts(): ExcelOptions {
  try {
    const saved = JSON.parse(localStorage.getItem("excel:opts") ?? "null");
    if (saved) return { ...DEFAULT_EXCEL_OPTIONS, ...saved };
  } catch {}
  return { ...DEFAULT_EXCEL_OPTIONS };
}

function showExportDialog() {
  const o = loadExcelOpts();
  const sorts: [SortKey, string][] = [["file", "按文件顺序"], ["studentId", "按学号"], ["name", "按姓名"], ["title", "按题目"]];
  const body = `<p class="sec-title">导出哪些列</p>
    <div class="col-grid">${EXCEL_COLUMNS.map((c) =>
      `<label><input type="checkbox" value="${c.key}" ${o.columns.includes(c.key) ? "checked" : ""} /> ${c.header}</label>`).join("")}</div>
    <div class="form-line"><span>排序</span><select id="xl-sort">${sorts.map(([v, t]) =>
      `<option value="${v}" ${o.sort === v ? "selected" : ""}>${t}</option>`).join("")}</select></div>
    <label class="opt-line"><input type="checkbox" id="xl-hl" ${o.highlight ? "checked" : ""} /> 需要核对的行标黄</label>
    ${state.roster ? `<label class="opt-line" style="margin-top:8px"><input type="checkbox" id="xl-roster" checked /> 附带「名单核对」表（${esc(state.roster.file)}）</label>` : ""}`;
  const dlg = openDialog("导出 Excel", body, [
    { label: "取消" },
    {
      label: "导出…",
      primary: true,
      run: () => {
        const columns = [...dlg.el.querySelectorAll<HTMLInputElement>(".col-grid input:checked")].map((i) => i.value) as ExcelOptions["columns"];
        if (!columns.length) {
          toast("至少要勾选一列。", "warn");
          return true;
        }
        const opts: ExcelOptions = {
          columns,
          sort: dlg.el.querySelector<HTMLSelectElement>("#xl-sort")!.value as SortKey,
          highlight: dlg.el.querySelector<HTMLInputElement>("#xl-hl")!.checked,
        };
        try {
          localStorage.setItem("excel:opts", JSON.stringify(opts));
        } catch {}
        if (dlg.el.querySelector<HTMLInputElement>("#xl-roster")?.checked) opts.roster = rosterResult();
        exportExcel(opts);
      },
    },
  ]);
}

// ---------------------------------------------------------------- 检查更新

let appVersion = "";

async function runUpdateCheck(manual: boolean) {
  try {
    const info = await checkForUpdate(appVersion);
    if (info) {
      toast(`发现新版本 v${info.version}（当前 v${appVersion}）${info.notes ? `\n${info.notes}` : ""}`, "ok", [
        { label: "去下载", primary: true, run: () => openUrl(info.url) },
      ]);
    } else if (manual) {
      toast(`已经是最新版本（v${appVersion}）。`);
    }
  } catch (e) {
    // 自动检查失败（比如连不上 GitHub）不打扰老师
    if (manual) toast(`检查更新失败：${errText(e)}\n可能是网络连不上 GitHub，可以稍后再试，或直接访问\n${RELEASES_URL}`, "warn");
  }
}

$("#check-update").onclick = () => runUpdateCheck(true);

// ---------------------------------------------------------------- 启动

applyTheme(document.documentElement.dataset.theme === "dark" ? "dark" : "light");
loadOpts();
state.roster = loadRoster();
setBusy(false);
updateHint();
getVersion().then((v) => {
  appVersion = v;
  ui.versionLabel.textContent = `当前版本 v${v}`;
  if (state.opts.autoUpdate) runUpdateCheck(false);
});

// 开发调试用：自动化测试无法操作系统的文件对话框，直接传入文件夹（正式版中会被移除）
if (import.meta.env.DEV) {
  Object.assign(window, { __start: start, __setView: setView, __dev: { state, showRoster, showRename, showExportDialog, runUpdateCheck } });
}
