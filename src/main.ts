import "./styles.css";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";
import { buildExcel, exportCoverImages, exportCoverPdf, type PaperRecord } from "./exporters";
import { extract, Method, needsCheck } from "./extractor";
import { basename, collectPdfs, dirname, isDir, joinPath, openPath, readFile, revealPath, writeFile } from "./native";
import { ocrPage } from "./ocr";
import { loadPdf, renderCover } from "./pdf";

type Paper = PaperRecord & { id: number; state: "pending" | "done" | "error" };
type Filter = "all" | "check" | "done";
type View = "table" | "gallery";
type Opt = "recursive" | "ocr";
type EditKey = "title" | "name" | "studentId" | "advisor";
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
  opts: { recursive: false, ocr: true } as Record<Opt, boolean>,
};

/** 表格的列：第 0 列是行号 */
const COLS: { key: "no" | "status" | EditKey | "method" | "file"; letter: string; editable?: boolean }[] = [
  { key: "no", letter: "" },
  { key: "status", letter: "A" },
  { key: "title", letter: "B", editable: true },
  { key: "name", letter: "C", editable: true },
  { key: "studentId", letter: "D", editable: true },
  { key: "advisor", letter: "E", editable: true },
  { key: "method", letter: "F" },
  { key: "file", letter: "G" },
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
  for (const b of $$<HTMLButtonElement>("#open-folder, #pick, #export-excel, #export-pdf, #export-png")) b.disabled = busy;
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
    kv("学生", [p.name, p.studentId].filter(Boolean).join(" · ")) + kv("导师", p.advisor) + kv("文件", `${p.file}${p.pages ? ` · ${p.pages} 页` : ""}`);

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
  if (confirming) p.method = Method.Confirmed;

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
    toast("全部核对完了，可以导出了。", "ok", [{ label: "导出 Excel", run: exportExcel, primary: true }]);
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
  let files: string[];
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
  state.baseFolder = sources.length === 1 && (await isDir(sources[0])) ? sources[0] : dirname(files[0]);
  state.papers = files.map((path, id) => ({
    id, path, file: relative(path, state.baseFolder), state: "pending",
    title: "", method: "", name: "", studentId: "", advisor: "", pages: 0,
  }));
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
      updatePaper(p);
      setStatus(`正在识别 ${++done} / ${files.length}`, done, files.length);
    }
  };
  await Promise.all([worker(), worker()]);
  setBusy(false);
  setStatus("就绪");
  showDetail();

  const check = state.papers.filter(isCheck);
  if (!check.length) {
    toast(`全部 ${files.length} 篇都已自动识别。`, "ok", [{ label: "导出 Excel", run: exportExcel, primary: true }]);
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

async function exportExcel() {
  const target = await save({
    title: "保存 Excel",
    defaultPath: joinPath(state.baseFolder, "论文题目汇总.xlsx"),
    filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
  });
  if (!target || isPaperFile(target)) return;
  try {
    await writeFile(target, await buildExcel(state.papers));
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

$("#export-excel").onclick = exportExcel;
$("#export-pdf").onclick = () => exportCovers("pdf");
$("#export-png").onclick = () => exportCovers("png");
$("#copy-titles").onclick = copyTitles;

// ---------------------------------------------------------------- 启动

applyTheme(document.documentElement.dataset.theme === "dark" ? "dark" : "light");
loadOpts();
setBusy(false);
updateHint();

// 开发调试用：自动化测试无法操作系统的文件对话框，直接传入文件夹（正式版中会被移除）
if (import.meta.env.DEV) Object.assign(window, { __start: start, __setView: setView });
