import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { buildExcel, DEFAULT_EXCEL_OPTIONS, sortRecords, type PaperRecord } from "../src/exporters";
import { planRenames, sanitize, type RenamePaper } from "../src/rename";
import { matchRoster, parseRoster, type RosterPaper } from "../src/roster";
import { isNewer } from "../src/update";

const enc = (s: string) => new TextEncoder().encode(s);

describe("名单", () => {
  it("CSV：自动找到学号、姓名列，跳过空行和重复", async () => {
    const csv = "班级,学号,姓名\n经济1班,2210000001,王五\n,,\n经济1班,2210000002,李 雷\n经济1班,2210000001,王五\n";
    expect(await parseRoster(enc(csv), "名单.csv")).toEqual([
      { id: "2210000001", name: "王五" },
      { id: "2210000002", name: "李雷" },
    ]);
  });

  it("GBK 编码的 CSV（中文 Excel 另存的默认格式）", async () => {
    // 「学号,姓名\n2210000001,王五」的 GBK 字节
    const gbk = new Uint8Array([
      0xd1, 0xa7, 0xba, 0xc5, 0x2c, 0xd0, 0xd5, 0xc3, 0xfb, 0x0a,
      0x32, 0x32, 0x31, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x31, 0x2c, 0xcd, 0xf5, 0xce, 0xe5,
    ]);
    expect(await parseRoster(gbk, "名单.csv")).toEqual([{ id: "2210000001", name: "王五" }]);
  });

  it("没有表头时按内容猜列", async () => {
    expect(await parseRoster(enc("王五,2210000001\n李雷,2210000002"), "a.csv")).toEqual([
      { id: "2210000001", name: "王五" },
      { id: "2210000002", name: "李雷" },
    ]);
  });

  it("xlsx", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("名单");
    ws.addRow(["序号", "学号", "姓名"]);
    ws.addRow([1, "2210000001", "王五"]);
    ws.addRow([2, 2210000002, "李雷"]); // 学号存成数字也要能读
    const data = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    expect(await parseRoster(data, "名单.xlsx")).toEqual([
      { id: "2210000001", name: "王五" },
      { id: "2210000002", name: "李雷" },
    ]);
  });

  it("比对：已交、未交、不一致、不在名单、文件名兜底", () => {
    const paper = (id: number, studentId: string, name: string, file = `${id}.pdf`): RosterPaper => ({
      id, studentId, name, file, title: `题目${id}`,
    });
    const r = matchRoster(
      [
        { id: "1001", name: "王五" },
        { id: "1002", name: "李雷" },
        { id: "1003", name: "韩梅梅" },
        { id: "1004", name: "赵六" },
      ],
      [
        paper(0, "1001", "王五"),
        paper(1, "1002", "李磊"), // 姓名写错
        paper(2, "", "", "1004+赵六+论文.pdf"), // 封面没读出来，只能看文件名
        paper(3, "9999", "路人"),
      ],
    );
    expect(r.missing.map((s) => s.name)).toEqual(["韩梅梅"]);
    expect(r.submitted.map((s) => s.student.name)).toEqual(["王五", "李雷", "赵六"]);
    expect(r.mismatched).toHaveLength(1);
    expect(r.mismatched[0].reason).toContain("李磊");
    expect(r.extra.map((p) => p.name)).toEqual(["路人"]);
  });
});

describe("批量重命名", () => {
  const p = (id: number, over: Partial<RenamePaper> = {}): RenamePaper => ({
    id, path: `C:\\论文\\${id}.pdf`, title: `题目${id}`, name: `学生${id}`, studentId: `100${id}`,
    advisor: "", college: "", major: "", ...over,
  });

  it("去掉非法字符、合并空项留下的分隔符", () => {
    expect(sanitize('A/B:C*D?"E"<F>|G')).toBe("A B C D E F G");
    expect(sanitize("_1001__题目_")).toBe("1001_题目");
    expect(sanitize("1001_研究——基于面板数据")).toBe("1001_研究——基于面板数据"); // 题目里的破折号保留
  });

  it("生成计划：正常、跳过、缺信息、冲突、已符合", () => {
    const plans = planRenames(
      [
        p(1),
        p(2, { title: "" }), // 缺题目
        p(3, { name: "同名", studentId: "", title: "同题" }),
        p(4, { name: "同名", studentId: "", title: "同题" }), // 和 3 改成同一个名字
        p(5),
        p(6, { path: "C:\\论文\\1006_学生6_题目6.pdf" }), // 已经符合
      ],
      "{学号}_{姓名}_{题目}",
      (x) => x.id === 5,
    );
    expect(plans.map((x) => x.status)).toEqual(["ok", "missing", "missing", "missing", "skip", "same"]);
    expect(plans[0].to).toBe("C:\\论文\\1001_学生1_题目1.pdf");
    expect(plans[1].note).toContain("题目");

    const dup = planRenames([p(3, { studentId: "" }), p(4, { studentId: "", name: "学生3", title: "题目3" })], "{姓名}_{题目}", () => false);
    expect(dup.map((x) => x.status)).toEqual(["conflict", "conflict"]);
  });

  it("目标名字被一篇不改名的论文占用时标为冲突", () => {
    const plans = planRenames(
      [p(1, { name: "甲" }), p(2, { path: "C:\\论文\\甲.pdf" })],
      "{姓名}",
      (x) => x.id === 2,
    );
    expect(plans[0].status).toBe("conflict");
  });
});

describe("Excel 导出选项", () => {
  const rec = (title: string, studentId: string, name = ""): PaperRecord => ({
    path: "", file: `${title}.pdf`, title, method: "封面题目栏", name, studentId,
    advisor: "", advisorTitle: "", college: "", major: "", pages: 1,
  });

  it("按学号排序，空的排最后", () => {
    const sorted = sortRecords([rec("c", ""), rec("b", "1002"), rec("a", "1001")], "studentId");
    expect(sorted.map((r) => r.title)).toEqual(["a", "b", "c"]);
  });

  it("只导出勾选的列", async () => {
    const data = await buildExcel([rec("题目一", "1001", "王五")], { ...DEFAULT_EXCEL_OPTIONS, columns: ["no", "title", "studentId"] });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data.buffer as ArrayBuffer);
    const ws = wb.worksheets[0];
    expect(ws.getRow(1).values).toEqual([undefined, "序号", "论文题目", "学号"]);
    expect(ws.getRow(2).values).toEqual([undefined, 1, "题目一", "1001"]);
  });
});

describe("Excel 名单核对表", () => {
  it("附带名单时多一张「名单核对」表，未交排在最前", async () => {
    const r = matchRoster(
      [{ id: "1001", name: "王五" }, { id: "1002", name: "李雷" }],
      [{ id: 0, studentId: "1001", name: "王五", title: "题目一", file: "a.pdf" }],
    );
    const rec: PaperRecord = {
      path: "", file: "a.pdf", title: "题目一", method: "封面题目栏", name: "王五", studentId: "1001",
      advisor: "", advisorTitle: "", college: "", major: "", pages: 1,
    };
    const data = await buildExcel([rec], { ...DEFAULT_EXCEL_OPTIONS, roster: r });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data.buffer as ArrayBuffer);
    const ws = wb.getWorksheet("名单核对")!;
    expect(ws.getRow(2).values).toEqual([undefined, "未交", "1002", "李雷"]);
    expect((ws.getRow(3).values as unknown[]).slice(1, 5)).toEqual(["已交", "1001", "王五", "题目一"]);
  });
});

describe("检查更新", () => {
  it("版本比较", () => {
    expect(isNewer("v1.10.0", "1.9.2")).toBe(true);
    expect(isNewer("v1.1.0", "1.1.0")).toBe(false);
    expect(isNewer("1.0.9", "1.1.0")).toBe(false);
  });
});
