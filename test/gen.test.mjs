import { test } from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const D = require("docx");
const JSZip = require("jszip");
const G = require("../src/gen.js");

const st = (o) => Object.assign({ zh: "標楷體", en: "Times New Roman", size: 13, bold: false, align: "left", lineType: "multiple", lineValue: 1, before: 0, after: 0, indent: 0 }, o);
const S = {
  styles: { h1: st({ size: 18, bold: true, align: "center" }), h2: st({ size: 16, bold: true }), h3: st({ size: 14, bold: true }),
    body: st({ align: "justify", lineType: "exact", lineValue: 25, after: 6, indent: 2 }), caption: st({ size: 12, bold: true, align: "center" }), table: st({ size: 11 }) },
  page: { top: 2.54, bottom: 2.54, left: 3.17, right: 3.17, header: true, pageNum: true },
};
// 2×1 透明 PNG
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAC0lEQVR42mNgQAcAAAkAAX+pdP4AAAAASUVORK5CYII=", "base64");
const P = {
  logos: { agency: { data: PNG, w: 2, h: 1 }, bidder: { data: PNG, w: 2, h: 1 } },
  projectName: "測試計畫", agency: "測試機關署", agencyShort: "貴署", bidder: "測試顧問公司", budget: 11900000, rocYear: 115, rocMonth: 1,
  period: "自決標次日起至民國115年12月10日止",
  workItems: [{ title: "調查作業", req: "辦理調查", subs: [{ title: "現地調查", req: "辦理30處" }, { title: "資料建置", req: "建置資料庫" }] }, { title: "查詢回復", req: "派駐人員", subs: [] }],
  milestones: [{ stage: "期中報告", when: "7月31日", tasks: ["完成15處"] }],
  deliverables: [{ name: "成果報告書", desc: "15份" }],
  evaluation: [{ item: "專案執行規劃", weight: 40, chapters: "第貳章～第肆章" }],
  priceItems: [{ title: "調查作業", qty: 1, unit: "式" }],
  references: [{ category: "法規與規定", text: "災害防救法。" }],
  months: Array.from({ length: 12 }, (_, i) => ({ y: 115, m: i + 1 })),
  content: { work: [{ subs: [{ blocks: [{ t: "p", text: "說明" }, { t: "figure", caption: "流程", steps: ["甲", "乙", "丙"] }, { t: "table", caption: "表", headers: ["a", "b"], rows: [["1", "2"]] }] }] }] },
};

test("金額大寫", () => {
  assert.equal(G.upperAmount(11900000), "壹仟壹佰玖拾萬元整");
  assert.equal(G.upperAmount(20304), "貳萬零參佰零肆元整");
  assert.equal(G.upperAmount(10000001), "壹仟萬零壹元整");
});

test("產生完整建議書", async () => {
  const buf = await D.Packer.toBuffer(G.buildProposal(D, P, S));
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file("word/document.xml").async("string");
  for (const s of ["服 務 建 議 書", "目　錄", "表 目 錄", "圖 目 錄", "第壹章　前言", "2-3　調查作業", "2-3-1　現地調查", "2-4　查詢回復",
    "甘特圖", "投標標價清單", "壹仟壹佰玖拾萬元整", "服務費用分析表", "參考文獻", "災害防救法", "附錄一", "【待補", "TOC \\h"]) assert.ok(xml.includes(s), `缺少：${s}`);
  assert.match(xml, /TOC[^<]*\\t[^<]*圖標題/);
  assert.match(xml, /TOC[^<]*\\t[^<]*表標題/);
  // 每章一節：頁首右側為該章名稱，頁尾為投標廠商
  const headers = await Promise.all(Object.keys(zip.files).filter((n) => /^word\/header\d+\.xml$/.test(n)).map((n) => zip.file(n).async("string")));
  const footers = await Promise.all(Object.keys(zip.files).filter((n) => /^word\/footer\d+\.xml$/.test(n)).map((n) => zip.file(n).async("string")));
  for (const t of ["目錄", "表目錄", "圖目錄", "第壹章　前言", "第捌章　標價清單及費用分析", "參考文獻", "附錄一　相關工作實績證明文件"])
    assert.ok(headers.some((h) => h.includes(`>${t}<`)), `頁首缺少章名：${t}`);
  assert.ok(headers.every((h) => h.includes("測試計畫")), "頁首缺少標案名稱");
  assert.ok(footers.length && footers.every((f) => f.includes("測試顧問公司") && f.includes("PAGE")), "頁尾缺少廠商名稱或頁碼");
  assert.ok(Object.keys(zip.files).some((n) => /^word\/media\//.test(n)), "未嵌入 LOGO 圖檔");
  assert.ok(headers.every((h) => h.includes("<w:drawing>")), "頁首缺少機關 LOGO");
  const styles = await zip.file("word/styles.xml").async("string");
  assert.ok(styles.includes('w:eastAsia="標楷體"'));
});
