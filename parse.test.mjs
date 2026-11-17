import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const TP = require("../src/parse.js");

test("中文數字", () => {
  assert.equal(TP.cnNum("一百一十五"), 115);
  assert.equal(TP.cnNum("一一六"), 116);
  assert.equal(TP.cnNum("十二"), 12);
  assert.equal(TP.cnNum("三十"), 30);
  assert.equal(TP.cnNum("10"), 10);
});

test("子項短標題", () => {
  const t = TP.shortTitle("廠商須依據本署最新之劃設作業流程與方法，辦理150條(處)之現地勘查及相關評估報告(以實做數量計價，至少完成140條)");
  assert.ok(t.length <= 30, t);
  assert.match(t, /現地勘查/);
  assert.equal(TP.shortTitle("教育訓練"), "教育訓練");
});

test("文件分類", () => {
  assert.equal(TP.classify("1-05-評選須知.odt", ""), "eval");
  assert.equal(TP.classify("1-04-投標須知.odt", ""), "notice");
  assert.equal(TP.classify("1-08-標價清單.odt", ""), "price");
  assert.equal(TP.classify("契約範本.odt", ""), "contract");
  assert.equal(TP.classify("1-02-採購說明.odt", ""), "spec");
});

test("純文字（PDF 路徑）解析", () => {
  const parsed = TP.extractPlain(readFileSync(new URL("./fixtures/spec.txt", import.meta.url), "utf8"));
  const P = TP.analyze([{ name: "採購說明.pdf", kind: "spec", parsed }]);
  assert.equal(P.projectName, "116年示範區防災調查委託專業服務");
  assert.equal(P.agency, "某某市政府");
  assert.equal(P.agencyShort, "貴府");
  assert.equal(P.budget, 3000000);
  assert.deepEqual(P.endRoc, { y: 116, m: 11, d: 30 });
  assert.equal(P.workItems.length, 2);
  assert.equal(P.workItems[0].title, "現況調查作業");
  assert.equal(P.workItems[0].subs.length, 2);
  assert.equal(P.workItems[1].subs.length, 2);
  assert.deepEqual(P.milestones.map((m) => m.stage).slice(0, 2), ["期中報告", "期末報告"]);
  assert.equal(P.deliverables[0].name, "成果報告書");
  assert.ok(P.background.includes("掌握轄內坡地災害潛勢"));
  assert.ok(P.references.some((r) => r.text.includes("坡地災害潛勢調查作業要點")));
});
