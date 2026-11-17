/* 招標文件解析：extractODT / extractDOCX / extractPlain → {lines, tables}; analyze(docs) → project */
(function (root) {
  "use strict";
  const NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
  const NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
  const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const clean = (s) => String(s || "").replace(/[\u00a0\u3000\t ]+/g, " ").replace(/\s+/g, " ").trim();

  // ---------- 編號前綴層級 ----------
  const PREFIX = [
    [/^[壹貳參肆伍陸柒捌玖拾]+\s*[、.．]\s*/, 1],
    [/^[一二三四五六七八九十]+\s*[、.．]\s*/, 2],
    [/^[(（]\s*[一二三四五六七八九十]+\s*[)）]\s*[、.．]?\s*/, 3],
    [/^\d{1,2}\s*[.、．](?!\d)\s*/, 4],
    [/^[(（]\s*\d{1,2}\s*[)）]\s*/, 5],
    [/^[A-Za-z]\s*[.、]\s*/, 6],
  ];
  function prefixLvl(t) { for (const [re, l] of PREFIX) if (re.test(t)) return l; return 0; }
  function stripPrefix(t) { for (const [re] of PREFIX) if (re.test(t)) return t.replace(re, ""); return t; }

  // ---------- ODT ----------
  function odtText(node) {
    let s = "";
    for (const c of Array.from(node.childNodes || [])) {
      if (c.nodeType === 3) s += c.nodeValue;
      else if (c.nodeType === 1) {
        const ln = c.localName;
        if (ln === "s") s += " ".repeat(parseInt(c.getAttributeNS(NS_TEXT, "c") || "1", 10));
        else if (ln === "tab") s += " ";
        else if (ln === "line-break") s += "\u2029";
        else if (ln === "annotation" || ln === "note" || ln === "tracked-changes") continue;
        else s += odtText(c);
      }
    }
    return s;
  }
  function extractODT(xml, DP) {
    const doc = new DP().parseFromString(xml, "application/xml");
    const lines = [], tables = [];
    const bodyEl = doc.getElementsByTagNameNS("urn:oasis:names:tc:opendocument:xmlns:office:1.0", "text")[0];
    function tableRows(tbl) {
      const rows = [];
      const walkRows = (n) => {
        for (const c of Array.from(n.childNodes || [])) {
          if (c.nodeType !== 1) continue;
          if (c.localName === "table-row") {
            const cells = [];
            for (const cc of Array.from(c.childNodes || [])) if (cc.nodeType === 1 && cc.localName === "table-cell") {
              const ps = []; const collect = (x) => { for (const y of Array.from(x.childNodes || [])) { if (y.nodeType !== 1) continue; if (y.localName === "p" || y.localName === "h") ps.push(clean(odtText(y))); else collect(y); } };
              collect(cc); cells.push(clean(ps.filter(Boolean).join(" ").replace(/\u2029/g, " ")));
            }
            rows.push(cells);
          } else if (/table-(header-)?rows|table-row-group/.test(c.localName)) walkRows(c);
        }
      };
      walkRows(tbl); return rows;
    }
    function walk(node, depth) {
      for (const c of Array.from(node.childNodes || [])) {
        if (c.nodeType !== 1) continue;
        const ln = c.localName;
        if (ln === "h") { const t = clean(odtText(c).replace(/\u2029/g, " ")); if (t) lines.push({ text: t, kind: "h", depth: parseInt(c.getAttributeNS(NS_TEXT, "outline-level") || "1", 10) }); }
        else if (ln === "p") { for (const part of odtText(c).split("\u2029")) { const t = clean(part); if (t) lines.push({ text: t, kind: "p", depth: 0 }); } }
        else if (ln === "list") {
          for (const li of Array.from(c.childNodes || [])) {
            if (li.nodeType !== 1 || (li.localName !== "list-item" && li.localName !== "list-header")) continue;
            let first = li.localName === "list-item";
            for (const x of Array.from(li.childNodes || [])) {
              if (x.nodeType !== 1) continue;
              if (x.localName === "p" || x.localName === "h") {
                for (const part of odtText(x).split("\u2029")) {
                  const t = clean(part); if (!t) continue;
                  lines.push({ text: t, kind: first ? "li" : "cont", depth: depth + 1 }); first = false;
                }
              } else if (x.localName === "list") walk({ childNodes: [x] }, depth + 1);
              else if (x.localName === "table") { tables.push(tableRows(x)); lines.push({ text: "", kind: "table", depth: 0 }); }
            }
          }
        } else if (ln === "table") { tables.push(tableRows(c)); lines.push({ text: "", kind: "table", depth: 0 }); }
        else if (ln === "section" || ln === "index-body" || ln === "table-of-content" || ln === "soft-page-break") walk(c, depth);
      }
    }
    if (bodyEl) walk(bodyEl, 0);
    return finish(lines, tables);
  }

  // ---------- DOCX ----------
  function extractDOCX(xml, DP) {
    const doc = new DP().parseFromString(xml, "application/xml");
    const lines = [], tables = [];
    const bodyEl = doc.getElementsByTagNameNS(NS_W, "body")[0];
    const ptext = (p) => { let s = ""; const ts = p.getElementsByTagNameNS(NS_W, "t"); for (let i = 0; i < ts.length; i++) s += ts[i].textContent; return clean(s); };
    const kids = (n, name) => Array.from(n.childNodes || []).filter((c) => c.nodeType === 1 && c.localName === name);
    function para(p) {
      const t = ptext(p); if (!t) return;
      const pPr = kids(p, "pPr")[0];
      let kind = "p", depth = 0;
      if (pPr) {
        const ps = kids(pPr, "pStyle")[0]; const sv = ps ? ps.getAttributeNS(NS_W, "val") || ps.getAttribute("w:val") : "";
        const hm = /^(Heading|標題)\s*(\d)$/i.exec(sv || "");
        if (hm) { kind = "h"; depth = +hm[2]; }
        const np = kids(pPr, "numPr")[0];
        if (np && kind === "p") { const il = kids(np, "ilvl")[0]; kind = "li"; depth = 1 + parseInt((il && (il.getAttributeNS(NS_W, "val") || il.getAttribute("w:val"))) || "0", 10); }
      }
      lines.push({ text: t, kind, depth });
    }
    function tbl(t) {
      const rows = kids(t, "tr").map((tr) => kids(tr, "tc").map((tc) => clean(kids(tc, "p").map(ptext).filter(Boolean).join(" "))));
      tables.push(rows); lines.push({ text: "", kind: "table", depth: 0 });
    }
    if (bodyEl) for (const c of Array.from(bodyEl.childNodes)) { if (c.nodeType !== 1) continue; if (c.localName === "p") para(c); else if (c.localName === "tbl") tbl(c); else if (c.localName === "sdt") { const pc = c.getElementsByTagNameNS(NS_W, "p"); for (let i = 0; i < pc.length; i++) para(pc[i]); } }
    return finish(lines, tables);
  }

  // ---------- 純文字（PDF/TXT） ----------
  function extractPlain(text) {
    const lines = String(text || "").split(/\r?\n/).map(clean).filter(Boolean).map((t) => ({ text: t, kind: "p", depth: 0 }));
    return finish(lines, []);
  }

  function finish(lines, tables) {
    // 統一層級：清單用 depth；一般段落以編號前綴推估；無編號者視為內文(99)
    const hasLists = lines.some((l) => l.kind === "li");
    lines.forEach((l) => {
      const pl = prefixLvl(l.text);
      if (l.kind === "li" || l.kind === "cont") { l.lvl = l.depth * 10; l.struct = l.kind === "li"; }
      else if (l.kind === "h") { l.lvl = l.depth * 10 - 5; l.struct = true; }
      else if (pl) { l.lvl = hasLists ? 10 + pl * 2 : pl * 10; l.struct = true; }
      else { l.lvl = 999; l.struct = false; }
      l.title = stripPrefix(l.text);
    });
    const text = lines.map((l) => l.text).join("\n") + "\n" + tables.map((t) => t.map((r) => r.join(" | ")).join("\n")).join("\n");
    return { lines, tables, text };
  }

  // ---------- 文件分類 ----------
  function classify(name, text) {
    const n = name || "", t = (text || "").slice(0, 1500);
    if (/評選須知|評選辦法|評審須知/.test(n + t.slice(0, 300))) return "eval";
    if (/投標須知/.test(n + t.slice(0, 200))) return "notice";
    if (/標價清單|價格清單|報價單/.test(n + t.slice(0, 200))) return "price";
    if (/契約/.test(n) || /契約(書|範本|草案)/.test(t.slice(0, 200))) return "contract";
    if (/招標公告|公告/.test(n)) return "announce";
    if (/採購說明|需求說明|工作說明|招標規範|規範|需求書|計畫書/.test(n + t.slice(0, 300))) return "spec";
    return "other";
  }
  const KIND_LABEL = { spec: "招標規範／採購說明", notice: "投標須知", eval: "評選須知", contract: "契約草案", price: "標價清單", announce: "招標公告", other: "其他文件" };

  // ---------- 中文數字 ----------
  function cnNum(s) {
    s = String(s).trim(); if (/^\d+$/.test(s)) return +s;
    const d = { 零: 0, 〇: 0, 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
    // 逐位寫法（例：一一六 → 116）
    if (!/[十百千]/.test(s)) { let v = 0; for (const ch of s) if (ch in d) v = v * 10 + d[ch]; return v; }
    let total = 0, cur = 0;
    for (const ch of s) {
      if (ch in d) cur = d[ch];
      else if (ch === "十") { total += (cur || 1) * 10; cur = 0; }
      else if (ch === "百") { total += (cur || 1) * 100; cur = 0; }
      else if (ch === "千") { total += (cur || 1) * 1000; cur = 0; }
    }
    return total + cur;
  }

  // ---------- 結構工具 ----------
  function children(lines, i) {
    const res = [], base = lines[i].lvl;
    for (let j = i + 1; j < lines.length; j++) {
      const L = lines[j];
      if (L.kind === "table") continue;
      if (L.struct && L.lvl <= base) break;
      if (!L.struct && base >= 999) break;
      res.push(j);
    }
    return res;
  }
  function findHead(lines, re, maxLen = 26) {
    for (let i = 0; i < lines.length; i++) { const L = lines[i]; if (L.struct && L.title.length <= maxLen && re.test(L.title)) return i; }
    return -1;
  }
  function shortTitle(t) {
    t = stripPrefix(clean(t)).replace(/[（(][^）)]*[）)]/g, "").replace(/[：:。；;，,]+$/, "").trim();
    if (t.length <= 26) return t;
    const clauses = t.split(/[，,。；;：:]/).map((x) => x.trim()).filter(Boolean);
    const V = "(前往|辦理|進行|協助|完成|建置|提出|檢視|製作|彙整|維護|更新|研擬|評估|輔導|說明|派駐|分析|調查|規劃|精進|盤點|印製|界定|蒐集|建立|執行|召開|編撰|訂定)";
    const startV = new RegExp("^(協助)?" + V);
    let c = clauses.find((x) => startV.test(x) && x.length >= 6);
    if (!c) {
      c = clauses[0] || t;
      c = c.replace(/^(受託)?廠商(須|應|需)?/, "").replace(/^(須|應|需|並|另|且|為)/, "");
      const mv = new RegExp("^(依據|依|參考|配合|就)[^，,]*?" + V).exec(c);
      if (mv) c = c.slice(mv.index + mv[0].length - mv[2].length);
    }
    c = c.trim() || t;
    if (c.length > 30) { const cut = c.slice(0, 28); const b = Math.max(cut.lastIndexOf("之"), cut.lastIndexOf("及"), cut.lastIndexOf("、")); c = (b >= 12 ? cut.slice(0, b) : cut) ; }
    return c;
  }

  // ---------- 主分析 ----------
  function analyze(docs) {
    const by = (k) => docs.filter((d) => d.kind === k);
    const all = docs.map((d) => d.parsed.text).join("\n");
    const spec = by("spec")[0] || by("announce")[0] || docs.find((d) => d.kind !== "contract") || docs[0];
    const P = { sources: docs.map((d) => ({ name: d.name, kind: d.kind })) };

    // 計畫名稱
    let m;
    if ((m = /本?標案名稱[：:]\s*([^\n。]+)/.exec(all))) P.projectName = clean(m[1]);
    else if ((m = /委託計畫名稱[：:]\s*([^\n]+)/.exec(all))) P.projectName = clean(m[1]);
    else if (spec && (m = /「([^」]{4,60})」/.exec(spec.parsed.text.slice(0, 600)))) P.projectName = clean(m[1]);
    else if ((m = /採購案[：:]\s*([^\s　]+)/.exec(all))) P.projectName = clean(m[1]);
    if (P.projectName) P.projectName = P.projectName.replace(/[「」]/g, "");

    // 機關
    const agRe = /^(.{2,30}?(?:署|局|部|府|處|會|所|中心|公所|院|館|隊|廳|室))$/;
    for (const d of docs) { const f = d.parsed.lines.slice(0, 6).find((l) => agRe.test(l.text.replace(/^#+\s*/, ""))); if (f) { P.agency = f.text.replace(/^#+\s*/, ""); break; } }
    if (!P.agency && (m = /機關名稱[：:]\s*([^\s，。]+)/.exec(all))) P.agency = m[1];
    if (P.agency) P.agencyShort = "貴" + P.agency.slice(-1).replace("所", "所").replace("心", "中心");

    // 預算
    if ((m = /NT\$\s*([\d,]{5,})/.exec(all))) P.budget = +m[1].replace(/,/g, "");
    else if ((m = /預算(?:金額|經費)?[^\d\n]{0,6}([\d,.]+)\s*萬元/.exec(all))) P.budget = Math.round(parseFloat(m[1].replace(/,/g, "")) * 10000);
    else if ((m = /預算(?:金額|經費)?[^\d\n]{0,6}([\d,]{5,})\s*元/.exec(all))) P.budget = +m[1].replace(/,/g, "");

    // 履約期限
    const numC = "[一二三四五六七八九十百零〇\\d]+";
    const pRe = new RegExp("(自[^。\\n]{0,20}?(?:至|起至))\\s*(?:民國)?\\s*(" + numC + ")\\s*年\\s*(" + numC + ")\\s*月\\s*(" + numC + ")\\s*日");
    if ((m = pRe.exec(all))) {
      P.endRoc = { y: cnNum(m[2]), m: cnNum(m[3]), d: cnNum(m[4]) };
      P.period = `${m[1].replace(/起?至$/, "")}起至民國${P.endRoc.y}年${P.endRoc.m}月${P.endRoc.d}日止`;
    } else if ((m = /(履約期限|履約期間)[：:為]?\s*([^。\n]{4,60})/.exec(all))) P.period = clean(m[2]);

    // 計畫需求／背景
    if (spec) {
      const L = spec.parsed.lines;
      const h = findHead(L, /(計畫需求|計畫緣起|計畫目的|需求說明|計畫背景|緣起)/);
      if (h >= 0) P.background = [L[h]].concat(children(L, h).map((j) => L[j])).filter((l) => l.kind === "cont" || !l.struct || l !== L[h]).map((l) => l.title).filter((t) => t !== L[h].title).join("\n");
    }

    // 工作項目
    P.workItems = [];
    if (spec) {
      const L = spec.parsed.lines;
      const h = findHead(L, /(計畫項目|工作項目|工作內容|服務項目|委託項目|工作範圍|服務內容|辦理事項)/);
      if (h >= 0) {
        const kids = children(L, h).filter((j) => L[j].struct);
        const minL = Math.min(...kids.map((j) => L[j].lvl));
        kids.forEach((j, idx) => {
          if (L[j].lvl !== minL) return;
          const sub = children(L, j);
          const item = { title: shortTitle(L[j].title), req: L[j].title, subs: [] };
          const conts = sub.filter((k) => !L[k].struct && L[k].lvl >= L[j].lvl).map((k) => L[k].title);
          if (conts.length && !sub.some((k) => L[k].struct)) item.req += "\n" + conts.join("\n");
          const sk = sub.filter((k) => L[k].struct);
          if (sk.length) {
            const sl = Math.min(...sk.map((k) => L[k].lvl));
            sk.forEach((k) => {
              if (L[k].lvl !== sl) return;
              const desc = children(L, k).map((q) => L[q].title);
              const req = [L[k].title].concat(desc).join("\n");
              item.subs.push({ title: shortTitle(L[k].title), req: req.length > 1500 ? req.slice(0, 1500) + "…" : req });
            });
          } else if (conts.length) {
            // 無子項時，以續行段落作為需求說明
            item.req = [L[j].title].concat(conts).join("\n");
          }
          P.workItems.push(item);
        });
      }
    }
    if (!P.workItems.length) {
      const pr = by("price")[0];
      if (pr) pr.parsed.tables.forEach((t) => t.forEach((r) => { const c = r.find((x, i) => i > 0 && x && x.length >= 4 && !/服務項目|合計|總標價|項次/.test(x)); if (c && r.length >= 4) P.workItems.push({ title: c, req: c, subs: [] }); }));
    }

    // 標價清單項目
    P.priceItems = [];
    const pr = by("price")[0];
    if (pr) pr.parsed.tables.forEach((t) => {
      const hi = t.findIndex((r) => r.some((c) => /服務項目|工作項目|項目/.test(c)));
      if (hi < 0) return;
      const ci = t[hi].findIndex((c) => /服務項目|工作項目|項目/.test(c) && !/項次/.test(c));
      const qi = t[hi].findIndex((c) => /數量/.test(c)), ui = t[hi].findIndex((c) => /單位/.test(c));
      t.slice(hi + 1).forEach((r) => { const x = r[ci]; if (x && !/合計|總計|總標價/.test(r.join(""))) P.priceItems.push({ title: x, qty: (qi >= 0 && +r[qi]) || 1, unit: (ui >= 0 && r[ui]) || "式" }); });
    });

    // 分段進度／查核點
    P.milestones = [];
    docs.filter((d) => d.kind === "spec" || d.kind === "announce").forEach((d) => {
      const L = d.parsed.lines;
      L.forEach((l, i) => {
        if (!l.struct || !/^(工作執行計畫書|期初|期中|期末|第[一二三四五]期)/.test(l.title) || l.title.length > 40) return;
        const tasks = children(L, i).filter((j) => L[j].struct).map((j) => L[j].title);
        if (!tasks.length) return;
        const mm = /^(.*?)(時需完成|時應完成|需完成|應完成|時需|完成工作)/.exec(l.title);
        P.milestones.push({ stage: clean((mm ? mm[1] : l.title).replace(/[：:]$/, "")), tasks, when: "" });
      });
    });
    if (/驗收|契約完成/.test(all) && P.milestones.length && P.endRoc) P.milestones.push({ stage: "成果報告書定稿", when: `${P.endRoc.y}年${P.endRoc.m}月${P.endRoc.d}日前`, tasks: ["依審查意見修正並提送成果報告書及電子檔", "辦理驗收"] });

    // 成果交付
    P.deliverables = [];
    if (spec) {
      const seen = new Set();
      spec.parsed.lines.forEach((l) => {
        const mm = /^(.{2,40}?)[：:]?\s*(\d+\s*份.*)$/.exec(l.title);
        if (!mm) return;
        let name = mm[1].replace(/^[(（][一二三四五六七八九十]+[)）]/, "").replace(/[：:]$/, "");
        if ((name.match(/[(（]/g) || []).length > (name.match(/[)）]/g) || []).length) name = name.replace(/[(（].*$/, "");
        if (!name || seen.has(name)) return;
        seen.add(name); P.deliverables.push({ name, desc: mm[2].length > 160 ? mm[2].slice(0, 160) + "…" : mm[2] });
      });
    }

    // 評選項目
    P.evaluation = [];
    for (const d of docs) {
      for (const t of d.parsed.tables) {
        const hi = t.findIndex((r) => r.some((c) => /配分/.test(c)) && r.some((c) => /評選項目|評分項目|評審項目/.test(c)));
        if (hi < 0) continue;
        t.slice(hi + 1).forEach((r) => {
          const item = r[0]; if (!item || /總分|合計|得分/.test(item) || /^\d+$/.test(item)) return;
          const w = r.slice(1).reverse().find((c) => /^\d{1,3}(\.\d+)?$/.test(c));
          if (!w || P.evaluation.find((e) => e.item === item)) return;
          P.evaluation.push({ item, weight: +w, desc: (r[1] || "").slice(0, 300) });
        });
        if (P.evaluation.length) break;
      }
      if (P.evaluation.length) break;
    }
    P.evaluation.forEach((e) => {
      const s = e.item;
      e.chapters = /CSR|社會責任/.test(s) ? "第柒章" : /價格|標價|經費|費用/.test(s) ? "第捌章" : /簡報|答詢/.test(s) ? "簡報及答詢" : /管理|履約|組織|人力/.test(s) ? "第伍章、第陸章" : /實績|經驗|規模/.test(s) ? "第陸章" : /規劃|技術|執行|方法|創新/.test(s) ? "第貳章～第肆章" : "【待補】";
    });

    // 內容要領
    P.outlineHints = [];
    const ev = by("eval")[0];
    if (ev) { const L = ev.parsed.lines; const h = findHead(L, /內容要領|內容要求|建議書內容/); if (h >= 0) P.outlineHints = children(L, h).filter((j) => L[j].struct).map((j) => L[j].title.slice(0, 200)).slice(0, 15); }

    // 法規與參考依據（僅取規範／評選須知，避免投標須知樣板文字）
    const refs = new Map();
    const srcDocs = docs.filter((d) => d.kind === "spec" || d.kind === "eval" || d.kind === "announce");
    const srcText = srcDocs.map((d) => d.parsed.text).join("\n");
    const add = (name, year, cat, issuer) => {
      name = clean(name).replace(/^(依據|依照|參照|參考|依|按|本署|本局|本府|本處|本機關)/, "");
      if (cat !== "研究報告") name = name.split(/[之為得將於的依照參]|如有|不得|過程|薪資|委員及|最新/).pop();
      if (name === P.projectName) return;
      if (name.length < 4 || name.length > 40 || /^(符合|於|其|有|條|及|與|並|之)/.test(name)) return;
      const r = refs.get(name) || { name, year: "", cat, issuer: "", n: 0 };
      r.n++; if (year) r.year = year; if (issuer) r.issuer = issuer; refs.set(name, r);
    };
    let lm;
    const q = /(本署|本局|本府|本處)?「([^」]{4,40})」(?:[（(](\d{2,3})年[)）])?/g;
    while ((lm = q.exec(srcText))) { const n = lm[2]; if (/^\d{2,3}年/.test(n)) add(n, "", "研究報告", P.agency); else if (/(法|辦法|準則|規則|要點|規定|手冊|規範|方法|原則|計畫)$/.test(n)) add(n, lm[3], "法規與規定", lm[1] ? P.agency : ""); }
    const law = /(?:依據|依照|參照|依|因應)?(本署|本局)?([\u4e00-\u9fff]{4,30}?(?:作業流程與方法|作業要點|作業規定|施行細則|辦法|準則|規則|要點|法))(?:(第\d+條)|[（(](\d{2,3})年[)）])/g;
    while ((lm = law.exec(srcText))) add(lm[2], lm[4], "法規與規定", lm[1] ? P.agency : "");
    const law2 = /(?:依據|依照|參照|依|因應)([\u4e00-\u9fff]{3,24}?(?:辦法|準則|規則|要點|法))(?=[之，,。、及並辦規定推])/g;
    while ((lm = law2.exec(srcText))) add(lm[1], "", "法規與規定", "");
    const roc = P.endRoc ? P.endRoc.y : "";
    const ad = (y) => (y ? `，${+y + 1911}` : "");
    P.references = Array.from(refs.values()).sort((a, b) => (a.cat === b.cat ? b.n - a.n : a.cat < b.cat ? -1 : 1)).slice(0, 16)
      .map((r) => ({ category: r.cat, text: r.cat === "研究報告" ? `${r.issuer || "【待補：機關】"}，${r.name}${ad(r.name.match(/^(\d{2,3})年/)[1])}。` : `${r.issuer ? r.issuer + "，" : ""}${r.name}${ad(r.year)}。` }));
    docs.forEach((d) => { if (d.kind !== "other") P.references.push({ category: "招標文件", text: `${P.agency || "招標機關"}，「${P.projectName || ""}」${KIND_LABEL[d.kind]}${roc ? `，${+roc + 1911}` : ""}。` }); });
    P.hasResident = /駐點/.test(all);
    return P;
  }

  root.TenderParse = { extractODT, extractDOCX, extractPlain, classify, analyze, KIND_LABEL, shortTitle, cnNum };
  if (typeof module !== "undefined") module.exports = root.TenderParse;
})(typeof window !== "undefined" ? window : globalThis);
