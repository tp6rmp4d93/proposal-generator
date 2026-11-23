/* 服務建議書 Word 產生器：buildProposal(docxNamespace, project, settings) -> Document */
(function (root) {
  "use strict";
  const CH = ["", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖", "拾", "拾壹", "拾貳"];
  const CN = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"];

  function upperAmount(n) {
    if (!n || isNaN(n)) return "";
    const d = "零壹貳參肆伍陸柒捌玖", u = ["", "拾", "佰", "仟"], big = ["", "萬", "億", "兆"];
    n = Math.round(n);
    const groups = []; while (n > 0) { groups.push(n % 10000); n = Math.floor(n / 10000); }
    let s = "", pendingZero = false;
    for (let g = groups.length - 1; g >= 0; g--) {
      const part = groups[g];
      if (!part) { if (s) pendingZero = true; continue; }
      if (s && (pendingZero || part < 1000)) s += "零";
      pendingZero = false;
      let p = "", zero = false;
      for (let i = 3; i >= 0; i--) {
        const x = Math.floor(part / Math.pow(10, i)) % 10;
        if (x) { if (zero) p += "零"; p += d[x] + u[i]; zero = false; } else if (p) zero = true;
      }
      s += p + big[g];
    }
    return s + "元整";
  }
  const money = (n) => (n || n === 0) && !isNaN(n) ? Number(n).toLocaleString("en-US") : "";

  function buildProposal(D, P, S) {
    const {
      Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle,
      ShadingType, VerticalAlign, HeadingLevel, PageBreak, TableOfContents, StyleLevel, Footer, Header,
      PageNumber, NumberFormat, LineRuleType, LevelFormat, TableLayoutType, HeightRule, ImageRun, Tab, TabStopType
    } = D;

    // ---------- 尺寸 ----------
    const cm = (v) => Math.round(v * 567);
    const pg = S.page;
    const W = 11906 - cm(pg.left) - cm(pg.right); // 內容寬度 (twips)
    const st = S.styles;
    const font = (s) => ({ ascii: s.en, hAnsi: s.en, cs: s.en, eastAsia: s.zh });
    const half = (pt) => Math.round(pt * 2);
    const align = (a) => ({ left: AlignmentType.LEFT, center: AlignmentType.CENTER, right: AlignmentType.RIGHT, justify: AlignmentType.JUSTIFIED }[a] || AlignmentType.LEFT);
    function spacing(s) {
      const o = { before: Math.round((s.before || 0) * 20), after: Math.round((s.after || 0) * 20) };
      if (s.lineType === "exact") { o.line = Math.round(s.lineValue * 20); o.lineRule = LineRuleType.EXACT; }
      else if (s.lineType === "atLeast") { o.line = Math.round(s.lineValue * 20); o.lineRule = LineRuleType.AT_LEAST; }
      else { o.line = Math.round((s.lineValue || 1) * 240); o.lineRule = LineRuleType.AUTO; }
      return o;
    }
    const indentOf = (s) => (s.indent ? { firstLine: Math.round(s.indent * s.size * 20) } : undefined);

    // ---------- 文字片段：【待補】自動黃色醒目 ----------
    function runs(text, opt = {}) {
      const out = [];
      String(text == null ? "" : text).split(/(【[^】]*】)/).forEach((seg) => {
        if (!seg) return;
        const isPh = /^【.*】$/.test(seg);
        out.push(new TextRun(Object.assign({ text: seg }, opt, isPh ? { highlight: "yellow" } : {})));
      });
      if (!out.length) out.push(new TextRun({ text: "" }));
      return out;
    }
    const body = (t) => new Paragraph({ style: "BodyPara", children: runs(t) });
    const plain = (t, o = {}) => new Paragraph(Object.assign({ children: runs(t) }, o));
    const blank = () => new Paragraph({ children: [] });

    // ---------- 標題與圖表編號 ----------
    let chap = 0, sec = 0, sub = 0, figN = 0, tabN = 0;
    // 每一章（含參考文獻、附錄）各為一個 Word 節，頁首右側顯示該章名稱
    const chunks = [];
    let out = [];
    function startChunk(title) {
      out = []; chunks.push({ title, children: out });
      out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(title)] }));
    }
    function H1(text) {
      chap++; sec = 0; sub = 0; figN = 0; tabN = 0;
      startChunk(`第${CH[chap]}章　${text}`);
    }
    function H2(text) { sec++; sub = 0; out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`${chap}-${sec}　${text}`)] })); }
    function H3(text) { sub++; out.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(`${chap}-${sec}-${sub}　${text}`)] })); }
    const nextFig = () => `圖 ${chap}-${++figN}`;
    const nextTab = () => `表 ${chap}-${++tabN}`;
    const figCaption = (no, t) => new Paragraph({ style: "FigureCaption", children: runs(`${no}　${t}`) });
    const tabCaption = (no, t) => new Paragraph({ style: "TableCaption", keepNext: true, children: runs(`${no}　${t}`) });
    const source = (t) => new Paragraph({ style: "SourceNote", children: runs(`（資料來源：${t || "本團隊整理"}）`) });

    // ---------- 表格 ----------
    const thin = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
    const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const allB = { top: thin, bottom: thin, left: thin, right: thin };
    const noB = { top: none, bottom: none, left: none, right: none };
    const HEAD = S.colors && S.colors.head || "D9E2F3";
    let keepAll = false;
    const tp = (t, o = {}) => new Paragraph({ style: "TableText", keepNext: keepAll || undefined, keepLines: keepAll || undefined, alignment: o.align ? align(o.align) : undefined, children: runs(t, o.bold ? { bold: true } : {}) });
    function cell(content, w, o = {}) {
      const paras = (Array.isArray(content) ? content : [content]).map((c) => (c instanceof Paragraph ? c : tp(c, o)));
      return new TableCell({
        width: { size: w, type: WidthType.DXA }, children: paras,
        columnSpan: o.span, rowSpan: o.rowSpan, verticalAlign: o.valign || VerticalAlign.CENTER,
        shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
        borders: o.borders || allB,
        margins: { top: 50, bottom: 50, left: 90, right: 90 },
      });
    }
    function widths(ratios) {
      const sum = ratios.reduce((a, b) => a + b, 0);
      const ws = ratios.map((r) => Math.floor((W * r) / sum));
      ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
      return ws;
    }
    // headers: string[]; rows: (string|{t,span,align,bold,fill})[][]
    function table(headers, rows, ratios, o = {}) {
      const ws = widths(ratios || headers.map(() => 1));
      const trs = [];
      if (headers && headers.length) trs.push(new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, ws[i], { bold: true, align: "center", fill: HEAD })) }));
      rows.forEach((r) => {
        let ci = 0;
        trs.push(new TableRow({
          cantSplit: true,
          children: r.map((c) => {
            const obj = typeof c === "object" && c !== null && !(c instanceof Paragraph) && !Array.isArray(c) ? c : { t: c };
            const span = obj.span || 1;
            let w = 0; for (let k = 0; k < span; k++) w += ws[ci + k] || 0; ci += span;
            return cell(obj.t == null ? "" : obj.t, w, { span: span > 1 ? span : undefined, align: obj.align || o.align, bold: obj.bold, fill: obj.fill });
          }),
        }));
      });
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows: trs, layout: TableLayoutType.FIXED });
    }
    function addTable(caption, headers, rows, ratios, src, o) {
      out.push(tabCaption(nextTab(), caption));
      out.push(table(headers, rows, ratios, o));
      out.push(source(src));
    }

    // ---------- 流程圖（以表格繪製，可直接在 Word 修改） ----------
    const BOX = S.colors && S.colors.box || "DEEAF6", BOX2 = "FFF2CC", BOX3 = "E2EFDA";
    const arrow = (w, span) => cell([new Paragraph({ style: "TableText", keepNext: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "▼", color: "4472C4" })] })], w, { span, borders: noB });
    // 縱向步驟流程
    function flowVertical(steps) {
      const ws = widths([1, 3, 1]);
      const rows = [];
      steps.forEach((s, i) => {
        if (i) rows.push(new TableRow({ children: [cell("", ws[0], { borders: noB }), arrow(ws[1]), cell("", ws[2], { borders: noB })] }));
        const fill = i === 0 || i === steps.length - 1 ? BOX2 : BOX;
        rows.push(new TableRow({ cantSplit: true, children: [cell("", ws[0], { borders: noB }), cell(s, ws[1], { align: "center", fill, bold: i === 0 || i === steps.length - 1 }), cell("", ws[2], { borders: noB })] }));
      });
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows, layout: TableLayoutType.FIXED });
    }
    // 計畫工作架構圖：啟動 → 各工作項目(並列) → 報告 → 成果
    function flowArchitecture(items) {
      const n = Math.max(1, Math.min(items.length, 6));
      const ws = widths(Array(n).fill(1));
      const full = (t, fill, bold) => new TableRow({ cantSplit: true, children: [cell(t, W, { span: n, align: "center", fill, bold })] });
      const arr = () => new TableRow({ children: [arrow(W, n)] });
      const rows = [
        full(`${P.projectName || "本計畫"}`, BOX2, true), arr(),
        full("契約簽訂 → 工作執行計畫書 → 期初簡報", BOX3), arr(),
        new TableRow({ cantSplit: true, children: items.slice(0, n).map((it, i) => cell(`${CN[i + 1]}、${it.title}`, ws[i], { align: "center", fill: BOX, bold: true })) }),
        new TableRow({ children: items.slice(0, n).map((it, i) => cell((it.subs || []).map((s, k) => tp(`${k + 1}. ${s.title}`)).concat((it.subs || []).length ? [] : [tp("【待補】")]), ws[i], { valign: VerticalAlign.TOP })) }),
        arr(),
        full("期中報告　→　期末報告（成果發表／說明會）", BOX3), arr(),
        full("成果報告書定稿・資料交付・驗收", BOX2, true),
      ];
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows, layout: TableLayoutType.FIXED });
    }
    // 組織架構圖
    function orgChart(groups) {
      const n = Math.max(1, Math.min(groups.length, 6));
      const ws = widths(Array(n).fill(1));
      const mid = (t, fill) => new TableRow({ cantSplit: true, children: [cell((Array.isArray(t) ? t : [t]).map((x, i) => tp(x, { align: "center", bold: i === 0 })), W, { span: n, fill })] });
      const arr = () => new TableRow({ children: [arrow(W, n)] });
      const rows = [
        mid(["計畫主持人", "【待補：姓名／職稱】"], BOX2), arr(),
        mid("協同主持人　【待補：姓名】　｜　品質管理及專家顧問群　【待補】", BOX3), arr(),
        new TableRow({ cantSplit: true, children: groups.slice(0, n).map((g, i) => cell(g.name, ws[i], { align: "center", bold: true, fill: BOX })) }),
        new TableRow({ children: groups.slice(0, n).map((g, i) => cell([tp(g.duty || ""), tp("組長：【待補】"), tp("組員：【待補】")], ws[i], { valign: VerticalAlign.TOP })) }),
      ];
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows, layout: TableLayoutType.FIXED });
    }
    // PDCA
    function pdca(p) {
      const ws = widths([1, 1]);
      const c = (k, t, fill) => cell([tp(k, { bold: true, align: "center" }), tp(t)], ws[0], { fill, valign: VerticalAlign.TOP });
      return new Table({
        width: { size: W, type: WidthType.DXA }, columnWidths: ws, layout: TableLayoutType.FIXED, rows: [
          new TableRow({ cantSplit: true, children: [c("計畫 Plan", p.plan, BOX), c("執行 Do", p.do, BOX3)] }),
          new TableRow({ cantSplit: true, children: [c("改善 Act", p.act, BOX2), c("查核 Check", p.check, "FBE4D5")] }),
        ],
      });
    }
    function figurePlaceholder(desc) {
      return new Table({
        width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
        rows: [new TableRow({ height: { value: cm(6), rule: HeightRule.ATLEAST }, children: [cell([tp("【圖片預留位置】", { align: "center", bold: true }), tp(desc ? `建議內容：${desc}` : "", { align: "center" })], W, { borders: { top: { style: BorderStyle.DASHED, size: 6, color: "7F7F7F" }, bottom: { style: BorderStyle.DASHED, size: 6, color: "7F7F7F" }, left: { style: BorderStyle.DASHED, size: 6, color: "7F7F7F" }, right: { style: BorderStyle.DASHED, size: 6, color: "7F7F7F" } } })] })],
      });
    }
    function addFigure(caption, build, src) {
      keepAll = true; const tbl = typeof build === "function" ? build() : build; keepAll = false;
      out.push(tbl);
      out.push(figCaption(nextFig(), caption));
      out.push(source(src));
    }

    // ---------- 清單 ----------
    let listInst = 0;
    function list(items, kind) {
      listInst++;
      (items || []).forEach((t) => out.push(new Paragraph({ style: "ListPara", numbering: { reference: kind === "bullet" ? "bul" : "num", level: 0, instance: listInst }, children: runs(t) })));
    }
    function blocks(bs) {
      (bs || []).forEach((b) => {
        if (!b) return;
        if (b.t === "p") out.push(body(b.text));
        else if (b.t === "list") list(b.items, b.style === "bullet" ? "bullet" : "num");
        else if (b.t === "table" && Array.isArray(b.headers) && Array.isArray(b.rows)) {
          const hs = b.headers.slice(0, 6);
          addTable(b.caption || "【待補：表名】", hs, b.rows.map((r) => hs.map((_, i) => (r[i] == null ? "" : String(r[i])))), null, b.source);
        } else if (b.t === "figure") {
          if (Array.isArray(b.steps) && b.steps.length > 1) addFigure(b.caption || "作業流程圖", () => flowVertical(b.steps));
          else addFigure(b.caption || "【待補：圖名】", () => figurePlaceholder(b.desc));
        }
      });
    }
    const paras = (arr) => (Array.isArray(arr) ? arr : arr ? [arr] : []).forEach((t) => out.push(body(t)));

    // ================= 內容 =================
    const C = P.content || {};
    const items = P.workItems || [];
    const short = P.agencyShort || "貴機關";
    const team = "本團隊";
    const groups = (C.groups && C.groups.length ? C.groups : items.map((it, i) => ({ name: `${CN[i + 1]}、${it.title.slice(0, 14)}組`, duty: `負責「${it.title}」相關工作` }))).slice(0, 6);

    // 第壹章 前言
    H1("前言");
    H2("計畫緣起與目的");
    paras(C.origin || [P.background || "【待補：計畫緣起】", `${team}彙整招標文件之需求，就計畫目的、工作範圍及執行策略提出本服務建議書，以協助${short}如期如質完成本計畫。【待補：補充本團隊觀點】`]);
    if (C.objectives && C.objectives.length) { out.push(body("本計畫之主要目標如下：")); list(C.objectives); }
    H2("委託工作範圍及項目");
    out.push(body(`依據「${P.projectName || "本計畫"}」採購說明之規定，本計畫委託工作項目及內容如下：`));
    items.forEach((it, i) => {
      out.push(new Paragraph({ style: "BodyPara", indent: { firstLine: 0 }, children: runs(`${CN[i + 1]}、${it.title}`, { bold: true }) }));
      listInst++;
      const lines2 = (it.subs && it.subs.length) ? it.subs : [{ req: it.req && it.req !== it.title ? it.req : "" }];
      lines2.forEach((s) => {
        const ls = String(s.req || s.title || "").split("\n").filter(Boolean);
        if (!ls.length) return;
        if (it.subs && it.subs.length) out.push(new Paragraph({ style: "ListPara", numbering: { reference: "num", level: 0, instance: listInst }, children: runs(ls[0]) }));
        else out.push(body(ls[0]));
        ls.slice(1).forEach((x) => out.push(new Paragraph({ style: "ListPara", numbering: { reference: "num", level: 1, instance: listInst }, children: runs(x) })));
      });
    });
    if (P.period) out.push(body(`履約期限：${P.period}。`));
    out.push(body(`本計畫工作架構如圖${chap}-${figN + 1}所示，各工作項目與本建議書章節之對照如表${chap}-${tabN + 1}所示。`));
    addFigure("本計畫工作架構圖", () => flowArchitecture(items));
    addTable("本計畫工作項目與章節對照表", ["項次", "工作項目", "對應章節"], items.map((it, i) => [{ t: CN[i + 1], align: "center" }, it.title, { t: `2-${i + 3}`, align: "center" }]), [1, 6, 2]);
    H2("建議書架構與評選項目對照");
    out.push(body(`為便於評選委員審閱，本建議書依評選須知「內容要領」編排，評選項目與章節對照如表${chap}-${tabN + 1}。`));
    const evalMap = (P.evaluation || []).map((e) => [e.item, { t: e.weight ? String(e.weight) : "", align: "center" }, e.chapters || "【待補】"]);
    addTable("評選項目與本建議書章節對照表", ["評選項目", "配分", "對應章節"], evalMap.length ? evalMap : [["【待補】", "", ""]], [5, 1.3, 3]);
    H2("承攬優勢");
    paras(C.advantages || [`${team}具備本計畫所需之專業技術與執行經驗，【待補：說明公司相關實績、核心技術、人力及設備優勢】。`]);

    // 第貳章 專案執行規劃
    H1("專案執行規劃");
    H2("對本計畫及問題之瞭解");
    paras(C.understanding || ["【待補：說明對計畫背景、現況與課題之瞭解】"]);
    if (C.issues && C.issues.length) {
      out.push(body(`${team}歸納本計畫之主要課題及對策如表${chap}-${tabN + 1}。`));
      addTable("本計畫課題與對策", ["課題", "對策"], C.issues.map((x) => [x.issue, x.strategy]), [1, 2]);
    }
    H2("工作項目規劃架構與執行流程");
    paras(C.planIntro || [`${team}依各工作項目之性質與期程，規劃整體執行流程如圖${chap}-${figN + 1}，各項工作之執行方法說明於後續各節。`]);
    addFigure("整體執行流程圖", () => flowVertical(C.mainFlow && C.mainFlow.length > 2 ? C.mainFlow : ["契約簽訂、工作啟動", "資料蒐集與工作執行計畫書", "各工作項目執行（詳後續各節）", "期中報告與檢討修正", "期末報告與成果發表", "成果報告書定稿及驗收"]));
    items.forEach((it, i) => {
      const W8 = (C.work && C.work[i]) || {};
      H2(it.title);
      paras(W8.overview || [`【工作需求】${it.req || it.title}`]);
      if (W8.flow && W8.flow.length > 2) {
        out.push(body(`本項工作之作業流程如圖${chap}-${figN + 1}。`));
        addFigure(`${it.title}作業流程圖`, () => flowVertical(W8.flow));
      }
      if (!(it.subs || []).length) {
        if (W8.blocks && W8.blocks.length) blocks(W8.blocks);
        else out.push(body("【待補：執行方法、作業步驟、使用資料與工具、人力配置及預期產出】"));
      }
      (it.subs || []).forEach((s, k) => {
        H3(s.title);
        const sb = W8.subs && W8.subs[k];
        if (sb && sb.blocks && sb.blocks.length) blocks(sb.blocks);
        else {
          out.push(body(`【工作需求】${s.req || s.title}`));
          out.push(body("【待補：執行方法、作業步驟、使用資料與工具、人力配置及預期產出】"));
        }
      });
    });
    H2("執行方式之創新性及可行性");
    paras(C.innovation || ["【待補：說明本團隊執行方式之創新作法】"]);
    if (C.feasibility && C.feasibility.length) addTable("執行方式可行性分析", ["面向", "說明"], C.feasibility.map((x) => [x.aspect, x.desc]), [1, 3]);
    H2("對委託工作之建議");
    if (C.suggestions && C.suggestions.length) list(C.suggestions); else out.push(body("【待補：對委託工作之修正或建議事項】"));

    // 第參章 預期成果與效益
    H1("預期成果與效益");
    H2("預期成果");
    out.push(body(`本計畫預期完成之成果項目如表${chap}-${tabN + 1}。`));
    const outc = (C.outcomes && C.outcomes.length ? C.outcomes.map((x) => [x.item, x.desc]) : items.map((it) => [it.title, "【待補：量化成果】"]));
    addTable("本計畫預期成果", ["成果項目", "內容說明"], outc, [2, 5]);
    if (P.deliverables && P.deliverables.length) {
      out.push(body(`依採購說明，本計畫應提送之成果文件如表${chap}-${tabN + 1}。`));
      addTable("成果提送項目與數量", ["項次", "成果項目", "說明"], P.deliverables.map((d, i) => [{ t: String(i + 1), align: "center" }, d.name, d.desc]), [1, 3, 6]);
    }
    H2("預期效益及其他附加價值");
    paras(C.benefits || ["【待補：預期效益】"]);
    if (C.addedValue && C.addedValue.length) { out.push(body("此外，本團隊另提供以下附加價值：")); list(C.addedValue); }

    // 第肆章 專案品質及進度控管
    H1("專案品質及進度控管");
    H2("預定工作進度");
    paras(C.schedule || [`本計畫預定工作進度如表${chap}-${tabN + 1}，${team}將依進度確實執行，並於各查核點提交成果。`]);
    // 甘特圖
    (function gantt() {
      const months = P.months && P.months.length ? P.months : [];
      const n = months.length || 12;
      const ms = months.length ? months : Array.from({ length: 12 }, (_, i) => ({ y: "", m: i + 1 }));
      const nameW = Math.round(W * (n > 12 ? 0.28 : 0.32));
      const mw = Math.floor((W - nameW) / n);
      const ws = [nameW].concat(Array(n).fill(mw)); ws[n] += W - ws.reduce((a, b) => a + b, 0);
      const gp = (t, o = {}) => new Paragraph({ style: "TableText", alignment: o.align ? align(o.align) : undefined, children: runs(t, Object.assign({ size: half(Math.min(st.table.size, 10)) }, o.bold ? { bold: true } : {})) });
      const gc = (t, w, o = {}) => cell([gp(t, o)], w, o);
      const rows = [];
      // 年度列
      const years = []; ms.forEach((m) => { const l = years[years.length - 1]; if (l && l.y === m.y) l.n++; else years.push({ y: m.y, n: 1 }); });
      rows.push(new TableRow({ tableHeader: true, children: [gc("工作項目", ws[0], { bold: true, align: "center", fill: HEAD, rowSpan: 2 })].concat(years.map((y) => { let w = 0; for (let k = 0; k < y.n; k++) w += mw; return gc(y.y ? `${y.y}年` : "年度", w, { span: y.n > 1 ? y.n : undefined, bold: true, align: "center", fill: HEAD }); })) }));
      rows.push(new TableRow({ tableHeader: true, children: ms.map((m, i) => gc(String(m.m), ws[i + 1], { align: "center", bold: true, fill: HEAD })) }));
      const bar = "8EAADB", bar2 = "B4C6E7";
      const span = (s) => (s && s.span ? s.span : [0, n - 1]);
      items.forEach((it, i) => {
        const sp = span(it);
        rows.push(new TableRow({ cantSplit: true, children: [gc(`${CN[i + 1]}、${it.title}`, ws[0], { bold: true })].concat(ms.map((_, k) => gc("", ws[k + 1], { fill: k >= sp[0] && k <= sp[1] ? bar : undefined }))) }));
        (it.subs || []).forEach((s, j) => {
          const ss = s.span || sp;
          rows.push(new TableRow({ cantSplit: true, children: [gc(`　${j + 1}. ${s.title}`, ws[0])].concat(ms.map((_, k) => gc("", ws[k + 1], { fill: k >= ss[0] && k <= ss[1] ? bar2 : undefined }))) }));
        });
      });
      // 查核點
      const ck = (P.checkpoints && P.checkpoints.length ? P.checkpoints : [{ name: "工作執行計畫書（期初）", idx: 0 }, { name: "期中報告", idx: Math.floor((n - 1) / 2) }, { name: "期末報告", idx: Math.max(0, n - 2) }, { name: "成果報告書定稿", idx: n - 1 }]);
      ck.forEach((c) => rows.push(new TableRow({ cantSplit: true, children: [gc(c.name, ws[0], { bold: true })].concat(ms.map((_, k) => gc(k === c.idx ? "★" : "", ws[k + 1], { align: "center" }))) })));
      // 累計進度
      rows.push(new TableRow({ cantSplit: true, children: [gc("預定累計進度(%)", ws[0], { bold: true, fill: "F2F2F2" })].concat(ms.map((_, k) => gc(String(Math.round(((k + 1) / n) * 100)), ws[k + 1], { align: "center", fill: "F2F2F2" }))) }));
      out.push(tabCaption(nextTab(), "本計畫預定工作進度表（甘特圖）"));
      out.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws, rows, layout: TableLayoutType.FIXED }));
      out.push(new Paragraph({ style: "SourceNote", children: runs("註：★為查核點；色塊為預定執行期間，可依實際規劃調整儲存格網底。") }));
    })();
    H2("查核點及分段進度");
    if (P.milestones && P.milestones.length) {
      out.push(body(`依採購說明之分段進度規定，本計畫查核時間點及應完成工作項目如表${chap}-${tabN + 1}。`));
      addTable("本計畫查核時間點及工作項目", ["項次", "查核點", "預定時間", "應完成工作項目"], P.milestones.map((m, i) => [{ t: String(i + 1), align: "center" }, m.stage, m.when || "【待補】", (m.tasks || []).map((t, k) => tp(`${k + 1}. ${t}`))]), [0.8, 1.8, 1.8, 6]);
    } else out.push(body("【待補：各查核點時間及應完成工作】"));
    H2("品質管理措施");
    paras(C.qc || [`${team}依 PDCA 循環建立品質管理機制，如圖${chap}-${figN + 1}。【待補】`]);
    addFigure("品質管理 PDCA 循環", () => pdca(C.pdca || { plan: "訂定工作執行計畫、品質目標與檢核表【待補】", do: "依標準作業程序執行各項工作【待補】", check: "內部審查、主持人覆核及進度查核【待補】", act: `檢討缺失、矯正預防並回饋${short}意見【待補】` }));
    if (C.risk && C.risk.length) addTable("風險管理及因應對策", ["潛在風險", "因應對策"], C.risk.map((r) => [r.risk, r.response]), [2, 5]);
    H2("資訊安全及個人資料保護");
    paras(C.infosec || [`${team}將遵守行政院資通安全相關規範及${short}委外廠商個人資料與資通安全保護責任規定，並於契約終止時依規定返還、刪除或銷毀機敏資料。【待補】`]);

    // 第伍章 專案管理及履約能力
    H1("專案管理及履約能力");
    H2("專案工作團隊組織及架構");
    paras(C.org || [`${team}依本計畫工作項目成立專案團隊，組織架構如圖${chap}-${figN + 1}。`]);
    addFigure("專案組織架構圖", () => orgChart(groups));
    out.push(body(`各組之主要工作項目如表${chap}-${tabN + 1}。`));
    addTable("各組工作人員之主要工作項目", ["組別", "主要工作內容", "負責人"], groups.map((g) => [g.name, g.duty || "", "【待補】"]), [2, 5, 1.5]);
    H2("專案成員學經歷");
    out.push(body(`本計畫主要工作人員之學經歷及於本計畫擔任之工作如表${chap}-${tabN + 1}（工作成員均已徵得當事人同意）。`));
    addTable("主要工作人員學經歷表", ["職稱", "姓名", "學歷(科系)", "專長／證照", "相關經歷(性質、年份)", "本計畫擔任工作"],
      ["計畫主持人", "協同主持人", "專案經理", "工程師", "工程師", "駐點人員"].map((r) => [r, "【待補】", "", "", "", ""]), [1.3, 1, 1.4, 1.4, 2, 1.6]);
    H2("執行中案件及人力配置");
    out.push(body(`${team}目前執行中之案件及人力配置如表${chap}-${tabN + 1}，經評估可充分投入本計畫所需人力。`));
    addTable("執行中案件及人力配置表", ["計畫名稱", "委託機關", "執行期間", "投入人力", "本計畫可投入比例"], [1, 2, 3].map(() => ["【待補】", "", "", "", ""]), [3, 2, 1.6, 1, 1.4]);
    H2("如期履約說明");
    paras(C.onTime || ["【待補：說明進度管控、人力調度及備援機制，確保如期如質履約】"]);

    // 第陸章 廠商規模與實績及主持人經驗
    H1("廠商規模與實績及主持人經驗");
    H2("投標廠商基本資料");
    addTable("投標廠商基本資料", ["項目", "內容"], [["廠商名稱", P.bidder || "【待補】"], ["統一編號", "【待補】"], ["負責人", "【待補】"], ["資本額", "【待補】"], ["營業項目", "【待補】"], ["員工人數", "【待補】"], ["組織概況", "【待補】"], ["地址／電話", "【待補】"]], [1, 3]);
    H2("與本計畫性質相關之工作實績");
    out.push(body(`${team}近年執行與本計畫性質相關之工作實績如表${chap}-${tabN + 1}，相關證明文件詳附錄一。`));
    addTable("相關工作實績表", ["項次", "計畫名稱", "委託機關", "執行期間", "契約金額(元)", "與本案關聯性"], [1, 2, 3, 4, 5].map((k) => [{ t: String(k), align: "center" }, "【待補】", "", "", "", ""]), [0.7, 3, 1.6, 1.4, 1.3, 1.6]);
    H2("計畫主持人相關工作經驗");
    out.push(body("【待補：計畫主持人學經歷、專長、證照及與本計畫性質相關之工作經驗】"));
    addTable("計畫主持人相關計畫經驗", ["年度", "計畫名稱", "委託機關", "擔任職務"], [1, 2, 3, 4].map(() => ["【待補】", "", "", ""]), [1, 4, 2, 1.4]);

    // 第柒章 CSR
    H1("廠商企業社會責任(CSR)指標");
    H2("全職從事本採購案員工之薪資");
    paras((C.csr && C.csr.salary) || ["【待補：說明履約期間全職從事本採購案員工之薪資，至少超過招標公告日勞動部公告最低工資之倍數，並檢附薪資證明文件】"]);
    addTable("全職從事本採購案員工薪資表", ["職稱", "人數", "月薪(元)", "為最低工資倍數", "證明文件"], [["【待補】", "", "", "", "附錄"], ["", "", "", "", ""]], [2, 1, 1.5, 1.6, 1.5]);
    H2("提供員工「工作與生活平衡」措施");
    paras((C.csr && C.csr.balance) || ["【待補：友善家庭、友善性別、彈性工時、員工協助方案、健康促進等措施，並檢附證明文件】"]);
    H2("辦理綠色採購");
    paras((C.csr && C.csr.green) || ["【待補：最近一年度綠色採購申報金額，並檢附「綠色採購金額證明文件」】"]);

    // 第捌章 標價清單
    H1("標價清單及費用分析");
    H2("標價清單");
    const total = P.budget ? Number(P.budget) : null;
    paras(C.budgetNote || [`${team}依招標文件所指定之工作項目、服務期限及對本計畫之瞭解，評估所需服務費用為新臺幣${total ? money(total) + "元整" : "【待補】元整"}（含稅），標價清單如表${chap}-${tabN + 1}，費用分析如表${chap}-${tabN + 2}。`]);
    const pItems = P.priceItems && P.priceItems.length ? P.priceItems : items.map((it) => ({ title: it.title, qty: 1, unit: "式" }));
    addTable("投標標價清單", ["項次", "服務項目", "數量", "單位", "單價(元)", "總價(元)", "說明"],
      pItems.map((p, i) => [{ t: CN[i + 1], align: "center" }, p.title, { t: String(p.qty || 1), align: "center" }, { t: p.unit || "式", align: "center" }, { t: "【待填】", align: "right" }, { t: "【待填】", align: "right" }, "詳採購說明"])
        .concat([[{ t: "合計", span: 5, align: "center", bold: true }, { t: total ? money(total) : "【待填】", align: "right", bold: true }, ""],
          [{ t: `總標價新臺幣　${total ? upperAmount(total) : "【待填】"}`, span: 7, bold: true }]]),
      [0.7, 4, 0.8, 0.8, 1.4, 1.5, 1.4]);
    H2("服務費用分析");
    const staff = ["計畫主持人", "協同主持人", "專案經理", "工程師", "助理工程師"].concat(P.hasResident ? ["駐點人員"] : []);
    const fa = [];
    fa.push([{ t: "一", align: "center", bold: true }, { t: "直接薪資", bold: true }, "", "", "", "", ""]);
    staff.forEach((s, i) => fa.push([{ t: `${i + 1}`, align: "center" }, s, { t: "人月", align: "center" }, "", "", "", ""]));
    fa.push(["", { t: "小計（一）", bold: true }, "", "", "", "", ""]);
    [["二", "管理費用（含勞健保、退休金等）", "式"], ["三", "公費", "式"]].forEach((r) => fa.push([{ t: r[0], align: "center", bold: true }, { t: r[1], bold: true }, { t: r[2], align: "center" }, "1", "", "", ""]));
    fa.push([{ t: "四", align: "center", bold: true }, { t: "其他直接費用", bold: true }, "", "", "", "", ""]);
    ["差旅費（現勘交通、住宿）", "印刷及報告製作費", "會議費（說明會、工作會議）", "保險費", "資料及設備使用費", "雜支"].forEach((s, i) => fa.push([{ t: `${i + 1}`, align: "center" }, s, { t: "式", align: "center" }, "1", "", "", ""]));
    fa.push(["", { t: "小計（四）", bold: true }, "", "", "", "", ""]);
    fa.push([{ t: "五", align: "center", bold: true }, { t: "營業稅（5%）", bold: true }, { t: "式", align: "center" }, "1", "", "", ""]);
    fa.push([{ t: "合計（一～五）", span: 5, align: "center", bold: true }, { t: total ? money(total) : "【待填】", align: "right", bold: true }, { t: "100%", align: "center" }]);
    addTable("服務費用分析表", ["項次", "費用項目", "單位", "數量", "單價(元)", "複價(元)", "百分比"], fa, [0.8, 3.6, 0.9, 0.9, 1.3, 1.7, 1]);
    out.push(new Paragraph({ style: "SourceNote", children: runs("註：依「機關委託技術服務廠商評選及計費辦法」之服務成本加公費法編列，金額請依實際估算填列。") }));

    // 參考文獻
    startChunk("參考文獻");
    const refs = P.references || [];
    const cats = [];
    refs.forEach((r) => { let c = cats.find((x) => x.name === r.category); if (!c) cats.push((c = { name: r.category, items: [] })); c.items.push(r.text); });
    ["研究報告", "期刊與論文"].forEach((k) => { if (!cats.find((c) => c.name === k)) cats.push({ name: k, items: ["【待補】"] }); });
    cats.forEach((c, i) => {
      out.push(new Paragraph({ style: "BodyPara", indent: { firstLine: 0 }, keepNext: true, children: [new TextRun({ text: `${CN[i + 1]}、【${c.name}】`, bold: true })] }));
      listInst++;
      c.items.forEach((t) => out.push(new Paragraph({ style: "RefPara", numbering: { reference: "num", level: 0, instance: listInst }, children: runs(t) })));
    });

    // 附錄
    ["相關工作實績證明文件", "主要工作人員學經歷證明文件", "企業社會責任(CSR)相關證明文件"].forEach((t, i) => {
      startChunk(`附錄${CN[i + 1]}　${t}`);
      out.push(body("【待補：檢附相關證明文件影本】"));
    });

    // ================= LOGO、封面與目錄 =================
    const color = String(pg.color || "2E7D32").replace(/^#/, "").toUpperCase();
    const logos = P.logos || {};
    // 依高度等比例縮放（單位：像素，96 dpi）
    function logo(l, h, maxW) {
      if (!l || !l.data || !l.w || !l.h) return null;
      let w = Math.round((h * l.w) / l.h);
      if (w > maxW) { h = Math.round((h * maxW) / w); w = maxW; }
      return new ImageRun({ type: "png", data: l.data, transformation: { width: w, height: h } });
    }
    const coverRun = (t, size, bold) => new TextRun({ text: t, size: half(size), bold, font: font(st.h1) });
    const agencyLogoCover = logo(logos.agency, 90, 260), bidderLogoCover = logo(logos.bidder, 56, 300);
    const cover = [
      ...Array(agencyLogoCover ? 2 : 4).fill(0).map(blank),
      agencyLogoCover ? new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [agencyLogoCover] }) : null,
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [coverRun(P.agency || "【待補：機關名稱】", 20, true)] }),
      blank(),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 600, lineRule: LineRuleType.AT_LEAST, after: 360 }, children: [coverRun(`「${P.projectName || "【待補：計畫名稱】"}」`, 24, true)] }),
      blank(),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [coverRun("服 務 建 議 書", 32, true)] }),
      ...Array(bidderLogoCover ? 7 : 9).fill(0).map(blank),
      bidderLogoCover ? new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [bidderLogoCover] }) : null,
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [coverRun(`投標廠商：${P.bidder || "【待補】"}`, 16, true)] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [coverRun(`中華民國 ${P.rocYear || "　　"} 年 ${P.rocMonth || "　　"} 月`, 16, true)] }),
    ].filter(Boolean);

    // 目錄頁抬頭：機關、標案名稱、服務建議書、目錄名稱
    const tocHead = (t) => [P.agency || "", P.projectName || "", "服務建議書"].filter(Boolean)
      .map((x) => new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { after: 40 }, children: [new TextRun({ text: x, bold: true, size: half(Math.max(st.h1.size - 2, 12)), font: font(st.h1) })] }))
      .concat([new Paragraph({ style: "TocTitle", children: [new TextRun(t)] })]);
    const tocPages = [
      { title: "目錄", children: tocHead("目　錄").concat([new TableOfContents("目錄", { hyperlink: true, headingStyleRange: "1-3" })]) },
      { title: "表目錄", children: tocHead("表 目 錄").concat([new TableOfContents("表目錄", { hyperlink: true, stylesWithLevels: [new StyleLevel("表標題", 1)] })]) },
      { title: "圖目錄", children: tocHead("圖 目 錄").concat([new TableOfContents("圖目錄", { hyperlink: true, stylesWithLevels: [new StyleLevel("圖標題", 1)] })]) },
    ];

    // 頁首：左＝機關 LOGO＋標案名稱；右＝本節名稱（第一層）
    const hfFont = font(st.body);
    const header = (title) => new Header({ children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: W }],
      spacing: { after: 0, line: 240, lineRule: LineRuleType.AUTO },
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color, space: 3 } },
      children: [logo(logos.agency, 30, 120), logos.agency ? new TextRun({ text: " " }) : null,
        new TextRun({ text: P.projectName || "", bold: true, size: 20, color, font: hfFont }),
        new TextRun({ children: [new Tab(), title || ""], bold: true, size: 20, color, font: hfFont })].filter(Boolean),
    })] });
    // 頁尾：左＝投標廠商 LOGO＋名稱；中＝頁碼（以無框線三欄表格排版，頁碼永遠置中）
    const footer = () => {
      // 有 LOGO：左 LOGO、右名稱；無 LOGO：名稱置左
      const showB = pg.footerBidder !== false;
      const nameRun = showB && P.bidder ? [new TextRun({ text: P.bidder, bold: true, size: 18, color: "404040", font: hfFont })] : [];
      const lg = showB ? logo(logos.bidder, 30, 190) : null;
      const left = lg ? [lg] : nameRun, right = lg ? nameRun : [];
      if (!left.length && !pg.pageNum) return undefined;
      const pn = pg.pageNum ? [new TextRun({ children: [PageNumber.CURRENT], size: 20, font: hfFont })] : [];
      const ws = [Math.round(W * 0.42), W - 2 * Math.round(W * 0.42), Math.round(W * 0.42)];
      const line = left.length || right.length ? { style: BorderStyle.SINGLE, size: 6, color: "808080" } : none;
      const fp = (kids, al) => new Paragraph({ alignment: al, spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO }, children: kids });
      const fc = (kids, w, al) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, borders: { top: line, bottom: none, left: none, right: none }, margins: { top: 60, bottom: 0, left: 0, right: 0 }, children: [fp(kids, al)] });
      return new Footer({ children: [new Table({
        width: { size: W, type: WidthType.DXA }, columnWidths: ws, layout: TableLayoutType.FIXED,
        borders: { top: line, bottom: noB.bottom, left: noB.left, right: noB.right, insideHorizontal: noB.top, insideVertical: noB.top },
        rows: [new TableRow({ children: [fc(left, ws[0], AlignmentType.LEFT), fc(pn, ws[1], AlignmentType.CENTER), fc(right, ws[2], AlignmentType.RIGHT)] })],
      }), new Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT }, children: [] })] });
    };
    const pageProps = (fmt, start) => ({
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: cm(pg.top), bottom: cm(pg.bottom), left: cm(pg.left), right: cm(pg.right), header: cm(1.2), footer: cm(1.2) },
        pageNumbers: fmt ? (start ? { start: 1, formatType: fmt } : { formatType: fmt }) : undefined,
      },
    });
    const section = (c, fmt, start) => {
      const f = footer();
      return { properties: pageProps(fmt, start), headers: pg.header ? { default: header(c.title) } : undefined, footers: f ? { default: f } : undefined, children: c.children };
    };

    const S2 = (k, extra) => Object.assign({ run: { font: font(st[k]), size: half(st[k].size), bold: !!st[k].bold }, paragraph: Object.assign({ spacing: spacing(st[k]), alignment: align(st[k].align), indent: indentOf(st[k]) }, extra || {}) });
    const bodyIndent = Math.round((st.body.indent || 0) * st.body.size * 20);

    return new Document({
      creator: P.bidder || "",
      title: `${P.projectName || ""} 服務建議書`,
      features: { updateFields: true },
      styles: {
        default: {
          document: { run: { font: font(st.body), size: half(st.body.size) } },
          heading1: S2("h1", { keepNext: true, keepLines: true }),
          heading2: S2("h2", { keepNext: true, keepLines: true }),
          heading3: S2("h3", { keepNext: true, keepLines: true }),
        },
        paragraphStyles: [
          Object.assign({ id: "BodyPara", name: "本文", basedOn: "Normal", next: "BodyPara", quickFormat: true }, S2("body")),
          Object.assign({ id: "ListPara", name: "條列", basedOn: "Normal", quickFormat: true }, S2("body", { indent: undefined })),
          Object.assign({ id: "RefPara", name: "參考文獻", basedOn: "Normal", quickFormat: true }, S2("body", { indent: undefined, alignment: AlignmentType.LEFT })),
          Object.assign({ id: "TableText", name: "表格文字", basedOn: "Normal", quickFormat: true }, S2("table", { indent: undefined })),
          Object.assign({ id: "TableCaption", name: "表標題", basedOn: "Normal", next: "Normal", quickFormat: true }, S2("caption", { keepNext: true, indent: undefined })),
          Object.assign({ id: "FigureCaption", name: "圖標題", basedOn: "Normal", next: "Normal", quickFormat: true }, S2("caption", { indent: undefined })),
          { id: "SourceNote", name: "資料來源", basedOn: "Normal", run: { font: font(st.table), size: half(Math.max(st.table.size - 1, 8)) }, paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 40, after: 200 } } },
          Object.assign({ id: "TocTitle", name: "目錄標題", basedOn: "Normal", next: "Normal" }, S2("h1", { keepNext: true })),
        ],
      },
      numbering: {
        config: [
          { reference: "num", levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(st.body.size * 20 * 2) + 120, hanging: Math.round(st.body.size * 20 * 1.6) } } } },
            { level: 1, format: LevelFormat.DECIMAL, text: "(%2)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(st.body.size * 20 * 4) + 120, hanging: Math.round(st.body.size * 20 * 2) } } } },
          ] },
          { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u25CF", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: Math.round(st.body.size * 20 * 2) + 120, hanging: Math.round(st.body.size * 20) } }, run: { size: 14 } } }] },
        ],
      },
      sections: [
        { properties: pageProps(null, false), children: cover },
        ...tocPages.map((c, i) => section(c, NumberFormat.UPPER_ROMAN, i === 0)),
        ...chunks.map((c, i) => section(c, NumberFormat.DECIMAL, i === 0)),
      ],
    });
  }

  root.ProposalGen = { buildProposal, upperAmount, CN, CH };
  if (typeof module !== "undefined") module.exports = root.ProposalGen;
})(typeof window !== "undefined" ? window : globalThis);
