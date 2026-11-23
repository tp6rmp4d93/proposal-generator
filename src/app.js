/* 服務建議書初稿產生器：介面與流程 */
"use strict";
const $ = (s, r = document) => r.querySelector(s);
const el = (tag, attrs = {}, ...kids) => { const e = document.createElement(tag); for (const [k, v] of Object.entries(attrs)) { if (k === "class") e.className = v; else if (k.startsWith("on")) e.addEventListener(k.slice(2), v); else if (v !== undefined && v !== null && v !== false) e.setAttribute(k, v === true ? "" : v); } kids.flat().forEach((c) => c != null && e.append(c)); return e; };
const toast = (m) => { const t = $("#toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove("on"), 2200); };
const CNN = ProposalGen.CN;

/* ================= 狀態 ================= */
const state = { bidder: "", docs: [], P: null, sample: null, downloads: null, blob: null, fileName: "", ctl: null, logos: { agency: null, bidder: null } };

/* ================= LOGO ================= */
// 一律轉成 PNG（支援 SVG、WebP 等），最長邊 800px，保留透明背景
const BIDDER_KEY = "proposalgen.bidder.v1";
function imageToPng(src) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth || 300, h = img.naturalHeight || 150;
      const k = Math.min(1, 800 / Math.max(w, h)); w = Math.max(1, Math.round(w * k)); h = Math.max(1, Math.round(h * k));
      const c = document.createElement("canvas"); c.width = w; c.height = h;
      c.getContext("2d").drawImage(img, 0, 0, w, h);
      const url = c.toDataURL("image/png");
      const bin = atob(url.split(",")[1]); const data = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
      res({ data, w, h, url });
    };
    img.onerror = () => rej(new Error("無法讀取這張圖片"));
    img.src = src;
  });
}
function renderLogos() {
  [["agency", "#lgAgencyBox"], ["bidder", "#lgBidderBox"]].forEach(([k, box]) => {
    const l = state.logos[k], b = $(box); b.textContent = "";
    if (l) b.append(el("img", { src: l.url, alt: k === "agency" ? "機關 LOGO" : "投標廠商 LOGO" })); else b.textContent = "無";
    document.querySelector(`[data-logo="${k}"][data-act="clear"]`).hidden = !l;
  });
  updatePreview();
}
function saveBidder() {
  try { localStorage.setItem(BIDDER_KEY, JSON.stringify({ name: state.bidder || "", logo: state.logos.bidder && state.logos.bidder.url.length < 1500000 ? state.logos.bidder.url : "" })); } catch (e) {}
}
let logoTarget = null;
document.addEventListener("click", (e) => {
  const b = e.target.closest("[data-logo]"); if (!b) return;
  if (b.dataset.act === "pick") { logoTarget = b.dataset.logo; $("#logoIn").click(); }
  else { state.logos[b.dataset.logo] = null; if (b.dataset.logo === "bidder") saveBidder(); renderLogos(); }
});
$("#logoIn").onchange = async () => {
  const f = $("#logoIn").files[0]; $("#logoIn").value = ""; if (!f || !logoTarget) return;
  const url = URL.createObjectURL(f);
  try { state.logos[logoTarget] = await imageToPng(url); if (logoTarget === "bidder") saveBidder(); renderLogos(); }
  catch (e) { toast(e.message); }
  URL.revokeObjectURL(url);
};
let savedBidder = { name: "", logo: "" };
try { savedBidder = Object.assign(savedBidder, JSON.parse(localStorage.getItem(BIDDER_KEY) || "{}")); } catch (e) {}
state.bidder = savedBidder.name || "";
$("#pBidder").value = state.bidder;
$("#pBidder").addEventListener("input", () => { state.bidder = $("#pBidder").value.trim(); saveBidder(); updatePreview(); updateGate(); });
if (savedBidder.logo) imageToPng(savedBidder.logo).then((l) => { state.logos.bidder = l; renderLogos(); }).catch(() => {});

/* ================= 排版設定 ================= */
const LS_KEY = "proposalgen.layout.v1";
const DEFAULT_LAYOUT = {
  styles: {
    h1: { zh: "標楷體", en: "Times New Roman", size: 18, bold: true, align: "center", lineType: "multiple", lineValue: 1, before: 12, after: 12, indent: 0 },
    h2: { zh: "標楷體", en: "Times New Roman", size: 16, bold: true, align: "left", lineType: "multiple", lineValue: 1, before: 12, after: 6, indent: 0 },
    h3: { zh: "標楷體", en: "Times New Roman", size: 14, bold: true, align: "left", lineType: "multiple", lineValue: 1, before: 6, after: 6, indent: 0 },
    body: { zh: "標楷體", en: "Times New Roman", size: 13, bold: false, align: "justify", lineType: "exact", lineValue: 25, before: 0, after: 6, indent: 2 },
    caption: { zh: "標楷體", en: "Times New Roman", size: 12, bold: true, align: "center", lineType: "exact", lineValue: 20, before: 6, after: 3, indent: 0 },
    table: { zh: "標楷體", en: "Times New Roman", size: 11, bold: false, align: "left", lineType: "multiple", lineValue: 1, before: 0, after: 0, indent: 0 },
  },
  page: { top: 2.54, bottom: 2.54, left: 3.17, right: 3.17, header: true, pageNum: true, footerBidder: true, color: "2E7D32" },
};
const STYLE_TABS = [["h1", "標題 1"], ["h2", "標題 2"], ["h3", "標題 3"], ["body", "內文"], ["caption", "圖表標題"], ["table", "表格文字"], ["page", "頁面"]];
const ZH_FONTS = ["標楷體", "新細明體", "細明體", "微軟正黑體", "華康中黑體", "華康楷書體", "思源黑體", "思源宋體", "Noto Sans TC", "Noto Serif TC"];
const EN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond", "Verdana", "Tahoma"];
const CSS_FONT = {
  "標楷體": '"標楷體","DFKai-SB","BiauKai","Kaiti TC","STKaiti","KaiTi"', "華康楷書體": '"DFKai-SB","標楷體","Kaiti TC","STKaiti"',
  "新細明體": '"PMingLiU","新細明體","Songti TC","Noto Serif TC"', "細明體": '"MingLiU","細明體","Songti TC","Noto Serif TC"',
  "微軟正黑體": '"Microsoft JhengHei","微軟正黑體","PingFang TC","Noto Sans TC"', "華康中黑體": '"DFHei","Microsoft JhengHei","PingFang TC","Noto Sans TC"',
  "思源黑體": '"Noto Sans TC","Source Han Sans TC","PingFang TC"', "思源宋體": '"Noto Serif TC","Source Han Serif TC","Songti TC"',
};
const SIZE_NAMES = { 42: "初號", 36: "小初", 26: "一號", 24: "小一", 22: "二號", 18: "小二", 16: "三號", 15: "小三", 14: "四號", 12: "小四", 10.5: "五號", 9: "小五" };
const clone = (o) => JSON.parse(JSON.stringify(o));
let layout = clone(DEFAULT_LAYOUT), savedJSON = null, activeTab = "h1";
try { const s = localStorage.getItem(LS_KEY); if (s) { const o = JSON.parse(s); if (o && o.styles && o.page) { layout = Object.assign(clone(DEFAULT_LAYOUT), o); for (const k in DEFAULT_LAYOUT.styles) layout.styles[k] = Object.assign({}, DEFAULT_LAYOUT.styles[k], o.styles[k] || {}); layout.page = Object.assign({}, DEFAULT_LAYOUT.page, o.page); savedJSON = JSON.stringify(layout); } } } catch (e) {}

function renderTabs() {
  const t = $("#tabs"); t.textContent = "";
  STYLE_TABS.forEach(([k, label]) => t.append(el("button", { role: "tab", "aria-selected": String(k === activeTab), onclick: () => { activeTab = k; renderTabs(); renderForm(); updatePreview(); } }, label)));
}
function fontInput(id, list, val, on) {
  const dl = el("datalist", { id: id + "L" }, list.map((f) => el("option", { value: f })));
  const i = el("input", { class: "inp", id, list: id + "L", value: val, oninput: (e) => on(e.target.value.trim()) });
  return [i, dl];
}
function numInput(val, on, attrs = {}) { return el("input", Object.assign({ class: "inp", type: "number", value: val, oninput: (e) => { const v = parseFloat(e.target.value); if (!isNaN(v)) on(v); } }, attrs)); }
function seg(options, cur, on) {
  return el("div", { class: "seg", role: "group" }, options.map(([v, l]) => el("button", { type: "button", "aria-pressed": String(cur === v), onclick: (e) => { on(v); [...e.currentTarget.parentNode.children].forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget))); } }, l)));
}
function renderForm() {
  const f = $("#styleForm"); f.textContent = "";
  const changed = () => { updatePreview(); updateSaveStatus(); };
  if (activeTab === "page") {
    const p = layout.page;
    f.append(el("div", { class: "grid2" },
      ["top", "bottom", "left", "right"].map((k) => el("label", { class: "f" }, { top: "上邊界 (cm)", bottom: "下邊界 (cm)", left: "左邊界 (cm)（裝訂側）", right: "右邊界 (cm)" }[k], numInput(p[k], (v) => { p[k] = v; changed(); }, { step: "0.1", min: "0.5", max: "6" })))),
      el("div", { class: "row", style: "margin-top:14px;gap:22px" },
        el("label", { class: "toggle" }, el("input", { type: "checkbox", checked: p.header, onchange: (e) => { p.header = e.target.checked; changed(); } }), "頁首：機關 LOGO＋標案名稱／章名"),
        el("label", { class: "toggle" }, el("input", { type: "checkbox", checked: p.footerBidder, onchange: (e) => { p.footerBidder = e.target.checked; changed(); } }), "頁尾：投標廠商 LOGO＋名稱"),
        el("label", { class: "toggle" }, el("input", { type: "checkbox", checked: p.pageNum, onchange: (e) => { p.pageNum = e.target.checked; changed(); } }), "頁碼（目錄 I、II；本文 1、2）")),
      el("label", { class: "f", style: "margin-top:14px;max-width:260px" }, "頁首文字與線條顏色",
        el("div", { class: "row" }, el("input", { type: "color", value: "#" + p.color, "aria-label": "頁首顏色", style: "width:48px;height:34px;border:1px solid var(--line);border-radius:4px;background:var(--field);padding:2px", oninput: (e) => { p.color = e.target.value.replace("#", "").toUpperCase(); changed(); } }), el("span", { class: "hint" }, "預設為綠色 #2E7D32"))),
      el("p", { class: "hint" }, "紙張固定為直式 A4，符合評選須知「直式 A4、左側裝訂」之要求。"));
    return;
  }
  const s = layout.styles[activeTab];
  const lineValCtl = numInput(s.lineValue, (v) => { s.lineValue = v; changed(); }, { step: s.lineType === "multiple" ? "0.1" : "1", min: "0.5" });
  const lineUnit = el("span", { class: "hint" });
  const syncLine = () => { lineValCtl.disabled = false; lineUnit.textContent = s.lineType === "multiple" ? "倍" : "pt"; lineValCtl.step = s.lineType === "multiple" ? "0.1" : "1"; };
  const lineSel = el("select", { class: "inp", onchange: (e) => {
    const v = e.target.value;
    if (v === "single") { s.lineType = "multiple"; s.lineValue = 1; } else if (v === "onehalf") { s.lineType = "multiple"; s.lineValue = 1.5; } else if (v === "double") { s.lineType = "multiple"; s.lineValue = 2; }
    else { s.lineType = v; if (v !== "multiple") s.lineValue = Math.max(12, Math.round(s.size * 1.6)); }
    lineValCtl.value = s.lineValue; syncLine(); changed();
  } }, [["single", "單行間距"], ["onehalf", "1.5 倍行高"], ["double", "2 倍行高"], ["multiple", "多倍行高"], ["atLeast", "最小行高"], ["exact", "固定行高"]].map(([v, l]) => el("option", { value: v }, l)));
  lineSel.value = s.lineType === "multiple" ? (s.lineValue === 1 ? "single" : s.lineValue === 1.5 ? "onehalf" : s.lineValue === 2 ? "double" : "multiple") : s.lineType;
  syncLine();
  const sizeName = el("span", { class: "hint" }, SIZE_NAMES[s.size] ? `（${SIZE_NAMES[s.size]}）` : "");
  f.append(el("div", { class: "grid2" },
    el("label", { class: "f" }, "中文字型", ...fontInput("fz", ZH_FONTS, s.zh, (v) => { s.zh = v; changed(); })),
    el("label", { class: "f" }, "英文／數字字型", ...fontInput("fe", EN_FONTS, s.en, (v) => { s.en = v; changed(); })),
  ));
  f.append(el("div", { class: "grid3", style: "margin-top:12px" },
    el("label", { class: "f" }, el("span", {}, "字級 (pt) ", sizeName), numInput(s.size, (v) => { s.size = v; sizeName.textContent = SIZE_NAMES[v] ? `（${SIZE_NAMES[v]}）` : ""; changed(); }, { step: "0.5", min: "6", max: "48" })),
    el("label", { class: "f" }, "行距", lineSel),
    el("label", { class: "f" }, el("span", {}, "行距值 ", lineUnit), lineValCtl),
    el("label", { class: "f" }, "與前段距離 (pt)", numInput(s.before, (v) => { s.before = v; changed(); }, { step: "1", min: "0" })),
    el("label", { class: "f" }, "與後段距離 (pt)", numInput(s.after, (v) => { s.after = v; changed(); }, { step: "1", min: "0" })),
    el("label", { class: "f" }, "首行縮排 (字元)", numInput(s.indent, (v) => { s.indent = v; changed(); }, { step: "0.5", min: "0", max: "6" })),
  ));
  f.append(el("div", { class: "row", style: "margin-top:14px;gap:18px" },
    el("label", { class: "toggle" }, el("input", { type: "checkbox", checked: s.bold, onchange: (e) => { s.bold = e.target.checked; changed(); } }), "粗體"),
    el("span", { class: "small muted" }, "對齊"),
    seg([["left", "靠左"], ["center", "置中"], ["justify", "左右對齊"], ["right", "靠右"]], s.align, (v) => { s.align = v; changed(); })));
}
function cssFont(s) { const z = CSS_FONT[s.zh] || `"${s.zh}"`; return `"${s.en}", ${z}, serif`; }
function updatePreview() {
  const k = 1.9, ptpx = 0.3528 * k; // 1pt → px at preview scale
  const sh = $("#sheet"), p = layout.page;
  sh.style.setProperty("--mt", p.top); sh.style.setProperty("--mb", p.bottom); sh.style.setProperty("--ml", p.left); sh.style.setProperty("--mr", p.right);
  sh.style.setProperty("--theme", "#" + (p.color || "2E7D32"));
  $("#pvHdr").style.display = p.header ? "" : "none";
  $("#pvProj").textContent = ($("#pName") && $("#pName").value) || "計畫名稱";
  const al = state.logos.agency, bl = state.logos.bidder, showB = p.footerBidder !== false;
  $("#pvAgencyLogo").hidden = !al; if (al) $("#pvAgencyLogo").src = al.url;
  $("#pvAgencyPh").hidden = !!al;
  $("#pvBidderLogo").hidden = !(bl && showB); if (bl) $("#pvBidderLogo").src = bl.url;
  $("#pvBidderPh").hidden = !showB || !!bl;
  $("#pvBidder").textContent = showB ? (state.bidder || "投標廠商名稱") : "";
  $("#pvBidder").style.color = state.bidder ? "" : "#9aa3ab";
  $("#pvBidder").style.marginLeft = bl ? "auto" : "";
  $("#pvPn").hidden = !p.pageNum;
  const hasB = showB;
  $("#pvFtr").style.display = hasB || p.pageNum ? "" : "none";
  $("#pvFtr").style.borderTopColor = hasB ? "" : "transparent";
  sh.querySelectorAll("[data-k]").forEach((e) => {
    const s = layout.styles[e.dataset.k];
    e.style.fontFamily = cssFont(s); e.style.fontSize = (s.size * ptpx).toFixed(2) + "px"; e.style.fontWeight = s.bold ? "700" : "400";
    if (e.tagName === "P") {
      if (!e.style.textAlign || e.dataset.k !== "table") e.style.textAlign = s.align === "justify" ? "justify" : s.align;
      e.style.lineHeight = s.lineType === "multiple" ? String(1.3 * s.lineValue) : s.lineType === "exact" ? (s.lineValue * ptpx).toFixed(2) + "px" : `max(1.3, ${(s.lineValue * ptpx).toFixed(2)}px)`;
      e.style.marginTop = (s.before * ptpx).toFixed(2) + "px"; e.style.marginBottom = (s.after * ptpx).toFixed(2) + "px";
      e.style.textIndent = s.indent ? s.indent + "em" : "0";
    }
    e.classList.toggle("pv-active", e.dataset.k === activeTab);
  });
  $("#pvNote").textContent = activeTab === "page" ? `邊界 ${p.top}/${p.bottom}/${p.left}/${p.right} cm` : `${STYLE_TABS.find((t) => t[0] === activeTab)[1]}：${layout.styles[activeTab].zh}／${layout.styles[activeTab].en} ${layout.styles[activeTab].size}pt`;
}
function isSaved() { return savedJSON === JSON.stringify(layout); }
function updateSaveStatus() {
  const s = $("#saveStatus");
  if (isSaved()) { s.className = "status ok"; s.textContent = "已儲存"; $("#s3").classList.add("done"); }
  else { s.className = "status warn"; s.textContent = savedJSON ? "有未儲存的變更" : "尚未儲存"; $("#s3").classList.remove("done"); }
  updateGate();
}
$("#saveLayout").onclick = () => {
  savedJSON = JSON.stringify(layout);
  try { localStorage.setItem(LS_KEY, savedJSON); toast("排版設定已儲存"); } catch (e) { toast("已儲存於本次開啟期間（瀏覽器未允許長期保存）"); }
  updateSaveStatus();
};
$("#resetLayout").onclick = () => { layout = clone(DEFAULT_LAYOUT); renderForm(); updatePreview(); updateSaveStatus(); toast("已還原預設值，記得儲存"); };

/* ================= 上傳與解析 ================= */
const drop = $("#drop"), fileIn = $("#fileIn");
drop.onclick = () => fileIn.click();
drop.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileIn.click(); } };
["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));
fileIn.onchange = () => { addFiles(fileIn.files); fileIn.value = ""; };

let pdfReady = null;
function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("load " + src)); document.head.append(s); }); }
function ensurePdf() {
  if (!pdfReady) pdfReady = loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js")
    .then(() => loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"))
    .then(() => { window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js"; });
  return pdfReady;
}
async function pdfText(buf) {
  await ensurePdf();
  const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
  const out = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const pg = await pdf.getPage(i); const tc = await pg.getTextContent();
    const rows = new Map();
    tc.items.forEach((it) => { const y = Math.round(it.transform[5] / 3) * 3; if (!rows.has(y)) rows.set(y, []); rows.get(y).push(it); });
    [...rows.keys()].sort((a, b) => b - a).forEach((y) => out.push(rows.get(y).sort((a, b) => a.transform[4] - b.transform[4]).map((x) => x.str).join("")));
  }
  return out.join("\n");
}
async function extractFile(f) {
  const ext = (f.name.split(".").pop() || "").toLowerCase();
  if (ext === "doc") throw new Error("舊版 .doc 無法在瀏覽器讀取，請以 Word 另存為 .docx 或 .odt 後再上傳");
  const buf = await f.arrayBuffer();
  if (ext === "odt") { const z = await JSZip.loadAsync(buf); const x = await z.file("content.xml").async("string"); return TenderParse.extractODT(x, DOMParser); }
  if (ext === "docx") { const z = await JSZip.loadAsync(buf); const x = await z.file("word/document.xml").async("string"); return TenderParse.extractDOCX(x, DOMParser); }
  if (ext === "pdf") return TenderParse.extractPlain(await pdfText(buf));
  if (ext === "txt" || ext === "md") return TenderParse.extractPlain(new TextDecoder().decode(buf));
  throw new Error("不支援的格式：." + ext);
}
async function addFiles(list) {
  for (const f of Array.from(list)) {
    if (state.docs.find((d) => d.name === f.name)) continue;
    const d = { name: f.name, kind: "other", parsed: null, error: "", busy: true };
    state.docs.push(d); renderFiles();
    try { d.parsed = await extractFile(f); d.kind = TenderParse.classify(f.name, d.parsed.text); if (!d.parsed.text.trim()) d.error = "讀不到文字（可能是掃描影像 PDF）"; }
    catch (e) { d.error = e.message || "讀取失敗"; }
    d.busy = false; renderFiles();
  }
  reanalyze();
}
function renderFiles() {
  const ul = $("#files"); ul.textContent = "";
  state.docs.forEach((d, i) => {
    const sel = el("select", { "aria-label": "文件類型", onchange: (e) => { d.kind = e.target.value; reanalyze(); } }, Object.entries(TenderParse.KIND_LABEL).map(([k, l]) => el("option", { value: k }, l)));
    sel.value = d.kind; sel.disabled = !!d.error || d.busy;
    ul.append(el("li", {}, el("div", {}, el("div", { class: "fn", title: d.name }, d.name), d.busy ? el("div", { class: "small muted" }, "讀取中…") : d.error ? el("div", { class: "err" }, d.error) : el("div", { class: "small muted" }, `${d.parsed.lines.length} 段・${d.parsed.tables.length} 個表格`)),
      sel, el("button", { class: "icon-btn", "aria-label": "移除 " + d.name, onclick: () => { state.docs.splice(i, 1); renderFiles(); reanalyze(); } }, "✕")));
  });
  $("#s1").classList.toggle("done", state.docs.some((d) => d.parsed && !d.error));
}
function monthStr(y, m) { return `${y}-${String(m).padStart(2, "0")}`; }
function reanalyze() {
  const ok = state.docs.filter((d) => d.parsed && !d.error);
  if (!ok.length) { state.P = null; renderP(); return; }
  const prev = state.P;
  const P = TenderParse.analyze(ok);
  P.workItems.forEach((w) => { w.auto = true; w.subs.forEach((s) => (s.auto = true)); });
  const now = new Date();
  P.rocYear = now.getFullYear() - 1911; P.rocMonth = now.getMonth() + 1;
  if (P.endRoc) { P.endMonth = monthStr(P.endRoc.y + 1911, P.endRoc.m); P.startMonth = monthStr(P.endRoc.y + 1911, 1); }
  else { P.startMonth = monthStr(now.getFullYear(), now.getMonth() + 1); P.endMonth = monthStr(now.getFullYear() + (now.getMonth() + 11 >= 12 ? 1 : 0), ((now.getMonth() + 11) % 12) + 1); }
  state.P = P; renderP();
}
const bind = (id, key, num) => { const i = $(id); i.value = state.P[key] == null ? "" : state.P[key]; i.oninput = () => { state.P[key] = num ? (i.value.replace(/[^\d]/g, "") ? +i.value.replace(/[^\d]/g, "") : null) : i.value; if (id === "#pName") updatePreview(); }; };
function renderP() {
  const P = state.P;
  $("#s2empty").hidden = !!P; $("#s2body").hidden = !P; $("#s2").classList.toggle("done", !!(P && P.projectName && P.workItems.length));
  if (!P) { updateGate(); return; }
  bind("#pName", "projectName"); bind("#pAgency", "agency"); bind("#pShort", "agencyShort");
  bind("#pBudget", "budget", true); bind("#pPeriod", "period"); bind("#pStart", "startMonth"); bind("#pEnd", "endMonth"); bind("#pRocY", "rocYear", true); bind("#pRocM", "rocMonth", true);
  const c = $("#chips"); c.textContent = "";
  const chip = (l, n) => c.append(el("span", { class: "chip" }, l, " ", el("b", {}, String(n))));
  chip("查核點", P.milestones.length); chip("成果交付", P.deliverables.length); chip("評選項目", P.evaluation.length); chip("標價項目", P.priceItems.length); chip("參考文獻", P.references.length);
  renderWI(); renderLogos(); updateGate();
}
function renderWI() {
  const box = $("#wiList"); box.textContent = ""; const P = state.P;
  if (!P.workItems.length) box.append(el("div", { class: "empty" }, "沒有自動找到工作項目。請手動新增，或使用「用 AI 重新解析」。"));
  P.workItems.forEach((w, i) => {
    const subs = el("div", { class: "wi-subs" });
    w.subs.forEach((s, k) => subs.append(el("div", {},
      el("div", { class: "wi-row" }, el("span", { class: "no" }, `${k + 1}.`), el("input", { class: "inp", value: s.title, "aria-label": `子項 ${k + 1} 標題`, title: s.req || "", oninput: (e) => { s.title = e.target.value; s.auto = false; } }),
        el("button", { class: "icon-btn", "aria-label": "刪除子項", onclick: () => { w.subs.splice(k, 1); renderWI(); } }, "✕")))));
    subs.append(el("button", { class: "link-btn", onclick: () => { w.subs.push({ title: "", req: "", auto: false }); renderWI(); } }, "＋ 子項"));
    box.append(el("div", { class: "wi" },
      el("div", { class: "wi-row" }, el("span", { class: "no" }, `${CNN[i + 1]}、`), el("input", { class: "inp", value: w.title, "aria-label": `工作項目 ${i + 1} 標題`, oninput: (e) => { w.title = e.target.value; w.auto = false; updateGate(); } }),
        el("button", { class: "icon-btn", "aria-label": "刪除工作項目", onclick: () => { P.workItems.splice(i, 1); renderWI(); updateGate(); } }, "✕")),
      w.req && w.req !== w.title && !w.subs.length ? el("div", { class: "req" }, w.req) : null, subs));
  });
}
$("#addWi").onclick = () => { if (!state.P) return; state.P.workItems.push({ title: "新工作項目", req: "", subs: [], auto: false }); renderWI(); updateGate(); };

/* ================= Claude（選用） ================= */
const RULES = (P) => `你是臺灣政府採購「技術服務建議書」的資深撰稿顧問。請以臺灣正體中文、專業且精練的公文語氣撰寫。投標廠商自稱「本團隊」，稱招標機關為「${P.agencyShort || "貴機關"}」（不要使用「本署」「本局」）。只根據提供的招標文件內容與一般專業知識撰寫；凡涉及投標廠商的人名、實績、證照、設備、數量、金額或日期等未提供的事實，一律寫成【待補：簡短說明】，不得捏造。不要使用 Markdown 符號，字串內不要換行。`;
const brief = (P) => JSON.stringify({ 計畫名稱: P.projectName, 招標機關: P.agency, 預算: P.budget, 履約期限: P.period, 工作項目: P.workItems.map((w) => ({ 標題: w.title, 子項: w.subs.map((s) => s.title) })), 評選項目: P.evaluation.map((e) => `${e.item}(${e.weight}分)`), 分段進度: P.milestones.map((m) => m.stage) });
function srcOf(kinds, max) { return state.docs.filter((d) => d.parsed && !d.error && kinds.includes(d.kind)).map((d) => `【${TenderParse.KIND_LABEL[d.kind]}】\n` + d.parsed.text).join("\n\n").slice(0, max); }
function keyLines(kinds, re, max) { return state.docs.filter((d) => d.parsed && kinds.includes(d.kind)).flatMap((d) => d.parsed.lines.map((l) => l.text).filter((t) => re.test(t))).join("\n").slice(0, max); }
const arr = (x) => (Array.isArray(x) ? x.filter((v) => v != null && v !== "").map((v) => (typeof v === "string" ? v : v)) : []);
const strs = (x) => arr(x).filter((v) => typeof v === "string");

async function initClaude() {
  const hint = $("#aiHint"), use = $("#useAI");
  try { state.downloads = window.claude ? await window.claude.use("downloads") : null; } catch (e) { state.downloads = null; }
  let s = null;
  try { s = window.claude ? await window.claude.use("sample") : null; } catch (e) { s = null; }
  state.sample = s;
  if (s) { use.checked = true; hint.textContent = "會使用你自己的 Claude 額度，約需 3～8 分鐘；關閉則產生章節架構與【待補】標記。"; $("#aiParse").hidden = false; }
  else { use.checked = false; use.disabled = true; hint.textContent = "此處無法使用 Claude 撰寫，將產生完整章節架構、表格與【待補】標記。"; }
}
function aiError(e) {
  const c = e && e.code;
  if (c === "not_granted" || c === "sampling_disabled" || c === "capability_disabled" || c === "not_declared") { state.sample = null; $("#useAI").checked = false; $("#useAI").disabled = true; $("#aiParse").hidden = true; $("#aiHint").textContent = "未允許使用 Claude，改為產生架構與【待補】標記。"; return "未允許使用 Claude"; }
  if (c === "rate_limited") return "使用量已達上限，請稍後再試";
  if (c === "cancelled") return "已停止";
  if (c === "invalid_json") return "回覆格式不完整";
  return "暫時無法取得回覆";
}

$("#aiParse").onclick = async () => {
  if (!state.sample || !state.P) return;
  const b = $("#aiParse"); b.disabled = true; b.textContent = "解析中…";
  const P = state.P;
  try {
    const src = srcOf(["spec", "announce", "other"], 60000) + "\n\n" + srcOf(["eval", "price"], 12000) + "\n\n" + keyLines(["notice"], /(標案名稱|預算|履約|期限|機關名稱)/, 3000);
    const r = await state.sample.json(`${RULES(P)}\n\n任務：從下列招標文件擷取結構化資訊。\n\n<<<\n${src}\n>>>\n\n只回覆一個 JSON 物件：{"projectName":"","agency":"","budget":數字或null,"period":"履約期限原文","endRoc":{"y":民國年,"m":月,"d":日}或null,"background":"計畫需求或緣起原文","workItems":[{"title":"工作項目短標題(≤20字)","req":"需求原文","subs":[{"title":"子項短標題(≤22字)","req":"子項需求原文"}]}],"milestones":[{"stage":"查核點名稱","tasks":["應完成工作"]}],"deliverables":[{"name":"","desc":""}],"evaluation":[{"item":"","weight":數字,"desc":""}]}`);
    if (r && Array.isArray(r.workItems) && r.workItems.length) {
      ["projectName", "agency", "period", "background"].forEach((k) => { if (r[k]) P[k] = r[k]; });
      if (r.budget) P.budget = +r.budget;
      if (r.endRoc && r.endRoc.y) { P.endRoc = r.endRoc; P.endMonth = monthStr(+r.endRoc.y + 1911, +r.endRoc.m || 12); }
      P.workItems = r.workItems.map((w) => ({ title: String(w.title || ""), req: String(w.req || w.title || ""), auto: true, subs: arr(w.subs).map((s) => ({ title: String(s.title || ""), req: String(s.req || s.title || ""), auto: true })) }));
      if (arr(r.milestones).length) P.milestones = r.milestones.map((m) => ({ stage: String(m.stage || ""), tasks: strs(m.tasks), when: "" }));
      if (arr(r.deliverables).length) P.deliverables = r.deliverables.map((d) => ({ name: String(d.name || ""), desc: String(d.desc || "") }));
      if (arr(r.evaluation).length) { const keep = P.evaluation; P.evaluation = r.evaluation.map((e) => ({ item: String(e.item || ""), weight: +e.weight || "", desc: String(e.desc || ""), chapters: (keep.find((k) => k.item === e.item) || {}).chapters || "【待補】" })); }
      renderP(); toast("已依 AI 解析結果更新，請再確認");
    } else toast("AI 未找到工作項目，保留原解析結果");
  } catch (e) { toast(aiError(e)); }
  b.disabled = false; b.textContent = "用 AI 重新解析";
};

/* ================= 開始條件 ================= */
function updateGate() {
  const P = state.P;
  const conds = [
    [!!(P && state.docs.some((d) => d.parsed && !d.error)), "上傳並解析招標文件"],
    [!!(P && P.workItems.length), "至少一個工作項目"],
    [isSaved(), "儲存排版設定"],
    [$("#confirm").checked, "勾選確認"],
  ];
  const n = $("#need"); n.textContent = "";
  conds.forEach(([ok, l]) => n.append(el("li", { class: ok ? "met" : "" }, l)));
  $("#start").disabled = !conds.every((c) => c[0]) || !!state.running;
}
$("#confirm").onchange = updateGate;

/* ================= 產生 ================= */
const logList = $("#logList");
function logStep(text) { const li = el("li", { class: "s-run" }, text); logList.append(li); return { done: (t) => { li.className = "s-done"; if (t) li.textContent = t; }, fail: (t) => { li.className = "s-fail"; li.textContent = t; } }; }
async function pool(tasks, n) { const res = []; let i = 0; const run = async () => { while (i < tasks.length) { const k = i++; res[k] = await tasks[k](); } }; await Promise.all(Array.from({ length: Math.min(n, tasks.length) }, run)); return res; }

function monthsOf(P) {
  const [sy, sm] = (P.startMonth || "").split("-").map(Number), [ey, em] = (P.endMonth || "").split("-").map(Number);
  if (!sy || !ey) return [];
  const out = []; let y = sy, m = sm;
  while ((y < ey || (y === ey && m <= em)) && out.length < 24) { out.push({ y: y - 1911, m }); m++; if (m > 12) { m = 1; y++; } }
  return out;
}

async function draft(P, months, signal) {
  const S = state.sample, C = { work: [] }, rules = RULES(P), b = brief(P);
  const spec = srcOf(["spec", "announce", "other"], 30000);
  const tasks = [];
  // A：前言、瞭解、成果效益
  tasks.push(async () => {
    const st = logStep("撰寫第壹章前言、第貳章規劃概述、第參章成果效益…");
    try {
      const r = await S.json(`${rules}\n\n任務：撰寫服務建議書的第壹章（前言）、第貳章前段（對本計畫及問題之瞭解、課題與對策、整體流程、創新性與可行性、建議）及第參章（預期成果與效益）。\n\n計畫資訊：${b}\n\n招標文件摘錄：\n<<<\n${spec}\n>>>\n\n只回覆一個 JSON 物件，格式如下：\n{"origin":["計畫緣起與目的，3-4段"],"objectives":["計畫目標，4-6項"],"advantages":["本團隊承攬優勢，2段，具體事實以【待補】標示"],"understanding":["對計畫及問題之瞭解，3段"],"issues":[{"issue":"課題","strategy":"對策"}],"planIntro":["說明整體執行構想，1段"],"mainFlow":["整體執行流程步驟，6-9步，每步≤18字"],"innovation":["創新作法，2-3段"],"feasibility":[{"aspect":"技術面","desc":""},{"aspect":"人力面","desc":""},{"aspect":"時程面","desc":""},{"aspect":"資料面","desc":""}],"suggestions":["對委託工作之建議，3-5項"],"outcomes":[{"item":"成果項目","desc":"量化內容"}],"benefits":["預期效益，2段"],"addedValue":["附加價值，3-5項"],"workTitles":[{"title":"工作項目短標題≤20字","subs":["子項短標題≤22字"]}],"references":[{"category":"法規與規定","text":"發行機關，名稱，西元年。"}]}\nworkTitles 的數量、順序及各子項數量須與計畫資訊中的工作項目完全一致。references 只列招標文件明確提到的法規、作業規定或前期計畫，不要自行新增文獻。`, { signal });
      Object.assign(C, { origin: strs(r.origin), objectives: strs(r.objectives), advantages: strs(r.advantages), understanding: strs(r.understanding), issues: arr(r.issues).filter((x) => x.issue), planIntro: strs(r.planIntro), mainFlow: strs(r.mainFlow), innovation: strs(r.innovation), feasibility: arr(r.feasibility).filter((x) => x.aspect), suggestions: strs(r.suggestions), outcomes: arr(r.outcomes).filter((x) => x.item), benefits: strs(r.benefits), addedValue: strs(r.addedValue) });
      ["origin", "objectives", "advantages", "understanding", "issues", "planIntro", "mainFlow", "innovation", "feasibility", "suggestions", "outcomes", "benefits", "addedValue"].forEach((k) => { if (!C[k] || !C[k].length) delete C[k]; });
      if (Array.isArray(r.workTitles)) r.workTitles.forEach((wt, i) => { const w = P.workItems[i]; if (!w || !wt) return; if (w.auto && wt.title) w.title = String(wt.title); arr(wt.subs).forEach((t, k) => { if (w.subs[k] && w.subs[k].auto && typeof t === "string" && t) w.subs[k].title = t; }); });
      if (arr(r.references).length) { const seen = new Set(P.references.map((x) => x.text)); r.references.forEach((x) => { if (x && x.text && !seen.has(x.text)) { P.references.push({ category: String(x.category || "法規與規定"), text: String(x.text) }); seen.add(x.text); } }); }
      st.done("第壹章、第貳章概述、第參章完成");
    } catch (e) { st.fail("前言與成果章節：" + aiError(e) + "，改用待補標記"); if (e && e.code === "cancelled") throw e; }
  });
  // B：各工作項目
  P.workItems.forEach((w, i) => tasks.push(async () => {
    const st = logStep(`撰寫工作項目${CNN[i + 1]}「${w.title}」執行方法…`);
    const n = w.subs.length;
    const subsTxt = n ? w.subs.map((s, k) => `${k + 1}. ${s.title}：${s.req || s.title}`).join("\n") : "（本項無子項）";
    try {
      const r = await S.json(`${rules}\n\n任務：撰寫服務建議書第貳章中「${w.title}」一節的執行方法。\n\n計畫名稱：${P.projectName}\n本項需求原文：${w.req || w.title}\n子項（依序）：\n${subsTxt}\n\n計畫背景摘要：${(P.background || "").slice(0, 1500)}\n\n只回覆一個 JSON 物件：{"overview":["本項工作概述與執行構想，1-2段"],"flow":["本項作業流程步驟，5-8步，每步≤16字"],${n ? `"subs":[{"blocks":[內容區塊]}]` : `"blocks":[內容區塊]`}}\n${n ? `subs 必須剛好 ${n} 個，順序對應上列子項；每個子項 3-6 個區塊。` : "blocks 3-8 個區塊。"}\n內容區塊可用：{"t":"p","text":"段落"}、{"t":"list","items":["條列"]}、{"t":"table","caption":"表名","headers":["欄位",最多5欄],"rows":[["儲存格"],最多8列]}、{"t":"figure","caption":"圖名","steps":["步驟",3-7步]}（流程圖）、{"t":"figure","caption":"圖名","desc":"建議放置的照片或地圖內容"}（預留圖位）。\n每個子項至少一段具體說明執行方法、作業步驟、使用資料與工具及品質確保；本節合計建議 1-3 個表格、1-2 個圖。`, { signal });
      C.work[i] = { overview: strs(r.overview), flow: strs(r.flow), subs: arr(r.subs).map((s) => ({ blocks: arr(s && s.blocks) })), blocks: arr(r.blocks) };
      if (!C.work[i].overview.length) delete C.work[i].overview;
      st.done(`工作項目${CNN[i + 1]}完成`);
    } catch (e) { st.fail(`工作項目${CNN[i + 1]}：` + aiError(e) + "，改用待補標記"); if (e && e.code === "cancelled") throw e; }
  }));
  // C：品質進度、組織、CSR、經費說明
  tasks.push(async () => {
    const st = logStep("撰寫品質及進度控管、專案組織、CSR 與經費說明…");
    const ml = months.map((m) => `${m.y}年${m.m}月`);
    const contract = keyLines(["contract", "spec"], /(資安|資通|個人資料|個資|保密|保險|逾期|驗收|履約期限|工作會議|駐點)/, 6000);
    try {
      const r = await S.json(`${rules}\n\n任務：撰寫服務建議書的專案品質及進度控管、專案管理及履約能力、企業社會責任(CSR)與經費說明。\n\n計畫資訊：${b}\n評選須知內容要領：${(P.outlineHints || []).join("；").slice(0, 3000)}\n分段進度：${JSON.stringify(P.milestones.map((m) => ({ 查核點: m.stage, 工作: m.tasks })))}\n契約與規範重點摘錄：\n${contract}\n進度表月份（索引從 0 起）：${JSON.stringify(ml)}\n\n只回覆一個 JSON 物件：{"schedule":["進度安排說明，1段"],"gantt":[{"span":[起,迄],"subs":[[起,迄]]}],"milestoneWhen":["各查核點預定時間"],"qc":["品質管理機制，2段"],"pdca":{"plan":"","do":"","check":"","act":""},"risk":[{"risk":"潛在風險","response":"因應對策"}],"infosec":["資訊安全及個資保護，2段"],"org":["專案組織說明，1-2段"],"groups":[{"name":"組名≤12字","duty":"職掌"}],"onTime":["如期履約保證措施，2段"],"csr":{"salary":["員工薪資承諾說明，1段"],"balance":["工作與生活平衡措施，1段"],"green":["綠色採購說明，1段"]},"budgetNote":["經費編列說明，1段"]}\ngantt 須與工作項目數量、順序一致，subs 數量須與各項子項數一致，月份索引介於 0 與 ${Math.max(0, ml.length - 1)}；milestoneWhen 須與分段進度數量一致；groups 3-6 組。`, { signal });
      const nM = months.length || 12, okSpan = (s) => Array.isArray(s) && s.length === 2 && s.every((x) => Number.isInteger(x) && x >= 0 && x < nM) && s[0] <= s[1];
      arr(r.gantt).forEach((g, i) => { const w = P.workItems[i]; if (!w || !g) return; if (okSpan(g.span)) w.span = g.span; arr(g.subs).forEach((ss, k) => { if (w.subs[k] && okSpan(ss)) w.subs[k].span = ss; }); });
      strs(r.milestoneWhen).forEach((t, i) => { if (P.milestones[i] && !P.milestones[i].when) P.milestones[i].when = t; });
      Object.assign(C, { schedule: strs(r.schedule), qc: strs(r.qc), risk: arr(r.risk).filter((x) => x.risk), infosec: strs(r.infosec), org: strs(r.org), groups: arr(r.groups).filter((g) => g.name).slice(0, 6), onTime: strs(r.onTime), budgetNote: strs(r.budgetNote) });
      if (r.pdca && r.pdca.plan) C.pdca = r.pdca;
      if (r.csr) C.csr = { salary: strs(r.csr.salary), balance: strs(r.csr.balance), green: strs(r.csr.green) };
      ["schedule", "qc", "risk", "infosec", "org", "groups", "onTime", "budgetNote"].forEach((k) => { if (!C[k] || !C[k].length) delete C[k]; });
      if (C.csr) ["salary", "balance", "green"].forEach((k) => { if (!C.csr[k].length) delete C.csr[k]; });
      st.done("品質進度、組織、CSR 與經費說明完成");
    } catch (e) { st.fail("管理章節：" + aiError(e) + "，改用待補標記"); if (e && e.code === "cancelled") throw e; }
  });
  await pool(tasks, 2);
  return C;
}

$("#start").onclick = async () => {
  const btn = $("#start");
  btn.classList.remove("stamped"); void btn.offsetWidth; btn.classList.add("stamped");
  state.running = true; updateGate();
  logList.textContent = ""; $("#log").classList.add("on"); $("#result").classList.remove("on");
  const P = JSON.parse(JSON.stringify(state.P));
  const months = monthsOf(P);
  P.months = months;
  P.logos = { agency: state.logos.agency, bidder: state.logos.bidder };
  P.bidder = state.bidder;
  // 查核點位置（甘特圖 ★）
  const n = months.length || 12;
  const ck = P.milestones.map((m) => ({ name: m.stage, idx: /期初|計畫書|第一期/.test(m.stage) ? 0 : /期中|第二期/.test(m.stage) ? Math.floor((n - 1) / 2) : /期末|第三期/.test(m.stage) ? Math.max(0, n - 2) : n - 1 }));
  P.checkpoints = ck.length ? ck : null;
  let C = {};
  if ($("#useAI").checked && state.sample) {
    state.ctl = new AbortController(); $("#stop").hidden = false;
    try { C = await draft(P, months, state.ctl.signal); }
    catch (e) { logStep("已停止 AI 撰寫").fail("已停止 AI 撰寫，其餘章節以【待補】標記產生"); }
    $("#stop").hidden = true;
  }
  P.content = C;
  const st = logStep("排版並產生 Word 檔…");
  try {
    const doc = ProposalGen.buildProposal(docx, P, layout);
    state.blob = await cleanDocx(await docx.Packer.toBlob(doc));
    state.fileName = `${(P.projectName || "服務建議書").replace(/[\\/:*?"<>|]/g, "")}_服務建議書初稿.docx`;
    st.done(`Word 檔完成（${Math.round(state.blob.size / 1024)} KB）`);
    const aiUsed = Object.keys(C).length > 1;
    $("#resultInfo").textContent = `${state.fileName}：${P.workItems.length} 個工作項目、${P.workItems.reduce((a, w) => a + w.subs.length, 0)} 個子項${aiUsed ? "，內文由 Claude 撰寫初稿" : "，內文以【待補】標記預留"}。`;
    $("#result").classList.add("on"); $("#s4").classList.add("done");
    $("#dlNote").textContent = "";
  } catch (e) { console.error(e); st.fail("產生 Word 檔失敗：" + (e && e.message || e)); }
  state.running = false; updateGate();
};
// 移除 docx.js 產生的非標準 highlightCs 標記，並去除 ZIP 資料夾項目，讓 Word／雲端共同編輯都能順利開啟
async function cleanDocx(blob) {
  const z = await JSZip.loadAsync(await blob.arrayBuffer(), { createFolders: false });
  const out = new JSZip();
  for (const name of Object.keys(z.files)) {
    const f = z.files[name]; if (f.dir) continue;
    if (/^word\/(document|styles|numbering|header\d*|footer\d*)\.xml$/.test(name)) out.file(name, (await f.async("string")).replace(/<w:highlightCs\b[^>]*\/>/g, ""), { createFolders: false });
    else out.file(name, await f.async("uint8array"), { createFolders: false });
  }
  return out.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression: "DEFLATE" });
}
$("#stop").onclick = () => { if (state.ctl) state.ctl.abort(); };
$("#dl").onclick = async () => {
  if (!state.blob) return;
  if (state.downloads) {
    try { await state.downloads.save({ filename: state.fileName, data: state.blob }); $("#dlNote").textContent = "已下載"; return; }
    catch (e) { if (e && e.code === "declined") { $("#dlNote").textContent = "已取消下載"; return; } if (e && e.code === "rate_limited") { $("#dlNote").textContent = "請稍候再按一次"; return; } }
  }
  const a = el("a", { href: URL.createObjectURL(state.blob), download: state.fileName }); document.body.append(a); a.click(); a.remove();
};

/* ================= 啟動 ================= */
renderTabs(); renderForm(); updatePreview(); updateSaveStatus(); renderP(); renderLogos();
if (window.claude && window.claude.use) initClaude(); else { $("#useAI").disabled = true; $("#aiHint").textContent = "此處無法使用 Claude 撰寫，將產生完整章節架構、表格與【待補】標記。"; }
