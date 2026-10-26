if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const $ = s => document.querySelector(s);
const IMG = ["jpg","jpeg","png","webp","gif","bmp","avif","svg"];
const MIME = { png: "image/png", jpg: "image/jpeg", webp: "image/webp" };
const OUT = {
  image: ["png","jpg","webp","bmp","pdf","docx"],
  pdf:   ["docx","docx-img","txt","html","png","jpg","webp"],
  docx:  ["pdf","html","txt"],
  sheet: ["csv","json","html","xlsx","docx","pdf","txt"],
  text:  ["pdf","docx","txt","html"],
  video: ["webm","mp4","mkv","mp3","wav","png","jpg","webp"],
  audio: ["mp3","wav","webm","mp4","ogg"]
};
let items = [];

const ext = f => f.name.split(".").pop().toLowerCase();
const base = f => f.name.replace(/\.[^.]+$/, "");

function kind(f) {
  const e = ext(f);
  if (f.type.startsWith("image/") || IMG.includes(e)) return "image";
  if (e === "pdf") return "pdf";
  if (e === "docx") return "docx";
  if (["xlsx","xls","ods"].includes(e)) return "sheet";
  if (f.type.startsWith("video/") || ["mp4","webm","mov","mkv","m4v"].includes(e)) return "video";
  if (f.type.startsWith("audio/") || ["mp3","wav","ogg","m4a","aac"].includes(e)) return "audio";
  if (["txt","md","html","htm","csv","json","log"].includes(e)) return "text";
  return null;
}

const size = n => n > 1e6 ? (n/1e6).toFixed(1)+" Mo" : Math.max(1,Math.round(n/1e3))+" Ko";

/* ---------- Helpers ---------- */
const loadImg = f => new Promise((res, rej) => {
  const i = new Image(); 
  i.onload = () => res(i); 
  i.onerror = () => rej(new Error("L'image est corrompue ou dans un format illisible."));
  i.src = URL.createObjectURL(/\.svg$/i.test(f.name) && f.type !== "image/svg+xml" ? new Blob([f], { type: "image/svg+xml" }) : f);
});

const toBlob = (c, t, q = .92) => new Promise(r => c.toBlob(r, t, q));

function drawCanvas(src, w, h, opaque) {
  const c = document.createElement("canvas"); c.width = w; c.height = h;
  const x = c.getContext("2d");
  if (opaque) { x.fillStyle = "#fff"; x.fillRect(0, 0, w, h); }
  x.drawImage(src, 0, 0, w, h); return c;
}

function textPdf(t) {
  const d = new jspdf.jsPDF(); const L = d.splitTextToSize(t, 180); let y = 15;
  for (const l of L) { if (y > 282) { d.addPage(); y = 15; } d.text(l, 15, y); y += 6; }
  return d.output("blob");
}

function wav(buf) {
  const n = buf.length, ch = buf.numberOfChannels, sr = buf.sampleRate;
  const o = new DataView(new ArrayBuffer(44 + n * ch * 2));
  const w = (p, s) => [...s].forEach((c, i) => o.setUint8(p + i, c.charCodeAt(0)));
  w(0,"RIFF"); o.setUint32(4, 36 + n*ch*2, true); w(8,"WAVEfmt "); o.setUint32(16,16,true);
  o.setUint16(20,1,true); o.setUint16(22,ch,true); o.setUint32(24,sr,true); o.setUint32(28,sr*ch*2,true);
  o.setUint16(32,ch*2,true); o.setUint16(34,16,true); w(36,"data"); o.setUint32(40,n*ch*2,true);
  const d = [...Array(ch)].map((_, i) => buf.getChannelData(i)); let p = 44;
  for (let i = 0; i < n; i++) for (let c = 0; c < ch; c++) {
    const s = Math.max(-1, Math.min(1, d[c][i])); o.setInt16(p, s < 0 ? s*32768 : s*32767, true); p += 2;
  }
  return new Blob([o], { type: "audio/wav" });
}

/* ---------- Conversions ---------- */
const LBL = { "docx-img": "docx (pages en images, non éditable)" };
const MRX = { webm: "webm", mp4: "mp4", mkv: "x-matroska", ogg: "ogg" };
const esc = s => s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const txtBlob = (t, m = "text/plain") => new Blob([t], { type: m });

function clean(html) {
  const d = new DOMParser().parseFromString(html, "text/html");
  d.querySelectorAll("script,iframe,object,embed").forEach(e => e.remove());
  d.querySelectorAll("*").forEach(e => [...e.attributes].forEach(a => /^on/i.test(a.name) && e.removeAttribute(a.name)));
  return d.body.innerHTML;
}

async function htmlPdf(html) {
  const box = document.createElement("div");
  box.style.cssText = "position:fixed;left:-9999px;top:0;width:595px;box-sizing:border-box;padding:24px;background:#fff;color:#000;font:12px sans-serif";
  box.innerHTML = clean(html); document.body.append(box);
  const d = new jspdf.jsPDF({ unit: "pt", format: "a4" });
  try { await new Promise(r => d.html(box, { callback: r, x: 0, y: 0, width: 595, windowWidth: 595, autoPaging: "text", margin: [28, 0, 28, 0], html2canvas: { useCORS: true, scale: 2 } })); }
  finally { box.remove(); }
  return d.output("blob");
}

async function imgBlock(c) {
  const bl = await toBlob(c, "image/jpeg", .92); 
  let w = 660, h = Math.round(660 * c.height / c.width);
  if (h > 950) { w = Math.round(w * 950 / h); h = 950; }
  return { img: new Uint8Array(await bl.arrayBuffer()), w, h };
}

async function makeDocx(blocks) {
  const z = new JSZip(), R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  let body = "", rels = "", n = 0;
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  for (const b of blocks) {
    if (b.br) body += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    else if (b.img) {
      n++; z.file(`word/media/i${n}.jpeg`, b.img);
      rels += `<Relationship Id="rI${n}" Type="${R}/image" Target="media/i${n}.jpeg"/>`;
      const cx = Math.round(b.w * 9525), cy = Math.round(b.h * 9525);
      body += `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${n}" name="i${n}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${n}" name="i${n}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rI${n}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    } else body += `<w:p><w:r><w:t xml:space="preserve">${esc(b.text)}</w:t></w:r></w:p>`;
  }
  z.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function renderPage(pdf, i, scale = 2) {
  const pg = await pdf.getPage(i), vp = pg.getViewport({ scale });
  const c = document.createElement("canvas"); c.width = vp.width; c.height = vp.height;
  const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height);
  await pg.render({ canvasContext: x, viewport: vp }).promise; return c;
}

async function pdfParas(pdf, prog) {
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const tc = await (await pdf.getPage(i)).getTextContent(), lines = [];
    for (const s of tc.items) {
      if (!s.str.trim()) continue;
      const y = s.transform[5]; let l = lines.find(l => Math.abs(l.y - y) < 3);
      if (!l) { l = { y, h: s.height || 10, it: [] }; lines.push(l); }
      l.it.push({ x: s.transform[4], w: s.width || 0, s: s.str });
    }
    lines.sort((a, b) => b.y - a.y);
    const ps = []; let prev = null;
    for (const l of lines) {
      l.it.sort((a, b) => a.x - b.x); let t = "", end = null;
      for (const o of l.it) { if (end !== null && o.x - end > 1) t += " "; t += o.s; end = o.x + o.w; }
      t = t.replace(/\s+/g, " ").trim();
      if (prev && prev.y - l.y < prev.h * 1.8) ps[ps.length - 1] += " " + t; else ps.push(t);
      prev = l;
    }
    pages.push(ps); prog(i / pdf.numPages * .9);
  }
  return pages;
}

async function bmp(c) {
  const w = c.width, h = c.height, d = c.getContext("2d").getImageData(0, 0, w, h).data, o = new DataView(new ArrayBuffer(54 + w * h * 4));
  o.setUint16(0, 0x4d42, true); o.setUint32(2, 54 + w*h*4, true); o.setUint32(10, 54, true); o.setUint32(14, 40, true);
  o.setInt32(18, w, true); o.setInt32(22, -h, true); o.setUint16(26, 1, true); o.setUint16(28, 32, true); o.setUint32(34, w*h*4, true);
  let p = 54; for (let i = 0; i < d.length; i += 4) { o.setUint8(p++, d[i+2]); o.setUint8(p++, d[i+1]); o.setUint8(p++, d[i]); o.setUint8(p++, d[i+3]); }
  return new Blob([o], { type: "image/bmp" });
}

async function decode(f) {
  const ctx = new (window.AudioContext || window.webkitAudioContext)();
  try { return await ctx.decodeAudioData(await f.arrayBuffer()); } 
  catch(e) { throw new Error("Format audio non supporté ou fichier altéré."); }
  finally { ctx.close(); }
}

function mp3(buf) {
  const ch = Math.min(2, buf.numberOfChannels), enc = new lamejs.Mp3Encoder(ch, buf.sampleRate, 128), out = [];
  const cv = a => { const r = new Int16Array(a.length); for (let i = 0; i < a.length; i++) { const s = Math.max(-1, Math.min(1, a[i])); r[i] = s < 0 ? s*32768 : s*32767; } return r; };
  const L = cv(buf.getChannelData(0)), R = ch > 1 ? cv(buf.getChannelData(1)) : null;
  for (let i = 0; i < L.length; i += 1152) {
    const m = ch > 1 ? enc.encodeBuffer(L.subarray(i, i + 1152), R.subarray(i, i + 1152)) : enc.encodeBuffer(L.subarray(i, i + 1152));
    if (m.length) out.push(m);
  }
  const e = enc.flush(); if (e.length) out.push(e);
  return new Blob(out, { type: "audio/mpeg" });
}

/* ---------- PDF → Word éditable, style conservé ---------- */
const SYM_RE = /symbol|wingding|dingbat|zapf|webdings|bullet/i;
const BUL = { "\uf0b7": "•", "\uf0a7": "▪", "\uf0d8": "➢", "\uf0fc": "✓", "\uf076": "❖", "\uf0a8": "◻", "\uf06e": "■", "\uf0a1": "○" };
const SYMA = { o: "◦", "§": "▪", "Ø": "➢", "ü": "✓", v: "❖", l: "●", n: "■", q: "❑" };
const MARK_RE = /^(?:[•●○◦▪■□▫‣⁃∙·–—\-➢✓❖◻]|\d{1,3}[.)]|\(\d{1,3}\)|[a-zA-Z][.)]|[ivxlcIVXLC]{1,5}[.)])$/;
const MARK_PREFIX = /^([•●○◦▪■□▫‣⁃∙·–—\-]|\d{1,3}[.)])\s+(\S[\s\S]*)$/;
const tw = v => Math.round(v * 20);

function fontInfo(pg, styles, fontName) {
  let name = "", b = false, i = false, mono = false, serif = false;
  try {
    if (pg.commonObjs.has(fontName)) {
      const f = pg.commonObjs.get(fontName);
      name = f.name || ""; b = !!(f.bold || f.black); i = !!f.italic; mono = !!f.isMonospace; serif = !!f.isSerifFont;
    }
  } catch (e) {}
  const ff = (styles[fontName] || {}).fontFamily || "";
  const n = name.replace(/^[A-Z]{6}\+/, ""), sym = SYM_RE.test(n);
  mono = mono || /mono/i.test(ff); serif = serif || /^serif/i.test(ff);
  b = b || /bold|black|heavy|demi/i.test(n); i = i || /italic|oblique|-it$/i.test(n);
  let fam = n.split(/[-,]/)[0].replace(/(PSMT|MT|PS)$/, "").replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^Helvetica$/, "Arial").trim();
  if (sym || !/^[A-Za-z][A-Za-z ]{2,40}$/.test(fam) || /^g d|^f \d/i.test(fam)) fam = mono ? "Courier New" : serif ? "Times New Roman" : "Arial";
  return { fam, b, i, sym };
}

// Couleur du texte : pixel le plus éloigné de la couleur de fond (la plus fréquente) dans la zone du mot
function sampleColor(px, cw, ch, x0, y0, x1, y1) {
  x0 = Math.max(0, x0 | 0); y0 = Math.max(0, y0 | 0); x1 = Math.min(cw, Math.ceil(x1)); y1 = Math.min(ch, Math.ceil(y1));
  const cnt = new Map(); let bgK = -1, bgN = 0;
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
    const i = (y * cw + x) * 4, k = ((px[i] >> 4) << 8) | ((px[i + 1] >> 4) << 4) | (px[i + 2] >> 4), n = (cnt.get(k) || 0) + 1;
    cnt.set(k, n); if (n > bgN) { bgN = n; bgK = k; }
  }
  if (bgK < 0) return "000000";
  const br = ((bgK >> 8) & 15) * 17, bg = ((bgK >> 4) & 15) * 17, bb = (bgK & 15) * 17;
  let best = 0;
  const dist = i => Math.abs(px[i] - br) + Math.abs(px[i + 1] - bg) + Math.abs(px[i + 2] - bb);
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) best = Math.max(best, dist((y * cw + x) * 4));
  if (best < 90) return (br + bg + bb) / 3 > 128 ? "000000" : "FFFFFF";
  const acc = new Map(), thr = best * .6;
  for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
    const i = (y * cw + x) * 4; if (dist(i) < thr) continue;
    const k = ((px[i] >> 5) << 6) | ((px[i + 1] >> 5) << 3) | (px[i + 2] >> 5), a = acc.get(k) || [0, 0, 0, 0];
    a[0]++; a[1] += px[i]; a[2] += px[i + 1]; a[3] += px[i + 2]; acc.set(k, a);
  }
  const top = [...acc.values()].sort((a, b) => b[0] - a[0])[0], R = Math.round(top[1] / top[0]), G = Math.round(top[2] / top[0]), B = Math.round(top[3] / top[0]);
  if (Math.max(R, G, B) - Math.min(R, G, B) < 40) { const l = (R + G + B) / 3; if (l < 70) return "000000"; if (l > 215) return "FFFFFF"; }
  return [R, G, B].map(v => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

// Graphismes vectoriels (traits, rectangles remplis, puces rondes) et images, en coordonnées de page (pt, origine en haut)
async function pageGfx(pg, vp) {
  const ops = await pg.getOperatorList(), O = pdfjsLib.OPS, segs = [], rects = [], dots = [], imgs = [], stk = [];
  let m = [1, 0, 0, 1, 0, 0], fc = "000000", sc = "000000", lw = 1, path = null;
  const mul = (a, b) => [a[0]*b[0]+a[1]*b[2], a[0]*b[1]+a[1]*b[3], a[2]*b[0]+a[3]*b[2], a[2]*b[1]+a[3]*b[3], a[4]*b[0]+a[5]*b[2]+b[4], a[4]*b[1]+a[5]*b[3]+b[5]];
  const pt = (x, y) => vp.convertToViewportPoint(m[0]*x + m[2]*y + m[4], m[1]*x + m[3]*y + m[5]);
  const hex = a => typeof a === "string" ? a.replace("#", "").toUpperCase() : [0, 1, 2].map(i => Math.round(a[i]).toString(16).padStart(2, "0")).join("").toUpperCase();
  const parse = (sub, co) => {
    const L = [], Rr = [], P = []; let ci = 0, cx = 0, cy = 0, sx = 0, sy = 0, curves = false;
    for (const o of sub) {
      if (o === O.moveTo) { cx = sx = co[ci++]; cy = sy = co[ci++]; P.push([cx, cy]); }
      else if (o === O.lineTo) { const x = co[ci++], y = co[ci++]; L.push([cx, cy, x, y]); cx = x; cy = y; P.push([x, y]); }
      else if (o === O.curveTo) { for (let k = 0; k < 3; k++) P.push([co[ci++], co[ci++]]); curves = true; cx = P[P.length - 1][0]; cy = P[P.length - 1][1]; }
      else if (o === O.curveTo2 || o === O.curveTo3) { for (let k = 0; k < 2; k++) P.push([co[ci++], co[ci++]]); curves = true; cx = P[P.length - 1][0]; cy = P[P.length - 1][1]; }
      else if (o === O.rectangle) { Rr.push([co[ci], co[ci + 1], co[ci + 2], co[ci + 3]]); P.push([co[ci], co[ci + 1]], [co[ci] + co[ci + 2], co[ci + 1] + co[ci + 3]]); ci += 4; }
      else if (o === O.closePath) { if (cx !== sx || cy !== sy) L.push([cx, cy, sx, sy]); cx = sx; cy = sy; }
    }
    return { L, R: Rr, P, curves };
  };
  const line = (a, b, w, col) => {
    const A = pt(a[0], a[1]), B = pt(b[0], b[1]), dx = B[0] - A[0], dy = B[1] - A[1];
    if (Math.abs(dy) <= .75 && Math.abs(dx) > 1) segs.push({ h: 1, y: (A[1] + B[1]) / 2, x0: Math.min(A[0], B[0]), x1: Math.max(A[0], B[0]), w, col });
    else if (Math.abs(dx) <= .75 && Math.abs(dy) > 1) segs.push({ v: 1, x: (A[0] + B[0]) / 2, y0: Math.min(A[1], B[1]), y1: Math.max(A[1], B[1]), w, col });
  };
  const fillRect = (x0, y0, x1, y1) => {
    const W = x1 - x0, H = y1 - y0;
    if (H <= 2.4 && W > 3) segs.push({ h: 1, y: (y0 + y1) / 2, x0, x1, w: Math.max(.5, H), col: fc });
    else if (W <= 2.4 && H > 3) segs.push({ v: 1, x: (x0 + x1) / 2, y0, y1, w: Math.max(.5, W), col: fc });
    else if (W > 2 && H > 2) {
      rects.push({ x0, y0, x1, y1, fill: fc });
      if (W >= 2.5 && W <= 8 && H >= 2.5 && H <= 8 && Math.abs(W - H) < 1.5) dots.push({ cx: (x0 + x1) / 2, cy: (y0 + y1) / 2, d: W, sq: 1 });
    }
  };
  const paint = (stroke, fill) => {
    if (!path) return; const P = path; path = null; const s = Math.hypot(m[0], m[1]) || 1;
    if (stroke) {
      const w = Math.max(.25, lw * s);
      P.L.forEach(l => line([l[0], l[1]], [l[2], l[3]], w, sc));
      P.R.forEach(([x, y, a, b]) => { line([x, y], [x + a, y], w, sc); line([x + a, y], [x + a, y + b], w, sc); line([x, y + b], [x + a, y + b], w, sc); line([x, y], [x, y + b], w, sc); });
    }
    if (fill) {
      P.R.forEach(([x, y, a, b]) => { const A = pt(x, y), B = pt(x + a, y + b); fillRect(Math.min(A[0], B[0]), Math.min(A[1], B[1]), Math.max(A[0], B[0]), Math.max(A[1], B[1])); });
      if (!P.R.length && P.L.length >= 3 && !P.curves && P.L.every(l => Math.abs(l[0] - l[2]) < .01 || Math.abs(l[1] - l[3]) < .01)) {
        const q = P.P.map(p => pt(p[0], p[1])), xs = q.map(p => p[0]), ys = q.map(p => p[1]);
        fillRect(Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys));
      }
    }
    if (P.curves && !P.R.length) {   // petit cercle = puce vectorielle
      const q = P.P.map(p => pt(p[0], p[1])), xs = q.map(p => p[0]), ys = q.map(p => p[1]);
      const W = Math.max(...xs) - Math.min(...xs), H = Math.max(...ys) - Math.min(...ys);
      if (W >= 2 && W <= 9 && H >= 2 && H <= 9 && Math.abs(W - H) < 1.5) dots.push({ cx: (Math.min(...xs) + Math.max(...xs)) / 2, cy: (Math.min(...ys) + Math.max(...ys)) / 2, d: W, hollow: !fill });
    }
  };
  for (let k = 0; k < ops.fnArray.length; k++) {
    const f = ops.fnArray[k], a = ops.argsArray[k];
    if (f === O.save) stk.push({ m, fc, sc, lw });
    else if (f === O.restore) { const s = stk.pop(); if (s) ({ m, fc, sc, lw } = s); }
    else if (f === O.transform) m = mul(a, m);
    else if (f === O.setLineWidth) lw = a[0];
    else if (f === O.setFillRGBColor) fc = hex(a);
    else if (f === O.setStrokeRGBColor) sc = hex(a);
    else if (f === O.constructPath) path = parse(a[0], a[1]);
    else if (f === O.stroke || f === O.closeStroke) paint(true, false);
    else if (f === O.fill || f === O.eoFill) paint(false, true);
    else if (f === O.fillStroke || f === O.eoFillStroke || f === O.closeFillStroke || f === O.closeEOFillStroke) paint(true, true);
    else if (f === O.endPath) path = null;
    else if (f === O.paintFormXObjectBegin) { stk.push({ m, fc, sc, lw }); if (a[0]) m = mul(a[0], m); }
    else if (f === O.paintFormXObjectEnd) { const s = stk.pop(); if (s) ({ m, fc, sc, lw } = s); }
    else if (f === O.paintImageXObject || f === O.paintInlineImageXObject || f === O.paintImageMaskXObject) {
      const q = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([u, v]) => pt(u, v)), xs = q.map(p => p[0]), ys = q.map(p => p[1]);
      const x0 = Math.min(...xs), x1 = Math.max(...xs), y0 = Math.min(...ys), y1 = Math.max(...ys);
      if (x1 - x0 >= 12 && y1 - y0 >= 12 && (x1 - x0) * (y1 - y0) <= .85 * vp.width * vp.height) imgs.push({ x0, x1, y0, y1 });
    }
  }
  return { segs, rects, dots, imgs };
}

function mergeSegs(arr, key, a0, a1) {
  arr = [...arr].sort((p, q) => p[key] - q[key] || p[a0] - q[a0]); const out = [];
  for (const s of arr) {
    const o = out.find(o => Math.abs(o[key] - s[key]) < 1.2 && s[a0] <= o[a1] + 2 && s[a1] >= o[a0] - 2);
    if (o) { o[a0] = Math.min(o[a0], s[a0]); o[a1] = Math.max(o[a1], s[a1]); if (s.w > o.w) { o.w = s.w; o.col = s.col; } } else out.push({ ...s });
  }
  return out;
}

function buildGrid(H, V, rects) {
  const clus = vals => { vals = [...vals].sort((a, b) => a - b); const o = []; for (const v of vals) if (!o.length || v - o[o.length - 1] > 2) o.push(v); return o; };
  const ys = clus(H.map(h => h.y)), xs = clus(V.map(v => v.x));
  if (ys.length < 2 || xs.length < 2 || (ys.length - 1) * (xs.length - 1) > 3000) return null;
  const nr = ys.length - 1, nc = xs.length - 1;
  const hAt = (y, a, b) => H.find(h => Math.abs(h.y - y) < 2.5 && h.x0 <= (a + b) / 2 + 1 && h.x1 >= (a + b) / 2 - 1);
  const vAt = (x, a, b) => V.find(v => Math.abs(v.x - x) < 2.5 && v.y0 <= (a + b) / 2 + 1 && v.y1 >= (a + b) / 2 - 1);
  const own = Array.from({ length: nr }, () => Array(nc).fill(null)), cells = []; let closed = 0, total = 0;
  for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
    if (own[r][c]) continue;
    let cs = 1, rs = 1;
    while (c + cs < nc && !own[r][c + cs] && !vAt(xs[c + cs], ys[r], ys[r + 1])) cs++;
    while (r + rs < nr && ![...Array(cs).keys()].some(k => hAt(ys[r + rs], xs[c + k], xs[c + k + 1]))) rs++;
    const cell = { r, c, cs, rs, x0: xs[c], x1: xs[c + cs], y0: ys[r], y1: ys[r + rs], items: [], imgs: [] };
    const ln = x => x ? { w: Math.min(24, Math.max(2, Math.round(x.w * 8))), col: x.col } : null;
    cell.bt = ln(hAt(ys[r], cell.x0, cell.x1)); cell.bb = ln(hAt(ys[r + rs], cell.x0, cell.x1));
    cell.bl = ln(vAt(xs[c], cell.y0, cell.y1)); cell.br = ln(vAt(xs[c + cs], cell.y0, cell.y1));
    total++; if (cell.bt && cell.bb && cell.bl && cell.br) closed++;
    for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) own[r + i][c + j] = cell;
    const cx = (cell.x0 + cell.x1) / 2, cy = (cell.y0 + cell.y1) / 2;
    const fr = rects.filter(q => cx > q.x0 && cx < q.x1 && cy > q.y0 && cy < q.y1 && (q.x1 - q.x0) * (q.y1 - q.y0) < (xs[nc] - xs[0] + 4) * (ys[nr] - ys[0] + 4) * 1.02)
      .sort((p, q) => (p.x1 - p.x0) * (p.y1 - p.y0) - (q.x1 - q.x0) * (q.y1 - q.y0))[0];
    cell.fill = fr && !/^(FFFFFF|FEFEFE|FDFDFD)$/.test(fr.fill) ? fr.fill : null;
    if (fr) cell.rect = fr;
    cells.push(cell);
  }
  if (nr * nc > 1 && closed / total < .7) return null;
  return { xs, ys, nr, nc, cells, own, x0: xs[0], x1: xs[nc], top: ys[0], bottom: ys[nr] };
}

function findTables(H, V, rects) {
  const nodes = [...H, ...V], par = nodes.map((_, i) => i), used = new Set(), tables = [];
  const find = i => par[i] === i ? i : (par[i] = find(par[i]));
  H.forEach((h, i) => V.forEach((v, j) => {
    if (h.y >= v.y0 - 2.5 && h.y <= v.y1 + 2.5 && v.x >= h.x0 - 2.5 && v.x <= h.x1 + 2.5) par[find(i)] = find(H.length + j);
  }));
  const comps = {}; nodes.forEach((s, i) => (comps[find(i)] = comps[find(i)] || []).push(s));
  for (const c of Object.values(comps)) {
    const h = c.filter(s => s.h), v = c.filter(s => s.v);
    if (h.length < 2 || v.length < 2) continue;
    const g = buildGrid(h, v, rects); if (g) { tables.push(g); c.forEach(s => used.add(s)); }
  }
  return { tables, used };
}

const MCTX = document.createElement("canvas").getContext("2d");
function markRatio(f, pre) {   // part de la largeur du texte occupée par le préfixe (mesure avec la vraie police si possible)
  try {
    MCTX.font = `${f.fi.i ? "italic " : ""}${f.fi.b ? "bold " : ""}${f.size}px "${f.fi.fam}", sans-serif`;
    const a = MCTX.measureText(pre).width, b = MCTX.measureText(f.str).width; if (a > 0 && b > 0) return Math.min(.9, a / b);
  } catch (e) {}
  return pre.length / f.str.length;
}

function buildLines(items, dots) {
  items.sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [], out = [];
  for (const it of items) { const l = lines.find(l => Math.abs(l.y - it.y) < Math.max(2, it.size * .35)); if (l) l.it.push(it); else lines.push({ y: it.y, it: [it] }); }
  lines.sort((a, b) => a.y - b.y);
  for (const l of lines) {
    l.it.sort((a, b) => a.x - b.x);
    let mk = null; const f = l.it[0], nx = l.it[1];
    const dot = dots.find(d => !d.used && d.cx < f.x - .5 && f.x - d.cx < 45 && Math.abs(d.cy - (l.y - f.size * .32)) < f.size * .6);
    if (dot) { dot.used = true; mk = { ch: dot.hollow ? "◦" : dot.sq ? "▪" : "•", x: dot.cx - dot.d / 2, textX: f.x }; }
    else if (nx && nx.x > f.x + 1 && nx.x - (f.x + f.w) < f.size * 3 && (f.fi.sym || MARK_RE.test(f.str) || (/^\d{1,2}$/.test(f.str) && (f.size !== nx.size || f.fi.fam !== nx.fi.fam) && nx.x - (f.x + f.w) > f.size * .5))) { mk = { ch: f.str, x: f.x, textX: nx.x, it: f }; l.it.shift(); }
    else { const mm = MARK_PREFIX.exec(f.str); if (mm) { const tx = f.x + f.w * markRatio(f, mm[1] + " "); mk = { ch: mm[1], x: f.x, textX: tx }; f.w -= tx - f.x; f.x = tx; f.str = mm[2]; } }
    if (!l.it.length) continue;
    l.runs = []; let end = null; const sz = {};
    for (const o of l.it) {
      const gap = end === null ? 0 : o.x - end, last0 = l.runs[l.runs.length - 1]; let txt = o.str, sp = false;
      if (gap > o.size * 2.5) l.runs.push({ tab: true, x: o.x });
      else if (end !== null && gap > o.size * .12 && !/^\s/.test(txt) && !/\s$/.test((last0 && last0.text) || "")) sp = true;
      const key = `${o.fi.fam}|${o.fi.b}|${o.fi.i}|${Math.round(o.size * 2)}|${o.color}|${o.u ? 1 : 0}|${o.shd || ""}`, last = l.runs[l.runs.length - 1];
      if (sp) { if (last && !last.tab && !last.u && o.u && last.key !== key) last.text += " "; else txt = " " + txt; }
      if (last && !last.tab && last.key === key) last.text += txt;
      else l.runs.push({ text: txt, key, fam: o.fi.fam, b: o.fi.b, i: o.fi.i, sz: Math.max(2, Math.round(o.size * 2)), color: o.color, u: o.u, shd: o.shd });
      const k = o.size.toFixed(1); sz[k] = (sz[k] || 0) + o.str.length; end = o.x + o.w;
    }
    const r0 = l.runs.find(r => !r.tab);
    l.mk = mk; l.textX = mk && mk.textX; l.mkRun = mk && (mk.it ? { text: mk.ch, fam: mk.it.fi.sym ? "Arial" : mk.it.fi.fam, b: mk.it.fi.b, i: mk.it.fi.i, sz: Math.max(2, Math.round(mk.it.size * 2)), color: mk.it.color } : { ...r0, text: mk.ch, u: false, ...(/^[•◦▪➢✓❖●■]$/.test(mk.ch) ? { fam: "Arial" } : {}) });
    l.x0 = mk ? mk.x : l.it[0].x; l.x1 = end; l.len = l.it.reduce((s, o) => s + o.str.length, 0);
    l.size = +Object.entries(sz).sort((a, b) => b[1] - a[1])[0][0];
    out.push(l);
  }
  return out;
}

function addLine(l, S) {
  const hasTab = l.runs.some(r => r.tab), lc = (l.x0 + l.x1) / 2, mk = l.mk, c = S.cur, bx = S.shade ? S.shade(l) : null;
  const center = !mk && !hasTab && Math.abs(lc - S.center) < 5 && l.x0 - S.left > 8 && S.right - l.x1 > 8;
  const right = !mk && !hasTab && !center && S.right - l.x1 < 3 && l.x0 - S.left > 25;
  const al = center ? "center" : right ? "right" : "left";
  if (c && c.box === bx && !mk && !hasTab && !c.hasTab && al === "left" && c.al === "left" && c.lastX1 >= S.right - c.size * 2.5
      && l.y - c.lastY > 0 && l.y - c.lastY < c.size * 1.9 && Math.abs(l.size - c.size) < .6
      && (c.textX != null ? Math.abs(l.x0 - c.textX) < 4 : l.x0 <= Math.max(c.x0, c.lastX0) + 3)) {
    if (c.lines === 1) { c.lead = l.y - c.lastY; c.before = Math.max(0, c.y0 - c.lead * .8 - c.prev0); }
    const last = c.runs[c.runs.length - 1], first = l.runs[0];
    if (/[A-Za-zÀ-ÿ]-$/.test(last.text) && /^[a-zà-ÿ]/.test(first.text)) last.text = last.text.slice(0, -1);
    else first.text = " " + first.text.replace(/^\s+/, "");
    c.runs.push(...l.runs); c.lastX0 = l.x0; c.lastX1 = l.x1; c.lastY = l.y; c.lines++; c.ends.push(l.x1);
  } else {
    const lead = Math.max(l.size, mk && mk.it ? mk.it.size : 0) * 1.2, runs = [], tabs = [];
    if (mk) { runs.push({ ...l.mkRun }, { tab: true }); tabs.push(l.textX - S.left); }
    l.runs.forEach(r => { runs.push(r); if (r.tab) tabs.push(r.x - S.left); });
    const ind = mk ? Math.max(0, l.textX - S.left) : al === "left" ? Math.max(0, l.x0 - S.left) : 0;
    S.cur = { runs, tabs, al, ends: [l.x1], size: l.size, lead, y0: l.y, prev0: S.prevBottom, x0: l.x0, lastX0: l.x0, lastX1: l.x1, lastY: l.y, lines: 1, hasTab, textX: mk ? l.textX : null,
      pb: S.pbNext, before: Math.max(0, l.y - lead * .8 - S.prevBottom), ind, hang: mk ? Math.min(ind, Math.max(0, l.textX - mk.x)) : 0, box: bx };
    S.paras.push(S.cur); S.pbNext = false;
  }
  S.prevBottom = S.cur.lastY + S.cur.lead * .2;
}

function flow(els, S) {
  for (const e of els) {
    if (e.l) addLine(e.l, S);
    else if (e.m) {
      S.cur = null; S.paras.push({ img: e.m, before: Math.max(0, e.m.y0 - S.prevBottom), ind: Math.max(0, e.m.x0 - S.left), pb: S.pbNext, maxW: S.right - S.left + 4 });
      S.pbNext = false; S.prevBottom = e.m.y1;
    } else if (e.tb) {
      S.cur = null; const gap = e.tb.top - S.prevBottom;
      if (gap > 1.5 || S.pbNext) S.paras.push({ spacer: Math.max(1, gap), pb: S.pbNext });
      S.pbNext = false; S.paras.push(e.tb); S.prevBottom = e.tb.bottom;
    } else if (e.rule) {
      const r = e.rule, bw = Math.max(.5, r.w), last = S.cur, lb = last ? last.lastY + last.lead * .2 : 0, gap = r.y - lb;
      const bdr = { sz: Math.min(24, Math.max(2, Math.round(bw * 8))), col: r.col };
      if (last && !last.bdr && !last.box && S.paras[S.paras.length - 1] === last && gap > -3 && gap < 20) { last.bdr = { ...bdr, space: Math.min(31, Math.max(0, Math.round(gap - bw / 2))) }; }
      else { S.paras.push({ spacer: Math.max(1, r.y - S.prevBottom), bdr: { ...bdr, space: 0 }, pb: S.pbNext }); S.pbNext = false; S.cur = null; }
      S.prevBottom = r.y + bw / 2;
    }
  }
  for (const p of S.paras) {   // texte justifié : toutes les lignes sauf la dernière finissent au bord droit
    if (!p.ends || p.lines < 2 || p.al !== "left" || p.hasTab) continue;
    const nl = p.ends.slice(0, -1);
    if (nl.every(x => x >= S.right - (nl.length > 1 ? 1.5 : .8))) p.al = "both";
  }
}

function tableBlock(t, items, imgs, dots) {
  const near = (list, pt) => list.find(c => pt[0] >= c.x0 - 1 && pt[0] <= c.x1 + 1 && pt[1] >= c.y0 - 1 && pt[1] <= c.y1 + 1) ||
    [...list].sort((a, b) => Math.hypot((a.x0 + a.x1) / 2 - pt[0], (a.y0 + a.y1) / 2 - pt[1]) - Math.hypot((b.x0 + b.x1) / 2 - pt[0], (b.y0 + b.y1) / 2 - pt[1]))[0];
  items.forEach(o => near(t.cells, [o.x + o.w / 2, o.y - o.size * .3]).items.push(o));
  imgs.forEach(m => near(t.cells, [(m.x0 + m.x1) / 2, (m.y0 + m.y1) / 2]).imgs.push(m));
  const offs = []; t.cells.forEach(c => { if (c.items.length) { const d = Math.min(...c.items.map(o => o.x)) - c.x0; if (d >= 0 && d < 20) offs.push(d); } });
  offs.sort((a, b) => a - b); const pad = offs.length ? Math.min(9, Math.max(1.5, offs[offs.length >> 1])) : 5.4;
  const cellOut = {};
  for (const c of t.cells) {
    const els = [...buildLines(c.items, dots).map(l => ({ l, top: l.y - l.size * .96 })), ...c.imgs.map(m => ({ m, top: m.y0 }))].sort((a, b) => a.top - b.top);
    const S = { paras: [], left: c.x0 + pad, right: c.x1 - pad, center: (c.x0 + c.x1) / 2, prevBottom: c.y0, pbNext: false, cur: null, shade: null };
    flow(els, S);
    let va = null;
    if (S.paras.length) {
      const p0 = S.paras[0], h = c.y1 - c.y0, cb = S.prevBottom, ct = c.y0 + p0.before, tp = ct - c.y0, bt = c.y1 - cb;
      if (tp > 2.5 && Math.abs(tp - bt) < 2.5) { va = "center"; p0.before = 0; }
      else if (tp > 2.5 && bt < 2.5) { va = "bottom"; p0.before = 0; }
      else p0.before = Math.min(p0.before, Math.max(0, h - (cb - ct)));
    }
    cellOut[c.r + "," + c.c] = { c, paras: S.paras, va };
  }
  const rows = [];
  for (let r = 0; r < t.nr; r++) {
    const cs = [];
    for (let k = 0; k < t.nc; k++) {
      const o = t.own[r][k]; if (o.c !== k) continue;
      if (o.r === r) cs.push({ ...cellOut[o.r + "," + o.c], vm: o.rs > 1 ? "restart" : null });
      else cs.push({ c: o, paras: [], va: null, vm: "cont", last: r === o.r + o.rs - 1 });
    }
    rows.push({ h: t.ys[r + 1] - t.ys[r], cells: cs });
  }
  return { tbl: { widths: t.xs.slice(1).map((x, i) => x - t.xs[i]), x0: t.x0, pad, rows }, top: t.top, bottom: t.bottom };
}

/* --- génération XML Word --- */
const rXml = r => r.br ? "<w:r><w:br/></w:r>" : r.field ? `<w:fldSimple w:instr=" ${r.field} ">${rXml({ ...r, field: null })}</w:fldSimple>` : r.tab ? "<w:r><w:tab/></w:r>" :
  `<w:r><w:rPr><w:rFonts w:ascii="${r.fam}" w:hAnsi="${r.fam}" w:eastAsia="${r.fam}" w:cs="${r.fam}"/>${r.b ? "<w:b/><w:bCs/>" : ""}${r.i ? "<w:i/><w:iCs/>" : ""}<w:color w:val="${r.color}"/><w:sz w:val="${r.sz}"/><w:szCs w:val="${r.sz}"/>${r.u ? '<w:u w:val="single"/>' : ""}${r.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${r.shd}"/>` : ""}</w:rPr><w:t xml:space="preserve">${esc(r.text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, ""))}</w:t></w:r>`;
const bdrXml = b => b ? `w:val="single" w:sz="${b.sz || b.w}" w:space="${b.space || 0}" w:color="${b.col}"` : 'w:val="nil"';
const EMPTY_P = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>';

function tblXml(t, lm, ctx) {
  const sum = t.widths.reduce((a, b) => a + b, 0);
  let x = `<w:tbl><w:tblPr><w:tblW w:w="${tw(sum)}" w:type="dxa"/><w:tblInd w:w="${tw(t.x0 - lm)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${tw(t.pad)}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${tw(t.pad)}" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${t.widths.map(w => `<w:gridCol w:w="${tw(w)}"/>`).join("")}</w:tblGrid>`;
  for (const r of t.rows) {
    x += `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${tw(Math.max(1, r.h - Math.max(0, ...r.cells.map(cl => ((cl.c.bb || {}).w || 0) / 8))))}" w:hRule="atLeast"/></w:trPr>`;
    for (const cl of r.cells) {
      const c = cl.c, w = c.x1 - c.x0, top = cl.vm === "cont" ? null : c.bt, bot = cl.vm === "cont" && !cl.last ? null : c.bb;
      x += `<w:tc><w:tcPr><w:tcW w:w="${tw(w)}" w:type="dxa"/>${c.cs > 1 ? `<w:gridSpan w:val="${c.cs}"/>` : ""}${cl.vm === "restart" ? '<w:vMerge w:val="restart"/>' : cl.vm === "cont" ? "<w:vMerge/>" : ""}<w:tcBorders><w:top ${bdrXml(top)}/><w:left ${bdrXml(c.bl)}/><w:bottom ${bdrXml(bot)}/><w:right ${bdrXml(c.br)}/></w:tcBorders>${c.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${c.fill}"/>` : ""}${cl.va ? `<w:vAlign w:val="${cl.va}"/>` : ""}</w:tcPr>${cl.paras.map(p => pXml(p, lm, ctx)).join("") || EMPTY_P}</w:tc>`;
    }
    x += "</w:tr>";
  }
  return x + "</w:tbl>";
}

function pXml(p, lm, ctx) {
  if (p.tbl) return tblXml(p.tbl, lm, ctx);
  let ppr = (p.keep ? "<w:keepNext/>" : "") + (p.pb ? "<w:pageBreakBefore/>" : "");
  let rInd = 0;
  if (p.bdr) ppr += `<w:pBdr><w:bottom w:val="single" w:sz="${p.bdr.sz}" w:space="${p.bdr.space}" w:color="${p.bdr.col}"/></w:pBdr>`;
  if (p.box) {
    const b = p.box, bw = b.bar ? b.bar.w : 0, inv = sp => `w:val="single" w:sz="2" w:space="${sp}" w:color="${b.fill}"`, cl = v => Math.min(31, Math.max(0, Math.round(v)));
    const padR = 6, edge = ctx.rm + 2; rInd = Math.max(0, edge - (b.x1 - padR));
    const spR = b.x1 >= edge ? cl(b.x1 - edge - .25) : padR;
    ppr += `<w:pBdr>${p.bTop ? `<w:top ${inv(p.bTop)}/>` : ""}<w:left ${b.bar ? `w:val="single" w:sz="${Math.max(2, Math.round(bw * 8))}" w:space="${cl(p.x0 - b.x0 - bw)}" w:color="${b.bar.col}"` : inv(cl(p.x0 - b.x0))}/>${p.bBot ? `<w:bottom ${inv(p.bBot)}/>` : ""}<w:right ${inv(spR)}/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="${b.fill}"/>`;
  }
  if (p.spacer) return `<w:p><w:pPr>${ppr}<w:spacing w:before="0" w:after="0" w:line="${tw(p.spacer)}" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`;
  if (p.tabs && p.tabs.length) ppr += "<w:tabs>" + p.tabs.map(x => `<w:tab w:val="left" w:pos="${tw(x)}"/>`).join("") + "</w:tabs>";
  ppr += `<w:spacing w:before="${tw(Math.min(p.before, 400))}" w:after="0" ${p.img ? 'w:line="240" w:lineRule="auto"' : `w:line="${tw(p.lead)}" w:lineRule="${p.lines > 1 ? "exact" : "atLeast"}"`}/>`;
  if (p.ind > 1 || p.hang > 0 || rInd > 0) ppr += `<w:ind w:left="${tw(p.ind)}"${rInd > 0 ? ` w:right="${tw(rInd)}"` : ""}${p.hang > 0 ? ` w:hanging="${tw(p.hang)}"` : ""}/>`;
  if (p.al && p.al !== "left") ppr += `<w:jc w:val="${p.al}"/>`;
  if (p.img) {
    const n = ++ctx.n; ctx.z.file(`word/media/i${n}.jpeg`, p.img.data);
    ctx.rels += `<Relationship Id="rI${n}" Type="${ctx.R}/image" Target="media/i${n}.jpeg"/>`;
    let w = p.img.x1 - p.img.x0, h = p.img.y1 - p.img.y0; if (p.maxW && w > p.maxW) { h *= p.maxW / w; w = p.maxW; }
    const cx = Math.round(w * 12700), cy = Math.round(h * 12700);
    return `<w:p><w:pPr>${ppr}</w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${n}" name="i${n}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${n}" name="i${n}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rI${n}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }
  const mx = Math.max(...p.runs.filter(r => !r.tab).map(r => r.sz));
  return `<w:p><w:pPr>${ppr}<w:rPr><w:sz w:val="${mx}"/><w:szCs w:val="${mx}"/></w:rPr></w:pPr>${p.runs.map(rXml).join("")}</w:p>`;
}

async function makeRichDocx(blocks, pgx) {
  const z = new JSZip(), W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const ctx = { z, R, n: 0, rels: "", rm: pgx.rm };
  let body = blocks.map(p => pXml(p, pgx.lm, ctx)).join("");
  if (blocks.length && blocks[blocks.length - 1].tbl) body += EMPTY_P;
  const HX = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`;
  const part = (tag, P) => `${HX}<w:${tag} ${NS}>${P.paras.map(p => pXml(p, pgx.lm, ctx)).join("") || EMPTY_P}</w:${tag}>`;
  if (pgx.hdr) z.file("word/header1.xml", part("hdr", pgx.hdr));
  if (pgx.ftr) z.file("word/footer1.xml", part("ftr", pgx.ftr));
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' + (pgx.hdr ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : "") + (pgx.ftr ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : "") + "</Types>");
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  z.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
  z.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  z.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rS1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rS2" Type="${R}/settings" Target="settings.xml"/>${pgx.hdr ? `<Relationship Id="rH1" Type="${R}/header" Target="header1.xml"/>` : ""}${pgx.ftr ? `<Relationship Id="rF1" Type="${R}/footer" Target="footer1.xml"/>` : ""}${ctx.rels}</Relationships>`);
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}<w:sectPr>${pgx.hdr ? '<w:headerReference w:type="default" r:id="rH1"/>' : ""}${pgx.ftr ? '<w:footerReference w:type="default" r:id="rF1"/>' : ""}<w:pgSz w:w="${tw(pgx.pw)}" w:h="${tw(pgx.ph)}"/><w:pgMar w:top="${tw(pgx.top)}" w:right="${tw(Math.max(18, pgx.pw - pgx.rm - 2))}" w:bottom="${tw(pgx.bot)}" w:left="${tw(pgx.lm)}" w:header="${tw(pgx.hdr ? pgx.hdr.top : 0)}" w:footer="${tw(pgx.ftr ? pgx.ph - pgx.ftr.bottom : 0)}" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function pdfLayout(pdf, prog) {
  const n = pdf.numPages, pgs = []; let pw = 595, ph = 842;
  for (let i = 1; i <= n; i++) {
    const pg = await pdf.getPage(i), vp = pg.getViewport({ scale: 1 });
    if (i === 1) { pw = vp.width; ph = vp.height; }
    const cv = await renderPage(pdf, i, 2), px = cv.getContext("2d").getImageData(0, 0, cv.width, cv.height).data;
    const gfx = await pageGfx(pg, vp), tc = await pg.getTextContent(), fc = {}, seen = new Set(), items = [];
    for (const s of tc.items) {
      if (s.str === undefined || !s.transform) continue;
      const fi = fc[s.fontName] || (fc[s.fontName] = fontInfo(pg, tc.styles, s.fontName));
      let str = s.str; if (str === "") { if (!fi.sym) continue; str = "•"; } else if (!str.trim()) continue;
      if (fi.sym && str.length === 1 && SYMA[str]) str = SYMA[str];
      str = str.replace(/[\uf000-\uf0ff]/g, c => BUL[c] || "•");
      const t = s.transform, size = Math.hypot(t[0], t[1]) || s.height || 10, y = vp.height - t[5], w = s.width || size * (fi.sym ? .35 : str.length * .5);
      const sig = `${Math.round(t[4])}|${Math.round(y)}|${str}`; if (seen.has(sig)) continue; seen.add(sig);
      const col = (x, wd) => sampleColor(px, cv.width, cv.height, x * 2, (y - size * .8) * 2, (x + wd) * 2, (y + size * .2) * 2);
      const toks = fi.sym ? [] : [...str.matchAll(/\S+/g)];
      if (toks.length) {   // un morceau par mot : couleur et soulignement précis au mot près
        MCTX.font = `${fi.i ? "italic " : ""}${fi.b ? "bold " : ""}${size}px "${fi.fam}", sans-serif`;
        const tot = MCTX.measureText(str).width || 0;
        for (const k of toks) {
          const a = tot ? MCTX.measureText(str.slice(0, k.index)).width / tot : k.index / str.length;
          const b = tot ? MCTX.measureText(str.slice(0, k.index + k[0].length)).width / tot : (k.index + k[0].length) / str.length;
          const x = t[4] + w * a, wd = w * (b - a); items.push({ x, y, w: wd, size, str: k[0], fi, color: col(x, wd) });
        }
      } else items.push({ x: t[4], y, w, size, str, fi, color: col(t[4], w) });
    }
    const H = mergeSegs(gfx.segs.filter(s => s.h), "y", "x0", "x1"), V = mergeSegs(gfx.segs.filter(s => s.v), "x", "y0", "y1");
    const rects = gfx.rects.filter(r => (r.x1 - r.x0) * (r.y1 - r.y0) < .7 * vp.width * vp.height && !/^(FFFFFF|FEFEFE|FDFDFD)$/.test(r.fill));
    const { tables, used } = findTables(H, V, rects), rest = H.filter(h => !used.has(h)), usedU = new Set();
    const ext = {}; items.forEach(o => { const k = Math.round(o.y), e = ext[k] || (ext[k] = { x0: 1e9, x1: -1e9 }); e.x0 = Math.min(e.x0, o.x); e.x1 = Math.max(e.x1, o.x + o.w); });
    const lineExt = o => { const k = Math.round(o.y), a = [ext[k - 1], ext[k], ext[k + 1]].filter(Boolean); return { x0: Math.min(...a.map(e => e.x0)), x1: Math.max(...a.map(e => e.x1)) }; };
    items.forEach(o => {
      const cx = o.x + o.w / 2, cy = o.y - o.size * .3, e = lineExt(o);
      const u = rest.find(h => h.w <= 1.6 && h.y > o.y + .5 && h.y < o.y + o.size * .3 + 1 && cx >= h.x0 && cx <= h.x1 && h.x0 >= e.x0 - 4 && h.x1 <= e.x1 + 4);
      if (u) { o.u = true; usedU.add(u); }
      const q = rects.find(r => cx > r.x0 && cx < r.x1 && cy > r.y0 && cy < r.y1 && r.y1 - r.y0 < o.size * 2.4 && !(r.x0 <= e.x0 + 8 && r.x1 >= e.x1 - 4));
      if (q && !/^(FFFFFF|FEFEFE)$/.test(q.fill)) o.shd = q.fill;
    });
    const rules = rest.filter(h => !usedU.has(h) && h.x1 - h.x0 >= 40);
    // crop images
    const imgs = [];
    for (const m of gfx.imgs) {
      const c = document.createElement("canvas"); c.width = Math.round((m.x1 - m.x0) * 2); c.height = Math.round((m.y1 - m.y0) * 2);
      c.getContext("2d").drawImage(cv, m.x0 * 2, m.y0 * 2, c.width, c.height, 0, 0, c.width, c.height);
      imgs.push({ ...m, data: new Uint8Array(await (await toBlob(c, "image/jpeg", .9)).arrayBuffer()) });
    }
    // répartition tableaux / flux
    const inT = (t, x, y) => x >= t.x0 - 3 && x <= t.x1 + 3 && y >= t.top - 3 && y <= t.bottom + 3;
    const flowItems = [], blocks = [];
    for (const t of tables) {
      const ti = items.filter(o => inT(t, o.x + o.w / 2, o.y - o.size * .3)), tm = imgs.filter(m => inT(t, (m.x0 + m.x1) / 2, (m.y0 + m.y1) / 2));
      ti.forEach(o => (o.inT = 1)); tm.forEach(m => (m.inT = 1));
      blocks.push(tableBlock(t, ti, tm, gfx.dots));
    }
    items.forEach(o => !o.inT && flowItems.push(o));
    const shades = rects.filter(r => !tables.some(t => inT(t, (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)));
    pgs.push({ lines: buildLines(flowItems, gfx.dots), imgs: imgs.filter(m => !m.inT), blocks, rules, shades, vh: vp.height }); prog(i / n * .9);
  }
  // en-têtes / pieds de page : lignes répétées en haut ou en bas sur au moins la moitié des pages
  const zone = (l, p) => l.y < p.vh * .1 ? "h" : l.y > p.vh * .9 ? "f" : "";
  const hfKey = (l, p) => zone(l, p) + "|" + l.runs.map(r => r.text || "\t").join("").replace(/\d+/g, "#").trim() + "|" + Math.round(l.y / 3);
  const hfFirst = {};
  if (n > 1) {
    const cnt = {}; pgs.forEach(p => new Set(p.lines.filter(l => zone(l, p)).map(l => hfKey(l, p))).forEach(k => (cnt[k] = (cnt[k] || 0) + 1)));
    const need = Math.max(2, Math.ceil(n * .5));
    pgs.forEach((p, pi) => { p.lines = p.lines.filter(l => { const z = zone(l, p), k = hfKey(l, p); if (!z || cnt[k] < need) return true; if (!hfFirst[k]) hfFirst[k] = { l, pi, z }; return false; }); });
  }
  const all = pgs.flatMap(p => p.lines), tbls = pgs.flatMap(p => p.blocks);
  if (!all.length && !tbls.length) throw new Error("Aucun texte trouvé (PDF scanné ?) — choisis « docx (pages en images) »");
  const long = all.filter(l => l.len > 25), cnt = {};
  long.forEach(l => { const k = Math.round(l.x0 / 2) * 2; cnt[k] = (cnt[k] || 0) + 1; });
  const top1 = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0];
  const lm = top1 ? +top1[0] : Math.min(...all.map(l => l.x0), ...tbls.map(b => b.tbl.x0), 36);
  const rm = Math.max(...(long.length ? long : all).map(l => l.x1), ...tbls.map(b => b.tbl.x0 + b.tbl.widths.reduce((a, c) => a + c, 0)), lm + 100);
  const els = pgs.map(p => [
    ...p.lines.map(l => ({ l, top: l.y - l.size * .96 })), ...p.imgs.map(m => ({ m, top: m.y0 })),
    ...p.blocks.map(b => ({ tb: b, top: b.top })), ...p.rules.map(r => ({ rule: r, top: r.y }))
  ].sort((a, b) => a.top - b.top));
  const mkPart = z => {
    const arr = Object.values(hfFirst).filter(a => a.z === z).sort((a, b) => a.l.y - b.l.y); if (!arr.length) return null;
    const top = Math.min(...arr.map(a => a.l.y - a.l.size * .96)), S2 = { paras: [], left: lm, right: rm, center: (lm + rm) / 2, prevBottom: top, pbNext: false, cur: null, shade: null };
    flow(arr.map(a => ({ l: a.l, top: a.l.y })), S2);
    const pi = arr[0].pi;
    S2.paras.forEach(p => { if (!p.runs) return; const out = []; for (const r of p.runs) { if (r.tab || !/\d/.test(r.text)) { out.push(r); continue; } r.text.split(/(\d+)/).forEach((t, k) => { if (!t) return; const f = k % 2 && (t === String(pi + 1) ? "PAGE" : t === String(n) && n > 1 ? "NUMPAGES" : null); out.push(f ? { ...r, text: t, field: f } : { ...r, text: t }); }); } p.runs = out; });
    return { paras: S2.paras, top, bottom: S2.prevBottom };
  };
  const hdrP = mkPart("h"), ftrP = mkPart("f");
  let topPt = Math.min(140, Math.max(20, Math.min(...els.map(e => e.length ? e[0].top : 1e9))));
  let botPt = Math.min(90, Math.max(10, Math.min(...pgs.map((p, i) => p.vh - Math.max(0, ...p.lines.map(l => l.y + l.size * .25), ...p.imgs.map(m => m.y1), ...p.blocks.map(b => b.bottom)))) - 3));
  if (hdrP) topPt = Math.max(topPt, hdrP.bottom + 2);
  if (ftrP) botPt = Math.max(botPt, ph - ftrP.top + 2);
  const S = { paras: [], left: lm, right: rm, center: (lm + rm) / 2, prevBottom: topPt, pbNext: false, cur: null, shade: null };
  pgs.forEach((p, pi) => {
    if (pi > 0) {   // saut de page forcé seulement si la page précédente n'était pas pleine (sinon Word coupe tout seul)
      const first = els[pi][0], lead = first && first.l ? first.l.size * 1.2 : 14, slack = (ph - botPt) - S.prevBottom;
      S.pbNext = slack >= lead * 1.3; S.natural = !S.pbNext;
    }
    S.prevBottom = topPt; S.cur = null;
    S.shade = l => {
      const cy = l.y - l.size * .3;
      const q = p.shades.filter(r => cy > r.y0 && cy < r.y1 && r.x0 <= l.x0 + 8 && r.x1 >= l.x1 - 4).sort((a, b) => (a.x1 - a.x0) * (a.y1 - a.y0) - (b.x1 - b.x0) * (b.y1 - b.y0))[0];
      if (!q) return null;
      if (!q.box) { const bar = p.shades.find(r => r !== q && r.x1 - r.x0 <= 8 && Math.abs(r.x0 - q.x0) < 3 && r.y0 >= q.y0 - 2 && r.y1 <= q.y1 + 2 && r.y1 - r.y0 > 8); q.box = { x0: q.x0, x1: q.x1, y0: q.y0, y1: q.y1, fill: q.fill, bar: bar ? { w: bar.x1 - bar.x0, col: bar.fill } : null }; }
      return q.box;
    };
    const i0 = S.paras.length; flow(els[pi], S);
    if (S.natural && S.paras[i0]) S.paras[i0].before = 0;
  });
  const sz = {}; all.forEach(l => { const k = l.size.toFixed(1); sz[k] = (sz[k] || 0) + l.len; });
  const bodySize = +(Object.entries(sz).sort((a, b) => b[1] - a[1])[0] || [11])[0];
  S.paras.forEach((p, i) => {   // titres : garder avec le paragraphe suivant
    if (p.runs && p.lines === 1 && p.textX == null && !p.hasTab && i < S.paras.length - 1 && p.runs.map(r => r.text || "").join("").length < 140
        && (p.size >= bodySize * 1.15 || p.runs.filter(r => !r.tab).every(r => r.b))) p.keep = true;
  });
  S.paras.forEach((p, i) => {
    if (!p.box) return; const pr = S.paras[i - 1], nx = S.paras[i + 1], b = p.box, cl = v => Math.min(31, Math.max(0, Math.round(v)));
    if (!(pr && pr.box === b)) { p.bTop = cl(p.y0 - p.lead * .8 - b.y0 - .25); if (p.bTop) p.before = Math.max(0, p.before - p.bTop - .25); }
    if (!(nx && nx.box === b)) { p.bBot = cl(b.y1 - (p.lastY + p.lead * .2) - .25); if (p.bBot && nx && nx.before !== undefined && !nx.pb) nx.before = Math.max(0, nx.before - p.bBot - .25); }
  });
  prog(.97);
  return { paras: S.paras, pw, ph, lm, rm, top: topPt, bot: botPt, hdr: hdrP, ftr: ftrP };
}
async function pdfToDocx(pdf, prog) { const L = await pdfLayout(pdf, prog); return makeRichDocx(L.paras, L); }

function layoutText(L) {
  const cellTxt = c => c.paras.filter(p => p.runs).map(p => p.runs.map(r => r.tab ? "\t" : r.text).join("")).join(" ");
  let out = "", prevList = false;
  for (const p of L.paras) {
    let t; const isList = p.textX != null || (p.runs && p.runs[1] && p.runs[1].tab && /^(?:[•◦▪–\-]|\d+[.)])$/.test(p.runs[0].text));
    if (p.tbl) t = p.tbl.rows.map(r => r.cells.filter(c => c.vm !== "cont").map(cellTxt).join("\t")).join("\n");
    else if (p.spacer || p.img) continue; else t = p.runs.map(r => r.tab ? "\t" : r.text).join("");
    if (!t.trim()) continue;
    out += (out ? (isList && prevList ? "\n" : "\n\n") : "") + t; prevList = isList;
  }
  return out;
}
function layoutHtml(L) {
  const runH = r => { if (r.tab) return "&emsp;"; let h = esc(r.text); const st = `font-family:'${r.fam}',sans-serif;font-size:${r.sz / 2}pt;color:#${r.color}${r.shd ? ";background:#" + r.shd : ""}`;
    h = `<span style="${st}">${h}</span>`; if (r.b) h = `<b>${h}</b>`; if (r.i) h = `<i>${h}</i>`; if (r.u) h = `<u>${h}</u>`; return h; };
  const pH = p => `<p style="margin:${Math.round(Math.min(p.before || 0, 40) * 10) / 10}pt 0 0 ${Math.round(Math.max(0, p.ind || 0) * 10) / 10}pt;text-align:${p.al === "both" ? "justify" : p.al || "left"}">${p.runs.map(runH).join("")}</p>`;
  return htmlDoc(L.paras.map(p => {
    if (p.tbl) return "<table>" + p.tbl.rows.map(r => "<tr>" + r.cells.filter(c => c.vm !== "cont").map(c => `<td colspan="${c.c.cs}"${c.c.fill ? ` style="background:#${c.c.fill}"` : ""}>${c.paras.filter(q => q.runs).map(q => q.runs.map(runH).join("")).join("<br>")}</td>`).join("") + "</tr>").join("") + "</table>";
    if (p.spacer || p.img) return ""; return pH(p);
  }).join("\n"));
}



/* ---------- Rendu PDF vectoriel (texte sélectionnable) à partir des blocs riches ---------- */
function blocksToPdf(paras) {
  const doc = new jspdf.jsPDF({ unit: "pt", format: "a4" }), PW = 595.28, PH = 841.89, M = 56.7, CW = PW - 2 * M, BOT = PH - M;
  let y = M;
  const fam = f => /courier|mono|consolas/i.test(f || "") ? "courier" : /times|serif|cambria|georgia|caladea|garamond|palatino/i.test(f || "") && !/sans/i.test(f) ? "times" : "helvetica";
  const sty = r => r.b && r.i ? "bolditalic" : r.b ? "bold" : r.i ? "italic" : "normal";
  const OK = /[\u2013\u2014\u2018\u2019\u201c\u201d\u2022\u2026\u20ac\u0152\u0153\u0160\u0161\u0178\u017d\u017e\u2020\u2021\u2030\u2122]/;
  const fix = t => t.replace(/[\u2190-\u21ff]/g, "->").replace(/[◦○]/g, "o").replace(/[▪■●➢❖◻]/g, "•").replace(/✓/g, "v").replace(/\u00a0/g, " ").replace(/[\u0100-\uffff]/g, ch => OK.test(ch) ? ch : "?");
  const setF = r => { doc.setFont(fam(r.fam), sty(r)); doc.setFontSize((r.sz || 22) / 2); };
  const rgb = h => [0, 2, 4].map(i => parseInt((h || "000000").slice(i, i + 2), 16));
  const page = () => { doc.addPage(); y = M; };
  function layout(p, width) {   // coupe les runs en lignes {it:[{x,w,t,r}], w, h}
    const ind = p.ind || 0, hang = p.hang || 0, lines = []; let line, x, tabI = 0;
    const nl = () => { line = { it: [], w: 0, h: 0 }; lines.push(line); x = lines.length === 1 ? Math.max(0, ind - hang) : ind; };
    nl();
    for (const r of p.runs || []) {
      if (r.br) { nl(); continue; }
      if (r.tab) { const st = (p.tabs || [])[tabI++]; x = Math.max(x + 1, st != null ? st : x + 36); continue; }
      setF(r); const sz = (r.sz || 22) / 2;
      for (const w of fix(r.text).split(/(\s+)/)) {
        if (!w) continue; const tw = doc.getTextWidth(w), sp = /^\s+$/.test(w);
        if (sp) { if (line.it.some(i => !i.sp)) { line.it.push({ x, w: tw, t: w, r, sp: 1 }); x += tw; } continue; }
        if (x + tw > width + .01 && line.it.some(i => !i.sp)) { while (line.it.length && line.it[line.it.length - 1].sp) line.it.pop(); nl(); }
        line.it.push({ x, w: tw, t: w, r }); x += tw; line.h = Math.max(line.h, sz * 1.2);
      }
    }
    for (const l of lines) { while (l.it.length && l.it[l.it.length - 1].sp) l.it.pop(); const e = l.it[l.it.length - 1]; l.w = e ? e.x + e.w : 0; if (!l.h) l.h = ((p.runs || [{ sz: 22 }])[0].sz || 22) / 2 * 1.2; }
    return lines;
  }
  function draw(lines, left, width, top, al) {   // dessine à partir de top, renvoie le bas
    let yy = top;
    for (const l of lines) {
      const off = al === "center" ? (width - l.w) / 2 : al === "right" ? width - l.w : 0, base = yy + l.h * .8;
      const segs = [];
      for (const i of l.it) { const s = segs[segs.length - 1]; if (s && s.r === i.r && Math.abs(i.x - (s.x + s.w)) < 1) { s.t += i.t; s.w += i.w; } else segs.push({ r: i.r, x: i.x, w: i.w, t: i.t }); }
      for (const i of segs) {
        setF(i.r); const X = left + off + i.x;
        if (i.r.shd) { doc.setFillColor(...rgb(i.r.shd)); doc.rect(X, yy, i.w, l.h, "F"); }
        doc.setTextColor(...rgb(i.r.color)); doc.text(i.t, X, base);
        if (i.r.u) { doc.setDrawColor(...rgb(i.r.color)); doc.setLineWidth(.5); doc.line(X, base + 1.5, X + i.w, base + 1.5); }
      }
      yy += l.h;
    }
    return yy;
  }
  const hOf = ls => ls.reduce((a, l) => a + l.h, 0);
  function table(t) {
    const X0 = Math.min(...t.rows.flatMap(r => r.cells.map(c => c.c.x0))), TW = Math.max(...t.rows.flatMap(r => r.cells.map(c => c.c.x1))) - X0, sc = Math.min(1, CW / TW), pad = Math.min(t.pad || 5, 8);
    for (const row of t.rows) {
      const cs = row.cells.filter(c => c.vm !== "cont").map(c => {
        const w = (c.c.x1 - c.c.x0) * sc, L = M + (c.c.x0 - X0) * sc;
        const ps = c.paras.filter(p => p.runs).map(p => ({ p, ls: layout(p, w - 2 * pad - (p.ind || 0) + (p.ind || 0)) }));
        return { c, w, L, ps, h: ps.reduce((a, q) => a + hOf(q.ls) + Math.min(q.p.before || 0, 6), 0) + 4 };
      });
      const rh = Math.max(row.h * 0.8, ...cs.map(c => c.h));
      if (y + rh > BOT) page();
      for (const q of cs) {
        if (q.c.c.fill) { doc.setFillColor(...rgb(q.c.c.fill)); doc.rect(q.L, y, q.w, rh, "F"); }
        let yy = y + 2; for (const z of q.ps) { yy += Math.min(z.p.before || 0, 6); yy = draw(z.ls, q.L + pad + (z.p.ind || 0) * 0, q.w - 2 * pad, yy, z.p.al); }
        for (const [k, b] of [["bt", [q.L, y, q.L + q.w, y]], ["bb", [q.L, y + rh, q.L + q.w, y + rh]], ["bl", [q.L, y, q.L, y + rh]], ["br", [q.L + q.w, y, q.L + q.w, y + rh]]]) {
          const bd = q.c.c[k]; if (bd) { doc.setDrawColor(...rgb(bd.col)); doc.setLineWidth(Math.max(.25, (bd.w || 4) / 8)); doc.line(...b); }
        }
      }
      y += rh;
    }
    y += 6;
  }
  for (const p of paras) {
    if (p.tbl) { table(p.tbl); continue; }
    if (p.spacer) { y += p.spacer; if (p.bdr) { doc.setDrawColor(...rgb(p.bdr.col)); doc.setLineWidth(p.bdr.sz / 8); doc.line(M, y, PW - M, y); } continue; }
    if (p.img) {
      let w = Math.min(p.img.x1 - p.img.x0, CW), h = (p.img.y1 - p.img.y0) * w / (p.img.x1 - p.img.x0); if (h > BOT - M) { w *= (BOT - M) / h; h = BOT - M; }
      y += p.before || 0; if (y + h > BOT) page(); doc.addImage(p.img.data, "JPEG", M + (p.ind || 0), y, w, h); y += h; continue;
    }
    if (!p.runs || !p.runs.length) continue;
    const ls = layout(p, CW - (p.al === "left" || !p.al ? 0 : 0)); if (y > M) y += Math.min(p.before || 0, 40);
    let ly = 0; // pagination ligne par ligne
    let rest = ls.slice(), first = true;
    while (rest.length) {
      if (y + rest[0].h > BOT) { page(); }
      const fit = []; let acc = 0; while (rest.length && y + acc + rest[0].h <= BOT) { acc += rest[0].h; fit.push(rest.shift()); }
      y = draw(fit, M, CW, y, p.al); first = false;
    }
  }
  return doc.output("blob");
}
const richPdf = blocks => blocksToPdf(blocks);

/* ---------- Texte / HTML / Markdown : rendu structuré ---------- */
function htmlToText(html) {
  const box = document.createElement("div");
  box.style.cssText = "position:fixed;left:-9999px;top:0;width:600px"; box.innerHTML = clean(html); document.body.append(box);
  const t = box.innerText; box.remove(); return t;
}
function mdToHtml(md) {
  const inl = s => esc(s).replace(/`([^`]+)`/g, "<code>$1</code>").replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, (m, a, b) => `<b>${a || b}</b>`)
    .replace(/(^|[^*])\*([^*\n]+)\*/g, "$1<i>$2</i>").replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  const out = []; let list = null, code = false, para = [];
  const flushP = () => { if (para.length) out.push("<p>" + para.map(inl).join("<br>") + "</p>"); para = []; };
  const flushL = () => { if (list) out.push(`</${list}>`); list = null; };
  for (const line of md.split(/\r?\n/)) {
    if (/^```/.test(line)) { flushP(); flushL(); out.push(code ? "</pre>" : "<pre>"); code = !code; continue; }
    if (code) { out.push(esc(line) + "\n"); continue; }
    let m;
    if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) { flushP(); flushL(); out.push(`<h${m[1].length}>${inl(m[2])}</h${m[1].length}>`); }
    else if ((m = /^\s*([-*+]|\d+[.)])\s+(.*)$/.exec(line))) { flushP(); const t = /\d/.test(m[1]) ? "ol" : "ul"; if (list !== t) { flushL(); out.push(`<${t}>`); list = t; } out.push(`<li>${inl(m[2])}</li>`); }
    else if (/^\s*$/.test(line)) { flushP(); flushL(); }
    else if (/^---+\s*$/.test(line)) { flushP(); flushL(); out.push("<hr>"); }
    else { flushL(); para.push(line); }
  }
  flushP(); flushL(); return out.join("\n");
}
const HTML_CSS = "<style>body{font-family:Arial,sans-serif;max-width:800px;margin:24px auto;line-height:1.5}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px 8px}pre,code{background:#eee}</style>";
const htmlDoc = body => "<!DOCTYPE html><meta charset=utf-8>" + HTML_CSS + body;

// HTML → blocs Word (titres, gras/italique/souligné, listes, tableaux)
async function htmlToBlocks(html) {
  const d = new DOMParser().parseFromString(clean(html), "text/html"), paras = [];
  for (const im of d.querySelectorAll("img")) {   // images (ex. issues de Word) → JPEG intégré
    try {
      const i = await new Promise((r, j) => { const x = new Image(); x.onload = () => r(x); x.onerror = j; x.src = im.getAttribute("src"); });
      const c = document.createElement("canvas"); c.width = i.naturalWidth; c.height = i.naturalHeight; const cx = c.getContext("2d"); cx.fillStyle = "#fff"; cx.fillRect(0, 0, c.width, c.height); cx.drawImage(i, 0, 0);
      const w = Math.min(481, i.naturalWidth * .75); im.__img = { data: new Uint8Array(await (await toBlob(c, "image/jpeg", .9)).arrayBuffer()), x0: 0, y0: 0, x1: w, y1: w * i.naturalHeight / i.naturalWidth };
    } catch (e) {}
  }
  const H = { H1: 20, H2: 16, H3: 13.5, H4: 12, H5: 11, H6: 10 };
  const run = (text, st) => ({ text, fam: st.mono ? "Courier New" : "Arial", b: st.b, i: st.i, u: st.u, sz: Math.round((st.size || 11) * 2), color: st.color || "000000" });
  const inline = (node, st, runs) => {
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { const t = c.textContent.replace(/\s+/g, " "); if (t) runs.push(run(t, st)); }
      else if (c.nodeType === 1) {
        const tg = c.tagName, s2 = { ...st };
        if (tg === "B" || tg === "STRONG" || /^H\d$/.test(tg)) s2.b = true; if (tg === "I" || tg === "EM") s2.i = true; if (tg === "U") s2.u = true;
        if (tg === "A") { s2.u = true; s2.color = "0563C1"; } if (tg === "CODE" || tg === "PRE") s2.mono = true;
        if (c.style && c.style.color) { const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(c.style.color); if (m) s2.color = [m[1], m[2], m[3]].map(v => (+v).toString(16).padStart(2, "0")).join("").toUpperCase(); }
        if (tg === "BR") runs.push({ text: "\n", br: 1 }); else if (!/^(UL|OL|TABLE)$/.test(tg)) inline(c, s2, runs);
      }
    }
  };
  const push = (runs, o = {}) => {
    while (runs.length && !runs[0].br && !runs[0].text.trim()) runs.shift(); if (runs[0] && runs[0].text) runs[0].text = runs[0].text.replace(/^\s+/, "");
    if (!runs.length) return; const last = runs[runs.length - 1]; if (last.text) last.text = last.text.replace(/\s+$/, "");
    const sz = Math.max(...runs.map(r => r.sz || 22));
    paras.push({ runs, tabs: o.tabs || [], al: o.al || "left", lines: 1, lead: sz / 2 * 1.2, before: o.before == null ? 4 : o.before, ind: o.ind || 0, hang: o.hang || 0, keep: o.keep });
  };
  const block = (node, ctx) => {
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { if (c.textContent.trim()) { const r = []; inline({ childNodes: [c] }, ctx.st, r); push(r); } continue; }
      if (c.nodeType !== 1) continue;
      const tg = c.tagName;
      if (tg === "IMG") { if (c.__img) paras.push({ img: c.__img, before: 4, ind: 0, maxW: 481 }); }
      else if (H[tg]) { const r = []; inline(c, { ...ctx.st, b: true, size: H[tg] }, r); push(r, { before: 10, keep: true }); }
      else if (tg === "UL" || tg === "OL") {
        let n = 0; for (const li of c.children) {
          if (li.tagName !== "LI") continue; n++; const r = []; const subs = [];
          for (const k of li.childNodes) { if (k.nodeType === 1 && /^(UL|OL)$/.test(k.tagName)) subs.push(k); else inline({ childNodes: [k] }, ctx.st, r); }
          const mk = tg === "OL" ? n + "." : ctx.lvl ? "◦" : "•", ind = 18 + 18 * (ctx.lvl || 0);
          r.unshift({ ...run(mk, ctx.st) }, { tab: true }); push(r, { ind, hang: 14, tabs: [ind], before: 2 });
          subs.forEach(sb => block({ childNodes: [sb] }, { ...ctx, lvl: (ctx.lvl || 0) + 1 }));
        }
      }
      else if (tg === "TABLE") {
        const rows = [...c.querySelectorAll("tr")]; if (!rows.length) continue;
        const nc = Math.max(...rows.map(r => r.children.length)), W = 480, cw = W / nc, bd = { w: 4, col: "808080" };
        paras.push({ tbl: { widths: Array(nc).fill(cw), x0: 56.7, pad: 5, rows: rows.map(tr => ({ h: 16, cells: [...tr.children].map((td, k) => {
          const r = []; inline(td, { ...ctx.st, b: td.tagName === "TH" || ctx.st.b }, r); const ps = []; const keep = paras.length;
          const sz = Math.max(...r.map(x => x.sz || 22)); if (r.some(x => x.text && x.text.trim())) ps.push({ runs: r.map(x => ({ ...x })), tabs: [], al: "left", lines: 1, lead: sz / 2 * 1.2, before: 2, ind: 0, hang: 0 });
          return { c: { x0: k * cw, x1: (k + 1) * cw, cs: 1, rs: 1, bt: bd, bb: bd, bl: bd, br: bd, fill: td.tagName === "TH" ? "E7E6E6" : null }, paras: ps, va: null, vm: null };
        }) })) }, top: 0, bottom: 0 });
        paras.push({ spacer: 6 });
      }
      else if (tg === "HR") paras.push({ spacer: 4, bdr: { sz: 6, col: "808080", space: 0 } });
      else if (/^(P|DIV|SECTION|ARTICLE|BLOCKQUOTE|PRE|FIGURE|HEADER|FOOTER|MAIN|BODY)$/.test(tg) && c.querySelector("p,div,ul,ol,table,h1,h2,h3,h4,h5,h6")) block(c, ctx);
      else if (tg === "PRE") { c.textContent.split("\n").forEach(l => push([run(l || " ", { mono: true, size: 10 })], { before: 0 })); }
      else { const r = []; inline(c, ctx.st, r); push(r, tg === "BLOCKQUOTE" ? { ind: 24 } : {}); c.querySelectorAll("img").forEach(im => { if (im.__img) paras.push({ img: im.__img, before: 4, ind: 0, maxW: 481 }); }); }
    }
  };
  block(d.body, { st: { size: 11 }, lvl: 0 });
  return paras;
}
const richDocx = paras => makeRichDocx(paras, { pw: 595.3, ph: 841.9, lm: 56.7, rm: 538.6, top: 56.7, bot: 56.7 });
function rowsToBlocks(rows) {   // tableau de lignes → bloc table Word
  const nc = Math.max(...rows.map(r => r.length), 1), cw = 480 / nc, bd = { w: 4, col: "808080" };
  const cell = (v, k, hdr) => ({ c: { x0: k * cw, x1: (k + 1) * cw, cs: 1, rs: 1, bt: bd, bb: bd, bl: bd, br: bd, fill: hdr ? "D9E2F3" : null },
    paras: v === "" || v == null ? [] : [{ runs: [{ text: String(v), fam: "Arial", b: hdr, i: false, sz: 20, color: "000000" }], tabs: [], al: "left", lines: 1, lead: 12, before: 2, ind: 0, hang: 0 }], va: null, vm: null });
  return [{ tbl: { widths: Array(nc).fill(cw), x0: 56.7, pad: 5, rows: rows.map((r, i) => ({ h: 16, cells: Array.from({ length: nc }, (_, k) => cell(r[k], k, i === 0)) })) }, top: 0, bottom: 0 }, { spacer: 4 }];
}

/* ---------- Enregistrement audio/vidéo (conversion de conteneur en temps réel) ---------- */
async function recordMedia(f, k, mime, prog) {
  let lastErr;
  for (let attempt = 0; attempt < 3; attempt++) {
    const el = document.createElement(k), url = URL.createObjectURL(f);
    try {
      el.src = url; el.playsInline = true; el.preload = "auto";
      await new Promise((r, j) => { el.onloadedmetadata = r; el.onerror = () => j(new Error("Le navigateur ne peut pas lire ce fichier (format ou codec non pris en charge).")); });
      await el.play();                                   // démarre le décodage pour que les pistes existent
      const cap = () => el.captureStream ? el.captureStream() : el.mozCaptureStream();
      let stream = cap();
      for (let t = 0; t < 30 && k === "video" && !stream.getVideoTracks().length; t++) { await new Promise(r => setTimeout(r, 100)); stream = cap(); }
      if (k === "video" && !stream.getVideoTracks().length) throw new Error("Aucune piste vidéo lisible dans ce fichier.");
      el.pause(); el.currentTime = 0; await new Promise(r => { el.onseeked = r; setTimeout(r, 500); });
      const rec = new MediaRecorder(stream, { mimeType: mime }), chunks = [], dur = isFinite(el.duration) ? el.duration : 600;
      rec.ondataavailable = ev => ev.data.size && chunks.push(ev.data);
      const done = new Promise(r => rec.onstop = r), ended = new Promise(r => el.onended = r);
      rec.start(); await el.play(); el.ontimeupdate = () => prog(Math.min(1, el.currentTime / dur));
      await Promise.race([ended, new Promise(r => setTimeout(r, dur * 1500 + 8000))]);
      await new Promise(r => setTimeout(r, 120)); rec.stop(); await done; el.pause();
      const blob = new Blob(chunks, { type: mime });
      if (blob.size >= 1500) return blob;
      lastErr = new Error("Enregistrement vide : ce navigateur n'a pas pu décoder le fichier (codec non pris en charge ?).");
    } catch (e) { lastErr = e; if (/lire ce fichier|piste vidéo/.test(e.message)) throw e; }
    finally { el.removeAttribute("src"); URL.revokeObjectURL(url); }
  }
  throw lastErr;
}

async function convert(it, fmt, prog) {
  const f = it.file, k = it.kind, b = base(f), e = fmt.split("-")[0], out = [];
  const add = (blob, name = `${b}.${e}`) => out.push({ name, blob });

  if (k === "image") {
    const img = await loadImg(f), w = img.naturalWidth || img.width || 800, h = img.naturalHeight || img.height || 600;
    const c = drawCanvas(img, w, h, ["jpg","bmp","pdf","docx"].includes(fmt));
    if (fmt === "pdf") { const d = new jspdf.jsPDF({ orientation: w > h ? "l" : "p", unit: "px", format: [w, h] }); d.addImage(c.toDataURL("image/jpeg", .92), "JPEG", 0, 0, w, h); add(d.output("blob")); }
    else if (fmt === "docx") add(await makeDocx([await imgBlock(c)]));
    else if (fmt === "bmp") add(await bmp(c));
    else add(await toBlob(c, MIME[fmt]));
  }
  else if (k === "pdf") {
    let pdf;
    try {
      pdf = await pdfjsLib.getDocument({ data: await f.arrayBuffer() }).promise;
    } catch(err) {
      if (err.name === 'PasswordException') {
        throw new Error("Ce PDF est protégé par un mot de passe.");
      }
      throw new Error("Fichier PDF illisible ou endommagé.");
    }
    const n = pdf.numPages;

    if (MIME[fmt]) {
      for (let i = 1; i <= n; i++) { add(await toBlob(await renderPage(pdf, i, 2), MIME[fmt]), `${b}-p${i}.${fmt}`); prog(i / n); }
    } else if (fmt === "docx") {
      // Word éditable : police, gras/italique, taille, couleur, alignement, retraits, images
      add(await pdfToDocx(pdf, prog));
    } else if (fmt === "docx-img") {
      // Alternative : une image par page (aspect exact, texte non modifiable ; utile pour les PDF scannés)
      const bl = [];
      for (let i = 1; i <= n; i++) {
        bl.push(await imgBlock(await renderPage(pdf, i, 2)));
        if (i < n) bl.push({ br: 1 });
        prog(i / n);
      }
      add(await makeDocx(bl));
    } else {
      const L = await pdfLayout(pdf, prog);
      if (fmt === "txt") add(txtBlob(layoutText(L)));
      else if (fmt === "html") add(txtBlob(layoutHtml(L), "text/html"));
    }
  }
  else if (k === "docx") {
    const ab = await f.arrayBuffer();
    if (fmt === "txt") add(txtBlob((await mammoth.extractRawText({ arrayBuffer: ab })).value));
    else { const html = (await mammoth.convertToHtml({ arrayBuffer: ab })).value; add(fmt === "html" ? txtBlob(htmlDoc(html), "text/html") : richPdf(await htmlToBlocks(html))); }
  }
  else if (k === "sheet" || (k === "text" && (["csv","json","xlsx"].includes(fmt) || (ext(f) === "csv" && fmt === "docx")))) {
    let wb;
    if (k === "sheet") wb = XLSX.read(await f.arrayBuffer());
    else {
      const raw = await f.text();
      if (ext(f) === "json") { wb = XLSX.utils.book_new(); const j = JSON.parse(raw); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Array.isArray(j) ? j : [j]), "Feuille1"); }
      else wb = XLSX.read(raw, { type: "string" });
    }
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (fmt === "csv") add(txtBlob(XLSX.utils.sheet_to_csv(ws), "text/csv"));
    else if (fmt === "json") add(txtBlob(JSON.stringify(XLSX.utils.sheet_to_json(ws), null, 2), "application/json"));
    else if (fmt === "txt") add(txtBlob(XLSX.utils.sheet_to_csv(ws, { FS: "\t" })));
    else if (fmt === "docx") add(await richDocx(rowsToBlocks(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }))));
    else if (fmt === "html") add(txtBlob(htmlDoc(XLSX.utils.sheet_to_html(ws).replace(/^[\s\S]*?<body>|<\/body>[\s\S]*$/g, "")), "text/html"));
    else if (fmt === "pdf") add(richPdf(rowsToBlocks(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }))));
    else add(new Blob([XLSX.write(wb, { type: "array", bookType: "xlsx" })], { type: "application/octet-stream" }));
  }
  else if (k === "text") {
    let raw = await f.text(); const e0 = ext(f), isH = ["html","htm"].includes(e0), isMd = e0 === "md";
    if (e0 === "json") { try { raw = JSON.stringify(JSON.parse(raw), null, 2); } catch (er) {} }
    const html = isH ? raw : isMd ? mdToHtml(raw) : null, plain = isH ? htmlToText(raw) : raw;
    if (fmt === "pdf") add(html ? richPdf(await htmlToBlocks(html)) : textPdf(plain));
    else if (fmt === "docx") add(await richDocx(html ? await htmlToBlocks(html) : plain.split(/\r?\n/).map(t => ({ runs: [{ text: t || " ", fam: e0 === "json" || e0 === "csv" ? "Courier New" : "Arial", sz: 22, color: "000000" }], tabs: [], al: "left", lines: 1, lead: 13.2, before: 0, ind: 0, hang: 0 }))));
    else if (fmt === "txt") add(txtBlob(plain));
    else add(txtBlob(html ? (isH ? raw : htmlDoc(html)) : htmlDoc("<pre>" + esc(raw) + "</pre>"), "text/html"));
  }
  else { // video / audio
    if (fmt === "wav") add(wav(await decode(f)));
    else if (fmt === "mp3") add(mp3(await decode(f)));
    else if (MIME[fmt]) {
      const v = document.createElement("video"); v.muted = true; v.src = URL.createObjectURL(f);
      await new Promise((r, j) => { v.onloadedmetadata = r; v.onerror = () => j(new Error("Format vidéo invalide.")); }); 
      v.currentTime = Math.min(1, v.duration / 2);
      await new Promise(r => v.onseeked = r);
      add(await toBlob(drawCanvas(v, v.videoWidth, v.videoHeight, fmt === "jpg"), MIME[fmt]));
    } else {
      add(await recordMedia(f, k, (k === "audio" ? "audio/" : "video/") + MRX[fmt], prog));
    }
  }
  return out;
}

/* ---------- Interface ---------- */
function formats(k, f) {
  let l = [...OUT[k]];
  if (k === "text") {
    const e = ext(f);
    if (e === "csv") l.push("json", "xlsx"); if (e === "json") l.push("csv", "xlsx");
    l = l.filter(x => x !== (e === "htm" ? "html" : e));
  }
  return l.filter(x => !MRX[x] || (window.MediaRecorder && MediaRecorder.isTypeSupported((k === "audio" ? "audio/" : "video/") + MRX[x])));
}

function removeItem(id) {
  items = items.filter(it => it.id !== id);
  const el = document.querySelector(`[data-id="${id}"]`);
  if (el) el.remove();
  $("#actions").style.display = items.length ? "flex" : "none";
}

function add(files) {
  for (const f of files) {
    const k = kind(f), id = Date.now() + Math.random();
    const it = { id, file: f, kind: k, fmt: k && formats(k, f)[0] };
    items.push(it);
    
    const row = document.createElement("div"); 
    row.className = "row"; 
    row.dataset.id = id;
    it.row = row;

    if (!k) {
      row.innerHTML = `
        <div class="name">${esc(f.name)}</div>
        <div class="status err">Format non pris en charge (.${ext(f)})</div>
        <button class="btn-remove" title="Supprimer">✕</button>
      `;
    } else {
      row.innerHTML = `
        <div>
          <div class="name"></div>
          <div class="meta">${k} · ${size(f.size)}</div>
        </div>
        <span class="meta">→</span>
        <select>${formats(k, f).map(x => `<option value="${x}">${LBL[x] || x}</option>`).join("")}</select>
        <button class="btn-convert">Convertir</button>
        <button class="btn-remove" title="Supprimer">✕</button>
        <div class="status"></div>
      `;
      row.querySelector(".name").textContent = f.name;
      row.querySelector("select").onchange = e => it.fmt = e.target.value;
      row.querySelector(".btn-convert").onclick = () => run(it);
    }
    row.querySelector(".btn-remove").onclick = () => removeItem(id);
    $("#list").append(row);
  }
  $("#actions").style.display = items.length ? "flex" : "none";
}

async function run(it) {
  if (!it.kind) return;
  const btn = it.row.querySelector(".btn-convert"), st = it.row.querySelector(".status");
  btn.disabled = true; st.className = "status"; st.innerHTML = 'Conversion…<div class="bar"><i></i></div>';
  const bar = st.querySelector("i");
  try {
    const res = await convert(it, it.fmt, p => bar.style.width = Math.round(p * 100) + "%");
    st.innerHTML = '<div class="dl-group"></div>';
    const dlGroup = st.querySelector(".dl-group");
    
    res.forEach(r => {
      const a = document.createElement("a"); 
      a.className = "btn-action btn-download"; 
      a.href = URL.createObjectURL(r.blob);
      a.download = r.name; 
      a.innerHTML = "⬇ Télécharger " + r.name + " (" + size(r.blob.size) + ")"; 
      dlGroup.append(a);
    });

    if (res.length > 1) {
      const a = document.createElement("a"); 
      a.className = "btn-action"; 
      a.textContent = "📦 Tout en ZIP"; 
      a.href = "#";
      a.onclick = async ev => {
        if (a.dataset.ok) return; ev.preventDefault();
        const z = new JSZip(); res.forEach(r => z.file(r.name, r.blob));
        a.href = URL.createObjectURL(await z.generateAsync({ type: "blob" })); 
        a.download = base(it.file) + ".zip"; 
        a.dataset.ok = 1; 
        a.click();
      };
      dlGroup.append(a);
    }
  } catch (e) {
    let msg = e.message || e;
    if (msg.includes("out of memory") || msg.includes("Allocation")) {
      msg = "Fichier trop volumineux pour la mémoire disponible du navigateur.";
    }
    st.className = "status err"; 
    st.textContent = "Erreur : " + msg;
  }
  btn.disabled = false;
}

/* ---------- Événements ---------- */

window.addEventListener("dragover", e => {
  e.preventDefault();
  document.body.classList.add("drag-over");
});

window.addEventListener("dragleave", e => {
  if (e.relatedTarget === null) {
    document.body.classList.remove("drag-over");
  }
});

window.addEventListener("drop", e => {
  e.preventDefault();
  document.body.classList.remove("drag-over");
  if (e.dataTransfer.files.length) {
    add(e.dataTransfer.files);
  }
});

window.addEventListener("paste", e => {
  const pastedFiles = [];
  if (e.clipboardData && e.clipboardData.items) {
    for (const item of e.clipboardData.items) {
      if (item.kind === "file") {
        pastedFiles.push(item.getAsFile());
      }
    }
  }
  if (pastedFiles.length) {
    add(pastedFiles);
  }
});

const drop = $("#drop"), input = $("#file");
drop.onclick = () => input.click();
input.onchange = () => { add(input.files); input.value = ""; };

$("#all").onclick = async () => { for (const it of items) if (it.kind) await run(it); };
$("#clear").onclick = () => { items.length = 0; $("#list").innerHTML = ""; $("#actions").style.display = "none"; };