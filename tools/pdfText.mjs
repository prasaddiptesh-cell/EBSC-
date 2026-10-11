// v3.9.3 Reading the text of a PDF (with where each piece sits on the page), on the device, no internet needed.
// Used for the FCCC Master Price List. Built on pdf-lib (already in the app) to open the file and unpack its streams;
// the page drawing commands are read here to find each piece of text, its x / y position and size.
// Handles the usual fonts made by Excel / Word / LibreOffice / browsers: simple fonts (WinAnsi, with or without
// Differences) and Type0 / Identity-H fonts, both through the font's ToUnicode table when it has one.
// Returns pages: [{ items: [{ x, y, size, str }] }] – never throws for a bad page (that page is just empty).
import { PDFDocument, PDFName, PDFArray, PDFDict, PDFRawStream, PDFRef, PDFNumber, decodePDFRawStream } from 'pdf-lib';

const WIN = { 128: '€', 130: '‚', 131: 'ƒ', 132: '„', 133: '…', 134: '†', 135: '‡', 136: 'ˆ', 137: '‰', 138: 'Š', 139: '‹', 140: 'Œ', 142: 'Ž', 145: '‘', 146: '’', 147: '“', 148: '”', 149: '•', 150: '–', 151: '—', 152: '˜', 153: '™', 154: 'š', 155: '›', 156: 'œ', 158: 'ž', 159: 'Ÿ' };
const winChar = (c) => WIN[c] || String.fromCharCode(c);
// Glyph names used in /Differences (the common ones).
const GLYPH = { space: ' ', period: '.', comma: ',', hyphen: '-', slash: '/', parenleft: '(', parenright: ')', percent: '%', ampersand: '&', colon: ':', semicolon: ';', quoteright: '’', quotesingle: "'", dollar: '$',
  zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9', endash: '–', emdash: '—', plus: '+', asterisk: '*', numbersign: '#', at: '@', underscore: '_', bullet: '•' };
const glyphChar = (n) => (GLYPH[n] !== undefined ? GLYPH[n] : n.length === 1 ? n : /^uni([0-9A-F]{4})$/i.test(n) ? String.fromCharCode(parseInt(n.slice(3), 16)) : '');

const bytesToLatin1 = (b) => { let s = ''; for (let i = 0; i < b.length; i += 8192) s += String.fromCharCode.apply(null, Array.from(b.subarray(i, i + 8192))); return s; };

function decodeStream(ctx, obj) {
  const o = obj instanceof PDFRef ? ctx.lookup(obj) : obj;
  if (!o) return '';
  if (o instanceof PDFRawStream) return bytesToLatin1(decodePDFRawStream(o).decode());
  if (o.getContents) return bytesToLatin1(o.getContents());
  return '';
}

// ---------------------------------------------------------------- ToUnicode CMap
function parseCMap(text) {
  const map = new Map();
  let codeLen = 1;
  const cs = /begincodespacerange([\s\S]*?)endcodespacerange/.exec(text);
  if (cs) { const m = /<([0-9a-fA-F]+)>/.exec(cs[1]); if (m) codeLen = Math.max(1, m[1].length / 2); }
  const hexStr = (h) => { let s = ''; for (let i = 0; i + 3 < h.length + 1; i += 4) { const c = parseInt(h.slice(i, i + 4), 16); if (!Number.isNaN(c)) s += String.fromCharCode(c); } return h.length === 2 ? String.fromCharCode(parseInt(h, 16)) : s; };
  let m;
  const charRe = /beginbfchar([\s\S]*?)endbfchar/g;
  while ((m = charRe.exec(text))) {
    const pr = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]*)>/g; let p;
    while ((p = pr.exec(m[1]))) map.set(parseInt(p[1], 16), hexStr(p[2]));
  }
  const rangeRe = /beginbfrange([\s\S]*?)endbfrange/g;
  while ((m = rangeRe.exec(text))) {
    const pr = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<([0-9a-fA-F]*)>|\[([^\]]*)\])/g; let p;
    while ((p = pr.exec(m[1]))) {
      const lo = parseInt(p[1], 16); const hi = parseInt(p[2], 16);
      if (hi - lo > 65535) continue;
      if (p[4] !== undefined) {
        const base = parseInt(p[4].slice(-4) || '0', 16); const pre = p[4].length > 4 ? hexStr(p[4].slice(0, -4)) : '';
        for (let c = lo; c <= hi; c++) map.set(c, pre + String.fromCharCode(base + (c - lo)));
      } else {
        const list = (p[5].match(/<([0-9a-fA-F]*)>/g) || []).map((x) => hexStr(x.slice(1, -1)));
        for (let c = lo; c <= hi && c - lo < list.length; c++) map.set(c, list[c - lo]);
      }
    }
  }
  return { map, codeLen };
}

function fontInfo(ctx, fontObj) {
  const f = fontObj instanceof PDFRef ? ctx.lookup(fontObj) : fontObj;
  const info = { two: false, map: null, diff: null, widths: null, dw: 500 };
  if (!(f instanceof PDFDict)) return info;
  const num = (v) => { const x = v instanceof PDFRef ? ctx.lookup(v) : v; return x instanceof PDFNumber ? x.asNumber() : null; };
  try {
    // Glyph widths (1/1000 of the size), so the end of each piece of text is known.
    const fc = num(f.get(PDFName.of('FirstChar')));
    let wa = f.get(PDFName.of('Widths')); if (wa instanceof PDFRef) wa = ctx.lookup(wa);
    if (wa instanceof PDFArray && fc !== null) { info.widths = new Map(); for (let i = 0; i < wa.size(); i++) info.widths.set(fc + i, num(wa.get(i)) || 0); }
    let desc = f.get(PDFName.of('DescendantFonts')); if (desc instanceof PDFRef) desc = ctx.lookup(desc);
    if (desc instanceof PDFArray && desc.size()) {
      let d0 = desc.get(0); if (d0 instanceof PDFRef) d0 = ctx.lookup(d0);
      if (d0 instanceof PDFDict) {
        const dw = num(d0.get(PDFName.of('DW'))); if (dw !== null) info.dw = dw;
        let w = d0.get(PDFName.of('W')); if (w instanceof PDFRef) w = ctx.lookup(w);
        if (w instanceof PDFArray) {
          info.widths = new Map(); let i = 0;
          while (i < w.size()) {
            const a = num(w.get(i)); let b = w.get(i + 1); if (b instanceof PDFRef) b = ctx.lookup(b);
            if (b instanceof PDFArray) { for (let k = 0; k < b.size(); k++) info.widths.set(a + k, num(b.get(k)) || 0); i += 2; } else { const hi = num(b); const ww = num(w.get(i + 2)); for (let c = a; c <= hi && c - a < 65536; c++) info.widths.set(c, ww || 0); i += 3; }
          }
        }
      }
    }
  } catch (e) { info.widths = null; }
  const sub = f.get(PDFName.of('Subtype'));
  info.two = !!(sub && String(sub) === '/Type0');
  const tu = f.get(PDFName.of('ToUnicode'));
  if (tu) { try { const c = parseCMap(decodeStream(ctx, tu)); info.map = c.map; if (c.codeLen === 2) info.two = true; } catch (e) { /* fall back */ } }
  let enc = f.get(PDFName.of('Encoding'));
  if (enc instanceof PDFRef) enc = ctx.lookup(enc);
  if (enc instanceof PDFDict) {
    const d = enc.get(PDFName.of('Differences'));
    const arr = d instanceof PDFRef ? ctx.lookup(d) : d;
    if (arr instanceof PDFArray) {
      info.diff = new Map(); let code = 0;
      for (let i = 0; i < arr.size(); i++) { const v = arr.get(i); if (v instanceof PDFNumber) code = v.asNumber(); else { info.diff.set(code, glyphChar(String(v).replace(/^\//, ''))); code += 1; } }
    }
  }
  return info;
}
// Width of the raw string in text space units (before the font size), 1/1000 em per glyph.
function rawWidth(raw, font) {
  let w = 0; let n = 0;
  if (font && font.two) { for (let i = 0; i + 1 < raw.length; i += 2) { const c = (raw.charCodeAt(i) << 8) | raw.charCodeAt(i + 1); w += font.widths && font.widths.has(c) ? font.widths.get(c) : font.dw; n++; } return { w: w / 1000, n }; }
  for (let i = 0; i < raw.length; i++) { const c = raw.charCodeAt(i); w += font && font.widths && font.widths.has(c) ? font.widths.get(c) : 500; n++; }
  return { w: w / 1000, n };
}
function decodeText(raw, font) {
  if (!font) return Array.from(raw).map((ch) => winChar(ch.charCodeAt(0))).join('');
  let out = '';
  if (font.two) {
    for (let i = 0; i + 1 < raw.length; i += 2) { const c = (raw.charCodeAt(i) << 8) | raw.charCodeAt(i + 1); out += font.map && font.map.has(c) ? font.map.get(c) : ''; }
    return out;
  }
  for (let i = 0; i < raw.length; i++) {
    const c = raw.charCodeAt(i);
    out += font.map && font.map.has(c) ? font.map.get(c) : font.diff && font.diff.has(c) ? font.diff.get(c) : winChar(c);
  }
  return out;
}

// ---------------------------------------------------------------- content stream tokens
const WS = /[\0\t\n\f\r ]/;
const DELIM = /[()<>[\]{}/%]/;
function* tokens(s) {
  let i = 0; const n = s.length;
  while (i < n) {
    const ch = s[i];
    if (WS.test(ch)) { i++; continue; }
    if (ch === '%') { while (i < n && s[i] !== '\n' && s[i] !== '\r') i++; continue; }
    if (ch === '(') {
      let depth = 1; let out = ''; i++;
      while (i < n && depth > 0) {
        const c = s[i];
        if (c === '\\') {
          const e = s[i + 1];
          const esc = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
          if (esc[e] !== undefined) { out += esc[e]; i += 2; } else if (/[0-7]/.test(e)) { let o = ''; let k = i + 1; while (k < n && o.length < 3 && /[0-7]/.test(s[k])) o += s[k++]; out += String.fromCharCode(parseInt(o, 8) & 255); i = k; } else if (e === '\r' || e === '\n') { i += 2; if (e === '\r' && s[i] === '\n') i++; } else { out += e || ''; i += 2; }
          continue;
        }
        if (c === '(') depth++;
        if (c === ')') { depth--; if (depth === 0) { i++; break; } }
        out += c; i++;
      }
      yield { t: 'str', v: out }; continue;
    }
    if (ch === '<' && s[i + 1] === '<') { yield { t: 'op', v: '<<' }; i += 2; continue; }
    if (ch === '>' && s[i + 1] === '>') { yield { t: 'op', v: '>>' }; i += 2; continue; }
    if (ch === '<') {
      let k = i + 1; let h = '';
      while (k < n && s[k] !== '>') { if (/[0-9a-fA-F]/.test(s[k])) h += s[k]; k++; }
      if (h.length % 2) h += '0';
      let out = ''; for (let j = 0; j < h.length; j += 2) out += String.fromCharCode(parseInt(h.slice(j, j + 2), 16));
      yield { t: 'str', v: out }; i = k + 1; continue;
    }
    if (ch === '[') { yield { t: '[' }; i++; continue; }
    if (ch === ']') { yield { t: ']' }; i++; continue; }
    if (ch === '/') { let k = i + 1; while (k < n && !WS.test(s[k]) && !DELIM.test(s[k])) k++; yield { t: 'name', v: s.slice(i + 1, k) }; i = k; continue; }
    if (/[0-9+\-.]/.test(ch)) { let k = i + 1; while (k < n && /[0-9.]/.test(s[k])) k++; yield { t: 'num', v: Number(s.slice(i, k)) || 0 }; i = k; continue; }
    let k = i; while (k < n && !WS.test(s[k]) && !DELIM.test(s[k])) k++;
    if (k === i) { i++; continue; }
    const word = s.slice(i, k); i = k;
    if (word === 'BI') { const end = s.indexOf('EI', i); i = end < 0 ? n : end + 2; continue; } // inline image: skip
    yield { t: 'op', v: word };
  }
}

const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];

function readContent(content, fonts, ctx, xobjects, depth = 0) {
  const items = [];
  let stack = [];
  let arr = null;
  let ctm = [1, 0, 0, 1, 0, 0];
  const saved = [];
  let tm = [1, 0, 0, 1, 0, 0]; let tlm = [1, 0, 0, 1, 0, 0];
  let font = null; let size = 10; let leading = 0; let tc = 0; let tw = 0; let th = 1;
  let last = null; // the item a following Tj continues (no move in between)
  const place = () => { const m = mul(tm, ctm); return { x: m[4], y: m[5], s: Math.abs(size * (Math.hypot(m[2], m[3]) || 1)) }; };
  const advance = (tx) => { tm = mul([1, 0, 0, 1, tx, 0], tm); };
  const show = (raw) => {
    const str = decodeText(raw, font);
    const p = place();
    const { w, n } = rawWidth(raw, font);
    const spaces = (str.match(/ /g) || []).length;
    advance((w * size + n * tc + spaces * tw) * th);
    const e = place();
    if (last) { last.str += str; last.end = Math.round(e.x * 10) / 10; return; }
    last = { x: Math.round(p.x * 10) / 10, end: Math.round(e.x * 10) / 10, y: Math.round(p.y * 10) / 10, size: Math.round(p.s * 10) / 10, str };
    items.push(last);
  };
  const moveTo = (tx, ty) => { tlm = mul([1, 0, 0, 1, tx, ty], tlm); tm = tlm.slice(); last = null; };
  for (const tk of tokens(content)) {
    if (tk.t === '[') { arr = []; continue; }
    if (tk.t === ']') { stack.push({ t: 'arr', v: arr || [] }); arr = null; continue; }
    if (arr && tk.t !== 'op') { arr.push(tk); continue; }
    if (tk.t !== 'op') { stack.push(tk); continue; }
    const num = (k) => { const x = stack[stack.length - k]; return x && x.t === 'num' ? x.v : 0; };
    switch (tk.v) {
      case 'q': saved.push(ctm.slice()); break;
      case 'Q': ctm = saved.pop() || [1, 0, 0, 1, 0, 0]; break;
      case 'cm': ctm = mul([num(6), num(5), num(4), num(3), num(2), num(1)], ctm); break;
      case 'BT': tm = [1, 0, 0, 1, 0, 0]; tlm = tm.slice(); last = null; break;
      case 'ET': last = null; break;
      case 'Tf': { const nm = stack[stack.length - 2]; size = num(1) || size; font = nm && nm.t === 'name' ? fonts(nm.v) : null; break; }
      case 'TL': leading = num(1); break;
      case 'Tc': tc = num(1); break;
      case 'Tw': tw = num(1); break;
      case 'Tz': th = num(1) / 100; break;
      case 'Td': moveTo(num(2), num(1)); break;
      case 'TD': leading = -num(1); moveTo(num(2), num(1)); break;
      case 'Tm': tlm = [num(6), num(5), num(4), num(3), num(2), num(1)]; tm = tlm.slice(); last = null; break;
      case 'T*': moveTo(0, -leading); break;
      case 'Tj': { const x = stack[stack.length - 1]; if (x && x.t === 'str') show(x.v); break; }
      case "'": { moveTo(0, -leading); const x = stack[stack.length - 1]; if (x && x.t === 'str') show(x.v); break; }
      case '"': { moveTo(0, -leading); const x = stack[stack.length - 1]; if (x && x.t === 'str') show(x.v); break; }
      case 'TJ': {
        const a = stack[stack.length - 1];
        if (a && a.t === 'arr') {
          for (const el of a.v) {
            if (el.t === 'str') show(el.v);
            else if (el.t === 'num') { advance((-el.v / 1000) * size * th); if (el.v < -250 && last && !/\s$/.test(last.str)) last.str += ' '; }
          }
        }
        break;
      }
      case 'Do': {
        // Form XObjects (some PDF makers put the page in one).
        const nm = stack[stack.length - 1];
        if (depth < 3 && nm && nm.t === 'name' && xobjects) {
          try {
            const xo = xobjects(nm.v);
            if (xo && xo.content) items.push(...readContent(xo.content, xo.fonts || fonts, ctx, xo.xobjects || xobjects, depth + 1).map((it) => ({ ...it })));
          } catch (e) { /* skip it */ }
        }
        break;
      }
      default: break;
    }
    stack = [];
  }
  return items.filter((it) => it.str && it.str.trim());
}

function resourcesOf(ctx, resDict) {
  const res = resDict instanceof PDFRef ? ctx.lookup(resDict) : resDict;
  const cache = new Map();
  const fontDict = res instanceof PDFDict ? res.lookup(PDFName.of('Font')) : null;
  const xoDict = res instanceof PDFDict ? res.lookup(PDFName.of('XObject')) : null;
  const fonts = (name) => {
    if (cache.has(name)) return cache.get(name);
    let fi = null;
    try { const fo = fontDict instanceof PDFDict ? fontDict.get(PDFName.of(name)) : null; fi = fo ? fontInfo(ctx, fo) : null; } catch (e) { fi = null; }
    cache.set(name, fi); return fi;
  };
  const xobjects = (name) => {
    const ref = xoDict instanceof PDFDict ? xoDict.get(PDFName.of(name)) : null;
    const xo = ref instanceof PDFRef ? ctx.lookup(ref) : ref;
    if (!(xo instanceof PDFRawStream)) return null;
    const sub = xo.dict.get(PDFName.of('Subtype'));
    if (!sub || String(sub) !== '/Form') return null;
    const inner = xo.dict.get(PDFName.of('Resources'));
    const r = inner ? resourcesOf(ctx, inner) : { fonts, xobjects: null };
    return { content: decodeStream(ctx, xo), fonts: r.fonts, xobjects: r.xobjects };
  };
  return { fonts, xobjects };
}

export async function pdfTextItems(bytes) {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
  const ctx = doc.context;
  const out = [];
  for (const page of doc.getPages()) {
    // Let the screen breathe between pages (a long list never freezes the till).
    await new Promise((r) => setTimeout(r, 0));
    out.push(readPage(page));
  }
  return out;
  function readPage(page) {
    try {
      const { Resources, Contents } = page.node.normalizedEntries();
      const r = resourcesOf(ctx, Resources);
      let content = '';
      const c = Contents instanceof PDFRef ? ctx.lookup(Contents) : Contents;
      if (c instanceof PDFArray) { for (let i = 0; i < c.size(); i++) content += `${decodeStream(ctx, c.get(i))}\n`; } else if (c) content = decodeStream(ctx, c);
      return { items: readContent(content, r.fonts, ctx, r.xobjects) };
    } catch (e) { return { items: [], problem: String(e && e.message || e) }; }
  }
}

// Pieces of text grouped into lines (top to bottom), each line's pieces left to right.
export function pdfLines(pages) {
  return pages.map((p) => {
    const items = p.items.slice().sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    for (const it of items) {
      const tol = Math.max(1.5, it.size * 0.35);
      const ln = lines.find((l) => Math.abs(l.y - it.y) <= tol);
      if (ln) ln.items.push(it); else lines.push({ y: it.y, items: [it] });
    }
    // Pieces drawn one after the other (a word split into letters, a number split in two) are joined back:
    // touching → no space; a word space apart → one space. Separate cells are further apart.
    lines.forEach((l) => {
      l.items.sort((a, b) => a.x - b.x);
      const out = [];
      for (const it of l.items) {
        const prev = out[out.length - 1];
        const gap = prev && prev.end !== undefined ? it.x - prev.end : Infinity;
        if (prev && gap < it.size * 0.15 && gap > -it.size) { prev.str += it.str; prev.end = Math.max(prev.end, it.end); continue; }
        if (prev && gap < it.size * 0.45 && gap > 0 && !/^\d/.test(it.str.trim()) && /[A-Za-z)]$/.test(prev.str.trim())) { prev.str += ` ${it.str}`; prev.end = Math.max(prev.end, it.end); continue; }
        out.push({ ...it });
      }
      l.items = out.map((it) => ({ ...it, str: it.str.replace(/\s+/g, ' ').trim() })).filter((it) => it.str);
    });
    return lines.sort((a, b) => b.y - a.y);
  });
}
