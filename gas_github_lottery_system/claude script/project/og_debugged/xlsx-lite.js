/* 內建離線 Excel / CSV 讀取器（自上週 draw_lots.html 抽出，邏輯未改動）
 * 瀏覽器：window.XlsxLite；Node：module.exports */
(function (root) {
const XlsxLite = (() => {
  /* ---- ZIP：讀取 .xlsx 內的檔案 ---- */
  async function inflateRaw(u8){
    const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  function readZip(buf){
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--){
      if (dv.getUint32(i, true) === 0x06054b50){ eocd = i; break; }
    }
    if (eocd < 0) throw new Error('不是有效的 .xlsx 檔');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const entries = {}, dec = new TextDecoder();
    for (let k = 0; k < count; k++){
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
      entries[name] = { method, csize, off };
      p += 46 + nl + el + cl;
    }
    return async (name) => {
      const e = entries[name];
      if (!e) return null;
      const start = e.off + 30 + dv.getUint16(e.off + 26, true) + dv.getUint16(e.off + 28, true);
      const raw = u8.subarray(start, start + e.csize);
      const data = e.method === 0 ? raw : await inflateRaw(raw);
      return new TextDecoder('utf-8').decode(data);
    };
  }

  /* ---- 迷你 XML 解析器 ---- */
  const ENT = {lt:'<', gt:'>', amp:'&', quot:'"', apos:"'"};
  const unesc = s => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
      e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1))
                   : (ENT[e] ?? m));
  function parseXML(s){
    const root = {name:'#root', attrs:{}, children:[]}, stack = [root];
    const re = /<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>]+)([^>]*?)(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(s))){
      const top = stack[stack.length - 1];
      if (m[1] !== undefined) top.children.push({text: m[1]});
      else if (m[2]) { if (stack.length > 1) stack.pop(); }
      else if (m[3]){
        const node = {name: m[3].split(':').pop(), attrs:{}, children:[]};
        m[4].replace(/([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g, (_, k, __, a, b) => {
          const v = unesc(a ?? b); node.attrs[k] = v; node.attrs[k.split(':').pop()] ??= v;
        });
        top.children.push(node);
        if (!m[5]) stack.push(node);
      }
      else if (m[6]) top.children.push({text: unesc(m[6])});
    }
    return root;
  }
  const kids = (n, name) => n.children.filter(c => c.name === name);
  function* all(n, name){ for (const c of n.children || []){ if (c.name === name) yield c; if (c.children) yield* all(c, name); } }
  const first = (n, name) => all(n, name).next().value;
  // 取文字（略過日文注音 rPh），並還原 Excel 的 _x000D_ 跳脫
  function textOf(n){
    let out = '';
    (function walk(x){ for (const c of x.children || []){
      if (c.text !== undefined){ if (x.name === 't' || x.name === 'v') out += c.text; }
      else if (c.name !== 'rPh') walk(c);
    }})(n);
    return out.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  }
  const colIndex = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '').toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  const colName = i => { let s = ''; i++; while (i > 0){ const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
  function resolve(base, target){
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const seg of target.split('/')){ if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
    return parts.join('/');
  }

  /* ---- 讀 .xlsx → { SheetNames, data: {工作表: [[列]]} } ---- */
  async function readXlsx(buf){
    const get = readZip(buf);
    const wbXml = await get('xl/workbook.xml');
    if (!wbXml) throw new Error('找不到工作簿內容（可能不是 .xlsx）');
    const wb = parseXML(wbXml);
    const rels = {};
    const relXml = await get('xl/_rels/workbook.xml.rels');
    if (relXml) for (const r of all(parseXML(relXml), 'Relationship')) rels[r.attrs.Id] = resolve('xl/workbook.xml', r.attrs.Target);
    const ssXml = await get('xl/sharedStrings.xml');
    const shared = ssXml ? [...all(parseXML(ssXml), 'si')].map(textOf) : [];

    const SheetNames = [], data = {};
    let idx = 1;
    for (const sh of all(wb, 'sheet')){
      const name = sh.attrs.name;
      const path = rels[sh.attrs['r:id'] ?? sh.attrs.id] || `xl/worksheets/sheet${idx}.xml`;
      idx++;
      const xml = await get(path);
      if (!xml) continue; // 圖表工作表等
      const rows = [];
      let rAuto = 0;
      for (const row of all(parseXML(xml), 'row')){
        const r = row.attrs.r ? +row.attrs.r - 1 : rAuto;
        rAuto = r + 1;
        const cells = [];
        let cAuto = 0;
        for (const c of kids(row, 'c')){
          const ci = c.attrs.r ? colIndex(c.attrs.r) : cAuto;
          cAuto = ci + 1;
          const t = c.attrs.t, v = first(c, 'v');
          let val = '';
          if (t === 's') val = v ? (shared[+textOf(v)] ?? '') : '';
          else if (t === 'inlineStr') { const is = first(c, 'is'); val = is ? textOf(is) : ''; }
          else if (t === 'b') val = v && textOf(v) === '1' ? 'TRUE' : 'FALSE';
          else val = v ? textOf(v) : '';
          cells[ci] = val;
        }
        rows[r] = Array.from(cells, x => x ?? '');
      }
      SheetNames.push(name);
      data[name] = rows.filter(r => r && r.some(x => String(x).trim() !== ''));
    }
    return { SheetNames, data };
  }

  /* ---- 文字解碼：先試 UTF-8，失敗改 Big5（台灣版 Excel 存的 CSV 常是 Big5） ---- */
  function decodeText(buf){
    try { return new TextDecoder('utf-8', {fatal:true}).decode(buf).replace(/^﻿/, ''); }
    catch { return new TextDecoder('big5').decode(buf); }
  }

  /* ---- CSV 解析（支援引號、逗號與換行） ---- */
  function parseCSV(text){
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++){
      const ch = text[i];
      if (q){
        if (ch === '"'){ if (text[i + 1] === '"'){ cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r'){
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length){ row.push(cell); rows.push(row); }
    return rows.filter(r => r.some(x => x.trim() !== ''));
  }

  return { readXlsx, parseCSV, decodeText, colName };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = XlsxLite;
else root.XlsxLite = XlsxLite;
})(typeof globalThis !== 'undefined' ? globalThis : this);
