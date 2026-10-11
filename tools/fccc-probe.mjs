// Reads the real FCCC Master Price List PDFs and writes what the app's PDF reader sees (for tuning the reader).
import fs from 'fs';
import { pdfTextItems, pdfLines } from './pdfText.mjs';
const pages = ['https://fccc.gov.fj/master-price-list/basic-food-items/', 'https://fccc.gov.fj/master-price-list/hardware/'];
let out = '';
for (const p of pages) {
  const html = await (await fetch(p)).text();
  const links = Array.from(html.matchAll(/href\s*=\s*["']([^"']+?\.pdf)["']/gi)).map((m) => m[1]);
  out += `=== PAGE ${p}\n${links.join('\n')}\n`;
  for (const url of links.slice(0, 2)) {
    try {
      const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
      const pg = await pdfTextItems(bytes);
      const ls = pdfLines(pg);
      out += `=== PDF ${url} pages ${pg.length}\n`;
      ls.slice(0, 3).forEach((lines, i) => { out += `--- page ${i + 1}\n`; lines.slice(0, 70).forEach((l) => { out += `${l.y} | ${l.items.map((t) => `${t.x}-${t.end}:${t.str}`).join(' | ')}\n`; }); });
    } catch (e) { out += `ERROR ${url} ${e.message}\n`; }
  }
}
fs.writeFileSync('fccc-probe.txt', out);
console.log(out.slice(0, 2000));
