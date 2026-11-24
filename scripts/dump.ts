/** Print a summary of a .docx body: blocks, breaks, sections, styles. pnpm tsx scripts/dump.ts file.docx [from] [to] */
import { nodeCompression } from '../src/util/node-compression.js';
import { readFileSync } from 'node:fs';
import { parseDocx } from '../src/docx/parser.js';
import type { Block } from '../src/ooxml/model.js';

const [file, fromS, toS] = process.argv.slice(2);
const doc = parseDocx(new Uint8Array(readFileSync(file)), nodeCompression.inflateRaw);
const from = Number(fromS ?? 0);
const to = Number(toS ?? 40);
const text = (b: Block): string => {
  if (b.type === 'table') return `[table ${b.rows.length}x${b.grid.length}]`;
  return b.content
    .map((i) => (i.type === 'text' ? i.text : i.type === 'tab' ? '→' : i.type === 'break' ? `<${i.breakType}>` : i.type === 'renderedPageBreak' ? '¶RPB¶' : i.type === 'drawing' ? `<drawing ${i.drawing.kind} ${i.drawing.inline ? 'inline' : 'float'} ${Math.round(i.drawing.width)}x${Math.round(i.drawing.height)}>` : i.type === 'field' ? `{${i.field}}` : ''))
    .join('');
};
doc.body.slice(from, to).forEach((b, i) => {
  let extra = '';
  if (b.type === 'paragraph') {
    const p = b.props;
    extra = `${p.styleId} sz=${b.markProps.size} ${b.markProps.fontAscii} sb=${p.spaceBefore} sa=${p.spaceAfter} line=${p.line}/${p.lineRule}${p.keepNext ? ' keepNext' : ''}${p.pageBreakBefore ? ' PBB' : ''}${b.sectPr ? ` SECT(${b.sectPr.type})` : ''}${b.numbering ? ` num="${b.numbering.text}"` : ''}`;
  }
  console.log(`${from + i}: ${extra} | ${text(b).slice(0, 100)}`);
});
const s = doc.finalSection;
console.log('final section', { w: s.pageWidth, h: s.pageHeight, mt: s.marginTop, mb: s.marginBottom, ml: s.marginLeft, mr: s.marginRight, header: s.header, footer: s.footer, type: s.type, titlePage: s.titlePage, cols: s.columns.count });
console.log('settings', doc.settings);
console.log('fonts', [...new Set(doc.fontRequests.map((r) => r.family))]);
console.log('warnings', doc.warnings);
