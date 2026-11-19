import type { DocumentInfo, ImageRef, Link } from '../docx/model.js';
import type { FontFace } from '../fonts/registry.js';
import { buildSfnt, subsetTrueType } from '../fonts/subset.js';
import { toPdfImage, type PdfImage } from '../image/image.js';
import type { DrawOp, GlyphOp, Page, PathSeg } from '../layout/types.js';
import { latin1, num, PdfObjects, pdfName, textString } from './objects.js';

export interface PdfOptions {
  info?: DocumentInfo;
  producer?: string;
  /** Compress content streams (on by default). */
  compress?: boolean;
}

interface FontUse {
  face: FontFace;
  resName: string;
  objNum: number;
  glyphs: Map<number, string>;
}

interface ImageUse {
  resName: string;
  objNum: number;
}

const ITALIC_SKEW = 0.2126; // tan(12°), the slant Word uses for synthetic italics

/** Write laid-out pages as a PDF file. */
export function renderPdf(pages: Page[], opts: PdfOptions = {}, warnings: string[] = []): Uint8Array {
  const pdf = new PdfObjects();
  const catalog = pdf.alloc();
  const pagesNum = pdf.alloc();
  const fonts = new Map<FontFace, FontUse>();
  const images = new Map<string, ImageUse | null>();
  const alphas = new Map<string, string>();
  const anchors = new Map<string, { page: number; x: number; y: number }>();
  const pageNums = pages.map(() => pdf.alloc());
  const compress = opts.compress !== false;

  // Bookmark destinations first, so links can point forward.
  pages.forEach((page, pi) => {
    for (const layer of [page.behind, page.headerFooter, page.body, page.front])
      for (const op of layer) if (op.kind === 'anchor' && !anchors.has(op.name)) anchors.set(op.name, { page: pi, x: op.x, y: page.height - op.y });
  });

  const fontFor = (face: FontFace): FontUse => {
    let use = fonts.get(face);
    if (!use) {
      use = { face, resName: `F${fonts.size + 1}`, objNum: pdf.alloc(), glyphs: new Map() };
      fonts.set(face, use);
    }
    return use;
  };

  const imageFor = (ref: ImageRef): ImageUse | null => {
    if (images.has(ref.id)) return images.get(ref.id)!;
    const img = toPdfImage(ref.data);
    if (!img) {
      warnings.push(`Image ${ref.id} uses an unsupported format and was skipped`);
      images.set(ref.id, null);
      return null;
    }
    const use = { resName: `Im${images.size + 1}`, objNum: writeImage(pdf, img) };
    images.set(ref.id, use);
    return use;
  };

  const alphaFor = (a: number): string => {
    const key = num(a);
    let name = alphas.get(key);
    if (!name) {
      name = `GS${alphas.size + 1}`;
      alphas.set(key, name);
    }
    return name;
  };

  pages.forEach((page, pi) => {
    const H = page.height;
    const out: string[] = [];
    const usedFonts = new Set<FontUse>();
    const usedImages = new Set<ImageUse>();
    const annots: string[] = [];
    const Y = (y: number) => H - y;
    const color = (hex: string, stroke: boolean) => {
      const v = parseInt(hex.slice(0, 6), 16) || 0;
      return `${num(((v >> 16) & 255) / 255)} ${num(((v >> 8) & 255) / 255)} ${num((v & 255) / 255)} ${stroke ? 'RG' : 'rg'}`;
    };
    const pathOps = (segs: PathSeg[]): string => {
      let s = '';
      for (const seg of segs) {
        const p = seg.pts;
        if (seg.op === 'M') s += `${num(p[0])} ${num(Y(p[1]))} m `;
        else if (seg.op === 'L') s += `${num(p[0])} ${num(Y(p[1]))} l `;
        else if (seg.op === 'C') s += `${num(p[0])} ${num(Y(p[1]))} ${num(p[2])} ${num(Y(p[3]))} ${num(p[4])} ${num(Y(p[5]))} c `;
        else s += 'h ';
      }
      return s;
    };

    const glyphs = (op: GlyphOp) => {
      if (!op.gids.length) return;
      const use = fontFor(op.face);
      usedFonts.add(use);
      const font = op.face.font;
      const toPdfUnits = 1000 / font.unitsPerEm;
      const th = op.scaleX || 1;
      // Text render mode, horizontal scaling, stroke colour and line width outlive ET, so stroked
      // (bold or outline) and scaled text saves and restores the graphics state; otherwise every
      // later run would be stroked or scaled too.
      const stroked = op.syntheticBold || op.outline;
      const alpha = op.alpha !== undefined && op.alpha < 1;
      const save = alpha || stroked || th !== 1;
      if (alpha) out.push(`q /${alphaFor(op.alpha!)} gs`);
      else if (save) out.push('q');
      let s = `BT /${use.resName} ${num(op.size)} Tf `;
      if (th !== 1) s += `${num(th * 100)} Tz `;
      s += color(op.color, false) + ' ';
      if (stroked) {
        s += `${color(op.color, true)} ${num(op.size * (op.outline ? 0.02 : 0.03))} w ${op.outline ? 1 : 2} Tr `;
      }
      const x0 = op.x + op.xs[0];
      s += `1 0 ${op.syntheticItalic ? ITALIC_SKEW : 0} 1 ${num(x0)} ${num(Y(op.y))} Tm [`;
      for (let i = 0; i < op.gids.length; i++) {
        const gid = op.gids[i];
        if (!use.glyphs.has(gid)) use.glyphs.set(gid, op.texts[i] ?? '');
        s += `<${gid.toString(16).padStart(4, '0')}>`;
        if (i + 1 < op.gids.length) {
          const w = font.advance(gid) * toPdfUnits;
          const d = op.xs[i + 1] - op.xs[i];
          const adj = w - (d * 1000) / (op.size * th);
          if (Math.abs(adj) > 0.05) s += num(adj);
        }
      }
      s += '] TJ ET';
      out.push(s);
      if (save) out.push('Q');
    };

    const draw = (op: DrawOp) => {
      switch (op.kind) {
        case 'glyphs':
          glyphs(op);
          break;
        case 'rect': {
          const a = op.alpha !== undefined && op.alpha < 1;
          out.push(`${a ? `q /${alphaFor(op.alpha!)} gs ` : ''}${color(op.fill, false)} ${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re f${a ? ' Q' : ''}`);
          break;
        }
        case 'line': {
          const cap = op.cap === 'round' ? 1 : op.cap === 'square' ? 2 : 0;
          const dash = op.dash ? `[${op.dash.map(num).join(' ')}] 0 d` : '[] 0 d';
          out.push(`${num(op.width)} w ${cap} J ${dash} ${color(op.color, true)} ${num(op.x1)} ${num(Y(op.y1))} m ${num(op.x2)} ${num(Y(op.y2))} l S`);
          break;
        }
        case 'path': {
          const path = pathOps(op.segs);
          if (op.fill) {
            const a = op.fillAlpha !== undefined && op.fillAlpha < 1;
            out.push(`${a ? `q /${alphaFor(op.fillAlpha!)} gs ` : ''}${color(op.fill, false)} ${path}f*${a ? ' Q' : ''}`);
          }
          if (op.stroke) {
            const a = op.strokeAlpha !== undefined && op.strokeAlpha < 1;
            const dash = op.dash ? `[${op.dash.map(num).join(' ')}] 0 d` : '[] 0 d';
            out.push(`${a ? `q /${alphaFor(op.strokeAlpha!)} gs ` : ''}${num(op.strokeWidth ?? 0.75)} w 0 J 1 j ${dash} ${color(op.stroke, true)} ${path}S${a ? ' Q' : ''}`);
          }
          break;
        }
        case 'image': {
          const use = imageFor(op.image);
          if (!use) break;
          usedImages.add(use);
          const c = op.crop ?? { l: 0, t: 0, r: 0, b: 0 };
          const fw = 1 - c.l - c.r;
          const fh = 1 - c.t - c.b;
          if (fw <= 0 || fh <= 0) break;
          const W = op.w / fw;
          const Hh = op.h / fh;
          const ix = op.x - c.l * W;
          const iy = op.y - c.t * Hh;
          let s = 'q ';
          if (c.l || c.t || c.r || c.b) s += `${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re W n `;
          s += `${num(W)} 0 0 ${num(Hh)} ${num(ix)} ${num(Y(iy + Hh))} cm /${use.resName} Do Q`;
          out.push(s);
          break;
        }
        case 'transform': {
          const [a, b, c, d, e, f] = op.matrix;
          out.push(`q ${num(a)} ${num(-b)} ${num(-c)} ${num(d)} ${num(c * H + e)} ${num(H - d * H - f)} cm`);
          break;
        }
        case 'clip':
          if (op.segs) out.push(`q ${pathOps(op.segs)}W* n`);
          else out.push(`q ${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.w)} ${num(op.h)} re W n`);
          break;
        case 'restore':
          out.push('Q');
          break;
        case 'link': {
          const action = linkAction(op.link, anchors, pageNums);
          if (!action) break;
          annots.push(`<< /Type /Annot /Subtype /Link /Rect [${num(op.x)} ${num(Y(op.y + op.h))} ${num(op.x + op.w)} ${num(Y(op.y))}] /Border [0 0 0] ${action} >>`);
          break;
        }
        case 'anchor':
        case 'float':
          break;
      }
    };
    for (const layer of [page.behind, page.headerFooter, page.body, page.front]) for (const op of layer) draw(op);

    const content = pdf.addStream('', latin1(out.join('\n')), compress);
    const fontRes = [...usedFonts].map((f) => `/${f.resName} ${f.objNum} 0 R`).join(' ');
    const imgRes = [...usedImages].map((i) => `/${i.resName} ${i.objNum} 0 R`).join(' ');
    const gsRes = [...alphas].map(([a, n]) => `/${n} << /Type /ExtGState /ca ${a} /CA ${a} >>`).join(' ');
    const annotNums = annots.map((a) => pdf.add(a));
    pdf.set(
      pageNums[pi],
      `<< /Type /Page /Parent ${pagesNum} 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Resources << /ProcSet [/PDF /Text /ImageB /ImageC /ImageI]${fontRes ? ` /Font << ${fontRes} >>` : ''}${imgRes ? ` /XObject << ${imgRes} >>` : ''}${gsRes ? ` /ExtGState << ${gsRes} >>` : ''} >> /Contents ${content} 0 R${annotNums.length ? ` /Annots [${annotNums.map((n) => `${n} 0 R`).join(' ')}]` : ''} >>`,
    );
  });

  let fi = 0;
  for (const use of fonts.values()) writeFont(pdf, use, fi++, compress);

  pdf.set(pagesNum, `<< /Type /Pages /Kids [${pageNums.map((n) => `${n} 0 R`).join(' ')}] /Count ${pages.length} >>`);
  pdf.set(catalog, `<< /Type /Catalog /Pages ${pagesNum} 0 R >>`);
  const info = opts.info ?? {};
  const entries = [`/Producer ${textString(opts.producer ?? 'Folio')}`];
  if (info.title) entries.push(`/Title ${textString(info.title)}`);
  if (info.creator) entries.push(`/Author ${textString(info.creator)}`);
  if (info.subject) entries.push(`/Subject ${textString(info.subject)}`);
  if (info.keywords) entries.push(`/Keywords ${textString(info.keywords)}`);
  const infoNum = pdf.add(`<< ${entries.join(' ')} >>`);
  return pdf.finish(catalog, infoNum);
}

function linkAction(link: Link, anchors: Map<string, { page: number; x: number; y: number }>, pageNums: number[]): string | null {
  if (link.anchor) {
    const a = anchors.get(link.anchor);
    if (a) return `/Dest [${pageNums[a.page]} 0 R /XYZ ${num(a.x)} ${num(a.y)} 0]`;
  }
  if (link.url) {
    let url = link.url;
    if (link.anchor && !anchors.has(link.anchor)) url += '#' + link.anchor;
    const ascii = url.replace(/[^\x20-\x7e]/g, (c) => encodeURIComponent(c)).replace(/[\\()]/g, (c) => '\\' + c);
    return `/A << /S /URI /URI (${ascii}) >>`;
  }
  return null;
}

function writeImage(pdf: PdfObjects, img: PdfImage): number {
  let smask = '';
  if (img.smask) {
    const n = pdf.addStream(`/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode`, img.smask, false);
    smask = ` /SMask ${n} 0 R`;
  }
  const parms = img.decodeParms ? ` /DecodeParms ${img.decodeParms}` : '';
  const decode = img.decode ? ` /Decode [${img.decode.join(' ')}]` : '';
  return pdf.addStream(
    `/Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace ${img.colorSpace} /BitsPerComponent ${img.bitsPerComponent} /Filter /${img.filter}${parms}${decode}${smask}`,
    img.data,
    false,
  );
}

function subsetTag(i: number): string {
  let s = '';
  let n = i + 1;
  for (let k = 0; k < 6; k++) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function writeFont(pdf: PdfObjects, use: FontUse, index: number, compress: boolean): void {
  const font = use.face.font;
  const upm = font.unitsPerEm;
  const scale = 1000 / upm;
  const psName = (font.names.postscriptName || use.face.familyName.replace(/\s+/g, '')).replace(/[^\x21-\x7e]/g, '');
  const baseFont = `${subsetTag(index)}+${psName}`;
  const gids = [...use.glyphs.keys()].sort((a, b) => a - b);

  let fileKey: string;
  let fileNum: number;
  let subtype: string;
  const sub = subsetTrueType(font, gids);
  if (sub) {
    fileNum = pdf.addStream(`/Length1 ${sub.length}`, sub, compress);
    fileKey = 'FontFile2';
    subtype = 'CIDFontType2';
  } else {
    // CFF outlines: embed the whole OpenType font (as a standalone file when in a collection).
    let data = font.data;
    if (font.dirOffset !== 0) {
      const tables = new Map<string, Uint8Array>();
      for (const tag of font.tables.keys()) tables.set(tag, font.tableBytes(tag)!);
      data = buildSfnt(tables, font.sfntVersion);
    }
    fileNum = pdf.addStream('/Subtype /OpenType', data, compress);
    fileKey = 'FontFile3';
    subtype = 'CIDFontType0';
  }

  // Widths
  let w = '';
  for (let i = 0; i < gids.length; ) {
    let j = i;
    const run: string[] = [];
    while (j < gids.length && gids[j] === gids[i] + (j - i)) {
      run.push(num(font.advance(gids[j]) * scale));
      j++;
    }
    w += `${gids[i]} [${run.join(' ')}] `;
    i = j;
  }

  let flags = 0;
  if (font.isFixedPitch) flags |= 1;
  flags |= 32; // nonsymbolic
  if (use.face.italic) flags |= 64;
  const [x0, y0, x1, y1] = font.bbox;
  const descriptor = pdf.add(
    `<< /Type /FontDescriptor /FontName ${pdfName(baseFont)} /Flags ${flags} /FontBBox [${num(x0 * scale)} ${num(y0 * scale)} ${num(x1 * scale)} ${num(y1 * scale)}] /ItalicAngle ${num(font.italicAngle)} /Ascent ${num(font.hheaAscender * scale)} /Descent ${num(font.hheaDescender * scale)} /CapHeight ${num((font.capHeight || font.hheaAscender * 0.7) * scale)} /StemV ${use.face.bold ? 120 : 80} /${fileKey} ${fileNum} 0 R >>`,
  );
  const cid = pdf.add(
    `<< /Type /Font /Subtype /${subtype} /BaseFont ${pdfName(baseFont)} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R /DW ${num(font.advance(0) * scale)} /W [${w}]${subtype === 'CIDFontType2' ? ' /CIDToGIDMap /Identity' : ''} >>`,
  );
  const toUnicode = pdf.addStream('', latin1(toUnicodeCMap(use.glyphs)), compress);
  pdf.set(use.objNum, `<< /Type /Font /Subtype /Type0 /BaseFont ${pdfName(baseFont)} /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${toUnicode} 0 R >>`);
}

function utf16Hex(s: string): string {
  let h = '';
  for (let i = 0; i < s.length; i++) h += s.charCodeAt(i).toString(16).padStart(4, '0');
  return h.toUpperCase();
}

function toUnicodeCMap(glyphs: Map<number, string>): string {
  const entries = [...glyphs].filter(([, t]) => t.length > 0).sort((a, b) => a[0] - b[0]);
  let s = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n';
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100);
    s += `${chunk.length} beginbfchar\n`;
    for (const [gid, text] of chunk) s += `<${gid.toString(16).padStart(4, '0').toUpperCase()}> <${utf16Hex(text)}>\n`;
    s += 'endbfchar\n';
  }
  s += 'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n';
  return s;
}
