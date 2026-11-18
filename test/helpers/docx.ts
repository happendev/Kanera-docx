/**
 * Builds small .docx packages in memory so tests can describe a document in WordprocessingML
 * without checking binary fixtures in. Entries are stored uncompressed.
 */

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface DocxSpec {
  /** Contents of `<w:body>` (without the final sectPr, see `sectPr`). */
  body: string;
  /** Final section properties; defaults to A4 with 1" margins. */
  sectPr?: string;
  /** Contents of `<w:styles>`. */
  styles?: string;
  /** Contents of `<w:numbering>`. */
  numbering?: string;
  /** Contents of `<w:settings>`. */
  settings?: string;
  /** Header/footer parts: relationship id → `<w:hdr>`/`<w:ftr>` inner XML. */
  headers?: Record<string, string>;
  footers?: Record<string, string>;
  /** Images: relationship id → { name, data }. */
  media?: Record<string, { name: string; data: Uint8Array }>;
  /** External hyperlinks: relationship id → URL. */
  links?: Record<string, string>;
  footnotes?: string;
}

export const A4_SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

export function makeDocx(spec: DocxSpec): Uint8Array {
  const files: Array<[string, string | Uint8Array]> = [];
  const rels: string[] = [];
  const overrides: string[] = ['<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'];

  if (spec.styles !== undefined) {
    files.push(['word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W_NS}>${spec.styles}</w:styles>`]);
    rels.push(`<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`);
    overrides.push('<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>');
  }
  if (spec.numbering !== undefined) {
    files.push(['word/numbering.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W_NS}>${spec.numbering}</w:numbering>`]);
    rels.push(`<Relationship Id="rIdNumbering" Type="${REL}/numbering" Target="numbering.xml"/>`);
    overrides.push('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>');
  }
  if (spec.settings !== undefined) {
    files.push(['word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}>${spec.settings}</w:settings>`]);
    rels.push(`<Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/>`);
  }
  if (spec.footnotes !== undefined) {
    files.push(['word/footnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}>${spec.footnotes}</w:footnotes>`]);
    rels.push(`<Relationship Id="rIdFootnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
  }
  for (const [id, xml] of Object.entries(spec.headers ?? {})) {
    files.push([`word/${id}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}>${xml}</w:hdr>`]);
    rels.push(`<Relationship Id="${id}" Type="${REL}/header" Target="${id}.xml"/>`);
  }
  for (const [id, xml] of Object.entries(spec.footers ?? {})) {
    files.push([`word/${id}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W_NS}>${xml}</w:ftr>`]);
    rels.push(`<Relationship Id="${id}" Type="${REL}/footer" Target="${id}.xml"/>`);
  }
  for (const [id, m] of Object.entries(spec.media ?? {})) {
    files.push([`word/media/${m.name}`, m.data]);
    rels.push(`<Relationship Id="${id}" Type="${REL}/image" Target="media/${m.name}"/>`);
  }
  for (const [id, url] of Object.entries(spec.links ?? {})) {
    rels.push(`<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${url}" TargetMode="External"/>`);
  }

  files.push([
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}><w:body>${spec.body}${spec.sectPr ?? A4_SECT}</w:body></w:document>`,
  ]);
  files.push(['word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`]);
  files.push([
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  ]);
  files.push([
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.join('')}</Types>`,
  ]);
  return zipStore(files.map(([name, data]) => [name, typeof data === 'string' ? new TextEncoder().encode(data) : data]));
}

/** Paragraph shorthand: `p('Hello')`, `p('Hello', '<w:jc w:val="center"/>', '<w:b/>')`. */
export function p(text: string, pPr = '', rPr = ''): string {
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of data) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zipStore(files: Array<[string, Uint8Array]>): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameBytes = new TextEncoder().encode(name);
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, nameBytes.length, true);
    chunks.push(new Uint8Array(local.buffer), nameBytes, data);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true);
    cd.setUint16(4, 20, true);
    cd.setUint16(6, 20, true);
    cd.setUint32(16, crc, true);
    cd.setUint32(20, data.length, true);
    cd.setUint32(24, data.length, true);
    cd.setUint16(28, nameBytes.length, true);
    cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const cdSize = central.reduce((n, c) => n + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  const all = [...chunks, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((n, c) => n + c.length, 0));
  let o = 0;
  for (const c of all) {
    out.set(c, o);
    o += c.length;
  }
  return out;
}
