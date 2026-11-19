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
  /** Contents of `<w:endnotes>`. */
  endnotes?: string;
  /** Contents of `<a:theme>`'s `<a:themeElements>`. */
  theme?: string;
  imports?: Record<string, { name: string; contentType: string; text: string }>;
  /** Custom XML data parts content controls can bind to: store item id → XML. */
  customXml?: Record<string, string>;
  /** Contents of the glossary document's `<w:docParts>` (building blocks such as placeholders). */
  glossary?: string;
  /** Contents of `<cp:coreProperties>` (title, author and the like). */
  coreProperties?: string;
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
  for (const [id, part] of Object.entries(spec.imports ?? {})) {
    files.push(['word/' + part.name, part.text]);
    rels.push('<Relationship Id="' + id + '" Type="' + REL + '/aFChunk" Target="' + part.name + '"/>');
    overrides.push('<Override PartName="/word/' + part.name + '" ContentType="' + part.contentType + '"/>');
  }
  if (spec.footnotes !== undefined) {
    files.push(['word/footnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${W_NS}>${spec.footnotes}</w:footnotes>`]);
    rels.push(`<Relationship Id="rIdFootnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`);
  }
  if (spec.endnotes !== undefined) {
    files.push(['word/endnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes ${W_NS}>${spec.endnotes}</w:endnotes>`]);
    rels.push(`<Relationship Id="rIdEndnotes" Type="${REL}/endnotes" Target="endnotes.xml"/>`);
  }
  if (spec.theme !== undefined) {
    files.push(['word/theme/theme1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme ${W_NS} name="Test"><a:themeElements>${spec.theme}</a:themeElements></a:theme>`]);
    rels.push(`<Relationship Id="rIdTheme" Type="${REL}/theme" Target="theme/theme1.xml"/>`);
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
  Object.entries(spec.customXml ?? {}).forEach(([storeId, xml], i) => {
    files.push([`customXml/item${i + 1}.xml`, xml]);
    files.push([
      `customXml/_rels/item${i + 1}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps${i + 1}.xml"/></Relationships>`,
    ]);
    files.push([
      `customXml/itemProps${i + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="${storeId}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>`,
    ]);
    rels.push(`<Relationship Id="rIdCustomXml${i + 1}" Type="${REL}/customXml" Target="../customXml/item${i + 1}.xml"/>`);
  });
  if (spec.glossary !== undefined) {
    files.push(['word/glossary/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W_NS}><w:docParts>${spec.glossary}</w:docParts></w:glossaryDocument>`]);
    rels.push(`<Relationship Id="rIdGlossary" Type="${REL}/glossaryDocument" Target="glossary/document.xml"/>`);
  }
  const packageRels = [`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`];
  if (spec.coreProperties !== undefined) {
    files.push([
      'docProps/core.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${spec.coreProperties}</cp:coreProperties>`,
    ]);
    packageRels.push('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>');
  }

  files.push([
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W_NS}><w:body>${spec.body}${spec.sectPr ?? A4_SECT}</w:body></w:document>`,
  ]);
  files.push(['word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`]);
  files.push([
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${packageRels.join('')}</Relationships>`,
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

export function zipStore(files: Array<[string, Uint8Array]>): Uint8Array {
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
