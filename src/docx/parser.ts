import { MAX_FONT_FACES } from '../fonts/sfnt.js';
import { fontFor } from '../fonts/script.js';
import { deobfuscateFont } from '../fonts/odttf.js';
import { attr, child, children, decodeEntities, descendant, descendants, elements, parseXml, path, resolveAlternateContent, textContent, type XmlElement } from '../util/xml.js';
import { parseDrawing, parseVml, loadImage, type DrawingContext } from './drawing.js';
import type {
  Block,
  Border,
  CellBorders,
  CellMargins,
  Columns,
  DocSettings,
  DocxDocument,
  HeaderFooterRefs,
  FontRequest,
  ImageRef,
  Inline,
  Link,
  Note,
  NumberingLabel,
  PageBorders,
  Paragraph,
  RunProps,
  Section,
  Table,
  TableCell,
  TableRow,
} from './model.js';
import { Numbering } from './numbering.js';
import type { InflateBytes } from '../util/compression.js';
import type { ArchiveLimits } from '../util/zip.js';
import { OpcPackage } from './package.js';
import {
  finalizePara,
  finalizeRun,
  mergePara,
  mergeRun,
  mergeTbl,
  mergeTc,
  mergeTr,
  parseParaDelta,
  parseRunDelta,
  parseTblDelta,
  parseTcDelta,
  parseTrDelta,
  themeColorFn,
  type ParaDelta,
  type PropContext,
  type RunDelta,
  type TableBordersDelta,
  type TableLook,
  type TcDelta,
} from './props.js';
import { StyleSheet } from './styles.js';
import { DEFAULT_THEME, parseTheme, type Theme } from './theme.js';
import { emu, hexColor, measure, num, onOff, parseBorder, twip, twipAttr } from './values.js';

function checkTableColumns(count: number): void {
  if (!Number.isSafeInteger(count) || count < 0 || count > 256) throw new Error('Table column limit exceeded');
}

interface TableContext {
  tableRun: RunDelta;
  tablePara: ParaDelta;
  override: boolean;
  /** Pre-2013 rule: the default paragraph style's 11 or 12 points give way to tableRun's size. */
  legacySize: boolean;
}

interface PartContext {
  part: string;
  table?: TableContext;
}

interface FieldFrame {
  instr: string;
  phase: 'instr' | 'result';
  kind?: 'page' | 'numpages' | 'sectionpages';
  format?: string;
  link?: Link;
  inline?: Inline & { type: 'field' };
  cached: string;
  /** A TOC field showing its result (the entries). */
  toc?: boolean;
  /** Where a TOC field's result starts, to tell an empty one. */
  tocStart?: { para: Paragraph; length: number };
  /** State of a legacy form checkbox field. */
  checkbox?: boolean;
  /** Selected entry of a legacy form drop-down field. */
  dropdown?: string;
}

const DEFAULT_SECTION_TWIPS = { w: 12240, h: 15840, margin: 1440, header: 720, footer: 720 };
const BUILT_IN_STYLES =
  '<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="279" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>';

export function parseDocx(data: Uint8Array, inflateRaw: InflateBytes, limits?: ArchiveLimits, embeddedFonts = true): DocxDocument {
  return new DocxParser(new OpcPackage(data, inflateRaw, limits), embeddedFonts).parse();
}

class DocxParser {
  private main: string;
  private theme: Theme;
  private ctx: PropContext;
  private styles: StyleSheet;
  private numbering: Numbering;
  private settings: DocSettings;
  private fieldStack: FieldFrame[] = [];
  private images = new Map<string, ImageRef>();
  private fontRequests = new Map<string, FontRequest>();
  private warnings: string[] = [];
  private headerCache = new Map<string, Block[]>();
  private prevHeaders: HeaderFooterRefs = {};
  private prevFooters: HeaderFooterRefs = {};
  private hyperlinkStack: Link[] = [];

  constructor(private readonly pkg: OpcPackage, private readonly readEmbeddedFonts: boolean) {
    this.main = this.pkg.mainDocumentPart();
    const settingsRel = this.pkg.relByType(this.main, '/settings');
    const settingsXml = settingsRel ? this.pkg.xml(settingsRel.target) : undefined;
    const fontLang = child(settingsXml, 'w:themeFontLang')?.attrs;
    const themeRel = this.pkg.relByType(this.main, '/theme');
    this.theme = themeRel ? parseTheme(this.pkg.xml(themeRel.target), { eastAsia: fontLang?.['w:eastAsia'], bidi: fontLang?.['w:bidi'] }) : DEFAULT_THEME;
    const mapping = child(settingsXml, 'w:clrSchemeMapping');
    const colorMapping: Record<string, string> = {};
    if (mapping) {
      for (const [k, v] of Object.entries(mapping.attrs)) {
        const key = k.replace(/^w:/, '');
        const val = v === 'dark1' ? 'dk1' : v === 'light1' ? 'lt1' : v === 'dark2' ? 'dk2' : v === 'light2' ? 'lt2' : v;
        colorMapping[key] = val;
      }
    }
    this.ctx = { theme: this.theme, colorMapping };
    this.settings = this.parseSettings(settingsXml);
    const stylesRel = this.pkg.relByType(this.main, '/styles');
    // Without a styles part Word falls back to its built-in defaults: Aptos 12 pt, 8 pt after and
    // 1.1625 lines (measured in Word's PDFs of such files).
    this.styles = new StyleSheet(stylesRel ? this.pkg.xml(stylesRel.target) : parseXml(BUILT_IN_STYLES), this.ctx);
    const numRel = this.pkg.relByType(this.main, '/numbering');
    this.numbering = new Numbering(numRel ? this.pkg.xml(numRel.target) : undefined, this.ctx, (styleId) => this.styles.paraStyle(styleId).numId);
  }

  private parseSettings(s: XmlElement | undefined): DocSettings {
    const compat = child(s, 'w:compat');
    let mode = 12;
    let override: boolean | undefined;
    for (const cs of children(compat, 'w:compatSetting')) {
      if (cs.attrs['w:name'] === 'compatibilityMode') mode = num(cs.attrs['w:val']) ?? mode;
      if (cs.attrs['w:name'] === 'overrideTableStyleFontSizeAndJustification') override = cs.attrs['w:val'] === '1';
    }
    // A package with no settings part at all opens in Word 2007 mode (its tables hang the cell margin
    // left of the indent); settings without a compat block are Word 2013+.
    if (!compat || !children(compat, 'w:compatSetting').length) mode = compat || !s ? 12 : 15;
    return {
      defaultTabStop: twipAttr(child(s, 'w:defaultTabStop'), 'w:val') ?? 36,
      evenAndOddHeaders: !!onOff(child(s, 'w:evenAndOddHeaders')),
      mirrorMargins: !!onOff(child(s, 'w:mirrorMargins')),
      compatibilityMode: mode,
      doNotExpandShiftReturn: !!onOff(child(compat, 'w:doNotExpandShiftReturn')),
      wpJustification: !!onOff(child(compat, 'w:wpJustification')),
      compressPunctuation: attr(child(s, 'w:characterSpacingControl'), 'w:val')?.startsWith('compressPunctuation') ?? false,
      // Word 2013 mode behaves as if the setting were on when a document leaves it out.
      overrideTableStyleFontSizeAndJustification: override ?? mode >= 15,
      autoHyphenation: !!onOff(child(s, 'w:autoHyphenation')),
      displayBackgroundShape: !!onOff(child(s, 'w:displayBackgroundShape')),
      doNotUseHTMLParagraphAutoSpacing: !!onOff(child(compat, 'w:doNotUseHTMLParagraphAutoSpacing')),
      adjustLineHeightInTable: !!onOff(child(compat, 'w:adjustLineHeightInTable')),
      footnotePosition: child(child(s, 'w:footnotePr'), 'w:pos')?.attrs['w:val'] ?? 'pageBottom',
      endnotePosition: child(child(s, 'w:endnotePr'), 'w:pos')?.attrs['w:val'] ?? 'docEnd',
      generated: !child(s, 'w:rsids'),
      updateFields: !!onOff(child(s, 'w:updateFields')),
    };
  }

  parse(): DocxDocument {
    const doc = this.pkg.xml(this.main);
    const body = child(doc, 'w:body');
    if (!body) throw new Error('Document has no body');
    const pctx: PartContext = { part: this.main };
    const blocks = this.parseBlocks(body, pctx);
    const finalSectPr = child(body, 'w:sectPr');
    const finalSection = this.parseSection(finalSectPr, pctx);
    this.buildTocs(blocks, finalSection, pctx);
    resolveLineSpacing(blocks, finalSection);

    const footnotes = new Map<string, Note>();
    const endnotes = new Map<string, Note>();
    let footnoteSeparator: Block[] | null = null;
    let footnoteContinuationSeparator: Block[] | null = null;
    let endnoteSeparator: Block[] | null = null;
    for (const [relType, elName, map] of [
      ['/footnotes', 'w:footnote', footnotes],
      ['/endnotes', 'w:endnote', endnotes],
    ] as const) {
      const rel = this.pkg.relByType(this.main, relType);
      if (!rel) continue;
      const root = this.pkg.xml(rel.target);
      for (const n of children(root, elName)) {
        const id = n.attrs['w:id'] ?? '';
        const type = n.attrs['w:type'];
        this.fieldStack = [];
        const blocks2 = this.parseBlocks(n, { part: rel.target });
        if (type === 'separator') {
          if (elName === 'w:footnote') footnoteSeparator = blocks2;
          else endnoteSeparator = blocks2;
        } else if (type === 'continuationSeparator') {
          if (elName === 'w:footnote') footnoteContinuationSeparator = blocks2;
        } else if (!type || type === 'normal') map.set(id, { id, blocks: blocks2 });
      }
    }

    const fontHints = new Map<string, { family?: string; pitch?: string; panose?: string; charset?: string; altName?: string }>();
    const embeddedFonts: DocxDocument['embeddedFonts'] = [];
    let embeddedBytes = 0;
    const ftRel = this.pkg.relByType(this.main, '/fontTable');
    if (ftRel) {
      const ft = this.pkg.xml(ftRel.target);
      for (const f of children(ft, 'w:font')) {
        const name = f.attrs['w:name'] ?? '';
        fontHints.set(name, {
          family: child(f, 'w:family')?.attrs['w:val'],
          pitch: child(f, 'w:pitch')?.attrs['w:val'],
          panose: child(f, 'w:panose1')?.attrs['w:val'],
          charset: child(f, 'w:charset')?.attrs['w:val'],
          altName: child(f, 'w:altName')?.attrs['w:val'],
        });
        if (!this.readEmbeddedFonts) continue;
        for (const [tag, bold, italic] of [
          ['w:embedRegular', false, false],
          ['w:embedBold', true, false],
          ['w:embedItalic', false, true],
          ['w:embedBoldItalic', true, true],
        ] as const) {
          const e = child(f, tag);
          if (!e) continue;
          const rel = this.pkg.rels(ftRel.target).get(e.attrs['r:id'] ?? '');
          const bytes = rel ? this.pkg.bytes(rel.target) : null;
          if (!bytes) continue;
          embeddedBytes += bytes.length;
          if (embeddedFonts.length >= MAX_FONT_FACES || embeddedBytes > 128 * 1024 * 1024) throw new Error('Embedded font limit exceeded');
          const key = e.attrs['w:fontKey'];
          embeddedFonts.push({ family: name, bold, italic, data: key ? deobfuscateFont(bytes, key) : bytes });
        }
      }
    }

    const info: DocxDocument['info'] = {};
    const core = this.pkg.xml('docProps/core.xml');
    if (core) {
      const get = (n: string) => textContent(child(core, n)).trim() || undefined;
      info.title = get('dc:title');
      info.subject = get('dc:subject');
      info.creator = get('dc:creator');
      info.keywords = get('cp:keywords');
      info.description = get('dc:description');
    }

    const bg = child(doc, 'w:background');
    return {
      body: blocks,
      finalSection,
      settings: { ...this.settings, defaultParagraphStyle: this.styles.defaultParagraphStyle ?? '' },
      footnotes,
      endnotes,
      footnoteSeparator,
      footnoteContinuationSeparator,
      endnoteSeparator,
      info,
      background: bg ? (hexColor(bg.attrs['w:color']) ?? null) : null,
      fontRequests: [...this.fontRequests.values()],
      embeddedFonts,
      fontHints,
      warnings: this.warnings,
    };
  }

  // -------------------------------------------------------------------------------------------
  // Blocks

  parseBlocks(container: XmlElement, pctx: PartContext): Block[] {
    const out: Block[] = [];
    this.collectBlocks(container, pctx, out);
    // Word does not lay out an empty paragraph whose mark is hidden (it joins the next paragraph).
    // The last paragraph of the body, a cell or a note stays: it closes its container.
    return out.filter(
      (b, i) => i === out.length - 1 || b.type !== 'paragraph' || !b.markProps.hidden || b.sectPr || !b.content.every((c) => c.type === 'renderedPageBreak'),
    );
  }

  private importChunk(el: XmlElement, pctx: PartContext): Block[] {
    const id = el.attrs['r:id'] ?? '';
    const rel = this.pkg.rels(pctx.part).get(id);
    const type = rel && !rel.external ? this.pkg.contentType(rel.target) : undefined;
    if (!rel || (type !== 'text/html' && type !== 'application/xhtml+xml' && type !== 'text/plain')) {
      this.warnings.push('Skipped unsupported or missing altChunk import: ' + (rel?.target ?? id));
      return [];
    }
    let text = this.pkg.zip.readText(rel.target);
    if (text === null) {
      this.warnings.push('Missing altChunk part: ' + rel.target);
      return [];
    }
    if (type !== 'text/plain') {
      // Import readable text only. HTML is never attached to a browser DOM or executed.
      this.warnings.push('Imported HTML altChunk as plain text: ' + rel.target);
      text = text.replace(/<!--[\s\S]*?-->/g, '')
        .replace(/<(head|script|style)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, '')
        .replace(/\s+/g, ' ')
        .replace(/<\/?(?:p|div|li|h[1-6]|br|tr|table|section|article)\b[^>]*>/gi, '\n')
        .replace(/<\/t[dh]\s*>/gi, ' ')
        .replace(/<\/?[a-z!](?:[^>"']|"[^"]*"|'[^']*')*>/gi, '');
      text = decodeEntities(text.replace(/&nbsp;/gi, '\u00a0'));
    }
    const lines = text.split(/\r\n?|\n/);
    const blocks: Block[] = [];
    for (const line of lines) {
      if (type !== 'text/plain' && !line.trim()) continue;
      const para = this.parseParagraph({ name: 'w:p', attrs: {}, children: [] }, pctx);
      this.pushText(para, type === 'text/plain' ? line : line.trim(), para.markProps);
      blocks.push(para);
    }
    return blocks;
  }

  private collectBlocks(container: XmlElement, pctx: PartContext, out: Block[]): void {
    for (const el of children(container)) {
      switch (el.name) {
        case 'w:p':
          out.push(this.parseParagraph(el, pctx));
          break;
        case 'w:altChunk':
          out.push(...this.importChunk(el, pctx));
          break;
        case 'w:tbl':
          out.push(this.parseTable(el, pctx));
          break;
        case 'w:sdt':
          this.collectBlocks(child(el, 'w:sdtContent') ?? el, pctx, out);
          break;
        case 'w:customXml':
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
          this.collectBlocks(el, pctx, out);
          break;
        case 'w:del':
        case 'w:moveFrom':
          break;
        case 'mc:AlternateContent': {
          const alternatives = resolveAlternateContent(el);
          if (!alternatives.length) this.warnings.push('Skipped alternate block without a supported choice or fallback in ' + pctx.part);
          for (const alt of alternatives) this.collectBlocks({ name: 'x', attrs: {}, children: [alt] }, pctx, out);
          break;
        }
        default:
          if (descendant(el, 'w:t') || descendant(el, 'w:drawing') || descendant(el, 'w:pict')) {
            this.warnings.push('Skipped unsupported content block ' + el.name + ' in ' + pctx.part);
          }
      }
    }
  }

  // -------------------------------------------------------------------------------------------
  // Paragraphs

  private parseParagraph(p: XmlElement, pctx: PartContext): Paragraph {
    const pPr = child(p, 'w:pPr');
    const direct = parseParaDelta(pPr, this.ctx);
    const styleId = direct.styleId ?? this.styles.defaultParagraphStyle;
    const isDefaultStyle = styleId === this.styles.defaultParagraphStyle;
    const styleP = this.styles.paraStyle(styleId);
    const numId = direct.numId ?? styleP.numId;
    const fromStyle = direct.numId === undefined && styleP.numId !== undefined;
    const ilvl = direct.ilvl ?? styleP.ilvl ?? 0;
    const hasNumbering = this.numbering.exists(numId);
    const lvl = hasNumbering ? this.numbering.level(numId!, ilvl) : undefined;

    let merged: ParaDelta = this.styles.defaultsPara;
    const tableP = pctx.table?.tablePara;
    if (tableP) merged = mergePara(merged, tableP);
    if (lvl && fromStyle) merged = mergePara(mergePara(merged, lvl.pPr), styleP);
    else if (lvl) merged = mergePara(mergePara(merged, styleP), lvl.pPr);
    else merged = mergePara(merged, styleP);
    if (tableP?.align !== undefined && pctx.table?.override && isDefaultStyle) merged = { ...merged, align: tableP.align };
    const directNoStyle = { ...direct };
    delete directNoStyle.styleId;
    merged = mergePara(merged, directNoStyle);

    const normalSize = this.styles.paraStyleRun(styleId).size ?? this.styles.defaultsRun.size ?? 10;
    // With overrideTableStyleFontSizeAndJustification on (Word 2013+), the default paragraph style's
    // size beats the table style's, as the style hierarchy says; only older documents let the table win.
    const tableOverridesSize = isDefaultStyle && !!pctx.table && pctx.table.legacySize && (normalSize === 11 || normalSize === 12);
    const markDirect = parseRunDelta(child(pPr, 'w:rPr'), this.ctx);
    const markDelta = this.styles.runDelta({ paraStyle: styleId, charStyle: markDirect.styleId, tableRun: pctx.table?.tableRun, direct: markDirect, tableOverridesSize });
    const markProps = finalizeRun(markDelta);
    this.noteFonts(markProps, '');
    const props = finalizePara(merged, markProps.size);
    props.styleId = styleId ?? '';

    const para: Paragraph = { type: 'paragraph', props, markProps, content: [], numbering: null };

    if (lvl && hasNumbering) {
      const text = this.numbering.next(numId!, ilvl);
      if (text !== null) {
        const labelProps = finalizeRun(mergeRun(markDelta, lvl.rPr));
        // Word never underlines or highlights the number from the paragraph mark alone.
        if (!lvl.rPr.underline) labelProps.underline = null;
        labelProps.highlight = null;
        let picture: ImageRef | undefined;
        if (lvl.picBulletId !== undefined) {
          const rid = this.numbering.picBullets.get(lvl.picBulletId);
          const numRel = this.pkg.relByType(this.main, '/numbering');
          if (rid && numRel) picture = loadImage(this.drawingCtx({ part: numRel.target }), rid);
        }
        const label: NumberingLabel = { text, props: labelProps, suffix: lvl.suffix, align: lvl.jc, picture, numId: numId! };
        para.numbering = label;
        this.noteFonts(labelProps, text);
      }
    }

    const sectPr = child(pPr, 'w:sectPr');
    this.parseInlines(p, para, pctx, styleId, tableOverridesSize);
    if (sectPr && pctx.part === this.main) para.sectPr = this.parseSection(sectPr, pctx);
    return para;
  }

  private parseInlines(container: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    for (const el of children(container)) {
      switch (el.name) {
        case 'w:r':
          this.parseRun(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'w:hyperlink': {
          const rid = el.attrs['r:id'];
          const link: Link = {};
          if (rid) {
            const rel = this.pkg.rels(pctx.part).get(rid);
            if (rel) link.url = rel.target;
          }
          if (el.attrs['w:anchor']) link.anchor = el.attrs['w:anchor'];
          this.hyperlinkStack.push(link);
          this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
          this.hyperlinkStack.pop();
          break;
        }
        case 'w:fldSimple': {
          const instr = el.attrs['w:instr'] ?? '';
          const info = analyzeField(instr);
          if (info.kind) {
            const firstRun = child(el, 'w:r');
            const props = this.runProps(child(firstRun, 'w:rPr'), styleId, pctx, tableOverridesSize);
            if (!this.suppressed())
              para.content.push({ type: 'field', field: info.kind, format: info.format, props, cached: textContent(el) });
          } else {
            if (info.link) this.hyperlinkStack.push(info.link);
            this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
            if (info.link) this.hyperlinkStack.pop();
          }
          break;
        }
        case 'w:ins':
        case 'w:moveTo':
        case 'w:smartTag':
        case 'w:customXml':
        case 'w:dir':
        case 'w:bdo':
          this.parseInlines(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'w:sdt': {
          const content = child(el, 'w:sdtContent') ?? el;
          const sdtPr = child(el, 'w:sdtPr');
          const bound = this.boundText(sdtPr);
          if (bound === null) this.parseInlines(content, para, pctx, styleId, tableOverridesSize);
          else this.parseRun(boundRun(content, sdtPr, bound), para, pctx, styleId, tableOverridesSize);
          break;
        }
        case 'w:bookmarkStart':
          if (el.attrs['w:name']) para.content.push({ type: 'bookmark', name: el.attrs['w:name'] });
          break;
        case 'm:oMathPara':
        case 'm:oMath':
          this.parseMath(el, para, pctx, styleId, tableOverridesSize);
          break;
        case 'mc:AlternateContent':
          this.parseInlines({ name: 'x', attrs: {}, children: resolveAlternateContent(el) }, para, pctx, styleId, tableOverridesSize);
          break;
      }
    }
  }

  /** Office Math is rendered linearly (base text of each math run) in the run's font. */
  private parseMath(el: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    for (const c of children(el)) {
      if (c.name === 'm:r') {
        const props = this.runProps(child(c, 'w:rPr'), styleId, pctx, tableOverridesSize);
        const text = children(c, 'm:t')
          .map((t) => textContent(t))
          .join('');
        if (text && !this.suppressed()) {
          if (!child(c, 'w:rPr')?.children.some((x) => typeof x !== 'string' && x.name === 'w:rFonts')) {
            props.fontAscii = props.fontHAnsi = 'Cambria Math';
          }
          props.italic = !child(child(c, 'm:rPr'), 'm:sty') && /[A-Za-z]/.test(text);
          this.pushText(para, text, props);
        }
      } else if (c.name === 'm:f') {
        // Fraction: num/den
        this.parseMath(child(c, 'm:num') ?? c, para, pctx, styleId, tableOverridesSize);
        const props = this.runProps(undefined, styleId, pctx, tableOverridesSize);
        this.pushText(para, '/', props);
        this.parseMath(child(c, 'm:den') ?? c, para, pctx, styleId, tableOverridesSize);
      } else if (c.name === 'w:r') {
        this.parseRun(c, para, pctx, styleId, tableOverridesSize);
      } else if (!c.name.endsWith('Pr')) {
        this.parseMath(c, para, pctx, styleId, tableOverridesSize);
      }
    }
  }

  private runProps(rPr: XmlElement | undefined, styleId: string | undefined, pctx: PartContext, tableOverridesSize: boolean): RunProps {
    const direct = parseRunDelta(rPr, this.ctx);
    let charStyle = direct.styleId;
    // Word shows table-of-contents entries in their TOC paragraph styles: the Hyperlink character
    // style that the TOC field's \h switch puts on every entry is not drawn (the link still works).
    if (charStyle && this.styles.styles.get(charStyle)?.name.toLowerCase() === 'hyperlink' && this.fieldStack.some((f) => f.toc)) charStyle = undefined;
    const delta = this.styles.runDelta({ paraStyle: styleId, charStyle, tableRun: pctx.table?.tableRun, direct, tableOverridesSize });
    const props = finalizeRun(delta);
    props.styleId = direct.styleId;
    return props;
  }

  private suppressed(): boolean {
    return this.fieldStack.some((f) => f.phase === 'instr' || (f.kind !== undefined && f.phase === 'result'));
  }

  private dataStores?: Map<string, XmlElement>;
  /** Tables of contents with no entries, for Word to fill in (see buildTocs). */
  private emptyTocs: Array<{ para: Paragraph; instr: string }> = [];
  private tocMarks = 0;

  private noteEmptyToc(para: Paragraph, instr: string, pctx: PartContext): void {
    if (this.settings.updateFields && pctx.part === this.main) this.emptyTocs.push({ para, instr });
  }

  /**
   * Word fills in an empty table of contents when it opens a document that asks it to update its
   * fields (an empty one otherwise stays empty in its PDFs): one entry per heading in the field's
   * levels, in place of the field's paragraph, each showing the page its heading lands on.
   */
  private buildTocs(blocks: Block[], finalSection: Section, pctx: PartContext): void {
    for (const { para, instr } of this.emptyTocs) {
      const at = blocks.indexOf(para);
      const levels = tocLevels(instr);
      if (at < 0 || !levels || para.content.some((c) => c.type === 'text')) continue;
      const section = blocks.slice(at).find((b): b is Paragraph => b.type === 'paragraph' && !!b.sectPr)?.sectPr ?? finalSection;
      const tabPos = Math.round((section.pageWidth - section.marginLeft - section.marginRight) * 20);
      const entries: Paragraph[] = [];
      const visit = (list: Block[]) => {
        for (const b of list) {
          if (b.type === 'table') {
            for (const row of b.rows) for (const cell of row.cells) visit(cell.content);
            continue;
          }
          const level = (b.props.outlineLevel ?? 9) + 1;
          const text = level >= levels[0] && level <= levels[1] ? headingText(b) : '';
          if (!text.trim()) continue;
          const mark = `_TocEntry${++this.tocMarks}`;
          b.content.unshift({ type: 'bookmark', name: mark });
          entries.push(this.tocEntry(level, text, mark, tabPos, /\\h\b/i.test(instr), pctx));
        }
      };
      visit(blocks);
      if (!entries.length) continue;
      entries[0].content.unshift(...para.content);
      entries[entries.length - 1].sectPr = para.sectPr;
      blocks.splice(at, 1, ...entries);
    }
  }

  private tocEntry(level: number, text: string, mark: string, tabPos: number, link: boolean, pctx: PartContext): Paragraph {
    const el = (name: string, attrs: Record<string, string> = {}, ...kids: Array<XmlElement | string>): XmlElement => ({ name, attrs, children: kids });
    const style = this.styles.byName(`toc ${level}`);
    // Word's own TOC styles when the document has none: 5pt after, and each level indented by the
    // size of the Normal font.
    const normalSize = this.styles.paraStyleRun(this.styles.defaultParagraphStyle).size ?? this.styles.defaultsRun.size ?? 10;
    const pPr = style
      ? [el('w:pStyle', { 'w:val': style.id })]
      : [el('w:spacing', { 'w:after': '100' }), el('w:ind', { 'w:left': String(Math.round((level - 1) * normalSize * 20)) })];
    pPr.push(el('w:tabs', {}, el('w:tab', { 'w:val': 'right', 'w:leader': 'dot', 'w:pos': String(tabPos) })));
    const runs = text.split('\t').flatMap((part, i) => [...(i ? [el('w:r', {}, el('w:tab'))] : []), el('w:r', {}, el('w:t', { 'xml:space': 'preserve' }, part))]);
    runs.push(el('w:r', {}, el('w:tab')));
    const p = el('w:p', {}, el('w:pPr', {}, ...pPr), ...(link ? [el('w:hyperlink', { 'w:anchor': mark }, ...runs)] : runs));
    const entry = this.parseBlocks(el('w:body', {}, p), pctx)[0] as Paragraph;
    const last = entry.content.filter((c) => c.type === 'text').pop();
    entry.content.push({ type: 'field', field: 'pageref', bookmark: mark, props: last?.props ?? entry.markProps, cached: '' });
    return entry;
  }

  private currentLink(): Link | undefined {
    for (let i = this.fieldStack.length - 1; i >= 0; i--) if (this.fieldStack[i].link) return this.fieldStack[i].link;
    return this.hyperlinkStack[this.hyperlinkStack.length - 1];
  }

  /**
   * Text Word shows in a content control bound to XML data (w:dataBinding): the bound value, or the
   * control's placeholder when the value is empty. Null keeps the cached content: the control is not
   * bound, its value would need formatting (dates, pictures) or the binding does not resolve.
   */
  private boundText(sdtPr: XmlElement | undefined): string | null {
    const binding = child(sdtPr, 'w:dataBinding');
    if (!binding || child(sdtPr, 'w:date') || child(sdtPr, 'w:picture')) return null;
    const store = this.dataStore(binding.attrs['w:storeItemID'] ?? '');
    const node = store && selectXPath(store, binding.attrs['w:xpath'] ?? '');
    if (!node) return null;
    let value = textContent(node);
    const list = child(sdtPr, 'w:dropDownList') ?? child(sdtPr, 'w:comboBox');
    const item = list && children(list, 'w:listItem').find((i) => i.attrs['w:value'] === value);
    if (item) value = item.attrs['w:displayText'] ?? value;
    if (value) return value;
    const docPart = path(sdtPr, 'w:placeholder', 'w:docPart')?.attrs['w:val'];
    return docPart ? this.placeholderText(docPart) : null;
  }

  /** XML data a content control can bind to, by store item id: Word's document properties and the custom XML parts. */
  private dataStore(id: string): XmlElement | undefined {
    if (!this.dataStores) {
      const stores = (this.dataStores = new Map());
      const add = (id: string | undefined, part: string | undefined) => {
        const xml = part ? this.pkg.xml(part) : undefined;
        if (id && xml) stores.set(id.toUpperCase(), xml);
      };
      add('{6C3C8BC8-F283-45AE-878A-BAB7291924A1}', this.pkg.relByType('', '/core-properties')?.target ?? 'docProps/core.xml');
      add('{6668398D-A668-4E3E-A5EB-62B293D839F1}', this.pkg.relByType('', '/extended-properties')?.target ?? 'docProps/app.xml');
      for (const rel of this.pkg.rels(this.main).values()) {
        if (rel.external || !rel.type.endsWith('/customXml')) continue;
        const props = this.pkg.relByType(rel.target, '/customXmlProps');
        const attrs = props ? this.pkg.xml(props.target)?.attrs : undefined;
        add(attrs && Object.entries(attrs).find(([k]) => k === 'itemID' || k.endsWith(':itemID'))?.[1], rel.target);
      }
    }
    return this.dataStores.get(id.toUpperCase());
  }

  /** Text of the first paragraph of a glossary building block (a content control's placeholder). */
  private placeholderText(name: string): string | null {
    const rel = this.pkg.relByType(this.main, '/glossaryDocument');
    const parts = children(child(rel ? this.pkg.xml(rel.target) : undefined, 'w:docParts'), 'w:docPart');
    const part = parts.find((d) => path(d, 'w:docPartPr', 'w:name')?.attrs['w:val'] === name);
    const para = child(child(part, 'w:docPartBody'), 'w:p');
    return para ? descendants(para, 'w:t').map((t) => textContent(t)).join('') : null;
  }

  private noteFonts(props: RunProps, text: string): void {
    for (const ch of text || ' ') {
      const cp = ch.codePointAt(0)!;
      const { family, bold, italic } = fontFor(cp, props);
      const lang = props.eastAsiaLang;
      const key = family + '\0' + Number(bold) + Number(italic) + '\0' + (lang ?? '');
      let request = this.fontRequests.get(key);
      if (!request) {
        request = { family, bold, italic, codePoints: new Set(), lang };
        this.fontRequests.set(key, request);
      }
      request.codePoints.add(cp);
    }
  }

  private pushText(para: Paragraph, text: string, props: RunProps): void {
    if (!text) return;
    const frame = this.fieldStack[this.fieldStack.length - 1];
    if (frame && frame.kind && frame.phase === 'result') {
      // Word shows a computed field (PAGE, NUMPAGES) in the formatting of its result, not of the
      // field characters around it.
      if (!frame.cached && frame.inline && !props.hidden) {
        frame.inline.props = props;
        this.noteFonts(props, text);
      }
      frame.cached += text;
    }
    if (this.suppressed() || props.hidden) return;
    this.noteFonts(props, text);
    const link = this.currentLink();
    const last = para.content[para.content.length - 1];
    if (last && last.type === 'text' && last.props === props && last.link === link) {
      last.text += text;
      return;
    }
    para.content.push({ type: 'text', text, props, link });
  }

  private parseRun(r: XmlElement, para: Paragraph, pctx: PartContext, styleId: string | undefined, tableOverridesSize: boolean): void {
    const props = this.runProps(child(r, 'w:rPr'), styleId, pctx, tableOverridesSize);
    // Breaks, tabs and pictures are measured with the run's font too, so load it even when the
    // run has no text.
    const visible = () => !this.suppressed() && !props.hidden;
    if (visible()) this.noteFonts(props, '');
    const items = children(r);
    for (let idx = 0; idx < items.length; idx++) {
      const el = items[idx];
      switch (el.name) {
        case 'w:t': {
          let t = textContent(el);
          if (el.attrs['xml:space'] !== 'preserve') t = t.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
          t = t.replace(/[\r\n]/g, '');
          this.pushText(para, t, props);
          break;
        }
        case 'w:tab':
          if (visible()) para.content.push({ type: 'tab', props, link: this.currentLink() });
          break;
        case 'w:ptab': {
          if (!visible()) break;
          const a = el.attrs['w:alignment'];
          const leader = el.attrs['w:leader'];
          para.content.push({
            type: 'ptab',
            props,
            alignment: a === 'center' ? 'center' : a === 'right' ? 'right' : 'left',
            relativeTo: el.attrs['w:relativeTo'] === 'indent' ? 'indent' : 'margin',
            leader: leader === 'dot' || leader === 'hyphen' || leader === 'underscore' || leader === 'middleDot' ? leader : 'none',
          });
          break;
        }
        case 'w:br': {
          if (!visible()) break;
          const type = el.attrs['w:type'];
          para.content.push({ type: 'break', breakType: type === 'page' ? 'page' : type === 'column' ? 'column' : 'line', clear: el.attrs['w:clear'], props });
          break;
        }
        case 'w:cr':
          if (visible()) para.content.push({ type: 'break', breakType: 'line', props });
          break;
        case 'w:softHyphen':
          this.pushText(para, '\u00ad', props);
          break;
        case 'w:noBreakHyphen':
          this.pushText(para, '\u2011', props);
          break;
        case 'w:sym': {
          const font = el.attrs['w:font'];
          const code = parseInt(el.attrs['w:char'] ?? '0', 16);
          if (!code) break;
          const symProps: RunProps = font ? { ...props, fontAscii: font, fontHAnsi: font, fontEastAsia: font, fontCs: font } : props;
          this.pushText(para, String.fromCodePoint(code), symProps);
          break;
        }
        case 'w:drawing': {
          if (!visible()) break;
          const d = parseDrawing(el, this.drawingCtx(pctx));
          if (d && !d.hidden) para.content.push({ type: 'drawing', drawing: d, props, link: this.currentLink() });
          break;
        }
        case 'w:pict':
        case 'w:object': {
          if (!visible()) break;
          const d = parseVml(el, this.drawingCtx(pctx));
          if (d && !d.hidden && d.kind !== 'unsupported') para.content.push({ type: 'drawing', drawing: d, props, link: this.currentLink() });
          break;
        }
        case 'mc:AlternateContent': {
          const alt = resolveAlternateContent(el);
          const fake: XmlElement = { name: 'w:r', attrs: {}, children: [child(r, 'w:rPr') ?? { name: 'w:rPr', attrs: {}, children: [] }, ...alt] };
          this.parseRun(fake, para, pctx, styleId, tableOverridesSize);
          break;
        }
        case 'w:fldChar': {
          const type = el.attrs['w:fldCharType'];
          if (type === 'begin') {
            // Legacy form checkboxes (FORMCHECKBOX) keep their state in the field's form data.
            const box = child(child(el, 'w:ffData'), 'w:checkBox');
            const state = child(box, 'w:checked') ?? child(box, 'w:default');
            const checked = state ? !['0', 'false', 'off'].includes(state.attrs['w:val'] ?? '1') : false;
            // Drop-downs (FORMDROPDOWN) store no result text: Word shows the selected list entry.
            const list = child(child(el, 'w:ffData'), 'w:ddList');
            const entries = children(list, 'w:listEntry').map((e) => e.attrs['w:val'] ?? '');
            const index = parseInt((child(list, 'w:result') ?? child(list, 'w:default'))?.attrs['w:val'] ?? '0', 10) || 0;
            this.fieldStack.push({ instr: '', phase: 'instr', cached: '', checkbox: box ? checked : undefined, dropdown: list ? (entries[index] ?? entries[0]) : undefined });
          }
          else if (type === 'separate') {
            const f = this.fieldStack[this.fieldStack.length - 1];
            if (f && f.phase === 'instr') {
              const info = analyzeField(f.instr);
              f.phase = 'result';
              f.link = info.link;
              f.toc = /^TOC\b/i.test(f.instr.trim());
              if (f.toc) f.tocStart = { para, length: para.content.length };
              if (info.kind) {
                f.kind = info.kind;
                f.format = info.format;
                // Only emit when every enclosing field is showing its result.
                const outerVisible = this.fieldStack.slice(0, -1).every((o) => o.phase === 'result' && !o.kind);
                if (outerVisible && !props.hidden) {
                  const inline: Inline & { type: 'field' } = { type: 'field', field: info.kind, format: info.format, props, cached: '' };
                  f.inline = inline;
                  para.content.push(inline);
                }
              }
            }
          } else if (type === 'end') {
            const f = this.fieldStack.pop();
            if (f?.checkbox !== undefined && !this.suppressed()) this.pushText(para, f.checkbox ? '\u2612' : '\u2610', { ...props, drawn: true });
            if (f?.dropdown && !this.suppressed()) this.pushText(para, f.dropdown, props);
            if (f) {
              if (f.phase === 'instr') {
                // Field without a result: dynamic fields still render.
                const info = analyzeField(f.instr);
                if (info.kind && !this.suppressed() && !props.hidden)
                  para.content.push({ type: 'field', field: info.kind, format: info.format, props, cached: '' });
                if (/^TOC\b/i.test(f.instr.trim())) this.noteEmptyToc(para, f.instr, pctx);
              } else if (f.inline) f.inline.cached = f.cached;
              else if (f.tocStart?.para === para && !para.content.slice(f.tocStart.length).some((c) => c.type === 'text')) this.noteEmptyToc(para, f.instr, pctx);
            }
          }
          break;
        }
        case 'w:instrText': {
          const f = this.fieldStack[this.fieldStack.length - 1];
          if (f && f.phase === 'instr') f.instr += textContent(el);
          break;
        }
        case 'w:footnoteReference':
        case 'w:endnoteReference':
          if (visible())
            para.content.push({
              type: 'noteRef',
              noteType: el.name === 'w:footnoteReference' ? 'footnote' : 'endnote',
              id: el.attrs['w:id'] ?? '',
              props,
              customMark: el.attrs['w:customMarkFollows'] === '1' || el.attrs['w:customMarkFollows'] === 'true' ? '' : undefined,
            });
          break;
        case 'w:footnoteRef':
        case 'w:endnoteRef':
          if (visible()) para.content.push({ type: 'noteNumber', props });
          break;
        case 'w:separator':
        case 'w:continuationSeparator':
          para.content.push({ type: 'separator', continuation: el.name === 'w:continuationSeparator', props });
          break;
        case 'w:lastRenderedPageBreak':
          para.content.push({ type: 'renderedPageBreak' });
          break;
        case 'w:ruby': {
          const base = child(el, 'w:rubyBase');
          if (base) this.parseInlines(base, para, pctx, styleId, tableOverridesSize);
          break;
        }
      }
    }
  }

  private drawingCtx(pctx: { part: string }): DrawingContext {
    return {
      pkg: this.pkg,
      part: pctx.part,
      theme: this.theme,
      colorMapping: this.ctx.colorMapping,
      images: this.images,
      parseBlocks: (container) => {
        // Text box content is a separate story: field state must not leak in or out.
        const saved = this.fieldStack;
        const savedLinks = this.hyperlinkStack;
        this.fieldStack = [];
        this.hyperlinkStack = [];
        const blocks = this.parseBlocks(container, { part: pctx.part });
        this.fieldStack = saved;
        this.hyperlinkStack = savedLinks;
        return blocks;
      },
    };
  }

  // -------------------------------------------------------------------------------------------
  // Tables

  private parseTable(tbl: XmlElement, pctx: PartContext): Table {
    const direct = parseTblDelta(child(tbl, 'w:tblPr'), this.ctx);
    const styleId = direct.styleId ?? this.styles.defaultTableStyle;
    const ts = this.styles.tableStyle(styleId);
    const tblD = mergeTbl(ts.tbl, direct);
    const look: TableLook = direct.look ?? tblD.look ?? { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
    const gridCols = children(child(tbl, 'w:tblGrid'), 'w:gridCol');
    checkTableColumns(gridCols.length);
    const grid = gridCols.map((g) => twipAttr(g, 'w:w') ?? 0);
    const tableBorders: TableBordersDelta = tblD.borders ?? {};
    // Word's 0.08" side margins come from the Normal Table style; a styles part without a default
    // table style leaves cells with none (as in Word's PDFs of such generated files).
    const side = this.styles.defaultTableStyle ? 5.4 : 0;
    const margins: CellMargins = { top: 0, bottom: 0, left: side, right: side, ...(tblD.cellMargins ?? {}) };

    // Gather rows/cells, unwrapping content controls and custom XML.
    const rowEls: XmlElement[] = [];
    const collectRows = (el: XmlElement) => {
      for (const c of children(el)) {
        if (c.name === 'w:tr') {
          if (!child(child(c, 'w:trPr'), 'w:del')) rowEls.push(c);
        }
        else if (c.name === 'w:sdt') collectRows(child(c, 'w:sdtContent') ?? c);
        else if (c.name === 'w:customXml' || c.name === 'w:ins') collectRows(c);
      }
    };
    collectRows(tbl);
    const cellEls = (tr: XmlElement): XmlElement[] => {
      const out: XmlElement[] = [];
      const walk = (el: XmlElement) => {
        for (const c of children(el)) {
          if (c.name === 'w:tc') out.push(c);
          else if (c.name === 'w:sdt') walk(child(c, 'w:sdtContent') ?? c);
          else if (c.name === 'w:customXml') walk(c);
        }
      };
      walk(tr);
      return out;
    };

    const nRows = rowEls.length;
    const nCols = Math.max(grid.length, 1);
    const rowBand = tblD.rowBandSize || 1;
    const colBand = tblD.colBandSize || 1;
    const rows: TableRow[] = [];
    let maxCols = 0;
    for (let r = 0; r < nRows; r++) {
      const tr = rowEls[r];
      const trD = mergeTr(ts.tr, parseTrDelta(child(tr, 'w:trPr')));
      const row: TableRow = {
        height: trD.height ?? 0,
        heightRule: trD.heightRule ?? 'atLeast',
        cantSplit: !!trD.cantSplit,
        isHeader: !!trD.isHeader,
        gridBefore: trD.gridBefore ?? 0,
        gridAfter: trD.gridAfter ?? 0,
        cellSpacing: 0,
        stamped: Object.keys(tr.attrs).some((k) => k.startsWith('w:rsid')),
        cells: [],
      };
      // Row-level exceptions to the table properties (w:tblPrEx) override borders, shading and
      // cell margins for this row's cells.
      const ex = parseTblDelta(child(tr, 'w:tblPrEx'), this.ctx);
      row.cellSpacing = trD.cellSpacing ?? ex.cellSpacing ?? tblD.cellSpacing ?? 0;
      const rowBorders: TableBordersDelta = ex.borders ? { ...tableBorders, ...ex.borders } : tableBorders;
      const rowShading = ex.shading !== undefined ? ex.shading : tblD.shading;
      const rowMargins: CellMargins = ex.cellMargins ? { ...margins, ...ex.cellMargins } : margins;
      checkTableColumns(row.gridBefore);
      checkTableColumns(row.gridAfter);
      let col = row.gridBefore;
      const tcs = cellEls(tr);
      for (const tc of tcs) {
        const tcDirect = parseTcDelta(child(tc, 'w:tcPr'), this.ctx);
        const span = Math.max(1, tcDirect.gridSpan ?? 1);
        const startCol = col;
        const endCol = col + span;
        checkTableColumns(endCol + row.gridAfter);
        col = endCol;

        const conds: Array<{ type: string; region: 'row' | 'col' | 'corner' | 'whole' }> = [{ type: 'wholeTable', region: 'whole' }];
        const isFirstRow = r === 0 && look.firstRow;
        const isLastRow = r === nRows - 1 && look.lastRow && nRows > 1;
        const isFirstCol = startCol === 0 && look.firstColumn;
        const isLastCol = endCol >= nCols && look.lastColumn;
        if (!look.noVBand) {
          const idx = startCol - (look.firstColumn ? 1 : 0);
          if (idx >= 0 && !isFirstCol && !isLastCol) conds.push({ type: Math.floor(idx / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert', region: 'col' });
        }
        if (!look.noHBand) {
          const idx = r - (look.firstRow ? 1 : 0);
          if (idx >= 0 && !isFirstRow && !isLastRow) conds.push({ type: Math.floor(idx / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz', region: 'row' });
        }
        if (isFirstCol) conds.push({ type: 'firstCol', region: 'col' });
        if (isLastCol) conds.push({ type: 'lastCol', region: 'col' });
        if (isFirstRow) conds.push({ type: 'firstRow', region: 'row' });
        if (isLastRow) conds.push({ type: 'lastRow', region: 'row' });
        if (isFirstRow && isFirstCol) conds.push({ type: 'nwCell', region: 'corner' });
        if (isFirstRow && isLastCol) conds.push({ type: 'neCell', region: 'corner' });
        if (isLastRow && isFirstCol) conds.push({ type: 'swCell', region: 'corner' });
        if (isLastRow && isLastCol) conds.push({ type: 'seCell', region: 'corner' });

        let pD: ParaDelta = ts.pPr;
        let rD: RunDelta = ts.rPr;
        let tcD: TcDelta = ts.tc;
        const isTop = r === 0;
        const isBottom = r === nRows - 1;
        const isLeft = startCol === 0;
        const isRight = endCol >= nCols;
        const pick = (b: Border | null | undefined, fallback: Border | null | undefined) => (b !== undefined ? b : fallback);
        let top = isTop ? rowBorders.top : rowBorders.insideH;
        let bottom = isBottom ? rowBorders.bottom : rowBorders.insideH;
        let left = isLeft ? rowBorders.left : rowBorders.insideV;
        let right = isRight ? rowBorders.right : rowBorders.insideV;
        let shading = rowShading;
        for (const c of conds) {
          const cf = ts.cond.get(c.type);
          if (!cf) continue;
          pD = mergePara(pD, cf.pPr);
          rD = mergeRun(rD, cf.rPr);
          tcD = mergeTc(tcD, { ...cf.tc, borders: undefined });
          if (cf.tbl.shading !== undefined) shading = cf.tbl.shading;
          const cb = { ...(cf.tbl.borders ?? {}), ...(cf.tc.borders ?? {}) };
          if (c.region === 'row') {
            top = pick(cb.top, top);
            bottom = pick(cb.bottom, bottom);
            left = isLeft ? pick(cb.left, left) : pick(cb.insideV, left);
            right = isRight ? pick(cb.right, right) : pick(cb.insideV, right);
          } else if (c.region === 'col') {
            left = pick(cb.left, left);
            right = pick(cb.right, right);
            top = isTop ? pick(cb.top, top) : pick(cb.insideH, top);
            bottom = isBottom ? pick(cb.bottom, bottom) : pick(cb.insideH, bottom);
          } else {
            top = isTop ? pick(cb.top, top) : pick(cb.insideH, top);
            bottom = isBottom ? pick(cb.bottom, bottom) : pick(cb.insideH, bottom);
            left = isLeft ? pick(cb.left, left) : pick(cb.insideV, left);
            right = isRight ? pick(cb.right, right) : pick(cb.insideV, right);
          }
        }
        tcD = mergeTc(tcD, tcDirect);
        const db = tcDirect.borders ?? {};
        top = pick(db.top, top);
        bottom = pick(db.bottom, bottom);
        left = pick(db.left, left);
        right = pick(db.right, right);
        const borders: CellBorders = {
          top: top ?? null,
          bottom: bottom ?? null,
          left: left ?? null,
          right: right ?? null,
          tl2br: db.tl2br ?? null,
          tr2bl: db.tr2bl ?? null,
        };
        const override = this.settings.overrideTableStyleFontSizeAndJustification;
        // Without that setting, a table style with paragraph or run formatting but no font size
        // still displaces a default paragraph style of 11 or 12 points: the text takes the document
        // defaults' size. A table style that sets a size leaves the paragraph style's in place.
        // (Fitted to Word's PDFs, where the two cases disagree with the setting's description.)
        const legacySize = !override && rD.size === undefined && (Object.keys(pD).length > 0 || Object.keys(rD).length > 0);
        const defaults = this.styles.defaultsRun;
        const tableRun = legacySize ? { ...rD, size: defaults.size ?? 10, sizeCs: defaults.sizeCs ?? defaults.size ?? 10 } : rD;
        const tableCtx: TableContext = { tableRun, tablePara: pD, override, legacySize };
        const cell: TableCell = {
          gridSpan: span,
          vMerge: tcDirect.vMerge ?? null,
          width: tcDirect.width ?? { type: 'auto', value: 0 },
          borders,
          shading: tcD.shading !== undefined ? tcD.shading : (shading ?? null),
          margins: { ...rowMargins, ...(tcD.margins ?? {}) },
          vAlign: tcD.vAlign ?? 'top',
          noWrap: !!tcD.noWrap,
          hideMark: !!tcD.hideMark,
          textDirection: tcD.textDirection ?? 'lrTb',
          content: this.parseBlocks(tc, { part: pctx.part, table: tableCtx }),
        };
        row.cells.push(cell);
      }
      checkTableColumns(col + row.gridAfter);
      maxCols = Math.max(maxCols, col + row.gridAfter);
      rows.push(row);
    }

    // Grid fallback: derive columns from cell widths when tblGrid is missing or short.
    let finalGrid = grid;
    if (finalGrid.length < maxCols || finalGrid.every((w) => w <= 0)) {
      finalGrid = new Array(maxCols).fill(0);
      for (const row of rows) {
        let c = row.gridBefore;
        for (const cell of row.cells) {
          const w = cell.width.type === 'dxa' ? cell.width.value / cell.gridSpan : 0;
          for (let k = 0; k < cell.gridSpan; k++) if (w > 0 && !finalGrid[c + k]) finalGrid[c + k] = w;
          c += cell.gridSpan;
        }
      }
    }

    return {
      type: 'table',
      props: {
        styleId,
        width: tblD.width ?? { type: 'auto', value: 0 },
        align: tblD.align ?? 'left',
        indent: tblD.indent ?? 0,
        layout: tblD.layout ?? 'autofit',
        cellMargins: margins,
        cellSpacing: tblD.cellSpacing ?? 0,
        floating: direct.floating ? parseFloating(direct.floating) : null,
        bidiVisual: !!tblD.bidiVisual,
      },
      grid: finalGrid,
      rows,
    };
  }

  // -------------------------------------------------------------------------------------------
  // Sections, headers and footers

  private parseSection(sectPr: XmlElement | undefined, pctx: PartContext): Section {
    const pgSz = child(sectPr, 'w:pgSz');
    const pgMar = child(sectPr, 'w:pgMar');
    // Page sizes and margins may carry units ("10mm"); bare numbers are twips.
    const tw = (v: string | undefined) => { const pt = measure(v); return pt === undefined ? undefined : pt * 20; };
    const t = (el: XmlElement | undefined, n: string, d: number) => twip(tw(attr(el, n)) ?? d);
    let pageWidth = t(pgSz, 'w:w', DEFAULT_SECTION_TWIPS.w);
    let pageHeight = t(pgSz, 'w:h', DEFAULT_SECTION_TWIPS.h);
    if (attr(pgSz, 'w:orient') === 'landscape' && pageWidth < pageHeight) [pageWidth, pageHeight] = [pageHeight, pageWidth];
    const topRaw = tw(attr(pgMar, 'w:top')) ?? DEFAULT_SECTION_TWIPS.margin;
    const bottomRaw = tw(attr(pgMar, 'w:bottom')) ?? DEFAULT_SECTION_TWIPS.margin;

    const cols = child(sectPr, 'w:cols');
    const colEls = children(cols, 'w:col');
    const count = Math.max(1, num(attr(cols, 'w:num')) ?? (colEls.length || 1));
    if (!Number.isSafeInteger(count) || count > 64 || colEls.length > 64) throw new Error('Section column limit exceeded');
    const columns: Columns = {
      count,
      space: twip(tw(attr(cols, 'w:space')) ?? 720),
      equalWidth: attr(cols, 'w:equalWidth') !== '0' && attr(cols, 'w:equalWidth') !== 'false',
      separator: attr(cols, 'w:sep') === '1' || attr(cols, 'w:sep') === 'true',
      cols: colEls.map((c) => ({ width: twip(tw(c.attrs['w:w']) ?? 0), space: twip(tw(c.attrs['w:space']) ?? 0) })),
    };
    if (colEls.length && !columns.equalWidth) columns.count = colEls.length;

    const headers: HeaderFooterRefs = { ...this.prevHeaders };
    const footers: HeaderFooterRefs = { ...this.prevFooters };
    for (const ref of children(sectPr, 'w:headerReference').concat(children(sectPr, 'w:footerReference'))) {
      const type = (ref.attrs['w:type'] ?? 'default') as keyof HeaderFooterRefs;
      const rel = this.pkg.rels(pctx.part).get(ref.attrs['r:id'] ?? '');
      if (!rel) continue;
      const blocks = this.headerFooter(rel.target, ref.name === 'w:headerReference' ? 'w:hdr' : 'w:ftr');
      if (ref.name === 'w:headerReference') headers[type] = blocks;
      else footers[type] = blocks;
    }
    this.prevHeaders = headers;
    this.prevFooters = footers;

    const typeVal = attr(child(sectPr, 'w:type'), 'w:val');
    const pgNum = child(sectPr, 'w:pgNumType');
    const grid = child(sectPr, 'w:docGrid');
    const fnPr = child(sectPr, 'w:footnotePr');
    return {
      pageWidth,
      pageHeight,
      marginTop: twip(Math.abs(topRaw)),
      marginBottom: twip(Math.abs(bottomRaw)),
      marginTopExact: topRaw < 0,
      marginBottomExact: bottomRaw < 0,
      marginLeft: t(pgMar, attr(pgMar, 'w:left') === undefined && attr(pgMar, 'w:start') !== undefined ? 'w:start' : 'w:left', DEFAULT_SECTION_TWIPS.margin),
      marginRight: t(pgMar, attr(pgMar, 'w:right') === undefined && attr(pgMar, 'w:end') !== undefined ? 'w:end' : 'w:right', DEFAULT_SECTION_TWIPS.margin),
      header: t(pgMar, 'w:header', DEFAULT_SECTION_TWIPS.header),
      footer: t(pgMar, 'w:footer', DEFAULT_SECTION_TWIPS.footer),
      gutter: t(pgMar, 'w:gutter', 0),
      columns,
      type: (['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'].includes(typeVal ?? '') ? typeVal : 'nextPage') as Section['type'],
      titlePage: !!onOff(child(sectPr, 'w:titlePg')),
      headers,
      footers,
      pageNumberStart: num(attr(pgNum, 'w:start')) ?? null,
      pageNumberFormat: attr(pgNum, 'w:fmt') ?? 'decimal',
      vAlign: (attr(child(sectPr, 'w:vAlign'), 'w:val') ?? 'top') as Section['vAlign'],
      gridType: attr(grid, 'w:type') ?? 'default',
      gridLinePitch: twip(num(attr(grid, 'w:linePitch')) ?? 360),
      gridCharSpace: (num(attr(grid, 'w:charSpace')) ?? 0) / 4096,
      pageBorders: this.parsePageBorders(child(sectPr, 'w:pgBorders')),
      footnoteRestart: (attr(child(fnPr, 'w:numRestart'), 'w:val') ?? 'continuous') as Section['footnoteRestart'],
      footnoteStart: num(attr(child(fnPr, 'w:numStart'), 'w:val')) ?? 1,
      footnoteFormat: attr(child(fnPr, 'w:numFmt'), 'w:val') ?? 'decimal',
    };
  }

  private parsePageBorders(el: XmlElement | undefined): PageBorders | null {
    if (!el) return null;
    const tc = themeColorFn(this.ctx);
    const side = (n: string) => parseBorder(child(el, n), tc) ?? null;
    const pb: PageBorders = {
      top: side('w:top'),
      bottom: side('w:bottom'),
      left: side('w:left'),
      right: side('w:right'),
      offsetFrom: el.attrs['w:offsetFrom'] === 'page' ? 'page' : 'text',
      display: (el.attrs['w:display'] ?? 'allPages') as PageBorders['display'],
      zOrder: el.attrs['w:zOrder'] === 'back' ? 'back' : 'front',
    };
    if (!pb.top && !pb.bottom && !pb.left && !pb.right) return null;
    return pb;
  }

  private headerFooter(part: string, rootName: string): Block[] {
    const cached = this.headerCache.get(part);
    if (cached) return cached;
    const root = this.pkg.xml(part);
    const savedFields = this.fieldStack;
    this.fieldStack = [];
    const blocks = root && root.name === rootName ? this.parseBlocks(root, { part }) : root ? this.parseBlocks(root, { part }) : [];
    this.fieldStack = savedFields;
    this.headerCache.set(part, blocks);
    return blocks;
  }
}

/**
 * The node a content control's binding points at. Word writes simple paths of named, indexed
 * steps (/ns0:root[1]/ns0:Company[1]); they are matched by local name, as the prefixes are the
 * binding's own. Anything else does not resolve.
 */
function selectXPath(root: XmlElement, xpath: string): XmlElement | undefined {
  const steps = xpath.split('/');
  if (steps.shift() !== '' || !steps.length) return undefined;
  let node: XmlElement | undefined;
  for (const step of steps) {
    const m = /^(?:[\w.-]+:)?([\w.-]+)(?:\[(\d+)\])?$/.exec(step);
    if (!m) return undefined;
    const local = (el: XmlElement) => el.name.slice(el.name.indexOf(':') + 1) === m[1];
    const matches = node ? elements(node).filter(local) : [root].filter(local);
    node = matches[Number(m[2] ?? 1) - 1];
    if (!node) return undefined;
  }
  return node;
}

/** A run showing a bound content control's value in the formatting of the control's content. */
function boundRun(content: XmlElement, sdtPr: XmlElement | undefined, text: string): XmlElement {
  const rPr = child(descendant(content, 'w:r'), 'w:rPr') ?? child(sdtPr, 'w:rPr');
  const runs: XmlElement[] = [];
  text.split(/\r?\n/).forEach((line, i) => {
    if (i) runs.push({ name: 'w:br', attrs: {}, children: [] });
    if (line) runs.push({ name: 'w:t', attrs: { 'xml:space': 'preserve' }, children: [line] });
  });
  return { name: 'w:r', attrs: {}, children: rPr ? [rPr, ...runs] : runs };
}


/**
 * Sizes spacing given in lines (w:beforeLines, w:afterLines). Word counts a line as the section's
 * grid pitch when the section has a line grid, and as 12pt otherwise. A section's properties sit
 * on its last paragraph, so the walk runs backwards.
 */
function resolveLineSpacing(blocks: Block[], finalSection: Section): void {
  let pitch = 12;
  const setPitch = (s: Section) => (pitch = s.gridType === 'lines' || s.gridType === 'linesAndChars' ? s.gridLinePitch : 12);
  const visit = (b: Block): void => {
    if (b.type === 'table') {
      for (const row of b.rows) for (const cell of row.cells) for (const c of cell.content) visit(c);
      return;
    }
    if (b.props.spaceBeforeLines) b.props.spaceBefore = b.props.spaceBeforeLines * pitch;
    if (b.props.spaceAfterLines) b.props.spaceAfter = b.props.spaceAfterLines * pitch;
  };
  setPitch(finalSection);
  for (let k = blocks.length - 1; k >= 0; k--) {
    const b = blocks[k];
    if (b.type === 'paragraph' && b.sectPr) setPitch(b.sectPr);
    visit(b);
  }
}
/** The outline levels a TOC field lists (\\o), or null for fields built from other switches. */
function tocLevels(instr: string): [number, number] | null {
  const o = /\\o\s*"?\s*(\d)\s*-\s*(\d)/i.exec(instr);
  if (o) return [Number(o[1]), Number(o[2])];
  return /\\[tfcabls]\b/i.test(instr) ? null : [1, 9];
}

/** A heading's text as its table-of-contents entry shows it, list number first. */
function headingText(p: Paragraph): string {
  const label = p.numbering?.text ? p.numbering.text + (p.numbering.suffix === 'tab' ? '\t' : p.numbering.suffix === 'space' ? ' ' : '') : '';
  let text = '';
  for (const c of p.content) {
    if (c.type === 'text' && !c.props.hidden) text += c.text;
    else if (c.type === 'tab') text += '\t';
    else if (c.type === 'field') text += c.cached;
  }
  return text.trim() ? label + text : '';
}

function parseFloating(el: XmlElement) {
  const t = (n: string) => measure(el.attrs[n]) ?? 0;
  return {
    horzAnchor: el.attrs['w:horzAnchor'] ?? 'text',
    vertAnchor: el.attrs['w:vertAnchor'] ?? 'margin',
    x: t('w:tblpX'),
    y: t('w:tblpY'),
    xAlign: el.attrs['w:tblpXSpec'],
    yAlign: el.attrs['w:tblpYSpec'],
    leftFromText: t('w:leftFromText'),
    rightFromText: t('w:rightFromText'),
    topFromText: t('w:topFromText'),
    bottomFromText: t('w:bottomFromText'),
  };
}

/** Classify a field instruction. */
export function analyzeField(instr: string): { kind?: 'page' | 'numpages' | 'sectionpages'; format?: string; link?: Link } {
  const s = instr.trim();
  const m = /^(\S+)/.exec(s);
  const name = (m?.[1] ?? '').toUpperCase();
  const fmt = /\\\*\s*(\w+)/.exec(s)?.[1];
  const format = fmt && fmt.toUpperCase() !== 'MERGEFORMAT' && fmt.toUpperCase() !== 'CHARFORMAT' ? fmt : undefined;
  if (name === 'PAGE') return { kind: 'page', format };
  if (name === 'NUMPAGES') return { kind: 'numpages', format };
  if (name === 'SECTIONPAGES') return { kind: 'sectionpages', format };
  if (name === 'HYPERLINK') {
    const link: Link = {};
    const url = /^HYPERLINK\s+"([^"]*)"/i.exec(s) ?? /^HYPERLINK\s+(\S+)/i.exec(s);
    if (url && !url[1].startsWith('\\')) link.url = url[1];
    const anchor = /\\l\s+"([^"]*)"/.exec(s);
    if (anchor) link.anchor = anchor[1];
    return { link };
  }
  return {};
}

export { emu, descendant };
