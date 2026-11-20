/**
 * Resolved document model. Style inheritance, numbering and table-style conditional formatting
 * are already applied, so the layout engine works with final values. Lengths are in points.
 */

export type Color = string | null; // 'RRGGBB' or null for automatic

export interface Border {
  style: string; // single, double, dotted, dashed, thick, ...
  width: number; // pt
  space: number; // pt
  color: Color;
}

export interface FontRequest {
  family: string;
  bold: boolean;
  italic: boolean;
  codePoints: Set<number>;
  /** East Asian language of the text, which picks the region of East Asian fallback fonts. */
  lang?: string;
}

export interface RunProps {
  fontAscii: string;
  fontHAnsi: string;
  fontEastAsia: string;
  fontCs: string;
  hint?: string;
  /** w:lang w:eastAsia: the language of East Asian text (zh-CN, ja-JP, ko-KR...). */
  eastAsiaLang?: string;
  /** Text standing in for a shape Word draws itself (a legacy form check box), not a font glyph. */
  drawn?: boolean;
  size: number;
  sizeCs: number;
  bold: boolean;
  italic: boolean;
  boldCs: boolean;
  italicCs: boolean;
  underline: string | null;
  underlineColor: Color;
  strike: boolean;
  dstrike: boolean;
  color: Color;
  highlight: Color;
  shading: Color;
  vertAlign: 'baseline' | 'superscript' | 'subscript';
  caps: boolean;
  smallCaps: boolean;
  spacing: number;
  scale: number;
  position: number;
  hidden: boolean;
  kern: number;
  rtl: boolean;
  /** Whether the text follows the section's character grid (w:snapToGrid). */
  snapToGrid?: boolean;
  cs: boolean;
  border: Border | null;
  outline: boolean;
  emboss: boolean;
  imprint: boolean;
  shadow: boolean;
  styleId?: string;
}

export type Alignment = 'left' | 'center' | 'right' | 'both' | 'distribute';

export interface TabStop {
  pos: number;
  type: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'num';
  leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface ParaBorders {
  top: Border | null;
  bottom: Border | null;
  left: Border | null;
  right: Border | null;
  between: Border | null;
  bar: Border | null;
}

export interface ParaProps {
  styleId: string;
  align: Alignment;
  indLeft: number;
  indRight: number;
  /** Positive: first-line indent; negative: hanging indent. */
  indFirstLine: number;
  spaceBefore: number;
  spaceAfter: number;
  /** Spacing given in lines (w:beforeLines, w:afterLines), which the section's line grid sizes. */
  spaceBeforeLines?: number;
  spaceAfterLines?: number;
  spaceBeforeAuto: boolean;
  spaceAfterAuto: boolean;
  /** For 'auto' this is a multiplier (1 = single); otherwise points. */
  line: number;
  lineRule: 'auto' | 'exact' | 'atLeast';
  contextualSpacing: boolean;
  keepNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
  tabs: TabStop[];
  borders: ParaBorders;
  shading: Color;
  outlineLevel: number | null;
  bidi: boolean;
  snapToGrid: boolean;
  /** Space between East Asian text and Latin letters (w:autoSpaceDE) or digits (w:autoSpaceDN). */
  autoSpaceDE: boolean;
  autoSpaceDN: boolean;
  /** Whether full-width punctuation may hang past the right edge instead of wrapping (w:overflowPunct). */
  overflowPunct: boolean;
  suppressAutoHyphens: boolean;
  textAlignment: string;
  frame: FrameProps | null;
}

export interface FrameProps {
  width?: number;
  height?: number;
  hRule?: string;
  x?: number;
  y?: number;
  xAlign?: string;
  yAlign?: string;
  hAnchor?: string;
  vAnchor?: string;
  wrap?: string;
  hSpace: number;
  vSpace: number;
  dropCap?: string;
  lines?: number;
}

export interface NumberingLabel {
  text: string;
  props: RunProps;
  suffix: 'tab' | 'space' | 'nothing';
  align: 'left' | 'center' | 'right';
  /** Picture bullet image, when the level uses one. */
  picture?: ImageRef;
  /** The list (w:numId) the paragraph belongs to. */
  numId: string;
}

export interface Link {
  url?: string;
  anchor?: string;
}

export interface TextInline {
  type: 'text';
  text: string;
  props: RunProps;
  link?: Link;
}
export interface TabInline {
  type: 'tab';
  props: RunProps;
  link?: Link;
}
/** Positional tab (w:ptab), common in headers and footers. */
export interface PTabInline {
  type: 'ptab';
  props: RunProps;
  alignment: 'left' | 'center' | 'right';
  relativeTo: 'margin' | 'indent';
  leader: TabStop['leader'];
}
export interface BreakInline {
  type: 'break';
  breakType: 'line' | 'page' | 'column';
  clear?: string;
  props: RunProps;
}
export interface DrawingInline {
  type: 'drawing';
  drawing: Drawing;
  props: RunProps;
  link?: Link;
}
export interface FieldInline {
  type: 'field';
  field: 'page' | 'numpages' | 'sectionpages' | 'pageref';
  format?: string;
  /** Bookmark whose page a page reference shows. */
  bookmark?: string;
  props: RunProps;
  /** Word's cached result, used for measuring before the value is known. */
  cached: string;
}
export interface NoteRefInline {
  type: 'noteRef';
  noteType: 'footnote' | 'endnote';
  id: string;
  props: RunProps;
  customMark?: string;
}
/** w:footnoteRef / w:endnoteRef inside the note text: renders the note's number. */
export interface NoteNumberInline {
  type: 'noteNumber';
  props: RunProps;
}
export interface BookmarkInline {
  type: 'bookmark';
  name: string;
}
export interface SeparatorInline {
  type: 'separator';
  continuation: boolean;
  props: RunProps;
}
/** Word's record of where it broke the page when it last laid the document out. */
export interface RenderedBreakInline {
  type: 'renderedPageBreak';
}

/**
 * Office Math (OMML) structure. Text keeps its run formatting; the other nodes carry the
 * formatting of their own characters (operators, delimiters, accents) in `props`.
 */
export type MathNode =
  | {
      k: 'text';
      text: string;
      props: RunProps;
      /** Letters stay upright (m:sty p or b, function names) instead of math italic. */
      upright?: boolean;
      /** Ordinary text (m:nor, or a w:r inside math): no math italics or operator spacing. */
      plain?: boolean;
    }
  | { k: 'row'; items: MathNode[] }
  | { k: 'frac'; type: 'bar' | 'skw' | 'lin' | 'noBar'; num: MathNode; den: MathNode; props: RunProps }
  | { k: 'script'; base: MathNode; sub?: MathNode; sup?: MathNode; pre?: boolean }
  | { k: 'rad'; deg?: MathNode; e: MathNode; props: RunProps }
  | { k: 'nary'; chr: string; under: boolean; sub?: MathNode; sup?: MathNode; e: MathNode; props: RunProps }
  | { k: 'delim'; beg: string; end: string; sep: string; items: MathNode[]; props: RunProps }
  | { k: 'func'; name: MathNode; e: MathNode }
  | { k: 'limit'; e: MathNode; lim: MathNode; over: boolean }
  | { k: 'accent'; chr: string; e: MathNode; props: RunProps }
  | { k: 'bar'; over: boolean; e: MathNode; props: RunProps }
  | { k: 'group'; chr: string; over: boolean; e: MathNode; props: RunProps }
  | { k: 'matrix'; rows: MathNode[][]; props: RunProps; /** An equation array (m:eqArr): rows aligned at their & marks. */ eqArr?: boolean }
  | { k: 'box'; e: MathNode; border: boolean; props: RunProps }
  | { k: 'phantom'; e: MathNode; show: boolean; zeroW: boolean; zeroAsc: boolean; zeroDesc: boolean };

/** An equation: inline (m:oMath in running text) or displayed on a line of its own (m:oMathPara). */
export interface MathInline {
  type: 'math';
  math: MathNode;
  display: boolean;
  /** A displayed equation's alignment, which overrides the paragraph's. */
  align?: 'left' | 'center' | 'right';
  props: RunProps;
}

export type Inline =
  | TextInline
  | TabInline
  | PTabInline
  | BreakInline
  | DrawingInline
  | FieldInline
  | NoteRefInline
  | NoteNumberInline
  | BookmarkInline
  | SeparatorInline
  | RenderedBreakInline
  | MathInline;

export interface Paragraph {
  type: 'paragraph';
  props: ParaProps;
  /** Formatting of the paragraph mark (affects empty-line height and list labels). */
  markProps: RunProps;
  content: Inline[];
  numbering: NumberingLabel | null;
  sectPr?: Section;
}

export interface CellMargins {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface TableWidth {
  type: 'dxa' | 'pct' | 'auto' | 'nil';
  value: number; // pt for dxa; fraction (0..1) for pct
}

export interface TableProps {
  styleId?: string;
  width: TableWidth;
  align: 'left' | 'center' | 'right';
  indent: number;
  layout: 'fixed' | 'autofit';
  cellMargins: CellMargins;
  cellSpacing: number;
  floating: FloatingTable | null;
  bidiVisual: boolean;
}

export interface FloatingTable {
  horzAnchor: string;
  vertAnchor: string;
  x: number;
  y: number;
  xAlign?: string;
  yAlign?: string;
  leftFromText: number;
  rightFromText: number;
  topFromText: number;
  bottomFromText: number;
}

export interface CellBorders {
  top: Border | null;
  bottom: Border | null;
  left: Border | null;
  right: Border | null;
  tl2br: Border | null;
  tr2bl: Border | null;
}

export interface TableCell {
  gridSpan: number;
  vMerge: 'restart' | 'continue' | null;
  width: TableWidth;
  borders: CellBorders;
  shading: Color;
  margins: CellMargins;
  vAlign: 'top' | 'center' | 'bottom';
  noWrap: boolean;
  /** w:hideMark: an empty cell's paragraph mark doesn't count toward the row height. */
  hideMark: boolean;
  textDirection: string;
  content: Block[];
}

export interface TableRow {
  height: number;
  heightRule: 'auto' | 'atLeast' | 'exact';
  cantSplit: boolean;
  isHeader: boolean;
  gridBefore: number;
  gridAfter: number;
  /** Space around each cell of the row (w:tblCellSpacing), so adjacent cells sit twice it apart. */
  cellSpacing: number;
  /** The row carries revision ids (w:rsidR and the like), which Word stamps on rows it saves. */
  stamped: boolean;
  cells: TableCell[];
}

export interface Table {
  type: 'table';
  props: TableProps;
  grid: number[];
  rows: TableRow[];
}

export type Block = Paragraph | Table;

export interface Columns {
  count: number;
  space: number;
  equalWidth: boolean;
  separator: boolean;
  cols: Array<{ width: number; space: number }>;
}

export interface HeaderFooterRefs {
  default?: Block[];
  first?: Block[];
  even?: Block[];
}

export interface PageBorders {
  top: Border | null;
  bottom: Border | null;
  left: Border | null;
  right: Border | null;
  offsetFrom: 'page' | 'text';
  display: 'allPages' | 'firstPage' | 'notFirstPage';
  zOrder: 'front' | 'back';
}

export interface Section {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  /** Negative top/bottom margins in Word mean "exact": the header must not push the body down. */
  marginTopExact: boolean;
  marginBottomExact: boolean;
  header: number;
  footer: number;
  gutter: number;
  columns: Columns;
  type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  titlePage: boolean;
  headers: HeaderFooterRefs;
  footers: HeaderFooterRefs;
  /** Which header/footer types this section defines itself (others inherit from the previous section). */
  pageNumberStart: number | null;
  pageNumberFormat: string;
  vAlign: 'top' | 'center' | 'both' | 'bottom';
  gridType: string;
  gridLinePitch: number;
  /** Extra width per character of a "linesAndChars" grid (w:charSpace), in points. */
  gridCharSpace: number;
  pageBorders: PageBorders | null;
  footnoteRestart: 'continuous' | 'eachSect' | 'eachPage';
  footnoteStart: number;
  footnoteFormat: string;
  /** Text runs down the page in lines that follow one another leftwards (w:textDirection tbRl), as vertical Japanese is set. */
  vertical?: boolean;
}

export interface ImageRef {
  id: string;
  data: Uint8Array;
  /** Metafile raster with a separate pixel array. In this case data holds the DIB header and palette. */
  dibBits?: Uint8Array;
  contentType?: string;
  svg?: ImageRef;
  /** Brightness and contrast (Word's picture corrections and Washout): each colour component c (0–1) is drawn as slope·c + offset, clipped. */
  adjust?: { slope: number; offset: number };
}

export interface Fill {
  color: Color;
  alpha: number;
  /** A gradient fill; `color` stands in for it where one colour must do. */
  gradient?: GradientFill;
}

/** A DrawingML gradient: colour stops at positions 0..1, along a line or out from a focus. */
export interface GradientFill {
  /** Colour stops, with their opacity where it is below 1. */
  stops: Array<{ pos: number; color: string; alpha?: number }>;
  /** A linear gradient's direction, in degrees clockwise from east. */
  angle: number;
  /** The angle is for the unit square, stretched with the shape (a:lin scaled). */
  scaled?: boolean;
  /** A path gradient runs from its focus rectangle (insets, as fractions of the box) out to the edges. */
  path?: 'circle' | 'rect' | 'shape';
  focus?: { l: number; t: number; r: number; b: number };
  /**
   * Blend in linear light, easing in and out from stop to stop. Word draws the gradients of charts,
   * SmartArt and drawing canvases so; an ordinary shape's gradient is a straight blend in sRGB.
   */
  eased?: boolean;
}

export interface LineStyle {
  color: Color;
  width: number;
  dash: string;
  alpha: number;
}

export interface PositionSpec {
  relativeFrom: string;
  offset?: number;
  align?: string;
  pct?: number;
}

export interface AnchorProps {
  behindDoc: boolean;
  wrap: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  wrapText: 'bothSides' | 'left' | 'right' | 'largest';
  distT: number;
  distB: number;
  distL: number;
  distR: number;
  /** The outline tight and through wrapping follow, as fractions of the object's extent. */
  polygon?: Array<[number, number]>;
  posH: PositionSpec;
  posV: PositionSpec;
  layoutInCell: boolean;
  allowOverlap: boolean;
  zIndex: number;
}

export interface TextBox {
  blocks: Block[];
  insetL: number;
  insetT: number;
  insetR: number;
  insetB: number;
  vAlign: 'top' | 'center' | 'bottom';
  autoFit: boolean;
  vertical?: string;
  /** Draw content past the box edges (floating tables, whose border halves sit outside the grid). */
  unclipped?: boolean;
  /** The text's area, as distances in from the shape's left, top, right and bottom; text is clipped to it. */
  rect?: [number, number, number, number];
  /** The text stays level when the shape turns, as in Word's legacy (VML) shapes. */
  upright?: boolean;
}

/** Text style of a chart element; colours are hex RGB. */
export interface ChartTextStyle {
  font: string;
  size: number;
  bold: boolean;
  italic: boolean;
  color: string;
  /** Text rotation in degrees, clockwise; undefined lets the chart choose (category labels turn when crowded). */
  rotation?: number;
}

export interface ChartLine {
  color: string;
  width: number;
  dash?: string;
  alpha?: number;
}

/** Fill and outline of a chart element: null draws none. */
export interface ChartShape {
  /** A gradient fill keeps its stops (positions 0..1) and direction; `color` stands in for it where one colour must do. */
  fill: { color: string; alpha: number; gradient?: GradientFill } | null;
  line: ChartLine | null;
  /** An outer shadow (a:outerShdw): the shape again, offset by (dx, dy) pt and blurred by `blur` pt. */
  shadow?: { dx: number; dy: number; blur: number; color: string; alpha: number };
}

/** A series' trendline (c:trendline), fitted to its values. */
export interface ChartTrendline {
  /** linear, exp, log, poly, power or movingAvg. */
  type: string;
  /** The polynomial's order or the moving average's period. */
  order: number;
  /** How far the line runs on past the last x value and back before the first. */
  forward: number;
  backward: number;
  /** A fixed y intercept (linear, exponential and polynomial fits). */
  intercept?: number;
  /** The legend entry's own name. */
  name?: string;
  line: ChartLine | null;
}

export interface ChartLabels {
  value: boolean;
  category: boolean;
  series: boolean;
  percent: boolean;
  /** c:dLblPos: ctr, inEnd, inBase, outEnd, t, b, l, r, bestFit. */
  position?: string;
  style: ChartTextStyle;
  format?: string;
  separator?: string;
  /** Points whose label is deleted. */
  hidden?: number[];
}

export interface ChartMarker {
  /** none, circle, square, diamond, triangle, x, star, plus, dash, dot. */
  symbol: string;
  size: number;
  shape: ChartShape;
}

export interface ChartSeries {
  name: string;
  categories: string[];
  /** Category values when the categories are numbers or dates, for number formats. */
  categoryNumbers?: Array<number | null>;
  categoryFormat?: string;
  values: Array<number | null>;
  /** Scatter and bubble x values. */
  xValues?: Array<number | null>;
  format: string;
  shape: ChartShape;
  marker?: ChartMarker;
  smooth: boolean;
  labels?: ChartLabels;
  /** Per-point formatting (c:dPt), by point index. */
  points: Map<number, ChartShape>;
  /** Pie slice explosion (percent of radius). */
  explosion: number;
  trendlines?: ChartTrendline[];
  errorBars?: ChartErrorBars[];
}

/** A series' error bars (c:errBars): an amount either side of each point. */
export interface ChartErrorBars {
  /** Along the value axis, or a scatter plot's x axis. */
  dir: 'x' | 'y';
  type: 'both' | 'plus' | 'minus';
  valType: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust';
  value: number;
  /** Custom amounts, by point. */
  plus?: Array<number | null>;
  minus?: Array<number | null>;
  endCap: boolean;
  line: ChartLine | null;
}

export interface ChartTitle {
  /** Paragraphs of runs. */
  paragraphs: Array<Array<{ text: string; style: ChartTextStyle }>>;
  overlay: boolean;
  /** Manual position as fractions of the chart (top-left of the title). */
  x?: number;
  y?: number;
  shape: ChartShape;
}

export interface ChartAxis {
  id: string;
  kind: 'cat' | 'val' | 'date' | 'ser';
  position: 'l' | 'r' | 't' | 'b';
  deleted: boolean;
  reverse: boolean;
  min?: number;
  max?: number;
  majorUnit?: number;
  logBase?: number;
  format: string;
  /** Use the series' own number format (c:numFmt sourceLinked). */
  sourceLinked: boolean;
  line: ChartLine | null;
  gridlines: ChartLine | null;
  minorGridlines: ChartLine | null;
  majorTick: 'out' | 'in' | 'cross' | 'none';
  labelPosition: 'nextTo' | 'low' | 'high' | 'none';
  style: ChartTextStyle;
  title?: ChartTitle;
  crossAxis: string;
  crosses: 'autoZero' | 'min' | 'max' | number;
  /** Categories sit between tick marks (bars) rather than on them. */
  between: boolean;
  labelSkip?: number;
}

export interface ChartPlot {
  type: 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar';
  /** Radar plots: lines (with or without markers) or filled shapes. */
  radarStyle?: 'standard' | 'marker' | 'filled';
  horizontal: boolean;
  grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  gapWidth: number;
  overlap: number;
  varyColors: boolean;
  holeSize: number;
  firstSliceAngle: number;
  /** Line and scatter plots: whether series without their own marker get one. */
  markers: boolean;
  /** Scatter: lines between points unless the style is marker-only. */
  lines: boolean;
  series: ChartSeries[];
  axes: string[];
}

export interface ChartLayout {
  /** The rectangle is the plot's inner area (without axis labels) rather than its outer one. */
  inner: boolean;
  x: number;
  y: number;
  w: number;
  h: number;
}

/** A chart drawn from its part's cached data. */
export interface Chart {
  title?: ChartTitle;
  plots: ChartPlot[];
  axes: ChartAxis[];
  legend?: { position: 'r' | 'l' | 't' | 'b' | 'tr'; overlay: boolean; style: ChartTextStyle; shape: ChartShape; hidden: number[]; layout?: ChartLayout };
  plotLayout?: ChartLayout;
  space: ChartShape;
  plotShape: ChartShape;
}

export interface Drawing {
  kind: 'image' | 'shape' | 'group' | 'unsupported';
  inline: boolean;
  width: number;
  height: number;
  effectL: number;
  effectT: number;
  effectR: number;
  effectB: number;
  anchor?: AnchorProps;
  // Position inside a group (pt, relative to group origin), for children of groups.
  x?: number;
  y?: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  image?: ImageRef;
  /** Crop fractions (0..1) of the source image. */
  crop?: { l: number; t: number; r: number; b: number };
  geometry: string;
  /** Adjust values for preset geometries (e.g. roundRect corner radius), raw OOXML units. */
  adjust?: Record<string, number>;
  /** Custom geometry subpaths in shape-relative coordinates (0..1), each saying whether the fill and the outline draw it. */
  path?: Array<{ fill: boolean; stroke: boolean; segs: Array<{ op: 'M' | 'L' | 'C' | 'Z'; pts: number[] }> }>;
  fill?: Fill | null;
  line?: LineStyle | null;
  textBox?: TextBox;
  children?: Drawing[];
  /** Watermark-style text drawn along the shape (VML textpath / WordArt). */
  textPath?: { text: string; font: string; bold: boolean; italic: boolean; color: Color; alpha: number };
  /** A chart, drawn in the drawing's box. */
  chart?: Chart;
  hidden: boolean;
  name?: string;
  /**
   * A VML horizontal rule (o:hr): aligned on its own line whatever the paragraph's alignment, and
   * when `pct` is set, that share of the text column wide.
   */
  hr?: { align: 'left' | 'center' | 'right'; pct: number };
}

export interface Note {
  id: string;
  blocks: Block[];
}

export interface DocumentInfo {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  created?: Date;
  modified?: Date;
}

export interface DocSettings {
  defaultTabStop: number;
  /** Word updates the document's fields when it opens it (w:updateFields). */
  updateFields?: boolean;
  evenAndOddHeaders: boolean;
  mirrorMargins: boolean;
  compatibilityMode: number;
  doNotExpandShiftReturn: boolean;
  /** WordPerfect justification (compat wpJustification): justified lines may also compress spaces. */
  wpJustification: boolean;
  /** East Asian punctuation may give up its blank half to fit a line (w:characterSpacingControl). */
  compressPunctuation?: boolean;
  /** Word 2013+ table-style rule (compat setting overrideTableStyleFontSizeAndJustification). */
  overrideTableStyleFontSizeAndJustification: boolean;
  autoHyphenation: boolean;
  displayBackgroundShape: boolean;
  doNotUseHTMLParagraphAutoSpacing: boolean;
  /** Lines in table cells snap to the document grid too (compat adjustLineHeightInTable). */
  adjustLineHeightInTable: boolean;
  footnotePosition: string;
  endnotePosition: string;
  /** Style id of paragraphs without a style (Normal), which Word's table row-end marks also take. */
  defaultParagraphStyle?: string;
  /** The settings carry no revision ids (w:rsids): another tool wrote the file, not Word. */
  generated?: boolean;
}

export interface DocxDocument {
  /** Body blocks; paragraphs carrying `sectPr` end a section. The final section is `finalSection`. */
  body: Block[];
  finalSection: Section;
  settings: DocSettings;
  footnotes: Map<string, Note>;
  endnotes: Map<string, Note>;
  footnoteSeparator: Block[] | null;
  footnoteContinuationSeparator: Block[] | null;
  endnoteSeparator: Block[] | null;
  info: DocumentInfo;
  background: Color;
  /** Equation layout measures ink bounds directly from glyph outlines. */
  hasMath: boolean;
  /** Families referenced anywhere, for font preloading. */
  fontRequests: FontRequest[];
  embeddedFonts: Array<{ family: string; bold: boolean; italic: boolean; data: Uint8Array }>;
  fontHints: Map<string, { family?: string; pitch?: string; panose?: string; charset?: string; altName?: string }>;
  warnings: string[];
}
