import { child, children, descendant, path, type XmlElement } from '../util/xml.js';
import type { Chart, ChartAxis, ChartLabels, ChartLayout, ChartLine, ChartMarker, ChartPlot, ChartSeries, ChartShape, ChartTextStyle, ChartTitle, ChartTrendline, ChartErrorBars } from './model.js';
import type { Theme } from './theme.js';
import { parseGradient } from './gradient.js';
import { emu } from './values.js';

type ColorOf = (el: XmlElement | undefined) => { color: string; alpha: number } | undefined;

interface ChartContext {
  theme: Theme;
  budget: ChartPointBudget;
  color: ColorOf;
  /** Default series colour scheme from the chart style (c:style 1-48). */
  scheme: number;
}

const isColor = (c: XmlElement) => /^a:(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(c.name);
const el = (name: string, attrs: Record<string, string> = {}, kids: XmlElement[] = []): XmlElement => ({ name, attrs, children: kids }) as XmlElement;
const scheme = (val: string, mods: Record<string, number> = {}) =>
  el(
    'a:schemeClr',
    { val },
    Object.entries(mods).map(([k, v]) => el(`a:${k}`, { val: String(v) })),
  );

/** A c: boolean element: present means true unless val says otherwise. */
function flag(e: XmlElement | undefined, dflt: boolean): boolean {
  if (!e) return dflt;
  const v = e.attrs.val;
  return v === undefined || (v !== '0' && v !== 'false');
}

const val = (e: XmlElement | undefined) => e?.attrs.val;
const numVal = (e: XmlElement | undefined) => {
  // Strict charts write percentages (gap width, overlap, hole size) with a % sign.
  const v = Number(e?.attrs.val?.replace(/%$/, ''));
  return e && Number.isFinite(v) ? v : undefined;
};

function fontName(ctx: ChartContext, typeface: string | undefined): string | undefined {
  if (!typeface) return undefined;
  if (typeface.startsWith('+mj')) return ctx.theme.majorLatin;
  if (typeface.startsWith('+mn')) return ctx.theme.minorLatin;
  return typeface;
}

/** Run properties (a:defRPr / a:rPr) over a base style. */
function runStyle(rPr: XmlElement | undefined, base: ChartTextStyle, ctx: ChartContext): ChartTextStyle {
  if (!rPr) return base;
  const s = { ...base };
  const sz = Number(rPr.attrs.sz);
  if (sz > 0) s.size = sz / 100;
  if (rPr.attrs.b !== undefined) s.bold = rPr.attrs.b === '1';
  if (rPr.attrs.i !== undefined) s.italic = rPr.attrs.i === '1';
  const fill = child(rPr, 'a:solidFill');
  const c = fill ? ctx.color(children(fill).find(isColor)) : undefined;
  if (c) s.color = c.color;
  s.font = fontName(ctx, child(rPr, 'a:latin')?.attrs.typeface) ?? s.font;
  return s;
}

/** A c:txPr (or c:rich) text style over a base: its first paragraph's defaults and its rotation. */
function textStyle(txPr: XmlElement | undefined, base: ChartTextStyle, ctx: ChartContext): ChartTextStyle {
  if (!txPr) return base;
  const s = runStyle(path(txPr, 'a:p', 'a:pPr', 'a:defRPr'), base, ctx);
  const rot = child(txPr, 'a:bodyPr')?.attrs.rot;
  // -60000000 is the "no custom angle" sentinel.
  if (rot !== undefined && rot !== '-60000000') s.rotation = Number(rot) / 60000;
  const vert = child(txPr, 'a:bodyPr')?.attrs.vert;
  if (vert === 'vert') s.rotation = 90;
  else if (vert === 'vert270') s.rotation = -90;
  return s;
}

function line(ln: XmlElement | undefined, ctx: ChartContext, dflt: ChartLine | null): ChartLine | null {
  if (!ln) return dflt;
  if (child(ln, 'a:noFill')) return null;
  const fill = child(ln, 'a:solidFill') ?? child(ln, 'a:gradFill');
  const c = fill ? ctx.color(children(fill).find(isColor) ?? descendant(fill, 'a:gs')?.children.find((k): k is XmlElement => typeof k !== 'string' && isColor(k))) : undefined;
  const width = ln.attrs.w !== undefined ? emu(Number(ln.attrs.w) || 0) : (dflt?.width ?? 0.75);
  const dash = val(child(ln, 'a:prstDash'));
  if (!c && !dflt) return null;
  return { color: c?.color ?? dflt!.color, width, ...(dash && dash !== 'solid' && { dash }), ...(c && c.alpha < 1 && { alpha: c.alpha }) };
}

/** A c:spPr over defaults: undefined parts keep the default. */
function shape(spPr: XmlElement | undefined, ctx: ChartContext, dflt: ChartShape): ChartShape {
  if (!spPr) return dflt;
  let fill = dflt.fill;
  if (child(spPr, 'a:noFill')) fill = null;
  else {
    const solid = child(spPr, 'a:solidFill');
    const grad = child(spPr, 'a:gradFill');
    const c = solid ? ctx.color(children(solid).find(isColor)) : undefined;
    if (c) fill = c;
    else if (grad) {
      const g = parseGradient(grad, ctx.color);
      if (g) fill = { color: g.mid.color, alpha: g.mid.alpha, ...(g.gradient.stops.length > 1 && { gradient: { ...g.gradient, eased: true } }) };
    }
  }
  const out: ChartShape = { fill, line: line(child(spPr, 'a:ln'), ctx, dflt.line) };
  const effects = child(spPr, 'a:effectLst');
  const shdw = effects ? child(effects, 'a:outerShdw') : undefined;
  const sc = shdw ? ctx.color(children(shdw).find(isColor)) : undefined;
  if (shdw && sc) {
    const dist = emu(Number(shdw.attrs.dist) || 0);
    const dir = ((Number(shdw.attrs.dir) || 0) / 60000) * (Math.PI / 180);
    out.shadow = { dx: dist * Math.cos(dir), dy: dist * Math.sin(dir), blur: emu(Number(shdw.attrs.blurRad) || 0), color: sc.color, alpha: sc.alpha };
  } else if (!effects && dflt.shadow) out.shadow = dflt.shadow;
  return out;
}

/** The default colour of series (or pie point) `i` in the chart style's colour scheme. */
function autoColor(i: number, ctx: ChartContext): string {
  const round = Math.floor(i / 6);
  // Later rounds darken, then lighten, the accent cycle (Office's chart palette).
  const rounds: Array<Record<string, number>> = [{}, { lumMod: 60000 }, { lumMod: 80000, lumOff: 20000 }, { lumMod: 80000 }, { lumMod: 60000, lumOff: 40000 }, { lumMod: 50000 }];
  const mods = rounds[round % rounds.length];
  if (ctx.scheme === 1) return ctx.color(scheme('tx1', { tint: [100000, 75000, 50000, 85000, 60000, 35000][i % 6] }))?.color ?? '808080';
  if (ctx.scheme >= 3 && ctx.scheme <= 8) {
    const shades: Array<Record<string, number>> = [{}, { shade: 76000 }, { tint: 77000 }, { shade: 53000 }, { tint: 54000 }, { shade: 88000 }, { tint: 88000 }];
    return ctx.color(scheme(`accent${ctx.scheme - 2}`, shades[i % shades.length]))?.color ?? '4472C4';
  }
  return ctx.color(scheme(`accent${(i % 6) + 1}`, mods))?.color ?? '4472C4';
}

/** Count dense cache slots, including gaps in sparse series, across one DOCX. */
export class ChartPointBudget {
  private remaining = 262144;

  reserve(count: number): void {
    if (count > 32768 || count > this.remaining) throw new Error('Chart point limit exceeded');
    this.remaining -= count;
  }
}

function pointNumber(raw: string | undefined): number {
  const value = Number(raw ?? 0);
  if (!Number.isSafeInteger(value) || value < 0) throw new Error('Invalid chart point count or index');
  return value;
}

interface Cache {
  values: Array<string | null>;
  format?: string;
}

/** Cached values of a c:cat / c:val / c:tx / c:xVal reference or literal. */
function cache(e: XmlElement | undefined, ctx: ChartContext): Cache {
  if (!e) return { values: [] };
  const ref = child(e, 'c:strRef') ?? child(e, 'c:numRef') ?? child(e, 'c:multiLvlStrRef');
  let c = ref ? (child(ref, 'c:strCache') ?? child(ref, 'c:numCache') ?? child(ref, 'c:multiLvlStrCache')) : (child(e, 'c:strLit') ?? child(e, 'c:numLit'));
  // Multi-level categories: the first level holds the innermost labels.
  if (c && child(c, 'c:lvl')) c = { ...child(c, 'c:lvl')!, children: [...children(c, 'c:ptCount'), ...children(child(c, 'c:lvl'), 'c:pt')] };
  if (!c) {
    const v = child(e, 'c:v');
    if (!v) return { values: [] };
    ctx.budget.reserve(1);
    return { values: [v.children.join('')] };
  }
  let count = pointNumber(child(c, 'c:ptCount')?.attrs.val);
  const points = children(c, 'c:pt');
  for (const pt of points) count = Math.max(count, pointNumber(pt.attrs.idx) + 1);
  ctx.budget.reserve(count);
  const values: Array<string | null> = new Array(count).fill(null);
  for (const pt of points) values[pointNumber(pt.attrs.idx)] = child(pt, 'c:v')?.children.join('') ?? '';
  const format = child(c, 'c:formatCode')?.children.join('');
  return { values, ...(format && { format }) };
}

const numbers = (c: Cache) => c.values.map((v) => (v === null || v.trim() === '' || !Number.isFinite(Number(v)) ? null : Number(v)));

function labels(dLbls: XmlElement | undefined, inherited: ChartLabels | undefined, base: ChartTextStyle, ctx: ChartContext): ChartLabels | undefined {
  if (!dLbls) return inherited;
  if (flag(child(dLbls, 'c:delete'), false)) return undefined;
  const style = textStyle(child(dLbls, 'c:txPr'), inherited?.style ?? base, ctx);
  const hidden = children(dLbls, 'c:dLbl')
    .filter((d) => flag(child(d, 'c:delete'), false))
    .map((d) => Number(val(child(d, 'c:idx'))) || 0);
  const show = (n: string, dflt: boolean) => (child(dLbls, n) ? flag(child(dLbls, n), false) : dflt);
  const out: ChartLabels = {
    value: show('c:showVal', inherited?.value ?? false),
    category: show('c:showCatName', inherited?.category ?? false),
    series: show('c:showSerName', inherited?.series ?? false),
    percent: show('c:showPercent', inherited?.percent ?? false),
    style,
  };
  const pos = val(child(dLbls, 'c:dLblPos')) ?? inherited?.position;
  if (pos) out.position = pos;
  const fmt = child(dLbls, 'c:numFmt');
  if (fmt && fmt.attrs.sourceLinked !== '1') out.format = fmt.attrs.formatCode;
  else if (inherited?.format) out.format = inherited.format;
  const sep = child(dLbls, 'c:separator')?.children.join('');
  if (sep !== undefined) out.separator = sep;
  if (hidden.length) out.hidden = hidden;
  return out.value || out.category || out.series || out.percent ? out : undefined;
}

function title(e: XmlElement | undefined, base: ChartTextStyle, ctx: ChartContext, auto?: string): ChartTitle | undefined {
  if (!e) return undefined;
  const style = textStyle(child(e, 'c:txPr'), base, ctx);
  const rich = path(e, 'c:tx', 'c:rich');
  const paragraphs: ChartTitle['paragraphs'] = [];
  if (rich) {
    const richStyle = textStyle(rich, style, ctx);
    if (richStyle.rotation !== undefined) style.rotation = richStyle.rotation;
    for (const p of children(rich, 'a:p')) {
      const pStyle = runStyle(path(p, 'a:pPr', 'a:defRPr'), richStyle, ctx);
      let runs: ChartTitle['paragraphs'][number] = [];
      for (const r of children(p)) {
        if (r.name === 'a:br') {
          paragraphs.push(runs);
          runs = [];
        } else if (r.name === 'a:r' || r.name === 'a:fld') {
          const text = children(r, 'a:t')
            .map((t) => t.children.join(''))
            .join('');
          // A line feed typed into the text breaks the line, as a:br does.
          const style = runStyle(child(r, 'a:rPr'), pStyle, ctx);
          text.split(/\r\n|[\n\r\v]/).forEach((part, k) => {
            if (k) {
              paragraphs.push(runs);
              runs = [];
            }
            if (part) runs.push({ text: part, style });
          });
        }
      }
      paragraphs.push(runs);
    }
  } else {
    const text = cache(path(e, 'c:tx'), ctx).values.join('') || auto;
    if (text) paragraphs.push([{ text, style }]);
  }
  while (paragraphs.length && !paragraphs[paragraphs.length - 1].length) paragraphs.pop();
  if (!paragraphs.some((p) => p.length)) return undefined;
  const ml = path(e, 'c:layout', 'c:manualLayout');
  return {
    paragraphs,
    overlay: flag(child(e, 'c:overlay'), false),
    ...(ml && numVal(child(ml, 'c:x')) !== undefined && { x: numVal(child(ml, 'c:x')), y: numVal(child(ml, 'c:y')) ?? 0 }),
    shape: shape(child(e, 'c:spPr'), ctx, { fill: null, line: null }),
  };
}

function layout(e: XmlElement | undefined): ChartLayout | undefined {
  const ml = path(e, 'c:layout', 'c:manualLayout');
  if (!ml) return undefined;
  const x = numVal(child(ml, 'c:x'));
  const y = numVal(child(ml, 'c:y'));
  const w = numVal(child(ml, 'c:w'));
  const h = numVal(child(ml, 'c:h'));
  // Only edge-mode rectangles (fractions of the chart) are placed; factor mode adjusts the automatic layout.
  if (x === undefined || y === undefined || w === undefined || h === undefined) return undefined;
  if (val(child(ml, 'c:xMode')) === 'factor' || val(child(ml, 'c:yMode')) === 'factor') return undefined;
  return { inner: val(child(ml, 'c:layoutTarget')) === 'inner', x, y, w, h };
}

const PLOT_TYPES: Record<string, ChartPlot['type']> = {
  'c:barChart': 'bar',
  'c:bar3DChart': 'bar',
  'c:lineChart': 'line',
  'c:line3DChart': 'line',
  'c:stockChart': 'line',
  'c:areaChart': 'area',
  'c:area3DChart': 'area',
  'c:pieChart': 'pie',
  'c:pie3DChart': 'pie',
  'c:ofPieChart': 'pie',
  'c:doughnutChart': 'doughnut',
  'c:scatterChart': 'scatter',
  'c:bubbleChart': 'scatter',
  'c:radarChart': 'radar',
};
const MARKERS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus'];

function plot(e: XmlElement, type: ChartPlot['type'], text: ChartTextStyle, ctx: ChartContext, seriesBase: number): ChartPlot {
  const grouping = val(child(e, 'c:grouping'));
  const stacked = grouping === 'stacked' || grouping === 'percentStacked';
  const scatterStyle = val(child(e, 'c:scatterStyle')) ?? 'marker';
  const p: ChartPlot = {
    type,
    horizontal: val(child(e, 'c:barDir')) === 'bar',
    grouping: grouping === 'stacked' || grouping === 'percentStacked' || grouping === 'clustered' ? grouping : type === 'bar' ? 'clustered' : 'standard',
    gapWidth: numVal(child(e, 'c:gapWidth')) ?? 150,
    overlap: numVal(child(e, 'c:overlap')) ?? (stacked && type === 'bar' ? 100 : 0),
    varyColors: flag(child(e, 'c:varyColors'), type === 'pie' || type === 'doughnut'),
    holeSize: numVal(child(e, 'c:holeSize')) ?? 50,
    firstSliceAngle: numVal(child(e, 'c:firstSliceAng')) ?? 0,
    markers: type === 'scatter' ? scatterStyle !== 'line' && scatterStyle !== 'smooth' && scatterStyle !== 'none' : flag(child(e, 'c:marker'), false),
    lines: type !== 'scatter' || scatterStyle === 'line' || scatterStyle === 'lineMarker' || scatterStyle === 'smooth' || scatterStyle === 'smoothMarker',
    series: [],
    axes: children(e, 'c:axId').map((a) => a.attrs.val ?? ''),
  };
  // Excel writes the marker style for plain radar charts too, turning each series' markers off.
  const radarStyle = val(child(e, 'c:radarStyle'));
  if (type === 'radar') p.radarStyle = radarStyle === 'filled' || radarStyle === 'standard' ? radarStyle : 'marker';
  const filled = p.radarStyle === 'filled';
  const traced = type === 'line' || type === 'scatter' || (type === 'radar' && !filled);
  const plotLabels = labels(child(e, 'c:dLbls'), undefined, text, ctx);
  const sers = children(e, 'c:ser').sort((a, b) => (numVal(child(a, 'c:order')) ?? 0) - (numVal(child(b, 'c:order')) ?? 0));
  for (const [k, s] of sers.entries()) {
    const idx = numVal(child(s, 'c:idx')) ?? seriesBase + k;
    const color = autoColor(idx, ctx);
    const isLine = type === 'line' || (type === 'scatter' && p.lines) || (type === 'radar' && !filled);
    // Without their own formatting, lines are 2.25pt in the series colour; bars and areas fill with it.
    const dflt: ChartShape = isLine ? { fill: null, line: { color, width: 2.25 } } : type === 'pie' || type === 'doughnut' ? { fill: { color, alpha: 1 }, line: { color: 'FFFFFF', width: 0.75 } } : { fill: { color, alpha: 1 }, line: null };
    const cat = cache(child(s, 'c:cat') ?? child(s, 'c:xVal'), ctx);
    const vals = cache(child(s, 'c:val') ?? child(s, 'c:yVal'), ctx);
    const ser: ChartSeries = {
      name: cache(child(s, 'c:tx'), ctx).values.join(' ') || `Series ${idx + 1}`,
      categories: cat.values.map((v) => v ?? ''),
      values: numbers(vals),
      format: vals.format ?? 'General',
      shape: shape(child(s, 'c:spPr'), ctx, type === 'scatter' && !p.lines ? { fill: null, line: null } : dflt),
      smooth: flag(child(s, 'c:smooth'), false) || (type === 'scatter' && scatterStyle.startsWith('smooth')),
      points: [],
      explosion: numVal(child(s, 'c:explosion')) ?? 0,
    };
    if (child(child(s, 'c:cat') ?? child(s, 'c:xVal'), 'c:numRef') || child(child(s, 'c:cat') ?? child(s, 'c:xVal'), 'c:numLit')) {
      ser.categoryNumbers = numbers(cat);
      ser.categoryFormat = cat.format ?? 'General';
    }
    if (type === 'scatter') ser.xValues = child(s, 'c:xVal') ? numbers(cat) : ser.values.map((_, i) => i + 1);
    // Markers: the series' own, else the plot's automatic ones (a symbol per series).
    const m = child(s, 'c:marker');
    const auto = (p.markers && (type === 'line' || type === 'scatter')) || p.radarStyle === 'marker';
    const symbol = val(child(m, 'c:symbol')) ?? (auto ? MARKERS[idx % MARKERS.length] : 'none');
    if (symbol !== 'none' && traced) {
      const mc = ser.shape.line?.color ?? ser.shape.fill?.color ?? color;
      const marker: ChartMarker = { symbol, size: numVal(child(m, 'c:size')) ?? (m ? 5 : 7), shape: shape(child(m, 'c:spPr'), ctx, { fill: { color: mc, alpha: 1 }, line: { color: mc, width: 0.75 } }) };
      ser.marker = marker;
    }
    for (const dPt of children(s, 'c:dPt')) {
      const i = numVal(child(dPt, 'c:idx')) ?? 0;
      const base: ChartShape = p.varyColors && (sers.length === 1 || type === 'pie' || type === 'doughnut') ? { ...ser.shape, fill: ser.shape.fill && { color: autoColor(i, ctx), alpha: 1 } } : ser.shape;
      ser.points.push({ idx: i, shape: shape(child(dPt, 'c:spPr'), ctx, base) });
    }
    const l = labels(child(s, 'c:dLbls'), plotLabels, text, ctx);
    if (l) ser.labels = l;
    // Without their own formatting, trendlines are 1.5pt dotted lines in the series colour.
    const trendColor = ser.shape.line?.color ?? ser.shape.fill?.color ?? color;
    const trends = children(s, 'c:trendline').map((t): ChartTrendline => {
      const type = val(child(t, 'c:trendlineType')) ?? 'linear';
      const name = child(t, 'c:name')?.children.join('');
      const intercept = numVal(child(t, 'c:intercept'));
      return {
        type,
        order: numVal(child(t, type === 'movingAvg' ? 'c:period' : 'c:order')) ?? 2,
        forward: numVal(child(t, 'c:forward')) ?? 0,
        backward: numVal(child(t, 'c:backward')) ?? 0,
        ...(intercept !== undefined && { intercept }),
        ...(name && { name }),
        line: line(path(t, 'c:spPr', 'a:ln'), ctx, { color: trendColor, width: 1.5, dash: 'sysDot' }),
      };
    });
    if (trends.length) ser.trendlines = trends;
    // Without their own formatting, error bars are 0.75pt lines in the text colour, capped.
    const bars = children(s, 'c:errBars').map((eb): ChartErrorBars => {
      const valType = val(child(eb, 'c:errValType'));
      const barType = val(child(eb, 'c:errBarType'));
      const custom = (name: string) => (child(eb, name) ? numbers(cache(child(eb, name), ctx)) : undefined);
      const plus = custom('c:plus');
      const minus = custom('c:minus');
      return {
        dir: val(child(eb, 'c:errDir')) === 'x' && type === 'scatter' ? 'x' : 'y',
        type: barType === 'plus' || barType === 'minus' ? barType : 'both',
        valType: valType === 'percentage' || valType === 'stdDev' || valType === 'stdErr' || valType === 'cust' ? valType : 'fixedVal',
        value: numVal(child(eb, 'c:val')) ?? 1,
        ...(plus && { plus }),
        ...(minus && { minus }),
        endCap: !flag(child(eb, 'c:noEndCap'), false),
        line: line(path(eb, 'c:spPr', 'a:ln'), ctx, { color: ctx.color(scheme('tx1'))?.color ?? '000000', width: 0.75 }),
      };
    });
    if (bars.length) ser.errorBars = bars;
    p.series.push(ser);
  }
  // Varied colours give each point of a lone series (or a pie's slices) its own accent colour.
  if (p.varyColors && (p.series.length === 1 || type === 'pie' || type === 'doughnut'))
    for (const ser of p.series) {
      const n = Math.max(ser.values.length, ser.categories.length);
      for (let i = 0; i < n; i++)
        if (!ser.points.some((q) => q.idx === i)) ser.points.push({ idx: i, shape: { ...ser.shape, fill: ser.shape.fill && { color: autoColor(i, ctx), alpha: ser.shape.fill.alpha } } });
    }
  return p;
}

function axis(e: XmlElement, text: ChartTextStyle, ctx: ChartContext, defaultLine: ChartLine): ChartAxis {
  const kind = e.name === 'c:valAx' ? 'val' : e.name === 'c:dateAx' ? 'date' : e.name === 'c:serAx' ? 'ser' : 'cat';
  const scaling = child(e, 'c:scaling');
  const pos = val(child(e, 'c:axPos'));
  const fmt = child(e, 'c:numFmt');
  const crosses = val(child(e, 'c:crosses'));
  const crossesAt = numVal(child(e, 'c:crossesAt'));
  const grid = (name: string) => {
    const g = child(e, name);
    return g ? line(path(g, 'c:spPr', 'a:ln'), ctx, defaultLine) : null;
  };
  const tick = val(child(e, 'c:majorTickMark'));
  const lblPos = val(child(e, 'c:tickLblPos'));
  const ax: ChartAxis = {
    id: val(child(e, 'c:axId')) ?? '',
    kind,
    position: pos === 'l' || pos === 'r' || pos === 't' || pos === 'b' ? pos : 'b',
    deleted: flag(child(e, 'c:delete'), false),
    reverse: val(child(scaling, 'c:orientation')) === 'maxMin',
    format: fmt?.attrs.formatCode ?? 'General',
    sourceLinked: fmt ? fmt.attrs.sourceLinked === '1' : true,
    line: line(path(e, 'c:spPr', 'a:ln'), ctx, defaultLine),
    gridlines: grid('c:majorGridlines'),
    minorGridlines: grid('c:minorGridlines'),
    majorTick: tick === 'in' || tick === 'cross' || tick === 'none' ? tick : 'out',
    labelPosition: lblPos === 'low' || lblPos === 'high' || lblPos === 'none' ? lblPos : 'nextTo',
    style: textStyle(child(e, 'c:txPr'), text, ctx),
    crossAxis: val(child(e, 'c:crossAx')) ?? '',
    crosses: crossesAt ?? (crosses === 'min' || crosses === 'max' ? crosses : 'autoZero'),
    between: val(child(e, 'c:crossBetween')) !== 'midCat',
  };
  const min = numVal(child(scaling, 'c:min'));
  const max = numVal(child(scaling, 'c:max'));
  const unit = numVal(child(e, 'c:majorUnit'));
  const log = numVal(child(scaling, 'c:logBase'));
  const skip = numVal(child(e, 'c:tickLblSkip'));
  if (min !== undefined) ax.min = min;
  if (max !== undefined) ax.max = max;
  if (unit !== undefined && unit > 0) ax.majorUnit = unit;
  if (log !== undefined && log > 1) ax.logBase = log;
  if (skip !== undefined && skip > 0) ax.labelSkip = skip;
  // Axis titles are bold; a vertical axis's title reads up the page unless it says otherwise.
  const t = title(child(e, 'c:title'), { ...text, bold: true }, ctx, 'Axis Title');
  if (t) {
    for (const p of t.paragraphs)
      for (const r of p)
        if (r.style.rotation === undefined && (ax.position === 'l' || ax.position === 'r')) r.style = { ...r.style, rotation: -90 };
    ax.title = t;
  }
  return ax;
}

/**
 * Read a chart part (c:chartSpace) into a chart drawn from its cached values. Elements without
 * their own formatting take the chart style's defaults: style 2 (Office's default) draws text 10pt
 * in the text colour, axes and gridlines in a 75% tint of it, and series in the accent colours.
 */
export function parseChart(space: XmlElement, theme: Theme, color: ColorOf, budget: ChartPointBudget): Chart | undefined {
  const chartEl = child(space, 'c:chart');
  const plotArea = child(chartEl, 'c:plotArea');
  if (!chartEl || !plotArea) return undefined;
  const style = numVal(descendant(space, 'c:style')) ?? 2;
  const ctx: ChartContext = { theme, color, budget, scheme: ((((style - 1) % 8) + 8) % 8) + 1 };
  const ink = color(scheme('tx1'))?.color ?? '000000';
  const text = textStyle(child(space, 'c:txPr'), { font: theme.minorLatin || 'Calibri', size: 10, bold: false, italic: false, color: ink }, ctx);
  delete text.rotation;
  const defaultLine: ChartLine = { color: color(scheme('tx1', { tint: 75000 }))?.color ?? '898989', width: 0.5 };
  const chart: Chart = {
    plots: [],
    axes: [],
    space: shape(child(space, 'c:spPr'), ctx, { fill: { color: color(scheme('bg1'))?.color ?? 'FFFFFF', alpha: 1 }, line: defaultLine }),
    plotShape: shape(child(plotArea, 'c:spPr'), ctx, { fill: null, line: null }),
  };
  let base = 0;
  for (const e of children(plotArea)) {
    const type = PLOT_TYPES[e.name];
    if (type) {
      const p = plot(e, type, text, ctx, base);
      base += p.series.length;
      chart.plots.push(p);
    } else if (e.name === 'c:catAx' || e.name === 'c:valAx' || e.name === 'c:dateAx' || e.name === 'c:serAx') chart.axes.push(axis(e, text, ctx, defaultLine));
  }
  // Surface charts are not drawn: leave the space empty rather than draw bare axes.
  if (!chart.plots.length) return undefined;
  const series = chart.plots.flatMap((p) => p.series);
  // A chart whose title was not deleted shows its only series' name as the title.
  const titleEl = child(chartEl, 'c:title');
  const autoDeleted = flag(child(chartEl, 'c:autoTitleDeleted'), false);
  const titleStyle: ChartTextStyle = { ...text, size: text.size * 1.8, bold: true };
  const t = titleEl
    ? title(titleEl, titleStyle, ctx, series.length === 1 ? series[0].name : 'Chart Title')
    : !autoDeleted && series.length === 1 && chart.plots[0]?.type !== 'scatter'
      ? title(el('c:title'), titleStyle, ctx, series[0].name)
      : undefined;
  if (t) chart.title = t;
  const legend = child(chartEl, 'c:legend');
  if (legend) {
    const pos = val(child(legend, 'c:legendPos'));
    const lay = layout(legend);
    chart.legend = {
      position: pos === 'l' || pos === 't' || pos === 'b' || pos === 'tr' ? pos : 'r',
      overlay: flag(child(legend, 'c:overlay'), false),
      style: textStyle(child(legend, 'c:txPr'), text, ctx),
      shape: shape(child(legend, 'c:spPr'), ctx, { fill: null, line: null }),
      hidden: children(legend, 'c:legendEntry')
        .filter((le) => flag(child(le, 'c:delete'), false))
        .map((le) => numVal(child(le, 'c:idx')) ?? 0),
      ...(lay && { layout: lay }),
    };
  }
  const lay = layout(plotArea);
  if (lay) chart.plotLayout = lay;
  return chart;
}

/** Every piece of text a chart may draw, with its style: for loading fonts before layout. */
export function chartTexts(chart: Chart): Array<{ style: ChartTextStyle; text: string }> {
  const out: Array<{ style: ChartTextStyle; text: string }> = [];
  const digits = '0123456789.,-+%E$ ';
  const title = (t: ChartTitle | undefined) => t?.paragraphs.forEach((p) => p.forEach((r) => out.push({ style: r.style, text: r.text })));
  title(chart.title);
  for (const a of chart.axes) {
    title(a.title);
    out.push({ style: a.style, text: digits + 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz:/' });
  }
  for (const p of chart.plots)
    for (const s of p.series) {
      const cats = s.categories.join(' ');
      for (const a of chart.axes) out.push({ style: a.style, text: cats });
      if (chart.legend) out.push({ style: chart.legend.style, text: s.name + ' ' + cats });
      if (s.labels) out.push({ style: s.labels.style, text: digits + s.name + ' ' + cats });
    }
  return out;
}

