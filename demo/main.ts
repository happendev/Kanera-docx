import * as pdfjs from 'pdfjs-dist';
import type { PDFDocumentLoadingTask, PDFDocumentProxy } from 'pdfjs-dist';
import type { ConvertRequest, ConvertResponse, OutputFormat } from './convert.worker.js';
import { renderMarkdown } from './markdown.js';
import { parseSamples, type Sample } from './samples.js';

type View = 'pages' | 'side';

const outputs = {
  pdf: { label: 'PDF', extension: 'pdf', mimeType: 'application/pdf' },
  text: { label: 'Plain text', extension: 'txt', mimeType: 'text/plain;charset=utf-8' },
  markdown: { label: 'Markdown', extension: 'md', mimeType: 'text/markdown;charset=utf-8' },
};

pdfjs.GlobalWorkerOptions.workerSrc = new URL('./pdf.worker.min.mjs', import.meta.url).href;
const pdfjsAssets = new URL('./pdfjs/', import.meta.url).href;

const $ = <T extends HTMLElement>(sel: string) => document.querySelector(sel) as T;
const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string) => {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text != null) node.textContent = text;
  return node;
};

const ui = {
  drop: $<HTMLLabelElement>('#drop'),
  file: $<HTMLInputElement>('#file'),
  search: $<HTMLInputElement>('#search'),
  samples: $<HTMLDivElement>('#samples'),
  pdfOptions: $<HTMLDetailsElement>('#pdf-options'),
  fontPack: $<HTMLSelectElement>('#opt-font-pack'),
  platform: $<HTMLSelectElement>('#opt-platform'),
  quality: $<HTMLSelectElement>('#opt-quality'),
  embedded: $<HTMLInputElement>('#opt-embedded'),
  docName: $<HTMLHeadingElement>('#doc-name'),
  docMeta: $<HTMLParagraphElement>('#doc-meta'),
  stats: $<HTMLUListElement>('#stats'),
  formats: $<HTMLDivElement>('#formats'),
  views: $<HTMLDivElement>('#views'),
  sideBySide: $<HTMLButtonElement>('#views [data-view="side"]'),
  download: $<HTMLAnchorElement>('#download'),
  open: $<HTMLAnchorElement>('#open'),
  warnings: $<HTMLDetailsElement>('#warnings'),
  pages: $<HTMLDivElement>('#pages'),
  textPreview: $<HTMLPreElement>('#text-preview'),
  markdownPreview: $<HTMLElement>('#markdown-preview'),
  status: $<HTMLDivElement>('#status'),
};

// ---------------------------------------------------------------------------------------------
// Conversion worker

const worker = new Worker(new URL('./convert.worker.js', import.meta.url), { type: 'module' });
let nextId = 1;
const pending = new Map<number, (r: ConvertResponse) => void>();
worker.onmessage = (e: MessageEvent<ConvertResponse>) => {
  pending.get(e.data.id)?.(e.data);
  pending.delete(e.data.id);
};

function convert(docx: ArrayBuffer, format: OutputFormat): Promise<ConvertResponse> {
  const id = nextId++;
  const quality = ui.quality.value;
  const request: ConvertRequest = format === 'pdf' ? {
    id,
    docx,
    format,
    options: {
      fontPack: ui.fontPack.value as 'full' | 'common',
      wordPlatform: ui.platform.value as 'windows' | 'mac',
      imageQuality: quality === 'auto' ? undefined : quality === 'off' ? false : Number(quality),
      useEmbeddedFonts: ui.embedded.checked,
    },
  } : { id, docx, format };
  return new Promise((resolve) => {
    pending.set(id, resolve);
    worker.postMessage(request, [docx]);
  });
}

// ---------------------------------------------------------------------------------------------
// State

interface Current {
  name: string;
  bytes: ArrayBuffer;
  sample?: Sample;
}

let samples: Sample[] = [];
let samplesReady = false;
const expandedSampleGroups = new Set<string>();
let current: Current | undefined;
let format: OutputFormat = 'pdf';
let view: View = 'pages';
let loadSequence = 0;
let generation = 0;
let outputUrl: string | undefined;
let docs: PDFDocumentProxy[] = [];
let tasks: PDFDocumentLoadingTask[] = [];
let observer: IntersectionObserver | undefined;

function setStatus(text: string, kind: 'busy' | 'error' | '' = '') {
  ui.status.textContent = text;
  ui.status.dataset.kind = kind;
  ui.status.hidden = !text;
}

function reportDemoError(err: unknown) {
  setStatus(`Unexpected error: ${err instanceof Error ? err.message : err}`, 'error');
}

const kb = (n: number) => (n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);

// ---------------------------------------------------------------------------------------------
// Samples

function renderSampleList() {
  const q = ui.search.value.trim().toLowerCase();
  ui.samples.replaceChildren();
  const groups: [string, Sample[]][] = [
    ['Featured', samples.filter((s) => s.featured)],
    ['Issue report samples', samples.filter((s) => !s.featured && s.suite === 'issue')],
    ['Feature tests', samples.filter((s) => !s.featured && s.suite === 'classic')],
  ];
  let shown = 0;
  for (const [label, list] of groups) {
    const items = list.filter((s) => !q || s.title.toLowerCase().includes(q) || s.id.toLowerCase().includes(q));
    if (!items.length) continue;
    const group = el('details', 'sample-group');
    group.open = q !== '' || expandedSampleGroups.has(label);
    if (!q) {
      group.addEventListener('toggle', () => {
        if (group.open) expandedSampleGroups.add(label);
        else expandedSampleGroups.delete(label);
      });
    }
    group.append(el('summary', '', `${label} · ${items.length}`));
    const ul = el('ul');
    for (const s of items) {
      const li = el('li');
      const a = el('a', '', s.title);
      a.href = `#sample=${encodeURIComponent(s.id)}`;
      if (current?.sample?.id === s.id) a.setAttribute('aria-current', 'true');
      li.append(a);
      ul.append(li);
    }
    group.append(ul);
    ui.samples.append(group);
    shown += items.length;
  }
  if (!shown) ui.samples.append(el('p', 'muted', 'No samples match.'));
}

async function loadSample(id: string) {
  const sample = samples.find((s) => s.id === id);
  if (!sample) return;
  const load = beginLoad(sample.title, `Downloading ${sample.title}…`);
  try {
    const res = await fetch(sample.docx);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const bytes = await res.arrayBuffer();
    if (load !== loadSequence) return;
    current = { name: sample.docx.slice(sample.docx.lastIndexOf('/') + 1), bytes, sample };
  } catch (err) {
    if (load === loadSequence) setStatus(`Couldn't download the sample (${err}).`, 'error');
    return;
  }
  renderSampleList();
  if (matchMedia('(max-width: 900px)').matches && samplesReady) document.querySelector('.viewer')?.scrollIntoView({ behavior: 'smooth' });
  await run();
}

async function loadFile(file: File) {
  if (!/\.(docx|docm|dotx|dotm)$/i.test(file.name)) {
    setStatus(`${file.name} isn't a Word document (.docx).`, 'error');
    return;
  }
  if (location.hash) history.replaceState(null, '', location.pathname + location.search);
  const load = beginLoad(file.name, `Reading ${file.name}…`);
  try {
    const bytes = await file.arrayBuffer();
    if (load !== loadSequence) return;
    current = { name: file.name, bytes };
  } catch (err) {
    if (load === loadSequence) setStatus(`Couldn't read the file (${err}).`, 'error');
    return;
  }
  renderSampleList();
  await run();
}

function beginLoad(title: string, status: string) {
  ++generation;
  current = undefined;
  clearOutput();
  ui.docName.textContent = title;
  ui.docMeta.textContent = '';
  setStatus(status, 'busy');
  return ++loadSequence;
}

// ---------------------------------------------------------------------------------------------
// Conversion and preview

function destroyTasks(loadingTasks: PDFDocumentLoadingTask[]) {
  for (const task of loadingTasks) {
    void task.destroy().catch((err) => console.warn('Could not clean up the PDF preview:', err));
  }
}

function clearOutput() {
  if (outputUrl) URL.revokeObjectURL(outputUrl);
  outputUrl = undefined;
  for (const link of [ui.download, ui.open]) {
    link.removeAttribute('href');
    link.setAttribute('aria-disabled', 'true');
  }
  ui.download.removeAttribute('download');
  ui.stats.replaceChildren();
  ui.views.hidden = true;
  renderWarnings([]);
  ui.textPreview.hidden = true;
  ui.textPreview.textContent = '';
  ui.markdownPreview.hidden = true;
  ui.markdownPreview.replaceChildren();
  observer?.disconnect();
  observer = undefined;
  ui.pages.replaceChildren();
  destroyTasks(tasks);
  tasks = [];
  docs = [];
}

async function run() {
  if (!current) return;
  const gen = ++generation;
  const { name, bytes, sample } = current;
  const output = outputs[format];
  clearOutput();
  ui.docName.textContent = sample ? sample.title : name;
  ui.docMeta.textContent = sample
    ? `${name} · ${kb(bytes.byteLength)} · reference by ${sample.reference.app}`
    : `${name} · ${kb(bytes.byteLength)} · your file stays on this device`;
  setStatus(`Converting to ${output.label}…`, 'busy');
  // Transfer a copy so the same document can be converted again when the format or options change.
  const result = await convert(bytes.slice(0), format);
  if (gen !== generation) return;
  if (!result.ok) {
    setStatus(`Conversion failed: ${result.error}`, 'error');
    return;
  }
  setStatus('');

  const blob = new Blob([result.format === 'pdf' ? result.pdf : result.text], { type: output.mimeType });
  outputUrl = URL.createObjectURL(blob);
  ui.download.href = outputUrl;
  ui.download.download = `${name.replace(/\.[^.]+$/, '')}.${output.extension}`;
  ui.download.removeAttribute('aria-disabled');
  renderWarnings(result.warnings);

  if (result.format !== 'pdf') {
    renderStats(result, blob.size);
    if (!result.text) {
      setStatus('No text found in this document.');
    } else if (result.format === 'markdown') {
      ui.markdownPreview.replaceChildren(renderMarkdown(result.text));
      ui.markdownPreview.hidden = false;
    } else {
      ui.textPreview.setAttribute('aria-label', 'Plain text output');
      ui.textPreview.textContent = result.text;
      ui.textPreview.hidden = false;
    }
    return;
  }

  ui.open.href = outputUrl;
  ui.open.removeAttribute('aria-disabled');

  const refPages = sample ? (sample.reference.kind === 'png' ? sample.reference.files.length : undefined) : undefined;
  renderStats(result, blob.size, refPages);

  const newTasks = [pdfjs.getDocument({ data: new Uint8Array(result.pdf), ...pdfjsOptions() })];
  if (sample?.reference.kind === 'pdf') newTasks.push(pdfjs.getDocument({ url: sample.reference.files[0], ...pdfjsOptions() }));
  tasks = newTasks;
  let loaded: PDFDocumentProxy[];
  try {
    loaded = await Promise.all(newTasks.map((t) => t.promise));
  } catch (err) {
    if (gen !== generation) return;
    destroyTasks(newTasks);
    tasks = [];
    setStatus(`Couldn't display the PDF: ${err instanceof Error ? err.message : err}`, 'error');
    return;
  }
  if (gen !== generation) return;
  const [ours, refDoc] = loaded;
  const reference: PDFDocumentProxy | string[] | undefined = refDoc ?? sample?.reference.files;
  if (refDoc) renderStats(result, blob.size, refDoc.numPages);
  docs = loaded;
  ui.views.hidden = !sample;
  ui.sideBySide.textContent = `Side by side with ${referenceName()}`;
  renderPages(ours, reference);
}

/** The application that made the current sample's reference rendering. */
function referenceName(): string {
  const app = current?.sample?.reference.app ?? '';
  return /\bWord\b/.test(app) || !app ? 'Word' : app.replace(/\s.*$/, '');
}

function pdfjsOptions() {
  return {
    cMapUrl: `${pdfjsAssets}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${pdfjsAssets}standard_fonts/`,
    wasmUrl: `${pdfjsAssets}wasm/`,
  };
}

function renderStats(r: Extract<ConvertResponse, { ok: true }>, size: number, refPages?: number) {
  const isPdf = r.format === 'pdf';
  const time = r.ms < 1000 ? `${Math.round(r.ms)} ms` : `${(r.ms / 1000).toFixed(2)} s`;
  const items: [string, string, string?][] = [
    [time, 'conversion time', isPdf ? 'Measured after the required font files are downloaded.' : undefined],
  ];
  if (isPdf) {
    items.push([
      `${r.pageCount}${refPages != null ? ` / ${refPages}` : ''}`,
      refPages != null ? `pages (ours / ${referenceName()})` : r.pageCount === 1 ? 'page' : 'pages',
    ], [kb(size), 'PDF size']);
    if (r.fonts.length) {
      const total = r.fonts.reduce((n, f) => n + f.bytes, 0);
      items.push([`${r.fonts.length} · ${kb(total)}`, 'font faces fetched', r.fonts.map((f) => f.file).join('\n')]);
    } else {
      items.push(['cached', 'fonts']);
    }
    if (r.renderedBreaks) items.push([`${r.matchedBreaks} / ${r.renderedBreaks}`, "Word's page breaks matched", "Page breaks Word recorded in the file (w:lastRenderedPageBreak) that start a page here too."]);
  } else {
    items.push([size < 1024 ? `${size} B` : kb(size), `${outputs[r.format].label} size`]);
  }
  ui.stats.replaceChildren(
    ...items.map(([value, label, title]) => {
      const li = el('li');
      if (title) li.title = title;
      li.append(el('strong', '', value), el('span', '', label));
      if (isPdf && refPages != null && label.startsWith('pages')) li.dataset.match = String(r.pageCount === refPages);
      return li;
    }),
  );
}

function renderWarnings(warnings: string[]) {
  ui.warnings.hidden = !warnings.length;
  ui.warnings.replaceChildren();
  if (!warnings.length) return;
  const summary = el('summary', '', `${warnings.length} warning${warnings.length === 1 ? '' : 's'}`);
  const ul = el('ul');
  for (const w of warnings) ul.append(el('li', '', w));
  ui.warnings.append(summary, ul);
}

type PageJob = () => Promise<void>;

function renderPages(ours: PDFDocumentProxy, reference: PDFDocumentProxy | string[] | undefined) {
  observer?.disconnect();
  const jobs = new Map<Element, PageJob>();
  const pageObserver = new IntersectionObserver(
    (entries) => {
      if (observer !== pageObserver) return;
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const job = jobs.get(entry.target);
        if (!job) continue;
        jobs.delete(entry.target);
        pageObserver.unobserve(entry.target);
        void job().catch((err) => {
          if (observer === pageObserver) setStatus(`Couldn't display a PDF page: ${err instanceof Error ? err.message : err}`, 'error');
        });
      }
    },
    { rootMargin: '800px 0px' },
  );
  observer = pageObserver;

  const refCount = !reference ? 0 : Array.isArray(reference) ? reference.length : reference.numPages;
  const showRef = view === 'side' && refCount > 0;
  const count = showRef ? Math.max(ours.numPages, refCount) : ours.numPages;
  ui.pages.dataset.view = showRef ? view : 'pages';
  ui.pages.replaceChildren();

  if (showRef) {
    const head = el('div', 'page-row heads');
    head.append(el('div', 'col-head', 'Kanera DOCX (this library)'), el('div', 'col-head', current?.sample?.reference.app ?? referenceName()));
    ui.pages.append(head);
  }

  for (let i = 0; i < count; i++) {
    const row = el('div', 'page-row');
    const oursSlot = pageSlot(i < ours.numPages ? `Page ${i + 1}` : undefined);
    const refSlot = showRef ? pageSlot(i < refCount ? `${referenceName()} page ${i + 1}` : undefined) : undefined;
    const sized = i < ours.numPages ? ours.getPage(i + 1) : undefined;
    row.append(oursSlot);
    if (refSlot) row.append(refSlot);
    ui.pages.append(row);

    // Size each slot before rendering so the scroll height is right from the start.
    void sized?.then((p) => {
      const vp = p.getViewport({ scale: 1 });
      oursSlot.style.aspectRatio = `${vp.width} / ${vp.height}`;
      refSlot?.style.setProperty('aspect-ratio', `${vp.width} / ${vp.height}`);
    }).catch(() => undefined);
    if (i < ours.numPages) jobs.set(row, async () => {
      await drawPdfPage(ours, i + 1, oursSlot);
      if (refSlot && i < refCount) await drawReference(reference!, i, refSlot);
    });
    else if (refSlot && i < refCount) jobs.set(row, () => drawReference(reference!, i, refSlot));
    pageObserver.observe(row);
  }
}

function pageSlot(label: string | undefined): HTMLDivElement {
  const slot = el('div', label ? 'page' : 'page missing');
  slot.setAttribute('aria-label', label ?? 'No page');
  if (!label) slot.append(el('span', '', 'No page'));
  return slot;
}

async function drawPdfPage(doc: PDFDocumentProxy, n: number, slot: HTMLElement) {
  const page = await doc.getPage(n);
  const base = page.getViewport({ scale: 1 });
  slot.style.aspectRatio = `${base.width} / ${base.height}`;
  const width = Math.max(slot.clientWidth, 200);
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const viewport = page.getViewport({ scale: (width / base.width) * dpr });
  const canvas = el('canvas');
  canvas.width = Math.round(viewport.width);
  canvas.height = Math.round(viewport.height);
  await page.render({ canvas, viewport }).promise;
  slot.replaceChildren(canvas);
}

async function drawReference(reference: PDFDocumentProxy | string[], i: number, slot: HTMLElement) {
  if (!Array.isArray(reference)) return drawPdfPage(reference, i + 1, slot);
  const img = el('img');
  img.alt = `${referenceName()} page ${i + 1}`;
  img.decoding = 'async';
  img.src = reference[i];
  await img.decode().catch(() => undefined);
  if (img.naturalWidth) slot.style.aspectRatio = `${img.naturalWidth} / ${img.naturalHeight}`;
  slot.replaceChildren(img);
}

function setView(next: View, rerender = true) {
  view = next;
  for (const b of ui.views.querySelectorAll<HTMLButtonElement>('button')) b.setAttribute('aria-pressed', String(b.dataset.view === next));
  try {
    localStorage.setItem('docx2pdf-demo-view', next);
  } catch {}
  if (rerender && docs[0]) {
    const ref = current?.sample?.reference.kind === 'png' ? current.sample.reference.files : docs[1];
    renderPages(docs[0], ref);
  }
}

// ---------------------------------------------------------------------------------------------
// Wiring

ui.file.addEventListener('change', () => {
  const f = ui.file.files?.[0];
  if (f) void loadFile(f).catch(reportDemoError);
  ui.file.value = '';
});
for (const type of ['dragenter', 'dragover'] as const) {
  window.addEventListener(type, (e) => {
    e.preventDefault();
    ui.drop.classList.add('dragging');
  });
}
for (const type of ['dragleave', 'drop'] as const) {
  window.addEventListener(type, (e) => {
    e.preventDefault();
    if (type === 'dragleave' && (e as DragEvent).relatedTarget) return;
    ui.drop.classList.remove('dragging');
  });
}
window.addEventListener('drop', (e) => {
  const f = e.dataTransfer?.files[0];
  if (f) void loadFile(f).catch(reportDemoError);
});

ui.search.addEventListener('input', renderSampleList);
for (const input of [ui.fontPack, ui.platform, ui.quality, ui.embedded]) {
  input.addEventListener('change', () => {
    if (format === 'pdf') void run().catch(reportDemoError);
  });
}
ui.formats.addEventListener('click', (e) => {
  const next = (e.target as HTMLElement).closest('button')?.dataset.format as OutputFormat | undefined;
  if (!next || next === format) return;
  format = next;
  for (const b of ui.formats.querySelectorAll('button')) b.setAttribute('aria-pressed', String(b.dataset.format === next));
  ui.download.textContent = `Download ${outputs[next].label}`;
  ui.open.hidden = ui.pdfOptions.hidden = next !== 'pdf';
  void run().catch(reportDemoError);
});
ui.views.addEventListener('click', (e) => {
  const b = (e.target as HTMLElement).closest('button');
  if (b?.dataset.view) setView(b.dataset.view as View);
});
for (const a of [ui.download, ui.open]) a.addEventListener('click', (e) => a.hasAttribute('aria-disabled') && e.preventDefault());

let resizeTimer = 0;
let lastWidth = ui.pages.clientWidth;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = window.setTimeout(() => {
    if (Math.abs(ui.pages.clientWidth - lastWidth) < 40 || !docs[0]) return;
    lastWidth = ui.pages.clientWidth;
    setView(view);
  }, 250);
});

function routeFromHash() {
  const m = /^#sample=(.+)$/.exec(location.hash);
  const id = m ? decodeURIComponent(m[1]) : undefined;
  if (id && id !== current?.sample?.id) void loadSample(id).catch(reportDemoError);
}
window.addEventListener('hashchange', routeFromHash);

async function init() {
  try {
    const saved = localStorage.getItem('docx2pdf-demo-view');
    if (saved === 'pages' || saved === 'side') view = saved;
  } catch {}
  setView(view, false);
  try {
    const res = await fetch('samples/samples.json');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    samples = parseSamples(await res.json());
  } catch {
    ui.samples.replaceChildren(el('p', 'muted', 'Could not load the sample list.'));
    return;
  }
  renderSampleList();
  if (!loadSequence) {
    if (location.hash.startsWith('#sample=')) routeFromHash();
    else if (samples[0]) await loadSample(samples[0].id);
  }
  samplesReady = true;
}

void init().catch(reportDemoError);
