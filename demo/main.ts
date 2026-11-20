import * as pdfjs from 'pdfjs-dist';
import type { PDFDocumentLoadingTask, PDFDocumentProxy } from 'pdfjs-dist';
import type { ConvertRequest, ConvertResponse } from './convert.worker.js';

interface Sample {
  id: string;
  title: string;
  suite: 'classic' | 'issue';
  featured: boolean;
  docx: string;
  reference: { kind: 'pdf' | 'png'; app: string; files: string[] };
}

type View = 'pages' | 'side' | 'overlay';

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
  fontPack: $<HTMLSelectElement>('#opt-font-pack'),
  platform: $<HTMLSelectElement>('#opt-platform'),
  quality: $<HTMLSelectElement>('#opt-quality'),
  embedded: $<HTMLInputElement>('#opt-embedded'),
  docName: $<HTMLHeadingElement>('#doc-name'),
  docMeta: $<HTMLParagraphElement>('#doc-meta'),
  stats: $<HTMLUListElement>('#stats'),
  views: $<HTMLDivElement>('#views'),
  overlayControl: $<HTMLLabelElement>('#overlay-control'),
  overlay: $<HTMLInputElement>('#overlay'),
  download: $<HTMLAnchorElement>('#download'),
  open: $<HTMLAnchorElement>('#open'),
  warnings: $<HTMLDetailsElement>('#warnings'),
  pages: $<HTMLDivElement>('#pages'),
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

function convert(docx: ArrayBuffer): Promise<ConvertResponse> {
  const id = nextId++;
  const quality = ui.quality.value;
  const request: ConvertRequest = {
    id,
    docx,
    options: {
      fontPack: ui.fontPack.value as 'full' | 'common',
      wordPlatform: ui.platform.value as 'windows' | 'mac',
      imageQuality: quality === 'auto' ? undefined : quality === 'off' ? false : Number(quality),
      embeddedFonts: ui.embedded.checked,
    },
  };
  return new Promise((resolve) => {
    pending.set(id, resolve);
    worker.postMessage(request);
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
let current: Current | undefined;
let view: View = 'pages';
let generation = 0;
let pdfUrl: string | undefined;
let docs: PDFDocumentProxy[] = [];
let tasks: PDFDocumentLoadingTask[] = [];
let observer: IntersectionObserver | undefined;

function setStatus(text: string, kind: 'busy' | 'error' | '' = '') {
  ui.status.textContent = text;
  ui.status.dataset.kind = kind;
  ui.status.hidden = !text;
}

const kb = (n: number) => (n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);

// ---------------------------------------------------------------------------------------------
// Samples

function renderSampleList() {
  const q = ui.search.value.trim().toLowerCase();
  ui.samples.replaceChildren();
  const groups: [string, Sample[]][] = [
    ['Featured', samples.filter((s) => s.featured)],
    ['Real-world documents', samples.filter((s) => !s.featured && s.suite === 'issue')],
    ['Feature tests', samples.filter((s) => !s.featured && s.suite === 'classic')],
  ];
  let shown = 0;
  for (const [label, list] of groups) {
    const items = list.filter((s) => !q || s.title.toLowerCase().includes(q) || s.id.toLowerCase().includes(q));
    if (!items.length) continue;
    const group = el('div', 'sample-group');
    group.append(el('h3', '', `${label} · ${items.length}`));
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
  setStatus(`Downloading ${sample.title}…`, 'busy');
  const res = await fetch(sample.docx);
  if (!res.ok) return setStatus(`Couldn't download the sample (HTTP ${res.status}).`, 'error');
  current = { name: sample.docx.slice(sample.docx.lastIndexOf('/') + 1), bytes: await res.arrayBuffer(), sample };
  renderSampleList();
  if (matchMedia('(max-width: 860px)').matches && samplesReady) document.querySelector('.viewer')?.scrollIntoView({ behavior: 'smooth' });
  await run();
}

async function loadFile(file: File) {
  if (!/\.(docx|docm|dotx|dotm)$/i.test(file.name)) {
    setStatus(`${file.name} isn't a Word document (.docx).`, 'error');
    return;
  }
  current = { name: file.name, bytes: await file.arrayBuffer() };
  if (location.hash) history.replaceState(null, '', location.pathname + location.search);
  renderSampleList();
  await run();
}

// ---------------------------------------------------------------------------------------------
// Conversion and preview

async function run() {
  if (!current) return;
  const gen = ++generation;
  const { name, sample } = current;
  ui.docName.textContent = sample ? sample.title : name;
  ui.docMeta.textContent = sample
    ? `${name} · ${kb(current.bytes.byteLength)} · reference by ${sample.reference.app}`
    : `${name} · ${kb(current.bytes.byteLength)} · your file stays on this device`;
  setStatus('Converting…', 'busy');
  // The worker keeps its own copy; keep ours for re-conversion when options change.
  const result = await convert(current.bytes.slice(0));
  if (gen !== generation) return;
  if (!result.ok) {
    setStatus(`Conversion failed: ${result.error}`, 'error');
    return;
  }
  setStatus('');

  if (pdfUrl) URL.revokeObjectURL(pdfUrl);
  const blob = new Blob([result.pdf], { type: 'application/pdf' });
  pdfUrl = URL.createObjectURL(blob);
  const pdfName = name.replace(/\.[^.]+$/, '') + '.pdf';
  ui.download.href = pdfUrl;
  ui.download.download = pdfName;
  ui.open.href = pdfUrl;
  ui.download.removeAttribute('aria-disabled');
  ui.open.removeAttribute('aria-disabled');

  const refPages = sample ? (sample.reference.kind === 'png' ? sample.reference.files.length : undefined) : undefined;
  renderStats(result, blob.size, refPages);
  renderWarnings(result.warnings);
  ui.views.hidden = !sample;

  const newTasks = [pdfjs.getDocument({ data: new Uint8Array(result.pdf), ...pdfjsOptions() })];
  if (sample?.reference.kind === 'pdf') newTasks.push(pdfjs.getDocument({ url: sample.reference.files[0], ...pdfjsOptions() }));
  let loaded: PDFDocumentProxy[];
  try {
    loaded = await Promise.all(newTasks.map((t) => t.promise));
  } catch (err) {
    for (const t of newTasks) void t.destroy();
    if (gen === generation) setStatus(`Couldn't display the PDF: ${err instanceof Error ? err.message : err}`, 'error');
    return;
  }
  if (gen !== generation) {
    for (const t of newTasks) void t.destroy();
    return;
  }
  const [ours, refDoc] = loaded;
  const reference: PDFDocumentProxy | string[] | undefined = refDoc ?? sample?.reference.files;
  if (refDoc) renderStats(result, blob.size, refDoc.numPages);
  for (const t of tasks) void t.destroy();
  tasks = newTasks;
  docs = loaded;
  await renderPages(ours, reference);
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
  const items: [string, string, string?][] = [
    [r.ms < 1000 ? `${Math.round(r.ms)} ms` : `${(r.ms / 1000).toFixed(2)} s`, 'conversion time', 'Includes downloading any font faces this document needs for the first time.'],
    [
      `${r.pageCount}${refPages != null ? ` / ${refPages}` : ''}`,
      refPages != null ? 'pages (ours / Word)' : r.pageCount === 1 ? 'page' : 'pages',
    ],
    [kb(size), 'PDF size'],
  ];
  if (r.fonts.length) {
    const total = r.fonts.reduce((n, f) => n + f.bytes, 0);
    items.push([`${r.fonts.length} · ${kb(total)}`, 'font faces fetched', r.fonts.map((f) => f.file).join('\n')]);
  } else {
    items.push(['cached', 'fonts']);
  }
  if (r.renderedBreaks) items.push([`${r.matchedBreaks} / ${r.renderedBreaks}`, "Word's page breaks matched", "Page breaks Word recorded in the file (w:lastRenderedPageBreak) that start a page here too."]);
  ui.stats.replaceChildren(
    ...items.map(([value, label, title]) => {
      const li = el('li');
      if (title) li.title = title;
      li.append(el('strong', '', value), el('span', '', label));
      if (refPages != null && label.startsWith('pages')) li.dataset.match = String(r.pageCount === refPages);
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

async function renderPages(ours: PDFDocumentProxy, reference: PDFDocumentProxy | string[] | undefined) {
  observer?.disconnect();
  const jobs = new Map<Element, PageJob>();
  observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const job = jobs.get(entry.target);
        if (!job) continue;
        jobs.delete(entry.target);
        observer?.unobserve(entry.target);
        void job();
      }
    },
    { rootMargin: '800px 0px' },
  );

  const refCount = !reference ? 0 : Array.isArray(reference) ? reference.length : reference.numPages;
  const showRef = view !== 'pages' && refCount > 0;
  const count = showRef ? Math.max(ours.numPages, refCount) : ours.numPages;
  ui.pages.dataset.view = showRef ? view : 'pages';
  ui.pages.replaceChildren();

  if (showRef && view === 'side') {
    const head = el('div', 'page-row heads');
    head.append(el('div', 'col-head', 'Kanera DOCX (this library)'), el('div', 'col-head', current?.sample?.reference.app ?? 'Word'));
    ui.pages.append(head);
  }

  for (let i = 0; i < count; i++) {
    const row = el('div', 'page-row');
    const oursSlot = pageSlot(i < ours.numPages ? `Page ${i + 1}` : undefined);
    const refSlot = showRef ? pageSlot(i < refCount ? `Word page ${i + 1}` : undefined) : undefined;
    const sized = i < ours.numPages ? ours.getPage(i + 1) : undefined;
    if (view === 'overlay' && showRef && refSlot) {
      const stack = el('div', 'overlay-stack');
      stack.append(oursSlot, refSlot);
      refSlot.classList.add('top');
      row.append(stack);
    } else {
      row.append(oursSlot);
      if (refSlot) row.append(refSlot);
    }
    ui.pages.append(row);

    // Size each slot before rendering so the scroll height is right from the start.
    void sized?.then((p) => {
      const vp = p.getViewport({ scale: 1 });
      oursSlot.style.aspectRatio = `${vp.width} / ${vp.height}`;
      refSlot?.style.setProperty('aspect-ratio', `${vp.width} / ${vp.height}`);
    });
    if (i < ours.numPages) jobs.set(row, async () => {
      await drawPdfPage(ours, i + 1, oursSlot);
      if (refSlot && i < refCount) await drawReference(reference!, i, refSlot);
    });
    else if (refSlot && i < refCount) jobs.set(row, () => drawReference(reference!, i, refSlot));
    observer.observe(row);
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
  img.alt = `Word page ${i + 1}`;
  img.decoding = 'async';
  img.src = reference[i];
  await img.decode().catch(() => undefined);
  if (img.naturalWidth) slot.style.aspectRatio = `${img.naturalWidth} / ${img.naturalHeight}`;
  slot.replaceChildren(img);
}

function setView(next: View, rerender = true) {
  view = next;
  for (const b of ui.views.querySelectorAll<HTMLButtonElement>('button')) b.setAttribute('aria-pressed', String(b.dataset.view === next));
  ui.overlayControl.hidden = next !== 'overlay';
  try {
    localStorage.setItem('docx2pdf-demo-view', next);
  } catch {}
  if (rerender && docs[0]) {
    const ref = current?.sample?.reference.kind === 'png' ? current.sample.reference.files : docs[1];
    void renderPages(docs[0], ref);
  }
}

// ---------------------------------------------------------------------------------------------
// Wiring

ui.file.addEventListener('change', () => {
  const f = ui.file.files?.[0];
  if (f) void loadFile(f);
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
  if (f) void loadFile(f);
});

ui.search.addEventListener('input', renderSampleList);
for (const input of [ui.fontPack, ui.platform, ui.quality, ui.embedded]) input.addEventListener('change', () => void run());
ui.views.addEventListener('click', (e) => {
  const b = (e.target as HTMLElement).closest('button');
  if (b?.dataset.view) setView(b.dataset.view as View);
});
ui.overlay.addEventListener('input', () => ui.pages.style.setProperty('--overlay', String(Number(ui.overlay.value) / 100)));
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
  if (id && id !== current?.sample?.id) void loadSample(id);
}
window.addEventListener('hashchange', routeFromHash);

async function init() {
  try {
    const saved = localStorage.getItem('docx2pdf-demo-view') as View | null;
    if (saved === 'pages' || saved === 'side' || saved === 'overlay') view = saved;
  } catch {}
  setView(view, false);
  ui.pages.style.setProperty('--overlay', String(Number(ui.overlay.value) / 100));
  const res = await fetch('samples/samples.json');
  samples = await res.json();
  renderSampleList();
  if (location.hash.startsWith('#sample=')) routeFromHash();
  else if (samples[0]) await loadSample(samples[0].id);
  samplesReady = true;
}

void init();
