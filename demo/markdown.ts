import DOMPurify from 'dompurify';
import { Marked } from 'marked';
import markedFootnote from 'marked-footnote';

export function renderMarkdown(markdown: string): DocumentFragment {
  const parser = new Marked(markedFootnote({
    prefixId: 'markdown-note-',
    keepLabels: true,
    headingClass: '',
  }));
  const html = parser.parse(markdown, { gfm: true, async: false });
  const preview = DOMPurify.sanitize(html, {
    RETURN_DOM_FRAGMENT: true,
    ALLOW_DATA_ATTR: false,
    ALLOWED_TAGS: [
      'a', 'blockquote', 'br', 'code', 'del', 'em', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr',
      'img', 'li', 'ol', 'p', 'pre', 'rp', 'rt', 'ruby', 'section', 'strong', 'sub', 'sup',
      'table', 'tbody', 'td', 'th', 'thead', 'tr', 'ul',
    ],
    ALLOWED_ATTR: ['align', 'alt', 'colspan', 'href', 'id', 'rowspan', 'start', 'title'],
  });
  // The Markdown contains archive-relative image references, so the preview shows their alt text.
  for (const image of preview.querySelectorAll('img')) image.replaceWith(image.alt);
  return preview;
}
