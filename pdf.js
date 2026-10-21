// Minimal PDF writer: one full-page JPEG per page. Text is drawn on a canvas
// first, so no fonts need embedding and Chinese prints exactly as previewed.
(function (root) {
  'use strict';

  // A4 landscape in PDF points (1/72 inch).
  const A4_LANDSCAPE = { width: 841.89, height: 595.28 };

  // pages: [{ jpeg: Uint8Array, width, height }] (pixel size of each JPEG).
  function build(pages, size = A4_LANDSCAPE) {
    const enc = new TextEncoder();
    const chunks = [];
    const offsets = [];
    let length = 0;
    const push = data => {
      const bytes = typeof data === 'string' ? enc.encode(data) : data;
      chunks.push(bytes);
      length += bytes.length;
    };
    const object = (n, body, stream) => {
      offsets[n] = length;
      push(`${n} 0 obj\n${body}\n`);
      if (stream) { push('stream\n'); push(stream); push('\nendstream\n'); }
      push('endobj\n');
    };

    // Objects: 1 catalog, 2 page tree, then 3 per page (page, content, image).
    const pageIds = pages.map((_, i) => 3 + i * 3);
    push('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
    object(1, '<< /Type /Catalog /Pages 2 0 R >>');
    object(2, `<< /Type /Pages /Kids [${pageIds.map(id => `${id} 0 R`).join(' ')}] /Count ${pages.length} >>`);
    pages.forEach((page, i) => {
      const [pageId, contentId, imageId] = [pageIds[i], pageIds[i] + 1, pageIds[i] + 2];
      const content = enc.encode(`q ${size.width} 0 0 ${size.height} 0 0 cm /Im0 Do Q`);
      object(pageId, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${size.width} ${size.height}] ` +
        `/Resources << /XObject << /Im0 ${imageId} 0 R >> >> /Contents ${contentId} 0 R >>`);
      object(contentId, `<< /Length ${content.length} >>`, content);
      object(imageId, `<< /Type /XObject /Subtype /Image /Width ${page.width} /Height ${page.height} ` +
        `/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${page.jpeg.length} >>`, page.jpeg);
    });

    const count = 3 + pages.length * 3;
    const xref = length;
    push(`xref\n0 ${count}\n0000000000 65535 f \n`);
    for (let n = 1; n < count; n++) push(String(offsets[n]).padStart(10, '0') + ' 00000 n \n');
    push(`trailer\n<< /Size ${count} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);

    const out = new Uint8Array(length);
    let pos = 0;
    chunks.forEach(c => { out.set(c, pos); pos += c.length; });
    return out;
  }

  const api = { build, A4_LANDSCAPE };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRPdf = api;
})(typeof window !== 'undefined' ? window : globalThis);
