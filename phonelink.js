// Packs data into the part of a link after "#", for opening on a phone.
// Browsers never send that part to the website, so the data goes from the
// PC to the phone (by QR code or a copied link) without being uploaded.
// JSON, deflate-compressed when the browser can, then base64url.
(function (root) {
  'use strict';

  const toB64 = bytes => {
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  };
  const fromB64 = text => {
    const bin = atob(text.replace(/-/g, '+').replace(/_/g, '/'));
    return Uint8Array.from(bin, ch => ch.charCodeAt(0));
  };
  const pipe = async (bytes, stream) =>
    new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer());

  // 'z' + compressed, or 'j' + plain JSON where CompressionStream is missing.
  async function encode(data) {
    const json = new TextEncoder().encode(JSON.stringify(data));
    if (typeof CompressionStream === 'function') return 'z' + toB64(await pipe(json, new CompressionStream('deflate-raw')));
    return 'j' + toB64(json);
  }

  async function decode(text) {
    const kind = text[0];
    let bytes = fromB64(text.slice(1));
    if (kind === 'z') {
      if (typeof DecompressionStream !== 'function') throw new Error('這支手機的瀏覽器太舊，請更新後再開');
      bytes = await pipe(bytes, new DecompressionStream('deflate-raw'));
    } else if (kind !== 'j') {
      throw new Error('連結內容無法辨識');
    }
    return JSON.parse(new TextDecoder().decode(bytes));
  }

  const api = { encode, decode };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ADRPhoneLink = api;
})(typeof window !== 'undefined' ? window : globalThis);
