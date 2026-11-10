// Dropdown lists drawn by the page instead of the browser. On some Windows
// PCs Chrome crashes the page (STATUS_BREAKPOINT) when its own dropdown list
// opens, especially in the --app window; this list is plain HTML.
//
// - Every <select> on the page opens this list instead of Chrome's (mouse
//   and keyboard); picking sets its value and fires `change`, so page code
//   works unchanged.
// - ADRDropdown.pick(anchor, items, current) opens it next to any element,
//   e.g. a plain button, and resolves to the picked value (null if closed).
//   items: [{ value, text, ask }]; an item with `ask` (a placeholder) asks
//   for a new value inside the list, without a prompt() dialog.
(function (root) {
  'use strict';

  let menu = null;   // { el, anchor, resolve }

  function close(value = null) {
    if (!menu) return;
    const { el, resolve } = menu;
    menu = null;
    el.remove();
    document.removeEventListener('mousedown', outside, true);
    window.removeEventListener('resize', onResize);
    window.removeEventListener('scroll', reposition, true);
    resolve(value);
  }
  const onResize = () => close();

  function outside(e) {
    if (menu && !menu.el.contains(e.target) && !menu.anchor.contains(e.target)) close();
  }

  function reposition() {
    if (!menu) return;
    const { el, anchor } = menu;
    if (!anchor.isConnected) return close();
    const r = anchor.getBoundingClientRect();
    if (r.bottom < 0 || r.top > window.innerHeight) return close();
    el.style.minWidth = r.width + 'px';
    el.style.left = Math.max(4, Math.min(r.left, window.innerWidth - el.offsetWidth - 4)) + 'px';
    const below = window.innerHeight - r.bottom;
    el.style.top = (below < el.offsetHeight + 8 && r.top > below
      ? Math.max(4, r.top - el.offsetHeight - 2) : r.bottom + 2) + 'px';
  }

  function ask(item) {
    const el = menu.el;
    el.replaceChildren();
    const input = document.createElement('input');
    input.className = 'dd-input';
    input.placeholder = item.ask;
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.className = 'dd-ok';
    ok.textContent = '確定';
    const done = () => close(input.value.trim() || null);
    // Enter that picks a character in a Chinese input method is not "done".
    input.addEventListener('keydown', e => {
      if (e.isComposing || e.keyCode === 229) return;
      if (e.key === 'Enter') { e.preventDefault(); done(); }
      if (e.key === 'Escape') close();
    });
    ok.addEventListener('click', done);
    const row = document.createElement('div');
    row.className = 'dd-ask';
    row.append(input, ok);
    el.append(row);
    reposition();
    input.focus();
  }

  function pick(anchor, items, current) {
    const wasOpenHere = menu && menu.anchor === anchor;
    close();
    if (wasOpenHere) return Promise.resolve(null);  // a second click closes it
    return new Promise(resolve => {
      const el = document.createElement('div');
      el.className = 'dd-menu';
      el.setAttribute('role', 'listbox');
      menu = { el, anchor, resolve };
      items.forEach(item => {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'dd-item' + (item.value === current && !item.ask ? ' current' : '');
        b.textContent = item.text;
        b.disabled = !!item.disabled;
        b.addEventListener('click', () => (item.ask ? ask(item) : close(item.value)));
        el.appendChild(b);
      });
      document.body.appendChild(el);
      reposition();
      const cur = el.querySelector('.current');
      if (cur) cur.scrollIntoView({ block: 'nearest' });
      document.addEventListener('mousedown', outside, true);
      window.addEventListener('resize', onResize);
      window.addEventListener('scroll', reposition, true);
    });
  }

  async function openSelect(select) {
    if (select.disabled) return;
    const items = [...select.options].map(o => ({ value: o.value, text: o.textContent, disabled: o.disabled }));
    const value = await pick(select, items, select.value);
    if (value == null) return;
    select.focus({ preventScroll: true });
    if (select.value !== value) {
      select.value = value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }

  // Mouse (and pen/touch, which also start with pointerdown): stop Chrome's
  // own list from opening and show ours instead.
  const isSelect = e => e.target.closest && e.target.closest('select');
  document.addEventListener('pointerdown', e => {
    const select = isSelect(e);
    if (select && e.button === 0) { e.preventDefault(); openSelect(select); }
  }, true);
  document.addEventListener('mousedown', e => { if (isSelect(e)) e.preventDefault(); }, true);
  document.addEventListener('click', e => { if (isSelect(e)) e.preventDefault(); }, true);

  // Keyboard: keys that would open Chrome's list open ours; arrow keys still
  // change the value directly, which never opens a list.
  document.addEventListener('keydown', e => {
    const select = e.target.tagName === 'SELECT' ? e.target : null;
    if (select && (e.key === ' ' || e.key === 'Enter' || e.key === 'F4' || (e.altKey && /Arrow(Down|Up)/.test(e.key)))) {
      e.preventDefault();
      openSelect(select);
    } else if (menu && e.key === 'Escape') {
      close();
    }
  }, true);

  root.ADRDropdown = { pick };
})(window);
