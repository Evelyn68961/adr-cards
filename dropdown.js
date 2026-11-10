// Replaces the browser's own dropdown list for every <select> on the page
// with one drawn by the page. On some Windows PCs Chrome crashes the page
// (STATUS_BREAKPOINT) when its native dropdown opens, especially in the
// --app window; this list is plain HTML, so it cannot do that.
// The <select> stays the source of truth: picking an item sets its value and
// fires `change`, so page code works unchanged. An <option data-ask="label">
// asks for a new value inside the list (no prompt() dialog) and adds it.
(function () {
  'use strict';

  let menu = null;

  function close() {
    if (!menu) return;
    menu.remove();
    menu = null;
    document.removeEventListener('mousedown', outside, true);
    window.removeEventListener('resize', close);
    window.removeEventListener('scroll', reposition, true);
  }

  function outside(e) {
    if (menu && !menu.contains(e.target) && e.target !== menu.select) close();
  }

  function reposition() {
    if (!menu) return;
    if (!menu.select.isConnected) return close();
    const r = menu.select.getBoundingClientRect();
    if (r.bottom < 0 || r.top > window.innerHeight) return close();
    const below = window.innerHeight - r.bottom;
    menu.style.minWidth = r.width + 'px';
    menu.style.left = Math.max(4, Math.min(r.left, window.innerWidth - menu.offsetWidth - 4)) + 'px';
    if (below < menu.offsetHeight + 8 && r.top > below) {
      menu.style.top = Math.max(4, r.top - menu.offsetHeight - 2) + 'px';
    } else {
      menu.style.top = (r.bottom + 2) + 'px';
    }
  }

  function choose(select, value) {
    close();
    select.focus({ preventScroll: true });
    if (select.value !== value) {
      select.value = value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }

  function ask(select, option) {
    menu.replaceChildren();
    const input = document.createElement('input');
    input.className = 'dd-input';
    input.placeholder = option.dataset.ask;
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.className = 'dd-ok';
    ok.textContent = '確定';
    const done = () => {
      const value = input.value.trim();
      if (!value) return close();
      if (![...select.options].some(o => o.value === value)) {
        select.insertBefore(new Option(value, value), option);
      }
      choose(select, value);
    };
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
    menu.append(row);
    reposition();
    input.focus();
  }

  function open(select) {
    close();
    if (select.disabled) return;
    menu = document.createElement('div');
    menu.className = 'dd-menu';
    menu.setAttribute('role', 'listbox');
    menu.select = select;
    [...select.options].forEach(option => {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'dd-item' + (option.value === select.value ? ' current' : '');
      item.textContent = option.textContent;
      item.disabled = option.disabled;
      item.addEventListener('click', () => (option.dataset.ask ? ask(select, option) : choose(select, option.value)));
      menu.appendChild(item);
    });
    document.body.appendChild(menu);
    reposition();
    const current = menu.querySelector('.current');
    if (current) current.scrollIntoView({ block: 'nearest' });
    document.addEventListener('mousedown', outside, true);
    window.addEventListener('resize', close);
    window.addEventListener('scroll', reposition, true);
  }

  // Mouse: stop the native list from opening and show ours instead.
  document.addEventListener('mousedown', e => {
    const select = e.target.closest && e.target.closest('select');
    if (!select || e.button !== 0) return;
    e.preventDefault();
    if (menu && menu.select === select) close();
    else open(select);
  }, true);

  // Keyboard: the keys that would open the native list open ours; arrow keys
  // still change the value directly, which never opens a list.
  document.addEventListener('keydown', e => {
    const select = e.target.tagName === 'SELECT' ? e.target : null;
    if (select && (e.key === ' ' || e.key === 'Enter' || e.key === 'F4' || (e.altKey && /Arrow(Down|Up)/.test(e.key)))) {
      e.preventDefault();
      open(select);
    } else if (menu && e.key === 'Escape') {
      close();
    }
  }, true);
})();
