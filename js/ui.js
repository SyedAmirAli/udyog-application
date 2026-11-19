/* Shared UI helpers: header, icons, toasts, dialogs, missing-field warning. */
(function () {
  const S = window.Schema;
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const icons = () => window.lucide && window.lucide.createIcons();
  const qs = (k) => new URLSearchParams(location.search).get(k);

  // opts.menu: show a drawer toggle (mobile/tablet) and move the nav links into the drawer.
  function header(active, right = '', opts = {}) {
    const link = (href, key, icon, text) => `<a href="${href}" class="btn ${active === key ? 'btn-soft' : 'btn-ghost'} max-sm:px-3 ${opts.menu ? 'max-lg:hidden' : ''}" title="${text}"><i data-lucide="${icon}"></i><span class="max-sm:hidden">${text}</span></a>`;
    return `
    <header class="no-print sticky top-0 z-30 border-b border-line bg-white/85 backdrop-blur-md">
      <div class="mx-auto flex h-14 max-w-7xl items-center gap-2 px-3 sm:h-16 sm:gap-3 sm:px-6">
        ${opts.menu ? `<button type="button" id="menu-btn" class="btn btn-ghost btn-icon -ml-1 lg:hidden" aria-label="মেনু খুলুন"><i data-lucide="menu" class="size-5!"></i></button>` : ''}
        <a href="index.html" class="flex min-w-0 items-center gap-2.5">
          <img src="assets/logo.png" alt="" class="size-8 shrink-0 sm:size-9">
          <span class="min-w-0 leading-tight">
            <span class="block truncate text-[17px] font-bold text-brand-800">উদ্যোগ <span class="text-brand-500">(UDYOG)</span></span>
            <span class="block truncate text-xs text-muted max-sm:hidden">অনলাইন আবেদন সহায়ক</span>
          </span>
        </a>
        <nav class="ml-auto flex shrink-0 items-center gap-1">
          ${link('index.html', 'home', 'house', 'হোম')}
          ${link('dashboard.html', 'dashboard', 'layout-dashboard', 'আমার আবেদন')}
          ${right}
        </nav>
      </div>
    </header>`;
  }

  // ---------- drawer (slide-in panel from the left, for phones/tablets) ----------
  function drawer(inner) {
    const el = document.createElement('div');
    el.id = 'drawer';
    el.className = 'no-print fixed inset-0 z-50 invisible lg:hidden';
    el.innerHTML = `
      <div data-drawer-close class="absolute inset-0 bg-brand-900/40 opacity-0 backdrop-blur-[2px] transition-opacity duration-300"></div>
      <aside role="dialog" aria-modal="true" aria-label="মেনু" class="absolute inset-y-0 left-0 flex w-[min(21rem,88vw)] -translate-x-full flex-col bg-canvas shadow-2xl transition-transform duration-300 ease-out">
        <div class="flex h-14 shrink-0 items-center gap-2.5 border-b border-line bg-white px-4">
          <img src="assets/logo.png" alt="" class="size-8">
          <span class="flex-1 text-[17px] font-bold text-brand-800">উদ্যোগ <span class="text-brand-500">(UDYOG)</span></span>
          <button type="button" data-drawer-close class="btn btn-ghost btn-icon -mr-1" aria-label="বন্ধ করুন"><i data-lucide="x" class="size-5!"></i></button>
        </div>
        <div class="flex-1 space-y-4 overflow-y-auto overscroll-contain p-4 pb-[max(1rem,env(safe-area-inset-bottom))]">
          ${inner}
          <nav class="card p-2">
            <a href="index.html" class="flex items-center gap-3 rounded-xl px-3 py-2.5 font-semibold hover:bg-brand-50"><i data-lucide="house" class="size-5 text-muted"></i>হোম</a>
            <a href="dashboard.html" class="flex items-center gap-3 rounded-xl px-3 py-2.5 font-semibold hover:bg-brand-50"><i data-lucide="layout-dashboard" class="size-5 text-muted"></i>আমার আবেদনসমূহ</a>
          </nav>
        </div>
      </aside>`;
    document.body.appendChild(el);
    const [shade, panel] = [el.firstElementChild, el.querySelector('aside')];
    let lastFocus = null;
    const api = {
      el,
      open() {
        lastFocus = document.activeElement;
        el.classList.remove('invisible');
        requestAnimationFrame(() => { shade.classList.remove('opacity-0'); panel.classList.remove('-translate-x-full'); });
        document.documentElement.style.overflow = 'hidden';
        panel.querySelector('[data-drawer-close]').focus({ preventScroll: true });
      },
      close() {
        shade.classList.add('opacity-0'); panel.classList.add('-translate-x-full');
        document.documentElement.style.overflow = '';
        setTimeout(() => el.classList.add('invisible'), 300);
        lastFocus?.focus?.({ preventScroll: true });
      },
    };
    el.addEventListener('click', (e) => { if (e.target.closest('[data-drawer-close], a[href], [data-step]')) api.close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !el.classList.contains('invisible')) api.close(); });
    document.getElementById('menu-btn')?.addEventListener('click', api.open);
    icons();
    return api;
  }

  // ---------- toast ----------
  function toast(msg, kind = 'ok') {
    let host = document.getElementById('toast-host');
    if (!host) {
      host = document.createElement('div');
      host.id = 'toast-host';
      host.className = 'no-print fixed inset-x-0 bottom-5 z-[60] flex flex-col items-center gap-2 px-4 pointer-events-none';
      document.body.appendChild(host);
    }
    const el = document.createElement('div');
    const icon = { ok: 'circle-check', warn: 'triangle-alert', err: 'circle-x' }[kind] || 'info';
    const tone = { ok: 'text-brand-400', warn: 'text-sun-400', err: 'text-berry-500' }[kind] || 'text-white';
    el.className = 'pointer-events-auto flex items-center gap-2.5 rounded-2xl bg-brand-900 px-4 py-3 text-sm font-medium text-white shadow-xl shadow-black/20 transition duration-300 translate-y-2 opacity-0';
    el.innerHTML = `<i data-lucide="${icon}" class="size-4.5 ${tone}"></i><span>${esc(msg)}</span>`;
    host.appendChild(el);
    icons();
    requestAnimationFrame(() => el.classList.remove('translate-y-2', 'opacity-0'));
    setTimeout(() => { el.classList.add('opacity-0'); setTimeout(() => el.remove(), 300); }, 2600);
  }

  // ---------- dialog ----------
  // opts: { title, body (html), icon, tone: 'brand'|'warn'|'danger', ok, cancel, wide }
  function dialog(opts) {
    return new Promise((resolve) => {
      const d = document.createElement('dialog');
      const tone = { brand: 'bg-brand-50 text-brand-700', warn: 'bg-sun-400/15 text-sun-500', danger: 'bg-berry-500/10 text-berry-500' }[opts.tone || 'brand'];
      const okCls = opts.tone === 'danger' ? 'bg-berry-500 text-white hover:bg-berry-500/90' : 'btn-primary';
      d.className = `no-print m-auto w-[calc(100%-2rem)] ${opts.wide ? 'max-w-lg' : 'max-w-md'} rounded-3xl border border-line bg-white p-0 text-ink shadow-2xl backdrop:bg-brand-900/40 backdrop:backdrop-blur-sm`;
      d.innerHTML = `
        <form method="dialog" class="p-6">
          <div class="flex items-start gap-3.5">
            <span class="grid size-11 shrink-0 place-items-center rounded-2xl ${tone}"><i data-lucide="${opts.icon || 'info'}" class="size-5.5"></i></span>
            <div class="min-w-0 flex-1 pt-1">
              <h2 class="text-lg font-bold leading-snug">${opts.title}</h2>
              <div class="mt-2 text-[15px] leading-relaxed text-muted">${opts.body || ''}</div>
            </div>
          </div>
          <div class="mt-6 flex flex-wrap justify-end gap-2">
            ${opts.cancel === false ? '' : `<button value="cancel" class="btn btn-outline">${opts.cancel || 'বাতিল'}</button>`}
            <button value="ok" class="btn ${okCls}" autofocus>${opts.ok || 'ঠিক আছে'}</button>
          </div>
        </form>`;
      document.body.appendChild(d);
      icons();
      d.addEventListener('close', () => { resolve(d.returnValue === 'ok'); d.remove(); });
      d.addEventListener('click', (e) => { if (e.target === d) d.close('cancel'); });
      d.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', () => d.close(b.dataset.close)));
      d.showModal();
    });
  }

  // Warning before print/download listing every required field still empty.
  // Resolves 'print' | 'fix' | null. Clicking an item opens that field in the form.
  function missingWarning(app) {
    const miss = S.missing(app.data || {});
    if (!miss.length) return Promise.resolve('print');
    const steps = S.STEPS;
    const byStep = {};
    miss.forEach((m) => (byStep[m.step] = byStep[m.step] || []).push(m));
    const list = Object.keys(byStep).map((si) => `
      <div class="mt-3 first:mt-0">
        <div class="mb-1.5 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-brand-600">
          <i data-lucide="${steps[si].icon}" class="size-3.5"></i>${steps[si].part} · ${steps[si].title}
        </div>
        <ul class="space-y-1">
          ${byStep[si].map((m) => `<li><a href="form.html?id=${encodeURIComponent(app.id)}&step=${m.step}&focus=${encodeURIComponent(m.id)}" class="group flex items-start gap-2 rounded-lg px-2 py-1.5 text-[14.5px] text-ink hover:bg-brand-50">
            <span class="mt-2 size-1.5 shrink-0 rounded-full bg-berry-500"></span><span class="flex-1">${esc(m.label)}</span>
            <i data-lucide="arrow-up-right" class="mt-0.5 size-4 text-muted opacity-0 transition group-hover:opacity-100"></i></a></li>`).join('')}
        </ul>
      </div>`).join('');
    return dialog({
      title: `<span class="text-berry-500">কিছু ঘর পূরণ করা হয়নি</span>`,
      icon: 'triangle-alert', tone: 'warn', wide: true,
      body: `<p>${S.toBn(miss.length)}টি প্রয়োজনীয় ঘর এখনো ফাঁকা। কোনো ঘরে ক্লিক করলে সরাসরি সেখানে যেতে পারবেন।</p>
        <div class="mt-3 max-h-[46vh] overflow-y-auto rounded-2xl border border-line bg-canvas/60 p-3">${list}</div>
        <p class="mt-4 font-medium text-ink">তবুও কি প্রিন্ট/ডাউনলোড করতে চান?</p>`,
      cancel: 'না, পূরণ করব',
      ok: '<i data-lucide="download"></i>হ্যাঁ, ডাউনলোড করুন',
    }).then((ok) => (ok ? 'print' : 'fix'));
  }

  function relTime(ts) {
    const s = Math.round((Date.now() - ts) / 1000);
    if (s < 45) return 'এইমাত্র';
    const m = Math.round(s / 60); if (m < 60) return S.toBn(m) + ' মিনিট আগে';
    const h = Math.round(m / 60); if (h < 24) return S.toBn(h) + ' ঘণ্টা আগে';
    const dd = Math.round(h / 24); if (dd < 30) return S.toBn(dd) + ' দিন আগে';
    return new Date(ts).toLocaleDateString('bn-BD');
  }

  function download(name, text, type = 'application/json') {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type }));
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  window.addEventListener('store:error', () => toast('ব্রাউজারের স্টোরেজ পূর্ণ! ছবি ছোট করুন বা পুরোনো আবেদন মুছুন।', 'err'));

  window.UI = { esc, icons, qs, header, drawer, toast, dialog, missingWarning, relTime, download };
})();
