/* Toolkit by GroupChatOn — site script: theme, menu, install, ads, sponsor cards, search, contact. */
(function () {
  'use strict';
  var doc = document, root = doc.documentElement, body = doc.body;
  /* Chrome/Edge 131+ print the @page footer (site name) instead of their own date/title/URL lines */
  (function () { var m = /Chrome\/(\d+)/.exec(navigator.userAgent); if (m && +m[1] >= 131) root.classList.add('tk-pmb'); })();
  var NO_ADS = body.getAttribute('data-ads') === '0';
  var AD_CLIENT = 'ca-pub-5935207738848708';
  var AD_SLOTS = { inArticle: '2020003182', end: '2358914532', home: '3496736383' };
  var PROMOS_URL = 'https://share.groupchaton.com/promos.json';
  var KOFI = { href: 'https://ko-fi.com/Q5Q71K1J5K', own: true, icon: '☕', title: 'Saved you an hour of work?',
    line: 'Toolkit is free, with no sign-up. A coffee on Ko-fi helps us make the next batch of templates.', cta: 'Buy us a coffee', small: 'Any amount helps' };
  function $(s, el) { return (el || doc).querySelector(s); }
  function $$(s, el) { return Array.prototype.slice.call((el || doc).querySelectorAll(s)); }
  function h(tag, attrs, kids) {
    var e = doc.createElement(tag);
    for (var k in attrs || {}) { if (attrs[k] == null) continue; if (k === 'text') e.textContent = attrs[k]; else e.setAttribute(k, attrs[k]); }
    (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? doc.createTextNode(c) : c); });
    return e;
  }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }

  /* theme */
  $$('[data-theme-toggle]').forEach(function (b) {
    b.addEventListener('click', function () {
      var t = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      root.setAttribute('data-theme', t); store('tk-theme', t);
    });
  });
  /* menu */
  var nav = $('#nav'), mb = $('.menu-btn');
  if (mb && nav) {
    mb.addEventListener('click', function () { var o = nav.classList.toggle('open'); mb.setAttribute('aria-expanded', String(o)); });
    doc.addEventListener('keydown', function (e) { if (e.key === 'Escape' && nav.classList.contains('open')) { nav.classList.remove('open'); mb.setAttribute('aria-expanded', 'false'); mb.focus(); } });
  }
  var here = location.pathname.split('/')[1];
  $$('#nav a').forEach(function (a) { if (here && a.getAttribute('href') === '/' + here + '/') a.setAttribute('aria-current', 'page'); });
  /* "/" focuses search */
  doc.addEventListener('keydown', function (e) {
    if (e.key !== '/' || /input|textarea|select/i.test((e.target.tagName || ''))) return;
    var q = $('input[type=search]'); e.preventDefault();
    if (q) q.focus(); else location.href = '/search/';
  });

  /* install as app */
  var deferred = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); deferred = e; $$('[data-install]').forEach(function (b) { b.hidden = false; }); });
  $$('[data-install]').forEach(function (b) {
    b.addEventListener('click', function () {
      if (deferred) { deferred.prompt(); deferred.userChoice.then(function () { deferred = null; }); }
      else alert('To install Toolkit: open your browser menu (⋮ or Share) and choose "Install app" or "Add to Home screen".');
    });
  });
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    window.addEventListener('load', function () { navigator.serviceWorker.register('/sw.js').catch(function () {}); });
  }

  /* ads: only where allowed, never beside buttons (slots sit between sections) */
  if (!NO_ADS) {
    $$('.ad-slot').forEach(function (s) {
      var id = AD_SLOTS[s.getAttribute('data-ad')]; if (!id) return;
      var fluid = s.getAttribute('data-ad') === 'inArticle';
      var ins = h('ins', { 'class': 'adsbygoogle', style: 'display:block' + (fluid ? ';text-align:center' : ''), 'data-ad-client': AD_CLIENT, 'data-ad-slot': id,
        'data-ad-format': fluid ? 'fluid' : 'auto', 'data-ad-layout': fluid ? 'in-article' : null, 'data-full-width-responsive': fluid ? null : 'true' });
      s.appendChild(h('span', { 'class': 'ad-label', text: 'Advertisement' })); s.appendChild(ins); s.classList.add('on');
      try { (window.adsbygoogle = window.adsbygoogle || []).push({}); } catch (e) {}
      setTimeout(function () { if (ins.getAttribute('data-ad-status') !== 'filled') { s.innerHTML = ''; s.classList.remove('on'); } /* 'unfilled', 'unfill-optimized' or blocked */ }, 9000);
    });
  } else { window.adsbygoogle = window.adsbygoogle || []; window.adsbygoogle.pauseAdRequests = 1; }

  /* sponsor + sister-site cards (shared list on share.groupchaton.com): home page 3 (owner), other pages at most 2, always below the tool */
  var slots = $$('.promo-slot'), MAX_PROMOS = body.classList.contains('pg-home') ? 3 : 2;
  function inIndia() { try { var z = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; return z === 'Asia/Kolkata' || z === 'Asia/Calcutta'; } catch (e) { return false; } }
  function pick(list) {
    var india = inIndia(), pool = [KOFI];
    list.forEach(function (p) {
      if (p.india && !india) return;
      if (p.own && !p.family) return;          // other sites' own Ko-fi wording: Toolkit has its own
      if (p.family === 'toolkit') return;        // never our own card here
      pool.push(p);
    });
    for (var i = pool.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)), t = pool[i]; pool[i] = pool[j]; pool[j] = t; }
    return pool;
  }
  function promoCard(p) {
    var own = !!p.own, img = p.img_icon || (p.imgs && p.imgs.length ? p.imgs[Math.floor(Math.random() * p.imgs.length)] : null);
    var pic = img ? h('img', { src: img, alt: '', loading: 'lazy', width: p.img_icon ? 64 : 112, height: p.img_icon ? 64 : 84, 'class': p.img_icon ? 'is-mark' : null })
      : h('span', { 'class': 'promo-ic', 'aria-hidden': 'true', text: p.icon === 'coffee' ? '☕' : (p.icon || '★') });
    var a = h('a', { 'class': 'promo', href: p.href, target: '_blank', rel: own ? 'noopener' : 'sponsored noopener' }, [pic,
      h('span', { 'class': 'promo-b' }, [h('b', { text: p.title }), h('span', { text: p.line }), p.small ? h('small', { text: p.small }) : null, h('span', { 'class': 'promo-cta', text: (p.cta || 'Open') + ' →' })])]);
    return [h('span', { 'class': 'promo-label', text: p.tag || (own ? 'From GroupChatOn' : 'Sponsored') }), a];
  }
  function fill(list) {
    var pool = pick(list || []);
    slots.slice(0, MAX_PROMOS).forEach(function (s, i) { var p = pool[i]; if (!p) return; s.innerHTML = ''; promoCard(p).forEach(function (n) { s.appendChild(n); }); });
  }
  if (slots.length && !NO_ADS) {
    var cached = store('tk-promos');
    fetch(PROMOS_URL, { cache: 'no-cache' }).then(function (r) { return r.json(); }).then(function (d) {
      store('tk-promos', JSON.stringify(d.promos || [])); fill(d.promos);
    }).catch(function () { try { fill(cached ? JSON.parse(cached) : []); } catch (e) { fill([]); } });
  }

  /* search page */
  var res = $('[data-search-results]');
  if (res) {
    var q = new URLSearchParams(location.search).get('q') || '', inp = $('#sq'), count = $('[data-search-count]');
    inp.value = q;
    fetch('/data/index.json').then(function (r) { return r.json(); }).then(function (idx) {
      function run(q) {
        var words = q.toLowerCase().replace(/[^\w\s]/g, ' ').split(/\s+/).filter(function (w) { return w.length > 1 && ['the', 'and', 'for', 'in', 'of', 'to', 'a', 'format', 'free'].indexOf(w) < 0; });
        var hits = idx.map(function (x) {
          var t = x.t.toLowerCase(), d = x.d.toLowerCase(), w = x.w.toLowerCase(), s = 0, all = true;
          words.forEach(function (k) { var m = t.indexOf(k) >= 0 ? 6 : d.indexOf(k) >= 0 ? 3 : w.indexOf(k) >= 0 ? 1 : 0; if (!m) all = false; s += m; });
          return { x: x, s: all ? s : 0 };
        }).filter(function (o) { return o.s > 0 || !words.length; }).sort(function (a, b) { return b.s - a.s; });
        res.innerHTML = '';
        count.textContent = words.length ? hits.length + (hits.length === 1 ? ' result' : ' results') + ' for "' + q + '"' + (hits.length ? '' : '. Try fewer words, or browse a toolkit below.') : 'Everything on Toolkit:';
        (hits.length ? hits : idx.map(function (x) { return { x: x }; })).forEach(function (o) {
          var x = o.x;
          res.appendChild(h('a', { 'class': 'card', href: x.u }, [h('span', { 'class': 'card-ic', 'aria-hidden': 'true', text: x.i }),
            h('span', { 'class': 'card-b' }, [h('b', { text: x.t }), h('small', { text: x.d }), h('em', { 'class': 'kind', text: x.kind })])]));
        });
      }
      run(q);
      $('[data-search]').addEventListener('submit', function (e) { e.preventDefault(); var v = inp.value.trim(); history.replaceState(null, '', '/search/?q=' + encodeURIComponent(v)); run(v); });
    });
  }
  /* 404: suggest a search from the address */
  var q404 = $('[data-404-q]');
  if (q404) q404.value = decodeURIComponent(location.pathname).replace(/\.html?$/, '').split(/[\/\-_]+/).filter(Boolean).join(' ');

  body.classList.add('tk-ready');
})();
