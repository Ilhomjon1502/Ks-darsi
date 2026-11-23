/* O'zbek lotin <-> kirill almashtirgich (Ўзбек лотин ва кирилл).
   Sayt matni lotinda yoziladi; "Кир" tugmasi sahifani kirillga o'giradi, "Lot" tugmasi qaytaradi.
   Tanlov brauzerda saqlanadi. Havola orqali: ?lang=cyr yoki ?lang=lat.
   Lotincha qoladigan narsalar: brendlar (Windows, Word, Google...), klavisha nomlari (Ctrl, Alt...),
   qisqartmalar (RAM, USB, PDF...), fayl nomlari, saytlar, formulalar va raqamli kodlar. */
(function () {
  'use strict';
  var KEY = 'ks-script';
  var store = {
    get: function (k) { try { return localStorage.getItem(k); } catch (_) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (_) {} }
  };
  var mode = (function () {
    var q = null;
    try { q = new URLSearchParams(location.search).get('lang'); } catch (_) {}
    if (q === 'cyr' || q === 'lat') { store.set(KEY, q); return q; }
    return store.get(KEY) === 'cyr' ? 'cyr' : 'lat';
  })();

  /* ---------- lotin -> kirill ---------- */
  var APOS = /['’ʻʼ`‘´]/;
  var isApos = function (ch) { return !!ch && APOS.test(ch); };
  var MAP = {
    a: 'а', b: 'б', c: 'к', d: 'д', e: 'е', f: 'ф', g: 'г', h: 'ҳ', i: 'и', j: 'ж', k: 'к', l: 'л', m: 'м',
    n: 'н', o: 'о', p: 'п', q: 'қ', r: 'р', s: 'с', t: 'т', u: 'у', v: 'в', w: 'в', x: 'х', y: 'й', z: 'з'
  };
  var isUp = function (c) { return c !== c.toLowerCase(); };

  function conv(w) {
    var out = '', i = 0, n = w.length;
    var allUp = n > 1 && w === w.toUpperCase();
    function put(s, c, nextC) {          // s kichik harfda; c - birinchi harf, nextC - ikkinchi harf
      if (!isUp(c)) return s;
      if (s.length === 1) return s.toUpperCase();
      return (allUp || (nextC && isUp(nextC))) ? s.toUpperCase() : s.charAt(0).toUpperCase() + s.slice(1);
    }
    while (i < n) {
      var c = w.charAt(i), l = c.toLowerCase(), x = w.charAt(i + 1), xl = x.toLowerCase(), x2 = w.charAt(i + 2);
      if (isApos(c)) {                   // harflar orasidagi tutuq belgisi -> ъ
        var prev = i > 0 && /[a-z]/i.test(w.charAt(i - 1)), next = i < n - 1 && /[a-z]/i.test(x);
        out += (prev && next) ? 'ъ' : c; i++; continue;
      }
      if ((l === 'o' || l === 'g') && isApos(x)) { out += put(l === 'o' ? 'ў' : 'ғ', c); i += 2; continue; }
      if (l === 's' && xl === 'h') { out += put('ш', c); i += 2; continue; }
      if (l === 'c' && xl === 'h') { out += put('ч', c); i += 2; continue; }
      if (l === 'n' && xl === 'g' && !isApos(x2)) { out += put('нг', c, x); i += 2; continue; }
      if (l === 'y' && !(xl === 'o' && isApos(x2))) {
        if (xl === 'o') { out += put('ё', c); i += 2; continue; }
        if (xl === 'u') { out += put('ю', c); i += 2; continue; }
        if (xl === 'a') { out += put('я', c); i += 2; continue; }
        if (xl === 'e') { out += put((i === 0 || w.charAt(i - 1).toLowerCase() === 'i') ? 'е' : 'йе', c, x); i += 2; continue; }
      }
      if (l === 't' && xl === 's' && (i === 0 || /^ts(iya|ion|ent|ikl|irk|el|ex)/i.test(w.slice(i)))) {
        out += put('ц', c); i += 2; continue;
      }
      if (l === 'e' && i === 0) { out += put('э', c); i++; continue; }
      if (MAP[l]) { out += put(MAP[l], c); i++; continue; }
      out += c; i++;
    }
    return out;
  }

  /* lotincha qoladigan so'zlar (kichik harfda) */
  var KEEP = {};
  ('codial windows word excel powerpoint microsoft office google chrome gmail gemini chatgpt copilot canva gamma designer ' +
   'docs drive telegram instagram youtube android apple linux macos ios iphone explorer paint store start download ' +
   'downloads new link family times roman emirates vibe coding co-working wi-fi wifi bluetooth backend frontend max ' +
   'deepfake crack docx xlsx pptx txt jpg png zip pdf ctrl alt shift tab enter esc backspace delete caps lock win prtsc space home end fn insert del').split(' ')
    .forEach(function (k) { KEEP[k] = 1; });
  var BRANDS = 'codial windows excel powerpoint explorer google chrome telegram instagram youtube gmail chatgpt canva gamma gemini copilot android linux office microsoft word paint apple'.split(' ')
    .sort(function (a, b) { return b.length - a.length; });
  var SUFFIX = /^['’ʻʼ`‘´]?(da|dan|ga|ni|ning|dagi|lar|larni|larning|ka|gacha|i|si)$/i;
  var ABBR = {};
  ('AI RAM USB PDF PIN ZIP SMM SMS IT ID GB MB KB TB UZ EN RU MS ERP CPU GPU SSD HDD BIOS URL QR PC IF SUM MAX MIN AVERAGE COUNT COUNTIF SUMIF ROUND VALUE NAME XLOOKUP').split(' ')
    .forEach(function (k) { ABBR[k] = 1; });

  var OVR = [['gallyutsinatsiya', 'галлюцинация'], ['qayer', 'қаер'], ['kompyuter', 'компьютер'], ['kalkulyator', 'калькулятор'], ['protsess', 'процесс']];

  function convPart(p) {
    if (!/[A-Za-z]/.test(p)) return p;
    if (/[A-Za-z][0-9]|[0-9][A-Za-z]/.test(p)) return p;           // F2, 3D, 2FA, ali2012
    if (/[a-z][A-Z]/.test(p)) return p;                              // iPhone, macOS, ChatGPT, YouTube
    if (p.length === 1) return p;                                    // klavisha harflari
    if (ABBR[p] || KEEP[p.toLowerCase()]) return p;
    var low = p.toLowerCase(), word = function (q) { return q.replace(/[A-Za-z][A-Za-z'’ʻʼ`‘´]*/g, conv); };
    for (var j = 0; j < OVR.length; j++) {            // imlosi qoidadan chiqadigan so'zlar
      if (low.indexOf(OVR[j][0]) === 0) {
        var rep = OVR[j][1];
        if (isUp(p.charAt(0))) rep = rep.charAt(0).toUpperCase() + rep.slice(1);
        return rep + word(p.slice(OVR[j][0].length));
      }
    }
    for (var i = 0; i < BRANDS.length; i++) {
      var b = BRANDS[i];
      if (low.length > b.length && low.indexOf(b) === 0 && SUFFIX.test(p.slice(b.length))) {
        return p.slice(0, b.length) + conv(p.slice(b.length));
      }
    }
    return p.replace(/[A-Za-z][A-Za-z'’ʻʼ`‘´]*/g, conv);
  }

  function convTok(tok) {
    var m = /^([^A-Za-z0-9]*)([\s\S]*?)([^A-Za-z0-9'’ʻʼ`‘´]*)$/.exec(tok);
    if (!m || !/[A-Za-z]/.test(m[2])) return tok;
    var core = m[2];
    if (/[@\\_:=$><#+%]|https?|[A-Za-z0-9]\.[A-Za-z0-9]|[A-Za-z0-9]\(|\)[A-Za-z0-9]/.test(core)) return tok;   // sayt, fayl, formula
    if (KEEP[core.toLowerCase()]) return tok;
    var res = core.split(/([-–\/])/).map(function (part, k) { return k % 2 ? part : convPart(part); }).join('');
    return m[1] + res + m[3];
  }

  function convText(s) { return s.replace(/[^\s]+/g, convTok); }

  /* ---------- DOM ---------- */
  var ATTRS = ['title', 'aria-label', 'placeholder', 'alt'];
  var origText = new WeakMap(), origAttr = new WeakMap();
  var changed = false;

  function skipEl(el) { return !el || (el.closest && el.closest('script,style,textarea,input,[translate="no"],.notranslate')); }

  function toCyrText(n) {
    var v = n.nodeValue;
    if (!v || !/[A-Za-z]/.test(v) || skipEl(n.parentElement)) return;
    var rec = origText.get(n);
    if (rec && v === rec.cyr) return;
    var c = convText(v);
    if (c === v) { origText.delete(n); return; }
    origText.set(n, { lat: v, cyr: c });
    n.nodeValue = c; changed = true;
  }
  function toCyrAttr(el, a) {
    var v = el.getAttribute(a);
    if (!v || !/[A-Za-z]/.test(v) || skipEl(el)) return;
    var m = origAttr.get(el) || {}, rec = m[a];
    if (rec && v === rec.cyr) return;
    var c = convText(v);
    if (c === v) { if (rec) delete m[a]; return; }
    m[a] = { lat: v, cyr: c }; origAttr.set(el, m);
    el.setAttribute(a, c); changed = true;
  }
  var SEL = '[title],[aria-label],[placeholder],[alt]';
  function walk(root) {
    if (!root) return;
    if (root.nodeType === 3) { toCyrText(root); return; }
    if (root.nodeType !== 1 && root.nodeType !== 9) return;
    var base = root.nodeType === 9 ? root.documentElement : root;
    var tw = document.createTreeWalker(base, NodeFilter.SHOW_TEXT), n;
    while ((n = tw.nextNode())) toCyrText(n);
    var els = base.querySelectorAll ? [].slice.call(base.querySelectorAll(SEL)) : [];
    if (base.matches && base.matches(SEL)) els.push(base);
    els.forEach(function (el) { ATTRS.forEach(function (a) { if (el.hasAttribute(a)) toCyrAttr(el, a); }); });
  }
  function revert() {
    var tw = document.createTreeWalker(document.documentElement, NodeFilter.SHOW_TEXT), n, rec;
    while ((n = tw.nextNode())) {
      rec = origText.get(n);
      if (rec && n.nodeValue === rec.cyr) n.nodeValue = rec.lat;
      origText.delete(n);
    }
    [].slice.call(document.querySelectorAll(SEL)).forEach(function (el) {
      var m = origAttr.get(el); if (!m) return;
      ATTRS.forEach(function (a) { if (m[a] && el.getAttribute(a) === m[a].cyr) el.setAttribute(a, m[a].lat); });
      origAttr.delete(el);
    });
  }

  var fitT = 0;
  function refit() {   // matn uzunligi o'zgargani uchun slayd shriftini qayta moslaymiz
    clearTimeout(fitT);
    fitT = setTimeout(function () { try { if (typeof fit === 'function') fit(); } catch (_) {} }, 40);
  }

  var mo = new MutationObserver(function (recs) {
    if (mode !== 'cyr') return;
    changed = false;
    recs.forEach(function (r) {
      if (r.type === 'characterData') toCyrText(r.target);
      else if (r.type === 'attributes') toCyrAttr(r.target, r.attributeName);
      else for (var k = 0; k < r.addedNodes.length; k++) walk(r.addedNodes[k]);
    });
    if (changed) refit();
  });
  mo.observe(document.documentElement, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ATTRS });

  /* ---------- tugma ---------- */
  var cyr = function () { return mode === 'cyr'; };
  function label(btn) {
    btn.textContent = cyr() ? 'Lot' : 'Кир';
    btn.setAttribute('aria-label', cyr() ? "Lotin yozuviga o'tish" : 'Кирилл ёзувига ўтиш');
    btn.title = btn.getAttribute('aria-label');
  }
  function setMode(m) {
    if (m === mode) return;
    mode = m; store.set(KEY, m);
    document.documentElement.lang = cyr() ? 'uz-Cyrl' : 'uz';
    changed = false;
    if (cyr()) walk(document); else revert();
    [].slice.call(document.querySelectorAll('[data-lang-t]')).forEach(label);
    refit();
  }
  document.addEventListener('click', function (e) {
    var b = e.target && e.target.closest && e.target.closest('[data-lang-t]');
    if (b) { e.preventDefault(); setMode(cyr() ? 'lat' : 'cyr'); }
  });
  var st = document.createElement('style');
  st.textContent = '.lang-btn{padding:0 10px;font-size:.85rem;letter-spacing:.02em}';
  document.head.appendChild(st);
  if (cyr()) document.documentElement.lang = 'uz-Cyrl';

  window.KSLang = {
    mode: function () { return mode; },
    set: setMode,
    convert: convText,
    btn: function () {
      var t = cyr() ? "Lotin yozuviga o'tish" : 'Кирилл ёзувига ўтиш';
      return '<button class="ibtn lang-btn" data-lang-t translate="no" aria-label="' + t + '" title="' + t + '">' + (cyr() ? 'Lot' : 'Кир') + '</button>';
    }
  };
})();
