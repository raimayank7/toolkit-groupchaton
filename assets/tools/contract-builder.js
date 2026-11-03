/* Contract builder (Contracts & Agreements toolkit): pick one of 35 agreement types, a country, answer that type's questions,
   switch optional clauses → a numbered agreement with definitions, clauses, schedules, signature and witness blocks and a
   "Before you sign" page. Edit text in the preview, Print / PDF, Copy text or download .docx (docx library, loaded when needed).
   Everything is made in the browser from the clause text below; nothing is uploaded. Starting drafts, never legal advice.
   The shared engine (tools.js) makes table documents, so this tool mounts its own form and preview into .tk-tool[data-tool=contract-builder]. */
(function () {
  'use strict';
  var doc = document, TOOL = 'contract-builder', KEY = 'tk-draft-' + TOOL;
  var ME = (doc.currentScript && doc.currentScript.src) || '/assets/tools/contract-builder.js', PACKV = '1';   /* bump PACKV when packs change */
  var DOCX = ['https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js', 'docx'];
  var CREDIT = 'toolkit.groupchaton.com';   /* = TKTools.MADE: the only line printed or put in files */
  var NS = (window.TKTools && TKTools.CREDIT) || 'Designed & Developed by: Nivesh Sahayak';   /* owner: on screen only (class foot-credit), never printed or in files */

  /* ---------------- helpers */
  function h(tag, a, kids) {
    var e = doc.createElement(tag);
    for (var k in a || {}) { var v = a[k]; if (v == null || v === false) continue; if (k === 'text') e.textContent = v; else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v); else e.setAttribute(k, v === true ? '' : v); }
    (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? doc.createTextNode(c) : c); });
    return e;
  }
  function store(v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(KEY) || 'null'); if (v === null) localStorage.removeItem(KEY); else localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) { return null; } }
  function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/[, ]/g, '')); return isFinite(n) ? n : 0; }
  function blank(v) { return v == null || String(v).trim() === ''; }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function iso(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function parse(s) { var p = String(s || '').split('-'); return p.length === 3 ? new Date(+p[0], +p[1] - 1, +p[2]) : null; }
  function addMonths(s, m, minusDay) { var d = parse(s); if (!d || !m) return ''; var day = d.getDate(); d.setDate(1); d.setMonth(d.getMonth() + Math.round(m)); d.setDate(Math.min(day, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate())); if (minusDay) d.setDate(d.getDate() - 1); return iso(d); }
  function addDays(s, n) { var d = parse(s); if (!d) return ''; d.setDate(d.getDate() + n); return iso(d); }
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function ord(n) { if (LANG) return LANG.ord ? LANG.ord(n) : String(n); var s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function w2(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''); }
  function w3(n) { var hu = Math.floor(n / 100), r = n % 100; return (hu ? ONES[hu] + ' Hundred' + (r ? ' and ' : '') : '') + (r ? w2(r) : ''); }
  function wordsIntl(n) { if (!n) return 'Zero'; var out = []; [[1e12, 'Trillion'], [1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand']].forEach(function (u) { if (n >= u[0]) { out.push(w3(Math.floor(n / u[0])) + ' ' + u[1]); n %= u[0]; } }); if (n) out.push(w3(n)); return out.join(' '); }
  function wordsIndian(n) {
    if (!n) return 'Zero'; var out = [], cr = Math.floor(n / 1e7), rest = n % 1e7;
    if (cr) out.push((cr > 99 ? wordsIndian(cr) : w2(cr)) + ' Crore');
    var lk = Math.floor(rest / 1e5), th = Math.floor(rest % 1e5 / 1e3), hu = Math.floor(rest % 1e3 / 100), re = rest % 100;
    if (lk) out.push(w2(lk) + ' Lakh'); if (th) out.push(w2(th) + ' Thousand'); if (hu) out.push(ONES[hu] + ' Hundred'); if (re) out.push(w2(re));
    return out.join(' ');
  }
  /* ---------------- languages: India drafts in English or any of the 22 scheduled languages of India. A language pack
     (cb-lang/<code>.js, loaded only when picked) calls TKContract.addLang(code, pack): pack.t maps each English template below
     ({0}, {1} = the parts filled in) or fixed sentence to its translation; pack.months, pack.num (or pack.words) give dates and
     amounts in words. Anything a pack lacks stays in English. [code, name in its own script, English name, script] */
  var LANGS = [['en', 'English', 'English', 'Latn'], ['as', 'অসমীয়া', 'Assamese', 'Beng'], ['bn', 'বাংলা', 'Bengali', 'Beng'],
    ['doi', 'डोगरी', 'Dogri', 'Deva'], ['gu', 'ગુજરાતી', 'Gujarati', 'Gujr'], ['hi', 'हिन्दी', 'Hindi', 'Deva'], ['kn', 'ಕನ್ನಡ', 'Kannada', 'Knda'],
    ['ks', 'کٲشُر', 'Kashmiri', 'Arab'], ['kok', 'कोंकणी', 'Konkani', 'Deva'], ['mai', 'मैथिली', 'Maithili', 'Deva'], ['ml', 'മലയാളം', 'Malayalam', 'Mlym'],
    ['mni', 'মৈতৈলোন্', 'Manipuri (Meitei, Bengali script)', 'Beng'], ['mr', 'मराठी', 'Marathi', 'Deva'], ['ne', 'नेपाली', 'Nepali', 'Deva'], ['or', 'ଓଡ଼ିଆ', 'Odia', 'Orya'],
    ['pa', 'ਪੰਜਾਬੀ', 'Punjabi', 'Guru'], ['sa', 'संस्कृतम्', 'Sanskrit', 'Deva'], ['sd', 'سنڌي', 'Sindhi', 'Arab'],
    ['ta', 'தமிழ்', 'Tamil', 'Taml'], ['te', 'తెలుగు', 'Telugu', 'Telu'], ['ur', 'اردو', 'Urdu', 'Arab']];
  var SCRIPT = {   /* script: [Google Noto font for the page, font set in Word files, right-to-left] */
    Latn: ['', 'Cambria'], Beng: ['Noto Serif Bengali', 'Nirmala UI'], Deva: ['Noto Serif Devanagari', 'Nirmala UI'], Gujr: ['Noto Serif Gujarati', 'Nirmala UI'],
    Guru: ['Noto Serif Gurmukhi', 'Nirmala UI'], Knda: ['Noto Serif Kannada', 'Nirmala UI'], Mlym: ['Noto Serif Malayalam', 'Nirmala UI'], Orya: ['Noto Serif Oriya', 'Nirmala UI'],
    Taml: ['Noto Serif Tamil', 'Nirmala UI'], Telu: ['Noto Serif Telugu', 'Nirmala UI'], Olck: ['Noto Sans Ol Chiki', 'Nirmala UI'], Arab: ['Noto Naskh Arabic', 'Arial', true] };
  var LANG = null, LANGINFO = null, PACKS = {}, SEEN = null;
  function langInfo(code) { for (var i = 0; i < LANGS.length; i++) if (LANGS[i][0] === code) return LANGS[i]; return null; }
  function tr(s) {
    if (SEEN && typeof s === 'string') SEEN.s[s] = 1;
    if (LANG && s === '.' && LANG.stop) return LANG.stop;   /* a bare full stop the engine adds: the language's own (। or ۔) */
    return LANG && typeof s === 'string' && LANG.t[s] != null ? LANG.t[s] : s;
  }
  function L(tpl, a) {
    if (SEEN) SEEN.t[tpl] = 1;
    var s = LANG && LANG.t[tpl] != null ? LANG.t[tpl] : tpl;
    s = s.replace(/\{(\d+)\}/g, function (m, i) { var v = a && a[+i]; return v == null ? '' : LANG ? tr(String(v)) : String(v); });
    if (SEEN) SEEN.o[s] = 1;
    return s;
  }
  function langWords(n) {   /* whole number in words in the picked language, Indian system (crore, lakh, thousand, hundred) */
    var P = LANG; if (P.words) return P.words(n); var u = P.num.u; if (!n) return P.num.zero || u[0];
    var out = [], cr = Math.floor(n / 1e7), r = n % 1e7, lk = Math.floor(r / 1e5), th = Math.floor(r % 1e5 / 1e3), hu = Math.floor(r % 1e3 / 100), re = r % 100;
    if (cr) out.push((cr > 99 ? langWords(cr) : u[cr]) + ' ' + P.num.c); if (lk) out.push(u[lk] + ' ' + P.num.l); if (th) out.push(u[th] + ' ' + P.num.th);
    if (hu) out.push(u[hu] + ' ' + P.num.h); if (re) out.push(u[re]);
    return out.join(' ');
  }
  function trDeep(o) { if (typeof o === 'string') return tr(o); if (Array.isArray(o)) return o.map(trDeep); if (o && typeof o === 'object') { var r = {}; for (var k in o) r[k] = trDeep(o[k]); return r; } return o; }
  function loadPack(code) {
    if (!code || code === 'en' || PACKS[code]) return Promise.resolve(PACKS[code] || null);
    return new Promise(function (ok, bad) {
      var v = (ME.match(/[?&]v=([^&]+)/) || [])[1] || '';
      var sc = doc.createElement('script'); sc.src = ME.replace(/[?#].*$/, '').replace(/[^\/]*$/, '') + 'cb-lang/' + code + '.js?v=' + v + '.' + PACKV; sc.async = true;
      sc.onload = function () { PACKS[code] ? ok(PACKS[code]) : bad(new Error('The language file did not load.')); };
      sc.onerror = function () { bad(new Error('Could not load the language. Check the internet connection and try again.')); };
      doc.head.appendChild(sc);
    });
  }
  function loadFont(code) {   /* Google Noto font for the script, so the preview and Print / PDF show it correctly on every device */
    var li = langInfo(code), f = li && SCRIPT[li[3]][0], id = 'ct-font-' + (li && li[3]);
    if (!f || doc.getElementById(id)) return;
    doc.head.appendChild(h('link', { id: id, rel: 'stylesheet', href: 'https://fonts.googleapis.com/css2?family=' + f.replace(/ /g, '+') + ':wght@400;700&display=swap' }));
  }
  function langMoney(whole, minor) { return LANG.money ? LANG.money(langWords(whole), minor ? langWords(minor) : '') : minor ? L('Rupees {0} and {1} Paise only', [langWords(whole), langWords(minor)]) : L('Rupees {0} only', [langWords(whole)]); }

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  /* ---------------- currencies and countries */
  var CUR = {   /* code: [symbol, major name, minor name, number locale] */
    INR: ['₹', 'Rupees', 'Paise', 'en-IN'], GBP: ['£', 'Pounds Sterling', 'Pence', 'en-GB'], USD: ['US$', 'US Dollars', 'Cents', 'en-US'],
    AED: ['AED', 'UAE Dirhams', 'Fils', 'en-US'], EUR: ['€', 'Euros', 'Cents', 'en-IE'], SAR: ['SAR', 'Saudi Riyals', 'Halalas', 'en-US'],
    SGD: ['S$', 'Singapore Dollars', 'Cents', 'en-SG'], AUD: ['A$', 'Australian Dollars', 'Cents', 'en-AU'], CAD: ['C$', 'Canadian Dollars', 'Cents', 'en-CA'],
    NZD: ['NZ$', 'New Zealand Dollars', 'Cents', 'en-NZ'], ZAR: ['R', 'South African Rand', 'Cents', 'en-ZA'], NGN: ['₦', 'Naira', 'Kobo', 'en-NG'],
    KES: ['KSh', 'Kenyan Shillings', 'Cents', 'en-KE'], PHP: ['₱', 'Philippine Pesos', 'Centavos', 'en-PH'], MYR: ['RM', 'Malaysian Ringgit', 'Sen', 'en-MY'],
    NPR: ['Rs', 'Nepalese Rupees', 'Paisa', 'en-IN'], BDT: ['Tk', 'Bangladeshi Taka', 'Poisha', 'en-IN'], LKR: ['Rs', 'Sri Lankan Rupees', 'Cents', 'en-US'], PKR: ['Rs', 'Pakistani Rupees', 'Paisa', 'en-US'] };
  var COUNTRY = {
    IN: { name: 'India', cur: 'INR', region: 'State', regionDef: 'Punjab', city: 'Ludhiana', date: 'dmy', wit: 'Two', idl: 'PAN, CIN or other ID (optional)',
      law: function (r) { return 'the laws of India'; },
      courts: function (x) { return L('the courts at {0}, {1}', [x.v('courts', 'city'), x.v('region', 'state')]); },
      arb: function (x) { return L('by a sole arbitrator appointed by both Parties together (or, if they cannot agree within 30 days, as provided in the Arbitration and Conciliation Act, 1996). The arbitration will be held under that Act; the seat and venue will be {0} and the language English. The award will be final and binding on the Parties', [x.v('courts', 'city')]); },
      org: { Company: 'a company incorporated under the Companies Act, 2013', LLP: 'a limited liability partnership registered under the Limited Liability Partnership Act, 2008', 'Partnership firm': 'a partnership firm under the Indian Partnership Act, 1932' },
      tax: 'GST', data: 'the Digital Personal Data Protection Act, 2023 and the rules made under it', esign: 'Information Technology Act, 2000',
      samples: ['Rajesh Kumar Sharma', '12, Green Park, Model Town, Ludhiana, Punjab 141002', 'Ananya Verma', 'House 45, Sector 21-C, Chandigarh 160022'] },
    UK: { name: 'United Kingdom', cur: 'GBP', region: 'Law of (England and Wales, Scotland or Northern Ireland)', regionDef: 'England and Wales', city: 'Manchester', date: 'dmy', wit: 'None', idl: 'Company number (optional)',
      law: function (r) { return L('the law of {0}', [r]); },
      courts: function (x) { return L('the courts of {0}', [x.v('region', 'England and Wales')]); },
      arb: function (x) { return L('by a sole arbitrator agreed by both Parties or, failing agreement within 30 days, appointed by the President of the Chartered Institute of Arbitrators. The arbitration will be conducted under the Arbitration Act 1996 (or, where Scots law governs this Agreement, the Arbitration (Scotland) Act 2010); the seat will be {0} and the language English', [x.v('courts', 'city')]); },
      org: { Company: 'a company registered in {region}', LLP: 'a limited liability partnership registered in {region}', 'Partnership firm': 'a partnership' },
      tax: 'VAT', data: 'the UK GDPR and the Data Protection Act 2018', esign: 'Electronic Communications Act 2000',
      samples: ['Oliver Bennett', '18 Albert Road, Didsbury, Manchester M20 2AB', 'Amelia Clarke', 'Flat 4, 27 Victoria Street, Leeds LS1 6DD'] },
    US: { name: 'United States', cur: 'USD', region: 'State whose law applies', regionDef: 'New York', city: 'New York', date: 'mdy', wit: 'None', idl: 'EIN or state file number (optional)',
      law: function (r) { return L('the laws of the State of {0}, without regard to its conflict-of-laws rules', [r]); },
      courts: function (x) { return L('the state and federal courts located in {0}, {1}', [x.v('courts', 'county or city'), x.v('region', 'State')]); },
      arb: function (x) { return L('by a single arbitrator administered by the American Arbitration Association under its Commercial Arbitration Rules (or, for a consumer, its Consumer Arbitration Rules). The place of arbitration will be {0}, {1}. Judgment on the award may be entered in any court having jurisdiction', [x.v('courts', 'city'), x.v('region', 'State')]); },
      org: { Company: 'a {region} corporation', LLP: 'a {region} limited liability partnership', 'Partnership firm': 'a {region} general partnership', LLC: 'a {region} limited liability company' },
      tax: 'sales tax', data: 'all applicable federal and state privacy and data security laws', esign: 'federal E-SIGN Act and the Uniform Electronic Transactions Act as adopted in the State',
      samples: ['Michael Carter', '415 Maple Avenue, Apt 3B, Brooklyn, NY 11215', 'Sarah Lopez', '2208 Riverside Drive, Austin, TX 78741'] },
    AE: { name: 'United Arab Emirates', cur: 'AED', region: 'Emirate', regionDef: 'Dubai', city: 'Dubai', date: 'dmy', wit: 'Two', idl: 'Trade licence or Emirates ID no. (optional)',
      law: function (r) { return L('the federal laws of the United Arab Emirates and, where applicable, the laws of the Emirate of {0}', [r]); },
      courts: function (x) { return L('the courts of the Emirate of {0}', [x.v('region', 'Emirate')]); },
      arb: function (x) { return L('under the Arbitration Rules of the Dubai International Arbitration Centre (DIAC) by a sole arbitrator appointed under those Rules, and in line with Federal Law No. 6 of 2018 on Arbitration. The seat of arbitration will be {0} and the language English', [x.v('courts', 'city')]); },
      org: { Company: 'a limited liability company licensed in the Emirate of {region}', LLP: 'a company licensed in the Emirate of {region}', 'Partnership firm': 'a partnership licensed in the Emirate of {region}' },
      tax: 'VAT', data: 'Federal Decree-Law No. 45 of 2021 on the Protection of Personal Data and any other applicable data law', esign: 'Federal Decree-Law No. 46 of 2021 on Electronic Transactions and Trust Services',
      samples: ['Ahmed Al Mansoori', 'Villa 12, Street 7, Al Barsha 2, Dubai', 'Fatima Khan', 'Apartment 1104, Marina Heights Tower, Dubai Marina, Dubai'] },
    INT: { name: 'Another country', cur: 'USD', region: 'Country (and state) whose law applies', regionDef: '', city: '', date: 'dmy', wit: 'Two', idl: 'Registration or ID number (optional)',
      law: function (r) { return L('the laws of {0}', [r || '[country]']); },
      courts: function (x) { return L('the courts at {0}, {1}', [x.v('courts', 'city'), x.v('region', 'country')]); },
      arb: function (x) { return L('by a sole arbitrator under the UNCITRAL Arbitration Rules, appointed by both Parties together or, failing agreement within 30 days, by the appointing authority under those Rules. The seat of arbitration will be {0} and the language English', [x.v('courts', 'city')]); },
      org: { Company: 'a company incorporated in {region}', LLP: 'a limited liability partnership registered in {region}', 'Partnership firm': 'a partnership' },
      tax: 'tax', data: 'all data protection laws that apply to it', esign: 'law on electronic signatures that applies',
      samples: ['Daniel Okafor', '14 Admiralty Way, Lekki Phase 1, Lagos, Nigeria', 'Maria Santos', '22 Rizal Avenue, Makati City, Metro Manila, Philippines'] }
  };
  var KINDS = ['Individual', 'Company', 'LLP', 'LLC', 'Partnership firm', 'Sole proprietorship', 'Trust, society or other organisation'];

  /* ---------------- questions (one library; each agreement type lists the keys it needs, with an optional own label) */
  var F = {
    property: ['Property address', 'wide', 'Flat 302, Green Avenue Apartments, Sector 21, Ludhiana, Punjab 141001'],
    propdesc: ['Property details: rooms, area, parking, fittings (one per line)', 'area', '2 bedrooms, hall, kitchen and 2 bathrooms, about 1,050 sq ft\nOne covered car parking space\n4 ceiling fans, 8 LED lights, 2 geysers\nModular kitchen with chimney and hob'],
    start: ['Start date', 'date', today()],
    term: ['Term (months)', 'num', '11'],
    rent: ['Monthly rent', 'num', '18000'],
    payday: ['Payment due by (day of each month)', 'num', '5'],
    deposit: ['Security deposit', 'num', '54000'],
    escalation: ['Increase each year or on renewal (%)', 'num', '5'],
    maint: ['Society or building maintenance paid by', 'who', 'B', ['B', 'A']],
    utilities: ['Electricity, water and gas bills paid by', 'who', 'B', ['B', 'A', 'S']],
    lockin: ['Lock-in period (months, 0 for none)', 'num', '6'],
    notice: ['Notice period (days)', 'num', '30'],
    occupants: ['Maximum number of people living there', 'num', '4'],
    use: ['Permitted use', 'wide', 'residence of the Tenant and family'],
    bizuse: ['Business to be carried on', 'wide', 'retail sale of ready-made garments'],
    fitout: ['Rent-free fit-out period (days)', 'num', '15'],
    room: ['Room or bed given', 'wide', 'Bed 2 in Room 3 (twin sharing), with a cupboard and study table'],
    meals: ['Meals included', 'sel', 'Breakfast and dinner', ['Breakfast and dinner', 'All three meals', 'No meals']],
    rules: ['House rules (one per line)', 'area', 'Main gate closes at 10:30 pm; inform the Owner if you will be late\nNo smoking, alcohol or drugs on the premises\nVisitors only in the common area between 9 am and 8 pm; no overnight guests\nKeep the room and shared areas clean; switch off lights, fans and AC when leaving\nNo cooking in the rooms'],
    price: ['Total price', 'num', '250000'],
    advance: ['Advance or token paid on signing', 'num', '50000'],
    balancedate: ['Balance to be paid by', 'date', ''],
    goods: ['Goods (one per line: description | quantity | unit price)', 'area', 'Steel office chair, model OC-22 | 20 | 3200\nOffice table 4 ft x 2 ft, laminated | 10 | 6500'],
    delivery: ['Delivery date', 'date', ''],
    deliveryplace: ['Delivery address', 'wide', ''],
    warranty: ['Warranty (months)', 'num', '12'],
    leadtime: ['Delivery time after each order (days)', 'num', '7'],
    products: ['Products', 'wide', 'packaged spices and masala blends sold under the "Rasoi Gold" brand'],
    vmake: ['Vehicle make and model', 'text', 'Maruti Suzuki Swift VXI (petrol)'],
    vreg: ['Registration number', 'text', 'PB10 AB 1234'],
    vyear: ['Year of manufacture', 'text', '2021'],
    vchassis: ['Chassis number', 'text', ''],
    vengine: ['Engine number', 'text', ''],
    vkm: ['Odometer reading (km)', 'num', '38500'],
    services: ['Services or scope of work (one per line)', 'area', 'Monthly bookkeeping in the Client\'s accounting software\nPayroll processing for up to 20 staff\nQuarterly management report on sales, expenses and dues'],
    milestones: ['Payment schedule (one per line: stage | when due | amount)', 'area', 'Advance on signing | On signing | 30000\nDesign approved | 15 days after signing | 30000\nWebsite live | On launch | 40000'],
    fee: ['Fee', 'num', '25000'],
    feebasis: ['Fee is charged', 'sel', 'per month', ['per month', 'for the whole work (fixed fee)', 'per hour', 'per day', 'per visit']],
    payterms: ['Invoices to be paid within (days)', 'num', '15'],
    tax: ['Amounts are', 'sel', 'plus tax', ['plus tax', 'inclusive of tax', 'no tax applies']],
    revisions: ['Rounds of changes included', 'num', '2'],
    position: ['Job title', 'text', 'Accounts Executive'],
    reportto: ['Reports to', 'text', 'Finance Manager'],
    workplace: ['Place of work', 'wide', 'the Employer\'s office at Ferozepur Road, Ludhiana, and any other place the Employer reasonably asks'],
    hours: ['Working hours', 'wide', '9:30 am to 6:00 pm, Monday to Saturday, with a one-hour lunch break'],
    salary: ['Monthly salary (gross)', 'num', '32000'],
    probation: ['Probation (months, 0 for none)', 'num', '6'],
    leave: ['Paid leave (days per year)', 'num', '18'],
    stipend: ['Monthly stipend (0 if unpaid)', 'num', '10000'],
    mentor: ['Mentor or supervisor', 'text', 'Ms. Kavita Rao, Senior Designer'],
    purpose: ['Purpose', 'wide', 'discussing a possible partnership to build and sell a mobile app for clinic appointments'],
    confyears: ['Confidentiality lasts after the end (years)', 'num', '3'],
    business: ['Name of the firm or business', 'text', 'Sharma & Verma Associates'],
    nature: ['Nature of business', 'wide', 'trading in electrical goods and fittings'],
    bizplace: ['Principal place of business', 'wide', 'Shop 14, Ghumar Mandi, Ludhiana, Punjab 141001'],
    capA: ['Capital or contribution of Party 1', 'num', '500000'],
    capB: ['Capital or contribution of Party 2', 'num', '500000'],
    shareA: ['Share of Party 1 (%) (Party 2 gets the rest)', 'num', '50'],
    remun: ['Monthly remuneration for each working partner (0 for none)', 'num', '0'],
    company: ['Company or LLP name (proposed if not yet formed)', 'text', 'Clinicly Technologies Private Limited'],
    roleA: ['Role of Party 1', 'wide', 'Chief Executive: product, sales, fundraising and finance'],
    roleB: ['Role of Party 2', 'wide', 'Chief Technology Officer: software development, hosting and security'],
    vest: ['Vesting period (years)', 'num', '4'],
    cliff: ['Cliff (months before any shares vest)', 'num', '12'],
    respA: ['What Party 1 will do (one per line)', 'area', 'Provide trainers and course material\nIssue certificates to students who complete the course'],
    respB: ['What Party 2 will do (one per line)', 'area', 'Provide a classroom with projector for 2 days a week\nSelect and enrol students'],
    costs: ['Costs', 'sel', 'Each Party bears its own costs', ['Each Party bears its own costs', 'Shared equally', 'As set out in the Schedule']],
    principal: ['Loan amount', 'num', '200000'],
    rate: ['Interest rate (% per year, 0 for none)', 'num', '9'],
    repay: ['Repayment', 'sel', 'Equal monthly instalments', ['Equal monthly instalments', 'In one amount on the due date', 'On demand']],
    instal: ['Number of monthly instalments', 'num', '12'],
    duedate: ['Final due date (one-amount repayment)', 'date', ''],
    paymode: ['Money paid by', 'sel', 'bank transfer', ['bank transfer', 'cheque', 'UPI', 'demand draft']],
    security: ['Security or guarantee (optional)', 'wide', ''],
    debtor: ['Name of the person or business whose debt is guaranteed', 'text', 'Verma Electricals'],
    obligation: ['What is guaranteed', 'wide', 'all amounts payable by the Debtor for goods supplied by the Creditor on credit'],
    poakind: ['Kind of power', 'sel', 'Special (for named acts)', ['Special (for named acts)', 'General']],
    matter: ['Property or matter covered', 'wide', 'my residential flat at Flat 302, Green Avenue Apartments, Sector 21, Ludhiana'],
    powers: ['Powers given (one per line)', 'area', 'To let out the flat, sign rent or leave and licence agreements and get them registered\nTo collect rent and the security deposit and deposit them only in my bank account\nTo pay property tax, society maintenance, electricity and water bills\nTo appear before the municipal corporation, the society and the police for tenant verification\nTo sign letters, applications and receipts needed for the above'],
    validity: ['Valid for (months, 0 until cancelled)', 'num', '24'],
    territory: ['Territory', 'wide', 'the districts of Ludhiana, Jalandhar and Moga in Punjab'],
    exclusive: ['Exclusivity', 'sel', 'Exclusive', ['Exclusive', 'Non-exclusive']],
    target: ['Minimum purchases per year', 'num', '1200000'],
    margin: ['Distributor margin or discount on list price (%)', 'num', '12'],
    commission: ['Commission (% of paid sales)', 'num', '5'],
    brand: ['Brand or business name', 'text', 'Chai Junction'],
    ffee: ['Initial franchise fee', 'num', '300000'],
    royalty: ['Royalty (% of monthly gross sales)', 'num', '6'],
    mktfee: ['Marketing fund (% of monthly gross sales)', 'num', '2'],
    location: ['Outlet address', 'wide', 'Shop 3, Ground Floor, City Square Mall, Jalandhar'],
    equip: ['Equipment or items (one per line)', 'area', 'Split air conditioner, 1.5 ton, Voltas, serial VX2231 (Reception)\nSplit air conditioner, 1.5 ton, Daikin, serial DK7740 (Cabin 1)\nSplit air conditioner, 2 ton, LG, serial LG1180 (Conference room)'],
    visits: ['Preventive maintenance visits per year', 'num', '4'],
    response: ['Response time to a breakdown call (hours)', 'num', '24'],
    coverage: ['Cover', 'sel', 'Non-comprehensive (labour only)', ['Non-comprehensive (labour only)', 'Comprehensive (parts included)']],
    hirerate: ['Hire charge', 'num', '1500'],
    hireper: ['Hire charge is', 'sel', 'per day', ['per day', 'per week', 'per month']],
    hiredays: ['Hire period (days)', 'num', '10'],
    eventname: ['Event', 'text', 'Wedding reception of Aarav and Diya'],
    eventdate: ['Event date', 'date', ''],
    venue: ['Venue', 'wide', 'Royal Palms Banquet Hall, Ferozepur Road, Ludhiana'],
    guests: ['Expected number of guests', 'num', '300'],
    project: ['Project', 'wide', 'a responsive website with an online booking system for the Client\'s clinic'],
    platform: ['Technology or platform', 'text', 'WordPress with a custom booking plugin'],
    support: ['Free bug-fix period after launch (months)', 'num', '3'],
    hosting: ['Hosting and domain', 'sel', 'In the Client\'s name and account', ['In the Client\'s name and account', 'Arranged by the Developer at the Client\'s cost']],
    work: ['Work being licensed', 'wide', 'the set of 40 product photographs titled "Monsoon Collection 2026"'],
    media: ['Where it may be used', 'wide', 'the Licensee\'s website, social media accounts and printed catalogues'],
    ltype: ['Licence', 'sel', 'Non-exclusive', ['Non-exclusive', 'Exclusive']],
    lterm: ['Licence period (months)', 'num', '24'],
    creditline: ['Credit line', 'text', 'Photo: Ananya Verma'],
    platforms: ['Platforms', 'text', 'Instagram and YouTube'],
    deliverables: ['Content to be posted (one per line)', 'area', '1 Instagram reel (30 to 60 seconds) featuring the product\n3 Instagram stories with a swipe-up or link sticker\n1 YouTube video (at least 5 minutes) with a dedicated review segment'],
    usage: ['Brand may reuse the content for (months)', 'num', '6'],
    excl: ['No competing brand content for (days)', 'num', '30'],
    mainproject: ['Main project or end client', 'wide', 'interior fit-out of the office of Northwind Logistics at Gurugram'],
    duties: ['Duties (one per line)', 'area', 'Cleaning the house and washing utensils\nCooking breakfast and lunch for the family of four\nLooking after the house when the family is away during the day'],
    weeklyoff: ['Weekly day off', 'text', 'Sunday'],
    livein: ['Live-in', 'sel', 'No (comes daily)', ['No (comes daily)', 'Yes (stays in the house)']],
    dispute: ['What the dispute was about', 'area', 'Unpaid invoices for interior work done at the First Party\'s office between January and March 2026, and the First Party\'s complaint about delay and quality'],
    setamt: ['Settlement amount', 'num', '150000'],
    setdate: ['Settlement amount to be paid by', 'date', ''],
    payer: ['Who pays', 'pay', 'AB', ['AB', 'BA']],
    casedetails: ['Pending case or complaint (if any)', 'wide', ''],
    origname: ['Agreement being ended', 'wide', 'Service Agreement for monthly bookkeeping'],
    origdate: ['Date of that agreement', 'date', ''],
    finalamt: ['Final amount payable (0 if none)', 'num', '0'],
    handover: ['Things to hand back or finish (one per line)', 'area', 'Return of all files, passwords and records of the Client\nFinal invoice for work done up to the end date'],
    titledocs: ['Title documents (one per line)', 'area', 'Registered sale deed in the Seller\'s favour\nLatest property tax receipt\nNo-dues certificate from the housing society\nEncumbrance certificate for the last 13 years'],
    courts: ['Courts at (city)', 'text', 'Ludhiana']
  };

  /* ---------------- context: how clause text reads answers */
  function ctx(T, S) {
    var V = S.values, c = COUNTRY[S.cc], cur = CUR[S.cur] || CUR.USD;
    var x = { T: T, S: S, c: c, cc: S.cc, cur: cur, A: T.A, B: T.B };
    x.raw = function (k) { return V[k]; };
    x.v = function (k, ph) { return blank(V[k]) ? '[' + (ph || (F[k] ? F[k][0] : k)).toLowerCase() + ']' : String(V[k]).trim(); };
    x.n = function (k) { return num(V[k]); };
    x.has = function (k) { return !blank(V[k]) && String(V[k]).trim() !== '0'; };
    x.ms = function (n) {   /* short money */
      if (typeof n === 'string') { if (blank(V[n])) return L('{0} [amount]', [cur[0]]); n = num(V[n]); }
      var f = new Intl.NumberFormat(cur[3], { minimumFractionDigits: n % 1 ? 2 : 0, maximumFractionDigits: 2 });
      return cur[0] + ' ' + f.format(n);
    };
    x.m = function (n) {   /* money with words */
      if (typeof n === 'string') { if (blank(V[n])) return L('{0} [amount] ([amount in words])', [cur[0]]); n = num(V[n]); }
      var neg = n < 0; n = Math.abs(n);
      var whole = Math.floor(n + 1e-9), minor = Math.round((n - whole) * 100);
      if (LANG) return (neg ? '-' : '') + x.ms(n) + (S.cur === 'INR' ? ' (' + langMoney(whole, minor) + ')' : '');
      var w = S.cur === 'INR' || S.cur === 'NPR' || S.cur === 'BDT' ? wordsIndian(whole) : wordsIntl(whole);
      return L('{0}{1} ({2}{3} {4}{5} only)', [neg ? 'minus ' : '', x.ms(n), neg ? 'minus ' : '', cur[1], w, minor ? L(' and {0} {1}', [w2(minor), cur[2]]) : '']);
    };
    x.date = function (s) {
      var d = parse(s); if (!d) return '[date]';
      if (LANG) return LANG.date ? LANG.date(d) : d.getDate() + ' ' + LANG.months[d.getMonth()] + ' ' + d.getFullYear();
      return c.date === 'mdy' ? MONTHS[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear() : d.getDate() + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear();
    };
    x.d = function (k) { return blank(V[k]) ? '[' + (F[k] ? F[k][0] : k).toLowerCase() + ']' : x.date(V[k]); };
    x.end = function (k) { var e = addMonths(V.start, num(V[k || 'term']), true); return e ? x.date(e) : '[end date]'; };
    x.lines = function (k) { return String(V[k] || '').split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean); };
    x.who = function (k) { var w = V[k] || (F[k] && F[k][2]); return w === 'A' ? L('the {0}', [x.A]) : w === 'S' ? L('the {0} and the {1} in equal shares', [x.A, x.B]) : L('the {0}', [x.B]); };
    x.days = function (k, ph) { return blank(V[k]) ? L('[{0}] days', [ph || 'number of']) : L(num(V[k]) === 1 ? '{0} day' : '{0} days', [num(V[k])]); };
    x.mo = function (k) { return blank(V[k]) ? L('[number of] months') : L(num(V[k]) === 1 ? '{0} month' : '{0} months', [num(V[k])]); };
    x.pct = function (k) { return blank(V[k]) ? '[ ]%' : num(V[k]) + '%'; };
    x.taxed = function () { return V.tax === 'inclusive of tax' ? L(' This amount includes all {0} and other taxes.', [c.tax]) : V.tax === 'no tax applies' ? '' : L(' {0} and any other tax on this amount, where it applies, will be added and paid against a proper tax invoice.', [c.tax.charAt(0).toUpperCase() + c.tax.slice(1)]); };
    x.on = function (k) { return !!(S.opts[T.id] || {})[k]; };
    x.region = function () { return x.v('region', c.region); };
    x.law = function () { return c.law(x.v('region', c.region)); };
    x.party = function (p) { return x.v(p + 'name', L('name of Party {0}', [p === 'a' ? '1' : '2'])); };
    x.inIN = S.cc === 'IN'; x.inUK = S.cc === 'UK'; x.inUS = S.cc === 'US'; x.inAE = S.cc === 'AE';
    return x;
  }
  function sec(h, items) { return { h: h, items: items }; }
  function sub(t, list) { return { t: t, sub: list }; }
  function tbl(head, rows, widths) { return { table: { head: head, rows: rows } }; }
  function emi(P, ratePct, n) { var r = ratePct / 1200; if (!n) return 0; return r ? P * r * Math.pow(1 + r, n) / (Math.pow(1 + r, n) - 1) : P / n; }
  function r2(n) { return Math.round((n + 1e-9) * 100) / 100; }

  F.evfee = ['Total price', 'num', '350000'];
  F.loanpurpose = ['Purpose of the loan (optional)', 'wide', 'buying stock for the Borrower\'s shop'];
  F.latepct = ['Late payment interest (% per month)', 'num', '1.5'];
  F.ncmonths = ['Non-compete lasts after the end (months)', 'num', '6'];
  F.ncfield = ['Competing business not allowed', 'wide', 'the same kind of business as this Agreement'];
  F.disp = ['If talking fails, disputes go to', 'sel', 'Arbitration', ['Arbitration', 'Courts']];
  F.liabcap = ['Liability cap', 'sel', 'Fees paid in the last 12 months', ['Fees paid in the last 12 months', 'Total value of this Agreement', 'Fixed amount']];
  F.capamt = ['Fixed cap amount (if chosen)', 'num', ''];

  /* ---------------- optional clauses: key → [label, short note, extra question keys, text] */
  var OPT = {
    conf: ['Confidentiality', 'Both sides keep each other\'s business information secret.', ['confyears'], function (x) {
      return sec('Confidentiality', [
        'Each Party will keep confidential all information about the other Party\'s business, customers, prices, finances, methods, technology and personal data that it receives because of this Agreement ("Confidential Information"), and will use it only to perform this Agreement.',
        'A Party may share Confidential Information only with its own employees, advisers and subcontractors who need it for this Agreement and are bound by similar duties of confidence. The Party is responsible for anything they do with it.',
        'This clause does not apply to information that is or becomes public through no fault of the receiving Party, that it already knew lawfully, that it receives lawfully from someone else free of any duty of confidence, or that it develops independently.',
        'A Party may disclose Confidential Information if required by law, a court or a regulator, after telling the other Party first where the law allows, so that it can seek protection.',
        L('These duties continue during this Agreement and for {0} after it ends, and for trade secrets for as long as they remain secret.', [x.has('confyears') ? L('{0} years', [x.n('confyears')]) : '[number of] years'])]);
    }],
    noncompete: ['Non-compete (where the law allows)', 'Limits competing business. India: only while the agreement runs; elsewhere: reasonable and lawful only.', ['ncmonths', 'ncfield'], function (x) {
      var who = x.T.nc === 'A' ? x.A : x.B, prot = who === x.A ? x.B : x.A, f = x.v('ncfield', 'competing business');
      if (x.inIN) return sec('Non-compete during this Agreement', [
        L('While this Agreement is in force, the {0} will not, directly or through anyone else, carry on, work for or hold an interest in {1} in competition with the {2}, without the {3}\'s written consent.', [who, f, prot, prot]),
        'The Parties understand that under section 27 of the Indian Contract Act, 1872 a restraint of trade that continues after an agreement ends is generally void (except in the limited cases the law allows, such as the sale of goodwill or among partners). This clause therefore applies only during the term of this Agreement. The duties of confidentiality and non-solicitation continue as stated in their own clauses.']);
      return sec('Non-compete', [
        L('During this Agreement and for {0} after it ends, the {1} will not, directly or through anyone else, carry on, work for or hold an interest in {2}{3} in competition with the {4}, without the {5}\'s written consent.', [x.mo('ncmonths'), who, f, x.T.terr !== false && x.raw('territory') ? L(' within {0}', [x.v('territory')]) : '', prot, prot]),
        L('The Parties agree this restriction is no wider than reasonably needed to protect the {0}\'s legitimate business interests, confidential information and goodwill. It applies only to the extent permitted by the law that governs this Agreement; if a court finds it too wide, it will apply with the smallest change needed to make it enforceable.{1}', [prot, x.inUS ? ' It does not apply in any State whose law prohibits such a restriction for this kind of relationship.' : ''])]);
    }],
    nonsolicit: ['Non-solicitation', 'No poaching each other\'s staff or customers for 12 months.', [], function (x) {
      return sec('Non-solicitation', [
        'During this Agreement and for 12 months after it ends, neither Party will, without the other\'s written consent, persuade or try to persuade any employee or contractor of the other Party who was involved with this Agreement to leave that Party. A general job advertisement open to the public is not a breach.',
        L('During the same period the {0} will not use Confidential Information or its contacts made through this Agreement to take away, or try to take away, the {1}\'s customers or suppliers.', [x.B, x.A]),
        'This clause applies to the extent permitted by the law that governs this Agreement.']);
    }],
    termination: ['Ending for breach or insolvency', 'Either side may end the agreement if the other breaks it and does not fix it, or goes bankrupt.', [], function (x) {
      var items = [];
      if (!x.T.ownEnd) items.push(L('Either Party may end this Agreement for any reason by giving the other at least {0} of written notice.', [x.days('notice')]));
      items.push(sub('Either Party may end this Agreement at once by written notice if the other Party:', [
        'commits a serious breach of this Agreement and, if the breach can be put right, does not put it right within 15 days after receiving written notice describing it;',
        'repeatedly breaches this Agreement in a way that shows it does not intend to keep to it; or',
        'becomes insolvent or bankrupt, has a receiver, administrator or liquidator appointed, makes an arrangement with its creditors, or stops carrying on business.']));
      items.push('When this Agreement ends, for any reason: each Party will return (or, if asked, destroy) the other\'s property, documents and Confidential Information; amounts properly due up to the end date must be paid; and the clauses meant to continue (including payment, confidentiality, ownership of work, indemnity, limit of liability and disputes) continue in force.');
      return sec(x.T.ownEnd ? 'Ending for breach or insolvency' : 'Ending this Agreement', items);
    }],
    latefee: ['Late payment interest', 'Interest on payments made after their due date.', ['latepct'], function (x) {
      var it = [L('If any amount payable under this Agreement is not paid by its due date, the Party that owes it will pay simple interest on the unpaid amount at {0} per month from the due date until it is paid, in addition to the amount itself.', [x.pct('latepct')]),
        'The Parties agree this rate is a reasonable estimate of the loss that late payment causes and is not a penalty. Charging interest does not stop the other Party from using its other rights for late payment.'];
      if (x.inIN) it.push('Where the Party to be paid is a micro or small enterprise registered under the Micro, Small and Medium Enterprises Development Act, 2006, the interest under that Act applies instead if it is higher.');
      if (x.inUK) it.push('Where the Late Payment of Commercial Debts (Interest) Act 1998 applies, the Parties agree this is a substantial remedy for late payment for the purposes of that Act.');
      return sec('Late payment', it);
    }],
    dispute: ['Disputes: talk first, then arbitration or courts', 'A 30-day talking period, then arbitration or the courts you choose.', ['disp'], function (x) {
      var arb = (x.raw('disp') || 'Arbitration') === 'Arbitration';
      return sec('Disputes', [
        'If a dispute arises out of or in connection with this Agreement, either Party may give the other a written notice describing it. The Parties (or their senior representatives) will meet, in person or online, within 14 days of the notice and try in good faith to settle it.',
        arb ? L('If the dispute is not settled within 30 days of the notice, it will be referred to and finally resolved by arbitration {0}.', [x.c.arb(x)]) : L('If the dispute is not settled within 30 days of the notice, either Party may take it to {0}, which will have exclusive jurisdiction.', [x.c.courts(x)]),
        'Each Party will keep performing its undisputed obligations while the dispute is being resolved. Nothing in this clause stops a Party from asking a court for urgent interim relief, such as an injunction.']);
    }],
    force: ['Events beyond control (force majeure)', 'No blame for delays caused by floods, epidemics, government orders and similar events.', [], function (x) {
      return sec('Events beyond control (force majeure)', [
        'Neither Party is responsible for a delay or failure to perform this Agreement (other than paying money that is due) caused by an event beyond its reasonable control, such as flood, earthquake, storm, fire, epidemic, war, riot, terrorism, a government order or lockdown, or a general failure of power or communications networks ("Force Majeure Event").',
        'The affected Party will tell the other promptly, explain the likely effect and how long it may last, and do what it reasonably can to limit the effect and resume performance.',
        L('If a Force Majeure Event prevents performance for more than 60 days in a row, either Party may end this Agreement by written notice. Amounts due for work done, goods supplied or the period of use before the event will still be paid{0}.', [x.T.fmRefund ? ', and advance payments for anything not supplied will be refunded within 30 days' : ''])]);
    }],
    ip: ['Ownership of work (intellectual property)', 'Who owns designs, writing, code or other work made under the agreement.', [], function (x) {
      var kind = x.T.ip || 'client';
      if (kind === 'owner') return sec('Intellectual property', [
        'Each Party keeps the intellectual property rights (such as copyright, trade marks, designs and know-how) that it owned before this Agreement or develops on its own. Nothing in this Agreement transfers ownership of them.',
        'A Party may use the other Party\'s name, logo or trade marks only as this Agreement allows or with the other\'s prior written approval, and must stop when this Agreement ends.']);
      if (kind === 'employer') return sec('Work created in employment', [
        'All work, writing, designs, software, inventions, data and other material that the Employee creates, alone or with others, in the course of employment or using the Employer\'s time, equipment or information ("Work Product") belong to the Employer from creation. To the extent they do not belong to the Employer automatically, the Employee assigns all rights in them to the Employer, worldwide and for the full term of protection.',
        'The Employee will tell the Employer about any Work Product and, during and after employment, sign any document the Employer reasonably needs to record or protect its rights, at the Employer\'s cost.',
        'This clause does not cover anything the Employee creates entirely in their own time, without the Employer\'s equipment or information, and which does not relate to the Employer\'s business.']);
      var it = [L('On full payment of the fees for them, all copyright and other intellectual property rights in the work created specially for the {0} under this Agreement ("Deliverables") will belong to the {1}, and the {2} hereby assigns those rights to the {3}{4}. Until full payment, the {5} may use the Deliverables only to review them.', [x.A, x.A, x.B, x.A, x.inIN ? ' worldwide and for the full term of copyright (this assignment is made under sections 18 and 19 of the Copyright Act, 1957)' : ' worldwide and for the full term of protection', x.A]),
        L('The {0} keeps ownership of its existing materials, know-how, tools, code libraries and general templates used in the Deliverables ("Background Materials") and gives the {1} a permanent, non-exclusive, royalty-free licence to use them as part of the Deliverables.', [x.B, x.A]),
        L('Materials owned by others (such as stock images, fonts, plugins or open-source software) are supplied under their own licence terms, which the {0} will list on delivery.', [x.B]),
        L('The {0} confirms that, to the best of its knowledge, the Deliverables (other than materials supplied by the {1}) do not infringe anyone else\'s rights.', [x.B, x.A])];
      if (x.T.portfolio !== false) it.push(L('Unless the {0} objects in writing, the {1} may show the Deliverables in its portfolio once they are public.', [x.A, x.B]));
      return sec('Ownership of work (intellectual property)', it);
    }],
    indemnity: ['Indemnity', 'Each side covers losses caused by its own breach, negligence or breach of law.', [], function (x) {
      var one = x.T.indem, who = one === 'A' ? x.A : one === 'B' ? x.B : null;
      var head = who ? L('The {0} will indemnify the {1} (and keep it indemnified)', [who, who === x.A ? x.B : x.A]) : 'Each Party (the "Indemnifying Party") will indemnify the other Party (and keep it indemnified)';
      return sec('Indemnity', [
        sub(L('{0} against losses, damages, fines, reasonable legal costs and claims by third parties arising from:', [head]), [
          L('{0} breach of this Agreement;', [who ? 'its' : 'the Indemnifying Party\'s']), L('{0} negligence, fraud or wilful misconduct, or that of its staff or subcontractors; and', [who ? 'its' : 'the Indemnifying Party\'s']),
          L('{0} breach of any law{1}.', [who ? 'its' : 'the Indemnifying Party\'s', x.T.ipIndem ? ', or any claim that work it supplied infringes someone else\'s intellectual property rights' : ''])]),
        'The Party claiming under this clause must tell the other promptly about any claim, take reasonable steps to limit its loss, not settle a third-party claim without the other\'s consent (which will not be unreasonably refused), and allow the other Party to take over the defence of a third-party claim at its own cost.']);
    }],
    liability: ['Limit of liability', 'Caps the money either side can claim and excludes indirect losses.', ['liabcap', 'capamt'], function (x) {
      var c = x.raw('liabcap') || 'Fees paid in the last 12 months';
      var cap = c === 'Fixed amount' ? x.m('capamt') : c === 'Total value of this Agreement' ? 'the total amount paid and payable under this Agreement' : 'the total amount paid and payable under this Agreement in the 12 months before the event that caused the claim';
      return sec('Limit of liability', [
        'Neither Party is liable to the other for loss of profit, revenue, business, contracts or goodwill, or for any indirect or consequential loss, even if it was foreseeable.',
        L('Each Party\'s total liability to the other arising out of or in connection with this Agreement, whether in contract, tort (including negligence), breach of statutory duty or otherwise, will not exceed {0}.', [cap]),
        'Nothing in this Agreement limits or excludes liability for fraud, for death or personal injury caused by negligence, for paying amounts due under this Agreement, for breach of confidentiality, under any indemnity for third-party claims, or for anything else that cannot be limited or excluded by law.']);
    }],
    data: ['Personal data', 'Handling customers\' or staff personal data under the data law of the country.', [], function (x) {
      return sec('Personal data', [
        L('Each Party will handle personal data it receives under this Agreement in line with {0}, use it only for this Agreement, keep it secure with reasonable technical and organisational measures, and tell the other Party without undue delay of any data breach affecting it.', [x.c.data]),
        L('Where the {0} processes personal data on behalf of the {1}, it will act only on the {2}\'s written instructions, make sure its staff are bound by confidentiality, not pass the data to another processor or outside the country without the {3}\'s consent, help the {4} respond to requests from individuals, and return or delete the data when this Agreement ends.', [x.B, x.A, x.A, x.A, x.A])]);
    }]
  };
  var OPT_ORDER = ['conf', 'noncompete', 'nonsolicit', 'ip', 'indemnity', 'liability', 'data', 'latefee', 'termination', 'force', 'dispute'];

  function general(x) {
    var T = x.T, it = [];
    it.push(L('Notices: notices under this Agreement must be in writing and delivered by hand, sent by {0}, or emailed to the addresses given above (or a new address a Party gives by notice). An emailed notice counts as received when sent, unless the sender receives an error message.', [x.inIN ? 'registered post or speed post' : x.inUK ? 'recorded delivery' : 'courier or registered mail']));
    if (x.inIN && !T.noStamp) it.push(L('Stamp duty: the stamp duty and any registration charges on this Agreement will be paid by {0}.', [x.who('stampby')]));
    it.push('Whole agreement: this Agreement (with its Schedules) is the whole agreement between the Parties about its subject and replaces all earlier discussions, proposals and arrangements about it.');
    it.push('Changes: a change to this Agreement is valid only if it is in writing and signed by both Parties.');
    it.push('Severability: if any part of this Agreement is found invalid or unenforceable, the rest of it stays in force, and the invalid part will apply with the smallest change needed to make it valid.');
    it.push('Waiver: a delay or failure to use a right under this Agreement is not a waiver of that right.');
    if (T.assign !== false) it.push(L('Transfer: neither Party may transfer (assign) its rights or obligations under this Agreement to anyone else without the other Party\'s written consent{0}.', [T.assign ? ', ' + T.assign : '']));
    if (T.rel !== false) it.push(L('Relationship: the Parties are independent. Nothing in this Agreement makes them partners, joint venturers, agents or employer and employee{0}, and neither may bind the other.', [T.rel ? ', ' + T.rel : '']));
    it.push(L('Copies and signatures: this Agreement may be signed in two or more copies, each of which is an original{0}', [T.noEsign ? '.' : L(', and signatures exchanged electronically are valid to the extent allowed by the {0}.', [x.c.esign])]));
    it.push(L('Governing law: this Agreement is governed by and will be interpreted under {0}.{1}', [x.law(), x.on('dispute') && OPT.dispute && T.opts.dispute != null ? '' : L(' {0} will have exclusive jurisdiction over any dispute about it.', [x.c.courts(x).replace(/^the /, 'The ')])]));
    return sec('General', it);
  }

  /* ---------------- shared building blocks used by several agreement types */
  function depositSec(x, holder, payer, what) {
    return sec('Security deposit', [
      L('On signing this Agreement the {0} has paid the {1} an interest-free security deposit of {2}{3}, receipt of which the {4} acknowledges.', [payer, holder, x.m('deposit'), x.raw('paymode') ? L(' by {0}', [x.raw('paymode')]) : '', holder]),
      L('The {0} will refund the deposit within 15 days after the {1} hands back {2}, after deducting only: unpaid {3}; unpaid bills and charges for the period of occupation; and the reasonable cost of repairing damage beyond normal wear and tear. The {4} will give a written list of any deductions with bills or estimates.', [holder, payer, what, x.T.feeWord || 'rent', holder]),
      L('The {0} may not adjust the deposit against {1} for the last months without the {2}\'s written consent.', [payer, x.T.feeWord || 'rent', holder]),
      L('If the {0} does not refund the amount due within that time, the {1} will also be entitled to simple interest at 1% per month on the unpaid amount until it is refunded.', [holder, payer])]);
  }
  function billsSec(x) {
    return sec('Bills and charges', [
      L('Electricity, water, piped gas and internet charges for the period of occupation will be paid by {0}, based on actual bills and meter readings. Meter readings on the start date are recorded in the Schedule.', [x.who('utilities')]),
      L('Society or building maintenance charges will be paid by the {0}. Property tax, building insurance and any charge on the ownership of the property will be paid by the {1}.', [x.raw('maint') === 'A' ? x.A : x.B, x.A]),
      'Each Party will give the other copies of paid bills when asked. Bills for the period up to the day the property is handed back are to be paid before the deposit is refunded.']);
  }
  function repairsSec(x, own, occ) {
    return sec('Repairs and upkeep', [
      L('The {0} will keep the property clean and in good condition, and will pay for day-to-day minor repairs such as replacing bulbs, fuses, tap washers and switch plates, and clearing blockages caused by use.', [occ]),
      L('The {0} will be responsible for structural repairs, roof and wall leakage and seepage, major plumbing and electrical wiring faults, and repairs to the fittings listed in the Schedule that fail through age or normal use, and will carry them out within a reasonable time after being told.', [own]),
      L('If the {0} does not start an urgent repair within 7 days of written notice, the {1} may have it done at reasonable cost and deduct the cost, supported by bills, from the next payment due.', [own, occ]),
      L('The {0} will not make structural changes, break walls or floors, or drill into load-bearing parts without the {1}\'s written consent. Removable fittings put in by the {2} may be taken away at the end if any damage is made good.', [occ, own, occ]),
      L('The {0} or its representative may inspect the property at reasonable times after giving at least 24 hours\' notice, and at any time in an emergency.', [own])]);
  }
  function premisesSched(x, title) {
    return { h: title || 'Schedule 1: The property and fittings', items: [L('Address: {0}', [x.v('property')]), sub('Description and fittings:', x.lines('propdesc').length ? x.lines('propdesc') : ['[rooms, area, parking, fittings]']),
      tbl(['Meter / item', 'Number', 'Reading or condition at start'], [['Electricity meter', '', ''], ['Water meter', '', ''], ['Keys handed over', '', ''], ['', '', '']])] };
  }
  function goodsRows(x) {
    var rows = [], tot = 0;
    x.lines('goods').forEach(function (l, i) { var p = l.split('|').map(function (s) { return s.trim(); }); var q = num(p[1]), u = num(p[2]), a = r2(q * u); tot += a; rows.push([String(i + 1), p[0] || '', p[1] || '', p[2] ? x.ms(u) : '', a ? x.ms(a) : '']); });
    return { rows: rows, total: r2(tot) };
  }
  function msRows(x) {
    var rows = [], tot = 0;
    x.lines('milestones').forEach(function (l, i) { var p = l.split('|').map(function (s) { return s.trim(); }); var a = num(p[2]); tot += a; rows.push([String(i + 1), p[0] || '', p[1] || '', p[2] ? x.ms(a) : '']); });
    return { rows: rows, total: r2(tot) };
  }
  function feeLine(x) {
    var b = x.raw('feebasis') || 'per month';
    return (b === 'for the whole work (fixed fee)' ? L('{0} for the whole of the work.{1}', [x.m('fee'), x.taxed()]) : L('{0} {1}.{2}', [x.m('fee'), b, x.taxed()]));
  }
  function invoiceLine(x, payer, payee) {
    return L('The {0} will send an invoice {1}, and the {2} will pay it within {3} of receiving it by bank transfer to the account named in the invoice{4}.', [payee, (x.raw('feebasis') || '') === 'per month' ? 'at the end of each month' : 'when each payment falls due', payer, x.days('payterms'), x.inIN ? ', after deducting tax at source (TDS) where the law requires it and giving the TDS certificate on time' : '']);
  }

  /* ---------------- the agreement types */
  var G = ['Property and rent', 'Buying and selling', 'Services and work', 'Jobs and people', 'Confidentiality', 'Business and partners', 'Money and loans', 'Authority, settlements and endings'];
  var TYPES = [];
  function type(t) { t.opts = t.opts || {}; TYPES.push(t); }

  /* ===== Property and rent */
  type({ id: 'rent', g: 0, name: 'Residential rent agreement (lease)', kw: 'tenancy rental lease house flat apartment landlord tenant 11 month',
    title: 'RENT AGREEMENT', A: 'Landlord', B: 'Tenant', ownEnd: true,
    blurb: 'Renting a house or flat to live in: rent, deposit, bills, repairs, lock-in, notice and handing back.',
    q: ['property', 'propdesc', 'start', 'term', 'rent', 'payday', 'deposit', 'paymode', 'maint', 'utilities', 'escalation', 'lockin', 'notice', 'occupants', 'use'],
    opts: { latefee: 0, dispute: 1, force: 0, indemnity: 0, termination: 1 },
    defs: function (x) { return [['Premises', 'the property described in Schedule 1, with its fittings and fixtures'], ['Rent', 'the monthly rent set out in the Rent clause'], ['Security Deposit', 'the deposit described in the Security deposit clause'], ['Term', 'the period described in the Term clause']]; },
    rec: function (x) { return ['The Landlord owns, or has the legal right to let, the Premises.', 'The Tenant wishes to rent the Premises for residential use, and the Landlord agrees to let them, on the terms of this Agreement.']; },
    body: function (x) { return [
      sec('Letting and term', [L('The Landlord lets the Premises to the Tenant for {0}, from {1} to {2} (the "Term").', [x.mo('term'), x.d('start'), x.end()]),
        L('The Term may be renewed only by a new written agreement signed by both Parties. Unless agreed otherwise, the rent on renewal will be {0} more than the rent at the end of the Term.', [x.pct('escalation')])]),
      sec('Rent', [L('The Tenant will pay rent of {0} per month, in advance, on or before the {1} day of each month, by {2} to the account the Landlord names in writing.', [x.m('rent'), x.has('payday') ? ord(x.n('payday')) : '[ ]', x.raw('paymode') || 'bank transfer']),
        'Rent for the first month is paid on signing. Rent for a part of a month is charged for the actual days.',
        L('The Landlord will give a receipt (or a written acknowledgment by message or email) for every payment.{0}', [x.inIN ? ' Where the rent paid in a year is more than the limit under the Income-tax Act, the Tenant may deduct tax at source as the law requires and will give the Landlord the TDS certificate.' : ''])]),
      depositSec(x, 'Landlord', 'Tenant', 'vacant possession of the Premises with all keys'),
      billsSec(x),
      sec('Use of the Premises', [L('The Tenant will use the Premises only for the {0}, with no more than {1} people living there.', [x.v('use'), x.has('occupants') ? x.n('occupants') : '[ ]']),
        'The Tenant will not sublet, assign, share or part with possession of any part of the Premises, or take paying guests, without the Landlord\'s written consent.',
        'The Tenant will not carry on any illegal, immoral or commercial activity, store dangerous or inflammable goods (other than normal household gas cylinders), or cause nuisance or annoyance to neighbours, and will follow the rules of the housing society or building.',
        'Pets may be kept only with the Landlord\'s written consent, and the Tenant will be responsible for any damage they cause.']),
      repairsSec(x, 'Landlord', 'Tenant'),
      sec('Lock-in, notice and ending', [
        x.n('lockin') > 0 ? L('Neither Party may end this Agreement during the first {0} (the "Lock-in Period") except for breach. If the Tenant leaves during the Lock-in Period without the Landlord\'s consent, the Tenant will pay the rent for the rest of the Lock-in Period.', [x.mo('lockin')]) : 'There is no lock-in period.',
        L('After {0}, either Party may end this Agreement by giving the other at least {1} of written notice, or by paying (or, for the Landlord, giving up) rent for the notice period instead.', [x.n('lockin') > 0 ? 'the Lock-in Period' : 'the start date', x.days('notice')]),
        'The Landlord may end this Agreement by 15 days\' written notice if the Tenant does not pay rent for two months in a row, or seriously breaks this Agreement and does not put it right within 15 days of a written notice.']),
      sec('Handing back the Premises', ['When this Agreement ends, the Tenant will hand back vacant and peaceful possession of the Premises with all keys and the fittings listed in Schedule 1, in the same condition as at the start except for normal wear and tear. Both Parties will inspect the Premises together and sign the handover list.',
        'If the Tenant stays on after this Agreement ends without a new agreement, the Tenant will pay damages for use and occupation at twice the daily rent for each day of overstay, without this creating a new tenancy.',
        'Belongings left behind more than 15 days after the end, after written notice to the Tenant, may be removed and disposed of by the Landlord at the Tenant\'s cost.']),
      sec('Promises by the Landlord', ['The Landlord confirms that it owns or has the right to let the Premises, that there is no legal dispute, loan default or order that stops it from letting them, and that the society or building rules allow letting.',
        'As long as the Tenant keeps to this Agreement, the Tenant may live in the Premises peacefully without interruption by the Landlord or anyone claiming through it.',
        x.inIN ? 'The Tenant will give the Landlord copies of identity documents of the adults living in the Premises, and the Landlord will complete tenant verification with the local police as the local rules require.' :
          x.inUK ? 'Where the law requires, the Landlord will carry out right-to-rent checks, protect the deposit in a government-approved scheme within 30 days and give the Tenant the prescribed information, the current gas safety record, the Energy Performance Certificate, the electrical safety report and any guide the Government requires.' :
            x.inAE ? 'The Landlord will register this tenancy with the official tenancy registration system of the Emirate (such as Ejari in Dubai or Tawtheeq in Abu Dhabi) and the Tenant will provide the documents needed. Rent increases on renewal will not exceed what the Emirate\'s rules allow.' :
              x.inUS ? 'The Landlord will make the disclosures that federal, state and local law require (for example the lead-based paint disclosure for homes built before 1978) and will keep the deposit as state law requires.' : 'Each Party will complete any registration, disclosure or verification that local law requires for a residential tenancy.'])]; },
    sched: function (x) { return [premisesSched(x)]; },
    notes: { IN: ['A lease for 12 months or more must be registered under the Registration Act, 1908, which is why many rent agreements are for 11 months. Some states require registration of every rent or leave and licence agreement (Maharashtra, for example), and registering also protects both sides.', 'Stamp duty on rent agreements depends on the state and is often worked out from the rent, deposit and term. Print the final draft on e-stamp or non-judicial stamp paper of that value.', 'States that have adopted the Model Tenancy Act may require the agreement to be reported to the Rent Authority; check your state\'s rule.'],
      UK: ['This draft is general. In England, the Renters\' Rights Act 2025 changes private tenancies (including ending fixed terms and "no-fault" evictions) as it comes into force: check the current rules and the Government\'s model tenancy on GOV.UK. Wales and Scotland have their own systems (occupation contracts and private residential tenancies).'],
      US: ['Landlord-tenant law is state and city law: deposit limits, return deadlines, notice periods, rent control and required disclosures differ. Use your state\'s rules or a local standard lease.'],
      AE: ['In Dubai the unified tenancy contract and Ejari registration are used, and in Abu Dhabi Tawtheeq. Use this draft as a list of points or an addendum to the official contract.'] } });

  type({ id: 'leave-licence', g: 0, name: 'Leave and licence agreement', kw: 'leave licence licensor licensee maharashtra mumbai pune rent 11 months',
    title: 'LEAVE AND LICENCE AGREEMENT', A: 'Licensor', B: 'Licensee', ownEnd: true, feeWord: 'licence fee',
    blurb: 'Permission to occupy a flat or room without creating a tenancy, common in Maharashtra and other states.',
    q: ['property', 'propdesc', 'start', 'term', ['rent', 'Monthly licence fee'], 'payday', 'deposit', 'paymode', 'maint', 'utilities', 'lockin', 'notice', 'occupants'],
    opts: { latefee: 0, dispute: 1, termination: 1 },
    defs: function (x) { return [['Licensed Premises', 'the premises described in Schedule 1, with the fittings listed there'], ['Licence Fee', 'the monthly fee set out in the Licence fee clause'], ['Licence Period', 'the period set out in the Grant of licence clause']]; },
    rec: function (x) { return ['The Licensor owns the Licensed Premises and is entitled to grant a licence to use them.', 'The Licensee has asked for permission to use the Licensed Premises for residence on leave and licence basis for a limited period, and the Licensor has agreed on the terms below.']; },
    body: function (x) { return [
      sec('Grant of licence', [L('The Licensor gives the Licensee leave and licence to use the Licensed Premises for residence for {0}, from {1} to {2} (the "Licence Period").', [x.mo('term'), x.d('start'), x.end()]),
        'This Agreement is a licence only. It does not create a tenancy, lease or any other interest in the Licensed Premises in favour of the Licensee, and legal possession stays with the Licensor at all times. The Licensee will not claim to be a tenant, sub-tenant or protected occupant.',
        'The licence is personal to the Licensee and the members of the family named to the Licensor, and cannot be transferred.']),
      sec('Licence fee', [L('The Licensee will pay a licence fee of {0} per month, in advance, by the {1} day of each month, by {2}.', [x.m('rent'), x.has('payday') ? ord(x.n('payday')) : '[ ]', x.raw('paymode') || 'bank transfer']),
        'The licence fee covers only the use of the Licensed Premises and the fittings listed in Schedule 1.']),
      depositSec(x, 'Licensor', 'Licensee', 'the Licensed Premises with all keys'),
      billsSec(x),
      sec('Use and conduct', [L('The Licensee will use the Licensed Premises only for residence, with no more than {0} people living there, and will not carry on any business from them.', [x.has('occupants') ? x.n('occupants') : '[ ]']),
        'The Licensee will not allow anyone else to occupy the Licensed Premises, will not make any alteration, and will follow the society\'s rules and pay any fine caused by the Licensee\'s conduct.',
        'The Licensor, or anyone authorised by it, may enter the Licensed Premises at reasonable times after notice to inspect them or carry out repairs.']),
      repairsSec(x, 'Licensor', 'Licensee'),
      sec('Ending the licence', [x.n('lockin') > 0 ? L('During the first {0} neither Party will end the licence except for breach; if the Licensee leaves during this period, the licence fee for the rest of it is payable.', [x.mo('lockin')]) : 'There is no lock-in period.',
        L('Either Party may end the licence earlier by giving at least {0} of written notice.', [x.days('notice')]),
        'The Licensor may revoke the licence by 15 days\' written notice if the Licensee does not pay the licence fee for two months, uses the premises for an unlawful purpose, or breaks this Agreement and does not put it right within 15 days of a written notice.',
        L('When the licence ends for any reason, the Licensee and every person living with the Licensee will vacate the Licensed Premises and hand over the keys. If the Licensee does not vacate, the Licensor may take steps under law to recover possession, and the Licensee will pay damages at twice the daily licence fee for every day of delay{0}.', [x.inIN ? ' (and, where section 24 of the Maharashtra Rent Control Act, 1999 applies, the damages provided in that section)' : ''])]),
      sec('Registration and verification', [x.inIN ? 'Where the law of the state requires it (for example section 55 of the Maharashtra Rent Control Act, 1999), this Agreement will be registered, and both Parties will attend or complete online registration. The Licensor will complete police verification of the Licensee.' : 'Each Party will complete any registration, verification or disclosure that local law requires for this licence.'])]; },
    sched: function (x) { return [premisesSched(x, 'Schedule 1: The Licensed Premises and fittings')]; },
    notes: { IN: ['In Maharashtra every leave and licence agreement must be registered, whatever its length; online registration is available through the IGR Maharashtra e-registration service. Stamp duty and registration fees are paid on the total licence fee and deposit as worked out under the state law.', 'Outside Maharashtra courts may treat a "leave and licence" as a tenancy if in substance it is one (exclusive possession, fixed rent, long period). Use the rent agreement if that is what you mean.'],
      all: ['Leave and licence is mainly an Indian concept. In other countries, use the residential rent agreement or a lodger agreement under local law.'] } });

  type({ id: 'pg', g: 0, name: 'Paying guest (PG) or room-sharing agreement', kw: 'pg paying guest hostel room sharing bed lodger boarding',
    title: 'PAYING GUEST AGREEMENT', A: 'Owner', B: 'Guest', ownEnd: true, feeWord: 'charges',
    blurb: 'A bed or room with shared areas and meals: charges, deposit, house rules, visitors and notice.',
    q: ['property', 'room', 'start', ['rent', 'Monthly charges'], 'meals', 'payday', 'deposit', 'paymode', 'notice', 'rules'],
    opts: { dispute: 0 },
    defs: function (x) { return [['House', L('the property at {0}', [x.v('property')])], ['Room', x.v('room')], ['Charges', 'the monthly charges in the Charges clause'], ['House Rules', 'the rules in Schedule 1']]; },
    rec: function (x) { return ['The Owner runs paying guest accommodation at the House.', 'The Guest wishes to stay in the House as a paying guest, and the Owner agrees, on the terms below.']; },
    body: function (x) { var meals = x.raw('meals') || 'Breakfast and dinner'; return [
      sec('Accommodation', [L('From {0} the Owner allows the Guest to stay in the Room and to use the shared kitchen, bathrooms, dining and common areas of the House with other guests.', [x.d('start')]),
        'The Guest is a paying guest and licensee only. The Guest does not get a tenancy or exclusive possession of any part of the House, and the Owner may enter the Room at reasonable times for cleaning, repairs and safety checks.',
        'The Owner may, with 7 days\' notice, move the Guest to another bed or room of the same standard in the House.']),
      sec('Charges', [L('The Guest will pay {0} per month in advance by the {1} day of each month, by {2}.', [x.m('rent'), x.has('payday') ? ord(x.n('payday')) : '[ ]', x.raw('paymode') || 'bank transfer']),
        L('The charges include: the bed with mattress, cupboard and study table; {0}; electricity for normal use of lights, fans and charging; drinking water; Wi-Fi; and cleaning of the common areas{1}. Air conditioning, laundry and extra services, if any, are charged as notified by the Owner in writing.', [meals === 'No meals' ? 'no meals' : meals.toLowerCase(), x.inIN ? '' : '']),
        'Charges for a part month at the start are worked out by days. No refund is given for days the Guest is away.']),
      depositSec(x, 'Owner', 'Guest', 'the Room and keys'),
      sec('House rules and visitors', ['The Guest will follow the House Rules in Schedule 1, which the Owner may update by giving 7 days\' notice of reasonable changes.',
        'Visitors are allowed only as the House Rules say. The Guest is responsible for the conduct of their visitors.',
        'The Guest will look after their own valuables. The Owner is not responsible for loss of the Guest\'s money or belongings unless caused by the Owner\'s or its staff\'s negligence.',
        'The Guest will report any illness, fire risk, damage or security problem to the Owner at once, and will pay for damage caused by the Guest or their visitors.']),
      sec('Identity and safety', [L('Before moving in the Guest will give a copy of a photo identity document, a recent photograph, a home address and an emergency contact. {0}', [x.inIN ? 'The Owner will complete police verification as local rules require.' : 'The Owner will keep these securely and use them only for the stay and as the law requires.']),
        'The Owner will keep fire extinguishers, working locks and a first-aid kit in the House, and will keep the House and common areas clean and safe.']),
      sec('Notice and ending', [L('Either Party may end this Agreement by giving at least {0} of notice in writing (a message or email is enough), or charges for the notice period instead.', [x.days('notice')]),
        'The Owner may end this Agreement at once, with refund of advance charges for unused days and the deposit after deductions, if the Guest seriously or repeatedly breaks the House Rules, behaves violently, brings in banned substances, or does not pay charges for 15 days after the due date.',
        'On leaving, the Guest will remove all belongings, return keys and access cards, and clear dues. Belongings left behind for more than 15 days after written notice may be disposed of.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: House Rules', items: [sub('', x.lines('rules').length ? x.lines('rules') : ['[house rules]'])] }]; },
    notes: { IN: ['Some cities require PG accommodation to be registered with the municipal corporation or police; check local rules. Stamp duty on the agreement depends on the state.'], all: ['In the UK and US a lodger or room-rental agreement under local law may be needed instead; consumer and housing rules can apply.'] } });

  type({ id: 'commercial', g: 0, name: 'Commercial lease (shop, office or warehouse)', kw: 'commercial lease shop office warehouse showroom business premises rent',
    title: 'COMMERCIAL LEASE AGREEMENT', A: 'Landlord', B: 'Tenant', ownEnd: true,
    blurb: 'Business premises: fit-out period, yearly rent increase, tax on rent, licences, signage, insurance and reinstatement.',
    q: ['property', 'propdesc', 'bizuse', 'start', ['term', 'Term (months)'], 'rent', 'escalation', 'payday', 'fitout', 'deposit', 'paymode', 'maint', 'utilities', 'lockin', 'notice', 'tax'],
    opts: { latefee: 1, dispute: 1, indemnity: 1, force: 1, termination: 1, liability: 0 },
    defs: function (x) { return [['Premises', 'the premises described in Schedule 1'], ['Permitted Use', x.v('bizuse')], ['Rent Commencement Date', 'the day after the fit-out period ends'], ['Lease Year', 'each period of 12 months from the start date']]; },
    rec: function (x) { return ['The Landlord owns the Premises and is entitled to lease them for commercial use.', 'The Tenant wishes to take the Premises on lease for its business, and the Landlord agrees on the terms below.']; },
    body: function (x) { var fit = x.n('fitout'); return [
      sec('Lease and term', [L('The Landlord leases the Premises to the Tenant for {0}, from {1} to {2}.', [x.mo('term'), x.d('start'), x.end()]),
        fit ? L('The first {0} days are a rent-free fit-out period for the Tenant to carry out approved interior works. Bills and maintenance charges are payable from the start date.', [fit]) : 'There is no rent-free fit-out period.',
        'The lease may be renewed on terms agreed in writing at least 60 days before the end of the term.']),
      sec('Rent and increases', [L('The Tenant will pay monthly rent of {0} in advance by the {1} day of each month from the Rent Commencement Date.{2}', [x.m('rent'), x.has('payday') ? ord(x.n('payday')) : '[ ]', x.taxed()]),
        L('The rent will increase by {0} at the start of each Lease Year after the first, as shown in Schedule 2.', [x.pct('escalation')]),
        x.inIN ? 'The Tenant will deduct tax at source from the rent where the Income-tax Act requires and give the Landlord TDS certificates on time. Where the Landlord is registered under GST, GST on rent will be charged on a proper tax invoice.' : 'Each Party will deal with any tax on rent as the law requires.']),
      depositSec(x, 'Landlord', 'Tenant', 'vacant possession of the Premises restored as required by this Agreement'),
      billsSec(x),
      sec('Use, licences and compliance', ['The Tenant will use the Premises only for the Permitted Use and will obtain, at its cost, every trade, shop and establishment, fire, signage and other licence that the business needs.',
        'The Landlord confirms the Premises may lawfully be used for commercial purposes and will give the documents (ownership proof, occupancy or completion certificate where available, consent letter) that the Tenant reasonably needs for its registrations.',
        'The Tenant will not store hazardous goods beyond what its business lawfully needs, overload floors or electrical lines, or cause nuisance to other occupants.']),
      sec('Fit-out, signage and alterations', ['The Tenant may carry out interior works, partitions, electrical and air-conditioning installations and signboards after getting the Landlord\'s written approval of the drawings, which will not be unreasonably refused, and any approval required from the building or authorities.',
        'At the end of the lease the Tenant may remove its fixtures and equipment and must restore the Premises to their original bare condition (except normal wear and tear and approved works the Landlord agrees in writing to keep).']),
      repairsSec(x, 'Landlord', 'Tenant'),
      sec('Insurance', ['The Landlord will insure the building structure. The Tenant will insure its fit-out, stock, equipment and its liability to the public. Neither Party will do anything that makes the other\'s insurance void.']),
      sec('Lock-in and notice', [x.n('lockin') > 0 ? L('Neither Party will end this lease in the first {0} except for breach. If the Tenant leaves earlier, it will pay the rent for the rest of the lock-in period.', [x.mo('lockin')]) : 'There is no lock-in period.',
        L('After that, the Tenant may end this lease by giving at least {0} of written notice. The Landlord may end it before the term only for the Tenant\'s breach: non-payment of rent for two months, or a serious breach not put right within 30 days of written notice.', [x.days('notice')])]),
      sec('Handing back', ['At the end of the lease the Tenant will hand back vacant possession with keys and access cards, restored as this Agreement requires. If the Tenant stays on without agreement it will pay twice the daily rent for each day of overstay.'])]; },
    sched: function (x) {
      var rows = [], r = x.n('rent'), years = Math.max(1, Math.ceil(x.n('term') / 12));
      for (var i = 0; i < Math.min(years, 10); i++) { rows.push([L('Lease Year {0}', [i + 1]), addMonths(x.raw('start'), i * 12) ? x.date(addMonths(x.raw('start'), i * 12)) : '[date]', r ? x.ms(r2(r)) : '[amount]']); r = r * (1 + x.n('escalation') / 100); }
      return [premisesSched(x, 'Schedule 1: The Premises'), { h: 'Schedule 2: Rent by lease year', items: [tbl(['Lease year', 'From', 'Monthly rent'], rows)] }]; },
    notes: { IN: ['A commercial lease for 12 months or more must be registered under the Registration Act, 1908 and stamped on the total rent, deposit and premium as the state law works it out. An unregistered long lease cannot be relied on for its terms.'],
      UK: ['Business tenancies in England and Wales may get security of tenure under the Landlord and Tenant Act 1954 unless the parties contract out by the statutory notice procedure before signing. Leases over 7 years must be registered at HM Land Registry and SDLT may be due.'],
      US: ['Commercial leases are governed by state law; check rules on recording, security deposits and personal guarantees.'], AE: ['Commercial tenancies must be registered (Ejari in Dubai, Tawtheeq in Abu Dhabi), and the trade licence is usually linked to the registered lease.'] } });

  type({ id: 'sell-property', g: 0, name: 'Agreement to sell property (starting draft)', kw: 'agreement to sell property flat plot house sale purchase token earnest money bayana',
    title: 'AGREEMENT TO SELL', A: 'Seller', B: 'Buyer', ownEnd: true, noEsign: true,
    blurb: 'Before the registered sale deed: price, advance, title checks, deadline for the balance and what happens if either side backs out.',
    q: ['property', 'propdesc', 'price', 'advance', 'paymode', 'balancedate', 'titledocs'],
    opts: { dispute: 1, force: 0 },
    defs: function (x) { return [['Property', 'the property described in Schedule 1, with all rights attached to it'], ['Sale Price', 'the total price in the Price clause'], ['Completion Date', L('{0}, or any later date agreed in writing', [x.d('balancedate')])], ['Sale Deed', 'the deed of sale to be signed and registered to transfer ownership']]; },
    rec: function (x) { return ['The Seller is the sole and absolute owner of the Property, as shown by the title documents listed in Schedule 2.', 'The Seller has agreed to sell, and the Buyer has agreed to buy, the Property free from all claims, on the terms below.']; },
    body: function (x) { var bal = x.n('price') - x.n('advance'); return [
      sec('Agreement to sell', [L('The Seller agrees to sell and the Buyer agrees to buy the Property for {0}.', [x.m('price')]),
        'This Agreement is an agreement to sell. Ownership will pass only when the Sale Deed is signed and registered and the full Sale Price is paid.']),
      sec('Price and payment', [L('The Buyer has paid {0} by {1} as advance and part payment on signing this Agreement, which the Seller acknowledges.', [x.m('advance'), x.raw('paymode') || 'bank transfer']),
        L('The balance of {0} will be paid on or before the Completion Date, at the time of registration of the Sale Deed{1}.', [x.has('price') && x.has('advance') ? x.m(bal) : '[balance amount]', x.inIN ? ', after deducting tax at source where the Income-tax Act requires (for example under section 194-IA)' : '']),
        'Where part of the price is paid by the Buyer\'s home loan, the Seller will give the bank the documents it reasonably needs, and the bank may pay that part directly to the Seller or the Seller\'s lender.']),
      sec('Title and the Seller\'s promises', [sub('The Seller confirms that:', ['it has clear and marketable title to the Property and full right to sell it;', 'the Property is free from any mortgage, charge, lien, lease, court order, acquisition notice or other claim, except those disclosed in writing to the Buyer;', 'all property taxes, society dues, electricity and water charges have been paid up to date; and', 'no other agreement to sell or transfer the Property has been made with anyone else.']),
        'The Seller has given the Buyer copies of the title documents in Schedule 2. The Buyer may verify the title, search the land and registration records and get a legal opinion within 30 days. If a defect in title is found that the Seller cannot cure before the Completion Date, the Buyer may cancel this Agreement and the Seller will refund all amounts paid within 15 days.',
        'Until completion the Seller will not sell, mortgage, let or otherwise deal with the Property, and will keep it in its present condition.']),
      sec('Completion', ['On the Completion Date, on receiving the balance, the Seller will sign the Sale Deed and appear for its registration, hand over vacant and peaceful possession with keys, and give all original title documents, no-dues certificates and the documents needed to transfer the property tax, electricity and society records to the Buyer.',
        'Stamp duty, registration fees and the costs of the Sale Deed will be paid by the Buyer, unless agreed otherwise. Each Party pays its own lawyer\'s fees.',
        'If the Property is on a loan, the Seller will pay off the loan out of the sale price and obtain the lender\'s release letter and original documents before or at completion.']),
      sec('If either Party fails', ['If the Buyer fails to pay the balance by the Completion Date (unless due to the Seller\'s default or a title defect), the Seller may, after 15 days\' written notice, cancel this Agreement and keep out of the advance a reasonable amount that reflects its actual loss, refunding the rest; keeping the whole advance is allowed only to the extent the law treats it as a reasonable earnest amount.',
        'If the Seller fails or refuses to complete the sale, the Buyer may, at its choice, ask the court for specific performance of this Agreement, or cancel it and get back the advance with simple interest at 12% per year from the date of payment, besides any other remedy the law allows.']),
      sec('Risk until completion', ['The risk of damage to the Property stays with the Seller until possession is handed over.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: The Property', items: [L('Address: {0}', [x.v('property')]), sub('Description, area and boundaries:', x.lines('propdesc').length ? x.lines('propdesc') : ['[survey or khasra number, area, boundaries on all four sides]'])] },
      { h: 'Schedule 2: Title documents', items: [sub('', x.lines('titledocs').length ? x.lines('titledocs') : ['[list each title document]'])] }]; },
    notes: { IN: ['An agreement to sell must be stamped as the state Stamp Act requires; where possession is given under it, it must be registered (section 17(1A) of the Registration Act, 1908) to rely on part performance. Many states charge stamp duty close to a sale deed when possession is handed over.', 'Do not pay large amounts in cash: cash receipts of ₹2 lakh or more in one transaction are prohibited (section 269ST of the Income-tax Act) and property cash of ₹20,000 or more is restricted (section 269SS).', 'Get the title checked by a lawyer and get an encumbrance certificate before paying the advance.'],
      UK: ['Property in the UK is sold through solicitors or licensed conveyancers by exchange of contracts on the Standard Conditions of Sale; use this draft only to list the points agreed.'],
      US: ['Real estate purchases use state-specific purchase contracts (often a broker or state bar form) with title insurance and escrow; use this only as a list of agreed points.'],
      AE: ['In Dubai, property sales between parties use the Dubai Land Department\'s standard contract (Form F) through a registered broker or trustee office; use this only as a list of points.'] } });

  type({ id: 'equipment-hire', g: 0, name: 'Equipment or goods hire agreement', kw: 'equipment hire rental machine generator scaffolding camera rent tools',
    title: 'EQUIPMENT HIRE AGREEMENT', A: 'Owner', B: 'Hirer', ownEnd: true, feeWord: 'hire charges',
    blurb: 'Renting out machines, tools or goods: hire charge, deposit, care, breakdowns, damage and late return.',
    q: ['equip', 'start', 'hiredays', 'hirerate', 'hireper', 'deposit', 'paymode', 'deliveryplace', 'tax'],
    opts: { latefee: 1, indemnity: 1, liability: 0, dispute: 1, force: 1 },
    defs: function (x) { return [['Equipment', 'the items listed in Schedule 1, with their accessories, manuals and keys'], ['Hire Period', 'the period in the Hire clause, as extended in writing'], ['Site', x.v('deliveryplace', 'site address')]]; },
    rec: function (x) { return ['The Owner owns the Equipment.', 'The Hirer wishes to hire the Equipment for use at the Site, and the Owner agrees on the terms below.']; },
    body: function (x) { return [
      sec('Hire', [L('The Owner hires the Equipment to the Hirer for {0} from {1} (the "Hire Period"). The Hire Period may be extended by written agreement at the same charges.', [x.days('hiredays', 'number of'), x.d('start')]),
        'The Equipment will be delivered to the Site in good working order. The Hirer will inspect it on delivery and record any existing damage on the delivery note; otherwise it is accepted as in good condition.']),
      sec('Hire charges and deposit', [L('The Hirer will pay hire charges of {0} {1} for the Hire Period, in advance{2}.{3}', [x.m('hirerate'), x.raw('hireper') || 'per day', (x.raw('hireper') || 'per day') === 'per month' ? ' each month' : '', x.taxed()]),
        L('The Hirer has paid a refundable security deposit of {0}, which the Owner will return within 7 days after the Equipment is returned, less unpaid charges and the cost of repairing damage or replacing missing parts.', [x.m('deposit')]),
        'Charges continue until the Equipment is returned to the Owner, even if the Hirer stops using it earlier.']),
      sec('Use and care', ['The Hirer will use the Equipment only at the Site, only for the purpose it is designed for, only by trained and competent people, and in line with the manufacturer\'s instructions and safety rules.',
        'The Hirer will not sell, lend, sub-hire, pledge or move the Equipment from the Site, alter it, or remove any identification mark, without the Owner\'s written consent. The Equipment remains the Owner\'s property at all times.',
        'The Hirer will carry out daily checks (such as fuel, oil and cleaning) and the Owner will carry out periodic servicing.']),
      sec('Breakdowns', ['The Hirer will stop using the Equipment and tell the Owner at once if it breaks down or becomes unsafe. The Owner will repair or replace it within 24 hours of being told, where reasonably possible.',
        'No hire charge is payable for any full day the Equipment cannot be used because of a breakdown not caused by the Hirer.']),
      sec('Risk, loss and damage', ['The Equipment is at the Hirer\'s risk from delivery until it is returned. The Hirer will pay for loss, theft or damage (other than normal wear and tear) at the cost of repair or, if it cannot be repaired, its replacement value as listed in Schedule 1.',
        'The Hirer will tell the Owner and, for theft, the police, within 24 hours of any loss or damage, and may insure the Equipment for its replacement value.']),
      sec('Return', [L('At the end of the Hire Period the Hirer will return the Equipment {0} clean, with all accessories, in the same condition as delivered except normal wear and tear.', [x.has('deliveryplace') ? 'for collection at the Site' : 'to the Owner']),
        'For each day of late return without agreement, the Hirer will pay the daily hire charge plus 50%.',
        'The Owner may take back the Equipment, after reasonable notice, if the Hirer does not pay the charges or misuses the Equipment.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Equipment', items: [tbl(['#', 'Equipment, model and serial number', 'Replacement value'], (x.lines('equip').length ? x.lines('equip') : ['[equipment]']).map(function (l, i) { return [String(i + 1), l, '']; }))] }]; },
    notes: { all: ['Where the hirer is a consumer, consumer protection law may limit some clauses (for example, fixed charges for damage).'] } });

  /* ===== Buying and selling */
  type({ id: 'goods-sale', g: 1, name: 'Sale of goods agreement', kw: 'sale of goods purchase order buy sell products supply one-off sales contract',
    title: 'SALE OF GOODS AGREEMENT', A: 'Seller', B: 'Buyer', fmRefund: true,
    blurb: 'A one-time sale of goods: list, price, delivery, inspection, when ownership and risk pass, warranty.',
    q: ['goods', 'tax', 'advance', 'paymode', 'payterms', 'delivery', 'deliveryplace', 'warranty'],
    opts: { latefee: 1, liability: 1, dispute: 1, force: 1, termination: 0, indemnity: 0 },
    defs: function (x) { return [['Goods', 'the goods described in Schedule 1'], ['Price', 'the total price in Schedule 1'], ['Delivery Date', x.d('delivery')], ['Delivery Address', x.v('deliveryplace')]]; },
    rec: function (x) { return ['The Seller deals in the Goods.', 'The Buyer wishes to buy the Goods, and the Seller agrees to sell them, on the terms below.']; },
    body: function (x) { var g = goodsRows(x); return [
      sec('Sale and price', [L('The Seller will sell and the Buyer will buy the Goods for {0}, made up as shown in Schedule 1.{1}', [g.total ? x.m(g.total) : x.m('price'), x.taxed()]),
        L('The Price includes packing and delivery to the Delivery Address unless Schedule 1 says otherwise.')]),
      sec('Payment', [x.has('advance') ? L('The Buyer has paid {0} as advance on signing. The balance will be paid within {1} after delivery and receipt of the Seller\'s invoice.', [x.m('advance'), x.days('payterms')]) : L('The Buyer will pay the Price within {0} after delivery and receipt of the Seller\'s invoice.', [x.days('payterms')]),
        L('Payment will be made by {0} to the account named in the invoice.', [x.raw('paymode') || 'bank transfer'])]),
      sec('Delivery', [L('The Seller will deliver the Goods to the Delivery Address on or before the Delivery Date, during normal business hours, with a delivery note{0}.', [x.inIN ? ', tax invoice and e-way bill where the law requires' : ' and invoice']),
        'If the Seller expects a delay, it will tell the Buyer at once with a new date. If delivery is more than 15 days late (other than for events beyond control), the Buyer may cancel the undelivered part and get back any advance paid for it.',
        'Delivery in instalments is allowed only if the Buyer agrees.']),
      sec('Inspection and acceptance', ['The Buyer will inspect the Goods within 7 days of delivery. If any Goods are damaged, short, or not as described, the Buyer will tell the Seller in writing within that time with photographs.',
        'The Seller will, at its own cost and at its choice, replace or repair the rejected Goods within 15 days or refund their price. The Goods are treated as accepted if the Buyer does not notify a problem within 7 days, except for hidden defects covered by the warranty.']),
      sec('Ownership and risk', ['Risk of loss or damage passes to the Buyer on delivery at the Delivery Address.',
        'Ownership of the Goods passes to the Buyer only when the Price has been paid in full. Until then the Buyer will keep the Goods separately identifiable and insured.']),
      sec('Warranty', [L('The Seller warrants that the Goods will be new (unless described otherwise), of satisfactory quality, fit for their usual purpose, match their description and any sample, and be free from defects in material and workmanship for {0} from delivery.', [x.mo('warranty')]),
        'The warranty does not cover damage from misuse, accident, normal wear, or repairs by anyone not approved by the Seller. Any manufacturer\'s warranty is passed on to the Buyer.',
        x.inIN ? 'The rights of the Buyer under the Sale of Goods Act, 1930 and the Consumer Protection Act, 2019 (where it applies) are not reduced by this Agreement.' : x.inUK ? 'Where the Buyer is a consumer, nothing in this Agreement affects the Buyer\'s statutory rights under the Consumer Rights Act 2015.' : x.inUS ? 'Except as stated in this clause, the Seller gives no other warranties, express or implied, to the extent permitted by the Uniform Commercial Code as adopted in the State.' : 'Nothing in this Agreement reduces any right the Buyer has under consumer protection law.'])]; },
    sched: function (x) { var g = goodsRows(x); var rows = g.rows.length ? g.rows : [['1', '[goods]', '', '', '']]; rows.push(['', 'Total', '', '', g.total ? x.ms(g.total) : '']); return [{ h: 'Schedule 1: Goods and price', items: [tbl(['#', 'Description', 'Quantity', 'Unit price', 'Amount'], rows)] }]; },
    notes: { all: ['For regular supplies over a period, use the vendor or supplier agreement instead.'] } });

  type({ id: 'vehicle-sale', g: 1, name: 'Vehicle sale agreement (car, bike, truck)', kw: 'vehicle car bike scooter motorcycle truck sale purchase second hand used rc transfer',
    title: 'VEHICLE SALE AGREEMENT', A: 'Seller', B: 'Buyer', ownEnd: true,
    blurb: 'Selling a used vehicle: details, price, documents, transfer of registration, and who is responsible for fines before and after handover.',
    q: ['vmake', 'vreg', 'vyear', 'vchassis', 'vengine', 'vkm', 'price', 'advance', 'paymode', 'delivery'],
    opts: { dispute: 0 },
    defs: function (x) { return [['Vehicle', 'the vehicle described in Schedule 1 with its keys, tools and accessories'], ['Handover', 'the date and time the Vehicle is delivered to the Buyer, recorded in Schedule 1']]; },
    rec: function (x) { return ['The Seller is the registered owner of the Vehicle.', 'The Seller has agreed to sell, and the Buyer has agreed to buy, the Vehicle on the terms below.']; },
    body: function (x) { return [
      sec('Sale and price', [L('The Seller sells the Vehicle to the Buyer for {0}.', [x.m('price')]), x.has('advance') ? L('The Buyer has paid {0} on signing, and will pay the balance by {1} at Handover on {2}.', [x.m('advance'), x.raw('paymode') || 'bank transfer', x.d('delivery')]) : L('The Buyer will pay the price by {0} at Handover on {1}.', [x.raw('paymode') || 'bank transfer', x.d('delivery')]),
        'The Buyer has inspected and test-driven the Vehicle and buys it in its present condition, as seen. The Seller has told the Buyer about every known accident, major repair and defect.']),
      sec('Seller\'s promises', [sub('The Seller confirms that:', ['the Vehicle is not stolen, and the chassis and engine numbers match the registration certificate;', 'there is no loan, hire purchase or hypothecation on the Vehicle, or, if there is, the Seller will pay it off and give the lender\'s no-objection certificate before Handover;', 'there is no pending tax, fine, challan or court case against the Vehicle, except as listed in Schedule 1; and', 'the odometer reading has not been tampered with.'])]),
      sec('Handover and documents', [L('At Handover the Seller will give the Buyer the Vehicle, all keys, and {0}.', [x.inIN ? 'the registration certificate (RC), valid insurance policy, pollution under control (PUC) certificate, service book, and signed Forms 29 and 30 (and the lender\'s NOC and Form 35, if there was a loan)' : x.inUK ? 'the V5C registration certificate (logbook) with the new keeper section completed, the MOT certificate if any, the service history and all keys' : x.inUS ? 'the certificate of title signed over to the Buyer, a bill of sale, the odometer disclosure statement and lien release (if any)' : x.inAE ? 'the registration card (Mulkiya), and will attend the transfer of ownership at the traffic authority' : 'the registration documents and any document needed to transfer ownership']),
        x.inIN ? 'The Seller will inform the registering authority of the sale within 14 days (or as the Motor Vehicles Act and rules require), and the Buyer will apply for transfer of registration within 30 days and pay the transfer fee. Until the RC is transferred, the Seller will cooperate with any signature or appearance needed.' : x.inUK ? 'The Seller will tell DVLA about the sale online or by post on the day of Handover. The Buyer will tax and insure the Vehicle before driving it.' : x.inUS ? 'The Buyer will register and title the Vehicle with the state motor vehicle agency within the time state law allows, and the Seller will remove its licence plates where state law requires.' : 'The Parties will complete the transfer of registration with the authority within 30 days of Handover.',
        'The Buyer will arrange insurance in its own name, or transfer the existing policy, from Handover.']),
      sec('Responsibility before and after Handover', ['The Seller is responsible for all accidents, offences, fines, tolls and claims relating to the Vehicle up to the date and time of Handover recorded in Schedule 1. The Buyer is responsible for everything after that time, even while the registration still shows the Seller\'s name, and will indemnify the Seller against it.',
        'If the Buyer fails to complete the transfer of registration on time, the Seller may report the sale to the authority and the police, and the Buyer will pay any cost caused.']),
      sec('Cancellation', ['If the Buyer does not pay the balance at Handover, the Seller may cancel the sale by written notice and refund the advance less reasonable costs actually caused. If the Seller cannot hand over a clear title, the Buyer may cancel and get back all amounts paid.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: The Vehicle', items: [tbl(['Item', 'Details'], [['Make and model', x.v('vmake')], ['Registration number', x.v('vreg')], ['Year of manufacture', x.v('vyear')], ['Chassis number', x.v('vchassis')], ['Engine number', x.v('vengine')], ['Odometer reading', x.has('vkm') ? L('{0} km', [x.n('vkm').toLocaleString('en-IN')]) : '[km]'], ['Handover date and time', ''], ['Pending fines or loan (if any)', '']])] }]; },
    notes: { IN: ['Check pending challans and loans on the Parivahan (VAHAN) portal before paying. Stamp duty on the agreement depends on the state.'], all: ['Take photographs of the vehicle and the odometer at handover and keep a copy of every signed transfer form.'] } });

  type({ id: 'supply', g: 1, name: 'Vendor or supplier agreement', kw: 'vendor supplier supply purchase orders framework raw material regular supply rate contract',
    title: 'SUPPLY AGREEMENT', A: 'Buyer', B: 'Supplier', ip: 'owner', indem: 'B', ipIndem: true,
    blurb: 'Regular supply against purchase orders: prices, lead time, quality, rejection, warranty and payment.',
    q: ['products', 'start', ['term', 'Term (months)'], 'leadtime', 'deliveryplace', 'payterms', 'warranty', 'tax', 'notice'],
    opts: { conf: 1, latefee: 1, indemnity: 1, liability: 1, termination: 1, force: 1, dispute: 1, data: 0 },
    defs: function (x) { return [['Products', L('{0}, as listed with prices in Schedule 1', [x.v('products')])], ['Order', 'a written purchase order from the Buyer for Products'], ['Specifications', 'the description, drawings, quality standards and samples agreed for each Product']]; },
    rec: function (x) { return ['The Supplier makes or deals in the Products.', 'The Buyer wishes to buy Products from time to time and the Supplier agrees to supply them, on the terms below.']; },
    body: function (x) { return [
      sec('Term and scope', [L('This Agreement starts on {0} and continues for {1}, unless ended earlier under its terms. It sets the terms for every Order during that time.', [x.d('start'), x.mo('term')]),
        'The Buyer is not obliged to buy any minimum quantity, and the Supplier may supply the same products to others, unless Schedule 1 says otherwise.']),
      sec('Orders', ['The Buyer will place Orders in writing (including by email) stating Products, quantity, price, delivery date and address.',
        'The Supplier will accept or reject each Order within 2 working days. An Order not rejected in that time is treated as accepted. Once accepted, an Order may be changed or cancelled only by agreement.',
        'If an Order conflicts with this Agreement, this Agreement prevails unless the Order expressly says it changes a named clause.']),
      sec('Prices', [L('Prices are as set out in Schedule 1 and stay fixed for the first 6 months.{0}', [x.taxed()]),
        'After that, the Supplier may change prices by giving at least 30 days\' written notice with reasons. Orders accepted before the change take effect are supplied at the old price.']),
      sec('Delivery', [L('The Supplier will deliver within {0} of accepting an Order to {1}, packed to prevent damage, with a delivery note{2} and test or quality certificates where agreed.', [x.days('leadtime'), x.v('deliveryplace', 'delivery address'), x.inIN ? ', tax invoice, e-way bill (where needed)' : ', invoice']),
        'Time of delivery matters. If a delivery is late by more than 7 days, the Buyer may cancel that delivery without cost, buy elsewhere, and claim the reasonable extra cost from the Supplier.']),
      sec('Quality, inspection and rejection', [L('Products must match the Specifications, be of good quality, and comply with the laws and standards that apply to them{0}.', [x.inIN ? ' (including BIS, FSSAI or Legal Metrology requirements where relevant)' : '']),
        'The Buyer may inspect Products on delivery and reject any that do not conform within 10 days by written notice. Rejected Products will be collected and replaced (or credited) by the Supplier at its cost within 7 days.',
        'Payment for Products does not mean they are accepted as free from hidden defects.']),
      sec('Warranty', [L('The Supplier warrants each Product against defects in design, material and workmanship for {0} from delivery, and will repair, replace or refund any defective Product at its cost.', [x.mo('warranty')])]),
      sec('Payment', [L('The Supplier will invoice each delivery. The Buyer will pay correct invoices within {0} of receiving the Products and invoice, and may withhold payment only for the disputed part of an invoice, giving reasons.', [x.days('payterms')])]),
      sec('Conduct', ['The Supplier will comply with all laws on labour, wages, safety, the environment and anti-bribery, will not use child or forced labour, and will not offer any gift or payment to the Buyer\'s staff to win business.',
        'The Supplier will keep adequate insurance for product liability and goods in transit.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Products and prices', items: [L('Products: {0}', [x.v('products')]), tbl(['#', 'Product and specification', 'Unit', 'Price per unit'], [['1', '', '', ''], ['2', '', '', ''], ['3', '', '', '']])] }]; },
    notes: { IN: ['If the Supplier is a registered micro or small enterprise (Udyam), the MSMED Act, 2006 requires payment within the agreed time and no later than 45 days; later payments attract compound interest at three times the bank rate.'] } });

  type({ id: 'distribution', g: 1, name: 'Distribution agreement', kw: 'distributor dealer stockist wholesale territory reseller appointment',
    title: 'DISTRIBUTION AGREEMENT', A: 'Company', B: 'Distributor', ip: 'owner', nc: 'B',
    blurb: 'Appointing a distributor who buys and resells in a territory: exclusivity, targets, margins, brand use and ending.',
    q: ['products', 'territory', 'exclusive', 'target', 'margin', 'start', ['term', 'Term (months)'], 'payterms', 'notice'],
    opts: { conf: 1, noncompete: 0, nonsolicit: 0, latefee: 1, termination: 1, indemnity: 1, liability: 1, force: 1, dispute: 1 },
    defs: function (x) { return [['Products', x.v('products')], ['Territory', x.v('territory')], ['Trade Marks', 'the names, logos and brands under which the Company sells the Products'], ['List Price', 'the Company\'s distributor price list in force from time to time']]; },
    rec: function (x) { return ['The Company makes and sells the Products under the Trade Marks.', 'The Distributor has the premises, staff and network to sell the Products in the Territory, and the Company wishes to appoint it as distributor on the terms below.']; },
    body: function (x) { var ex = (x.raw('exclusive') || 'Exclusive') === 'Exclusive'; return [
      sec('Appointment', [L('The Company appoints the Distributor as its {0} distributor of the Products in the Territory for {1} from {2}.', [ex ? 'exclusive' : 'non-exclusive', x.mo('term'), x.d('start')]),
        ex ? 'During this Agreement the Company will not appoint another distributor for the Products in the Territory. The Company may still sell directly to government buyers, national chains and online marketplaces, and will pay no commission on such sales unless agreed.' : 'The Company may appoint other distributors in the Territory and sell directly there.',
        'The Distributor buys the Products and resells them in its own name and at its own risk. It is not the Company\'s agent and cannot make promises or contracts for the Company.']),
      sec('Distributor\'s duties', [sub('The Distributor will:', ['actively promote and sell the Products throughout the Territory and not actively seek customers outside it;', 'keep enough stock to meet normal demand and store it as the Company advises;', 'keep a sales team, serve retailers regularly and handle customer complaints promptly;', 'give the Company a monthly report of sales, stock and market information by the 7th of each month; and', 'not change the Products, their packaging, labels or expiry dates.'])]),
      sec('Targets', [L('The Distributor aims to buy Products worth at least {0} (at List Price, before tax) in each year of this Agreement.', [x.m('target')]),
        L('If the Distributor buys less than 75% of the target in a year, the Company may, by written notice, {0}end this Agreement with {1} of notice.', [ex ? 'make the appointment non-exclusive or ' : '', x.days('notice')])]),
      sec('Orders, prices and payment', [L('The Distributor will buy Products at List Price less a margin of {0}, unless a scheme says otherwise. The Company may change the List Price by 30 days\' written notice.{1}', [x.pct('margin'), x.taxed()]),
        'Orders are subject to the Company\'s acceptance and stock. Delivery is ex-warehouse unless agreed; risk passes on dispatch and ownership on payment.',
        L('The Distributor will pay invoices within {0}. The Company may stop supplies while any invoice is overdue.', [x.days('payterms')]),
        'The Distributor is free to set its own resale prices; any price the Company suggests is only a recommendation.']),
      sec('Trade Marks and marketing', ['The Distributor may use the Trade Marks only to identify and promote the Products in the Territory, following the Company\'s brand guidelines, and will not register any similar name, mark or domain.',
        'Marketing materials and schemes will be agreed in advance. Any shared cost will be agreed in writing.']),
      sec('Returns and ending', ['Expired, damaged-in-transit or defective Products will be replaced or credited as per the Company\'s returns policy.',
        'When this Agreement ends the Company may buy back saleable stock in original packing at the price paid by the Distributor, and the Distributor will stop using the Trade Marks and return all marketing material and customer lists the Company provided.'])]; },
    notes: { all: ['Exclusive territories, resale price controls and non-compete terms can raise competition law issues (for example under India\'s Competition Act, 2002, UK and EU competition law, or US antitrust law). Keep resale prices free and restrictions reasonable.'] } });

  type({ id: 'franchise', g: 1, name: 'Franchise agreement (outline)', kw: 'franchise franchisor franchisee outlet royalty brand business format',
    title: 'FRANCHISE AGREEMENT (OUTLINE)', A: 'Franchisor', B: 'Franchisee', ip: 'owner', nc: 'B', indem: 'B',
    blurb: 'An outline for a single outlet: fees, royalty, standards, training, supplies, brand use, audits and de-branding at the end.',
    q: ['brand', 'location', 'territory', 'ffee', 'royalty', 'mktfee', 'start', ['term', 'Term (months)'], 'payterms', 'tax'],
    opts: { conf: 1, noncompete: 1, nonsolicit: 1, latefee: 1, termination: 1, indemnity: 1, liability: 1, data: 1, dispute: 1, force: 1 },
    defs: function (x) { return [['System', L('the Franchisor\'s method of running a {0} business, including recipes or processes, layouts, suppliers, software and know-how', [x.v('brand')])], ['Marks', L('the name "{0}" and related logos and trade marks', [x.v('brand')])], ['Outlet', L('the franchised business at {0}', [x.v('location')])], ['Manual', 'the operations manual given by the Franchisor, as updated'], ['Gross Sales', 'all sales of the Outlet, less taxes collected and genuine refunds']]; },
    rec: function (x) { return ['The Franchisor has developed the System and owns the Marks.', 'The Franchisee wishes to run one Outlet using the System and the Marks, and the Franchisor agrees on the terms below.']; },
    body: function (x) { return [
      sec('Grant', [L('The Franchisor grants the Franchisee the right to run one Outlet at {0} using the System and the Marks for {1} from {2}.', [x.v('location'), x.mo('term'), x.d('start')]),
        L('The Franchisor will not open or license another {0} outlet within {1} during the term, but may sell through delivery apps and online.', [x.v('brand'), x.v('territory', 'protected area')]),
        'The Franchisee may renew for a further term on the Franchisor\'s then-current terms if it has met this Agreement and gives notice at least 6 months before the end.']),
      sec('Fees', [L('Initial franchise fee: {0}, payable on signing, for the grant, site approval, layout and initial training. It is not refundable once initial training has started, except as the law requires.{1}', [x.m('ffee'), x.taxed()]),
        L('Royalty: {0} of Gross Sales each month, and marketing fund contribution: {1} of Gross Sales each month, both payable with a sales statement by the 10th of the following month.', [x.pct('royalty'), x.pct('mktfee')]),
        'The Franchisor will use the marketing fund for brand advertising and promotion and give a yearly summary of how it was spent.']),
      sec('Setting up the Outlet', ['The Franchisee will fit out the Outlet at its cost to the Franchisor\'s approved design and opens it within 90 days of signing, after the Franchisor confirms it meets the standards.',
        L('The Franchisee will obtain all licences and registrations for the Outlet in its own name{0}.', [x.inIN ? ' (such as FSSAI, shop and establishment, GST and fire)' : ''])]),
      sec('Training and support', ['The Franchisor will train the Franchisee and its manager before opening, give the Manual, help with opening, and give continuing support by phone, visits and updates to the System.']),
      sec('Operating standards', ['The Franchisee will run the Outlet strictly in line with the Manual: product quality, menu or range, hygiene, uniforms, customer service, opening hours and software.',
        'The Franchisee will buy key ingredients, materials and equipment only from the Franchisor or approved suppliers, at the prices they set, and may buy other items from any supplier meeting the specifications.',
        'The Franchisee may set its own selling prices, though the Franchisor may recommend them.',
        'The Franchisor may inspect the Outlet at any time during business hours and, by mystery shoppers or audits, check standards. The Franchisee will put right any failure within the time the Franchisor reasonably sets.']),
      sec('Records and audit', ['The Franchisee will record all sales in the approved billing software, keep proper accounts and give the Franchisor access to sales data. The Franchisor may audit the books on 7 days\' notice; if Gross Sales were under-reported by more than 2%, the Franchisee pays the shortfall, interest and the audit cost.']),
      sec('Marks and know-how', ['The Franchisee may use the Marks only at the Outlet and as the Manual allows. All goodwill belongs to the Franchisor. The System and Manual are confidential.']),
      sec('Transfer', ['The Franchisee may not sell or transfer the Outlet or this Agreement without the Franchisor\'s written consent, which may require the buyer to be trained and a transfer fee paid. The Franchisor may transfer this Agreement to a buyer of its business.']),
      sec('End of the franchise', ['When this Agreement ends, the Franchisee will stop using the Marks and System at once, remove all signs and branding within 15 days, return the Manual and confidential material, and pay all amounts due.',
        'The Franchisor may, within 30 days, buy the Outlet\'s equipment and stock at their written-down value.'])]; },
    notes: { US: ['US franchisors must give the Franchise Disclosure Document at least 14 days before signing or payment (FTC Franchise Rule), and some states require registration. This outline is not a substitute.'], UK: ['There is no franchise statute in the UK, but the British Franchise Association\'s code is widely followed.'], AE: ['UAE franchises usually need the Marks registered with the Ministry of Economy and may need a commercial agency or licence review.'], IN: ['India has no franchise-specific law; the Contract Act, trade mark law, competition law and FEMA (for foreign franchisors) apply. Register the Marks before franchising.'] } });

  type({ id: 'agent', g: 1, name: 'Commission or sales agent agreement', kw: 'commission agent sales agent broker referral representative intermediary',
    title: 'SALES AGENT AGREEMENT', A: 'Principal', B: 'Agent', ip: 'owner', nc: 'B',
    blurb: 'An agent who finds customers for a commission: no authority to bind, commission on paid sales, reports and ending.',
    q: ['products', 'territory', 'exclusive', 'commission', ['payterms', 'Commission paid within (days) after the customer pays'], 'start', ['term', 'Term (months)'], 'notice', 'tax'],
    opts: { conf: 1, noncompete: 0, nonsolicit: 1, termination: 1, indemnity: 0, liability: 1, dispute: 1, data: 1 },
    defs: function (x) { return [['Products', x.v('products')], ['Territory', x.v('territory')], ['Customer', 'a buyer in the Territory introduced by the Agent and accepted by the Principal'], ['Paid Sales', 'the value of Products sold to Customers, excluding taxes, freight and returns, actually received by the Principal']]; },
    rec: function (x) { return ['The Principal sells the Products.', 'The Agent has contacts in the Territory and will look for customers for the Principal in return for commission, on the terms below.']; },
    body: function (x) { return [
      sec('Appointment', [L('The Principal appoints the Agent as its {0} sales agent for the Products in the Territory for {1} from {2}.', [(x.raw('exclusive') || 'Exclusive') === 'Exclusive' ? 'exclusive' : 'non-exclusive', x.mo('term'), x.d('start')]),
        'The Agent will find and introduce customers and pass on their enquiries and orders. The Agent has no authority to accept orders, agree prices or terms, collect money (unless authorised in writing), or bind the Principal in any way, and will not say otherwise.',
        'The Principal may accept or reject any order at its discretion and will tell the Agent which orders it accepts.']),
      sec('Commission', [L('The Principal will pay the Agent commission of {0} of Paid Sales to Customers{1}.{2}', [x.pct('commission'), (x.raw('exclusive') || 'Exclusive') === 'Exclusive' ? ', and on Paid Sales to other buyers in the Territory during this Agreement' : '', x.taxed()]),
        L('Commission is earned only when the Customer pays, and is payable within {0} after the end of the month in which payment is received, with a statement of the sales it covers. Commission paid on sales later refunded will be adjusted.', [x.days('payterms')]),
        'Commission is also payable on repeat orders from a Customer within 6 months after this Agreement ends, if they result mainly from the Agent\'s work during it.',
        'The Agent bears its own travel, office and other expenses unless the Principal agrees in writing in advance.']),
      sec('Agent\'s duties', ['The Agent will act in good faith, follow the Principal\'s reasonable instructions, describe the Products accurately using only approved material, give monthly reports of visits and prospects, and not take any secret commission or payment from customers.']),
      sec('Principal\'s duties', ['The Principal will give the Agent product information, samples and price lists, tell it promptly about accepted and rejected orders and expected lower volumes, and keep records to check commission.']),
      sec('Ending', [L('Either Party may end this Agreement by giving at least {0} of written notice.{1}', [x.days('notice'), x.inUK ? ' Where the Commercial Agents (Council Directive) Regulations 1993 apply, the notice will be no shorter than those Regulations require (one month for each year of the agency, up to three months), and the Agent may be entitled to an indemnity or compensation as those Regulations provide.' : '']),
        'On ending, the Agent will return samples, price lists and customer data, and commission earned up to the end will be paid.'])]; },
    notes: { UK: ['Self-employed agents who sell goods may be protected by the Commercial Agents Regulations 1993, which give minimum notice and compensation or indemnity on termination that cannot be excluded.'], AE: ['A registered commercial agency in the UAE is governed by the Commercial Agencies law and must be registered with the Ministry of Economy; this draft is for an unregistered commission arrangement.'] } });

  /* ===== Services and work */
  type({ id: 'service', g: 2, name: 'Service agreement', kw: 'service agreement contract services provider client consultancy outsourcing bookkeeping cleaning security',
    title: 'SERVICE AGREEMENT', A: 'Client', B: 'Service Provider', ip: 'client', indem: 'mutual', ipIndem: true,
    blurb: 'Ongoing or one-off services: scope, standards, fees, invoices, changes, staff and ending.',
    q: ['services', 'start', ['term', 'Term (months)'], 'fee', 'feebasis', 'payterms', 'tax', 'notice'],
    opts: { conf: 1, ip: 1, indemnity: 1, liability: 1, data: 1, latefee: 1, termination: 1, force: 1, dispute: 1, nonsolicit: 0, noncompete: 0 },
    defs: function (x) { return [['Services', 'the services described in Schedule 1, and any changes agreed in writing'], ['Fees', 'the fees in the Fees clause'], ['Effective Date', x.d('start')]]; },
    rec: function (x) { return ['The Service Provider is in the business of providing services of the kind described in Schedule 1.', 'The Client wishes to engage the Service Provider to provide the Services, on the terms below.']; },
    body: function (x) { return [
      sec('Services and term', [L('The Service Provider will provide the Services from the Effective Date for {0}, and after that until either Party ends this Agreement by notice.', [x.mo('term')]),
        'The Service Provider will perform the Services with reasonable skill, care and diligence, using suitably qualified people, in line with good industry practice, the law, and the Client\'s reasonable site and security rules that it has been told about.',
        'The Service Provider will name a contact person who will report on progress as reasonably requested.']),
      sec('Client\'s duties', ['The Client will give the information, access, approvals and decisions the Service Provider reasonably needs, on time. The Service Provider is not responsible for delay caused by the Client failing to do so.']),
      sec('Fees and invoices', [L('The Client will pay {0}', [feeLine(x)]), invoiceLine(x, 'Client', 'Service Provider'),
        'Expenses such as travel or third-party costs are payable only if the Client approved them in writing in advance, at cost against receipts.',
        'If the Client disputes part of an invoice in good faith, it will pay the undisputed part on time and tell the Service Provider the reason within 7 days.']),
      sec('Changes', ['Either Party may ask for a change to the Services. A change takes effect only when both Parties agree in writing its effect on scope, fees and timing.']),
      sec('People', [L('The Service Provider is responsible for its own staff and subcontractors, including their pay, taxes, {0}insurance and conduct, and will comply with labour laws. Nobody it uses becomes an employee of the Client.', [x.inIN ? 'PF, ESI, ' : x.inUK ? 'PAYE, National Insurance, pension contributions, ' : '']),
        'The Service Provider may use subcontractors with the Client\'s prior consent and remains responsible for their work.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: The Services', items: [sub('', x.lines('services').length ? x.lines('services') : ['[describe each service, how often, and the standard expected]'])] }]; },
    notes: { all: ['Describe the Services in measurable words (how often, by when, to what standard). Most service disputes start with vague scope.'] } });

  type({ id: 'freelance', ex: { services: 'Logo design: 3 first concepts, then the chosen one refined\nBrand colour palette and font guide (PDF)\nSocial media profile images and banners for 3 platforms', milestones: 'Advance | On signing | 15000\nLogo concepts | 10 days after signing | 15000\nFinal files | On approval | 10000' }, g: 2, name: 'Freelance or consultancy agreement', kw: 'freelance freelancer consultant consultancy independent contractor designer writer developer project',
    title: 'FREELANCE SERVICES AGREEMENT', A: 'Client', B: 'Freelancer', ip: 'client', indem: 'mutual', fmRefund: true,
    blurb: 'A project for a freelancer or consultant: deliverables, revisions, acceptance, payments by stage, kill fee and ownership.',
    q: ['services', 'milestones', 'start', 'revisions', 'payterms', 'tax', 'notice'],
    opts: { conf: 1, ip: 1, latefee: 1, liability: 1, termination: 1, dispute: 1, indemnity: 0, force: 0, data: 0 },
    defs: function (x) { return [['Project', 'the work described in Schedule 1'], ['Deliverables', 'the work products listed in Schedule 1'], ['Revision', 'one consolidated set of change requests to a Deliverable within the agreed brief']]; },
    rec: function (x) { return ['The Freelancer is an independent professional offering the skills needed for the Project.', 'The Client wishes to hire the Freelancer for the Project, on the terms below.']; },
    body: function (x) { var ms = msRows(x); return [
      sec('The work', [L('The Freelancer will carry out the Project and deliver the Deliverables, starting on {0}, by the dates in Schedule 2 (which move by any delay in the Client\'s feedback or materials).', [x.d('start')]),
        'The Freelancer decides how, when and where to do the work and uses its own equipment, but will meet the agreed deadlines and keep the Client informed.']),
      sec('Revisions and acceptance', [L('The fee includes {0} Revisions for each Deliverable. Further revisions, or changes outside the brief, will be quoted and charged separately.', [x.has('revisions') ? x.n('revisions') : '[number of]']),
        'The Client will review each Deliverable and either approve it or send one consolidated list of revisions within 7 days. If the Client sends nothing in that time, the Deliverable is treated as approved.']),
      sec('Fees and payment', [L('The Client will pay the fees in Schedule 2{0}.{1}', [ms.total ? L(', totalling {0}', [x.m(ms.total)]) : '', x.taxed()]),
        L('The Freelancer will invoice each stage when it falls due, and the Client will pay within {0} by bank transfer{1}. Payments made in advance are non-refundable once the work for that stage has started, except where the Freelancer fails to deliver.', [x.days('payterms'), x.inIN ? ' (after any TDS the law requires)' : '']),
        'The Freelancer may pause work while any invoice is more than 15 days overdue, after notice.']),
      sec('Cancellation', ['The Client may cancel the Project at any time by written notice. The Client will then pay for all work done up to cancellation (at least the fees for stages completed plus a fair share of the current stage) and any approved expenses, and the Freelancer will hand over the work paid for.',
        L('The Freelancer may end this Agreement by {0} of written notice, refunding fees paid for work not done.', [x.days('notice')])]),
      sec('Independent contractor', [L('The Freelancer is self-employed and responsible for its own taxes, insurance and registrations. Nothing in this Agreement creates employment.{0}', [x.inUK ? ' The Client will tell the Freelancer if it decides that the off-payroll working rules (IR35) apply.' : ''])])]; },
    sched: function (x) { var ms = msRows(x); var rows = ms.rows.length ? ms.rows : [['1', '[stage]', '[when due]', '[amount]']]; if (ms.total) rows.push(['', 'Total', '', x.ms(ms.total)]);
      return [{ h: 'Schedule 1: Project and deliverables', items: [sub('', x.lines('services').length ? x.lines('services') : ['[describe the deliverables, formats and brief]'])] }, { h: 'Schedule 2: Payment schedule', items: [tbl(['#', 'Stage', 'When due', 'Amount'], rows)] }]; },
    notes: { all: ['Ask for an advance before starting a new client\'s project and keep every approval in writing (email is fine).'] } });

  type({ id: 'webdev', ex: { services: 'Home, About, Services, Doctors and Contact pages\nOnline appointment booking with email and SMS confirmation\nAdmin panel to manage doctors\' time slots\nMobile-friendly design and basic on-page SEO' }, g: 2, name: 'Website or app development agreement', kw: 'website web app mobile application software development developer coding design ecommerce',
    title: 'WEBSITE AND APP DEVELOPMENT AGREEMENT', A: 'Client', B: 'Developer', ip: 'client', indem: 'mutual', ipIndem: true,
    blurb: 'Building a website or app: specification, milestones, testing and acceptance, bug-fix period, hosting, source code and passwords.',
    q: ['project', 'platform', 'services', 'milestones', 'start', 'revisions', 'support', 'hosting', 'payterms', 'tax'],
    opts: { conf: 1, ip: 1, indemnity: 1, liability: 1, data: 1, latefee: 1, termination: 1, force: 1, dispute: 1 },
    defs: function (x) { return [['Project', x.v('project')], ['Specification', 'the features and pages listed in Schedule 1, as changed in writing'], ['Acceptance', 'the Client\'s written approval after testing'], ['Defect', 'a failure of the delivered software to work as the Specification says'], ['Launch', 'the date the website or app is made live for the public']]; },
    rec: function (x) { return ['The Developer designs and builds websites and applications.', L('The Client wishes the Developer to build the Project using {0}, on the terms below.', [x.v('platform')])]; },
    body: function (x) { return [
      sec('Scope', [L('The Developer will design, build, test and launch the Project to the Specification in Schedule 1, starting on {0}, following the milestones in Schedule 2.', [x.d('start')]),
        'Anything not in the Specification (new features, pages, integrations, content writing, data entry) is a change and will be quoted separately before work starts.',
        L('Design work includes {0} rounds of changes to each design before approval.', [x.has('revisions') ? x.n('revisions') : '[number of]'])]),
      sec('Client materials', ['The Client will supply text, images, logos, product data and access to accounts by the dates agreed. The Client confirms it has the right to use everything it supplies and will indemnify the Developer against claims about it.']),
      sec('Testing and acceptance', ['When a milestone is ready, the Developer will make it available on a test site. The Client will test it within 10 days and either accept it or list Defects in writing.',
        'The Developer will fix listed Defects and resubmit. Requests for things not in the Specification are changes, not Defects. A milestone is accepted when the Client approves it, uses it live, or does not respond within 10 days.']),
      sec('Hosting, domain and accounts', [(x.raw('hosting') || '').indexOf('Client') === 0 ? 'The domain name, hosting, app store developer accounts, payment gateway and third-party service accounts will be held in the Client\'s name and paid by the Client. The Developer will have access only as needed and will return it at the end.' : 'The Developer will arrange hosting and the domain at the Client\'s cost, but the domain name and app store listings will be registered in the Client\'s name. The Developer will transfer hosting to the Client\'s account on request.',
        'The Developer will follow reasonable security practice: strong passwords, up-to-date software at launch, HTTPS, regular backups during the project, and no hidden access after handover.']),
      sec('Launch, handover and source code', ['After final payment the Developer will hand over the complete source code, database, design files, and all admin usernames and passwords, and will change or remove its own access if asked.',
        'Third-party components (themes, plugins, libraries, fonts, APIs) remain under their own licences. The Developer will list paid licences and their renewal costs.']),
      sec('Bug-fix period and maintenance', [L('For {0} after Launch the Developer will fix Defects reported in writing free of charge, within 3 working days for critical Defects that stop the site or app working and 10 working days for others.', [x.mo('support')]),
        'This does not cover problems caused by changes made by others, hosting failures, third-party updates, or new features. Maintenance after the bug-fix period can be agreed separately, for example under a maintenance agreement.']),
      sec('Fees', [L('The Client will pay the amounts in Schedule 2 as each milestone falls due.{0}', [x.taxed()]), invoiceLine(x, 'Client', 'Developer')])]; },
    sched: function (x) { var ms = msRows(x); var rows = ms.rows.length ? ms.rows : [['1', '[milestone]', '[when due]', '[amount]']]; if (ms.total) rows.push(['', 'Total', '', x.ms(ms.total)]);
      return [{ h: 'Schedule 1: Specification', items: [L('Platform: {0}', [x.v('platform')]), sub('Features and pages:', x.lines('services').length ? x.lines('services') : ['[list pages, features, integrations, devices and browsers supported]'])] }, { h: 'Schedule 2: Milestones and payments', items: [tbl(['#', 'Milestone', 'When due', 'Amount'], rows)] }]; },
    notes: { all: ['Write the Specification in detail (pages, features, who can log in, payments, languages, devices). It decides what counts as a defect and what is a paid change.'] } });

  type({ id: 'amc', g: 2, name: 'Maintenance agreement (AMC)', kw: 'amc annual maintenance contract service repair ac lift generator computer equipment upkeep',
    title: 'ANNUAL MAINTENANCE AGREEMENT', A: 'Client', B: 'Service Provider', ip: 'owner', indem: 'B',
    blurb: 'Annual maintenance of equipment: preventive visits, breakdown response time, parts, exclusions and service reports.',
    q: ['equip', 'coverage', 'visits', 'response', 'start', ['term', 'Term (months)'], ['fee', 'AMC charge for the term'], 'payterms', 'tax'],
    opts: { liability: 1, indemnity: 0, latefee: 0, termination: 1, force: 1, dispute: 1 },
    defs: function (x) { return [['Equipment', 'the items listed in Schedule 1 at the Client\'s premises'], ['Preventive Maintenance', 'scheduled inspection, cleaning, adjustment and testing'], ['Breakdown', 'a failure that stops the Equipment working normally'], ['Response Time', x.has('response') ? L('{0} hours', [x.n('response')]) : '[ ] hours']]; },
    rec: function (x) { return ['The Client owns the Equipment.', 'The Service Provider has the skills, staff and spares to maintain it and will do so on the terms below.']; },
    body: function (x) { var comp = (x.raw('coverage') || '').indexOf('Comprehensive') === 0; return [
      sec('Term and charge', [L('This Agreement covers {0} from {1}. The charge for the term is {2}, payable in advance within {3} of the invoice.{4}', [x.mo('term'), x.d('start'), x.m('fee'), x.days('payterms'), x.taxed()]),
        'The Agreement may be renewed for a further term at a charge agreed in writing before the end.']),
      sec('Preventive maintenance', [L('The Service Provider will make {0} Preventive Maintenance visits during the term at roughly equal intervals, on dates agreed with the Client, following the manufacturer\'s recommended checklist.', [x.has('visits') ? x.n('visits') : '[number of]'])]),
      sec('Breakdown calls', ['The Client may report a Breakdown by phone, email or message at any time. The Service Provider will attend within the Response Time during working hours (and as soon as possible outside them), and will try to restore working within 48 hours.',
        'There is no limit on the number of breakdown calls for covered faults.']),
      sec('Parts', [comp ? 'This is a comprehensive agreement: spare parts needed for repair are included in the charge, except consumables (filters, gas top-ups, belts, batteries, bulbs, ink, toner) and the parts excluded in Schedule 1.' : 'This is a non-comprehensive (labour only) agreement: spare parts are charged extra at the Service Provider\'s quoted price, which the Client will approve before replacement.',
        'Replaced parts will be genuine or of equal quality and carry a warranty of at least 3 months. Old parts will be returned to the Client if asked.']),
      sec('Exclusions', [sub('The agreement does not cover repairs needed because of:', ['misuse, accident, fire, flood, pests or abnormal power supply;', 'moving the Equipment, or repairs or changes by anyone else;', 'consumables and cosmetic parts such as covers and panels; or', 'equipment beyond economic repair, which the Service Provider will report in writing with reasons.'])]),
      sec('Service reports', ['After every visit the technician will leave a signed service report showing the work done, parts used and any recommendations. The Client will sign it or note any problem.']),
      sec('Staff and safety', ['The Service Provider\'s technicians will be trained, carry identity cards, follow the Client\'s safety and security rules, and be the Service Provider\'s responsibility.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Equipment covered', items: [tbl(['#', 'Equipment, make, model, serial number and location', 'Parts excluded'], (x.lines('equip').length ? x.lines('equip') : ['[equipment]']).map(function (l, i) { return [String(i + 1), l, '']; }))] }]; },
    notes: { all: ['Inspect the equipment with the service provider at the start and list anything already faulty, so it is not argued about later.'] } });

  type({ id: 'event', ex: { services: 'Buffet dinner for the guests: 2 starters, 4 main dishes, 2 desserts (menu attached)\nService staff of 15 in uniform from 7 pm to 11:30 pm\nCrockery, cutlery, counters and buffet decoration' }, g: 2, name: 'Event or vendor booking agreement', kw: 'event booking vendor caterer decorator dj photographer wedding party venue tent banquet',
    title: 'EVENT SERVICES BOOKING AGREEMENT', A: 'Client', B: 'Vendor', ip: 'owner', fmRefund: true,
    blurb: 'Booking a caterer, decorator, photographer or other event vendor: services, payments, guest numbers, cancellation scale.',
    q: ['eventname', 'eventdate', 'venue', 'guests', 'services', 'evfee', 'advance', 'balancedate', 'paymode', 'tax'],
    opts: { force: 1, liability: 1, dispute: 0, indemnity: 0 },
    defs: function (x) { return [['Event', L('{0} on {1}', [x.v('eventname'), x.d('eventdate')])], ['Venue', x.v('venue')], ['Services', 'the services listed in Schedule 1'], ['Price', 'the total price in the Price clause']]; },
    rec: function (x) { return ['The Vendor provides event services.', 'The Client wishes to book the Vendor for the Event and the Vendor agrees, on the terms below.']; },
    body: function (x) { var bal = x.n('evfee') - x.n('advance'); return [
      sec('Booking', ['The Vendor will provide the Services at the Venue for the Event. On receiving the advance, the Vendor will hold the date for the Client and will not take another booking that clashes with it.']),
      sec('Price and payments', [L('The total price is {0} for {1}.{2}', [x.m('evfee'), x.has('guests') ? L('{0} guests', [x.n('guests')]) : 'the number of guests stated', x.taxed()]),
        L('The Client has paid an advance of {0}. The balance{1}{2} is payable by {3}, by {4}.', [x.m('advance'), x.has('evfee') && x.has('advance') && bal > 0 ? L(' of {0}', [x.m(bal)]) : '', bal > 0 ? '' : '', x.d('balancedate'), x.raw('paymode') || 'bank transfer']),
        'The Client will confirm the final guest count at least 7 days before the Event. Charges per guest for numbers above the count are payable at the same rate; numbers below the count are charged at the confirmed count.']),
      sec('What each Party will do', ['The Vendor will arrive in time to set up, bring the staff, equipment and materials needed, keep its staff in clean uniforms and good conduct, and clear its materials from the Venue after the Event.',
        'The Client will arrange access to the Venue, power and water supply, and the Venue\'s permission for the Vendor\'s work, and will tell the Vendor the Venue\'s rules.',
        'Permissions for music, sound, fireworks, liquor and similar items are the responsibility of the Party that provides them, as listed in Schedule 1.']),
      sec('Changes', ['Changes to the Services (menu, decor, timings) may be made up to 10 days before the Event, and later only if the Vendor can reasonably manage them. Changes may change the price, agreed in writing.']),
      sec('Cancellation and postponement', [sub('If the Client cancels, the Vendor will refund from the amounts paid:', ['cancellation more than 60 days before the Event: everything paid, less 10% of the price;', 'cancellation 30 to 60 days before: everything paid, less 25% of the price;', 'cancellation less than 30 days before: everything paid, less 50% of the price, and less than 7 days before: no refund,']),
        'but if the Vendor books the same date for another client, it will refund the amounts paid less only its actual costs. Refunds are made within 15 days.',
        'One postponement to a date within 6 months is allowed without charge if requested at least 30 days before the Event and the Vendor is available.',
        'If the Vendor cancels or fails to turn up, it will refund everything paid within 7 days and pay the Client\'s reasonable extra cost of a replacement vendor, up to the price.']),
      sec('Damage and safety', ['The Vendor is responsible for damage its staff cause to the Venue, and the Client for damage its guests cause to the Vendor\'s equipment. The Vendor will follow food safety, fire and electrical safety practice.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Services', items: [sub('', x.lines('services').length ? x.lines('services') : ['[menu, decor items, hours of service, staff numbers, equipment, permissions]'])] }]; },
    notes: { IN: ['Caterers need an FSSAI registration or licence; keep a copy. Stamp duty on the agreement depends on the state.'], all: ['Attach the menu, decor photos or a shot list to Schedule 1 so there is no argument about what was promised.'] } });

  type({ id: 'content-licence', g: 2, name: 'Content licence agreement (photos, writing, music, video)', kw: 'content licence license copyright photos images music video article software usage rights royalty',
    title: 'CONTENT LICENCE AGREEMENT', A: 'Licensor', B: 'Licensee', ip: 'owner', indem: 'A', ipIndem: true,
    blurb: 'Letting someone use your photos, writing, music or video without selling the copyright: scope, media, period, fee and credit.',
    q: ['work', 'media', 'territory', 'ltype', 'lterm', ['fee', 'Licence fee'], 'payterms', 'creditline', 'start', 'tax'],
    opts: { indemnity: 1, liability: 1, termination: 1, dispute: 1, latefee: 0 },
    defs: function (x) { return [['Work', x.v('work')], ['Permitted Media', x.v('media')], ['Territory', x.v('territory', 'territory, for example worldwide')], ['Licence Period', L('{0} from {1}', [x.mo('lterm'), x.d('start')])]]; },
    rec: function (x) { return ['The Licensor owns the copyright in the Work, or has the right to license it.', 'The Licensee wishes to use the Work in the Permitted Media, and the Licensor agrees to license it, on the terms below.']; },
    body: function (x) { var ex = x.raw('ltype') === 'Exclusive'; return [
      sec('Licence', [L('The Licensor grants the Licensee a {0}, non-transferable licence to reproduce, publish and display the Work in the Permitted Media in the Territory during the Licence Period.', [ex ? 'exclusive' : 'non-exclusive']),
        ex ? 'During the Licence Period the Licensor will not license the Work to anyone else for the Permitted Media in the Territory, but may still use it in its own portfolio.' : 'The Licensor may license the Work to others and use it itself.',
        'The Licensee may resize, crop or format the Work to fit the Permitted Media, but will not otherwise alter, distort or combine it in a way that harms the Licensor\'s reputation without written approval.',
        'Any use not expressly allowed (for example other media, merchandise, resale, use as a trade mark, or training of software models) needs a separate written licence. Copyright and all other rights stay with the Licensor.']),
      sec('Fee', [L('The Licensee will pay a licence fee of {0} within {1} of the Licensor\'s invoice.{2}', [x.m('fee'), x.days('payterms'), x.taxed()]), 'The licence starts only when the fee is paid in full.']),
      sec('Credit', [L('Where practical, the Licensee will show the credit "{0}" next to the Work or in the credits. Accidental omission is not a breach if corrected within 7 days of notice.{1}', [x.v('creditline'), x.inIN ? ' The Licensor\'s moral rights under section 57 of the Copyright Act, 1957 are not affected.' : ''])]),
      sec('Licensor\'s promises', ['The Licensor confirms that it owns or controls the rights licensed, that the Work is original, and that, to its knowledge, using it as allowed will not infringe anyone\'s copyright, privacy or other rights, and that it has releases from any identifiable people or private property shown where needed.']),
      sec('End of the licence', ['When the Licence Period ends or this Agreement ends early, the Licensee will stop new use of the Work and, within 30 days, remove it from websites and digital media under its control. Printed material already distributed need not be recalled.'])]; },
    notes: { IN: ['Under section 19 of the Copyright Act, 1957, if a licence does not say the period or territory, it may be treated as 5 years and India only, so keep those fields filled in.'], all: ['Attach thumbnails or a list of file names of the Work to the agreement.'] } });

  type({ id: 'influencer', g: 2, name: 'Influencer or brand collaboration agreement', kw: 'influencer creator brand collaboration sponsored post instagram youtube social media marketing paid partnership',
    title: 'INFLUENCER COLLABORATION AGREEMENT', A: 'Brand', B: 'Creator', ip: 'owner', indem: 'mutual', nc: 'B',
    blurb: 'Sponsored posts: deliverables, approval, ad disclosure rules, usage rights, exclusivity days and payment.',
    q: ['platforms', 'deliverables', 'start', ['fee', 'Fee for all content'], 'advance', 'payterms', 'usage', 'excl', 'tax'],
    opts: { conf: 1, termination: 1, liability: 1, dispute: 1, indemnity: 0, latefee: 0 },
    defs: function (x) { return [['Content', 'the posts, videos and stories listed in Schedule 1'], ['Platforms', x.v('platforms')], ['Products', 'the Brand\'s products or services featured in the Content'], ['Campaign Start', x.d('start')]]; },
    rec: function (x) { return ['The Creator publishes content on the Platforms.', 'The Brand wishes the Creator to make and publish sponsored Content about the Products, on the terms below.']; },
    body: function (x) { return [
      sec('Content and timing', ['The Creator will make and publish the Content listed in Schedule 1 on its own accounts on the Platforms, from the Campaign Start, on the dates agreed in writing.',
        'The Brand will give a brief with key messages, do\'s and don\'ts and, where relevant, free samples of the Products at least 7 days before the first posting date. The Creator keeps its own style and voice.']),
      sec('Approval', ['The Creator will share a draft of each piece of Content at least 3 days before posting. The Brand may ask for factual corrections and up to two rounds of reasonable changes within 2 days; silence counts as approval.',
        'The Creator will not make any claim about the Products that the Brand has not confirmed or that the Creator cannot honestly support, and the Brand is responsible for the accuracy of claims it supplies.']),
      sec('Advertising disclosure', [L('Every piece of Content will clearly show that it is paid or sponsored, at the start and in a way viewers cannot miss, using the Platform\'s paid-partnership label and words such as "Ad" or "Sponsored", as required by {0}.', [x.inIN ? 'the ASCI guidelines for influencer advertising and the Consumer Protection Act, 2019 guidelines' : x.inUK ? 'the CAP Code and Competition and Markets Authority guidance' : x.inUS ? 'the FTC Endorsement Guides' : x.inAE ? 'the UAE media regulator\'s advertising rules (and any advertiser permit required)' : 'the advertising rules of the countries where the Content is aimed'])]),
      sec('Keeping the Content live', ['The Creator will keep each post live for at least 90 days and will not edit its sponsored message without the Brand\'s approval.',
        'The Creator will send screenshots of reach, views and engagement for each post within 7 days of posting and again after 30 days.']),
      sec('Fee', [L('The Brand will pay {0} for all the Content.{1}{2}', [x.m('fee'), x.taxed(), x.has('advance') ? L(' {0} is payable on signing and the rest within {1} after the last piece is posted.', [x.m('advance'), x.days('payterms')]) : L(' It is payable within {0} after the last piece is posted.', [x.days('payterms')])]),
        'Products sent for the Content may be kept by the Creator unless the Brand asks for their return in writing.']),
      sec('Use of the Content by the Brand', [L('The Creator keeps ownership of the Content. The Brand may share and repost it on its own social media accounts and website, crediting the Creator, for {0} from posting. Paid advertising using the Content, or other uses, needs written agreement and may need an extra fee.', [x.mo('usage')])]),
      sec('Exclusivity and conduct', [L('For {0} after each post the Creator will not publish paid content for a directly competing product without the Brand\'s consent.', [x.days('excl', 'number of')]),
        'Each Party will avoid conduct that brings the other into public disrepute. If the Creator publicly behaves in a way that seriously harms the Brand\'s reputation, the Brand may end this Agreement and ask for removal of unpublished and future Content, paying for Content already posted.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Content', items: [L('Platforms: {0}', [x.v('platforms')]), sub('', x.lines('deliverables').length ? x.lines('deliverables') : ['[each post, format, length and date]'])] }]; },
    notes: { IN: ['Influencers in India must disclose material connections clearly (ASCI guidelines, Department of Consumer Affairs "Endorsements Know-hows!"). TDS under section 194R may apply to free products above the yearly limit.'], AE: ['Influencers advertising in the UAE need an advertiser permit from the UAE Media Council.'] } });

  type({ id: 'subcontract', ex: { services: 'Gypsum false ceiling in 2,400 sq ft of office area\nElectrical wiring and light fittings in the ceiling\nHandover of the area ready for painting' }, g: 2, name: 'Subcontractor agreement', kw: 'subcontractor subcontract contractor outsourcing back to back site work labour contractor',
    title: 'SUBCONTRACT AGREEMENT', A: 'Contractor', B: 'Subcontractor', ip: 'client', indem: 'B', portfolio: false,
    blurb: 'Passing part of your work to a subcontractor: same standards as your client contract, no direct dealing with the client, safety, defects.',
    q: ['mainproject', 'services', 'start', ['term', 'Duration (months)'], 'fee', 'feebasis', 'payterms', 'tax', 'notice'],
    opts: { conf: 1, nonsolicit: 1, ip: 1, indemnity: 1, liability: 1, termination: 1, force: 1, dispute: 1 },
    defs: function (x) { return [['Main Contract', L('the Contractor\'s contract with its client for {0}', [x.v('mainproject')])], ['End Client', 'the Contractor\'s client under the Main Contract'], ['Subcontract Work', 'the work described in Schedule 1']]; },
    rec: function (x) { return [L('The Contractor has a Main Contract for {0}.', [x.v('mainproject')]), 'The Contractor wishes the Subcontractor to carry out part of that work, and the Subcontractor agrees, on the terms below.']; },
    body: function (x) { return [
      sec('Subcontract Work', [L('The Subcontractor will carry out the Subcontract Work from {0} and complete it within {1}, or by the dates in any programme the Contractor issues.', [x.d('start'), x.mo('term')]),
        'The Subcontractor will do the work to the standards, specifications and timelines of the Main Contract that the Contractor has shared with it, and follow the Contractor\'s reasonable instructions.']),
      sec('Relationship with the End Client', ['The Subcontractor works for the Contractor only. It will not deal directly with the End Client about pricing, extra work or future work, and will refer any request from the End Client to the Contractor.',
        'The Subcontractor will not, for 12 months after this Agreement ends, offer the End Client the same kind of work directly, to the extent the law allows.']),
      sec('Price and payment', [L('The Contractor will pay {0}', [feeLine(x)]), invoiceLine(x, 'Contractor', 'Subcontractor'),
        'The Contractor may hold back payment for work that is defective or not done, after telling the Subcontractor why in writing.']),
      sec('Site, safety and staff', [L('The Subcontractor will provide its own trained workers, tools and safety equipment, follow all safety, labour and environmental laws and site rules, and is responsible for its workers\' wages, {0}.', [x.inIN ? 'PF, ESI, labour registration and insurance' : 'taxes, insurance and work permits']),
        'The Subcontractor will keep public liability and workers\' compensation insurance (or the equivalent) for the duration.']),
      sec('Defects', ['The Subcontractor will fix, at its cost, any defect in its work notified within 12 months after completion (or the defects period of the Main Contract, if longer).']),
      sec('Variations and delays', ['Extra or changed work is payable only if ordered by the Contractor in writing. If the Subcontractor is late through its own fault, it will pay the Contractor\'s resulting costs, including any delay damages the End Client charges for that delay, up to 10% of the subcontract price.'])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Subcontract Work', items: [sub('', x.lines('services').length ? x.lines('services') : ['[scope, quantities, specifications, programme]'])] }]; },
    notes: { UK: ['For construction work, the Housing Grants, Construction and Regeneration Act 1996 gives rights to stage payments, payment notices and adjudication, and makes "pay when paid" clauses ineffective in most cases.'], all: ['Attach the relevant parts of the main contract (specifications, safety rules) to Schedule 1.'] } });

  /* ===== Jobs and people */
  type({ id: 'employment', g: 3, name: 'Employment or appointment agreement', kw: 'employment contract appointment letter employee employer job offer staff salary probation',
    title: 'EMPLOYMENT AGREEMENT', A: 'Employer', B: 'Employee', ip: 'employer', nc: 'B', rel: false, ownEnd: true, assign: false,
    blurb: 'Hiring an employee: job, pay, hours, probation, leave, rules, confidentiality, notice, return of property.',
    q: ['position', 'reportto', 'workplace', 'start', 'hours', 'salary', 'probation', 'leave', 'notice'],
    opts: { conf: 1, ip: 1, noncompete: 0, nonsolicit: 1, data: 1, dispute: 0 },
    defs: function (x) { return [['Start Date', x.d('start')], ['Salary', 'the gross monthly salary in the Salary clause'], ['Policies', 'the Employer\'s written rules and policies as given to the Employee and updated from time to time']]; },
    rec: function (x) { return [L('The Employer wishes to employ the Employee as {0}, and the Employee accepts, on the terms below.', [x.v('position')])]; },
    body: function (x) { return [
      sec('Job and start', [L('The Employee is employed as {0}, reporting to the {1}, from the Start Date.{2}', [x.v('position'), x.v('reportto'), x.inUS ? ' Employment is "at will": either Party may end it at any time with or without cause, subject to the notice requested below as a courtesy and to applicable law.' : '']),
        'The Employee will perform the duties of the job and other reasonable duties the Employer assigns that suit the Employee\'s skills, follow lawful instructions and the Policies, and work honestly and diligently.',
        L('The place of work is {0}.', [x.v('workplace')])]),
      sec('Probation', [x.n('probation') > 0 ? L('The first {0} are a probation period, which the Employer may extend once by up to 3 months by written notice. During probation either Party may end the employment by 15 days\' written notice (or salary instead). The Employer will confirm the employment in writing after successful probation.', [x.mo('probation')]) : 'There is no probation period.']),
      sec('Working hours', [L('Normal working hours are {0}. The Employee may need to work reasonable extra hours when the work requires; overtime will be paid {1}.', [x.v('hours'), x.inIN || x.inAE ? 'where the law requires it' : x.inUS ? 'where the Employee is non-exempt under the Fair Labor Standards Act and state law' : 'where the law requires it or the Policies provide'])]),
      sec('Salary and deductions', [L('The Employer will pay a gross salary of {0} per month, by bank transfer by the 7th of the following month{1}. The salary break-up (if any) is in Schedule 1.', [x.m('salary'), x.inAE ? ' through the Wage Protection System' : '']),
        x.inIN ? 'The Employer will deduct the Employee\'s share of provident fund and ESI where they apply, professional tax and income tax at source, and pay the employer\'s contributions, as the law requires.' : x.inUK ? 'The Employer will deduct income tax and National Insurance under PAYE, enrol the Employee in a workplace pension where the law requires, and pay statutory sick, maternity, paternity and other statutory pay as applicable.' : x.inUS ? 'The Employer will withhold federal, state and local taxes and make the deductions required by law or authorised by the Employee in writing. Benefits, if any, are as described in the Employer\'s plan documents.' : x.inAE ? 'On leaving after at least one year\'s continuous service, the Employee will be entitled to end-of-service gratuity as the UAE Labour Law provides (unless covered by an approved alternative savings scheme).' : 'The Employer will make the deductions the law requires.',
        'The salary will be reviewed once a year. Any increase is at the Employer\'s discretion unless the law or a policy says otherwise.']),
      sec('Leave and holidays', [L('The Employee is entitled to {0} days of paid leave a year (in addition to public holidays observed by the Employer), taken with prior approval, and to sick leave and other leave as the Policies and the law provide.{1}', [x.has('leave') ? x.n('leave') : '[ ]', x.inUK ? ' This includes at least the statutory minimum of 5.6 weeks a year for a full-time worker (pro rata for part-time), which may include bank holidays.' : x.inAE ? ' This includes at least the annual leave required by the UAE Labour Law (30 days a year after one year\'s service).' : ''])]),
      sec('Conduct and other work', ['The Employee will not, during employment, work for or have an interest in any other business without the Employer\'s written consent, which will not be unreasonably refused for work that does not compete or affect the Employee\'s duties.',
        'The Employee will not accept gifts or payments from customers or suppliers beyond small customary gifts, and will report any conflict of interest.']),
      sec('Ending the employment', [L('After probation, either Party may end the employment by giving {0} of written notice{1}, or salary for the notice period instead. The Employer may ask the Employee not to attend work during notice (garden leave) on full pay.', [x.days('notice'), x.inUK ? ' or, if longer, the statutory minimum notice' : '']),
        L('The Employer may end the employment without notice for serious misconduct, such as theft, fraud, violence, harassment or serious breach of the Policies, after a fair inquiry {0}.', [x.inIN ? 'following the principles of natural justice and the standing orders that apply' : 'as the law requires']),
        L('On leaving, the Employee will return all property, documents, devices, ID cards and data of the Employer, and hand over work. The Employer will settle final dues within the time the law requires and give {0}.', [x.inIN ? 'an experience and relieving letter' : 'a letter confirming the dates and role of employment'])])]; },
    sched: function (x) { return [{ h: 'Schedule 1: Salary details', items: [tbl(['Component', 'Monthly amount'], [['Basic salary', ''], [x.inIN ? 'House rent allowance' : 'Allowances', ''], ['Other allowances', ''], ['Gross salary', x.has('salary') ? x.ms('salary') : '']])] }]; },
    notes: { IN: ['Many states need appointment letters or standing orders under the Industrial Employment (Standing Orders) Act or shop and establishment law; the new Labour Codes (in force from 21 November 2025) also require an appointment letter for every employee.', 'Non-compete clauses after employment ends are generally void in India (section 27, Contract Act); confidentiality and non-solicitation are more likely to be upheld.'],
      UK: ['Employers must give a written statement of particulars on or before day one (Employment Rights Act 1996, section 1). Check the Employment Rights Act 2025 changes as they come into force.'],
      US: ['Several states (including California, Minnesota, North Dakota and Oklahoma) ban most employee non-competes; others limit them by salary or require notice.'],
      AE: ['Employment in the mainland UAE must be on the MOHRE standard contract; free zones have their own rules. Use this draft to agree terms, then complete the official contract.'] } });

  type({ id: 'internship', g: 3, name: 'Internship agreement', kw: 'internship intern trainee student apprentice stipend training certificate',
    title: 'INTERNSHIP AGREEMENT', A: 'Organisation', B: 'Intern', ip: 'employer', rel: false, ownEnd: true, assign: false,
    blurb: 'A learning placement: duration, stipend, mentor, hours, confidentiality, work ownership, certificate.',
    q: ['position', 'mentor', 'workplace', 'start', ['term', 'Duration (months)'], 'stipend', 'hours', 'notice'],
    opts: { conf: 1, ip: 1, data: 0 },
    defs: function (x) { return [['Internship', 'the training placement described in this Agreement'], ['Mentor', x.v('mentor')], ['Internship Period', L('{0} from {1}', [x.mo('term'), x.d('start')])]]; },
    rec: function (x) { return ['The Organisation offers practical training to students and new graduates.', L('The Intern wishes to gain experience as {0}, and the Organisation agrees to offer an internship, on the terms below.', [x.v('position')])]; },
    body: function (x) { return [
      sec('The internship', [L('The Organisation offers the Intern an internship as {0} for the Internship Period, at {1}, under the guidance of the Mentor.', [x.v('position'), x.v('workplace')]),
        'The main purpose is learning. The Organisation will give the Intern meaningful tasks, regular feedback and a chance to learn its work, and will not use the Intern to replace a regular employee.',
        L('Working hours are {0}. The Intern will be allowed time off for examinations and college requirements with notice.', [x.v('hours')])]),
      sec('Stipend', [x.n('stipend') > 0 ? L('The Organisation will pay a stipend of {0} per month, by bank transfer by the 7th of the following month, after any tax deduction the law requires.', [x.m('stipend')]) : L('The internship is unpaid, except for reimbursement of pre-approved work expenses.{0}', [x.inUK ? ' The Parties have checked that the Intern is not a "worker" entitled to the National Minimum Wage (for example because the internship is a required part of a higher-education course, or the Intern only shadows).' : x.inUS ? ' The Parties have checked that the internship meets the primary-beneficiary test of the US Department of Labor, so that no wages are due.' : ''])]),
      sec('Conduct', ['The Intern will follow the Organisation\'s rules, attend regularly, inform the Mentor in advance of any absence and treat information and colleagues with respect.']),
      sec('Certificate and ending', ['On completing the internship, the Organisation will give the Intern a certificate stating the period and the work done, and, if asked, a fair reference.',
        L('Either Party may end the internship early by giving {0} of written notice. The Organisation may end it at once for serious misconduct. The Intern will return all property and data on leaving.', [x.days('notice')]),
        'The internship does not create a promise of employment, though the Organisation may offer a job separately.'])]; },
    notes: { IN: ['Interns are generally not "workers" under Indian labour law if the placement is genuine training; apprentices under the Apprentices Act, 1961 have separate contracts and stipend rules.'], all: ['Check whether your college needs its own internship form signed as well.'] } });

  type({ id: 'domestic', g: 3, name: 'Domestic worker or caretaker agreement', kw: 'domestic worker maid cook driver nanny caretaker housekeeper household help servant ayah',
    title: 'DOMESTIC WORK AGREEMENT', A: 'Householder', B: 'Worker', rel: false, ownEnd: true, assign: false, noStamp: false,
    blurb: 'Help at home: duties, hours, weekly off, salary, leave, food and stay for live-in staff, dignity and safety, notice.',
    q: ['duties', 'livein', 'hours', 'weeklyoff', 'salary', 'leave', 'start', 'notice'],
    opts: { conf: 0 },
    defs: function (x) { return [['Home', 'the Householder\'s home at the address above'], ['Duties', 'the work listed in Schedule 1']]; },
    rec: function (x) { return ['The Householder wishes to employ the Worker in the Home, and the Worker agrees, on the terms below. The Worker confirms they are at least 18 years old.']; },
    body: function (x) { var live = (x.raw('livein') || '').indexOf('Yes') === 0; return [
      sec('Work and hours', [L('From {0} the Worker will do the Duties in Schedule 1. Any other regular work will be agreed and may need extra pay.', [x.d('start')]),
        L('Working hours are {0}, with reasonable breaks for meals and rest. The Worker will have {1} off every week, and at least 8 hours of continuous rest each night{2}.', [x.v('hours'), x.v('weeklyoff'), live ? '' : ''])]),
      sec('Pay', [L('The Householder will pay {0} per month by the 7th of the following month, by bank transfer or UPI (or in cash against a signed receipt), and will give a simple pay record.', [x.m('salary')]),
        'Extra work on the weekly off or beyond the agreed hours will be paid extra at a rate agreed in advance.',
        'Pay will be reviewed every year on the anniversary of the start date.']),
      sec('Leave', [L('The Worker is entitled to {0} days of paid leave a year, plus the main public holidays agreed with the Householder, and paid sick leave of up to 7 days a year on telling the Householder.', [x.has('leave') ? x.n('leave') : '[ ]'])]),
      live ? sec('Food and stay', ['The Householder will provide the Worker with a clean, private and lockable room or space, access to a toilet and bathroom, bedding, and three nutritious meals a day, free of charge and without deduction from pay.']) : null,
      sec('Dignity and safety', ['The Householder will treat the Worker with respect, will not verbally or physically abuse, harass or confine the Worker, and will not keep the Worker\'s identity documents, phone or wages.',
        'The Householder will give safety equipment such as gloves for cleaning chemicals, and pay for first aid and reasonable medical treatment for injuries suffered at work.',
        'The Worker will do the work honestly and carefully, keep the household\'s private matters confidential, and not let anyone into the Home without the Householder\'s permission.']),
      sec('Ending', [L('Either Party may end this Agreement by giving {0} of notice (or the pay for the notice period instead). The Householder may end it without notice for theft, violence or similar serious misconduct.', [x.days('notice')]),
        L('On leaving, the Worker will be paid all wages and leave due up to the last day{0}.', [live ? ', and given reasonable time (at least 3 days) to move out' : ''])])].filter(Boolean); },
    sched: function (x) { return [{ h: 'Schedule 1: Duties', items: [sub('', x.lines('duties').length ? x.lines('duties') : ['[list each task]'])] }]; },
    notes: { IN: ['Some states have minimum wages and welfare boards for domestic workers (for example Kerala, Tamil Nadu, Maharashtra, Karnataka). The worker can register on the e-Shram portal for social security benefits. Do tenant-style police verification where your city requires it.'],
      UK: ['A domestic worker is an employee: the householder must pay at least the National Minimum Wage (unless the family-worker exemption applies), give a written statement, holiday pay and may need to run payroll.'],
      AE: ['Domestic workers in the UAE are covered by Federal Decree-Law No. 9 of 2022 and must be hired through licensed recruiters (Tadbeer) on the MOHRE standard contract.'], US: ['Household employers may owe "nanny tax" (Social Security, Medicare, federal unemployment) above the yearly threshold and must follow state minimum wage and domestic workers\' bill of rights laws where they exist.'] } });

  /* ===== Confidentiality */
  function ndaCore(x, mutual) {
    var disc = mutual ? 'a Party disclosing information (the "Discloser")' : 'the Discloser', rec = mutual ? 'the Party receiving it (the "Recipient")' : 'the Recipient';
    return [
      sec('Confidential Information', [L('"Confidential Information" means all information, in any form (written, oral, electronic, visual or by samples), that {0} gives to {1} or that the Recipient learns about the Discloser, in connection with the Purpose, including business plans, customer and supplier details, prices, finances, product designs, software, source code, formulas, processes and know-how, whether or not marked "confidential", if a reasonable person would understand it to be confidential.', [disc, rec]),
        L('Confidential Information also includes the fact that the Parties are in discussions and the terms discussed{0}.', [mutual ? '' : ''])]),
      sec('Duties of the Recipient', [sub('The Recipient will:', ['use Confidential Information only for the Purpose;', L('not disclose it to anyone except its directors, employees, professional advisers{0} who need to know it for the Purpose and who are bound by duties of confidentiality at least as strict as this Agreement;', [mutual ? ' and, with the Discloser\'s consent, contractors' : '']), 'protect it with at least the care it uses for its own confidential information, and never less than reasonable care;', 'not copy it except as needed for the Purpose, and not reverse engineer, decompile or analyse any samples or software; and', 'tell the Discloser at once if it learns of any unauthorised use or disclosure, and help to stop it.']),
        'The Recipient is responsible for any breach of this Agreement by the people it shares Confidential Information with.']),
      sec('Exceptions', [sub('The duties do not apply to information that the Recipient can show:', ['is or becomes public other than through a breach of this Agreement;', 'was lawfully known to it, without a duty of confidence, before the Discloser gave it;', 'it received lawfully from someone else who was free to disclose it; or', 'it developed independently without using the Confidential Information.']),
        'The Recipient may disclose Confidential Information if required by law, a court, a regulator or a stock exchange, but only to the extent required and, where the law allows, after telling the Discloser so that it can seek protection.']),
      sec('Return and destruction', ['When the Discloser asks, or when the Purpose ends, the Recipient will promptly return or destroy all Confidential Information and copies, and confirm this in writing. It may keep one copy where the law requires or in automatic backups that are not routinely accessed, which stay confidential.']),
      sec('No licence, no warranty, no obligation', ['All Confidential Information remains the property of the Discloser. Nothing in this Agreement gives any licence or right to it, except to use it for the Purpose.',
        'Confidential Information is given "as is" without any promise of accuracy or completeness.',
        'Neither Party is obliged to enter into any further agreement or transaction. Each may end discussions at any time.']),
      sec('Term', [L('This Agreement starts on {0}. Information may be shared under it for {1}, and the duties of confidentiality continue for {2} after that period ends, and for trade secrets for as long as they remain trade secrets.', [x.d('start'), x.mo('term'), x.has('confyears') ? L('{0} years', [x.n('confyears')]) : '[number of] years'])]),
      sec('Remedies', ['The Recipient accepts that damages alone may not be enough to remedy a breach. The Discloser may seek an injunction or specific performance, in addition to any other remedy, without proving actual loss to the extent the law allows.'])];
  }
  type({ id: 'nda-mutual', g: 4, name: 'Non-disclosure agreement (NDA), mutual', kw: 'nda mutual non disclosure confidentiality agreement two way both sides partnership discussions',
    title: 'MUTUAL NON-DISCLOSURE AGREEMENT', A: 'First Party', B: 'Second Party', rel: '',
    blurb: 'Both sides share secrets while discussing a deal: definition, duties, exceptions, return, how long it lasts.',
    q: ['purpose', 'start', ['term', 'Information may be shared for (months)'], 'confyears'],
    opts: { nonsolicit: 0, data: 0, dispute: 1, liability: 0 },
    defs: function (x) { return [['Purpose', x.v('purpose')], ['Discloser', 'the Party disclosing Confidential Information'], ['Recipient', 'the Party receiving Confidential Information']]; },
    rec: function (x) { return [L('The Parties wish to explore {0} (the "Purpose").', [x.v('purpose')]), 'For the Purpose each Party will share confidential information with the other, and each agrees to protect the other\'s information on the terms below.']; },
    body: function (x) { return ndaCore(x, true); },
    notes: { all: ['Share only what is needed for the Purpose and mark documents "Confidential". An NDA reduces risk but cannot undo a leak.'] } });
  type({ id: 'nda-oneway', g: 4, name: 'Non-disclosure agreement (NDA), one-way', kw: 'nda one way unilateral non disclosure confidentiality investor freelancer vendor employee',
    title: 'NON-DISCLOSURE AGREEMENT', A: 'Discloser', B: 'Recipient', rel: '',
    blurb: 'Only you share secrets (with a freelancer, investor, buyer or vendor): the other side keeps them confidential.',
    q: ['purpose', 'start', ['term', 'Information may be shared for (months)'], 'confyears'],
    opts: { nonsolicit: 0, data: 0, dispute: 1 },
    defs: function (x) { return [['Purpose', x.v('purpose')]]; },
    rec: function (x) { return [L('The Discloser intends to share confidential information with the Recipient for {0} (the "Purpose").', [x.v('purpose')]), 'The Recipient agrees to keep that information confidential on the terms below.']; },
    body: function (x) { return ndaCore(x, false); },
    notes: { all: ['A one-way NDA protects only the Discloser. If you will also receive the other side\'s secrets, use the mutual NDA.'] } });

  /* ===== Business and partners */
  type({ id: 'partnership', g: 5, name: 'Partnership deed', kw: 'partnership deed firm partners business capital profit sharing indian partnership act',
    title: 'PARTNERSHIP DEED', A: 'First Partner', B: 'Second Partner', rel: false, ownEnd: true, assign: false,
    blurb: 'Two partners starting a firm: name, business, capital, profit sharing, remuneration, bank, accounts, retirement and dissolution.',
    q: ['business', 'nature', 'bizplace', 'start', 'capA', 'capB', 'shareA', 'remun', ['notice', 'Notice to retire (days)']],
    opts: { conf: 0, noncompete: 1, dispute: 1 }, nc: 'both', terr: false,
    defs: function (x) { return [['Firm', L('the partnership firm {0} formed by this Deed', [x.v('business')])], ['Partners', 'the First Partner and the Second Partner, and any partner admitted later'], ['Accounting Year', x.inIN ? 'each year from 1 April to 31 March' : 'each year of 12 months ending on the date the Partners choose']]; },
    rec: function (x) { return ['The Partners have agreed to carry on business in partnership and wish to record the terms in writing.']; },
    body: function (x) { var a = x.has('shareA') ? x.n('shareA') : 50, b = Math.round((100 - a) * 100) / 100; return [
      sec('Name, business and place', [L('The Partners carry on business in partnership under the name {0} from {1}.', [x.v('business'), x.d('start')]),
        L('The business of the Firm is {0}, and any other lawful business the Partners agree in writing.', [x.v('nature')]),
        L('The principal place of business is {0}. The Firm may open branches elsewhere if the Partners agree.', [x.v('bizplace')]),
        'The partnership is at will. It continues until ended under this Deed.']),
      sec('Capital', [L('The First Partner contributes {0} and the Second Partner {1} as capital. Further capital will be brought in as the Partners agree.', [x.m('capA'), x.m('capB')]),
        L('Capital accounts will be kept for each Partner. {0}', [x.inIN ? 'Interest on capital may be paid at a rate agreed by the Partners, not exceeding the rate the Income-tax Act allows as a deduction (currently 12% a year).' : 'Interest on capital will be paid only if the Partners agree in writing.'])]),
      sec('Profit and loss', [L('The net profit or loss of the Firm, after interest on capital and partners\' remuneration, will be shared: First Partner {0}%, Second Partner {1}%.', [a, b]),
        'Each Partner may draw money on account of profit as agreed, and drawings will be adjusted at the end of each Accounting Year.']),
      sec('Working partners and remuneration', [L('Both Partners are working partners and will give their full time and attention to the business.{0}', [x.n('remun') > 0 ? L(' Each will be paid remuneration of {0} per month{1}.', [x.m('remun'), x.inIN ? ', within the limits of section 40(b) of the Income-tax Act' : '']) : ' No remuneration is payable unless the Partners agree in writing.'])]),
      sec('Bank and money', ['The Firm\'s bank accounts will be opened in its name and operated jointly by both Partners, or by either Partner alone for payments up to the limit the Partners set in writing.',
        'All money received for the Firm will be deposited in its bank accounts.']),
      sec('Books of account', ['Proper books of account will be kept at the principal place of business, open to each Partner at all times. Accounts will be closed and, where the law requires, audited each Accounting Year, and signed by both Partners.']),
      sec('Duties and limits', [sub('Each Partner will act honestly and in the Firm\'s interest, and will not, without the other Partner\'s written consent:', ['borrow or lend money, or give any guarantee or security, on the Firm\'s behalf;', 'sell, mortgage or charge any property of the Firm, other than stock in the ordinary course of business;', 'admit a new partner, or transfer their share in the Firm;', 'start or settle any legal case on behalf of the Firm, other than recovering ordinary debts; or', 'carry on any business that competes with the Firm while a partner.'])]),
      sec('Retirement and death', [L('A Partner may retire by giving at least {0} of written notice to the other.', [x.days('notice')]),
        'On the retirement or death of a Partner the Firm will not dissolve if the remaining Partner chooses to continue the business (with or without a new partner). The retiring Partner, or the legal heirs of a deceased Partner, will be paid their capital balance and share of profit up to that date, as shown by accounts made up to that date, in instalments within 12 months, with interest at 6% a year on the unpaid balance. Goodwill will be valued as the Partners agree or, failing agreement, by a valuer.',
        'The legal heir of a deceased Partner may be admitted as a partner if the remaining Partner agrees.']),
      sec('Dissolution', ['The Firm may be dissolved by agreement or as the law allows. On dissolution, the assets will be realised, debts paid, capital returned and any surplus or shortfall shared in the profit-sharing ratio.'])]; },
    notes: { IN: ['Partnership deeds are stamped under the state Stamp Act. Register the Firm with the Registrar of Firms (Form 1 under the Indian Partnership Act, 1932): an unregistered firm cannot sue third parties on its contracts (section 69). Get a PAN and GST registration in the Firm\'s name.', 'For more than two partners, edit the parties, capital and profit-sharing clauses and add each partner\'s signature.'],
      UK: ['General partnerships in England and Wales are governed by the Partnership Act 1890; register the partnership for Self Assessment with HMRC. Consider an LLP for limited liability.'],
      US: ['Partnership law is state law (most states follow the Revised Uniform Partnership Act). Get an EIN and check state registration of the business name.'] } });

  type({ id: 'founders', g: 5, name: 'Founders\' agreement (company or LLP, starting draft)', kw: 'founders agreement co-founder startup equity vesting cliff shareholders llp company',
    title: 'FOUNDERS\' AGREEMENT', A: 'First Founder', B: 'Second Founder', rel: false, ownEnd: true, assign: false, nc: 'both', terr: false,
    blurb: 'Co-founders of a startup: roles, equity, vesting with cliff, IP to the company, decisions, leaving founder, deadlock.',
    q: ['company', 'nature', 'roleA', 'roleB', 'shareA', 'vest', 'cliff', 'capA', 'capB', 'start'],
    opts: { conf: 1, noncompete: 1, nonsolicit: 1, dispute: 1 },
    defs: function (x) { return [['Company', x.v('company')], ['Business', x.v('nature')], ['Founder Shares', 'the shares (or LLP contribution and profit share) held by each Founder'], ['Vested', 'the part of a Founder\'s Founder Shares that is no longer subject to buy-back under the Vesting clause'], ['Leaving Founder', 'a Founder who stops working full-time for the Company for any reason']]; },
    rec: function (x) { return ['The Founders are starting (or have started) the Company to carry on the Business.', 'They wish to agree how they will work together, own the Company and deal with a founder leaving, before investors come in.']; },
    body: function (x) { var a = x.has('shareA') ? x.n('shareA') : 50, b = Math.round((100 - a) * 100) / 100; return [
      sec('Roles and time', [L('The First Founder will be responsible for: {0}. The Second Founder will be responsible for: {1}.', [x.v('roleA'), x.v('roleB')]),
        L('Each Founder will work full-time for the Company from {0} and will not take other paid work without the other\'s consent.', [x.d('start')]),
        'Founders will draw no salary until the Company raises funds or earns enough, as they decide together.']),
      sec('Ownership', [L('The Founder Shares will be held: First Founder {0}%, Second Founder {1}% (before any investment or employee stock pool).', [a, b]),
        L('Initial contributions: First Founder {0}, Second Founder {1}, paid as share capital or a loan to the Company as the Founders agree.', [x.m('capA'), x.m('capB')]),
        'If the Company is not yet formed, the Founders will form it promptly and issue the Founder Shares in these proportions.']),
      sec('Vesting', [L('Each Founder\'s Founder Shares vest over {0} from {1}, with a cliff of {2}: nothing vests before the cliff, the cliff portion vests at the cliff date, and the rest vests in equal monthly portions after it.', [x.has('vest') ? L('{0} years', [x.n('vest')]) : '[number of] years', x.d('start'), x.mo('cliff')]),
        'If a Founder becomes a Leaving Founder, the Company (or the other Founder) may buy back the unvested Founder Shares at their face value (or the lowest price the law allows) within 90 days. If the Founder leaves because of serious misconduct or breach of this Agreement, the buy-back may also cover vested shares at the lower of face value and fair value, to the extent the law allows.',
        'All Founder Shares fully vest if the Company is sold or merged and the Founder is not offered a comparable role.']),
      sec('Intellectual property', ['Each Founder assigns to the Company all rights in the code, designs, brand, content, domain names and other work they have created for the Business, before or after this Agreement, and will sign any document needed to record this. Nothing is kept back for personal use.']),
      sec('Decisions', ['Day-to-day decisions within each Founder\'s area are theirs. The following need both Founders\' written agreement: raising money or issuing shares, borrowing above an amount the Founders set, hiring or firing senior staff, changing the Business, selling the Company or its main assets, and any payment to a Founder or a Founder\'s relative.',
        'If the Founders cannot agree on such a decision after discussing it in good faith for 30 days, they will try mediation before any other step.']),
      sec('Transfer of shares', ['No Founder may sell or transfer Founder Shares without first offering them to the other Founder at the same price and terms (right of first refusal), and the other Founder may join any sale to a third party on the same terms (tag-along).']),
      sec('Formal documents', [L('When the Company is formed or raises investment, the Founders will put these terms into its articles, a shareholders\' agreement{0}. Until then, this Agreement binds the Founders personally.', [x.inIN ? ' or the LLP agreement (filed with the Registrar within 30 days of incorporation)' : ''])])]; },
    notes: { IN: ['Share buy-backs and transfers in a private company must follow the Companies Act, 2013 and the articles; vesting is usually done through a share transfer or reverse-vesting arrangement in the shareholders\' agreement. Stamp duty applies to share issues and transfers.'], all: ['Founders\' agreements are worth a lawyer\'s review before you sign: investors will read them later.'] } });

  type({ id: 'mou', ex: { purpose: 'running a 3-month digital skills course for final-year students' }, g: 5, name: 'Memorandum of understanding (MOU)', kw: 'mou memorandum of understanding cooperation collaboration non binding intent letter of intent',
    title: 'MEMORANDUM OF UNDERSTANDING', A: 'First Party', B: 'Second Party', ownEnd: true,
    blurb: 'Recording a plan to work together: purpose, roles, costs, contacts, which parts are binding and which are not.',
    q: ['purpose', 'respA', 'respB', 'costs', 'start', ['term', 'Duration (months)'], 'notice'],
    opts: { conf: 1, ip: 1, dispute: 1 }, ip: 'owner',
    defs: function (x) { return [['Purpose', x.v('purpose')], ['Project', 'the joint activities described in this MOU'], ['Binding Clauses', 'the clauses listed in the Legal effect clause']]; },
    rec: function (x) { return [L('The Parties share an interest in {0}.', [x.v('purpose')]), 'They wish to record their understanding of how they will work together before entering into any detailed agreement.']; },
    body: function (x) { return [
      sec('Purpose', [L('The Parties will work together on {0} (the "Project").', [x.v('purpose')])]),
      sec('Roles', [sub('The First Party will:', x.lines('respA').length ? x.lines('respA') : ['[tasks]']), sub('The Second Party will:', x.lines('respB').length ? x.lines('respB') : ['[tasks]'])]),
      sec('Working together', ['Each Party will name a contact person who will be responsible for the Project. The contact persons will meet (in person or online) at least once a month to review progress and agree next steps, and will record decisions in writing.']),
      sec('Costs', [(x.raw('costs') || 'Each Party bears its own costs') === 'Each Party bears its own costs' ? 'Each Party will bear its own costs of the Project, unless the Parties agree otherwise in writing before the cost is incurred.' : (x.raw('costs') === 'Shared equally' ? 'Costs of the Project agreed in advance in writing will be shared equally. Each Party bears its own internal costs.' : 'Costs will be shared as set out in a schedule the Parties sign.')]),
      sec('Publicity', ['Neither Party will use the other\'s name or logo, or make a public announcement about the Project, without the other\'s prior written approval.']),
      sec('Term and ending', [L('This MOU starts on {0} and lasts for {1}, unless replaced by a detailed agreement or ended earlier.', [x.d('start'), x.mo('term')]), L('Either Party may end this MOU by giving {0} of written notice.', [x.days('notice')])]),
      sec('Legal effect', ['Except for the Binding Clauses, this MOU records the Parties\' intentions only. It is not a legally binding contract and does not oblige either Party to enter into any further agreement.',
        'The Binding Clauses are: Costs, Publicity, Term and ending, this Legal effect clause, Confidentiality, Intellectual property, Disputes and General. They are legally binding.'])]; },
    notes: { all: ['If money, property or exclusive rights change hands, follow the MOU with a full agreement before work starts.'], IN: ['An MOU may need stamp duty as an agreement under your state\'s Stamp Act if it is to be relied on.'] } });

  type({ id: 'collaboration', ex: { purpose: 'jointly producing and selling an online course on GST for small businesses', respA: 'Write and record the course videos\nAnswer students\' questions in the course forum', respB: 'Build and run the course website and payments\nRun marketing campaigns and handle customer support' }, g: 5, name: 'Business collaboration or joint venture agreement', kw: 'joint venture jv collaboration co-marketing revenue share project partnership contractual',
    title: 'COLLABORATION AGREEMENT', A: 'First Party', B: 'Second Party', ip: 'owner', ownEnd: false, nc: 'both', rel: ' except as this Agreement expressly provides',
    blurb: 'Two businesses doing a project together without forming a company: contributions, revenue share, accounts, decisions, exit.',
    q: ['purpose', 'respA', 'respB', 'capA', 'capB', ['shareA', 'Share of net revenue for Party 1 (%)'], 'start', ['term', 'Term (months)'], 'notice'],
    opts: { conf: 1, ip: 1, noncompete: 0, nonsolicit: 1, indemnity: 1, liability: 1, termination: 1, force: 1, dispute: 1, data: 1 },
    defs: function (x) { return [['Project', x.v('purpose')], ['Net Revenue', 'all money received from the Project, less taxes, refunds, payment charges and the Project costs both Parties approved in writing'], ['Background IP', 'intellectual property a Party owned before, or creates outside, the Project'], ['Project IP', 'intellectual property created for the Project']]; },
    rec: function (x) { return ['Each Party brings skills and resources that the other needs for the Project.', 'They wish to carry out the Project together, sharing its revenue, without forming a separate company or partnership.']; },
    body: function (x) { var a = x.has('shareA') ? x.n('shareA') : 50, b = Math.round((100 - a) * 100) / 100; return [
      sec('The Project', [L('The Parties will carry out the Project together from {0} for {1}.', [x.d('start'), x.mo('term')])]),
      sec('Contributions', [sub(L('The First Party will contribute {0} (in money or agreed value of resources) and will:', [x.m('capA')]), x.lines('respA').length ? x.lines('respA') : ['[tasks]']), sub(L('The Second Party will contribute {0} and will:', [x.m('capB')]), x.lines('respB').length ? x.lines('respB') : ['[tasks]'])]),
      sec('Revenue sharing and accounts', [L('Net Revenue will be shared: First Party {0}%, Second Party {1}%.', [a, b]),
        'The Party that receives Project money (the "Collecting Party") will keep separate records, send a statement of revenue and costs by the 10th of each month, and pay the other Party its share within 15 days of the statement.',
        'Each Party may inspect the Project records on 7 days\' notice, once a quarter. If an underpayment of more than 5% is found, the Collecting Party pays the shortfall and the cost of the inspection.']),
      sec('Management', ['A committee of one representative from each Party will meet monthly to plan and review the Project. Decisions on budgets, pricing, new partners and anything outside the plan need both representatives\' agreement.',
        'Neither Party may sign contracts or take on debts in the other\'s name.']),
      sec('Project IP', ['Each Party keeps its Background IP and lets the other use it free of charge only for the Project during this Agreement.', 'Project IP will be owned jointly in equal shares unless the Parties agree otherwise in writing; neither may license or sell it to others without the other\'s consent.']),
      sec('Exit', ['On ending, the Parties will settle accounts within 30 days, finish or hand over customer commitments fairly, and each may continue to use the jointly owned Project IP for its own business unless they agree a buy-out.'])]; },
    notes: { all: ['If the venture will hold property, hire staff or raise money, a separate company or LLP with a shareholders\' or LLP agreement is usually better.'], IN: ['Revenue sharing between two businesses may have GST consequences; check with your accountant.'] } });

  /* ===== Money and loans */
  function loanSchedule(x) {
    var P = x.n('principal'), n = Math.round(x.n('instal')), r = x.n('rate'), rows = [];
    if ((x.raw('repay') || 'Equal monthly instalments') !== 'Equal monthly instalments' || !P || !n) return null;
    var e = r2(emi(P, r, n)), bal = P;
    for (var i = 1; i <= Math.min(n, 120); i++) {
      var int = r2(bal * r / 1200), pr = i === n ? r2(bal) : r2(e - int), pay = i === n ? r2(pr + int) : e; bal = r2(bal - pr);
      rows.push([String(i), x.raw('start') ? x.date(addMonths(x.raw('start'), i)) : '[date]', x.ms(pay), x.ms(pr), x.ms(int), x.ms(Math.max(0, bal))]);
    }
    return { emi: e, rows: rows };
  }
  type({ id: 'loan', g: 6, name: 'Loan agreement between individuals', kw: 'loan agreement friend family personal loan lend borrow money interest emi instalments repayment',
    title: 'LOAN AGREEMENT', A: 'Lender', B: 'Borrower', rel: false, ownEnd: true,
    blurb: 'Lending money to a friend, relative or small business: amount, interest, instalment table, default, cheques and security.',
    q: ['principal', 'start', 'paymode', 'rate', 'repay', 'instal', 'duedate', 'loanpurpose', 'security'],
    opts: { dispute: 0, latefee: 0 },
    defs: function (x) { return [['Loan', 'the amount lent under this Agreement'], ['Disbursement Date', x.d('start')], ['Outstanding Amount', 'the part of the Loan, interest and other amounts not yet repaid']]; },
    rec: function (x) { return [L('The Borrower has asked the Lender for a loan{0}.', [x.has('loanpurpose') ? L(' for {0}', [x.v('loanpurpose')]) : '']), 'The Lender has agreed to lend, and the Borrower to repay, on the terms below.']; },
    body: function (x) { var sch = loanSchedule(x), mode = x.raw('repay') || 'Equal monthly instalments'; return [
      sec('The loan', [L('The Lender lends the Borrower {0}, paid on the Disbursement Date by {1}. The Borrower confirms receiving it.', [x.m('principal'), x.raw('paymode') || 'bank transfer']),
        x.inIN ? 'The loan is not given or repaid in cash: loans and repayments of ₹20,000 or more must be by account payee cheque, bank draft or electronic transfer (sections 269SS and 269T of the Income-tax Act).' : 'Payments under this Agreement will be made by bank transfer or another traceable method.']),
      sec('Interest', [x.n('rate') > 0 ? L('Interest is payable at {0}% per year on the Outstanding Amount, calculated monthly on the reducing balance.', [x.n('rate')]) : 'The Loan is interest-free, except for default interest under the Default clause.']),
      sec('Repayment', [mode === 'Equal monthly instalments' ? L('The Borrower will repay the Loan with interest in {0} equal monthly instalments{1}, starting one month after the Disbursement Date, as shown in Schedule 1.', [x.has('instal') ? x.n('instal') : '[ ]', sch ? L(' of {0}', [x.ms(sch.emi)]) : '']) :
        mode === 'On demand' ? 'The Borrower will repay the Loan with interest within 30 days after the Lender asks for repayment in writing.' : L('The Borrower will repay the Loan with all interest in one amount on or before {0}.', [x.d('duedate')]),
        'The Borrower may repay all or part of the Loan early at any time without penalty; interest stops on the amount repaid.',
        'Payments are applied first to default interest, then to interest, then to the Loan. The Lender will give a receipt or acknowledgment for each payment.']),
      sec('Default', [sub('The whole Outstanding Amount becomes payable at once, if the Lender gives written notice, when:', ['an instalment or payment is more than 30 days late;', 'the Borrower gave false information to get the Loan;', 'the Borrower becomes insolvent or bankrupt; or', 'any cheque given for repayment is dishonoured.']),
        'On any overdue amount the Borrower will pay additional simple interest of 2% per year above the agreed rate until it is paid, which the Parties agree is a reasonable pre-estimate of loss.',
        'The Borrower will pay the Lender\'s reasonable costs of recovering overdue amounts.']),
      sec('Security and cheques', [x.has('security') ? L('As security for the Loan: {0}.', [x.v('security')]) : 'The Loan is unsecured.',
        x.inIN ? 'If agreed, the Borrower will give post-dated cheques for the instalments. Dishonour of a cheque may have consequences under section 138 of the Negotiable Instruments Act, 1881.' : 'If agreed, the Borrower will set up a standing order or direct debit for the instalments.']),
      sec('Borrower\'s promises', ['The Borrower confirms that it has the legal capacity to borrow, that the information it gave the Lender is true, and that it will tell the Lender promptly of any change of address or anything that may affect repayment.'])]; },
    sched: function (x) { var s = loanSchedule(x); return s ? [{ h: 'Schedule 1: Repayment schedule', items: [tbl(['#', 'Due date', 'Instalment', LANG && LANG.t['Principal (loan amount repaid)'] != null ? LANG.t['Principal (loan amount repaid)'].replace(/\s*\([^)]*\)\s*$/, '') : 'Principal', 'Interest', 'Balance'], s.rows)] }] : []; },
    notes: { IN: ['Loan agreements are stamped under the state Stamp Act. Lending money regularly as a business may need a licence under state money-lending laws; this draft is for an occasional personal loan.', 'Keep bank proof of the transfer: it is the best evidence of the loan.'],
      UK: ['Lending to an individual in the course of a business is regulated consumer credit needing FCA authorisation. Loans between friends and family are usually not regulated.'],
      US: ['Check your state\'s usury limit on interest. For loans to family above the IRS limit, charge at least the Applicable Federal Rate to avoid tax issues.'] } });

  type({ id: 'promissory', g: 6, name: 'Promissory note (on demand)', kw: 'promissory note pronote borrow money promise to pay on demand negotiable instrument',
    title: 'PROMISSORY NOTE', A: 'Maker', B: 'Payee', sign: ['A'], noEsign: true,
    blurb: 'A short signed promise by a borrower to pay a sum, on demand or by a date, with interest.',
    q: ['principal', 'rate', 'start', 'repay', 'duedate'],
    opts: {},
    short: function (x) { var demand = (x.raw('repay') || '') !== 'In one amount on the due date';
      return [x.ms('principal') + '  ' + (x.v('place', 'place')) + ', ' + x.d('start'),
        L('{0}, I, {1}, of {2}, promise to pay {3}, of {4}{5}, the sum of {6}{7}, for value received in {8}.', [demand ? 'On demand' : L('On or before {0}', [x.d('duedate')]), x.party('a'), x.v('aaddr', 'address'), x.party('b'), x.v('baddr', 'address'), x.inIN ? ', or order' : ', or order', x.m('principal'), x.n('rate') > 0 ? L(', together with interest at {0}% per year from the date of this note until payment', [x.n('rate')]) : '', x.raw('paymode') || 'bank transfer']),
        'If not paid when due, I will also pay the reasonable costs of collection.'];
    },
    notes: { IN: ['Under the Indian Stamp Act a promissory note must be stamped (for small on-demand notes usually with a revenue stamp, signed across). It must not be made payable "to bearer" (section 31 of the RBI Act).', 'A promissory note is signed only by the maker (the borrower).'],
      US: ['Promissory notes are governed by state law (UCC Article 3 for negotiable notes). Check usury limits.'], UK: ['A promissory note under the Bills of Exchange Act 1882 must be an unconditional written promise signed by the maker.'] } });

  type({ id: 'guarantee', g: 6, name: 'Personal guarantee', kw: 'guarantee guarantor surety personal guarantee credit supply rent security',
    title: 'DEED OF GUARANTEE', A: 'Guarantor', B: 'Creditor', ownEnd: true, noEsign: true,
    blurb: 'A person promising to pay if a customer, tenant or borrower does not: what is covered, limit, demand, ending.',
    q: ['debtor', 'obligation', ['principal', 'Maximum amount guaranteed'], 'start'],
    opts: { dispute: 0 },
    defs: function (x) { return [['Debtor', x.v('debtor')], ['Guaranteed Obligations', x.v('obligation')], ['Maximum Amount', L('{0}, plus interest and recovery costs on that amount', [x.m('principal')])]]; },
    rec: function (x) { return ['The Creditor has agreed to do business with, or give credit to, the Debtor only if the Guarantor guarantees the Guaranteed Obligations.', 'The Guarantor has a relationship with the Debtor and agrees to give this guarantee.']; },
    body: function (x) { return [
      sec('Guarantee', [L('In return for the Creditor dealing with the Debtor, the Guarantor guarantees that the Debtor will pay the Guaranteed Obligations when due. If the Debtor does not pay, the Guarantor will pay the Creditor the unpaid amount within 15 days of a written demand, up to the Maximum Amount.{0}', [x.inIN ? ' This is a contract of guarantee under section 126 of the Indian Contract Act, 1872.' : '']),
        L('This is a continuing guarantee covering all Guaranteed Obligations arising from {0} until it ends under the Ending clause, not only the first transaction.', [x.d('start')])]),
      sec('Creditor\'s steps first', ['Before making a demand on the Guarantor, the Creditor will send the Debtor a written reminder and allow 15 days to pay. The Creditor need not first start legal action against the Debtor.']),
      sec('Effect of changes', [L('The Guarantor\'s liability is not reduced if the Creditor gives the Debtor more time, or if the Debtor\'s other security is released, unless the Creditor has varied the Guaranteed Obligations without the Guarantor\'s consent in a way that increases the Guarantor\'s risk{0}.', [x.inIN ? ', in which case sections 133 to 139 of the Indian Contract Act apply' : ''])]),
      sec('Ending', ['The Guarantor may end this guarantee for future transactions by written notice to the Creditor. It continues to cover all Guaranteed Obligations arising before the notice is received.',
        'When all Guaranteed Obligations have been paid and the guarantee has ended, the Creditor will confirm the release in writing.']),
      sec('Guarantor\'s confirmations', ['The Guarantor confirms that it has read this guarantee, understands that it may have to pay the Debtor\'s debts, has had the chance to take independent legal advice, and signs freely.'])]; },
    notes: { UK: ['A guarantee must be in writing and signed by the guarantor (Statute of Frauds 1677). It is usually signed as a deed with a witness. A personal guarantee given to secure an individual\'s consumer credit has extra rules.'], all: ['A guarantor should never sign without understanding the debt they may have to pay. Ask the creditor for the debtor\'s account statement whenever needed.'] } });

  /* ===== Authority, settlements and endings */
  type({ id: 'poa', g: 7, name: 'Power of attorney (starting draft)', kw: 'power of attorney poa gpa spa general special attorney authorise agent nri property',
    title: 'POWER OF ATTORNEY', A: 'Principal', B: 'Attorney', ownEnd: true, noEsign: true, rel: false, assign: false, sign: ['A', 'B'], deed: true,
    blurb: 'Giving someone authority to act for you for named acts or generally: powers, limits, ratification, validity and revocation.',
    q: ['poakind', 'matter', 'powers', 'start', 'validity'],
    opts: {},
    defs: function (x) { return [['Matter', x.v('matter')], ['Powers', 'the powers listed in this deed']]; },
    rec: function (x) { return [L('The Principal cannot attend to the Matter in person{0} and wishes to appoint the Attorney, whom the Principal trusts, to act for the Principal.', [x.inIN ? ' as needed' : ''])]; },
    body: function (x) { var gen = (x.raw('poakind') || '').indexOf('General') === 0; return [
      sec('Appointment', [L('By this deed the Principal appoints the Attorney as the Principal\'s {0} attorney to do the acts listed below in the Principal\'s name and on the Principal\'s behalf, in relation to the Matter.', [gen ? 'general' : 'special'])]),
      sec('Powers', [sub('The Attorney may:', x.lines('powers').length ? x.lines('powers') : ['[list each power]']), gen ? 'And generally do all lawful acts in relation to the Matter that the Principal could do personally, except those excluded below.' : 'The Attorney may do only the acts listed above.']),
      sec('Limits', [sub('The Attorney may not:', ['sell, gift, mortgage or transfer any property of the Principal unless that power is expressly listed above;', 'receive money except into the Principal\'s own bank account;', 'make gifts, lend the Principal\'s money, or benefit personally, except reasonable expenses supported by bills;', 'appoint anyone else to act in the Attorney\'s place, unless this deed expressly allows it.']),
        'The Attorney will act honestly, keep accounts of everything done and money handled, and report to the Principal whenever asked.']),
      sec('Ratification', ['The Principal agrees to ratify and confirm everything the Attorney lawfully does under this deed.']),
      sec('Validity and revocation', [x.n('validity') > 0 ? L('This power is valid from {0} for {1}, unless revoked earlier.', [x.d('start'), x.mo('validity')]) : L('This power is valid from {0} until revoked.', [x.d('start')]),
        L('The Principal may revoke this power at any time by written notice to the Attorney (and, where it was registered, by registering the revocation). It also ends on the death{0} of the Principal{1}.', [x.inUK || x.inUS ? '' : ' or mental incapacity', x.inUS ? ', and, unless this deed says it is durable, on the Principal\'s incapacity' : x.inUK ? ', and on the Principal\'s loss of mental capacity (an ordinary power of attorney is not a lasting power of attorney)' : '']),
        'Anyone dealing with the Attorney in good faith without notice of revocation may rely on this deed.'])]; },
    notes: { IN: ['A power of attorney is stamped as the state Stamp Act requires and attested by a notary; one that authorises sale or other transfer of immovable property usually must be registered, and a sale by power of attorney does not transfer ownership (Supreme Court, Suraj Lamp & Industries v. State of Haryana, 2011). An NRI signing abroad should have it attested by the Indian Embassy or apostilled and then adjudicated for stamp duty within 3 months of arrival in India.'],
      UK: ['An ordinary power of attorney must be signed as a deed in front of a witness (Powers of Attorney Act 1971). To cover loss of mental capacity, use a lasting power of attorney registered with the Office of the Public Guardian.'],
      US: ['Most states have a statutory power of attorney form (often under the Uniform Power of Attorney Act) and require notarisation; real-estate powers may need recording with the county. Use your state\'s form where one exists.'],
      AE: ['Powers of attorney in the UAE must be notarised by a notary public (often in Arabic or bilingual) and, if signed abroad, legalised and attested by the UAE embassy and the Ministry of Foreign Affairs.'] } });

  type({ id: 'settlement', g: 7, name: 'Settlement and release agreement', kw: 'settlement release compromise dispute full and final payment withdraw case',
    title: 'SETTLEMENT AGREEMENT', A: 'First Party', B: 'Second Party', ownEnd: true,
    blurb: 'Ending a dispute: settlement amount and date, full and final release, withdrawing cases, no admission, confidentiality.',
    q: ['dispute', 'casedetails', 'setamt', 'setdate', 'payer', 'paymode'],
    opts: { conf: 1, dispute: 0 },
    defs: function (x) { return [['Dispute', 'the disagreement described in the Background, and every claim arising from the same facts'], ['Settlement Amount', x.m('setamt')], ['Payment Date', x.d('setdate')]]; },
    rec: function (x) { return [L('A dispute has arisen between the Parties about: {0}.{1}', [x.v('dispute'), x.has('casedetails') ? L(' Pending case or complaint: {0}.', [x.v('casedetails')]) : '']), 'Without admitting liability, the Parties wish to settle the Dispute fully and finally on the terms below.']; },
    body: function (x) { var payA = (x.raw('payer') || 'AB') === 'AB', P = payA ? x.A : x.B, R = payA ? x.B : x.A; return [
      sec('Settlement payment', [L('The {0} will pay the {1} the Settlement Amount by {2} on or before the Payment Date.', [P, R, x.raw('paymode') || 'bank transfer']),
        'Each Party bears its own legal and other costs of the Dispute.']),
      sec('Full and final settlement', [L('On receiving the Settlement Amount in full, the {0} releases the {1}, and the {2} releases the {3}, from all claims, known or unknown, arising out of the Dispute, and neither will bring any further claim about it.', [R, P, P, R]),
        'This release does not cover obligations under this Agreement, or claims that the law does not allow to be released.']),
      sec('Cases and complaints', [L('Within 15 days after the Settlement Amount is paid, the Parties will withdraw or jointly apply to settle every case, complaint or notice between them about the Dispute{0}, and sign any document needed for this{1}.', [x.has('casedetails') ? L(', including {0}', [x.v('casedetails')]) : '', x.inIN ? ' (such as a compromise application under Order XXIII of the Code of Civil Procedure, or a joint application in a Lok Adalat or before the mediator)' : ''])]),
      sec('No admission', ['This Agreement is a compromise. It is not an admission of liability or wrongdoing by either Party.']),
      sec('Non-disparagement', ['Neither Party will make any false or damaging public statement about the other in relation to the Dispute. This does not stop either Party from giving truthful information required by law.']),
      sec('If payment is not made', [L('If the Settlement Amount is not paid in full by the Payment Date, the {0} may, by written notice, either sue for the unpaid Settlement Amount under this Agreement or treat this Agreement as ended and continue its original claims, giving credit for any amount received.', [R])])]; },
    notes: { all: ['If there is a pending court case, the settlement usually has to be recorded before that court to end the case. Settling employment claims may need special formalities (for example a settlement agreement with independent legal advice in the UK).'] } });

  type({ id: 'termination', g: 7, name: 'Mutual termination agreement', kw: 'termination agreement end contract mutual consent cancel agreement closure exit handover',
    title: 'MUTUAL TERMINATION AGREEMENT', A: 'First Party', B: 'Second Party', ownEnd: true,
    blurb: 'Ending an existing agreement by consent: end date, final payment, handover, surviving clauses, release.',
    q: ['origname', 'origdate', ['start', 'End date'], 'finalamt', 'payer', 'handover'],
    opts: { conf: 1, dispute: 0 },
    defs: function (x) { return [['Original Agreement', L('the {0} between the Parties dated {1}', [x.v('origname'), x.d('origdate')])], ['End Date', x.d('start')]]; },
    rec: function (x) { return ['The Parties entered into the Original Agreement.', 'They have agreed to end it by mutual consent on the terms below.']; },
    body: function (x) { var payA = (x.raw('payer') || 'AB') === 'AB', P = payA ? x.A : x.B, R = payA ? x.B : x.A; return [
      sec('Ending', ['The Original Agreement ends on the End Date by mutual consent. No notice period or early-ending charge under the Original Agreement applies.']),
      sec('Final settlement', [x.n('finalamt') > 0 ? L('The {0} will pay the {1} {2} within 15 days after the End Date, in full and final settlement of all amounts under the Original Agreement.', [P, R, x.m('finalamt')]) : 'No further amount is payable by either Party under the Original Agreement, apart from what is set out in this Agreement.']),
      sec('Handover', [sub('Before or on the End Date the Parties will complete:', x.lines('handover').length ? x.lines('handover') : ['[things to return or finish]'])]),
      sec('Surviving obligations', ['Clauses of the Original Agreement on confidentiality, ownership of work, indemnity for third-party claims and limitation of liability continue after the End Date, as that agreement provides.']),
      sec('Release', ['Once the final settlement is paid and the handover completed, each Party releases the other from all other claims under the Original Agreement, except claims about the obligations that survive and claims that cannot be released by law.'])]; },
    notes: { all: ['Check the original agreement\'s ending clause for anything that must still happen (return of deposits, notices to third parties, cancellation of registrations).'] } });

  /* ---------------- putting an agreement together */
  F.place = ['City where it is signed', 'text', 'Ludhiana'];
  F.date = ['Date of signing', 'date', today()];
  F.region = ['State', 'text', 'Punjab'];
  F.stampby = ['Stamp duty and registration paid by', 'who', 'S', ['S', 'A', 'B']];
  F.wit = ['Witnesses', 'sel', 'Two', ['None', 'One', 'Two']];
  var EX = { commercial: { term: '36', deposit: '360000' }, franchise: { term: '60' }, distribution: { term: '24' }, supply: { term: '12' }, service: { term: '12' },
    agent: { term: '12' }, collaboration: { term: '12' }, mou: { term: '12' }, subcontract: { term: '3' }, amc: { term: '12' }, 'nda-mutual': { term: '24' },
    'nda-oneway': { term: '12' }, internship: { term: '3' }, 'equipment-hire': { deposit: '10000' } };   /* more per-type samples (see applyEx) */
  var TYPE = {}; TYPES.forEach(function (t) { TYPE[t.id] = t; if (EX[t.id]) t.ex = Object.assign({}, t.ex, EX[t.id]); });
  var NOTES = {
    IN: ['Print the final agreement on non-judicial stamp paper or an e-stamp certificate (or have it franked) of the value your state\'s Stamp Act requires, in the name of one of the parties and dated on or before the signing date. An unstamped or under-stamped document cannot be used as evidence until the duty and a penalty are paid.',
      'Both parties initial every page and sign the last page in full, in front of two witnesses who write their names and addresses. Documents that must be registered are presented at the sub-registrar\'s office within 4 months of signing.'],
    UK: ['Most agreements in the UK need no stamp duty and no witness. A document signed as a deed (for example a power of attorney or a guarantee) must be signed in front of a witness, who also signs. Land deals may need SDLT (England and Northern Ireland), LBTT (Scotland) or LTT (Wales) and registration at the Land Registry.'],
    US: ['Most agreements in the US need no witness or stamp. Powers of attorney and documents to be recorded (such as deeds) usually need a notary. Some states have mandatory forms, notices or disclosures for leases, consumer contracts, employment and powers of attorney.'],
    AE: ['Contracts may be in English, but UAE courts work in Arabic and rely on a certified Arabic translation; a bilingual contract should say which language prevails. Tenancy, employment and some commercial contracts must use or be registered on official government systems.'],
    INT: ['Check your country\'s rules on stamp duty, registration, witnesses, notarisation and the language of contracts before signing.']
  };
  var DISCLAIM = 'This is a starting draft made from a template, not legal advice. It has not been checked against your facts or the latest law. Laws, stamp duty and registration rules differ by country and state and change over time. Read every clause, change what does not fit, fill every [square bracket], and have a lawyer check it before anyone signs, especially for property, large amounts or long terms.';

  function partyDesc(x, p, role) {
    var V = x.S.values, kind = V[p + 'kind'] || 'Individual', name = x.party(p), addr = x.v(p + 'addr', 'address'), id = (V[p + 'id'] || '').trim();
    var signer = x.v(p + 'sign', 'name of signatory'), title = x.v(p + 'title', 'designation'), s;
    if (kind === 'Individual') s = L('{0}{1}, residing at {2}{3}', [name, V[p + 'rel'] ? ', ' + V[p + 'rel'].trim() : '', addr, id ? L(' (ID: {0})', [id]) : '']);
    else if (kind === 'Sole proprietorship') s = L('{0}, a sole proprietorship owned by {1}, with its place of business at {2}{3}', [name, signer, addr, id ? ' (' + id + ')' : '']);
    else if (kind === 'Trust, society or other organisation') s = L('{0}{1}, having its office at {2}, acting through its authorised representative {3}, {4}', [name, id ? L(' (registration no. {0})', [id]) : '', addr, signer, title]);
    else { var o = (x.c.org[kind] || x.c.org.Company || 'a ' + kind.toLowerCase()).replace('{region}', x.v('region', x.c.region)); s = L('{0}, {1}{2}, having its {3} at {4}, acting through {5}, its {6}', [name, o, id ? L(' (registration no. {0})', [id]) : '', kind === 'Partnership firm' ? 'principal place of business' : 'registered office', addr, signer, title]); }
    if (V[p + 'email']) s += L(', email {0}', [V[p + 'email'].trim()]);
    return L('{0} (the "{1}")', [s, role]);
  }

  function compose(S) {
    LANG = S.cc === 'IN' && S.lang && S.lang !== 'en' ? PACKS[S.lang] || null : null; LANGINFO = LANG ? langInfo(S.lang) : null;
    var T = TYPE[S.type] || TYPES[0], x = ctx(T, S), M = { title: T.title, parts: [] };
    var place = x.v('place', 'city'), date = x.d('date');
    if (T.short) {
      M.short = T.short(x);
    } else {
      M.intro = (T.deed ? L('This {0} is made at {1} on {2} by the Principal in favour of the Attorney named below.', [cap(T.title), place, date]) : L('This {0} is made at {1} on {2} between:', [cap(T.title), place, date]));
      M.parties = [L(T.deed ? '{0}; in favour of' : '{0}; and', [partyDesc(x, 'a', T.A)]), L('{0}.', [partyDesc(x, 'b', T.B)])];
      M.together = T.deed ? '' : L('The {0} and the {1} are together called the "Parties" and each a "Party".', [T.A, T.B]);
      M.rec = T.rec ? T.rec(x) : [];
      var secs = [];
      var defs = T.defs ? T.defs(x) : [];
      secs.push(sec('Definitions and interpretation', [sub(L('In this {0}, unless the context requires otherwise:', [T.deed ? 'deed' : 'Agreement']), [['"' + (T.deed ? 'Deed' : 'Agreement') + '"', L('this document, including its Schedules and any written changes signed by {0}', [T.deed ? 'the Principal' : 'both Parties'])]].concat(defs).map(function (d) { return L('"{0}" means {1};', [d[0].replace(/"/g, ''), d[1]]); })),
        'Headings are for convenience only. Words in the singular include the plural and the other way round; "including" means "including without limitation"; and "days" means calendar days.']));
      secs = secs.concat(T.body(x));
      OPT_ORDER.forEach(function (k) { if (T.opts[k] != null && x.on(k)) secs.push(OPT[k][3](x)); });
      if (!T.deed) secs.push(general(x)); else secs.push(sec('Governing law', [L('This deed is governed by {0}.', [x.law()])]));
      M.secs = secs;
      M.sched = T.sched ? T.sched(x) : [];
    }
    M.sign = (T.sign || ['A', 'B']).map(function (p) { var k = p.toLowerCase(), kind = S.values[k + 'kind'] || 'Individual'; return { role: p === 'A' ? T.A : T.B, name: x.party(k), org: kind !== 'Individual', signer: S.values[k + 'sign'] || '', title: S.values[k + 'title'] || '' }; });
    if (T.deed) M.sign[1].accept = true;
    M.wit = { None: 0, One: 1, Two: 2 }[S.values.wit || x.c.wit] || 0;
    if ((T.deed || T.id === 'guarantee') && S.cc === 'UK' && M.wit < 1) M.wit = 1;   /* a deed needs a witness in the UK */
    M.notary = !!T.deed;
    var n = [DISCLAIM].concat(T.short ? [] : NOTES[S.cc] || [], (T.notes && T.notes[S.cc]) || [], (T.notes && T.notes.all) || []);
    if (T.opts.noncompete != null && x.on('noncompete') && !x.inIN) n.push('Non-compete clauses are restricted in many places (for example most of them are void in California and limited in several other US states; in the UK they must be reasonable). Remove the clause if it is not allowed where you are.');
    M.notes = S.notes === false ? [] : n;
    if (LANG) {   /* fixed sentences (headings, plain clauses, notes) are swapped here; sentences with answers in them were made by L() */
      M.lang = LANGINFO[0]; M.tnote = LANG.note; M.title = tr(M.title);
      ['short', 'intro', 'parties', 'together', 'rec', 'secs', 'sched', 'notes'].forEach(function (k) { if (M[k] != null) M[k] = trDeep(M[k]); });
    }
    M.file = (M.lang ? M.lang + '-' : '') + T.title.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') + (S.values.bname ? '-' + String(S.values.bname).replace(/[^\w ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 30) : '');
    return M;
  }
  function cap(s) { return s.toLowerCase().replace(/(^|\s|\()([a-z])/g, function (m, a, b) { return a + b.toUpperCase(); }).replace(/ (Of|And|Or|To|For|The|In) /g, function (m) { return m.toLowerCase(); }).replace(/\(Outline\)/, '(outline)'); }

  /* ---------------- render the model as the preview document */
  function textWithPh(el, s) {
    String(s).split(/(\[[^\]]*\])/).forEach(function (part) { if (!part) return; el.appendChild(/^\[.*\]$/.test(part) ? h('span', { 'class': 'ct-ph', text: part }) : doc.createTextNode(part)); });
    return el;
  }
  function P(cls, s, num) { var p = h('p', { 'class': cls }); if (num) { p.appendChild(h('b', { text: num })); p.appendChild(doc.createTextNode(' ')); } return textWithPh(p, s); }
  function table(t) {
    var el = h('table', { 'class': 'ct-tb' }), tr = h('tr');
    t.head.forEach(function (c) { tr.appendChild(h('th', { text: c })); });
    el.appendChild(h('thead', {}, [tr])); var tb = h('tbody');
    t.rows.forEach(function (r) { var row = h('tr'); r.forEach(function (c) { row.appendChild(textWithPh(h('td'), c == null ? '' : c)); }); tb.appendChild(row); });
    el.appendChild(tb); return el;
  }
  var LET = 'abcdefghijklmnopqrstuvwxyz';
  function render(M) {
    var li = M.lang && langInfo(M.lang), sc = li && SCRIPT[li[3]];
    var el = h('div', { 'class': 'tk-doc ct-doc' + (li ? ' ct-intl ct-s-' + li[3] : ''), id: 'ct-doc', lang: li ? li[0] : 'en', dir: sc && sc[2] ? 'rtl' : null,
      'data-page': li ? tr('Page ') : null, 'data-initials': li ? tr('Initials:') : null,
      style: li ? '--ct-font:\'' + sc[0] + '\',\'' + sc[1] + '\',Cambria,Georgia,serif' : null, 'data-wfont': sc ? sc[1] : null });
    el.appendChild(h('h2', { 'class': 'ct-title', text: M.title }));
    if (M.tnote) el.appendChild(h('p', { 'class': 'ct-tnote', text: M.tnote }));
    if (M.short) {
      M.short.forEach(function (s, i) { el.appendChild(P(i ? 'ct-p' : 'ct-p ct-right', s)); });
    } else {
      el.appendChild(P('ct-p', M.intro));
      M.parties.forEach(function (s, i) { el.appendChild(P('ct-cl', s, '(' + (i + 1) + ')')); });
      if (M.together) el.appendChild(P('ct-p', M.together));
      if (M.rec.length) { el.appendChild(h('h3', { text: tr('Background') })); M.rec.forEach(function (s, i) { el.appendChild(P('ct-cl', s, String.fromCharCode(65 + i) + '.')); }); }
      el.appendChild(P('ct-p', tr(M.parties.length && M.together ? 'The Parties agree as follows:' : 'The Principal declares as follows:')));
      M.secs.forEach(function (s, i) {
        var n = i + 1; el.appendChild(h('h3', { text: n + '. ' + s.h }));
        s.items.forEach(function (it, j) {
          var num = n + '.' + (j + 1);
          if (typeof it === 'string') el.appendChild(P('ct-cl', it, num));
          else if (it.table) el.appendChild(table(it.table));
          else { el.appendChild(P('ct-cl', it.t, num)); it.sub.forEach(function (sb, k) { el.appendChild(P('ct-sub', sb, '(' + LET[k % 26] + ')')); }); }
        });
      });
      M.sched.forEach(function (s) {
        el.appendChild(h('h3', { 'class': 'ct-sch', text: s.h }));
        s.items.forEach(function (it) {
          if (typeof it === 'string') el.appendChild(P('ct-p', it));
          else if (it.table) el.appendChild(table(it.table));
          else { if (it.t) el.appendChild(P('ct-p', it.t)); it.sub.forEach(function (sb, k) { el.appendChild(P('ct-sub', sb, (k + 1) + '.')); }); }
        });
      });
    }
    el.appendChild(h('h3', { 'class': 'ct-sigh', text: tr('Signatures') }));
    el.appendChild(P('ct-p', tr(M.short ? 'Signed by the Maker:' : 'Signed on the date and at the place written above.')));
    var sg = h('div', { 'class': 'ct-sig' });
    M.sign.forEach(function (s) {
      var b = h('div', { 'class': 'ct-sb' });
      b.appendChild(h('p', { 'class': 'ct-sl', text: s.accept ? L('Accepted by the {0}', [s.role]) : (s.org ? L('For and on behalf of the {0}', [s.role]) : L('The {0}', [s.role])) }));
      b.appendChild(textWithPh(h('p', { 'class': 'ct-sn' }), s.name));
      b.appendChild(h('p', { text: tr('Signature: ______________________') }));
      b.appendChild(h('p', { text: s.org ? L('Name and title: {0}', [s.signer ? s.signer + (s.title ? ', ' + s.title : '') : '______________']) : tr('Date: ______________') }));
      if (s.org) b.appendChild(h('p', { text: tr('Date: ______________') }));
      sg.appendChild(b);
    });
    for (var w = 1; w <= M.wit; w++) {
      sg.appendChild(h('div', { 'class': 'ct-sb' }, [h('p', { 'class': 'ct-sl', text: L('Witness {0}', [w]) }), h('p', { text: tr('Signature: ______________________') }), h('p', { text: tr('Name: ______________________') }), h('p', { text: tr('Address: ______________________') })]));
    }
    el.appendChild(sg);
    if (M.notary) el.appendChild(h('div', { 'class': 'ct-nt' }, [h('p', { 'class': 'ct-sl', text: tr('Notary or attestation') }), h('p', { text: tr('Signed before me by the Principal, who is known to me or proved identity to my satisfaction, on ______________.') }), h('p', { text: tr('Notary\'s signature, seal and registration number: ______________________') })]));
    if (M.notes.length) {
      var nb = h('div', { 'class': 'ct-notes' }, [h('h3', { text: tr('Before you sign (remove this page from the final copy)') })]);
      var ul = h('ul'); M.notes.forEach(function (s, i) { ul.appendChild(h('li', { 'class': i ? null : 'ct-first', text: s })); }); nb.appendChild(ul); el.appendChild(nb);
    }
    el.appendChild(h('p', { 'class': 'foot', contenteditable: 'false' }, [CREDIT]));
    return el;
  }

  /* ---------------- Word (.docx) from the preview as shown (so typed edits are kept) */
  function loadDocx() {
    if (window.docx) return Promise.resolve(window.docx);
    return new Promise(function (ok, bad) {
      var s = h('script', { src: DOCX[0], async: true });
      s.onload = function () { window.docx ? ok(window.docx) : bad(new Error('The Word maker did not load.')); };
      s.onerror = function () { bad(new Error('Could not load the Word maker. Check the internet connection and try again.')); };
      doc.head.appendChild(s);
    });
  }
  var WL = null;   /* Word settings of the document being written: { font, lang, rtl } */
  function runsOf(node, D, base) {
    var out = [];
    (function walk(n, bold) {
      if (n.nodeType === 3) { if (n.nodeValue) out.push(new D.TextRun(Object.assign({ text: n.nodeValue.replace(/\s+/g, ' '), bold: bold || base.b, size: base.size || 22, color: base.color }, WL ? WL.run : {}))); return; }
      if (n.nodeType !== 1) return;
      if (n.tagName === 'BR') { out.push(new D.TextRun({ text: '', break: 1 })); return; }
      var b = bold || /^(B|STRONG)$/.test(n.tagName);
      Array.prototype.forEach.call(n.childNodes, function (c) { walk(c, b); });
    })(node, false);
    return out;
  }
  function toDocx(el, file, title) {
    return loadDocx().then(function (D) {
      var AL = D.AlignmentType, kids = [], nob = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, line = { style: D.BorderStyle.SINGLE, size: 4, color: '94A3B8' };
      var lc = el.getAttribute('lang') || 'en', rtl = el.getAttribute('dir') === 'rtl', wf = el.getAttribute('data-wfont');
      WL = lc === 'en' ? null : { rtl: rtl, run: { font: { ascii: wf, hAnsi: wf, cs: wf, eastAsia: wf }, language: { value: lc + '-IN', bidirectional: lc + '-IN' }, rightToLeft: rtl || undefined } };
      if (rtl) AL = Object.assign({}, AL, { LEFT: AL.START || 'start', RIGHT: AL.END || 'end' });
      function para(node, o) { o = o || {}; return new D.Paragraph({ alignment: o.al || AL.JUSTIFIED, bidirectional: rtl || undefined, indent: o.indent, spacing: { after: o.after == null ? 100 : o.after, before: o.before || 0 }, keepNext: o.keepNext, pageBreakBefore: o.pbb, children: runsOf(node, D, o) }); }
      function tableOf(t) {
        return new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: Array.prototype.map.call(t.querySelectorAll('tr'), function (tr) {
          return new D.TableRow({ children: Array.prototype.map.call(tr.children, function (c) { var th = c.tagName === 'TH';
            return new D.TableCell({ shading: th ? { fill: 'F1F5F9', type: D.ShadingType.CLEAR, color: 'auto' } : undefined, borders: { top: line, bottom: line, left: line, right: line },
              children: [new D.Paragraph({ bidirectional: rtl || undefined, children: runsOf(c, D, { b: th, size: 20 }) })] }); }) });
        }) });
      }
      function blockTable(blocks) {
        var rows = [];
        for (var i = 0; i < blocks.length; i += 2) {
          rows.push(new D.TableRow({ cantSplit: true, children: [blocks[i], blocks[i + 1]].map(function (b) {
            return new D.TableCell({ width: { size: 50, type: D.WidthType.PERCENTAGE }, borders: { top: nob, bottom: nob, left: nob, right: nob },
              children: b ? Array.prototype.map.call(b.children, function (p, k) { return new D.Paragraph({ bidirectional: rtl || undefined, spacing: { after: 80, before: k ? 0 : 240 }, children: runsOf(p, D, { b: p.classList.contains('ct-sl'), size: 21 }) }); }) : [new D.Paragraph('')] }); }) }));
        }
        return new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: rows });
      }
      Array.prototype.forEach.call(el.children, function (n) {
        var c = n.classList, tag = n.tagName;
        if (c.contains('foot')) return;
        if (tag === 'H2') kids.push(para(n, { al: AL.CENTER, b: true, size: 30, after: 240 }));
        else if (tag === 'H3') kids.push(para(n, { al: AL.LEFT, b: true, size: 23, before: 200, after: 100, keepNext: true, pbb: c.contains('ct-sch') && false }));
        else if (tag === 'P') {
          if (c.contains('ct-tnote')) kids.push(para(n, { al: AL.LEFT, size: 19, color: '7C2D12', after: 200 }));
          else if (c.contains('ct-sub')) kids.push(para(n, { indent: { left: 1134, hanging: 425 } }));
          else if (c.contains('ct-cl')) kids.push(para(n, { indent: { left: 567, hanging: 567 } }));
          else if (c.contains('ct-right')) kids.push(para(n, { al: AL.RIGHT }));
          else kids.push(para(n));
        } else if (tag === 'TABLE') { kids.push(tableOf(n)); kids.push(new D.Paragraph({ text: '' })); }
        else if (c.contains('ct-sig')) kids.push(blockTable(Array.prototype.slice.call(n.children)));
        else if (c.contains('ct-nt')) { kids.push(new D.Paragraph({ text: '', spacing: { after: 200 } })); Array.prototype.forEach.call(n.children, function (p) { kids.push(para(p, { b: p.classList.contains('ct-sl'), al: AL.LEFT })); }); }
        else if (c.contains('ct-notes')) {
          var first = true;
          Array.prototype.forEach.call(n.querySelectorAll('h3,li'), function (p) {
            kids.push(p.tagName === 'H3' ? para(p, { b: true, size: 23, al: AL.LEFT, pbb: true, color: '9F1239' }) : new D.Paragraph({ bullet: { level: 0 }, bidirectional: rtl || undefined, spacing: { after: 80 }, children: runsOf(p, D, { size: 20, color: '334155' }) }));
            first = false;
          });
        } else {   /* anything typed in edit mode */
          String(n.innerText || n.textContent || '').split(/\n+/).forEach(function (t) { if (t.trim()) kids.push(new D.Paragraph({ bidirectional: rtl || undefined, spacing: { after: 100 }, children: [new D.TextRun(Object.assign({ text: t, size: 22 }, WL ? WL.run : {}))] })); });
        }
      });
      var d = new D.Document({ creator: 'Nivesh Sahayak', lastModifiedBy: 'Nivesh Sahayak', keywords: 'Toolkit by GroupChatOn; Nivesh Sahayak; https://www.niveshsahayak.in/', description: NS + ' · https://www.niveshsahayak.in/', title: title,
        styles: { default: { document: { run: WL ? Object.assign({ size: 22 }, WL.run) : { font: 'Cambria', size: 22 } } } },
        sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1247, right: 1134 } } },
          footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: AL.CENTER, children: [new D.TextRun({ text: el.getAttribute('data-page') || 'Page ', size: 16, color: '94A3B8' }), new D.TextRun({ children: [D.PageNumber.CURRENT], size: 16, color: '94A3B8' }), new D.TextRun({ text: '   ·   ' + (el.getAttribute('data-initials') || 'Initials:') + ' ________   ________   ·   ' + CREDIT, size: 14, color: '94A3B8' })] })] }) },
          children: kids }] });
      return D.Packer.toBlob(d).then(function (blob) {
        var a = h('a', { href: URL.createObjectURL(blob), download: file + '.docx' }); doc.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
      });
    });
  }

  /* ---------------- styles for this tool only (the shared CSS styles the form fields and hides the form when printing) */
  var CSS = '.ct-list{max-height:300px;overflow:auto;border:1px solid var(--line);border-radius:12px;padding:6px;background:var(--bg)}' +
    '.ct-gh{margin:8px 6px 4px;font-size:.74rem;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:var(--muted)}' +
    '.ct-opt{display:block;width:100%;text-align:left;min-height:44px;padding:8px 10px;margin:2px 0;border:1px solid transparent;border-radius:10px;background:none;color:var(--ink);font:inherit;font-size:.93rem;cursor:pointer}' +
    '.ct-opt:hover{background:var(--soft)}.ct-opt[aria-pressed=true]{background:var(--soft);border-color:var(--brand);font-weight:700}' +
    '.ct-sel{margin:8px 0 0;font-size:.9rem;color:var(--muted)}.ct-sel b{color:var(--ink)}' +
    '.ct-search{width:100%;min-height:44px;padding:9px 12px;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--ink);font:inherit;margin-bottom:8px}' +
    '.tk-f textarea{min-height:96px;resize:vertical;line-height:1.45}' +
    '.ct-ochk{border:1px solid var(--line);border-radius:12px;padding:6px 10px;margin:6px 0;background:var(--bg)}.ct-ochk small{display:block;color:var(--muted);font-weight:500;margin:-6px 0 4px 28px;font-size:.82rem}' +
    '.ct-ochk .tk-row{margin:4px 0 6px}' +
    '.ct-warn{background:var(--in);border-radius:12px;padding:10px 12px;font-size:.9rem;margin:0 0 10px}' +
    '.ct-edit-bar{display:none;background:var(--soft);border-radius:12px;padding:10px 12px;font-size:.9rem;margin:0 0 10px}.ct-edit-bar.on{display:block}.ct-edit-bar button{margin-top:6px}' +
    '.ct-scroll{max-height:calc(100vh - 220px);min-height:320px;overflow:auto;border-radius:6px}' +
    '@media (max-width:960px){.ct-scroll{max-height:none}}' +
    '.ct-doc{font:13.5px/1.6 Cambria,Georgia,\'Times New Roman\',serif;padding:26px clamp(14px,4vw,34px)}' +
    '.ct-doc[contenteditable=true]{outline:3px dashed #BE185D;outline-offset:-3px;cursor:text}' +
    '.ct-doc .ct-title{text-align:center;font:700 18px/1.3 Cambria,Georgia,serif;letter-spacing:.06em;margin:0 0 14px;color:#14213D}' +
    '.ct-doc h3{font:700 14px/1.35 Cambria,Georgia,serif;text-align:start;margin:16px 0 6px;color:#14213D;break-after:avoid}' +
    '.ct-doc p{margin:0 0 6px;text-align:justify}.ct-doc .ct-cl{padding-inline-start:2.4em;text-indent:-2.4em}.ct-doc .ct-sub{padding-inline-start:4.6em;text-indent:-2em}' +
    '.ct-doc .ct-right{text-align:end}.ct-doc .ct-ph{background:#FEF3C7;border-radius:3px;padding:0 2px;color:#92400E}' +
    '.ct-doc .ct-tb{border-collapse:collapse;width:100%;margin:6px 0 10px;font-size:12.5px}.ct-doc .ct-tb th,.ct-doc .ct-tb td{border:1px solid #CBD5E1;padding:4px 6px;text-align:start;vertical-align:top;background:none;color:#14213D}.ct-doc .ct-tb th{background:#F1F5F9;font-weight:700}' +
    '.ct-doc .ct-sig{display:grid;grid-template-columns:1fr 1fr;gap:6px 24px;margin-top:8px}@media (max-width:520px){.ct-doc .ct-sig{grid-template-columns:1fr}.ct-doc .ct-tb{font-size:11px}.ct-doc .ct-tb th,.ct-doc .ct-tb td{padding:3px 4px}}' +
    '.ct-doc .ct-sb{break-inside:avoid;padding-top:16px}.ct-doc .ct-sb p,.ct-doc .ct-nt p{text-align:start;margin:0 0 8px}.ct-doc .ct-sl{font-weight:700}.ct-doc .ct-sn{font-style:italic}' +
    '.ct-doc .ct-nt{margin-top:18px;border-top:1px solid #CBD5E1;padding-top:8px;break-inside:avoid}' +
    '.ct-doc .ct-notes{margin-top:22px;border:1px solid #F9A8D4;background:#FFF1F2;border-radius:8px;padding:10px 14px}.ct-doc .ct-notes h3{margin-top:4px;color:#9F1239}' +
    '.ct-doc .ct-notes ul{margin:0;padding-inline-start:1.2em}.ct-doc .ct-notes li{margin:0 0 6px;font-size:12.5px}.ct-doc .ct-notes .ct-first{font-weight:700}' +
    '.ct-doc.ct-intl,.ct-doc.ct-intl .ct-title,.ct-doc.ct-intl h3{font-family:var(--ct-font);line-height:1.8}.ct-doc.ct-intl .ct-title{letter-spacing:0;line-height:1.5}' +
    '.ct-doc.ct-s-Arab{font-size:15px;line-height:2}.ct-doc.ct-s-Arab .ct-cl,.ct-doc.ct-s-Arab .ct-sub{text-indent:0}' +
    '.ct-doc .ct-tnote{border:1px solid #FCD34D;background:#FFFBEB;border-radius:6px;padding:6px 10px;font-size:12.5px;color:#14213D;text-align:start;margin:0 0 12px}' +
    '.ct-lang-note{font-size:.84rem;color:var(--muted);margin:2px 0 0}' +
    '@media print{.ct-doc .ct-tnote{background:none}}' +
    '@media print{.ct-scroll{max-height:none!important;overflow:visible!important}.ct-warn,.ct-edit-bar{display:none!important}.ct-doc{padding:0!important;outline:0!important;font-size:11.5pt}' +
    '.ct-doc .ct-ph{background:none;color:inherit}.ct-doc .ct-notes{break-before:page;background:none}.ct-doc h3{break-after:avoid}.ct-doc .ct-cl,.ct-doc .ct-sub{orphans:3;widows:3}}';

  /* ---------------- the form */
  function sampleValues(cc, keep) {
    var c = COUNTRY[cc], v = {};
    Object.keys(F).forEach(function (k) { v[k] = F[k][2]; });
    v.place = c.city || ''; v.courts = c.city || ''; v.region = c.regionDef; v.wit = c.wit; v.date = today();
    v.akind = 'Individual'; v.aname = c.samples[0]; v.aaddr = c.samples[1]; v.bkind = 'Individual'; v.bname = c.samples[2]; v.baddr = c.samples[3];
    ['rel', 'id', 'sign', 'title', 'email'].forEach(function (k) { v['a' + k] = ''; v['b' + k] = ''; });
    v.balancedate = addMonths(today(), 2); v.delivery = addDays(today(), 14); v.eventdate = addMonths(today(), 3); v.setdate = addDays(today(), 30); v.duedate = addMonths(today(), 12); v.origdate = addMonths(today(), -12);
    if (cc !== 'IN') {   /* India-flavoured samples read oddly elsewhere: keep only the neutral ones */
      ['property', 'workplace', 'bizplace', 'location', 'venue', 'territory', 'matter', 'products', 'brand', 'mainproject', 'mentor', 'powers', 'titledocs'].forEach(function (k) { v[k] = ''; });
      v.business = 'Bennett & Clarke Partners'; v.vreg = ''; v.vmake = ''; v.purpose = 'discussing a possible partnership to build and sell a booking app for clinics';
    }
    return Object.assign(v, keep || {});
  }
  function samplePairs() {   /* [answer key, English sample] that a language pack may translate */
    var c = COUNTRY.IN, out = [['place', c.city], ['courts', c.city], ['region', c.regionDef], ['aname', c.samples[0]], ['aaddr', c.samples[1]], ['bname', c.samples[2]], ['baddr', c.samples[3]]];
    Object.keys(F).forEach(function (k) { if (/^(text|wide|area)$/.test(F[k][1]) && F[k][2]) out.push([k, F[k][2]]); });
    TYPES.forEach(function (t) { if (t.ex) Object.keys(t.ex).forEach(function (k) { if (/[a-z]/i.test(t.ex[k])) out.push([k, t.ex[k]]); }); });
    return out;
  }
  function sampleSwap(S, from, to) {
    if (from === to) return;
    samplePairs().forEach(function (e) {
      var k = e[0], en = e[1], was = from && from.t[en] != null ? from.t[en] : en;
      if (S.values[k] === was) S.values[k] = to && to.t[en] != null ? to.t[en] : en;
    });
  }
  function packOf(S) { return S.cc === 'IN' && S.lang && S.lang !== 'en' ? PACKS[S.lang] || null : null; }
  function applyEx(S) {   /* sample answers that only suit this type replace the shared sample (never the user's own text) */
    var t = TYPE[S.type], keys = {};
    TYPES.forEach(function (o) { if (o.ex) Object.keys(o.ex).forEach(function (k) { keys[k] = 1; }); });
    var P = packOf(S), sm = function (en) { return P && P.t[en] != null ? P.t[en] : en; };
    Object.keys(keys).forEach(function (k) { var v = S.values[k]; if (blank(v) || v === F[k][2] || v === sm(F[k][2]) || TYPES.some(function (o) { return o.ex && (o.ex[k] === v || sm(o.ex[k]) === v); })) S.values[k] = sm((t.ex && t.ex[k]) || F[k][2]); });
  }
  function initial() {
    var s = store();
    if (s && s.values && TYPE[s.type] && COUNTRY[s.cc]) { s.values = Object.assign(sampleValues(s.cc), s.values); s.opts = s.opts || {}; return s; }
    return { type: 'rent', cc: 'IN', cur: 'INR', values: sampleValues('IN'), opts: {}, notes: true, edited: null };   /* rent has no own samples */
  }

  function mount(box) {
    var spot = box.querySelector('.tk-mount'); if (!spot || spot.getAttribute('data-ct')) return; spot.setAttribute('data-ct', '1');
    if (!doc.getElementById('ct-css')) doc.head.appendChild(h('style', { id: 'ct-css', text: CSS }));
    var S = initial(), timer, docEl;
    TYPES.forEach(function (t) { if (!S.opts[t.id]) { S.opts[t.id] = {}; for (var k in t.opts) S.opts[t.id][k] = !!t.opts[k]; } });
    var form = h('form', { 'class': 'tk-form', novalidate: true, onsubmit: function (e) { e.preventDefault(); } });
    var msg = h('p', { 'class': 'tk-msg', role: 'status', 'aria-live': 'polite' });
    var scroll = h('div', { 'class': 'ct-scroll' });
    var editBar = h('div', { 'class': 'ct-edit-bar', role: 'status' });
    var editBtn;
    function save() { store(S); }
    function T() { return TYPE[S.type]; }

    /* fields */
    function fieldEl(key, label, onChange) {
      var def = F[key] || [label, 'text', ''], lab = label || def[0], type = def[1], val = S.values[key], inp;
      if (type === 'sel' || type === 'who' || type === 'pay') {
        var t = T(), name = function (o) { return o === 'A' ? 'The ' + t.A : o === 'B' ? 'The ' + t.B : o === 'S' ? 'Both, in equal shares' : o === 'AB' ? 'The ' + t.A + ' pays the ' + t.B : o === 'BA' ? 'The ' + t.B + ' pays the ' + t.A : o; };
        inp = h('select', { onchange: function () { S.values[key] = inp.value; onChange(); } }, def[3].map(function (o) { return h('option', { value: o, selected: String(o) === String(val == null ? def[2] : val) ? true : null, text: name(o) }); }));
      } else if (type === 'area') {
        inp = h('textarea', { rows: 4, oninput: function () { S.values[key] = inp.value; onChange(); } }); inp.value = val == null ? '' : val;
      } else {
        inp = h('input', { type: type === 'num' ? 'text' : type === 'date' ? 'date' : 'text', inputmode: type === 'num' ? 'decimal' : null, value: val == null ? '' : val, autocomplete: 'off',
          oninput: function () { S.values[key] = inp.value; onChange(); } });
      }
      var wide = type === 'wide' || type === 'area';
      return h('label', { 'class': 'tk-f' + (wide ? ' wide' : ''), style: wide ? 'grid-column:1/-1' : null }, [lab, inp]);
    }
    function changed() { save(); schedule(); }
    function schedule() { clearTimeout(timer); timer = setTimeout(preview, 150); }

    var list, search, selInfo;
    function typeList() {
      list = h('div', { 'class': 'ct-list', role: 'group', 'aria-label': 'Agreement types' });
      G.forEach(function (g, gi) {
        var grp = h('div', { 'class': 'ct-grp' }, [h('p', { 'class': 'ct-gh', text: g })]);
        TYPES.filter(function (t) { return t.g === gi; }).forEach(function (t) {
          grp.appendChild(h('button', { type: 'button', 'class': 'ct-opt', 'data-id': t.id, 'data-kw': (t.name + ' ' + t.kw + ' ' + g).toLowerCase(), 'aria-pressed': t.id === S.type ? 'true' : 'false', text: t.name,
            onclick: function () { pick(t.id); } }));
        });
        list.appendChild(grp);
      });
      return list;
    }
    function filter() {
      var q = search.value.toLowerCase().trim().split(/\s+/).filter(Boolean);
      Array.prototype.forEach.call(list.querySelectorAll('.ct-grp'), function (g) {
        var any = false;
        Array.prototype.forEach.call(g.querySelectorAll('.ct-opt'), function (b) { var ok = q.every(function (w) { return b.getAttribute('data-kw').indexOf(w) >= 0; }); b.hidden = !ok; if (ok) any = true; });
        g.hidden = !any;
      });
    }
    function pick(id) {
      if (S.type === id) return;
      if (S.edited && !confirm('You have typed your own changes into the draft. Switching the agreement type will replace them. Continue?')) return;
      S.type = id; S.edited = null; applyEx(S); save();
      Array.prototype.forEach.call(list.querySelectorAll('.ct-opt'), function (b) { b.setAttribute('aria-pressed', b.getAttribute('data-id') === id ? 'true' : 'false'); });
      buildRest(); preview(true);
      msg.textContent = T().name + ' selected.';
    }

    var rest = h('div');
    function buildRest() {
      var t = T(), c = COUNTRY[S.cc];
      rest.innerHTML = '';
      selInfo.innerHTML = ''; selInfo.appendChild(h('b', { text: t.name + ': ' })); selInfo.appendChild(doc.createTextNode(t.blurb));
      /* country and signing */
      var cSel = h('select', { onchange: function () { switchCountry(cSel.value); } }, Object.keys(COUNTRY).map(function (k) { return h('option', { value: k, selected: k === S.cc ? true : null, text: COUNTRY[k].name }); }));
      var curSel = h('select', { onchange: function () { S.cur = curSel.value; changed(); } }, Object.keys(CUR).map(function (k) { return h('option', { value: k, selected: k === S.cur ? true : null, text: k + ' (' + CUR[k][0] + ')' }); }));
      var lSel = S.cc === 'IN' ? h('select', { onchange: function () { switchLang(lSel.value, lSel); } }, LANGS.map(function (l) { return h('option', { value: l[0], lang: l[0], selected: l[0] === (S.lang || 'en') ? true : null, text: l[0] === 'en' ? 'English' : l[1] + ' (' + l[2] + ')' }); })) : null;
      var gen = [h('label', { 'class': 'tk-f' }, ['Country', cSel]), lSel ? h('label', { 'class': 'tk-f' }, ['Language of the agreement', lSel, h('span', { 'class': 'ct-lang-note', text: 'English or ' + (LANGS.length - 1) + ' languages of India' })]) : null, fieldEl('region', c.region, changed), h('label', { 'class': 'tk-f' }, ['Currency', curSel]),
        fieldEl('place', null, changed), t.short ? null : fieldEl('courts', S.cc === 'US' ? 'Courts in (county or city)' : S.cc === 'UK' ? 'Courts or arbitration city' : null, changed), fieldEl('date', null, changed),
        S.cc === 'IN' && !t.short && !t.noStamp ? fieldEl('stampby', null, changed) : null, fieldEl('wit', null, changed)].filter(Boolean);
      var notesChk = h('input', { type: 'checkbox', checked: S.notes !== false ? true : null, onchange: function () { S.notes = notesChk.checked; changed(); } });
      rest.appendChild(h('fieldset', {}, [h('legend', { text: '2. Country and signing' }), h('div', { 'class': 'tk-row' }, gen),
        h('label', { 'class': 'tk-chk' }, [notesChk, 'Add the "Before you sign" notes page to the draft'])]));
      /* parties */
      ['a', 'b'].forEach(function (p, i) {
        var role = i ? t.B : t.A, kind = S.values[p + 'kind'] || 'Individual', row = h('div', { 'class': 'tk-row' });
        var kSel = h('select', { onchange: function () { S.values[p + 'kind'] = kSel.value; save(); buildRest(); preview(); } }, KINDS.map(function (k) { return h('option', { value: k, selected: k === kind ? true : null, text: k }); }));
        row.appendChild(h('label', { 'class': 'tk-f' }, ['Who is the ' + role + '?', kSel]));
        row.appendChild(fieldEl(p + 'name', kind === 'Individual' ? 'Full name' : 'Name of the business', changed));
        if (kind === 'Individual') row.appendChild(fieldEl(p + 'rel', S.cc === 'IN' ? 'Son, daughter or wife of (optional)' : 'Extra description (optional)', changed));
        row.appendChild(fieldEl(p + 'addr', 'Full address', changed));
        if (kind !== 'Individual') { row.appendChild(fieldEl(p + 'sign', kind === 'Sole proprietorship' ? 'Owner\'s name' : 'Person signing', changed)); if (kind !== 'Sole proprietorship') row.appendChild(fieldEl(p + 'title', 'Their title', changed)); }
        row.appendChild(fieldEl(p + 'id', kind === 'Individual' ? (S.cc === 'IN' ? 'PAN or other ID (optional)' : 'ID number (optional)') : c.idl, changed));
        row.appendChild(fieldEl(p + 'email', 'Email for notices (optional)', changed));
        rest.appendChild(h('fieldset', {}, [h('legend', { text: (i + 3) + '. ' + role + ' (Party ' + (i + 1) + ')' }), row]));
      });
      /* the type's own questions */
      rest.appendChild(h('fieldset', {}, [h('legend', { text: '5. Details of this agreement' }), h('p', { 'class': 'muted', style: 'margin:0 0 8px;font-size:.86rem', text: 'Leave a box empty if you do not know yet: the draft shows [square brackets] there.' }),
        h('div', { 'class': 'tk-row' }, t.q.map(function (q) { return Array.isArray(q) ? fieldEl(q[0], q[1], changed) : fieldEl(q, null, changed); }))]));
      /* optional clauses */
      var keys = OPT_ORDER.filter(function (k) { return t.opts[k] != null; });
      if (keys.length) {
        var box = h('div');
        keys.forEach(function (k) {
          var o = OPT[k], on = !!S.opts[t.id][k], extra = h('div', { 'class': 'tk-row', hidden: on ? null : true }, o[2].map(function (q) { return fieldEl(q, null, changed); }));
          var cb = h('input', { type: 'checkbox', checked: on ? true : null, onchange: function () { S.opts[t.id][k] = cb.checked; extra.hidden = !cb.checked; changed(); } });
          box.appendChild(h('div', { 'class': 'ct-ochk' }, [h('label', { 'class': 'tk-chk' }, [cb, o[0]]), h('small', { text: o[1] }), o[2].length ? extra : null]));
        });
        rest.appendChild(h('fieldset', {}, [h('legend', { text: '6. Optional clauses' }), box]));
      }
      rest.appendChild(h('div', { 'class': 'tk-acts' }, [
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Fill sample', onclick: function () { if (S.edited && !confirm('Replace your typed changes with the sample?')) return; S.values = sampleValues(S.cc); sampleSwap(S, null, packOf(S)); S.edited = null; applyEx(S); save(); buildRest(); preview(true); msg.textContent = 'Sample answers filled in.'; } }),
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Clear answers', onclick: function () {
          if (!confirm('Clear all answers and typed changes on this device?')) return;
          var keep = { place: '', courts: '', region: COUNTRY[S.cc].regionDef, wit: COUNTRY[S.cc].wit, date: today(), start: today() };
          var v = {}; Object.keys(S.values).forEach(function (k) { v[k] = ''; }); ['akind', 'bkind'].forEach(function (k) { v[k] = 'Individual'; });
          ['maint', 'utilities', 'stampby', 'payer', 'meals', 'feebasis', 'tax', 'repay', 'paymode', 'exclusive', 'poakind', 'coverage', 'hireper', 'hosting', 'ltype', 'livein', 'costs', 'disp', 'liabcap'].forEach(function (k) { v[k] = F[k][2]; });
          S.values = Object.assign(v, keep); S.edited = null; save(); buildRest(); preview(true); msg.textContent = 'Answers cleared.'; } })]));
    }
    function switchLang(code, sel) {
      var before = S.lang || 'en';
      if (code === before) return;
      if (S.edited && !confirm('You have typed your own changes into the draft. Changing the language will replace them. Continue?')) { sel.value = before; return; }
      var li = langInfo(code); msg.textContent = code === 'en' ? '' : 'Loading ' + li[2] + '…'; sel.disabled = true; loadFont(code);
      loadPack(code).then(function (pk) {
        sampleSwap(S, packOf(S), pk); S.lang = code; S.edited = null; save(); buildRest(); preview(true);
        msg.textContent = code === 'en' ? 'Agreement in English.' : 'Agreement now in ' + li[2] + ' (' + li[1] + '). The form stays in English; you can type answers in either language.';
      }, function (er) { sel.value = before; sel.disabled = false; msg.textContent = er.message; });
    }
    function switchCountry(cc) {
      var old = COUNTRY[S.cc], c = COUNTRY[cc];
      if (S.cc === 'IN' && cc !== 'IN') sampleSwap(S, packOf(S), null);
      if (S.values.aname === old.samples[0]) { S.values.aname = c.samples[0]; S.values.aaddr = c.samples[1]; }
      if (S.values.bname === old.samples[2]) { S.values.bname = c.samples[2]; S.values.baddr = c.samples[3]; }
      if (!S.values.place || S.values.place === old.city) S.values.place = c.city;
      if (!S.values.courts || S.values.courts === old.city) S.values.courts = c.city;
      S.values.region = c.regionDef; S.values.wit = c.wit; S.cc = cc; S.cur = c.cur; if (cc === 'IN') sampleSwap(S, null, packOf(S)); save(); buildRest(); preview(); msg.textContent = 'Country changed to ' + c.name + '. Check the state or region and currency.';
    }

    function preview(force) {
      if (S.edited && !force) { editBar.classList.add('on'); return; }
      if (force) S.edited = null;
      editBar.classList.remove('on');
      var M = compose(S); docEl = render(M); docEl.setAttribute('data-file', M.file); docEl.setAttribute('data-title', M.title);
      scroll.innerHTML = ''; scroll.appendChild(docEl); setEditing(false);
    }
    function setEditing(on) {
      if (!docEl) return;
      docEl.setAttribute('contenteditable', on ? 'true' : 'false');
      if (on) { docEl.setAttribute('spellcheck', 'true'); docEl.focus(); }
      if (editBtn) { editBtn.lastChild.textContent = on ? ' Done editing' : ' Edit text'; editBtn.setAttribute('aria-pressed', on ? 'true' : 'false'); }
    }
    scroll.addEventListener('input', function () { S.edited = { type: S.type, html: docEl.innerHTML }; save(); });

    /* build */
    search = h('input', { type: 'search', 'class': 'ct-search', placeholder: 'Search: rent, NDA, loan, freelance, vehicle…', 'aria-label': 'Search agreement types', oninput: filter,
      onkeydown: function (e) { if (e.key === 'Enter') { e.preventDefault(); var b = list.querySelector('.ct-opt:not([hidden])'); if (b && !b.closest('[hidden]')) b.click(); } } });
    selInfo = h('p', { 'class': 'ct-sel' });
    form.appendChild(h('fieldset', {}, [h('legend', { text: '1. Choose the agreement (' + TYPES.length + ' types)' }), search, typeList(), selInfo]));
    form.appendChild(rest);
    buildRest();

    var icon = function (d) { return h('span', { 'aria-hidden': 'true', text: d }); };
    editBtn = h('button', { type: 'button', 'class': 'btn btn-2', 'aria-pressed': 'false', onclick: function () { setEditing(docEl.getAttribute('contenteditable') !== 'true'); msg.textContent = docEl.getAttribute('contenteditable') === 'true' ? 'Type straight into the draft. Your changes are kept on this device.' : 'Editing finished.'; } }, [icon('✏️'), ' Edit text']);
    var acts = h('div', { 'class': 'tk-out-acts' }, [
      h('button', { type: 'button', 'class': 'btn', onclick: function () { setEditing(false); if (doc.fonts && doc.fonts.ready && S.lang && S.lang !== 'en') doc.fonts.ready.then(function () { window.print(); }); else window.print(); } }, [icon('🖨'), ' Print / PDF']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { var b = e.currentTarget; b.disabled = true; setEditing(false); msg.textContent = 'Making the Word file…';
        toDocx(docEl, docEl.getAttribute('data-file') || 'agreement', docEl.getAttribute('data-title') || 'Agreement').then(function () { msg.textContent = 'Word file ready. Check your downloads folder.'; }, function (er) { msg.textContent = (er && er.message) || 'Something went wrong. Please try again.'; }).then(function () { b.disabled = false; }); } }, [icon('📘'), ' Word']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function () { var cl = docEl.cloneNode(true), f = cl.querySelector('.foot'); if (f) f.remove(); doc.body.appendChild(cl); cl.style.cssText = 'position:fixed;left:-9999px'; var t = cl.innerText.trim(); cl.remove();
        var ok = function () { msg.textContent = 'Text copied. Paste it into your email, Word or Google Docs.'; };
        if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(t).then(ok, fallback); else fallback();
        function fallback() { var ta = h('textarea', { style: 'position:fixed;top:-1000px;opacity:0' }); ta.value = t; doc.body.appendChild(ta); ta.select(); try { doc.execCommand('copy'); ok(); } catch (e) { msg.textContent = 'Could not copy. Select the text in the draft and copy it.'; } ta.remove(); } } }, [icon('📋'), ' Copy text']),
      editBtn]);
    editBar.appendChild(doc.createTextNode('You have typed your own changes into the draft, so it no longer updates from the answers. '));
    editBar.appendChild(h('button', { type: 'button', 'class': 'btn btn-2', text: 'Rebuild from answers (drops my typed changes)', onclick: function () { preview(true); save(); msg.textContent = 'Draft rebuilt from your answers.'; } }));
    var warn = h('p', { 'class': 'ct-warn' }, [h('b', { text: 'Starting draft, not legal advice. ' }), 'Laws, stamp duty and registration rules differ by country and state. Read every clause and have a lawyer check it before anyone signs.']);
    var out = h('div', { 'class': 'tk-out' }, [acts, msg, warn, editBar, scroll]);
    spot.appendChild(h('div', { 'class': 'tk-app ct-app' }, [form, out]));
    function start() {
      if (S.edited && S.edited.type === S.type && S.edited.html) {
        preview(true); docEl.innerHTML = S.edited.html; S.edited = { type: S.type, html: docEl.innerHTML }; editBar.classList.add('on');
      } else { S.edited = null; preview(true); }
    }
    if (packOf(S) === null && S.cc === 'IN' && S.lang && S.lang !== 'en') {
      var li0 = langInfo(S.lang); loadFont(S.lang);
      if (!li0) { S.lang = 'en'; start(); }
      else { scroll.appendChild(h('p', { 'class': 'muted', style: 'padding:20px', text: 'Loading ' + li0[2] + '…' }));
        loadPack(S.lang).then(start, function (er) { var ed = S.edited; S.lang = 'en'; S.edited = ed; buildRest(); start(); msg.textContent = er.message + ' Showing English.'; }); }
    } else start();
    doc.body.classList.add('pg-tool');
    /* scroll the chosen type into view inside the list */
    var cur = list.querySelector('[aria-pressed=true]'); if (cur) list.scrollTop = Math.max(0, cur.offsetTop - list.offsetTop - 40);
  }

  window.TKContract = { types: TYPES, compose: compose, countries: COUNTRY, sample: sampleValues, langs: LANGS, samples: samplePairs,
    addLang: function (code, pack) { PACKS[code] = pack; },
    _seen: function (on) { SEEN = on ? { t: {}, s: {}, o: {} } : null; return SEEN; }, _seenNow: function () { return SEEN; }, _render: function (M) { return render(M); }, _applyEx: applyEx };
  function mountAll() { Array.prototype.forEach.call(doc.querySelectorAll('.tk-tool[data-tool="' + TOOL + '"]'), mount); }
  if (doc.readyState !== 'loading') mountAll(); else doc.addEventListener('DOMContentLoaded', mountAll);
})();
