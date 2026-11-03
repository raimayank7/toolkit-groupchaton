/* Toolkit online tools: GST invoice, quick bill / quotation, report card, salary slip, expense claim.
   One engine: a form (profile saved on this device + the document + line items) → a document model →
   rendered as a live preview, printed (Print / PDF), or written as .xlsx (ExcelJS) and .docx (docx), both loaded only when needed.
   Nothing leaves the browser. */
(function () {
  'use strict';
  var doc = document;
  var LIBS = { xlsx: ['https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js', 'ExcelJS'],
               docx: ['https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js', 'docx'] };
  var STATES = ['Andaman and Nicobar Islands', 'Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chandigarh', 'Chhattisgarh',
    'Dadra and Nagar Haveli and Daman and Diu', 'Delhi', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jammu and Kashmir', 'Jharkhand',
    'Karnataka', 'Kerala', 'Ladakh', 'Lakshadweep', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha',
    'Puducherry', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal', 'Other Territory'];
  var CREDIT = 'Designed & Developed by: Nivesh Sahayak';   /* owner: shown under the on-screen preview only, never printed or put in the files (file properties carry it) */
  var MADE = 'toolkit.groupchaton.com';   /* owner: the only line printed or put in files, short and simple */

  /* ---------------- helpers */
  function h(tag, a, kids) {
    var e = doc.createElement(tag);
    for (var k in a || {}) { var v = a[k]; if (v == null || v === false) continue; if (k === 'text') e.textContent = v; else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v); else e.setAttribute(k, v === true ? '' : v); }
    (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? doc.createTextNode(c) : c); });
    return e;
  }
  function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/,/g, '')); return isFinite(n) ? n : 0; }
  function r2(n) { return Math.round((n + 1e-9) * 100) / 100; }
  var F2 = new Intl.NumberFormat('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  var F0 = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
  function money(n, d) { return '₹ ' + (d === 0 ? F0 : F2).format(n); }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function dmy(iso) { if (!iso) return ''; var p = String(iso).split('-'); return p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : iso; }
  function monthName(iso) { if (!iso) return ''; var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, 1).toLocaleString('en-IN', { month: 'long', year: 'numeric' }); }
  function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || 'null'); localStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
  var ONES = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var TENS = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function w2(n) { return n < 20 ? ONES[n] : TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : ''); }
  function words(amount) {   /* Indian style: crore, lakh, thousand */
    amount = r2(Math.abs(amount)); var n = Math.floor(amount), p = Math.round((amount - n) * 100), out = [];
    if (!n && !p) return '';
    var cr = Math.floor(n / 1e7), lk = Math.floor(n % 1e7 / 1e5), th = Math.floor(n % 1e5 / 1e3), hu = Math.floor(n % 1e3 / 100), re = n % 100;
    if (cr) out.push((cr > 99 ? words(cr).replace(/^Rupees | Only$/g, '') : w2(cr)) + ' Crore');
    if (lk) out.push(w2(lk) + ' Lakh'); if (th) out.push(w2(th) + ' Thousand'); if (hu) out.push(ONES[hu] + ' Hundred'); if (re) out.push(w2(re));
    return 'Rupees ' + (out.length ? out.join(' ') : 'Zero') + (p ? ' and ' + w2(p) + ' Paise' : '') + ' Only';
  }
  function loadLib(key) {
    var L = LIBS[key];
    if (window[L[1]]) return Promise.resolve(window[L[1]]);
    return new Promise(function (ok, bad) {
      var s = doc.createElement('script'); s.src = L[0]; s.async = true;
      s.onload = function () { window[L[1]] ? ok(window[L[1]]) : bad(new Error('library did not load')); };
      s.onerror = function () { bad(new Error('Could not load the file maker. Check the internet connection and try again.')); };
      doc.head.appendChild(s);
    });
  }
  function save(blob, name) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name }); doc.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function safe(s) { return String(s || 'document').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'document'; }
  function nextNo(s) { var m = String(s || '').match(/^(.*?)(\d+)(\D*)$/); if (!m) return s || ''; var n = String(+m[2] + 1); while (n.length < m[2].length) n = '0' + n; return m[1] + n + m[3]; }
  function gradeOf(p) { return p >= 91 ? 'A1' : p >= 81 ? 'A2' : p >= 71 ? 'B1' : p >= 61 ? 'B2' : p >= 51 ? 'C1' : p >= 41 ? 'C2' : p >= 33 ? 'D' : 'E'; }

  /* ---------------- the five tools: fields, lines and the document they make */
  var GST = ['0', '5', '12', '18', '28', '40'];
  var TOOLS = {
    invoice: {
      name: 'GST invoice', no: 'no',
      profile: [['name', 'Business name', 'Sharma Traders'], ['addr', 'Address', '12, Main Bazaar, Ludhiana, Punjab 141001', 'wide'], ['gstin', 'GSTIN', '03ABCDE1234F1Z5'],
        ['state', 'State', 'Punjab', 'select', STATES], ['phone', 'Phone', ''], ['email', 'Email', ''], ['bank', 'Bank name', ''], ['acct', 'Account number', ''], ['ifsc', 'IFSC', '']],
      fields: [['no', 'Invoice number', 'INV-001'], ['date', 'Invoice date', today(), 'date'], ['cname', 'Customer name', 'Gupta Enterprises', 'wide'],
        ['caddr', 'Customer address', 'Model Town, Ludhiana', 'wide'], ['cgstin', 'Customer GSTIN (if registered)', ''], ['cstate', 'Customer state (place of supply)', 'Punjab', 'select', STATES]],
      groups: [{ key: 'items', title: 'Items', add: 'Add line',
        cols: [['desc', 'Description', '', 'wide'], ['hsn', 'HSN/SAC', ''], ['qty', 'Qty', '1', 'number'], ['unit', 'Unit', 'Nos'], ['rate', 'Rate (₹)', '', 'number'], ['disc', 'Disc %', '', 'number'], ['gst', 'GST %', '18', 'select', GST]],
        sample: [{ desc: 'Printer toner cartridge', hsn: '8443', qty: '2', unit: 'Pcs', rate: '1850', gst: '18' }, { desc: 'A4 copier paper (500 sheets)', hsn: '4802', qty: '10', unit: 'Ream', rate: '260', disc: '5', gst: '18' }] }],
      make: function (P, D, G) {
        var intra = (P.state || '') === (D.cstate || ''), taxable = 0, tax = 0, rows = [], byRate = {};
        G.items.forEach(function (l, i) {
          if (!l.desc && !num(l.rate)) return;
          var t = r2(num(l.qty) * num(l.rate) * (1 - num(l.disc) / 100)), g = r2(t * num(l.gst) / 100);
          taxable += t; tax += g; byRate[l.gst] = (byRate[l.gst] || 0) + g;
          rows.push([rows.length + 1, l.desc, l.hsn, num(l.qty), l.unit, num(l.rate), num(l.disc) ? num(l.disc) + '%' : '', t, num(l.gst) + '%', g]);
        });
        taxable = r2(taxable); tax = r2(tax);
        var cg = intra ? r2(tax / 2) : 0, sg = intra ? r2(tax - cg) : 0, ig = intra ? 0 : tax, raw = r2(taxable + tax), gt = Math.round(raw);
        var tot = [['Taxable value', taxable]];
        if (intra) tot.push(['CGST', cg], ['SGST', sg]); else tot.push(['IGST', ig]);
        tot.push(['Round off', r2(gt - raw)], ['Grand total', gt, 1]);
        return { file: 'Invoice-' + safe(D.no), title: 'TAX INVOICE', org: P.name, sub: [P.addr, [P.gstin ? 'GSTIN: ' + P.gstin : '', P.state ? 'State: ' + P.state : '', P.phone ? 'Phone: ' + P.phone : '', P.email].filter(Boolean).join('  ·  ')],
          kv: [['Invoice No.', D.no], ['Date', dmy(D.date)], ['Bill to', D.cname], ['Place of supply', D.cstate], ['Address', D.caddr], ['Buyer GSTIN', D.cgstin || 'Unregistered'],
               ['Supply type', intra ? 'Intra-state (CGST + SGST)' : 'Inter-state (IGST)'], ['', '']],
          cols: [['#', 'c'], ['Description', ''], ['HSN/SAC', ''], ['Qty', 'r'], ['Unit', ''], ['Rate', 'r$'], ['Disc', 'r'], ['Taxable value', 'r$'], ['GST', 'r'], ['GST amount', 'r$']],
          rows: rows, totals: tot, words: words(gt),
          notes: [P.bank ? 'Bank: ' + P.bank + (P.acct ? '   A/c: ' + P.acct : '') + (P.ifsc ? '   IFSC: ' + P.ifsc : '') : '',
                  'Declaration: We declare that this invoice shows the actual price of the goods or services described and that all particulars are true and correct.'].filter(Boolean),
          sign: ['Customer\'s signature', 'For ' + (P.name || '') + '\n\nAuthorised signatory'] };
      }
    },
    bill: {
      name: 'bill', no: 'no',
      profile: [['name', 'Shop name', 'Verma General Store'], ['addr', 'Address', 'Shop 4, Sadar Bazaar, Ambala', 'wide'], ['phone', 'Phone', '98120 00000'], ['gstin', 'GSTIN (optional)', '']],
      fields: [['kind', 'Document', 'Bill', 'select', ['Bill', 'Quotation']], ['no', 'Number', 'B-001'], ['date', 'Date', today(), 'date'], ['valid', 'Valid until (quotation)', '', 'date'],
        ['cname', 'Customer name', ''], ['cphone', 'Customer phone', ''], ['disc', 'Discount (₹)', '', 'number'], ['terms', 'Note / terms', 'Thank you! Goods once sold can be exchanged within 7 days with the bill.', 'wide']],
      groups: [{ key: 'items', title: 'Items', add: 'Add item', cols: [['desc', 'Item', '', 'wide'], ['qty', 'Qty', '1', 'number'], ['rate', 'Rate (₹)', '', 'number']],
        sample: [{ desc: 'Basmati rice 5 kg', qty: '1', rate: '610' }, { desc: 'Sunflower oil 1 L', qty: '2', rate: '155' }, { desc: 'Tea 250 g', qty: '1', rate: '140' }] }],
      make: function (P, D, G) {
        var sub = 0, rows = [];
        G.items.forEach(function (l) { if (!l.desc && !num(l.rate)) return; var a = r2(num(l.qty) * num(l.rate)); sub += a; rows.push([rows.length + 1, l.desc, num(l.qty), num(l.rate), a]); });
        sub = r2(sub); var disc = Math.min(sub, num(D.disc)), raw = r2(sub - disc), gt = Math.round(raw), q = D.kind === 'Quotation';
        var tot = [['Sub-total', sub]]; if (disc) tot.push(['Discount', -disc]); if (gt !== raw) tot.push(['Round off', r2(gt - raw)]); tot.push([q ? 'Total' : 'Amount payable', gt, 1]);
        return { file: (q ? 'Quotation-' : 'Bill-') + safe(D.no), title: q ? 'QUOTATION' : 'BILL / CASH MEMO', org: P.name,
          sub: [P.addr, [P.phone ? 'Phone: ' + P.phone : '', P.gstin ? 'GSTIN: ' + P.gstin : ''].filter(Boolean).join('  ·  ')],
          kv: [[q ? 'Quotation No.' : 'Bill No.', D.no], ['Date', dmy(D.date)], ['Customer', D.cname || (q ? '' : 'Cash')], [q ? 'Valid until' : 'Phone', q ? dmy(D.valid) : D.cphone]],
          cols: [['#', 'c'], ['Item', ''], ['Qty', 'r'], ['Rate', 'r$'], ['Amount', 'r$']], rows: rows, totals: tot, words: words(gt),
          notes: D.terms ? [D.terms] : [], sign: ['', 'For ' + (P.name || '')] };
      }
    },
    report: {
      name: 'report card', no: 'roll',
      profile: [['name', 'School / centre name', 'Govt. Senior Secondary School'], ['addr', 'Address or board', 'Ludhiana, Punjab', 'wide']],
      fields: [['student', 'Student name', 'Aarav Sharma', 'wide'], ['roll', 'Roll number', '1'], ['cls', 'Class & section', 'VIII-B'], ['exam', 'Exam', 'Half-yearly examination'],
        ['session', 'Session', ''], ['present', 'Days present', '', 'number'], ['days', 'Working days', '', 'number'], ['remarks', 'Teacher\'s remarks', 'Works sincerely. Keep practising mathematics every day.', 'wide']],
      groups: [{ key: 'subjects', title: 'Subjects', add: 'Add subject', cols: [['sub', 'Subject', '', 'wide'], ['max', 'Max marks', '100', 'number'], ['got', 'Marks obtained', '', 'number']],
        sample: [{ sub: 'English', max: '100', got: '78' }, { sub: 'Hindi', max: '100', got: '84' }, { sub: 'Mathematics', max: '100', got: '69' }, { sub: 'Science', max: '100', got: '81' }, { sub: 'Social Science', max: '100', got: '74' }] }],
      make: function (P, D, G) {
        var tm = 0, tg = 0, rows = [], fail = false;
        G.subjects.forEach(function (l) {
          if (!l.sub) return; var mx = num(l.max), g = num(l.got), p = mx ? g / mx * 100 : 0;
          tm += mx; tg += g; if (p < 33) fail = true;
          rows.push([l.sub, mx, g, Math.round(p) + '%', gradeOf(p)]);
        });
        var pc = tm ? tg / tm * 100 : 0;
        var att = num(D.days) ? num(D.present) + ' of ' + num(D.days) + ' days (' + Math.round(num(D.present) / num(D.days) * 100) + '%)' : '';
        return { file: 'Report-Card-' + safe(D.student), title: 'REPORT CARD' + (D.exam ? '  ·  ' + D.exam.toUpperCase() : ''), org: P.name, sub: [P.addr],
          kv: [['Student', D.student], ['Roll No.', D.roll], ['Class', D.cls], ['Session', D.session], ['Attendance', att], ['', '']],
          cols: [['Subject', ''], ['Maximum', 'r'], ['Obtained', 'r'], ['Percent', 'r'], ['Grade', 'c']], rows: rows,
          totals: [['Total', tg + ' / ' + tm], ['Percentage', (Math.round(pc * 10) / 10) + '%'], ['Overall grade', rows.length ? gradeOf(pc) : ''], ['Result', rows.length ? (fail ? 'Needs improvement' : 'Pass') : '', 1]],
          notes: D.remarks ? ['Remarks: ' + D.remarks] : [], sign: ['Class teacher', 'Principal', 'Parent'] };
      }
    },
    payslip: {
      name: 'salary slip', no: 'code',
      profile: [['name', 'Company name', 'Sharma Traders Pvt. Ltd.'], ['addr', 'Address', '12, Main Bazaar, Ludhiana, Punjab 141001', 'wide']],
      fields: [['month', 'Salary month', today().slice(0, 7), 'month'], ['emp', 'Employee name', 'Rahul Mehra'], ['code', 'Employee code', 'E001'], ['desig', 'Designation', 'Accountant'],
        ['dept', 'Department', 'Accounts'], ['pan', 'PAN', ''], ['uan', 'UAN', ''], ['acct', 'Bank account', ''], ['paid', 'Days paid', '26', 'number'], ['wdays', 'Working days', '26', 'number'],
        ['pf', 'Calculate PF (12% of basic, up to ₹15,000)', true, 'check'], ['esi', 'Calculate ESI (0.75% if gross ≤ ₹21,000)', true, 'check']],
      groups: [{ key: 'earn', title: 'Earnings (for the days paid)', add: 'Add earning', cols: [['label', 'Earning', '', 'wide'], ['amt', 'Amount (₹)', '', 'number']],
        sample: [{ label: 'Basic', amt: '14000' }, { label: 'House rent allowance', amt: '5600' }, { label: 'Other allowances', amt: '2400' }] },
        { key: 'ded', title: 'Other deductions', add: 'Add deduction', cols: [['label', 'Deduction', '', 'wide'], ['amt', 'Amount (₹)', '', 'number']],
        sample: [{ label: 'Professional tax', amt: '200' }, { label: 'TDS', amt: '' }, { label: 'Advance', amt: '' }] }],
      make: function (P, D, G) {
        var earn = G.earn.filter(function (l) { return l.label || num(l.amt); }), gross = 0, basic = 0;
        earn.forEach(function (l) { gross += num(l.amt); if (/basic/i.test(l.label)) basic += num(l.amt); });
        var ded = [];
        if (D.pf) ded.push(['Provident fund (12%)', Math.round(Math.min(basic, 15000) * 0.12)]);
        if (D.esi && gross > 0 && gross <= 21000) ded.push(['ESI (0.75%)', Math.ceil(gross * 0.0075)]);
        G.ded.forEach(function (l) { if (num(l.amt)) ded.push([l.label || 'Deduction', num(l.amt)]); });
        var td = 0; ded.forEach(function (d) { td += d[1]; });
        var n = Math.max(earn.length, ded.length), rows = [];
        for (var i = 0; i < n; i++) rows.push([earn[i] ? earn[i].label : '', earn[i] ? num(earn[i].amt) : '', ded[i] ? ded[i][0] : '', ded[i] ? ded[i][1] : '']);
        rows.push(['Gross earnings', r2(gross), 'Total deductions', r2(td)]);
        var net = r2(gross - td);
        return { file: 'Payslip-' + safe(D.emp) + '-' + (D.month || ''), title: 'PAYSLIP FOR ' + monthName(D.month).toUpperCase(), org: P.name, sub: [P.addr],
          kv: [['Employee', D.emp], ['Employee code', D.code], ['Designation', D.desig], ['Department', D.dept], ['PAN', D.pan], ['UAN', D.uan], ['Bank account', D.acct], ['Days paid', num(D.paid) + ' of ' + num(D.wdays)]],
          cols: [['Earnings', ''], ['Amount', 'r$'], ['Deductions', ''], ['Amount', 'r$']], rows: rows, lastBold: true,
          totals: [['Net pay', net, 1]], words: words(net), notes: ['This is a computer-generated payslip.'], sign: ['Employee signature', 'Authorised signatory'] };
      }
    },
    expense: {
      name: 'expense claim', no: 'purpose',
      profile: [['name', 'Your name', 'Rahul Mehra'], ['code', 'Employee code', 'E001'], ['dept', 'Department', 'Accounts'], ['org', 'Company', 'Sharma Traders Pvt. Ltd.']],
      fields: [['purpose', 'Purpose of the expense', 'Client visit to Delhi for GST audit', 'wide'], ['date', 'Claim date', today(), 'date'], ['advance', 'Advance taken (₹)', '', 'number'], ['mgr', 'Approving manager', '']],
      groups: [{ key: 'bills', title: 'Bills', add: 'Add bill',
        cols: [['date', 'Date', '', 'date'], ['cat', 'Category', 'Travel', 'select', ['Travel', 'Lodging', 'Food', 'Local conveyance', 'Phone & internet', 'Office supplies', 'Other']], ['desc', 'Description', '', 'wide'], ['bill', 'Bill no.', ''], ['amt', 'Amount (₹)', '', 'number']],
        sample: [{ date: today(), cat: 'Travel', desc: 'Train Ludhiana to New Delhi', bill: 'PNR 2451', amt: '1265' }, { date: today(), cat: 'Lodging', desc: 'Hotel, one night', bill: 'H-1093', amt: '3200' }, { date: today(), cat: 'Food', desc: 'Meals', bill: '', amt: '860' }] }],
      make: function (P, D, G) {
        var rows = [], tot = 0, by = {};
        G.bills.forEach(function (l) { if (!l.desc && !num(l.amt)) return; var a = num(l.amt); tot += a; by[l.cat] = (by[l.cat] || 0) + a; rows.push([rows.length + 1, dmy(l.date), l.cat, l.desc, l.bill, a]); });
        tot = r2(tot); var adv = num(D.advance), net = r2(tot - adv);
        var t = Object.keys(by).map(function (k) { return [k, r2(by[k])]; });
        t.push(['Total claimed', tot, 1]); if (adv) t.push(['Less advance', -adv]); t.push([net < 0 ? 'Refund to company' : 'Payable to employee', Math.abs(net), 1]);
        return { file: 'Expense-Claim-' + safe(P.name) + '-' + (D.date || ''), title: 'EXPENSE REIMBURSEMENT CLAIM', org: P.org || P.name, sub: [],
          kv: [['Employee', P.name], ['Employee code', P.code], ['Department', P.dept], ['Claim date', dmy(D.date)], ['Purpose', D.purpose], ['Manager', D.mgr]],
          cols: [['#', 'c'], ['Date', ''], ['Category', ''], ['Description', ''], ['Bill no.', ''], ['Amount', 'r$']], rows: rows, totals: t, words: words(Math.abs(net)),
          notes: ['I confirm these expenses were incurred for official work and the bills are attached.'], sign: ['Employee', 'Approved by' + (D.mgr ? ' (' + D.mgr + ')' : ''), 'Accounts'] };
      }
    }
  };

  /* ---------------- rendering the document model */
  function cellText(v, align) {
    if (typeof v === 'number') return align.indexOf('$') >= 0 ? money(v) : (Number.isInteger(v) ? F0.format(v) : String(v));
    return v == null ? '' : String(v);
  }
  function renderDoc(m) {
    var el = h('div', { 'class': 'tk-doc', id: 'tk-doc' });
    el.appendChild(h('h3', { text: m.org || 'Your name' }));
    m.sub.forEach(function (s) { if (s) el.appendChild(h('p', { 'class': 'sub', text: s })); });
    el.appendChild(h('div', { 'class': 'bar', text: m.title }));
    var kv = h('table', { 'class': 'kv' }), tb = h('tbody');
    for (var i = 0; i < m.kv.length; i += 2) {
      var a = m.kv[i], b = m.kv[i + 1] || ['', ''];
      tb.appendChild(h('tr', {}, [h('td', { text: a[0] }), h('td', { 'class': 'b', text: a[1] || '' }), h('td', { text: b[0] }), h('td', { 'class': 'b', text: b[1] || '' })]));
    }
    kv.appendChild(tb); el.appendChild(kv);
    var t = h('table'), th = h('tr');
    m.cols.forEach(function (c) { th.appendChild(h('th', { 'class': c[1].indexOf('r') >= 0 ? 'r' : null, text: c[0] })); });
    t.appendChild(h('thead', {}, [th]));
    var body = h('tbody');
    if (!m.rows.length) body.appendChild(h('tr', {}, [h('td', { colspan: m.cols.length, text: 'Add lines on the left to see them here.' })]));
    m.rows.forEach(function (r, ri) {
      var tr = h('tr', { 'class': m.lastBold && ri === m.rows.length - 1 ? 'b' : null });
      r.forEach(function (v, i) { var al = m.cols[i][1]; tr.appendChild(h('td', { 'class': al.indexOf('r') >= 0 ? 'r' : null, text: cellText(v, al) })); });
      body.appendChild(tr);
    });
    t.appendChild(body); el.appendChild(t);
    if (m.totals.length) {
      var tt = h('table', { 'class': 'tot' }), tbb = h('tbody');
      m.totals.forEach(function (x) { tbb.appendChild(h('tr', { 'class': x[2] ? 'gt' : null }, [h('td', { text: x[0] }), h('td', { 'class': 'r', text: typeof x[1] === 'number' ? money(x[1]) : x[1] })])); });
      tt.appendChild(tbb); el.appendChild(tt);
    }
    if (m.words) el.appendChild(h('p', { 'class': 'words', text: m.words }));
    m.notes.forEach(function (n) { el.appendChild(h('p', { 'class': 'sub', style: 'text-align:left', text: n })); });
    el.appendChild(h('div', { 'class': 'sign' }, m.sign.map(function (s) { return h('div', { style: 'white-space:pre-line' }, ['\n\n' + (s ? '____________________\n' + s.split('\n').pop() : '')]); })));
    el.appendChild(h('p', { 'class': 'foot' }, [MADE]));   /* owner: just the site address under every form, on screen and in print */
    return el;
  }

  /* .xlsx */
  function toXlsx(m) {
    return loadLib('xlsx').then(function (ExcelJS) {
      var wb = new ExcelJS.Workbook(); wb.creator = 'Nivesh Sahayak'; wb.lastModifiedBy = 'Nivesh Sahayak'; wb.company = 'Nivesh Sahayak'; wb.description = CREDIT + ' · https://www.niveshsahayak.in/'; wb.keywords = 'Toolkit by GroupChatOn; Nivesh Sahayak; https://www.niveshsahayak.in/';
      var n = m.cols.length, ws = wb.addWorksheet('Document', { pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } } });
      ws.headerFooter.oddFooter = '&C&8' + MADE;
      var INR = '[>=10000000]₹ ##\\,##\\,##\\,##0.00;[>=100000]₹ ##\\,##\\,##0.00;₹ ##,##0.00';
      var blue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } }, thin = { style: 'thin', color: { argb: 'FFCBD5E1' } }, box = { top: thin, left: thin, bottom: thin, right: thin };
      var r = 1;
      function line(text, font, fill) { ws.mergeCells(r, 1, r, n); var c = ws.getCell(r, 1); c.value = text; c.font = font || {}; c.alignment = { horizontal: 'center', wrapText: true }; if (fill) c.fill = fill; r++; }
      line(m.org || '', { bold: true, size: 16, color: { argb: 'FF2563EB' } });
      m.sub.forEach(function (s) { if (s) line(s, { size: 9 }); });
      r++; line(m.title, { bold: true, color: { argb: 'FFFFFFFF' } }, blue); r++;
      for (var i = 0; i < m.kv.length; i += 2) {
        var a = m.kv[i], b = m.kv[i + 1] || ['', ''], half = Math.max(1, Math.floor(n / 2));
        ws.getCell(r, 1).value = a[0]; ws.getCell(r, 1).font = { color: { argb: 'FF55607A' } };
        if (half > 1) ws.mergeCells(r, 2, r, half); ws.getCell(r, 2).value = a[1] || ''; ws.getCell(r, 2).font = { bold: true };
        if (n > 2) { ws.getCell(r, half + 1).value = b[0]; ws.getCell(r, half + 1).font = { color: { argb: 'FF55607A' } }; if (n > half + 2) ws.mergeCells(r, half + 2, r, n); ws.getCell(r, half + 2).value = b[1] || ''; ws.getCell(r, half + 2).font = { bold: true }; }
        r++;
      }
      r++;
      m.cols.forEach(function (c, j) { var x = ws.getCell(r, j + 1); x.value = c[0]; x.font = { bold: true, color: { argb: 'FFFFFFFF' } }; x.fill = blue; x.border = box; x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
      r++;
      m.rows.forEach(function (row, ri) {
        row.forEach(function (v, j) { var x = ws.getCell(r, j + 1); x.value = v === '' ? null : v; x.border = box; if (typeof v === 'number' && m.cols[j][1].indexOf('$') >= 0) x.numFmt = INR; if (m.lastBold && ri === m.rows.length - 1) x.font = { bold: true }; if (m.cols[j][1].indexOf('r') >= 0) x.alignment = { horizontal: 'right' }; });
        r++;
      });
      r++;
      m.totals.forEach(function (x) {
        if (n > 2) ws.mergeCells(r, Math.max(1, n - 2), r, n - 1);
        var k = ws.getCell(r, Math.max(1, n - 2)), v = ws.getCell(r, n); k.value = x[0]; k.alignment = { horizontal: 'right' }; v.value = x[1];
        if (typeof x[1] === 'number') v.numFmt = INR; v.alignment = { horizontal: 'right' }; if (x[2]) { k.font = v.font = { bold: true }; } r++;
      });
      if (m.words) { r++; ws.mergeCells(r, 1, r, n); ws.getCell(r, 1).value = m.words; ws.getCell(r, 1).font = { bold: true }; r++; }
      m.notes.forEach(function (t) { ws.mergeCells(r, 1, r, n); var c = ws.getCell(r, 1); c.value = t; c.font = { size: 9, color: { argb: 'FF55607A' } }; c.alignment = { wrapText: true }; ws.getRow(r).height = 28; r++; });
      r += 2; m.sign.forEach(function (s, j) { var col = m.sign.length === 1 ? n : Math.round(1 + j * (n - 1) / Math.max(1, m.sign.length - 1)); ws.getCell(r, col).value = (s || '').split('\n').pop(); ws.getCell(r, col).font = { size: 9 }; });
      m.cols.forEach(function (c, j) {
        var w = c[0].length + 2; m.rows.forEach(function (row) { w = Math.max(w, String(cellText(row[j], c[1])).length + 2); });
        ws.getColumn(j + 1).width = Math.min(46, Math.max(8, w));
      });
      return wb.xlsx.writeBuffer().then(function (buf) { save(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), m.file + '.xlsx'); });
    });
  }

  /* .docx */
  function toDocx(m) {
    return loadLib('docx').then(function (D) {
      var P = D.Paragraph, R = D.TextRun, AL = D.AlignmentType;
      var bd = { style: D.BorderStyle.SINGLE, size: 4, color: 'CBD5E1' }, nob = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
      function p(text, o) { o = o || {}; return new P({ alignment: o.al, spacing: { after: o.after == null ? 60 : o.after }, children: [new R({ text: text || '', bold: o.b, size: o.size || 20, color: o.color })] }); }
      function cell(text, o) { o = o || {}; return new D.TableCell({ shading: o.fill ? { fill: o.fill, type: D.ShadingType.CLEAR, color: 'auto' } : undefined, borders: o.nob ? { top: nob, bottom: nob, left: nob, right: nob } : { top: bd, bottom: bd, left: bd, right: bd },
        children: [new P({ alignment: o.r ? AL.RIGHT : AL.LEFT, children: [new R({ text: text == null ? '' : String(text), bold: o.b, color: o.color, size: 18 })] })] }); }
      var kids = [p(m.org, { b: true, size: 32, color: '2563EB', al: AL.CENTER, after: 20 })];
      m.sub.forEach(function (s) { if (s) kids.push(p(s, { size: 17, al: AL.CENTER, after: 20 })); });
      kids.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: [new D.TableRow({ children: [cell(m.title, { fill: '2563EB', b: true, color: 'FFFFFF', nob: true })] })] }));
      kids.push(p('', { after: 60 }));
      var kvRows = [];
      for (var i = 0; i < m.kv.length; i += 2) { var a = m.kv[i], b = m.kv[i + 1] || ['', '']; kvRows.push(new D.TableRow({ children: [cell(a[0], { nob: true, color: '55607A' }), cell(a[1], { nob: true, b: true }), cell(b[0], { nob: true, color: '55607A' }), cell(b[1], { nob: true, b: true })] })); }
      kids.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: kvRows })); kids.push(p('', { after: 60 }));
      var rows = [new D.TableRow({ tableHeader: true, children: m.cols.map(function (c) { return cell(c[0], { fill: '2563EB', b: true, color: 'FFFFFF', r: c[1].indexOf('r') >= 0 }); }) })];
      m.rows.forEach(function (row, ri) { rows.push(new D.TableRow({ children: row.map(function (v, j) { return cell(cellText(v, m.cols[j][1]), { r: m.cols[j][1].indexOf('r') >= 0, b: m.lastBold && ri === m.rows.length - 1 }); }) })); });
      kids.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: rows })); kids.push(p('', { after: 60 }));
      if (m.totals.length) kids.push(new D.Table({ width: { size: 50, type: D.WidthType.PERCENTAGE }, alignment: AL.RIGHT,
        rows: m.totals.map(function (x) { return new D.TableRow({ children: [cell(x[0], { nob: true, b: x[2] }), cell(typeof x[1] === 'number' ? money(x[1]) : x[1], { nob: true, r: true, b: x[2] })] }); }) }));
      if (m.words) kids.push(p(m.words, { b: true, after: 80 }));
      m.notes.forEach(function (t) { kids.push(p(t, { size: 17, color: '55607A' })); });
      kids.push(p('', { after: 600 }));
      kids.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: [new D.TableRow({ children: m.sign.map(function (s) { return cell(s ? '____________________\n' + s.split('\n').pop() : '', { nob: true }); }) })] }));
      var d = new D.Document({ creator: 'Nivesh Sahayak', lastModifiedBy: 'Nivesh Sahayak', keywords: 'Toolkit by GroupChatOn; Nivesh Sahayak; https://www.niveshsahayak.in/', description: CREDIT + ' · https://www.niveshsahayak.in/', title: m.title,
        sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 900, bottom: 900, left: 1000, right: 1000 } } },
          footers: { default: new D.Footer({ children: [p(MADE, { size: 14, color: '94A3B8', al: AL.CENTER })] }) }, children: kids }] });
      return D.Packer.toBlob(d).then(function (blob) { save(blob, m.file + '.docx'); });
    });
  }

  /* ---------------- the form */
  function field(def, val, onChange) {
    var key = def[0], lab = def[1], type = def[3] || 'text', opts = def[4], wide = def[3] === 'wide';
    if (type === 'check') {
      var cb = h('input', { type: 'checkbox', checked: val ? true : null, onchange: function () { onChange(key, cb.checked); } });
      return h('label', { 'class': 'tk-chk' }, [cb, lab]);
    }
    var inp;
    if (type === 'select') {
      inp = h('select', { onchange: function () { onChange(key, inp.value); } }, opts.map(function (o) { return h('option', { value: o, selected: String(o) === String(val) ? true : null, text: o }); }));
    } else {
      inp = h('input', { type: type === 'wide' ? 'text' : type, value: val == null ? '' : val, inputmode: type === 'number' ? 'decimal' : null, step: type === 'number' ? 'any' : null,
        oninput: function () { onChange(key, inp.value); } });
    }
    return h('label', { 'class': 'tk-f' + (wide ? ' wide' : ''), style: wide ? 'grid-column:1/-1' : null }, [lab, inp]);
  }

  function mount(box) {
    var id = box.getAttribute('data-tool'), T = TOOLS[id]; if (!T) return;
    var spot = box.querySelector('.tk-mount');
    var PK = 'tk-prof-' + id, DK = 'tk-draft-' + id, NK = 'tk-last-' + id;
    function defaults(list) { var o = {}; list.forEach(function (f) { o[f[0]] = f[2]; }); return o; }
    var prof = Object.assign(defaults(T.profile), store(PK) || {});
    var draft = store(DK);
    var data, groups;
    function fresh(sample) {
      data = defaults(T.fields); groups = {};
      T.groups.forEach(function (g) { groups[g.key] = sample ? JSON.parse(JSON.stringify(g.sample)) : [defaults(g.cols)]; });
      var last = store(NK); if (last && T.no === 'no') data.no = nextNo(last);
    }
    if (draft && draft.data) { data = Object.assign(defaults(T.fields), draft.data); groups = draft.groups; } else fresh(true);
    var form = h('form', { 'class': 'tk-form', novalidate: true, onsubmit: function (e) { e.preventDefault(); } });
    var out = h('div', { 'class': 'tk-out' });
    var msg = h('p', { 'class': 'tk-msg', role: 'status', 'aria-live': 'polite' });
    var preview = h('div');
    var timer;
    function persist() { store(PK, prof); store(DK, { data: data, groups: groups }); }
    function update() { clearTimeout(timer); timer = setTimeout(function () { persist(); preview.innerHTML = ''; preview.appendChild(renderDoc(T.make(prof, data, groups))); }, 120); }
    function build() {
      form.innerHTML = '';
      form.appendChild(h('fieldset', {}, [h('legend', { text: 'Your details (saved on this device)' }), h('div', { 'class': 'tk-row' }, T.profile.map(function (f) { return field(f, prof[f[0]], function (k, v) { prof[k] = v; update(); }); }))]));
      form.appendChild(h('fieldset', {}, [h('legend', { text: 'This ' + T.name }), h('div', { 'class': 'tk-row' }, T.fields.map(function (f) { return field(f, data[f[0]], function (k, v) { data[k] = v; update(); }); }))]));
      T.groups.forEach(function (g) {
        var list = h('div', { 'class': 'tk-lines' });
        groups[g.key].forEach(function (line, i) {
          list.appendChild(h('div', { 'class': 'tk-line' }, [
            h('button', { type: 'button', 'class': 'tk-x', 'aria-label': 'Remove line ' + (i + 1), title: 'Remove', text: '×', onclick: function () { groups[g.key].splice(i, 1); if (!groups[g.key].length) groups[g.key].push(defaults(g.cols)); build(); update(); } }),
            h('div', { 'class': 'tk-row' }, g.cols.map(function (c) { return field(c, line[c[0]], function (k, v) { line[k] = v; update(); }); }))]));
        });
        form.appendChild(h('fieldset', {}, [h('legend', { text: g.title }), list,
          h('div', { 'class': 'tk-acts' }, [h('button', { type: 'button', 'class': 'btn btn-2', text: '+ ' + g.add, onclick: function () { groups[g.key].push(defaults(g.cols)); build(); update(); var f = form.querySelectorAll('.tk-lines')[T.groups.indexOf(g)]; var ins = f && f.lastChild.querySelector('input,select'); if (ins) ins.focus(); } })])]));
      });
      form.appendChild(h('div', { 'class': 'tk-acts' }, [
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'New ' + T.name + ' (keep my details)', onclick: function () { fresh(false); build(); update(); msg.textContent = 'Cleared. Your details are kept.'; } }),
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Fill sample', onclick: function () { fresh(true); build(); update(); } })]));
    }
    function done(kind) { if (T.no === 'no' && data.no) store(NK, data.no); msg.textContent = kind + ' ready. Check your downloads folder.'; }
    function fail(e) { msg.textContent = (e && e.message) || 'Something went wrong. Please try again.'; }
    var icon = function (d) { return h('span', { 'aria-hidden': 'true', text: d }); };
    out.appendChild(h('div', { 'class': 'tk-out-acts' }, [
      h('button', { type: 'button', 'class': 'btn', onclick: function () { persist(); if (T.no === 'no' && data.no) store(NK, data.no); window.print(); } }, [icon('🖨'), ' Print / PDF']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { var b = e.currentTarget; b.disabled = true; msg.textContent = 'Making the Excel file…'; toXlsx(T.make(prof, data, groups)).then(function () { done('Excel file'); }, fail).then(function () { b.disabled = false; }); } }, [icon('📗'), ' Excel']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { var b = e.currentTarget; b.disabled = true; msg.textContent = 'Making the Word file…'; toDocx(T.make(prof, data, groups)).then(function () { done('Word file'); }, fail).then(function () { b.disabled = false; }); } }, [icon('📘'), ' Word'])]));
    out.appendChild(msg); out.appendChild(preview);
    build();
    spot.appendChild(h('div', { 'class': 'tk-app' }, [form, out]));
    preview.appendChild(renderDoc(T.make(prof, data, groups)));
    doc.body.classList.add('pg-tool');
  }
  /* tools in their own files (assets/tools/<id>.js, loaded after this one) call TKTools.register(id, definition) with the same
     shape as TOOLS above; lib gives them the helpers. Mounting waits for DOMContentLoaded, which comes after every deferred script. */
  window.TKTools = { words: words, tools: TOOLS, CREDIT: CREDIT, MADE: MADE, register: function (id, def) { TOOLS[id] = def; },
    lib: { h: h, num: num, r2: r2, money: money, today: today, dmy: dmy, monthName: monthName, words: words, nextNo: nextNo, gradeOf: gradeOf, store: store, STATES: STATES, GST: GST } };
  function mountAll() { Array.prototype.forEach.call(doc.querySelectorAll('.tk-tool[data-tool]'), mount); }
  if (doc.readyState === 'complete') mountAll(); else doc.addEventListener('DOMContentLoaded', mountAll);
})();
