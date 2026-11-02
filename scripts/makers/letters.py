"""Letters & Applications toolkit (kit id: letters): four Word files of letter formats.
   The wording comes from the online letter writer (assets/tools/letter-writer.js, the DATA literal between /*DATA*/ and /*END*/),
   so the Word templates and the online tool always say the same thing. Blanks are written as [Label] and highlighted yellow."""
from make_files import *
import ast, re
from docx.enum.text import WD_COLOR_INDEX, WD_BREAK

_JS = open(os.path.join(HERE, 'assets', 'tools', 'letter-writer.js'), encoding='utf-8').read()
LD = ast.literal_eval(_JS[_JS.index('/*DATA*/') + 8:_JS.index('/*END*/')])
LF, LR, LT, LTONE = LD['F'], LD['R'], {t['id']: t for t in LD['T']}, LD['TONE']
HIFONT = 'Nirmala UI'
HI_LABEL = {'sname': 'आपका नाम', 'cls': 'कक्षा', 'roll': 'अनुक्रमांक', 'd1': 'आरंभ तिथि', 'd2': 'अंतिम तिथि', 'illness': 'बीमारी का नाम',
            'child': 'बच्चे का नाम', 'reason': 'कारण', 'event': 'समारोह', 'place': 'स्थान', 'earner': 'कमाने वाले सदस्य', 'job': 'उनका काम',
            'income': 'मासिक आय', 'marks': 'अंक', 'since': 'वर्ष', 'session': 'सत्र', 'purpose': 'प्रयोजन', 'doc': 'दस्तावेज का नाम',
            'notice': 'सूचना अवधि', 'ldate': 'अंतिम कार्य दिवस', 'sdesig': 'आपका पद', 'dept': 'विभाग', 'cover': 'सहकर्मी का नाम',
            'acno': 'खाता संख्या', 'toacc': 'दूसरी खाता संख्या', 'newaddr': 'नया पता', 'card': 'कार्ड के अंतिम 4 अंक', 'area': 'मोहल्ला / वार्ड',
            'details': 'अन्य विवरण', 'conno': 'उपभोक्ता / कनेक्शन संख्या', 'item': 'वस्तु का विवरण', 'value': 'अनुमानित मूल्य', 'time': 'समय',
            'person': 'व्यक्ति का नाम', 'prel': 'आपसे संबंध', 'idp': 'उनका पहचान प्रमाण', 'task': 'कार्य', 'ndays': 'दिनों की संख्या',
            'sphone': 'मोबाइल नंबर', 'saddr': 'आपका पता', 'semail': 'ईमेल', 'empid': 'कर्मचारी कोड', 'what': 'क्या हुआ', 'month': 'माह',
            'amt': 'राशि', 'by': 'अंतिम तिथि'}

E_HI = {'Medical certificate': 'चिकित्सा प्रमाण पत्र', 'Income certificate': 'आय प्रमाण पत्र', 'Address proof (self-attested)': 'नए पते का प्रमाण (स्व-सत्यापित)',
        'Unused cheque leaves\nDebit card\nCopy of ID proof': 'अप्रयुक्त चेक\nडेबिट कार्ड\nपहचान प्रमाण की प्रति', 'Copy of the latest bill': 'नवीनतम बिल की प्रति'}

def _label(t, k, hi):
    if k == 'ndays' and not hi: return 'number of days'
    if hi and k in HI_LABEL: return HI_LABEL[k]
    lab = (t.get('lb') or {}).get(k) or (LF[k][0] if k in LF else k)
    return re.sub(r' \(.*\)$', '', lab)

def _fill(t, s, hi):
    """Template → text with [Label] blanks. Optional parts stay in, both gender forms are shown as a/b."""
    li = 1 if hi else 0
    def key(m):
        k = m.group(1)
        if k == 'amtw': return '[amount in words]'
        if k == 'ndays': return '[' + _label(t, k, hi) + ']'
        d = LF.get(k)
        if d and d[1] == 's': return '[' + ' / '.join(o[li] for o in d[4]) + ']'
        if d and d[1] == 'm': return '₹ [' + _label(t, k, hi) + ']'
        if d and d[1] == 'd': return '[' + (HI_LABEL.get(k, 'दिनांक') if hi else _label(t, k, hi)) + ']'
        if d and d[1] == 'l': return '1. [' + _label(t, k, hi) + ']\n2. [ ]\n3. [ ]'
        return '[' + _label(t, k, hi) + ']'
    s = re.sub(r'\(\((?:\w+:)?([^|()]*)\|([^()]*)\)\)', lambda m: m.group(1) if m.group(1) == m.group(2) else m.group(1) + '/' + m.group(2), s)
    s = re.sub(r'\[\[([\s\S]*?)\]\]', lambda m: m.group(1), s)
    s = re.sub(r'\{(\w+)\}', key, s)
    return re.sub(r' {2,}', ' ', s)

def letter_blocks(tid, hi=False):
    """The letter as (text, style) blocks in the Indian layout, formal tone, default optional sentences, every detail a [blank]."""
    t = LT[tid]; r = LR[t['r']]; li = 1 if hi else 0
    role = dict(r); role['x'] = t.get('x', r['x']); cl = t.get('cl', r['cl'])
    f = lambda s: _fill(t, s, hi)
    B = []
    if r.get('head'):
        B.append(('[Business / organisation name]' if not hi else '[संस्था का नाम]', 'head'))
        B.append(('[Address] · Phone [ ] · [Email]' if not hi else '[पता] · फ़ोन [ ] · [ईमेल]', 'headsub'))
        B.append((('Ref: [ ]\nDate: [DD-MM-YYYY]') if not hi else 'पत्रांक: [ ]\nदिनांक: [दिनांक]', ''))
    else:
        B.append(('Date: [DD-MM-YYYY]' if not hi else 'दिनांक: [दिनांक]', ''))
    if t.get('tw'):
        subj = f(t['s'][li])
        B.append((subj if hi else subj.upper(), 'title'))
        B.append(('सर्वसंबंधित के लिए' if hi else 'TO WHOM IT MAY CONCERN', 'center'))
    else:
        to = [f(t['to'][li])]
        if f(t['org'][li]): to.append(f(t['org'][li]))
        to.append('[शहर]' if hi else '[City]')
        B.append((('सेवा में,\n' if hi else 'To,\n') + '\n'.join(to), ''))
        B.append((('विषय: ' if hi else 'Subject: ') + f(t['s'][li]), 'bold'))
        named = t.get('rs') == 'n'
        if t.get('sal'): sal = f(t['sal'][li])
        elif hi: sal = ('प्रिय ' if cl == 'warm' else 'आदरणीय ') + '[प्राप्तकर्ता का नाम] जी' if named else 'महोदय/महोदया'
        else: sal = 'Dear [Recipient\'s name]' if named else ('Respected Sir/Madam' if r.get('resp') else 'Dear Sir/Madam')
        B.append((sal + ',', ''))
    paras = [f(p) for p in t['b'][li]]
    opts = ' '.join(f(o[2 + li]) for o in t.get('o', []) if o[4])
    if opts: paras.insert(t['oa'] if 'oa' in t else len(paras), opts)
    tone = (t.get('t') or {}).get('formal') or (LTONE.get(t['k']) or {}).get('formal')
    if tone: paras.append(f(tone[li]))
    B += [(p, '') for p in paras]
    if not r.get('head') and cl != 'warm' and t['k'] in ('req', 'comp', 'info'): B.append(('सधन्यवाद,' if hi else 'Thanking you,', ''))
    if hi:
        close = {'obed': 'आपका आज्ञाकारी शिष्य / आपकी आज्ञाकारिणी शिष्या', 'warm': 'सादर,'}.get(cl) or ('प्रार्थी / प्रार्थिनी' if cl == 'app' and t['k'] in ('req', 'comp') else 'भवदीय / भवदीया')
    else:
        close = {'obed': 'Yours obediently,', 'warm': 'With warm regards,'}.get(cl) or ('Yours sincerely,' if t.get('sal') or t.get('rs') == 'n' else 'Yours faithfully,')
    sig = [] if t.get('tw') else [close]
    if r.get('head'): sig.append('[संस्था का नाम] की ओर से' if hi else 'For [Business / organisation name]')
    sig += ['', '', '[' + _label(t, 'sname', hi) + ']']
    LBL = {'cls': ('Class: ', 'कक्षा: '), 'roll': ('Roll No.: ', 'अनुक्रमांक: '), 'dept': ('Department: ', 'विभाग: '), 'empid': ('Employee ID: ', 'कर्मचारी कोड: '),
           'sphone': ('Mobile: ', 'मोबाइल: '), 'semail': ('Email: ', 'ईमेल: '), 'flat': ('Flat: ', 'फ्लैट: ')}
    for k in role['x']:
        if k == '_parent': sig.append('[बच्चे का नाम], कक्षा [ ] के अभिभावक' if hi else 'Parent of [Child\'s name], Class [ ]')
        elif k == 'saddr': sig.append('[' + _label(t, 'saddr', hi) + ']')
        elif k == 'sdesig': sig.append('[' + _label(t, 'sdesig', hi) + ']')
        elif k == 'sorg': sig.append('[Organisation]')
        elif k in LBL: sig.append(LBL[k][li] + '[ ]')
    if t['c'] == 'bank' and 'acno' in t['f']: sig.append(('खाता संख्या: ' if hi else 'A/c No.: ') + '[ ]')
    B.append(('\n'.join(sig), 'sig'))
    enc = E_HI.get(t.get('e')) if hi else t.get('e')
    if enc:
        B.append((('संलग्न:\n' if hi else 'Enclosures:\n') + '\n'.join('%d. %s' % (i + 1, s) for i, s in enumerate(enc.split('\n'))), 'small'))
    return B

def _fields_line(tid, hi):
    t = LT[tid]; r = LR[t['r']]; ks = []
    for k in list(t.get('a', r['a'])) + t['f']:
        if k not in ks and k in LF and LF[k][1] not in ('g', 'k'): ks.append(k)
    return ', '.join(_label(t, k, hi) for k in ks)

def _runs(p, text, hi=False, bold=False, size=None, color=None):
    """Add text to a paragraph: line breaks kept, [blanks] highlighted yellow."""
    for li, line in enumerate(text.split('\n')):
        if li: p.add_run().add_break()
        for j, part in enumerate(re.split(r'(\[[^\[\]\n]{1,70}\])', line)):
            if not part: continue
            r = p.add_run(part); r.bold = bold
            if size: r.font.size = Pt(size)
            if color: r.font.color.rgb = RGBColor.from_string(color)
            if j % 2: r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            if hi:   # Devanagari is a complex script: Word uses the cs font, szCs and bCs for it
                r.font.name = HIFONT
                rpr = r._element.get_or_add_rPr(); rf = rpr.find(qn('w:rFonts'))
                for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'): rf.set(qn(a), HIFONT)
                b = rpr.find(qn('w:b'))
                if b is not None: b.addnext(OxmlElement('w:bCs'))
                sz = rpr.find(qn('w:sz'))
                if sz is not None: x = OxmlElement('w:szCs'); x.set(qn('w:val'), sz.get(qn('w:val'))); sz.addnext(x)

def _p(d, text='', hi=False, bold=False, size=None, color=None, align=None, after=6):
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(after)
    if align == 'c': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text: _runs(p, text, hi, bold, size, color)
    return p

def write_letter(d, tid, hi, n):
    t = LT[tid]
    _p(d, '%d. %s' % (n, t['h'] if hi else t['n']), hi, bold=True, size=14, color=BLUE, after=2)
    _p(d, ('भरें: ' if hi else 'Fill in: ') + _fields_line(tid, hi), hi, size=9, color=MUTED, after=10)
    for text, st in letter_blocks(tid, hi):
        if st == 'head': _p(d, text, hi, bold=True, size=16, color=BLUE, align='c', after=0)
        elif st == 'headsub':
            p = _p(d, text, hi, size=9, align='c', after=10)
            pPr = p._p.get_or_add_pPr(); b = OxmlElement('w:pBdr'); bt = OxmlElement('w:bottom')
            for k, v in (('w:val', 'single'), ('w:sz', '10'), ('w:space', '1'), ('w:color', BLUE)): bt.set(qn(k), v)
            b.append(bt); pPr.append(b)
        elif st == 'title': _p(d, text, hi, bold=True, size=13, align='c', after=4)
        elif st == 'center': _p(d, text, hi, bold=True, align='c', after=10)
        elif st == 'bold': _p(d, text, hi, bold=True, after=8)
        elif st == 'sig': _p(d, text, hi, after=8)
        elif st == 'small': _p(d, text, hi, size=9.5, after=4)
        else: _p(d, text, hi, after=8)

def _contents(d, ids, hi):
    _p(d, 'इस फ़ाइल में' if hi else 'In this file', hi, bold=True, size=13, color=BLUE, after=4)
    t = d.add_table(rows=0, cols=2); t.style = 'Table Grid'
    for i, tid in enumerate(ids):
        c = t.add_row().cells; c[0].text = str(i + 1); c[0].width = Cm(1.2); c[1].width = Cm(15.5)
        c[1].text = ''; _runs(c[1].paragraphs[0], LT[tid]['h'] if hi else LT[tid]['n'], hi, size=10)
        for r in c[0].paragraphs[0].runs: r.font.size = Pt(10)
    _p(d, after=4)

def _bullets(d, items, hi=False):
    for s in items:
        p = d.add_paragraph(style='List Bullet'); p.paragraph_format.space_after = Pt(2); _runs(p, s, hi, size=10.5)

def _grid(d, headers, rows, widths, hi=False):
    t = table(d, headers, rows, widths)
    if hi:
        for row in t.rows[1:]:
            for c in row.cells:
                txt = c.text; c.text = ''; _runs(c.paragraphs[0], txt, True, size=10)
    return t

def _brk(d): d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def _templates(d, ids, hi):
    for i, tid in enumerate(ids):
        _brk(d); write_letter(d, tid, hi, i + 1)

def _cover(d, title, sub, hi=False):
    _p(d, title, hi, bold=True, size=22, color=BLUE, after=2)
    _p(d, sub, hi, size=11, color=MUTED, after=12)

def _howto_box(d, hi=False):
    tips = (['पीले [कोष्ठक] में लिखी जगह पर अपनी जानकारी लिखें और पीला रंग हटा दें (Home › Text Highlight Colour › No Colour)।',
             'जहाँ दो रूप दिए हैं (रहा/रही, का छात्र/की छात्रा), वहाँ जो आप पर लागू हो वह रखें और दूसरा मिटा दें।',
             'जो वाक्य आपके मामले में लागू न हो उसे हटा दें। हस्ताक्षर अपने नाम के ऊपर करें और एक प्रति अपने पास रखें।',
             'ऑनलाइन लेटर राइटर (toolkit.groupchaton.com) में 78 प्रकार के पत्र हिंदी और अंग्रेज़ी में अपने-आप भरे जाते हैं।'] if hi else
            ['Replace every yellow [bracket] with your own detail, then remove the highlight (Home › Text Highlight Colour › No Colour).',
             'Where two forms are shown (he/she, son/daughter), keep the one that fits and delete the other.',
             'Delete any sentence that does not fit your case. Sign above your name and keep a copy of what you send.',
             'The online letter writer at toolkit.groupchaton.com fills 78 letter types in English or Hindi from a short form.'])
    _p(d, 'इस फ़ाइल का उपयोग कैसे करें' if hi else 'How to use this file', hi, bold=True, size=13, color=BLUE, after=4)
    _bullets(d, tips, hi)
    _p(d, after=6)

def _hi_doc(d):
    st = d.styles['Normal']; st.font.name = HIFONT
    st.element.rPr.rFonts.set(qn('w:eastAsia'), HIFONT); st.element.rPr.rFonts.set(qn('w:cs'), HIFONT)
    sz = st.element.rPr.find(qn('w:sz')); x = OxmlElement('w:szCs'); x.set(qn('w:val'), sz.get(qn('w:val'))); sz.addnext(x)

GUIDE_IDS = ['sick-leave', 'leave-parent', 'fee-concession', 'tc-request', 'bonafide', 'job-application', 'casual-leave', 'sick-leave-office',
             'resignation', 'salary-cert-request', 'account-closure', 'bank-address-change', 'bank-statement', 'loan-application',
             'sanitation-complaint', 'rti', 'police-theft', 'authority-letter', 'society-complaint', 'payment-reminder']
HINDI_IDS = ['sick-leave', 'leave-parent', 'leave-family', 'fee-concession', 'tc-request', 'char-school', 'duplicate-doc', 'casual-leave',
             'resignation', 'account-closure', 'bank-address-change', 'cheque-book', 'card-block', 'sanitation-complaint', 'water-complaint',
             'electricity-complaint', 'police-theft', 'lost-report', 'char-request', 'authority-letter']
OFFICE_IDS = ['job-application', 'casual-leave', 'sick-leave-office', 'wfh', 'resignation', 'relieving-request', 'salary-cert-request', 'increment',
              'salary-advance', 'transfer-request', 'noc-employer', 'joining-report', 'maternity-leave', 'appointment', 'warning', 'experience-cert',
              'relieving', 'internship-cert', 'salary-cert', 'increment-letter']
BANKGOV_IDS = ['account-closure', 'bank-address-change', 'cheque-book', 'bank-statement', 'loan-application', 'card-block', 'stop-payment',
               'mobile-update', 'nominee', 'failed-txn', 'loan-noc', 'sanitation-complaint', 'road-light', 'water-complaint', 'electricity-complaint',
               'rti', 'police-theft', 'lost-report', 'passport-cover', 'aadhaar-cover', 'landlord-repair', 'vacate-notice', 'deposit-refund',
               'rent-increase', 'rent-reminder', 'society-complaint', 'society-noc']

def letters_guide():
    d = doc_new('letters-format-guide-word')
    _cover(d, 'Formal Letter Format Guide', 'How to lay out a formal letter (Indian, UK and US styles) and the 20 letters people write most often, ready to fill in.')
    _howto_box(d)
    _p(d, 'The 9 parts of a formal letter', bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['#', 'Part', 'What to write', 'Example'], [
        ['1', 'Sender\'s address', 'Your address. In Indian school applications it goes under your name at the end; in UK and US letters at the top.', 'House 24, Model Town, Ludhiana 141002'],
        ['2', 'Date', 'The day you write. Keep one style all through the letter.', 'Date: 03-11-2026 / 3 November 2026'],
        ['3', 'Receiver\'s address', 'Designation or name, office, city. Start with "To," in the Indian style.', 'To, The Branch Manager, State Bank of India, Model Town, Ludhiana'],
        ['4', 'Subject', 'One line that says what the letter is about. Bold or underlined.', 'Subject: Request for a new cheque book'],
        ['5', 'Salutation', 'Dear Sir/Madam, Respected Sir (school), or Dear Mr. Gupta when you know the name.', 'Respected Sir,'],
        ['6', 'Opening', 'Who you are and why you are writing, in one or two sentences.', 'I am a student of Class VIII-B of your school.'],
        ['7', 'Body', 'The facts, dates, numbers and exactly what you are asking for. Short paragraphs.', 'Kindly grant me leave from 03-11-2026 to 05-11-2026.'],
        ['8', 'Closing line', 'A polite line of thanks or the action you expect.', 'Thanking you,'],
        ['9', 'Complimentary close and signature', 'Yours faithfully / sincerely / obediently, your signature, name and details.', 'Yours faithfully, (signature) Riya Sharma'],
    ], [0.8, 3.2, 7.2, 5.8])
    _p(d, 'Indian, UK and US layouts compared', bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['', 'Indian', 'UK / international', 'US (block)'], [
        ['Your address', 'At the end, under your name (applications); top left in business letters', 'Top right, without your name', 'Top left, without your name'],
        ['Date', 'Date: 03-11-2026, below "To" block or above it', '3 November 2026, under your address', 'November 3, 2026, on the left'],
        ['Receiver', '"To," then designation, office, city', 'Designation or name and address, left', 'Name, title and address, left'],
        ['Subject', '"Subject:" before the salutation', 'Bold heading after the salutation', '"Subject:" or "RE:" before the salutation'],
        ['Salutation', 'Respected Sir/Madam, / Dear Sir/Madam,', 'Dear Sir or Madam, / Dear Mr Gupta,', 'Dear Sir or Madam: / Dear Mr. Gupta:'],
        ['Before closing', 'Thanking you,', 'Not used', 'Not used'],
        ['Closing', 'Yours faithfully, / Yours obediently, (students)', 'Yours faithfully, (no name) / Yours sincerely, (named)', 'Sincerely,'],
    ], [2.8, 4.8, 4.8, 4.6])
    _p(d, 'Salutation and closing: which goes with which', bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['If you start with', 'End with', 'Used for'], [
        ['Dear Sir / Dear Madam / Dear Sir/Madam', 'Yours faithfully,', 'Offices, banks, companies when you do not know the name'],
        ['Respected Sir / Respected Madam', 'Yours obediently, (students) or Yours faithfully,', 'Principal, teacher, senior officials in India'],
        ['Dear Mr. Gupta / Dear Ms. Kaur', 'Yours sincerely,', 'A person whose name you know'],
        ['Dear Rahul', 'Yours sincerely, / With warm regards,', 'Employer to employee, letters to people you know'],
    ], [5.5, 5.0, 6.5])
    _p(d, 'Getting the tone right', bold=True, size=13, color=BLUE, after=4)
    _bullets(d, ['Formal: plain and respectful. "I request you to look into the matter and take the necessary action at the earliest."',
                 'Polite and warm: adds thanks and understanding. "I would be grateful if you could look into this at your earliest convenience. Thank you for your help."',
                 'Firm: states a deadline and the next step, without rudeness. "If the matter is not resolved within 15 days, I will be compelled to take it up with the higher authorities."',
                 'Never threaten, insult or exaggerate. Facts, dates and numbers make a letter strong; adjectives do not.'])
    _p(d, 'Before you sign: 10 common mistakes', bold=True, size=13, color=BLUE, after=4)
    _bullets(d, ['No date, or two different date styles in one letter.', 'A vague subject such as "Application" or "Request".',
                 '"Dear Sir" with "Yours sincerely" (use "Yours faithfully").', 'Account, roll, invoice or consumer number missing, so the office cannot find your record.',
                 'Asking for two unrelated things in one letter.', 'Long paragraphs that hide the actual request.', 'Text speak and short forms (pls, thx, u).',
                 'Enclosures promised in the letter but not attached.', 'No signature above the name, or no phone number to contact you.',
                 'No copy kept, and no "received" stamp or signature on your copy.'])
    _p(d, after=4)
    _contents(d, GUIDE_IDS, False)
    _templates(d, GUIDE_IDS, False)
    dsave(d, 'Letter-Format-Guide-20-Templates.docx')

def letters_hindi():
    d = doc_new('letters-hindi-word'); _hi_doc(d)
    _cover(d, 'हिंदी प्रार्थना पत्र: प्रारूप और 20 नमूने', 'विद्यालय, कार्यालय, बैंक, नगर निगम और पुलिस के लिए औपचारिक पत्र, भरने के लिए तैयार। Hindi application formats.', True)
    _howto_box(d, True)
    _p(d, 'प्रार्थना पत्र के भाग', True, bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['#', 'भाग', 'क्या लिखें', 'उदाहरण'], [
        ['1', 'सेवा में', 'पत्र किसे भेजा जा रहा है: पद, कार्यालय और शहर।', 'सेवा में, प्रधानाचार्य महोदय, राजकीय विद्यालय, लुधियाना'],
        ['2', 'दिनांक', 'पत्र लिखने की तिथि।', 'दिनांक: 03-11-2026'],
        ['3', 'विषय', 'एक पंक्ति में पत्र का उद्देश्य।', 'विषय: तीन दिन के अवकाश हेतु प्रार्थना पत्र'],
        ['4', 'संबोधन', 'महोदय / महोदया / आदरणीय ... जी', 'महोदय,'],
        ['5', 'आरंभ', 'सविनय निवेदन है कि ... (आप कौन हैं)', 'सविनय निवेदन है कि मैं कक्षा आठवीं का छात्र हूँ।'],
        ['6', 'मुख्य भाग', 'तथ्य, तिथियाँ और आपका अनुरोध, छोटे अनुच्छेदों में।', 'कृपया मुझे तीन दिन का अवकाश प्रदान करें।'],
        ['7', 'समापन', 'सधन्यवाद / आपकी अति कृपा होगी।', 'सधन्यवाद,'],
        ['8', 'अभिवादन शब्द', 'आपका आज्ञाकारी शिष्य / प्रार्थी / भवदीय', 'आपकी आज्ञाकारिणी शिष्या'],
        ['9', 'नाम और विवरण', 'हस्ताक्षर, नाम, कक्षा, पता या मोबाइल नंबर।', 'रिया शर्मा, कक्षा आठवीं-ब'],
    ], [0.8, 2.6, 7.0, 6.6], True)
    _p(d, 'समापन शब्द कब लिखें', True, bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['कौन लिख रहा है', 'पुरुष / लड़का', 'महिला / लड़की'], [
        ['विद्यार्थी, प्रधानाचार्य या शिक्षक को', 'आपका आज्ञाकारी शिष्य', 'आपकी आज्ञाकारिणी शिष्या'],
        ['कोई अनुरोध या शिकायत करने वाला', 'प्रार्थी', 'प्रार्थिनी'],
        ['सामान्य औपचारिक पत्र', 'भवदीय', 'भवदीया'],
    ], [6.0, 5.2, 5.8], True)
    _p(d, after=4)
    _contents(d, HINDI_IDS, True)
    _templates(d, HINDI_IDS, True)
    dsave(d, 'Hindi-Prarthana-Patra-Formats.docx')

def letters_office():
    d = doc_new('letters-office-word')
    _cover(d, 'Office & HR Letter Formats', '13 letters for employees and 7 for employers, ready to fill in. Employer letters use a simple letterhead; replace it with yours.')
    _howto_box(d)
    _p(d, 'Good practice for office letters', bold=True, size=13, color=BLUE, after=4)
    _bullets(d, ['Address leave and work requests to your reporting manager, and HR matters (certificates, relieving letter, NOC) to the HR manager.',
                 'Put your employee ID, designation and department under your name so the letter reaches the right file.',
                 'Give dates, not "next week". Leave, resignation and joining letters are read later as records.',
                 'Employer letters: number every letter (Ref) and keep a copy signed by the employee as proof of receipt.',
                 'Notice periods, probation and leave rules in these templates are common practice; use what your appointment letter or company policy says.'])
    _p(d, after=4)
    _contents(d, OFFICE_IDS, False)
    _templates(d, OFFICE_IDS, False)
    dsave(d, 'Office-and-HR-Letter-Formats.docx')

def letters_bankgov():
    d = doc_new('letters-bank-govt-word')
    _cover(d, 'Bank, Government & Housing Letter Formats', '27 letters for banks, government offices, the police, landlords and housing societies, ready to fill in.')
    _howto_box(d)
    _p(d, 'What each office needs to see', bold=True, size=13, color=BLUE, after=4)
    _grid(d, ['Office', 'Address it to', 'Always include'], [
        ['Bank branch', 'The Branch Manager', 'Account number, account type, your registered mobile, signature as in bank records'],
        ['Municipal corporation / council', 'The Municipal Commissioner or Executive Engineer', 'Exact locality and ward, since when the problem exists, photos if possible'],
        ['Electricity / water department', 'The Assistant Engineer / SDO or Executive Engineer', 'Consumer or connection number, copy of the latest bill'],
        ['Any public authority (RTI)', 'The Public Information Officer', 'Clear numbered questions, the period, the ₹10 fee details (or BPL proof)'],
        ['Police station', 'The Station House Officer (SHO)', 'Date, time and place, description and value of the item, your ID'],
        ['Passport office / Aadhaar centre', 'The Passport Officer / Officer-in-charge', 'ARN or last 4 digits of Aadhaar, list of self-attested documents'],
        ['Landlord / housing society', 'The landlord by name / The Secretary', 'Property or flat number, dates, amounts, what the rent agreement says'],
    ], [4.2, 5.0, 7.8])
    _p(d, after=4)
    _contents(d, BANKGOV_IDS, False)
    _templates(d, BANKGOV_IDS, False)
    dsave(d, 'Bank-Government-Housing-Letter-Formats.docx')

MAKERS = [letters_guide, letters_hindi, letters_office, letters_bankgov]
