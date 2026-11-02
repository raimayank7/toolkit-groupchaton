"""Events & Weddings toolkit files (kit id `events`). python -X utf8 scripts/make_files.py --only events
   Global: a Settings sheet holds currency, number style (Indian 12,34,567 / international 1,234,567) and date style;
   conditional formats switch the number and date formats from those settings (names EvNum, EvDate, EvCur)."""
from make_files import *
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.workbook.defined_name import DefinedName

PINK = 'DB2777'
MON = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'   # Indian grouping (default)
INTL = '[>=1000000]#\\,###\\,##0;[>=1000]#\\,##0;0'   # international grouping, same on every PC locale
CURS = ['₹', '$', '€', '£', 'AED', 'SAR', 'QAR', 'KWD', 'OMR', 'BHD', 'SGD', 'MYR', 'A$', 'C$', 'NZ$', 'NPR', 'PKR', 'BDT', 'LKR', 'ZAR', 'KES', 'NGN', 'Rs']
NUMS = ['Indian (12,34,567)', 'International (1,234,567)']
DATES = [('31-12-2026', None), ('31-Dec-2026', 'dd-mmm-yyyy'), ('12/31/2026', 'mm/dd/yyyy'), ('2026-12-31', 'yyyy-mm-dd')]
YN = ['Yes', 'No']
_fid = [300]

def ev_lists(wb, extra=None):
    cols = {1: CURS, 2: NUMS, 3: [d for d, _ in DATES], 4: YN}
    cols.update(extra or {})
    ws = wb.create_sheet('Lists')
    for col, vals in cols.items():
        for i, v in enumerate(vals, 1): ws.cell(i, col, v)
    ws.sheet_state = 'hidden'
    return ws

def ev_settings(wb, rows, cur=True, extra_title=None):
    """Settings sheet: rows = [(label, value, fmt)] typed from row 3, then currency / number style / date style.
       Returns {label: absolute ref}. Defines names EvCur, EvNum, EvDate used by headings and conditional formats."""
    ws = sheet(wb, 'Settings', [30, 30, 4, 30, 14], tab=PINK)
    title(ws, 'Settings', 'Yellow cells: type or pick. Everything else in this file follows these settings.')
    refs, r = {}, 3
    std = [('Currency', '₹', None, '=Lists!$A$1:$A$%d' % len(CURS)), ('Number style', NUMS[0], None, '=Lists!$B$1:$B$2'),
           ('Date style', DATES[0][0], '@', '=Lists!$C$1:$C$4')]
    for lab, val, fmt, *src in [(a, b, c) for a, b, c in rows] + (std if cur else std[2:]):
        label(ws, 'A%d' % r, lab); cell(ws, 'B%d' % r, val, 'in', fmt, align='left')
        if src: dv(ws, 'B%d' % r, src[0])
        refs[lab] = 'Settings!$B$%d' % r; r += 1
    cell(ws, 'A%d' % (r + 1), 'Currency is shown in headings and totals; you can also type a currency not in the list. '
         'Number style changes how big amounts are grouped. Date style changes how every date in this file is shown.', size=9, color=MUTED, wrap=True)
    ws.merge_cells('A%d:B%d' % (r + 1, r + 3)); ws.row_dimensions[r + 1].height = 30
    for nm, lab in (('EvCur', 'Currency'), ('EvNum', 'Number style'), ('EvDate', 'Date style')):
        if lab in refs: wb.defined_names[nm] = DefinedName(nm, attr_text=refs[lab])
    return ws, refs

def _dxf(code):
    _fid[0] += 1
    return DifferentialStyle(numFmt=NumberFormat(numFmtId=_fid[0], formatCode=code))

def styles(ws, money=(), dates=()):
    """Money ranges get the international format when EvNum says so; date ranges follow EvDate."""
    for rng in money:
        ws.conditional_formatting.add(rng, Rule(type='expression', formula=['EvNum="%s"' % NUMS[1]], dxf=_dxf(INTL)))
    for rng in dates:
        for lab, code in DATES:
            if code: ws.conditional_formatting.add(rng, Rule(type='expression', formula=['EvDate="%s"' % lab], dxf=_dxf(code)))

def cur(text):
    """heading with the currency: '=\"Estimated (\"&EvCur&\")\"'"""
    return '="%s ("&EvCur&")"' % text

def fill(ws, r0, rows, col=1):
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v is not None: ws.cell(r0 + i, col + j, v)

def status_colors(ws, rng, first, rules):
    for word, bg, fg in rules:
        ws.conditional_formatting.add(rng, FormulaRule(formula=['%s="%s"' % (first, word)], fill=PatternFill('solid', fgColor=bg), font=Font(bold=True, color=fg)))

RED, AMBER, GREEN = ('FECACA', 'B91C1C'), ('FEF3C7', '92400E'), ('DCFCE7', '166534')
EVENT_DAY = TODAY + datetime.timedelta(days=150)
D = lambda n: TODAY + datetime.timedelta(days=n)

# ------------------------------------------------------------------ budget planner
CATS = ['Venue & hall', 'Catering & food', 'Decoration & flowers', 'Photography & video', 'Attire & jewellery (bride)', 'Attire (groom)',
        'Makeup, hair & mehndi', 'Music, DJ & band', 'Invitations & stationery', 'Gifts, favours & return gifts', 'Rituals, priest & pooja',
        'Transport & baraat', 'Guest stay & hotel', 'Entertainment & activities', 'Miscellaneous & tips', 'Contingency']

def budget_planner():
    wb = new_wb('ev-wedding-budget-planner-excel')
    howto(wb, first=True)
    st, R = ev_settings(wb, [('Event name', "Riya & Arjun's Wedding", None), ('Event date', EVENT_DAY, DATE), ('Total budget', 2500000, MON), ('Expected guests', 350, '0')])
    label(st, 'D2', 'Categories (rename freely)')
    for i in range(20): cell(st, 'D%d' % (3 + i), CATS[i] if i < len(CATS) else None, 'in')
    styles(st, ['B5'], ['B4'])
    bd = sheet(wb, 'Budget', [24, 30, 22, 13, 13, 13, 13, 13, 12, 11, 30], tab=PINK)
    sm = sheet(wb, 'Summary', [28, 14, 14, 14, 14, 14, 12, 14], tab=PINK)
    ev_lists(wb)
    ws = bd
    title(ws, 'Budget lines')
    ws['A2'] = '=Settings!B3&"  ·  budget "&EvCur&" "&TEXT(Settings!B5,"#,##0")'; ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    cell(ws, 'C3', 'Totals (visible rows)', 'soft', bold=True, align='right')
    for col in 'DEFGH': cell(ws, '%s3' % col, '=SUBTOTAL(109,%s5:%s204)' % (col, col), 'soft', MON, bold=True)
    head(ws, 4, ['Category', 'Item / details', 'Vendor', cur('Estimated'), cur('Actual'), cur('Advance paid'), cur('Expected cost'), cur('Balance due'), 'Next due date', 'Status', 'Notes'])
    f = {'G': '=IF(AND($D{r}="",$E{r}=""),"",IF($E{r}<>"",$E{r},$D{r}))', 'H': '=IF($G{r}="","",$G{r}-N($F{r}))',
         'J': '=IF($G{r}="","",IF($H{r}<=0,"Paid",IF($I{r}="",IF($E{r}="","Estimate","Booked"),IF($I{r}<TODAY(),"Overdue",IF($I{r}-TODAY()<=7,"Due soon","Booked")))))'}
    grid(ws, 5, 204, 'ABCDEFGHIJK', {'G': 'calc', 'H': 'calc', 'J': 'calc'}, {c: MON for c in 'DEFGH'} | {'I': DATE}, f)
    dv(ws, 'A5:A204', '=Settings!$D$3:$D$22', 'Pick a category (edit the list on the Settings sheet)')
    fill(ws, 5, [('Venue & hall', 'Banquet lawn, 2 days (sangeet + wedding)', 'Green Meadows Resort', 450000, 480000, 150000, None, None, D(20)),
                 ('Catering & food', 'Wedding dinner, 350 plates at 1,100', 'Shree Caterers', 385000, None, 50000, None, None, D(40)),
                 ('Catering & food', 'Sangeet snacks and dinner, 200 plates', 'Shree Caterers', 160000, None, None),
                 ('Decoration & flowers', 'Mandap, stage, entrance and lights', 'Bloom Decorators', 220000, 240000, 60000, None, None, D(5)),
                 ('Photography & video', 'Photos, video, highlights film, album', 'Pixel Stories', 180000, 175000, 50000, None, None, D(-3)),
                 ('Attire & jewellery (bride)', 'Wedding lehenga and reception saree', '', 250000, 235000, 235000),
                 ('Attire (groom)', 'Sherwani, safa and reception suit', '', 90000, None, None),
                 ('Makeup, hair & mehndi', 'Bridal makeup (3 looks) + mehndi artist', 'Glam by Neha', 75000, 72000, 20000),
                 ('Music, DJ & band', 'DJ for sangeet, dhol for baraat', 'DJ Rocky', 60000, None, None),
                 ('Invitations & stationery', '300 cards + e-invite design', 'Print House', 45000, 41000, 41000),
                 ('Gifts, favours & return gifts', 'Return gifts and welcome bags', '', 80000, None, None),
                 ('Rituals, priest & pooja', 'Pandit ji, samagri, havan items', '', 31000, None, None),
                 ('Transport & baraat', 'Horse, band, 4 cars for guests', '', 70000, None, None),
                 ('Guest stay & hotel', '20 rooms for 2 nights', 'Hotel Comfort Inn', 200000, None, 40000),
                 ('Contingency', 'Kept aside for surprises (about 10%)', '', 200000, None, None)])
    ws['K6'] = 'Final numbers to caterer 7 days before'; ws['K9'] = 'Balance on delivery of album'
    status_colors(ws, 'J5:J204', '$J5', [('Overdue',) + RED, ('Due soon',) + AMBER, ('Paid',) + GREEN])
    styles(ws, ['D3:H204'], ['I5:I204'])
    ws.freeze_panes = 'C5'; ws.auto_filter.ref = 'A4:K204'
    # Summary
    ws = sm
    title(ws, 'Budget summary')
    ws['A2'] = '=Settings!B3&"  ·  "&TEXT(Settings!B4,"dd-mm-yyyy")&"  ·  "&IF(Settings!B4>=TODAY(),Settings!B4-TODAY()&" days to go","event date passed")'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    B = "Budget!$%s$5:$%s$204"
    kpi = [('Total budget', '=Settings!B5'), ('Estimated total', '=SUM(%s)' % (B % ('D', 'D'))), ('Expected total (actual or estimate)', '=SUM(%s)' % (B % ('G', 'G'))),
           ('Paid so far', '=SUM(%s)' % (B % ('F', 'F'))), ('Balance due', '=SUM(%s)' % (B % ('H', 'H'))),
           ('Budget left (negative = over budget)', '=B4-B6'), ('Expected cost per guest', '=IF(N(Settings!B6)=0,0,B6/Settings!B6)'),
           ('Due in the next 30 days', '=SUMIFS(%s,%s,">="&TODAY(),%s,"<="&TODAY()+30)' % (B % ('H', 'H'), B % ('I', 'I'), B % ('I', 'I'))),
           ('Overdue', '=SUMIFS(%s,%s,"Overdue")' % (B % ('H', 'H'), B % ('J', 'J')))]
    for i, (k, fml) in enumerate(kpi, 4):
        cell(ws, 'A%d' % i, k, 'box', bold=True); cell(ws, 'B%d' % i, fml, 'soft' if i in (6, 8, 9) else 'calc', MON, bold=True)
        cell(ws, 'C%d' % i, '=EvCur', align='left', size=9, color=MUTED)
    ws.conditional_formatting.add('B9', CellIsRule(operator='lessThan', formula=['0'], fill=PatternFill('solid', fgColor=RED[0]), font=Font(bold=True, color=RED[1])))
    ws.conditional_formatting.add('B12', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor=RED[0]), font=Font(bold=True, color=RED[1])))
    head(ws, 14, ['Category', cur('Estimated'), cur('Actual'), cur('Expected'), cur('Paid'), cur('Balance'), '% of budget', cur('Expected − estimate')])
    for i in range(20):
        r, s = 15 + i, 3 + i
        cell(ws, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (s, s), 'box')
        for col, src in zip('BCDEF', 'DEGFH'):
            cell(ws, '%s%d' % (col, r), '=IF($A{r}="","",SUMIF(Budget!$A$5:$A$204,$A{r},Budget!${c}$5:${c}$204))'.format(r=r, c=src), 'calc', MON)
        cell(ws, 'G%d' % r, '=IF(OR($A{r}="",N(Settings!$B$5)=0),"",D{r}/Settings!$B$5)'.format(r=r), 'calc', '0.0%')
        cell(ws, 'H%d' % r, '=IF($A{r}="","",D{r}-B{r})'.format(r=r), 'calc', MON)
    cell(ws, 'A35', 'Total', 'soft', bold=True)
    for col in 'BCDEFH': cell(ws, '%s35' % col, '=SUM(%s15:%s34)' % (col, col), 'soft', MON, bold=True)
    cell(ws, 'G35', '=IF(N(Settings!$B$5)=0,"",D35/Settings!$B$5)', 'soft', '0.0%', bold=True)
    ws.conditional_formatting.add('H15:H35', CellIsRule(operator='greaterThan', formula=['0'], font=Font(bold=True, color=RED[1])))
    styles(ws, ['B4:B12', 'B15:F35', 'H15:H35'])
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Expected cost by category'; ch.legend = None
    ch.add_data(Reference(ws, min_col=4, min_row=14, max_row=14 + len(CATS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=15, max_row=14 + len(CATS)))
    ch.series[0].graphicalProperties.solidFill = PINK
    ch.height, ch.width = 11, 16; ch.y_axis.delete = False; ch.x_axis.delete = False
    ws.add_chart(ch, 'J3')
    protect(ws)
    save(wb, 'Event-Wedding-Budget-Planner.xlsx')

# ------------------------------------------------------------------ guest list
SIDES = ["Bride's side", "Groom's side", 'Both', 'Company', 'Friends', 'Other']
RSVP = ['Yes', 'No', 'Maybe', 'Pending']
MEALS = ['Veg', 'Non-veg', 'Jain', 'Vegan', 'Gluten-free', 'Kids meal']

def guest_list():
    wb = new_wb('ev-guest-list-rsvp-excel')
    howto(wb, first=True)
    st, R = ev_settings(wb, [('Event name', "Riya & Arjun's Wedding", None), ('Event date', EVENT_DAY, DATE), ('Seats per table', 10, '0'), ('Hotel for guests', 'Hotel Comfort Inn', None)], cur=False)
    label(st, 'D2', 'Functions (up to 5, rename freely)')
    for i, n in enumerate(['Mehndi', 'Sangeet', 'Haldi', 'Wedding', 'Reception']): cell(st, 'D%d' % (3 + i), n, 'in')
    styles(st, dates=['B4'])
    gs = sheet(wb, 'Guests', [24, 13, 18, 15, 22, 12, 8, 8, 8, 8, 8, 8, 9, 7, 8, 8, 11, 7, 8, 22], tab=PINK)
    sm = sheet(wb, 'Summary', [26, 12, 12, 12, 12, 12, 12, 12], tab=PINK)
    se = sheet(wb, 'Seating', [12, 10, 10, 10], tab=PINK)
    ev_lists(wb, {5: SIDES, 6: RSVP, 7: MEALS, 8: ['Y']})
    ws = gs
    title(ws, 'Guest list', 'One row per guest or family. Y under each function they are invited to.')
    head(ws, 3, ['Guest / family name', 'Side', 'Group / relation', 'Phone', 'Email', 'City'] + ['=Settings!D%d&""' % (3 + i) for i in range(5)] +
         ['Invite sent', 'RSVP', 'Adults', 'Children', 'People', 'Meal', 'Table', 'Room', 'Notes'])
    grid(ws, 4, 603, [L(c) for c in range(1, 21)], {'P': 'calc'}, None, {'P': '=IF($A{r}="","",N($N{r})+N($O{r}))'})
    for col in 'GHIJK':
        for r in range(4, 604): ws['%s%d' % (col, r)].alignment = Alignment(horizontal='center')
    dv(ws, 'B4:B603', '=Lists!$E$1:$E$6'); dv(ws, 'G4:K603', '=Lists!$H$1:$H$1'); dv(ws, 'L4:L603', '=Lists!$D$1:$D$2')
    dv(ws, 'M4:M603', '=Lists!$F$1:$F$4'); dv(ws, 'Q4:Q603', '=Lists!$G$1:$G$6')
    fill(ws, 4, [('Sharma family (Mama ji)', "Bride's side", 'Maternal uncle', '+91 98140 11111', '', 'Ludhiana', 'Y', 'Y', 'Y', 'Y', 'Y', 'Yes', 'Yes', 4, 1, None, 'Veg', 1, '101'),
                 ('Mehta family', "Bride's side", 'Neighbours', '+91 98150 22222', '', 'Ludhiana', None, 'Y', None, 'Y', 'Y', 'Yes', 'Yes', 2, 2, None, 'Veg', 2),
                 ('Kapoor family (Bua ji)', "Groom's side", 'Paternal aunt', '+91 98110 33333', '', 'Delhi', 'Y', 'Y', 'Y', 'Y', 'Y', 'Yes', 'Yes', 3, 0, None, 'Jain', 1, '102'),
                 ('Rahul & Sneha Verma', "Groom's side", 'College friends', '+91 99880 44444', 'rahul@example.com', 'Pune', None, 'Y', None, 'Y', None, 'Yes', 'Maybe', 2, 0, None, 'Non-veg', 3, '103'),
                 ('Ms. Anita Rao', 'Company', "Riya's manager", '+91 98450 55555', 'anita@example.com', 'Bengaluru', None, None, None, 'Y', 'Y', 'Yes', 'Yes', 1, 0, None, 'Veg', 3),
                 ('Singh family', "Groom's side", 'Father\'s friend', '+91 98720 66666', '', 'Chandigarh', None, None, None, 'Y', 'Y', 'Yes', 'Pending', 2, 1, None, 'Non-veg'),
                 ('James & Emma Wilson', 'Friends', "Arjun's colleagues (UK)", '+44 7700 900123', 'james@example.co.uk', 'London', None, 'Y', None, 'Y', 'Y', 'Yes', 'Yes', 2, 0, None, 'Gluten-free', 2, '104'),
                 ('Gupta family', 'Both', 'Family friends', '+91 98760 77777', '', 'Ludhiana', 'Y', 'Y', None, 'Y', None, 'No', 'Pending', 4, 2, None, 'Veg'),
                 ('Mr. & Mrs. Khanna', "Bride's side", 'Father\'s colleague', '+91 98140 88888', '', 'Ludhiana', None, None, None, 'Y', None, 'Yes', 'No', 2, 0, None, 'Veg')])
    ws.freeze_panes = 'B4'; ws.auto_filter.ref = 'A3:T603'
    status_colors(ws, 'M4:M603', '$M4', [('No',) + RED, ('Pending',) + AMBER, ('Yes',) + GREEN])
    G = lambda c: 'Guests!$%s$4:$%s$603' % (c, c)
    ws = sm
    title(ws, 'Guest summary', 'People = adults + children. Confirmed = RSVP Yes.')
    head(ws, 3, ['Function', 'Invited', 'Confirmed', 'Maybe', 'Pending', 'Declined', 'Not yet invited'])
    for i in range(5):
        r, fc = 4 + i, L(7 + i)
        cell(ws, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (3 + i, 3 + i), 'box', bold=True)
        cell(ws, 'B%d' % r, '=IF($A{r}="","",SUMIFS({p},{f},"Y"))'.format(r=r, p=G('P'), f=G(fc)), 'calc')
        for col, w in zip('CDEF', ['Yes', 'Maybe', 'Pending', 'No']):
            cell(ws, '%s%d' % (col, r), '=IF($A{r}="","",SUMIFS({p},{f},"Y",{m},"{w}"))'.format(r=r, p=G('P'), f=G(fc), m=G('M'), w=w), 'calc')
        cell(ws, 'G%d' % r, '=IF($A{r}="","",SUMIFS({p},{f},"Y",{l},"<>Yes"))'.format(r=r, p=G('P'), f=G(fc), l=G('L')), 'calc')
    head(ws, 11, ['Side', 'Rows (families)', 'People invited', 'Confirmed', 'Pending'])
    for i, s in enumerate(SIDES, 12):
        cell(ws, 'A%d' % i, s, 'box')
        cell(ws, 'B%d' % i, '=COUNTIFS(%s,A%d,%s,"<>")' % (G('B'), i, G('A')), 'calc')
        cell(ws, 'C%d' % i, '=SUMIFS(%s,%s,A%d)' % (G('P'), G('B'), i), 'calc')
        cell(ws, 'D%d' % i, '=SUMIFS(%s,%s,A%d,%s,"Yes")' % (G('P'), G('B'), i, G('M')), 'calc')
        cell(ws, 'E%d' % i, '=SUMIFS(%s,%s,A%d,%s,"Pending")+SUMIFS(%s,%s,A%d,%s,"")' % (G('P'), G('B'), i, G('M'), G('P'), G('B'), i, G('M')), 'calc')
    cell(ws, 'A18', 'Total', 'soft', bold=True)
    for col in 'BCDE': cell(ws, '%s18' % col, '=SUM(%s12:%s17)' % (col, col), 'soft', bold=True)
    head(ws, 20, ['Meal (confirmed people)', 'All'] + ['=IF(Settings!D%d="","",Settings!D%d)' % (3 + i, 3 + i) for i in range(5)])
    for i, m in enumerate(MEALS, 21):
        cell(ws, 'A%d' % i, m, 'box')
        cell(ws, 'B%d' % i, '=SUMIFS(%s,%s,A%d,%s,"Yes")' % (G('P'), G('Q'), i, G('M')), 'calc')
        for k in range(5):
            cell(ws, '%s%d' % (L(3 + k), i), '=IF(Settings!$D$%d="","",SUMIFS(%s,%s,$A%d,%s,"Yes",%s,"Y"))' % (3 + k, G('P'), G('Q'), i, G('M'), G(L(7 + k))), 'calc')
    cell(ws, 'A27', 'No meal chosen yet', 'box')
    cell(ws, 'B27', '=SUMIFS(%s,%s,"",%s,"Yes")' % (G('P'), G('Q'), G('M')), 'calc')
    head(ws, 29, ['Totals', ''])
    tots = [('Rows on the list', '=COUNTA(%s)' % G('A')), ('People on the list', '=SUM(%s)' % G('P')), ('Confirmed people', '=SUMIFS(%s,%s,"Yes")' % (G('P'), G('M'))),
            ('People still to reply (Pending or empty)', '=SUM(%s)-SUMIFS(%s,%s,"Yes")-SUMIFS(%s,%s,"No")-SUMIFS(%s,%s,"Maybe")' % (G('P'), G('P'), G('M'), G('P'), G('M'), G('P'), G('M'))),
            ('Invitations not sent yet (rows)', '=COUNTIFS(%s,"<>",%s,"<>Yes")' % (G('A'), G('L'))),
            ('Hotel rooms allotted', '=SUMPRODUCT((%s<>"")/COUNTIF(%s,%s&""))' % (G('S'), G('S'), G('S'))),
            ('Children (confirmed)', '=SUMIFS(%s,%s,"Yes")' % (G('O'), G('M')))]
    for i, (k, fml) in enumerate(tots, 30):
        cell(ws, 'A%d' % i, k, 'box'); cell(ws, 'B%d' % i, fml, 'calc', bold=True)
    protect(ws)
    ws = se
    title(ws, 'Seating', 'Seats taken = people on rows with that table number (except RSVP No).')
    head(ws, 3, ['Table', 'Seats', 'Taken', 'Free'])
    for i in range(60):
        r = 4 + i
        cell(ws, 'A%d' % r, i + 1, 'in', align='center')
        cell(ws, 'B%d' % r, '=Settings!$B$5', 'in', align='center')
        cell(ws, 'C%d' % r, '=IF($A{r}="","",SUMIFS({p},{t},$A{r},{m},"<>No"))'.format(r=r, p=G('P'), t=G('R'), m=G('M')), 'calc', align='center')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",N($B{r})-$C{r})'.format(r=r), 'calc', align='center')
    ws.conditional_formatting.add('D4:D63', CellIsRule(operator='lessThan', formula=['0'], fill=PatternFill('solid', fgColor=RED[0]), font=Font(bold=True, color=RED[1])))
    ws.freeze_panes = 'A4'
    save(wb, 'Event-Guest-List-RSVP.xlsx')

# ------------------------------------------------------------------ vendor tracker
SERVICES = ['Venue', 'Caterer', 'Decorator', 'Florist', 'Photographer', 'Videographer', 'Makeup artist', 'Mehndi artist', 'DJ / sound', 'Band / dhol',
            'Priest / officiant', 'Tent & furniture', 'Lights & generator', 'Transport', 'Hotel', 'Invitations / printing', 'Choreographer', 'Wedding planner', 'Cake / sweets', 'Other']
MODES = ['Cash', 'Bank transfer', 'UPI', 'Card', 'Cheque']

def vendor_tracker():
    wb = new_wb('ev-vendor-payment-tracker-excel')
    howto(wb, first=True)
    st, R = ev_settings(wb, [('Event name', "Riya & Arjun's Wedding", None), ('Event date', EVENT_DAY, DATE)])
    styles(st, dates=['B4'])
    vn = sheet(wb, 'Vendors', [17, 22, 16, 15, 20, 24, 13, 9, 12, 13, 13, 8, 12, 13, 13, 10, 32], tab=PINK)
    py = sheet(wb, 'Payments', [12, 24, 13, 15, 14, 30], tab=PINK)
    sm = sheet(wb, 'Summary', [34, 14, 14, 14, 8], tab=PINK)
    ct = sheet(wb, 'Contacts', [17, 24, 18, 18, 10])
    ev_lists(wb, {5: SERVICES, 6: MODES})
    ws = vn
    title(ws, 'Vendors', 'One row per vendor. Paid, balance and status calculate from the Payments sheet.')
    cell(ws, 'F3', 'Totals (visible rows)', 'soft', bold=True, align='right')
    for col in 'GJKN': cell(ws, '%s3' % col, '=SUBTOTAL(109,%s5:%s154)' % (col, col), 'soft', MON, bold=True)
    head(ws, 4, ['Service', 'Vendor name', 'Contact person', 'Phone', 'Email', 'Quotes compared', cur('Contract amount'), 'Contract signed', 'Contract date',
                 cur('Paid so far'), cur('Balance'), '% paid', 'Next due date', cur('Next amount'), 'Status', 'Arrival on the day', 'What is included / cancellation terms'])
    f = {'J': '=IF($B{r}="","",SUMIF(Payments!$B$4:$B$503,$B{r},Payments!$C$4:$C$503))', 'K': '=IF($B{r}="","",N($G{r})-$J{r})',
         'L': '=IF(OR($B{r}="",N($G{r})=0),"",$J{r}/$G{r})',
         'O': '=IF($B{r}="","",IF(N($G{r})=0,"Not booked",IF($K{r}<=0,"Paid in full",IF($M{r}="","Booked",IF($M{r}<TODAY(),"Overdue",IF($M{r}-TODAY()<=7,"Due in 7 days","Booked"))))))'}
    grid(ws, 5, 154, [L(c) for c in range(1, 18)], {c: 'calc' for c in 'JKLO'}, {'G': MON, 'J': MON, 'K': MON, 'N': MON, 'L': '0%', 'I': DATE, 'M': DATE, 'P': 'h:mm AM/PM'}, f)
    dv(ws, 'A5:A154', '=Lists!$E$1:$E$%d' % len(SERVICES)); dv(ws, 'H5:H154', '=Lists!$D$1:$D$2')
    fill(ws, 5, [('Venue', 'Green Meadows Resort', 'Mr. Bedi (manager)', '+91 98140 10101', 'events@example.com', '3 venues seen; this one has rooms', 480000, 'Yes', D(-60), None, None, None, D(20), 165000, None, datetime.time(6, 0), '2 days, lawn + hall, power backup, 20 rooms at 4,500; 50% refund if cancelled 90+ days before'),
                 ('Caterer', 'Shree Caterers', 'Ramesh ji', '+91 98150 20202', '', '4 tastings; per plate 950 to 1,250', 545000, 'Yes', D(-45), None, None, None, D(40), 200000, None, datetime.time(11, 0), '550 plates over 2 functions, live counters, waiters, crockery; final count 7 days before'),
                 ('Decorator', 'Bloom Decorators', 'Kiran', '+91 99880 30303', '', '2 quotes', 240000, 'Yes', D(-30), None, None, None, D(5), 90000, None, datetime.time(7, 0), 'Mandap, stage, entrance, fairy lights, fresh flowers; removal next morning'),
                 ('Photographer', 'Pixel Stories', 'Aman', '+91 98110 40404', 'aman@example.com', '3 portfolios compared', 175000, 'Yes', D(-50), None, None, None, D(-3), 50000, None, datetime.time(14, 0), '2 photographers, 1 video, highlights film, 1 album; data in 30 days'),
                 ('DJ / sound', 'DJ Rocky', 'Rocky', '+91 97790 50505', '', '', 60000, 'No', None, None, None, None, None, None, None, datetime.time(16, 0), 'Sangeet 5 hours, extra hour 6,000'),
                 ('Makeup artist', 'Glam by Neha', 'Neha', '+91 98720 60606', '', '2 trials done', 72000, 'Yes', D(-20), None, None, None, D(148), 52000, None, datetime.time(13, 0), '3 looks (sangeet, wedding, reception) + 2 family members'),
                 ('Band / dhol', 'Royal Band', '', '+91 98760 70707', '', '', None, 'No', None, None, None, None, None, None, None, None, 'Waiting for quote')])
    status_colors(ws, 'O5:O154', '$O5', [('Overdue',) + RED, ('Due in 7 days',) + AMBER, ('Paid in full',) + GREEN])
    ws.conditional_formatting.add('H5:H154', FormulaRule(formula=['AND($B5<>"",$H5<>"Yes")'], fill=PatternFill('solid', fgColor=AMBER[0])))
    styles(ws, ['G3:G154', 'J3:K154', 'N3:N154'], ['I5:I154', 'M5:M154'])
    ws.freeze_panes = 'C5'; ws.auto_filter.ref = 'A4:Q154'
    ws = py
    title(ws, 'Payments made', 'One row per payment. Use the vendor name exactly as on the Vendors sheet (pick from the list).')
    head(ws, 3, ['Date', 'Vendor', cur('Amount'), 'Mode', 'Receipt no.', 'Note'])
    grid(ws, 4, 503, 'ABCDEF', {}, {'A': DATE, 'C': MON})
    dv(ws, 'B4:B503', '=Vendors!$B$5:$B$154'); dv(ws, 'D4:D503', '=Lists!$F$1:$F$5')
    fill(ws, 4, [(D(-60), 'Green Meadows Resort', 150000, 'Bank transfer', 'GM-0412', 'Booking advance'), (D(-30), 'Green Meadows Resort', 165000, 'Bank transfer', 'GM-0533', 'Second instalment'),
                 (D(-45), 'Shree Caterers', 50000, 'UPI', 'SC-118', 'Advance'), (D(-30), 'Bloom Decorators', 60000, 'Cheque', 'BD-77', 'Advance, cheque 004512'),
                 (D(-50), 'Pixel Stories', 50000, 'UPI', 'PS-22', 'Booking'), (D(-20), 'Glam by Neha', 20000, 'Cash', 'GN-5', 'Advance')])
    styles(ws, ['C4:C503'], ['A4:A503'])
    ws.freeze_panes = 'A4'
    V = lambda c: 'Vendors!$%s$5:$%s$154' % (c, c)
    ws = sm
    title(ws, 'Vendor summary', 'Totals from the Vendors and Payments sheets.')
    rows = [('Total contracted', '=SUM(%s)' % V('G')), ('Paid so far', '=SUM(%s)' % V('J')), ('Balance to pay', '=SUM(%s)' % V('K')),
            ('Due in the next 30 days', '=SUMIFS(%s,%s,">="&TODAY(),%s,"<="&TODAY()+30,%s,">0")' % (V('N'), V('M'), V('M'), V('K'))),
            ('Overdue payments', '=SUMIFS(%s,%s,"Overdue")' % (V('N'), V('O')))]
    for i, (k, fml) in enumerate(rows, 3):
        cell(ws, 'A%d' % i, k, 'box', bold=True); cell(ws, 'B%d' % i, fml, 'soft' if i in (5, 7) else 'calc', MON, bold=True); cell(ws, 'C%d' % i, '=EvCur', size=9, color=MUTED)
    for i, (k, fml) in enumerate([('Vendors on the list', '=COUNTA(%s)' % V('B')), ('Booked (contract amount typed)', '=COUNTIFS(%s,"<>",%s,">0")' % (V('B'), V('G'))),
                                  ('Contract not signed yet', '=COUNTIFS(%s,"<>",%s,"<>Yes")' % (V('B'), V('H'))), ('Overdue', '=COUNTIF(%s,"Overdue")' % V('O'))], 9):
        cell(ws, 'A%d' % i, k, 'box'); cell(ws, 'B%d' % i, fml, 'calc', bold=True)
    ws.conditional_formatting.add('B7', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor=RED[0]), font=Font(bold=True, color=RED[1])))
    ws.conditional_formatting.add('B12', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor=RED[0]), font=Font(bold=True, color=RED[1])))
    head(ws, 14, ['Service', cur('Contracted'), cur('Paid'), cur('Balance')])
    for i, s in enumerate(SERVICES, 15):
        cell(ws, 'A%d' % i, s, 'box')
        for col, src in zip('BCD', 'GJK'): cell(ws, '%s%d' % (col, i), '=SUMIF(%s,$A%d,%s)' % (V('A'), i, V(src)), 'calc', MON)
    cell(ws, 'A35', 'Total', 'soft', bold=True)
    for col in 'BCD': cell(ws, '%s35' % col, '=SUM(%s15:%s34)' % (col, col), 'soft', MON, bold=True)
    styles(ws, ['B3:B7', 'B15:D35'])
    protect(ws)
    ws = ct
    title(ws, 'Day-of contacts')
    ws['A2'] = '=Settings!B3&"  ·  "&TEXT(Settings!B4,"dd-mm-yyyy")'; ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 3, ['Service', 'Vendor', 'Contact person', 'Phone', 'Arrives'])
    for i in range(150):
        r, s = 4 + i, 5 + i
        for col, src in zip('ABCD', 'ABCD'): cell(ws, '%s%d' % (col, r), '=IF(Vendors!$B%d="","",Vendors!%s%d&"")' % (s, src, s), 'box')
        cell(ws, 'E%d' % r, '=IF(OR(Vendors!$B{s}="",Vendors!$P{s}=""),"",Vendors!$P{s})'.format(s=s), 'box', 'h:mm AM/PM')
    ws.print_area = 'A1:E60'
    protect(ws)
    save(wb, 'Event-Vendor-Tracker.xlsx')

# ------------------------------------------------------------------ planning checklist
PERIODS = [('12+ months before', -365), ('9-12 months before', -300), ('6-9 months before', -210), ('4-6 months before', -150), ('2-3 months before', -75),
           ('1 month before', -30), ('2 weeks before', -14), ('Last week', -5), ('Day before', -1), ('On the day', 0), ('After the event', 14)]
TASKS = [
    ('12+ months before', 'Planning', 'Agree the total budget and who pays for which part'),
    ('12+ months before', 'Rituals', 'Fix the date; for a Hindu or Sikh wedding ask the pandit or granthi for muhurat dates first'),
    ('12+ months before', 'Guests', 'Make a rough guest count for each function to size the venue'),
    ('12+ months before', 'Venue & vendors', 'Visit venues: capacity, parking, rooms, kitchen, power backup, rain cover'),
    ('12+ months before', 'Venue & vendors', 'Book the venue with a written agreement and cancellation terms'),
    ('12+ months before', 'Planning', 'Decide the functions (roka or engagement, mehndi, sangeet, haldi, wedding, reception) and their dates'),
    ('12+ months before', 'Venue & vendors', 'Hire a wedding planner or name one family coordinator'),
    ('9-12 months before', 'Venue & vendors', 'Book the caterer after a food tasting'),
    ('9-12 months before', 'Venue & vendors', 'Book the photographer and videographer (see full albums, not only highlights)'),
    ('9-12 months before', 'Venue & vendors', 'Book the decorator and agree on a theme and colours'),
    ('9-12 months before', 'Travel & stay', 'Reserve a block of hotel rooms for guests from out of town'),
    ('9-12 months before', 'Guests', 'Start the guest list with phone numbers and addresses'),
    ('9-12 months before', 'Planning', 'Make one shared folder for contracts, receipts and lists'),
    ('6-9 months before', 'Guests', 'Send save-the-dates to guests who must travel or take leave'),
    ('6-9 months before', 'Venue & vendors', 'Book the DJ, band, dhol or live musicians'),
    ('6-9 months before', 'Attire & beauty', 'Book makeup, hair and mehndi artists after a trial'),
    ('6-9 months before', 'Attire & beauty', 'Order the bride\'s outfits; ask how long alterations take'),
    ('6-9 months before', 'Attire & beauty', 'Order jewellery and wedding rings'),
    ('6-9 months before', 'Rituals', 'Book the priest, pandit, granthi, officiant or registrar'),
    ('6-9 months before', 'Venue & vendors', 'Book a choreographer for sangeet performances (if wanted)'),
    ('6-9 months before', 'Paperwork', 'Check passports and visas for a destination wedding or the honeymoon'),
    ('4-6 months before', 'Guests', 'Design and order invitation cards and e-invites'),
    ('4-6 months before', 'Travel & stay', 'Plan transport: baraat, guest shuttles, the couple\'s car'),
    ('4-6 months before', 'Attire & beauty', 'Order the groom\'s outfits (sherwani, suit, safa or turban)'),
    ('4-6 months before', 'Travel & stay', 'Book honeymoon travel and stay'),
    ('4-6 months before', 'Guests', 'Plan return gifts, favours and welcome bags'),
    ('4-6 months before', 'Venue & vendors', 'Book tent, furniture, generator and sound if the venue does not provide them'),
    ('4-6 months before', 'Attire & beauty', 'Buy or rent outfits for close family'),
    ('2-3 months before', 'Guests', 'Send invitations: cards by hand or post, e-invites by WhatsApp or email'),
    ('2-3 months before', 'Venue & vendors', 'Finalise the menu for each function, with veg, Jain, vegan and children\'s options'),
    ('2-3 months before', 'Attire & beauty', 'Final fittings of all outfits'),
    ('2-3 months before', 'Rituals', 'Get the list of puja and ritual items (samagri) from the priest'),
    ('2-3 months before', 'Planning', 'Plan the sangeet programme and start rehearsals'),
    ('2-3 months before', 'Rituals', 'Arrange shagun envelopes, trays and items for the rituals'),
    ('2-3 months before', 'Paperwork', 'Find out the documents needed to register the marriage in your state or country'),
    ('2-3 months before', 'Attire & beauty', 'Book skin and hair treatments for 1 to 2 weeks before'),
    ('1 month before', 'Guests', 'Call or message guests who have not replied'),
    ('1 month before', 'Venue & vendors', 'Confirm every vendor in writing: date, times, what is included, balance'),
    ('1 month before', 'Guests', 'Make the seating plan and hotel room allocation'),
    ('1 month before', 'Planning', 'Write the event-day run-sheet'),
    ('1 month before', 'Planning', 'Arrange money for vendor balances and tips in labelled envelopes'),
    ('1 month before', 'Travel & stay', 'Plan airport and station pick-ups for guests'),
    ('2 weeks before', 'Venue & vendors', 'Give the final headcount to the caterer (by the contract deadline)'),
    ('2 weeks before', 'Planning', 'Share the run-sheet with vendors and coordinators'),
    ('2 weeks before', 'Travel & stay', 'Send the rooming list to the hotel'),
    ('2 weeks before', 'Attire & beauty', 'Collect outfits after alterations and try them on'),
    ('2 weeks before', 'Day-of', 'Pack the emergency kit (see the run-sheet)'),
    ('2 weeks before', 'Venue & vendors', 'Give the DJ the song list and a do-not-play list'),
    ('Last week', 'Venue & vendors', 'Walk through the venue with the decorator and venue manager'),
    ('Last week', 'Attire & beauty', 'Pack a bag for each function for the couple'),
    ('Last week', 'Planning', 'Hand vendor balance envelopes to the coordinator'),
    ('Last week', 'Venue & vendors', 'Confirm arrival times with every vendor'),
    ('Last week', 'Guests', 'Name family members to welcome guests and to receive gifts'),
    ('Last week', 'Rituals', 'Mehndi and haldi arrangements ready (if held this week)'),
    ('Day before', 'Rituals', 'Check all ritual items against the priest\'s list'),
    ('Day before', 'Day-of', 'Charge phones, power banks and camera batteries'),
    ('Day before', 'Day-of', 'Keep IDs, rings, documents and shagun envelopes in one bag with a trusted person'),
    ('Day before', 'Day-of', 'Eat well and sleep early'),
    ('On the day', 'Day-of', 'Coordinator follows the run-sheet; the couple stays off the phone'),
    ('On the day', 'Day-of', 'Receive vendors and check setup against the plan'),
    ('On the day', 'Day-of', 'One person records gifts and envelopes in the gift register'),
    ('On the day', 'Day-of', 'Pay vendor balances after the service and take receipts'),
    ('On the day', 'Day-of', 'Keep water and snacks for the couple and elders'),
    ('After the event', 'Planning', 'Count and deposit cash gifts; finish the gift register'),
    ('After the event', 'Guests', 'Send thank-you messages to guests'),
    ('After the event', 'Attire & beauty', 'Return rented outfits, jewellery and decor items'),
    ('After the event', 'Planning', 'Settle every vendor account and close the budget'),
    ('After the event', 'Paperwork', 'Register the marriage and collect the certificate'),
    ('After the event', 'Paperwork', 'Update names and addresses on IDs and bank accounts if needed'),
    ('After the event', 'Venue & vendors', 'Collect photos and video; leave reviews for good vendors'),
]
TCATS = ['Planning', 'Venue & vendors', 'Guests', 'Attire & beauty', 'Rituals', 'Paperwork', 'Travel & stay', 'Day-of']
TSTATUS = ['To do', 'In progress', 'Done', 'Not needed']

def checklist():
    assert len(TASKS) == 70, len(TASKS)
    wb = new_wb('ev-wedding-planning-checklist-excel')
    howto(wb, first=True)
    st, R = ev_settings(wb, [('Event name', "Riya & Arjun's Wedding", None), ('Event date', EVENT_DAY, DATE)], cur=False)
    styles(st, dates=['B4'])
    head(st, 2, ['Period', 'Days from event'], col=4)
    for i, (p, n) in enumerate(PERIODS, 3): cell(st, 'D%d' % i, p, 'box'); cell(st, 'E%d' % i, n, 'in', align='center')
    cell(st, 'D15', 'Days from event: negative = before the event. Change them to move every due date in a period.', size=9, color=MUTED, wrap=True)
    st.merge_cells('D15:E17')
    ck = sheet(wb, 'Checklist', [18, 58, 16, 14, 12, 12, 8, 9, 26], tab=PINK)
    pg = sheet(wb, 'Progress', [22, 9, 9, 9, 9, 10], tab=PINK)
    ev_lists(wb, {5: [p for p, _ in PERIODS], 6: TCATS, 7: TSTATUS})
    ws = ck
    title(ws, 'Planning checklist')
    ws['A2'] = '=Settings!B3&"  ·  event date "&TEXT(Settings!B4,"dd-mm-yyyy")'; ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 3, ['When', 'Task', 'Area', 'Owner', 'Due date', 'Status', 'Days left', 'Overdue', 'Notes'])
    last = 4 + len(TASKS) + 39
    grid(ws, 4, last, 'ABCDEFGHI', {'G': 'calc', 'H': 'calc'}, {'E': DATE},
         {'E': '=IF($A{r}="","",Settings!$B$4+IFERROR(VLOOKUP($A{r},Settings!$D$3:$E$13,2,FALSE),0))',
          'G': '=IF(OR($E{r}="",$F{r}="Done",$F{r}="Not needed"),"",$E{r}-TODAY())',
          'H': '=IF(AND($E{r}<>"",$F{r}<>"Done",$F{r}<>"Not needed",N($E{r})<TODAY()),"Overdue","")'})
    for i, (p, a, t) in enumerate(TASKS, 4):
        ws['A%d' % i], ws['B%d' % i], ws['C%d' % i], ws['F%d' % i] = p, t, a, 'To do'
        ws['B%d' % i].alignment = Alignment(wrap_text=True, vertical='center')
    for i in (4, 5, 6, 8): ws['F%d' % i] = 'Done'
    ws['F7'] = 'In progress'; ws['D4'], ws['D5'], ws['D7'] = 'Both families', 'Papa', 'Riya'
    dv(ws, 'A4:A%d' % last, '=Lists!$E$1:$E$%d' % len(PERIODS)); dv(ws, 'C4:C%d' % last, '=Lists!$F$1:$F$%d' % len(TCATS)); dv(ws, 'F4:F%d' % last, '=Lists!$G$1:$G$4')
    ws.conditional_formatting.add('A4:I%d' % last, FormulaRule(formula=['$H4="Overdue"'], fill=PatternFill('solid', fgColor='FEE2E2')))
    ws.conditional_formatting.add('A4:I%d' % last, FormulaRule(formula=['OR($F4="Done",$F4="Not needed")'], font=Font(color='94A3B8')))
    styles(ws, dates=['E4:E%d' % last])
    ws.freeze_panes = 'C4'; ws.auto_filter.ref = 'A3:I%d' % last
    ws = pg
    title(ws, 'Progress', 'Tasks marked Not needed are left out.')
    head(ws, 3, ['Period', 'Tasks', 'Done', 'Open', 'Overdue', '% done'])
    C = lambda c: 'Checklist!$%s$4:$%s$%d' % (c, c, last)
    for i, (p, _) in enumerate(PERIODS, 4):
        cell(ws, 'A%d' % i, p, 'box')
        cell(ws, 'B%d' % i, '=COUNTIFS(%s,$A%d,%s,"<>Not needed")' % (C('A'), i, C('F')), 'calc', align='center')
        cell(ws, 'C%d' % i, '=COUNTIFS(%s,$A%d,%s,"Done")' % (C('A'), i, C('F')), 'calc', align='center')
        cell(ws, 'D%d' % i, '=B%d-C%d' % (i, i), 'calc', align='center')
        cell(ws, 'E%d' % i, '=COUNTIFS(%s,$A%d,%s,"Overdue")' % (C('A'), i, C('H')), 'calc', align='center')
        cell(ws, 'F%d' % i, '=IF(B%d=0,"",C%d/B%d)' % (i, i, i), 'calc', '0%', align='center')
    cell(ws, 'A15', 'All tasks', 'soft', bold=True)
    for col in 'BCDE': cell(ws, '%s15' % col, '=SUM(%s4:%s14)' % (col, col), 'soft', bold=True, align='center')
    cell(ws, 'F15', '=IF(B15=0,"",C15/B15)', 'soft', '0%', bold=True, align='center')
    ws.conditional_formatting.add('E4:E15', CellIsRule(operator='greaterThan', formula=['0'], font=Font(bold=True, color=RED[1])))
    cell(ws, 'A17', '="Days to go: "&MAX(0,Settings!B4-TODAY())', bold=True, color=PINK, size=13)
    protect(ws)
    save(wb, 'Event-Wedding-Planning-Checklist.xlsx')

# ------------------------------------------------------------------ gift / shagun register
GTYPES = ['Cash / shagun envelope', 'Bank or UPI transfer', 'Gold / jewellery', 'Gift item', 'Gift card / voucher', 'Other']

def gift_register():
    wb = new_wb('ev-gift-shagun-register-excel')
    howto(wb, first=True)
    st, R = ev_settings(wb, [('Event name', "Riya & Arjun's Wedding", None), ('Event date', EVENT_DAY, DATE)])
    label(st, 'D2', 'Functions (rename freely)')
    for i in range(8): cell(st, 'D%d' % (3 + i), (['Roka', 'Mehndi', 'Sangeet', 'Wedding', 'Reception'] + [None] * 3)[i], 'in')
    styles(st, dates=['B4'])
    rg = sheet(wb, 'Register', [12, 13, 26, 13, 20, 12, 15, 20, 26, 12, 12, 12, 10, 20, 6], tab=PINK)
    sm = sheet(wb, 'Summary', [28, 10, 14, 14, 14], tab=PINK)
    lk = sheet(wb, 'Look-up', [12, 13, 26, 20, 20, 26, 12, 12], tab=PINK)
    ev_lists(wb, {5: SIDES, 6: GTYPES})
    ws = rg
    title(ws, 'Gift and shagun register', 'One row per envelope or gift. Write the name as on the envelope or card.')
    last = 803
    head(ws, 3, ['Date', 'Function', 'Guest / family name', 'Side', 'Relation', 'City', 'Phone', 'Gift type', 'Description', cur('Cash'), cur('Gift value (est.)'), cur('Total'), 'Thank-you sent', 'Notes', 'Look-up #'])
    grid(ws, 4, last, 'ABCDEFGHIJKLMNO', {'L': 'calc', 'O': 'calc'}, {'A': DATE, 'J': MON, 'K': MON, 'L': MON},
         {'L': '=IF(AND($J{r}="",$K{r}=""),"",N($J{r})+N($K{r}))',
          'O': '=IF(OR($C{r}="",TRIM(\'Look-up\'!$B$3)=""),"",IF(ISNUMBER(SEARCH(TRIM(\'Look-up\'!$B$3),$C{r})),MAX($O$3:$O{{p}})+1,""))'})
    for r in range(4, last + 1): ws['O%d' % r].value = ws['O%d' % r].value.replace('{p}', str(r - 1))
    dv(ws, 'B4:B%d' % last, '=Settings!$D$3:$D$10'); dv(ws, 'D4:D%d' % last, '=Lists!$E$1:$E$6'); dv(ws, 'H4:H%d' % last, '=Lists!$F$1:$F$6'); dv(ws, 'M4:M%d' % last, '=Lists!$D$1:$D$2')
    fill(ws, 4, [(EVENT_DAY - datetime.timedelta(days=2), 'Sangeet', 'Sharma family (Mama ji)', "Bride's side", 'Maternal uncle', 'Ludhiana', '98140 11111', 'Cash / shagun envelope', 'Envelope', 21000, None, None, 'Yes'),
                 (EVENT_DAY - datetime.timedelta(days=2), 'Sangeet', 'Mehta family', "Bride's side", 'Neighbours', 'Ludhiana', '', 'Gift item', 'Dinner set (48 pieces)', None, 6500, None, 'Yes'),
                 (EVENT_DAY, 'Wedding', 'Kapoor family (Bua ji)', "Groom's side", 'Paternal aunt', 'Delhi', '98110 33333', 'Gold / jewellery', 'Gold earrings', None, 45000, None, 'No'),
                 (EVENT_DAY, 'Wedding', 'Ms. Anita Rao', 'Company', "Riya's manager", 'Bengaluru', '', 'Gift card / voucher', 'Shopping voucher', None, 5000, None, 'No'),
                 (EVENT_DAY, 'Wedding', 'Singh family', "Groom's side", "Father's friend", 'Chandigarh', '', 'Cash / shagun envelope', 'Envelope', 5100, None, None, 'No'),
                 (EVENT_DAY, 'Wedding', 'Rahul & Sneha Verma', "Groom's side", 'College friends', 'Pune', '', 'Bank or UPI transfer', 'UPI on the day', 11000, None, None, 'No'),
                 (EVENT_DAY + datetime.timedelta(days=1), 'Reception', 'James & Emma Wilson', 'Friends', "Arjun's colleagues", 'London', '', 'Gift item', 'Tea set and card', None, 4000, None, 'No'),
                 (EVENT_DAY + datetime.timedelta(days=1), 'Reception', 'Sharma family (Mama ji)', "Bride's side", 'Maternal uncle', 'Ludhiana', '', 'Cash / shagun envelope', 'Envelope for the couple', 11000, None, None, 'No')])
    styles(ws, ['J4:L%d' % last], ['A4:A%d' % last])
    ws.conditional_formatting.add('M4:M%d' % last, FormulaRule(formula=['AND($C4<>"",$M4<>"Yes")'], fill=PatternFill('solid', fgColor=AMBER[0])))
    ws.column_dimensions['O'].hidden = True
    ws.freeze_panes = 'D4'; ws.auto_filter.ref = 'A3:N%d' % last
    G = lambda c: 'Register!$%s$4:$%s$%d' % (c, c, last)
    ws = sm
    title(ws, 'Gift summary', 'Totals from the Register sheet.')
    for i, (k, fml, fm) in enumerate([('Entries', '=COUNTA(%s)' % G('C'), '0'), ('Cash total', '=SUM(%s)' % G('J'), MON), ('Gift value (estimated)', '=SUM(%s)' % G('K'), MON),
                                      ('Cash + gifts', '=SUM(%s)' % G('L'), MON), ('Thank-yous still to send', '=COUNTIFS(%s,"<>",%s,"<>Yes")' % (G('C'), G('M')), '0')], 3):
        cell(ws, 'A%d' % i, k, 'box', bold=True); cell(ws, 'B%d' % i, fml, 'soft', fm, bold=True); ws.merge_cells('B%d:C%d' % (i, i))
    def block(r0, heading, labels, crit):
        head(ws, r0, [heading, 'Entries', cur('Cash'), cur('Gift value'), cur('Total')])
        for i, lab in enumerate(labels, r0 + 1):
            cell(ws, 'A%d' % i, lab, 'box')
            cell(ws, 'B%d' % i, '=IF($A{r}="","",COUNTIFS({c},$A{r},{n},"<>"))'.format(r=i, c=G(crit), n=G('C')), 'calc', align='center')
            for col, src in zip('CDE', 'JKL'): cell(ws, '%s%d' % (col, i), '=IF($A{r}="","",SUMIF({c},$A{r},{s}))'.format(r=i, c=G(crit), s=G(src)), 'calc', MON)
        return r0 + len(labels) + 2
    r = block(10, 'By function', ['=IF(Settings!D%d="","",Settings!D%d)' % (3 + i, 3 + i) for i in range(8)], 'B')
    r = block(r, 'By side', SIDES, 'D')
    r = block(r, 'By gift type', GTYPES, 'H')
    styles(ws, ['B4:C6', 'C10:E%d' % r])
    protect(ws)
    ws = lk
    title(ws, 'Look up a guest', 'Type part of a name (for example "Sharma") to see everything they gave.')
    label(ws, 'A3', 'Name'); ws.merge_cells('B3:D3'); cell(ws, 'B3', 'Sharma', 'in')
    label(ws, 'E3', 'Total from matches'); cell(ws, 'F3', '=SUMIFS(%s,%s,">0")' % (G('L'), G('O')), 'soft', MON, bold=True); cell(ws, 'G3', '=EvCur', size=9, color=MUTED)
    head(ws, 5, ['Date', 'Function', 'Guest / family name', 'Relation', 'Gift type', 'Description', cur('Cash'), cur('Gift value')])
    for i in range(1, 61):
        r = 5 + i
        for col, src in zip('ABCDEFGH', 'ABCEHIJK'):
            cell(ws, '%s%d' % (col, r), '=IFERROR(INDEX(Register!$%s$1:$%s$%d,MATCH(%d,Register!$O$1:$O$%d,0))&"","")' % (src, src, last, i, last) if col not in 'AGH' else
                 '=IFERROR(N(INDEX(Register!$%s$1:$%s$%d,MATCH(%d,Register!$O$1:$O$%d,0))),"")' % (src, src, last, i, last), 'box',
                 DATE if col == 'A' else (MON if col in 'GH' else None))
    ws.conditional_formatting.add('A6:A65', CellIsRule(operator='equal', formula=['0'], font=Font(color='FFFFFF')))
    ws.conditional_formatting.add('G6:H65', CellIsRule(operator='equal', formula=['0'], font=Font(color='FFFFFF')))
    styles(ws, ['F3', 'G6:H65'], ['A6:A65'])
    protect(ws)
    save(wb, 'Event-Gift-Shagun-Register.xlsx')

# ------------------------------------------------------------------ run-sheet (Word)
def run_sheet():
    d = doc_new('ev-event-day-run-sheet-word')
    para(d, 'EVENT DAY RUN-SHEET', bold=True, size=20, color=PINK, space=0)
    para(d, '[Event name] · [Day, DD Month YYYY]', bold=True, size=13, space=8)
    kv(d, [('Venue', '[Name]'), ('Address / map link', '[ ]'), ('Event coordinator', '[Name]'), ('Coordinator phone', '[ ]'),
           ('Backup coordinator', '[Name, phone]'), ('Venue manager', '[Name, phone]'), ('Guests expected', '[ ]'), ('Ceremony / muhurat time', '[ ]')])
    para(d, 'Key contacts', bold=True, size=13, color=PINK)
    table(d, ['Role', 'Name', 'Phone', 'On site from'],
          [['Bride\'s family point person', '[ ]', '[ ]', '[ ]'], ['Groom\'s family point person', '[ ]', '[ ]', '[ ]'], ['Caterer', '[ ]', '[ ]', '[ ]'],
           ['Decorator', '[ ]', '[ ]', '[ ]'], ['Photographer / video', '[ ]', '[ ]', '[ ]'], ['DJ / sound', '[ ]', '[ ]', '[ ]'], ['Makeup artist', '[ ]', '[ ]', '[ ]'],
           ['Priest / officiant', '[ ]', '[ ]', '[ ]'], ['Transport', '[ ]', '[ ]', '[ ]'], ['Hotel desk', '[ ]', '[ ]', '[ ]']], [5, 4.5, 4, 3.5], fill_rows=2)
    para(d, 'Vendor arrival and setup', bold=True, size=13, color=PINK)
    table(d, ['Time', 'Vendor', 'Setup task', 'Received by'],
          [['06:00', 'Decorator', 'Mandap, stage, entrance, lights', '[ ]'], ['09:00', 'Tent / sound / generator', 'Power backup tested, mics checked', '[ ]'],
           ['11:00', 'Caterer', 'Kitchen setup, counters, water station', '[ ]'], ['13:00', 'Makeup artist', 'Bridal room ready, lighting, mirror', '[ ]'],
           ['14:00', 'Photographer', 'Getting-ready shots; shot list from the couple', '[ ]'], ['16:00', 'DJ', 'Sound check, playlist and do-not-play list', '[ ]']], [2, 4, 7.5, 3.5], fill_rows=2)
    para(d, 'Timeline (sample wedding day: replace with your own)', bold=True, size=13, color=PINK)
    table(d, ['Time', 'Activity', 'Where', 'Responsible', 'Notes'],
          [['06:00', 'Venue opens; decoration setup starts', 'Hall and lawn', 'Venue manager', ''],
           ['10:00', 'Haldi (if on the same day)', 'Home / lawn', 'Bride\'s family', 'Old clothes, towels ready'],
           ['13:00', 'Lunch for family and vendors', 'Dining area', 'Caterer', 'Count vendor meals'],
           ['14:00', 'Bride\'s makeup and hair begins', 'Bridal room', 'Makeup artist', '3 to 4 hours'],
           ['17:30', 'Groom gets ready; safa / turban tying', 'Groom\'s room', 'Groom\'s family', ''],
           ['18:30', 'Baraat gathers; dhol starts', '[Meeting point]', 'Groom\'s point person', 'Horse / car ready'],
           ['19:00', 'Baraat procession to the venue', 'Road to gate', 'Transport + band', 'Keep 15 minutes buffer'],
           ['19:45', 'Milni and welcome of the groom\'s family', 'Main gate', 'Bride\'s family', 'Garlands, shagun envelopes'],
           ['20:00', 'Jaimala / varmala on stage', 'Stage', 'Coordinator', 'Photographer in position'],
           ['20:30', 'Dinner opens for guests', 'Dining area', 'Caterer', 'Elders served first'],
           ['21:30', 'Family photos on stage', 'Stage', 'Photographer', 'Call families in order of a list'],
           ['23:00', 'Wedding ceremony / pheras (muhurat)', 'Mandap', 'Priest', 'Ritual items checked'],
           ['01:00', 'Ceremony ends; vidaai or send-off', 'Gate', 'Both families', 'Car decorated and fuelled']], [1.8, 5.2, 3, 3.3, 3.7], fill_rows=2)
    para(d, 'Alternative sample: ceremony and reception', bold=True, color=PINK)
    table(d, ['Time', 'Activity', 'Responsible'],
          [['15:30', 'Guests arrive; ushers seat them', 'Ushers'], ['16:00', 'Ceremony', 'Officiant'], ['16:45', 'Group photos', 'Photographer'],
           ['17:00', 'Drinks and canapés', 'Caterer'], ['18:30', 'Couple\'s entrance; dinner served', 'MC / DJ'], ['19:45', 'Speeches and toasts', 'MC'],
           ['20:15', 'First dance, cake cutting', 'DJ'], ['20:30', 'Dancing', 'DJ'], ['23:00', 'Send-off; venue pack-down', 'Coordinator']], [2, 10, 5])
    para(d, 'Emergency kit', bold=True, size=13, color=PINK)
    for s in ['Safety pins, needle and thread in outfit colours, fabric tape', 'Pain relief, antacid, plasters, ORS, any regular medicines of the couple and elders',
              'Phone chargers and power banks, extension cord', 'Tissues, wet wipes, deodorant, comb, hair pins, spare dupatta or tie',
              'Copies of vendor contracts, run-sheet and the guest list', 'Cash for tips and small purchases; balance envelopes for vendors (with the coordinator)',
              'Umbrellas and torches; water bottles and dry snacks']:
        d.add_paragraph(s, style='List Bullet')
    para(d, 'Backup plan', bold=True, size=13, color=PINK)
    table(d, ['If this happens', 'We will', 'Who decides'],
          [['Rain or storm', 'Move to [covered area / hall]; decorator has a rain cover', '[ ]'], ['Power cut', 'Generator starts within [ ] minutes; check fuel in the morning', '[ ]'],
           ['A vendor is late or does not come', 'Call the backup contact: [name, phone]', '[ ]'], ['Baraat or ceremony runs late', 'Hold dinner / adjust music; tell the caterer and priest', '[ ]'],
           ['Someone feels unwell', 'First aid kit at [place]; nearest hospital: [name, address, phone]', '[ ]'], ['Ring, documents or envelopes missing', 'Check with the person holding the day bag: [name]', '[ ]']], [5, 9, 3])
    para(d, 'After the event', bold=True, size=13, color=PINK)
    for s in ['Collect gifts, envelopes and the gift register; two people count cash', 'Pay remaining vendor balances and take receipts',
              'Check the venue for left items; return rented decor and furniture', 'Hand over the run-sheet notes to whoever settles accounts']:
        d.add_paragraph(s, style='List Bullet')
    para(d, 'Run-sheet version [1] · printed [DD-MM-YYYY] · share the PDF on the family group too.', size=9, italic=True, color='55607A')
    dsave(d, 'Event-Day-Run-Sheet.docx')

MAKERS = [budget_planner, guest_list, vendor_tracker, checklist, gift_register, run_sheet]
