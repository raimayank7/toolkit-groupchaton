"""Landlords & Rentals toolkit files (kit id: rentals). Run: python -X utf8 scripts/make_files.py --only rentals
   Global kit: money cells use the currency symbol typed on Settings, dates use the computer's own short date format."""
from make_files import *

DT = 'mm-dd-yy'        # Excel built-in format 14 = the computer's regional short date (05-04-2026, 4/5/2026, 2026-04-05 …)
MON = 'mmm yyyy'
MONEY = '#,##0.00'
MONEY0 = '#,##0;-#,##0;"–"'
RED = PatternFill('solid', fgColor='FECACA')
GREEN = PatternFill('solid', fgColor='BBF7D0')
ORANGE = PatternFill('solid', fgColor='FED7AA')
FY = datetime.date(TODAY.year if TODAY.month >= 4 else TODAY.year - 1, 4, 1)   # sample ledger year = Indian financial year
PROPS = ['Green View Apartments', 'Market Road Shops', 'Lake Side House']
MODES = ['Cash', 'UPI', 'Bank transfer', 'Cheque', 'Card', 'Other']

def mon(d, n=0):
    m = d.month - 1 + n
    return datetime.date(d.year + m // 12, m % 12 + 1, 1)

def cur(sym):   # header text with the currency symbol from Settings
    return lambda text: '="%s ("&%s&")"' % (text, sym)

def words_any(ref, major, minor, style):
    """Amount in words: Indian (crore, lakh) or International (billion, million) style, any currency name; 0-99 words in Lists!A1:A100."""
    W = 'Lists!$A$1:$A$100'
    n = 'INT(ROUND(%s,2))' % ref
    def g3(e): return 'IF(INT((%s)/100)>0,INDEX(%s,INT((%s)/100)+1)&" Hundred ","")&IF(MOD(%s,100)>0,INDEX(%s,MOD(%s,100)+1)&" ","")' % (e, W, e, e, W, e)
    def grp(e, word): return 'IF((%s)>0,%s&"%s ","")' % (e, g3(e), word)
    ind = '&'.join([grp('INT(%s/10000000)' % n, 'Crore'), grp('INT(MOD(%s,10000000)/100000)' % n, 'Lakh'),
                    grp('INT(MOD(%s,100000)/1000)' % n, 'Thousand'), g3('MOD(%s,1000)' % n)])
    intl = '&'.join([grp('INT(%s/1000000000)' % n, 'Billion'), grp('INT(MOD(%s,1000000000)/1000000)' % n, 'Million'),
                     grp('INT(MOD(%s,1000000)/1000)' % n, 'Thousand'), g3('MOD(%s,1000)' % n)])
    pa = 'ROUND(MOD(ROUND(%s,2),1)*100,0)' % ref
    return ('=IF(N(%s)<=0,"",%s&" "&IF(%s=0,"Zero",TRIM(IF(%s="International",%s,%s)))&IF(%s>0," and "&INDEX(%s,%s+1)&" "&%s,"")&" Only")'
            % (ref, major, n, style, intl, ind, pa, W, pa, minor))

def put(ws, r0, rows, col=1):
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v is not None: ws.cell(r0 + i, col + j, v)

# ------------------------------------------------------------------ rent ledger
def rent_ledger():
    wb = new_wb('rent-ledger-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 24, 4, 30, 4, 18])
    un = sheet(wb, 'Units', [11, 24, 13, 24, 15, 12, 12, 9, 12, 12, 13, 26], tab=BLUE)
    pm = sheet(wb, 'Payments', [12, 11, 12, 12, 14, 18, 24, 24], tab=BLUE)
    lg = sheet(wb, 'Ledger', [11, 22, 22, 11, 8, 13, 13, 13, 10, 12, 12, 8] + [9] * 12 + [70, 11, 11], tab=BLUE)
    sm = sheet(wb, 'Summary', [26, 8, 10, 14, 14, 14, 14, 14, 16], tab=BLUE)
    C = cur('Settings!$B$4')
    # Settings
    title(st, 'Settings', 'Fill once. Everything else uses these.')
    for i, (k, v, fmt) in enumerate([('Your name (for reminders)', 'Anil Kapoor', None), ('Currency symbol', '₹', None),
                                     ('Ledger year starts (month)', FY, MON), ('Position as on', '=TODAY()', DT)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    cell(st, 'A7', 'Leave "Position as on" as =TODAY() for today\'s position, or type a date to see the position on that day.', size=9, color=MUTED)
    label(st, 'A9', 'Ledger runs from'); cell(st, 'B9', '=DATE(YEAR(B5),MONTH(B5),1)', 'calc', MON, align='left')
    label(st, 'A10', 'to'); cell(st, 'B10', '=EDATE(B9,11)', 'calc', MON, align='left')
    label(st, 'D2', 'Your properties (edit freely)')
    for i in range(3, 23): cell(st, 'D%d' % i, PROPS[i - 3] if i - 3 < len(PROPS) else None, 'in')
    label(st, 'F2', 'Payment modes')
    for i in range(3, 11): cell(st, 'F%d' % i, MODES[i - 3] if i - 3 < len(MODES) else None, 'in')
    protect(st)
    # Units
    title(un, 'Units and tenants', 'One row per flat, shop or room. Leave the tenant empty when a unit is vacant.')
    head(un, 3, ['Unit code', 'Property', 'Flat / shop no.', 'Tenant name', 'Phone', C('Monthly rent'), C('Deposit held'), 'Rent due day',
                 'Lease start', 'Lease end', C('Arrears brought forward'), 'Notes'])
    grid(un, 4, 203, 'ABCDEFGHIJKL', {}, {'F': MONEY, 'G': MONEY, 'I': DT, 'J': DT, 'K': MONEY})
    dv(un, 'B4:B203', '=Settings!$D$3:$D$22', 'Pick a property (add names on the Settings sheet)')
    D = datetime.timedelta
    units = [('GV-101', PROPS[0], 'Flat 101', 'Rohit Malhotra', '98140 12345', 18000, 36000, 5, mon(FY, -3), mon(FY, 8) - D(days=1), None, ''),
             ('GV-102', PROPS[0], 'Flat 102', 'Priya Nair', '98150 23456', 16500, 33000, 5, mon(FY, 2), mon(FY, 13) - D(days=1), None, ''),
             ('GV-103', PROPS[0], 'Flat 103', None, None, 17000, None, 5, None, None, None, 'Vacant: painting done'),
             ('MR-S1', PROPS[1], 'Shop 1', 'Sandhu Electricals', '98720 34567', 25000, 75000, 10, mon(FY, -12), mon(FY, 24) - D(days=1), None, 'Commercial'),
             ('MR-S2', PROPS[1], 'Shop 2', 'Kaur Bakery', '98880 45678', 22000, 66000, 10, mon(FY, -6), mon(FY, 18) - D(days=1), 22000, 'One month old dues'),
             ('LS-01', PROPS[2], 'Whole house', 'David & Mary Thomas', '99880 56789', 40000, 80000, 1, mon(FY, 0), mon(FY, 11) + D(days=30), None, '')]
    put(un, 4, units)
    un.freeze_panes = 'B4'
    # Payments (samples: monthly rent up to this month, with a few gaps)
    title(pm, 'Rent received', 'One row per payment. "Rent for month" = any date in the month the rent is for. Tenant and property fill in.')
    head(pm, 3, ['Date paid', 'Unit code', 'Rent for month', C('Amount'), 'Mode', 'Reference', 'Tenant', 'Property'])
    look = lambda col: '=IF($B{r}="","",IFERROR(INDEX(Units!$%s:$%s,MATCH($B{r},Units!$A:$A,0))&"","Unit code not found"))' % (col, col)
    grid(pm, 4, 1003, 'ABCDEFGH', {'G': 'calc', 'H': 'calc'}, {'A': DT, 'C': MON, 'D': MONEY}, {'G': look('D'), 'H': look('B')})
    dv(pm, 'B4:B1003', '=Units!$A$4:$A$203', 'Pick the unit code'); dv(pm, 'E4:E1003', '=Settings!$F$3:$F$10')
    pm.conditional_formatting.add('C4:C1003', FormulaRule(formula=['AND($D4<>"",$C4="")'], fill=RED))
    rows, last = [], mon(TODAY)
    skip = {('GV-102', mon(TODAY, -1)), ('MR-S2', mon(TODAY, -2))}
    for u in units:
        if not u[3]: continue
        m, upto = max(FY, mon(u[8])), (last if TODAY.day >= u[7] else mon(last, -1))
        while m <= upto and m <= mon(FY, 11):
            if (u[0], m) not in skip and not (u[0] == 'GV-102' and m == upto):
                paid = m + D(days=u[7] - 1 + (2 if u[0] == 'MR-S1' else 0))
                amt = u[5] // 2 if (u[0] == 'MR-S2' and m == upto) else u[5]
                mode = 'Cheque' if u[0].startswith('MR') else ('Cash' if u[0] == 'GV-101' else 'Bank transfer')
                rows.append((min(paid, TODAY), u[0], m, amt, mode, ('CHQ %06d' % (412000 + len(rows))) if mode == 'Cheque' else (None if mode == 'Cash' else 'UTR …%04d' % (3100 + len(rows)))))
            m = mon(m, 1)
    rows.append((TODAY, 'LS-01', mon(TODAY, 1), 40000, 'Bank transfer', 'Advance for next month'))
    rows.sort(key=lambda x: x[0])
    put(pm, 4, rows)
    pm.freeze_panes = 'A4'
    # Ledger
    ws = lg
    title(ws, 'Rent ledger')
    ws['A2'] = '="As on "&TEXT(Settings!B6,"dd-mmm-yyyy")&"   ·   ledger year "&TEXT(Settings!B9,"mmm yyyy")&" to "&TEXT(Settings!B10,"mmm yyyy")'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 4, ['Unit', 'Property', 'Tenant', C('Rent'), 'Months due', C('Rent due'), C('Paid'), C('Balance'), 'Status', C('Deposit'),
                 'Lease ends', 'Days left'] + [None] * 12 + ['Reminder message (copy into WhatsApp, SMS or email)', 'Due from', 'Due up to'])
    for k in range(12):
        c = ws.cell(4, 13 + k, '=Settings!$B$9' if k == 0 else '=EDATE(Settings!$B$9,%d)' % k); c.number_format = 'mmm yy'
    AS = 'Settings!$B$6'
    M0 = 'DATE(YEAR(%s),MONTH(%s),1)' % (AS, AS)
    for i in range(200):
        r, s = 5 + i, 4 + i
        f = lambda t: t.format(r=r, s=s, AS=AS, M0=M0)
        cell(ws, 'A%d' % r, f('=IF(Units!A{s}="","",Units!A{s})'), 'box')
        cell(ws, 'B%d' % r, f('=IF($A{r}="","",Units!B{s}&"")'), 'box')
        cell(ws, 'C%d' % r, f('=IF($A{r}="","",IF(Units!D{s}="","(vacant)",Units!D{s}&""))'), 'box')
        cell(ws, 'D%d' % r, f('=IF($A{r}="","",N(Units!F{s}))'), 'calc', MONEY)
        cell(ws, 'Z%d' % r, f('=IF($A{r}="","",IF(Units!$I{s}="",Settings!$B$9,MAX(Settings!$B$9,DATE(YEAR(Units!$I{s}),MONTH(Units!$I{s}),1))))'), 'calc', MON)
        cell(ws, 'AA%d' % r, f('=IF($A{r}="","",MIN(IF(DAY({AS})>=MAX(1,N(Units!$H{s})),{M0},EDATE({M0},-1)),Settings!$B$10,'
                               'IF(Units!$J{s}="",2958465,DATE(YEAR(Units!$J{s}),MONTH(Units!$J{s}),1))))'), 'calc', MON)
        cell(ws, 'E%d' % r, f('=IF($A{r}="","",IF(OR(Units!$D{s}="",$AA{r}<$Z{r}),0,(YEAR($AA{r})-YEAR($Z{r}))*12+MONTH($AA{r})-MONTH($Z{r})+1))'), 'calc', align='center')
        cell(ws, 'F%d' % r, f('=IF($A{r}="","",N(Units!K{s})+E{r}*D{r})'), 'calc', MONEY)
        cell(ws, 'G%d' % r, f('=IF($A{r}="","",SUMIFS(Payments!$D:$D,Payments!$B:$B,$A{r},Payments!$C:$C,"<"&EDATE(Settings!$B$10,1)))'), 'calc', MONEY)
        cell(ws, 'H%d' % r, f('=IF($A{r}="","",F{r}-G{r})'), 'calc', MONEY, bold=True)
        cell(ws, 'I%d' % r, f('=IF($A{r}="","",IF(H{r}>0.005,"Arrears",IF(H{r}<-0.005,"Advance",IF(Units!D{s}="","Vacant","Paid up"))))'), 'calc', align='center')
        cell(ws, 'J%d' % r, f('=IF($A{r}="","",N(Units!G{s}))'), 'calc', MONEY)
        cell(ws, 'K%d' % r, f('=IF(OR($A{r}="",Units!J{s}=""),"",Units!J{s})'), 'calc', DT)
        cell(ws, 'L%d' % r, f('=IF(K{r}="","",K{r}-{AS})'), 'calc', '0', align='center')
        for k in range(12):
            col = L(13 + k)
            cell(ws, '%s%d' % (col, r), f('=IF($A{r}="","",SUMIFS(Payments!$D:$D,Payments!$B:$B,$A{r},Payments!$C:$C,">="&%s$4,Payments!$C:$C,"<"&EDATE(%s$4,1)))' % (col, col)), 'box', MONEY0, align='right')
        cell(ws, 'Y%d' % r, f('=IF(OR($A{r}="",N(H{r})<=0.005,Units!D{s}=""),"","Dear "&C{r}&", rent of "&Settings!$B$4&" "&TEXT(H{r},"#,##0")&" is pending for "&Units!C{s}&", "&B{r}&" as on "&TEXT({AS},"dd-mmm-yyyy")&". Kindly pay at the earliest. Thank you. "&Settings!$B$3)'), size=9)
    ws.conditional_formatting.add('M5:X204', FormulaRule(formula=['AND($A5<>"",$C5<>"(vacant)",M$4>=$Z5,M$4<=$AA5,N(M5)<$D5)'], fill=RED))
    ws.conditional_formatting.add('M5:X204', FormulaRule(formula=['AND($A5<>"",$D5>0,N(M5)>=$D5)'], fill=GREEN))
    ws.conditional_formatting.add('I5:I204', CellIsRule(operator='equal', formula=['"Arrears"'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('I5:I204', CellIsRule(operator='equal', formula=['"Advance"'], font=Font(bold=True, color='1D4ED8')))
    ws.conditional_formatting.add('L5:L204', FormulaRule(formula=['AND(L5<>"",L5<0)'], fill=RED))
    ws.conditional_formatting.add('L5:L204', FormulaRule(formula=['AND(L5<>"",L5>=0,L5<=60)'], fill=ORANGE))
    cell(ws, 'A206', 'Red month = rent due but not fully paid · green = paid in full · orange "Days left" = lease ends within 60 days, red = lease has ended.', size=9, color=MUTED)
    ws.freeze_panes = 'D5'
    protect(ws)
    # Summary
    ws = sm
    title(ws, 'Summary by property')
    ws['A2'] = '="As on "&TEXT(Settings!B6,"dd-mmm-yyyy")&"   ·   amounts in "&Settings!B4'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 3, ['Property', 'Units', 'Occupied', C('Monthly rent roll'), C('Rent due so far'), C('Collected'), C('Arrears'), C('Deposits held'), C('Received this month')])
    for i in range(20):
        r, s = 4 + i, 3 + i
        f = lambda t: t.format(r=r, s=s)
        cell(ws, 'A%d' % r, f('=IF(Settings!D{s}="","",Settings!D{s})'), 'box')
        cell(ws, 'B%d' % r, f('=IF($A{r}="","",COUNTIF(Units!$B$4:$B$203,$A{r}))'), 'calc', align='center')
        cell(ws, 'C%d' % r, f('=IF($A{r}="","",COUNTIFS(Units!$B$4:$B$203,$A{r},Units!$D$4:$D$203,"<>"))'), 'calc', align='center')
        cell(ws, 'D%d' % r, f('=IF($A{r}="","",SUMIFS(Units!$F$4:$F$203,Units!$B$4:$B$203,$A{r},Units!$D$4:$D$203,"<>"))'), 'calc', MONEY)
        cell(ws, 'E%d' % r, f('=IF($A{r}="","",SUMIFS(Ledger!$F$5:$F$204,Ledger!$B$5:$B$204,$A{r}))'), 'calc', MONEY)
        cell(ws, 'F%d' % r, f('=IF($A{r}="","",SUMIFS(Ledger!$G$5:$G$204,Ledger!$B$5:$B$204,$A{r}))'), 'calc', MONEY)
        cell(ws, 'G%d' % r, f('=IF($A{r}="","",SUMIFS(Ledger!$H$5:$H$204,Ledger!$B$5:$B$204,$A{r},Ledger!$H$5:$H$204,">0"))'), 'calc', MONEY)
        cell(ws, 'H%d' % r, f('=IF($A{r}="","",SUMIFS(Ledger!$J$5:$J$204,Ledger!$B$5:$B$204,$A{r}))'), 'calc', MONEY)
        cell(ws, 'I%d' % r, f('=IF($A{r}="","",SUMIFS(Payments!$D$4:$D$1003,Payments!$H$4:$H$1003,$A{r},Payments!$A$4:$A$1003,">="&DATE(YEAR(Settings!$B$6),MONTH(Settings!$B$6),1),Payments!$A$4:$A$1003,"<"&EDATE(DATE(YEAR(Settings!$B$6),MONTH(Settings!$B$6),1),1)))'), 'calc', MONEY)
    cell(ws, 'A24', 'Total', 'soft', bold=True)
    for col in 'BCDEFGHI': cell(ws, '%s24' % col, '=SUM(%s4:%s23)' % (col, col), 'soft', MONEY if col > 'C' else None, bold=True, align='center' if col <= 'C' else None)
    ws.conditional_formatting.add('G4:G23', CellIsRule(operator='greaterThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    protect(ws)
    save(wb, 'Rentals-Rent-Ledger.xlsx')

# ------------------------------------------------------------------ rent receipt book
def receipt_book():
    wb = new_wb('rent-receipt-book-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 44])
    rg = sheet(wb, 'Register', [10, 12, 22, 34, 12, 12, 12, 13, 18, 18, 14], tab=BLUE)
    rc = sheet(wb, 'Receipt', [17, 20, 14, 18, 13, 18, 3, 12, 12, 12, 12], tab=BLUE)
    lists_sheet(wb, extra={3: ['India', 'Other'], 4: ['Indian', 'International']})
    title(st, 'Settings', 'Landlord details and how amounts are written.')
    for i, (k, v) in enumerate([('Landlord name', 'Suresh Kumar Gupta'), ('Landlord address', '45, Shastri Nagar, Jaipur, Rajasthan 302016'),
                                ('Landlord phone', '98290 00000'), ('Landlord PAN (India) or tax ID', 'ABCPG1234K'), ('Country', 'India'),
                                ('Currency symbol', '₹'), ('Currency name in words', 'Rupees'), ('Minor unit in words', 'Paise'), ('Words style', 'Indian')], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in')
    dv(st, 'B7', '=Lists!$C$1:$C$2'); dv(st, 'B11', '=Lists!$D$1:$D$2')
    cell(st, 'A13', 'Examples: ₹ Rupees Paise (Indian) · $ Dollars Cents · £ Pounds Pence · € Euros Cents · AED Dirhams Fils (International).', size=9, color=MUTED)
    label(st, 'A15', 'PAN check')
    st.merge_cells('B15:B16')
    cell(st, 'B15', '=IF(B7<>"India","Not needed outside India.",IF(MAX(Register!$K$4:$K$503)>100000,IF(B6="","⚠ A tenant\'s rent in this register is above ₹1,00,000 for the year: employers usually ask for the landlord\'s PAN. Type it above.","OK: PAN is filled in."),"OK: no tenant\'s rent in this register is above ₹1,00,000."))',
         'calc', wrap=True)
    st.row_dimensions[15].height = st.row_dimensions[16].height = 24
    protect(st)
    title(rg, 'Receipt register', 'One row per receipt. Never reuse a number; write "Cancelled" in Notes instead of deleting.')
    C = cur('Settings!$B$8')
    head(rg, 3, ['Receipt no.', 'Receipt date', 'Tenant name', 'Property address', 'Rent from', 'Rent to', C('Amount'), 'Paid by', 'Reference (UTR / cheque)', 'Notes', C('Tenant\'s total in register')])
    grid(rg, 4, 503, 'ABCDEFGHIJK', {'K': 'calc'}, {'B': DT, 'E': DT, 'F': DT, 'G': MONEY, 'K': MONEY},
         {'K': '=IF($C{r}="","",SUMIFS($G$4:$G$503,$C$4:$C$503,$C{r},$J$4:$J$503,"<>Cancelled"))'})
    dv(rg, 'H4:H503', '"Cash,UPI,Bank transfer,Cheque,Other"')
    rows = []
    for k in range(7):
        m = mon(FY, k)
        rows.append(('R-%03d' % (k + 1), m + datetime.timedelta(days=2), 'Neha Arora', 'Flat 2B, Lotus Residency, Malviya Nagar, Jaipur 302017',
                     m, mon(m, 1) - datetime.timedelta(days=1), 18000, 'UPI', 'UPI ref 6104%04d' % (2211 + k * 37), None))
    rows.append(('R-008', mon(FY, 6) + datetime.timedelta(days=4), 'Karan Singh', 'Room 3, 45 Shastri Nagar, Jaipur 302016',
                 mon(FY, 4), mon(FY, 7) - datetime.timedelta(days=1), 19500, 'Cash', None, 'Three months together'))
    put(rg, 4, rows)
    rg.freeze_panes = 'A4'
    # Receipt: helpers in H (white text), two copies on one A4 page
    ws = rc
    ws.merge_cells('A1:F1'); cell(ws, 'A1', 'Rent receipt: pick a receipt number', bold=True, size=13, color=BLUE)
    label(ws, 'A3', 'Receipt no.'); cell(ws, 'B3', 'R-001', 'in', align='center')
    dv(ws, 'B3', '=Register!$A$4:$A$503', 'Pick a receipt number from the Register sheet')
    cell(ws, 'C3', '=IF(H2=0,"Not found on the Register sheet","")', color='B91C1C', size=9)
    ws.merge_cells('C3:F3')
    def X(col): return 'IF($H$2=0,"",INDEX(Register!$%s$4:$%s$503,$H$2))' % (col, col)
    helpers = [('H2', '=IFERROR(MATCH($B$3,Register!$A$4:$A$503,0),0)'), ('H3', '=IF($H$2=0,0,N(INDEX(Register!$G$4:$G$503,$H$2)))'),
               ('H4', words_any('$H$3', 'Settings!$B$9', 'Settings!$B$10', 'Settings!$B$11')),
               ('H5', '=Settings!$B$8&" "&IF(Settings!$B$11="Indian",TEXT($H$3,"[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00"),TEXT($H$3,"#,##0.00"))'),
               ('H6', '=AND(Settings!$B$7="India",' + X('H') + '="Cash",$H$3>5000)')]
    for ref, f in helpers:
        cell(ws, ref, f, color='FFFFFF', size=8)
    def copy(r0, which):
        ws.merge_cells('A%d:D%d' % (r0, r0)); cell(ws, 'A%d' % r0, 'RENT RECEIPT', bold=True, size=16, color=BLUE)
        ws.merge_cells('E%d:F%d' % (r0, r0)); cell(ws, 'E%d' % r0, which, size=9, color=MUTED, align='right')
        ws.merge_cells('A%d:F%d' % (r0 + 1, r0 + 1)); cell(ws, 'A%d' % (r0 + 1), '=Settings!$B$3', bold=True, size=12)
        ws.merge_cells('A%d:F%d' % (r0 + 2, r0 + 2)); cell(ws, 'A%d' % (r0 + 2), '=Settings!$B$4&IF(Settings!$B$5<>"","   ·   Phone "&Settings!$B$5,"")', size=9, color=MUTED)
        r = r0 + 4
        label(ws, 'A%d' % r, 'Receipt No.'); cell(ws, 'B%d' % r, '=IF($H$2=0,"",$B$3)', 'soft', bold=True)
        label(ws, 'E%d' % r, 'Date'); cell(ws, 'F%d' % r, '=' + X('B'), 'soft', DT, align='left')
        r += 1
        ws.merge_cells('A%d:F%d' % (r, r + 2))
        cell(ws, 'A%d' % r, ('=IF($H$2=0,"","Received with thanks from "&%s&" the sum of "&$H$5&" ("&$H$4&") by "&IF(%s="UPI","UPI",LOWER(%s))&IF(%s<>""," (ref. "&%s&")","")'
                             '&" towards rent of the property at "&%s&" for the period shown below.")') % (X('C'), X('H'), X('H'), X('I'), X('I'), X('D')), wrap=True, size=11)
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical='top')
        for k in range(3): ws.row_dimensions[r + k].height = 20
        r += 3
        label(ws, 'A%d' % r, 'Amount'); ws.merge_cells('B%d:C%d' % (r, r)); cell(ws, 'B%d' % r, '=IF($H$2=0,"",$H$5)', 'soft', bold=True, size=12)
        label(ws, 'D%d' % r, 'Rent period'); cell(ws, 'E%d' % r, '=' + X('E'), 'soft', DT, align='center'); cell(ws, 'F%d' % r, '=IF($H$2=0,"","to  "&TEXT(%s,"dd-mmm-yyyy"))' % X('F'), 'soft')
        r += 1
        label(ws, 'A%d' % r, 'In words'); ws.merge_cells('B%d:F%d' % (r, r)); cell(ws, 'B%d' % r, '=$H$4', size=9)
        r += 1
        cell(ws, 'A%d' % r, '=IF(Settings!$B$7="India","Landlord PAN",IF(Settings!$B$6="","","Landlord tax ID"))', bold=True)
        cell(ws, 'B%d' % r, '=IF(Settings!$B$6="",IF(Settings!$B$7="India","Not given",""),Settings!$B$6)', bold=True)
        r += 2
        ws.merge_cells('E%d:F%d' % (r, r + 2))
        cell(ws, 'E%d' % r, '=IF($H$6,"Affix ₹1 revenue stamp and sign across","")', size=8, color=MUTED, align='center', wrap=True)
        for rr in range(r, r + 3):
            for cc in 'EF': ws['%s%d' % (cc, rr)].border = Border(left=Side(style='dashed', color='94A3B8') if cc == 'E' else None, right=Side(style='dashed', color='94A3B8') if cc == 'F' else None,
                                                                    top=Side(style='dashed', color='94A3B8') if rr == r else None, bottom=Side(style='dashed', color='94A3B8') if rr == r + 2 else None)
        r += 3
        ws.merge_cells('A%d:C%d' % (r, r)); cell(ws, 'A%d' % r, '="Tenant: "&' + X('C'), size=9)
        ws.merge_cells('D%d:F%d' % (r, r)); cell(ws, 'D%d' % r, "Landlord's signature", size=9, align='right')
        return r
    end = copy(5, 'Tenant copy')
    ws.merge_cells('A%d:F%d' % (end + 2, end + 2)); cell(ws, 'A%d' % (end + 2), '✂  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -', color='94A3B8', align='center')
    end2 = copy(end + 4, 'Landlord copy')
    ws.print_area = 'A1:F%d' % end2
    protect(ws)
    save(wb, 'Rentals-Rent-Receipt-Book.xlsx')

# ------------------------------------------------------------------ property expense & maintenance log
def expense_log():
    wb = new_wb('rent-expense-log-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [26, 18, 4, 28, 4, 28])
    ex = sheet(wb, 'Expenses', [12, 22, 12, 22, 32, 22, 13, 10, 12], tab=BLUE)
    rp = sheet(wb, 'Repairs', [12, 22, 12, 16, 34, 10, 18, 15, 12, 12, 10, 24], tab=BLUE)
    sm = sheet(wb, 'Summary', [26, 16, 16, 16, 18, 12, 14], tab=BLUE)
    lists_sheet(wb, extra={2: ['Owner', 'Tenant'], 3: ['High', 'Medium', 'Low'], 4: ['Open', 'In progress', 'Waiting for parts', 'Done']})
    CATS = ['Repairs (general)', 'Plumbing', 'Electrical', 'Painting', 'Carpentry', 'Appliance repair', 'Cleaning & pest control', 'Property tax',
            'Society / HOA charges', 'Insurance', 'Utilities (paid by owner)', 'Loan interest', 'Agent / brokerage', 'Legal & documentation', 'Furniture & fittings', 'Other']
    title(st, 'Settings', 'Currency, accounting year, your properties and expense categories.')
    label(st, 'A3', 'Currency symbol'); cell(st, 'B3', '₹', 'in', align='left')
    label(st, 'A4', 'Accounting year starts'); cell(st, 'B4', FY, 'in', DT, align='left')
    label(st, 'A5', 'Year ends'); cell(st, 'B5', '=EDATE(B4,12)-1', 'calc', DT, align='left')
    label(st, 'D2', 'Your properties'); label(st, 'F2', 'Expense categories (edit freely)')
    for i in range(3, 23):
        cell(st, 'D%d' % i, PROPS[i - 3] if i - 3 < len(PROPS) else None, 'in')
        cell(st, 'F%d' % i, CATS[i - 3] if i - 3 < len(CATS) else None, 'in')
    protect(st)
    C = cur('Settings!$B$3')
    title(ex, 'Expenses', 'One row per bill. "Charge to: Tenant" = the tenant must pay it back; it is not counted as your cost.')
    head(ex, 3, ['Date', 'Property', 'Unit', 'Category', 'Description', 'Vendor / paid to', C('Amount'), 'Charge to', 'Bill no.'])
    grid(ex, 4, 1003, 'ABCDEFGHI', {}, {'A': DT, 'G': MONEY})
    dv(ex, 'B4:B1003', '=Settings!$D$3:$D$22'); dv(ex, 'D4:D1003', '=Settings!$F$3:$F$22'); dv(ex, 'H4:H1003', '=Lists!$B$1:$B$2')
    D = lambda n: FY + datetime.timedelta(days=n)
    put(ex, 4, [(D(9), PROPS[0], 'Flat 101', 'Plumbing', 'Kitchen sink trap replaced', 'Raju Plumbing Works', 850, 'Owner', 'RP-221'),
                (D(20), PROPS[0], 'All', 'Society / HOA charges', 'Quarterly maintenance, 3 flats', 'Green View RWA', 13500, 'Owner', 'GV/Q1/88'),
                (D(35), PROPS[1], 'Shop 2', 'Electrical', 'New MCB and wiring check', 'Sharma Electric', 2400, 'Owner', '1174'),
                (D(48), PROPS[2], 'Whole house', 'Property tax', 'Municipal property tax, full year', 'Municipal Corporation', 9600, 'Owner', 'PT-2026-5531'),
                (D(66), PROPS[0], 'Flat 103', 'Painting', 'Repaint before new tenant', 'Colour Home Painters', 14500, 'Owner', 'CH-093'),
                (D(80), PROPS[2], 'Whole house', 'Appliance repair', 'Broken geyser knob (damage)', 'Service centre', 650, 'Tenant', 'SC-4471'),
                (D(95), PROPS[1], 'All', 'Insurance', 'Shop building fire insurance', 'Insurer', 7200, 'Owner', 'POL-88123'),
                (D(130), PROPS[2], 'Whole house', 'Cleaning & pest control', 'Termite treatment', 'PestAway', 3800, 'Owner', 'PA-640')])
    ex.freeze_panes = 'A4'; ex.auto_filter.ref = 'A3:I1003'
    title(rp, 'Repair requests', 'Days open count until the job is done. Open jobs older than 7 days turn red.')
    head(rp, 3, ['Reported on', 'Property', 'Unit', 'Reported by', 'Problem', 'Priority', 'Assigned to', 'Status', 'Fixed on', C('Cost'), 'Days open', 'Notes'])
    grid(rp, 4, 503, 'ABCDEFGHIJKL', {'K': 'calc'}, {'A': DT, 'I': DT, 'J': MONEY},
         {'K': '=IF($A{r}="","",IF($I{r}<>"",$I{r}-$A{r},TODAY()-$A{r}))'})
    dv(rp, 'B4:B503', '=Settings!$D$3:$D$22'); dv(rp, 'F4:F503', '=Lists!$C$1:$C$3'); dv(rp, 'H4:H503', '=Lists!$D$1:$D$4')
    T0 = lambda **k: TODAY - datetime.timedelta(**k)
    put(rp, 4, [(T0(days=40), PROPS[0], 'Flat 101', 'Rohit (tenant)', 'Kitchen sink leaking', 'High', 'Raju Plumbing Works', 'Done', T0(days=38), 850, None, ''),
                (T0(days=12), PROPS[1], 'Shop 1', 'Sandhu Electricals', 'Shutter lock stiff', 'Medium', 'Local locksmith', 'Open', None, None, None, 'Call again'),
                (T0(days=4), PROPS[2], 'Whole house', 'David Thomas', 'Bathroom exhaust fan not working', 'Low', 'Sharma Electric', 'Waiting for parts', None, None, None, ''),
                (T0(days=2), PROPS[0], 'Flat 102', 'Priya (tenant)', 'Balcony door does not close', 'Medium', 'Carpenter', 'In progress', None, None, None, '')])
    rp.conditional_formatting.add('A4:L503', FormulaRule(formula=['AND($A4<>"",$H4<>"Done",N($K4)>7)'], fill=PatternFill('solid', fgColor='FEE2E2')))
    rp.conditional_formatting.add('A4:L503', FormulaRule(formula=['$H4="Done"'], font=Font(color='94A3B8')))
    rp.freeze_panes = 'A4'; rp.auto_filter.ref = 'A3:L503'
    ws = sm
    title(ws, 'Summary for the year')
    ws['A2'] = '="From "&TEXT(Settings!B4,"dd-mmm-yyyy")&" to "&TEXT(Settings!B5,"dd-mmm-yyyy")&"   ·   amounts in "&Settings!B3'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    IN = 'Expenses!$A$4:$A$1003,">="&Settings!$B$4,Expenses!$A$4:$A$1003,"<="&Settings!$B$5'
    head(ws, 4, ['Property', C('Owner expenses'), C('Rent received (type)'), C('Net income'), C('To recover from tenants'), 'Open repairs', 'Expenses % of rent'])
    rent_in = {0: 360000, 1: 516000, 2: 280000}
    for i in range(20):
        r, s = 5 + i, 3 + i
        f = lambda t: t.format(r=r, s=s, IN=IN)
        cell(ws, 'A%d' % r, f('=IF(Settings!D{s}="","",Settings!D{s})'), 'box')
        cell(ws, 'B%d' % r, f('=IF($A{r}="","",SUMIFS(Expenses!$G$4:$G$1003,Expenses!$B$4:$B$1003,$A{r},Expenses!$H$4:$H$1003,"<>Tenant",{IN}))'), 'calc', MONEY)
        cell(ws, 'C%d' % r, rent_in.get(i), 'in', MONEY)
        cell(ws, 'D%d' % r, f('=IF($A{r}="","",N(C{r})-B{r})'), 'calc', MONEY, bold=True)
        cell(ws, 'E%d' % r, f('=IF($A{r}="","",SUMIFS(Expenses!$G$4:$G$1003,Expenses!$B$4:$B$1003,$A{r},Expenses!$H$4:$H$1003,"Tenant",{IN}))'), 'calc', MONEY)
        cell(ws, 'F%d' % r, f('=IF($A{r}="","",COUNTIFS(Repairs!$B$4:$B$503,$A{r},Repairs!$H$4:$H$503,"<>Done",Repairs!$A$4:$A$503,"<>"))'), 'calc', align='center')
        cell(ws, 'G%d' % r, f('=IF(OR($A{r}="",N(C{r})=0),"",B{r}/C{r})'), 'calc', '0.0%', align='center')
    cell(ws, 'A25', 'Total', 'soft', bold=True)
    for col in 'BCDEF': cell(ws, '%s25' % col, '=SUM(%s5:%s24)' % (col, col), 'soft', MONEY if col != 'F' else None, bold=True, align='center' if col == 'F' else None)
    cell(ws, 'G25', '=IF(C25=0,"",B25/C25)', 'soft', '0.0%', bold=True, align='center')
    ws.conditional_formatting.add('D5:D25', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    head(ws, 28, ['Category', C('Owner expenses'), C('Charged to tenants')])
    for i in range(20):
        r, s = 29 + i, 3 + i
        f = lambda t: t.format(r=r, s=s, IN=IN)
        cell(ws, 'A%d' % r, f('=IF(Settings!F{s}="","",Settings!F{s})'), 'box')
        cell(ws, 'B%d' % r, f('=IF($A{r}="","",SUMIFS(Expenses!$G$4:$G$1003,Expenses!$D$4:$D$1003,$A{r},Expenses!$H$4:$H$1003,"<>Tenant",{IN}))'), 'calc', MONEY)
        cell(ws, 'C%d' % r, f('=IF($A{r}="","",SUMIFS(Expenses!$G$4:$G$1003,Expenses!$D$4:$D$1003,$A{r},Expenses!$H$4:$H$1003,"Tenant",{IN}))'), 'calc', MONEY)
    cell(ws, 'A49', 'Total', 'soft', bold=True)
    for col in 'BC': cell(ws, '%s49' % col, '=SUM(%s29:%s48)' % (col, col), 'soft', MONEY, bold=True)
    ws.merge_cells('A51:G52')
    cell(ws, 'A51', 'Net income here = rent received minus your own expenses, before income tax. Tax rules on repairs, improvements and standard deductions differ by country: ask your accountant.', wrap=True, size=9, color=MUTED)
    protect(ws)
    save(wb, 'Rentals-Property-Expense-Log.xlsx')

# ------------------------------------------------------------------ PG / hostel bed register
def bed_register():
    wb = new_wb('rent-pg-bed-register-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [26, 26, 4, 24, 4, 24])
    bd = sheet(wb, 'Beds', [11, 9, 8, 12, 7, 13, 26, 12], tab=BLUE)
    rs = sheet(wb, 'Residents', [22, 14, 15, 18, 26, 30, 22, 9, 12, 12, 12, 12, 18, 13, 14, 22], tab=BLUE)
    pm = sheet(wb, 'Payments', [12, 22, 12, 12, 14, 18, 10], tab=BLUE)
    du = sheet(wb, 'Dues', [22, 9, 12, 12, 16, 9, 13, 13, 13, 14, 70, 11, 11], tab=BLUE)
    lists_sheet(wb, extra={2: ['Single', 'Double', 'Triple', 'Four', 'Dormitory'], 3: ['Yes', 'No'], 4: ['Done', 'Pending', 'Not required']})
    title(st, 'Settings', 'PG or hostel details, rent due day and your lists.')
    for i, (k, v, fmt) in enumerate([('PG / hostel name', 'Sunrise PG & Hostel', None), ('Currency symbol', '₹', None),
                                     ('Position as on', '=TODAY()', DT), ('Rent due day of the month', 5, '0')], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    label(st, 'D2', 'ID types'); label(st, 'F2', 'Food plans')
    IDS = ['Aadhaar card', 'Passport', 'Driving licence', 'Voter ID', 'College / company ID', 'Other']
    FOOD = ['No food', 'Breakfast only', 'Breakfast & dinner', 'All meals']
    for i in range(3, 13):
        cell(st, 'D%d' % i, IDS[i - 3] if i - 3 < len(IDS) else None, 'in')
        cell(st, 'F%d' % i, FOOD[i - 3] if i - 3 < len(FOOD) else None, 'in')
    label(st, 'A9', 'Payment modes'); cell(st, 'A10', 'Cash, UPI, Bank transfer, Card (fixed list)', size=9, color=MUTED)
    protect(st)
    C = cur('Settings!$B$4')
    AS = 'Settings!$B$5'
    # Beds
    ws = bd
    title(ws, 'Beds', 'One row per bed. Current resident and status fill in from the Residents sheet.')
    for i, (k, f, fmt) in enumerate([('Beds', '=COUNTA(A7:A206)', '0'), ('Occupied', '=COUNTIF(H7:H206,"Occupied")', '0'), ('Vacant', '=COUNTIF(H7:H206,"Vacant")', '0'),
                                     ('Occupancy', '=IF(B3=0,0,B4/B3)', '0%')], 3):
        label(ws, 'A%d' % i, k); cell(ws, 'B%d' % i, f, 'soft', fmt, bold=True, align='center')
    ws.merge_cells('D3:F3'); label(ws, 'D3', '="Rent roll, occupied beds ("&Settings!$B$4&")"')
    cell(ws, 'G3', '=SUMIFS(F7:F206,H7:H206,"Occupied")', 'soft', MONEY, bold=True)
    ws.merge_cells('D4:F4'); label(ws, 'D4', '="Rent lost on vacant beds ("&Settings!$B$4&")"')
    cell(ws, 'G4', '=SUMIFS(F7:F206,H7:H206,"Vacant")', 'soft', MONEY, bold=True)
    head(ws, 6, ['Bed code', 'Room', 'Floor', 'Sharing', 'AC', C('Rent per bed'), 'Current resident', 'Status'])
    grid(ws, 7, 206, 'ABCDEFGH', {'G': 'calc', 'H': 'calc'}, {'F': MONEY},
         {'G': '=IF($A{r}="","",IFERROR(LOOKUP(2,1/((Residents!$H$4:$H$503=$A{r})*(Residents!$J$4:$J$503<>"")*(Residents!$J$4:$J$503<=%s)'
               '*(((Residents!$K$4:$K$503="")+(Residents!$K$4:$K$503>=%s))>0)),Residents!$A$4:$A$503),""))' % (AS, AS),
          'H': '=IF($A{r}="","",IF($G{r}="","Vacant","Occupied"))'})
    dv(ws, 'D7:D206', '=Lists!$B$1:$B$5'); dv(ws, 'E7:E206', '=Lists!$C$1:$C$2')
    beds = []
    for room, share, ac, rent in (('101', 'Double', 'Yes', 9500), ('102', 'Triple', 'No', 7000), ('201', 'Single', 'Yes', 14000), ('202', 'Double', 'No', 8000)):
        n = {'Single': 1, 'Double': 2, 'Triple': 3}[share]
        for b in 'ABC'[:n]: beds.append(('%s-%s' % (room, b), room, room[0], share, ac, rent))
    put(ws, 7, beds)
    ws.conditional_formatting.add('H7:H206', CellIsRule(operator='equal', formula=['"Vacant"'], fill=GREEN, font=Font(bold=True, color='166534')))
    ws.freeze_panes = 'B7'
    protect(ws)
    # Residents
    title(rs, 'Residents', 'One row per stay. Type the check-out date when someone leaves; the bed becomes vacant. Keep this sheet private.')
    head(rs, 3, ['Name', 'Phone', 'ID type', 'ID number', 'Emergency contact (name, phone)', 'Home address', 'College / workplace', 'Bed', C('Monthly rent'),
                 'Check-in', 'Check-out', C('Deposit'), 'Food plan', 'Police verification', C('Extra charges / adjustment'), 'Notes'])
    grid(rs, 4, 503, 'ABCDEFGHIJKLMNOP', {}, {'I': MONEY, 'J': DT, 'K': DT, 'L': MONEY, 'O': MONEY})
    dv(rs, 'C4:C503', '=Settings!$D$3:$D$12'); dv(rs, 'H4:H503', '=Beds!$A$7:$A$206', 'Pick the bed code'); dv(rs, 'M4:M503', '=Settings!$F$3:$F$12'); dv(rs, 'N4:N503', '=Lists!$D$1:$D$3')
    M = mon(TODAY)
    people = [('Riya Shah', '98111 10001', 'Aadhaar card', 'XXXX XXXX 4821', 'Mr. Shah (father) 98111 20001', 'Surat, Gujarat', 'Symbiosis College', '101-A', 9500, mon(M, -5), None, 19000, 'Breakfast & dinner', 'Done', None),
              ('Ananya Iyer', '98111 10002', 'Passport', 'XXXXX6620', 'Mrs. Iyer (mother) 98111 20002', 'Chennai, Tamil Nadu', 'Infosys, Hinjewadi', '101-B', 9500, mon(M, -3) + datetime.timedelta(days=9), None, 19000, 'All meals', 'Done', -3000),
              ('Pooja Verma', '98111 10003', 'Voter ID', 'XXX1187', 'Mr. Verma (father) 98111 20003', 'Indore, Madhya Pradesh', 'Fergusson College', '102-A', 7000, mon(M, -8), mon(M, -1) + datetime.timedelta(days=14), 14000, 'Breakfast only', 'Done', None),
              ('Simran Kaur', '98111 10004', 'Aadhaar card', 'XXXX XXXX 9930', 'Mr. Singh (brother) 98111 20004', 'Ludhiana, Punjab', 'Deloitte', '201-A', 14000, mon(M, -2), None, 28000, 'No food', 'Pending', 1200),
              ('Fatima Khan', '98111 10005', 'Driving licence', 'XX-0420-77', 'Mrs. Khan (mother) 98111 20005', 'Lucknow, Uttar Pradesh', 'Pune University', '102-B', 7000, mon(M, -1), None, 14000, 'All meals', 'Pending', None)]
    put(rs, 4, people)
    rs.freeze_panes = 'B4'
    # Payments
    title(pm, 'Payments', 'One row per payment, with the month it is for.')
    head(pm, 3, ['Date paid', 'Resident', 'For month', C('Amount'), 'Mode', 'Reference', 'Bed'])
    grid(pm, 4, 1003, 'ABCDEFG', {'G': 'calc'}, {'A': DT, 'C': MON, 'D': MONEY},
         {'G': '=IF($B{r}="","",IFERROR(INDEX(Residents!$H$4:$H$503,MATCH($B{r},Residents!$A$4:$A$503,0))&"",""))'})
    dv(pm, 'B4:B1003', '=Residents!$A$4:$A$503'); dv(pm, 'E4:E1003', '"Cash,UPI,Bank transfer,Card"')
    pays = []
    for p in people:
        m = mon(p[9]); stop = mon(p[10]) if p[10] else (M if TODAY.day >= 5 else mon(M, -1))
        while m <= stop:
            if not (p[0] == 'Simran Kaur' and m == stop) and not (p[0] == 'Fatima Khan' and m == stop):
                amt = p[8] + (p[14] if p[14] and p[14] < 0 and m == mon(p[9]) else 0)   # pro-rata first month
                pays.append((min(TODAY, m + datetime.timedelta(days=3)), p[0], m, amt, 'UPI' if len(pays) % 3 else 'Cash', None))
            m = mon(m, 1)
    pays.append((TODAY, 'Fatima Khan', stop, 4000, 'Cash', 'Part payment'))
    pays.sort(key=lambda x: x[0])
    put(pm, 4, pays)
    pm.freeze_panes = 'A4'
    # Dues
    ws = du
    title(ws, 'Rent dues by resident')
    ws['A2'] = '="As on "&TEXT(Settings!B5,"dd-mmm-yyyy")&"   ·   rent due on day "&Settings!B6&" of each month"'
    ws['A2'].font = Font(size=10, italic=True, color=MUTED)
    head(ws, 3, ['Resident', 'Bed', C('Rent'), 'Check-in', 'Stay', 'Months due', C('Due (incl. extras)'), C('Paid'), C('Balance'), 'Phone',
                 'Reminder message (copy into WhatsApp or SMS)', 'From', 'Up to'])
    M0 = 'DATE(YEAR(%s),MONTH(%s),1)' % (AS, AS)
    for i in range(500):
        r = 4 + i
        f = lambda t: t.format(r=r, AS=AS, M0=M0)
        cell(ws, 'A%d' % r, f('=IF(Residents!A{r}="","",Residents!A{r})'), 'box')
        cell(ws, 'B%d' % r, f('=IF($A{r}="","",Residents!H{r}&"")'), 'box', align='center')
        cell(ws, 'C%d' % r, f('=IF($A{r}="","",IF(N(Residents!I{r})>0,Residents!I{r},IFERROR(INDEX(Beds!$F$7:$F$206,MATCH(Residents!H{r},Beds!$A$7:$A$206,0)),0)))'), 'calc', MONEY)
        cell(ws, 'D%d' % r, f('=IF(OR($A{r}="",Residents!J{r}=""),"",Residents!J{r})'), 'calc', DT)
        cell(ws, 'E%d' % r, f('=IF($A{r}="","",IF(Residents!J{r}="","No check-in date",IF(Residents!J{r}>{AS},"Booked",IF(AND(Residents!K{r}<>"",Residents!K{r}<{AS}),"Left","Staying"))))'), 'calc', align='center')
        cell(ws, 'L%d' % r, f('=IF(OR($A{r}="",Residents!$J{r}=""),"",DATE(YEAR(Residents!$J{r}),MONTH(Residents!$J{r}),1))'), 'calc', MON)
        cell(ws, 'M%d' % r, f('=IF(L{r}="","",MAX(L{r},IF(Residents!$K{r}<>"",MIN(DATE(YEAR(Residents!$K{r}),MONTH(Residents!$K{r}),1),{M0}),'
                              'IF(DAY({AS})>=MAX(1,N(Settings!$B$6)),{M0},EDATE({M0},-1)))))'), 'calc', MON)
        cell(ws, 'F%d' % r, f('=IF($A{r}="","",IF(OR(L{r}="",Residents!$J{r}>{AS}),0,(YEAR(M{r})-YEAR(L{r}))*12+MONTH(M{r})-MONTH(L{r})+1))'), 'calc', align='center')
        cell(ws, 'G%d' % r, f('=IF($A{r}="","",F{r}*C{r}+IF(F{r}>0,N(Residents!O{r}),0))'), 'calc', MONEY)
        cell(ws, 'H%d' % r, f('=IF($A{r}="","",SUMIF(Payments!$B$4:$B$1003,$A{r},Payments!$D$4:$D$1003))'), 'calc', MONEY)
        cell(ws, 'I%d' % r, f('=IF($A{r}="","",G{r}-H{r})'), 'calc', MONEY, bold=True)
        cell(ws, 'J%d' % r, f('=IF($A{r}="","",Residents!B{r}&"")'), 'box')
        cell(ws, 'K%d' % r, f('=IF(OR($A{r}="",N(I{r})<=0.005),"","Hi "&A{r}&", your rent balance at "&Settings!$B$3&" is "&Settings!$B$4&" "&TEXT(I{r},"#,##0")&" as on "&TEXT({AS},"dd-mmm-yyyy")&". Please pay at the earliest. Thank you.")'), size=9)
    ws.conditional_formatting.add('I4:I503', CellIsRule(operator='greaterThan', formula=['0.005'], fill=RED, font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('A4:K503', FormulaRule(formula=['$E4="Left"'], font=Font(color='94A3B8')))
    ws.freeze_panes = 'B4'
    protect(ws)
    save(wb, 'Rentals-PG-Hostel-Bed-Register.xlsx')

# ------------------------------------------------------------------ Word: rent agreement / lease draft
def box_note(d, lines, fill='FFF4E5'):
    t = d.add_table(rows=1, cols=1); t.style = 'Table Grid'
    c = t.rows[0].cells[0]; shade(c, fill); c.text = ''
    for i, (txt, b) in enumerate(lines):
        p = c.paragraphs[0] if i == 0 else c.add_paragraph()
        r = p.add_run(txt); r.bold = b; r.font.size = Pt(9.5)
    para(d, space=4)

def clause(d, n, head_, *texts):
    para(d, '%d. %s' % (n, head_), bold=True, color=BLUE, space=2)
    for t in texts: para(d, t, space=4)

def lease():
    d = doc_new('rent-lease-agreement-word')
    box_note(d, [('STARTING DRAFT – CHECK LOCAL LAW BEFORE SIGNING', True),
                 ('This is a general template, not legal advice. Tenancy law, deposit limits, notice periods, stamp duty and registration rules differ by country, state and city. '
                  'Have this draft checked by a lawyer or a qualified local agent, adapt it, and delete this box before printing.', False)])
    para(d, 'RENT AGREEMENT (RESIDENTIAL LEASE)', bold=True, size=16, align='c', color=BLUE, space=10)
    para(d, 'This Rent Agreement is made on [DD-MM-YYYY] at [City] between:', space=6)
    kv(d, [('Landlord', '[Full name], [son / daughter / wife] of [ ], aged [ ]'), ('Landlord address', '[Permanent address]'),
           ('Landlord ID / PAN', '[ ]'), ('Landlord phone, email', '[ ]'),
           ('Tenant', '[Full name], [son / daughter / wife] of [ ], aged [ ]'), ('Tenant permanent address', '[ ]'),
           ('Tenant ID', '[ ]'), ('Tenant phone, email', '[ ]')], cols=1)
    para(d, 'The Landlord and the Tenant agree as follows:', space=6)
    clause(d, 1, 'Property', 'The Landlord lets to the Tenant the residential property at [full address with flat number, floor, building, locality, city, PIN / postcode] '
           '("the Property"), described in Schedule 1, together with the fittings and furniture listed in Schedule 2.')
    clause(d, 2, 'Term', 'The tenancy starts on [DD-MM-YYYY] and ends on [DD-MM-YYYY], a period of [11] months, unless ended earlier under clause 18 or renewed under clause 17.')
    clause(d, 3, 'Rent', 'The monthly rent is [currency] [amount] ([amount in words]), payable in advance on or before the [5th] day of each month. '
           'The rent [includes / does not include] society or maintenance charges of [currency] [ ] per month.')
    clause(d, 4, 'Payment and receipts', 'Rent will be paid by [bank transfer / UPI / cheque] to the Landlord\'s account [account name, number, bank, IFSC / sort code]. '
           'The Landlord will give a signed receipt for every payment on request, including any receipt the Tenant needs for a house rent allowance claim.')
    clause(d, 5, 'Late payment', 'If rent is not received within [7] days of the due date, the Tenant will pay a late fee of [currency] [ ] [per day / once per month], '
           'unless the delay is agreed in writing. [Delete this clause if late fees are not allowed where the Property is.]')
    clause(d, 6, 'Security deposit', 'The Tenant has paid a refundable security deposit of [currency] [amount] ([amount in words]) by [mode, reference] on [DD-MM-YYYY], '
           'equal to [two] months\' rent. No interest is payable on the deposit [unless local law requires it]. The deposit cannot be used by the Tenant as rent for the last months.')
    clause(d, 7, 'Return of the deposit', 'The Landlord will return the deposit within [15] days after the Tenant hands back vacant possession and all keys, '
           'after deducting only: unpaid rent, unpaid utility bills, and the reasonable cost of repairing damage beyond normal wear and tear, each shown in writing with bills. '
           'The condition on the move-in inspection checklist signed by both parties will be the starting point.')
    clause(d, 8, 'Utilities and other charges', 'The Tenant will pay electricity, water, gas, internet and [ ] bills for the tenancy period directly or as billed. '
           'Opening meter readings are recorded on the move-in checklist. Property tax, building insurance and [ ] will be paid by the Landlord.')
    clause(d, 9, 'Use of the Property', 'The Property will be used only as a private residence for the Tenant and the family members named here: [names]. '
           'It will not be used for any business, illegal or immoral purpose, and the Tenant will follow the rules of the housing society or building.')
    clause(d, 10, 'Guests', 'Guests may stay for up to [14] days at a time. Longer stays need the Landlord\'s written consent.')
    clause(d, 11, 'Subletting', 'The Tenant will not sublet, assign or part with possession of the whole or any part of the Property, or take paying guests, without the Landlord\'s written consent.')
    clause(d, 12, 'Day-to-day repairs', 'The Tenant will keep the Property clean and in good condition and will pay for minor repairs up to [currency] [ ] each, '
           'such as bulbs, tap washers and blocked drains caused by misuse.')
    clause(d, 13, 'Major repairs', 'The Landlord will carry out structural repairs and repairs to wiring, plumbing, roof, walls and appliances provided by the Landlord, '
           'within a reasonable time after the Tenant reports them in writing. In an emergency the Tenant may arrange an urgent repair and claim the cost against bills.')
    clause(d, 14, 'Alterations', 'The Tenant will not make structural changes, drill into tiled areas, paint or fix anything permanent without the Landlord\'s written consent, '
           'and will restore the Property to its original condition at the end if asked.')
    clause(d, 15, 'Landlord\'s entry', 'The Landlord or the Landlord\'s agent may inspect the Property at a reasonable time after giving at least [24] hours\' notice, '
           'or at any time in an emergency.')
    clause(d, 16, 'Pets and smoking', 'Pets are [allowed / not allowed / allowed with written consent]. Smoking inside the Property is [allowed / not allowed].')
    clause(d, 17, 'Rent increase and renewal', 'The tenancy may be renewed by mutual written agreement before it ends, with a rent increase of [ ] percent '
           '[or as agreed at the time]. A fresh agreement will be signed for each renewal.')
    clause(d, 18, 'Lock-in and notice', 'Neither party may end the tenancy during the first [ ] months (the lock-in period) except for a serious breach. '
           'After that, either party may end it by giving [one] month\'s written notice. The Landlord may end it earlier as allowed by law if the Tenant fails to pay rent for [two] months or seriously breaches this Agreement.')
    clause(d, 19, 'Handing back the Property', 'At the end of the tenancy the Tenant will hand back vacant possession, all keys, remotes and access cards, '
           'with the Property clean and in the condition recorded at move-in, except for normal wear and tear. Both parties will sign the move-out section of the checklist.')
    clause(d, 20, 'Law and disputes', 'This Agreement is governed by the laws in force where the Property is. The parties will first try to settle any dispute by talking to each other; '
           'if that fails, by [mediation / the Rent Authority / the courts at (city)].')
    para(d, 'Signed by both parties on the date written above, each having read and understood every page.', space=24)
    t = d.add_table(rows=3, cols=2)
    for i, (a, b) in enumerate([('LANDLORD', 'TENANT'), ('\n\n____________________________\nSignature', '\n\n____________________________\nSignature'), ('Name: [ ]', 'Name: [ ]')]):
        t.rows[i].cells[0].text, t.rows[i].cells[1].text = a, b
    para(d, space=14)
    para(d, 'Witnesses', bold=True, color=BLUE)
    t = d.add_table(rows=2, cols=2)
    for i in range(2):
        for j in range(2): t.rows[i].cells[j].text = ('\n____________________________\nWitness %d: name and address\n[ ]' % (j + 1)) if i == 0 else ''
    d.add_page_break()
    para(d, 'SCHEDULE 1 – THE PROPERTY', bold=True, size=13, color=BLUE, space=6)
    kv(d, [('Type', '[Flat / house / room / studio]'), ('Address', '[ ]'), ('Floor', '[ ]'), ('Area', '[ ] sq ft / sq m'),
           ('Rooms', '[ ] bedrooms, [ ] bathrooms, kitchen, [ ]'), ('Parking', '[Covered / open, slot number]'), ('Furnishing', '[Unfurnished / semi-furnished / furnished]')], cols=1)
    para(d, 'SCHEDULE 2 – FITTINGS AND FURNITURE', bold=True, size=13, color=BLUE, space=6)
    table(d, ['#', 'Item', 'Qty', 'Condition at start', 'Remarks'],
          [['1', 'Ceiling fans', '[ ]', '[Good]', ''], ['2', 'Lights / tube lights', '[ ]', '', ''], ['3', 'Air conditioner', '[ ]', '', ''], ['4', 'Geyser / water heater', '[ ]', '', ''],
           ['5', 'Wardrobes', '[ ]', '', ''], ['6', 'Bed with mattress', '[ ]', '', ''], ['7', 'Refrigerator', '[ ]', '', ''], ['8', 'Curtains and rods', '[ ]', '', '']],
          [1, 6, 1.5, 4, 4.5], fill_rows=6)
    para(d, 'NOTES FOR INDIA (delete if the Property is elsewhere)', bold=True, size=12, color=BLUE, space=6)
    for s in ['Many agreements are made for 11 months because, under the Registration Act, a lease from year to year or for more than one year must be registered at the sub-registrar\'s office.',
              'Stamp duty still applies to an 11-month agreement. The amount and the method (stamp paper or e-stamping) depend on the state; an unstamped agreement may not be accepted as evidence.',
              'Many cities ask landlords to complete tenant police verification. Do it at the local police station or online where your state police offer it.',
              'Some states have adopted the Model Tenancy Act, 2021 or their own rent laws, which can cap the security deposit and require informing a Rent Authority. Check what applies where the Property is.',
              'Employers usually ask for the landlord\'s PAN on rent receipts when the rent for the year is above ₹1,00,000.']:
        d.add_paragraph(s, style='List Bullet')
    dsave(d, 'Rentals-Rent-Agreement-Draft.docx')

# ------------------------------------------------------------------ Word: move-in / move-out checklist
def checklist():
    d = doc_new('rent-inspection-checklist-word')
    para(d, 'MOVE-IN / MOVE-OUT INSPECTION CHECKLIST', bold=True, size=16, color=BLUE, space=2)
    para(d, 'Walk through together, write the condition of every item and take dated photos. Both sign at move-in and again at move-out.', size=9, italic=True, space=8)
    kv(d, [('Property address', '[ ]'), ('Unit', '[ ]'), ('Landlord', '[ ]'), ('Tenant', '[ ]'), ('Move-in date', '[DD-MM-YYYY]'), ('Move-out date', '[DD-MM-YYYY]')])
    para(d, 'Condition words: New · Good · Fair (worn) · Damaged · Missing · Not working', size=9, color=MUTED, space=6)
    H = ['Item', 'Condition at move-in', 'Condition at move-out', 'Photo no.']
    W = [4.8, 5, 5, 2.2]
    rooms = [('Entrance and hall', ['Main door, lock and handle', 'Doorbell, peephole', 'Walls and ceiling', 'Floor', 'Lights and switches', 'Shoe rack / cupboard']),
             ('Living room', ['Walls and ceiling (paint, cracks, damp)', 'Floor / tiles', 'Windows, glass, grills, mosquito mesh', 'Curtains and rods', 'Fans and lights', 'Switches and sockets', 'Furniture (list each)']),
             ('Kitchen', ['Counter top and sink', 'Taps and drain (no leaks)', 'Cabinets and drawers', 'Chimney / exhaust fan', 'Gas pipeline / cylinder and regulator', 'Stove / hob', 'Refrigerator', 'Water purifier', 'Tiles and walls']),
             ('Bedroom 1', ['Door and lock', 'Walls and ceiling', 'Floor', 'Windows and curtains', 'Wardrobe (doors, keys, shelves)', 'Bed and mattress', 'Fan, lights, AC (remote)', 'Switches and sockets']),
             ('Bedroom 2', ['Door and lock', 'Walls and ceiling', 'Floor', 'Windows and curtains', 'Wardrobe', 'Fan, lights, AC (remote)', 'Switches and sockets']),
             ('Bathroom 1', ['Door and lock', 'Wash basin and mirror', 'Taps, shower, health faucet', 'Toilet and flush', 'Geyser / water heater', 'Exhaust fan', 'Tiles, grout and drain']),
             ('Bathroom 2', ['Wash basin and mirror', 'Taps and shower', 'Toilet and flush', 'Geyser', 'Tiles and drain']),
             ('Balcony, utility and outside', ['Balcony door and railing', 'Washing machine point and drain', 'Clothes drying stand', 'Parking slot', 'Letter box', 'Garden / terrace'])]
    for name, items in rooms:
        para(d, name, bold=True, color=BLUE, space=2)
        table(d, H, [[it, '', '', ''] for it in items], W, fill_rows=1)
    para(d, 'Meter readings', bold=True, color=BLUE, space=2)
    table(d, ['Meter', 'Meter / account no.', 'Reading at move-in', 'Reading at move-out'],
          [['Electricity', '', '', ''], ['Water', '', '', ''], ['Gas (piped)', '', '', ''], ['Gas cylinders (number full / empty)', '', '', '']], [4.8, 4, 4, 4.2])
    para(d, 'Keys, remotes and cards', bold=True, color=BLUE, space=2)
    table(d, ['Item', 'Number given at move-in', 'Number returned at move-out', 'Remarks'],
          [['Main door keys', '', '', ''], ['Bedroom / cupboard keys', '', '', ''], ['Letter box key', '', '', ''], ['AC / TV remotes', '', '', ''], ['Gate remote / access cards', '', '', '']], [4.8, 4, 4, 4.2])
    para(d, 'Move-in confirmation', bold=True, color=BLUE, space=2)
    para(d, 'We have inspected the Property together and agree that the move-in column is a true record of its condition on the move-in date.', size=10, space=18)
    para(d, 'Landlord: ____________________   Tenant: ____________________   Date: ____________', space=14)
    para(d, 'Deposit settlement at move-out', bold=True, color=BLUE, space=2)
    table(d, ['Description', 'Amount', 'Bill / reason'],
          [['Security deposit held', '[ ]', ''], ['Less: unpaid rent', '[ ]', ''], ['Less: unpaid utility bills', '[ ]', ''], ['Less: repair of damage beyond normal wear and tear', '[ ]', '[Item, bill no.]'],
           ['Less: cleaning (only if agreed in the rent agreement)', '[ ]', ''], ['Amount returned to tenant', '[ ]', '[Mode, reference, date]']], [8.5, 3, 5.5])
    para(d, 'We agree the move-out column and the deposit settlement above. The tenant has handed back all keys listed and vacant possession.', size=10, space=18)
    para(d, 'Landlord: ____________________   Tenant: ____________________   Date: ____________', space=10)
    para(d, 'Normal wear and tear (fading paint, light marks, worn carpets in walkways) is usually not charged to the tenant. Deposit rules differ by place: check your agreement and local law.', size=8, italic=True, color=MUTED)
    dsave(d, 'Rentals-Move-In-Move-Out-Checklist.docx')

MAKERS = [rent_ledger, receipt_book, expense_log, bed_register, lease, checklist]
