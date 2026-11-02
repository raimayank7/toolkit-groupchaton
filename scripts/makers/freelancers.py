"""Freelancers & Agencies (worldwide): files for content/kits/freelancers.py. Any currency (amounts use plain number
   formats and the headers show the currency code) and a date style picked on each Settings sheet (conditional number formats)."""
from make_files import *
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.formatting.rule import Rule
from openpyxl.workbook.defined_name import DefinedName

TEAL = '0F766E'
NUM = '#,##0.00'
HRS = '0.00'
TIME = 'hh:mm'
SYSDATE = 'mm-dd-yy'   # built-in format 14: Excel shows it in the computer's own short date style
CURS = ['USD', 'EUR', 'GBP', 'INR', 'AED', 'CAD', 'AUD', 'SGD', 'NZD', 'CHF', 'JPY', 'ZAR', 'SAR', 'QAR', 'MYR', 'PHP', 'PKR', 'BDT', 'LKR',
        'NPR', 'NGN', 'KES', 'EGP', 'BRL', 'MXN', 'HKD', 'SEK', 'NOK', 'DKK', 'PLN', 'IDR', 'THB', 'Other']
DSTYLES = ['Computer setting', 'DD-MM-YYYY', 'MM/DD/YYYY', 'YYYY-MM-DD', 'DD MMM YYYY']
DFMT = {'DD-MM-YYYY': 'dd-mm-yyyy', 'MM/DD/YYYY': 'mm"/"dd"/"yyyy', 'YYYY-MM-DD': 'yyyy-mm-dd', 'DD MMM YYYY': 'dd mmm yyyy'}
YESNO = ['Yes', 'No']
D = lambda n: TODAY + datetime.timedelta(days=n)
M = lambda n: MONTH1 + datetime.timedelta(days=n)

def fl_lists(wb, extra=None):
    cols = {1: CURS, 2: DSTYLES, 3: YESNO}
    cols.update(extra or {})
    ws = wb.create_sheet('Lists')
    for col, vals in cols.items():
        for i, v in enumerate(vals, 1): ws.cell(i, col, v)
    ws.sheet_state = 'hidden'
    return ws

def lst(col, n): return '=Lists!$%s$1:$%s$%d' % (col, col, n)

def date_style(wb, ref):
    """Settings cell holding the date style → name DateStyle; dated ranges get one conditional number format per style."""
    wb.defined_names['DateStyle'] = DefinedName('DateStyle', attr_text=ref)

def dates(ws, rng):
    for i, (k, f) in enumerate(DFMT.items()):
        dx = DifferentialStyle(numFmt=NumberFormat(numFmtId=300 + i, formatCode=f))
        ws.conditional_formatting.add(rng, Rule(type='expression', dxf=dx, formula=['DateStyle="%s"' % k]))

def settings_note(ws, ref, text):
    cell(ws, ref, text, size=9, color=MUTED, wrap=True)

def rows(ws, r1, data, col=1):
    for i, row in enumerate(data, r1):
        for j, v in enumerate(row):
            if v is not None: ws.cell(i, col + j, v)

# ------------------------------------------------------------------ time tracker
def fl_time_tracker():
    wb = new_wb('fl-time-tracker-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [28, 24, 3, 60])
    cl = sheet(wb, 'Clients', [30, 11, 14, 14, 30, 30])
    lg = sheet(wb, 'Time log', [12, 26, 22, 34, 8, 8, 8, 9, 9, 10, 10, 9, 11, 12, 10, 6, 6], tab=TEAL)
    sm = sheet(wb, 'Summary', [36, 13, 13, 13, 16, 16, 16], tab=TEAL)
    ts = sheet(wb, 'Timesheet', [12, 22, 32, 8, 8, 9, 9, 10, 13, 6], tab=TEAL)
    fl_lists(wb)
    date_style(wb, 'Settings!$B$5')
    # Settings
    title(st, 'Settings', 'Fill once; change the period whenever you report or invoice.')
    for i, (k, v, fmt) in enumerate([('Your name or business', 'Alex Morgan Studio', None), ('Home currency', 'USD', None), ('Date style', 'Computer setting', None),
                                     ('Report period: first day', MONTH1, SYSDATE), ('Report period: last day', '=EOMONTH(B6,0)', SYSDATE)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    dv(st, 'B4', lst('A', len(CURS))); dv(st, 'B5', lst('B', len(DSTYLES)))
    dates(st, 'B6:B7')
    settings_note(st, 'D3', 'Home currency is the one you are paid in at home. Each client can still be billed in their own currency (Clients sheet).')
    settings_note(st, 'D5', '"Computer setting" shows dates the way your computer\'s regional settings do. Pick a fixed style to show every date the same way.')
    settings_note(st, 'D6', 'The Summary and Timesheet sheets show only the sessions between these two dates. The last day starts as the month end of the first day.')
    for r in (3, 5, 6): st.row_dimensions[r].height = 32
    # Clients
    title(cl, 'Clients', 'Rate per hour in the client\'s currency. Billing increment: minutes each session is rounded up to (0 = exact).')
    head(cl, 3, ['Client', 'Currency', 'Hourly rate', 'Billing increment (min)', 'Contact email', 'Notes'])
    grid(cl, 4, 103, 'ABCDEF', {}, {'C': NUM})
    dv(cl, 'B4:B103', lst('A', len(CURS)))
    rows(cl, 4, [('Northwind Labs (US)', 'USD', 65, 15, 'ap@northwind.example', 'Monthly invoice, Net 30'),
                 ('Brightside Media (UK)', 'GBP', 48, 6, 'finance@brightside.example', 'Timesheet with every invoice'),
                 ('Kaufmann GmbH (DE)', 'EUR', 55, 15, 'buchhaltung@kaufmann.example', 'Reverse charge VAT note'),
                 ('Desert Rose Events (UAE)', 'AED', 220, 30, 'accounts@desertrose.example', '')])
    cl.freeze_panes = 'A4'
    # Time log
    title(lg, 'Time log', 'One row per work session. Grey columns calculate themselves. Billable: leave Yes, or pick No for unpaid time.')
    head(lg, 3, ['Date', 'Client', 'Project', 'Task / what you did', 'Start', 'End', 'Break (min)', 'Hours worked', 'Billable',
                 'Rate override', 'Billed hours', 'Currency', 'Rate', 'Amount', 'Invoiced', 'In period', 'Row'])
    look = lambda col, r: 'IFERROR(INDEX(Clients!$%s$4:$%s$103,MATCH($B%d,Clients!$A$4:$A$103,0)),"")' % (col, col, r)
    f = {'H': '=IF(OR($E{r}="",$F{r}=""),"",MAX(0,ROUND((MOD($F{r}-$E{r},1)*1440-N($G{r}))/60,2)))',
         'K': '=IF($H{r}="","",IF($I{r}="No",0,IF(N(%s)>0,CEILING(MAX(0,ROUND(MOD($F{r}-$E{r},1)*1440-N($G{r}),0)),N(%s)),MAX(0,ROUND(MOD($F{r}-$E{r},1)*1440-N($G{r}),0)))/60))',
         'L': '=IF($B{r}="","",%s&"")',
         'M': '=IF($B{r}="","",IF($J{r}<>"",N($J{r}),N(%s)))',
         'N': '=IF(OR($K{r}="",$M{r}=""),"",ROUND($K{r}*$M{r},2))',
         'P': '=IF(AND(ISNUMBER($A{r}),$A{r}>=Settings!$B$6,$A{r}<=Settings!$B$7),1,0)',
         'Q': '=IF(AND($P{r}=1,$B{r}<>"",$B{r}=Timesheet!$B$4),ROW(),"")'}
    for r in range(4, 1004):
        for col, kind, fmt in (('A', 'in', SYSDATE), ('B', 'in', None), ('C', 'in', None), ('D', 'in', None), ('E', 'in', TIME), ('F', 'in', TIME), ('G', 'in', '0'),
                               ('I', 'in', None), ('J', 'in', NUM), ('O', 'in', None)):
            cell(lg, '%s%d' % (col, r), None, kind, fmt)
        cell(lg, 'H%d' % r, f['H'].format(r=r), 'calc', HRS)
        cell(lg, 'K%d' % r, f['K'].replace('%s', look('D', r)).format(r=r), 'calc', HRS)
        cell(lg, 'L%d' % r, f['L'].replace('%s', look('B', r)).format(r=r), 'calc')
        cell(lg, 'M%d' % r, f['M'].replace('%s', look('C', r)).format(r=r), 'calc', NUM)
        cell(lg, 'N%d' % r, f['N'].format(r=r), 'calc', NUM)
        cell(lg, 'P%d' % r, f['P'].format(r=r), 'calc')
        cell(lg, 'Q%d' % r, f['Q'].format(r=r), 'calc')
    lg.column_dimensions['P'].hidden = lg.column_dimensions['Q'].hidden = True
    dv(lg, 'B4:B1003', '=Clients!$A$4:$A$103', 'Pick a client from the Clients sheet')
    dv(lg, 'I4:I1003', lst('C', 2)); dv(lg, 'O4:O1003', lst('C', 2))
    t = datetime.time
    rows(lg, 4, [(M(0), 'Northwind Labs (US)', 'Mobile app UI', 'Onboarding screens, first draft', t(9, 0), t(12, 40), 10, None, 'Yes', None, None, None, None, None, 'Yes'),
                 (M(0), 'Brightside Media (UK)', 'Blog content', 'Article: remote team rituals', t(14, 0), t(16, 7), 0, None, 'Yes', None, None, None, None, None, 'No'),
                 (M(1), 'Kaufmann GmbH (DE)', 'Website translation', 'Product pages DE to EN', t(8, 30), t(13, 0), 30, None, 'Yes', None, None, None, None, None, 'No'),
                 (M(1), 'Northwind Labs (US)', 'Mobile app UI', 'Client call and revisions', t(18, 0), t(19, 5), 0, None, 'Yes', None, None, None, None, None, 'No'),
                 (M(2), 'Desert Rose Events (UAE)', 'Event brochure', 'Layout and image search', t(10, 15), t(13, 45), 15, None, 'Yes', None, None, None, None, None, 'No'),
                 (M(2), 'Brightside Media (UK)', 'Blog content', 'Pitch for next quarter (unpaid)', t(16, 0), t(16, 50), 0, None, 'No', None, None, None, None, None, 'No'),
                 (M(3), 'Northwind Labs (US)', 'Mobile app UI', 'Icon set, rush job', t(22, 30), t(1, 15), 0, None, 'Yes', 80, None, None, None, None, 'No'),
                 (M(3), 'Kaufmann GmbH (DE)', 'Website translation', 'Terms and FAQ pages', t(9, 0), t(11, 20), 0, None, 'Yes', None, None, None, None, None, 'No')])
    lg.freeze_panes = 'C4'
    lg.auto_filter.ref = 'A3:O1003'
    lg.conditional_formatting.add('A4:O1003', FormulaRule(formula=['$I4="No"'], font=Font(color='94A3B8')))
    lg.conditional_formatting.add('O4:O1003', FormulaRule(formula=['AND($N4<>"",N($N4)>0,$O4<>"Yes")'], fill=PatternFill('solid', fgColor='FDE68A')))
    dates(lg, 'A4:A1003')
    # Summary
    ws = sm
    title(ws, 'Hours and earnings by client')
    ws['A2'] = '=Settings!B3'; ws['A2'].font = Font(size=10, color=MUTED)
    label(ws, 'A3', 'Period (change it on the Settings sheet)'); cell(ws, 'B3', '=Settings!B6', 'calc', SYSDATE, align='left'); cell(ws, 'C3', '=Settings!B7', 'calc', SYSDATE, align='left')
    dates(ws, 'B3:C3')
    label(ws, 'A4', 'Total hours worked (all clients)'); cell(ws, 'C4', '=SUM(C6:C105)', 'soft', HRS, bold=True)
    head(ws, 5, ['Client', 'Currency', 'Hours worked', 'Billable hours', 'Amount earned', 'Not yet invoiced', 'Earned per hour worked'])
    for i in range(100):
        r, s = 6 + i, 4 + i
        S = lambda col, extra='': 'SUMIFS(\'Time log\'!$%s$4:$%s$1003,\'Time log\'!$B$4:$B$1003,$A%d,\'Time log\'!$P$4:$P$1003,1%s)' % (col, col, r, extra)
        cell(ws, 'A%d' % r, '=IF(Clients!A%d="","",Clients!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Clients!B%d&"")' % (r, s), 'box', align='center')
        cell(ws, 'C%d' % r, '=IF($A%d="","",%s)' % (r, S('H')), 'calc', HRS)
        cell(ws, 'D%d' % r, '=IF($A%d="","",%s)' % (r, S('K')), 'calc', HRS)
        cell(ws, 'E%d' % r, '=IF($A%d="","",%s)' % (r, S('N')), 'calc', NUM)
        cell(ws, 'F%d' % r, '=IF($A%d="","",%s)' % (r, S('N', ',\'Time log\'!$O$4:$O$1003,"<>Yes"')), 'calc', NUM)
        cell(ws, 'G%d' % r, '=IF(OR($A{r}="",N($C{r})=0),"",ROUND($E{r}/$C{r},2))'.format(r=r), 'calc', NUM)
    ws.conditional_formatting.add('F6:F105', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='FDE68A'), font=Font(bold=True)))
    cell(ws, 'A107', 'Amounts are in each client\'s own currency and are not added together. "Not yet invoiced" = billable work in the period without Yes in the Invoiced column.', size=9, color=MUTED)
    ws.freeze_panes = 'B6'
    protect(ws)
    # Timesheet (printable)
    ws = ts
    ws.merge_cells('A1:I1'); cell(ws, 'A1', '=Settings!B3', bold=True, size=16, color=TEAL, align='center')
    ws.merge_cells('A2:I2'); c = cell(ws, 'A2', 'TIMESHEET', bold=True, size=12, color='FFFFFF', align='center'); c.fill = PatternFill('solid', fgColor=TEAL)
    label(ws, 'A4', 'Client'); ws.merge_cells('B4:D4'); cell(ws, 'B4', 'Northwind Labs (US)', 'in')
    dv(ws, 'B4', '=Clients!$A$4:$A$103', 'Pick the client for this timesheet')
    label(ws, 'A5', 'Period'); ws.merge_cells('B5:D5'); cell(ws, 'B5', '=Settings!B6', 'calc', SYSDATE, align='left')
    label(ws, 'E5', 'to'); ws.merge_cells('F5:G5'); cell(ws, 'F5', '=Settings!B7', 'calc', SYSDATE, align='left')
    label(ws, 'E4', 'Currency'); ws.merge_cells('F4:G4'); cell(ws, 'F4', '=IFERROR(INDEX(Clients!$B$4:$B$103,MATCH($B$4,Clients!$A$4:$A$103,0))&"","")', 'calc', align='left')
    head(ws, 7, ['Date', 'Project', 'Task', 'Start', 'End', 'Hours worked', 'Billed hours', 'Rate', 'Amount'])
    src = {'A': ('A', SYSDATE), 'B': ('C', None), 'C': ('D', None), 'D': ('E', TIME), 'E': ('F', TIME), 'F': ('H', HRS), 'G': ('K', HRS), 'H': ('M', NUM), 'I': ('N', NUM)}
    for n in range(40):
        r = 8 + n
        cell(ws, 'J%d' % r, '=IFERROR(SMALL(\'Time log\'!$Q$4:$Q$1003,%d),"")' % (n + 1), size=8, color='94A3B8')
        for col, (sc, fmt) in src.items():
            cell(ws, '%s%d' % (col, r), '=IF($J{r}="","",INDEX(\'Time log\'!${c}:${c},$J{r}))'.format(r=r, c=sc), 'box', fmt, wrap=(col == 'C'))
    ws.column_dimensions['J'].hidden = True
    cell(ws, 'E48', 'Total', 'soft', bold=True)
    for col, fmt in (('F', HRS), ('G', HRS), ('I', NUM)): cell(ws, '%s48' % col, '=SUM(%s8:%s47)' % (col, col), 'soft', fmt, bold=True)
    ws.merge_cells('A49:I49')
    cell(ws, 'A49', '=IF(COUNT(\'Time log\'!$Q$4:$Q$1003)>40,"More than 40 sessions in this period: shorten the period on the Settings sheet and print in two parts.","")', bold=True, color='B91C1C', size=9)
    ws.merge_cells('A51:D51'); cell(ws, 'A51', 'Prepared by: ______________________', size=10)
    ws.merge_cells('E51:I51'); cell(ws, 'E51', 'Approved by client: ______________________', size=10)
    dates(ws, 'A8:A47'); dates(ws, 'B5'); dates(ws, 'F5')
    ws.print_area = 'A1:I52'
    protect(ws)
    save(wb, 'Freelance-Time-Tracker.xlsx')

# ------------------------------------------------------------------ client & project tracker
STATUS = ['Lead', 'Proposal sent', 'In progress', 'On hold', 'Delivered', 'Paid', 'Cancelled']
PRICING = ['Fixed fee', 'Hourly', 'Retainer (monthly)', 'Milestones']
METHODS = ['Bank transfer', 'Card', 'Online payment link', 'Platform payout', 'Cheque', 'Cash', 'Other']

def fl_client_tracker():
    wb = new_wb('fl-client-project-tracker-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [28, 24, 3, 60])
    cl = sheet(wb, 'Clients', [30, 16, 11, 22, 30, 14])
    pj = sheet(wb, 'Projects', [28, 26, 9, 16, 13, 12, 12, 14, 12, 12, 13, 13, 13, 13, 12, 10, 17, 30], tab=TEAL)
    py = sheet(wb, 'Payments', [12, 28, 14, 20, 30, 26])
    ov = sheet(wb, 'Overview', [28, 9, 9, 14, 14, 14, 14, 14, 80], tab=TEAL)
    fl_lists(wb, {4: STATUS, 5: PRICING, 6: METHODS})
    date_style(wb, 'Settings!$B$4')
    title(st, 'Settings', 'Fill once.')
    for i, (k, v, fmt) in enumerate([('Your name or business', 'Alex Morgan Studio', None), ('Date style', 'Computer setting', None), ('Position as on', '=TODAY()', SYSDATE)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    dv(st, 'B4', lst('B', len(DSTYLES))); dates(st, 'B5')
    settings_note(st, 'D4', '"Computer setting" follows your computer\'s regional date style.')
    settings_note(st, 'D5', 'Leave =TODAY() to see today\'s position, or type a date to see the position on that day.')
    # Clients
    title(cl, 'Clients', 'Payment terms = days the client has to pay after the invoice date (0 = due on receipt).')
    head(cl, 3, ['Client', 'Country', 'Currency', 'Contact person', 'Email', 'Payment terms (days)'])
    grid(cl, 4, 203, 'ABCDEF', {})
    dv(cl, 'C4:C203', lst('A', len(CURS)))
    rows(cl, 4, [('Northwind Labs', 'United States', 'USD', 'Jordan Lee', 'ap@northwind.example', 30),
                 ('Brightside Media', 'United Kingdom', 'GBP', 'Priya Shah', 'finance@brightside.example', 14),
                 ('Kaufmann GmbH', 'Germany', 'EUR', 'Lena Kaufmann', 'buchhaltung@kaufmann.example', 30),
                 ('Desert Rose Events', 'United Arab Emirates', 'AED', 'Omar Haddad', 'accounts@desertrose.example', 15),
                 ('Maple & Pine Co.', 'Canada', 'CAD', 'Chris Martin', 'billing@maplepine.example', 0)])
    cl.freeze_panes = 'A4'
    # Projects
    title(pj, 'Projects', 'Project names must be unique: payments are matched by name. Grey columns calculate themselves.')
    head(pj, 3, ['Project', 'Client', 'Currency', 'Pricing', 'Agreed fee', 'Start', 'Deadline', 'Status', 'Invoice no.', 'Invoice date',
                 'Invoiced amount', 'Received', 'Balance due', 'Not yet invoiced', 'Payment due', 'Days overdue', 'Alert', 'Notes'])
    asof = 'Settings!$B$5'
    lk = lambda col: 'IFERROR(INDEX(Clients!$%s$4:$%s$203,MATCH($B{r},Clients!$A$4:$A$203,0)),"")' % (col, col)
    f = {'C': '=IF($B{r}="","",%s&"")' % lk('C'),
         'L': '=IF($A{r}="","",SUMIF(Payments!$B$4:$B$1003,$A{r},Payments!$C$4:$C$1003))',
         'M': '=IF(OR($A{r}="",$K{r}=""),"",ROUND(N($K{r})-N($L{r}),2))',
         'N': '=IF(OR($A{r}="",$E{r}="",$H{r}="Cancelled"),"",MAX(0,ROUND(N($E{r})-N($K{r}),2)))',
         'O': '=IF(OR($J{r}="",$B{r}=""),"",$J{r}+N(%s))' % lk('F'),
         'P': '=IF(OR($O{r}="",N($M{r})<=0),"",MAX(0,%s-$O{r}))' % asof,
         'Q': ('=IF($A{r}="","",IF(N($P{r})>0,"Payment overdue",IF(AND(ISNUMBER($G{r}),$G{r}<%s,OR($H{r}="Lead",$H{r}="Proposal sent",$H{r}="In progress",$H{r}="On hold")),'
               '"Deadline passed",IF(AND($K{r}<>"",N($M{r})<=0),IF(N($N{r})>0,"Invoices paid","Paid in full"),""))))') % asof}
    grid(pj, 4, 503, 'ABCDEFGHIJKLMNOPQR', {c: 'calc' for c in 'CLMNOPQ'},
         {'E': NUM, 'F': SYSDATE, 'G': SYSDATE, 'J': SYSDATE, 'K': NUM, 'L': NUM, 'M': NUM, 'N': NUM, 'O': SYSDATE}, f)
    dv(pj, 'B4:B503', '=Clients!$A$4:$A$203', 'Pick a client from the Clients sheet')
    dv(pj, 'D4:D503', lst('E', len(PRICING))); dv(pj, 'H4:H503', lst('D', len(STATUS)))
    rows(pj, 4, [('Mobile app UI redesign', 'Northwind Labs', None, 'Fixed fee', 6400, D(-60), D(-5), 'Delivered', 'AM-2026-031', D(-45), 3200),
                 ('Mobile app icon set', 'Northwind Labs', None, 'Fixed fee', 900, D(-20), D(10), 'In progress'),
                 ('Blog content Q3', 'Brightside Media', None, 'Retainer (monthly)', 1800, D(-90), None, 'In progress', 'AM-2026-034', D(-25), 600),
                 ('Website translation DE-EN', 'Kaufmann GmbH', None, 'Hourly', 2750, D(-75), D(-45), 'Delivered', 'AM-2026-036', D(-42), 2475),
                 ('Event brochure', 'Desert Rose Events', None, 'Milestones', 8800, D(-60), D(-12), 'Paid', 'AM-2026-029', D(-50), 8800),
                 ('Packaging concept', 'Maple & Pine Co.', None, 'Fixed fee', 3500, None, None, 'Proposal sent'),
                 ('Brand guidelines', 'Brightside Media', None, 'Fixed fee', 2200, None, None, 'Cancelled')])
    pj.freeze_panes = 'B4'
    pj.auto_filter.ref = 'A3:R503'
    pj.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Payment overdue"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    pj.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Deadline passed"'], fill=PatternFill('solid', fgColor='FDE68A'), font=Font(bold=True, color='92400E')))
    pj.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Paid in full"'], fill=PatternFill('solid', fgColor='BBF7D0')))
    pj.conditional_formatting.add('Q4:Q503', CellIsRule(operator='equal', formula=['"Invoices paid"'], fill=PatternFill('solid', fgColor='DCFCE7')))
    pj.conditional_formatting.add('A4:R503', FormulaRule(formula=['$H4="Cancelled"'], font=Font(color='94A3B8')))
    dates(pj, 'F4:G503'); dates(pj, 'J4:J503'); dates(pj, 'O4:O503')
    # Payments
    title(py, 'Payments received', 'One row per payment, including deposits and part payments. Type the amount that reached you.')
    head(py, 3, ['Date', 'Project', 'Amount received', 'Method', 'Reference / note', 'Client'])
    grid(py, 4, 1003, 'ABCDEF', {'F': 'calc'}, {'A': SYSDATE, 'C': NUM},
         {'F': '=IF($B{r}="","",IFERROR(INDEX(Projects!$B$4:$B$503,MATCH($B{r},Projects!$A$4:$A$503,0))&"","Project not found"))'})
    dv(py, 'B4:B1003', '=Projects!$A$4:$A$503', 'Pick the project this payment is for')
    dv(py, 'D4:D1003', lst('F', len(METHODS)))
    rows(py, 4, [(D(-58), 'Mobile app UI redesign', 3200, 'Bank transfer', '50% deposit'),
                 (D(-48), 'Event brochure', 4400, 'Bank transfer', 'Milestone 1'),
                 (D(-30), 'Event brochure', 4400, 'Bank transfer', 'Milestone 2'),
                 (D(-12), 'Blog content Q3', 600, 'Online payment link', 'Invoice AM-2026-034'),
                 (D(-4), 'Website translation DE-EN', 1200, 'Bank transfer', 'Part payment')])
    py.freeze_panes = 'A4'
    dates(py, 'A4:A1003')
    py.conditional_formatting.add('F4:F1003', CellIsRule(operator='equal', formula=['"Project not found"'], font=Font(bold=True, color='B91C1C')))
    # Overview
    ws = ov
    title(ws, 'Clients overview')
    ws['A2'] = '=Settings!B3&"  ·  amounts in each client\'s currency"'; ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Client', 'Currency', 'Active projects', 'Agreed (not cancelled)', 'Invoiced', 'Received', 'Outstanding', 'Overdue', 'Reminder for overdue payments (copy into an email)'])
    P = lambda col, extra='': "SUMIFS(Projects!$%s$4:$%s$503,Projects!$B$4:$B$503,$A{r}%s)" % (col, col, extra)
    for i in range(200):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Clients!A%d="","",Clients!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Clients!C%d&"")' % (r, s), 'box', align='center')
        act = '+'.join('COUNTIFS(Projects!$B$4:$B$503,$A{r},Projects!$H$4:$H$503,"%s")' % k for k in ('In progress', 'On hold', 'Delivered'))
        cell(ws, 'C%d' % r, ('=IF($A{r}="","",%s)' % act).format(r=r), 'calc', align='center')
        cell(ws, 'D%d' % r, ('=IF($A{r}="","",%s)' % P('E', ',Projects!$H$4:$H$503,"<>Cancelled"')).format(r=r), 'calc', NUM)
        cell(ws, 'E%d' % r, ('=IF($A{r}="","",%s)' % P('K')).format(r=r), 'calc', NUM)
        cell(ws, 'F%d' % r, ('=IF($A{r}="","",%s)' % P('L')).format(r=r), 'calc', NUM)
        cell(ws, 'G%d' % r, ('=IF($A{r}="","",%s)' % P('M')).format(r=r), 'calc', NUM)
        cell(ws, 'H%d' % r, ('=IF($A{r}="","",%s)' % P('M', ',Projects!$P$4:$P$503,">0"')).format(r=r), 'calc', NUM)
        cell(ws, 'I%d' % r, ('=IF(OR($A{r}="",N($H{r})<=0),"","Dear "&IF(Clients!$D${s}<>"",Clients!$D${s},$A{r})&", I hope you are well. According to my records, "&$B{r}&" "&TEXT($H{r},"#,##0.00")'
                             '&" is now overdue on my invoice(s) to "&$A{r}&". Could you please let me know when the payment will be made? Thank you, "&Settings!$B$3&".")').format(r=r, s=s), size=9)
    ws.conditional_formatting.add('H5:H204', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'B5'
    # status counts + chart, to the right
    head(ws, 4, ['Status', 'Projects'], col=11)
    ws.column_dimensions['K'].width, ws.column_dimensions['L'].width = 16, 10
    for i, s in enumerate(STATUS, 5):
        cell(ws, 'K%d' % i, s, 'box'); cell(ws, 'L%d' % i, '=COUNTIF(Projects!$H$4:$H$503,K%d)' % i, 'calc', align='center')
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Projects by status'; ch.legend = None
    ch.add_data(Reference(ws, min_col=12, min_row=4, max_row=4 + len(STATUS)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=11, min_row=5, max_row=4 + len(STATUS)))
    ch.series[0].graphicalProperties.solidFill = TEAL
    ch.height, ch.width = 7, 12
    ws.add_chart(ch, 'K14')
    protect(ws)
    save(wb, 'Freelance-Client-Project-Tracker.xlsx')

# ------------------------------------------------------------------ income & expense with tax set-aside
CATS = ['Software & subscriptions', 'Computer & equipment', 'Internet & phone', 'Home office / rent share', 'Coworking space', 'Travel',
        'Meals with clients', 'Training, courses & books', 'Marketing & website', 'Bank & payment fees', 'Accountant & legal fees', 'Insurance',
        'Subcontractors', 'Office supplies', 'Other']

def fl_income_expense():
    wb = new_wb('fl-income-expense-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 20, 3, 30, 3, 50])
    inc = sheet(wb, 'Income', [12, 24, 13, 28, 9, 13, 11, 14, 14, 14, 10], tab=TEAL)
    ex = sheet(wb, 'Expenses', [12, 26, 22, 30, 13, 11, 14, 10], tab=TEAL)
    sm = sheet(wb, 'Summary', [14, 15, 15, 15, 16, 15, 16, 16], tab=TEAL)
    fl_lists(wb)
    date_style(wb, 'Settings!$B$7')
    title(st, 'Settings', 'Fill once a year.')
    for i, (k, v, fmt) in enumerate([('Your name or business', 'Alex Morgan Studio', None), ('Home currency', 'USD', None),
                                     ('First day of your tax year', datetime.date(TODAY.year, 1, 1), SYSDATE), ('Keep aside for income tax', 0.25, '0%'),
                                     ('Date style', 'Computer setting', None)], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    dv(st, 'B4', lst('A', len(CURS))); dv(st, 'B7', lst('B', len(DSTYLES))); dates(st, 'B5')
    head(st, 2, ['Expense categories (rename freely)'], col=4)
    for i in range(20): cell(st, 'D%d' % (3 + i), CATS[i] if i < len(CATS) else None, 'in')
    for r, t_ in ((3, 'Home currency = the currency your bank account and tax return use.'),
                  (5, 'For example 1 January, 1 April (India) or 6 April (UK). The Summary shows 12 months from this date.'),
                  (6, 'The share of each month\'s profit to move to a tax savings account. Ask a local tax adviser what suits your income.'),
                  (7, '"Computer setting" follows your computer\'s regional date style.')):
        settings_note(st, 'F%d' % r, t_); st.row_dimensions[r].height = 30
    # Income
    title(inc, 'Income received', 'One row per payment received. Foreign currency: type the rate you actually got (home currency for 1 unit).')
    hc = 'Settings!$B$4'
    head(inc, 3, ['Date received', 'Client', 'Invoice no.', 'Description', 'Currency', 'Amount', 'Exchange rate', 'Amount (home)', 'Sales tax included (home)', 'Net income (home)', 'Month'])
    for col, k in (('H', 'Amount'), ('I', 'Sales tax included'), ('J', 'Net income')):
        inc['%s3' % col] = '="%s ("&%s&")"' % (k, hc)
    grid(inc, 4, 1003, 'ABCDEFGHIJK', {c: 'calc' for c in 'HJK'}, {'A': SYSDATE, 'F': NUM, 'G': '0.0000', 'H': NUM, 'I': NUM, 'J': NUM, 'K': 'mmm yyyy'},
         {'H': '=IF($F{r}="","",IF(OR($E{r}="",$E{r}=%s),$F{r},IF(N($G{r})=0,"",ROUND($F{r}*$G{r},2))))' % hc,
          'J': '=IF($H{r}="","",ROUND($H{r}-N($I{r}),2))', 'K': '=IF(ISNUMBER($A{r}),DATE(YEAR($A{r}),MONTH($A{r}),1),"")'})
    dv(inc, 'E4:E1003', lst('A', len(CURS)))
    md = lambda m_, d_: datetime.date(TODAY.year, m_, d_)
    rows(inc, 4, [(md(1, 12), 'Northwind Labs', 'AM-2026-001', 'Website audit', 'USD', 2400, None, None, 0),
                  (md(1, 28), 'Brightside Media', 'AM-2026-003', 'Blog content January', 'GBP', 600, 1.2650, None, 0),
                  (md(2, 9), 'Kaufmann GmbH', 'AM-2026-005', 'Translation, product pages', 'EUR', 1450, 1.0820, None, 0),
                  (md(2, 25), 'Maple & Pine Co.', 'AM-2026-007', 'Logo refresh', 'USD', 1800, None, None, 0),
                  (md(3, 14), 'Desert Rose Events', 'AM-2026-009', 'Event brochure, milestone 1', 'AED', 4400, 0.2723, None, 0),
                  (md(3, 30), 'Local client (with sales tax)', 'AM-2026-010', 'Workshop', 'USD', 1080, None, None, 80)])
    inc.freeze_panes = 'B4'
    inc.conditional_formatting.add('G4:G1003', FormulaRule(formula=['AND($F4<>"",$E4<>"",$E4<>%s,N($G4)=0)' % hc], fill=PatternFill('solid', fgColor='FECACA')))
    dates(inc, 'A4:A1003')
    # Expenses
    title(ex, 'Business expenses', 'Amounts in your home currency. Business use: leave empty for 100%, or type the work share, e.g. 60%.')
    head(ex, 3, ['Date', 'Category', 'Supplier', 'Description', 'Amount', 'Business use', 'Allowable amount', 'Month'])
    for col, k in (('E', 'Amount'), ('G', 'Allowable amount')): ex['%s3' % col] = '="%s ("&%s&")"' % (k, hc)
    grid(ex, 4, 1003, 'ABCDEFGH', {c: 'calc' for c in 'GH'}, {'A': SYSDATE, 'E': NUM, 'F': '0%', 'G': NUM, 'H': 'mmm yyyy'},
         {'G': '=IF($E{r}="","",ROUND($E{r}*IF($F{r}="",1,N($F{r})),2))', 'H': '=IF(ISNUMBER($A{r}),DATE(YEAR($A{r}),MONTH($A{r}),1),"")'})
    dv(ex, 'B4:B1003', '=Settings!$D$3:$D$22', 'Pick a category (rename them on the Settings sheet)')
    rows(ex, 4, [(md(1, 3), 'Software & subscriptions', 'Design software', 'Annual plan', 660, None),
                 (md(1, 15), 'Internet & phone', 'Mobile operator', 'Phone bill January', 50, 0.6),
                 (md(2, 2), 'Computer & equipment', 'Electronics store', 'External monitor', 280, None),
                 (md(2, 20), 'Bank & payment fees', 'Bank', 'Incoming transfer fees', 18, None),
                 (md(3, 5), 'Coworking space', 'City coworking', 'Hot desk, March', 190, None),
                 (md(3, 22), 'Training, courses & books', 'Online course', 'Advanced motion design', 120, None)])
    ex.freeze_panes = 'B4'
    dates(ex, 'A4:A1003')
    # Summary
    ws = sm
    title(ws, 'Monthly summary for the tax year')
    ws['A2'] = '=Settings!B3&"  ·  amounts in "&Settings!B4&"  ·  keeping "&TEXT(Settings!B6,"0%")&" of profit aside for tax"'; ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Month', 'Net income', 'Expenses', 'Profit', 'Keep aside for income tax', 'Sales tax collected', 'Total to keep aside', 'Profit so far this year'])
    for n in range(12):
        r = 5 + n
        cell(ws, 'A%d' % r, '=EDATE(DATE(YEAR(Settings!$B$5),MONTH(Settings!$B$5),1),%d)' % n, 'box', 'mmm yyyy', bold=True)
        cell(ws, 'B%d' % r, '=SUMIFS(Income!$J$4:$J$1003,Income!$K$4:$K$1003,$A%d)' % r, 'calc', NUM)
        cell(ws, 'C%d' % r, '=SUMIFS(Expenses!$G$4:$G$1003,Expenses!$H$4:$H$1003,$A%d)' % r, 'calc', NUM)
        cell(ws, 'D%d' % r, '=B{r}-C{r}'.format(r=r), 'calc', NUM)
        cell(ws, 'E%d' % r, '=ROUND(MAX(0,D{r})*N(Settings!$B$6),2)'.format(r=r), 'calc', NUM)
        cell(ws, 'F%d' % r, '=SUMIFS(Income!$I$4:$I$1003,Income!$K$4:$K$1003,$A%d)' % r, 'calc', NUM)
        cell(ws, 'G%d' % r, '=E{r}+F{r}'.format(r=r), 'calc', NUM)
        cell(ws, 'H%d' % r, '=ROUND(SUM($D$5:D{r}),2)'.format(r=r), 'calc', NUM)
    cell(ws, 'A17', 'Year total', 'soft', bold=True)
    for col in 'BCDEFG': cell(ws, '%s17' % col, '=SUM(%s5:%s16)' % (col, col), 'soft', NUM, bold=True)
    ws.conditional_formatting.add('D5:D17', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('H5:H16', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    label(ws, 'A19', 'On the whole year')
    cell(ws, 'A20', 'Keep aside on the year\'s profit', 'box'); ws.merge_cells('A20:C20'); cell(ws, 'D20', '=ROUND(MAX(0,D17)*N(Settings!$B$6),2)', 'soft', NUM, bold=True)
    ws.merge_cells('E20:H21')
    cell(ws, 'E20', 'Monthly figures set aside nothing in a loss month, so the monthly total can be higher than the yearly figure. This is a savings guide, not a tax calculation.', size=9, color=MUTED, wrap=True)
    cell(ws, 'A21', 'Months with income recorded', 'box'); ws.merge_cells('A21:C21'); cell(ws, 'D21', '=COUNTIF(B5:B16,">0")', 'calc', align='center')
    label(ws, 'A23', 'Expenses by category (whole year)')
    head(ws, 24, ['Category', '', '', 'Allowable', 'Share'])
    ws.merge_cells('A24:C24')
    yr = 'Expenses!$A$4:$A$1003,">="&Settings!$B$5,Expenses!$A$4:$A$1003,"<"&EDATE(Settings!$B$5,12)'
    for i in range(20):
        r = 25 + i
        ws.merge_cells('A%d:C%d' % (r, r))
        cell(ws, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (3 + i, 3 + i), 'box')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIFS(Expenses!$G$4:$G$1003,Expenses!$B$4:$B$1003,$A{r},{y}))'.format(r=r, y=yr), 'calc', NUM)
        cell(ws, 'E%d' % r, '=IF(OR($A{r}="",N($C$17)=0),"",N(D{r})/$C$17)'.format(r=r), 'calc', '0%')
    ch = BarChart(); ch.type = 'col'; ch.style = 10; ch.title = 'Income and expenses by month'
    ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=4, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=16))
    ch.series[0].graphicalProperties.solidFill = TEAL; ch.series[1].graphicalProperties.solidFill = 'F59E0B'
    ch.height, ch.width = 8, 18
    ws.add_chart(ch, 'J4')
    protect(ws)
    save(wb, 'Freelance-Income-Expense-Tracker.xlsx')

# ------------------------------------------------------------------ Word: proposal and agreement
def h2(d, text): return para(d, text, bold=True, size=13, color=TEAL, space=4)

def fl_proposal():
    d = doc_new('fl-proposal-quote-word')
    letterhead(d, '[Your Name / Agency]', '[City, Country] · [email] · [phone] · [website or portfolio]', gst=False)
    para(d, 'PROJECT PROPOSAL & QUOTE', bold=True, size=16, align='c', space=2)
    para(d, '[Project name]', bold=True, size=13, align='c', color=TEAL, space=8)
    kv(d, [('Prepared for', '[Client name, company]'), ('Proposal no.', 'P-[2026-001]'), ('Contact', '[Name, email]'), ('Date', '[date]'),
           ('Currency', '[USD / EUR / GBP / INR / AED / …]'), ('Valid until', '[date, e.g. 30 days from today]')])
    h2(d, '1. Your goals')
    para(d, '[Write the client\'s goals in their own words from your call or brief. For example: "Northwind wants new users to finish sign-up in under two minutes, '
            'because 40% leave on the second screen today."]', italic=True)
    for s in ['[Goal 1: what success looks like]', '[Goal 2]', '[Goal 3, including any deadline that matters to the client, e.g. a launch date]']:
        d.add_paragraph(s, style='List Bullet')
    h2(d, '2. Proposed approach')
    para(d, '[Two or three short paragraphs: how you will solve the problem, why this approach, and what you have done before that is similar. Link one or two relevant samples.]', italic=True, space=8)
    h2(d, '3. Scope and deliverables')
    table(d, ['#', 'Deliverable', 'What is included'],
          [['1', '[e.g. Discovery workshop]', '[one 90-minute online session, summary notes]'], ['2', '[e.g. Design of 8 screens]', '[two rounds of revisions, source files]'],
           ['3', '[e.g. Handover]', '[assets exported for developers, 30-minute walkthrough]']], [1, 6, 10], fill_rows=2)
    para(d, 'Out of scope (not included in this price)', bold=True)
    for s in ['[e.g. copywriting, stock photos or fonts that need a licence]', '[e.g. development or publishing]', '[e.g. more than two rounds of revisions]']:
        d.add_paragraph(s, style='List Bullet')
    para(d, 'Anything out of scope can be added later at [rate per hour / per day] or with a separate quote.', size=10, space=8)
    h2(d, '4. Timeline')
    table(d, ['Phase', 'Dates', 'Output', 'Needed from you'],
          [['[Kick-off]', '[week 1]', '[brief confirmed]', '[access, brand files]'], ['[Draft]', '[weeks 2–3]', '[first version]', '[feedback within 3 working days]'],
           ['[Final]', '[week 4]', '[final files]', '[sign-off]']], [3.5, 3, 5, 5.5], fill_rows=1)
    para(d, 'Dates move by the same number of days if feedback or materials arrive late.', size=10, space=8)
    h2(d, '5. Investment')
    table(d, ['Item', 'Qty', 'Unit', 'Rate', 'Amount'],
          [['[Deliverable or phase]', '[1]', '[project]', '[ ]', '[ ]'], ['[Deliverable or phase]', '[ ]', '[hours / days]', '[ ]', '[ ]'],
           ['Optional: [add-on, e.g. social media versions]', '[ ]', '[ ]', '[ ]', '[ ]']], [7, 1.6, 2.6, 2.6, 3.2], fill_rows=1)
    kv(d, [('Subtotal', '[currency] [ ]'), ('[VAT / GST / Sales tax] @ [ ]% or "Not applicable"', '[currency] [ ]'), ('Total', '[currency] [ ]'), ('Transfer fees', '[paid by the client / shared]')], cols=1)
    para(d, 'Payment schedule', bold=True)
    table(d, ['When', 'Share', 'Amount'], [['On signing (deposit, before work starts)', '[50]%', '[ ]'], ['[On delivery of first draft / milestone]', '[25]%', '[ ]'],
                                           ['On final delivery', '[25]%', '[ ]']], [10, 3, 4])
    para(d, 'Invoices are due within [14] days. Payment by bank transfer to the account on the invoice [or other method you accept].', size=10, space=8)
    h2(d, '6. Revisions and changes')
    para(d, 'This price includes [two] rounds of revisions per deliverable. A round means one set of combined feedback. Further rounds or new requests are charged at '
            '[rate] or quoted separately before any work starts.', space=8)
    h2(d, '7. Next steps')
    for s in ['Reply with any questions, or sign below to accept.', 'I send the service agreement and the deposit invoice.', 'Work starts on [date] once the deposit is received.']:
        d.add_paragraph(s, style='List Number')
    para(d, '', space=6)
    h2(d, 'Acceptance')
    para(d, 'I accept this proposal, its scope, price and payment schedule.', space=18)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = 'For the client', 'For [Your Name / Agency]'
    t.rows[1].cells[0].text = '\n\nSignature: ____________________\nName: ____________________\nDate: ____________'
    t.rows[1].cells[1].text = '\n\nSignature: ____________________\nName: ____________________\nDate: ____________'
    dsave(d, 'Freelance-Project-Proposal.docx')

def clause(d, num_title, texts):
    h2(d, num_title)
    for i, s in enumerate(texts, 1):
        p = para(d, space=4); r = p.add_run('%s.%d  ' % (num_title.split('.')[0], i)); r.bold = True; p.add_run(s)

def fl_agreement():
    d = doc_new('fl-service-agreement-word')
    para(d, 'FREELANCE SERVICE AGREEMENT', bold=True, size=18, align='c', color=TEAL, space=2)
    para(d, 'Starting draft: adapt it to your project and have it checked by a lawyer in the country whose law applies. Delete this line and every option in [brackets] that does not apply.',
         size=9, italic=True, align='c', color=MUTED, space=10)
    para(d, 'This agreement is made on [date] between:', space=6)
    kv(d, [('Freelancer', '[Full legal name or business name]'), ('Address', '[Address, country]'), ('Tax / registration no.', '[if any]'), ('Email', '[ ]')], cols=1)
    para(d, 'and', align='c', space=6)
    kv(d, [('Client', '[Company or person]'), ('Address', '[Address, country]'), ('Tax / VAT no.', '[if any]'), ('Contact person', '[Name, email]')], cols=1)
    para(d, 'together "the parties". The parties agree as follows.', space=8)
    clause(d, '1. Services', ['The Freelancer will provide the services and deliverables described in Schedule A [or: in Proposal no. [ ] dated [ ], attached] ("the Services").',
                              'The work starts on [date] and the planned completion date is [date]. Dates depend on the Client supplying materials, access and feedback on time.',
                              'Anything not described in Schedule A is outside this agreement until both parties agree it in writing (email is enough).'])
    clause(d, '2. Independent contractor', ['The Freelancer works as an independent business, not as an employee. The Freelancer decides how, when and where the work is done, uses their own equipment, and is responsible for their own taxes and social contributions.',
                                            'The Freelancer may use subcontractors for parts of the work, remains responsible for their work, and makes sure they keep the Client\'s information confidential.',
                                            'Nothing in this agreement stops the Freelancer from working for other clients.'])
    clause(d, '3. Fees and payment', ['The Client will pay [amount and currency, e.g. USD 6,400] for the Services [or: [currency] [rate] per [hour / day], estimated at [ ] hours].',
                                      'A deposit of [50]% is due on signing; work starts when it is received. The rest is invoiced [on delivery / at the milestones in Schedule A / monthly in arrears].',
                                      'Invoices are payable within [14 / 30] days of the invoice date by [bank transfer] to the account shown on the invoice. Bank and transfer charges are paid by [the Client / each party for its own bank].',
                                      'Prices [exclude / include] VAT, GST or sales tax. Where tax applies it is added at the rate in force on the invoice date [or: the Client accounts for tax under the reverse charge, where that applies].',
                                      'If an invoice is not paid on time, the Freelancer may pause work after giving [7] days\' written notice, and may charge late payment interest of [ ]% per month [or: as allowed by the law that applies], plus reasonable costs of recovering the debt.',
                                      'Expenses such as stock images, fonts, software licences or travel are charged only if the Client approves them in advance.'])
    clause(d, '4. Revisions and changes', ['The fee includes [two] rounds of revisions per deliverable. A round is one combined set of feedback.',
                                           'Extra revisions, new deliverables or changes to the agreed scope are charged at [currency] [rate] per [hour] or quoted before the work starts.',
                                           'If the Client does not give feedback within [10] working days of a delivery, the delivery is treated as approved [or: the Freelancer may invoice for the work done so far].'])
    clause(d, '5. Ownership of the work', ['When the Client has paid all invoices under this agreement in full, ownership of the final deliverables (including copyright, as far as the law allows) passes to the Client.',
                                           'Until full payment, the Freelancer keeps ownership and the Client may use the deliverables only for review.',
                                           'The Freelancer keeps ownership of drafts, rejected concepts, and their own pre-existing tools, code libraries, templates and know-how, and gives the Client a permanent licence to use any of them that are built into the final deliverables.',
                                           'Third-party materials (fonts, stock images, open-source software) are used under their own licences, which the Freelancer will name on request.',
                                           'The Freelancer may show the finished work in a portfolio and mention the Client by name [unless the Client asks otherwise in writing / after the public launch].'])
    clause(d, '6. Confidentiality and data', ['Each party will keep the other\'s confidential information secret and use it only for this project, during the agreement and for [two] years after it ends. This does not cover information that is already public or that the law requires to be disclosed.',
                                              'If the Freelancer handles personal data for the Client, the Freelancer will follow the Client\'s reasonable written instructions and the data protection law that applies, and will delete or return the data when the project ends.'])
    clause(d, '7. Responsibilities', ['The Freelancer will do the work with reasonable skill and care, and the deliverables will be the Freelancer\'s own original work except for third-party materials named under clause 5.',
                                      'The Client confirms that materials it supplies (text, images, logos, data) may be used for the project.',
                                      'Neither party is liable for indirect losses such as loss of profit or business. Each party\'s total liability under this agreement is limited to the total fees paid or payable under it [except where the law does not allow such a limit].'])
    clause(d, '8. Ending the agreement', ['Either party may end this agreement by giving [14] days\' written notice.',
                                          'Either party may end it at once if the other seriously breaks it and does not fix the problem within [7] days of being asked in writing.',
                                          'If the agreement ends early, the Client pays for all work done up to the end date and for approved expenses; the deposit is [not refundable / set against that amount]. Ownership of work that is paid for passes as in clause 5.'])
    clause(d, '9. General', ['This agreement, with its schedules, is the whole agreement between the parties about the Services. Changes must be in writing and agreed by both.',
                             'Neither party is responsible for delays caused by events outside its reasonable control, such as natural disasters, war or failure of public networks.',
                             'This agreement is governed by the laws of [country / state]. The parties will first try to settle any dispute by talking in good faith; if that fails within [30] days, the courts of [place] [or: mediation / arbitration at [ ]] will decide.',
                             'The agreement may be signed electronically and in separate copies.'])
    h2(d, 'Signatures')
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = 'Freelancer', 'Client'
    for c in t.rows[1].cells: c.text = '\n\nSignature: ____________________\nName: ____________________\nTitle: ____________________\nDate: ____________'
    para(d, '', space=10)
    h2(d, 'Schedule A: Services, deliverables and payment milestones')
    table(d, ['#', 'Deliverable / milestone', 'Due date', 'Fee', 'When invoiced'],
          [['1', '[e.g. Deposit]', '[on signing]', '[ ]', '[on signing]'], ['2', '[Deliverable]', '[date]', '[ ]', '[on delivery]'], ['3', '[Deliverable]', '[date]', '[ ]', '[on delivery]']],
          [1, 7, 3, 2.6, 3.4], fill_rows=2)
    para(d, 'This draft is a general template, not legal advice. Laws on contracts, late payment, independent contractors and copyright differ between countries.', size=8, italic=True, color=MUTED)
    dsave(d, 'Freelance-Service-Agreement.docx')

MAKERS = [fl_time_tracker, fl_client_tracker, fl_income_expense, fl_proposal, fl_agreement]
