"""Home & Personal Finance toolkit (kit id home-money): six workbooks, HM-*.xlsx.
   Global by design: the currency symbol is typed on each Settings sheet (shown in headings), amounts are plain numbers with
   thousand separators, and dates use Excel's built-in short date (format 14) so they show in the computer's own date style."""
from make_files import *
from openpyxl.chart import BarChart, Reference

NUM, NUM0, DT, MON, MMY = '#,##0.00', '#,##0', 'mm-dd-yy', 'mmmm yyyy', 'mmm-yyyy'
DASH = '#,##0;-#,##0;"–"'
PINK = 'DB2777'
RED, AMBER, GREEN = PatternFill('solid', fgColor='FECACA'), PatternFill('solid', fgColor='FDE68A'), PatternFill('solid', fgColor='D1FAE5')

def addm(d, k):
    """d moved k months (day kept, capped at 28 so every month has it)."""
    m = d.month - 1 + k
    return datetime.date(d.year + m // 12, m % 12 + 1, min(d.day, 28))

def settings(wb, rows, widths=(34, 24)):
    st = sheet(wb, 'Settings', list(widths))
    title(st, 'Settings', 'Yellow cells: type or pick. The other sheets read from here.')
    for i, (k, v, fmt) in enumerate(rows, 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    return st

def fill_rows(ws, r0, rows, col=1):
    for i, row in enumerate(rows, r0):
        for j, v in enumerate(row):
            if v is not None: ws.cell(i, col + j, v)

def month_head(ws, r, c0, n, start):
    """n month headers from the month in `start` (a cell reference), white on blue."""
    for k in range(n):
        cell(ws, '%s%d' % (L(c0 + k), r), '=EDATE(DATE(YEAR(%s),MONTH(%s),1),%d)' % (start, start, k), fmt='mmm-yy', bold=True, color='FFFFFF', align='center').fill = F_HEAD

def red_if(ws, rng, formula, fill=RED, bold=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill, font=Font(bold=True, color='991B1B') if bold else None))

def upcoming(ws, r0, n, key_rng, show, helper_col):
    """n rows from r0 listing the rows of another sheet in date order: key_rng holds date+ROW()/100000 (unique), show = [(col, src_rng, fmt)]."""
    hk, hi = L(helper_col), L(helper_col + 1)
    for k in range(1, n + 1):
        r = r0 + k - 1
        cell(ws, '%s%d' % (hk, r), '=IFERROR(SMALL(%s,%d),"")' % (key_rng, k))
        cell(ws, '%s%d' % (hi, r), '=IF(%s%d="","",MATCH(%s%d,%s,0))' % (hk, r, hk, r, key_rng))
        for col, src, fmt in show:
            f = ('=IF($%s%d="","",%d)' % (hi, r, k)) if src is None else ('=IF($%s%d="","",INDEX(%s,$%s%d))' % (hi, r, src, hi, r))
            cell(ws, '%s%d' % (col, r), f, 'box', fmt)
    ws.column_dimensions[hk].hidden = ws.column_dimensions[hi].hidden = True

# ------------------------------------------------------------------ 1. Monthly household budget
def hm_budget():
    wb = new_wb('hm-household-budget-excel')
    howto(wb, first=True)
    settings(wb, [('Household name', 'The Sharma family', None), ('Currency symbol', '₹', None), ('Month to view', MONTH1, MON),
                  ('Savings target (% of income)', 0.2, '0%'), ('Year starts (first month)', datetime.date(TODAY.year, 1, 1), MON)])
    bu = sheet(wb, 'Budget', [34, 14, 16], tab=PINK)
    inc = sheet(wb, 'Income', [13, 30, 18, 14, 32, 12])
    sp = sheet(wb, 'Spending', [13, 28, 34, 14, 16, 12, 12], tab=PINK)
    mo = sheet(wb, 'Month', [36, 16, 4, 14, 16, 16], tab=BLUE)
    yr = sheet(wb, 'Year', [30, 11] + [11] * 12 + [13], tab=BLUE)
    lists_sheet(wb, extra={3: ['Fixed', 'Variable', 'Savings'], 4: ['Cash', 'UPI', 'Debit card', 'Credit card', 'Bank transfer', 'Auto-debit', 'Cheque']})
    cur = 'Amounts in the currency set on the Settings sheet'
    # Budget
    title(bu, 'Budget plan', cur)
    head(bu, 3, ['Category', 'Type', 'Monthly plan'])
    grid(bu, 4, 53, 'ABC', {}, {'C': NUM0})
    cats = [('Rent / home loan EMI', 'Fixed', 18000), ('Electricity', 'Variable', 2200), ('Water and cooking gas', 'Variable', 1100), ('Mobile and internet', 'Fixed', 1500),
            ('Groceries', 'Variable', 9000), ('Milk, fruit and vegetables', 'Variable', 4500), ('School fees and tuition', 'Fixed', 7500), ('Fuel and transport', 'Variable', 4000),
            ('House help', 'Fixed', 11500), ('Insurance premiums (monthly share)', 'Fixed', 3500), ('Medicines and doctor', 'Variable', 1500), ('Eating out and entertainment', 'Variable', 3000),
            ('Clothes and personal care', 'Variable', 2500), ('Gifts, festivals and functions', 'Variable', 2000), ('Subscriptions and apps', 'Fixed', 900),
            ('Emergency fund', 'Savings', 5000), ('Recurring deposit / SIP', 'Savings', 15000), ('Other', 'Variable', 1500)]
    fill_rows(bu, 4, cats)
    dv(bu, 'B4:B53', '=Lists!$C$1:$C$3')
    cell(bu, 'A55', 'Total planned', 'soft', bold=True); cell(bu, 'C55', '=SUM(C4:C53)', 'soft', NUM0, bold=True)
    for i, k in enumerate(['Fixed', 'Variable', 'Savings'], 56):
        cell(bu, 'A%d' % i, k, 'box'); cell(bu, 'B%d' % i, '=IF($C$55=0,"",C%d/$C$55)' % i, 'calc', '0%'); cell(bu, 'C%d' % i, '=SUMIF($B$4:$B$53,A%d,$C$4:$C$53)' % i, 'calc', NUM0)
    bu.freeze_panes = 'A4'
    # Income
    title(inc, 'Income', cur)
    head(inc, 3, ['Date', 'Source', 'Earner', 'Amount', 'Note', 'Month'])
    grid(inc, 4, 203, 'ABCDEF', {'F': 'calc'}, {'A': DT, 'D': NUM0, 'F': MMY}, {'F': '=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'})
    p1 = addm(MONTH1, -1)
    fill_rows(inc, 4, [(p1, 'Salary', 'Rohit', 58000, 'Net, after deductions'), (p1.replace(day=2), 'Salary', 'Neha', 36000, ''),
                       (p1.replace(day=15), 'Tuition classes at home', 'Neha', 4500, 'Side income'),
                       (MONTH1, 'Salary', 'Rohit', 58000, 'Net, after deductions'), (MONTH1.replace(day=2), 'Salary', 'Neha', 36000, ''),
                       (MONTH1.replace(day=15), 'Tuition classes at home', 'Neha', 4500, 'Side income'), (MONTH1.replace(day=20), 'Savings account interest', 'Neha', 620, 'Quarterly credit')])
    inc.freeze_panes = 'A4'
    # Spending
    title(sp, 'Spending', cur)
    head(sp, 3, ['Date', 'Category', 'Description', 'Amount', 'Paid by', 'Month', 'Type'])
    grid(sp, 4, 1003, 'ABCDEFG', {'F': 'calc', 'G': 'calc'}, {'A': DT, 'D': NUM0, 'F': MMY},
         {'F': '=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))',
          'G': '=IF($B{r}="","",IFERROR(INDEX(Budget!$B$4:$B$53,MATCH($B{r},Budget!$A$4:$A$53,0)),"Unplanned"))'})
    dv(sp, 'B4:B1003', '=Budget!$A$4:$A$53', 'Pick a category from the Budget sheet')
    dv(sp, 'E4:E1003', '=Lists!$D$1:$D$7')
    D = lambda base, n: base.replace(day=n)
    rows = []
    for base, f in ((p1, 1.0), (MONTH1, 1.0)):
        rows += [(D(base, 1), 'Rent / home loan EMI', 'Home loan EMI', 18000, 'Auto-debit'), (D(base, 1), 'Recurring deposit / SIP', 'Monthly SIP', 15000, 'Auto-debit'),
                 (D(base, 2), 'House help', 'Cook and cleaner salary', 11500, 'UPI'), (D(base, 3), 'Groceries', 'Monthly ration', 6400, 'UPI'),
                 (D(base, 5), 'School fees and tuition', 'Maths tuition', 2500, 'UPI'), (D(base, 7), 'Mobile and internet', 'Broadband + two mobiles', 1497, 'Credit card'),
                 (D(base, 9), 'Fuel and transport', 'Petrol', 2200, 'Credit card'), (D(base, 10), 'Milk, fruit and vegetables', 'Vegetable vendor (10 days)', 1450, 'Cash'),
                 (D(base, 12), 'Electricity', 'Electricity bill', 2650, 'UPI'), (D(base, 14), 'Eating out and entertainment', 'Family dinner', 2350, 'Credit card'),
                 (D(base, 16), 'Medicines and doctor', 'Paediatrician visit + medicines', 1180, 'UPI'), (D(base, 18), 'Groceries', 'Top-up shopping', 1900, 'UPI'),
                 (D(base, 20), 'Insurance premiums (monthly share)', 'Set aside for health insurance renewal', 3500, 'Bank transfer'),
                 (D(base, 22), 'Eating out and entertainment', 'Movie and snacks', 1250, 'UPI'), (D(base, 24), 'Gifts, festivals and functions', 'Wedding gift', 3100, 'Cash'),
                 (D(base, 26), 'Emergency fund', 'Moved to emergency fund', 5000, 'Bank transfer'), (D(base, 27), 'Clothes and personal care', 'Salon', 900, 'UPI'),
                 (D(base, 28), 'Car service', 'Car service (not in plan)', 4200, 'Credit card')]
    fill_rows(sp, 4, rows)
    sp.freeze_panes = 'A4'
    # Month view
    ws = mo
    title(ws, 'Month view', cur)
    label(ws, 'A3', 'Month'); cell(ws, 'B3', '=DATE(YEAR(Settings!B5),MONTH(Settings!B5),1)', 'soft', MON, bold=True, align='left')
    SS = lambda crit: 'SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,$B$3,Spending!$G$4:$G$1003,"%s")' % crit
    kpi = [('Income', '=SUMIFS(Income!$D$4:$D$203,Income!$F$4:$F$203,$B$3)', NUM0), ('Fixed spending', '=' + SS('Fixed'), NUM0),
           ('Variable and unplanned spending', '=%s+%s' % (SS('Variable'), SS('Unplanned')), NUM0), ('Moved to savings', '=' + SS('Savings'), NUM0),
           ('Left over (income − all of the above)', '=B4-B5-B6-B7', NUM0), ('Savings rate (income not spent)', '=IF(N(B4)<=0,"",(B4-B5-B6)/B4)', '0.0%'),
           ('Your target', '=Settings!B6', '0%'),
           ('How the month is going', '=IF(B9="","Add this month\'s income to see the rate",IF(B9>=B10,"On target","Below target by "&ROUND((B10-B9)*100,1)&" points"))', None)]
    for i, (k, f, fmt) in enumerate(kpi, 4):
        label(ws, 'A%d' % i, k, bold=(i in (8, 9, 11))); cell(ws, 'B%d' % i, f, 'calc', fmt, bold=(i in (8, 9, 11)), align='right' if fmt else 'left')
    red_if(ws, 'B8', 'B8<0'); red_if(ws, 'B11', 'AND(B9<>"",B9<B10)', AMBER); ws.conditional_formatting.add('B11', FormulaRule(formula=['AND(B9<>"",B9>=B10)'], fill=GREEN))
    head(ws, 3, ['Type', 'Plan', 'Actual'], col=4)
    for i, k in enumerate(['Fixed', 'Variable', 'Savings', 'Unplanned'], 4):
        cell(ws, 'D%d' % i, k, 'box')
        cell(ws, 'E%d' % i, '=SUMIF(Budget!$B$4:$B$53,D%d,Budget!$C$4:$C$53)' % i, 'calc', NUM0)
        cell(ws, 'F%d' % i, '=SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,$B$3,Spending!$G$4:$G$1003,D%d)' % i, 'calc', NUM0)
    ch = BarChart(); ch.type = 'col'; ch.style = 10; ch.title = 'Plan vs actual'; ch.y_axis.title = None
    ch.add_data(Reference(ws, min_col=5, max_col=6, min_row=3, max_row=7), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=4, min_row=4, max_row=7))
    ch.series[0].graphicalProperties.solidFill = '94A3B8'; ch.series[1].graphicalProperties.solidFill = PINK
    ch.height, ch.width = 7.5, 13
    ws.add_chart(ch, 'H3')
    head(ws, 13, ['Category', 'Type', '', 'Plan', 'Actual', 'Left (plan − actual)'])
    ws.cell(13, 3).fill = PatternFill(None)
    for i in range(50):
        r, s = 14 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Budget!A%d="","",Budget!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Budget!B%d&"")' % (r, s), 'box')
        cell(ws, 'D%d' % r, '=IF($A%d="","",N(Budget!C%d))' % (r, s), 'calc', NUM0)
        cell(ws, 'E%d' % r, '=IF($A{r}="","",SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,$B$3,Spending!$B$4:$B$1003,$A{r}))'.format(r=r), 'calc', NUM0)
        cell(ws, 'F%d' % r, '=IF($A{r}="","",D{r}-E{r})'.format(r=r), 'calc', NUM0)
    red_if(ws, 'A14:F63', 'AND($A14<>"",$B14<>"Savings",$E14>$D14)')
    cell(ws, 'A64', 'Unplanned (category not on the Budget sheet)', 'box')
    cell(ws, 'E64', '=' + SS('Unplanned'), 'calc', NUM0)
    cell(ws, 'A65', 'Total', 'soft', bold=True)
    for col in 'DE': cell(ws, '%s65' % col, '=SUM(%s14:%s64)' % (col, col), 'soft', NUM0, bold=True)
    cell(ws, 'F65', '=D65-E65', 'soft', NUM0, bold=True)
    ws.freeze_panes = 'A14'
    protect(ws)
    # Year view
    ws = yr
    title(ws, 'Year view: spending by month', cur)
    head(ws, 3, ['Category', 'Type'] + [''] * 12 + ['Total'])
    month_head(ws, 3, 3, 12, 'Settings!$B$7')
    for i in range(50):
        r, s = 4 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Budget!A%d="","",Budget!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Budget!B%d&"")' % (r, s), 'box', size=9)
        for k in range(12):
            c = L(3 + k)
            cell(ws, '%s%d' % (c, r), '=IF($A{r}="","",SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,{c}$3,Spending!$B$4:$B$1003,$A{r}))'.format(r=r, c=c), 'calc', DASH)
        cell(ws, 'O%d' % r, '=IF($A{r}="","",SUM(C{r}:N{r}))'.format(r=r), 'calc', DASH)
    cell(ws, 'A54', 'Unplanned', 'box')
    lines = [(56, 'Total spending (not counting savings)', 'SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,{c}$3,Spending!$G$4:$G$1003,"<>Savings")'),
             (57, 'Moved to savings', 'SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,{c}$3,Spending!$G$4:$G$1003,"Savings")'),
             (58, 'Income', 'SUMIFS(Income!$D$4:$D$203,Income!$F$4:$F$203,{c}$3)')]
    for k in range(12):
        c = L(3 + k)
        cell(ws, '%s54' % c, '=SUMIFS(Spending!$D$4:$D$1003,Spending!$F$4:$F$1003,{c}$3,Spending!$G$4:$G$1003,"Unplanned")'.format(c=c), 'calc', DASH)
        for r, _, f in lines: cell(ws, '%s%d' % (c, r), '=' + f.format(c=c), 'soft', DASH, bold=True)
        cell(ws, '%s59' % c, '=IF(N({c}58)<=0,"",({c}58-{c}56)/{c}58)'.format(c=c), 'soft', '0%', bold=True)
    cell(ws, 'O54', '=SUM(C54:N54)', 'calc', DASH)
    for r, k, _ in lines:
        cell(ws, 'A%d' % r, k, 'soft', bold=True); cell(ws, 'O%d' % r, '=SUM(C%d:N%d)' % (r, r), 'soft', DASH, bold=True)
    cell(ws, 'A59', 'Savings rate', 'soft', bold=True); cell(ws, 'O59', '=IF(N(O58)<=0,"",(O58-O56)/O58)', 'soft', '0%', bold=True)
    ws.freeze_panes = 'C4'
    ws.page_setup.orientation = 'landscape'
    protect(ws)
    save(wb, 'HM-Household-Budget.xlsx')

# ------------------------------------------------------------------ 2. Loan & EMI tracker
def hm_loan():
    wb = new_wb('hm-loan-emi-excel')
    howto(wb, first=True)
    lo = sheet(wb, 'Loan', [40, 20, 4, 60], tab=PINK)
    sc = sheet(wb, 'Schedule', [6, 13, 16, 14, 14, 14, 14, 16, 16, 15], tab=BLUE)
    al = sheet(wb, 'All loans', [26, 20, 13, 15, 9, 13, 9, 13, 12, 30])
    lists_sheet(wb, extra={3: ['Home', 'Vehicle', 'Personal', 'Education', 'Gold', 'Credit card', 'Business', 'Other']})
    first = addm(MONTH1, -26).replace(day=5)
    ws = lo
    title(ws, 'Loan details', 'Yellow cells: type your loan. Grey cells: results.')
    ins = [('Loan name', 'Home loan', None, ''), ('Currency symbol', '₹', None, ''), ('Loan amount', 3500000, NUM0, 'The amount borrowed (or the balance today, if you start mid-way).'),
           ('Interest rate per year', 0.085, '0.00%', 'For example 8.5%. Fixed rate assumed.'), ('Tenure (months)', 240, None, '1 to 360 months (20 years = 240).'),
           ('First EMI date', first, DT, 'Later EMIs fall on the same day each month.'), ('Extra payment every month', 0, NUM0, 'Optional: paid with every EMI to close the loan sooner.'),
           ('One-time prepayment amount', 200000, NUM0, 'Optional: a lump sum, such as a bonus.'), ('One-time prepayment in EMI number', 36, None, 'Paid together with this EMI number.')]
    for i, (k, v, fmt, note) in enumerate(ins, 3):
        label(ws, 'A%d' % i, k); cell(ws, 'B%d' % i, v, 'in', fmt, align='right' if fmt and fmt != DT else 'left'); cell(ws, 'D%d' % i, note, size=9, color=MUTED)
    v = DataValidation(type='whole', operator='between', formula1='1', formula2='360', showErrorMessage=True, errorTitle='Tenure', error='Type the number of months, from 1 to 360.')
    ws.add_data_validation(v); v.add('B7')
    label(ws, 'A13', 'Results'); ws['A13'].font = Font(bold=True, size=12, color=BLUE)
    res = [('EMI (monthly instalment)', '=IF(OR(N(B5)<=0,N(B7)<=0),"",ROUND(PMT(B6/12,B7,-B5),2))', NUM, 'Reducing-balance method, as banks use.'),
           ('Total interest without prepayment', '=IF(B14="","",ROUND(B14*B7-B5,0))', NUM0, 'EMI × months − loan amount.'),
           ('Total interest with your prepayments', '=SUM(Schedule!E5:E364)', NUM0, 'From the Schedule sheet.'),
           ('Interest saved', '=IF(B14="","",B15-B16)', NUM0, ''),
           ('Months to close the loan', '=COUNT(Schedule!C5:C364)', None, ''),
           ('Months saved', '=IF(B14="","",B7-B18)', None, ''),
           ('Loan closes on', '=IF(B18=0,"",INDEX(Schedule!B5:B364,B18))', DT, 'Date of the last EMI.'),
           ('EMIs marked paid', '=COUNT(Schedule!J5:J364)', None, 'Type the date in "Paid on" on the Schedule sheet.'),
           ('Balance still owed', '=IF(B21=0,N(B5),INDEX(Schedule!H5:H364,B21))', NUM0, 'Closing balance after the last EMI marked paid.')]
    for i, (k, f, fmt, note) in enumerate(res, 14):
        big = i in (14, 17, 22)
        label(ws, 'A%d' % i, k, bold=big); cell(ws, 'B%d' % i, f, 'soft' if big else 'calc', fmt, bold=big, align='right'); cell(ws, 'D%d' % i, note, size=9, color=MUTED)
    ws.merge_cells('A24:D26')
    cell(ws, 'A24', 'Assumes a fixed rate and that prepayments shorten the loan while the EMI stays the same. Your lender\'s statement is the final word; '
         'ask about prepayment charges first. A calculator and record only, not financial advice.', wrap=True, size=9, color=MUTED)
    protect(ws)
    # Schedule
    ws = sc
    title(ws, 'Repayment schedule', 'Amounts in the currency set on the Loan sheet. Type the date in the last column as each EMI is paid.')
    head(ws, 4, ['#', 'EMI date', 'Opening balance', 'EMI paid', 'Interest', 'Principal', 'Prepayment', 'Closing balance', 'Interest so far', 'Paid on (type the date)'])
    for k in range(1, 361):
        r = 4 + k
        cell(ws, 'C%d' % r, '=IF(OR(N(Loan!$B$5)<=0,Loan!$B$14=""),"",Loan!$B$5)' if k == 1 else '=IF(H{p}="","",IF(H{p}<0.005,"",H{p}))'.format(p=r - 1), 'calc', NUM)
        cell(ws, 'A%d' % r, '=IF($C{r}="","",{k})'.format(r=r, k=k), 'box', align='center')
        cell(ws, 'B%d' % r, '=IF($C{r}="","",EDATE(Loan!$B$8,{k}-1))'.format(r=r, k=k), 'box', DT)
        cell(ws, 'E%d' % r, '=IF($C{r}="","",ROUND($C{r}*Loan!$B$6/12,2))'.format(r=r), 'calc', NUM)
        cell(ws, 'F%d' % r, '=IF($C{r}="","",IF({k}>=Loan!$B$7,$C{r},MIN($C{r},MAX(0,Loan!$B$14-$E{r}))))'.format(r=r, k=k), 'calc', NUM)
        cell(ws, 'D%d' % r, '=IF($C{r}="","",$E{r}+$F{r})'.format(r=r), 'calc', NUM)
        cell(ws, 'G%d' % r, '=IF($C{r}="","",MAX(0,MIN($C{r}-$F{r},N(Loan!$B$9)+IF({k}=N(Loan!$B$11),N(Loan!$B$10),0))))'.format(r=r, k=k), 'calc', NUM)
        cell(ws, 'H%d' % r, '=IF($C{r}="","",ROUND($C{r}-$F{r}-$G{r},2))'.format(r=r), 'calc', NUM)
        cell(ws, 'I%d' % r, '=IF($C{r}="","",SUM($E$5:$E{r}))'.format(r=r), 'calc', NUM)
        d = addm(first, k - 1).replace(day=first.day)
        cell(ws, 'J%d' % r, d if d <= TODAY else None, 'in', DT)
    ws.conditional_formatting.add('A5:J364', FormulaRule(formula=['AND($A5<>"",$J5<>"")'], fill=GREEN))
    ws.conditional_formatting.add('G5:G364', FormulaRule(formula=['N($G5)>0'], fill=AMBER, font=Font(bold=True)))
    ws.freeze_panes = 'C5'
    ws.print_title_rows = '4:4'
    protect(ws)
    # All loans
    ws = al
    title(ws, 'All loans in the family', 'One row per loan. Type the outstanding amount from your latest statement.')
    label(ws, 'A3', 'Monthly take-home income'); cell(ws, 'B3', 156000, 'in', NUM0, align='right')
    label(ws, 'A4', 'Total EMI per month'); cell(ws, 'B4', '=SUM(F9:F28)', 'calc', NUM0, align='right')
    label(ws, 'A5', 'Total outstanding'); cell(ws, 'B5', '=SUM(D9:D28)', 'calc', NUM0, align='right')
    label(ws, 'A6', 'EMIs as a share of take-home income'); cell(ws, 'B6', '=IF(N(B3)<=0,"",B4/B3)', 'soft', '0%', bold=True, align='right')
    cell(ws, 'D6', 'Lenders look at this share when you apply for a new loan; a lower share leaves more room for savings and emergencies.', size=9, color=MUTED)
    head(ws, 8, ['Loan', 'Lender', 'Type', 'Outstanding', 'Rate', 'EMI', 'EMI day', 'Ends on', 'Account (last 4)', 'Notes'])
    grid(ws, 9, 28, 'ABCDEFGHIJ', {}, {'D': NUM0, 'E': '0.00%', 'F': NUM0, 'H': DT})
    fill_rows(ws, 9, [('Home loan', 'State Bank of India', 'Home', 3180000, 0.085, 30374, 5, addm(first, 239), '4417', 'Floating rate, linked to repo rate'),
                      ('Car loan', 'HDFC Bank', 'Vehicle', 412000, 0.092, 14650, 10, addm(MONTH1, 31), '0932', ''),
                      ('Laptop on EMI', 'Credit card (no-cost EMI)', 'Credit card', 38000, 0, 6334, 18, addm(MONTH1, 6), '7781', 'Check processing fee')])
    dv(ws, 'C9:C28', '=Lists!$C$1:$C$8')
    ws.freeze_panes = 'A9'
    save(wb, 'HM-Loan-EMI-Tracker.xlsx')

# ------------------------------------------------------------------ 3. Bills & subscriptions
FREQ = [('Monthly', 1), ('Every 2 months', 2), ('Quarterly', 3), ('Every 4 months', 4), ('Half-yearly', 6), ('Yearly', 12), ('One-time', 0)]
BILL_CATS = ['Electricity', 'Water', 'Cooking gas', 'Mobile', 'Internet and broadband', 'DTH / cable TV', 'Streaming (video and music)', 'Software and apps',
             'Rent / maintenance', 'Insurance premium', 'Loan EMI', 'School / tuition fees', 'Credit card', 'Property tax', 'Club / gym', 'Newspaper and magazines',
             'Domestic help', 'Other']

def hm_bills():
    wb = new_wb('hm-bills-due-excel')
    howto(wb, first=True)
    settings(wb, [('Currency symbol', '₹', None), ('As on (today unless you type a date)', '=TODAY()', DT), ('"Due soon" warning (days before)', 7, None),
                  ('Calendar starts (month)', MONTH1, MON)])
    bl = sheet(wb, 'Bills', [26, 22, 20, 15, 12, 15, 13, 13, 9, 15, 10, 9, 13, 9, 12, 13, 12, 5, 5], tab=PINK)
    up = sheet(wb, 'Upcoming', [6, 30, 13, 13, 10, 12, 10], tab=BLUE)
    ca = sheet(wb, 'Calendar', [28, 11] + [10] * 12, tab=BLUE)
    sm = sheet(wb, 'Summary', [46, 16])
    lists_sheet(wb, extra={3: [f for f, _ in FREQ], 4: [m for _, m in FREQ], 5: BILL_CATS, 6: ['Yes', 'No'], 7: ['Keep', 'Review', 'Cancel']})
    ws = bl
    title(ws, 'Bills and subscriptions', 'Amounts in the currency set on the Settings sheet')
    cell(ws, 'D2', 'Each time you pay, type the date in "Last paid on".', size=9, color=MUTED)
    head(ws, 3, ['Bill / subscription', 'Category', 'Provider', 'Account no. (last 4)', 'Amount', 'How often', 'First due date', 'Last paid on', 'Auto-pay', 'Pay from', 'Keep?',
                 'Every (months)', 'Next due', 'Days left', 'Status', 'Yearly cost', 'Monthly average', 'k', 'key'])
    REF = 'IF($H{r}="",$G{r}-1,$H{r}+MIN(45,15*$L{r}))'
    E = lambda n: 'EDATE($G{r},$L{r}*(' + '$R{r}' + ('+%d' % n if n else '') + '))'
    f = {'L': '=IF($F{r}="","",IFERROR(INDEX(Lists!$D$1:$D$7,MATCH($F{r},Lists!$C$1:$C$7,0)),1))',
         'R': '=IF(OR($G{r}="",$L{r}="",N($L{r})=0),"",IF(' + REF + '<$G{r},0,INT(DATEDIF($G{r},' + REF + ',"m")/$L{r})))',
         'M': ('=IF(OR($A{r}="",$G{r}="",$L{r}=""),"",IF($L{r}=0,IF($H{r}="",$G{r},""),IF(%s>%s,%s,IF(%s>%s,%s,%s))))'
               % (E(0), REF, E(0), E(1), REF, E(1), E(2))),
         'N': '=IF($M{r}="","",$M{r}-Settings!$B$4)',
         'O': '=IF($A{r}="","",IF($M{r}="","Paid",IF($N{r}<0,"Overdue",IF($N{r}=0,"Due today",IF($N{r}<=Settings!$B$5,"Due soon","OK")))))',
         'P': '=IF(OR($A{r}="",$L{r}=""),"",IF($L{r}=0,0,N($E{r})*12/$L{r}))',
         'Q': '=IF($P{r}="","",$P{r}/12)',
         'S': '=IF($M{r}="","",$M{r}+ROW()/100000)'}
    grid(ws, 4, 203, 'ABCDEFGHIJKLMNOPQRS', {c: 'calc' for c in 'LMNOPQRS'},
         {'E': NUM0, 'G': DT, 'H': DT, 'M': DT, 'P': NUM0, 'Q': NUM0}, f)
    ws.column_dimensions['R'].hidden = ws.column_dimensions['S'].hidden = True
    m = lambda k, d: addm(MONTH1, k).replace(day=d)
    fill_rows(ws, 4, [
        ('Electricity', 'Electricity', 'State power board', '4521', 2400, 'Every 2 months', m(-6, 12), m(-2, 11), 'No', 'UPI', 'Keep'),
        ('Mobile postpaid (2 lines)', 'Mobile', 'Airtel', '7788', 1198, 'Monthly', m(-12, 18), m(-1, 17), 'Yes', 'Credit card', 'Keep'),
        ('Home broadband', 'Internet and broadband', 'ACT Fibernet', '3310', 2997, 'Quarterly', m(-9, 5), m(-3, 4), 'No', 'UPI', 'Keep'),
        ('Video streaming', 'Streaming (video and music)', 'Netflix', None, 649, 'Monthly', m(-8, 22), m(-1, 22), 'Yes', 'Credit card', 'Review'),
        ('Second video app (rarely used)', 'Streaming (video and music)', 'Prime Video', None, 1499, 'Yearly', m(-14, 9), m(-2, 9), 'Yes', 'Credit card', 'Cancel'),
        ('Music streaming', 'Streaming (video and music)', 'Spotify', None, 119, 'Monthly', m(-20, 3), m(0, 3), 'Yes', 'Credit card', 'Keep'),
        ('Cloud storage 200 GB', 'Software and apps', 'Google One', None, 2100, 'Yearly', m(-14, 15), m(-2, 15), 'Yes', 'Credit card', 'Keep'),
        ('Gym membership', 'Club / gym', 'FitZone', None, 4500, 'Quarterly', m(-5, 1), m(-2, 1), 'No', 'UPI', 'Review'),
        ('Car insurance', 'Insurance premium', 'ICICI Lombard', '0912', 14800, 'Yearly', m(-11, 20), m(-11, 18), 'No', 'Credit card', 'Keep'),
        ('Society maintenance', 'Rent / maintenance', 'Green Park RWA', 'Flat 2B', 3500, 'Monthly', m(-12, 10), m(-1, 8), 'No', 'Bank transfer', 'Keep'),
        ('Property tax', 'Property tax', 'Municipal corporation', '6620', 6200, 'Yearly', m(-16, 28), m(-4, 20), 'No', 'UPI', 'Keep'),
        ('DTH recharge', 'DTH / cable TV', 'Tata Play', '5566', 350, 'Monthly', m(-6, 2), m(-2, 1), 'No', 'UPI', 'Review'),
        ('School fees (term)', 'School / tuition fees', 'City Public School', 'Adm 2231', 28000, 'Quarterly', m(-7, 10), m(-1, 9), 'No', 'Bank transfer', 'Keep')])
    dv(ws, 'B4:B203', '=Lists!$E$1:$E$18'); dv(ws, 'F4:F203', '=Lists!$C$1:$C$7'); dv(ws, 'I4:I203', '=Lists!$F$1:$F$2'); dv(ws, 'K4:K203', '=Lists!$G$1:$G$3')
    ws.conditional_formatting.add('O4:O203', CellIsRule(operator='equal', formula=['"Overdue"'], fill=RED, font=Font(bold=True, color='991B1B')))
    ws.conditional_formatting.add('O4:O203', FormulaRule(formula=['OR($O4="Due soon",$O4="Due today")'], fill=AMBER))
    ws.freeze_panes = 'B4'
    # Upcoming
    ws = up
    title(ws, 'Upcoming bills in date order', 'The next 30 due dates, earliest first, so overdue bills are at the top.')
    head(ws, 3, ['#', 'Bill', 'Due date', 'Amount', 'Days left', 'Status', 'Auto-pay'])
    B = lambda c: 'Bills!$%s$4:$%s$203' % (c, c)
    upcoming(ws, 4, 30, B('S'), [('A', None, None), ('B', B('A'), None), ('C', B('M'), DT), ('D', B('E'), NUM0), ('E', B('N'), None), ('F', B('O'), None), ('G', B('I'), None)], 9)
    ws.conditional_formatting.add('F4:F33', CellIsRule(operator='equal', formula=['"Overdue"'], fill=RED, font=Font(bold=True, color='991B1B')))
    ws.conditional_formatting.add('F4:F33', FormulaRule(formula=['OR($F4="Due soon",$F4="Due today")'], fill=AMBER))
    protect(ws)
    # Calendar
    ws = ca
    title(ws, 'Bills calendar: money needed each month', 'Amounts in the currency set on the Settings sheet')
    head(ws, 3, ['Bill', 'Amount'] + [''] * 12)
    month_head(ws, 3, 3, 12, 'Settings!$B$6')
    cell(ws, 'A4', 'Total due in the month', 'soft', bold=True)
    for k in range(12): cell(ws, '%s4' % L(3 + k), '=SUM(%s5:%s204)' % (L(3 + k), L(3 + k)), 'soft', NUM0, bold=True)
    for i in range(200):
        r, s = 5 + i, 4 + i
        G, Lm = 'Bills!$G%d' % s, 'Bills!$L%d' % s
        cell(ws, 'A%d' % r, '=IF(Bills!A%d="","",Bills!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",N(Bills!E%d))' % (r, s), 'calc', NUM0)
        for k in range(12):
            c = L(3 + k)
            cell(ws, '%s%d' % (c, r), ('=IF(OR($A{r}="",{G}="",{L}=""),"",IF({L}=0,IF(AND(YEAR({G})=YEAR({c}$3),MONTH({G})=MONTH({c}$3)),$B{r},0),'
                                        'IF(AND({c}$3>=DATE(YEAR({G}),MONTH({G}),1),MOD((YEAR({c}$3)-YEAR({G}))*12+MONTH({c}$3)-MONTH({G}),{L})=0),$B{r},0)))').format(r=r, c=c, G=G, L=Lm),
                 'calc', DASH)
    ws.freeze_panes = 'C5'
    ws.page_setup.orientation = 'landscape'
    protect(ws)
    # Summary
    ws = sm
    title(ws, 'Summary', 'Amounts in the currency set on the Settings sheet')
    rows = [('Bills and subscriptions listed', '=COUNTIF(Bills!A4:A203,"?*")', None), ('Yearly cost of all bills', '=SUM(Bills!P4:P203)', NUM0),
            ('Monthly average', '=B4/12', NUM0), ('Overdue now', '=COUNTIF(Bills!O4:O203,"Overdue")', None),
            ('Due today or soon', '=COUNTIF(Bills!O4:O203,"Due today")+COUNTIF(Bills!O4:O203,"Due soon")', None), ('On auto-pay', '=COUNTIF(Bills!I4:I203,"Yes")', None),
            ('Yearly cost of items marked Review', '=SUMIF(Bills!K4:K203,"Review",Bills!P4:P203)', NUM0),
            ('Yearly cost of items marked Cancel', '=SUMIF(Bills!K4:K203,"Cancel",Bills!P4:P203)', NUM0),
            ('Heaviest month in the calendar (amount)', '=MAX(Calendar!C4:N4)', NUM0),
            ('Heaviest month', '=IF(B11=0,"",INDEX(Calendar!C3:N3,MATCH(B11,Calendar!C4:N4,0)))', MON)]
    for i, (k, f2, fmt) in enumerate(rows, 3):
        label(ws, 'A%d' % i, k, bold=False); cell(ws, 'B%d' % i, f2, 'calc', fmt, bold=True, align='right')
    red_if(ws, 'B6', 'B6>0', bold=True)
    ch = BarChart(); ch.type = 'col'; ch.style = 10; ch.title = 'Money needed each month'; ch.legend = None
    ch.add_data(Reference(ca, min_col=3, max_col=14, min_row=4), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(ca, min_col=3, max_col=14, min_row=3))
    ch.series[0].graphicalProperties.solidFill = PINK
    ch.height, ch.width = 7.5, 16
    ws.add_chart(ch, 'D3')
    protect(ws)
    save(wb, 'HM-Bills-and-Subscriptions.xlsx')

# ------------------------------------------------------------------ 4. Family documents & policies
DOCS = ['Aadhaar / national ID card', 'PAN / tax ID card', 'Passport', 'Driving licence', 'Voter ID', 'Birth certificate', 'Marriage certificate',
        'School / college certificates', 'Property sale deed / title', 'Rent or lease agreement', 'Vehicle registration (RC)', 'Vehicle pollution / fitness',
        'Bank account / passbook', 'Fixed deposit / bond', 'Investment statements', 'Pension / retirement account', 'Income tax returns', 'Will',
        'Medical records', 'Vaccination record', 'Warranty card', 'Gas connection papers', 'Other']
POL = ['Health', 'Life (term)', 'Life (other)', 'Vehicle', 'Home / property', 'Travel', 'Personal accident', 'Other']
PAY = [('Yearly', 1), ('Half-yearly', 2), ('Quarterly', 4), ('Monthly', 12), ('Single premium', 0)]

def hm_documents():
    wb = new_wb('hm-family-documents-excel')
    howto(wb, first=True)
    settings(wb, [('As on (today unless you type a date)', '=TODAY()', DT), ('Warn me this many days before expiry', 60, None), ('Currency symbol', '₹', None)])
    me = sheet(wb, 'Members', [22, 14, 13, 11, 16, 34])
    do = sheet(wb, 'Documents', [16, 28, 11, 22, 12, 12, 30, 28, 18, 24, 9, 12, 5], tab=PINK)
    po = sheet(wb, 'Policies', [24, 16, 18, 11, 22, 14, 12, 14, 13, 16, 22, 24, 13, 9, 12, 5], tab=PINK)
    co = sheet(wb, 'Contacts', [30, 24, 18, 40])
    sm = sheet(wb, 'Summary', [36, 14, 4, 30, 13, 18, 5, 5], tab=BLUE)
    lists_sheet(wb, extra={3: DOCS, 4: POL, 5: [p for p, _ in PAY], 6: [n for _, n in PAY], 7: ['Yes', 'No']})
    warn = 'Never type passwords, PINs, CVV, OTPs or full account / card numbers here. The last 4 digits are enough to find a document.'
    # Members
    title(me, 'Family members', 'One row per person. Names here appear in the drop-down on the Documents sheet.')
    head(me, 3, ['Name', 'Relation', 'Date of birth', 'Blood group', 'Phone', 'Notes'])
    grid(me, 4, 23, 'ABCDEF', {}, {'C': DT})
    fill_rows(me, 4, [('Rohit Sharma', 'Self', datetime.date(1984, 3, 12), 'B+', '98xxx xx101', ''), ('Neha Sharma', 'Spouse', datetime.date(1987, 8, 23), 'O+', '98xxx xx102', 'Allergic to penicillin'),
                      ('Aarav Sharma', 'Son', datetime.date(2014, 1, 5), 'B+', '', ''), ('Kamla Devi', 'Mother', datetime.date(1956, 11, 30), 'A+', '98xxx xx104', 'Diabetic; medicines list in Medical records')])
    # Documents
    ws = do
    title(ws, 'Documents index', 'One row per document: whose it is, where the original is and where the scanned copy is.')
    ws.merge_cells('A3:L3'); c = cell(ws, 'A3', warn, bold=True, color='991B1B', size=10); c.fill = RED
    head(ws, 4, ['Member', 'Document', 'No. (last 4)', 'Issued by', 'Issued on', 'Valid until', 'Original kept at', 'Digital copy at', 'Nominee / joint holder', 'Notes', 'Days left', 'Status', 'key'])
    grid(ws, 5, 304, 'ABCDEFGHIJKLM', {'K': 'calc', 'L': 'calc', 'M': 'calc'}, {'E': DT, 'F': DT},
         {'K': '=IF($F{r}="","",$F{r}-Settings!$B$3)',
          'L': '=IF(AND($A{r}="",$B{r}=""),"",IF($F{r}="","No expiry",IF($K{r}<0,"Expired",IF($K{r}<=Settings!$B$4,"Renew soon","Valid"))))',
          'M': '=IF($F{r}="","",$F{r}+ROW()/100000)'})
    ws.column_dimensions['M'].hidden = True
    T = lambda n: TODAY + datetime.timedelta(days=n)
    fill_rows(ws, 5, [('Rohit Sharma', 'Passport', '4471', 'Passport Office, Chandigarh', T(-3605), T(45), 'Steel almirah, top shelf, blue folder', 'Google Drive › Family › IDs', '', ''),
                      ('Rohit Sharma', 'PAN / tax ID card', '812K', 'Income Tax Department', datetime.date(2006, 7, 1), None, 'Wallet (card); copy in blue folder', 'Google Drive › Family › IDs', '', ''),
                      ('Neha Sharma', 'Driving licence', '0921', 'RTO Ludhiana', T(-2900), T(800), 'Handbag', 'Phone › DigiLocker', '', ''),
                      ('Aarav Sharma', 'Passport', '6620', 'Passport Office, Chandigarh', T(-1846), T(-20), 'Steel almirah, top shelf, blue folder', 'Google Drive › Family › IDs', '', 'Minor passport: 5 years'),
                      ('Rohit Sharma', 'Property sale deed / title', '—', 'Sub-Registrar office', datetime.date(2016, 9, 14), None, 'Bank locker 214, SBI Model Town', 'Scanned: external disk, folder House', 'Neha Sharma', 'Original, plus 2 certified copies at home'),
                      ('Rohit Sharma', 'Vehicle registration (RC)', '3301', 'RTO Ludhiana', datetime.date(2021, 2, 2), T(1900), 'Car glovebox (copy); original in blue folder', 'Phone › DigiLocker', '', ''),
                      ('Rohit Sharma', 'Vehicle pollution / fitness', '—', 'Authorised testing centre', T(-150), T(30), 'Car glovebox', '', '', '6-monthly check'),
                      ('Kamla Devi', 'Will', '—', 'Registered with Sub-Registrar', datetime.date(2022, 5, 10), None, 'With family lawyer (see Contacts)', '', 'Rohit Sharma', 'Executor: Rohit Sharma')])
    dv(ws, 'A5:A304', '=Members!$A$4:$A$23'); dv(ws, 'B5:B304', '=Lists!$C$1:$C$23')
    ws.conditional_formatting.add('L5:L304', CellIsRule(operator='equal', formula=['"Expired"'], fill=RED, font=Font(bold=True, color='991B1B')))
    ws.conditional_formatting.add('L5:L304', CellIsRule(operator='equal', formula=['"Renew soon"'], fill=AMBER))
    ws.freeze_panes = 'C5'
    # Policies
    ws = po
    title(ws, 'Insurance policies', 'Amounts in the currency set on the Settings sheet')
    ws.merge_cells('A3:O3'); c = cell(ws, 'A3', warn, bold=True, color='991B1B', size=10); c.fill = RED
    head(ws, 4, ['Policy', 'Type', 'Insurer', 'Policy no. (last 4)', 'Covers (people / asset)', 'Sum insured', 'Premium', 'Paid', 'Next renewal / due', 'Nominee',
                 'Agent / helpline', 'Policy kept at', 'Yearly premium', 'Days left', 'Status', 'key'])
    grid(ws, 5, 104, 'ABCDEFGHIJKLMNOP', {c: 'calc' for c in 'MNOP'}, {'F': NUM0, 'G': NUM0, 'I': DT, 'M': NUM0},
         {'M': '=IF(OR($G{r}="",$H{r}=""),"",N($G{r})*IFERROR(INDEX(Lists!$F$1:$F$5,MATCH($H{r},Lists!$E$1:$E$5,0)),1))',
          'N': '=IF($I{r}="","",$I{r}-Settings!$B$3)',
          'O': '=IF($A{r}="","",IF($I{r}="","No date",IF($N{r}<0,"Overdue",IF($N{r}<=Settings!$B$4,"Renew soon","OK"))))',
          'P': '=IF($I{r}="","",$I{r}+ROW()/100000)'})
    ws.column_dimensions['P'].hidden = True
    fill_rows(ws, 5, [('Family health floater', 'Health', 'Star Health', '2210', 'Rohit, Neha, Aarav', 1000000, 28500, 'Yearly', T(50), 'Neha Sharma', 'Agent Mr. Gill 98xxx xx210', 'Blue folder; e-copy in email'),
                      ('Term life cover', 'Life (term)', 'HDFC Life', '7734', 'Rohit', 15000000, 18200, 'Yearly', T(140), 'Neha Sharma', 'Helpline on policy', 'Blue folder'),
                      ('Car insurance', 'Vehicle', 'ICICI Lombard', '0912', 'Car (Hyundai i20)', 650000, 14800, 'Yearly', T(18), '', 'App / helpline', 'Car glovebox + email'),
                      ('Home contents', 'Home / property', 'Bajaj Allianz', '5521', 'Flat 2B contents', 1500000, 3200, 'Yearly', T(210), 'Neha Sharma', '', 'Email'),
                      ('Mother\'s health policy', 'Health', 'Niva Bupa', '3398', 'Kamla Devi', 300000, 11000, 'Half-yearly', T(95), 'Rohit Sharma', 'Agent Mr. Gill 98xxx xx210', 'Blue folder')])
    dv(ws, 'B5:B104', '=Lists!$D$1:$D$8'); dv(ws, 'H5:H104', '=Lists!$E$1:$E$5')
    ws.conditional_formatting.add('O5:O104', CellIsRule(operator='equal', formula=['"Overdue"'], fill=RED, font=Font(bold=True, color='991B1B')))
    ws.conditional_formatting.add('O5:O104', CellIsRule(operator='equal', formula=['"Renew soon"'], fill=AMBER))
    ws.freeze_panes = 'B5'
    # Contacts
    title(co, 'Emergency and useful contacts', 'Print this sheet and keep it with the original documents.')
    head(co, 3, ['Who', 'Name', 'Phone', 'Notes'])
    grid(co, 4, 33, 'ABCD', {})
    fill_rows(co, 4, [(w,) for w in ['Family doctor', 'Nearest hospital (emergency)', 'Ambulance', 'Health insurance helpline', 'Car insurance helpline', 'Bank (main account)',
                                      'Tax adviser / accountant', 'Family lawyer', 'Society office / landlord', 'Employer HR (Rohit)', 'Employer HR (Neha)', 'Trusted neighbour', 'Relative to inform']])
    # Summary
    ws = sm
    title(ws, 'Summary', 'Amounts in the currency set on the Settings sheet')
    rows = [('Documents listed', '=COUNTIF(Documents!B5:B304,"?*")', None), ('Documents expired', '=COUNTIF(Documents!L5:L304,"Expired")', None),
            ('Documents to renew soon', '=COUNTIF(Documents!L5:L304,"Renew soon")', None), ('Policies listed', '=COUNTIF(Policies!A5:A104,"?*")', None),
            ('Policies overdue or to renew soon', '=COUNTIF(Policies!O5:O104,"Overdue")+COUNTIF(Policies!O5:O104,"Renew soon")', None),
            ('Total yearly premium', '=SUM(Policies!M5:M104)', NUM0)]
    for i, (k, f2, fmt) in enumerate(rows, 3):
        label(ws, 'A%d' % i, k, bold=False); cell(ws, 'B%d' % i, f2, 'calc', fmt, bold=True, align='right')
    red_if(ws, 'B4', 'B4>0', bold=True); red_if(ws, 'B5', 'B5>0', AMBER); red_if(ws, 'B7', 'B7>0', AMBER)
    head(ws, 10, ['Policy type', 'Sum insured', 'Yearly premium'], col=1)
    ws['C10'].fill = F_HEAD
    for i, p in enumerate(POL, 11):
        cell(ws, 'A%d' % i, p, 'box')
        cell(ws, 'B%d' % i, '=SUMIF(Policies!$B$5:$B$104,A%d,Policies!$F$5:$F$104)' % i, 'calc', NUM0)
        cell(ws, 'C%d' % i, '=SUMIF(Policies!$B$5:$B$104,A%d,Policies!$M$5:$M$104)' % i, 'calc', NUM0)
    ws.column_dimensions['C'].width = 15
    head(ws, 3, ['Next document expiries', 'Date', 'Whose'], col=4)
    DR = lambda c: 'Documents!$%s$5:$%s$304' % (c, c)
    upcoming(ws, 4, 10, DR('M'), [('D', DR('B'), None), ('E', DR('F'), DT), ('F', DR('A'), None)], 7)
    head(ws, 16, ['Next policy renewals', 'Date', 'Premium'], col=4)
    PR = lambda c: 'Policies!$%s$5:$%s$104' % (c, c)
    upcoming(ws, 17, 10, PR('P'), [('D', PR('A'), None), ('E', PR('I'), DT), ('F', PR('G'), NUM0)], 7)
    red_if(ws, 'D4:F13', 'AND($E4<>"",$E4<Settings!$B$3)')
    red_if(ws, 'D17:F26', 'AND($E17<>"",$E17<Settings!$B$3)')
    protect(ws)
    save(wb, 'HM-Family-Documents-Index.xlsx')

# ------------------------------------------------------------------ 5. House help salary & advances
WORK = ['Cook', 'Cleaning', 'Driver', 'Nanny / child care', 'Elder care', 'Gardener', 'Guard / watchman', 'Laundry / ironing', 'All-round help', 'Other']

def hm_house_help():
    wb = new_wb('hm-house-help-excel')
    howto(wb, first=True)
    settings(wb, [('Household name', 'Sharma household, Flat 2B', None), ('Currency symbol', '₹', None), ('Year starts (for yearly totals)', datetime.date(TODAY.year, 1, 1), MON)])
    stf = sheet(wb, 'Staff', [20, 18, 14, 26, 12, 13, 10, 10, 18, 20, 40])
    pr = sheet(wb, 'Pay register', [12, 20, 10, 12, 12, 13, 13, 26, 13, 10, 10, 12, 13, 5], tab=PINK)
    ad = sheet(wb, 'Advances', [13, 20, 14, 32, 14, 26])
    sm = sheet(wb, 'Summary', [20, 14, 12, 15, 14, 14, 16], tab=BLUE)
    rc = sheet(wb, 'Receipt', [30, 24, 4, 20, 14, 5], tab=BLUE)
    lists_sheet(wb, extra={3: WORK, 4: ['Cash', 'UPI', 'Bank transfer', 'Cheque'], 5: ['Yes', 'No']})
    cur = 'Amounts in the currency set on the Settings sheet'
    # Staff
    title(stf, 'People who help at home', cur)
    head(stf, 3, ['Name', 'Work', 'Phone', 'Address', 'Started on', 'Monthly salary', 'Paid leaves a month', 'ID proof seen', 'Police verification', 'Emergency contact',
                  'Notes (weekly off, timings, notice agreed)'])
    grid(stf, 4, 33, 'ABCDEFGHIJK', {}, {'E': DT, 'F': NUM0})
    fill_rows(stf, 4, [('Sunita Devi', 'Cook', '98xxx xx301', 'Sector 14', datetime.date(2023, 4, 1), 7000, 2, 'Yes', 'Done (Apr 2023)', 'Husband 98xxx xx302', 'Two meals; Thursday off; one month notice'),
                       ('Ramesh Kumar', 'Driver', '98xxx xx303', 'Shivpuri', datetime.date(2022, 8, 16), 16000, 2, 'Yes', 'Done (Aug 2022)', 'Brother 98xxx xx304', '8 am to 7 pm; Sunday off; overtime agreed per hour after 8 pm'),
                       ('Kamla', 'Cleaning', '98xxx xx305', 'Sector 14', datetime.date(2024, 1, 10), 4500, 2, 'Yes', 'Applied', '', 'Sweeping, mopping, dishes; Tuesday off')])
    dv(stf, 'B4:B33', '=Lists!$C$1:$C$10'); dv(stf, 'H4:H33', '=Lists!$E$1:$E$2')
    stf.freeze_panes = 'B4'
    # Pay register
    ws = pr
    title(ws, 'Pay register', cur)
    head(ws, 3, ['Month', 'Name', 'Leaves taken', 'Bonus / extra', 'Advance recovered', 'Paid on', 'Mode', 'Note', 'Monthly salary', 'Paid leaves', 'Extra leaves',
                 'Leave deduction', 'Net paid', 'key'])
    SL = lambda col: 'IFERROR(N(INDEX(Staff!$%s$4:$%s$33,MATCH($B{r},Staff!$A$4:$A$33,0))),0)' % (col, col)
    grid(ws, 4, 603, 'ABCDEFGHIJKLMN', {c: 'calc' for c in 'IJKLMN'}, {'A': MMY, 'D': NUM0, 'E': NUM0, 'F': DT, 'I': NUM0, 'L': NUM0, 'M': NUM0},
         {'I': '=IF($B{r}="","",' + SL('F') + ')', 'J': '=IF($B{r}="","",' + SL('G') + ')',
          'K': '=IF($B{r}="","",MAX(0,N($C{r})-$J{r}))',
          'L': '=IF(OR($B{r}="",$A{r}=""),"",ROUND($I{r}/DAY(EOMONTH($A{r},0))*$K{r},0))',
          'M': '=IF($B{r}="","",$I{r}-N($L{r})+N($D{r})-N($E{r}))',
          'N': '=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'})
    ws.column_dimensions['N'].hidden = True
    p2, p1 = addm(MONTH1, -2), addm(MONTH1, -1)
    fill_rows(ws, 4, [(p2, 'Sunita Devi', 2, None, None, p2.replace(day=28) + datetime.timedelta(days=3), 'UPI', ''),
                      (p2, 'Ramesh Kumar', 1, 1200, 2000, p2.replace(day=28) + datetime.timedelta(days=3), 'Bank transfer', 'Overtime 6 hours'),
                      (p2, 'Kamla', 3, None, None, p2.replace(day=28) + datetime.timedelta(days=3), 'Cash', ''),
                      (p1, 'Sunita Devi', 3, 1000, None, p1.replace(day=28) + datetime.timedelta(days=3), 'UPI', 'Festival bonus'),
                      (p1, 'Ramesh Kumar', 2, 3000, 2000, p1.replace(day=28) + datetime.timedelta(days=3), 'Bank transfer', 'Festival bonus'),
                      (p1, 'Kamla', 2, 1000, None, p1.replace(day=28) + datetime.timedelta(days=3), 'Cash', 'Festival bonus')])
    dv(ws, 'B4:B603', '=Staff!$A$4:$A$33'); dv(ws, 'G4:G603', '=Lists!$D$1:$D$4')
    ws.conditional_formatting.add('L4:L603', CellIsRule(operator='greaterThan', formula=['0'], fill=AMBER))
    ws.freeze_panes = 'C4'
    # Advances
    title(ad, 'Advances given', cur)
    head(ad, 3, ['Date', 'Name', 'Amount given', 'Reason', 'Recover per month (plan)', 'Note'])
    grid(ad, 4, 203, 'ABCDEF', {}, {'A': DT, 'C': NUM0, 'E': NUM0})
    fill_rows(ad, 4, [(addm(MONTH1, -3).replace(day=15), 'Ramesh Kumar', 10000, 'Child\'s school admission', 2000, 'Agreed in writing')])
    dv(ad, 'B4:B203', '=Staff!$A$4:$A$33')
    # Summary
    ws = sm
    title(ws, 'Summary', 'Yearly figures cover 12 months from the month on the Settings sheet.')
    head(ws, 4, ['Name', 'Monthly salary', 'Months paid this year', 'Paid this year', 'Advances given', 'Advance recovered', 'Advance outstanding'])
    Y1, Y2 = 'DATE(YEAR(Settings!$B$5),MONTH(Settings!$B$5),1)', 'EDATE(DATE(YEAR(Settings!$B$5),MONTH(Settings!$B$5),1),12)'
    PRg = lambda c: "'Pay register'!$%s$4:$%s$603" % (c, c)
    for i in range(30):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Staff!A%d="","",Staff!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",N(Staff!F%d))' % (r, s), 'calc', NUM0)
        cell(ws, 'C%d' % r, '=IF($A{r}="","",COUNTIFS({B},$A{r},{N},">="&{y1},{N},"<"&{y2}))'.format(r=r, B=PRg('B'), N=PRg('N'), y1=Y1, y2=Y2), 'calc', align='center')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIFS({M},{B},$A{r},{N},">="&{y1},{N},"<"&{y2}))'.format(r=r, M=PRg('M'), B=PRg('B'), N=PRg('N'), y1=Y1, y2=Y2), 'calc', NUM0)
        cell(ws, 'E%d' % r, '=IF($A{r}="","",SUMIF(Advances!$B$4:$B$203,$A{r},Advances!$C$4:$C$203))'.format(r=r), 'calc', NUM0)
        cell(ws, 'F%d' % r, '=IF($A{r}="","",SUMIF({B},$A{r},{E}))'.format(r=r, B=PRg('B'), E=PRg('E')), 'calc', NUM0)
        cell(ws, 'G%d' % r, '=IF($A{r}="","",E{r}-F{r})'.format(r=r), 'calc', NUM0, bold=True)
    ws.conditional_formatting.add('G5:G34', CellIsRule(operator='greaterThan', formula=['0'], fill=AMBER))
    cell(ws, 'A36', 'Total', 'soft', bold=True)
    for col in 'BDEFG': cell(ws, '%s36' % col, '=SUM(%s5:%s34)' % (col, col), 'soft', NUM0, bold=True)
    protect(ws)
    # Receipt
    ws = rc
    ws.merge_cells('A1:E1'); cell(ws, 'A1', 'SALARY PAYMENT RECEIPT', bold=True, size=16, color=BLUE, align='center')
    ws.merge_cells('A2:E2'); cell(ws, 'A2', '=Settings!B3', align='center', size=10)
    label(ws, 'A4', 'Name'); cell(ws, 'B4', 'Ramesh Kumar', 'in', align='left'); dv(ws, 'B4', '=Staff!$A$4:$A$33', 'Pick a name')
    label(ws, 'A5', 'Month'); cell(ws, 'B5', p1, 'in', MON, align='left')
    cell(ws, 'F4', "=IFERROR(MATCH(1,INDEX(({B}=$B$4)*({N}=DATE(YEAR($B$5),MONTH($B$5),1)),0),0),\"\")".format(B=PRg('B'), N=PRg('N')))
    ws.column_dimensions['F'].hidden = True
    ws.merge_cells('A6:E6'); cell(ws, 'A6', '=IF($F$4="","No payment found for this name and month on the Pay register.","")', bold=True, color='991B1B', size=9)
    G = lambda c: '=IF($F$4="","",INDEX(%s,$F$4))' % PRg(c)
    lines = [('Work', '=IFERROR(INDEX(Staff!$B$4:$B$33,MATCH($B$4,Staff!$A$4:$A$33,0))&"","")', None), ('Monthly salary', G('I'), NUM0),
             ('Leaves taken', G('C'), None), ('Paid leaves allowed', G('J'), None), ('Less: leave deduction', G('L'), NUM0), ('Add: bonus / extra', G('D'), NUM0),
             ('Less: advance recovered', G('E'), NUM0), ('Net amount paid', G('M'), NUM0), ('Paid on', G('F'), DT), ('Mode of payment', G('G'), None),
             ('Advance still outstanding', '=IFERROR(INDEX(Summary!$G$5:$G$34,MATCH($B$4,Summary!$A$5:$A$34,0)),"")', NUM0)]
    for i, (k, f2, fmt) in enumerate(lines, 8):
        big = k == 'Net amount paid'
        label(ws, 'A%d' % i, k, bold=big); cell(ws, 'B%d' % i, f2, 'soft' if big else 'box', fmt, bold=big, align='right', size=13 if big else None)
    ws.merge_cells('A20:E20'); cell(ws, 'A20', '=IF($F$4="","","Received the above amount with thanks.")', size=10)
    ws['A23'], ws['D23'] = '____________________', '____________________'
    cell(ws, 'A24', 'Signature of the person paid', size=9); cell(ws, 'D24', '="For "&Settings!B3', size=9)
    ws.print_area = 'A1:E25'
    protect(ws)
    save(wb, 'HM-House-Help-Register.xlsx')

# ------------------------------------------------------------------ 6. Home inventory
INV_CATS = ['Electronics', 'Appliances', 'Furniture', 'Kitchenware', 'Jewellery and watches', 'Clothing and bags', 'Art and collectibles', 'Sports and bicycles',
            'Musical instruments', 'Tools', 'Books and media', 'Other']

def hm_inventory():
    wb = new_wb('hm-home-inventory-excel')
    howto(wb, first=True)
    st = settings(wb, [('Household and address', 'Flat 2B, Green Park, Ludhiana', None), ('Currency symbol', '₹', None), ('As on (today unless you type a date)', '=TODAY()', DT),
                       ('Insurer', 'Bajaj Allianz', None), ('Contents policy no. (last 4)', '5521', None), ('Contents sum insured', 1500000, NUM0)], widths=(34, 30, 4, 26))
    label(st, 'D2', 'Rooms (edit freely)')
    for i, rm in enumerate(['Living room', 'Kitchen', 'Main bedroom', 'Second bedroom', 'Kids\' room', 'Study', 'Balcony / store', 'Garage / parking'] + [None] * 12, 3):
        cell(st, 'D%d' % i, rm, 'in')
    it = sheet(wb, 'Items', [16, 28, 18, 22, 16, 12, 6, 13, 13, 13, 24, 8, 9, 24, 13, 14, 17, 8, 11], tab=PINK)
    sm = sheet(wb, 'Summary', [30, 10, 16, 16, 4, 24, 10, 16, 16], tab=BLUE)
    lists_sheet(wb, extra={3: INV_CATS, 4: ['Yes', 'No']})
    ws = it
    title(ws, 'Home contents', 'Amounts in the currency set on the Settings sheet')
    cell(ws, 'D2', 'One row per item, or one row per group of small items.', size=9, color=MUTED)
    head(ws, 3, ['Room', 'Item', 'Category', 'Brand / model', 'Serial no.', 'Bought on', 'Qty', 'Price paid (each)', 'Value now (each, optional)', 'Warranty until',
                 'Receipt kept at', 'Photo taken', 'Receipt kept', 'Notes', 'Total paid', 'Value for claim', 'Warranty', 'Age (years)', 'Proof'])
    Q = 'IF($G{r}="",1,N($G{r}))'
    grid(ws, 4, 503, 'ABCDEFGHIJKLMNOPQRS', {c: 'calc' for c in 'OPQRS'}, {'F': DT, 'H': NUM0, 'I': NUM0, 'J': DT, 'O': NUM0, 'P': NUM0, 'R': '0.0'},
         {'O': '=IF($B{r}="","",' + Q + '*N($H{r}))',
          'P': '=IF($B{r}="","",' + Q + '*IF($I{r}="",N($H{r}),N($I{r})))',
          'Q': '=IF($B{r}="","",IF($J{r}="","None noted",IF($J{r}>=Settings!$B$5,"Under warranty","Expired")))',
          'R': '=IF(OR($B{r}="",$F{r}=""),"",ROUND((Settings!$B$5-$F{r})/365.25,1))',
          'S': '=IF($B{r}="","",IF(AND($L{r}="Yes",$M{r}="Yes"),"Complete","Missing"))'})
    T = lambda n: TODAY + datetime.timedelta(days=n)
    fill_rows(ws, 4, [
        ('Living room', 'LED TV 55 inch', 'Electronics', 'Samsung UA55CU8000', 'SN-0X7731', T(-700), 1, 52990, 38000, T(30), 'Blue folder › Bills', 'Yes', 'Yes', 'Wall mount included'),
        ('Living room', 'Sofa set (3+1+1)', 'Furniture', 'Local carpenter', '', T(-1500), 1, 68000, 30000, None, 'Blue folder › Bills', 'Yes', 'Yes', ''),
        ('Living room', 'Split AC 1.5 ton', 'Appliances', 'Daikin FTKF50', 'DK5521980', T(-420), 1, 46500, None, T(1400), 'Email (Croma)', 'Yes', 'Yes', 'Compressor warranty 5 years'),
        ('Kitchen', 'Refrigerator 340 L', 'Appliances', 'LG GL-T382', '210KR0331', T(-900), 1, 36990, 26000, T(-170), 'Blue folder › Bills', 'Yes', 'Yes', ''),
        ('Kitchen', 'Microwave oven', 'Appliances', 'IFB 25BC4', 'IFB77120', T(-1200), 1, 11500, 6000, None, '', 'Yes', 'No', 'Bill not found'),
        ('Kitchen', 'Mixer grinder', 'Appliances', 'Preethi Zodiac', '', T(-300), 1, 9200, None, T(430), 'Phone › Photos › Bills', 'No', 'Yes', ''),
        ('Main bedroom', 'Gold jewellery (set)', 'Jewellery and watches', '22 carat, 38 g', '', datetime.date(2016, 11, 20), 1, 118000, 280000, None, 'Bank locker 214', 'Yes', 'Yes', 'Ask insurer if it must be declared'),
        ('Main bedroom', 'Wardrobe (sliding)', 'Furniture', 'Godrej Interio', '', T(-2000), 1, 42000, 22000, None, 'Blue folder › Bills', 'No', 'Yes', ''),
        ('Study', 'Laptop', 'Electronics', 'Dell Inspiron 14', 'CN-0H5T2', T(-250), 1, 64990, None, T(115), 'Email (Dell)', 'Yes', 'Yes', ''),
        ('Study', 'Books', 'Books and media', 'About 150 books', '', None, 150, 300, None, None, '', 'Yes', 'No', 'Estimated'),
        ('Balcony / store', 'Washing machine', 'Appliances', 'Bosch 7 kg front load', 'WAJ240ZBIN-88', T(-1100), 1, 33990, 20000, T(-30), 'Blue folder › Bills', 'Yes', 'Yes', ''),
        ('Garage / parking', 'Bicycle', 'Sports and bicycles', 'Firefox Rapide', 'FF2210', T(-500), 1, 18500, None, None, '', 'No', 'No', '')])
    dv(ws, 'A4:A503', '=Settings!$D$3:$D$22'); dv(ws, 'C4:C503', '=Lists!$C$1:$C$12'); dv(ws, 'L4:M503', '=Lists!$D$1:$D$2')
    ws.conditional_formatting.add('S4:S503', CellIsRule(operator='equal', formula=['"Missing"'], fill=AMBER))
    ws.freeze_panes = 'C4'
    # Summary
    ws = sm
    title(ws, 'Summary', 'Amounts in the currency set on the Settings sheet')
    I = lambda c: 'Items!$%s$4:$%s$503' % (c, c)
    rows = [('Total value for a claim', '=SUM(%s)' % I('P'), NUM0, True), ('Total price paid', '=SUM(%s)' % I('O'), NUM0, False),
            ('Contents sum insured (Settings)', '=N(Settings!B8)', NUM0, False), ('Value ÷ sum insured', '=IF(N(B5)=0,"",B3/B5)', '0%', True),
            ('Items listed', '=COUNTIF(%s,"?*")' % I('B'), None, False), ('Items with photo and receipt', '=IF(B7=0,"",COUNTIF(%s,"Complete")/B7)' % I('S'), '0%', True),
            ('Items under warranty', '=COUNTIF(%s,"Under warranty")' % I('Q'), None, False)]
    for i, (k, f2, fmt, b) in enumerate(rows, 3):
        label(ws, 'A%d' % i, k, bold=b); cell(ws, 'B%d' % i, f2, 'soft' if b else 'calc', fmt, bold=b, align='right')
    ws.merge_cells('B3:D3'); ws.merge_cells('B4:D4'); ws.merge_cells('B5:D5'); ws.merge_cells('B6:D6'); ws.merge_cells('B7:D7'); ws.merge_cells('B8:D8'); ws.merge_cells('B9:D9')
    red_if(ws, 'B6', 'AND(B6<>"",B6>1)', AMBER)
    head(ws, 11, ['Room', 'Items', 'Price paid', 'Value for claim'])
    for i in range(20):
        r, s = 12 + i, 3 + i
        cell(ws, 'A%d' % r, '=IF(Settings!D%d="","",Settings!D%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A{r}="","",COUNTIF({A},$A{r}))'.format(r=r, A=I('A')), 'calc', align='center')
        cell(ws, 'C%d' % r, '=IF($A{r}="","",SUMIF({A},$A{r},{O}))'.format(r=r, A=I('A'), O=I('O')), 'calc', NUM0)
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIF({A},$A{r},{P}))'.format(r=r, A=I('A'), P=I('P')), 'calc', NUM0)
    head(ws, 11, ['Category', 'Items', 'Price paid', 'Value for claim'], col=6)
    for i, c in enumerate(INV_CATS, 12):
        cell(ws, 'F%d' % i, c, 'box')
        cell(ws, 'G%d' % i, '=COUNTIF({C},F{r})'.format(r=i, C=I('C')), 'calc', align='center')
        cell(ws, 'H%d' % i, '=SUMIF({C},F{r},{O})'.format(r=i, C=I('C'), O=I('O')), 'calc', NUM0)
        cell(ws, 'I%d' % i, '=SUMIF({C},F{r},{P})'.format(r=i, C=I('C'), P=I('P')), 'calc', NUM0)
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Value for claim by category'; ch.legend = None
    ch.add_data(Reference(ws, min_col=9, min_row=11, max_row=23), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=6, min_row=12, max_row=23))
    ch.series[0].graphicalProperties.solidFill = PINK
    ch.height, ch.width = 8.5, 15
    ws.add_chart(ch, 'K3')
    protect(ws)
    save(wb, 'HM-Home-Inventory.xlsx')

MAKERS = [hm_budget, hm_loan, hm_bills, hm_documents, hm_house_help, hm_inventory]
