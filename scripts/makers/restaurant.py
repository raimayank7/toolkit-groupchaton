"""Restaurant & Cloud Kitchen toolkit files (kit id: restaurant).  python -X utf8 scripts/make_files.py --only restaurant"""
from make_files import *

RED = 'DC2626'

# ------------------------------------------------------------------ daily sales & cash register
def rest_sales():
    wb = new_wb('rest-daily-sales-excel')
    howto(wb, first=True)
    st = sheet(wb, 'Settings', [30, 24, 15])
    ds = sheet(wb, 'Daily Sales', [11, 6, 7, 12, 11, 11, 11, 11, 11, 11, 13, 10, 12, 13, 13, 13, 12, 26], tab=RED)
    sm = sheet(wb, 'Month Summary', [24, 15, 9, 13, 14, 15, 16, 24], tab=RED)
    ds.page_setup.orientation = 'landscape'
    title(st, 'Settings', 'Fill once a month. Payment-mode names and commission are used on the other sheets.')
    for i, (k, v, fmt) in enumerate([('Restaurant name', 'Spice Route Family Restaurant', None), ('Month (first day)', MONTH1, 'mmmm yyyy'),
                                     ('Opening cash float in drawer', 2000, INR), ('GST on aggregator commission', 0.18, '0%')], 3):
        label(st, 'A%d' % i, k); cell(st, 'B%d' % i, v, 'in', fmt, align='left')
    head(st, 8, ['Payment mode (rename if you like)', 'Money goes to', 'Commission %'])
    for i, (m, to, c) in enumerate([('Cash', 'Cash drawer', 0), ('UPI', 'Bank (same day)', 0), ('Card', 'Bank (next day)', 0),
                                    ('Swiggy', 'Bank (weekly payout)', 0.25), ('Zomato', 'Bank (weekly payout)', 0.25), ('Other', 'Bank', 0)], 9):
        cell(st, 'A%d' % i, m, 'in'); cell(st, 'B%d' % i, to, 'in'); cell(st, 'C%d' % i, c, 'in', '0.0%')
    st.merge_cells('A16:C18')
    cell(st, 'A16', 'Commission is an estimate: aggregator payouts also deduct payment gateway fees, GST on fees, TCS, TDS and any ads or offers you run. '
         'Match each payout with the statement from the platform.', wrap=True, size=9, color=MUTED)

    ws = ds
    title(ws, 'Daily sales and cash drawer')
    ws['A2'] = '=Settings!B3&"  ·  "&TEXT(Settings!B4,"mmmm yyyy")'; ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Date', 'Day', 'Bills', 'Opening cash'] + [''] * 6 + ['Total sales', 'Average bill', 'Cash paid out', 'Cash deposited / handed over',
                                                              'Expected cash in drawer', 'Counted cash', 'Short (−) / Excess (+)', 'Notes'])
    for j in range(6): ws.cell(4, 5 + j).value = '=Settings!A%d' % (9 + j)
    first = 'DATE(YEAR(Settings!$B$4),MONTH(Settings!$B$4),%d)'
    for n in range(31):
        r = 5 + n
        d = first % (n + 1)
        cell(ws, 'A%d' % r, '=IF(MONTH(%s)=MONTH(Settings!$B$4),%s,"")' % (d, d), 'calc', DATE)
        cell(ws, 'B%d' % r, '=IF($A{r}="","",TEXT($A{r},"ddd"))'.format(r=r), 'calc', align='center')
        cell(ws, 'C%d' % r, None, 'in', align='center')
        cell(ws, 'D%d' % r, ('=IF($A5="","",Settings!$B$5)' if n == 0 else '=IF($A{r}="","",IF(P{p}<>"",P{p},O{p}))'.format(r=r, p=r - 1)), 'calc', INR0)
        for col in 'EFGHIJMNP': cell(ws, '%s%d' % (col, r), None, 'in', INR0)
        cell(ws, 'K%d' % r, '=IF($A{r}="","",SUM(E{r}:J{r}))'.format(r=r), 'calc', INR0, bold=True)
        cell(ws, 'L%d' % r, '=IF(OR($A{r}="",N(C{r})=0),"",K{r}/C{r})'.format(r=r), 'calc', INR0)
        cell(ws, 'O%d' % r, '=IF($A{r}="","",D{r}+N(E{r})-N(M{r})-N(N{r}))'.format(r=r), 'calc', INR0)
        cell(ws, 'Q%d' % r, '=IF(OR($A{r}="",P{r}=""),"",P{r}-O{r})'.format(r=r), 'calc', '+#,##0;−#,##0;0')
        cell(ws, 'R%d' % r, None, 'in')
    # bills, cash, UPI, card, Swiggy, Zomato, other, paid out, deposited, counted cash vs expected, note  (1st of the month may be any weekday)
    sample = [(64, 16200, 22900, 10400, 11900, 11100, 0, 1750, 13500, 0, ''), (38, 8200, 12600, 5100, 6800, 6250, 0, 900, 6000, 0, 'Gas cylinder paid in cash'),
              (35, 7400, 11900, 4800, 5600, 5100, 1200, 650, 6500, 0, 'Party order advance under Other'), (40, 9100, 13850, 5900, 6900, 6400, 0, 1200, 7500, 20, 'Excess ₹20: change not taken by a guest'),
              (44, 9900, 14700, 6300, 7400, 6900, 0, 1100, 8000, 0, ''), (62, 15600, 21400, 9800, 11200, 10400, 0, 2300, 12000, -50, 'Busy dinner; ₹50 short')]
    cash = 2000
    for i, row in enumerate(sample):
        r = 5 + i
        cash = cash + row[1] - row[7] - row[8] + row[9]   # counted cash = expected + difference
        for col, v in zip('CEFGHIJMNR', row[:9] + row[10:]): ws['%s%d' % (col, r)] = v if v != '' else None
        ws['P%d' % r] = cash
    cell(ws, 'A37', 'Month total', 'soft', bold=True); ws.merge_cells('A37:B37')
    for col in 'CEFGHIJKMNQ': cell(ws, '%s37' % col, '=SUM(%s5:%s35)' % (col, col), 'soft', INR0 if col != 'C' else '0', bold=True)
    cell(ws, 'L37', '=IF(N(C37)=0,"",K37/C37)', 'soft', INR0, bold=True)
    ws.conditional_formatting.add('Q5:Q35', CellIsRule(operator='lessThan', formula=['0'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('Q5:Q35', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='DCFCE7')))
    ws.conditional_formatting.add('A5:B35', FormulaRule(formula=['AND($A5<>"",WEEKDAY($A5,2)>5)'], font=Font(bold=True, color=RED)))
    ws.freeze_panes = 'C5'
    ws.print_title_rows = '4:4'
    protect(ws)

    ws = sm
    title(ws, 'Month summary')
    ws['A2'] = '=Settings!B3&"  ·  "&TEXT(Settings!B4,"mmmm yyyy")'; ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Payment mode', 'Sales', 'Share', 'Commission %', 'Commission', 'GST on commission', 'Net to receive', 'Money goes to'])
    for i in range(6):
        r, col = 5 + i, L(5 + i)
        cell(ws, 'A%d' % r, '=Settings!A%d' % (9 + i), 'box')
        cell(ws, 'B%d' % r, "='Daily Sales'!%s37" % col, 'calc', INR0)
        cell(ws, 'C%d' % r, '=IF(N($B$11)=0,"",B{r}/$B$11)'.format(r=r), 'calc', '0.0%', align='center')
        cell(ws, 'D%d' % r, '=N(Settings!C%d)' % (9 + i), 'calc', '0.0%', align='center')
        cell(ws, 'E%d' % r, '=ROUND(B{r}*D{r},2)'.format(r=r), 'calc', INR)
        cell(ws, 'F%d' % r, '=ROUND(E{r}*N(Settings!$B$6),2)'.format(r=r), 'calc', INR)
        cell(ws, 'G%d' % r, '=B{r}-E{r}-F{r}'.format(r=r), 'calc', INR)
        cell(ws, 'H%d' % r, '=Settings!B%d&""' % (9 + i), 'box')
    cell(ws, 'A11', 'Total', 'soft', bold=True)
    for col, fmt in zip('BCEFG', (INR0, '0.0%', INR, INR, INR)):
        cell(ws, '%s11' % col, '=SUM(%s5:%s10)' % (col, col), 'soft', fmt, bold=True)
    label(ws, 'A13', 'Key figures for the month')
    D_ = "'Daily Sales'!"
    figs = [('Bills', '=%sC37' % D_, '0'), ('Average bill value', '=IF(N(B14)=0,"",B11/B14)', INR0),
            ('Days with sales', '=COUNTIF(%sK5:K35,">0")' % D_, '0'), ('Average sales per day', '=IF(N(B16)=0,"",B11/B16)', INR0),
            ('Best day', '=IF(MAX(%sK5:K35)=0,"",INDEX(%sA5:A35,MATCH(MAX(%sK5:K35),%sK5:K35,0)))' % (D_, D_, D_, D_), DATE),
            ('Sales on the best day', '=MAX(%sK5:K35)' % D_, INR0), ('Cash paid out of the counter', '=%sM37' % D_, INR0),
            ('Cash deposited / handed over', '=%sN37' % D_, INR0), ('Total cash short (−) / excess (+)', '=%sQ37' % D_, '+#,##0;−#,##0;0'),
            ('Days the cash was short', '=COUNTIF(%sQ5:Q35,"<0")' % D_, '0')]
    for i, (k, f, fmt) in enumerate(figs, 14):
        cell(ws, 'A%d' % i, k, 'box'); cell(ws, 'B%d' % i, f, 'calc', fmt, align='right')
    ch = BarChart(); ch.type = 'bar'; ch.style = 10; ch.title = 'Sales by payment mode'; ch.legend = None
    ch.add_data(Reference(ws, min_col=2, min_row=4, max_row=10), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=10))
    ch.series[0].graphicalProperties.solidFill = RED
    ch.height, ch.width = 7.5, 14
    ws.add_chart(ch, 'D13')
    protect(ws)
    save(wb, 'Restaurant-Daily-Sales-Register.xlsx')

# ------------------------------------------------------------------ recipe costing & menu pricing
def rest_recipe():
    wb = new_wb('rest-recipe-costing-excel')
    howto(wb, first=True)
    ig = sheet(wb, 'Ingredients', [28, 9, 14, 10, 16, 12, 24])
    rc = sheet(wb, 'Recipes', [26, 26, 12, 8, 14, 13], tab=RED)
    mc = sheet(wb, 'Menu Costing', [26, 14, 10, 12, 11, 12, 13, 12, 10, 13, 12, 15], tab=RED)
    mc.page_setup.orientation = 'landscape'
    title(ig, 'Ingredients', 'Price for ONE purchase unit. Yield = usable part after peeling, cleaning or trimming (100% if nothing is thrown away).')
    head(ig, 3, ['Ingredient', 'Unit', 'Price per unit (₹)', 'Yield %', 'Cost per usable unit (₹)', 'Price date', 'Supplier'])
    grid(ig, 4, 303, 'ABCDEFG', {'E': 'calc'}, {'C': INR, 'D': '0%', 'E': INR, 'F': DATE},
         {'E': '=IF(OR($A{r}="",N($C{r})=0),"",ROUND($C{r}/IF(N($D{r})<=0,1,$D{r}),2))'})
    dv(ig, 'B4:B303', '"kg,L,pc,dozen,pack"')
    ings = [('Paneer', 'kg', 380, 1), ('Butter', 'kg', 520, 1), ('Fresh cream', 'L', 240, 1), ('Tomato', 'kg', 35, 0.9), ('Onion', 'kg', 30, 0.9),
            ('Cashew', 'kg', 900, 1), ('Ginger-garlic paste', 'kg', 140, 1), ('Masala and whole spices', 'kg', 600, 1), ('Cumin seeds', 'kg', 400, 1),
            ('Refined oil', 'L', 150, 1), ('Basmati rice', 'kg', 110, 1), ('Whole urad dal', 'kg', 130, 1), ('Maida', 'kg', 40, 1),
            ('Chicken curry cut', 'kg', 240, 0.95), ('Milk', 'L', 60, 1), ('Curd', 'kg', 80, 1), ('Sugar', 'kg', 44, 1)]
    for i, (n, u, p, y) in enumerate(ings, 4):
        ig['A%d' % i], ig['B%d' % i], ig['C%d' % i], ig['D%d' % i], ig['F%d' % i] = n, u, p, y, MONTH1
        if n in ('Tomato', 'Onion'): ig['G%d' % i] = 'Local mandi'
    ig.freeze_panes = 'A4'
    protect(ig)

    title(rc, 'Recipe lines', 'One line per ingredient. Quantity for the WHOLE batch, in the ingredient\'s unit (150 g = 0.15 kg).')
    head(rc, 3, ['Dish', 'Ingredient', 'Qty for the batch', 'Unit', 'Cost per unit (₹)', 'Line cost (₹)'])
    look = lambda col: '=IF($B{r}="","",IFERROR(INDEX(Ingredients!$%s:$%s,MATCH($B{r},Ingredients!$A:$A,0)),""))' % (col, col)
    grid(rc, 4, 603, 'ABCDEF', {'D': 'calc', 'E': 'calc', 'F': 'calc'}, {'C': '0.000', 'E': INR, 'F': INR},
         {'D': '=IF($B{r}="","",IFERROR(INDEX(Ingredients!$B:$B,MATCH($B{r},Ingredients!$A:$A,0))&"",""))', 'E': look('E'),
          'F': '=IF(OR($C{r}="",N($E{r})=0),"",ROUND(N($C{r})*$E{r},2))'})
    dv(rc, 'A4:A603', "='Menu Costing'!$A$7:$A$106", 'Pick a dish from the Menu Costing sheet')
    dv(rc, 'B4:B603', '=Ingredients!$A$4:$A$303', 'Pick an ingredient from the Ingredients sheet')
    recipes = [('Paneer Butter Masala', [('Paneer', 1.5), ('Butter', 0.2), ('Fresh cream', 0.3), ('Tomato', 2), ('Onion', 0.8), ('Cashew', 0.15),
                                         ('Ginger-garlic paste', 0.1), ('Masala and whole spices', 0.08), ('Refined oil', 0.15)]),
               ('Dal Makhani', [('Whole urad dal', 0.8), ('Butter', 0.25), ('Fresh cream', 0.25), ('Tomato', 1), ('Ginger-garlic paste', 0.08), ('Masala and whole spices', 0.05)]),
               ('Jeera Rice', [('Basmati rice', 1.2), ('Butter', 0.1), ('Cumin seeds', 0.02)]),
               ('Butter Naan', [('Maida', 1.6), ('Curd', 0.2), ('Butter', 0.2), ('Milk', 0.2), ('Sugar', 0.03)]),
               ('Chicken Curry', [('Chicken curry cut', 2.5), ('Onion', 1.2), ('Tomato', 0.8), ('Ginger-garlic paste', 0.12), ('Masala and whole spices', 0.1),
                                  ('Refined oil', 0.25), ('Curd', 0.3)]),
               ('Sweet Lassi', [('Curd', 2.5), ('Sugar', 0.3), ('Milk', 0.5)])]
    r = 4
    for dish, lines in recipes:
        for ing, q in lines:
            rc['A%d' % r], rc['B%d' % r], rc['C%d' % r] = dish, ing, q; r += 1
    rc.freeze_panes = 'A4'
    protect(rc)

    ws = mc
    title(ws, 'Menu costing and pricing', 'Yellow cells: type. Grey cells calculate.')
    label(ws, 'A3', 'Target food cost %'); cell(ws, 'B3', 0.30, 'in', '0%', align='center')
    label(ws, 'A4', 'GST % on the menu price'); cell(ws, 'B4', 0.05, 'in', '0%', align='center')
    ws.merge_cells('D3:L4')
    cell(ws, 'D3', 'Food cost % = cost per portion ÷ menu price before GST. Suggested price = cost per portion ÷ target food cost %, rounded up to the next ₹5. '
         'Red rows cost more than your target.', wrap=True, size=9, color=MUTED)
    head(ws, 6, ['Dish', 'Category', 'Portions per recipe', 'Recipe cost (₹)', 'Extra per plate (₹)', 'Cost per portion (₹)', 'Menu price before GST (₹)',
                 'Price with GST (₹)', 'Food cost %', 'Gross margin per plate (₹)', 'Suggested price (₹)', 'Status'])
    f = {'D': '=IF($A{r}="","",SUMIF(Recipes!$A:$A,$A{r},Recipes!$F:$F))',
         'F': '=IF(OR($A{r}="",N($C{r})=0),"",ROUND($D{r}/$C{r}+N($E{r}),2))',
         'H': '=IF(N($G{r})=0,"",ROUND($G{r}*(1+N($B$4)),2))',
         'I': '=IF(OR($F{r}="",N($G{r})=0),"",$F{r}/$G{r})',
         'J': '=IF(OR($F{r}="",N($G{r})=0),"",$G{r}-$F{r})',
         'K': '=IF(OR($F{r}="",N($B$3)<=0),"",CEILING($F{r}/$B$3,5))',
         'L': '=IF($I{r}="","",IF($I{r}>$B$3,"Food cost high","OK"))'}
    grid(ws, 7, 106, 'ABCDEFGHIJKL', {c: 'calc' for c in 'DFHIJKL'},
         {'C': '0', 'D': INR, 'E': INR, 'F': INR, 'G': INR0, 'H': INR, 'I': '0.0%', 'J': INR, 'K': INR0}, f)
    for i, row in enumerate([('Paneer Butter Masala', 'Main course', 10, 4, 280), ('Dal Makhani', 'Main course', 10, 3, 240), ('Jeera Rice', 'Rice', 10, 2, 180),
                             ('Butter Naan', 'Breads', 20, 1, 50), ('Chicken Curry', 'Main course', 10, 4, 320), ('Sweet Lassi', 'Beverages', 10, 2, 90)], 7):
        for col, v in zip('ABCEG', row): ws['%s%d' % (col, i)] = v
    ws.conditional_formatting.add('A7:L106', FormulaRule(formula=['$L7="Food cost high"'], fill=PatternFill('solid', fgColor='FEE2E2')))
    ws.conditional_formatting.add('L7:L106', CellIsRule(operator='equal', formula=['"Food cost high"'], font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'B7'
    ws.print_title_rows = '6:6'
    protect(ws)
    save(wb, 'Restaurant-Recipe-Costing.xlsx')

# ------------------------------------------------------------------ kitchen stock & wastage
TYPES = ['Purchase', 'Issue to kitchen', 'Wastage', 'Staff meal', 'Return to supplier']
REASONS = ['Spoiled / expired', 'Over-production', 'Cooking error / burnt', 'Dropped / spilled', 'Customer return', 'Trim / peel waste', 'Pest or storage damage', 'Other']

def rest_stock():
    wb = new_wb('rest-stock-wastage-excel')
    howto(wb, first=True)
    it = sheet(wb, 'Items', [28, 14, 8, 12, 12, 11, 14])
    en = sheet(wb, 'Entries', [12, 28, 18, 9, 22, 14, 8, 13], tab=RED)
    sk = sheet(wb, 'Stock', [28, 8, 10, 11, 10, 10, 10, 10, 11, 13, 9, 11, 10, 10, 13], tab=RED)
    wr_ = sheet(wb, 'Wastage Report', [30, 14, 10, 15, 12], tab=RED)
    lists_sheet(wb, extra={5: TYPES, 6: REASONS})
    sk.page_setup.orientation = 'landscape'
    title(it, 'Items', 'Opening stock = what you hold today. Par level = the least you want in stock. Rate = purchase price per unit.')
    head(it, 3, ['Item', 'Category', 'Unit', 'Storage', 'Opening stock', 'Par level', 'Rate (₹ per unit)'])
    grid(it, 4, 303, 'ABCDEFG', {}, {'G': INR})
    dv(it, 'B4:B303', '"Vegetables,Fruit,Dairy,Meat & fish,Dry store,Spices,Oil & ghee,Beverages,Packaging,Cleaning"')
    dv(it, 'C4:C303', '"kg,L,pc,dozen,pack,portion"')
    dv(it, 'D4:D303', '"Chiller,Freezer,Dry store,Room"')
    items = [('Tomato', 'Vegetables', 'kg', 'Chiller', 12, 10, 35), ('Onion', 'Vegetables', 'kg', 'Dry store', 40, 25, 30), ('Coriander leaves', 'Vegetables', 'kg', 'Chiller', 1, 1, 80),
             ('Paneer', 'Dairy', 'kg', 'Chiller', 6, 4, 380), ('Milk', 'Dairy', 'L', 'Chiller', 10, 10, 60), ('Butter', 'Dairy', 'kg', 'Chiller', 5, 3, 520),
             ('Chicken curry cut', 'Meat & fish', 'kg', 'Freezer', 15, 8, 240), ('Basmati rice', 'Dry store', 'kg', 'Dry store', 50, 20, 110),
             ('Refined oil', 'Oil & ghee', 'L', 'Dry store', 30, 15, 150), ('Takeaway container 500 ml', 'Packaging', 'pc', 'Dry store', 400, 200, 6)]
    for i, row in enumerate(items, 4):
        for j, v in enumerate(row): it.cell(i, j + 1, v)
    it.freeze_panes = 'A4'

    title(en, 'Stock entries', 'One row per movement. Wastage rows also need a reason. Unit and value fill in from the Items sheet.')
    head(en, 3, ['Date', 'Item', 'Type', 'Qty', 'Wastage reason', 'Recorded by', 'Unit', 'Value (₹)'])
    grid(en, 4, 1503, 'ABCDEFGH', {'G': 'calc', 'H': 'calc'}, {'A': DATE, 'H': INR},
         {'G': '=IF($B{r}="","",IFERROR(INDEX(Items!$C:$C,MATCH($B{r},Items!$A:$A,0))&"",""))',
          'H': '=IF(OR($B{r}="",$D{r}=""),"",ROUND(N($D{r})*IFERROR(INDEX(Items!$G:$G,MATCH($B{r},Items!$A:$A,0)),0),2))'})
    dv(en, 'B4:B1503', '=Items!$A$4:$A$303'); dv(en, 'C4:C1503', '=Lists!$E$1:$E$5'); dv(en, 'E4:E1503', '=Lists!$F$1:$F$8')
    d = MONTH1
    D = lambda n: d + datetime.timedelta(days=n)
    ents = [(D(0), 'Tomato', 'Purchase', 15, None, 'Store'), (D(0), 'Tomato', 'Issue to kitchen', 11, None, 'Store'), (D(0), 'Paneer', 'Issue to kitchen', 2.5, None, 'Store'),
            (D(0), 'Tomato', 'Wastage', 1.2, 'Spoiled / expired', 'Sunil'), (D(0), 'Paneer', 'Staff meal', 0.3, None, 'Rakesh'),
            (D(1), 'Chicken curry cut', 'Issue to kitchen', 6, None, 'Store'), (D(1), 'Onion', 'Purchase', 25, None, 'Store'), (D(1), 'Basmati rice', 'Purchase', 25, None, 'Store'), (D(1), 'Coriander leaves', 'Wastage', 0.4, 'Spoiled / expired', 'Geeta'),
            (D(1), 'Basmati rice', 'Wastage', 1.5, 'Over-production', 'Sunil'), (D(2), 'Milk', 'Purchase', 20, None, 'Store'), (D(2), 'Milk', 'Issue to kitchen', 18, None, 'Store'),
            (D(2), 'Paneer', 'Wastage', 0.4, 'Cooking error / burnt', 'Imran'), (D(3), 'Onion', 'Wastage', 1.8, 'Trim / peel waste', 'Geeta'),
            (D(3), 'Chicken curry cut', 'Purchase', 10, None, 'Store'), (D(3), 'Takeaway container 500 ml', 'Issue to kitchen', 140, None, 'Store'),
            (D(4), 'Chicken curry cut', 'Wastage', 0.5, 'Customer return', 'Priya'), (D(4), 'Paneer', 'Return to supplier', 1, None, 'Store')]
    for i, row in enumerate(ents, 4):
        for j, v in enumerate(row): en.cell(i, j + 1, v)
    en.freeze_panes = 'A4'

    ws = sk
    title(ws, 'Stock position', 'Closing = opening + purchased − returned − issued − wasted − staff meals. Type a physical count in the yellow Count column.')
    label(ws, 'A3', 'Stock value'); cell(ws, 'J3', '=SUM(J5:J304)', 'soft', INR, bold=True)
    label(ws, 'L3', 'Variance value'); cell(ws, 'O3', '=SUM(O5:O304)', 'soft', INR, bold=True)
    head(ws, 4, ['Item', 'Unit', 'Opening', 'Purchased', 'Returned', 'Issued', 'Wasted', 'Staff meals', 'Closing', 'Value (₹)', 'Par level', 'Status',
                 'Count', 'Variance', 'Variance value (₹)'])
    S = lambda t: '=IF($A{r}="","",SUMIFS(Entries!$D:$D,Entries!$B:$B,$A{r},Entries!$C:$C,"%s"))' % t
    for i in range(300):
        r, s = 5 + i, 4 + i
        rate = 'N(Items!G%d)' % s
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Items!C%d&"")' % (r, s), 'box')
        cell(ws, 'C%d' % r, '=IF($A%d="","",N(Items!E%d))' % (r, s), 'calc')
        for col, t in zip('DEFGH', ['Purchase', 'Return to supplier', 'Issue to kitchen', 'Wastage', 'Staff meal']):
            cell(ws, '%s%d' % (col, r), S(t).format(r=r), 'calc')
        cell(ws, 'I%d' % r, '=IF($A{r}="","",C{r}+D{r}-E{r}-F{r}-G{r}-H{r})'.format(r=r), 'calc', bold=True)
        cell(ws, 'J%d' % r, '=IF($A{r}="","",ROUND(I{r}*{rate},2))'.format(r=r, rate=rate), 'calc', INR)
        cell(ws, 'K%d' % r, '=IF($A%d="","",N(Items!F%d))' % (r, s), 'calc')
        cell(ws, 'L%d' % r, '=IF($A{r}="","",IF(I{r}<=K{r},"Order now","OK"))'.format(r=r), 'calc', align='center')
        cell(ws, 'M%d' % r, None, 'in')
        cell(ws, 'N%d' % r, '=IF(OR($A{r}="",M{r}=""),"",M{r}-I{r})'.format(r=r), 'calc')
        cell(ws, 'O%d' % r, '=IF(N{r}="","",ROUND(N{r}*{rate},2))'.format(r=r, rate=rate), 'calc', INR)
    ws['M5'], ws['M6'] = 14.5, 62.5   # sample counts: tomato 0.3 kg and onion 0.7 kg short of the book stock
    ws.conditional_formatting.add('L5:L304', CellIsRule(operator='equal', formula=['"Order now"'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('N5:O304', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    ws.conditional_formatting.add('I5:I304', CellIsRule(operator='lessThan', formula=['0'], font=Font(bold=True, color='B91C1C')))
    ws.freeze_panes = 'B5'
    ws.print_title_rows = '4:4'
    protect(ws)

    ws = wr_
    title(ws, 'Wastage report', 'Type the from and to dates; everything below follows them.')
    label(ws, 'A3', 'From'); cell(ws, 'B3', MONTH1, 'in', DATE, align='left')
    label(ws, 'A4', 'To'); cell(ws, 'B4', '=EOMONTH(B3,0)', 'in', DATE, align='left')
    per = ',Entries!$A:$A,">="&$B$3,Entries!$A:$A,"<="&$B$4'
    head(ws, 6, ['Wastage reason', 'Entries', 'Value (₹)', 'Share of wastage'])
    for i in range(len(REASONS) + 1):
        r = 7 + i
        crit = '$A%d' % r if i < len(REASONS) else '""'
        cell(ws, 'A%d' % r, ('=Lists!F%d' % (i + 1)) if i < len(REASONS) else 'No reason written', 'box')
        cell(ws, 'B%d' % r, '=COUNTIFS(Entries!$C:$C,"Wastage",Entries!$E:$E,%s,Entries!$B:$B,"<>"%s)' % (crit, per), 'calc', '0', align='center')
        cell(ws, 'C%d' % r, '=SUMIFS(Entries!$H:$H,Entries!$C:$C,"Wastage",Entries!$E:$E,%s%s)' % (crit, per), 'calc', INR)
        cell(ws, 'D%d' % r, '=IF(N($C$16)=0,"",C{r}/$C$16)'.format(r=r), 'calc', '0.0%', align='center')
    cell(ws, 'A16', 'Total wastage', 'soft', bold=True)
    cell(ws, 'B16', '=SUM(B7:B15)', 'soft', '0', bold=True, align='center'); cell(ws, 'C16', '=SUM(C7:C15)', 'soft', INR, bold=True)
    figs = [('Purchases in the period', '=SUMIFS(Entries!$H:$H,Entries!$C:$C,"Purchase"%s)' % per, INR),
            ('Wastage as % of purchases', '=IF(N(B18)=0,"",C16/B18)', '0.0%'),
            ('Staff meals in the period', '=SUMIFS(Entries!$H:$H,Entries!$C:$C,"Staff meal"%s)' % per, INR)]
    for i, (k, f, fmt) in enumerate(figs, 18):
        cell(ws, 'A%d' % i, k, 'box'); cell(ws, 'B%d' % i, f, 'calc', fmt, align='right')
    head(ws, 22, ['Item', 'Qty wasted', 'Unit', 'Value (₹)'])
    for i in range(300):
        r, s = 23 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Items!A%d="","",Items!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A{r}="","",SUMIFS(Entries!$D:$D,Entries!$B:$B,$A{r},Entries!$C:$C,"Wastage"{p}))'.format(r=r, p=per), 'calc', '0.00', align='center')
        cell(ws, 'C%d' % r, '=IF($A%d="","",Items!C%d&"")' % (r, s), 'box', align='center')
        cell(ws, 'D%d' % r, '=IF($A{r}="","",SUMIFS(Entries!$H:$H,Entries!$B:$B,$A{r},Entries!$C:$C,"Wastage"{p}))'.format(r=r, p=per), 'calc', INR)
    ws.conditional_formatting.add('D23:D322', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='FEE2E2')))
    protect(ws)
    save(wb, 'Restaurant-Kitchen-Stock-Wastage.xlsx')

# ------------------------------------------------------------------ staff shift roster
ROLES = ['Manager', 'Head chef', 'Cook', 'Tandoor cook', 'Helper', 'Dishwasher', 'Captain', 'Steward', 'Cashier', 'Delivery rider', 'Cleaner']

def rest_roster():
    wb = new_wb('rest-shift-roster-excel')
    howto(wb, first=True)
    sh = sheet(wb, 'Shifts', [9, 26, 9, 9, 9, 9, 11, 11])
    sf = sheet(wb, 'Staff', [24, 16, 15, 14])
    ro = sheet(wb, 'Roster', [22, 14] + [11] * 7 + [8, 8, 7, 15, 3, 7, 34], tab=RED)
    cv = sheet(wb, 'Coverage', [9, 26] + [11] * 7, tab=RED)
    ro.page_setup.orientation = 'landscape'
    T_ = lambda hh, mm: datetime.time(hh, mm)
    title(sh, 'Shifts', 'Times in 24-hour format (17:00). A shift that ends after midnight is fine. Use Start 2 / End 2 for split shifts.')
    head(sh, 3, ['Code', 'Shift name', 'Start', 'End', 'Start 2', 'End 2', 'Break (min)', 'Paid hours'])
    grid(sh, 4, 23, 'ABCDEFGH', {'H': 'calc'}, {c: 'hh:mm' for c in 'CDEF'} | {'H': '0.0'},
         {'H': '=IF($A{r}="","",IF(OR($C{r}="",$D{r}=""),0,ROUND(MOD($D{r}-$C{r},1)*24+IF(OR($E{r}="",$F{r}=""),0,MOD($F{r}-$E{r},1)*24)-N($G{r})/60,2)))'})
    shifts = [('M', 'Morning', T_(7, 0), T_(15, 0), None, None, 30), ('G', 'General', T_(10, 0), T_(19, 0), None, None, 60),
              ('E', 'Evening', T_(15, 0), T_(23, 30), None, None, 30), ('S', 'Split (lunch + dinner)', T_(10, 30), T_(15, 0), T_(18, 30), T_(23, 0), 0),
              ('N', 'Closing', T_(17, 0), T_(1, 0), None, None, 30), ('OFF', 'Weekly off', None, None, None, None, None), ('L', 'Leave', None, None, None, None, None)]
    for i, row in enumerate(shifts, 4):
        for j, v in enumerate(row): sh.cell(i, j + 1, v)
    protect(sh)

    title(sf, 'Staff', 'Weekly hour limit: 48 is common; check your state\'s shops and establishments rules.')
    head(sf, 3, ['Name', 'Role', 'Phone', 'Weekly hour limit'])
    grid(sf, 4, 63, 'ABCD', {})
    dv(sf, 'B4:B63', '"%s"' % ','.join(ROLES))
    staff = [('Rakesh Kumar', 'Head chef'), ('Sunil Thapa', 'Cook'), ('Imran Ali', 'Tandoor cook'), ('Geeta Devi', 'Helper'), ('Manoj Yadav', 'Dishwasher'),
             ('Priya Nair', 'Captain'), ('Aman Singh', 'Steward'), ('Rohit Verma', 'Steward'), ('Neha Arora', 'Cashier'), ('Vikas Rawat', 'Delivery rider')]
    for i, (n, ro_) in enumerate(staff, 4):
        sf['A%d' % i], sf['B%d' % i], sf['D%d' % i] = n, ro_, 48

    ws = ro
    title(ws, 'Weekly staff roster', 'Pick a shift code for each person and day. Legend on the right.')
    label(ws, 'A3', 'Week starting'); ws.merge_cells('C3:D3')
    cell(ws, 'C3', TODAY - datetime.timedelta(days=TODAY.weekday()), 'in', 'ddd dd-mm-yyyy', align='left')
    head(ws, 5, ['Name', 'Role'] + [''] * 7 + ['Hours', 'Days on', 'Offs', 'Check'])
    for n in range(7):
        c = ws.cell(5, 3 + n); c.value = '=$C$3+%d' % n; c.number_format = 'ddd dd-mmm'
    for i in range(60):
        r, s = 6 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Staff!A%d="","",Staff!A%d)' % (s, s), 'box')
        cell(ws, 'B%d' % r, '=IF($A%d="","",Staff!B%d&"")' % (r, s), 'box', size=9)
        for n in range(7): cell(ws, '%s%d' % (L(3 + n), r), None, 'in', align='center')
        cell(ws, 'J%d' % r, '=IF($A{r}="","",SUMPRODUCT(SUMIF(Shifts!$A$4:$A$23,C{r}:I{r},Shifts!$H$4:$H$23)))'.format(r=r), 'calc', '0.0', align='center')
        cell(ws, 'K%d' % r, '=IF($A{r}="","",COUNTA(C{r}:I{r})-COUNTIF(C{r}:I{r},"OFF")-COUNTIF(C{r}:I{r},"L"))'.format(r=r), 'calc', align='center')
        cell(ws, 'L%d' % r, '=IF($A{r}="","",COUNTIF(C{r}:I{r},"OFF"))'.format(r=r), 'calc', align='center')
        cell(ws, 'M%d' % r, ('=IF($A{r}="","",IF(COUNTA(C{r}:I{r})<7,"Fill all days",IF(L{r}=0,"No weekly off",'
                             'IF(AND(N(Staff!D{s})>0,J{r}>N(Staff!D{s})),"Over hours","OK"))))').format(r=r, s=s), 'calc', align='center')
    week = [('S', 'S', 'S', 'OFF', 'S', 'S', 'S'), ('M', 'M', 'OFF', 'M', 'E', 'E', 'E'), ('E', 'OFF', 'E', 'E', 'E', 'E', 'E'), ('M', 'M', 'M', 'OFF', 'M', 'M', 'M'),
            ('E', 'E', 'E', 'E', 'OFF', 'N', 'N'), ('G', 'G', 'G', 'G', 'OFF', 'E', 'E'), ('M', 'OFF', 'M', 'M', 'M', 'E', 'E'), ('OFF', 'E', 'E', 'E', 'N', 'N', 'N'),
            ('G', 'G', 'OFF', 'G', 'G', 'G', 'L'), ('N', 'N', 'N', 'OFF', 'N', 'N', 'N')]
    for i, codes in enumerate(week):
        for n, c in enumerate(codes): ws.cell(6 + i, 3 + n, c)
    dv(ws, 'C6:I65', '=Shifts!$A$4:$A$23', 'Pick a shift code')
    cell(ws, 'A67', 'Total hours', 'soft', bold=True); cell(ws, 'J67', '=SUM(J6:J65)', 'soft', '0.0', bold=True, align='center')
    ws.conditional_formatting.add('C6:I65', CellIsRule(operator='equal', formula=['"OFF"'], fill=PatternFill('solid', fgColor='E2E8F0'), font=Font(color='64748B')))
    ws.conditional_formatting.add('C6:I65', CellIsRule(operator='equal', formula=['"L"'], fill=PatternFill('solid', fgColor='FFEDD5'), font=Font(bold=True, color='C2410C')))
    ws.conditional_formatting.add('M6:M65', FormulaRule(formula=['AND(M6<>"",M6<>"OK")'], fill=PatternFill('solid', fgColor='FECACA'), font=Font(bold=True, color='B91C1C')))
    head(ws, 5, ['Code', 'Shift and time'], col=15)
    for i in range(20):
        r, s = 6 + i, 4 + i
        cell(ws, 'O%d' % r, '=IF(Shifts!A%d="","",Shifts!A%d)' % (s, s), 'box', align='center', bold=True)
        cell(ws, 'P%d' % r, ('=IF(O{r}="","",Shifts!B{s}&IF(Shifts!C{s}="",""," "&TEXT(Shifts!C{s},"hh:mm")&"-"&TEXT(Shifts!D{s},"hh:mm"))'
                             '&IF(Shifts!E{s}="","",", "&TEXT(Shifts!E{s},"hh:mm")&"-"&TEXT(Shifts!F{s},"hh:mm")))').format(r=r, s=s), 'box', size=9)
    ws.freeze_panes = 'C6'
    ws.print_area = 'A1:P67'
    protect(ws)

    ws = cv
    title(ws, 'Coverage: people on duty')
    ws['A2'] = '="Week starting "&TEXT(Roster!C3,"dd-mm-yyyy")'; ws['A2'].font = Font(size=10, color=MUTED)
    head(ws, 4, ['Code', 'Shift'] + [''] * 7)
    for n in range(7):
        c = ws.cell(4, 3 + n); c.value = '=Roster!$C$3+%d' % n; c.number_format = 'ddd dd-mmm'
    for i in range(20):
        r, s = 5 + i, 4 + i
        cell(ws, 'A%d' % r, '=IF(Shifts!A%d="","",Shifts!A%d)' % (s, s), 'box', align='center', bold=True)
        cell(ws, 'B%d' % r, '=IF($A%d="","",Shifts!B%d&"")' % (r, s), 'box')
        for n in range(7):
            col = L(3 + n)
            cell(ws, '%s%d' % (col, r), '=IF($A{r}="","",COUNTIF(Roster!{c}$6:{c}$65,$A{r}))'.format(r=r, c=col), 'calc', '0;;"–"', align='center')
    cell(ws, 'A26', '', 'soft'); cell(ws, 'B26', 'On duty (all shifts)', 'soft', bold=True)
    for n in range(7):
        col = L(3 + n)
        cell(ws, '%s26' % col, '=COUNTA(Roster!{c}6:{c}65)-COUNTIF(Roster!{c}6:{c}65,"OFF")-COUNTIF(Roster!{c}6:{c}65,"L")'.format(c=col), 'soft', '0', bold=True, align='center')
    head(ws, 28, ['', 'Role on duty'] + [''] * 7)
    for n in range(7):
        c = ws.cell(28, 3 + n); c.value = '=Roster!$C$3+%d' % n; c.number_format = 'ddd dd-mmm'
    for i, role in enumerate(ROLES):
        r = 29 + i
        cell(ws, 'B%d' % r, role, 'box')
        for n in range(7):
            col = L(3 + n)
            cell(ws, '%s%d' % (col, r), ('=COUNTIFS(Roster!$B$6:$B$65,$B{r},Roster!{c}$6:{c}$65,"<>OFF",Roster!{c}$6:{c}$65,"<>L",Roster!{c}$6:{c}$65,"<>")').format(r=r, c=col),
                 'calc', '0;;"–"', align='center')
    ws.conditional_formatting.add('C26:I26', CellIsRule(operator='lessThan', formula=['1'], fill=PatternFill('solid', fgColor='FECACA')))
    protect(ws)
    save(wb, 'Restaurant-Staff-Shift-Roster.xlsx')

# ------------------------------------------------------------------ hygiene checklist & temperature log (Word)
DAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

def rest_hygiene():
    d = doc_new('rest-hygiene-checklist-word')
    letterhead(d, '[Restaurant Name]', '[Address] · Phone [ ] · FSSAI licence / registration no. [ ]', gst=False)
    para(d, 'KITCHEN HYGIENE CHECKLIST', bold=True, size=15, align='c', space=6)
    kv(d, [('Week from', '[DD-MM-YYYY]'), ('To', '[DD-MM-YYYY]'), ('Food safety supervisor', '[Name]'), ('Manager', '[Name]')])
    para(d, 'Tick ✓ when done, ✗ when not done and write what was done about it in the corrective action table. The shift supervisor initials each day.', size=9, italic=True, space=6)
    sections = [
        ('1. Before opening', [
            'Staff in clean uniform or apron, hair covered, nails short, no rings or watches on hands',
            'Nobody on duty with fever, diarrhoea, vomiting, jaundice or an open wound (cuts covered with a waterproof dressing)',
            'Hand-wash basin has running water, soap and paper towels or a clean cloth',
            'Floors, platforms and drains clean and dry; no insects, rodents or droppings seen',
            'Fridge and freezer temperatures read and written on the temperature log',
            'Drinking-water filter working; ice made only from potable water',
            'Cleaning chemicals labelled and kept away from food']),
        ('2. Receiving and storage', [
            'Supplies from known suppliers; packets intact, within use-by date, chilled items still cold',
            'Raw meat and fish stored below and apart from cooked and ready-to-eat food',
            'Veg and non-veg kept apart, with separate chopping boards and knives',
            'Containers covered and labelled with item and date; first in, first out',
            'Dry store on racks off the floor, dry and ventilated',
            'Expired or spoiled stock removed and written in the wastage register']),
        ('3. During service', [
            'Hands washed before handling food and after raw meat, toilet, waste, phone or cash',
            'Meat, poultry and eggs cooked through (about 75 °C in the centre)',
            'Hot food held above 60 °C; cold food at 5 °C or below; nothing left out over 2 hours',
            'Fruit and vegetables washed in potable water before cutting',
            'Frying oil changed before it turns dark, foams or smokes (FSSAI limit: total polar compounds 25%)',
            'Tasting only with a clean spoon; no eating, smoking, spitting or chewing in the kitchen',
            'Delivery food in food-grade, sealed containers marked veg or non-veg']),
        ('4. Closing', [
            'Leftovers cooled quickly, covered, labelled and refrigerated, or thrown away',
            'Utensils, boards and equipment washed, sanitised and dried',
            'Platforms, burners, tandoor area and floors cleaned; drains cleared',
            'Waste separated (wet and dry), bins emptied and closed, bin area clean',
            'Fridge and freezer temperatures read again and written',
            'Gas valves off; doors closed and windows screened against pests'])]
    W = [8.6] + [1.2] * 7
    for name, checks in sections:
        para(d, name, bold=True, color=RED, space=2)
        table(d, ['Check'] + DAYS, [[c] + [''] * 7 for c in checks], W)
    table(d, ['Supervisor initials'] + DAYS, [[''] * 8], W)
    para(d, '5. Weekly and monthly checks', bold=True, color=RED, space=2)
    table(d, ['Check', 'How often', 'Date done', 'By', 'Remarks'], [[c, f, '', '', ''] for c, f in [
        ('Deep clean of walls, ceiling, exhaust hood and filters', 'Weekly'), ('Fridges and freezers defrosted and deep cleaned', 'Weekly'),
        ('Grease trap cleaned', 'Weekly'), ('Pest control by an approved agency, report filed', 'Monthly or as per contract'),
        ('Water tanks cleaned', 'As planned (e.g. every 6 months)'), ('Drinking water tested at a recognised lab', 'As your licence requires'),
        ('Thermometers checked (ice water reads 0 °C)', 'Monthly'), ('Medical fitness certificates of food handlers on file', 'Yearly'),
        ('Trained food safety supervisor (FoSTaC) on the premises', 'Every shift'), ('FSSAI licence valid; Food Safety Display Board visible to guests', 'Monthly'),
        ('Fire extinguisher and first-aid box checked', 'Monthly'), ('Hygiene refresher for all kitchen staff', 'Monthly')]], [6.6, 3.6, 2.2, 1.8, 2.8])
    para(d, '6. Corrective action', bold=True, color=RED, space=2)
    table(d, ['Date', 'What was wrong', 'Action taken', 'By'], [], [2.4, 6, 6.2, 2.4], fill_rows=5)
    para(d, 'Checked by (manager): ____________________     Date: ____________', space=10)
    para(d, 'Based on common good hygiene and sanitary practices for food businesses (FSSAI, Schedule 4). This is a working checklist, not an official FSSAI form; '
            'follow your licence conditions and the latest FSSAI rules.', size=8, italic=True, color=MUTED)
    dsave(d, 'Restaurant-Kitchen-Hygiene-Checklist.docx')

    d = doc_new('rest-hygiene-checklist-word')
    s = d.sections[0]; s.left_margin = s.right_margin = Cm(1.5); s.top_margin = s.bottom_margin = Cm(1.4)
    para(d, '[Restaurant Name]', bold=True, size=16, align='c', color=BLUE, space=0)
    para(d, 'FSSAI licence / registration no. [ ]', size=9, align='c', space=4)
    para(d, 'TEMPERATURE LOG', bold=True, size=14, align='c', space=4)
    kv(d, [('Month', '[Month YYYY]'), ('Fridge 1', '[e.g. Dairy chiller]'), ('Fridge 2', '[e.g. Veg chiller]'), ('Freezer', '[e.g. Meat freezer]')])
    para(d, 'Targets: fridges 5 °C or below · freezer −18 °C or below · hot holding above 60 °C · meat and poultry cooked to about 75 °C in the centre. '
            'Write the reading (for example 3 °C), not a tick. AM = at opening, PM = before closing.', size=9, space=4)
    table(d, ['Day', 'Fridge 1 AM', 'Fridge 1 PM', 'Fridge 2 AM', 'Fridge 2 PM', 'Freezer AM', 'Freezer PM', 'Hot holding °C', 'Cooked item and core °C', 'Initials'],
          [[str(n)] + [''] * 9 for n in range(1, 32)], [1.0, 1.55, 1.55, 1.55, 1.55, 1.55, 1.55, 1.6, 4.2, 1.5])
    para(d, 'If a fridge reads above 5 °C twice in a row or the freezer above −15 °C: move the food to another unit, check the door seal and thermostat, call for service, '
            'decide whether any food must be thrown away, and write it in the corrective action table of the hygiene checklist.', size=9, space=6)
    para(d, 'Checked by (manager): ____________________     Date: ____________', space=4)
    para(d, 'A working log based on common FSSAI good hygiene practices; not an official FSSAI form.', size=8, italic=True, color=MUTED)
    dsave(d, 'Restaurant-Temperature-Log.docx')

MAKERS = [rest_sales, rest_recipe, rest_stock, rest_roster, rest_hygiene]
