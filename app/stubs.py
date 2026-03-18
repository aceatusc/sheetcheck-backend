"""
stubs.py — Three canned demo tasks for the SheetCheck landing page.

Each stub is a complete set of segments that can be triggered without
hitting the LLM. Activated by sending "stub:KEY" as the message.
"""

# ── Demo 1: P&L Dashboard (original) ─────────────────────────────────────

STUB_PNL = [
    {
        "id": "seg-1", "description": "Write header row labels",
        "sheet_context": ["A1:E1"], "predecessors": [],
        "explanation": "Creates five column headers — Month, Revenue, Expenses, Profit, Growth % — in row 1.",
        "qa_pairs": [
            {"q":"Why 5 columns?","a":"Revenue, Expenses, and Profit give a complete P&L; Growth % adds a trend indicator."},
            {"q":"Why start in A1?","a":"Top-left anchors the table and keeps formulas simple."},
        ],
        "edit_suggestions": ["Add more columns like Net Margin","Rename Growth % to MoM Growth","Start table at a different cell"],
        "parameters": [],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A1:E1").values=[["Month","Revenue","Expenses","Profit","Growth %"]]; await ctx.sync(); });',
    },
    {
        "id": "seg-2", "description": "Style header row",
        "sheet_context": ["A1:E1"], "predecessors": ["seg-1"],
        "explanation": "Applies dark background with blue bold text to A1:E1.",
        "qa_pairs": [{"q":"Why dark background?","a":"High contrast helps users quickly identify column names."}],
        "edit_suggestions": ["Use a lighter header background","Change the font color","Increase the font size"],
        "parameters": [
            {"label":"Background color","key":"#1a1d27","value":"#1a1d27","type":"color"},
            {"label":"Font color","key":"#4f8ef7","value":"#4f8ef7","type":"color"},
            {"label":"Font size","key":"11","value":11,"type":"number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const h=s.getRange("A1:E1"); h.format.fill.color="#1a1d27"; h.format.font.color="#4f8ef7"; h.format.font.bold=true; h.format.font.size=11; h.format.horizontalAlignment="Center"; await ctx.sync(); });',
    },
    {
        "id": "seg-3", "description": "Fill in monthly data rows",
        "sheet_context": ["A2:E7"], "predecessors": ["seg-1"],
        "explanation": "Writes six months of raw data into A2:E7. Column E left blank for growth formulas.",
        "qa_pairs": [{"q":"Why leave E blank?","a":"Growth % needs the previous month to exist first."}],
        "edit_suggestions": ["Extend to 12 months","Change the starting revenue values","Add a different starting month"],
        "parameters": [
            {"label":"Jan Revenue","key":"142000","value":142000,"type":"number"},
            {"label":"Jan Expenses","key":"98000","value":98000,"type":"number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:E7").values=[["Jan",142000,98000,44000,""],["Feb",158000,104000,54000,""],["Mar",175000,110000,65000,""],["Apr",163000,107000,56000,""],["May",191000,115000,76000,""],["Jun",210000,121000,89000,""]]; await ctx.sync(); });',
    },
    {
        "id": "seg-4", "description": "Add Growth % formulas",
        "sheet_context": ["E3:E7","D2:D7"], "predecessors": ["seg-3"],
        "explanation": "Inserts IFERROR month-over-month profit growth formulas in E3:E7.",
        "qa_pairs": [
            {"q":"Why skip E2?","a":"January has no prior month; leaving it blank is cleaner."},
            {"q":"Why IFERROR?","a":"Protects against division-by-zero if any profit value is 0."},
        ],
        "edit_suggestions": ["Show growth as Revenue % instead","Format as integer %","Include E2 with N/A label"],
        "parameters": [{"label":"Number format","key":"0.0%","value":"0.0%","type":"text"}],
        "code": "await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('E3:E7').formulas=[['=IFERROR((D3-D2)/D2,\"\")'],['=IFERROR((D4-D3)/D3,\"\")'],['=IFERROR((D5-D4)/D4,\"\")'],['=IFERROR((D6-D5)/D5,\"\")'],['=IFERROR((D7-D6)/D6,\"\")']];\ns.getRange('E3:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });",
    },
    {
        "id": "seg-5", "description": "Format currency columns",
        "sheet_context": ["B2:D7"], "predecessors": ["seg-3"],
        "explanation": "Applies $#,##0 number format to Revenue, Expenses, and Profit.",
        "qa_pairs": [{"q":"Why $#,##0?","a":"Financial summaries round to whole dollars for readability."}],
        "edit_suggestions": ["Show cents with $#,##0.00","Use € for Euros","Remove the currency symbol"],
        "parameters": [{"label":"Currency format","key":"$#,##0","value":"$#,##0","type":"select","options":["$#,##0","$#,##0.00","€#,##0","£#,##0","#,##0"]}],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("B2:D7").numberFormat=[["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"]]; await ctx.sync(); });',
    },
    {
        "id": "seg-6", "description": "Zebra-stripe rows",
        "sheet_context": ["A2:E7"], "predecessors": ["seg-3"],
        "explanation": "Alternates row backgrounds for readability.",
        "qa_pairs": [{"q":"Why alternating colors?","a":"Zebra striping reduces eye-tracking errors on wide rows."}],
        "edit_suggestions": ["Use a stronger stripe color","Apply stripes to the header too","Use a warm color palette"],
        "parameters": [
            {"label":"Even row color","key":"#f5f7ff","value":"#f5f7ff","type":"color"},
            {"label":"Odd row color","key":"#ffffff","value":"#ffffff","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=7;i++){s.getRange("A"+i+":E"+i).format.fill.color=i%2===0?"#f5f7ff":"#ffffff";} await ctx.sync(); });',
    },
    {
        "id": "seg-7", "description": "Colour-code Profit column",
        "sheet_context": ["D2:D7"], "predecessors": ["seg-3","seg-5"],
        "explanation": "Colours font green+bold for profit ≥ threshold, red otherwise.",
        "qa_pairs": [{"q":"Why bold for high values?","a":"Double encoding helps users with color vision deficiency."}],
        "edit_suggestions": ["Change the profit threshold","Use different colors","Apply to a wider range"],
        "parameters": [
            {"label":"Profit threshold","key":"60000","value":60000,"type":"number"},
            {"label":"High color","key":"#1a7a4a","value":"#1a7a4a","type":"color"},
            {"label":"Low color","key":"#b94040","value":"#b94040","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const p=s.getRange("D2:D7"); p.load("values"); await ctx.sync(); p.values.forEach((row,i)=>{const c=s.getRange("D"+(i+2)); c.format.font.color=row[0]>=60000?"#1a7a4a":"#b94040"; c.format.font.bold=row[0]>=60000;}); await ctx.sync(); });',
    },
    {
        "id": "seg-8", "description": "Add totals row and auto-fit",
        "sheet_context": ["A8:E8","A1:E8"], "predecessors": ["seg-3","seg-5","seg-6"],
        "explanation": "Appends a dark TOTAL row with SUM formulas, then auto-fits all columns.",
        "qa_pairs": [
            {"q":"Why autofitColumns last?","a":"All data must exist before fitting."},
            {"q":"Why SUM not SUBTOTAL?","a":"For a simple 6-row table SUM is clearer and more expected."},
        ],
        "edit_suggestions": ["Change TOTAL label to GRAND TOTAL","Use a different totals row color","Add an average row"],
        "parameters": [
            {"label":"Row label","key":"TOTAL","value":"TOTAL","type":"select","options":["TOTAL","GRAND TOTAL","SUM","NET"]},
            {"label":"Background color","key":"#1a1d27","value":"#1a1d27","type":"color"},
            {"label":"Font color","key":"#ffffff","value":"#ffffff","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A8").values=[["TOTAL"]]; s.getRange("B8:D8").formulas=[["=SUM(B2:B7)","=SUM(C2:C7)","=SUM(D2:D7)"]]; s.getRange("B8:D8").numberFormat=[["$#,##0","$#,##0","$#,##0"]]; const t=s.getRange("A8:E8"); t.format.fill.color="#1a1d27"; t.format.font.color="#ffffff"; t.format.font.bold=true; s.getRange("A1:E8").getEntireColumn().format.autofitColumns(); await ctx.sync(); });',
    },
]

# ── Demo 2: Tax Filing Summary ────────────────────────────────────────────
# Designed to showcase every SheetCheck feature:
#   Parameters: number (threshold), select (tax bracket), color (status colors)
#   Edit + branching: change tax rate mid-chain → new branch in graph
#   Ask panel: explain ROUND() formula choice
#   Step-back: undo to before tax formulas and try a different rate
#   Rubric gate: meaningful hard/soft requirements for a tax summary

STUB_SALES = [
    {
        "id": "seg-1", "description": "Write tax summary headers",
        "sheet_context": ["A1:F1"], "predecessors": [],
        "explanation": "Creates six column headers: Category, Description, Gross Amount, Tax Rate, Tax Owed, Status.",
        "qa_pairs": [
            {"q":"Why include both Category and Description?","a":"Category groups income types for the summary; Description holds the specific source name."},
            {"q":"Why a Status column?","a":"Quickly shows whether tax has been withheld, is still owed, or generates a refund."},
        ],
        "edit_suggestions": ["Add a Deductions column","Add a Notes column","Split Tax Owed into Federal and State"],
        "parameters": [],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A1:F1").values=[["Category","Description","Gross Amount","Tax Rate","Tax Owed","Status"]]; await ctx.sync(); });',
    },
    {
        "id": "seg-2", "description": "Style header row",
        "sheet_context": ["A1:F1"], "predecessors": ["seg-1"],
        "explanation": "Applies a dark navy background with gold bold text — a professional government-form aesthetic.",
        "qa_pairs": [
            {"q":"Why navy and gold?","a":"These colours evoke official tax document styling and distinguish this from typical financial dashboards."},
        ],
        "edit_suggestions": ["Use a darker background","Switch to a grey professional look","Use IRS blue (#003087)"],
        "parameters": [
            {"label":"Background color","key":"#1a2744","value":"#1a2744","type":"color"},
            {"label":"Font color","key":"#f5c842","value":"#f5c842","type":"color"},
            {"label":"Font size","key":"11","value":11,"type":"number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const h=s.getRange("A1:F1"); h.format.fill.color="#1a2744"; h.format.font.color="#f5c842"; h.format.font.bold=true; h.format.font.size=11; h.format.horizontalAlignment="Center"; await ctx.sync(); });',
    },
    {
        "id": "seg-3", "description": "Fill in income sources",
        "sheet_context": ["A2:D8"], "predecessors": ["seg-1"],
        "explanation": "Populates seven income rows across four categories: Employment, Freelance, Investment, and Rental.",
        "qa_pairs": [
            {"q":"Why seven rows?","a":"Covers the most common individual income types without overwhelming the template."},
            {"q":"Why leave Tax Owed blank?","a":"Tax Owed is calculated by formula in the next step once rates are confirmed."},
        ],
        "edit_suggestions": ["Add a retirement income row","Change the income amounts","Add a business income category"],
        "parameters": [
            {"label":"Primary salary","key":"95000","value":95000,"type":"number"},
            {"label":"Freelance income","key":"28000","value":28000,"type":"number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:D8").values=[["Employment","W-2 Salary",95000,0.22],["Employment","Bonus",18000,0.22],["Freelance","Consulting",28000,0.24],["Freelance","Design work",9500,0.24],["Investment","Dividends",4200,0.15],["Investment","Capital Gains",11000,0.15],["Rental","Property rent",24000,0.22]]; await ctx.sync(); });',
    },
    {
        "id": "seg-4", "description": "Apply tax rate formatting",
        "sheet_context": ["D2:D8"], "predecessors": ["seg-3"],
        "explanation": "Formats column D as percentage and highlights each bracket with a subtle background so the three rates stand out.",
        "qa_pairs": [
            {"q":"Why colour-code the brackets?","a":"At a glance the user can see which rows share the same rate — important when verifying entries."},
            {"q":"Why 22%, 24%, 15%?","a":"These map to the most common 2024 federal brackets for middle-income filers."},
        ],
        "edit_suggestions": ["Change the 24% bracket to 32%","Add a 37% top bracket row","Make the bracket colors more distinct"],
        "parameters": [
            {"label":"Standard bracket (22%)","key":"#eef4ff","value":"#eef4ff","type":"color"},
            {"label":"High bracket (24%)","key":"#fff8e6","value":"#fff8e6","type":"color"},
            {"label":"Capital gains bracket (15%)","key":"#eefff4","value":"#eefff4","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("D2:D8").numberFormat="0%"; s.getRange("A2:F3").format.fill.color="#eef4ff"; s.getRange("A4:F5").format.fill.color="#fff8e6"; s.getRange("A6:F7").format.fill.color="#eefff4"; s.getRange("A8:F8").format.fill.color="#eef4ff"; await ctx.sync(); });',
    },
    {
        "id": "seg-5", "description": "Calculate Tax Owed with ROUND",
        "sheet_context": ["E2:E8"], "predecessors": ["seg-3"],
        "explanation": "Fills E2:E8 with =ROUND(C*D,2) formulas so each row shows the exact tax liability rounded to cents.",
        "qa_pairs": [
            {"q":"Why ROUND to 2 decimals?","a":"Tax authorities require cent-level precision; floating-point arithmetic without ROUND can produce $0.001 rounding errors that fail validation."},
            {"q":"Why not just C*D?","a":"Bare multiplication accumulates floating-point error across rows; ROUND ensures the SUM of individual amounts matches the rounded total."},
        ],
        "edit_suggestions": ["Add a 10% self-employment tax surcharge on Freelance rows","Use ROUNDUP instead of ROUND","Apply a $2000 standard deduction before calculating"],
        "parameters": [
            {"label":"Rounding decimals","key":"2","value":2,"type":"select","options":["0","1","2"]},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=8;i++){ s.getRange("E"+i).formulas=[["=ROUND(C"+i+"*D"+i+",2)"]]; } s.getRange("C2:C8").numberFormat="$#,##0"; s.getRange("E2:E8").numberFormat="$#,##0.00"; await ctx.sync(); });',
    },
    {
        "id": "seg-6", "description": "Fill Status column with withholding logic",
        "sheet_context": ["F2:F8"], "predecessors": ["seg-5"],
        "explanation": "Labels each row: 'Withheld' if the source typically withholds tax, 'Owed' if the filer must pay, 'Refund' if overpaid. Colour-codes each status.",
        "qa_pairs": [
            {"q":"Why hard-code Withheld for W-2?","a":"W-2 employers are legally required to withhold; the sheet correctly reflects that the liability is pre-paid."},
            {"q":"Why colour the status cells?","a":"A tax preparer reviewing the sheet can instantly see which rows require action vs. which are already settled."},
        ],
        "edit_suggestions": ["Add a 'Partial' status for estimated payments","Change the Owed color to orange","Mark all Freelance rows as Owed"],
        "parameters": [
            {"label":"Withheld color","key":"#d4edda","value":"#d4edda","type":"color"},
            {"label":"Owed color","key":"#fdecea","value":"#fdecea","type":"color"},
            {"label":"Withheld font","key":"#1a7a4a","value":"#1a7a4a","type":"color"},
            {"label":"Owed font","key":"#b94040","value":"#b94040","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const statuses=[["Withheld"],["Withheld"],["Owed"],["Owed"],["Owed"],["Owed"],["Withheld"]]; s.getRange("F2:F8").values=statuses; [2,3,8].forEach(r=>{ s.getRange("F"+r).format.fill.color="#d4edda"; s.getRange("F"+r).format.font.color="#1a7a4a"; s.getRange("F"+r).format.font.bold=true; }); [4,5,6,7].forEach(r=>{ s.getRange("F"+r).format.fill.color="#fdecea"; s.getRange("F"+r).format.font.color="#b94040"; s.getRange("F"+r).format.font.bold=true; }); await ctx.sync(); });',
    },
    {
        "id": "seg-7", "description": "Add summary section and auto-fit",
        "sheet_context": ["A10:F12","A1:F8"], "predecessors": ["seg-5","seg-6"],
        "explanation": "Appends a three-row summary below the data: Total Income, Total Tax Owed, and Effective Rate. Auto-fits columns.",
        "qa_pairs": [
            {"q":"Why show effective rate?","a":"The marginal bracket rate is misleading — the effective rate (total tax ÷ total income) is what the filer actually pays and is more useful for planning."},
            {"q":"Why leave row 9 blank?","a":"A visual separator between the detail rows and the summary block improves readability."},
        ],
        "edit_suggestions": ["Add a Total Tax Withheld row","Show the refund or amount owed net","Add a tax planning target rate row"],
        "parameters": [
            {"label":"Summary background","key":"#1a2744","value":"#1a2744","type":"color"},
            {"label":"Summary font","key":"#f5c842","value":"#f5c842","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A10").values=[["SUMMARY"]]; s.getRange("A10:F10").format.fill.color="#1a2744"; s.getRange("A10:F10").format.font.color="#f5c842"; s.getRange("A10:F10").format.font.bold=true; s.getRange("A11:B11").values=[["Total Income","=SUM(C2:C8)"]]; s.getRange("A12:B12").values=[["Total Tax Owed","=SUM(E2:E8)"]]; s.getRange("A13:B13").values=[["Effective Rate","=B12/B11"]]; s.getRange("B11:B12").numberFormat="$#,##0.00"; s.getRange("B13").numberFormat="0.00%"; s.getRange("A11:B13").format.font.bold=true; s.getRange("A1:F13").getEntireColumn().format.autofitColumns(); await ctx.sync(); });',
    },
]

# ── Demo 3: Inventory Summary ─────────────────────────────────────────────

STUB_INVENTORY = [
    {
        "id": "seg-1", "description": "Write inventory table headers",
        "sheet_context": ["A1:G1"], "predecessors": [],
        "explanation": "Creates headers: SKU, Product, Category, Stock, Reorder Point, Unit Cost, Stock Value.",
        "qa_pairs": [{"q":"Why include Reorder Point?","a":"It lets the sheet flag items that need restocking automatically."}],
        "edit_suggestions": ["Add a Supplier column","Remove Unit Cost","Add a Last Updated column"],
        "parameters": [],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A1:G1").values=[["SKU","Product","Category","Stock","Reorder Pt","Unit Cost","Stock Value"]]; await ctx.sync(); });',
    },
    {
        "id": "seg-2", "description": "Style header row",
        "sheet_context": ["A1:G1"], "predecessors": ["seg-1"],
        "explanation": "Applies a deep purple header with white bold text.",
        "qa_pairs": [{"q":"Why purple?","a":"Inventory themes often use purple to distinguish from finance (blue) and sales (teal)."}],
        "edit_suggestions": ["Use a different header color","Make the header taller","Right-align numeric headers"],
        "parameters": [
            {"label":"Background color","key":"#2d1b69","value":"#2d1b69","type":"color"},
            {"label":"Font color","key":"#f3e8ff","value":"#f3e8ff","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const h=s.getRange("A1:G1"); h.format.fill.color="#2d1b69"; h.format.font.color="#f3e8ff"; h.format.font.bold=true; h.format.horizontalAlignment="Center"; await ctx.sync(); });',
    },
    {
        "id": "seg-3", "description": "Fill in product data",
        "sheet_context": ["A2:G9"], "predecessors": ["seg-1"],
        "explanation": "Populates 8 products across 3 categories with stock levels and unit costs.",
        "qa_pairs": [{"q":"Why 3 categories?","a":"Enough variety to demonstrate category filtering without overwhelming the demo."}],
        "edit_suggestions": ["Add more products","Change the categories","Update the stock levels"],
        "parameters": [
            {"label":"Reorder threshold","key":"20","value":20,"type":"number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:F9").values=[["SKU001","Wireless Mouse","Electronics",45,20,29.99],["SKU002","USB-C Cable","Electronics",8,15,9.99],["SKU003","Desk Lamp","Furniture",32,10,45.00],["SKU004","Office Chair","Furniture",5,3,299.99],["SKU005","Notebook A5","Stationery",120,30,4.99],["SKU006","Ballpoint Pens","Stationery",6,25,1.99],["SKU007","Monitor Stand","Furniture",18,5,79.99],["SKU008","Webcam HD","Electronics",22,10,89.99]]; await ctx.sync(); });',
    },
    {
        "id": "seg-4", "description": "Calculate Stock Value column",
        "sheet_context": ["G2:G9"], "predecessors": ["seg-3"],
        "explanation": "Fills G2:G9 with Stock × Unit Cost formulas.",
        "qa_pairs": [{"q":"Why a formula not a static value?","a":"Formulas update automatically when stock or cost changes."}],
        "edit_suggestions": ["Include a discount factor","Round to 2 decimal places","Show value in thousands"],
        "parameters": [],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=9;i++){s.getRange("G"+i).formulas=[["=D"+i+"*F"+i]];} await ctx.sync(); });',
    },
    {
        "id": "seg-5", "description": "Highlight low-stock items",
        "sheet_context": ["A2:G9"], "predecessors": ["seg-3"],
        "explanation": "Colours rows red where Stock ≤ Reorder Point to flag items needing restocking.",
        "qa_pairs": [{"q":"Why whole-row highlighting?","a":"Easier to scan than just coloring a single cell when you have many columns."}],
        "edit_suggestions": ["Use yellow for warning instead","Add a STATUS text column","Change the threshold logic"],
        "parameters": [
            {"label":"Low stock color","key":"#fdecea","value":"#fdecea","type":"color"},
            {"label":"Low stock font","key":"#b94040","value":"#b94040","type":"color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const r=s.getRange("A2:F9"); r.load("values"); await ctx.sync(); r.values.forEach((row,i)=>{ if(row[3]<=row[4]){ const lr=s.getRange("A"+(i+2)+":G"+(i+2)); lr.format.fill.color="#fdecea"; lr.format.font.color="#b94040"; }}); await ctx.sync(); });',
    },
    {
        "id": "seg-6", "description": "Format currency and auto-fit",
        "sheet_context": ["F2:G9","A1:G9"], "predecessors": ["seg-4","seg-5"],
        "explanation": "Formats Unit Cost and Stock Value as currency, then auto-fits all columns.",
        "qa_pairs": [{"q":"Why format last?","a":"Formatting after data entry avoids Excel auto-converting formatted cells."}],
        "edit_suggestions": ["Use a different currency","Add borders between columns","Bold the Stock Value column"],
        "parameters": [
            {"label":"Currency format","key":"$#,##0.00","value":"$#,##0.00","type":"select","options":["$#,##0.00","$#,##0","€#,##0.00","£#,##0.00"]},
        ],
        "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("F2:G9").numberFormat="$#,##0.00"; s.getRange("A1:G9").getEntireColumn().format.autofitColumns(); await ctx.sync(); });',
    },
]

# ── Registry ──────────────────────────────────────────────────────────────

STUBS = {
    "pnl":       STUB_PNL,
    "sales":     STUB_SALES,
    "inventory": STUB_INVENTORY,
}

STUB_META = {
    "pnl":       {"label": "📊 P&L Dashboard",    "description": "Monthly revenue, expenses & profit with growth % and colour coding"},
    "sales":     {"label": "🧾 Tax Filing",       "description": "Income sources with tax brackets, ROUND formulas, withholding status and effective rate summary"},
    "inventory": {"label": "📦 Inventory Summary", "description": "Product stock levels with low-stock alerts and value calculations"},
}

# ── Per-stub rubrics and verify results ──────────────────────────────────

STUB_RUBRICS = {
    "pnl": {
        "stub_key": "pnl",
        "hard_requirements": [
            {"id":"h1","label":"Header row with Month, Revenue, Expenses, Profit, Growth %","checked":False},
            {"id":"h2","label":"6 months of data in rows 2–7","checked":False},
            {"id":"h3","label":"Currency format on Revenue, Expenses, Profit","checked":False},
        ],
        "soft_requirements": [
            {"id":"s1","label":"Visual hierarchy between header and data rows","checked":False},
            {"id":"s2","label":"Growth % column shows trend direction","checked":False},
            {"id":"s3","label":"Totals row at the bottom","checked":False},
            {"id":"s4","label":"Column widths readable without scrolling","checked":False},
        ],
    },
    "sales": {
        "stub_key": "sales",
        "hard_requirements": [
            {"id":"h1","label":"Headers include at least Category, Description, Gross Amount, Tax Rate, Tax Owed, Status","checked":False},
        ],
        "soft_requirements": [
            {"id":"s1","label":"Tax brackets visually colour-coded by rate","checked":False},
            {"id":"s2","label":"Status column distinguishes Withheld vs Owed rows","checked":False},
        ],
    },
    "inventory": {
        "stub_key": "inventory",
        "hard_requirements": [
            {"id":"h1","label":"Headers include SKU, Product, Category, Stock, Reorder Pt","checked":False},
            {"id":"h2","label":"Stock Value column uses Stock × Unit Cost formula","checked":False},
            {"id":"h3","label":"Low-stock items (Stock ≤ Reorder Pt) are flagged","checked":False},
        ],
        "soft_requirements": [
            {"id":"s1","label":"Currency format on Unit Cost and Stock Value","checked":False},
            {"id":"s2","label":"Column widths auto-fitted","checked":False},
            {"id":"s3","label":"Low-stock rows visually distinct from healthy stock","checked":False},
        ],
    },
}

STUB_VERIFIES = {
    "pnl": [
        {"id":"h1","met":True, "reasoning":"Row 1 contains the five required headers.","references":["A1:E1"]},
        {"id":"h2","met":True, "reasoning":"Rows 2–7 are populated with Jan–Jun data.","references":["A2:A7"]},
        {"id":"h3","met":True, "reasoning":"$#,##0 format applied to B2:D7.","references":["B2:D7"]},
        {"id":"s1","met":True, "reasoning":"Dark header row contrasts clearly with data rows.","references":["A1:E1"]},
        {"id":"s2","met":True, "reasoning":"E3:E7 contains growth % formulas.","references":["E3:E7"]},
        {"id":"s3","met":True, "reasoning":"Row 8 contains TOTAL with SUM formulas.","references":["A8:E8"]},
        {"id":"s4","met":False,"reasoning":"autofitColumns called but column A may be too narrow for long labels.","references":["A1"]},
    ],
    "sales": [
        {"id":"h1","met":True, "reasoning":"Row 1 contains all six required headers.","references":["A1:F1"]},
        {"id":"h2","met":True, "reasoning":"E2:E8 contains =ROUND(C*D,2) formulas for each row.","references":["E2:E8"]},
        {"id":"h3","met":True, "reasoning":"Seven income rows populated across Employment, Freelance, Investment, Rental.","references":["A2:D8"]},
        {"id":"s1","met":True, "reasoning":"Rows shaded by bracket: blue (22%), yellow (24%), green (15%).","references":["A2:F8"]},
        {"id":"s2","met":True, "reasoning":"F2:F8 contains Withheld/Owed labels with colour coding.","references":["F2:F8"]},
        {"id":"s3","met":True, "reasoning":"Rows 11–13 contain Total Income, Total Tax Owed, Effective Rate.","references":["A11:B13"]},
        {"id":"s4","met":True, "reasoning":"autofitColumns called on A1:F13.","references":["A1:F13"]},
    ],
    "inventory": [
        {"id":"h1","met":True, "reasoning":"Row 1 contains SKU, Product, Category, Stock, Reorder Pt, Unit Cost, Stock Value.","references":["A1:G1"]},
        {"id":"h2","met":True, "reasoning":"G2:G9 contains =D*F formulas for Stock Value.","references":["G2:G9"]},
        {"id":"h3","met":True, "reasoning":"Rows where Stock ≤ Reorder Pt are highlighted red.","references":["A2:G9"]},
        {"id":"s1","met":True, "reasoning":"$#,##0.00 format applied to F2:G9.","references":["F2:G9"]},
        {"id":"s2","met":True, "reasoning":"autofitColumns called on A1:G9.","references":["A1:G9"]},
        {"id":"s3","met":True, "reasoning":"Low-stock rows use red fill and font, clearly distinct.","references":["A2:G9"]},
    ],
}
