# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
import os
from enum import Enum
import json
from dotenv import load_dotenv

load_dotenv('../.env')

# Shared secret — must match SHARED_SECRET in front-end/modules/llmClient.js
SHARED_SECRET = "my-super-secret-2025"

# ---------------------------------------------------------------------------
# Provider / Model enum
# ---------------------------------------------------------------------------

class Provider(str, Enum):
    ANTHROPIC = "anthropic"
    OPENAI    = "openai"
    MISTRAL   = "mistralai"
    GOOGLE    = "google"

class Model(str, Enum):
    CLAUDE  = "claude-opus-4-5"
    GPT4O   = "gpt-4o"
    MINISTRAL = "ministral-3b-2512"
    MISTRAL_SMALL = "mistral-small-2506"
    MISTRAL_LARGE = "mistral-large-2512"
    GEMINI_2_5 = "gemini-2.5-flash"
    GEMINI_3_FLASH = "gemini-3-flash-preview"
    GEMINI_3_1_FLASH_LITE = "gemini-3.1-flash-lite-preview"
    GEMINI_3_1_PRO = "gemini-3.1-pro-preview"

# Per-endpoint config: each API can use a different provider/model
class EndpointConfig:
    def __init__(self, provider: Provider, model: Model):
        self.provider = provider
        self.model    = model

ENDPOINT_MODELS = {
    "code":           EndpointConfig(Provider.GOOGLE, Model.GEMINI_3_FLASH),
    "ask":            EndpointConfig(Provider.MISTRAL, Model.MISTRAL_SMALL),
    "edit":           EndpointConfig(Provider.GOOGLE, Model.GEMINI_3_FLASH),
    "rubric_scaffold":EndpointConfig(Provider.MISTRAL, Model.MINISTRAL),
    "rubric_verify":  EndpointConfig(Provider.MISTRAL, Model.MISTRAL_SMALL),
    "chat":           EndpointConfig(Provider.MISTRAL, Model.MISTRAL_SMALL),
}

# API keys
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "PLACEHOLDER_ANTHROPIC_KEY")
OPENAI_API_KEY    = os.getenv("OPENAI_API_KEY",    "PLACEHOLDER_OPENAI_KEY")
MISTRAL_API_KEY   = os.getenv("MISTRAL_API_KEY",   "PLACEHOLDER_MISTRAL_KEY")
GEMINI_API_KEY    = os.getenv("GEMINI_API_KEY",     "PLACEHOLDER_GEMINI_KEY")

# ---------------------------------------------------------------------------
# System prompts (one per endpoint)
# ---------------------------------------------------------------------------

# _ARRAY_RULES = """
# CRITICAL: 2-D ARRAY DIMENSION RULES (violations cause runtime errors)

# Every array assigned to .values, .formulas, or .numberFormat MUST be a 2-D
# array whose dimensions EXACTLY match the range: [rows][cols].

# Count the rows and columns of the range BEFORE writing the array.

#   Range A1:E1  -> 1 row,  5 cols  -> [[ v1, v2, v3, v4, v5 ]]
#   Range A1:A5  -> 5 rows, 1 col   -> [[ v1 ], [ v2 ], [ v3 ], [ v4 ], [ v5 ]]
#   Range B2:D4  -> 3 rows, 3 cols  -> [[ v,v,v ], [ v,v,v ], [ v,v,v ]]

# CORRECT examples:

#   // 1 row, 3 cols
#   sheet.getRange("A1:C1").values = [["Month", "Revenue", "Profit"]];

#   // 4 rows, 1 col — numberFormat must repeat per row
#   sheet.getRange("B2:B5").numberFormat = [["$#,##0"], ["$#,##0"], ["$#,##0"], ["$#,##0"]];

#   // 3 rows, 2 cols
#   sheet.getRange("C2:D4").numberFormat = [
#       ["0.0%", "$#,##0"],
#       ["0.0%", "$#,##0"],
#       ["0.0%", "$#,##0"],
#   ];

# WRONG examples (these ALL throw "array doesn't match range dimensions"):

#   // Wrong: only 1 row given for a 4-row range
#   sheet.getRange("B2:B5").numberFormat = [["$#,##0"]];

#   // Wrong: only 1 col given for a 3-col range
#   sheet.getRange("B2:D7").numberFormat = [
#       ["$#,##0"], ["$#,##0"], ["$#,##0"],
#       ["$#,##0"], ["$#,##0"], ["$#,##0"],
#   ];

#   // Wrong: flat array instead of 2-D
#   sheet.getRange("A1:C1").values = ["Month", "Revenue", "Profit"];

# SELF-CHECK before finalising each segment:
#   1. Write down the range address, e.g. "B2:D7"
#   2. Count: rows = last_row - first_row + 1, cols = last_col - first_col + 1
#      Example: B2:D7 -> rows = 7-2+1 = 6, cols = D(4)-B(2)+1 = 3
#   3. Verify your array has exactly that many inner arrays, each with exactly that many elements.
#   4. If the counts do not match, fix the array before emitting JSON.
# """
_ARRAY_RULES = """"""

# _CODE_RULES = """
# Code field rules:
# - Executed as: new AsyncFunction(code)()
# - MUST use: await Excel.run(async (ctx) => { ... await ctx.sync(); })
# - Do not import anything; Excel and Office globals are available
# - Do not include markdown backticks or text outside the JSON array
# - If you use sheet.getUsedRange(), you MUST call await ctx.sync() immediately before it if you have just assigned values to the sheet. You cannot reference the "Used Range" of data that hasn't been synchronized yet.
# - THE MERGE RULE:
# If a range is merged (e.g., A1:E1), you MUST still provide an array matching the full dimensions.
# Wrong: range("A1:E1").values = [["Title"]]
# Right: range("A1:E1").values = [["Title", "", "", "", ""]]
# - THE SINGLE-VALUE SHORTCUT:
# To avoid dimension math errors when applying the SAME value or format to a large range, do not use an array. Assign the value directly to the property.
# Better: range.numberFormat = "yyyy-mm-dd"; (This applies the format to every cell in the range automatically without needing a 2-D array).
# Better: range.values = "Pending"; (This fills the entire range with the word "Pending").
# """
_CODE_RULES = """"""

# ---------------------------------------------------------------------------
# Stub segments
# ---------------------------------------------------------------------------

STUB_SEGMENTS = [
    {
        "id":            "seg-1",
        "description":   "Write header row labels",
        "sheet_context": ["A1:E1"],
        "explanation":   "Creates the five column headers — Month, Revenue, Expenses, Profit, and Growth % — in row 1.",
        "predecessors":  [],
        "affordances":   [
            {"id":"aff-1","label":"Header background color","type":"color","value":"#1a1d27","options":[]},
            {"id":"aff-2","label":"Font color","type":"color","value":"#4f8ef7","options":[]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"5-column layout (default)","probability":0.7,
             "code": 'await Excel.run(async (ctx) => { const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A1:E1").values = [["Month","Revenue","Expenses","Profit","Growth %"]]; await ctx.sync(); });'},
            {"id":"alt-2","label":"6-column with Notes","probability":0.2,
             "code": 'await Excel.run(async (ctx) => { const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A1:F1").values = [["Month","Revenue","Expenses","Profit","Growth %","Notes"]]; await ctx.sync(); });'},
            {"id":"alt-3","label":"Minimal 3-column","probability":0.1,
             "code": 'await Excel.run(async (ctx) => { const s = ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A1:C1").values = [["Month","Revenue","Profit"]]; await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why 5 columns?","a":"Revenue, Expenses, and Profit together give a complete P&L picture; Growth % adds a trend indicator."},
            {"q":"Why start in A1?","a":"Starting at the top-left keeps the table anchored and makes formulas simpler."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:E1").clear(); await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet = ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A1:E1").values = [["Month","Revenue","Expenses","Profit","Growth %"]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-2",
        "description":   "Style header row",
        "sheet_context": ["A1:E1"],
        "explanation":   "Applies dark background with blue bold text to A1:E1, making headers visually distinct.",
        "predecessors":  ["seg-1"],
        "affordances":   [
            {"id":"aff-3","label":"Background color","type":"color","value":"#1a1d27","options":[]},
            {"id":"aff-4","label":"Font color","type":"color","value":"#4f8ef7","options":[]},
            {"id":"aff-5","label":"Font size","type":"number","value":"11","options":[]},
            {"id":"aff-6","label":"Alignment","type":"dropdown","value":"Center","options":["Left","Center","Right"]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"Dark header (default)","probability":0.6,
             "code": 'await Excel.run(async (ctx) => { const h = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:E1"); h.format.fill.color="#1a1d27"; h.format.font.color="#4f8ef7"; h.format.font.bold=true; h.format.font.size=11; h.format.horizontalAlignment="Center"; await ctx.sync(); });'},
            {"id":"alt-2","label":"Light blue header","probability":0.25,
             "code": 'await Excel.run(async (ctx) => { const h = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:E1"); h.format.fill.color="#dbeafe"; h.format.font.color="#1e3a8a"; h.format.font.bold=true; h.format.font.size=11; await ctx.sync(); });'},
            {"id":"alt-3","label":"Green accent header","probability":0.15,
             "code": 'await Excel.run(async (ctx) => { const h = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:E1"); h.format.fill.color="#14532d"; h.format.font.color="#bbf7d0"; h.format.font.bold=true; h.format.font.size=11; await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why dark background?","a":"High contrast between header and data rows helps users quickly identify column names."},
            {"q":"Can I change these colors?","a":"Yes — use the affordance controls in the Edit panel to pick any color."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { const h = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:E1"); h.format.fill.color=""; h.format.font.color=""; h.format.font.bold=false; await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); const header=sheet.getRange("A1:E1"); header.format.fill.color="#1a1d27"; header.format.font.color="#4f8ef7"; header.format.font.bold=true; header.format.font.size=11; header.format.horizontalAlignment="Center"; await ctx.sync(); });',
    },
    {
        "id":            "seg-3",
        "description":   "Fill in monthly data rows",
        "sheet_context": ["A2:E7"],
        "explanation":   "Writes six months of raw data into A2:E7. Column E is left blank for Growth % formulas in the next step.",
        "predecessors":  ["seg-1"],
        "affordances":   [],
        "alternatives":  [
            {"id":"alt-1","label":"Jan–Jun data (default)","probability":0.5,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:E7").values=[["Jan",142000,98000,44000,""],["Feb",158000,104000,54000,""],["Mar",175000,110000,65000,""],["Apr",163000,107000,56000,""],["May",191000,115000,76000,""],["Jun",210000,121000,89000,""]]; await ctx.sync(); });'},
            {"id":"alt-2","label":"Jul–Dec data","probability":0.3,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:E7").values=[["Jul",220000,130000,90000,""],["Aug",235000,128000,107000,""],["Sep",198000,122000,76000,""],["Oct",215000,135000,80000,""],["Nov",241000,140000,101000,""],["Dec",268000,148000,120000,""]]; await ctx.sync(); });'},
            {"id":"alt-3","label":"Placeholder zeros","probability":0.2,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A2:E7").values=[["Jan",0,0,0,""],["Feb",0,0,0,""],["Mar",0,0,0,""],["Apr",0,0,0,""],["May",0,0,0,""],["Jun",0,0,0,""]]; await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why leave column E blank?","a":"Growth % requires the previous month's value to exist first; we add those formulas in the dedicated next step."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("A2:E7").clear(); await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A2:E7").values=[["Jan",142000,98000,44000,""],["Feb",158000,104000,54000,""],["Mar",175000,110000,65000,""],["Apr",163000,107000,56000,""],["May",191000,115000,76000,""],["Jun",210000,121000,89000,""]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-4",
        "description":   "Add Growth % formulas",
        "sheet_context": ["E3:E7","D2:D7"],
        "explanation":   "Inserts IFERROR month-over-month profit growth formulas in E3:E7, formatted as percentages.",
        "predecessors":  ["seg-3"],
        "affordances":   [
            {"id":"aff-7","label":"Number format","type":"dropdown","value":"0.0%","options":["0%","0.0%","0.00%"]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"Profit growth (default)","probability":0.5,
             "code": "await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('E3:E7').formulas=[['=IFERROR((D3-D2)/D2,\"\")'],['=IFERROR((D4-D3)/D3,\"\")'],['=IFERROR((D5-D4)/D4,\"\")'],['=IFERROR((D6-D5)/D5,\"\")'],['=IFERROR((D7-D6)/D6,\"\")']];\ns.getRange('E3:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });"},
            {"id":"alt-2","label":"Revenue growth","probability":0.3,
             "code": "await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('E3:E7').formulas=[['=IFERROR((B3-B2)/B2,\"\")'],['=IFERROR((B4-B3)/B3,\"\")'],['=IFERROR((B5-B4)/B4,\"\")'],['=IFERROR((B6-B5)/B5,\"\")'],['=IFERROR((B7-B6)/B6,\"\")']];\ns.getRange('E3:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });"},
            {"id":"alt-3","label":"Profit margin %","probability":0.2,
             "code": "await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange('E2:E7').formulas=[['=IFERROR(D2/B2,\"\")'],['=IFERROR(D3/B3,\"\")'],['=IFERROR(D4/B4,\"\")'],['=IFERROR(D5/B5,\"\")'],['=IFERROR(D6/B6,\"\")'],['=IFERROR(D7/B7,\"\")']];\ns.getRange('E2:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });"},
        ],
        "qa_pairs": [
            {"q":"Why skip E2?","a":"January has no prior month, so the growth formula would divide by zero; IFERROR handles it but leaving E2 blank is cleaner."},
            {"q":"Why IFERROR?","a":"Protects against division-by-zero if any profit value is 0."},
        ],
        "undo_code": "await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange('E2:E7').clear(); await ctx.sync(); });",
        "code": "await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange('E3:E7').formulas=[['=IFERROR((D3-D2)/D2,\"\")'],['=IFERROR((D4-D3)/D3,\"\")'],['=IFERROR((D5-D4)/D4,\"\")'],['=IFERROR((D6-D5)/D5,\"\")'],['=IFERROR((D7-D6)/D6,\"\")']];\nsheet.getRange('E3:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });",
    },
    {
        "id":            "seg-5",
        "description":   "Format currency columns",
        "sheet_context": ["B2:D7"],
        "explanation":   "Applies $#,##0 number format to Revenue, Expenses, and Profit columns.",
        "predecessors":  ["seg-3"],
        "affordances":   [
            {"id":"aff-8","label":"Currency format","type":"dropdown","value":"$#,##0","options":["$#,##0","$#,##0.00","#,##0","#,##0.00"]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"Dollar no decimals (default)","probability":0.6,
             "code": 'await Excel.run(async (ctx) => { const r=ctx.workbook.worksheets.getActiveWorksheet().getRange("B2:D7"); r.numberFormat=[["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"]]; await ctx.sync(); });'},
            {"id":"alt-2","label":"Dollar with cents","probability":0.25,
             "code": 'await Excel.run(async (ctx) => { const r=ctx.workbook.worksheets.getActiveWorksheet().getRange("B2:D7"); r.numberFormat=[["$#,##0.00","$#,##0.00","$#,##0.00"],["$#,##0.00","$#,##0.00","$#,##0.00"],["$#,##0.00","$#,##0.00","$#,##0.00"],["$#,##0.00","$#,##0.00","$#,##0.00"],["$#,##0.00","$#,##0.00","$#,##0.00"],["$#,##0.00","$#,##0.00","$#,##0.00"]]; await ctx.sync(); });'},
            {"id":"alt-3","label":"Thousands only","probability":0.15,
             "code": 'await Excel.run(async (ctx) => { const r=ctx.workbook.worksheets.getActiveWorksheet().getRange("B2:D7"); r.numberFormat=[["#,##0","#,##0","#,##0"],["#,##0","#,##0","#,##0"],["#,##0","#,##0","#,##0"],["#,##0","#,##0","#,##0"],["#,##0","#,##0","#,##0"],["#,##0","#,##0","#,##0"]]; await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why $#,##0 over $#,##0.00?","a":"Financial summaries typically round to whole dollars for readability; individual transactions warrant cents."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("B2:D7").numberFormat=null; await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("B2:D7").numberFormat=[["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-6",
        "description":   "Zebra-stripe rows",
        "sheet_context": ["A2:E7"],
        "explanation":   "Alternates row backgrounds for readability across A2:E7.",
        "predecessors":  ["seg-3"],
        "affordances":   [
            {"id":"aff-9","label":"Even row color","type":"color","value":"#f5f7ff","options":[]},
            {"id":"aff-10","label":"Odd row color","type":"color","value":"#ffffff","options":[]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"Blue/white stripes (default)","probability":0.5,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=7;i++){s.getRange("A"+i+":E"+i).format.fill.color=i%2===0?"#f5f7ff":"#ffffff";} await ctx.sync(); });'},
            {"id":"alt-2","label":"Green/white stripes","probability":0.3,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=7;i++){s.getRange("A"+i+":E"+i).format.fill.color=i%2===0?"#f0fdf4":"#ffffff";} await ctx.sync(); });'},
            {"id":"alt-3","label":"No stripes (uniform)","probability":0.2,
             "code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("A2:E7").format.fill.color="#ffffff"; await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why alternating colors?","a":"Zebra striping reduces eye tracking errors when reading across wide rows."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("A2:E7").format.fill.color=""; await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=7;i++){sheet.getRange("A"+i+":E"+i).format.fill.color=i%2===0?"#f5f7ff":"#ffffff";} await ctx.sync(); });',
    },
    {
        "id":            "seg-7",
        "description":   "Colour-code Profit column",
        "sheet_context": ["D2:D7"],
        "explanation":   "Reads D2:D7 and colours font green+bold for profit ≥ threshold, red otherwise.",
        "predecessors":  ["seg-3","seg-5"],
        "affordances":   [
            {"id":"aff-11","label":"Profit threshold ($)","type":"number","value":"60000","options":[]},
            {"id":"aff-12","label":"High-profit color","type":"color","value":"#1a7a4a","options":[]},
            {"id":"aff-13","label":"Low-profit color","type":"color","value":"#b94040","options":[]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"Green/red by threshold (default)","probability":0.6,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const p=s.getRange("D2:D7"); p.load("values"); await ctx.sync(); p.values.forEach((r,i)=>{const c=s.getRange("D"+(i+2)); c.format.font.color=r[0]>=60000?"#1a7a4a":"#b94040"; c.format.font.bold=r[0]>=60000;}); await ctx.sync(); });'},
            {"id":"alt-2","label":"Traffic-light 3-tier","probability":0.25,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const p=s.getRange("D2:D7"); p.load("values"); await ctx.sync(); p.values.forEach((r,i)=>{const c=s.getRange("D"+(i+2)); if(r[0]>=70000){c.format.font.color="#1a7a4a";c.format.font.bold=true;}else if(r[0]>=50000){c.format.font.color="#b45309";}else{c.format.font.color="#b94040";}}); await ctx.sync(); });'},
            {"id":"alt-3","label":"Bold only (no color)","probability":0.15,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); const p=s.getRange("D2:D7"); p.load("values"); await ctx.sync(); p.values.forEach((r,i)=>{s.getRange("D"+(i+2)).format.font.bold=r[0]>=60000;}); await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why $60,000 threshold?","a":"It sits roughly at the median of the sample data, giving a balanced split; the affordance lets you adjust it."},
            {"q":"Why bold for high values?","a":"Double encoding (color + weight) helps users with color vision deficiency."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { const r=ctx.workbook.worksheets.getActiveWorksheet().getRange("D2:D7"); r.format.font.color=""; r.format.font.bold=false; await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); const profit=sheet.getRange("D2:D7"); profit.load("values"); await ctx.sync(); profit.values.forEach((row,i)=>{const cell=sheet.getRange("D"+(i+2)); cell.format.font.color=row[0]>=60000?"#1a7a4a":"#b94040"; cell.format.font.bold=row[0]>=60000;}); await ctx.sync(); });',
    },
    {
        "id":            "seg-8",
        "description":   "Add totals row and auto-fit",
        "sheet_context": ["A8:E8","A1:E8"],
        "explanation":   "Appends a dark TOTAL row with SUM formulas in B8:D8, then auto-fits all columns.",
        "predecessors":  ["seg-3","seg-5","seg-6"],
        "affordances":   [
            {"id":"aff-14","label":"Totals row label","type":"dropdown","value":"TOTAL","options":["TOTAL","SUM","GRAND TOTAL","Total"]},
            {"id":"aff-15","label":"Totals row background","type":"color","value":"#1a1d27","options":[]},
        ],
        "alternatives":  [
            {"id":"alt-1","label":"SUM row dark styled (default)","probability":0.6,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A8").values=[["TOTAL"]]; s.getRange("B8:D8").formulas=[["=SUM(B2:B7)","=SUM(C2:C7)","=SUM(D2:D7)"]]; s.getRange("B8:D8").numberFormat=[["$#,##0","$#,##0","$#,##0"]]; const t=s.getRange("A8:E8"); t.format.fill.color="#1a1d27"; t.format.font.color="#ffffff"; t.format.font.bold=true; s.getRange("A1:E8").getEntireColumn().format.autofitColumns(); await ctx.sync(); });'},
            {"id":"alt-2","label":"AVERAGE row","probability":0.25,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A8").values=[["AVERAGE"]]; s.getRange("B8:D8").formulas=[["=AVERAGE(B2:B7)","=AVERAGE(C2:C7)","=AVERAGE(D2:D7)"]]; s.getRange("B8:D8").numberFormat=[["$#,##0","$#,##0","$#,##0"]]; const t=s.getRange("A8:E8"); t.format.fill.color="#374151"; t.format.font.color="#ffffff"; t.format.font.bold=true; s.getRange("A1:E8").getEntireColumn().format.autofitColumns(); await ctx.sync(); });'},
            {"id":"alt-3","label":"Both SUM and AVERAGE rows","probability":0.15,
             "code": 'await Excel.run(async (ctx) => { const s=ctx.workbook.worksheets.getActiveWorksheet(); s.getRange("A8").values=[["TOTAL"]]; s.getRange("B8:D8").formulas=[["=SUM(B2:B7)","=SUM(C2:C7)","=SUM(D2:D7)"]]; s.getRange("A9").values=[["AVERAGE"]]; s.getRange("B9:D9").formulas=[["=AVERAGE(B2:B7)","=AVERAGE(C2:C7)","=AVERAGE(D2:D7)"]]; for(const r of[s.getRange("A8:E8"),s.getRange("A9:E9")]){r.format.fill.color="#1a1d27";r.format.font.color="#fff";r.format.font.bold=true;} s.getRange("B8:D9").numberFormat=[["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"]]; s.getRange("A1:E9").getEntireColumn().format.autofitColumns(); await ctx.sync(); });'},
        ],
        "qa_pairs": [
            {"q":"Why autofitColumns last?","a":"All data must exist before fitting, otherwise columns are sized to empty cells."},
            {"q":"Why SUM not SUBTOTAL?","a":"SUBTOTAL respects filters, but for a simple 6-row table SUM is clearer and more expected."},
        ],
        "undo_code": 'await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("A8:E9").clear(); await ctx.sync(); });',
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A8").values=[["TOTAL"]]; sheet.getRange("B8:D8").formulas=[["=SUM(B2:B7)","=SUM(C2:C7)","=SUM(D2:D7)"]]; sheet.getRange("B8:D8").numberFormat=[["$#,##0","$#,##0","$#,##0"]]; const t=sheet.getRange("A8:E8"); t.format.fill.color="#1a1d27"; t.format.font.color="#ffffff"; t.format.font.bold=true; sheet.getRange("A1:E8").getEntireColumn().format.autofitColumns(); await ctx.sync(); });',
    },
]

STUB_RUBRIC = {
    "hard_requirements": [
        {"id": "h1", "label": "Header row present with correct column labels", "checked": False},
        {"id": "h2", "label": "All 6 months of data filled in rows 2–7", "checked": False},
        {"id": "h3", "label": "Currency format applied to Revenue, Expenses, Profit", "checked": False},
    ],
    "soft_requirements": [
        {"id": "s1", "label": "Visual hierarchy differentiates headers from data", "checked": False},
        {"id": "s2", "label": "Growth % column shows trend direction", "checked": False},
        {"id": "s3", "label": "Totals row present at bottom", "checked": False},
        {"id": "s4", "label": "Column widths are readable without scrolling", "checked": False},
    ],
}

STUB_ASK = {
    "answer": "[test] The LLM would explain the design choice for this step in detail.",
    "follow_up_questions": [
        "Why was this range chosen?",
        "Can this be done differently?",
    ],
}

STUB_EDIT = STUB_SEGMENTS[1]  # Return seg-2 as an edited segment stub

STUB_VERIFY = [
    {"id": "h1", "met": True,  "reasoning": "Row 1 contains Month, Revenue, Expenses, Profit, Growth % labels.", "references": ["A1:E1"]},
    {"id": "h2", "met": True,  "reasoning": "Rows 2–7 are populated with Jan–Jun data.", "references": ["A2:A7"]},
    {"id": "h3", "met": True,  "reasoning": "$#,##0 format applied to B2:D7.", "references": ["B2:D7"]},
    {"id": "s1", "met": True,  "reasoning": "Dark header row contrasts with white/blue data rows.", "references": ["A1:E1"]},
    {"id": "s2", "met": True,  "reasoning": "E3:E7 contains growth % formulas with percentage formatting.", "references": ["E3:E7"]},
    {"id": "s3", "met": True,  "reasoning": "Row 8 contains TOTAL with SUM formulas.", "references": ["A8:E8"]},
    {"id": "s4", "met": False, "reasoning": "autofitColumns was called but column A may still be too narrow for 'GRAND TOTAL'.", "references": ["A8"]},
]


SYSTEM_PROMPTS = {

"code": """You are an Excel automation assistant. The user will describe a change they want made to their spreadsheet. You must respond with ONLY a valid JSON array
of "code segments" — no prose, no markdown fences, no explanation outside the JSON structure itself.

Each segment shape:
{
  "id":            "seg-N",
  "description":   "Short imperative label",
  "sheet_context": ["<range address>", ...],
  "explanation":   "One or two sentences: inputs to outputs",
  "predecessors":  ["seg-id", ...],
  "affordances":   [{"id":"aff-N","label":"Label","type":"dropdown|number|color|toggle","value":"default","options":["a","b"]}],
  "alternatives":  [
    {"id":"alt-1","label":"Alternative 1","probability":0.6,"code":"...office.js..."},
    {"id":"alt-2","label":"Alternative 2","probability":0.25,"code":"...office.js..."},
    {"id":"alt-3","label":"Alternative 3","probability":0.15,"code":"...office.js..."}
  ],
  "qa_pairs":      [{"q":"Why ...?","a":"Because ..."},{"q":"...","a":"..."}],
  "code":          "await Excel.run(async (ctx) => { ... await ctx.sync(); });",
  "undo_code":     "await Excel.run(async (ctx) => { /* undo */ await ctx.sync(); });",
}

Rules:
- predecessors: list ids of segments this one depends on semantically (can be empty [])
- affordances: dynamic UI controls the user can tweak that affect code behaviour. Be creative — expose colors, formulas, thresholds, labels, chart types, etc. as affordances whenever the code has a "magic value". Each affordance has a placeholder comment in the code like: /* AFFORDANCE:aff-N */
- alternatives: always generate exactly 3 implementations with probabilities summing to 1.0
- qa_pairs: 2–3 Q&A pairs explaining design choices for this step
- undo_code: Office.js that reverses exactly what code does (clear values/formats etc.)
""" + _ARRAY_RULES + _CODE_RULES + """
Worksheet context and optional rubric are provided below.
""" + f"Example: {json.dumps(STUB_SEGMENTS)}\n",

"ask": """You are an Excel assistant answering a follow-up question about a specific step in a spreadsheet automation plan.

The user will provide:
- The current step's description and explanation
- Their follow-up question

Respond with a single JSON object:
{
  "answer": "Clear, concise answer (1-3 sentences)",
  "follow_up_questions": ["Short suggested follow-up 1", "Short suggested follow-up 2"]
}

Respond with ONLY the JSON object, no markdown, no extra text.
""",

"edit": """You are an Excel automation assistant. The user wants to modify a specific step.

You will receive the original segment and user feedback/preferred alternative.
Respond with ONLY a single updated segment JSON object (same shape as a code segment from /code).
Segment shape:
{
  "id":            "seg-N",
  "description":   "Short imperative label",
  "sheet_context": ["<range address>", ...],
  "explanation":   "One or two sentences: inputs to outputs",
  "predecessors":  ["seg-id", ...],
  "affordances":   [{"id":"aff-N","label":"Label","type":"dropdown|number|color|toggle","value":"default","options":["a","b"]}],
  "alternatives":  [
    {"id":"alt-1","label":"Alternative 1","probability":0.6,"code":"...office.js..."},
    {"id":"alt-2","label":"Alternative 2","probability":0.25,"code":"...office.js..."},
    {"id":"alt-3","label":"Alternative 3","probability":0.15,"code":"...office.js..."}
  ],
  "qa_pairs":      [{"q":"Why ...?","a":"Because ..."},{"q":"...","a":"..."}],
  "code":          "await Excel.run(async (ctx) => { ... await ctx.sync(); });",
  "undo_code":     "await Excel.run(async (ctx) => { /* undo */ await ctx.sync(); });",
}
Include updated alternatives, affordances, undo_code, qa_pairs reflecting the edit.
""" + _ARRAY_RULES + _CODE_RULES + f"Example: {json.dumps(STUB_SEGMENTS)}\n",

"rubric_scaffold": """You are an Excel task evaluator. Generate an initial rubric for a spreadsheet task.

The user will describe their task. Respond with ONLY a JSON object:
{
  "hard_requirements": [
    {"id": "h1", "label": "Requirement text", "checked": false}
  ],
  "soft_requirements": [
    {"id": "s1", "label": "Requirement text", "checked": false}
  ]
}

Generate 2-3 hard and 3-4 soft requirements. Hard = must-have correctness criteria. Soft = quality/style preferences.
Respond with ONLY the JSON, no markdown.
""",

"rubric_verify": """You are an Excel task evaluator. Evaluate whether a completed worksheet satisfies each rubric requirement.

You will receive the rubric and worksheet state. Respond with ONLY a JSON array:
[
  {
    "id": "h1",
    "met": true,
    "reasoning": "One sentence explanation",
    "references": ["A1:E1", "B2:D7"]
  }
]

Include every rubric item (hard and soft). Be precise about cell references.
Respond with ONLY the JSON array, no markdown.
""",

"chat": """You are a helpful Excel and spreadsheet assistant. Answer the user's question clearly and concisely.
Keep responses brief and practical. You may use markdown formatting.
""",
}
