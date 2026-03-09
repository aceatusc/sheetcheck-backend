# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
import os
from dotenv import load_dotenv

load_dotenv('../.env')

# Shared secret checked against the X-Addin-Secret header sent by llmClient.js
# Must match SHARED_SECRET in modules/llmClient.js
SHARED_SECRET = "my-super-secret-2025"

# Pick your LLM provider: "anthropic" | "openai" | "mistralai"
LLM_PROVIDER = "mistralai"

# API keys — prefer env vars in production
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "PLACEHOLDER_ANTHROPIC_KEY")
OPENAI_API_KEY    = os.getenv("OPENAI_API_KEY",    "PLACEHOLDER_OPENAI_KEY")
MISTRAL_API_KEY   = os.getenv("MISTRAL_API_KEY",    "PLACEHOLDER_OPENAI_KEY")

# Models
ANTHROPIC_MODEL = "claude-opus-4-5"
OPENAI_MODEL    = "gpt-4o"
MISTRAL_MODEL   = "mistral-small-2506" # "mistral-large-2512"

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are an Excel automation assistant. The user will describe a change they
want made to their spreadsheet. You must respond with ONLY a valid JSON array
of "code segments" — no prose, no markdown fences, no explanation outside the
JSON structure itself.

Each segment has this exact shape:
{
  "id":            "<unique string, e.g. seg-1>",
  "description":   "<short imperative label, e.g. 'Write header row'>",
  "sheet_context": ["<range address>", ...],
  "explanation":   "<one or two sentences explaining inputs to outputs>",
  "code":          "<async Office.js code string — see rules below>"
}

CRITICAL: 2-D ARRAY DIMENSION RULES (violations cause runtime errors)

Every array assigned to .values, .formulas, or .numberFormat MUST be a 2-D
array whose dimensions EXACTLY match the range: [rows][cols].

Count the rows and columns of the range BEFORE writing the array.

  Range A1:E1  -> 1 row,  5 cols  -> [[ v1, v2, v3, v4, v5 ]]
  Range A1:A5  -> 5 rows, 1 col   -> [[ v1 ], [ v2 ], [ v3 ], [ v4 ], [ v5 ]]
  Range B2:D4  -> 3 rows, 3 cols  -> [[ v,v,v ], [ v,v,v ], [ v,v,v ]]

CORRECT examples:

  // 1 row, 3 cols
  sheet.getRange("A1:C1").values = [["Month", "Revenue", "Profit"]];

  // 4 rows, 1 col — numberFormat must repeat per row
  sheet.getRange("B2:B5").numberFormat = [["$#,##0"], ["$#,##0"], ["$#,##0"], ["$#,##0"]];

  // 3 rows, 2 cols
  sheet.getRange("C2:D4").numberFormat = [
      ["0.0%", "$#,##0"],
      ["0.0%", "$#,##0"],
      ["0.0%", "$#,##0"],
  ];

WRONG examples (these ALL throw "array doesn't match range dimensions"):

  // Wrong: only 1 row given for a 4-row range
  sheet.getRange("B2:B5").numberFormat = [["$#,##0"]];

  // Wrong: only 1 col given for a 3-col range
  sheet.getRange("B2:D7").numberFormat = [
      ["$#,##0"], ["$#,##0"], ["$#,##0"],
      ["$#,##0"], ["$#,##0"], ["$#,##0"],
  ];

  // Wrong: flat array instead of 2-D
  sheet.getRange("A1:C1").values = ["Month", "Revenue", "Profit"];

SELF-CHECK before finalising each segment:
  1. Write down the range address, e.g. "B2:D7"
  2. Count: rows = last_row - first_row + 1, cols = last_col - first_col + 1
     Example: B2:D7 -> rows = 7-2+1 = 6, cols = D(4)-B(2)+1 = 3
  3. Verify your array has exactly that many inner arrays, each with exactly that many elements.
  4. If the counts do not match, fix the array before emitting JSON.

Additional code rules:
- Executed as: new AsyncFunction(code)()
- MUST use Office.js: await Excel.run(async (ctx) => { ... await ctx.sync(); })
- Do not import anything; Excel and Office globals are already available
- Do not include markdown backticks or any text outside the JSON array

Worksheet context provided by the user is attached below as JSON.
Use it to write accurate range addresses and avoid overwriting existing data.
"""
# TODO: include Examples if needed


# ---------------------------------------------------------------------------
# Stub segments
# ---------------------------------------------------------------------------

STUB_SEGMENTS = [
    {
        "id":            "seg-1",
        "description":   "Write header row labels",
        "sheet_context": ["A1:E1"],
        "explanation":   "Creates the five column headers — Month, Revenue, Expenses, Profit, and Growth % — in row 1. These labels define the structure of the entire table.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("A1:E1").values = [["Month", "Revenue", "Expenses", "Profit", "Growth %"]];
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-2",
        "description":   "Style header row (bold, background, font color)",
        "sheet_context": ["A1:E1"],
        "explanation":   "Applies a dark background (#1a1d27) with blue bold text to A1:E1, making the headers visually distinct from the data rows below.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet  = ctx.workbook.worksheets.getActiveWorksheet();
                const header = sheet.getRange("A1:E1");
                header.format.fill.color          = "#1a1d27";
                header.format.font.color          = "#4f8ef7";
                header.format.font.bold           = true;
                header.format.font.size           = 11;
                header.format.horizontalAlignment = "Center";
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-3",
        "description":   "Fill in monthly data rows",
        "sheet_context": ["A2:E7", "A2:A7"],
        "explanation":   "Writes six months of raw data into A2:E7. Columns B–D hold numeric values for Revenue, Expenses, and Profit. Column E is left blank here — Growth % formulas are added in the next step.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                const data = [
                    ["Jan", 142000, 98000,  44000, ""],
                    ["Feb", 158000, 104000, 54000, ""],
                    ["Mar", 175000, 110000, 65000, ""],
                    ["Apr", 163000, 107000, 56000, ""],
                    ["May", 191000, 115000, 76000, ""],
                    ["Jun", 210000, 121000, 89000, ""],
                ];
                sheet.getRange("A2:E7").values = data;
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-4",
        "description":   "Add Growth % formulas",
        "sheet_context": ["E3:E7", "D2:D7"],
        "explanation":   "Inserts IFERROR formulas in E3:E7 that compute month-over-month profit growth: (current − previous) ÷ previous. E2 is skipped because Jan has no prior month. Results are formatted as percentages with one decimal place.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("E3:E7").formulas = [
                    ['=IFERROR((D3-D2)/D2, "")'],
                    ['=IFERROR((D4-D3)/D3, "")'],
                    ['=IFERROR((D5-D4)/D4, "")'],
                    ['=IFERROR((D6-D5)/D5, "")'],
                    ['=IFERROR((D7-D6)/D6, "")'],
                ];
                sheet.getRange("E3:E7").numberFormat = [["0.0%"],["0.0%"],["0.0%"],["0.0%"],["0.0%"]];
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-5",
        "description":   "Format Revenue, Expenses, Profit as currency",
        "sheet_context": ["B2:D7"],
        "explanation":   "Applies the $#,##0 number format to the three numeric columns so values render as dollar amounts with comma separators (e.g. $142,000). Display-only — underlying values remain plain numbers.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("B2:D7").numberFormat = [
                    ["$#,##0", "$#,##0", "$#,##0"],
                    ["$#,##0", "$#,##0", "$#,##0"],
                    ["$#,##0", "$#,##0", "$#,##0"],
                    ["$#,##0", "$#,##0", "$#,##0"],
                    ["$#,##0", "$#,##0", "$#,##0"],
                    ["$#,##0", "$#,##0", "$#,##0"],
                ];
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-6",
        "description":   "Zebra-stripe data rows",
        "sheet_context": ["A2:E7"],
        "explanation":   "Alternates row backgrounds between #f5f7ff (even) and white (odd) across A2:E7, improving readability when scanning across wide rows.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                for (let i = 2; i <= 7; i++) {
                    sheet.getRange("A" + i + ":E" + i).format.fill.color =
                        i % 2 === 0 ? "#f5f7ff" : "#ffffff";
                }
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-7",
        "description":   "Colour Profit column by value",
        "sheet_context": ["D2:D7"],
        "explanation":   "Reads D2:D7 then colours font green+bold for profit ≥ $60,000, red otherwise — giving an instant visual signal of high vs low-performing months.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet  = ctx.workbook.worksheets.getActiveWorksheet();
                const profit = sheet.getRange("D2:D7");
                profit.load("values");
                await ctx.sync();
                profit.values.forEach((row, i) => {
                    const cell = sheet.getRange("D" + (i + 2));
                    cell.format.font.color = row[0] >= 60000 ? "#1a7a4a" : "#b94040";
                    cell.format.font.bold  = row[0] >= 60000;
                });
                await ctx.sync();
            });
        """,
    },
    {
        "id":            "seg-8",
        "description":   "Add totals row and auto-fit columns",
        "sheet_context": ["A8:E8", "A1:E8"],
        "explanation":   "Appends a TOTAL row in row 8 with SUM formulas for B–D, styled like the header. autofitColumns() then resizes all five columns to fit their widest content.",
        "code":          """
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("A8").values          = [["TOTAL"]];
                sheet.getRange("B8:D8").formulas     = [["=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)"]];
                sheet.getRange("B8:D8").numberFormat = [["$#,##0", "$#,##0", "$#,##0"]];
                const totalsRow = sheet.getRange("A8:E8");
                totalsRow.format.fill.color = "#1a1d27";
                totalsRow.format.font.color = "#ffffff";
                totalsRow.format.font.bold  = true;
                sheet.getRange("A1:E8").getEntireColumn().format.autofitColumns();
                await ctx.sync();
            });
        """,
    },
]
