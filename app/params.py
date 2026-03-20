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
    CLAUDE_OPUS  = "claude-opus-4-6"
    CLAUDE_SONNET = "claude-sonnet-4-6"
    CLAUDE_HAIKU = "claude-haiku-4-5"
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
    "code":           EndpointConfig(Provider.GOOGLE, Model.GEMINI_3_1_PRO),
    "ask":            EndpointConfig(Provider.MISTRAL, Model.MISTRAL_SMALL),
    "edit":           EndpointConfig(Provider.GOOGLE, Model.GEMINI_3_1_FLASH_LITE),
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
# Stub segments
# ---------------------------------------------------------------------------

STUB_SEGMENTS = [
    {
        "id":            "seg-1",
        "description":   "Write header row labels",
        "sheet_context": ["A1:E1"],
        "explanation":   "Creates the five column headers — Month, Revenue, Expenses, Profit, and Growth % — in row 1.",
        "predecessors":  [],
        "qa_pairs": [
            {"q":"Why 5 columns?","a":"Revenue, Expenses, and Profit together give a complete P&L picture; Growth % adds a trend indicator."},
            {"q":"Why start in A1?","a":"Starting at the top-left keeps the table anchored and makes formulas simpler."},
        ],
        "edit_suggestions": ["Add more columns like Net Margin", "Rename Growth % to MoM Growth", "Start table at a different cell"],
        "parameters":    [],
        "code": 'await Excel.run(async (ctx) => { const sheet = ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A1:E1").values = [["Month","Revenue","Expenses","Profit","Growth %"]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-2",
        "description":   "Style header row",
        "sheet_context": ["A1:E1"],
        "explanation":   "Applies dark background with blue bold text to A1:E1, making headers visually distinct.",
        "predecessors":  ["seg-1"],
        "qa_pairs": [
            {"q":"Why dark background?","a":"High contrast between header and data rows helps users quickly identify column names."},
        ],
        "edit_suggestions": ["Use a lighter header background", "Change the font color", "Increase the font size"],
        "parameters": [
            {"label": "Background color", "key": "#1a1d27", "value": "#1a1d27", "type": "color"},
            {"label": "Font color",       "key": "#4f8ef7", "value": "#4f8ef7", "type": "color"},
            {"label": "Font size",        "key": "11",      "value": 11,        "type": "number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); const header=sheet.getRange("A1:E1"); header.format.fill.color="#1a1d27"; header.format.font.color="#4f8ef7"; header.format.font.bold=true; header.format.font.size=11; header.format.horizontalAlignment="Center"; await ctx.sync(); });',
    },
    {
        "id":            "seg-3",
        "description":   "Fill in monthly data rows",
        "sheet_context": ["A2:E7"],
        "explanation":   "Writes six months of raw data into A2:E7. Column E is left blank for Growth % formulas in the next step.",
        "predecessors":  ["seg-1"],
        "qa_pairs": [
            {"q":"Why leave column E blank?","a":"Growth % requires the previous month's value to exist first; we add those formulas in the dedicated next step."},
        ],
        "edit_suggestions": ["Extend to 12 months", "Change the starting revenue values", "Add a different starting month"],
        "parameters": [
            {"label": "Jan Revenue",  "key": "142000", "value": 142000, "type": "number"},
            {"label": "Jan Expenses", "key": "98000",  "value": 98000,  "type": "number"},
        ],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("A2:E7").values=[["Jan",142000,98000,44000,""],["Feb",158000,104000,54000,""],["Mar",175000,110000,65000,""],["Apr",163000,107000,56000,""],["May",191000,115000,76000,""],["Jun",210000,121000,89000,""]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-4",
        "description":   "Add Growth % formulas",
        "sheet_context": ["E3:E7","D2:D7"],
        "explanation":   "Inserts IFERROR month-over-month profit growth formulas in E3:E7, formatted as percentages.",
        "predecessors":  ["seg-3"],
        "qa_pairs": [
            {"q":"Why skip E2?","a":"January has no prior month, so the growth formula would divide by zero; IFERROR handles it but leaving E2 blank is cleaner."},
            {"q":"Why IFERROR?","a":"Protects against division-by-zero if any profit value is 0."},
        ],
        "edit_suggestions": ["Show growth as Revenue % instead of Profit %", "Format as integer % instead of 1 decimal", "Include E2 with an N/A label"],
        "parameters": [
            {"label": "Number format", "key": "0.0%", "value": "0.0%", "type": "text"},
        ],
        "code": "await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange('E3:E7').formulas=[['=IFERROR((D3-D2)/D2,\"\")'],['=IFERROR((D4-D3)/D3,\"\")'],['=IFERROR((D5-D4)/D4,\"\")'],['=IFERROR((D6-D5)/D5,\"\")'],['=IFERROR((D7-D6)/D6,\"\")']];\nsheet.getRange('E3:E7').numberFormat=[['0.0%'],['0.0%'],['0.0%'],['0.0%'],['0.0%']]; await ctx.sync(); });",
    },
    {
        "id":            "seg-5",
        "description":   "Format currency columns",
        "sheet_context": ["B2:D7"],
        "explanation":   "Applies $#,##0 number format to Revenue, Expenses, and Profit columns.",
        "predecessors":  ["seg-3"],
        "qa_pairs": [
            {"q":"Why $#,##0 over $#,##0.00?","a":"Financial summaries typically round to whole dollars for readability; individual transactions warrant cents."},
        ],
        "edit_suggestions": ["Show cents with $#,##0.00", "Use €#,##0 for Euros", "Remove the currency symbol"],
        "parameters": [
            {"label": "Currency format", "key": "$#,##0", "value": "$#,##0", "type": "select", "options": ["$#,##0", "$#,##0.00", "€#,##0", "£#,##0", "#,##0"]},
        ],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); sheet.getRange("B2:D7").numberFormat=[["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"],["$#,##0","$#,##0","$#,##0"]]; await ctx.sync(); });',
    },
    {
        "id":            "seg-6",
        "description":   "Zebra-stripe rows",
        "sheet_context": ["A2:E7"],
        "explanation":   "Alternates row backgrounds for readability across A2:E7.",
        "predecessors":  ["seg-3"],
        "qa_pairs": [
            {"q":"Why alternating colors?","a":"Zebra striping reduces eye tracking errors when reading across wide rows."},
        ],
        "edit_suggestions": ["Use a stronger stripe color", "Apply stripes to the header too", "Use a warm color palette instead"],
        "parameters": [
            {"label": "Even row color", "key": "#f5f7ff", "value": "#f5f7ff", "type": "color"},
            {"label": "Odd row color",  "key": "#ffffff", "value": "#ffffff", "type": "color"},
        ],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); for(let i=2;i<=7;i++){sheet.getRange("A"+i+":E"+i).format.fill.color=i%2===0?"#f5f7ff":"#ffffff";} await ctx.sync(); });',
    },
    {
        "id":            "seg-7",
        "description":   "Colour-code Profit column",
        "sheet_context": ["D2:D7"],
        "explanation":   "Reads D2:D7 and colours font green+bold for profit ≥ threshold, red otherwise.",
        "predecessors":  ["seg-3","seg-5"],
        "qa_pairs": [
            {"q":"Why bold for high values?","a":"Double encoding (color + weight) helps users with color vision deficiency."},
        ],
        "edit_suggestions": ["Change the profit threshold", "Use different colors", "Apply to a wider range"],
        "parameters":    [{"label": "Profit threshold", "key": "threshold", "value": 60000, "type": "number"}, {"label": "High color", "key": "highColor", "value": "#1a7a4a", "type": "color"}, {"label": "Low color", "key": "lowColor", "value": "#b94040", "type": "color"}],
        "code": 'await Excel.run(async (ctx) => { const sheet=ctx.workbook.worksheets.getActiveWorksheet(); const profit=sheet.getRange("D2:D7"); profit.load("values"); await ctx.sync(); profit.values.forEach((row,i)=>{const cell=sheet.getRange("D"+(i+2)); cell.format.font.color=row[0]>=60000?"#1a7a4a":"#b94040"; cell.format.font.bold=row[0]>=60000;}); await ctx.sync(); });',
    },
    {
        "id":            "seg-8",
        "description":   "Add totals row and auto-fit",
        "sheet_context": ["A8:E8","A1:E8"],
        "explanation":   "Appends a dark TOTAL row with SUM formulas in B8:D8, then auto-fits all columns.",
        "predecessors":  ["seg-3","seg-5","seg-6"],
        "qa_pairs": [
            {"q":"Why autofitColumns last?","a":"All data must exist before fitting, otherwise columns are sized to empty cells."},
            {"q":"Why SUM not SUBTOTAL?","a":"SUBTOTAL respects filters, but for a simple 6-row table SUM is clearer and more expected."},
        ],
        "edit_suggestions": ["Change TOTAL label to GRAND TOTAL", "Use a different totals row color", "Add an average row below totals"],
        "parameters": [
            {"label": "Row label",        "key": "TOTAL",   "value": "TOTAL",   "type": "select", "options": ["TOTAL", "GRAND TOTAL", "SUM", "NET"]},
            {"label": "Background color", "key": "#1a1d27", "value": "#1a1d27", "type": "color"},
            {"label": "Font color",       "key": "#ffffff", "value": "#ffffff", "type": "color"},
        ],
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

STUB_EDIT = STUB_SEGMENTS[5:]  # Return seg-2 + seg-3 as the edited chain stub

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
  "qa_pairs":      [{"q":"Why ...?","a":"Because ..."},{"q":"...","a":"..."}],
  "edit_suggestions": ["Suggestion 1", "Suggestion 2", "Suggestion 3"],
  "parameters":    [{"label": "Human label", "key": "varName", "value": 42, "type": "number"}],
  "code":          "await Excel.run(async (ctx) => { ... await ctx.sync(); });",
}

Rules:
- predecessors: list ids of segments this one depends on semantically (can be empty [])
- qa_pairs: 2-3 Q&A pairs explaining design choices for this step
- edit_suggestions: 2-3 short prompts suggesting edits the user might want (e.g. "Change threshold to 50000", "Use red/green instead")
- parameters: ANY numeric/string constant hardcoded in the code that a user might want to tweak (thresholds, colors, counts, ranges). Each must have a unique key matching a variable or literal in the code, a human-readable label, its current value, and type. Types: "number" (numeric input), "color" (color picker, value must be a hex string like "#ff0000"), "select" (dropdown — also include an "options" array of string choices), "text" (plain text). Empty array [] if no meaningful parameters exist.
Worksheet context and optional are provided below.
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
"edit": """You are an Excel automation assistant. The user wants to modify a step in a multi-step spreadsheet automation.

You will receive:
- The segment to edit (the step the user is currently viewing)
- The remaining_segments that follow it (may be empty)
- The user's feedback describing the change

Your job is to:
1. Produce an updated version of the edited segment reflecting the user's feedback
2. Regenerate all remaining_segments so they are consistent with the edit

Respond with ONLY a JSON array of code segments starting from the edited step, followed by the regenerated remainder.
The array must have the same length as 1 + len(remaining_segments).
Each segment must have the same structure as segments from /code (id, description, explanation, code, sheet_context, qa_pairs, predecessors, edit_suggestions, parameters). Preserve and update parameters[] to reflect any new constants introduced.
Do NOT include markdown fences or any text outside the JSON array.
""" + f"Example single-segment edit response: {json.dumps([STUB_EDIT])}\n",

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

Generate 1-2 hard and 1-2 soft requirements. Hard = must-have correctness criteria. Soft = nice-to have criteria.
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
