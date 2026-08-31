# equity-research-report

A Claude Agent Skill that turns a ticker into two deliverables: a multi-page equity research PDF and a companion Excel model with live formulas. Financials come from SEC EDGAR's XBRL API rather than scraped summaries, and the workbook has to pass two blocking verification gates before it ships.

The target output is what a sell-side or buy-side analyst would hand you, not a formatted summary of the first page of search results.

## What it produces

**PDF report** — nine sections built with `reportlab.platypus`: cover and executive summary, business model mechanics, industry/competitive/regulatory landscape, financial analysis, valuation, technical analysis, catalyst calendar, key risks, and a sources-and-limitations section. Vector line, bar and scenario charts are generated inline, no image embedding. Pages are rendered to PNG and visually checked before delivery.

**Excel model** — native openpyxl charts that reference cells and update when inputs change, heat-mapped sensitivity grids, frozen panes on wide tabs, and depending on which template fits: a single-driver grid, explicit multi-year Bear/Base/Bull projections with two independent valuation methods, or a full 3-statement model with DCF, unlevered FCF, an SBC/dilution schedule and a reverse DCF.

## How it works

**Step 1 — Research.** Four parallel research passes: fundamentals and financials, technical and market data, industry/competitive/regulatory landscape, and comps and valuation multiples. EDGAR XBRL is queried first for anything the company has actually filed; the `frames` API pulls peer multiples measured over an identical period, which is what hand-built comps tables usually get wrong. Every material number carries a source URL. Conflicting figures are reported side by side instead of silently resolved, and gaps are stated rather than filled in.

Before anything gets built, the skill identifies the single variable that drives the company's valuation — net interest margin for a bank, reserve yield times circulation for a stablecoin issuer, net revenue retention for SaaS, commodity price against production cost for a miner. That variable gets a dedicated model in both deliverables.

**Step 2 — Build.** Four starter scripts, each marking replaceable content with `# TICKER-SPECIFIC` so the styling and helper scaffolding stays intact. Template choice depends on whether the company has one cleanly isolable driver or a valuation that turns on the interaction of growth, margins, working capital and capital allocation.

**Step 3 — Verify.** Two gates, both blocking. The workbook recalculates to zero errors, then `verify_model.py` evaluates every formula itself and checks balance-sheet tie-out to exactly 0 in every period, Bear ≤ Base ≤ Bull on every row, diluted ≥ basic shares, upside/downside sign consistency, and share prices within a sane multiple of the current price (which is how $ vs $mm unit mismatches get caught). A clean recalculation only proves formulas evaluate; it says nothing about whether the model is right.

## Contents

| Path | Lines | Purpose |
|---|---|---|
| `SKILL.md` | — | The workflow: research protocol, report structure, PDF and Excel conventions, verification gates |
| `scripts/build_pdf_template.py` | 644 | Report generator with styled section headers, tables, stat boxes, disclaimer boxes and chart helpers |
| `scripts/build_xlsx_template.py` | 587 | Default workbook: quarterly financials with live charts, single-driver sensitivity grid, valuation, technicals, catalysts, sources |
| `scripts/build_scenario_projections_template.py` | 1559 | Multi-year Bear/Base/Bull, two valuation methods, tied-out mini balance sheet, unlevered FCF, dilution schedule, reverse DCF |
| `scripts/build_3statement_dcf_template.py` | 1081 | Full 3-statement build with DCF, for companies where the drivers interact |
| `scripts/verify_model.py` | 367 | Mechanical correctness checks; runs against every workbook before delivery |
| `references/sec_edgar_fred.md` | 206 | Ticker to CIK, XBRL line-item pulls, the `frames` peer API, FRED macro inputs, and the User-Agent header that EDGAR 403s without |
| `references/modeling_lessons.md` | 171 | Bugs that produce a clean-recalculating, wrong model — SBC hitting equity, interest running off prior balances, terminal value share, openpyxl formula-string traps |

The PDF and default XLSX templates were extracted from a Circle Internet Group (CRCL) report, the 3-statement template was built and debugged against NVIDIA, and the scenario template came out of a CRCL rebuild following a structured review. Swap in the target company's actual drivers; don't force CRCL's line items onto a business where they don't apply.

## Install

**Claude Code** — clone into personal scope:

```bash
git clone https://github.com/USER/equity-research-report ~/.claude/skills/equity-research-report
```

Or into project scope at `.claude/skills/` if you want it committed alongside a repo. Claude Code scans both directories at session start and loads it automatically.

**claude.ai** — zip the skill folder and upload it as a custom skill in settings. See the [Agent Skills docs](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview) for the current upload path.

Then ask for it in plain language: *"build me an equity research report on NVDA"*, *"fundamental and technical analysis of CRCL"*, *"build a 3-statement model for CAT"*.

## Requirements

- Python with `openpyxl` and `reportlab`
- Network access to `sec.gov` and `fred.stlouisfed.org` for primary-source data
- Optional: LibreOffice headless or the `formulas` package for recalculation; `pdftoppm` or PyMuPDF for page rendering

The bundled `xlsx` and `pdf` skills exist on claude.ai but not in Claude Code. The templates here are self-contained, so their absence changes nothing except that the conventions sections in `SKILL.md` become the whole ruleset rather than a supplement.

## Limitations

Built from public secondary sources and filings at a point in time, not from a live market data terminal. Non-US filers, private companies and pre-revenue names often have no usable XBRL, so those fall back to general research. Options positioning, short interest and technical levels come from published commentary and go stale quickly.

## Disclaimer

Output is not investment, legal or tax advice. Every report carries this on the cover and again at the end, along with a data-quality section listing conflicting figures, unconfirmed dates and anything needing primary-source verification. Verify decision-critical numbers against the filings yourself.

## License

MIT.
