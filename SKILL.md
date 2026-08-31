---
name: equity-research-report
description: Build a professional-grade fundamental + technical equity research report (PDF) plus a companion Excel model for a public company/ticker — the kind of depth you'd get from a real sell-side or buy-side analyst, not a surface-level summary or a first-page-of-Google writeup. Use this whenever the user asks for a stock or company analysis, an "equity research report," investment research, a "Bloomberg-style" or "Wall Street style" writeup, or just names a ticker and asks for a report/deep-dive/analysis on it — even if they don't spell out the format. Also trigger for "analyze this stock," "fundamental and technical analysis of X," "due diligence on Y," "build me a research report," "financial modeling," "build a DCF," or "build a 3-statement model."
---

# Equity Research Report (Fundamental + Technical)

Produces two deliverables — a multi-page PDF report and a companion Excel model with live formulas
— that mirror how a real analyst report is structured. The differentiator versus a generic writeup:
every driver is sourced, the single biggest valuation lever gets its own model, disagreement across
sources is surfaced rather than smoothed over, and settled facts are separated from estimates and
single-source claims.

## Reference files

Read these when the step calls for them, not upfront:

- `references/sec_edgar_fred.md` — pulling exact, filing-sourced financials from SEC EDGAR's XBRL
  API, peer multiples from the `frames` API, and macro inputs from FRED. Read during Step 1.
- `references/modeling_lessons.md` — the accumulated bugs that produce a clean-recalculating,
  wrong model. Read before restructuring any build script, and when a delivered model gets pushed
  back on.

## Before you start

1. **Confirm the ticker/company** and any obvious scope constraints, but don't block on
   clarifying questions the user hasn't asked — start researching immediately.
2. **Do the research before touching the output-format skills.** If `xlsx` and `pdf` skills are
   available in this environment, read their SKILL.md files only after you have the substantive
   content in hand — reading them first anchors you on document mechanics before you know what goes
   in the document. Those two are bundled on claude.ai but are **not** present in Claude Code; the
   templates here are self-contained, so their absence changes nothing except that the conventions
   sections below become the whole ruleset rather than a supplement.
3. **Pull historical financials from SEC EDGAR before secondary sources** if the company files with
   the SEC. Non-US filers, private companies and pre-revenue names may not have usable XBRL —
   fall back to general research for those.

## Step 1 — Research (parallel, sourced, conflict-aware)

Spawn multiple parallel research agents rather than one broad pass. A good split for a public
equity:

1. **Fundamentals & financials** — how the company *actually* makes money (not the one-line
   pitch), the last 4-8 quarters, guidance, balance sheet and capital structure, ownership and
   insider activity, analyst ratings and consensus, risks and catalysts flagged by management.
   Pull statement line items from EDGAR's XBRL API first — see the reference doc.
2. **Technical & market data** — price and trend, 52-week range, moving averages and
   support/resistance from recent chart commentary, volume, implied volatility, options
   positioning (put/call skew, open interest), short interest, institutional/insider ownership,
   and a chronological catalyst list with what moved the stock and why.
3. **Industry, competitive & regulatory landscape** — market sizing and growth, competitor share
   and recent shifts, the regulatory framework that actually governs this business and what
   changed in the last year (not a static description), emerging disruptors.
4. **Comps & valuation multiples** — a peer set with *current* EV/Revenue, EV/EBITDA and P/E,
   sourced same-day where possible. Cross-check across ≥2 sources: financial data sites serve
   stale cached snapshots, and a number that looks right on one source and is 2-3x off on another
   is a routine failure mode, not a rare one. The EDGAR `frames` API gives every peer measured
   over an identical period, which hand-built comps tables most often get wrong.

Instruct every research agent explicitly to:
- Cite a source URL for every material fact or number.
- Flag when a figure is an estimate, a projection, or an individual's back-of-envelope claim
  circulating on social media, rather than a disclosed fact.
- Report conflicting numbers side by side instead of silently picking one — YTD returns, market
  cap and exact dates are where sources disagree most.
- Note gaps explicitly ("could not confirm X — verify against primary source Y") rather than
  guessing.

**Identify the one variable that drives this company's valuation** before building — net interest
margin for a bank, reserve yield × circulation for a stablecoin issuer, net revenue retention for
SaaS, commodity price × production cost for a miner. That variable gets a dedicated section in both
deliverables. A dedicated driver model is what separates a professional report from a template dump
of generic ratios.

## Step 2 — Build

Read the `xlsx` and `pdf` skills' SKILL.md now if they exist, then adapt the bundled starter
scripts rather than writing from scratch. Every script marks replaceable content with `# TICKER-SPECIFIC`; keep the
helper and styling scaffolding as-is.

| Script | Use when |
|---|---|
| `build_xlsx_template.py` | Default workbook: cover, quarterly financials with live charts, a single-driver sensitivity grid, valuation, ratings, technicals, catalysts, sources. |
| `build_pdf_template.py` | The report itself. Section headers, data tables, stat boxes, disclaimer boxes, and line/bar/scenario charts already styled. |
| `build_scenario_projections_template.py` | 1-2 dominant drivers, but the user wants more than one grid: explicit multi-year Bear/Base/Bull, two independent valuation methods, a tied-out mini balance sheet, real unlevered FCF, an SBC/dilution schedule, a reverse DCF, and quantified competitive risk. |
| `build_3statement_dcf_template.py` | Valuation depends on the *interaction* of growth, margins, working capital and capital allocation (mega-cap tech, an industrial, a consumer staple), or the user explicitly asks for "financial modeling", a "DCF", or a "3-statement model". |
| `verify_model.py` | Every workbook, before delivery. See Step 3. |

Choosing between the three model templates: prefer the lighter single-driver grid when the company
has one cleanly-isolable driver — a full 3-statement build there is overkill and invites false
precision. For a genuinely comprehensive report, do both: the single-driver grid for intuition, the
full build for a defensible valuation range.

`build_xlsx_template.py` and `build_pdf_template.py` were extracted from a real Circle Internet
Group (CRCL) report — a stablecoin issuer with an unusually clean single-driver story. Swap in
whatever the target company's actual driver and structure are; don't force CRCL's line items onto a
company where they don't apply. `build_3statement_dcf_template.py` was built and debugged against
NVIDIA; `build_scenario_projections_template.py` came out of a CRCL rebuild done in response to a
structured review.

Before restructuring any of them, read `references/modeling_lessons.md`. Those scripts encode fixes
for bugs that took real debugging to find and that are easy to silently reintroduce.

## Report structure (mirror this in both the PDF and the Excel model)

1. **Cover / executive summary** — snapshot stats, a 3-4 paragraph thesis, what the report covers,
   the disclaimer.
2. **Business model mechanics** — how the company makes money in enough detail that the reader
   understands the unit economics, not just "they sell software." Call out the single biggest cost
   or revenue-share relationship if one dominates.
3. **Industry, competitive & regulatory landscape** — market sizing, competitor share, the
   regulatory environment and what's in motion, emerging disruptors. Separate live threats from
   speculative ones.
4. **Financial analysis** — a *quarterly* trend table and chart (the deceleration story lives in
   quarter-over-quarter detail, not annual totals), guidance, the dedicated driver model, balance
   sheet and capital structure.
5. **Valuation** — a comps table and chart (noting which comps are true comps vs. reference-only),
   and a bear/base/bull scenario table and chart. Label it a sanity-check framework unless you
   built a real DCF. Include sell-side targets and note the spread as a signal of how contested the
   valuation is.
6. **Technical analysis** — a price chart with moving averages, levels from recent chart
   commentary, volatility and options positioning, short interest and ownership, relative
   performance.
7. **Catalyst calendar** — historical catalysts with the stock's actual reaction, forward catalysts
   with why each matters.
8. **Key risks** — specific and named, not "competition" and "regulation."
9. **Sources & limitations** — every source, plus an explicit list of data-quality flags:
   conflicting figures, unconfirmed dates, estimates presented elsewhere as facts, anything needing
   primary-source verification.

## PDF conventions

See the `pdf` skill for library basics where it exists. Project-specific:

- Build with `reportlab.platypus` (`SimpleDocTemplate` + `Paragraph`/`Table`/`TableStyle`), not raw
  canvas drawing.
- **Include the charts.** `build_pdf_template.py` ships `line_chart`, `bar_chart` and
  `scenario_chart` helpers built on `reportlab.graphics` — no extra dependency, vector output. A
  quarterly trend, a comps bar with the subject highlighted, and a scenario chart with a dashed
  current-price line are what make a report read as analyst work rather than a wall of tables. A
  technical-analysis section with no price chart is the most conspicuous possible omission.
- Navy section-header bars, zebra-striped tables, an amber disclaimer box on the cover and closing
  page. One `PageBreak()` per major section. A footer with title and page number via
  `onFirstPage`/`onLaterPages`.
- `keepWithNext=1` on heading styles, so a sub-heading never strands at the foot of a page with its
  table or chart on the next one.
- **Render pages to PNG and look at them before delivering** (`pdftoppm -png -r 80 -f N -l N`, or
  PyMuPDF where pdftoppm isn't available). Text overflow, orphaned headers and chart labels
  colliding with bars are easy to introduce, easy to catch visually, and near-impossible to catch by
  reading the generation script.

## Excel conventions

See the `xlsx` skill for its ruleset where it exists (colors, number formats, recalculation) — the
templates already encode those conventions, so they are not a prerequisite. Project-specific
mechanics — never starting a text cell with `=`, never baking a grid axis into a formula string,
`DefinedName` for cross-sheet anchors, `ISBLANK`-aware ratio guards, tracking row numbers as
variables — are all in `references/modeling_lessons.md`, section 4. Beyond those:

- **Native openpyxl charts, not images.** They reference cells, so they follow when someone changes
  an input. A picture of a chart in a live model goes stale the first time anyone touches it.
- **Heat-map the sensitivity grids** with `ColorScaleRule`. The shape of the gradient tells the
  reader which driver dominates; a wall of similar numbers hides it.
- **Freeze panes** on every wide tab, so row labels stay visible while scrolling across periods.

## Step 3 — Verify and deliver

Both gates are blocking. Do not deliver on one alone.

1. **Recalculate** the workbook until `total_errors: 0`. On claude.ai this is
   `/mnt/skills/public/xlsx/scripts/recalc.py <file.xlsx> 60`. Elsewhere — Claude Code included —
   that script does not exist; use LibreOffice headless
   (`soffice --headless --convert-to xlsx --outdir <dir> <file.xlsx>`) or let step 2 do it, since
   `verify_model.py` evaluates every formula itself and reports errors as FAIL.
2. **Run `scripts/verify_model.py <file.xlsx> --price <current price>`** and resolve every FAIL.

The second gate exists because a clean recalculation proves formulas *evaluate*, not that the model
is right. It mechanically checks what SKILL.md used to ask you to check by eye: balance-sheet
tie-out to exactly 0 in every period, Bear ≤ Base ≤ Bull on every row, diluted ≥ basic shares,
upside/downside signs, and share prices within a sane multiple of the current price (which catches
$ vs $mm unit mismatches). It obtains values via LibreOffice if available, else the `formulas`
package, else cached values — so it works whether or not this environment ships a recalc tool.

Read its WARNs rather than clearing them: a Bull case with fewer shares than Bear is a correct
inversion, and the script cannot know that.

Then:
- Visually spot-check 2-3 PDF pages (cover, a chart-heavy page, the last page) by rendering to PNG.
- Deliver both files with a one-sentence summary — the files speak for themselves; don't re-narrate
  their contents in chat.

## Always include

- A disclaimer, on the cover and again at the end: not investment/legal/tax advice, built from
  public secondary sources at a point in time rather than a live market data terminal, and
  decision-critical figures should be verified against primary sources.
- A visible sources/limitations section. Don't bury data-quality caveats or omit them because
  they're inconvenient. Surfacing "these two sources disagree by 2x" is a feature of a professional
  report, not a flaw.
