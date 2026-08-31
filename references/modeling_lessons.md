# Modeling lessons — bugs that a clean recalculation will not catch

Read this **before restructuring any of the build scripts**, and when a delivered model has been
pushed back on. Everything here was caught in a real build, in production, usually after the
number had already been written into a report. The scripts encode the fixes; this document
explains why they are shaped the way they are, so you don't "simplify" one away.

The organising principle: **a formula that evaluates is not a formula that is right.** Every
item below produces a workbook with zero formula errors and a wrong answer.

---

## 1. Structural accounting — the 3-statement template

### Stock-based compensation must hit equity, not just cash
SBC is a non-cash expense added back in the cash flow statement (correct), but it must *also*
appear in the equity roll-forward:

    equity_close = equity_open + net income + SBC − buybacks − dividends

because SBC credits additional paid-in capital in real accounting. Omit the equity-side add-back
and the balance sheet gap grows by exactly the SBC amount every period — a distinctive,
diagnostic symptom. If your balance check is off by a constant that matches an SBC line, this is
why.

### Any balance-sheet line that jumps needs an offsetting cash-flow line
If debt increases (a bond issuance), that increase must appear as a financing inflow
(`+ Net debt issuance/(repayment)`, formula-driven off the period-over-period balance sheet
delta). Otherwise cash does not reconcile and the balance sheet breaks from that period forward.
The general rule: **for every balance-sheet delta there is a cash-flow line, or the item is held
flat.** Held flat is a legitimate simplification; unexplained movement is not.

### Interest must run off *prior*-period balances
Computing interest income on the same period's closing cash makes cash depend on interest which
depends on cash — a circular reference this toolchain's recalc step cannot resolve. Use the
beginning-balance convention (prior period's ending cash and debt). It is standard practice, it
is acyclic by construction, and it avoids needing iterative calculation turned on.

Do not leave interest income hardcoded as a flat or declining series while the modelled cash
balance compounds. A model that throws off billions in cumulative FCF and shows interest income
*falling* is visibly unowned, and it is the first thing a reader with a finance background
notices.

### Summary-level research data rarely sums to the reported total
Press-release-level balance sheet detail is almost always missing smaller line items (prepaid
expenses, deferred tax assets, lease liabilities). Rather than silently under-stating totals, add
an explicit, honestly-labelled **reconciling plug** sized so the historical anchor period matches
the reported total. Pull `Assets`, `Liabilities` and `StockholdersEquity` from SEC EDGAR's XBRL
API first (see `sec_edgar_fred.md`) so the plug is sized against a verified total rather than an
estimate. Replace with fully itemized 10-K data for production use.

### Pass-through / custodial balances never tie exactly
If the company carries a real-world pass-through item — reserve assets vs. a stablecoin's
circulation liability, a bank's deposits vs. cash, an insurer's segregated assets — the reported
asset and liability figures for it will not match exactly. Reclassify that real anchor-period gap
into "other operating assets" **before** forcing the two pass-through lines equal in the forward
model, or the balance sheet carries a permanent, silent hole. Hand-derive the balance identity
from the real anchor figures before running the build script.

---

## 2. Projection logic — the scenario template

### Don't grow an already-elapsed period's run-rate by a full additional year
Modelling the current fiscal year mid-way through, with some quarters of actuals in hand:
anchoring "this year's" projection to the in-hand run-rate and *then* applying a full year's
growth rate double-counts growth that has already happened. A real early draft had its **Bear**
case showing higher revenue than the prior full fiscal year, from an intentionally conservative
assumption, purely from this.

Anchor growth-rate-driven projections to the last genuinely **complete** prior period's actual —
matching how the company itself discloses YoY growth — and derive an "average balance for the
year" as the average of beginning and ending balance rather than conflating a mid-year snapshot
with a full-year anchor.

### Ground year one against the hardest available data
The first forecast year's revenue growth should be checked against the most recent actual quarter
plus next-quarter guidance, annualized — not an arbitrary round number or an unconfirmed
consensus figure. Building the NVIDIA demo, a "conservative-looking" growth assumption turned out
to imply a *lower* full-year revenue figure than the company's own already-reported H1 actual plus
guidance summed to.

### Bear < Base < Bull is a claim to be checked, not a property you get for free
Read the actual output values at every driver and every output line. Cost lines and share counts
legitimately invert (a bigger Bull buyback means fewer Bull shares); revenue, EBITDA and price
targets do not. `verify_model.py` flags every inversion as a WARN precisely because the call
requires reading.

---

## 3. Valuation

### Reconcile the two methods — don't just flag the gap
On a delivered model, the exit-multiple method and the DCF diverged in the Base case (~$53 vs
~$32) because the exit multiple was picked independently of what the DCF's own WACC/terminal-
growth pair implies as a fair terminal multiple. A note saying "if these diverge, check
consistency" is not enough. **Calibrate:** back out the DCF-implied terminal multiple (terminal
value grossed up by `(1+WACC)^N`, divided by Year-N EBITDA) and set the Base-case exit multiple
close to it. Leave Bear/Bull with a real spread around that anchor — those should diverge on
genuinely different growth/quality assumptions, not on an uncalibrated multiple.

### Watch the terminal value share
With a short explicit forecast window, Gordon-growth terminal value routinely carries 80-90% of
enterprise value, and the DCF becomes an exit-multiple opinion wearing a DCF's clothing. Both
templates now print "Terminal value as % of EV" as a visible diagnostic. Read it. Above ~85%,
either extend the horizon or say plainly in the report that Method 1 is carrying the valuation
argument.

### Sign conventions
Upside/downside must be `(target / current) - 1`, **not** `(current / target) - 1`. The inverted
version rendered a Bear-case target far below the current price as "+529% upside" instead of a
~-84% downside. A recalc error count will never catch this — the formula is valid, just wrong. A
target below the current price reads as negative, always.

### Blend, weight, and say what the market is pricing in
Once reconciled, add a **Blended** row (average of the two methods), an **Upside/(downside) to
blended** row, and a **probability-weighted target** (e.g. 25/50/25, weights stated explicitly as
the model's own judgment) to both the Valuation tab and the Cover dashboard.

Add a 2-3 sentence qualitative note on the Cover dashboard — not only in the PDF — stating plainly
which scenario the current price looks closest to, and what would have to be true (which driver or
margin assumption, moving which direction) for the current price to be justified. This is the
single highest-value qualitative sentence on the dashboard and the easiest to skip when the
dashboard is treated as numbers-only.

### Wire the biggest risk into the Bear case
Don't leave the main competitive/distribution risk as a disconnected sensitivity table. Put a
short table on the risk-quantification tab showing the Bear case's key driver and margin
assumptions side by side with Base, so it is visible that Bear already embeds the risk being
quantified elsewhere. This beats building a fourth scenario column throughout the model.

---

## 4. openpyxl mechanics

### Never start a plain-text cell with `=`
A note like `"= price x shares out"` is parsed as a formula and throws `#VALUE!`. Write
`"Formula: price x shares out"`. Three separate `#VALUE!` errors in one build traced back to this
one pattern.

### Never bake a grid axis into the formula string
If a sensitivity grid writes its axis value both as a visible label *and* as a Python literal
inside the formula, editing the label silently changes nothing — the grid looks adjustable and
isn't. Write the axis as a number in a cell and reference it (`$B12`, `C$5`). This bug shipped in
two separate grids in this skill before being caught.

### Guard Gordon growth where g ≥ WACC
`FCF/(WACC−g)` with g above WACC produces a large *negative* price, not an error, so `IFERROR`
does not catch it. Wrap the cell in `IF(g>=wacc,"n/m",...)`.

### Wrap margin/ratio formulas against blanks, not just errors
Use `IF(ISBLANK(...),"n/a",IFERROR(...))` rather than bare `IFERROR`, so a genuinely unreported
quarter renders `"n/a"` instead of a misleading `0%`. Division against a blank numerator is not a
formula error, so `IFERROR` alone won't catch it.

### Track row numbers as variables, never hand-counted offsets
An off-by-one in a manually-counted row index silently overwrites the wrong row with the wrong
formula. This happened building a Cover tab — a market-cap formula landed on the enterprise-value
row — and was caught only by inspecting cached values after recalculation.

### Use `DefinedName` for cross-sheet anchors
Current price, shares outstanding, market cap. Other tabs then reference `=Cover_Price*Cover_Shares`
instead of a hardcoded coordinate that breaks the moment a row is inserted above it.

### Excel's 31-character sheet-name cap is enforced by silent truncation
A longer name is truncated on save, and every formula referencing the untruncated name becomes
`#REF!` — errors that look unrelated to the real cause. Check every `create_sheet()` argument.

### Watch for column-letter self-references
A formula that accidentally references its own column produces silent `#VALUE!` errors. Comment
cross-year references explicitly when building them in a loop.
