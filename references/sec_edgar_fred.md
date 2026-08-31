# SEC EDGAR & FRED — structured primary-source data

Use this alongside general web research during Step 1 (Fundamentals & financials). It closes the
biggest data-quality gap this skill has hit in practice: secondary-source aggregators
(stockanalysis.com and similar) give press-release-level summaries that often don't sum to the
company's actual reported totals — that's exactly why the 3-statement DCF template has
"reconciling plug" rows. SEC EDGAR's XBRL API returns the precise figures tagged in the filing
itself, so pulling from it first can shrink or eliminate the need for that plug, and gives every
historical figure a citation stronger than "aggregator site, accessed [date]": an actual filing,
form type, and period.

## Critical: send a User-Agent header

SEC's fair-access policy **rejects any request that does not declare a `User-Agent`**, and it
rejects it with a bare `403` that looks exactly like a firewall block. This is the single most
common reason an EDGAR pull "doesn't work," and it is trivially fixable:

```bash
# 403 — no User-Agent
curl -s -o /dev/null -w '%{http_code}\n' \
  https://data.sec.gov/api/xbrl/companyconcept/CIK0001045810/us-gaap/Assets.json
# 200 — identical URL, User-Agent declared
curl -s -o /dev/null -w '%{http_code}\n' -A 'equity-research-skill contact@example.com' \
  https://data.sec.gov/api/xbrl/companyconcept/CIK0001045810/us-gaap/Assets.json
```

Both lines above were run directly; they return `403` and `200` respectively. SEC asks that the
User-Agent identify you and give a contact address. Use a project/tool name plus a contact you
are willing to publish — do not put the user's personal email address in an outbound header
unless they have asked you to.

Also respect the **10 requests/second** rate limit. Exceeding it gets your IP throttled, and the
throttle response is another opaque error. A `time.sleep(0.15)` between calls in a loop is enough.

### Prefer a script; fall back to WebFetch

Pull EDGAR data with `bash`/`curl`/`python requests` and parse the JSON directly. That is faster
than one `WebFetch` per tag (a 15-tag pull is one script run instead of 15 sequential calls) and
strictly more accurate, because `WebFetch` returns a *model's summary* of the page — a number
that has been read and retyped rather than parsed. For financial statement line items that
distinction matters.

**If** the environment you are running in genuinely blocks outbound network access from the
shell, `WebFetch` on the same URLs works and is the correct fallback. Test which situation you
are in with one call before committing to an approach:

```bash
curl -s -o /dev/null -w '%{http_code}\n' -A 'equity-research-skill contact@example.com' \
  https://www.sec.gov/files/company_tickers.json
```

`200` means script the whole pull. Anything else — and only then — fall back to per-URL
`WebFetch` calls. (An earlier version of this document asserted that `sec.gov` was unreachable
from the sandboxed shell and mandated `WebFetch` for every call. That conclusion was drawn from
a `403` that was actually the missing User-Agent above, so the ban was both wrong and expensive.
Verify, don't assume — in either direction.)

## Step 1 — ticker → CIK

GET `https://www.sec.gov/files/company_tickers.json` and find the entry whose `ticker` matches. Zero-pad `cik_str` to 10 digits for the
endpoints below (e.g. `1045810` → `CIK0001045810`).

## Step 2 — filing history / fiscal year end / most recent filings

GET `https://data.sec.gov/submissions/CIK{10-digit}.json`. Read out: the
company's registered name, its fiscal year end, and the most recent 10-K and most recent 10-Q —
each with accession number, filing date, and period-of-report (the actual fiscal period covered,
which is what you need for the model's period labels, not the filing date).

## Step 3 — pull specific line items across years

GET `https://data.sec.gov/api/xbrl/companyconcept/CIK{10-digit}/us-gaap/{TAG}.json`, one call per
tag. Filter `units.USD` to `form == "10-K"` and take the most recent entries by `end` date.

```python
import requests, time
UA = {"User-Agent": "equity-research-skill contact@example.com"}
def concept(cik10, tag):
    r = requests.get(f"https://data.sec.gov/api/xbrl/companyconcept/{cik10}/us-gaap/{tag}.json",
                     headers=UA, timeout=30)
    time.sleep(0.15)                      # stay under the 10 req/s limit
    if r.status_code == 404:
        return None                       # this filer does not use this tag — see note below
    r.raise_for_status()
    rows = [u for u in r.json()["units"]["USD"] if u.get("form") == "10-K"]
    # Dedupe: the same period is restated in later filings, so keep the newest `filed` per `end`.
    best = {}
    for u in rows:
        if u["end"] not in best or u["filed"] > best[u["end"]]["filed"]:
            best[u["end"]] = u
    return sorted(best.values(), key=lambda u: u["end"], reverse=True)
```

**The dedupe matters.** A `companyconcept` response carries the same fiscal period multiple
times, once per filing that reported it — a naive "last 3 entries" grab silently mixes restated
and original figures, and the mismatch surfaces later as a balance sheet that will not tie.

Useful `us-gaap` tags, organized by where they land in the 3-statement template:

**Income statement**
- `Revenues` (older/simpler filers) or `RevenueFromContractWithCustomerExcludingAssessedTax`
  (most filers post-ASC 606) — try both if one 404s
- `CostOfRevenue` or `CostOfGoodsAndServicesSold`
- `ResearchAndDevelopmentExpense`
- `SellingGeneralAndAdministrativeExpense`
- `OperatingIncomeLoss`
- `InterestExpense` / `InvestmentIncomeInterest` (or a single net figure if the filer reports one)
- `IncomeTaxExpenseBenefit`
- `NetIncomeLoss`
- `EarningsPerShareDiluted`
- `WeightedAverageNumberOfDilutedSharesOutstanding`

**Balance sheet**
- `Assets`, `AssetsCurrent`
- `CashAndCashEquivalentsAtCarryingValue`, `ShortTermInvestments`
- `AccountsReceivableNetCurrent`
- `InventoryNet`
- `PropertyPlantAndEquipmentNet`
- `Goodwill`, `IntangibleAssetsNetExcludingGoodwill`
- `Liabilities`, `LiabilitiesCurrent`
- `AccountsPayableCurrent`
- `LongTermDebtNoncurrent`
- `StockholdersEquity`

**Cash flow**
- `DepreciationDepletionAndAmortization`
- `ShareBasedCompensation`
- `PaymentsForRepurchaseOfCommonStock`
- `PaymentsOfDividends`
- `PaymentsToAcquirePropertyPlantAndEquipment`
- `NetCashProvidedByUsedInOperatingActivities`

Every tag above was spot-checked live against NVIDIA's filings while writing this doc —
`Assets` returned $206,803M for FY2026A (matching the model's anchor period exactly),
`ShareBasedCompensation` returned $6,386M for FY2026A (also an exact match to figures used
elsewhere in this skill's example model). Tag names are standardized but not universal: a bank
won't file `InventoryNet`, some companies use firm-specific extension tags instead of the
standard `us-gaap` ones, and non-US filers on 20-F may not use `us-gaap` at all. If a
`companyconcept` URL comes back empty/404, that filer doesn't use that tag — fall back to reading
the actual filing or a secondary source for that one line item, and note the gap explicitly in
the Sources & Assumptions tab rather than silently guessing or leaving it blank.

## Step 3b — peer multiples in one call (the `frames` API)

For the comps/valuation research agent, `companyconcept` is the wrong shape: it is one company
across time, when what you need is one metric across companies at a point in time. That is what
`frames` does — one call returns every filer that reported a given tag for a given period:

    https://data.sec.gov/api/xbrl/frames/us-gaap/Assets/USD/CY2025Q4I.json

Period codes: `CY2025Q4I` for an instantaneous (balance sheet) value at that quarter end,
`CY2025Q4` for a duration (income statement) value over that quarter, `CY2025` for a full year.
Each entry carries `cik`, `entityName`, and `val`. Filter to your peer CIKs and you have a
consistent, filing-sourced peer set for the denominator of every multiple — no aggregator, no
stale cache, and every peer measured over the identical period, which is the thing hand-built
comps tables most often get wrong.

Verified live: the URL above returns `200`.

## Step 4 (optional) — browse all tags at once

`https://data.sec.gov/api/xbrl/companyfacts/CIK{10-digit}.json` returns every tagged fact for the
company in one large payload. Only reach for this if you need to discover an unfamiliar tag name
for a line item that isn't in the list above — prefer the targeted `companyconcept` calls for
anything you already know you need, since the full companyfacts payload is large and slower to
work through.

## Risk-free rate and other macro inputs (FRED)

FRED serves a plain CSV with **no API key required**, and it accepts multiple series in one call:

    https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10,DGS3MO

```bash
curl -s 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10,DGS3MO' | tail -5
```

The last non-blank row is the most recent observation. Blank values are non-trading days — skip
them rather than treating a blank as zero.

Series worth knowing: `DGS10` (10-year Treasury, the standard risk-free rate for a DCF),
`DGS3MO` / `DGS1MO` (short rates — the right anchor for a cash-yield assumption),
`DGS30`, `DFF` (effective fed funds), `T10YIE` (10-year breakeven inflation),
`BAMLH0A0HYM2` (high-yield OAS, a useful credit-stress read for a levered name).

If the shell has no network access, `WebFetch` the human-readable `https://fred.stlouisfed.org/series/DGS10`
page instead, which shows the latest observation in the page text. Do NOT `WebFetch` the CSV
endpoint — it comes back as unparsed binary to the summarizing model. That is a limitation of
fetching-through-a-model, not of the endpoint: by script the CSV is the better source.

Whatever you pull, record the value AND its observation date in the model's sources tab. A
risk-free rate without a date is not a citation.

## What this fixes, and what it doesn't

Pulling `Assets`, `Liabilities`, and `StockholdersEquity` directly from XBRL gives you the real,
exact reported totals to reconcile the itemized balance sheet against — you may still need a
reconciling-plug line if you don't itemize every sub-line item (most secondary research won't),
but now it's sized against a verified total instead of an estimated one, which is a meaningfully
stronger place to be.

This does **not** replace reading the filing for forward guidance, segment detail, or qualitative
risk factors — XBRL only covers tagged, quantitative face-of-financial-statements data. It's also
not a live feed: `data.sec.gov` only reflects what's actually been filed as of your research date,
so the current quarter won't appear until that 10-Q is filed (same staleness caveat that applies
to any secondary source, just with better precision on what *has* been filed).
