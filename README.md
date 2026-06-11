# Portfolio AI — Financial Intelligence Pipeline

A Python pipeline that ingests heterogeneous brokerage PDFs — Vanguard, Merrill Lynch, Fidelity — normalizes 130+ positions across 23 accounts into a unified operational database, enriches with live market data, layers an AI operator interface for plain-English portfolio intelligence, and pushes real-time alerts to your phone.

Built for a real user managing a ~$5.9M multi-broker portfolio across retirement accounts, taxable accounts, HSAs, international assets, and cash instruments.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     STAGE 1 — INGESTION                     │
│                                                             │
│   Vanguard PDF    Merrill Lynch PDF    Fidelity PDF         │
│        │                 │                  │               │
│        └─────────────────┴──────────────────┘               │
│                          │                                  │
│                   read_pdf_text()                           │
│               (pypdf — raw text extraction)                 │
│                          │                                  │
│        ┌─────────────────┼──────────────────┐               │
│        │                 │                  │               │
│ parse_vanguard()  parse_benefits_  parse_fidelity()         │
│  23 accounts       online()          TOD · Trust            │
│  regex patterns    CTVA 401K · HSA   joint accounts         │
│        │                 │                  │               │
│        └─────────────────┴──────────────────┘               │
│                          │                                  │
│                   clean_numeric()                           │
│              (strips $, commas, nulls)                      │
│                          │                                  │
│              SQLite — vanguard_portfolio.db                 │
│        portfolio_summary · accounts · holdings              │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                  STAGE 2 — NORMALIZE + ENRICH               │
│                                                             │
│              fetch_market_data() via yfinance               │
│         beta · live price · dividend yield · ticker type    │
│    split-adjusted (VGT 8:1, MGK 5:1 — effective Apr 2026)  │
│                          │                                  │
│                    enrich_excel()                           │
│         unified output.xlsx — columns A through O           │
│         cols N & O are live Excel formulas (=J*E, =N-G)    │
│                          │                                  │
│              build_summary_sheet()                          │
│     KPIs · gain/loss vs cost basis · breakdown by type      │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│                   STAGE 3 — AI Q&A LAYER                    │
│                                                             │
│           User types a plain-English question               │
│                          │                                  │
│              load_portfolio_context()                       │
│         serialize SQLite → structured prompt string         │
│                          │                                  │
│           Anthropic API — claude-sonnet-4-6                 │
│     system prompt + portfolio data + question → answer      │
│                                                             │
│   "What's my tech exposure?"                                │
│   → "Tech represents 34% of the portfolio — VGT, VIGAX,    │
│      FSELX total $612k across 4 accounts."                  │
│                                                             │
│            Multi-turn · streaming · context-grounded        │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│               STAGE 4 — REAL-TIME ALERTS (ntfy)             │
│                                                             │
│         Every run → push notification to phone              │
│                                                             │
│   detect_significant_changes()                              │
│     diff current vs previous market_data_cache.json         │
│     price change ≥ ±1%  → alert with $ portfolio impact     │
│     yield change ≥ ±0.5pp → alert                           │
│                                                             │
│   send_ntfy_notification()                                  │
│     always sends: portfolio value + net gain/loss           │
│     urgent: fires when thresholds breached                  │
│                                                             │
│   Portfolio updated                                         │
│   Value:   $5,900,000  (+0.07% vs last run)                 │
│   Net G/L: +$860,000  vs cost basis                         │
└─────────────────────────────────────────────────────────────┘
```

---

## Files

| File | Purpose |
|------|---------|
| `consolidatedfinancialdatabaseproject.py` | Stages 1, 2 & 4 — PDF ingestion, SQLite normalization, market data enrichment, Excel output, ntfy alerts |
| `stage3_ai_qa.py` | Stage 3 — AI Q&A layer, Anthropic API integration, interactive CLI |

---

## Setup

**Requirements**

```bash
pip install pypdf yfinance pandas openpyxl numpy anthropic requests
```

**API key (Stage 3)**

```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

**ntfy push notifications (Stage 4)**

1. Install the [ntfy app](https://ntfy.sh) on iOS or Android
2. Subscribe to a topic name (e.g. `databaseautomation`)
3. Set `NTFY_TOPIC` at the top of `consolidatedfinancialdatabaseproject.py` to match

---

## Usage

**Step 1 — Place your brokerage PDFs in the project directory**

```
04182026 Holdings | Vanguard.pdf
04222026 Benefits OnLine | Investments.pdf
04272026 Fidelity Portfolio Positions.pdf
```

**Step 2 — Run Stages 1, 2 & 4**

```bash
python consolidatedfinancialdatabaseproject.py
```

```
Building unified positions matrix…
Fetching live market data (Beta, Price, Yield)…
  3-month T-bill rate: 3.63%
  Fetching market data for 32 tickers…
  Excel enriched → output.xlsx
Building portfolio summary sheet…
  Summary sheet written → output.xlsx
Checking for significant price/yield changes…
  Value:   $    5,900,000  (+0.07% vs last run)
  Net G/L: +$     860,000  vs cost basis
  No significant changes vs last run.
  ntfy notification sent → topic: databaseautomation
  Market data cached → market_data_cache.json
------------------------------------------------------------
FINISHED SUCCESSFULLY!
```

Outputs:
- `vanguard_portfolio.db` — normalized SQLite database
- `output.xlsx` — enriched positions matrix (columns A–O)
- `market_data_cache.json` — cached market data for next-run diff

**Step 3 — Run the AI Q&A layer**

```bash
python stage3_ai_qa.py
```

```
============================================================
  Portfolio AI — powered by Claude
============================================================
Portfolio loaded. Ask anything about the data.

You: What is my total tech exposure?

Claude: Tech holdings represent approximately 34% of the investable
portfolio. The primary positions are VGT ($273k, Saving Invested Dow),
VIGAX ($256k across two accounts), FSELX ($264k), and NVDA ($57k across
Fidelity and Robinhood). Combined market value is approximately $612,000
across four accounts.

You: Which account has the highest unrealized gain?
```

---

## How it works

### Stage 1 — Heterogeneous ingestion

Each brokerage produces a structurally different PDF. `read_pdf_text()` extracts raw text via `pypdf.PdfReader`. Three dedicated parsers (`parse_vanguard`, `parse_benefits_online`, `parse_fidelity`) apply regex patterns tuned to each brokerage's layout, then insert normalized rows into a shared SQLite schema: `portfolio_summary`, `accounts`, and `holdings`. `clean_numeric()` strips all currency formatting into clean floats.

### Stage 2 — Live enrichment

`fetch_market_data()` queries yfinance for beta, live price, and dividend yield across all tradeable symbols using a three-layer fallback: `trailingAnnualDividendYield` → `rate/price` computation → `dividendYield` with format detection, hard-capped at 15% to filter bad data. Non-market instruments (cash, money markets, T-bills, HSA balances) receive the current 3-month T-bill rate as yield. `enrich_excel()` writes columns I–O into the unified spreadsheet with live Excel formulas in N (`=J*E`) and O (`=N-G`). Holdings are split-adjusted for the April 21, 2026 Vanguard splits (VGT 8:1, MGK 5:1).

### Stage 3 — AI operator layer

`load_portfolio_context()` serializes the entire SQLite database into a structured text block — every account, every holding, cost basis, and balance. This context is injected into the Anthropic API alongside the user's question. Claude reasons over the actual numbers, not its training data. Conversation history is maintained across turns so follow-up questions resolve correctly. Responses stream word-by-word.

### Stage 4 — Real-time alerts

`detect_significant_changes()` diffs current market data against `market_data_cache.json` from the previous run. Price moves ≥ ±1% or yield changes ≥ ±0.5 percentage points trigger alerts with per-position dollar impact. The function also computes full portfolio totals (value, gain/loss, % change) so everything is self-contained. `send_ntfy_notification()` fires on every run — quiet ping for routine updates, high-priority buzz for threshold breaches.

---

## Why this architecture

Brokerage data is siloed by design. Each institution exports a different format, uses different field names, and structures its PDFs differently. The core challenge is the same one faced by any financial data platform: ingest from heterogeneous sources, normalize to a common schema, and make the unified data queryable by non-technical operators.

The AI layer demonstrates that LLMs become genuinely useful for financial decision-making when grounded in proprietary data — not when asked to answer from training data alone. The alert system closes the loop: the pipeline doesn't just produce a report, it watches the portfolio and surfaces what matters.

---

## Limitations & known constraints

- PDF parsing is layout-sensitive. Brokerage statement format changes will break regex patterns and require parser updates.
- `pypdf.extract_text()` reconstructs reading order from `(x, y)` coordinates, which can misorder multi-column tables. Parsers are written defensively for this reason.
- yfinance beta is unavailable for ETFs and mutual funds via `info['beta']`. The fallback computes 3-year rolling beta from monthly returns against SPY.
- The Anthropic API context window limits prompt length. At ~130 positions the context is well within limits, but portfolios with 500+ positions would require chunking or summarization.
- ntfy push notifications require the ntfy app subscribed to the configured topic. The pipeline continues without error if the notification fails.

---

## Author

Tanuj Mangalam · Purdue University, Computer Engineering · [GitHub](https://github.com/)
