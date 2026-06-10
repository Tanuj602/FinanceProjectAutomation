# Portfolio AI — Financial Intelligence Pipeline

A Python pipeline that ingests heterogeneous brokerage PDFs — Vanguard, Merrill Lynch, Fidelity — normalizes 130+ positions across 23 accounts into a unified operational database, enriches with live market data, and layers an AI operator interface for plain-English portfolio intelligence.

Built for a real user managing a multi-broker portfolio across retirement accounts, taxable accounts, HSAs, international assets, and cash instruments.

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
│                          │                                  │
│                    enrich_excel()                           │
│         unified output.xlsx — columns A through O          │
│                          │                                  │
│              build_summary_sheet()                          │
│         KPIs · gain/loss vs snapshot · breakdown by type    │
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
```

---

## Files

| File | Purpose |
|------|---------|
| `databaseproject.py` | Stages 1 & 2 — PDF ingestion, parsing, SQLite normalization, market data enrichment, Excel output |
| `stage3_ai_qa.py` | Stage 3 — AI Q&A layer, Anthropic API integration, interactive CLI |

---

## Setup

**Requirements**

```bash
pip install pypdf yfinance pandas openpyxl numpy anthropic
```

**API key**

```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

---

## Usage

**Step 1 — Place your brokerage PDFs in the project directory**

```
04182026 Holdings | Vanguard.pdf
04222026 Benefits OnLine | Investments.pdf
04272026 Fidelity Portfolio Positions.pdf
```

**Step 2 — Run Stages 1 & 2**

```bash
python databaseproject.py
```

Outputs:
- `vanguard_portfolio.db` — normalized SQLite database
- `output.xlsx` — enriched positions matrix with live market data

**Step 3 — Run Stage 3**

```bash
python stage3_ai_qa.py
```

```
============================================================
  Portfolio AI — powered by Claude
============================================================
Portfolio loaded. Ask anything about the data.

Example questions:
  • What is my total tech exposure?
  • Which account has the highest unrealized gains?
  • How much is in dividend ETFs vs growth ETFs?
  • What percentage of the portfolio is international?
  • Which single position is the largest holding?

You: What is my total tech exposure?

Claude: Tech holdings represent approximately 34% of the investable
portfolio. The primary positions are VGT ($273k, Saving Invested Dow),
VIGAX ($256k across two accounts), FSELX ($264k), and NVDA ($57k across
Fidelity and Robinhood. Combined market value is approximately $612,000
across four accounts.
```

---

## How it works

### Stage 1 — Heterogeneous ingestion

Each brokerage produces a structurally different PDF. `read_pdf_text()` extracts raw text via `pypdf.PdfReader`. Three dedicated parsers (`parse_vanguard`, `parse_benefits_online`, `parse_fidelity`) apply regex patterns tuned to each brokerage's layout, then insert normalized rows into a shared SQLite schema: `portfolio_summary`, `accounts`, and `holdings`. `clean_numeric()` strips all currency formatting into clean floats.

### Stage 2 — Live enrichment

`fetch_market_data()` queries yfinance for beta, live price, and dividend yield across all tradeable symbols. Non-market instruments (cash, money markets, T-bills, HSA balances) are classified and handled separately. `enrich_excel()` writes columns I–O into the unified spreadsheet. `build_summary_sheet()` computes KPIs against the April snapshot.

### Stage 3 — AI operator layer

`load_portfolio_context()` serializes the entire SQLite database into a structured text block — every account, every holding, cost basis, and balance. This context is injected into the Anthropic API alongside the user's question. Claude reasons over the actual numbers, not its training data. Conversation history is maintained across turns so follow-up questions resolve correctly.

---

## Why this architecture

Brokerage data is siloed by design. Each institution exports a different format, uses different field names, and structures its PDFs differently. The core challenge is the same one faced by any financial data platform: ingest from heterogeneous sources, normalize to a common schema, and make the unified data queryable by non-technical operators.

The AI layer demonstrates that LLMs become genuinely useful for financial decision-making when grounded in proprietary data — not when asked to answer from training data alone.

---

## Limitations & known constraints

- PDF parsing is layout-sensitive. Brokerage statement format changes will break regex patterns and require parser updates.
- `pypdf.extract_text()` reconstructs reading order from `(x, y)` coordinates, which can misorder multi-column tables. Parsers are written defensively for this reason.
- yfinance beta is unavailable for ETFs and mutual funds via `info['beta']`. The fallback computes 3-year rolling beta from monthly returns against SPY.
- The Anthropic API context window limits prompt length. At ~130 positions the context is well within limits, but portfolios with 500+ positions would require chunking or summarization.

---

## Author

Tanuj Mangalam · Purdue University, Computer Engineering · [GitHub](https://github.com/)
