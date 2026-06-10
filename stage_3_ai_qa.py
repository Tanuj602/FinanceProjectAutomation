"""
stage3_ai_qa.py  —  Portfolio AI Q&A Layer
============================================
Builds on top of databaseproject.py (Stages 1 & 2).

Run databaseproject.py first to generate vanguard_portfolio.db.
Then run this file:  python stage3_ai_qa.py

The script loads the normalized portfolio from SQLite, serializes it
as context, and lets you ask plain-English questions via the Anthropic API.
Claude reasons over the actual numbers — not its training data.
"""

import os
import sqlite3
import anthropic

# ── Config ────────────────────────────────────────────────────────────────────

DB_PATH    = "vanguard_portfolio.db"
MODEL      = "claude-sonnet-4-6"
MAX_TOKENS = 1024

SYSTEM_PROMPT = """You are a personal financial assistant for a multi-account investment portfolio.

You have been given a complete snapshot of the portfolio — every account, every holding,
symbol, quantity, cost basis, current balance, and ticker type.

Your job:
- Answer questions accurately using ONLY the data provided. Never invent numbers.
- Be specific: name the accounts and tickers that drive your answer.
- Format dollar values with commas (e.g. $1,234,567). Format percentages to 1 decimal place.
- If a question cannot be answered from the data, say so clearly.
- Keep answers concise — 3 to 6 sentences unless a breakdown is explicitly requested.

You are talking to the portfolio owner. Treat the data as current and authoritative."""


# ── Database helpers ──────────────────────────────────────────────────────────

def load_portfolio_context(db_path: str) -> str:
    """
    Reads the SQLite database produced by databaseproject.py and serializes
    it into a structured text block suitable for injection into a prompt.
    Returns a string Claude can reason over.
    """
    if not os.path.exists(db_path):
        raise FileNotFoundError(
            f"Database not found: {db_path}\n"
            "Run databaseproject.py first to generate the portfolio database."
        )

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row

    # ── Portfolio summary ─────────────────────────────────────────────────
    summary = conn.execute(
        "SELECT as_of_date, total_net_worth FROM portfolio_summary ORDER BY id DESC LIMIT 1"
    ).fetchone()

    # ── All accounts ──────────────────────────────────────────────────────
    accounts = conn.execute(
        "SELECT account_id, account_name, total_balance FROM accounts ORDER BY total_balance DESC"
    ).fetchall()

    # ── All holdings with account name joined ────────────────────────────
    holdings = conn.execute("""
        SELECT
            a.account_name,
            h.symbol,
            h.name,
            h.price,
            h.quantity,
            h.cost_basis,
            h.current_balance,
            h.unrealized_gain_loss
        FROM holdings h
        JOIN accounts a ON h.account_id = a.account_id
        ORDER BY h.current_balance DESC
    """).fetchall()

    conn.close()

    # ── Serialize to prompt-friendly text ────────────────────────────────
    lines = []

    if summary:
        lines.append(f"PORTFOLIO SNAPSHOT — as of {summary['as_of_date']}")
        lines.append(f"Total net worth: ${summary['total_net_worth']:,.2f}")
        lines.append("")

    lines.append("=== ACCOUNTS ===")
    for a in accounts:
        lines.append(f"  {a['account_name']} | Balance: ${a['total_balance']:,.2f}")
    lines.append("")

    lines.append("=== HOLDINGS (sorted by current balance, largest first) ===")
    lines.append(f"{'Account':<42} {'Symbol':<10} {'Qty':>10} {'Balance':>14} {'Cost Basis':>14} {'Unreal G/L':>14}")
    lines.append("-" * 108)

    for h in holdings:
        lines.append(
            f"{h['account_name']:<42} "
            f"{h['symbol']:<10} "
            f"{h['quantity']:>10,.3f} "
            f"${h['current_balance']:>13,.2f} "
            f"${h['cost_basis']:>13,.2f} "
            f"${h['unrealized_gain_loss']:>13,.2f}"
        )

    return "\n".join(lines)


# ── Q&A engine ────────────────────────────────────────────────────────────────

def ask_portfolio(question: str, context: str, client: anthropic.Anthropic) -> str:
    """
    Sends a question + full portfolio context to Claude.
    Returns the answer as a plain string.
    """
    message = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": (
                    f"Here is the complete portfolio data:\n\n"
                    f"{context}\n\n"
                    f"---\n\n"
                    f"Question: {question}"
                ),
            }
        ],
    )
    return message.content[0].text


# ── Conversation loop ─────────────────────────────────────────────────────────

def run_qa_loop():
    """
    Interactive CLI loop. Type a question, get an answer.
    Type 'quit' or press Ctrl-C to exit.
    """
    print("\n" + "=" * 60)
    print("  Portfolio AI — powered by Claude")
    print("=" * 60)
    print("Loading portfolio data…")

    context = load_portfolio_context(DB_PATH)
    client  = anthropic.Anthropic()   # reads ANTHROPIC_API_KEY from env

    # Count holdings for the welcome message
    holding_count = context.count("\n") - context.count("===")
    print(f"Portfolio loaded. Ask anything about the data.\n")
    print("Example questions:")
    print("  • What is my total tech exposure?")
    print("  • Which account has the highest unrealized gains?")
    print("  • How much is in dividend ETFs vs growth ETFs?")
    print("  • What percentage of the portfolio is international?")
    print("  • Which single position is the largest holding?\n")
    print("Type 'quit' to exit.\n")
    print("-" * 60)

    conversation_history = []

    while True:
        try:
            question = input("\nYou: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.")
            break

        if not question:
            continue
        if question.lower() in ("quit", "exit", "q"):
            print("Goodbye.")
            break

        # Multi-turn: keep conversation history so follow-up questions work
        conversation_history.append({
            "role": "user",
            "content": (
                f"Portfolio data:\n\n{context}\n\n---\n\nQuestion: {question}"
                if not conversation_history          # inject context only on first turn
                else question
            ),
        })

        print("\nClaude: ", end="", flush=True)

        # Stream the response so it feels responsive
        with client.messages.stream(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=conversation_history,
        ) as stream:
            full_response = ""
            for text in stream.text_stream:
                print(text, end="", flush=True)
                full_response += text

        print()  # newline after streamed response

        conversation_history.append({
            "role": "assistant",
            "content": full_response,
        })


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    run_qa_loop()
    