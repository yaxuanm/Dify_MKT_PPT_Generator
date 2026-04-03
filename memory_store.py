"""
Persistent memory for the PPT agent: usage events, aggregates, optional LLM-refreshed insights.
"""
import os
import sqlite3
from datetime import datetime, timezone
from threading import Lock

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
DB_PATH = os.path.join(DATA_DIR, "memory.db")
_lock = Lock()
_EVENT_COUNT_SINCE_REFRESH = 0

# LLM insight refresh: off by default (saves API $). Set MEMORY_LLM_REFRESH=1 to enable.
LLM_REFRESH_ENABLED = os.environ.get("MEMORY_LLM_REFRESH", "0").lower() in ("1", "true", "yes")
# When enabled: refresh every N recorded events (higher = fewer Haiku calls)
REFRESH_EVERY_N = int(os.environ.get("MEMORY_REFRESH_EVERY", "80"))
# Cap memory block size in system prompt (tokens ≈ chars/4)
MEMORY_MAX_CHARS = int(os.environ.get("MEMORY_MAX_CHARS", "380"))
# Aggregate last N rows only (smaller DB read)
AGGREGATE_LIMIT = int(os.environ.get("MEMORY_AGGREGATE_LIMIT", "120"))


def _conn():
    os.makedirs(DATA_DIR, exist_ok=True)
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    with _lock:
        with _conn() as c:
            c.execute(
                """CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT NOT NULL,
                session_id TEXT,
                endpoint TEXT,
                has_image INTEGER DEFAULT 0,
                message_preview TEXT,
                layout_type TEXT,
                success INTEGER NOT NULL,
                error_type TEXT,
                error_detail TEXT,
                output_file TEXT
            )"""
            )
            c.execute(
                """CREATE TABLE IF NOT EXISTS insights (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                updated_ts TEXT,
                text TEXT NOT NULL DEFAULT ''
            )"""
            )
            c.execute(
                """CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT NOT NULL,
                session_id TEXT,
                rating INTEGER,
                note TEXT
            )"""
            )


def record_event(
    *,
    session_id,
    endpoint,
    has_image=False,
    message_preview="",
    layout_type=None,
    success=True,
    error_type=None,
    error_detail=None,
    output_file=None,
):
    global _EVENT_COUNT_SINCE_REFRESH
    preview = (message_preview or "")[:500]
    with _lock:
        with _conn() as c:
            c.execute(
                """INSERT INTO events (ts, session_id, endpoint, has_image, message_preview,
                layout_type, success, error_type, error_detail, output_file)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (
                    datetime.now(timezone.utc).isoformat(),
                    session_id or "",
                    endpoint,
                    1 if has_image else 0,
                    preview,
                    layout_type,
                    1 if success else 0,
                    error_type,
                    (error_detail or "")[:2000] if error_detail else None,
                    output_file,
                ),
            )
        _EVENT_COUNT_SINCE_REFRESH += 1


def record_feedback(session_id, rating, note=""):
    with _lock:
        with _conn() as c:
            c.execute(
                "INSERT INTO feedback (ts, session_id, rating, note) VALUES (?,?,?,?)",
                (
                    datetime.now(timezone.utc).isoformat(),
                    session_id or "",
                    rating,
                    (note or "")[:2000],
                ),
            )


def _aggregate_recent(limit_events=None):
    lim = limit_events if limit_events is not None else AGGREGATE_LIMIT
    with _lock:
        with _conn() as c:
            rows = c.execute(
                """SELECT success, layout_type, error_type, endpoint
                FROM events ORDER BY id DESC LIMIT ?""",
                (lim,),
            ).fetchall()
    total = len(rows)
    if total == 0:
        return None
    ok = sum(1 for r in rows if r[0])
    layouts = {}
    errors = {}
    for suc, layout, err, _ep in rows:
        if suc and layout:
            layouts[layout] = layouts.get(layout, 0) + 1
        if not suc and err:
            errors[err] = errors.get(err, 0) + 1
    top_layouts = sorted(layouts.items(), key=lambda x: -x[1])[:4]
    top_errors = sorted(errors.items(), key=lambda x: -x[1])[:3]
    return {
        "total_recent": total,
        "success_count": ok,
        "fail_count": total - ok,
        "success_rate": round(100.0 * ok / total, 1) if total else 0,
        "top_layouts": top_layouts,
        "top_errors": top_errors,
    }


def get_saved_insights_text():
    with _lock:
        with _conn() as c:
            row = c.execute("SELECT text FROM insights WHERE id=1").fetchone()
    if row and row[0]:
        return row[0].strip()
    return ""


def save_insights_text(text: str):
    ts = datetime.now(timezone.utc).isoformat()
    # Stored short; prompt only uses truncated slice
    cap = int(os.environ.get("MEMORY_INSIGHTS_STORE_MAX", "1200"))
    with _lock:
        with _conn() as c:
            c.execute(
                "INSERT OR REPLACE INTO insights (id, updated_ts, text) VALUES (1, ?, ?)",
                (ts, text[:cap]),
            )


def _truncate(s: str, max_chars: int) -> str:
    s = (s or "").strip()
    if len(s) <= max_chars:
        return s
    return s[: max_chars - 1] + "…"


def format_memory_block():
    """Compact stats for system prompt (minimal tokens). No extra API call."""
    agg = _aggregate_recent()
    lines = []
    if not agg:
        lines.append("[mem] no history yet → output raw JSON only, pick layout from intent.")
    else:
        # One dense line: rate + layouts
        lr = ",".join(f"{k}:{v}" for k, v in agg["top_layouts"]) or "—"
        lines.append(
            f"[mem] ok {agg['success_rate']}% n={agg['total_recent']} layouts {lr}"
        )
        if agg["top_errors"]:
            er = ",".join(f"{k[:24]}:{v}" for k, v in agg["top_errors"])
            lines.append(f"[mem] errs {er}")
        lines.append("[mem] on parse fails: never wrap JSON in ```")
    insight = get_saved_insights_text()
    if insight:
        lines.append("[mem] hints " + _truncate(insight, 200))
    out = "\n".join(lines)
    return _truncate(out, MEMORY_MAX_CHARS)


def maybe_refresh_insights_llm(client, model: str = None):
    """
    Optional Haiku call: set MEMORY_LLM_REFRESH=1. Default off to save tokens/cost.
    """
    global _EVENT_COUNT_SINCE_REFRESH
    if not LLM_REFRESH_ENABLED:
        return
    if REFRESH_EVERY_N <= 0:
        return
    if _EVENT_COUNT_SINCE_REFRESH < REFRESH_EVERY_N:
        return
    _EVENT_COUNT_SINCE_REFRESH = 0

    agg = _aggregate_recent()
    if not agg or agg["total_recent"] < 8:
        return

    recent_errors = []
    with _lock:
        with _conn() as c:
            recent_errors = c.execute(
                """SELECT error_type, error_detail FROM events
                WHERE success=0 ORDER BY id DESC LIMIT 8"""
            ).fetchall()

    err_blob = "; ".join(
        f"{a}:{(b or '')[:80]}" for a, b in recent_errors if a or b
    )[:900]

    user_prompt = f"""Stats: ok {agg['success_rate']}% n={agg['total_recent']} layouts {agg['top_layouts']} errs {agg['top_errors']}
Errors: {err_blob}
Reply with 3-5 comma-separated short rules (English, no sentences over 80 chars). Total under 450 characters."""

    primary = model or os.environ.get("MEMORY_HAIKU_MODEL", "claude-3-5-haiku-20241022")
    raw_fb = os.environ.get("MEMORY_HAIKU_FALLBACK", "claude-3-haiku-20240307").strip()
    fallbacks = [x.strip() for x in raw_fb.split(",") if x.strip()]
    seen = set()
    chain = []
    for m in [primary] + fallbacks:
        if m and m not in seen:
            seen.add(m)
            chain.append(m)
    text = ""
    for m in chain:
        try:
            r = client.messages.create(
                model=m,
                max_tokens=200,
                system="Output only the rules text, no quotes.",
                messages=[{"role": "user", "content": user_prompt}],
            )
            text = r.content[0].text.strip()
            if text:
                save_insights_text(text)
            break
        except Exception:
            continue


def memory_summary_json():
    """For GET /api/memory/summary debugging."""
    agg = _aggregate_recent()
    with _lock:
        with _conn() as c:
            n = c.execute("SELECT COUNT(*) FROM events").fetchone()[0]
            fb = c.execute("SELECT COUNT(*) FROM feedback").fetchone()[0]
    return {
        "events_total": n,
        "feedback_total": fb,
        "aggregate_recent": agg,
        "insights_preview": (get_saved_insights_text() or "")[:500],
    }
