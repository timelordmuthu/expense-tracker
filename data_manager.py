"""
data_manager.py  —  JSON flat-file storage layer.
No Streamlit, no UI logic.
"""

import json
import os
from datetime import date, datetime

DATA_DIR  = "data"
DATA_FILE = os.path.join(DATA_DIR, "expenses.json")

DEFAULT_CATEGORIES = [
    "Food & Dining", "Transport", "Utilities", "Healthcare",
    "Entertainment", "Shopping", "Education", "Other",
]

# ── helpers ───────────────────────────────────────────────────────────────────

def _ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)

def _default_data() -> dict:
    return {
        "income":         [],
        "expenses":       [],
        "bonuses":        [],
        "rewards":        [],
        "tasks":          [],
        "categories":     list(DEFAULT_CATEGORIES),
        "streak":         0,
        "last_entry":     None,
        "points":         0,
        "income_sources": [],
    }

def _parse_date(value):
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return None

# ── public API ────────────────────────────────────────────────────────────────

def init_data() -> dict:
    _ensure_data_dir()
    if not os.path.exists(DATA_FILE):
        fresh = _default_data()
        save_data(fresh)
        return fresh
    return load_data()

def load_data() -> dict:
    _ensure_data_dir()
    if not os.path.exists(DATA_FILE):
        return _default_data()
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except (json.JSONDecodeError, OSError):
        return _default_data()
    for key, val in _default_data().items():
        data.setdefault(key, val)
    return data

def save_data(data: dict) -> None:
    _ensure_data_dir()
    with open(DATA_FILE, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=4, ensure_ascii=False, default=str)

def add_entry(section: str, entry) -> dict:
    """Append entry to section list, advance streak, award +10 pts."""
    data = load_data()
    if section not in data or not isinstance(data[section], list):
        raise KeyError(f"Unknown list section: '{section}'")
    if isinstance(entry, dict):
        entry.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
    data[section].append(entry)
    if section in ("income", "expenses", "bonuses"):
        data = _advance_streak(data)
        data["points"] = data.get("points", 0) + 10
    save_data(data)
    return data

def mark_paid(section: str, index: int) -> dict:
    """Flip entry status to paid/received and persist."""
    data = load_data()
    entries = data.get(section, [])
    if 0 <= index < len(entries):
        entries[index]["status"] = "received" if section == "income" else "paid"
        data[section] = entries
        save_data(data)
    return data

def update_streak(data: dict = None) -> dict:
    if data is None:
        data = load_data()
    data = _advance_streak(data)
    save_data(data)
    return data

def delete_entry(section: str, index: int) -> dict:
    """Remove the entry at *index* from *section* list and persist."""
    data = load_data()
    entries = data.get(section, [])
    if 0 <= index < len(entries):
        entries.pop(index)
        data[section] = entries
        save_data(data)
    return data

def award_points(amount: int) -> dict:
    data = load_data()
    data["points"] = data.get("points", 0) + amount
    save_data(data)
    return data

def _advance_streak(data: dict) -> dict:
    today      = date.today()
    last_entry = _parse_date(data.get("last_entry"))
    if last_entry is None:
        data["streak"]     = 1
        data["last_entry"] = today.isoformat()
    elif last_entry == today:
        pass
    elif (today - last_entry).days == 1:
        data["streak"]    += 1
        data["last_entry"] = today.isoformat()
        if data["streak"] % 7 == 0:
            data["points"] = data.get("points", 0) + data["streak"] * 5
    else:
        data["streak"]     = 1
        data["last_entry"] = today.isoformat()
    return data
