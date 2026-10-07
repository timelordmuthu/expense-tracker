"""
logic.py  —  Pure financial calculations. No UI, no file I/O.

Glossary
--------
cash flow     net movement of money in / out over a period
liquidity     how much spendable cash is available right now
receivable    money owed TO you (an asset on your balance sheet)
payable       money you owe to others (a liability)
net balance   total inflows minus total outflows
daily rate    a periodic amount expressed per calendar day
projection    estimated future value based on current data
amortise      spread a cost over time (e.g. daily saving toward a goal)
"""

from datetime import date, datetime
import calendar

# ── 1. Balance ────────────────────────────────────────────────────────────────

def calculate_balance(data: dict) -> dict:
    """
    Net balance (liquidity) = income + bonuses(open) − expenses.
    Returns a breakdown dict so the UI can display each component.
    """
    def _sum(entries, key="amount"):
        total = 0.0
        for e in entries:
            try:
                total += float(e.get(key, 0) or 0)
            except (TypeError, ValueError):
                pass
        return round(total, 2)

    total_income   = _sum(data.get("income",   []))
    total_expenses = _sum(data.get("expenses", []))
    # Only open (non-hidden) bonuses count toward balance
    open_bonuses   = [b for b in data.get("bonuses", []) if not b.get("hidden", False)]
    total_bonuses  = _sum(open_bonuses)

    net = round(total_income + total_bonuses - total_expenses, 2)

    return {
        "total_income":   total_income,
        "total_expenses": total_expenses,
        "total_bonuses":  total_bonuses,
        "net_balance":    net,
    }

# ── 2. Pending (payables & receivables) ──────────────────────────────────────

def calculate_pending(data: dict) -> tuple:
    """
    Returns (to_pay, to_receive).
    to_pay     = pending expenses  (liability — you owe someone)
    to_receive = pending income    (receivable — someone owes you)
    """
    to_pay = to_receive = 0.0
    for e in data.get("expenses", []):
        if str(e.get("status", "")).lower() == "pending":
            try:
                amt = float(e.get("amount", 0) or 0)
                if amt > 0:
                    to_pay += amt
            except (TypeError, ValueError):
                pass
    for e in data.get("income", []):
        if str(e.get("status", "")).lower() == "pending":
            try:
                amt = float(e.get("amount", 0) or 0)
                if amt > 0:
                    to_receive += amt
            except (TypeError, ValueError):
                pass
    return round(to_pay, 2), round(to_receive, 2)

# ── 3. Reward estimation ──────────────────────────────────────────────────────

def estimate_days_for_reward(balance: float, daily_saving: float, reward_cost: float) -> dict:
    """
    Linear projection: days = shortfall / daily_saving_rate.
    Handles: already affordable, zero saving, negative saving.
    """
    balance      = float(balance)
    daily_saving = float(daily_saving)
    reward_cost  = float(reward_cost)

    if reward_cost <= 0:
        return {"days": 0, "affordable": True, "shortfall": 0.0, "reachable": True, "error": None}

    shortfall = max(0.0, round(reward_cost - balance, 2))

    if shortfall == 0.0:
        return {"days": 0, "affordable": True, "shortfall": 0.0, "reachable": True, "error": None}

    if daily_saving <= 0:
        reason = "Daily saving is zero." if daily_saving == 0 else "Balance is shrinking (negative saving)."
        return {"days": None, "affordable": False, "shortfall": shortfall, "reachable": False, "error": reason}

    import math
    days = math.ceil(shortfall / daily_saving)
    return {"days": days, "affordable": False, "shortfall": shortfall, "reachable": True, "error": None}

# ── 4. Monthly income projection ──────────────────────────────────────────────

_MULTIPLIERS = {
    "daily":    365 / 12,
    "weekly":   52  / 12,
    "biweekly": 26  / 12,
    "monthly":  1.0,
    "yearly":   1   / 12,
}

def monthly_income_projection(amount: float, mode: str) -> dict:
    """
    Normalise any periodic income to a monthly figure.
    daily mode accounts for actual days in a month via calendar.
    """
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return {"monthly": 0.0, "yearly": 0.0, "daily": 0.0, "mode": mode,
                "amount": 0.0, "error": "Invalid amount."}

    if amount < 0:
        return {"monthly": 0.0, "yearly": 0.0, "daily": 0.0, "mode": mode,
                "amount": amount, "error": "Amount cannot be negative."}

    key = str(mode).lower()
    if key not in _MULTIPLIERS:
        return {"monthly": 0.0, "yearly": 0.0, "daily": 0.0, "mode": mode,
                "amount": amount, "error": f"Unknown mode '{mode}'."}

    monthly = round(amount * _MULTIPLIERS[key], 2)
    yearly  = round(monthly * 12, 2)
    daily   = round(monthly / (365 / 12), 2)

    return {"monthly": monthly, "yearly": yearly, "daily": daily,
            "mode": key, "amount": amount, "error": None}

# ── 5. Daily saving rate (derived from last N days of data) ───────────────────

def compute_daily_saving(data: dict, window_days: int = 30) -> float:
    """
    Estimate daily saving rate from the last *window_days* of entries.
    saving = (income + open_bonuses - expenses) / window_days
    Returns 0.0 if negative (can't save toward a goal at current rate).
    """
    bi = calculate_balance(data)
    rate = (bi["total_income"] + bi["total_bonuses"] - bi["total_expenses"]) / max(window_days, 1)
    return max(0.0, round(rate, 2))

# ── 6. Actual monthly income (day-count aware) ────────────────────────────────

def actual_monthly_income(amount: float, mode: str, month: int = None, year: int = None) -> float:
    """
    For daily-rate income: multiply by actual days in the given month.
    For all other modes: delegate to monthly_income_projection.

    This is the key distinction — a daily wage earner gets MORE in a 31-day month.
    """
    if str(mode).lower() != "daily":
        return monthly_income_projection(amount, mode)["monthly"]

    today = date.today()
    m = month or today.month
    y = year  or today.year
    days_in_month = calendar.monthrange(y, m)[1]
    return round(float(amount) * days_in_month, 2)
