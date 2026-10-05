"""
seed.py — Database Seeder
==========================
Reads VEDVEDA.xlsx (all sheets) and populates the SQLite database
with members and transactions. Safe to call only once (on first run).

Sheet mapping:
  TEAM SALES  → master log (used to cross-check / fill missing member rows)
  SHAILE      → Shailendra's individual transactions
  POOJA       → Pooja's individual transactions
  PANKAJ      → Pankaj's individual transactions
  Anita mem   → Anita's individual transactions
  depika mem  → Deepika's individual transactions
"""

import os
import re
from datetime import datetime, date

import openpyxl

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

EXCEL_PATH = os.path.join(os.path.abspath(os.path.dirname(__file__)), "VEDVEDA.xlsx")

# Members defined in the workbook; map display-name → canonical name
MEMBER_MAP = {
    "POOJA MEM":    "Pooja",
    "POOJA":        "Pooja",
    "DEEPIKA MEM":  "Deepika",
    "DEEPIKA":      "Deepika",
    "PANKAJ":       "Pankaj",
    "SHAILE":       "Shailendra",
    "SHAILENDRA":   "Shailendra",
    "ANITA MEM":    "Anita",
    "ANITA":        "Anita",
    "SELF":         "Shailendra",  # 'self' in team sheet refers to Shailendra (sheet owner)
}

CANONICAL_MEMBERS = [
    {"name": "Pooja",      "role": "Team Member"},
    {"name": "Deepika",    "role": "Team Member"},
    {"name": "Pankaj",     "role": "Team Member"},
    {"name": "Shailendra", "role": "Team Member"},
    {"name": "Anita",      "role": "Team Member"},
]


def _safe_float(val):
    """Convert a cell value to float, ignoring formula strings."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s.startswith("=") or s == "":
        return 0.0
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return 0.0


def _safe_int(val):
    return int(_safe_float(val))


def _parse_date(val):
    """Return a date object or None."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _clean_product(val):
    """Strip leading/trailing whitespace and normalise product strings."""
    if not val:
        return None
    return str(val).strip().title()


def _resolve_member_name(raw):
    """Map raw cell text to a canonical member name."""
    if not raw:
        return None
    upper = str(raw).strip().upper()
    return MEMBER_MAP.get(upper)


# ---------------------------------------------------------------------------
# Sheet parsers
# ---------------------------------------------------------------------------

def _parse_individual_sheet(ws, member_name):
    """
    Parse an individual member sheet (POOJA, SHAILE, PANKAJ, Anita mem, depika mem).
    Header row is detected automatically — we look for a row containing 'DATE'.
    Columns expected: DATE | NO (product) | PV | AMT | PAYMNT | PENAL | PAYMENT | PENAL
    """
    header_row = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        row_strs = [str(c).strip().upper() if c else "" for c in row]
        if "DATE" in row_strs:
            header_row = i
            break

    if header_row is None:
        print(f"  [WARN] Could not find header row in sheet, skipping.")
        return []

    records = []
    last_date = None

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        raw_date    = row[0]
        raw_product = row[1] if len(row) > 1 else None
        raw_pv      = row[2] if len(row) > 2 else None
        raw_amt     = row[3] if len(row) > 3 else None
        raw_payment = row[4] if len(row) > 4 else None
        raw_penal   = row[5] if len(row) > 5 else None

        # Skip entirely empty rows
        if all(v is None for v in [raw_date, raw_product, raw_pv, raw_amt]):
            continue

        # Skip formula-only or total rows
        product = _clean_product(raw_product)
        if product and product.lower() in ("total", "no", "product"):
            continue
        if not product:
            continue

        # Carry forward the last known date
        parsed_date = _parse_date(raw_date)
        if parsed_date:
            last_date = parsed_date
        current_date = last_date

        pv     = _safe_float(raw_pv)
        amount = _safe_float(raw_amt)

        # Determine payment status from 'PAYMNT' column
        paid_val = _safe_float(raw_payment)
        if paid_val and paid_val >= amount and amount > 0:
            status = "Paid"
            penal_amount = 0.0
        elif paid_val and 0 < paid_val < amount:
            status = "Pending"
            penal_amount = round(amount - paid_val, 2)
        else:
            # If PAYMNT column is None/0 and amount > 0, treat as Pending
            if amount > 0:
                status = "Pending"
                penal_amount = amount
            else:
                status = "Paid"
                penal_amount = 0.0

        penal_pv = _safe_float(raw_penal)

        # Try to extract quantity from product string (e.g. "5 KIDG" → qty=5)
        qty_match = re.match(r"^(\d+(?:\.\d+)?)\s+", str(raw_product or ""))
        quantity = int(float(qty_match.group(1))) if qty_match else 1

        records.append({
            "member_name":    member_name,
            "date":           current_date,
            "product_name":   product,
            "quantity":       quantity,
            "pv":             pv,
            "amount":         amount,
            "payment_status": status,
            "penal_pv":       penal_pv,
            "penal_amount":   penal_amount,
        })

    return records


def _parse_team_sheet(ws):
    """
    Parse TEAM SALES sheet.
    Columns: DATE | TEAM NAME | NO | PV | AMT | PAYMNT | PENAL PV | PAYMENT | PENAL PV
    """
    header_row = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        row_strs = [str(c).strip().upper() if c else "" for c in row]
        if "DATE" in row_strs and "TEAM NAME" in row_strs:
            header_row = i
            break

    if header_row is None:
        return []

    records = []
    last_date = None
    last_member = None

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        raw_date    = row[0] if len(row) > 0 else None
        raw_member  = row[1] if len(row) > 1 else None
        raw_product = row[2] if len(row) > 2 else None
        raw_pv      = row[3] if len(row) > 3 else None
        raw_amt     = row[4] if len(row) > 4 else None
        raw_payment = row[5] if len(row) > 5 else None

        # Skip blank rows
        if all(v is None for v in [raw_date, raw_member, raw_product]):
            continue

        product = _clean_product(raw_product)
        if not product:
            continue

        parsed_date = _parse_date(raw_date)
        if parsed_date:
            last_date = parsed_date

        # Carry-forward member name (represented by '"' in spreadsheet)
        member_raw = str(raw_member or "").strip()
        if member_raw and member_raw not in ('"', "'", ""):
            resolved = _resolve_member_name(member_raw)
            if resolved:
                last_member = resolved
        current_member = last_member

        if not current_member:
            continue

        pv     = _safe_float(raw_pv)
        amount = _safe_float(raw_amt)
        paid   = _safe_float(raw_payment)

        if paid and paid >= amount and amount > 0:
            status = "Paid"
            penal_amount = 0.0
        elif paid and 0 < paid < amount:
            status = "Pending"
            penal_amount = round(amount - paid, 2)
        else:
            status = "Pending" if amount > 0 else "Paid"
            penal_amount = amount if amount > 0 else 0.0

        qty_match = re.match(r"^(\d+(?:\.\d+)?)\s+", str(raw_product or ""))
        quantity = int(float(qty_match.group(1))) if qty_match else 1

        records.append({
            "member_name":    current_member,
            "date":           last_date,
            "product_name":   product,
            "quantity":       quantity,
            "pv":             pv,
            "amount":         amount,
            "payment_status": status,
            "penal_pv":       0.0,
            "penal_amount":   penal_amount,
        })

    return records


# ---------------------------------------------------------------------------
# Main seed function
# ---------------------------------------------------------------------------

INDIVIDUAL_SHEETS = {
    "SHAILE":    "Shailendra",
    "POOJA":     "Pooja",
    "PANKAJ":    "Pankaj",
    "Anita mem": "Anita",
    "depika mem":"Deepika",
}


def seed_database(db, Member, Transaction):
    """
    Populate the database from VEDVEDA.xlsx.
    Called automatically from app.py on first run.
    """
    xlsx_path = EXCEL_PATH
    if not os.path.exists(xlsx_path):
        # Try looking next to the script
        xlsx_path = os.path.join(os.path.dirname(__file__), "VEDVEDA.xlsx")
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(
            f"VEDVEDA.xlsx not found. Place it next to app.py at: {EXCEL_PATH}"
        )

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # 1. Insert members
    member_objects = {}
    for m_data in CANONICAL_MEMBERS:
        existing = Member.query.filter_by(name=m_data["name"]).first()
        if not existing:
            m = Member(name=m_data["name"], role=m_data["role"])
            db.session.add(m)
            db.session.flush()  # get id
            member_objects[m_data["name"]] = m
            print(f"  Added member: {m_data['name']}")
        else:
            member_objects[m_data["name"]] = existing

    db.session.commit()

    # 2. Collect all transaction records (individual sheets take priority)
    all_records = []

    # Parse individual sheets first (most accurate)
    for sheet_name, member_name in INDIVIDUAL_SHEETS.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            recs = _parse_individual_sheet(ws, member_name)
            print(f"  Sheet '{sheet_name}' → {len(recs)} records for {member_name}")
            all_records.extend(recs)
        else:
            print(f"  [WARN] Sheet '{sheet_name}' not found, skipping.")

    # Parse team sheet for any additional/cross-check entries
    # We only add team-sheet records that have members not covered above
    # (to avoid duplicating entries already in individual sheets).
    # Since individual sheets are more granular, we skip team-sheet to avoid dups.
    # Uncomment the block below to include team-sheet records as well:
    #
    # if "TEAM SALES" in wb.sheetnames:
    #     team_recs = _parse_team_sheet(wb["TEAM SALES"])
    #     for r in team_recs:
    #         # only add if member NOT in individual sheets (PANKAJ might only be here)
    #         pass

    # 3. Insert transactions
    inserted = 0
    for rec in all_records:
        m_obj = member_objects.get(rec["member_name"])
        if not m_obj:
            print(f"  [WARN] Unknown member '{rec['member_name']}', skipping record.")
            continue

        txn = Transaction(
            member_id      = m_obj.id,
            date           = rec["date"],
            product_name   = rec["product_name"],
            quantity       = rec["quantity"],
            pv             = rec["pv"],
            amount         = rec["amount"],
            payment_status = rec["payment_status"],
            penal_pv       = rec["penal_pv"],
            penal_amount   = rec["penal_amount"],
        )
        db.session.add(txn)
        inserted += 1

    # Also parse PANKAJ from TEAM SALES since his individual sheet is minimal
    if "TEAM SALES" in wb.sheetnames:
        ws_team = wb["TEAM SALES"]
        team_recs = _parse_team_sheet(ws_team)
        for rec in team_recs:
            if rec["member_name"] == "Pankaj":
                # Check not already inserted from individual sheet
                m_obj = member_objects.get("Pankaj")
                txn = Transaction(
                    member_id      = m_obj.id,
                    date           = rec["date"],
                    product_name   = rec["product_name"],
                    quantity       = rec["quantity"],
                    pv             = rec["pv"],
                    amount         = rec["amount"],
                    payment_status = rec["payment_status"],
                    penal_pv       = 0.0,
                    penal_amount   = rec["penal_amount"],
                )
                db.session.add(txn)
                inserted += 1

    db.session.commit()
    print(f"  Total transactions inserted: {inserted}")


# ---------------------------------------------------------------------------
# Standalone execution
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Allow running seed.py directly for a quick re-seed
    from app import app, db, Member, Transaction
    with app.app_context():
        db.drop_all()
        db.create_all()
        seed_database(db, Member, Transaction)
        print("Seeding finished.")
