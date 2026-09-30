"""
financial_analysis.py — Bank CSV  →  professional Excel financial analysis.

HOW TO RUN (VS Code terminal):
    python financial_analysis.py                     # uses INPUT_FILE below
    python financial_analysis.py my_statement.csv    # analyse a different CSV

WHAT IT DOES
    Reads a bank statement CSV, classifies EVERY transaction, groups money
    received / paid by person or company, calculates GST and an income tax
    estimate, and writes one Excel workbook with many sheets.

GOLDEN RULES FOLLOWED BY THIS PROGRAM
    * No transaction is ever deleted or ignored — every CSV row appears in
      the "All Transactions" sheet and in either Income or Expense detail.
    * The original bank description is kept exactly as it appears.
    * When there is not enough evidence, the transaction is NOT treated as
      confirmed business income or a confirmed deductible expense — it goes
      to "Review Required" with the reason.
    * Totals are validated against the original CSV (difference must be $0.00).

ALL FIGURES ARE ESTIMATES ONLY – ACCOUNTANT REVIEW REQUIRED.
"""

import csv
import re
import sys
from collections import defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# 1. SETTINGS — change these to suit your business
# =============================================================================
INPUT_FILE = "bank_transactions.csv"
OUTPUT_FILE = "Business_Financial_Analysis.xlsx"

# Income tax estimate settings.
#   "company"    -> profit × COMPANY_TAX_RATE
#   "individual" -> Australian resident progressive rates (INDIVIDUAL_TAX_BRACKETS)
#   "custom"     -> profit × CUSTOM_TAX_RATE
ENTITY_TYPE = "company"
# 25% is the "base rate entity" company rate (aggregated turnover under $50m and
# no more than 80% passive income). Other companies pay 30%. Confirm with your
# accountant which rate applies to your company.
COMPANY_TAX_RATE = 0.25
CUSTOM_TAX_RATE = 0.30
# Resident individual rates for 2024-25: (income threshold, rate above threshold).
INDIVIDUAL_TAX_BRACKETS = [(0, 0.0), (18_200, 0.16), (45_000, 0.30), (135_000, 0.37), (190_000, 0.45)]

# Dates in Australian bank CSVs are day/month/year.
DATE_DAY_FIRST = True

# Directors / personal names — never treated automatically as business income or expense.
DIRECTOR_NAMES = ["Ganzorig Batzorig", "Munkhchimeg Lkhagvatseren"]

# People who get their own separate section (money received FROM / paid TO).
SPECIFIC_PEOPLE = ["Munkhbat Ulziikhutag", "Davaadulam Dashdavaa", "Davaadolgor Dashdavaa", "Deleg"]

# Known customers (confirmed business income when they pay you).
KNOWN_CUSTOMERS = ["Richard Crookes"]

# Names that are definitely the same person/company (strong evidence).
# Standard name -> other spellings seen in the bank statement.
NAME_ALIASES = {
    "Richard Crookes": ["RICHARD CROOKES CONSTRUCTIONS", "RICHARD CROOKES PTY LTD",
                        "RICHARD CROOKES CONSTRUCTION", "RCC"],
}

# GST status of customers. Only add a name after checking your invoices.
#   GST_INCLUSIVE_PAYERS: invoices to them INCLUDED GST  -> GST = amount / 11
#   NO_GST_PAYERS:        invoices to them had NO GST     -> GST = $0
GST_INCLUSIVE_PAYERS = []          # e.g. ["Richard Crookes"]
NO_GST_PAYERS = []

# Words that indicate wages when found in an OUTGOING payment to an individual.
SALARY_WORDS = ["payment", "wage", "wages", "salary", "payroll", "tsalin", "tsalind"]
# Words that indicate borrowing / lending.
LOAN_WORDS = ["loan", "lend", "lent", "borrow", "borrowed"]
# Words that suggest money being returned.
RETURN_WORDS = ["back", "repay", "repayment", "return", "refund"]
# Words in an INCOMING reference that suggest business work.
BUSINESS_WORK_WORDS = ["clean", "cleaning", "construction", "const", "invoice", "inv", "job", "abn"]
# Repeated payments to the same individual (without a wage keyword) -> "possible salary".
REPEATED_PAYMENT_MIN = 3

# =============================================================================
# 2. FIXED LABELS (used everywhere — do not change spelling)
# =============================================================================
BUS_INC = "BUSINESS INCOME"
BUS_EXP = "BUSINESS EXPENSE"
PER_INC = "PERSONAL INCOME"
PER_EXP = "PERSONAL EXPENSE"
DIRECTOR = "PERSONAL / DIRECTOR"
SALARY = "SALARY / WAGES"
LOAN = "LOAN"
TRANSFER = "TRANSFER"
TAX = "TAX"
GST = "GST"
REVIEW = "REVIEW REQUIRED"

GST_INCLUDED = "GST INCLUDED"
NO_GST = "NO GST"
GST_UNKNOWN = "GST UNKNOWN"
GST_NA = "NOT APPLICABLE"

INCOME = "INCOME"
EXPENSE = "EXPENSE"

# Review reason labels (each reason starts with one of these).
R_DIRECTOR = "POSSIBLE PERSONAL / DIRECTOR"
R_LOAN = "REVIEW REQUIRED – POSSIBLE LOAN"
R_UNKNOWN = "UNKNOWN PERSON/COMPANY"
R_GST = "UNKNOWN GST STATUS"
R_SALARY = "POSSIBLE SALARY"
R_TRANSFER = "POSSIBLE INTERNAL TRANSFER"
R_EXPENSE = "UNCLEAR BUSINESS EXPENSE"
R_REFUND = "REFUND / REVERSAL"
R_DUPLICATE = "POSSIBLE DUPLICATE"
R_SIMILAR = "POSSIBLE SAME NAME"
R_UNREADABLE = "UNREADABLE DATA"
REVIEW_REASON_LABELS = [R_DIRECTOR, R_LOAN, R_UNKNOWN, R_GST, R_SALARY, R_TRANSFER,
                        R_EXPENSE, R_REFUND, R_DUPLICATE, R_SIMILAR, R_UNREADABLE]

# Words that mean a name is an organisation, not an individual.
ORGANISATION_WORDS = {
    "pty", "ltd", "limited", "inc", "corp", "co", "company", "group", "trading", "services",
    "service", "dental", "hire", "construction", "constructions", "cleaning", "australia",
    "bank", "insurance", "pathology", "clinic", "school", "club", "council", "holdings",
    "enterprises", "solutions", "pl", "lt", "store", "stores", "shop", "centre", "center",
    "hospital", "pharmacy", "refunds", "refund", "fund", "trust", "super",
}

# =============================================================================
# 3. MERCHANT RULES
#    (regex, standard name or None, level, category, GST status, note)
#    level:  "business"  = confirmed business expense (deductible)
#            "possible"  = probably business, but business use must be confirmed -> review
#            "personal"  = personal spending (not in business profit)
#            "fine"      = fines/penalties (not deductible) -> review
#            "transfer"  = looks like moving money to another account -> review
#    First matching rule wins, so specific rules come before general ones.
# =============================================================================
MERCHANT_RULES = [
    # --- confirmed business ---
    (r"CLEANMAX", "Cleanmax Trading", "business", "Cleaning Materials", GST_INCLUDED, "Cleaning supplier"),
    (r"CLEANERS ROOM", "Cleaners Room Australia", "business", "Cleaning Materials", GST_INCLUDED, "Cleaning supplier"),
    (r"TAXCARE", "Taxcare Pty Ltd", "business", "Accounting", GST_INCLUDED, "Accounting / BAS service"),
    (r"ESTIMATEONE", "EstimateOne", "business", "Software / Subscriptions", GST_INCLUDED, "Construction tender platform"),
    (r"A ?& ?F TOOLS", "A & F Tools", "business", "Tools & Equipment", GST_INCLUDED, "Tools supplier"),
    (r"KENNARDS|COATES HIRE|EQUIPMENT HIRE", None, "business", "Equipment Hire", GST_INCLUDED, "Equipment hire"),
    (r"\bICARE\b|WORKERS COMP", None, "business", "Workers Compensation", GST_INCLUDED, "Workers compensation insurance"),
    (r"PUBLIC LIABILITY", None, "business", "Public Liability Insurance", GST_INCLUDED, "Public liability insurance"),
    (r"\bXERO\b|\bMYOB\b|QUICKBOOKS", None, "business", "Software / Subscriptions", GST_INCLUDED, "Accounting software"),
    (r"SYDNEY TOOLS|TOOLS WAREHOUSE|TOTAL TOOLS", None, "business", "Tools & Equipment", GST_INCLUDED, "Tools supplier"),
    # --- possibly business (review) ---
    (r"BUNNINGS", "Bunnings", "possible", "Tools & Equipment", GST_INCLUDED,
     "Hardware store – could be business tools/materials or personal"),
    (r"CLEARCHOICE PRODUCTS", "ClearChoice Products", "possible", "Cleaning Materials", GST_INCLUDED,
     "Possible cleaning supplier – confirm"),
    (r"OFFICEWORKS", "Officeworks", "possible", "Office Expenses", GST_INCLUDED, "Office supplies – confirm business use"),
    (r"EG GROUP|METRO PETROLEUM|COLES EXPRESS|\bSHELL\b(?! COVE)|CALTEX|AMPOL|\bBP\b|7-ELEVEN|UNITED PETROLEUM",
     None, "possible", "Fuel", GST_INCLUDED, "Fuel – confirm business vehicle use"),
    (r"ETOLL|E-TOLL|LINKT", "Tolls (E-Toll / Linkt)", "possible", "Tolls", GST_INCLUDED, "Tolls – confirm business travel"),
    (r"PARKING|\bEPARK\b|WILSON PARK|SECURE PARK", None, "possible", "Parking", GST_INCLUDED,
     "Parking – confirm business travel"),
    (r"VODAFONE", "Vodafone", "possible", "Phone", GST_INCLUDED, "Phone – confirm business-use percentage"),
    (r"TELSTRA|OPTUS|BELONG|AMAYSIM|BOOST MOBILE", None, "possible", "Phone", GST_INCLUDED,
     "Phone – confirm business-use percentage"),
    (r"\bTPG\b|AUSSIE BROADBAND|IINET|SUPERLOOP", None, "possible", "Internet", GST_INCLUDED,
     "Internet – confirm business-use percentage"),
    (r"^DIRECT DEBIT \d* ?IAL|\bIAL-", "Insurance Australia Ltd (IAL)", "possible", "Insurance", GST_INCLUDED,
     "Insurance direct debit – policy type unknown (vehicle / public liability / personal?)"),
    (r"ALLIANZ|AAMI|NRMA|\bQBE\b|\bGIO\b|SUNCORP|\bYOUI\b|BUDGET DIRECT", None, "possible", "Insurance",
     GST_INCLUDED, "Insurance – policy type unknown"),
    (r"AUTO REPAIR|MECHANIC|TYRE|SMASH REPAIR|AUTO SERVICE", None, "possible", "Vehicle Maintenance", GST_INCLUDED,
     "Vehicle repair – confirm business vehicle"),
    (r"\bAUDI\b|MERCEDES|TOYOTA|HYUNDAI|MAZDA|\bKIA\b", None, "possible", "Vehicle Maintenance", GST_INCLUDED,
     "Car dealer – service, repair or purchase? Confirm business vehicle"),
    (r"SERVICE NSW", "Service NSW", "possible", "Vehicle Expenses", GST_UNKNOWN,
     "Service NSW – vehicle registration or licence? (rego has no GST)"),
    (r"SPEEDWAY|\bLIBERTY\b(?= .*PETROL)", None, "possible", "Fuel", GST_INCLUDED, "Fuel – confirm business vehicle use"),
    (r"\bTAXI|\bUBER\b(?! ?\*? ?EATS)|\bDIDI\b", None, "possible", "Other Business Expenses", GST_INCLUDED,
     "Taxi / rideshare – confirm business travel"),
    (r"JB HI ?FI|BING LEE|HARVEY NORMAN|THE GOOD GUYS", None, "possible", "Office Expenses", GST_INCLUDED,
     "Electronics store – business equipment or personal?"),
    (r"JUMPYSIM|JUMP SIM", None, "possible", "Phone", GST_INCLUDED, "Phone – confirm business-use percentage"),
    (r"TRAINING|CAREERS", None, "possible", "Other Business Expenses", GST_INCLUDED,
     "Training course – confirm it relates to the business"),
    (r"APPLE\.COM/BILL", "Apple.com/bill", "possible", "Software / Subscriptions", GST_INCLUDED,
     "Apple subscription – could be personal (iCloud, apps)"),
    (r"GOOGLE|MICROSOFT|ADOBE|CANVA|DROPBOX", None, "possible", "Software / Subscriptions", GST_INCLUDED,
     "Software subscription – confirm business use"),
    (r"\bRENT\b|REAL ESTATE", None, "possible", "Rent", GST_UNKNOWN, "Rent – business premises or home?"),
    (r"TRANSPORTFORNSW TAP|\bOPAL\b", "Transport for NSW (Opal)", "possible", "Other Business Expenses", GST_INCLUDED,
     "Public transport – confirm business travel"),
    (r"\bEBAY\b|AMAZON", None, "possible", "Other Business Expenses", GST_INCLUDED, "Online purchase – item unknown"),
    (r"ALTOGETHER GROUP", "Altogether Group", "possible", "Other Business Expenses", GST_INCLUDED,
     "Utility provider – home or business?"),
    # --- fines (not deductible) ---
    (r"SDRO|INFRNGMNT|INFRINGE|\bFINE\b|PENALTY", "Revenue NSW (fines)", "fine", "Fines / Penalties", GST_NA,
     "Fines and penalties are generally NOT tax deductible"),
    # --- transfers to other accounts ---
    (r"REVOLUT", "Revolut", "transfer", "Transfer between related accounts", GST_NA,
     "Card payment to Revolut – possibly a top-up of another account you own"),
    # --- personal ---
    (r"WOOLWORTHS", "Woolworths", "personal", "Groceries", GST_NA, ""),
    (r"\bCOLES\b(?! EXPRESS)", "Coles", "personal", "Groceries", GST_NA, ""),
    (r"\bALDI\b", "Aldi", "personal", "Groceries", GST_NA, ""),
    (r"COSTCO", "Costco", "personal", "Groceries", GST_NA, ""),
    (r"\bIGA\b|HARRIS FARM|FRESH ASIANA|MONGOL SHOP|THAI KEE|TONG LI|CITY MART|EZYMART|STOP N SAVE|KMALL|YOUSAVE|"
     r"TOP HAPPIGO", None, "personal", "Groceries", GST_NA, ""),
    (r"UBER ?\*? ?EATS|MENULOG|DOORDASH", None, "personal", "Food & Dining", GST_NA, ""),
    (r"CAFE|COFFEE|ESPRESSO|BAKERY|CAKE|SUSHI|MCDONALD|\bKFC\b|HUNGRY JACK|PIZZA|RESTAURANT|KITCHEN|GRILL|HOTEL|"
     r"\bBAR\b|FOOD & BEVER|VENDING|\bVEND\b|GER HOUSE|FARINA|THREE CHEFS|YOGURBERRY|NOMADS|AZUKI|SANDWICH|"
     r"NATURALLY TASTY|DUMPLING|NOODLE|KEBAB|BURGER", None, "personal", "Food & Dining", GST_NA, ""),
    (r"DAN MURPHY|LIQUORLAND|CELLARS|\bBWS\b", None, "personal", "Alcohol", GST_NA, ""),
    (r"KIDS|KIDDO|SMIGGLE|FUNSTART|FOREVER FUN|ZONE BOWLING|AQUATIC|LEISURE C|TIMEZONE|HAPPY FACES|HOBBYCO", None,
     "personal",
     "Children / Family", GST_NA, ""),
    (r"\bZARA|UNIQLO|\bMYER\b|DAVID JONES|\bKMART\b|BIG W|\bTARGET\b|\bIKEA\b|CASTLERY|BED BATH|ADAIRS|"
     r"VICTORIAS BASEMENT", None, "personal",
     "Household / Clothing", GST_NA, ""),
    (r"PHARMACY|CHEMIST|MYHEALTH|MEDICAL|DENTAL|PATHOLOGY|\bBUPA\b|MEDIBANK|RADIOLOGY|I-MED|^DR |DOUGLASS HANLY",
     None, "personal", "Health", GST_NA, ""),
    (r"BEAUTY|HAIR|NAILS|BARBER", None, "personal", "Personal Care", GST_NA, ""),
    (r"PRIMEVIDEO|NETFLIX|SPOTIFY|DISNEY|\bSTAN\b|CINEMA|HOYTS", None, "personal", "Entertainment / Subscriptions",
     GST_NA, ""),
    (r"TATTS|LOTTO|SPORTSBET|LADBROKES", None, "personal", "Gambling", GST_NA, ""),
]
PERSONAL_NOTE = "Treated as personal spending – change if this was a business purchase"

# =============================================================================
# 4. EXCEL STYLE
# =============================================================================
AUD = '"$"#,##0.00;[Red]-"$"#,##0.00'
DATE_FMT = "DD/MM/YYYY"
PCT = "0.0%"
BLUE = "1F4E79"
F_HEADER = PatternFill("solid", fgColor=BLUE)
F_SECTION = PatternFill("solid", fgColor="D9E1F2")
F_SUBTOTAL = PatternFill("solid", fgColor="EDEDED")
F_GRAND = PatternFill("solid", fgColor="BDD7EE")
F_REVIEW = PatternFill("solid", fgColor="FCE4D6")
F_OK = PatternFill("solid", fgColor="C6EFCE")
F_BAD = PatternFill("solid", fgColor="FFC7CE")
F_NOTE = PatternFill("solid", fgColor="FFF2CC")
FONT_HEADER = Font(bold=True, color="FFFFFF")
FONT_BOLD = Font(bold=True)
FONT_TITLE = Font(bold=True, size=16, color=BLUE)
FONT_SUB = Font(bold=True, size=12, color=BLUE)
THIN = Side(style="thin", color="BFBFBF")
BORDER_TOP = Border(top=Side(style="thin", color="000000"))
BORDER_DOUBLE = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

ALL_TX = "All Transactions"


# =============================================================================
# 5. READING THE CSV
# =============================================================================
COLUMN_ALIASES = {
    "date": ["date", "transaction date", "posting date", "posted date", "value date", "txn date", "огноо"],
    "description": ["description", "details", "transaction details", "narrative", "memo", "particulars",
                    "transaction description", "reference", "тайлбар", "гүйлгээний тайлбар"],
    "amount": ["amount", "transaction amount", "value", "amount (aud)", "дүн"],
    "credit": ["credit", "credits", "credit amount", "deposit", "deposits", "money in", "paid in", "income", "орлого"],
    "debit": ["debit", "debits", "debit amount", "withdrawal", "withdrawals", "money out", "paid out", "expense",
              "зарлага"],
    "balance": ["balance", "running balance", "closing balance", "үлдэгдэл", "дансны үлдэгдэл"],
}
DATE_FORMATS = ["%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
                "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%Y-%m-%d %H:%M:%S"]


def parse_amount(value):
    """
    Convert a bank amount text to a number.
    Returns None for an empty cell. Raises ValueError if it is not a number.
    Handles: "1,500.00", "-85.20", "$1,234.50", "(500)", "500-", "250.00 CR", "250.00 DR".
    """
    text = str(value if value is not None else "").strip()
    if text == "":
        return None
    text = re.sub(r"(?i)(aud|[$\s '])", "", text)
    negative = False
    if text.upper().endswith("DR"):
        negative, text = True, text[:-2]
    elif text.upper().endswith("CR"):
        text = text[:-2]
    if text.startswith("(") and text.endswith(")"):
        negative, text = True, text[1:-1]
    if text.endswith("-"):
        negative, text = True, text[:-1]
    if text.startswith("+"):
        text = text[1:]
    if text.startswith("-"):
        negative, text = not negative, text[1:]
    if text == "" or re.search(r"[^\d.,]", text):
        raise ValueError(value)
    if "," in text and "." in text and text.rfind(",") > text.rfind("."):
        text = text.replace(".", "").replace(",", ".")      # 1.234,56
    else:
        text = text.replace(",", "")                          # 1,234.56
    number = float(text)
    return -number if negative else number


def parse_date(value):
    """Convert a date text to a datetime. Returns None if it cannot be read."""
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(text, dayfirst=DATE_DAY_FIRST).to_pydatetime()
    except (ValueError, TypeError):
        return None


def looks_like_number(value):
    """True if the text is a number (used to detect CSVs without a header row)."""
    try:
        return parse_amount(value) is not None
    except ValueError:
        return False


def read_csv_rows(path):
    """
    Read the CSV file exactly as text (no values changed).
    Returns (header or None, data_rows, csv_line_numbers).
    """
    raw = Path(path).read_bytes()
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines()
    sample = [ln for ln in lines[:20] if ln.strip()]
    if not sample:
        raise ValueError("The CSV file is empty.")
    # Choose the delimiter that appears consistently in the first lines.
    delimiter = max([",", ";", "\t", "|"], key=lambda d: min(ln.count(d) for ln in sample))

    all_rows, line_numbers = [], []
    for line_no, row in enumerate(csv.reader(lines, delimiter=delimiter), start=1):
        if any(cell.strip() for cell in row):          # skip completely blank lines only
            all_rows.append(row)
            line_numbers.append(line_no)
    if not all_rows:
        raise ValueError("The CSV file has no rows.")

    first = all_rows[0]
    known = {a for aliases in COLUMN_ALIASES.values() for a in aliases}
    has_header = any(c.strip().lower() in known for c in first) or not (
        any(parse_date(c) for c in first) and any(looks_like_number(c) for c in first))
    if has_header:
        return [c.strip() for c in first], all_rows[1:], line_numbers[1:]
    return None, all_rows, line_numbers


def detect_columns(header, rows):
    """
    Work out which column is Date / Description / Amount / Credit / Debit / Balance.
    Uses the header names if present, otherwise inspects the values.
    Returns a dict like {"date": 0, "amount": 1, "description": 2, "balance": 3}.
    """
    mapping = {}
    if header:
        lowered = [h.lower() for h in header]
        for key, aliases in COLUMN_ALIASES.items():          # exact names first
            for i, name in enumerate(lowered):
                if i not in mapping.values() and name in aliases:
                    mapping[key] = i
                    break
        for key, aliases in COLUMN_ALIASES.items():          # then partial names
            if key in mapping:
                continue
            for i, name in enumerate(lowered):
                if i not in mapping.values() and any(len(a) > 3 and a in name for a in aliases):
                    mapping[key] = i
                    break
        print(f"   Columns found in header: {header}")
    else:
        sample = rows[:50]
        width = max(len(r) for r in sample)
        numeric, text_cols = [], []
        for i in range(width):
            values = [r[i] for r in sample if i < len(r) and r[i].strip()]
            if not values:
                continue
            dates = sum(parse_date(v) is not None and not looks_like_number(v) for v in values) / len(values)
            nums = sum(looks_like_number(v) for v in values) / len(values)
            if "date" not in mapping and dates >= 0.8:
                mapping["date"] = i
            elif nums >= 0.8:
                numeric.append(i)
            else:
                text_cols.append(i)
        if numeric:
            mapping["amount"] = numeric[0]
        if len(numeric) > 1:
            mapping["balance"] = numeric[-1]
        if text_cols:
            mapping["description"] = max(
                text_cols, key=lambda c: sum(len(r[c]) for r in sample if c < len(r)))
        print("   No header row found – columns detected from their contents.")

    if "date" not in mapping:
        raise ValueError("Could not find a Date column in the CSV.")
    if "amount" not in mapping and not ("credit" in mapping or "debit" in mapping):
        raise ValueError("Could not find an Amount (or Debit/Credit) column in the CSV.")
    if "description" not in mapping:
        print("   WARNING: no Description column found – descriptions will be blank.")
    readable = {k: (header[v] if header else f"column {v + 1}") for k, v in mapping.items()}
    print(f"   Column mapping: {readable}")
    return mapping


def cell(row, index):
    """Safe access to a CSV cell (returns '' if the row is short)."""
    return row[index] if index is not None and index < len(row) else ""


def row_amount(row, mapping):
    """
    Signed amount of one CSV row: positive = money IN, negative = money OUT.
    Returns (amount or None, problem text or None).
    """
    try:
        if "amount" in mapping:
            value = parse_amount(cell(row, mapping["amount"]))
            return value, (None if value is not None else "Amount is empty")
        credit = parse_amount(cell(row, mapping.get("credit"))) if "credit" in mapping else None
        debit = parse_amount(cell(row, mapping.get("debit"))) if "debit" in mapping else None
        if credit is None and debit is None:
            return None, "Debit and Credit are both empty"
        return abs(credit or 0.0) - abs(debit or 0.0), None
    except ValueError:
        return None, "Amount is not a number"


def original_csv_totals(rows, mapping):
    """
    INDEPENDENT check used for validation: total money in / out straight from the CSV rows.
    """
    total_in = total_out = 0.0
    unreadable = 0
    for row in rows:
        amount, _problem = row_amount(row, mapping)
        if amount is None:
            unreadable += 1
        elif amount > 0:
            total_in += amount
        else:
            total_out += -amount
    return {"rows": len(rows), "in": round(total_in, 2), "out": round(total_out, 2), "unreadable": unreadable}


# =============================================================================
# 6. NAMES — find and standardise the person / company
# =============================================================================
def words(text):
    """Lower-case words of a text (letters, digits, hyphens)."""
    return re.findall(r"[a-z0-9][a-z0-9\-']*", str(text).lower())


def similar(a, b):
    """Similarity of two words (0–1)."""
    return SequenceMatcher(None, a, b).ratio()


def is_same_person(found_name, known_name):
    """
    Strong-evidence name match: case-insensitive, small spelling differences,
    and bank-truncated surnames (e.g. "MUNKHCHIMEG LKHAGVATS").
    One-word names (e.g. "Deleg") must match a whole word exactly.
    """
    found, known = words(found_name), words(known_name)
    if not found or not known:
        return False
    if len(known) == 1:
        return known[0] in found
    n = len(known)
    for start in range(len(found) - n + 1):
        window = found[start:start + n]
        for order in (known, list(reversed(known))):
            head = all(window[i] == order[i] or similar(window[i], order[i]) >= 0.85 for i in range(n - 1))
            tail = (window[-1] == order[-1] or similar(window[-1], order[-1]) >= 0.85
                    or (len(window[-1]) >= 4 and order[-1].startswith(window[-1])))
            if head and tail:
                return True
    return False


def is_organisation(name):
    """True if the name contains words like PTY, LTD, GROUP, SERVICES."""
    return any(w in ORGANISATION_WORDS for w in words(name))


def clean_merchant(text):
    """Remove card numbers, 'Value Date', store numbers etc. from a card purchase description."""
    s = re.sub(r"\s*(AU\s+)?Card\s+xx\d+.*$", "", text, flags=re.I)
    s = re.sub(r"\s*Value Date:.*$", "", s, flags=re.I)
    s = re.sub(r"^(SQ \*|SMP\*|NXT\*|ZLR\*|V\*|PAYPAL \*|SP \*?)", "", s.strip(), flags=re.I)
    s = re.sub(r"\*+", " ", s)
    s = re.sub(r"\b[A-Z]?\d{3,}[A-Z]?\b", "", s, flags=re.I)
    s = re.sub(r"\s+(AU|NSWAU|NSW|AUS)$", "", s.strip(), flags=re.I)
    return re.sub(r"\s+", " ", s).strip().upper() or text.strip().upper()


REF_STOP_WORDS = {"gst", "abn", "inv", "invoice", "ref", "payment", "transfer", "netbank", "commbank",
                  "payid", "app", "bpay", "osko", "from", "for"}


def name_tokens(text, uppercase_run):
    """
    Take the name at the start of `text`.
    uppercase_run=True : bank writes names in CAPITALS followed by a mixed-case reference.
    Stops at reference words (GST, NETBANK, PAYID, ...) or tokens containing digits.
    """
    tokens = []
    for tok in text.split():
        clean = tok.strip(",.;:()")
        low = clean.lower()
        if not clean or low in REF_STOP_WORDS or re.search(r"\d", clean):
            break
        if uppercase_run and not clean.isupper():
            break
        tokens.append(clean)
        if not uppercase_run and len(tokens) >= 4:
            break
    return " ".join(tokens)


def extract_counterparty(description):
    """
    Find who the money came from / went to.
    Returns (raw_name, kind). kind is one of:
      ato, bnpl, bank_fee, own_account, person_or_company_in, person_or_company_out,
      direct_credit, direct_debit, bpay, card, other
    """
    s = description.strip()
    u = s.upper()
    if re.search(r"TAX OFFICE PAYMENTS|AUSTRALIAN TAXATION OFFICE|\bATO\b", u):
        return "Australian Taxation Office", "ato"
    m = re.search(r"STEPPAY|ZIPPAY|ZIPMONEY|ZIP\.CO|AFTERPAY|HUMM\b|LATITUDE PAY", u)
    if m:
        brand = {"STEPPAY": "StepPay"}.get(m.group(0), "Zip Pay" if "ZIP" in m.group(0) else m.group(0).title())
        return brand, "bnpl"
    if u.startswith("INTERNATIONAL TRANSACTION FEE") or re.match(r"^(ACCOUNT|MONTHLY|OVERDRAWN|DISHONOUR)\w* FEE", u):
        return "Commonwealth Bank (fees)", "bank_fee"
    m = re.match(r"TRANSFER TO (XX\d+|CBA A/C|.*\bA/C\b)", u)
    if m:
        return f"Own account ({m.group(1).strip()})", "own_account"
    if re.match(r"TRANSFER FROM (COMMBANK APP|NETBANK)\b", u):
        return "Own account (not named)", "own_account"
    m = re.match(r"^(?:PAYMENT|WAGES?|SALARY|PAYROLL)\b\s*(?:TO|FROM)?\s*[-:–]?\s*(.+)", s, flags=re.I)
    if m:
        name = name_tokens(m.group(1), uppercase_run=False)
        if name:
            return name, "person_or_company_out"
    m = re.match(r"(?:FAST )?TRANSFER FROM\s+(.+)", s, flags=re.I)
    if m:
        rest = m.group(1)
        first = rest.split()[0] if rest.split() else ""
        name = name_tokens(rest, uppercase_run=first.isupper())
        return (name or rest), "person_or_company_in"
    m = re.match(r"TRANSFER TO\s+(.+)", s, flags=re.I)
    if m:
        return (name_tokens(m.group(1), uppercase_run=False) or m.group(1)), "person_or_company_out"
    m = re.match(r"DIRECT CREDIT\s+\d*\s*(.+)", s, flags=re.I)
    if m:
        rest = re.sub(r"\s+\d+$", "", m.group(1).strip())
        half = rest.split()
        n = len(half) // 2
        if n and len(half) % 2 == 0 and " ".join(half[:n]).upper() == " ".join(half[n:]).upper():
            rest = " ".join(half[:n])               # "Richard Crookes RICHARD CROOKES"
        return rest, "direct_credit"
    m = re.match(r"DIRECT DEBIT\s+\d*\s*([A-Za-z][A-Za-z &\-]*)", s, flags=re.I)
    if m:
        return m.group(1).strip(" -"), "direct_debit"
    m = re.match(r"(.+?)\s+(?:NETBANK |COMMBANK APP )?BPAY\b", s, flags=re.I)
    if m:
        return m.group(1).strip(), "bpay"
    m = re.match(r"(?:RETURN|ADJUST PURCHASE|REFUND)\s+(.+)", s, flags=re.I)
    if m:
        return clean_merchant(m.group(1)), "card"
    if re.search(r"CARD XX\d+|VALUE DATE", u) or re.search(r"\sAU$", u):
        return clean_merchant(s), "card"
    return clean_merchant(s), "other"


def match_rule(description):
    """Return the first MERCHANT_RULE that matches the description, or None."""
    upper = description.upper()
    for rule in MERCHANT_RULES:
        if re.search(rule[0], upper):
            return rule
    return None


def standardise_name(raw_name, kind, rule):
    """
    Produce the standard person/company name.
    Only strong evidence is used (known names, aliases, merchant rules, same spelling).
    """
    for known in DIRECTOR_NAMES + SPECIFIC_PEOPLE + KNOWN_CUSTOMERS:
        if is_same_person(raw_name, known):
            return known
    for standard, aliases in NAME_ALIASES.items():
        upper = raw_name.upper()
        if upper == standard.upper() or any(upper == a.upper() or upper.startswith(a.upper() + " ")
                                            for a in aliases):
            return standard
    if rule and rule[1] and kind not in ("person_or_company_in", "person_or_company_out"):
        return rule[1]
    if kind in ("person_or_company_in", "person_or_company_out", "direct_credit"):
        return " ".join("-".join(p.capitalize() for p in w.split("-")) for w in raw_name.split())
    return raw_name.upper() if kind in ("card", "other") else raw_name


# =============================================================================
# 7. CLASSIFICATION
# =============================================================================
def has_word(text, word_list):
    """True if any of the words appears as a whole word in the text."""
    low = " " + re.sub(r"[^a-z0-9]+", " ", text.lower()) + " "
    return any(f" {w} " in low for w in word_list)


def reference_text(description, name):
    """The description with the person/company name removed (what is left is the payment reference)."""
    if not name:
        return description
    idx = description.lower().find(name.lower())
    return (description[:idx] + " " + description[idx + len(name):]) if idx >= 0 else description


PERSON_KINDS = ("person_or_company_in", "person_or_company_out")


def new_result():
    """Empty classification result for one transaction."""
    return {"Classification": REVIEW, "Suggested Classification": "", "Category": "",
            "Business / Personal": "UNKNOWN", "GST Status": GST_UNKNOWN, "GST Amount": 0.0,
            "Potential GST": 0.0, "Tax Deductible": "N/A", "Loan Type": "", "Counterparty Type": "Unknown",
            "reasons": [], "notes": []}


def classify(tx, stats):
    """
    Classify ONE transaction. Rules are applied in a fixed order so they never conflict:
      1 unreadable  2 directors  3 ATO  4 own-account transfers  5 buy-now-pay-later
      6 refunds     7 merchant rules  8 income from people/companies
      9 payments to individuals (salary)  10 anything else -> review
    """
    r = new_result()
    desc, amount, name, kind = tx["Original Description"], tx["Amount"], tx["Person / Company"], tx["kind"]
    direction = tx["Direction"]
    ref = reference_text(desc, tx["raw_name"])
    rule = tx["rule"]

    # 1. Unreadable rows are kept, but flagged.
    if amount is None or direction not in (INCOME, EXPENSE):
        r["reasons"].append(f"{R_UNREADABLE}: {tx['problem'] or 'amount is zero'} – check the original CSV row")
        r["Suggested Classification"] = REVIEW
        return r

    is_income = direction == INCOME

    # 2. Directors / personal names.
    if name in DIRECTOR_NAMES:
        r.update({"Counterparty Type": "Director / Personal", "Business / Personal": "DIRECTOR",
                  "GST Status": GST_NA, "Tax Deductible": "N/A" if is_income else "NO"})
        if is_income and has_word(ref, LOAN_WORDS):
            r.update({"Classification": LOAN, "Category": "Director loan – money borrowed",
                      "Loan Type": "Borrowed (director loan)"})
            r["reasons"].append(f"{R_DIRECTOR}: money received from director with loan reference – confirm loan")
        elif is_income:
            loan_type = ("Possible repayment received" if has_word(ref, RETURN_WORDS)
                         else "Possible loan / capital contribution")
            r.update({"Classification": DIRECTOR, "Category": "Director – money received", "Loan Type": loan_type})
            r["reasons"].append(f"{R_DIRECTOR}: money received from director – NOT treated as business income")
        else:
            r.update({"Classification": DIRECTOR, "Category": "Director – money paid",
                      "Loan Type": "Possible loan repayment / drawings"})
            r["reasons"].append(f"{R_DIRECTOR}: payment to director – NOT treated as a business expense")
        r["Suggested Classification"] = r["Classification"]
        return r

    # A director's name inside someone else's reference is worth a look.
    for director in DIRECTOR_NAMES:
        if is_same_person(ref, director) and name != director:
            r["reasons"].append(f"{R_DIRECTOR}: reference mentions {director}")

    # 3. Australian Taxation Office.
    if kind == "ato":
        r.update({"Counterparty Type": "Government (ATO)", "Business / Personal": "BUSINESS", "GST Status": GST_NA,
                  "Tax Deductible": "NO" if not is_income else "N/A"})
        if re.search(r"\bBAS\b", desc, re.I):
            r.update({"Classification": GST, "Category": "GST Payments (BAS)"})
            r["notes"].append("BAS payment – may also include PAYG instalments/withholding")
        elif re.search(r"INSTAL", desc, re.I):
            r.update({"Classification": TAX, "Category": "Tax Payments (PAYG instalment)"})
        else:
            r.update({"Classification": TAX, "Category": "ATO Payments"})
        if is_income:
            r["Category"] = "ATO refund"
            r["reasons"].append(f"{R_REFUND}: money received from the ATO – confirm what it relates to")
        r["Suggested Classification"] = r["Classification"]
        return r

    # 4. Transfers between your own accounts.
    if kind == "own_account" or (rule and rule[2] == "transfer"):
        r.update({"Counterparty Type": "Own / related account", "Business / Personal": "NOT APPLICABLE",
                  "GST Status": GST_NA, "Category": "Transfer between related accounts",
                  "Tax Deductible": "N/A" if is_income else "NO"})
        if kind == "own_account" and not is_income:
            r["Classification"] = TRANSFER
            r["notes"].append("Transfer to your own account – not income or expense")
        else:
            r["Classification"] = TRANSFER if kind == "own_account" else REVIEW
            detail = rule[5] if rule else "source account not shown in the description"
            r["reasons"].append(f"{R_TRANSFER}: {detail}")
        r["Suggested Classification"] = TRANSFER
        return r

    # 5. Buy-now-pay-later (StepPay, Zip, Afterpay) — these are loan repayments.
    if kind == "bnpl":
        r.update({"Counterparty Type": "Finance (buy now pay later)", "Business / Personal": "NOT APPLICABLE",
                  "GST Status": GST_NA})
        if is_income:
            r.update({"Classification": REVIEW, "Suggested Classification": LOAN,
                      "Category": "Refund from buy-now-pay-later provider"})
            r["reasons"].append(f"{R_REFUND}: refund/adjustment from {name}")
        else:
            r.update({"Classification": LOAN, "Suggested Classification": LOAN, "Tax Deductible": "NO",
                      "Category": "Loan repayment – buy now pay later", "Loan Type": "Repayment"})
            r["notes"].append("Repayment of buy-now-pay-later credit – the underlying purchases are not shown "
                              "in the bank statement")
        return r

    # 6. Refunds / returns / reversals of purchases (money in).
    if is_income and re.match(r"^(RETURN|ADJUST PURCHASE|REFUND)\b", desc.strip(), re.I) or \
            (is_income and re.search(r"\bREFUNDS?\b", desc, re.I)):
        r.update({"Classification": REVIEW, "Suggested Classification": "Refund / purchase reversal",
                  "Category": (rule[3] if rule else "Refund"), "Counterparty Type": "Supplier / Merchant"})
        r["reasons"].append(f"{R_REFUND}: refund or purchase reversal – match it to the original purchase")
        return r

    # 7. Merchant rules (card purchases, direct debits, BPAY, payments to companies).
    name_rule = match_rule(tx["raw_name"]) if kind in PERSON_KINDS else None
    if name_rule:
        rule = name_rule
    if rule and not is_income and (kind not in PERSON_KINDS or is_organisation(name) or name_rule):
        level, category, gst, note = rule[2], rule[3], rule[4], rule[5]
        r["Category"] = category
        if note:
            r["notes"].append(note)
        overseas = bool(re.search(r"\b(AE|US|GB|SG|HK|CN)\b\s*$|DUBAI|SHARJAH|\bUSD\b", desc.upper()))
        if level == "business":
            gst_status = NO_GST if overseas else gst
            r.update({"Classification": BUS_EXP, "Suggested Classification": BUS_EXP, "Business / Personal": "BUSINESS",
                      "Counterparty Type": "Supplier", "Tax Deductible": "YES", "GST Status": gst_status,
                      "GST Amount": abs(amount) / 11 if gst_status == GST_INCLUDED else 0.0})
        elif level == "possible":
            r.update({"Classification": REVIEW, "Suggested Classification": BUS_EXP, "Counterparty Type": "Supplier",
                      "Tax Deductible": "REVIEW", "GST Status": GST_UNKNOWN,
                      "Potential GST": abs(amount) / 11 if gst == GST_INCLUDED and not overseas else 0.0})
            r["reasons"].append(f"{R_EXPENSE}: {note or 'confirm business use'}")
        elif level == "fine":
            r.update({"Classification": REVIEW, "Suggested Classification": BUS_EXP, "Tax Deductible": "NO",
                      "GST Status": GST_NA, "Counterparty Type": "Government"})
            r["reasons"].append(f"{R_EXPENSE}: {note}")
        else:  # personal
            r.update({"Classification": PER_EXP, "Suggested Classification": PER_EXP, "Business / Personal": "PERSONAL",
                      "Counterparty Type": "Merchant (personal)", "Tax Deductible": "NO", "GST Status": GST_NA})
            r["notes"].append(PERSONAL_NOTE)
        return r

    # 8. Money received from a person or company.
    if is_income:
        specific = name in SPECIFIC_PEOPLE
        if re.search(r"\bGST\b", desc, re.I):
            r.update({"Classification": BUS_INC, "Suggested Classification": BUS_INC, "Counterparty Type": "Customer",
                      "Business / Personal": "BUSINESS", "GST Status": GST_INCLUDED, "GST Amount": amount / 11})
            r["notes"].append("Reference mentions GST (and an ABN) – treated as a GST-inclusive customer payment")
            return r
        if name in KNOWN_CUSTOMERS or name in GST_INCLUSIVE_PAYERS or name in NO_GST_PAYERS:
            r.update({"Classification": BUS_INC, "Suggested Classification": BUS_INC, "Counterparty Type": "Customer",
                      "Business / Personal": "BUSINESS"})
            if name in GST_INCLUSIVE_PAYERS:
                r.update({"GST Status": GST_INCLUDED, "GST Amount": amount / 11})
            elif name in NO_GST_PAYERS:
                r["GST Status"] = NO_GST
            else:
                r["reasons"].append(f"{R_GST}: confirm whether invoices to {name} included GST "
                                    f"(then add the name to GST_INCLUSIVE_PAYERS or NO_GST_PAYERS)")
            return r
        if kind in PERSON_KINDS and not is_organisation(name):
            r.update({"Counterparty Type": "Specific person" if specific else "Individual",
                      "GST Status": GST_UNKNOWN, "Business / Personal": "UNKNOWN"})
            if has_word(ref, LOAN_WORDS):
                r.update({"Classification": LOAN, "Suggested Classification": LOAN, "Category": "Money borrowed",
                          "Loan Type": "Borrowed"})
                r["reasons"].append(f"{R_LOAN}: reference mentions a loan – confirm the loan terms")
                return r
            clues = []
            if has_word(ref, BUSINESS_WORK_WORDS):
                clues.append("reference mentions business work (could be business income)")
            if has_word(ref, RETURN_WORDS):
                clues.append("reference suggests money being returned")
            if has_word(ref, SALARY_WORDS):
                clues.append("reference mentions wages/payment")
            r.update({"Classification": REVIEW, "Suggested Classification": "Possible loan / possible business income",
                      "Category": "Incoming transfer from individual", "Loan Type": "Possible loan"})
            r["reasons"].append(f"{R_LOAN}: incoming transfer from an individual – NOT counted as confirmed "
                                f"business income" + (f" ({'; '.join(clues)})" if clues else ""))
            return r
        r.update({"Classification": REVIEW, "Suggested Classification": BUS_INC, "Counterparty Type": "Unknown",
                  "Category": "Unknown income source", "GST Status": GST_UNKNOWN})
        r["reasons"].append(f"{R_UNKNOWN}: money received from {name} – source not confirmed")
        return r

    # 9. Payments to individuals.
    if kind in PERSON_KINDS and not is_organisation(name):
        specific = name in SPECIFIC_PEOPLE
        r["GST Status"] = GST_NA
        if has_word(ref, SALARY_WORDS):
            r.update({"Classification": SALARY, "Suggested Classification": SALARY, "Category": "Salary / Wages",
                      "Counterparty Type": "Worker", "Business / Personal": "BUSINESS", "Tax Deductible": "YES"})
            r["notes"].append("Wage keyword found – PAYG withholding and superannuation obligations not checked")
            return r
        if has_word(ref, LOAN_WORDS + ["repay", "repayment"]):
            r.update({"Classification": REVIEW, "Suggested Classification": LOAN, "Category": "Money lent / loan repaid",
                      "Counterparty Type": "Specific person" if specific else "Individual",
                      "Loan Type": "Possible repayment", "Tax Deductible": "NO"})
            r["reasons"].append(f"{R_LOAN}: payment reference mentions a loan/repayment")
            return r
        s = stats.get(name, {})
        r.update({"Classification": REVIEW, "Counterparty Type": "Specific person" if specific else "Individual",
                  "Tax Deductible": "REVIEW", "Business / Personal": "UNKNOWN"})
        if s.get("salary", 0) > 0 or s.get("out", 0) >= REPEATED_PAYMENT_MIN:
            r.update({"Suggested Classification": SALARY, "Category": "Salary / Wages"})
            r["reasons"].append(f"{R_SALARY}: {s.get('out', 0)} payments to this person "
                                f"({s.get('salary', 0)} with a wage keyword) – this one has no wage keyword")
        else:
            r.update({"Suggested Classification": "Unknown", "Category": "Payment to individual"})
            r["reasons"].append(f"{R_UNKNOWN}: payment to an individual – purpose unclear")
        return r

    # 10. Everything else (unknown companies/merchants, bank fees).
    if kind == "bank_fee":
        r.update({"Classification": REVIEW, "Suggested Classification": BUS_EXP, "Category": "Bank Fees",
                  "Counterparty Type": "Bank", "GST Status": NO_GST, "Tax Deductible": "REVIEW"})
        r["reasons"].append(f"{R_EXPENSE}: bank fee (usually charged on an overseas card purchase) – "
                            f"business purpose unknown")
        return r
    r.update({"Classification": REVIEW, "Suggested Classification": "Unknown", "Category": "Other Business Expenses",
              "Counterparty Type": "Supplier / Merchant", "Tax Deductible": "REVIEW"})
    what = "supplier" if kind in ("person_or_company_out", "bpay", "direct_debit") else "merchant"
    r["reasons"].append(f"{R_EXPENSE}: {what} not recognised – business or personal?")
    return r


# =============================================================================
# 8. BUILD THE TRANSACTION TABLE
# =============================================================================
def build_transactions(rows, line_numbers, mapping):
    """Turn CSV rows into transaction dicts (nothing is removed)."""
    transactions = []
    for row, line_no in zip(rows, line_numbers):
        amount, problem = row_amount(row, mapping)
        date_raw = cell(row, mapping["date"])
        date = parse_date(date_raw)
        if date is None:
            problem = (problem + "; " if problem else "") + f"date '{date_raw}' could not be read"
        balance = None
        if "balance" in mapping:
            try:
                balance = parse_amount(cell(row, mapping["balance"]))
            except ValueError:
                balance = None
        description = cell(row, mapping.get("description"))
        direction = INCOME if amount is not None and amount > 0 else (
            EXPENSE if amount is not None and amount < 0 else "UNREADABLE" if amount is None else "ZERO")
        transactions.append({
            "Row #": line_no, "Date": date, "date_raw": date_raw, "Original Description": description,
            "Direction": direction, "Amount": amount,
            "Money In": amount if amount is not None and amount > 0 else 0.0,
            "Money Out": -amount if amount is not None and amount < 0 else 0.0,
            "Balance": balance, "problem": problem,
        })
    return transactions


def identify_counterparties(transactions):
    """Find and standardise the person/company for every transaction."""
    for tx in transactions:
        raw_name, kind = extract_counterparty(tx["Original Description"])
        rule = match_rule(tx["Original Description"])
        tx.update({"raw_name": raw_name, "kind": kind, "rule": rule,
                   "Person / Company": standardise_name(raw_name, kind, rule)})

    # Strong evidence merge: the bank truncates long names (~20 characters).
    names = sorted({tx["Person / Company"] for tx in transactions}, key=len, reverse=True)
    merge = {}
    for short in names:
        if len(short) < 18:
            continue
        for full in names:
            if len(full) > len(short) and full.upper().startswith(short.upper()):
                merge[short] = full
                break
    for tx in transactions:
        tx["Person / Company"] = merge.get(tx["Person / Company"], tx["Person / Company"])


def flag_similar_names(transactions):
    """
    Names that look alike but are NOT merged (not enough evidence) get a review note,
    e.g. "Tamir" and "Tamir Ganburged".
    """
    people = defaultdict(int)
    for tx in transactions:
        if tx["kind"] in ("person_or_company_in", "person_or_company_out", "direct_credit"):
            people[tx["Person / Company"]] += 1
    names = list(people)
    lookalike = defaultdict(set)
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            wa, wb = words(a), words(b)
            if not wa or not wb or a.upper() == b.upper():
                continue
            shorter, longer = (wa, wb) if len(wa) <= len(wb) else (wb, wa)
            prefix_same = shorter == longer[:len(shorter)]
            first_alike = similar(wa[0], wb[0]) >= 0.85
            if prefix_same or (first_alike and similar(a.lower(), b.lower()) >= 0.88):
                lookalike[a].add(b)
                lookalike[b].add(a)
    for tx in transactions:
        others = lookalike.get(tx["Person / Company"])
        if others:
            tx["result"]["reasons"].append(
                f"{R_SIMILAR}: '{tx['Person / Company']}' looks similar to {', '.join(sorted(others))} – not merged")


def payment_stats(transactions):
    """Count payments to each individual (used to spot possible workers)."""
    stats = defaultdict(lambda: {"out": 0, "salary": 0, "in": 0})
    for tx in transactions:
        name = tx["Person / Company"]
        if tx["kind"] in PERSON_KINDS and tx["Direction"] == EXPENSE:
            stats[name]["out"] += 1
            if has_word(reference_text(tx["Original Description"], tx["raw_name"]), SALARY_WORDS):
                stats[name]["salary"] += 1
        elif tx["Direction"] == INCOME:
            stats[name]["in"] += 1
    return stats


def flag_duplicates(transactions):
    """Identical date + description + amount: KEEP both, but flag for checking."""
    counts = defaultdict(int)
    for tx in transactions:
        counts[(tx["date_raw"], tx["Original Description"], tx["Amount"])] += 1
    duplicates = 0
    for tx in transactions:
        n = counts[(tx["date_raw"], tx["Original Description"], tx["Amount"])]
        tx["Possible Duplicate"] = "YES" if n > 1 else ""
        if n > 1:
            duplicates += 1
            tx["result"]["reasons"].append(
                f"{R_DUPLICATE}: {n} identical transactions (same date, description, amount) – all kept; "
                f"confirm they are separate payments")
    return duplicates


def balance_check(transactions):
    """
    Check the running balance: previous balance + amount = balance.
    A break can mean a missing transaction in the CSV. Returns the number of breaks (or None).
    """
    rows = [tx for tx in transactions if tx["Balance"] is not None and tx["Amount"] is not None]
    if len(rows) < 2:
        return None
    newest_first = (rows[0]["Date"] or datetime.min) >= (rows[-1]["Date"] or datetime.min)
    breaks = 0
    for i in range(1, len(rows)):
        if newest_first:
            ok = abs(rows[i]["Balance"] + rows[i - 1]["Amount"] - rows[i - 1]["Balance"]) < 0.01
        else:
            ok = abs(rows[i - 1]["Balance"] + rows[i]["Amount"] - rows[i]["Balance"]) < 0.01
        breaks += 0 if ok else 1
    return breaks


def finalise(transactions):
    """Copy classification results into flat columns and build the pandas DataFrame."""
    records = []
    for tx in transactions:
        r = tx["result"]
        amount = tx["Amount"] if tx["Amount"] is not None else 0.0
        gst = round(r["GST Amount"], 2)
        signed_gst = gst if amount >= 0 else -gst
        review = bool(r["reasons"]) or r["Classification"] == REVIEW
        records.append({
            "Row #": tx["Row #"], "Date": tx["Date"], "Original Description": tx["Original Description"],
            "Direction": tx["Direction"], "Amount": amount, "Money In": tx["Money In"], "Money Out": tx["Money Out"],
            "Balance": tx["Balance"], "Person / Company": tx["Person / Company"],
            "Counterparty Type": r["Counterparty Type"], "Classification": r["Classification"],
            "Suggested Classification": r["Suggested Classification"] or r["Classification"],
            "Category": r["Category"], "Business / Personal": r["Business / Personal"],
            "GST Status": r["GST Status"], "GST Amount": gst, "Amount excl GST": round(amount - signed_gst, 2),
            "Potential GST": round(r["Potential GST"], 2), "Tax Deductible": r["Tax Deductible"],
            "Loan Type": r["Loan Type"], "Review Required": "YES" if review else "NO",
            "Reason for Review": "; ".join(r["reasons"]) or ("Insufficient information" if review else ""),
            "Notes": "; ".join(r["notes"]), "Possible Duplicate": tx.get("Possible Duplicate", ""),
            "Year": tx["Date"].year if tx["Date"] else None, "Month": tx["Date"].month if tx["Date"] else None,
        })
    df = pd.DataFrame(records)
    df["_sort_date"] = df["Date"].fillna(pd.Timestamp.max)
    df = df.sort_values(["_sort_date", "Row #"], ascending=[True, False]).drop(columns="_sort_date")
    return df.reset_index(drop=True)


# =============================================================================
# 9. EXCEL HELPERS
# =============================================================================
def style_header(ws, row, first_col, last_col):
    """Blue bold header row."""
    for c in range(first_col, last_col + 1):
        cl = ws.cell(row=row, column=c)
        cl.font, cl.fill = FONT_HEADER, F_HEADER
        cl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_value(ws, row, col, value, fmt=None, bold=False, fill=None, border=None):
    """Write one cell with optional number format and style."""
    if value is not None and not isinstance(value, str) and pd.api.types.is_scalar(value) and pd.isna(value):
        value = None
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    cl = ws.cell(row=row, column=col, value=value)
    if fmt:
        cl.number_format = fmt
    if bold:
        cl.font = FONT_BOLD
    if fill:
        cl.fill = fill
    if border:
        cl.border = border
    return cl


def set_widths(ws, widths):
    """Set column widths from a list (index 0 = column A)."""
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title(ws, text, subtitle=None):
    """Sheet title in row 1 (and optional subtitle in row 2)."""
    ws.cell(row=1, column=1, value=text).font = FONT_TITLE
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = Font(italic=True, color="7F7F7F")


def section_title(ws, row, text, width_cols):
    """A shaded section heading across several columns."""
    for c in range(1, width_cols + 1):
        ws.cell(row=row, column=c).fill = F_SECTION
    cl = ws.cell(row=row, column=1, value=text)
    cl.font = FONT_SUB


FORMATS = {"Date": DATE_FMT, "Amount": AUD, "Money In": AUD, "Money Out": AUD, "Balance": AUD,
           "GST Amount": AUD, "Amount excl GST": AUD, "Potential GST": AUD, "Amount Received": AUD,
           "Amount Paid": AUD, "Amount Borrowed": AUD, "Amount Repaid": AUD}
WIDTHS = {"Row #": 7, "Date": 11, "Original Description": 60, "Direction": 10, "Amount": 14, "Money In": 13,
          "Money Out": 13, "Balance": 13, "Person / Company": 30, "Counterparty Type": 20, "Classification": 20,
          "Suggested Classification": 26, "Category": 30, "Business / Personal": 14, "GST Status": 15,
          "GST Amount": 12, "Amount excl GST": 14, "Potential GST": 12, "Tax Deductible": 11, "Loan Type": 28,
          "Review Required": 10, "Reason for Review": 70, "Notes": 50, "Possible Duplicate": 10, "Year": 7,
          "Month": 7, "Amount Received": 15, "Amount Paid": 15}


def write_flat_table(ws, df, columns, start_row=1, total_columns=(), freeze=True):
    """
    Write a DataFrame as a table: bold header, freeze, auto filter, formats,
    and a SUBTOTAL row (totals follow the filter) two rows below the data.
    Returns (first_data_row, last_data_row).
    """
    for j, col in enumerate(columns, start=1):
        ws.cell(row=start_row, column=j, value=col)
    style_header(ws, start_row, 1, len(columns))
    for i, rec in enumerate(df[columns].itertuples(index=False), start=start_row + 1):
        for j, (col, value) in enumerate(zip(columns, rec), start=1):
            write_value(ws, i, j, value, FORMATS.get(col))
    first, last = start_row + 1, start_row + max(len(df), 1)
    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(columns))}{last}"
    set_widths(ws, [WIDTHS.get(c, 14) for c in columns])
    if total_columns:
        total_row = last + 2
        write_value(ws, total_row, 1, "TOTAL (follows filter)", bold=True)
        for col in total_columns:
            j = columns.index(col) + 1
            letter = get_column_letter(j)
            write_value(ws, total_row, j, f"=SUBTOTAL(109,{letter}{first}:{letter}{last})", AUD, True,
                        F_GRAND, BORDER_DOUBLE)
    return first, last


def add_class_colours(ws, column_letter, first, last, width_cols):
    """Colour whole rows by classification / review status (conditional formatting)."""
    rng = f"A{first}:{get_column_letter(width_cols)}{last}"
    colours = {REVIEW: "FCE4D6", DIRECTOR: "E4DFEC", LOAN: "FFF2CC", TRANSFER: "EDEDED", BUS_INC: "E2EFDA",
               PER_EXP: "F2F2F2", SALARY: "DDEBF7"}
    for label, colour in colours.items():
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'${column_letter}{first}="{label}"'], fill=PatternFill("solid", fgColor=colour)))


CRITERIA_COLUMNS = {"GST_Status": "GST Status", "Tax_Deductible": "Tax Deductible",
                    "Possible_Duplicate": "Possible Duplicate", "Review_Required": "Review Required",
                    "Business_Personal": "Business / Personal"}


def criteria_column(key):
    """Keyword-argument name -> real column name (e.g. Business_Personal -> 'Business / Personal')."""
    return CRITERIA_COLUMNS.get(key, key.replace("_", " "))


class AllTxRef:
    """Builds SUMIFS/COUNTIFS formulas that point at the 'All Transactions' sheet."""

    def __init__(self, columns, first, last):
        self.columns, self.first, self.last = columns, first, last

    def rng(self, column):
        letter = get_column_letter(self.columns.index(column) + 1)
        return f"'{ALL_TX}'!${letter}${self.first}:${letter}${self.last}"

    def sumifs(self, sum_col, **criteria):
        parts = [self.rng(sum_col)]
        for col, value in criteria.items():
            parts += [self.rng(criteria_column(col)), f'"{value}"' if isinstance(value, str) else str(value)]
        return f"SUMIFS({','.join(parts)})"

    def countifs(self, **criteria):
        parts = []
        for col, value in criteria.items():
            parts += [self.rng(criteria_column(col)), f'"{value}"' if isinstance(value, str) else str(value)]
        return f"COUNTIFS({','.join(parts)})"


# =============================================================================
# 10. EXCEL SHEETS
# =============================================================================
ALL_COLUMNS = ["Row #", "Date", "Original Description", "Direction", "Amount", "Money In", "Money Out", "Balance",
               "Person / Company", "Counterparty Type", "Classification", "Suggested Classification", "Category",
               "Business / Personal", "GST Status", "GST Amount", "Amount excl GST", "Potential GST",
               "Tax Deductible", "Loan Type", "Review Required", "Reason for Review", "Notes",
               "Possible Duplicate", "Year", "Month"]


def sheet_all_transactions(wb, df):
    """Every single transaction with every classification column."""
    ws = wb.create_sheet(ALL_TX)
    first, last = write_flat_table(ws, df, ALL_COLUMNS, total_columns=("Amount", "Money In", "Money Out",
                                                                        "GST Amount", "Amount excl GST"))
    class_letter = get_column_letter(ALL_COLUMNS.index("Classification") + 1)
    add_class_colours(ws, class_letter, first, last, len(ALL_COLUMNS))
    return AllTxRef(ALL_COLUMNS, first, last)


def sheet_detail(wb, df, name, direction):
    """Income Detail / Expense Detail — flat list of all incoming or outgoing transactions."""
    ws = wb.create_sheet(name)
    part = df[df["Direction"] == direction]
    amount_col = "Money In" if direction == INCOME else "Money Out"
    columns = ["Row #", "Date", "Original Description", amount_col, "Person / Company", "Counterparty Type",
               "Classification", "Category", "Business / Personal", "GST Status", "GST Amount", "Amount excl GST"]
    if direction == EXPENSE:
        columns += ["Tax Deductible"]
    columns += ["Review Required", "Reason for Review", "Notes"]
    first, last = write_flat_table(ws, part, columns, total_columns=(amount_col, "GST Amount", "Amount excl GST"))
    add_class_colours(ws, get_column_letter(columns.index("Classification") + 1), first, last, len(columns))
    return part[amount_col].sum()


def write_grouped(ws, df, start_row, amount_col, detail_columns, total_label, width_cols):
    """
    Write one section per person/company: heading, table, TOTAL row (SUM formula).
    Groups are sorted from the highest total to the lowest.
    Returns (list of summary dicts, next free row).
    """
    groups = (df.groupby("Person / Company")[amount_col].sum().sort_values(ascending=False))
    row = start_row
    summary = []
    for name, total in groups.items():
        part = df[df["Person / Company"] == name].sort_values(["Date", "Row #"])
        ctype = part["Counterparty Type"].mode().iat[0] if not part.empty else ""
        section_title(ws, row, f"{name}   ({ctype})", width_cols)
        row += 1
        for j, col in enumerate(detail_columns, start=1):
            ws.cell(row=row, column=j, value="Amount" if col == amount_col else col)
        style_header(ws, row, 1, len(detail_columns))
        row += 1
        first = row
        for rec in part[detail_columns].itertuples(index=False):
            for j, (col, value) in enumerate(zip(detail_columns, rec), start=1):
                write_value(ws, row, j, value, FORMATS.get(col))
            row += 1
        last = row - 1
        amount_letter = get_column_letter(detail_columns.index(amount_col) + 1)
        write_value(ws, row, 2, f"{total_label} {name.upper()}", bold=True, fill=F_SUBTOTAL, border=BORDER_TOP)
        write_value(ws, row, 1, f"{len(part)} txn", bold=True, fill=F_SUBTOTAL, border=BORDER_TOP)
        total_cell = write_value(ws, row, detail_columns.index(amount_col) + 1,
                                 f"=SUM({amount_letter}{first}:{amount_letter}{last})", AUD, True, F_SUBTOTAL,
                                 BORDER_TOP)
        summary.append({"name": name, "type": ctype, "count": len(part), "total": total,
                        "cell": f"{amount_letter}{row}", "first": first, "last": last,
                        "date_letter": get_column_letter(detail_columns.index("Date") + 1)})
        _ = total_cell
        row += 2
    return summary, row


def sheet_grouped_by_person(wb, df, name, direction):
    """Income by Person-Company / Expenses by Person-Company."""
    ws = wb.create_sheet(name)
    part = df[df["Direction"] == direction]
    if direction == INCOME:
        amount_col, label, verb = "Money In", "TOTAL RECEIVED FROM", "Total Received"
        detail = ["Date", "Original Description", "Money In", "Category", "Classification", "GST Status",
                  "Business / Personal", "Notes", "Reason for Review"]
        heading = "INCOME BY PERSON / COMPANY — who paid me, when, how much"
    else:
        amount_col, label, verb = "Money Out", "TOTAL PAID TO", "Total Paid"
        detail = ["Date", "Original Description", "Money Out", "Category", "Classification", "GST Status",
                  "Business / Personal", "Tax Deductible", "Notes", "Reason for Review"]
        heading = "EXPENSES BY PERSON / COMPANY — who I paid, when, how much"
    title(ws, heading, "Summary first (highest total first), then every transaction grouped by person/company. "
                       "Original bank description kept exactly.")
    n_groups = part["Person / Company"].nunique()
    summary_header = 4
    start_sections = summary_header + n_groups + 4
    summary, _ = write_grouped(ws, part, start_sections, amount_col, detail, label, len(detail))

    headers = ["Rank", "Person / Company", "Type", "Number of Transactions", verb, "% of Total",
               "First Date", "Last Date"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=summary_header, column=j, value=h)
    style_header(ws, summary_header, 1, len(headers))
    grand_row = summary_header + n_groups + 1
    for i, s in enumerate(summary, start=1):
        r = summary_header + i
        write_value(ws, r, 1, i)
        write_value(ws, r, 2, s["name"], bold=True)
        write_value(ws, r, 3, s["type"])
        write_value(ws, r, 4, f"=COUNT({s['date_letter']}{s['first']}:{s['date_letter']}{s['last']})")
        write_value(ws, r, 5, f"={s['cell']}", AUD)
        write_value(ws, r, 6, f"=IF($E${grand_row}=0,0,E{r}/$E${grand_row})", PCT)
        write_value(ws, r, 7, f"=MIN({s['date_letter']}{s['first']}:{s['date_letter']}{s['last']})", DATE_FMT)
        write_value(ws, r, 8, f"=MAX({s['date_letter']}{s['first']}:{s['date_letter']}{s['last']})", DATE_FMT)
    write_value(ws, grand_row, 2, "GRAND TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    write_value(ws, grand_row, 4, f"=SUM(D{summary_header + 1}:D{grand_row - 1})", None, True, F_GRAND, BORDER_DOUBLE)
    write_value(ws, grand_row, 5, f"=SUM(E{summary_header + 1}:E{grand_row - 1})", AUD, True, F_GRAND, BORDER_DOUBLE)
    ws.auto_filter.ref = f"A{summary_header}:{get_column_letter(len(headers))}{grand_row - 1}"
    ws.freeze_panes = "A2"
    widths = [max(WIDTHS.get(c, 14), w) for c, w in zip(detail, [12, 60, 15, 30, 22, 15, 14, 16, 45, 45])]
    widths[0] = max(widths[0], 8)
    widths[1] = max(widths[1], 34)
    set_widths(ws, widths + [14] * max(0, len(headers) - len(widths)))
    return {"cell": f"'{name}'!$E${grand_row}", "total": part[amount_col].sum()}


def write_person_block(ws, row, name, df, width_cols=7):
    """
    For one person: MONEY RECEIVED FROM them, MONEY PAID TO them, totals and net.
    Returns next free row.
    """
    part = df[df["Person / Company"] == name].sort_values(["Date", "Row #"])
    ws.cell(row=row, column=1, value=name.upper()).font = FONT_TITLE
    row += 1
    totals = {}
    for direction, label, amount_col in ((INCOME, "MONEY RECEIVED FROM", "Money In"),
                                         (EXPENSE, "MONEY PAID TO", "Money Out")):
        sub = part[part["Direction"] == direction]
        section_title(ws, row, f"{label} {name.upper()}", width_cols)
        row += 1
        cols = ["Date", "Original Description", amount_col, "Classification", "Category", "Loan Type",
                "Reason for Review"]
        for j, col in enumerate(cols, start=1):
            ws.cell(row=row, column=j, value="Amount" if col == amount_col else col)
        style_header(ws, row, 1, len(cols))
        row += 1
        first = row
        if sub.empty:
            write_value(ws, row, 2, "No transactions found")
            row += 1
        for rec in sub[cols].itertuples(index=False):
            for j, (col, value) in enumerate(zip(cols, rec), start=1):
                write_value(ws, row, j, value, FORMATS.get(col))
            row += 1
        write_value(ws, row, 2, f"TOTAL {'RECEIVED FROM' if direction == INCOME else 'PAID TO'} {name.upper()} "
                                f"({len(sub)} transactions)", bold=True, fill=F_SUBTOTAL, border=BORDER_TOP)
        totals[direction] = write_value(ws, row, 3, f"=SUM(C{first}:C{row - 1})" if not sub.empty else 0, AUD,
                                        True, F_SUBTOTAL, BORDER_TOP).coordinate
        row += 2
    write_value(ws, row, 2, f"NET AMOUNT (received − paid) — {name}", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    write_value(ws, row, 3, f"={totals[INCOME]}-{totals[EXPENSE]}", AUD, True, F_GRAND, BORDER_DOUBLE)
    return row + 3


def sheet_people(wb, df, sheet_name, heading, people, extra_note):
    """Personal-Director and Specific People sheets: one block per person."""
    ws = wb.create_sheet(sheet_name)
    title(ws, heading, extra_note)
    row = 4
    for name in people:
        row = write_person_block(ws, row, name, df)
    set_widths(ws, [12, 62, 15, 22, 30, 30, 70])
    ws.freeze_panes = "A2"
    return ws, row


def sheet_personal_director(wb, df):
    """Director-related transactions (separately) + other personal spending by category."""
    ws, row = sheet_people(
        wb, df, "Personal-Director", "PERSONAL / DIRECTOR RELATED TRANSACTIONS",
        DIRECTOR_NAMES, "Money received from directors is NOT counted as business income, and payments to them are "
                        "NOT counted as business expenses. Review each one with your accountant.")
    personal = df[df["Classification"].isin([PER_EXP, PER_INC])]
    section_title(ws, row, "OTHER PERSONAL SPENDING (not included in business profit)", 7)
    row += 1
    cats = personal.groupby("Category")["Money Out"].agg(["count", "sum"]).sort_values("sum", ascending=False)
    for j, h in enumerate(["Category", "Transactions", "Total Paid"], start=1):
        ws.cell(row=row, column=j, value=h)
    style_header(ws, row, 1, 3)
    row += 1
    first = row
    for cat, rec in cats.iterrows():
        write_value(ws, row, 1, cat or "Personal")
        write_value(ws, row, 2, int(rec["count"]))
        write_value(ws, row, 3, float(rec["sum"]), AUD)
        row += 1
    write_value(ws, row, 1, "TOTAL PERSONAL SPENDING", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    write_value(ws, row, 2, f"=SUM(B{first}:B{row - 1})", None, True, F_GRAND, BORDER_DOUBLE)
    write_value(ws, row, 3, f"=SUM(C{first}:C{row - 1})", AUD, True, F_GRAND, BORDER_DOUBLE)
    row += 3
    section_title(ws, row, "PERSONAL SPENDING – EVERY TRANSACTION", 7)
    row += 1
    cols = ["Date", "Original Description", "Money Out", "Person / Company", "Category", "Notes"]
    for j, h in enumerate(cols, start=1):
        ws.cell(row=row, column=j, value="Amount" if h == "Money Out" else h)
    style_header(ws, row, 1, len(cols))
    row += 1
    first = row
    for rec in personal.sort_values(["Category", "Date"])[cols].itertuples(index=False):
        for j, (col, value) in enumerate(zip(cols, rec), start=1):
            write_value(ws, row, j, value, FORMATS.get(col))
        row += 1
    write_value(ws, row, 2, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    write_value(ws, row, 3, f"=SUM(C{first}:C{max(first, row - 1)})", AUD, True, F_GRAND, BORDER_DOUBLE)


def sheet_salary(wb, df):
    """Salary / wages per worker, plus possible salary payments that need review."""
    ws = wb.create_sheet("Salary-Wages")
    title(ws, "SALARY / WAGES BY WORKER",
          "Confirmed = payment to an individual with a wage keyword (payment, wage, salary, payroll). "
          "Possible salary payments are listed separately and NOT included in the totals.")
    confirmed = df[df["Classification"] == SALARY]
    possible = df[(df["Classification"] == REVIEW) & (df["Suggested Classification"] == SALARY)]
    workers = confirmed.groupby("Person / Company")["Money Out"].sum().sort_values(ascending=False)

    summary_row = 4
    headers = ["Worker", "Number of Payments", "Total Paid", "First Payment", "Last Payment",
               "Possible salary payments (review)", "Possible salary amount (review)"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=summary_row, column=j, value=h)
    style_header(ws, summary_row, 1, len(headers))
    all_workers = list(workers.index) + [n for n in possible["Person / Company"].unique() if n not in workers.index]
    row = summary_row + len(all_workers) + 4
    refs = {}
    for label, data in (("CONFIRMED SALARY / WAGES", confirmed), ("POSSIBLE SALARY – REVIEW REQUIRED "
                                                                  "(not included in totals)", possible)):
        section_title(ws, row, label, 7)
        row += 2
        order = data.groupby("Person / Company")["Money Out"].sum().sort_values(ascending=False)
        for name in order.index:
            part = data[data["Person / Company"] == name].sort_values("Date")
            ws.cell(row=row, column=1, value=name.upper()).font = FONT_SUB
            row += 1
            cols = ["Date", "Original Description", "Money Out", "Reason for Review"]
            for j, h in enumerate(cols, start=1):
                ws.cell(row=row, column=j, value="Amount" if h == "Money Out" else h)
            style_header(ws, row, 1, len(cols))
            row += 1
            first = row
            for rec in part[cols].itertuples(index=False):
                for j, (col, value) in enumerate(zip(cols, rec), start=1):
                    write_value(ws, row, j, value, FORMATS.get(col))
                row += 1
            write_value(ws, row, 1, f"=COUNT(A{first}:A{row - 1})", None, True, F_SUBTOTAL, BORDER_TOP)
            write_value(ws, row, 2, f"TOTAL PAID TO {name.upper()} (number of payments in column A)", bold=True,
                        fill=F_SUBTOTAL, border=BORDER_TOP)
            write_value(ws, row, 3, f"=SUM(C{first}:C{row - 1})", AUD, True, F_SUBTOTAL, BORDER_TOP)
            refs.setdefault(name, {})[label[:4]] = (first, row - 1, row)
            row += 2
    for i, name in enumerate(all_workers, start=1):
        r = summary_row + i
        write_value(ws, r, 1, name, bold=True)
        conf = refs.get(name, {}).get("CONF")
        poss = refs.get(name, {}).get("POSS")
        if conf:
            f, l, t = conf
            write_value(ws, r, 2, f"=A{t}")
            write_value(ws, r, 3, f"=C{t}", AUD)
            write_value(ws, r, 4, f"=MIN(A{f}:A{l})", DATE_FMT)
            write_value(ws, r, 5, f"=MAX(A{f}:A{l})", DATE_FMT)
        else:
            write_value(ws, r, 2, 0)
            write_value(ws, r, 3, 0, AUD)
        if poss:
            write_value(ws, r, 6, f"=A{poss[2]}")
            write_value(ws, r, 7, f"=C{poss[2]}", AUD)
    total_row = summary_row + len(all_workers) + 1
    write_value(ws, total_row, 1, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col, fmt in ((2, None), (3, AUD), (6, None), (7, AUD)):
        letter = get_column_letter(col)
        write_value(ws, total_row, col, f"=SUM({letter}{summary_row + 1}:{letter}{total_row - 1})", fmt, True,
                    F_GRAND, BORDER_DOUBLE)
    ws.auto_filter.ref = f"A{summary_row}:G{total_row - 1}"
    ws.freeze_panes = "A2"
    set_widths(ws, [30, 62, 15, 14, 14, 18, 18])


def loan_rows(df):
    """Transactions that are (or may be) borrowing / lending."""
    mask = (df["Loan Type"] != "") | (df["Classification"] == LOAN)
    part = df[mask].copy()
    if part.empty:
        return part.assign(Status="", **{"Amount Borrowed": 0.0, "Amount Repaid": 0.0})
    part["Status"] = part.apply(
        lambda r: "Confirmed" if r["Classification"] == LOAN and r["Review Required"] == "NO"
        else ("Loan keyword – confirm" if r["Classification"] == LOAN else "Possible – review"), axis=1)
    part["Amount Borrowed"] = part["Money In"]
    part["Amount Repaid"] = part["Money Out"]
    return part.sort_values(["Person / Company", "Date"])


def sheet_loans(wb, df):
    """LOANS / BORROWED MONEY — summary per person (SUMIFS formulas) + detail."""
    ws = wb.create_sheet("Loans")
    title(ws, "LOANS / BORROWED MONEY",
          "Money borrowed is NOT business income. 'Possible' rows are unconfirmed and need review. "
          "Buy-now-pay-later repayments show only the repayments (the amounts borrowed are not in the bank statement).")
    part = loan_rows(df)
    people = part.groupby("Person / Company")["Amount Borrowed"].sum().sort_values(ascending=False).index.tolist()
    detail_start = 4 + len(people) + 4
    cols = ["Person / Company", "Date", "Original Description", "Loan Type", "Status", "Amount Borrowed",
            "Amount Repaid", "Classification", "Reason for Review"]
    section_title(ws, detail_start - 1, "DETAIL", len(cols))
    for j, h in enumerate(cols, start=1):
        ws.cell(row=detail_start, column=j, value=h)
    style_header(ws, detail_start, 1, len(cols))
    row = detail_start + 1
    for rec in part[cols].itertuples(index=False):
        for j, (col, value) in enumerate(zip(cols, rec), start=1):
            write_value(ws, row, j, value, FORMATS.get(col))
        row += 1
    last = max(row - 1, detail_start + 1)
    rng = lambda letter: f"${letter}${detail_start + 1}:${letter}${last}"  # noqa: E731

    headers = ["Person / Company", "Amount Borrowed", "Amount Repaid", "Outstanding / Net Movement",
               "Transactions", "Status"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=4, column=j, value=h)
    style_header(ws, 4, 1, len(headers))
    for i, name in enumerate(people, start=5):
        statuses = sorted(part.loc[part["Person / Company"] == name, "Status"].unique())
        write_value(ws, i, 1, name, bold=True)
        write_value(ws, i, 2, f'=SUMIFS({rng("F")},{rng("A")},A{i})', AUD)
        write_value(ws, i, 3, f'=SUMIFS({rng("G")},{rng("A")},A{i})', AUD)
        write_value(ws, i, 4, f"=B{i}-C{i}", AUD)
        write_value(ws, i, 5, f'=COUNTIF({rng("A")},A{i})')
        write_value(ws, i, 6, ", ".join(statuses))
    total = 5 + len(people)
    write_value(ws, total, 1, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in "BCDE":
        write_value(ws, total, "ABCDE".index(col) + 1, f"=SUM({col}5:{col}{total - 1})",
                    None if col == "E" else AUD, True, F_GRAND, BORDER_DOUBLE)
    ws.auto_filter.ref = f"A{detail_start}:{get_column_letter(len(cols))}{last}"
    ws.freeze_panes = "A2"
    set_widths(ws, [30, 14, 60, 32, 22, 16, 16, 22, 70])


def label_rows(ws, start_row, rows, label_width_col=1, value_col=2):
    """Write (label, formula/value, format, style) rows. Returns {key: cell}."""
    cells = {}
    r = start_row
    for item in rows:
        if item is None:
            r += 1
            continue
        key, label, value, fmt, style = (list(item) + [None] * 5)[:5]
        write_value(ws, r, label_width_col, label, bold=style in ("bold", "total", "grand"))
        cl = write_value(ws, r, value_col, value, fmt, bold=style in ("bold", "total", "grand"),
                         fill=F_GRAND if style == "grand" else (F_SUBTOTAL if style == "total" else None),
                         border=BORDER_DOUBLE if style == "grand" else (BORDER_TOP if style == "total" else None))
        if key:
            cells[key] = cl.coordinate
        r += 1
    return cells, r


def sheet_gst(wb, ref):
    """GST SUMMARY — confirmed and unknown GST kept separate."""
    ws = wb.create_sheet("GST Summary")
    title(ws, "GST SUMMARY — ESTIMATE ONLY – ACCOUNTANT REVIEW REQUIRED",
          "GST on GST-inclusive amounts = gross ÷ 11 (not 10% of the gross). GST UNKNOWN amounts are NOT in the "
          "confirmed totals.")
    section_title(ws, 4, "INCOME (business income only)", 4)
    for j, h in enumerate(["", "Gross Amount", "GST", "Amount excl GST"], start=1):
        ws.cell(row=5, column=j, value=h)
    style_header(ws, 5, 1, 4)
    lines = [
        ("Income with confirmed GST (GST INCLUDED)", {"GST_Status": GST_INCLUDED}),
        ("Income with no GST", {"GST_Status": NO_GST}),
        ("Income with unknown GST status", {"GST_Status": GST_UNKNOWN}),
    ]
    for i, (label, crit) in enumerate(lines, start=6):
        write_value(ws, i, 1, label, bold=True)
        write_value(ws, i, 2, "=" + ref.sumifs("Amount", Classification=BUS_INC, **crit), AUD)
        write_value(ws, i, 3, "=" + ref.sumifs("GST Amount", Classification=BUS_INC, **crit), AUD)
        write_value(ws, i, 4, "=" + ref.sumifs("Amount excl GST", Classification=BUS_INC, **crit), AUD)
    write_value(ws, 9, 1, "TOTAL CONFIRMED BUSINESS INCOME", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in "BCD":
        write_value(ws, 9, "ABCD".index(col) + 1, f"=SUM({col}6:{col}8)", AUD, True, F_GRAND, BORDER_DOUBLE)

    section_title(ws, 11, "EXPENSES (confirmed deductible business expenses only)", 4)
    for j, h in enumerate(["", "Gross Expense", "GST Credit", "Expense excl GST"], start=1):
        ws.cell(row=12, column=j, value=h)
    style_header(ws, 12, 1, 4)
    write_value(ws, 13, 1, "Business expenses with GST (credit claimed)", bold=True)
    write_value(ws, 13, 2, "=-" + ref.sumifs("Amount", Classification=BUS_EXP, GST_Status=GST_INCLUDED), AUD)
    write_value(ws, 13, 3, "=" + ref.sumifs("GST Amount", Classification=BUS_EXP, GST_Status=GST_INCLUDED), AUD)
    write_value(ws, 13, 4, "=-" + ref.sumifs("Amount excl GST", Classification=BUS_EXP, GST_Status=GST_INCLUDED),
                AUD)
    write_value(ws, 14, 1, "Business expenses without GST", bold=True)
    write_value(ws, 14, 2, "=-" + ref.sumifs("Amount", Classification=BUS_EXP, GST_Status=NO_GST), AUD)
    write_value(ws, 14, 3, 0, AUD)
    write_value(ws, 14, 4, "=B14", AUD)
    write_value(ws, 15, 1, "Salary / wages (no GST credit)", bold=True)
    write_value(ws, 15, 2, "=-" + ref.sumifs("Amount", Classification=SALARY), AUD)
    write_value(ws, 15, 3, 0, AUD)
    write_value(ws, 15, 4, "=B15", AUD)
    write_value(ws, 16, 1, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in "BCD":
        write_value(ws, 16, "ABCD".index(col) + 1, f"=SUM({col}13:{col}15)", AUD, True, F_GRAND, BORDER_DOUBLE)

    section_title(ws, 18, "NET GST", 4)
    cells, _ = label_rows(ws, 19, [
        ("coll", "GST collected (confirmed)", "=C6", AUD, "bold"),
        ("cred", "GST credits (confirmed)", "=C13", AUD, "bold"),
        ("net", "ESTIMATED NET GST PAYABLE (confirmed)", "=B19-B20", AUD, "grand"),
        None,
        (None, "POTENTIAL / UNCONFIRMED (not included above)", None, None, "bold"),
        ("pot_inc", "Potential GST if unknown-status income was GST-inclusive (÷ 11)", "=B8/11", AUD),
        ("pot_cred", "Potential GST credits on expenses awaiting review",
         "=" + ref.sumifs("Potential GST", Classification=REVIEW), AUD),
        ("pot_net", "Potential net GST payable if all of the above were confirmed", "=B21+B24-B25", AUD, "total"),
        None,
        ("bas", "For reference: GST/BAS payments made to the ATO in this period",
         "=-" + ref.sumifs("Amount", Classification=GST), AUD),
    ])
    ws.cell(row=30, column=1, value="Estimated GST — preliminary only. The real BAS depends on whether each sale "
                                    "was GST-inclusive, GST-free income, input tax credits and other conditions."
            ).fill = F_NOTE
    set_widths(ws, [70, 18, 16, 18])
    ws.freeze_panes = "A2"
    return {"collected": "'GST Summary'!$B$19", "credits": "'GST Summary'!$B$20", "net": "'GST Summary'!$B$21",
            "unknown_income": "'GST Summary'!$B$8"}


def sheet_monthly(wb, df, ref):
    """MONTHLY SUMMARY — business performance month by month (SUMIFS formulas)."""
    ws = wb.create_sheet("Monthly Summary")
    headers = ["Year", "Month", "Month Name", "Business Income (gross)", "Business Income excl GST",
               "Business Expenses excl GST", "Salary / Wages", "GST Collected", "GST Credits",
               "Profit Before Tax", "Personal Income", "Personal Expenses", "Items for Review"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 1, len(headers))
    months = df.dropna(subset=["Year"])[["Year", "Month"]].drop_duplicates().sort_values(["Year", "Month"])
    row = 2
    for rec in months.itertuples(index=False):
        y, m = int(rec.Year), int(rec.Month)
        write_value(ws, row, 1, y)
        write_value(ws, row, 2, m)
        write_value(ws, row, 3, datetime(y, m, 1).strftime("%B"))
        crit = {"Year": f"A{row}", "Month": f"B{row}"}

        def sif(col, **extra):
            parts = [ref.rng(col)]
            for k, v in {**{k: v for k, v in crit.items()}, **extra}.items():
                parts += [ref.rng(criteria_column(k)), v if k in crit else f'"{v}"']
            return f"SUMIFS({','.join(parts)})"

        def cif(**extra):
            parts = []
            for k, v in {**crit, **extra}.items():
                parts += [ref.rng(criteria_column(k)), v if k in crit else f'"{v}"']
            return f"COUNTIFS({','.join(parts)})"

        write_value(ws, row, 4, "=" + sif("Amount", Classification=BUS_INC), AUD)
        write_value(ws, row, 5, "=" + sif("Amount excl GST", Classification=BUS_INC), AUD)
        write_value(ws, row, 6, "=-" + sif("Amount excl GST", Classification=BUS_EXP, Tax_Deductible="YES"), AUD)
        write_value(ws, row, 7, "=-" + sif("Amount", Classification=SALARY), AUD)
        write_value(ws, row, 8, "=" + sif("GST Amount", Classification=BUS_INC), AUD)
        write_value(ws, row, 9, "=" + sif("GST Amount", Classification=BUS_EXP, Tax_Deductible="YES"), AUD)
        write_value(ws, row, 10, f"=E{row}-F{row}-G{row}", AUD, bold=True)
        write_value(ws, row, 11, "=" + sif("Money In", Business_Personal="PERSONAL"), AUD)
        write_value(ws, row, 12, "=" + sif("Money Out", Business_Personal="PERSONAL"), AUD)
        write_value(ws, row, 13, "=" + cif(Review_Required="YES"))
        row += 1
    last = row - 1
    write_value(ws, row, 1, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in range(4, 14):
        letter = get_column_letter(col)
        write_value(ws, row, col, f"=SUM({letter}2:{letter}{last})", None if col == 13 else AUD, True, F_GRAND,
                    BORDER_DOUBLE)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last}"
    ws.conditional_formatting.add(f"J2:J{last}", CellIsRule(operator="lessThan", formula=["0"], fill=F_BAD))
    ws.conditional_formatting.add(f"J2:J{last}", CellIsRule(operator="greaterThan", formula=["0"], fill=F_OK))
    set_widths(ws, [7, 7, 12, 17, 17, 17, 15, 14, 13, 16, 15, 16, 12])


def tax_formula(profit_cell):
    """Excel formula for the income tax estimate, based on ENTITY_TYPE."""
    if ENTITY_TYPE == "individual":
        parts, prev = [], 0.0
        for threshold, rate in INDIVIDUAL_TAX_BRACKETS:
            if rate - prev:
                parts.append(f"MAX(0,{profit_cell}-{threshold})*{round(rate - prev, 4)}")
            prev = rate
        return "=" + "+".join(parts)
    rate = COMPANY_TAX_RATE if ENTITY_TYPE == "company" else CUSTOM_TAX_RATE
    return f"=MAX(0,{profit_cell})*{rate}"


def tax_explanation():
    """Plain-English description of the tax settings used."""
    if ENTITY_TYPE == "individual":
        return "Sole trader / individual – Australian resident progressive rates 2024-25 (INDIVIDUAL_TAX_BRACKETS)."
    if ENTITY_TYPE == "company":
        return (f"Company tax rate {COMPANY_TAX_RATE:.0%} (COMPANY_TAX_RATE). 25% applies to base rate entities, "
                f"otherwise 30% – confirm with your accountant.")
    return f"Custom tax rate {CUSTOM_TAX_RATE:.0%} (CUSTOM_TAX_RATE)."


def sheet_tax(wb, ref):
    """ESTIMATED TAX — kept separate from GST."""
    ws = wb.create_sheet("Tax Estimate")
    title(ws, "ESTIMATED AUSTRALIAN INCOME TAX", "ESTIMATE ONLY – ACCOUNTANT REVIEW REQUIRED")
    ws["A2"].font = Font(bold=True, color="C00000")
    cells, r = label_rows(ws, 4, [
        (None, "Entity type (ENTITY_TYPE)", ENTITY_TYPE, None, "bold"),
        (None, "Tax setting used", tax_explanation(), None),
        None,
        ("inc", "Business Income excluding GST", "=" + ref.sumifs("Amount excl GST", Classification=BUS_INC), AUD,
         "bold"),
        ("exp", "Business Expenses excluding GST (confirmed deductible)",
         "=-" + ref.sumifs("Amount excl GST", Classification=BUS_EXP, Tax_Deductible="YES"), AUD, "bold"),
        ("sal", "Salary / Wages", "=-" + ref.sumifs("Amount", Classification=SALARY, Tax_Deductible="YES"), AUD,
         "bold"),
        ("profit", "Estimated Taxable Profit", "=B7-B8-B9", AUD, "total"),
        ("tax", "Estimated Income Tax", None, AUD, "bold"),
        ("net", "Net Profit After Estimated Income Tax", "=B10-B11", AUD, "grand"),
        None,
        (None, "NOT INCLUDED ABOVE (needs review)", None, None, "bold"),
        ("rev_in", "Incoming money awaiting review (possible loans / unconfirmed income)",
         "=" + ref.sumifs("Money In", Classification=REVIEW), AUD),
        ("rev_out", "Outgoing money awaiting review (possible business expenses / salary)",
         "=" + ref.sumifs("Money Out", Classification=REVIEW), AUD),
        ("dir_in", "Money received from directors", "=" + ref.sumifs("Money In", Classification=DIRECTOR), AUD),
        ("dir_out", "Money paid to directors", "=" + ref.sumifs("Money Out", Classification=DIRECTOR), AUD),
    ])
    write_value(ws, 11, 2, tax_formula("B10"), AUD, bold=True)
    ws.cell(row=r + 1, column=1, value="Personal transactions, loans, loan repayments, internal transfers and "
                                       "ATO/GST payments are excluded from the profit calculation.").fill = F_NOTE
    ws.cell(row=r + 2, column=1, value="If business income has unknown GST status it is counted at its gross "
                                       "amount above (GST not confirmed) – see the GST Summary sheet.").fill = F_NOTE
    set_widths(ws, [72, 60])
    ws.freeze_panes = "A2"
    return {"income": "'Tax Estimate'!$B$7", "expenses": "'Tax Estimate'!$B$8", "salary": "'Tax Estimate'!$B$9",
            "profit": "'Tax Estimate'!$B$10", "tax": "'Tax Estimate'!$B$11", "net": "'Tax Estimate'!$B$12"}


def sheet_review(wb, df):
    """REVIEW REQUIRED — every transaction that needs a human decision, with the reason."""
    ws = wb.create_sheet("Review Required")
    part = df[df["Review Required"] == "YES"].copy()
    part["Income/Expense"] = part["Direction"]
    part["Suggested Category"] = part["Category"]
    cols = ["Row #", "Date", "Original Description", "Amount", "Income/Expense", "Person / Company",
            "Suggested Classification", "Suggested Category", "Reason for Review"]
    first, last = write_flat_table(ws, part, cols, total_columns=("Amount",))
    ws.conditional_formatting.add(f"D{first}:D{last}", CellIsRule(operator="lessThan", formula=["0"],
                                                                  font=Font(color="C00000")))
    # Summary by reason (to the right of the table).
    col = len(cols) + 2
    ws.cell(row=1, column=col, value="Reason for Review")
    ws.cell(row=1, column=col + 1, value="Transactions")
    style_header(ws, 1, col, col + 1)
    reason_letter = get_column_letter(cols.index("Reason for Review") + 1)
    for i, label in enumerate(REVIEW_REASON_LABELS, start=2):
        write_value(ws, i, col, label, bold=True)
        write_value(ws, i, col + 1, f'=COUNTIF(${reason_letter}${first}:${reason_letter}${last},"*{label}*")')
    write_value(ws, len(REVIEW_REASON_LABELS) + 2, col, "Transactions needing review", bold=True, fill=F_GRAND)
    write_value(ws, len(REVIEW_REASON_LABELS) + 2, col + 1, f"=COUNTA(A{first}:A{last})", None, True, F_GRAND)
    ws.cell(row=len(REVIEW_REASON_LABELS) + 3, column=col,
            value="(one transaction can have more than one reason)").font = Font(italic=True, color="7F7F7F")
    ws.column_dimensions[get_column_letter(col)].width = 38
    ws.column_dimensions[get_column_letter(col + 1)].width = 13
    return len(part)


def sheet_person_summary(wb, df, ref):
    """PERSON / COMPANY SUMMARY — totals per person/company (SUMIFS on All Transactions)."""
    ws = wb.create_sheet("Person-Company Summary")
    grouped = df.groupby("Person / Company").agg(received=("Money In", "sum"), paid=("Money Out", "sum"),
                                                 ctype=("Counterparty Type", lambda s: s.mode().iat[0]))
    grouped = grouped.sort_values(["received", "paid"], ascending=False)
    headers = ["Person / Company", "Total Received", "Total Paid", "Number of Incoming Transactions",
               "Number of Outgoing Transactions", "Net Movement", "Category"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)
    style_header(ws, 1, 1, len(headers))
    for i, (name, rec) in enumerate(grouped.iterrows(), start=2):
        write_value(ws, i, 1, name, bold=True)
        crit_name = f"{ref.rng('Person / Company')},A{i}"
        write_value(ws, i, 2, f"=SUMIFS({ref.rng('Money In')},{crit_name})", AUD)
        write_value(ws, i, 3, f"=SUMIFS({ref.rng('Money Out')},{crit_name})", AUD)
        write_value(ws, i, 4, f'=COUNTIFS({crit_name},{ref.rng("Direction")},"{INCOME}")')
        write_value(ws, i, 5, f'=COUNTIFS({crit_name},{ref.rng("Direction")},"{EXPENSE}")')
        write_value(ws, i, 6, f"=B{i}-C{i}", AUD)
        write_value(ws, i, 7, rec["ctype"])
    last = len(grouped) + 1
    total = last + 1
    write_value(ws, total, 1, "TOTAL", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in range(2, 7):
        letter = get_column_letter(col)
        write_value(ws, total, col, f"=SUM({letter}2:{letter}{last})", AUD if col in (2, 3, 6) else None, True,
                    F_GRAND, BORDER_DOUBLE)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{last}"
    ws.conditional_formatting.add(f"F2:F{last}", CellIsRule(operator="lessThan", formula=["0"],
                                                            font=Font(color="C00000")))
    set_widths(ws, [38, 16, 16, 16, 16, 16, 26])
    return {"received": f"'Person-Company Summary'!$B${total}", "paid": f"'Person-Company Summary'!$C${total}"}


def sheet_dashboard(ws, df, ref, raw, checks, gst_cells, tax_cells, grouped_cells, summary_cells, source):
    """Dashboard: validation, key business figures, GST, classification breakdown."""
    title(ws, "BUSINESS FINANCIAL ANALYSIS — DASHBOARD",
          f"Source: {source}  ·  Generated {datetime.now():%d/%m/%Y %H:%M}  ·  "
          f"Period: {df['Date'].min():%d/%m/%Y} – {df['Date'].max():%d/%m/%Y}")
    ws.cell(row=3, column=1, value="ESTIMATE ONLY – ACCOUNTANT REVIEW REQUIRED").font = Font(bold=True, color="C00000",
                                                                                               size=12)
    section_title(ws, 5, "VALIDATION — every dollar in the CSV must be in this workbook", 4)
    cells, r = label_rows(ws, 6, [
        ("csv_in", "Original CSV total incoming money", raw["in"], AUD, "bold"),
        ("pro_in", "Processed total income (All Transactions)", "=" + ref.sumifs("Money In", Direction=INCOME), AUD),
        ("grp_in", "Income by Person-Company grand total", f"={grouped_cells['in']}", AUD),
        ("sum_in", "Person-Company Summary total received", f"={summary_cells['received']}", AUD),
        ("diff_in", "INCOME DIFFERENCE", None, AUD, "grand"),
        None,
        ("csv_out", "Original CSV total outgoing money", raw["out"], AUD, "bold"),
        ("pro_out", "Processed total expenses (All Transactions)", "=" + ref.sumifs("Money Out", Direction=EXPENSE),
         AUD),
        ("grp_out", "Expenses by Person-Company grand total", f"={grouped_cells['out']}", AUD),
        ("sum_out", "Person-Company Summary total paid", f"={summary_cells['paid']}", AUD),
        ("diff_out", "EXPENSE DIFFERENCE", None, AUD, "grand"),
        None,
        ("n_csv", "TOTAL TRANSACTIONS IN ORIGINAL CSV", raw["rows"], None, "bold"),
        ("n_pro", "TOTAL TRANSACTIONS PROCESSED", f"=COUNTA({ref.rng('Row #')})", None, "bold"),
        ("n_uncl", "UNCLASSIFIED TRANSACTIONS", f"=COUNTBLANK({ref.rng('Classification')})", None, "bold"),
        ("n_dup", "DUPLICATE TRANSACTIONS (identical date, description, amount – all kept)",
         '=' + ref.countifs(Possible_Duplicate="YES"), None, "bold"),
        ("n_unread", "Rows with unreadable amount", raw["unreadable"], None),
        ("n_bal", "Running balance breaks (possible missing rows in CSV)",
         checks["balance_breaks"] if checks["balance_breaks"] is not None else "No balance column", None),
    ])
    c = cells
    ws[c["diff_in"]] = f"=ROUND({c['pro_in']}-{c['csv_in']},2)"
    ws[c["diff_out"]] = f"=ROUND({c['pro_out']}-{c['csv_out']},2)"
    status = (f'=IF(AND(ABS({c["diff_in"]})<0.005,ABS({c["diff_out"]})<0.005,{c["n_csv"]}={c["n_pro"]},'
              f'{c["n_uncl"]}=0,ABS({c["grp_in"]}-{c["csv_in"]})<0.005,ABS({c["sum_in"]}-{c["csv_in"]})<0.005,'
              f'ABS({c["grp_out"]}-{c["csv_out"]})<0.005,ABS({c["sum_out"]}-{c["csv_out"]})<0.005),'
              f'"✔ ALL TOTALS MATCH – no transaction is missing",'
              f'"⚠ WARNING – totals do not match, check the Validation figures")')
    write_value(ws, 6, 4, "VALIDATION STATUS", bold=True)
    write_value(ws, 7, 4, status, bold=True)
    ws.conditional_formatting.add("D7", FormulaRule(formula=['LEFT(D7,1)="✔"'], fill=F_OK))
    ws.conditional_formatting.add("D7", FormulaRule(formula=['LEFT(D7,1)="⚠"'], fill=F_BAD))
    for key in ("diff_in", "diff_out"):
        ws.conditional_formatting.add(c[key], CellIsRule(operator="equal", formula=["0"], fill=F_OK))
        ws.conditional_formatting.add(c[key], CellIsRule(operator="notEqual", formula=["0"], fill=F_BAD))

    section_title(ws, 26, "BUSINESS RESULT (confirmed transactions only)", 4)
    label_rows(ws, 27, [
        (None, "CONFIRMED BUSINESS INCOME (gross)", "=" + ref.sumifs("Amount", Classification=BUS_INC), AUD, "grand"),
        (None, "Business income excluding GST", f"={tax_cells['income']}", AUD),
        (None, "Allowable business expenses excluding GST", f"={tax_cells['expenses']}", AUD),
        (None, "Salary / wages", f"={tax_cells['salary']}", AUD),
        (None, "PROFIT BEFORE TAX", f"={tax_cells['profit']}", AUD, "total"),
        (None, "Estimated income tax", f"={tax_cells['tax']}", AUD),
        (None, "ESTIMATED NET PROFIT AFTER TAX", f"={tax_cells['net']}", AUD, "grand"),
    ])
    section_title(ws, 35, "GST (see GST Summary)", 4)
    label_rows(ws, 36, [
        (None, "GST Collected (confirmed)", f"={gst_cells['collected']}", AUD),
        (None, "GST Credits (confirmed)", f"={gst_cells['credits']}", AUD),
        (None, "ESTIMATED NET GST PAYABLE", f"={gst_cells['net']}", AUD, "grand"),
        (None, "Business income with UNKNOWN GST status (not in GST total)", f"={gst_cells['unknown_income']}", AUD),
    ])

    section_title(ws, 41, "ALL TRANSACTIONS BY CLASSIFICATION", 4)
    for j, h in enumerate(["Classification", "Transactions", "Money In", "Money Out"], start=1):
        ws.cell(row=42, column=j, value=h)
    style_header(ws, 42, 1, 4)
    classes = [BUS_INC, BUS_EXP, SALARY, PER_INC, PER_EXP, DIRECTOR, LOAN, TRANSFER, TAX, GST, REVIEW]
    for i, label in enumerate(classes, start=43):
        write_value(ws, i, 1, label, bold=True)
        write_value(ws, i, 2, "=" + ref.countifs(Classification=label))
        write_value(ws, i, 3, "=" + ref.sumifs("Money In", Classification=label), AUD)
        write_value(ws, i, 4, "=" + ref.sumifs("Money Out", Classification=label), AUD)
    total = 43 + len(classes)
    write_value(ws, total, 1, "TOTAL (must equal the CSV)", bold=True, fill=F_GRAND, border=BORDER_DOUBLE)
    for col in "BCD":
        write_value(ws, total, "ABCD".index(col) + 1, f"=SUM({col}43:{col}{total - 1})",
                    None if col == "B" else AUD, True, F_GRAND, BORDER_DOUBLE)
    write_value(ws, total + 2, 1, "Transactions needing review (any reason)", bold=True)
    write_value(ws, total + 2, 2, "=" + ref.countifs(Review_Required="YES"), bold=True)
    ws.cell(row=total + 4, column=1, value="Sheets: All Transactions · Income/Expense Detail · Income/Expenses by "
                                           "Person-Company · Person-Company Summary · Salary-Wages · "
                                           "Personal-Director · Specific People · Loans · GST Summary · "
                                           "Monthly Summary · Tax Estimate · Review Required").font = Font(
        italic=True, color="7F7F7F")
    set_widths(ws, [72, 20, 18, 60])
    ws.freeze_panes = "A4"


# =============================================================================
# 11. MAIN PROGRAM
# =============================================================================
def analyse(input_file=INPUT_FILE, output_file=OUTPUT_FILE):
    """Run the whole analysis and write the Excel workbook. Returns a results dict."""
    input_path = Path(input_file)
    output_path = Path(output_file).resolve()
    if not input_path.exists():
        csvs = sorted(p.name for p in Path(".").glob("*.csv"))
        raise FileNotFoundError(f"Input file '{input_file}' not found. CSV files in this folder: {csvs or 'none'}. "
                                f"Rename your file to {INPUT_FILE}, change INPUT_FILE at the top of the program, "
                                f"or run:  python financial_analysis.py your_file.csv")

    print("Reading CSV...")
    header, rows, line_numbers = read_csv_rows(input_path)
    mapping = detect_columns(header, rows)
    raw = original_csv_totals(rows, mapping)
    print(f"   {raw['rows']} transactions found in {input_path.name}")

    print("Cleaning transactions...")
    transactions = build_transactions(rows, line_numbers, mapping)

    print("Identifying income and expenses...")
    n_in = sum(tx["Direction"] == INCOME for tx in transactions)
    n_out = sum(tx["Direction"] == EXPENSE for tx in transactions)
    print(f"   {n_in} incoming, {n_out} outgoing, {len(transactions) - n_in - n_out} unreadable/zero")

    print("Grouping people and companies...")
    identify_counterparties(transactions)
    stats = payment_stats(transactions)
    for tx in transactions:
        tx["result"] = classify(tx, stats)
    flag_similar_names(transactions)
    duplicates = flag_duplicates(transactions)
    breaks = balance_check(transactions)

    print("Calculating GST...")
    df = finalise(transactions)
    print("Calculating business expenses...")

    print("Creating Excel report...")
    wb = Workbook()
    dashboard = wb.active
    dashboard.title = "Dashboard"
    ref = sheet_all_transactions(wb, df)
    sheet_detail(wb, df, "Income Detail", INCOME)
    sheet_detail(wb, df, "Expense Detail", EXPENSE)
    grouped_in = sheet_grouped_by_person(wb, df, "Income by Person-Company", INCOME)
    grouped_out = sheet_grouped_by_person(wb, df, "Expenses by Person-Company", EXPENSE)
    summary_cells = sheet_person_summary(wb, df, ref)
    sheet_salary(wb, df)
    sheet_personal_director(wb, df)
    sheet_people(wb, df, "Specific People", "SPECIFIC PEOPLE — money received FROM and paid TO each person",
                 SPECIFIC_PEOPLE, "Each person is shown separately (never combined). Net = received − paid.")
    sheet_loans(wb, df)
    gst_cells = sheet_gst(wb, ref)
    sheet_monthly(wb, df, ref)
    tax_cells = sheet_tax(wb, ref)
    n_review = sheet_review(wb, df)
    checks = {"balance_breaks": breaks}
    sheet_dashboard(dashboard, df, ref, raw, checks, gst_cells, tax_cells,
                    {"in": grouped_in["cell"], "out": grouped_out["cell"]}, summary_cells, input_path.name)

    print("Validating totals...")
    processed_in = round(df["Money In"].sum(), 2)
    processed_out = round(df["Money Out"].sum(), 2)
    results = {
        "csv_rows": raw["rows"], "processed_rows": len(df), "csv_in": raw["in"], "csv_out": raw["out"],
        "processed_in": processed_in, "processed_out": processed_out,
        "grouped_in": round(grouped_in["total"], 2), "grouped_out": round(grouped_out["total"], 2),
        "income_difference": round(processed_in - raw["in"], 2),
        "expense_difference": round(processed_out - raw["out"], 2),
        "unclassified": int((df["Classification"] == "").sum()), "duplicates": duplicates,
        "unreadable": raw["unreadable"], "balance_breaks": breaks, "review": n_review, "df": df,
        "output": output_path,
    }
    ok = (results["income_difference"] == 0 and results["expense_difference"] == 0
          and results["csv_rows"] == results["processed_rows"] and results["unclassified"] == 0
          and abs(results["grouped_in"] - raw["in"]) < 0.005 and abs(results["grouped_out"] - raw["out"]) < 0.005)
    results["ok"] = ok
    print(f"   Original CSV incoming:   ${raw['in']:,.2f}   processed: ${processed_in:,.2f}   "
          f"INCOME DIFFERENCE: ${results['income_difference']:,.2f}")
    print(f"   Original CSV outgoing:   ${raw['out']:,.2f}   processed: ${processed_out:,.2f}   "
          f"EXPENSE DIFFERENCE: ${results['expense_difference']:,.2f}")
    print(f"   Transactions in CSV: {raw['rows']}   processed: {len(df)}   unclassified: {results['unclassified']}   "
          f"duplicates (kept): {duplicates}   need review: {n_review}")
    if breaks:
        print(f"   NOTE: {breaks} running-balance breaks – the CSV itself may be missing rows.")
    if not ok:
        print("   ⚠ WARNING: TOTALS DO NOT MATCH – check the Dashboard validation section!")
    else:
        print("   ✔ All totals match – no transaction is missing.")

    try:
        wb.save(output_path)
    except PermissionError:
        raise PermissionError(f"Cannot write {output_path}. Close the file in Excel and run again.")
    print("REPORT COMPLETE")
    print(f"Excel file: {output_path}")
    return results


if __name__ == "__main__":
    arguments = sys.argv[1:]
    try:
        analyse(arguments[0] if arguments else INPUT_FILE, arguments[1] if len(arguments) > 1 else OUTPUT_FILE)
    except (FileNotFoundError, ValueError, PermissionError) as error:
        print(f"ERROR: {error}")
        sys.exit(1)
