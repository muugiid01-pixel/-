"""
database.py — SQLite database-тай ажиллах бүх функц.

Хүснэгтүүд:
    transactions — гүйлгээнүүд
    settings     — GST/татварын тохиргоо (ганц мөр)
    categories   — зардлын ангиллын жагсаалт
"""

import sqlite3
from contextlib import contextmanager
from datetime import datetime

import pandas as pd

import config

TRANSACTION_COLUMNS = [
    "id", "date", "description", "amount", "transaction_type", "finance_type",
    "category", "person", "balance", "source_file", "created_at",
    "dup_key", "manual_override",
]
EDITABLE_FIELDS = {
    "date", "description", "amount", "transaction_type", "finance_type",
    "category", "person", "balance",
}


class DatabaseError(Exception):
    """Database-ийн алдааг Монгол мессежтэйгээр дамжуулна."""


@contextmanager
def get_connection(db_path=None):
    """SQLite холболт нээж, амжилттай бол commit, алдаа гарвал rollback хийнэ."""
    conn = None
    try:
        conn = sqlite3.connect(db_path or config.DB_PATH)
        conn.row_factory = sqlite3.Row
        yield conn
        conn.commit()
    except sqlite3.Error as exc:
        if conn is not None:
            conn.rollback()
        raise DatabaseError(f"{config.ERRORS['db_error']}: {exc}") from exc
    finally:
        if conn is not None:
            conn.close()


def init_db(db_path=None):
    """Хүснэгтүүдийг (байхгүй бол) үүсгэж, анхдагч утгуудыг оруулна."""
    with get_connection(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                description TEXT NOT NULL,
                amount REAL NOT NULL,
                transaction_type TEXT NOT NULL,
                finance_type TEXT NOT NULL,
                category TEXT NOT NULL,
                person TEXT DEFAULT '',
                balance REAL,
                source_file TEXT,
                created_at TEXT NOT NULL,
                dup_key TEXT NOT NULL,
                manual_override INTEGER DEFAULT 0
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_tx_dup ON transactions(dup_key)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_tx_date ON transactions(date)")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                gst_rate REAL NOT NULL,
                gst_mode TEXT NOT NULL,
                entity_type TEXT NOT NULL,
                tax_rate REAL NOT NULL,
                tax_year TEXT NOT NULL
            )
            """
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS categories (name TEXT PRIMARY KEY)"
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS uploaded_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                file_hash TEXT NOT NULL UNIQUE,
                stored_path TEXT,
                row_count INTEGER NOT NULL,
                inserted INTEGER NOT NULL,
                duplicates INTEGER NOT NULL,
                uploaded_at TEXT NOT NULL
            )
            """
        )
        d = config.DEFAULT_SETTINGS
        conn.execute(
            "INSERT OR IGNORE INTO settings (id, gst_rate, gst_mode, entity_type, tax_rate, tax_year) "
            "VALUES (1, ?, ?, ?, ?, ?)",
            (d["gst_rate"], d["gst_mode"], d["entity_type"], d["tax_rate"], d["tax_year"]),
        )
        conn.executemany(
            "INSERT OR IGNORE INTO categories (name) VALUES (?)",
            [(c,) for c in config.DEFAULT_CATEGORIES],
        )


# ---------------------------------------------------------------------------
# Transactions
# ---------------------------------------------------------------------------
def insert_transactions(df, db_path=None):
    """
    Боловсруулсан гүйлгээнүүдийг database-д хадгална.

    Duplicate шалгалт: Огноо + Тайлбар + Дүн (dup_key) ижил гүйлгээ database-д
    аль хэдийн байвал дахин оруулахгүй. Нэг өдөр яг ижил хоёр гүйлгээ
    (жишээ нь хоёр удаа ижил кофе) байж болох тул тоогоор нь харьцуулна:
    файлд 2, database-д 2 байвал 0 шинэ; файлд 2, database-д 1 байвал 1 шинэ.

    Returns:
        (inserted_count, duplicate_count)
    """
    if df.empty:
        return 0, 0
    now = datetime.now().isoformat(timespec="seconds")
    inserted = duplicates = 0

    with get_connection(db_path) as conn:
        existing = {
            row["dup_key"]: row["n"]
            for row in conn.execute(
                "SELECT dup_key, COUNT(*) AS n FROM transactions GROUP BY dup_key"
            )
        }
        seen_in_file = {}
        rows = []
        for rec in df.to_dict("records"):
            key = rec["dup_key"]
            seen_in_file[key] = seen_in_file.get(key, 0) + 1
            if seen_in_file[key] <= existing.get(key, 0):
                duplicates += 1
                continue
            balance = rec.get("balance")
            rows.append((
                rec["date"], rec["description"], float(rec["amount"]),
                rec["transaction_type"], rec["finance_type"], rec["category"],
                rec.get("person") or "",
                None if balance is None or pd.isna(balance) else float(balance),
                rec.get("source_file", ""), now, key,
            ))
        conn.executemany(
            "INSERT INTO transactions (date, description, amount, transaction_type, "
            "finance_type, category, person, balance, source_file, created_at, dup_key) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            rows,
        )
        inserted = len(rows)
    return inserted, duplicates


def get_transactions(db_path=None):
    """Бүх гүйлгээг огноогоор эрэмбэлж DataFrame хэлбэрээр буцаана."""
    with get_connection(db_path) as conn:
        df = pd.read_sql_query(
            "SELECT * FROM transactions ORDER BY date DESC, id ASC", conn
        )
    if df.empty:
        return pd.DataFrame(columns=TRANSACTION_COLUMNS)
    df["person"] = df["person"].fillna("")
    return df


def update_transaction(tx_id, fields, db_path=None):
    """
    Нэг гүйлгээг гараар засна (manual correction).

    Орлого ↔ Зарлага сольсон бол дүнгийн тэмдгийг автоматаар тааруулна
    (Орлого = эерэг, Зарлага = сөрөг) — ингэснээр Priority 1 дүрэм хадгалагдана.
    """
    fields = {k: v for k, v in fields.items() if k in EDITABLE_FIELDS}
    if not fields:
        return
    if "amount" in fields or "transaction_type" in fields:
        with get_connection(db_path) as conn:
            row = conn.execute(
                "SELECT amount, transaction_type FROM transactions WHERE id = ?", (tx_id,)
            ).fetchone()
        if row is not None:
            amount = float(fields.get("amount", row["amount"]))
            if "transaction_type" in fields:
                amount = abs(amount) if fields["transaction_type"] == config.INCOME else -abs(amount)
            else:
                fields["transaction_type"] = config.EXPENSE if amount < 0 else config.INCOME
            fields["amount"] = amount

    assignments = ", ".join(f"{k} = ?" for k in fields)
    with get_connection(db_path) as conn:
        conn.execute(
            f"UPDATE transactions SET {assignments}, manual_override = 1 WHERE id = ?",
            [*fields.values(), tx_id],
        )


def delete_transactions(ids, db_path=None):
    """Сонгосон гүйлгээнүүдийг устгана."""
    ids = [int(i) for i in ids]
    if not ids:
        return 0
    placeholders = ",".join("?" for _ in ids)
    with get_connection(db_path) as conn:
        cur = conn.execute(f"DELETE FROM transactions WHERE id IN ({placeholders})", ids)
        return cur.rowcount


def delete_all_transactions(db_path=None):
    """Бүх гүйлгээг устгана (Settings → Database цэвэрлэх)."""
    with get_connection(db_path) as conn:
        conn.execute("DELETE FROM transactions")
        conn.execute("DELETE FROM uploaded_files")


def reclassify_transactions(classify_fn, db_path=None):
    """
    Гараар засаагүй (manual_override = 0) бүх гүйлгээг шинэ дүрмээр дахин ангилна.

    Returns:
        int: шинэчлэгдсэн мөрийн тоо
    """
    with get_connection(db_path) as conn:
        rows = conn.execute(
            "SELECT id, description, amount FROM transactions WHERE manual_override = 0"
        ).fetchall()
        updates = []
        for row in rows:
            r = classify_fn(row["description"], row["amount"])
            updates.append((r["transaction_type"], r["finance_type"], r["category"],
                            r["person"], row["id"]))
        conn.executemany(
            "UPDATE transactions SET transaction_type = ?, finance_type = ?, category = ?, "
            "person = ? WHERE id = ?",
            updates,
        )
    return len(updates)


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
def get_settings(db_path=None):
    """Тохиргоог dict хэлбэрээр буцаана."""
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT * FROM settings WHERE id = 1").fetchone()
    if row is None:
        return dict(config.DEFAULT_SETTINGS)
    return {k: row[k] for k in ("gst_rate", "gst_mode", "entity_type", "tax_rate", "tax_year")}


def save_settings(settings, db_path=None):
    """Тохиргоог хадгална."""
    merged = {**get_settings(db_path), **settings}
    with get_connection(db_path) as conn:
        conn.execute(
            "UPDATE settings SET gst_rate = ?, gst_mode = ?, entity_type = ?, tax_rate = ?, "
            "tax_year = ? WHERE id = 1",
            (float(merged["gst_rate"]), merged["gst_mode"], merged["entity_type"],
             float(merged["tax_rate"]), merged["tax_year"]),
        )


# ---------------------------------------------------------------------------
# Categories
# ---------------------------------------------------------------------------
def get_categories(db_path=None):
    """Ангиллын жагсаалт (анхдагч дараалал + шинээр нэмсэн нь)."""
    with get_connection(db_path) as conn:
        names = [r["name"] for r in conn.execute("SELECT name FROM categories")]
    default = [c for c in config.DEFAULT_CATEGORIES if c in names]
    custom = sorted(c for c in names if c not in config.DEFAULT_CATEGORIES)
    return default + custom


def add_category(name, db_path=None):
    """Шинэ ангилал нэмнэ. Хоосон эсвэл давхардсан бол False буцаана."""
    name = (name or "").strip()
    if not name:
        return False
    with get_connection(db_path) as conn:
        cur = conn.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (name,))
        return cur.rowcount > 0


def delete_category(name, db_path=None):
    """
    Ангилал устгана. Тухайн ангилалтай гүйлгээнүүд "Бусад" болно.
    Системийн үндсэн ангиллыг устгахгүй.
    """
    protected = {config.OTHER_CATEGORY, config.PERSONAL_CATEGORY,
                 config.INCOME_CATEGORY, config.SALARY_CATEGORY}
    if name in protected:
        return False
    with get_connection(db_path) as conn:
        conn.execute("UPDATE transactions SET category = ? WHERE category = ?",
                     (config.OTHER_CATEGORY, name))
        cur = conn.execute("DELETE FROM categories WHERE name = ?", (name,))
        return cur.rowcount > 0


# ---------------------------------------------------------------------------
# Uploaded files — нэг удаа upload хийсэн файлыг бүртгэж, дахин оруулахгүй
# ---------------------------------------------------------------------------
def find_uploaded_file(file_hash, db_path=None):
    """Ижил агуулгатай (hash) файл өмнө нь хадгалагдсан бол түүний мэдээллийг буцаана."""
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT * FROM uploaded_files WHERE file_hash = ?", (file_hash,)).fetchone()
    return dict(row) if row else None


def record_uploaded_file(file_name, file_hash, stored_path, row_count, inserted, duplicates,
                         db_path=None):
    """Upload хийсэн файлыг бүртгэнэ."""
    with get_connection(db_path) as conn:
        conn.execute(
            "INSERT OR IGNORE INTO uploaded_files (file_name, file_hash, stored_path, row_count, "
            "inserted, duplicates, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (file_name, file_hash, stored_path, int(row_count), int(inserted), int(duplicates),
             datetime.now().isoformat(timespec="seconds")),
        )


def get_uploaded_files(db_path=None):
    """Хадгалагдсан файлуудын жагсаалт (шинэ нь эхэндээ)."""
    with get_connection(db_path) as conn:
        return pd.read_sql_query("SELECT * FROM uploaded_files ORDER BY uploaded_at DESC", conn)


def delete_uploaded_file(file_id, db_path=None):
    """
    Файлын бүртгэл болон тухайн файлаас орсон бүх гүйлгээг устгана.

    Returns:
        (deleted_transactions, stored_path)
    """
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT file_name, stored_path FROM uploaded_files WHERE id = ?",
                           (int(file_id),)).fetchone()
        if row is None:
            return 0, None
        cur = conn.execute("DELETE FROM transactions WHERE source_file = ?", (row["file_name"],))
        conn.execute("DELETE FROM uploaded_files WHERE id = ?", (int(file_id),))
        return cur.rowcount, row["stored_path"]


# ---------------------------------------------------------------------------
# Database viewer — бүх хүснэгтийг харах
# ---------------------------------------------------------------------------
DB_TABLES = {
    "transactions": "Гүйлгээ",
    "uploaded_files": "Upload хийсэн файлууд",
    "settings": "Тохиргоо",
    "categories": "Ангилал",
}


def read_table(table, db_path=None):
    """Нэг хүснэгтийн бүх мөр, бүх баганыг DataFrame хэлбэрээр буцаана (зөвхөн DB_TABLES)."""
    if table not in DB_TABLES:
        raise DatabaseError(f"{config.ERRORS['db_error']}: '{table}' хүснэгт байхгүй")
    order = {"transactions": " ORDER BY date DESC, id ASC",
             "uploaded_files": " ORDER BY uploaded_at DESC"}.get(table, "")
    with get_connection(db_path) as conn:
        return pd.read_sql_query(f"SELECT * FROM {table}{order}", conn)


def table_counts(db_path=None):
    """Хүснэгт бүрийн мөрийн тоо."""
    with get_connection(db_path) as conn:
        return {t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0] for t in DB_TABLES}
