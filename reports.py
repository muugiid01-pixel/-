"""
reports.py — Dashboard, тайлан, хүн тус бүрийн тооцоо, export-д зориулсан
DataFrame дээрх тооцооллууд.
"""

from datetime import timedelta

import pandas as pd

import config
from categorizer import name_or_alias_in_text
from tax_calculator import profit_breakdown


def prepare(df):
    """Огноог datetime болгож, дүнг float болгосон хуулбар буцаана."""
    df = df.copy()
    if df.empty:
        df["date_dt"] = pd.Series(dtype="datetime64[ns]")
        return df
    df["date_dt"] = pd.to_datetime(df["date"])
    df["amount"] = df["amount"].astype(float)
    return df


def _sum(df, finance_type=None, transaction_type=None):
    """Нөхцөлд таарсан гүйлгээний дүнгийн абсолют нийлбэр."""
    mask = pd.Series(True, index=df.index)
    if finance_type:
        mask &= df["finance_type"] == finance_type
    if transaction_type:
        mask &= df["transaction_type"] == transaction_type
    return float(df.loc[mask, "amount"].abs().sum())


def current_balance(df):
    """
    Одоогийн дансны үлдэгдэл = хамгийн сүүлийн өдрийн эцсийн үлдэгдэл.

    Нэг өдөр хэд хэдэн гүйлгээ байвал: эцсийн гүйлгээний үлдэгдэл нь өөр
    ямар ч гүйлгээний "өмнөх үлдэгдэл" (balance − amount) болдоггүй.
    """
    if df.empty or "balance" not in df or df["balance"].isna().all():
        return None
    data = prepare(df).dropna(subset=["balance"])
    last_day = data[data["date_dt"] == data["date_dt"].max()]
    if len(last_day) == 1:
        return float(last_day["balance"].iloc[0])
    previous = set((last_day["balance"] - last_day["amount"]).round(2))
    for bal in last_day["balance"].round(2):
        if bal not in previous:
            return float(bal)
    return float(last_day["balance"].iloc[0])


def summarize(df, settings):
    """
    KPI-уудыг тооцоолно. Personal гүйлгээ Business ашиг, GST, татварт орохгүй.

    Returns:
        dict: бүх KPI болон profit breakdown
    """
    df = prepare(df)
    summary = {
        "total_income": _sum(df, transaction_type=config.INCOME),
        "total_expense": _sum(df, transaction_type=config.EXPENSE),
        "business_income": _sum(df, config.BUSINESS, config.INCOME),
        "business_expense": _sum(df, config.BUSINESS, config.EXPENSE),
        "personal_income": _sum(df, config.PERSONAL, config.INCOME),
        "personal_expense": _sum(df, config.PERSONAL, config.EXPENSE),
        "transaction_count": int(len(df)),
        "current_balance": current_balance(df),
    }
    summary.update(profit_breakdown(summary["business_income"], summary["business_expense"], settings))
    return summary


def expense_by_category(df, finance_type=None):
    """Зарлагыг ангиллаар нэгтгэнэ (их → бага)."""
    df = prepare(df)
    data = df[df["transaction_type"] == config.EXPENSE]
    if finance_type:
        data = data[data["finance_type"] == finance_type]
    if data.empty:
        return pd.DataFrame(columns=["category", "amount", "count"])
    out = (data.assign(amount=data["amount"].abs())
           .groupby("category").agg(amount=("amount", "sum"), count=("id", "count"))
           .reset_index().sort_values("amount", ascending=False))
    return out


def monthly_summary(df, finance_type=None):
    """Сар бүрийн орлого/зарлага."""
    df = prepare(df)
    if finance_type:
        df = df[df["finance_type"] == finance_type]
    if df.empty:
        return pd.DataFrame(columns=["month", config.INCOME, config.EXPENSE, "net"])
    df = df.assign(month=df["date_dt"].dt.to_period("M").astype(str), abs_amount=df["amount"].abs())
    pivot = df.pivot_table(index="month", columns="transaction_type", values="abs_amount",
                           aggfunc="sum", fill_value=0).reset_index()
    for col in (config.INCOME, config.EXPENSE):
        if col not in pivot:
            pivot[col] = 0.0
    pivot["net"] = pivot[config.INCOME] - pivot[config.EXPENSE]
    return pivot[["month", config.INCOME, config.EXPENSE, "net"]].sort_values("month")


def person_transactions(df, person_name):
    """
    Тухайн хүний бүх гүйлгээ. Person талбар эсвэл тайлбарт нэр нь байвал
    (fuzzy) хамаарна.
    """
    df = prepare(df)
    if df.empty:
        return df
    by_field = df["person"].fillna("").str.lower() == person_name.lower()
    by_text = df["description"].apply(lambda d: name_or_alias_in_text(d, person_name))
    return df[by_field | by_text].sort_values("date_dt", ascending=False)


def person_summary(df, person_name):
    """
    Хүн тус бүрийн тайлан:
        Өгсөн  = тухайн хүн рүү гарсан зарлага
        Авсан  = тухайн хүнээс орж ирсэн орлого
        Зөрүү  = Авсан − Өгсөн
    """
    tx = person_transactions(df, person_name)
    given = float(tx.loc[tx["transaction_type"] == config.EXPENSE, "amount"].abs().sum()) if not tx.empty else 0.0
    received = float(tx.loc[tx["transaction_type"] == config.INCOME, "amount"].abs().sum()) if not tx.empty else 0.0
    return {
        "person": person_name,
        "given": given,
        "received": received,
        "difference": received - given,
        "count": int(len(tx)),
        "transactions": tx,
    }


def party_monthly(df, person_name):
    """Тухайн харилцагчтай хийсэн гүйлгээний сар бүрийн авсан/өгсөн дүн."""
    tx = person_transactions(df, person_name)
    if tx.empty:
        return pd.DataFrame(columns=["month", "Авсан", "Өгсөн"])
    tx = tx.assign(month=tx["date_dt"].dt.to_period("M").astype(str), abs_amount=tx["amount"].abs())
    pivot = tx.pivot_table(index="month", columns="transaction_type", values="abs_amount",
                           aggfunc="sum", fill_value=0).reset_index()
    pivot = pivot.rename(columns={config.INCOME: "Авсан", config.EXPENSE: "Өгсөн"})
    for col in ("Авсан", "Өгсөн"):
        if col not in pivot:
            pivot[col] = 0.0
    return pivot[["month", "Авсан", "Өгсөн"]].sort_values("month")


def people_overview(df, people=None):
    """Бүртгэлтэй бүх хүний товч тайлан (хүснэгт)."""
    rows = []
    for name in people or config.TRACKED_PEOPLE:
        s = person_summary(df, name)
        rows.append({"Хүн": name, "Өгсөн": s["given"], "Авсан": s["received"],
                     "Зөрүү": s["difference"], "Гүйлгээ": s["count"]})
    return pd.DataFrame(rows)


def period_range(period, reference_date, start=None, end=None):
    """
    Тайлангийн хугацааны эхлэл/төгсгөл.

    period: "Өдөр" | "7 хоног" | "Сар" | "Custom"
    """
    ref = pd.Timestamp(reference_date).normalize()
    if period == "Өдөр":
        return ref, ref
    if period == "7 хоног":
        return ref - timedelta(days=6), ref
    if period == "Сар":
        first = ref.replace(day=1)
        last = (first + pd.offsets.MonthEnd(1)).normalize()
        return first, last
    return pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize()


def filter_by_date(df, start, end):
    """Огнооны хязгаарт багтах гүйлгээнүүд (хоёр талдаа хамааруулна)."""
    df = prepare(df)
    if df.empty:
        return df
    mask = (df["date_dt"] >= pd.Timestamp(start)) & (df["date_dt"] <= pd.Timestamp(end))
    return df[mask]


def filter_transactions(df, start=None, end=None, finance_types=None, transaction_types=None,
                        categories=None, person=None, search=None):
    """Гүйлгээний хүснэгтийн filter ба тайлбараар хайлт."""
    df = prepare(df)
    if df.empty:
        return df
    if start is not None and end is not None:
        df = df[(df["date_dt"] >= pd.Timestamp(start)) & (df["date_dt"] <= pd.Timestamp(end))]
    if finance_types:
        df = df[df["finance_type"].isin(finance_types)]
    if transaction_types:
        df = df[df["transaction_type"].isin(transaction_types)]
    if categories:
        df = df[df["category"].isin(categories)]
    if person:
        df = df[df["person"].fillna("").str.lower() == person.lower()]
    if search:
        df = df[df["description"].str.contains(search, case=False, na=False, regex=False)]
    return df


def export_dataframe(df):
    """Export-д зориулсан багануудтай DataFrame (Date, Description, ...)."""
    if df.empty:
        return pd.DataFrame(columns=list(config.EXPORT_COLUMNS.values()))
    out = df.sort_values(["date", "id"])[list(config.EXPORT_COLUMNS.keys())].copy()
    out["transaction_type"] = out["transaction_type"].map(
        {config.INCOME: "Income", config.EXPENSE: "Expense"}).fillna(out["transaction_type"])
    return out.rename(columns=config.EXPORT_COLUMNS)


def export_csv_bytes(df):
    """Export CSV-г Excel-д зөв харагдах UTF-8 (BOM-той) bytes болгоно."""
    return export_dataframe(df).to_csv(index=False).encode("utf-8-sig")
