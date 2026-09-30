"""
app.py — "Орлого, Зарлагын Хяналтын Систем" Streamlit web application.

Ажиллуулах:  streamlit run app.py
"""

import hashlib
import os
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import config
import database as db
import reports
from categorizer import classify_transaction
from data_processor import CSVProcessingError, process_csv

C = config.COLORS
PAGES = [
    "Dashboard", "CSV Upload", "Гүйлгээ", "Business", "Personal",
    "People", "Тайлан", "Tax Estimate", "Export", "Database", "Settings",
]
PAGE_ICONS = {
    "Dashboard": "📊", "CSV Upload": "📤", "Гүйлгээ": "📋", "Business": "💼",
    "Personal": "👤", "People": "👥", "Тайлан": "🗓️", "Tax Estimate": "🧾",
    "Export": "⬇️", "Database": "🗄️", "Settings": "⚙️",
}
UPLOAD_DIR = "uploads"
GST_NOTE = "Estimated GST — зөвхөн урьдчилсан тооцоо"
TAX_NOTE = "Estimated Australian Tax — татварын албан ёсны зөвлөгөө биш, урьдчилсан тооцоо."


# ---------------------------------------------------------------------------
# Дизайн
# ---------------------------------------------------------------------------
def inject_css():
    """Corporate dashboard загварын CSS."""
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: {C['light_grey']}; }}
        section[data-testid="stSidebar"] {{
            background-color: {C['white']};
            border-right: 1px solid {C['border']};
        }}
        h1, h2, h3 {{ color: {C['text']}; }}
        .app-header {{
            background: linear-gradient(90deg, {C['primary']} 0%, {C['primary_dark']} 100%);
            color: white; padding: 18px 24px; border-radius: 12px; margin-bottom: 18px;
        }}
        .app-header h1 {{ color: white; margin: 0; font-size: 1.6rem; }}
        .app-header p {{ color: #DCE7FB; margin: 4px 0 0 0; font-size: 0.95rem; }}
        .kpi-card {{
            background: {C['white']}; border: 1px solid {C['border']};
            border-left: 5px solid var(--accent); border-radius: 10px;
            padding: 14px 16px; margin-bottom: 12px;
            box-shadow: 0 1px 3px rgba(16, 24, 40, 0.06); min-height: 96px;
        }}
        .kpi-label {{ color: {C['muted']}; font-size: 0.82rem; font-weight: 600;
                      text-transform: uppercase; letter-spacing: .02em; }}
        .kpi-value {{ color: {C['text']}; font-size: 1.45rem; font-weight: 700; margin-top: 4px; }}
        .kpi-sub {{ color: {C['muted']}; font-size: 0.78rem; margin-top: 2px; }}
        .note-box {{
            background: #FFF7E6; border: 1px solid #FCD9A0; color: #8A5300;
            padding: 10px 14px; border-radius: 8px; font-size: 0.9rem; margin: 6px 0 14px 0;
        }}
        .section-card {{
            background: {C['white']}; border: 1px solid {C['border']};
            border-radius: 10px; padding: 16px 18px; margin-bottom: 14px;
        }}
        .calc-row {{ display: flex; justify-content: space-between; padding: 7px 0;
                     border-bottom: 1px dashed {C['border']}; font-size: 0.98rem; }}
        .calc-row.total {{ font-weight: 700; border-bottom: 2px solid {C['text']}; }}
        .calc-row.final {{ font-weight: 800; color: {C['primary_dark']}; border-bottom: none;
                           font-size: 1.08rem; }}
        .badge {{ display: inline-block; padding: 2px 10px; border-radius: 999px;
                  font-size: 0.78rem; font-weight: 600; color: white; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def page_header(title, subtitle=""):
    """Хуудасны толгой хэсэг."""
    st.markdown(
        f'<div class="app-header"><h1>{title}</h1>'
        f'{f"<p>{subtitle}</p>" if subtitle else ""}</div>',
        unsafe_allow_html=True,
    )


def money(value):
    """Мөнгөн дүнг "$1,234.50" / "-$1,234.50" хэлбэрээр форматлана."""
    if value is None or pd.isna(value):
        return "—"
    sign = "-" if value < 0 else ""
    return f"{sign}{config.CURRENCY_SYMBOL}{abs(value):,.2f}"


def kpi_card(label, value, accent, sub=""):
    """Нэг KPI card-ийн HTML."""
    return (
        f'<div class="kpi-card" style="--accent:{accent}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'<div class="kpi-sub">{sub}</div></div>'
    )


def kpi_grid(cards, per_row=4):
    """KPI card-уудыг мөр мөрөөр байрлуулна."""
    for i in range(0, len(cards), per_row):
        cols = st.columns(per_row)
        for col, card in zip(cols, cards[i:i + per_row]):
            col.markdown(kpi_card(*card), unsafe_allow_html=True)


def note(text):
    """Анхааруулах тэмдэглэл."""
    st.markdown(f'<div class="note-box">⚠️ {text}</div>', unsafe_allow_html=True)


def style_fig(fig, height=360):
    """Plotly графикт нэгдсэн загвар өгнө."""
    fig.update_layout(
        template="plotly_white", height=height, margin=dict(l=10, r=10, t=50, b=10),
        font=dict(family="Inter, Segoe UI, sans-serif", color=C["text"]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=None),
        paper_bgcolor=C["white"], plot_bgcolor=C["white"],
    )
    return fig


def empty_state():
    """Гүйлгээ байхгүй үеийн мессеж."""
    st.info("Одоогоор гүйлгээ алга. Зүүн цэсний **CSV Upload** хэсгээс банкны хуулгаа оруулна уу.")


def amount_color(val):
    """Хүснэгтэд эерэг дүнг ногоон, сөрөг дүнг улаанаар харуулна."""
    if isinstance(val, (int, float)):
        return f"color: {C['income'] if val >= 0 else C['expense']}; font-weight: 600"
    return ""


def display_table(df, height=420):
    """Гүйлгээний хүснэгтийг Монгол багана, өнгөт дүнтэйгээр харуулна."""
    if df.empty:
        st.caption("Гүйлгээ олдсонгүй.")
        return
    view = df[["date", "description", "finance_type", "transaction_type", "category", "person", "amount"]]
    view = view.rename(columns={
        "date": "Огноо", "description": "Тайлбар", "finance_type": "Business/Personal",
        "transaction_type": "Төрөл", "category": "Ангилал", "person": "Хүн", "amount": "Дүн",
    })
    styled = view.style.map(amount_color, subset=["Дүн"]).format({"Дүн": "{:,.2f}"})
    st.dataframe(styled, width="stretch", height=height, hide_index=True)


# ---------------------------------------------------------------------------
# Графикууд
# ---------------------------------------------------------------------------
def chart_income_vs_expense(income, expense, title):
    """Орлого ба зарлагын баганан график."""
    fig = go.Figure(go.Bar(
        x=[config.INCOME, config.EXPENSE], y=[income, expense],
        marker_color=[C["income"], C["expense"]],
        text=[money(income), money(expense)], textposition="outside",
    ))
    fig.update_layout(title=title, showlegend=False)
    return style_fig(fig)


def chart_business_vs_personal(s):
    """Business ба Personal-ийн орлого/зарлагыг харьцуулна."""
    fig = go.Figure()
    fig.add_bar(name=config.BUSINESS, x=[config.INCOME, config.EXPENSE],
                y=[s["business_income"], s["business_expense"]], marker_color=C["business"])
    fig.add_bar(name=config.PERSONAL, x=[config.INCOME, config.EXPENSE],
                y=[s["personal_income"], s["personal_expense"]], marker_color=C["personal"])
    fig.update_layout(title="Business vs Personal", barmode="group")
    return style_fig(fig)


def chart_categories(df, finance_type=None, title="Зарлагын ангилал"):
    """Зарлагын ангиллын donut график."""
    data = reports.expense_by_category(df, finance_type)
    if data.empty:
        return None
    fig = px.pie(data, names="category", values="amount", hole=0.5, title=title,
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_traces(textposition="inside", textinfo="percent")
    return style_fig(fig, 400)


def chart_monthly(df, finance_type=None, title="Сар бүрийн орлого/зарлага"):
    """Сар бүрийн орлого, зарлага, цэвэр дүн."""
    data = reports.monthly_summary(df, finance_type)
    if data.empty:
        return None
    fig = go.Figure()
    fig.add_bar(name=config.INCOME, x=data["month"], y=data[config.INCOME], marker_color=C["income"])
    fig.add_bar(name=config.EXPENSE, x=data["month"], y=data[config.EXPENSE], marker_color=C["expense"])
    fig.add_scatter(name="Цэвэр", x=data["month"], y=data["net"], mode="lines+markers",
                    line=dict(color=C["primary"], width=3))
    fig.update_layout(title=title, barmode="group")
    return style_fig(fig, 400)


def chart_people(df):
    """Бүртгэлтэй хүмүүст өгсөн/авсан мөнгө."""
    data = reports.people_overview(df)
    if data.empty or (data["Өгсөн"].sum() == 0 and data["Авсан"].sum() == 0):
        return None
    fig = go.Figure()
    fig.add_bar(name="Өгсөн", x=data["Хүн"], y=data["Өгсөн"], marker_color=C["expense"])
    fig.add_bar(name="Авсан", x=data["Хүн"], y=data["Авсан"], marker_color=C["income"])
    fig.update_layout(title="Person payments", barmode="group")
    return style_fig(fig)


def chart_profit_waterfall(s):
    """Цэвэр ашгийн задаргааны waterfall график."""
    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "total", "relative", "total"],
        x=["Business Revenue", "Business Expenses", "Estimated GST",
           "Profit Before Tax", "Estimated Tax", "Net Profit"],
        y=[s["business_revenue"], -s["business_expenses"], -s["estimated_gst"],
           0, -s["estimated_tax"], 0],
        text=[money(s["business_revenue"]), money(-s["business_expenses"]),
              money(-s["estimated_gst"]), money(s["profit_before_tax"]),
              money(-s["estimated_tax"]), money(s["net_profit"])],
        textposition="outside",
        increasing=dict(marker=dict(color=C["income"])),
        decreasing=dict(marker=dict(color=C["expense"])),
        totals=dict(marker=dict(color=C["primary"])),
        connector=dict(line=dict(color=C["border"])),
    ))
    fig.update_layout(title="Profit breakdown", showlegend=False)
    return style_fig(fig, 420)


def show_chart(fig):
    """График байвал харуулна, үгүй бол мэдээлэл алга гэж бичнэ."""
    if fig is None:
        st.caption("Графикт харуулах мэдээлэл алга.")
    else:
        st.plotly_chart(fig, use_container_width=True)


def profit_calc_block(s):
    """Цэвэр ашгийн тооцооны дарааллыг текстээр харуулна."""
    rows = [
        ("Business Revenue", money(s["business_revenue"]), ""),
        ("− Business Expenses", money(-s["business_expenses"]), ""),
        ("− Estimated GST", money(-s["estimated_gst"]), ""),
        ("= Profit Before Income Tax", money(s["profit_before_tax"]), "total"),
        ("− Estimated Australian Tax", money(-s["estimated_tax"]), ""),
        ("= Estimated Net Profit", money(s["net_profit"]), "final"),
    ]
    html = "".join(f'<div class="calc-row {cls}"><span>{label}</span><span>{val}</span></div>'
                   for label, val, cls in rows)
    st.markdown(f'<div class="section-card">{html}</div>', unsafe_allow_html=True)
    st.caption("Personal гүйлгээ энэ тооцоонд орохгүй.")


# ---------------------------------------------------------------------------
# Хуудсууд
# ---------------------------------------------------------------------------
def page_dashboard(df, settings):
    """Dashboard — KPI, GST, цэвэр ашиг, графикууд."""
    page_header("Dashboard", "Орлого, зарлага, GST болон татварын ерөнхий зураглал")
    if df.empty:
        empty_state()
        return
    s = reports.summarize(df, settings)

    kpi_grid([
        ("Нийт орлого", money(s["total_income"]), C["income"]),
        ("Нийт зарлага", money(s["total_expense"]), C["expense"]),
        ("Business орлого", money(s["business_income"]), C["business"]),
        ("Business зарлага", money(s["business_expense"]), C["business"]),
        ("Personal орлого", money(s["personal_income"]), C["personal"]),
        ("Personal зарлага", money(s["personal_expense"]), C["personal"]),
        ("Нийт гүйлгээ", f"{s['transaction_count']:,}", C["primary"]),
        ("Одоогийн дансны үлдэгдэл", money(s["current_balance"]), C["primary"]),
        ("Estimated GST", money(s["estimated_gst"]), C["gst"], "Урьдчилсан тооцоо"),
        ("Estimated Tax", money(s["estimated_tax"]), C["tax"], settings["entity_type"]),
        ("Estimated Net Profit", money(s["net_profit"]), C["profit"], "Business-ийн цэвэр ашиг"),
    ])

    st.subheader("GST")
    note(GST_NOTE)
    kpi_grid([
        ("Business Revenue", money(s["business_revenue"]), C["business"]),
        ("Estimated GST", money(s["estimated_gst"]), C["gst"], config.GST_MODES[settings["gst_mode"]]),
        ("Revenue after estimated GST", money(s["revenue_after_gst"]), C["income"]),
    ], per_row=3)

    st.subheader("Цэвэр ашгийн тооцоо")
    note(TAX_NOTE)
    left, right = st.columns([1, 1.3])
    with left:
        profit_calc_block(s)
    with right:
        show_chart(chart_profit_waterfall(s))

    st.subheader("Графикууд")
    c1, c2 = st.columns(2)
    with c1:
        show_chart(chart_income_vs_expense(s["business_income"], s["business_expense"],
                                           "Business Income vs Expense"))
    with c2:
        show_chart(chart_income_vs_expense(s["personal_income"], s["personal_expense"],
                                           "Personal Income vs Expense"))
    c3, c4 = st.columns(2)
    with c3:
        show_chart(chart_business_vs_personal(s))
    with c4:
        show_chart(chart_categories(df))
    show_chart(chart_monthly(df))
    show_chart(chart_people(df))


def save_uploaded_csv(raw_bytes, file_name):
    """
    CSV-г боловсруулж database-д шууд хадгална, эх файлын хуулбарыг uploads/ хавтаст үлдээнэ.
    Ижил агуулгатай файл өмнө нь хадгалагдсан бол дахин боловсруулахгүй.

    Returns:
        dict: үр дүн (status = "saved" | "exists")
    """
    file_hash = hashlib.sha256(raw_bytes).hexdigest()
    existing = db.find_uploaded_file(file_hash)
    if existing:
        return {"status": "exists", "name": file_name, "info": existing}

    processed, mapping, warnings = process_csv(raw_bytes, file_name)
    inserted, duplicates = db.insert_transactions(processed)

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    stored_path = os.path.join(UPLOAD_DIR, f"{datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(file_name)}")
    with open(stored_path, "wb") as fh:
        fh.write(raw_bytes)
    db.record_uploaded_file(file_name, file_hash, stored_path, len(processed), inserted, duplicates)
    return {"status": "saved", "name": file_name, "rows": len(processed), "inserted": inserted,
            "duplicates": duplicates, "mapping": mapping, "warnings": warnings}


def show_upload_result(result, df):
    """Сүүлд upload хийсэн файлын үр дүнг харуулна."""
    if result["status"] == "exists":
        info = result["info"]
        st.info(f"ℹ️ **{result['name']}** файл {info['uploaded_at'].replace('T', ' ')}-нд аль хэдийн "
                f"хадгалагдсан байна ({info['inserted']:,} гүйлгээ). Дахин оруулсангүй.")
        return
    st.success(f"✅ **{result['name']}** — {result['rows']:,} гүйлгээ уншиж, "
               f"{result['inserted']:,} шинэ гүйлгээ database-д хадгаллаа.")
    if result["duplicates"]:
        st.warning(f"⚠️ {config.ERRORS['duplicate']} ({result['duplicates']:,} гүйлгээ)")
    labels = {"date": "Огноо", "description": "Тайлбар", "amount": "Дүн", "income": "Орлого",
              "expense": "Зарлага", "balance": "Үлдэгдэл"}
    found = ", ".join(f"{labels.get(k, k)} → `{v}`" for k, v in result["mapping"].items())
    st.caption(f"Танигдсан баганууд: {found}")
    for w in result["warnings"]:
        st.warning(w)
    new_rows = df[df["source_file"] == result["name"]]
    if not new_rows.empty:
        with st.expander("Хадгалсан гүйлгээнүүд (автомат ангилал)"):
            display_table(reports.prepare(new_rows), height=360)


def page_upload(df):
    """CSV Upload — нэг удаа upload хийхэд database-д шууд хадгална."""
    page_header("CSV Upload", "Банкны хуулгыг нэг удаа оруулахад database-д хадгалагдана")

    result = st.session_state.pop("upload_result", None)
    if result:
        show_upload_result(result, df)

    files = db.get_uploaded_files()
    if not df.empty:
        st.markdown(
            f'<div class="section-card">💾 Database-д <b>{len(df):,}</b> гүйлгээ, '
            f'<b>{len(files)}</b> файл хадгалагдсан байна. Програм нээх бүрт автоматаар ачаалагдана — '
            f'<b>дахин upload хийх шаардлагагүй</b>. Засвараа <b>Гүйлгээ</b> хуудсанд хийнэ.</div>',
            unsafe_allow_html=True,
        )

    st.subheader("Нэмэлт CSV оруулах" if not df.empty else "CSV оруулах")
    st.caption("Огноо, Тайлбар, Дүн (эсвэл Орлого/Зарлага), Дансны үлдэгдэл баганыг автоматаар таньна. "
               "Header-гүй CommBank CSV-г ч таньна. `-` тэмдэгтэй дүн = Зарлага, эерэг = Орлого. "
               "Өмнө орсон гүйлгээг давхар оруулахгүй.")
    key = st.session_state.setdefault("uploader_key", 0)
    col1, col2 = st.columns([3, 1])
    with col1:
        uploaded = st.file_uploader("CSV файл сонгох", type=["csv"], key=f"uploader_{key}")
    with col2:
        st.write("")
        st.write("")
        use_sample = st.button("Sample CSV ашиглах", width="stretch")

    raw, name = None, None
    if uploaded is not None:
        raw, name = uploaded.getvalue(), uploaded.name
    elif use_sample:
        try:
            with open("sample_transactions.csv", "rb") as fh:
                raw, name = fh.read(), "sample_transactions.csv"
        except FileNotFoundError:
            st.error("❌ sample_transactions.csv файл олдсонгүй.")

    if raw is not None:
        try:
            with st.spinner("Боловсруулж, хадгалж байна…"):
                st.session_state["upload_result"] = save_uploaded_csv(raw, name)
        except CSVProcessingError as exc:
            st.error(f"❌ {exc}")
            return
        except (db.DatabaseError, OSError) as exc:
            st.error(f"❌ {config.ERRORS['db_error']}: {exc}")
            return
        st.session_state["uploader_key"] = key + 1
        st.rerun()

    if not files.empty:
        st.subheader("Хадгалагдсан файлууд")
        view = files[["file_name", "uploaded_at", "row_count", "inserted", "duplicates"]].rename(columns={
            "file_name": "Файл", "uploaded_at": "Огноо", "row_count": "Мөр",
            "inserted": "Хадгалсан", "duplicates": "Давхардсан"})
        view["Огноо"] = view["Огноо"].str.replace("T", " ")
        st.dataframe(view, width="stretch", hide_index=True)
        with st.expander("🗑️ Файл устгах (тухайн файлаас орсон бүх гүйлгээ устана)"):
            labels = {r.id: f"{r.file_name} — {r.uploaded_at.replace('T', ' ')}" for r in files.itertuples()}
            fid = st.selectbox("Файл", list(labels), format_func=labels.get)
            if st.button("Устгах", key="delete_file"):
                try:
                    n, path = db.delete_uploaded_file(fid)
                except db.DatabaseError as exc:
                    st.error(f"❌ {exc}")
                    return
                if path and os.path.exists(path):
                    os.remove(path)
                st.success(f"✅ Файл болон {n:,} гүйлгээ устгагдлаа.")
                st.rerun()


def page_transactions(df, categories):
    """Гүйлгээ — filter, хайлт, гараар засах, устгах."""
    page_header("Гүйлгээ", "Шүүх, хайх, ангилал засах, устгах")
    if df.empty:
        empty_state()
        return

    data = reports.prepare(df)
    min_d, max_d = data["date_dt"].min().date(), data["date_dt"].max().date()

    with st.container(border=True):
        f1, f2, f3 = st.columns([1.4, 1, 1])
        with f1:
            dr = st.date_input("Огноо", value=(min_d, max_d), min_value=min_d, max_value=max_d)
        with f2:
            fin = st.multiselect("Business / Personal", [config.BUSINESS, config.PERSONAL])
        with f3:
            typ = st.multiselect("Орлого / Зарлага", [config.INCOME, config.EXPENSE])
        f4, f5, f6 = st.columns([1.4, 1, 1])
        with f4:
            search = st.text_input("🔍 Тайлбараар хайх")
        with f5:
            cats = st.multiselect("Ангилал", sorted(data["category"].unique()))
        with f6:
            people = sorted(p for p in data["person"].unique() if p)
            person = st.selectbox("Хүний нэр", ["Бүгд"] + people)

    start, end = (dr if isinstance(dr, tuple) and len(dr) == 2 else (min_d, max_d))
    filtered = reports.filter_transactions(
        df, start, end, fin, typ, cats, None if person == "Бүгд" else person, search.strip()
    )

    inc = filtered.loc[filtered["transaction_type"] == config.INCOME, "amount"].abs().sum()
    exp = filtered.loc[filtered["transaction_type"] == config.EXPENSE, "amount"].abs().sum()
    kpi_grid([
        ("Олдсон гүйлгээ", f"{len(filtered):,}", C["primary"]),
        ("Орлого", money(inc), C["income"]),
        ("Зарлага", money(exp), C["expense"]),
        ("Зөрүү", money(inc - exp), C["profit"]),
    ])

    st.markdown("**Засах:** нүдэн дээр дараад өөрчилнө. Устгах гүйлгээгээ **Устгах** баганаар "
                "тэмдэглээд **Хадгалах** товч дарна. Төрлийг сольвол дүнгийн тэмдэг автоматаар солигдоно.")

    editor_df = filtered[["id", "date_dt", "description", "finance_type", "transaction_type",
                          "category", "person", "amount"]].copy()
    editor_df["date_dt"] = editor_df["date_dt"].dt.date
    editor_df.insert(0, "delete", False)
    editor_df = editor_df.set_index("id")
    cat_options = sorted(set(categories) | set(data["category"].unique()))

    edited = st.data_editor(
        editor_df,
        key=f"editor_{start}_{end}_{fin}_{typ}_{cats}_{person}_{search}",
        width="stretch", height=480, hide_index=True, num_rows="fixed",
        column_config={
            "delete": st.column_config.CheckboxColumn("Устгах", width="small"),
            "date_dt": st.column_config.DateColumn("Огноо", format="YYYY-MM-DD"),
            "description": st.column_config.TextColumn("Тайлбар", width="large"),
            "finance_type": st.column_config.SelectboxColumn(
                "Business/Personal", options=[config.BUSINESS, config.PERSONAL], required=True),
            "transaction_type": st.column_config.SelectboxColumn(
                "Төрөл", options=[config.INCOME, config.EXPENSE], required=True),
            "category": st.column_config.SelectboxColumn("Ангилал", options=cat_options, required=True),
            "person": st.column_config.TextColumn("Хүн"),
            "amount": st.column_config.NumberColumn("Дүн", format="%.2f"),
        },
    )

    b1, b2, _ = st.columns([1, 1.3, 3])
    if b1.button("💾 Хадгалах", type="primary", width="stretch"):
        save_edits(editor_df, edited)
    if b2.button("🔄 Автоматаар дахин ангилах", width="stretch",
                 help="Гараар засаагүй гүйлгээнүүдийг дүрмээр дахин ангилна."):
        try:
            n = db.reclassify_transactions(classify_transaction)
            st.success(f"✅ {n:,} гүйлгээ дахин ангилагдлаа.")
            st.rerun()
        except db.DatabaseError as exc:
            st.error(f"❌ {exc}")


def save_edits(original, edited):
    """data_editor-ийн өөрчлөлтийг database-д хадгална."""
    field_map = {"date_dt": "date", "description": "description", "finance_type": "finance_type",
                 "transaction_type": "transaction_type", "category": "category",
                 "person": "person", "amount": "amount"}
    try:
        to_delete = edited.index[edited["delete"]].tolist()
        updated = 0
        for tx_id in edited.index:
            if tx_id in to_delete:
                continue
            changes = {}
            for col, field in field_map.items():
                old, new = original.at[tx_id, col], edited.at[tx_id, col]
                if pd.isna(new) and pd.isna(old):
                    continue
                if new is None and field in ("description", "person"):
                    new = ""
                if old != new:
                    changes[field] = new.isoformat() if isinstance(new, date) else new
            if changes:
                if "amount" in changes and pd.isna(changes["amount"]):
                    st.error(f"❌ {config.ERRORS['not_numeric']} (ID {tx_id})")
                    return
                db.update_transaction(int(tx_id), changes)
                updated += 1
        deleted = db.delete_transactions(to_delete)
    except db.DatabaseError as exc:
        st.error(f"❌ {exc}")
        return
    st.success(f"✅ {updated} гүйлгээ засагдаж, {deleted} гүйлгээ устгагдлаа.")
    st.rerun()


def page_finance_type(df, settings, finance_type):
    """Business эсвэл Personal хуудас."""
    is_business = finance_type == config.BUSINESS
    color = C["business"] if is_business else C["personal"]
    page_header(finance_type, "Бизнесийн орлого, зарлага, ашиг" if is_business
                else "Хувийн орлого, зарлага (Business тооцоонд орохгүй)")
    if df.empty:
        empty_state()
        return
    part = df[df["finance_type"] == finance_type]
    inc = part.loc[part["transaction_type"] == config.INCOME, "amount"].abs().sum()
    exp = part.loc[part["transaction_type"] == config.EXPENSE, "amount"].abs().sum()

    if is_business:
        s = reports.summarize(df, settings)
        kpi_grid([
            ("Business орлого", money(inc), C["income"]),
            ("Business зарлага", money(exp), C["expense"]),
            ("Estimated GST", money(s["estimated_gst"]), C["gst"]),
            ("Estimated Net Profit", money(s["net_profit"]), C["profit"]),
        ])
        note(f"{GST_NOTE}. {TAX_NOTE}")
    else:
        kpi_grid([
            ("Personal орлого", money(inc), C["income"]),
            ("Personal зарлага", money(exp), C["expense"]),
            ("Зөрүү", money(inc - exp), color),
            ("Гүйлгээ", f"{len(part):,}", color),
        ])

    c1, c2 = st.columns(2)
    with c1:
        show_chart(chart_income_vs_expense(inc, exp, f"{finance_type} Income vs Expense"))
    with c2:
        show_chart(chart_categories(df, finance_type, f"{finance_type} зарлагын ангилал"))
    show_chart(chart_monthly(df, finance_type, f"{finance_type} — сар бүрийн орлого/зарлага"))

    if is_business:
        cat = reports.expense_by_category(df, config.BUSINESS)
        if not cat.empty:
            st.markdown("**Зарлагын ангилал:**")
            st.dataframe(cat.rename(columns={"category": "Ангилал", "amount": "Дүн", "count": "Тоо"})
                         .style.format({"Дүн": "{:,.2f}"}), width="stretch", hide_index=True)
    st.markdown(f"**{finance_type} гүйлгээнүүд:**")
    display_table(reports.prepare(part))


def party_table(tx):
    """Харилцагчийн гүйлгээг Огноо / Тайлбар / Ангилал / Дүн баганатай хүснэгтээр харуулна."""
    if tx.empty:
        st.caption("Гүйлгээ алга.")
        return
    view = tx[["date", "description", "category", "amount"]].rename(
        columns={"date": "Огноо", "description": "Тайлбар", "category": "Ангилал", "amount": "Дүн"})
    st.dataframe(view.style.map(amount_color, subset=["Дүн"]).format({"Дүн": "{:,.2f}"}),
                 width="stretch", hide_index=True, height=min(38 * (len(view) + 1) + 3, 400))


def focus_parties_section(df):
    """
    Сонгосон харилцагчдын (config.FOCUS_PARTIES) орлого, зарлагыг тусад нь харуулна:
    нэгтгэл хүснэгт, дараа нь харилцагч бүрт tab — Орлого ба Зарлага тусдаа хүснэгт.
    """
    st.subheader("Сонгосон харилцагчид — орлого / зарлага тусдаа")
    options = list(dict.fromkeys(
        config.FOCUS_PARTIES + config.TRACKED_PEOPLE + sorted(p for p in df["person"].unique() if p)))
    parties = st.multiselect("Харилцагч", options, default=config.FOCUS_PARTIES)
    if not parties:
        return

    summaries = [reports.person_summary(df, name) for name in parties]
    overview = pd.DataFrame([{"Харилцагч": s["person"], "Орлого (авсан)": s["received"],
                              "Зарлага (өгсөн)": s["given"], "Зөрүү": s["difference"],
                              "Гүйлгээ": s["count"]} for s in summaries])
    st.dataframe(overview.style.format({"Орлого (авсан)": "{:,.2f}", "Зарлага (өгсөн)": "{:,.2f}",
                                        "Зөрүү": "{:,.2f}"}).map(amount_color, subset=["Зөрүү"]),
                 width="stretch", hide_index=True)

    for tab, s in zip(st.tabs([s["person"] for s in summaries]), summaries):
        with tab:
            kpi_grid([
                ("Орлого (авсан)", money(s["received"]), C["income"]),
                ("Зарлага (өгсөн)", money(s["given"]), C["expense"]),
                ("Зөрүү", money(s["difference"]), C["profit"]),
                ("Гүйлгээний тоо", f"{s['count']}", C["primary"]),
            ])
            tx = s["transactions"]
            if tx.empty:
                st.info(f"{s['person']}-тэй холбоотой гүйлгээ олдсонгүй.")
                continue
            left, right = st.columns(2)
            with left:
                st.markdown(f"<h5 style='color:{C['income']}'>🟢 Орлого</h5>", unsafe_allow_html=True)
                party_table(tx[tx["transaction_type"] == config.INCOME])
            with right:
                st.markdown(f"<h5 style='color:{C['expense']}'>🔴 Зарлага</h5>", unsafe_allow_html=True)
                party_table(tx[tx["transaction_type"] == config.EXPENSE])

            monthly = reports.party_monthly(df, s["person"])
            fig = go.Figure()
            fig.add_bar(name="Орлого (авсан)", x=monthly["month"], y=monthly["Авсан"], marker_color=C["income"])
            fig.add_bar(name="Зарлага (өгсөн)", x=monthly["month"], y=monthly["Өгсөн"], marker_color=C["expense"])
            fig.update_layout(title=f"{s['person']} — сар бүрээр", barmode="group")
            show_chart(style_fig(fig, 320))
            st.download_button(f"⬇️ {s['person']} — CSV татах", reports.export_csv_bytes(tx),
                               file_name=f"{s['person'].replace(' ', '_')}.csv", mime="text/csv",
                               key=f"dl_{s['person']}")


def page_people(df):
    """People — хүн тус бүрийн гүйлгээний тайлан."""
    page_header("People", "Тусгайлан бүртгэсэн хүмүүсийн гүйлгээ")
    if df.empty:
        empty_state()
        return

    focus_parties_section(df)
    st.divider()
    st.subheader("Бүх бүртгэлтэй хүмүүс")

    overview = reports.people_overview(df)
    st.dataframe(overview.style.format({"Өгсөн": "{:,.2f}", "Авсан": "{:,.2f}", "Зөрүү": "{:,.2f}"})
                 .map(amount_color, subset=["Зөрүү"]), width="stretch", hide_index=True)
    show_chart(chart_people(df))

    others = sorted(p for p in df["person"].unique()
                    if p and p not in config.TRACKED_PEOPLE and p not in config.PERSONAL_NAMES)
    extra = st.selectbox("Өөр хүн харах (Person талбараас)", ["—"] + others)
    names = list(config.TRACKED_PEOPLE) + ([extra] if extra != "—" else [])

    for name in names:
        s = reports.person_summary(df, name)
        with st.expander(f"👤 {name} — Өгсөн {money(s['given'])} · Авсан {money(s['received'])} · "
                         f"Гүйлгээ {s['count']}", expanded=(name == extra)):
            kpi_grid([
                ("Нийт өгсөн мөнгө", money(s["given"]), C["expense"]),
                ("Нийт авсан мөнгө", money(s["received"]), C["income"]),
                ("Зөрүү", money(s["difference"]), C["profit"]),
                ("Гүйлгээний тоо", f"{s['count']}", C["primary"]),
            ])
            tx = s["transactions"]
            if tx.empty:
                st.caption("Гүйлгээ олдсонгүй.")
            else:
                view = tx[["date", "description", "transaction_type", "category", "amount"]].rename(
                    columns={"date": "Огноо", "description": "Тайлбар", "transaction_type": "Төрөл",
                             "category": "Ангилал", "amount": "Дүн"})
                st.dataframe(view.style.map(amount_color, subset=["Дүн"]).format({"Дүн": "{:,.2f}"}),
                             width="stretch", hide_index=True)


def page_reports(df, settings):
    """Тайлан — Өдөр / 7 хоног / Сар / Custom хугацаагаар."""
    page_header("Тайлан", "Хугацаагаар тайлан гаргах")
    if df.empty:
        empty_state()
        return
    data = reports.prepare(df)
    min_d, max_d = data["date_dt"].min().date(), data["date_dt"].max().date()

    period = st.radio("Хугацаа", ["Өдөр", "7 хоног", "Сар", "Custom"], horizontal=True)
    if period == "Custom":
        c1, c2 = st.columns(2)
        start_in = c1.date_input("Эхлэх огноо", value=min_d, min_value=min_d, max_value=max_d)
        end_in = c2.date_input("Дуусах огноо", value=max_d, min_value=min_d, max_value=max_d)
        start, end = reports.period_range(period, max_d, start_in, end_in)
    else:
        label = {"Өдөр": "Өдөр сонгох", "7 хоног": "7 хоногийн сүүлийн өдөр", "Сар": "Сарын аль нэг өдөр"}
        ref = st.date_input(label[period], value=max_d, min_value=min_d, max_value=max_d)
        start, end = reports.period_range(period, ref)

    if start > end:
        st.error("❌ Огноо буруу: эхлэх огноо дуусах огнооноос хойш байна.")
        return

    part = reports.filter_by_date(df, start, end)
    st.markdown(f"#### 📅 {start.date()} — {end.date()}")
    s = reports.summarize(part, settings)
    kpi_grid([
        ("Business Income", money(s["business_income"]), C["income"]),
        ("Business Expense", money(s["business_expense"]), C["expense"]),
        ("Personal Income", money(s["personal_income"]), C["personal"]),
        ("Personal Expense", money(s["personal_expense"]), C["personal"]),
        ("Estimated GST", money(s["estimated_gst"]), C["gst"]),
        ("Estimated Tax", money(s["estimated_tax"]), C["tax"]),
        ("Estimated Net Profit", money(s["net_profit"]), C["profit"]),
        ("Transaction Count", f"{s['transaction_count']:,}", C["primary"]),
    ])
    note(f"{GST_NOTE}. {TAX_NOTE}")

    if part.empty:
        st.info("Энэ хугацаанд гүйлгээ алга.")
        return

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Expense Categories**")
        cat = reports.expense_by_category(part)
        st.dataframe(cat.rename(columns={"category": "Ангилал", "amount": "Дүн", "count": "Тоо"})
                     .style.format({"Дүн": "{:,.2f}"}), width="stretch", hide_index=True)
    with c2:
        st.markdown("**Person Payments**")
        people = reports.people_overview(part)
        st.dataframe(people.style.format({"Өгсөн": "{:,.2f}", "Авсан": "{:,.2f}", "Зөрүү": "{:,.2f}"}),
                     width="stretch", hide_index=True)

    profit_calc_block(s)
    with st.expander("Хугацааны гүйлгээнүүд"):
        display_table(part)
    st.download_button("⬇️ Тайлангийн гүйлгээг CSV-ээр татах", reports.export_csv_bytes(part),
                       file_name=f"report_{start.date()}_{end.date()}.csv", mime="text/csv")


def page_tax(df, settings):
    """Tax Estimate — GST ба Австралийн татварын дэлгэрэнгүй."""
    page_header("Tax Estimate", "GST болон Australian Tax — урьдчилсан тооцоо")
    note(TAX_NOTE)
    note(f"{GST_NOTE}. Бодит BAS тооцоо нь гүйлгээ GST-inclusive эсэх, GST-free орлого, "
         "input tax credits зэрэг нөхцөлөөс хамаарна.")
    if df.empty:
        empty_state()
        return
    s = reports.summarize(df, settings)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Одоогийн тохиргоо**")
        rate_text = (f"Progressive ({settings['tax_year']})"
                     if settings["entity_type"] == config.ENTITY_INDIVIDUAL else f"{settings['tax_rate']:.2f}%")
        st.markdown(
            f"""
            <div class="section-card">
            <div class="calc-row"><span>GST rate</span><span>{settings['gst_rate']:.2f}%</span></div>
            <div class="calc-row"><span>GST mode</span><span>{config.GST_MODES[settings['gst_mode']]}</span></div>
            <div class="calc-row"><span>Entity type</span><span>{settings['entity_type']}</span></div>
            <div class="calc-row"><span>Tax rate</span><span>{rate_text}</span></div>
            <div class="calc-row"><span>Effective tax rate</span><span>{s['effective_tax_rate']:.2f}%</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption("Тохиргоог **Settings** хэсгээс өөрчилнө.")
    with c2:
        st.markdown("**Тооцооны дараалал**")
        profit_calc_block(s)

    show_chart(chart_profit_waterfall(s))

    if settings["entity_type"] == config.ENTITY_INDIVIDUAL:
        st.markdown(f"**Resident individual tax brackets ({settings['tax_year']})**")
        brackets = config.INDIVIDUAL_TAX_BRACKETS[settings["tax_year"]]
        rows = []
        for i, (low, rate) in enumerate(brackets):
            high = brackets[i + 1][0] if i + 1 < len(brackets) else None
            rows.append({"Орлого": f"{money(low + (1 if low else 0))} – {money(high) if high else 'дээш'}",
                         "Хувь": f"{rate * 100:.0f}%"})
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        if not config.INCLUDE_MEDICARE_LEVY:
            st.caption("Medicare levy (2%) оруулаагүй. config.py → INCLUDE_MEDICARE_LEVY.")


def page_export(df):
    """Export — боловсруулсан мэдээллийг CSV-ээр татах."""
    page_header("Export", "Боловсруулсан гүйлгээг CSV файлаар татах")
    if df.empty:
        empty_state()
        return
    scope = st.radio("Юуг export хийх вэ?", ["Бүгд", config.BUSINESS, config.PERSONAL], horizontal=True)
    part = df if scope == "Бүгд" else df[df["finance_type"] == scope]
    export_df = reports.export_dataframe(part)
    st.caption(f"{len(export_df):,} мөр · Баганууд: {', '.join(export_df.columns)}")
    st.dataframe(export_df.head(200), width="stretch", hide_index=True, height=420)
    st.download_button("⬇️ CSV татах", reports.export_csv_bytes(part),
                       file_name=f"transactions_{scope.lower() if scope != 'Бүгд' else 'all'}_{date.today()}.csv",
                       mime="text/csv", type="primary")


def page_database():
    """Database — SQLite-д хадгалагдсан бүх хүснэгт, бүх баганыг харуулна (зөвхөн харах)."""
    page_header("Database", "finance.db-д хадгалагдсан бүх мэдээлэл")
    try:
        counts = db.table_counts()
    except db.DatabaseError as exc:
        st.error(f"❌ {exc}")
        return

    size_kb = os.path.getsize(config.DB_PATH) / 1024 if os.path.exists(config.DB_PATH) else 0
    kpi_grid([(label, f"{counts[t]:,} мөр", C["primary"]) for t, label in db.DB_TABLES.items()])
    st.caption(f"Файл: `{os.path.abspath(config.DB_PATH)}` · Хэмжээ: {size_kb:,.0f} KB. "
               "Энэ хуудас зөвхөн харуулна — засварыг **Гүйлгээ** болон **Settings** хуудсанд хийнэ.")

    for tab, (table, label) in zip(st.tabs(list(db.DB_TABLES.values())), db.DB_TABLES.items()):
        with tab:
            data = db.read_table(table)
            if data.empty:
                st.caption("Хоосон.")
                continue
            if table == "transactions":
                q = st.text_input("🔍 Бүх баганаас хайх", key="db_search")
                if q.strip():
                    mask = data.astype(str).apply(
                        lambda col: col.str.contains(q.strip(), case=False, regex=False)).any(axis=1)
                    data = data[mask]
                st.caption(f"{len(data):,} мөр · manual_override = 1 бол гараар засагдсан.")
            height = 600 if table == "transactions" else min(38 * (len(data) + 1) + 3, 600)
            st.dataframe(data, width="stretch", hide_index=True, height=height)
            st.download_button(f"⬇️ {label} — CSV татах", data.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"{table}.csv", mime="text/csv", key=f"dl_db_{table}")

    if os.path.exists(config.DB_PATH):
        with open(config.DB_PATH, "rb") as fh:
            st.download_button("💾 Database файлыг бүтнээр нь татах (нөөц)", fh.read(),
                               file_name=f"finance_backup_{date.today()}.db",
                               mime="application/octet-stream")


def page_settings(settings, categories):
    """Settings — GST, татвар, ангилал, database."""
    page_header("Settings", "GST, татвар, ангиллын тохиргоо")

    st.subheader("GST")
    note(GST_NOTE)
    gst_rate = st.number_input("GST rate (%)", min_value=0.0, max_value=100.0,
                               value=float(settings["gst_rate"]), step=0.5)
    modes = list(config.GST_MODES)
    gst_mode = st.radio("GST тооцох арга", modes, index=modes.index(settings["gst_mode"]),
                        format_func=lambda m: config.GST_MODES[m])

    st.subheader("Australian Tax")
    note(TAX_NOTE)
    entity = st.selectbox("Entity Type", config.ENTITY_TYPES,
                          index=config.ENTITY_TYPES.index(settings["entity_type"]))
    tax_rate = float(settings["tax_rate"])
    tax_year = settings["tax_year"]
    if entity == config.ENTITY_COMPANY:
        options = list(config.COMPANY_TAX_RATE_OPTIONS) + ["Өөрөө оруулах"]
        current = next((k for k, v in config.COMPANY_TAX_RATE_OPTIONS.items() if v == tax_rate),
                       "Өөрөө оруулах")
        choice = st.selectbox("Company tax rate", options, index=options.index(current))
        if choice == "Өөрөө оруулах":
            tax_rate = st.number_input("Tax rate (%)", 0.0, 100.0, tax_rate, 0.5)
        else:
            tax_rate = config.COMPANY_TAX_RATE_OPTIONS[choice]
    elif entity == config.ENTITY_CUSTOM:
        tax_rate = st.number_input("Custom Tax Rate (%)", 0.0, 100.0, tax_rate, 0.5)
    else:
        years = list(config.INDIVIDUAL_TAX_BRACKETS)
        tax_year = st.selectbox("Санхүүгийн жил (tax brackets)", years,
                                index=years.index(tax_year) if tax_year in years else 0)
        st.caption("Progressive tax: шатлалуудыг config.py → INDIVIDUAL_TAX_BRACKETS-д шинэчилнэ.")

    if st.button("💾 Тохиргоо хадгалах", type="primary"):
        try:
            db.save_settings({"gst_rate": gst_rate, "gst_mode": gst_mode, "entity_type": entity,
                              "tax_rate": tax_rate, "tax_year": tax_year})
            st.success("✅ Тохиргоо хадгалагдлаа.")
        except db.DatabaseError as exc:
            st.error(f"❌ {exc}")

    st.divider()
    st.subheader("Зарлагын ангилал")
    st.write(", ".join(f"`{c}`" for c in categories))
    c1, c2 = st.columns(2)
    with c1:
        new_cat = st.text_input("Шинэ ангилал нэмэх")
        if st.button("➕ Нэмэх"):
            try:
                if db.add_category(new_cat):
                    st.success(f"✅ '{new_cat.strip()}' ангилал нэмэгдлээ.")
                    st.rerun()
                else:
                    st.warning("Ангиллын нэр хоосон эсвэл аль хэдийн байна.")
            except db.DatabaseError as exc:
                st.error(f"❌ {exc}")
    with c2:
        del_cat = st.selectbox("Ангилал устгах", ["—"] + categories)
        if st.button("🗑️ Устгах") and del_cat != "—":
            try:
                if db.delete_category(del_cat):
                    st.success(f"✅ '{del_cat}' устгагдлаа. Энэ ангилалтай гүйлгээ 'Бусад' болсон.")
                    st.rerun()
                else:
                    st.warning("Системийн үндсэн ангиллыг устгах боломжгүй.")
            except db.DatabaseError as exc:
                st.error(f"❌ {exc}")

    st.divider()
    st.subheader("Database")
    if st.button("🔄 Бүх гүйлгээг дахин ангилах (гараар засснаас бусад)"):
        try:
            n = db.reclassify_transactions(classify_transaction)
            st.success(f"✅ {n:,} гүйлгээ дахин ангилагдлаа.")
        except db.DatabaseError as exc:
            st.error(f"❌ {exc}")
    confirm = st.checkbox("Бүх гүйлгээг устгахыг баталж байна")
    if st.button("⚠️ Бүх гүйлгээг устгах", disabled=not confirm):
        try:
            db.delete_all_transactions()
            st.success("✅ Бүх гүйлгээ устгагдлаа.")
            st.rerun()
        except db.DatabaseError as exc:
            st.error(f"❌ {exc}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    """Програмын эхлэл: тохиргоо, sidebar, хуудас сонголт."""
    st.set_page_config(page_title=config.APP_TITLE, page_icon="💼", layout="wide")
    inject_css()

    try:
        db.init_db()
        df = db.get_transactions()
        settings = db.get_settings()
        categories = db.get_categories()
    except db.DatabaseError as exc:
        st.error(f"❌ {exc}")
        st.stop()

    with st.sidebar:
        st.markdown(f"<h2 style='color:{C['primary']};margin-bottom:0'>💼 Санхүү</h2>"
                    f"<p style='color:{C['muted']};margin-top:2px'>{config.APP_TITLE}</p>",
                    unsafe_allow_html=True)
        page = st.radio("Цэс", PAGES, format_func=lambda p: f"{PAGE_ICONS[p]}  {p}",
                        label_visibility="collapsed")
        st.divider()
        st.caption(f"Нийт гүйлгээ: **{len(df):,}**")
        st.markdown(
            f'<span class="badge" style="background:{C["income"]}">Орлого</span> '
            f'<span class="badge" style="background:{C["expense"]}">Зарлага</span> '
            f'<span class="badge" style="background:{C["business"]}">Business</span> '
            f'<span class="badge" style="background:{C["personal"]}">Personal</span>',
            unsafe_allow_html=True,
        )

    if page == "Dashboard":
        page_dashboard(df, settings)
    elif page == "CSV Upload":
        page_upload(df)
    elif page == "Гүйлгээ":
        page_transactions(df, categories)
    elif page == "Business":
        page_finance_type(df, settings, config.BUSINESS)
    elif page == "Personal":
        page_finance_type(df, settings, config.PERSONAL)
    elif page == "People":
        page_people(df)
    elif page == "Тайлан":
        page_reports(df, settings)
    elif page == "Tax Estimate":
        page_tax(df, settings)
    elif page == "Export":
        page_export(df)
    elif page == "Database":
        page_database()
    elif page == "Settings":
        page_settings(settings, categories)


if __name__ == "__main__":
    main()
