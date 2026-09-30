# Орлого, Зарлагын Хяналтын Систем (MVP)

Банкны CSV дансны хуулгыг автоматаар боловсруулж, орлого/зарлага, Business/Personal,
зардлын ангилал, хүн тус бүрийн гүйлгээ, GST болон татварын урьдчилсан тооцоо,
цэвэр ашиг, тайлан, CSV export гаргадаг Streamlit web application.

```
CSV Upload → Data Cleaning → Орлого/Зарлага → Business/Personal → Зардлын ангилал
→ Хүн тус бүрийн гүйлгээ → GST Estimate → Tax Estimate → Цэвэр ашиг
→ Dashboard → Report → CSV Export
```

---

## 1. Архитектур

| Файл | Үүрэг |
|---|---|
| `app.py` | Streamlit UI — sidebar цэс, 10 хуудас, графикууд |
| `config.py` | **Бүх тохиргоо нэг дор**: нэрс, keyword, багана таних, татварын шатлал, өнгө |
| `data_processor.py` | CSV унших, багана таних, огноо/дүн цэвэрлэх |
| `categorizer.py` | Ангиллын дүрмүүд (Priority 1–5), нэр таних (fuzzy) |
| `database.py` | SQLite: transactions, settings, categories хүснэгт; duplicate шалгалт |
| `tax_calculator.py` | GST (Mode A/B), Company/Individual/Custom татвар, ашгийн задаргаа |
| `reports.py` | KPI, сар бүрийн нэгтгэл, хүн тус бүрийн тайлан, хугацааны filter, export |
| `tests.py` | 34 unit test |
| `sample_transactions.csv` | Туршилтын өгөгдөл |

### Ангиллын дүрмийн дараалал

1. **Priority 1** — Дүнгийн тэмдэг: `-500,000` = Зарлага, `500,000` = Орлого (`+` шаардлагагүй).
   Тусдаа Орлого/Зарлага (Credit/Debit) баганатай бол тэдгээрийг ашиглана.
2. **Priority 2** — `Ganzorig Batzorig`, `Munkhchimeg Lkhagvatseren` илэрвэл → **Personal**.
3. **Priority 3** — `Munkhbat Ulziikhutag`, `Davaadulam Dashdavaa`, `Davaadolgor Dashdavaa`, `Deleg` → Person талбар.
4. **Priority 4** — Тайлбарт `payment` үг + хувь хүний нэр + зарлага → **Цалин / Ажилтны төлбөр**.
   (`Repayment`, `PTY LTD` зэрэг байгууллага тооцогдохгүй.)
5. **Priority 5** — Keyword-based зардлын ангилал (`config.py → CATEGORY_KEYWORDS`).

Нэр таних нь том/жижиг үсэг, жижиг алдаа (`Munkbat`), банкны таслагдсан нэр
(`MUNKHCHIMEG LKHAGVATS`), нэр/овгийн солигдсон дарааллыг таньна.

---

## 2. Суулгах (алхам алхмаар)

### Алхам 1 — Python суулгах
Python 3.9 эсвэл түүнээс дээш хувилбар хэрэгтэй. Terminal нээгээд шалгана:

```bash
python3 --version
```

Байхгүй бол https://www.python.org/downloads/ -оос татаж суулгана
(Windows дээр суулгахдаа **"Add Python to PATH"**-г чагтлаарай).

### Алхам 2 — Project folder руу орох

```bash
cd /path/to/Myproject0926
```

### Алхам 3 — Virtual environment үүсгэх
Virtual environment нь энэ project-ийн сангуудыг компьютерын бусад хэсгээс тусгаарлана.

```bash
python3 -m venv .venv
```

Идэвхжүүлэх:

```bash
# macOS / Linux
source .venv/bin/activate

# Windows (PowerShell)
.venv\Scripts\Activate.ps1
```

Идэвхжсэн бол terminal-ийн эхэнд `(.venv)` гэж харагдана.

### Алхам 4 — Сангууд суулгах

```bash
pip install -r requirements.txt
```

### Алхам 5 — Програм ажиллуулах

```bash
streamlit run app.py
```

Browser автоматаар `http://localhost:8501` хаягаар нээгдэнэ. Зогсоохдоо terminal дээр `Ctrl + C`.

---

## 3. CSV upload хийх

1. Зүүн цэснээс **CSV Upload** сонгоно.
2. **CSV файл сонгох** дээр дарж банкны хуулгаа оруулна
   (эсвэл туршихын тулд **Sample CSV ашиглах** товч).
3. Файл **шууд database-д хадгалагдана** — тусдаа хадгалах товч дарах шаардлагагүй.
   Эх файлын хуулбар `uploads/` хавтаст үлдэнэ.
4. Дараагийн удаа програм нээхэд өгөгдөл автоматаар ачаалагдана — **дахин upload хийхгүй**.
   Засвараа **Гүйлгээ** хуудсанд хийнэ. Шинэ хуулга гарвал зөвхөн тэр файлыг нэмж оруулна.

Ижил файлыг дахин оруулбал "аль хэдийн хадгалагдсан" гэж мэдэгдээд алгасна.
Давхцсан хугацаатай шинэ файл оруулбал зөвхөн шинэ гүйлгээнүүд нэмэгдэнэ.
**Хадгалагдсан файлууд** жагсаалтаас файлыг устгавал түүнээс орсон гүйлгээнүүд хамт устана.

Дэмжих CSV хэлбэр:

- Header-тэй: `Date, Description, Amount, Balance` эсвэл
  `Огноо, Гүйлгээний тайлбар, Орлого, Зарлага, Дансны үлдэгдэл` гэх мэт.
- Header-гүй (CommBank маяг): `29/6/2025,-14.99,StepPay Repayment,119.72`.
- Огноо: `29/06/2025`, `2025-06-29`, `2025.06.29` …
- Дүн: `1,500,000`, `-85,000`, `$1,234.50`, `(500)`, `250.00 DR` …
- Encoding: UTF-8, Cyrillic (cp1251).

Алдаа гарвал Монгол хэлээр тайлбарлана (CSV формат буруу, Amount багана олдохгүй,
Огноо буруу, Дүн numeric биш, CSV хоосон, Database error, Duplicate transaction).

---

## 4. Database хэрхэн ажилладаг вэ

- Бүх мэдээлэл project folder дотор **`finance.db`** (SQLite) файлд хадгалагдана. Анх ажиллуулахад автоматаар үүснэ.
- `transactions` хүснэгт: `id, date, description, amount, transaction_type, finance_type,
  category, person, balance, source_file, created_at` (+ `dup_key`, `manual_override` туслах баганууд).
- `settings` хүснэгт: `gst_rate, gst_mode, entity_type, tax_rate, tax_year`.
- `categories` хүснэгт: ангиллын жагсаалт (Settings-ээс шинээр нэмнэ).
- `uploaded_files` хүснэгт: upload хийсэн файлууд (нэр, hash, хадгалсан огноо, гүйлгээний тоо).

**Duplicate detection:** Ижил CSV-г дахин upload хийхэд `Огноо + Тайлбар + Дүн` ижил гүйлгээг
давхар оруулахгүй. Нэг өдөр яг ижил хоёр гүйлгээ (жишээ нь хоёр кофе) байвал хоёуланг нь хадгална.

**Гараар засах:** **Гүйлгээ** хуудсанд хүснэгтийн нүдэн дээр дарж Business↔Personal,
Орлого↔Зарлага, Ангилал, Хүн, Тайлбар, Дүн, Огноог засна; **Устгах** баганаар тэмдэглэж устгана;
**💾 Хадгалах** дарна. Төрлийг сольвол дүнгийн тэмдэг автоматаар тохирно.
Гараар зассан гүйлгээг "Автоматаар дахин ангилах" өөрчлөхгүй.

Бүх өгөгдлийг цэвэрлэх: **Settings → Database → Бүх гүйлгээг устгах**, эсвэл `finance.db`-г устгах.

---

## 5. GST / Tax settings

**Settings** хуудсанд:

| Тохиргоо | Тайлбар |
|---|---|
| GST rate (%) | Анхдагч 10% |
| GST mode A | `Estimated GST = Business Revenue × GST rate` |
| GST mode B | `Estimated GST = GST-inclusive Business Revenue ÷ 11` |
| Entity Type: Company | 25% (base rate entity), 30% эсвэл өөрөө оруулна |
| Entity Type: Sole Trader / Individual | Progressive tax — санхүүгийн жилийн шатлалаар |
| Entity Type: Custom | Custom Tax Rate (%) |

Цэвэр ашгийн тооцоо (зөвхөн Business гүйлгээ):

```
  Business Revenue
− Business Expenses
− Estimated GST
= Profit Before Income Tax
− Estimated Australian Tax
= Estimated Net Profit
```

Татварын шатлалыг шинэчлэх бол `config.py → INDIVIDUAL_TAX_BRACKETS`-д шинэ жил нэмнэ.
Medicare levy оруулах бол `INCLUDE_MEDICARE_LEVY = True`.

> ⚠️ **Estimated GST — зөвхөн урьдчилсан тооцоо.** Бодит BAS тооцоо нь гүйлгээ GST-inclusive эсэх,
> GST-free income, input tax credits зэргээс хамаарна.
> ⚠️ **Estimated Australian Tax — татварын албан ёсны зөвлөгөө биш, урьдчилсан тооцоо.**

---

## 6. Тест ажиллуулах

Virtual environment идэвхтэй үед:

```bash
python tests.py
```

Шалгах зүйлс: эерэг дүн = Орлого, сөрөг = Зарлага; Ganzorig/Munkhchimeg = Personal;
Munkhbat/Davaadulam/Davaadolgor/Deleg payment = Цалин; GST Mode A/B; Company/Individual татвар;
Business/Personal тусдаа; цэвэр ашиг; duplicate; гараар засах; export багана.
Төгсгөлд `OK` гэж гарвал бүх тест амжилттай.

---

## 7. Тохиргоо өөрчлөх (config.py)

- Personal нэр нэмэх → `PERSONAL_NAMES`
- Хянах хүн нэмэх → `TRACKED_PEOPLE`
- Keyword нэмэх → `CATEGORY_KEYWORDS`
- Мөнгөн тэмдэг → `CURRENCY_SYMBOL`

Өөрчилсний дараа **Settings → Бүх гүйлгээг дахин ангилах** дарж хуучин гүйлгээнд хэрэгжүүлнэ.

---

## 8. Excel санхүүгийн тайлан (financial_analysis.py)

Банкны CSV-г бүрэн шинжилж, 15 sheet-тэй **Business_Financial_Analysis.xlsx** файл гаргана.
Streamlit програмаас тусдаа ажилладаг.

### Ажиллуулах

```bash
cd ~/Desktop/AI/Myproject0926
source .venv/bin/activate
pip install -r requirements.txt          # openpyxl нэмэгдсэн
python financial_analysis.py CSVData-TPAR24-25.csv
```

CSV-гээ `bank_transactions.csv` гэж нэрлэвэл зүгээр `python financial_analysis.py` гэхэд хангалттай.

### Тохиргоо (файлын эхэнд)

| Хувьсагч | Утга |
|---|---|
| `INPUT_FILE`, `OUTPUT_FILE` | Оролтын CSV, гаралтын Excel |
| `ENTITY_TYPE` | `"company"`, `"individual"`, `"custom"` |
| `COMPANY_TAX_RATE` | 0.25 (base rate entity) эсвэл 0.30 |
| `DIRECTOR_NAMES`, `SPECIFIC_PEOPLE`, `KNOWN_CUSTOMERS` | Хүмүүсийн жагсаалт |
| `GST_INCLUSIVE_PAYERS`, `NO_GST_PAYERS` | Нэхэмжлэл шалгаад GST-тэй эсэхийг баталгаажуулсан харилцагчид |
| `NAME_ALIASES` | Нэг компанийн өөр бичлэгүүд |
| `MERCHANT_RULES` | Дэлгүүр/нийлүүлэгчийн ангиллын дүрэм |

### Sheet-үүд

Dashboard · All Transactions · Income Detail · Expense Detail · Income by Person-Company ·
Expenses by Person-Company · Person-Company Summary · Salary-Wages · Personal-Director ·
Specific People · Loans · GST Summary · Monthly Summary · Tax Estimate · Review Required

### Гол зарчим

- Ямар ч гүйлгээ устахгүй. Dashboard дээр CSV-ийн нийт дүн ба тайлангийн дүнгийн зөрүү **$0.00** байх ёстой.
- Нотолгоо хангалтгүй гүйлгээ баталгаажсан бизнесийн орлого/зардалд орохгүй — **Review Required**-д шалтгаантай орно.
- GST = GST-тэй дүн ÷ 11. GST UNKNOWN дүн баталгаажсан GST-д орохгүй.
- Бүх тоо **ESTIMATE ONLY – ACCOUNTANT REVIEW REQUIRED**.
