"""
config.py — Системийн бүх тохиргоо нэг дор.

Энд байгаа утгуудыг өөрчилснөөр ангиллын дүрэм, татварын шатлал,
өнгө, анхдагч тохиргоог кодын бусад хэсэгт хүрэлгүйгээр шинэчилж болно.
"""

# ---------------------------------------------------------------------------
# Ерөнхий
# ---------------------------------------------------------------------------
APP_TITLE = "Орлого, Зарлагын Хяналтын Систем"
DB_PATH = "finance.db"
CURRENCY_SYMBOL = "$"

# Гүйлгээний төрөл (transaction_type)
INCOME = "Орлого"
EXPENSE = "Зарлага"

# Санхүүгийн төрөл (finance_type)
BUSINESS = "Business"
PERSONAL = "Personal"

# ---------------------------------------------------------------------------
# Өнгө (UI)
# ---------------------------------------------------------------------------
COLORS = {
    "primary": "#2F6FDB",      # Medium blue
    "primary_dark": "#1E4FA8",
    "white": "#FFFFFF",
    "light_grey": "#F3F5F8",
    "border": "#E2E6EC",
    "text": "#1F2937",
    "muted": "#6B7280",
    "income": "#16A34A",       # Орлого = ногоон
    "expense": "#DC2626",      # Зарлага = улаан
    "business": "#2F6FDB",     # Business = цэнхэр
    "personal": "#8B5CF6",     # Personal = нил ягаан
    "gst": "#F59E0B",
    "tax": "#EA580C",
    "profit": "#0EA5E9",
}

# ---------------------------------------------------------------------------
# Хүмүүс
# ---------------------------------------------------------------------------
# Priority 2: Эдгээр нэртэй гүйлгээ автоматаар Personal болно.
PERSONAL_NAMES = [
    "Ganzorig Batzorig",
    "Munkhchimeg Lkhagvatseren",
]

# Priority 3: Тусгайлан хянах хүмүүс (People хэсэгт тайлан гарна).
TRACKED_PEOPLE = [
    "Munkhbat Ulziikhutag",
    "Davaadulam Dashdavaa",
    "Davaadolgor Dashdavaa",
    "Deleg",
    "Modun Pty Ltd",
]

# Нэрийн өөр бичлэгүүд (банк богиносгож бичдэг). Стандарт нэр → хувилбарууд.
NAME_ALIASES = {
    "Modun Pty Ltd": ["Modun"],
}

# People хуудсанд тусад нь (орлого, зарлагыг салгаж) харуулах харилцагчид.
FOCUS_PARTIES = [
    "Davaadulam Dashdavaa",
    "Munkhbat Ulziikhutag",
    "Modun Pty Ltd",
]

# Нэр таних fuzzy тохирлын босго (0–1). Их байх тусам хатуу.
NAME_MATCH_THRESHOLD = 0.85
# Банк нэрийг тасалж бичсэн үед (жишээ нь "LKHAGVATS") овгийн эхний хэдэн
# үсэг таарвал хүлээн зөвшөөрөх вэ.
NAME_PREFIX_MIN_LENGTH = 4

# ---------------------------------------------------------------------------
# Ангилал
# ---------------------------------------------------------------------------
SALARY_CATEGORY = "Цалин / Ажилтны төлбөр"
PERSONAL_CATEGORY = "Personal"
INCOME_CATEGORY = "Орлого"
OTHER_CATEGORY = "Бусад"

DEFAULT_CATEGORIES = [
    "Хоол",
    "Тээвэр",
    "Түрээс",
    SALARY_CATEGORY,
    "Худалдан авалт",
    "Үйл ажиллагааны зардал",
    "Материал",
    "Даатгал",
    "Машин / Тээврийн хэрэгсэл",
    "Хүүхэд / Боловсрол",
    "Utility / Төлбөр",
    "Татвар / ATO",
    PERSONAL_CATEGORY,
    INCOME_CATEGORY,
    OTHER_CATEGORY,
]

# Priority 5: Keyword-based ангилал. Дарааллаар шалгана — эхэлж таарсан
# ангилал сонгогдоно. Бүх keyword жижиг үсгээр бичигдэнэ.
CATEGORY_KEYWORDS = [
    ("Татвар / ATO", [
        "tax office", "ato ", "australian taxation office", "татвар",
    ]),
    (SALARY_CATEGORY, [
        "salary", "wage", "wages", "payroll", "цалин",
    ]),
    ("Даатгал", [
        "insurance", "insur", "ial-", "aami", "allianz", "bupa", "medibank",
        "nib ", "hcf", "ahm ", "qbe", "suncorp", "youi", "даатгал",
    ]),
    ("Түрээс", [
        "rent", "rental", "lease", "real estate", "realestate", "property",
        "түрээс",
    ]),
    ("Хүүхэд / Боловсрол", [
        "school", "childcare", "child care", "kindergarten", "preschool",
        "early learning", "tutor", "education", "university", "tafe",
        "college", "kids", "kiddo", "happy faces", "funstart", "сургууль", "цэцэрлэг", "хүүхэд",
    ]),
    ("Хоол", [
        "uber eats", "ubereats", "doordash", "menulog", "woolworths", "coles",
        "aldi", "iga ", "supermarket", "food", "bever", "bakery", "cafe",
        "coffee", "restaurant", "pizza", "kfc", "mcdonald", "hungry jack",
        "sushi", "thai", "kitchen", "grill", "bbq", "noodle", "dumpling",
        "mongol shop", "fresh", "chefs", "stop n save", "costco", "kmall", "asiana", "ger house", "replenish", "butcher",
        "хоол", "дэлгүүр",
    ]),
    ("Тээвэр", [
        "transport nsw", "transportfornsw", "etoll", "e-toll", "opal", "uber", "taxi", "didi",
        "ola ", "train", "linkt", "toll", "parking", "qantas", "jetstar",
        "virgin australia", "airline", "такси", "автобус",
    ]),
    ("Машин / Тээврийн хэрэгсэл", [
        "caltex", "ampol", "bp ", "shell", "7-eleven", "fuel", "petrol",
        "mechanic", "tyre", "tire", "rego", "service nsw", "car wash",
        "carwash", "apex", "nrma", "auto", "бензин", "машин",
    ]),
    ("Utility / Төлбөр", [
        "vodafone", "telstra", "optus", "tpg", "aussie broadband", "internet",
        "mobile", "energy", "agl", "origin", "electricity", "water", "gas ",
        "altogether", "apple.com/bill", "netflix", "spotify", "google",
        "microsoft", "primevideo", "prime video", "subscription", "council", "bill", "интернет", "утас",
        "цахилгаан",
    ]),
    ("Материал", [
        "bunnings", "material", "hardware", "timber", "paint", "cleanmax",
        "supplies", "supply", "tools", "материал",
    ]),
    ("Үйл ажиллагааны зардал", [
        "fee", "officeworks", "xero", "myob", "software", "advertising",
        "facebook", "meta ", "accountant", "asic", "abn", "domain", "hosting",
        "vending", "pathology", "postage", "auspost", "зардал",
    ]),
    ("Худалдан авалт", [
        "kmart", "target", "big w", "amazon", "ebay", "castlery", "ikea",
        "jb hi-fi", "harvey norman", "myer", "david jones", "zippay", "zip.co",
        "steppay", "zipmoney", "afterpay", "revolut", "zara", "mart", "shop", "store", "purchase",
        "худалдан",
    ]),
]

# ---------------------------------------------------------------------------
# "Payment + хүний нэр" дүрэм (Priority 4)
# ---------------------------------------------------------------------------
# Хүлээн авагчийн нэрийг таслах "эхлэл" хэллэгүүд.
RECIPIENT_PREFIXES = [
    "transfer to", "payment to", "paid to", "pay to", "payment",
]
# Нэрийн дараа ирж, нэрийг дуусгадаг банкны үгс.
RECIPIENT_STOP_WORDS = {
    "netbank", "commbank", "app", "payid", "bpay", "card", "value", "date",
    "au", "from", "ref", "reference", "via", "on", "for", "payment",
    "payments", "transfer", "osko", "fast", "phone", "online", "internet",
    "banking", "mobile", "the",
}
# Эдгээр үг орсон бол хувь хүн биш, байгууллага гэж үзнэ.
BUSINESS_NAME_WORDS = {
    "pty", "ltd", "limited", "inc", "llc", "co", "company", "corp", "group",
    "trading", "store", "stores", "shop", "services", "service", "dental",
    "medical", "clinic", "bank", "office", "bakery", "cafe", "restaurant",
    "supermarket", "market", "school", "childcare", "centre", "center",
    "australia", "club", "council", "insurance", "energy", "tax", "cba",
    "account", "savings", "expenses", "hospital", "pharmacy", "university",
}
MAX_PERSON_NAME_WORDS = 3

# ---------------------------------------------------------------------------
# CSV баганын нэрийг таних (жижиг үсгээр)
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "date": [
        "date", "transaction date", "posting date", "value date", "огноо",
        "гүйлгээний огноо", "он сар өдөр",
    ],
    "description": [
        "description", "details", "narrative", "memo", "particulars",
        "transaction details", "reference", "гүйлгээний тайлбар", "тайлбар",
        "утга", "гүйлгээний утга",
    ],
    "amount": [
        "amount", "дүн", "гүйлгээний дүн", "value", "transaction amount",
    ],
    "income": [
        "income", "credit", "credit amount", "deposit", "deposits", "in",
        "money in", "орлого", "кредит",
    ],
    "expense": [
        "expense", "debit", "debit amount", "withdrawal", "withdrawals", "out",
        "money out", "зарлага", "дебит",
    ],
    "balance": [
        "balance", "running balance", "closing balance", "дансны үлдэгдэл",
        "үлдэгдэл",
    ],
}

# ---------------------------------------------------------------------------
# GST
# ---------------------------------------------------------------------------
GST_MODE_SIMPLE = "A"      # Business Revenue × GST rate
GST_MODE_INCLUSIVE = "B"   # GST-inclusive revenue ÷ 11 (rate/(1+rate))
GST_MODES = {
    GST_MODE_SIMPLE: "Mode A — Business Revenue × GST rate",
    GST_MODE_INCLUSIVE: "Mode B — GST-inclusive орлого ÷ 11",
}

# ---------------------------------------------------------------------------
# Австралийн татвар (урьдчилсан тооцоонд зориулсан)
# ---------------------------------------------------------------------------
ENTITY_COMPANY = "Company"
ENTITY_INDIVIDUAL = "Sole Trader / Individual"
ENTITY_CUSTOM = "Custom"
ENTITY_TYPES = [ENTITY_COMPANY, ENTITY_INDIVIDUAL, ENTITY_CUSTOM]

# Company-д санал болгох хувь (хэрэглэгч өөрөө сонгоно/оруулна).
COMPANY_TAX_RATE_OPTIONS = {
    "Base rate entity (25%)": 25.0,
    "Full company rate (30%)": 30.0,
}

# Resident individual tax brackets. Мөр бүр: (доод хязгаар, хувь).
# Хувь нь тухайн доод хязгаараас дээших доллар бүрт үйлчилнэ.
# Шинэ жилийн шатлал гарвал энд нэг мөр нэмэхэд хангалттай.
INDIVIDUAL_TAX_BRACKETS = {
    "2024-25": [(0, 0.0), (18_200, 0.16), (45_000, 0.30), (135_000, 0.37), (190_000, 0.45)],
    "2025-26": [(0, 0.0), (18_200, 0.16), (45_000, 0.30), (135_000, 0.37), (190_000, 0.45)],
    "2026-27": [(0, 0.0), (18_200, 0.15), (45_000, 0.30), (135_000, 0.37), (190_000, 0.45)],
}
# Medicare levy — энгийн тооцоонд оруулах эсэх.
INCLUDE_MEDICARE_LEVY = False
MEDICARE_LEVY_RATE = 0.02

# ---------------------------------------------------------------------------
# Анхдагч тохиргоо (Settings table)
# ---------------------------------------------------------------------------
DEFAULT_SETTINGS = {
    "gst_rate": 10.0,
    "gst_mode": GST_MODE_SIMPLE,
    "entity_type": ENTITY_COMPANY,
    "tax_rate": 25.0,
    "tax_year": "2024-25",
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
EXPORT_COLUMNS = {
    "date": "Date",
    "description": "Description",
    "finance_type": "Business/Personal",
    "transaction_type": "Income/Expense",
    "category": "Category",
    "person": "Person",
    "amount": "Amount",
    "balance": "Balance",
}

# ---------------------------------------------------------------------------
# Монгол алдааны мессеж
# ---------------------------------------------------------------------------
ERRORS = {
    "bad_format": "CSV формат буруу байна. Файлыг CSV (таслалаар тусгаарласан) хэлбэрээр хадгалаад дахин оролдоно уу.",
    "no_amount": "Дүн (Amount) багана олдсонгүй. CSV-д 'Amount/Дүн' эсвэл 'Орлого/Зарлага' (Credit/Debit) багана байх шаардлагатай.",
    "no_date": "Огнооны багана олдсонгүй. CSV-д 'Date/Огноо' багана байх шаардлагатай.",
    "bad_date": "Огноо буруу байна. Огноог уншиж чадсангүй.",
    "not_numeric": "Дүн numeric биш байна. Дүнгийн баганад зөвхөн тоо байх ёстой.",
    "empty": "CSV файл хоосон байна.",
    "db_error": "Database алдаа гарлаа",
    "duplicate": "Давхардсан гүйлгээ илэрлээ — database-д дахин оруулсангүй.",
}
