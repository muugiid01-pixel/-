"""
categorizer.py — Гүйлгээг автоматаар ангилах дүрмүүд.

Дүрмүүд дараах priority-аар ажиллана (хоорондоо зөрчилдөхгүй):

    Priority 1: Дүнгийн тэмдгээр Орлого/Зарлага тодорхойлно.
    Priority 2: Ganzorig Batzorig / Munkhchimeg Lkhagvatseren → Personal.
    Priority 3: Тусгайлан бүртгэсэн хүний нэрийг Person талбарт оруулна.
    Priority 4: "Payment + хүний нэр + зарлага" → Цалин / Ажилтны төлбөр.
    Priority 5: Keyword-based зардлын ангилал.
"""

import re
from difflib import SequenceMatcher

import config

# Үгийн эхлэлийг тодорхойлох тэмдэгтүүд (латин, кирилл, тоо).
_WORD_CHARS = "a-z0-9а-яёөү"
_PAYMENT_RE = re.compile(r"\bpayment\b", re.IGNORECASE)


def normalize_text(text):
    """Текстийг жижиг үсэгтэй болгож, үсэг/тооноос бусад тэмдэгтийг зайгаар солино."""
    if text is None:
        return ""
    text = str(text).lower()
    text = re.sub(rf"[^{_WORD_CHARS}]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokenize(text):
    """Текстийг цэвэрлээд үгсийн жагсаалт болгоно."""
    return normalize_text(text).split()


def _similar(a, b):
    """Хоёр үгийн төстэй байдлыг 0–1 хооронд буцаана."""
    return SequenceMatcher(None, a, b).ratio()


def _token_matches(text_token, name_token):
    """Үг яг таарах эсвэл жижиг spelling зөрүүтэй (fuzzy) таарах эсэх."""
    if text_token == name_token:
        return True
    return _similar(text_token, name_token) >= config.NAME_MATCH_THRESHOLD


def _last_token_matches(text_token, name_token):
    """
    Овог/сүүлийн үгийг шалгана. Банк нэрийг тасалж бичдэг тул
    ("LKHAGVATSEREN" → "LKHAGVATS") угтвар таарахыг ч зөвшөөрнө.
    """
    if _token_matches(text_token, name_token):
        return True
    return (
        len(text_token) >= config.NAME_PREFIX_MIN_LENGTH
        and name_token.startswith(text_token)
    )


def name_in_text(text, full_name):
    """
    Тайлбарт тухайн хүний нэр байгаа эсэхийг шалгана.

    - Том/жижиг үсэг хамаарахгүй.
    - Жижиг spelling зөрүүг (fuzzy) таньна.
    - Таслагдсан овгийг таньна.
    - Нэр, овог солигдсон дараалалтай байсан ч таньна.
    - Нэг үгтэй нэрийг (жишээ нь "Deleg") зөвхөн бүтэн үгээр таньна,
      тиймээс "Delger" гэх мэт өөр нэртэй андуурахгүй.
    """
    text_tokens = tokenize(text)
    name_tokens = tokenize(full_name)
    if not text_tokens or not name_tokens:
        return False

    if len(name_tokens) == 1:
        target = name_tokens[0]
        return any(
            tok == target
            or (len(tok) == len(target) and _similar(tok, target) >= config.NAME_MATCH_THRESHOLD)
            for tok in text_tokens
        )

    orders = [name_tokens, list(reversed(name_tokens))]
    for order in orders:
        n = len(order)
        for i in range(len(text_tokens) - n + 1):
            window = text_tokens[i:i + n]
            head_ok = all(_token_matches(window[k], order[k]) for k in range(n - 1))
            if head_ok and _last_token_matches(window[-1], order[-1]):
                return True
    return False


def name_or_alias_in_text(text, name):
    """Стандарт нэр эсвэл config.NAME_ALIASES дахь хувилбаруудын аль нэг нь тайлбарт байгаа эсэх."""
    return any(name_in_text(text, variant)
               for variant in [name, *config.NAME_ALIASES.get(name, [])])


def is_organization(name):
    """Нэрэнд PTY, LTD гэх мэт байгууллагын үг байвал True."""
    return any(tok in config.BUSINESS_NAME_WORDS for tok in tokenize(name))


def match_person(text, names):
    """Жагсаалтаас тайлбарт таарсан эхний хүний стандарт нэрийг буцаана."""
    for name in names:
        if name_or_alias_in_text(text, name):
            return name
    return None


def detect_personal_name(description):
    """Priority 2: Personal гэж тооцогдох нэр (Ganzorig/Munkhchimeg) байвал буцаана."""
    return match_person(description, config.PERSONAL_NAMES)


def detect_tracked_person(description):
    """Priority 3: Тусгайлан хянах хүмүүсийн нэр байвал буцаана."""
    return match_person(description, config.TRACKED_PEOPLE)


def has_payment_word(description):
    """Тайлбарт бие даасан "payment" үг байгаа эсэх ("Repayment"-г тооцохгүй)."""
    return bool(_PAYMENT_RE.search(str(description or "")))


def extract_recipient_name(description):
    """
    "Transfer To John Smith NetBank payment", "Payment - John Smith" гэх мэт
    тайлбараас хувь хүний нэрийг гаргаж авна.

    Байгууллагын нэр (PTY LTD, SHOP, DENTAL гэх мэт), дансны дугаар эсвэл
    хэт урт нэр байвал None буцаана.
    """
    text = str(description or "")
    lowered = text.lower()
    for prefix in config.RECIPIENT_PREFIXES:
        match = re.search(rf"\b{re.escape(prefix)}\b", lowered)
        if not match:
            continue
        remainder = text[match.end():]
        words = []
        for raw in remainder.split():
            word = raw.strip("-–:,.;()[]\"'")
            if not word:
                continue
            if word.lower() in config.RECIPIENT_STOP_WORDS:
                break
            if not word.replace("-", "").replace("'", "").isalpha():
                break
            words.append(word)
        if not words or len(words) > config.MAX_PERSON_NAME_WORDS:
            return None
        if any(w.lower() in config.BUSINESS_NAME_WORDS for w in words):
            return None
        return " ".join(w.capitalize() for w in words)
    return None


def determine_transaction_type(amount):
    """
    Priority 1: Дүнгийн тэмдгээр төрөл тодорхойлно.
    Сөрөг (-) дүн = Зарлага, бусад (эерэг) дүн = Орлого.
    """
    return config.EXPENSE if float(amount) < 0 else config.INCOME


def _keyword_found(padded_text, keyword):
    """Keyword үгийн эхнээс таарч байгаа эсэх ("rent" нь "current"-д таарахгүй)."""
    pattern = rf"(?<![{_WORD_CHARS}]){re.escape(keyword)}"
    return re.search(pattern, padded_text) is not None


def keyword_category(description):
    """Priority 5: Keyword-ээр зардлын ангилал олно. Олдохгүй бол "Бусад"."""
    padded = f" {str(description or '').lower()} "
    for category, keywords in config.CATEGORY_KEYWORDS:
        if any(_keyword_found(padded, kw) for kw in keywords):
            return category
    return config.OTHER_CATEGORY


def classify_transaction(description, amount):
    """
    Нэг гүйлгээг бүх priority дүрмээр ангилна.

    Returns:
        dict: transaction_type, finance_type, category, person
    """
    transaction_type = determine_transaction_type(amount)
    is_expense = transaction_type == config.EXPENSE
    result = {
        "transaction_type": transaction_type,
        "finance_type": config.BUSINESS,
        "category": config.INCOME_CATEGORY if not is_expense else config.OTHER_CATEGORY,
        "person": "",
    }

    # Priority 2 — Personal нэрс хамгийн түрүүнд.
    personal_name = detect_personal_name(description)
    if personal_name:
        result["finance_type"] = config.PERSONAL
        result["category"] = config.PERSONAL_CATEGORY
        result["person"] = personal_name
        return result

    # Priority 3 — бүртгэлтэй хүн.
    tracked = detect_tracked_person(description)
    if tracked:
        result["person"] = tracked

    # Priority 4 — Payment + хүний нэр + зарлага (байгууллага бол цалин биш).
    if is_expense and has_payment_word(description):
        recipient = tracked if tracked and not is_organization(tracked) else None
        recipient = recipient or (None if tracked else extract_recipient_name(description))
        if recipient:
            result["category"] = config.SALARY_CATEGORY
            result["person"] = recipient
            return result

    # Priority 5 — keyword ангилал (зөвхөн зарлагад).
    if is_expense:
        result["category"] = keyword_category(description)
    return result


def classify_dataframe(df):
    """DataFrame-ийн мөр бүрийг ангилж, шинэ багануудыг нэмж буцаана."""
    df = df.copy()
    results = [classify_transaction(d, a) for d, a in zip(df["description"], df["amount"])]
    for key in ("transaction_type", "finance_type", "category", "person"):
        df[key] = [r[key] for r in results]
    return df
