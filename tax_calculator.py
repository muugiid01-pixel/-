"""
tax_calculator.py — GST болон Австралийн орлогын татварын урьдчилсан тооцоо.

АНХААР: Энэ нь зөвхөн урьдчилсан тооцоо. Бодит BAS/татварын тайлан нь
GST-inclusive эсэх, GST-free орлого, input tax credits, хөнгөлөлт гэх мэт
олон нөхцөлөөс хамаарна. Татвар, шатлалын утгууд config.py-д байгаа.
"""

import config


def calculate_gst(business_revenue, gst_rate=10.0, gst_mode=config.GST_MODE_SIMPLE):
    """
    Estimated GST тооцоо.

    Mode A: Business Revenue × GST rate
    Mode B: GST-inclusive Business Revenue × rate / (1 + rate)
            (10%-ийн үед яг ÷ 11)
    Орлого 0 эсвэл сөрөг бол 0 буцаана.
    """
    revenue = max(float(business_revenue), 0.0)
    rate = float(gst_rate) / 100.0
    if gst_mode == config.GST_MODE_INCLUSIVE:
        return revenue * rate / (1.0 + rate)
    return revenue * rate


def calculate_individual_tax(taxable_income, tax_year=None):
    """
    Sole Trader / Individual — Australian resident progressive tax.

    Шатлалыг config.INDIVIDUAL_TAX_BRACKETS-аас авна. Хүсвэл Medicare levy
    нэмнэ (config.INCLUDE_MEDICARE_LEVY).
    """
    income = max(float(taxable_income), 0.0)
    brackets = config.INDIVIDUAL_TAX_BRACKETS.get(
        tax_year, config.INDIVIDUAL_TAX_BRACKETS[config.DEFAULT_SETTINGS["tax_year"]]
    )
    tax = 0.0
    for i, (threshold, rate) in enumerate(brackets):
        upper = brackets[i + 1][0] if i + 1 < len(brackets) else float("inf")
        if income > threshold:
            tax += (min(income, upper) - threshold) * rate
    if config.INCLUDE_MEDICARE_LEVY:
        tax += income * config.MEDICARE_LEVY_RATE
    return tax


def calculate_flat_tax(taxable_income, tax_rate):
    """Company эсвэл Custom — тогтмол хувиар татвар (алдагдалтай бол 0)."""
    return max(float(taxable_income), 0.0) * float(tax_rate) / 100.0


def calculate_income_tax(profit_before_tax, settings):
    """Entity type-аас хамааран зөв функцийг сонгож татвар бодно."""
    entity = settings.get("entity_type", config.ENTITY_COMPANY)
    if entity == config.ENTITY_INDIVIDUAL:
        return calculate_individual_tax(profit_before_tax, settings.get("tax_year"))
    return calculate_flat_tax(profit_before_tax, settings.get("tax_rate", 0.0))


def effective_tax_rate(profit_before_tax, tax):
    """Бодит (effective) татварын хувь (%)."""
    return (tax / profit_before_tax * 100.0) if profit_before_tax > 0 else 0.0


def profit_breakdown(business_revenue, business_expenses, settings):
    """
    Цэвэр ашгийн бүрэн задаргаа. Зөвхөн Business гүйлгээ оролцоно.

        Business Revenue
      − Business Expenses
      − Estimated GST
      = Profit Before Income Tax
      − Estimated Australian Tax
      = Estimated Net Profit
    """
    revenue = float(business_revenue)
    expenses = abs(float(business_expenses))
    gst = calculate_gst(revenue, settings.get("gst_rate", 10.0), settings.get("gst_mode"))
    profit_before_tax = revenue - expenses - gst
    tax = calculate_income_tax(profit_before_tax, settings)
    return {
        "business_revenue": revenue,
        "business_expenses": expenses,
        "estimated_gst": gst,
        "revenue_after_gst": revenue - gst,
        "profit_before_tax": profit_before_tax,
        "estimated_tax": tax,
        "effective_tax_rate": effective_tax_rate(profit_before_tax, tax),
        "net_profit": profit_before_tax - tax,
    }
