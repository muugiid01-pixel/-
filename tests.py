"""
tests.py — Ангилал, CSV боловсруулалт, database, GST/татвар, тайлангийн тест.

Ажиллуулах:  python tests.py      (эсвэл  python -m unittest tests -v)
"""

import os
import tempfile
import unittest

import config
import database as db
import reports
from categorizer import (classify_transaction, extract_recipient_name, name_in_text)
from data_processor import CSVProcessingError, parse_amount, process_csv
from tax_calculator import (calculate_gst, calculate_individual_tax, calculate_income_tax,
                            profit_breakdown)

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE_CSV = os.path.join(HERE, "sample_transactions.csv")
SAL = config.SALARY_CATEGORY


class TestRequiredCases(unittest.TestCase):
    """Даалгаврын TEST DATA хэсгийн 8 заавал шалгах тохиолдол."""

    def check(self, description, amount, tx_type, finance=None, category=None, person=None):
        """Нэг гүйлгээг ангилж хүлээгдэж буй үр дүнтэй харьцуулна."""
        r = classify_transaction(description, amount)
        self.assertEqual(r["transaction_type"], tx_type, description)
        if finance:
            self.assertEqual(r["finance_type"], finance, description)
        if category:
            self.assertEqual(r["category"], category, description)
        if person:
            self.assertEqual(r["person"], person, description)

    def test_1_salary_payment_income(self):
        self.check("Salary Payment", 5_000_000, config.INCOME)

    def test_2_supermarket_expense(self):
        self.check("Supermarket", -250_000, config.EXPENSE)

    def test_3_payment_munkhbat(self):
        self.check("Payment Munkhbat Ulziikhutag", -500_000, config.EXPENSE,
                   config.BUSINESS, SAL, "Munkhbat Ulziikhutag")

    def test_4_payment_davaadulam(self):
        self.check("Payment Davaadulam Dashdavaa", -600_000, config.EXPENSE,
                   config.BUSINESS, SAL, "Davaadulam Dashdavaa")

    def test_5_transfer_ganzorig_personal(self):
        self.check("Transfer Ganzorig Batzorig", -1_000_000, config.EXPENSE, config.PERSONAL)

    def test_6_transfer_munkhchimeg_personal(self):
        self.check("Transfer Munkhchimeg Lkhagvatseren", -800_000, config.EXPENSE, config.PERSONAL)

    def test_7_from_munkhchimeg_personal_income(self):
        self.check("Transfer from Munkhchimeg Lkhagvatseren", 300_000, config.INCOME, config.PERSONAL)

    def test_8_client_payment_business_income(self):
        self.check("Client Payment", 1_500_000, config.INCOME, config.BUSINESS)


class TestClassificationRules(unittest.TestCase):
    """Priority дүрэм болон нэр танихын нэмэлт шалгалт."""

    def test_positive_is_income_negative_is_expense(self):
        for amount, expected in [(500_000, config.INCOME), (-500_000, config.EXPENSE),
                                 (1_500_000, config.INCOME), (-85_000, config.EXPENSE)]:
            self.assertEqual(classify_transaction("Anything", amount)["transaction_type"], expected)

    def test_davaadolgor_and_deleg_payments(self):
        r = classify_transaction("Payment Davaadolgor Dashdavaa", -450_000)
        self.assertEqual((r["category"], r["person"]), (SAL, "Davaadolgor Dashdavaa"))
        r = classify_transaction("Payment Deleg", -350_000)
        self.assertEqual((r["category"], r["person"]), (SAL, "Deleg"))

    def test_payment_to_unregistered_person(self):
        r = classify_transaction("Payment - John Smith", -500_000)
        self.assertEqual((r["finance_type"], r["category"]), (config.BUSINESS, SAL))

    def test_personal_rule_beats_payment_rule(self):
        r = classify_transaction("Payment Ganzorig Batzorig", -200_000)
        self.assertEqual((r["finance_type"], r["category"]), (config.PERSONAL, config.PERSONAL_CATEGORY))

    def test_name_case_truncation_and_typos(self):
        self.assertTrue(name_in_text("Fast Transfer From MUNKHCHIMEG LKHAGVATS Lend",
                                     "Munkhchimeg Lkhagvatseren"))
        self.assertTrue(name_in_text("payment DAVAADULAM DASHDAVAA", "Davaadulam Dashdavaa"))
        self.assertTrue(name_in_text("Payment Munkbat Ulziikhutag", "Munkhbat Ulziikhutag"))
        self.assertTrue(name_in_text("Batzorig Ganzorig transfer", "Ganzorig Batzorig"))
        self.assertTrue(name_in_text("pay DELEG", "Deleg"))

    def test_names_not_confused(self):
        self.assertFalse(name_in_text("Payment Davaadolgor Dashdavaa", "Davaadulam Dashdavaa"))
        self.assertFalse(name_in_text("Transfer To Chingerel CommBank App Anand Ganzorig",
                                      "Ganzorig Batzorig"))
        self.assertFalse(name_in_text("Payment Delger", "Deleg"))

    def test_modun_company_tracked_not_salary(self):
        for desc in ["Transfer To MODUN PTY LTD NetBank payment", "Payment Modun", "Modun Pty Ltd invoice"]:
            r = classify_transaction(desc, -1000)
            self.assertEqual(r["person"], "Modun Pty Ltd", desc)
            self.assertNotEqual(r["category"], SAL, desc)
        self.assertEqual(classify_transaction("Fast Transfer From MODUN PTY LTD", 5000)["person"],
                         "Modun Pty Ltd")

    def test_repayment_is_not_payment(self):
        r = classify_transaction("StepPay Repayment", -80)
        self.assertNotEqual(r["category"], SAL)

    def test_company_is_not_person(self):
        self.assertIsNone(extract_recipient_name("Transfer To Funstart pty ltd NetBank payment"))
        self.assertEqual(extract_recipient_name("Transfer To Munkhzaya NetBank payment"), "Munkhzaya")

    def test_keyword_categories(self):
        cases = {"Office Rent January": "Түрээс", "Shell Fuel Station": "Машин / Тээврийн хэрэгсэл",
                 "Uber trip": "Тээвэр", "Bunnings building material": "Материал",
                 "Telstra internet bill": "Utility / Төлбөр", "Allianz insurance": "Даатгал",
                 "Kids school fee": "Хүүхэд / Боловсрол", "Coles Supermarket": "Хоол",
                 "Random thing": config.OTHER_CATEGORY}
        for desc, cat in cases.items():
            self.assertEqual(classify_transaction(desc, -100)["category"], cat, desc)


class TestCSVProcessing(unittest.TestCase):
    """CSV унших, цэвэрлэх."""

    def test_parse_amount(self):
        self.assertEqual(parse_amount("1,500,000"), 1_500_000)
        self.assertEqual(parse_amount("-85,000"), -85_000)
        self.assertEqual(parse_amount("$1,234.50"), 1234.5)
        self.assertEqual(parse_amount("(500)"), -500)
        with self.assertRaises(ValueError):
            parse_amount("29/6/2025")

    def test_sample_csv_classification(self):
        df, mapping, warnings = process_csv(SAMPLE_CSV)
        self.assertEqual(len(df), 28)
        self.assertEqual(warnings, [])
        self.assertTrue(((df["amount"] > 0) == (df["transaction_type"] == config.INCOME)).all())
        self.assertTrue(((df["amount"] < 0) == (df["transaction_type"] == config.EXPENSE)).all())

    def test_split_income_expense_columns(self):
        csv = ("Огноо,Гүйлгээний тайлбар,Орлого,Зарлага,Дансны үлдэгдэл\n"
               "2025.03.01,Client Payment,\"1,000,000\",,1000000\n"
               "2025.03.02,Payment Deleg,,\"200,000\",800000\n").encode("utf-8")
        df, mapping, _ = process_csv(csv)
        self.assertEqual(df["amount"].tolist(), [1_000_000, -200_000])
        self.assertEqual(df["category"].iloc[1], SAL)

    def test_headerless_commbank_format(self):
        csv = b"27/6/2025,250,Fast Transfer From MUNKHCHIMEG LKHAGVATS Lend,643.62\n" \
              b"27/6/2025,-980,Transfer To Saruul Erdene-Ochir NetBank payment Lakemba,613.62\n"
        df, mapping, _ = process_csv(csv)
        self.assertEqual(df["date"].iloc[0], "2025-06-27")
        self.assertEqual(df["finance_type"].iloc[0], config.PERSONAL)
        self.assertEqual(df["category"].iloc[1], SAL)

    def test_errors(self):
        with self.assertRaises(CSVProcessingError):
            process_csv(b"")
        with self.assertRaises(CSVProcessingError):
            process_csv(b"Date,Description\n2025-01-01,Coffee\n")
        with self.assertRaises(CSVProcessingError):
            process_csv(b"Date,Description,Amount\nhello,Coffee,10\nworld,Tea,5\n")
        with self.assertRaises(CSVProcessingError):
            process_csv(b"Date,Description,Amount\n2025-01-01,Coffee,abc\n")


class TestDatabase(unittest.TestCase):
    """SQLite хадгалалт, duplicate, гараар засах."""

    def setUp(self):
        self.tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        self.tmp.close()
        self.path = self.tmp.name
        db.init_db(self.path)

    def tearDown(self):
        os.remove(self.path)

    def test_duplicate_upload_is_skipped(self):
        df, _, _ = process_csv(SAMPLE_CSV)
        self.assertEqual(db.insert_transactions(df, self.path), (28, 0))
        self.assertEqual(db.insert_transactions(df, self.path), (0, 28))
        self.assertEqual(len(db.get_transactions(self.path)), 28)

    def test_same_day_identical_transactions_kept(self):
        csv = b"Date,Description,Amount\n2025-01-01,Coffee,-5\n2025-01-01,Coffee,-5\n"
        df, _, _ = process_csv(csv)
        self.assertEqual(db.insert_transactions(df, self.path), (2, 0))
        self.assertEqual(db.insert_transactions(df, self.path), (0, 2))

    def test_manual_correction_saved(self):
        df, _, _ = process_csv(SAMPLE_CSV)
        db.insert_transactions(df, self.path)
        tx = db.get_transactions(self.path)
        row = tx[tx["description"] == "Supermarket"].iloc[0]
        db.update_transaction(int(row["id"]), {"finance_type": config.PERSONAL,
                                               "transaction_type": config.INCOME,
                                               "category": "Бусад", "person": "Test"}, self.path)
        new = db.get_transactions(self.path).set_index("id").loc[row["id"]]
        self.assertEqual(new["finance_type"], config.PERSONAL)
        self.assertEqual(new["transaction_type"], config.INCOME)
        self.assertEqual(new["amount"], 250_000)
        self.assertEqual(new["person"], "Test")
        self.assertEqual(new["manual_override"], 1)

    def test_uploaded_file_registry(self):
        df, _, _ = process_csv(SAMPLE_CSV, "sample_transactions.csv")
        ins, dup = db.insert_transactions(df, self.path)
        db.record_uploaded_file("sample_transactions.csv", "abc123", None, len(df), ins, dup, self.path)
        self.assertEqual(db.find_uploaded_file("abc123", self.path)["inserted"], 28)
        self.assertIsNone(db.find_uploaded_file("other", self.path))
        file_id = int(db.get_uploaded_files(self.path).iloc[0]["id"])
        self.assertEqual(db.delete_uploaded_file(file_id, self.path)[0], 28)
        self.assertTrue(db.get_transactions(self.path).empty)
        self.assertTrue(db.get_uploaded_files(self.path).empty)

    def test_read_all_tables(self):
        df, _, _ = process_csv(SAMPLE_CSV)
        db.insert_transactions(df, self.path)
        counts = db.table_counts(self.path)
        self.assertEqual(counts["transactions"], 28)
        self.assertEqual(counts["settings"], 1)
        for table in db.DB_TABLES:
            self.assertEqual(len(db.read_table(table, self.path)), counts[table])
        self.assertIn("manual_override", db.read_table("transactions", self.path).columns)
        with self.assertRaises(db.DatabaseError):
            db.read_table("sqlite_master", self.path)

    def test_settings_and_categories(self):
        db.save_settings({"gst_rate": 12.5, "gst_mode": "B"}, self.path)
        self.assertEqual(db.get_settings(self.path)["gst_rate"], 12.5)
        self.assertTrue(db.add_category("Маркетинг", self.path))
        self.assertIn("Маркетинг", db.get_categories(self.path))


class TestTaxAndReports(unittest.TestCase):
    """GST, татвар, цэвэр ашиг, хүн тус бүрийн тайлан — sample CSV дээр."""

    @classmethod
    def setUpClass(cls):
        df, _, _ = process_csv(SAMPLE_CSV)
        cls.df = df.assign(id=range(1, len(df) + 1))
        cls.company = {"gst_rate": 10.0, "gst_mode": "A", "entity_type": config.ENTITY_COMPANY,
                       "tax_rate": 25.0, "tax_year": "2024-25"}

    def test_business_and_personal_separate(self):
        s = reports.summarize(self.df, self.company)
        self.assertAlmostEqual(s["business_income"], 12_300_000)
        self.assertAlmostEqual(s["business_expense"], 6_182_500)
        self.assertAlmostEqual(s["personal_income"], 550_000)
        self.assertAlmostEqual(s["personal_expense"], 1_800_000)
        self.assertEqual(s["current_balance"], 6_867_500)

    def test_gst_modes(self):
        self.assertAlmostEqual(calculate_gst(12_300_000, 10, "A"), 1_230_000)
        self.assertAlmostEqual(calculate_gst(12_300_000, 10, "B"), 12_300_000 / 11)
        self.assertEqual(calculate_gst(-5, 10, "A"), 0)

    def test_net_profit_company(self):
        s = reports.summarize(self.df, self.company)
        self.assertAlmostEqual(s["estimated_gst"], 1_230_000)
        self.assertAlmostEqual(s["profit_before_tax"], 4_887_500)
        self.assertAlmostEqual(s["estimated_tax"], 1_221_875)
        self.assertAlmostEqual(s["net_profit"], 3_665_625)

    def test_individual_progressive_tax(self):
        self.assertAlmostEqual(calculate_individual_tax(100_000, "2024-25"), 20_788)
        self.assertAlmostEqual(calculate_individual_tax(18_000, "2024-25"), 0)
        self.assertAlmostEqual(calculate_individual_tax(100_000, "2026-27"), 20_520)
        settings = {**self.company, "entity_type": config.ENTITY_INDIVIDUAL}
        self.assertAlmostEqual(calculate_income_tax(100_000, settings), 20_788)

    def test_custom_rate_and_loss(self):
        settings = {**self.company, "entity_type": config.ENTITY_CUSTOM, "tax_rate": 20}
        self.assertAlmostEqual(calculate_income_tax(1000, settings), 200)
        p = profit_breakdown(100, 500, settings)
        self.assertEqual(p["estimated_tax"], 0)
        self.assertAlmostEqual(p["net_profit"], 100 - 500 - 10)

    def test_person_report(self):
        m = reports.person_summary(self.df, "Munkhbat Ulziikhutag")
        self.assertEqual((m["given"], m["received"], m["difference"], m["count"]),
                         (1_000_000, 200_000, -800_000, 3))
        d = reports.person_summary(self.df, "Deleg")
        self.assertEqual((d["given"], d["received"], d["count"]), (350_000, 100_000, 2))
        self.assertEqual(reports.person_summary(self.df, "Davaadulam Dashdavaa")["count"], 2)
        self.assertEqual(reports.person_summary(self.df, "Davaadolgor Dashdavaa")["count"], 1)

    def test_focus_party_income_expense_split(self):
        import pandas as pd
        extra = pd.DataFrame([
            {"id": 900, "date": "2025-03-01", "description": "Transfer To MODUN PTY LTD NetBank",
             "amount": -700.0, "transaction_type": config.EXPENSE, "finance_type": config.BUSINESS,
             "category": "Бусад", "person": "Modun Pty Ltd", "balance": None},
            {"id": 901, "date": "2025-03-05", "description": "Fast Transfer From Modun Pty Ltd",
             "amount": 2000.0, "transaction_type": config.INCOME, "finance_type": config.BUSINESS,
             "category": "Орлого", "person": "Modun Pty Ltd", "balance": None},
        ])
        s = reports.person_summary(pd.concat([self.df, extra], ignore_index=True), "Modun Pty Ltd")
        self.assertEqual((s["received"], s["given"], s["difference"], s["count"]), (2000, 700, 1300, 2))

    def test_export_columns(self):
        out = reports.export_dataframe(self.df)
        self.assertEqual(list(out.columns), ["Date", "Description", "Business/Personal",
                                             "Income/Expense", "Category", "Person",
                                             "Amount", "Balance"])

    def test_report_periods(self):
        start, end = reports.period_range("Сар", "2025-01-15")
        self.assertEqual((str(start.date()), str(end.date())), ("2025-01-01", "2025-01-31"))
        start, end = reports.period_range("7 хоног", "2025-01-10")
        self.assertEqual(str(start.date()), "2025-01-04")
        self.assertEqual(len(reports.filter_by_date(self.df, *reports.period_range("Өдөр", "2025-01-05"))), 2)



# ---------------------------------------------------------------------------
# financial_analysis.py — Excel report
# ---------------------------------------------------------------------------
import contextlib
import io

import financial_analysis as fa


def run_analysis(csv_text):
    """Write csv_text to a temp file, run the analysis quietly, return results."""
    folder = tempfile.mkdtemp()
    src = os.path.join(folder, "bank.csv")
    with open(src, "w", encoding="utf-8") as fh:
        fh.write(csv_text)
    with contextlib.redirect_stdout(io.StringIO()):
        return fa.analyse(src, os.path.join(folder, "report.xlsx"))


BANK_CSV = """Date,Description,Amount,Balance
01/02/2025,Direct Credit 348524 RICHARD CROOKES RICHARD CROOKES,"5,000.00",5000.00
03/02/2025,Fast Transfer From ODSUREN LKHAGVAJAV GST 73 653 557 147,"11,000.00",16000.00
04/02/2025,Fast Transfer From MUNKHCHIMEG LKHAGVATS Lend,2000.00,18000.00
05/02/2025,Fast Transfer From MUNKHCHIMEG LKHAGVATS Transfer,500.00,18500.00
06/02/2025,Transfer To Munkhchimeg Lkhagvatseren NetBank,-300.00,18200.00
07/02/2025,Transfer To Munkhzaya NetBank payment,-1000.00,17200.00
08/02/2025,Transfer To Munkhbat Ulziikhutag NetBank oko,-400.00,16800.00
09/02/2025,Transfer To Munkhbat Ulziikhutag NetBank oko,-400.00,16400.00
10/02/2025,Transfer To Munkhbat Ulziikhutag NetBank oko,-400.00,16000.00
11/02/2025,Fast Transfer From DAVAADULAM DASHDAVAA Muugii back,3000.00,19000.00
12/02/2025,StepPay Repayment,-50.00,18950.00
12/02/2025,StepPay Repayment,-50.00,18900.00
13/02/2025,Transfer to xx8106 CommBank app Expenses,-100.00,18800.00
14/02/2025,TAX OFFICE PAYMENTS CommBank app BPAY 75556 2025 Q1 BAS,-1500.00,17300.00
15/02/2025,CLEANMAX TRADING PTY LTD MELROSE AU Card xx4394 Value Date: 13/02/2025,-110.00,17190.00
16/02/2025,WOOLWORTHS/LAKESIDE PDE AU Card xx4394 Value Date: 14/02/2025,-80.00,17110.00
17/02/2025,BUNNINGS 775000 MARSDEN PARK AU Card xx4394 Value Date: 15/02/2025,-55.00,17055.00
18/02/2025,Payment - John Smith,-750.00,16305.00
"""


class TestFinancialAnalysis(unittest.TestCase):
    """Classification rules, GST, profit and validation of the Excel report."""

    @classmethod
    def setUpClass(cls):
        cls.r = run_analysis(BANK_CSV)
        cls.df = cls.r["df"].set_index("Original Description")

    def row(self, description):
        return self.df.loc[description] if not isinstance(self.df.loc[description], type(self.df)) \
            else self.df.loc[description].iloc[0]

    def test_validation_totals_match(self):
        r = self.r
        self.assertTrue(r["ok"])
        self.assertEqual((r["income_difference"], r["expense_difference"]), (0, 0))
        self.assertEqual(r["csv_rows"], r["processed_rows"])
        self.assertEqual(r["csv_rows"], 18)
        self.assertEqual(r["unclassified"], 0)
        self.assertEqual(r["duplicates"], 2)          # two StepPay rows: both kept
        self.assertEqual(r["balance_breaks"], 0)
        self.assertTrue(os.path.exists(r["output"]))

    def test_known_customer_income_gst_unknown(self):
        row = self.row("Direct Credit 348524 RICHARD CROOKES RICHARD CROOKES")
        self.assertEqual((row["Classification"], row["Person / Company"], row["GST Status"]),
                         (fa.BUS_INC, "Richard Crookes", fa.GST_UNKNOWN))
        self.assertEqual(row["GST Amount"], 0)

    def test_gst_inclusive_is_divided_by_11(self):
        row = self.row("Fast Transfer From ODSUREN LKHAGVAJAV GST 73 653 557 147")
        self.assertEqual(row["GST Status"], fa.GST_INCLUDED)
        self.assertAlmostEqual(row["GST Amount"], 1000.00)
        self.assertAlmostEqual(row["Amount excl GST"], 10000.00)

    def test_director_transactions_not_business(self):
        self.assertEqual(self.row("Fast Transfer From MUNKHCHIMEG LKHAGVATS Lend")["Classification"], fa.LOAN)
        self.assertEqual(self.row("Fast Transfer From MUNKHCHIMEG LKHAGVATS Transfer")["Classification"], fa.DIRECTOR)
        paid = self.row("Transfer To Munkhchimeg Lkhagvatseren NetBank")
        self.assertEqual((paid["Classification"], paid["Review Required"]), (fa.DIRECTOR, "YES"))

    def test_salary_and_possible_salary(self):
        self.assertEqual(self.row("Transfer To Munkhzaya NetBank payment")["Classification"], fa.SALARY)
        self.assertEqual(self.row("Payment - John Smith")["Classification"], fa.SALARY)
        possible = self.row("Transfer To Munkhbat Ulziikhutag NetBank oko")
        self.assertEqual((possible["Classification"], possible["Suggested Classification"]), (fa.REVIEW, fa.SALARY))

    def test_incoming_individual_is_possible_loan(self):
        row = self.row("Fast Transfer From DAVAADULAM DASHDAVAA Muugii back")
        self.assertEqual(row["Classification"], fa.REVIEW)
        self.assertIn(fa.R_LOAN, row["Reason for Review"])

    def test_other_classes(self):
        self.assertEqual(self.row("StepPay Repayment")["Classification"], fa.LOAN)
        self.assertEqual(self.row("Transfer to xx8106 CommBank app Expenses")["Classification"], fa.TRANSFER)
        self.assertEqual(self.row("TAX OFFICE PAYMENTS CommBank app BPAY 75556 2025 Q1 BAS")["Classification"],
                         fa.GST)
        cleanmax = self.row("CLEANMAX TRADING PTY LTD MELROSE AU Card xx4394 Value Date: 13/02/2025")
        self.assertEqual((cleanmax["Classification"], cleanmax["Tax Deductible"]), (fa.BUS_EXP, "YES"))
        self.assertAlmostEqual(cleanmax["GST Amount"], 10.00)
        self.assertEqual(self.row("WOOLWORTHS/LAKESIDE PDE AU Card xx4394 Value Date: 14/02/2025")["Classification"],
                         fa.PER_EXP)
        bunnings = self.row("BUNNINGS 775000 MARSDEN PARK AU Card xx4394 Value Date: 15/02/2025")
        self.assertEqual((bunnings["Classification"], bunnings["GST Amount"]), (fa.REVIEW, 0))

    def test_profit_excludes_personal_loans_transfers(self):
        df = self.r["df"]
        income_ex = df.loc[df.Classification == fa.BUS_INC, "Amount excl GST"].sum()
        self.assertAlmostEqual(income_ex, 5000 + 10000)
        expenses_ex = -df.loc[(df.Classification.isin([fa.BUS_EXP, fa.SALARY])), "Amount excl GST"].sum()
        self.assertAlmostEqual(expenses_ex, 100 + 1000 + 750)

    def test_debit_credit_columns(self):
        r = run_analysis("Transaction Date,Details,Debit,Credit\n"
                         "2025-03-01,Client invoice 55,,1100.00\n2025-03-02,Coles Ryde,45.50,\n")
        self.assertTrue(r["ok"])
        self.assertEqual((r["csv_in"], r["csv_out"]), (1100.0, 45.5))
        self.assertEqual(list(r["df"]["Direction"]), [fa.INCOME, fa.EXPENSE])

    def test_unreadable_rows_are_kept(self):
        r = run_analysis("Date,Description,Amount\n01/03/2025,Coffee,-5\nbad date,Tea,abc\n")
        self.assertEqual(r["processed_rows"], 2)
        self.assertIn(fa.R_UNREADABLE, r["df"]["Reason for Review"].iloc[-1])


if __name__ == "__main__":
    unittest.main(verbosity=2)
