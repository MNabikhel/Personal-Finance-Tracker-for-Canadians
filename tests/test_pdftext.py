"""Runs the PDF statement parser over statement text, inside LibreOffice Basic.

A PDF statement reaches ``modPdfText`` as lines of text - what Power Query or
Word recovered from the page - and the module has to find the transactions in
them without knowing which bank printed the page.  These tests hand it text in
the shapes Canadian statements actually take: a credit card with transaction
and posting dates, ``CR`` credits and a year printed only in the header; a
chequing account with a running balance, a date printed once for a day's
transactions and no sign on the amounts; and the summary lines, totals and
notices that surround both.

Getting the text out of the PDF is Excel's job (``modPdf``) and cannot run
here; what is under test is everything after that.
"""

from __future__ import annotations

import os
import sys
import unittest
from typing import List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests import libreoffice, vbahost  # noqa: E402

MODULES = vbahost.IMPORT_MODULES + ["modPdfText"]

CARD_STATEMENT = """\
RBC Royal Bank
Visa Infinite Avion
STATEMENT FROM FEB 20 TO MAR 19, 2026
Your account number: 4514 XXXX XXXX 1234
Previous Statement Balance $1,203.45
Payments & credits -$1,215.45
Purchases & debits $151.33
Minimum payment $10.00
Payment due date APR 10, 2026
Credit limit $12,000.00
TRANSACTION POSTING ACTIVITY DESCRIPTION AMOUNT ($)
DATE DATE
FEB 21 FEB 23 PAYMENT - THANK YOU / PAIEMENT - MERCI -1,203.45
FEB 22 FEB 23 TIM HORTONS #3324 TORONTO ON $5.80
MAR 03 MAR 04 LOBLAWS #4861 TORONTO ON 99.09
MAR 05 MAR 06 AMAZON.CA*2K4XY1 AMAZON.CA ON 23.45
MAR 10 MAR 11 RETURN LOBLAWS #4861 TORONTO ON 12.00 CR
MAR 15 MAR 16 NETFLIX.COM 866-579-7172 ON 22.99
Foreign currency transactions are converted at the rate on the posting date.
TOTAL PURCHASES 151.33
NEW BALANCE $139.33
Card expires MAR 2029
"""

CARD_EXPECTED = [
    ("2026-02-21", "1203.45", "PAYMENT - THANK YOU / PAIEMENT - MERCI"),
    ("2026-02-22", "-5.80", "TIM HORTONS #3324 TORONTO ON"),
    ("2026-03-03", "-99.09", "LOBLAWS #4861 TORONTO ON"),
    ("2026-03-05", "-23.45", "AMAZON.CA*2K4XY1 AMAZON.CA ON"),
    ("2026-03-10", "12.00", "RETURN LOBLAWS #4861 TORONTO ON"),
    ("2026-03-15", "-22.99", "NETFLIX.COM 866-579-7172 ON"),
]

# December purchases on a January statement belong to the year before.
YEAR_END_STATEMENT = """\
Statement period: December 20, 2025 to January 19, 2026
Payment due date: February 9, 2026
Minimum payment: $10.00
DEC 22 DEC 23 LCBO/RAO #0079 TORONTO ON 28.27
JAN 02 JAN 03 PRESTO FARE/TRANSIT TORONTO ON 101.66
"""

ACCOUNT_STATEMENT = """\
Royal Bank of Canada
Your account statement
From March 1, 2026 to March 31, 2026
Date Description Withdrawals ($) Deposits ($) Balance ($)
Opening Balance 1,000.00
03 Mar Payroll Deposit NORTHWIND LOGISTICS 2,483.18 3,483.18
Contactless Interac purchase - 3324 TIM HORTONS 5.80 3,477.38
05 Mar Online Banking transfer - 1234 450.00 3,027.38
e-Transfer received JANE DOE 100.00 3,127.38
12 Mar Monthly fee 4.00 3,123.38
Closing Balance 3,123.38
Total withdrawals 459.80
Total deposits 2,583.18
Please review this statement and report any errors within 30 days.
"""

ACCOUNT_EXPECTED = [
    ("2026-03-03", "2483.18", "Payroll Deposit NORTHWIND LOGISTICS"),
    ("2026-03-03", "-5.80", "Contactless Interac purchase - 3324 TIM HORTONS"),
    ("2026-03-05", "-450.00", "Online Banking transfer - 1234"),
    ("2026-03-05", "100.00", "e-Transfer received JANE DOE"),
    ("2026-03-12", "-4.00", "Monthly fee"),
]

# Nothing on the page says what kind of statement it is; the running balance
# moving by each amount has to.
BARE_ACCOUNT_STATEMENT = """\
Statement date 2026-04-30
04/02 HYDRO ONE PREAUTHORIZED 97.76 902.24
04/03 CANADA FED / FED CCB 619.75 1,521.99
04/05 SPOTIFY P2C5 MISC PAYMENT 16.99 1,505.00
"""


# A reference number ahead of each date, and a trailing minus for credits.
SCOTIA_STATEMENT = """\
Scotiabank Visa
Statement date: March 19, 2026
Minimum payment due: $10.00
REF# TRANS DATE POST DATE DESCRIPTION AMOUNT
001 MAR 3 MAR 4 TIM HORTONS #3324 TORONTO ON 5.80
002 MAR 5 MAR 6 PAYMENT THANK YOU 200.00-
"""

# French: month names, "1 234,56" amounts and "Solde" for the balance lines.
FRENCH_STATEMENT = """\
Relevé de compte
Période du 1 mars 2026 au 31 mars 2026
Solde précédent 1 000,00
03 mars Dépôt salaire 2 483,18 3 483,18
05 mars Achat TIM HORTONS 5,80 3 477,38
Solde final 3 477,38
"""


def setUpModule():
    try:
        libreoffice.context()
    except libreoffice.Unavailable as exc:  # pragma: no cover - environment
        raise unittest.SkipTest(str(exc))


def _lines(text: str) -> str:
    """Basic that builds a Collection of the statement's lines."""
    return f"modPdfText.SplitLines({vbahost.basic_string(text)})"


def _run(body: str) -> List[List[str]]:
    return vbahost.rows(vbahost.run(body, MODULES))


def _read(text: str, kind: str = "") -> List[List[str]]:
    """Every transaction read from ``text``: date, amount, description."""
    kind_expr = vbahost.basic_string(kind) if kind else \
        "modPdfText.DetectKind(lines, anchorYear, anchorMonth)"
    body = f'''
Sub Run()
    Dim lines As Collection, records As Collection
    Dim anchorYear As Long, anchorMonth As Long
    Dim readCount As Long, badCount As Long
    Dim txn As clsTxn
    Dim i As Long
    Set lines = {_lines(text)}
    anchorYear = modPdfText.StatementAnchor(lines, anchorMonth)
    Set records = modPdfText.ReadStatement(lines, {kind_expr}, anchorYear, _
                                           anchorMonth, readCount, badCount)
    Emit "counts", readCount, badCount
    For i = 1 To records.Count
        Set txn = records.Item(i)
        Emit Stamp(txn.TxnDate), Money(txn.Amount), txn.Description, txn.Merchant
    Next i
End Sub
'''
    return _run(body)


class CardStatementTests(unittest.TestCase):
    def setUp(self):
        self.rows = _read(CARD_STATEMENT)
        self.counts, self.records = self.rows[0], self.rows[1:]

    def test_every_transaction_line_is_read_and_nothing_else(self):
        self.assertEqual([tuple(row[:3]) for row in self.records], CARD_EXPECTED)

    def test_charges_are_money_out_and_credits_money_in(self):
        amounts = {row[2]: row[1] for row in self.records}
        self.assertEqual(amounts["TIM HORTONS #3324 TORONTO ON"], "-5.80")
        self.assertEqual(amounts["RETURN LOBLAWS #4861 TORONTO ON"], "12.00")
        self.assertEqual(amounts["PAYMENT - THANK YOU / PAIEMENT - MERCI"], "1203.45")

    def test_the_year_comes_from_the_statement_not_the_card_expiry(self):
        # "MAR 2029" is on the page too; a month and year alone must not count.
        self.assertTrue(all(row[0].startswith("2026-") for row in self.records))

    def test_merchants_are_cleaned_as_the_csv_import_cleans_them(self):
        body = "\n".join(
            ["Sub Run()"]
            + [f"    Emit {vbahost.basic_string(row[2])}, "
               f"modRules.CleanMerchant({vbahost.basic_string(row[2])})"
               for row in self.records]
            + ["End Sub"])
        cleaned = {row[0]: row[1] for row in _run(body)}
        for row in self.records:
            with self.subTest(row[2]):
                self.assertEqual(row[3], cleaned[row[2]])
        self.assertNotEqual(cleaned["TIM HORTONS #3324 TORONTO ON"],
                            "TIM HORTONS #3324 TORONTO ON")

    def test_the_counts_describe_the_lines(self):
        # Six lines began with a date and carried money; none was unreadable.
        self.assertEqual(self.counts, ["counts", "6", "0"])

    def test_it_is_recognised_as_a_card_statement(self):
        body = f'''
Sub Run()
    Dim lines As Collection
    Dim anchorMonth As Long
    Set lines = {_lines(CARD_STATEMENT)}
    Emit modPdfText.DetectKind(lines, 2026, 4)
End Sub
'''
        self.assertEqual(_run(body)[0], ["Credit card"])


class YearEndTests(unittest.TestCase):
    def test_december_lines_on_a_january_statement_are_last_year(self):
        rows = _read(YEAR_END_STATEMENT)[1:]
        self.assertEqual([(row[0], row[1]) for row in rows],
                         [("2025-12-22", "-28.27"), ("2026-01-02", "-101.66")])


class AccountStatementTests(unittest.TestCase):
    def setUp(self):
        self.rows = _read(ACCOUNT_STATEMENT)
        self.counts, self.records = self.rows[0], self.rows[1:]

    def test_every_transaction_is_read_with_its_sign_from_the_balance(self):
        self.assertEqual([tuple(row[:3]) for row in self.records], ACCOUNT_EXPECTED)

    def test_a_line_without_a_date_takes_the_date_of_the_line_above(self):
        by_description = {row[2]: row[0] for row in self.records}
        self.assertEqual(by_description["Contactless Interac purchase - 3324 TIM HORTONS"],
                         "2026-03-03")
        self.assertEqual(by_description["e-Transfer received JANE DOE"], "2026-03-05")

    def test_balances_and_totals_are_not_transactions(self):
        descriptions = [row[2] for row in self.records]
        for word in ("Opening", "Closing", "Total"):
            self.assertFalse(any(word in d for d in descriptions), descriptions)

    def test_it_is_recognised_as_an_account_statement(self):
        body = f'''
Sub Run()
    Dim lines As Collection
    Set lines = {_lines(ACCOUNT_STATEMENT)}
    Emit modPdfText.DetectKind(lines, 2026, 3)
End Sub
'''
        self.assertEqual(_run(body)[0], ["Bank account"])


class BareStatementTests(unittest.TestCase):
    """No words give the kind away, and the dates are 04/02 with no year."""

    def test_the_running_balance_identifies_an_account_statement(self):
        body = f'''
Sub Run()
    Dim lines As Collection
    Set lines = {_lines(BARE_ACCOUNT_STATEMENT)}
    Emit modPdfText.DetectKind(lines, 2026, 4)
End Sub
'''
        self.assertEqual(_run(body)[0], ["Bank account"])

    def test_numeric_dates_and_balance_signs(self):
        rows = _read(BARE_ACCOUNT_STATEMENT)[1:]
        self.assertEqual([(row[0], row[1]) for row in rows],
                         [("2026-04-02", "-97.76"), ("2026-04-03", "619.75"),
                          ("2026-04-05", "-16.99")])

    def test_read_as_a_card_statement_the_same_lines_are_all_charges(self):
        # What the user gets when they answer "No" to the preview and ask for
        # the other reading: every line's first amount, as money out.
        rows = _read(BARE_ACCOUNT_STATEMENT, "Credit card")[1:]
        self.assertEqual([row[1] for row in rows], ["-97.76", "-619.75", "-16.99"])


class OtherBanksTests(unittest.TestCase):
    def test_a_reference_number_before_the_date_is_stepped_over(self):
        rows = _read(SCOTIA_STATEMENT)[1:]
        self.assertEqual([tuple(row[:3]) for row in rows],
                         [("2026-03-03", "-5.80", "TIM HORTONS #3324 TORONTO ON"),
                          ("2026-03-05", "200.00", "PAYMENT THANK YOU")])

    def test_a_french_statement_is_read(self):
        rows = _read(FRENCH_STATEMENT)
        self.assertEqual([tuple(row[:3]) for row in rows[1:]],
                         [("2026-03-03", "2483.18", "Dépôt salaire"),
                          ("2026-03-05", "-5.80", "Achat TIM HORTONS")])

    def test_a_french_statement_is_recognised_by_its_balance(self):
        body = f'''
Sub Run()
    Dim lines As Collection
    Set lines = {_lines(FRENCH_STATEMENT)}
    Emit modPdfText.DetectKind(lines, 2026, 3)
End Sub
'''
        self.assertEqual(_run(body)[0], ["Bank account"])

    def test_payment_apps_and_debit_interest_are_not_bank_deposits(self):
        # On account PDFs without a running balance, wording is the only sign
        # clue. The word "pay" in Apple Pay/Google Pay and "interest" in an
        # overdraft charge must not turn spending into income.
        rows = _read(
            "Statement date: September 30, 2026\n"
            "09/03 APPLE PAY GROCERY 25.00\n"
            "09/04 GOOGLE PAY TRANSIT 12.50\n"
            "09/05 OVERDRAFT INTEREST 4.25",
            "Bank account",
        )[1:]
        self.assertEqual(
            [(row[2], row[1]) for row in rows],
            [
                ("APPLE PAY GROCERY", "-25.00"),
                ("GOOGLE PAY TRANSIT", "-12.50"),
                ("OVERDRAFT INTEREST", "-4.25"),
            ],
        )


class AnchorTests(unittest.TestCase):
    def _anchor(self, text: str) -> List[str]:
        body = f'''
Sub Run()
    Dim lines As Collection
    Dim anchorMonth As Long, anchorYear As Long
    Set lines = {_lines(text)}
    anchorYear = modPdfText.StatementAnchor(lines, anchorMonth)
    Emit anchorYear, anchorMonth
End Sub
'''
        return _run(body)[0]

    def test_the_latest_full_date_on_the_page_wins(self):
        self.assertEqual(self._anchor(CARD_STATEMENT), ["2026", "4"])

    def test_numeric_and_day_first_dates_are_read(self):
        self.assertEqual(self._anchor("Statement date 2026-04-30"), ["2026", "4"])
        self.assertEqual(self._anchor("Issued 30 April 2026"), ["2026", "4"])
        self.assertEqual(self._anchor("Issued 04/30/2026"), ["2026", "4"])

    def test_a_page_with_no_full_date_says_so(self):
        # The caller then asks the user for the year.
        self.assertEqual(self._anchor("MAR 03 MAR 04 LOBLAWS 99.09\nMAR 2029"), ["0", "0"])


class TokenTests(unittest.TestCase):
    def test_money_tokens(self):
        cases = {
            "5.80": True, "$5.80": True, "-1,203.45": True, "-$1,203.45": True,
            "1,203.45-": True, "(12.00)": True, "1234,56": True, "12,000.00": True,
            "3324": False, "#3324": False, "2026": False, "2026-03-03": False,
            "AMAZON.CA": False, "1.3": False, "10.5": False, "19.99%": False,
            "866-579-7172": False, "03/03": False,
        }
        body = "\n".join(
            ["Sub Run()"]
            + [f"    Emit {vbahost.basic_string(token)}, "
               f"Flag(modPdfText.IsMoneyToken({vbahost.basic_string(token)}))"
               for token in cases]
            + ["End Sub"])
        got = {row[0]: row[1] == "yes" for row in _run(body)}
        self.assertEqual(got, cases)

    def test_leading_dates(self):
        # line -> the date read from its start and how many words it took.
        cases = {
            "MAR 03 MAR 04 LOBLAWS": ("2026-03-03", "2"),
            "Mar. 3 LOBLAWS": ("2026-03-03", "2"),
            "03 Mar LOBLAWS": ("2026-03-03", "2"),
            "03 Mar 2025 LOBLAWS": ("2025-03-03", "3"),
            "Mar 3, 2025 LOBLAWS": ("2025-03-03", "3"),
            "03/03 LOBLAWS": ("2026-03-03", "1"),
            "2025-03-03 LOBLAWS": ("2025-03-03", "1"),
            "03/03/2025 LOBLAWS": ("2025-03-03", "1"),
            "DEC 30 LATE FEE": ("2025-12-30", "2"),
            "LOBLAWS 99.09": ("", "0"),
            "5.80 LOBLAWS": ("", "0"),
            "MARKET 5 LOBLAWS": ("", "0"),
            "Total 99.09": ("", "0"),
        }
        lines = ["Sub Run()", "    Dim tokens As Variant, used As Long, got As Date"]
        for line in cases:
            lines += [
                f"    tokens = Split({vbahost.basic_string(line)}, \" \")",
                "    got = modPdfText.TxnDateAt(tokens, 0, 2026, 4, used)",
                f"    If used = 0 Then",
                f"        Emit {vbahost.basic_string(line)}, \"\", used",
                "    Else",
                f"        Emit {vbahost.basic_string(line)}, Stamp(got), used",
                "    End If",
            ]
        lines.append("End Sub")
        got = {row[0]: (row[1], row[2]) for row in _run("\n".join(lines))}
        self.assertEqual(got, cases)

    def test_summary_lines(self):
        cases = {
            "Previous Statement Balance": True, "NEW BALANCE": True,
            "Opening Balance": True, "TOTAL PURCHASES": True,
            "Minimum payment": True, "Credit limit": True,
            "PAYMENT - THANK YOU": False, "TOTAL WINE & MORE": True,
            "LOBLAWS #4861": False, "": True,
        }
        body = "\n".join(
            ["Sub Run()"]
            + [f"    Emit {vbahost.basic_string(text)}, "
               f"Flag(modPdfText.IsSummary({vbahost.basic_string(text)}))"
               for text in cases]
            + ["End Sub"])
        got = {row[0]: row[1] == "yes" for row in _run(body)}
        self.assertEqual(got, cases)


if __name__ == "__main__":
    unittest.main()
