"""Checks the built .xlsm: its package, its shape, and what its formulas say.

Two different things are being verified here.  The package tests confirm the
file really is a macro-enabled workbook - the content types, the relationship
to vbaProject.bin, the tables and defined names the macros address by name -
because a workbook that Excel refuses to open, or that opens with the macros
detached, fails no matter how good the VBA is.

The recalculation tests then open the workbook in LibreOffice, throw away every
cached result and recompute from scratch, and compare what comes out with the
same totals worked out independently in Python from the sample transactions.
LibreOffice is a separate implementation of both OOXML and the spreadsheet
function library, so agreement is real evidence the formulas are right rather
than a restatement of them.
"""

from __future__ import annotations

import io
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import build as builder  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from tests import libreoffice  # noqa: E402
from tools import data, package, sample, workbook  # noqa: E402

TODAY = date(2026, 9, 1)
MONTH = workbook.report_month(TODAY)          # the month the workbook opens on

CENTS = 2

_blob: bytes = b""
_path: str = ""
_records: List[sample.Txn] = []


def setUpModule():
    global _blob, _path, _records
    _blob = builder.build_package(TODAY)
    _records = sample.build(TODAY)
    handle, _path = tempfile.mkstemp(prefix="cft-workbook-", suffix=".xlsm")
    with os.fdopen(handle, "wb") as stream:
        stream.write(_blob)


def tearDownModule():
    if _path and os.path.exists(_path):
        os.unlink(_path)


# --- What the sample transactions add up to, worked out independently -------

TYPE_OF = {row[0]: row[2] for row in data.CATEGORIES}
GROUP_OF = {row[0]: row[1] for row in data.CATEGORIES}
ESSENTIAL_OF = {row[0]: row[3] for row in data.CATEGORIES}

# Which account belongs to whom, as the Accounts sheet lists it.  This is what
# the "Paid By" column works out from the account the money actually left.
PAID_BY = {
    sample.ACCOUNT_ALEX: sample.PERSON_A,
    sample.ACCOUNT_SAM: sample.PERSON_B,
    sample.ACCOUNT_JOINT: "Joint",
    sample.ACCOUNT_CARD: "Joint",
}

DEFAULT_SPLIT_A = Decimal("0.5")


def _round(value: Decimal) -> Decimal:
    """A spreadsheet ROUND: to the cent, with ties away from zero.

    Not Python's round(), which breaks ties towards the even digit and so
    disagrees with the sheet by a cent on exactly the half-cent amounts that
    splitting a bill in two produces.
    """
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _month(records: List[sample.Txn]) -> List[sample.Txn]:
    return [txn for txn in records if txn.month == MONTH]


def _of_type(records: List[sample.Txn], kind: str) -> List[sample.Txn]:
    return [txn for txn in records if TYPE_OF.get(txn.category, "Expense") == kind]


def _amount(txn: sample.Txn) -> Decimal:
    return Decimal(f"{txn.amount:.2f}")


def _split_a(txn: sample.Txn) -> Decimal:
    if txn.owner == sample.PERSON_A:
        return Decimal(1)
    if txn.owner == sample.PERSON_B:
        return Decimal(0)
    return DEFAULT_SPLIT_A


def _share_a(txn: sample.Txn) -> Decimal:
    return _round(_amount(txn) * _split_a(txn))


def _share_b(txn: sample.Txn) -> Decimal:
    return _round(_amount(txn) - _share_a(txn))


def _sum(values) -> float:
    return float(_round(sum(values, Decimal(0))))


def _total(records: List[sample.Txn]) -> float:
    return _sum(_amount(txn) for txn in records)


# --- Reading a recalculated copy of the workbook ----------------------------


class Recalculated:
    """The workbook open in LibreOffice, recomputed from the formulas alone."""

    def __init__(self, path: str):
        self.document = libreoffice.open_document(path)
        self.recalculate()

    def recalculate(self) -> None:
        self.document.calculateAll()

    def close(self) -> None:
        self.document.close(True)

    def cell(self, sheet: str, ref: str):
        return self.document.Sheets.getByName(sheet).getCellRangeByName(ref)

    def number(self, sheet: str, ref: str) -> float:
        return round(self.cell(sheet, ref).getValue(), CENTS)

    def text(self, sheet: str, ref: str) -> str:
        return self.cell(sheet, ref).getString()

    def set(self, sheet: str, ref: str, value: str) -> None:
        self.cell(sheet, ref).setString(value)
        self.recalculate()


class PackageTests(unittest.TestCase):
    """The parts of the package that make it a macro-enabled workbook."""

    def setUp(self):
        self.described = package.describe(_blob)

    def test_it_is_a_readable_zip_with_no_broken_members(self):
        with zipfile.ZipFile(io.BytesIO(_blob)) as archive:
            self.assertIsNone(archive.testzip())

    def test_the_vba_project_is_present_and_declared(self):
        self.assertIn("xl/vbaProject.bin", self.described["names"])
        self.assertTrue(self.described["vba_project"],
                        "the project part must not be empty")
        self.assertTrue(self.described["bin_default"],
                        "a .bin content type is needed or Excel drops the project")
        self.assertTrue(self.described["vba_relationship"],
                        "the workbook part must point at the project")

    def test_the_workbook_part_is_macro_enabled(self):
        # Without this the file is a .xlsx wearing a .xlsm extension and Excel
        # refuses to open it at all.
        self.assertTrue(self.described["macro_content_type"])
        self.assertFalse(self.described["sheet_content_type"],
                         "the plain spreadsheet content type must be replaced")

    def test_two_builds_of_one_source_are_the_same_bytes(self):
        self.assertEqual(builder.build_package(TODAY), _blob)

    def test_the_document_is_dated_from_the_build_not_the_clock(self):
        with zipfile.ZipFile(io.BytesIO(_blob)) as archive:
            core = archive.read("docProps/core.xml").decode("utf-8")
        self.assertIn(f"{TODAY.isoformat()}T00:00:00Z", core)
        self.assertEqual(core.count(f"{TODAY.isoformat()}T00:00:00Z"), 2,
                         "created and modified should agree")

    def test_the_app_name_matches_the_one_the_macros_use(self):
        # It is the title of every message box the workbook shows, so the two
        # spellings drifting apart would be visible to the user.
        source = (os.path.join(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__))), "vba", "modConst.bas"))
        with open(source, encoding="utf-8") as stream:
            text = stream.read()
        self.assertIn(f'APP_NAME As String = "{workbook.APP_NAME}"', text)


class ShapeTests(unittest.TestCase):
    """The names the macros address the workbook by."""

    @classmethod
    def setUpClass(cls):
        cls.wb = load_workbook(io.BytesIO(_blob), data_only=False)

    def test_every_sheet_the_macros_name_exists(self):
        self.assertEqual(set(workbook.CODE_NAMES) - set(self.wb.sheetnames), set())

    def test_sheet_code_names_survive_the_round_trip(self):
        # The VBA project's module stream binds a document module to a sheet by
        # code name; if these are lost the sheet modules bind to nothing.
        got = {name: self.wb[name].sheet_properties.codeName
               for name in workbook.CODE_NAMES}
        self.assertEqual(got, workbook.CODE_NAMES)

    def test_the_tables_the_macros_read_are_all_there(self):
        wanted = {
            workbook.SH_TXN: "tblTxn",
            workbook.SH_ACCOUNTS: "tblAccounts",
            workbook.SH_CATEGORIES: "tblCategories",
            workbook.SH_RULES: "tblRules",
            workbook.SH_FORMATS: "tblFormats",
            workbook.SH_LOG: "tblLog",
            workbook.SH_ENGINE: "tblTemplates",
        }
        for sheet, table in wanted.items():
            with self.subTest(table):
                self.assertIn(table, self.wb[sheet].tables)

    def test_the_ledger_has_every_column_the_macros_write(self):
        table = self.wb[workbook.SH_TXN].tables["tblTxn"]
        self.assertEqual([column.name for column in table.tableColumns],
                         workbook.TXN_HEADERS)

    def test_the_defined_names_the_macros_read_are_all_there(self):
        for key in ("ReportMonth", "ReportView", "PersonA", "PersonB",
                    "DefaultSplitA", "HouseholdMode", "Configured",
                    "TransferWindowDays", "SkipDuplicates", "CategoryList",
                    "AccountList", "FormatList", "TopMerchants", "TaxYear"):
            with self.subTest(key):
                self.assertIn(key, self.wb.defined_names)

    def test_the_formula_templates_match_the_built_rows(self):
        # The macros copy these onto rows they add, so a template that has
        # drifted from the built rows would quietly produce two kinds of row.
        engine = self.wb[workbook.SH_ENGINE]
        templates = {}
        for row in range(4, 4 + len(workbook.TXN_FORMULAS)):
            templates[engine[f"A{row}"].value] = engine[f"B{row}"].value
        self.assertEqual(templates, {header: workbook.formula_r1c1(header)
                                     for header in workbook.TXN_FORMULAS})

    def test_a_row_needing_a_category_is_flagged_across_its_whole_width(self):
        # A cellIs rule compares the cell being formatted, so spreading one
        # across the ledger only ever colours Category itself - which looks
        # right in the builder and does nothing on the sheet.
        ws = self.wb[workbook.SH_TXN]
        category = workbook.col_of("Category")
        first = workbook.col_of(workbook.TXN_HEADERS[0])
        last = workbook.col_of(workbook.TXN_HEADERS[-1])

        for ranges, rules in ws.conditional_formatting._cf_rules.items():
            if str(ranges.sqref).startswith(f"{first}") and last in str(ranges.sqref):
                for rule in rules:
                    if rule.type == "expression":
                        self.assertEqual(
                            rule.formula,
                            [f'${category}{workbook.TXN_FIRST_ROW + 1}'
                             f'="Uncategorized"'])
                        return
        self.fail("no whole-row rule for uncategorised transactions")

    def test_the_sample_data_is_actually_in_the_ledger(self):
        ws = self.wb[workbook.SH_TXN]
        first = workbook.TXN_FIRST_ROW + 1
        column = workbook.col_of("Amount")
        amounts = [ws[f"{column}{first + offset}"].value
                   for offset in range(len(_records))]
        self.assertEqual(amounts, [txn.amount for txn in _records])


class RecalculationTests(unittest.TestCase):
    """Every number on the Dashboard, recomputed from the formulas."""

    sheet = workbook.SH_DASHBOARD

    @classmethod
    def setUpClass(cls):
        try:
            libreoffice.context()
        except libreoffice.Unavailable as exc:  # pragma: no cover - environment
            raise unittest.SkipTest(str(exc))
        cls.book = Recalculated(_path)

    @classmethod
    def tearDownClass(cls):
        cls.book.close()

    def test_it_opens_on_the_last_complete_month(self):
        self.assertEqual(self.book.text(self.sheet, "C6"), MONTH)
        self.assertEqual(self.book.text(self.sheet, "F6"), "Household")

    def test_money_in_out_and_saved(self):
        month = _month(_records)
        self.assertEqual(self.book.number(self.sheet, "C9"),
                         _total(_of_type(month, "Income")))
        self.assertEqual(self.book.number(self.sheet, "C10"),
                         -_total(_of_type(month, "Expense")))
        self.assertEqual(self.book.number(self.sheet, "C11"),
                         -_total(_of_type(month, "Saving")))

    def test_transfers_are_left_out_of_both_sides(self):
        # A credit card payment moves money without spending it.  Counting it
        # would double every card purchase, so it must appear in neither total.
        month = _month(_records)
        transfers = _of_type(month, "Transfer")
        self.assertTrue(transfers, "the sample month is meant to contain transfers")
        counted = (self.book.number(self.sheet, "C9")
                   + self.book.number(self.sheet, "C10")
                   + self.book.number(self.sheet, "C11"))
        everything = abs(_total(month))
        self.assertNotAlmostEqual(counted, everything, places=2)

    def test_net_cash_flow_and_savings_rate(self):
        money_in = self.book.number(self.sheet, "C9")
        money_out = self.book.number(self.sheet, "C10")
        saved = self.book.number(self.sheet, "C11")
        self.assertEqual(self.book.number(self.sheet, "C12"),
                         round(money_in - money_out - saved, CENTS))
        self.assertAlmostEqual(self.book.cell(self.sheet, "C13").getValue(),
                               (money_in - money_out) / money_in, places=6)

    def test_essential_and_discretionary_split_the_spending(self):
        month = _of_type(_month(_records), "Expense")
        essential = [txn for txn in month
                     if ESSENTIAL_OF.get(txn.category) == "Yes"]
        self.assertEqual(self.book.number(self.sheet, "F9"), -_total(essential))
        self.assertEqual(
            self.book.number(self.sheet, "F10"),
            round(self.book.number(self.sheet, "C10")
                  - self.book.number(self.sheet, "F9"), CENTS))

    def test_the_transaction_count(self):
        self.assertEqual(self.book.number(self.sheet, "F12"), len(_month(_records)))

    def test_nothing_is_left_uncategorised(self):
        self.assertEqual(self.book.number(self.sheet, "I6"), 0)

    def test_spending_by_group_accounts_for_everything_that_left(self):
        # The breakdown covers every group except Income and Transfers, so it
        # takes in savings as well as spending - money put into an RRSP did go
        # somewhere.  What it must not do is lose or double count a group.
        month = _month(_records)
        by_group: Dict[str, List[sample.Txn]] = {}
        for txn in month:
            if TYPE_OF.get(txn.category, "Expense") in ("Income", "Transfer"):
                continue
            by_group.setdefault(GROUP_OF.get(txn.category, "Other"), []).append(txn)

        got: Dict[str, float] = {}
        for offset, group in enumerate(data.SPENDING_GROUPS):
            got[group] = self.book.number(self.sheet, f"C{17 + offset}")
            with self.subTest(group):
                self.assertEqual(got[group], -_total(by_group.get(group, [])))

        self.assertEqual(
            round(sum(got.values()), CENTS),
            round(self.book.number(self.sheet, "C10")
                  + self.book.number(self.sheet, "C11"), CENTS),
            "the groups must account for the spending and the saving together")


class ReportsTests(unittest.TestCase):
    """The twelve months behind the Dashboard."""

    sheet = workbook.SH_REPORTS

    @classmethod
    def setUpClass(cls):
        try:
            libreoffice.context()
        except libreoffice.Unavailable as exc:  # pragma: no cover - environment
            raise unittest.SkipTest(str(exc))
        cls.book = Recalculated(_path)

    @classmethod
    def tearDownClass(cls):
        cls.book.close()

    def _column(self, index: int) -> str:
        from openpyxl.utils import get_column_letter
        return get_column_letter(workbook.REPORT_FIRST_COL + index)

    def test_the_twelve_months_end_on_the_report_month(self):
        months = [self.book.text(self.sheet, f"{self._column(index)}4")
                  for index in range(workbook.REPORT_MONTHS)]
        self.assertEqual(months[-1], MONTH)
        self.assertEqual(len(set(months)), workbook.REPORT_MONTHS)
        self.assertEqual(months, sorted(months), "months must read left to right")

    def test_each_month_column_agrees_with_the_transactions(self):
        for index in range(workbook.REPORT_MONTHS):
            column = self._column(index)
            month = self.book.text(self.sheet, f"{column}4")
            rows = [txn for txn in _records if txn.month == month]
            with self.subTest(month):
                self.assertEqual(self.book.number(self.sheet, f"{column}5"),
                                 _total(_of_type(rows, "Income")))
                self.assertEqual(self.book.number(self.sheet, f"{column}6"),
                                 -_total(_of_type(rows, "Expense")))
                self.assertEqual(self.book.number(self.sheet, f"{column}12"),
                                 len(rows))

    def test_the_total_column_sums_the_twelve(self):
        from openpyxl.utils import get_column_letter
        total = get_column_letter(workbook.REPORT_FIRST_COL + workbook.REPORT_MONTHS)
        for row in (5, 6, 7):
            with self.subTest(row=row):
                months = sum(self.book.number(self.sheet, f"{self._column(i)}{row}")
                             for i in range(workbook.REPORT_MONTHS))
                self.assertEqual(self.book.number(self.sheet, f"{total}{row}"),
                                 round(months, CENTS))


class CoupleModeTests(unittest.TestCase):
    """The part that only matters when two people share the workbook."""

    @classmethod
    def setUpClass(cls):
        try:
            libreoffice.context()
        except libreoffice.Unavailable as exc:  # pragma: no cover - environment
            raise unittest.SkipTest(str(exc))
        cls.book = Recalculated(_path)

    @classmethod
    def tearDownClass(cls):
        cls.book.close()

    def test_paid_by_follows_the_account_not_the_owner(self):
        # Someone can pay for something that is not theirs; that is the whole
        # reason there is anything to settle.
        month = _of_type(_month(_records), "Expense")
        for person, column in ((sample.PERSON_A, "C"), (sample.PERSON_B, "D")):
            paid = [txn for txn in month if PAID_BY[txn.account] == person]
            with self.subTest(person):
                self.assertEqual(self.book.number(workbook.SH_HOUSEHOLD, f"{column}7"),
                                 -_total(paid))

    def test_a_fair_share_divides_joint_spending_by_the_household_split(self):
        month = _of_type(_month(_records), "Expense")
        personal = [txn for txn in month
                    if PAID_BY[txn.account] in (sample.PERSON_A, sample.PERSON_B)]
        self.assertEqual(self.book.number(workbook.SH_HOUSEHOLD, "C8"),
                         -_sum(_share_a(txn) for txn in personal))
        self.assertEqual(self.book.number(workbook.SH_HOUSEHOLD, "D8"),
                         -_sum(_share_b(txn) for txn in personal))

    def test_the_settlement_says_who_owes_whom(self):
        balance = self.book.number(workbook.SH_HOUSEHOLD, "C9")
        settlement = self.book.text(workbook.SH_HOUSEHOLD, "C11")
        self.assertEqual(
            balance,
            round(self.book.number(workbook.SH_HOUSEHOLD, "C7")
                  - self.book.number(workbook.SH_HOUSEHOLD, "C8"), CENTS))
        if balance > 0:
            self.assertIn(f"{sample.PERSON_B} owes {sample.PERSON_A}", settlement)
        elif balance < 0:
            self.assertIn(f"{sample.PERSON_A} owes {sample.PERSON_B}", settlement)
        else:
            self.assertIn("Even", settlement)
        self.assertRegex(settlement, r"\$[\d,]+\.\d\d|Even")

    def test_the_two_balances_cancel(self):
        # Whatever one of them is owed, the other owes.  If these do not cancel
        # the shares are not a partition of the spending.
        self.assertEqual(
            round(self.book.number(workbook.SH_HOUSEHOLD, "C9")
                  + self.book.number(workbook.SH_HOUSEHOLD, "D9"), CENTS), 0.0)

    def test_the_view_selector_switches_every_report_to_one_person(self):
        month = _month(_records)
        try:
            for person, share in ((sample.PERSON_A, _share_a),
                                  (sample.PERSON_B, _share_b)):
                self.book.set(workbook.SH_DASHBOARD, "F6", person)
                with self.subTest(person):
                    self.assertEqual(
                        self.book.number(workbook.SH_DASHBOARD, "C9"),
                        _sum(share(txn) for txn in _of_type(month, "Income")))
                    self.assertEqual(
                        self.book.number(workbook.SH_DASHBOARD, "C10"),
                        -_sum(share(txn) for txn in _of_type(month, "Expense")))
        finally:
            self.book.set(workbook.SH_DASHBOARD, "F6", "Household")

    def test_the_two_personal_views_add_up_to_the_household_one(self):
        # Share A and Share B are meant to be a partition of the amount, so
        # neither person's view can lose or invent money.
        totals = {}
        try:
            for view in ("Household", sample.PERSON_A, sample.PERSON_B):
                self.book.set(workbook.SH_DASHBOARD, "F6", view)
                totals[view] = (self.book.number(workbook.SH_DASHBOARD, "C9"),
                                self.book.number(workbook.SH_DASHBOARD, "C10"))
        finally:
            self.book.set(workbook.SH_DASHBOARD, "F6", "Household")

        for index, label in enumerate(("money in", "money out")):
            with self.subTest(label):
                self.assertAlmostEqual(
                    totals[sample.PERSON_A][index] + totals[sample.PERSON_B][index],
                    totals["Household"][index], places=1)


if __name__ == "__main__":
    unittest.main()
