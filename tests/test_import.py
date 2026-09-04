"""Runs the shipped import path over the sample bank exports.

The four files under ``samples/`` are written by ``tools.sample`` in each
bank's real export shape - RBC's split description columns, BMO's notice lines
and 8-digit dates, Amex's inverted signs, Tangerine's memo column - from a
known list of transactions.  Reading them back with the workbook's own macros
and recovering that list is the end-to-end check that the import works: the
date order, the sign convention, the duplicate key and the categorisation all
have to be right at once or the recovered records will not match.

The macros normally read the Bank Formats and Rules sheets for this, which
needs Excel; ``tests.refdata`` renders those same rows from ``tools.data``
instead, so what runs is the shipped code over the shipped reference data.
"""

from __future__ import annotations

import os
import sys
import unittest
from datetime import date
from typing import Dict, List

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests import libreoffice, refdata, vbahost  # noqa: E402
from tools import sample  # noqa: E402

TODAY = date(2026, 9, 1)
SAMPLE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "samples")

# file -> the bank format that reads it, the account it belongs to and whose
# money that account holds.
SAMPLES = {
    "rbc-chequing-alex.csv": ("RBC Chequing/Savings/Card", sample.ACCOUNT_ALEX, "Alex"),
    "tangerine-chequing-sam.csv": ("Tangerine", sample.ACCOUNT_SAM, "Sam"),
    "bmo-joint-chequing.csv": ("BMO", sample.ACCOUNT_JOINT, "Joint"),
    "amex-cobalt-joint.csv": ("Amex Canada", sample.ACCOUNT_CARD, "Joint"),
}

# The rows that were written to each file, in the order they were written.
EXPECTED: Dict[str, List[sample.Txn]] = {}


def setUpModule():
    try:
        libreoffice.context()
    except libreoffice.Unavailable as exc:  # pragma: no cover - environment
        raise unittest.SkipTest(str(exc))

    records = sample.build(TODAY)
    for name, (_profile, account, _owner) in SAMPLES.items():
        EXPECTED[name] = [txn for txn in records if txn.account == account]
    missing = [name for name in SAMPLES if not os.path.exists(_path(name))]
    if missing:
        raise unittest.SkipTest(f"run build.py first: {missing} are not present")


def _path(name: str) -> str:
    return os.path.join(SAMPLE_DIR, name)


def _run(body: str) -> List[List[str]]:
    return vbahost.rows(vbahost.run(
        body, vbahost.IMPORT_MODULES,
        extra={"Formats": refdata.profiles_basic(), "Rules": refdata.rules_basic()}))


class FormatDetectionTests(unittest.TestCase):
    """Which bank wrote the file, decided from its opening lines alone."""

    def test_each_sample_is_recognised(self):
        lines = ["Sub Run()", "    Dim rows As Collection",
                 "    Dim matched As clsProfile"]
        for name in SAMPLES:
            lines += [
                f"    Set rows = modParse.SplitRows("
                f"modParse.ReadTextFile({vbahost.basic_string(_path(name))}), \",\")",
                "    Set matched = modProfiles.MatchProfile(SheetProfiles(), rows)",
                f"    If matched Is Nothing Then",
                f"        Emit {vbahost.basic_string(name)}, \"(unrecognised)\"",
                "    Else",
                f"        Emit {vbahost.basic_string(name)}, matched.Name",
                "    End If",
            ]
        lines.append("End Sub")

        got = dict((row[0], row[1]) for row in _run("\n".join(lines)))
        self.assertEqual(got, {name: profile
                              for name, (profile, _, _) in SAMPLES.items()})

    def test_a_file_that_matches_nothing_is_left_undecided(self):
        # Better to ask the user than to guess: an unrecognised file has to
        # come back as Nothing so ImportOneFile falls through to the picker.
        body = '''
Sub Run()
    Dim rows As Collection
    Dim matched As clsProfile
    Set rows = modParse.SplitRows("Widget,Colour,Price" & Chr$(10) & "bolt,red,3", ",")
    Set matched = modProfiles.MatchProfile(SheetProfiles(), rows)
    Emit "matched", Flag(Not matched Is Nothing)
End Sub
'''
        self.assertEqual(_run(body)[0], ["matched", "no"])


class ReadBackTests(unittest.TestCase):
    """Every row of every sample file, recovered through the real read path."""

    read: Dict[str, List[List[str]]] = {}

    @classmethod
    def setUpClass(cls):
        for name, (profile, account, owner) in SAMPLES.items():
            body = f'''
Sub Run()
    Dim profile As clsProfile
    Dim rows As Collection
    Dim records As Collection
    Dim rules As Collection
    Dim rule As clsRule
    Dim txn As clsTxn
    Dim readCount As Long, badCount As Long
    Dim i As Long
    Dim category As String

    Set profile = ProfileNamed({vbahost.basic_string(profile)})
    Set rules = modRules.ByPriority(SheetRules())
    Set rows = modParse.SplitRows( _
        modParse.ReadTextFile({vbahost.basic_string(_path(name))}), profile.Delimiter())
    Set records = modImport.ReadRecords(rows, profile, _
        {vbahost.basic_string(account)}, {vbahost.basic_string(owner)}, _
        {vbahost.basic_string(name)}, readCount, badCount)

    Emit "counts", rows.Count, records.Count, readCount, badCount
    For i = 1 To records.Count
        Set txn = records.Item(i)
        Set rule = modRules.FirstMatch(rules, txn.Merchant, txn.Description, _
                                       txn.Account, txn.Amount)
        If rule Is Nothing Then category = "" Else category = rule.Category
        Emit "row", Stamp(txn.TxnDate), Money(txn.Amount), txn.Description, _
             txn.MatchKey(), category, txn.Merchant, txn.Account, txn.Owner
    Next i
End Sub
'''
            cls.read[name] = _run(body)

    def rows_for(self, name: str) -> List[List[str]]:
        return [row[1:] for row in self.read[name] if row[0] == "row"]

    def counts_for(self, name: str) -> List[str]:
        return next(row[1:] for row in self.read[name] if row[0] == "counts")

    def test_every_written_row_comes_back(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                rows = self.rows_for(name)
                self.assertEqual(len(rows), len(expected))
                _read, kept, considered, unreadable = self.counts_for(name)
                self.assertEqual(int(kept), len(expected))
                self.assertEqual(int(unreadable), 0,
                                 "every row past the header is a transaction")
                self.assertEqual(int(considered), len(expected))

    def test_dates_survive_each_banks_written_order(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                got = [row[0] for row in self.rows_for(name)]
                self.assertEqual(got, [txn.when.isoformat() for txn in expected])

    def test_money_out_is_negative_whatever_the_bank_called_it(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                got = [row[1] for row in self.rows_for(name)]
                self.assertEqual(got, [f"{txn.amount:.2f}" for txn in expected])

    def test_amex_purchases_are_written_positive_and_read_negative(self):
        # The sign flip is the one thing a credit-card profile gets wrong most
        # easily, and getting it wrong inverts the whole report.
        name = "amex-cobalt-joint.csv"
        with open(_path(name), encoding="utf-8") as stream:
            raw = stream.read().splitlines()[1:]
        written = [float(line.split(",")[-2]) for line in raw if line]
        self.assertTrue(any(value > 0 for value in written),
                        "the fixture is meant to hold Amex's positive purchases")
        read_back = [float(row[1]) for row in self.rows_for(name)]
        self.assertEqual(read_back, [-value for value in written])

    def test_descriptions_are_rebuilt_from_however_many_columns(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                got = [row[2] for row in self.rows_for(name)]
                self.assertEqual(got, [txn.description for txn in expected])

    def test_the_duplicate_key_agrees_with_the_builders(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                got = [row[3] for row in self.rows_for(name)]
                self.assertEqual(got, [
                    sample.match_key(txn.account, txn.when, txn.amount,
                                     txn.description)
                    for txn in expected
                ])

    def test_the_rules_place_every_transaction(self):
        for name, expected in EXPECTED.items():
            with self.subTest(name):
                got = [row[4] for row in self.rows_for(name)]
                self.assertNotIn("", got, "no transaction should go unmatched")
                wrong = [
                    (txn.description, want, mine)
                    for txn, want, mine in
                    zip(expected, (txn.category for txn in expected), got)
                    if want != mine
                ]
                self.assertEqual(wrong, [])

    def test_the_account_and_owner_come_from_the_file_not_the_row(self):
        for name, (_profile, account, owner) in SAMPLES.items():
            with self.subTest(name):
                rows = self.rows_for(name)
                self.assertEqual({row[6] for row in rows}, {account})
                self.assertEqual({row[7] for row in rows}, {owner})


class SkippedRowTests(unittest.TestCase):
    """Header rows and bank notices are skipped, not imported as transactions."""

    def test_a_header_that_is_read_anyway_is_counted_unreadable(self):
        # Skip Rows is a user-editable number, so it will sometimes be wrong.
        # Rows that are not transactions must fall out rather than land in the
        # ledger as a transaction dated whenever DateSerial felt like.
        cases = {"rbc-chequing-alex.csv": 1, "bmo-joint-chequing.csv": 3,
                 "tangerine-chequing-sam.csv": 1, "amex-cobalt-joint.csv": 1}
        lines = ["Sub Run()", "    Dim profile As clsProfile",
                 "    Dim rows As Collection", "    Dim records As Collection",
                 "    Dim readCount As Long, badCount As Long"]
        for name in cases:
            profile = SAMPLES[name][0]
            lines += [
                f"    Set profile = ProfileNamed({vbahost.basic_string(profile)})",
                "    profile.SkipRows = 0",
                f"    Set rows = modParse.SplitRows(modParse.ReadTextFile("
                f"{vbahost.basic_string(_path(name))}), profile.Delimiter())",
                "    readCount = 0: badCount = 0",
                f"    Set records = modImport.ReadRecords(rows, profile, \"A\", \"O\", "
                f"{vbahost.basic_string(name)}, readCount, badCount)",
                f"    Emit {vbahost.basic_string(name)}, records.Count, badCount",
            ]
        lines.append("End Sub")

        got = dict((row[0], (int(row[1]), int(row[2]))) for row in _run("\n".join(lines)))
        for name, header_rows in cases.items():
            with self.subTest(name):
                kept, unreadable = got[name]
                self.assertEqual(unreadable, header_rows)
                self.assertEqual(kept, len(EXPECTED[name]),
                                 "the transactions still all arrive")

class NoHeaderFormatTests(unittest.TestCase):
    """Recurring exports can be identified by their account even without headers."""

    CASES = [
        ("TD (no header)",
         "03/11/2026,LOBLAWS,48.14,,1000.00", "-48.14", "LOBLAWS"),
        ("CIBC (no header)",
         "2026-03-11,LOBLAWS,48.14,,CARD", "-48.14", "LOBLAWS"),
        ("Simplii (no header)",
         "2026-03-11,PAYROLL,,2483.18", "2483.18", "PAYROLL"),
        ("Scotiabank (no header)",
         "03/11/2026,-48.14,LOBLAWS,TORONTO", "-48.14", "LOBLAWS TORONTO"),
    ]

    def test_each_no_header_layout_parses_when_selected_from_the_account(self):
        lines = [
            "Sub Run()",
            "    Dim profile As clsProfile",
            "    Dim rows As Collection, records As Collection",
            "    Dim readCount As Long, badCount As Long",
            "    Dim txn As clsTxn",
        ]
        for profile, text, _amount, _description in self.CASES:
            lines += [
                f"    Set profile = ProfileNamed({vbahost.basic_string(profile)})",
                f"    Set rows = modParse.SplitRows({vbahost.basic_string(text)}, "
                "profile.Delimiter())",
                "    Set records = modImport.ReadRecords(rows, profile, "
                '"Account", "Owner", "statement.csv", readCount, badCount)',
                "    Set txn = records.Item(1)",
                f"    Emit {vbahost.basic_string(profile)}, Money(txn.Amount), "
                "txn.Description, badCount",
            ]
        lines.append("End Sub")

        got = {row[0]: row[1:] for row in _run("\n".join(lines))}
        self.assertEqual(got, {
            profile: [amount, description, "0"]
            for profile, _text, amount, description in self.CASES
        })


class DuplicateTests(unittest.TestCase):
    """Re-downloading an overlapping statement must not double the ledger."""

    NAME = "bmo-joint-chequing.csv"

    def _body(self) -> str:
        profile, account, owner = SAMPLES[self.NAME]
        return f'''
Function Twin() As clsTxn
    Dim txn As clsTxn
    Set txn = New clsTxn
    txn.TxnDate = DateSerial(2026, 3, 11)
    txn.Description = "TIM HORTONS #123 TORONTO ON"
    txn.Amount = -4.75
    txn.Account = "RBC Chequing (Alex)"
    Set Twin = txn
End Function

Sub Run()
    Dim profile As clsProfile
    Dim rows As Collection, records As Collection
    Dim existing As Collection, fresh As Collection, again As Collection
    Dim pair As Collection
    Dim readCount As Long, badCount As Long, dupes As Long
    Dim i As Long

    Set profile = ProfileNamed({vbahost.basic_string(profile)})
    Set rows = modParse.SplitRows( _
        modParse.ReadTextFile({vbahost.basic_string(_path(self.NAME))}), _
        profile.Delimiter())
    Set records = modImport.ReadRecords(rows, profile, _
        {vbahost.basic_string(account)}, {vbahost.basic_string(owner)}, _
        {vbahost.basic_string(self.NAME)}, readCount, badCount)

    Set existing = New Collection
    Set fresh = modImport.WithoutDuplicates(records, existing, dupes)
    Emit "first", fresh.Count, dupes

    For i = 1 To fresh.Count
        modUtil.BumpVal existing, fresh.Item(i).MatchKey()
    Next i
    dupes = 0
    Set again = modImport.WithoutDuplicates(records, existing, dupes)
    Emit "second", again.Count, dupes

    ' Two identical coffees bought on one day are two transactions.
    Set pair = New Collection
    pair.Add Twin()
    pair.Add Twin()
    dupes = 0
    Emit "pair", modImport.WithoutDuplicates(pair, New Collection, dupes).Count, dupes

    ' ... and if the ledger already holds one of them, only the second is new.
    Set existing = New Collection
    modUtil.BumpVal existing, Twin().MatchKey()
    dupes = 0
    Emit "pairagain", modImport.WithoutDuplicates(pair, existing, dupes).Count, dupes
End Sub
'''

    def test_a_second_import_of_the_same_file_adds_nothing(self):
        got = dict((row[0], [int(value) for value in row[1:]])
                   for row in _run(self._body()))
        total = len(EXPECTED[self.NAME])
        self.assertEqual(got["first"], [total, 0])
        self.assertEqual(got["second"], [0, total])
        self.assertEqual(got["pair"], [2, 0], "identical same-day purchases both land")
        self.assertEqual(got["pairagain"], [1, 1])


class RuleCollisionTests(unittest.TestCase):
    """Names that contain a shorter pattern belonging to somewhere else.

    A "Contains" rule on a short pattern is the main way this rule set goes
    wrong: MOBIL is a gas station and also the tail of FREEDOM MOBILE, MEC is a
    co-op and also the start of MECHANIC.  Each pair below is a real Canadian
    merchant that a careless pattern swallows.
    """

    CASES = [
        ("FREEDOM MOBILE PREAUTHORIZED DEBIT", -45.00, "Mobile Phone"),
        ("MOBIL 1234 TORONTO ON", -62.40, "Fuel"),
        ("PUBLIC MOBILE PREAUTHORIZED PAYMENT", -29.00, "Mobile Phone"),
        ("KOODO MOBILE PREAUTHORIZED PAYMENT", -50.00, "Mobile Phone"),
        ("IDP PURCHASE - 1234 MEC TORONTO ON", -189.99, "Fitness & Sports"),
        ("BROADVIEW AUTO MECHANIC LTD", -430.00, None),
        ("IDP PURCHASE - 4612 IGA EXTRA MONTREAL QC", -88.20, "Groceries"),
        ("IDP PURCHASE - 4612 METRO PLUS MONTREAL QC", -71.05, "Groceries"),
        ("STM OPUS MONTREAL QC", -97.00, "Public Transit"),
        ("IDP PURCHASE - 1234 SAQ SELECTION MONTREAL QC", -42.75, "Alcohol"),
        ("SHELL 4321 TORONTO ON", -78.10, "Fuel"),
        ("OCS.CA TORONTO ON", -55.00, "Cannabis"),
        ("TELUS MOBILITY PREAUTHORIZED PAYMENT", -85.00, "Mobile Phone"),
        ("SUN LIFE ASSURANCE MISC PAYMENT", -74.50, "Life & Disability Insurance"),
        ("CANADA LIFE GROUP BENEFITS", -61.20, "Life & Disability Insurance"),
    ]

    def test_the_more_specific_merchant_wins(self):
        lines = ["Sub Run()", "    Dim rules As Collection", "    Dim rule As clsRule",
                 "    Set rules = modRules.ByPriority(SheetRules())"]
        for index, (description, amount, _) in enumerate(self.CASES, start=1):
            literal = vbahost.basic_string(description)
            lines += [
                f"    Set rule = modRules.FirstMatch(rules, "
                f"modRules.CleanMerchant({literal}), {literal}, \"Chequing\", {amount})",
                "    If rule Is Nothing Then",
                f"        Emit {index}, \"(none)\"",
                "    Else",
                f"        Emit {index}, rule.Category",
                "    End If",
            ]
        lines.append("End Sub")

        got = [row[1] for row in _run("\n".join(lines))]
        want = [category or "(none)" for _, _, category in self.CASES]
        self.assertEqual(
            [(case[0], expected, actual)
             for case, expected, actual in zip(self.CASES, want, got)
             if expected != actual],
            [])


class RuleOrderTests(unittest.TestCase):
    def test_priority_decides_and_ties_keep_sheet_order(self):
        body = '''
Sub Run()
    Dim ordered As Collection
    Dim i As Long
    Dim line As String
    Set ordered = modRules.ByPriority(SheetRules())
    For i = 2 To ordered.Count
        If ordered.Item(i - 1).Priority > ordered.Item(i).Priority Then
            Emit "outoforder", i
        ElseIf ordered.Item(i - 1).Priority = ordered.Item(i).Priority Then
            If ordered.Item(i - 1).RowIndex > ordered.Item(i).RowIndex Then
                Emit "unstable", i
            End If
        End If
    Next i
    Emit "count", ordered.Count
End Sub
'''
        rows = _run(body)
        self.assertEqual([row for row in rows if row[0] != "count"], [])
        self.assertEqual(int(rows[-1][1]), len(refdata.data.seed_rules()))

    def test_a_registered_plan_beats_the_generic_transfer_rule(self):
        # "TRANSFER TO RRSP WEALTHSIMPLE" matches both the RRSP rule and the
        # internal-transfer rule; if the transfer rule wins, a year of RRSP
        # contributions disappears from the savings total.
        cases = [
            ("TRANSFER TO RRSP WEALTHSIMPLE", -450.0, "RRSP Contribution"),
            ("TFSA CONTRIBUTION TANGERINE INVESTMENT", -350.0, "TFSA Contribution"),
            ("TRANSFER TO SAVINGS 1234", -200.0, "Internal Transfer"),
            ("TRANSFER FROM SAVINGS 1234", 200.0, "Internal Transfer"),
            ("PAYMENT - THANK YOU / PAIEMENT - MERCI", 812.44, "Credit Card Payment"),
        ]
        lines = ["Sub Run()", "    Dim rules As Collection", "    Dim rule As clsRule",
                 "    Set rules = modRules.ByPriority(SheetRules())"]
        for description, amount, _ in cases:
            literal = vbahost.basic_string(description)
            lines += [
                f"    Set rule = modRules.FirstMatch(rules, "
                f"modRules.CleanMerchant({literal}), {literal}, \"Chequing\", {amount})",
                "    If rule Is Nothing Then",
                f"        Emit {literal}, \"(none)\"",
                "    Else",
                f"        Emit {literal}, rule.Category",
                "    End If",
            ]
        lines.append("End Sub")

        got = [row[1] for row in _run("\n".join(lines))]
        self.assertEqual(got, [category for _, _, category in cases])


class RefundTests(unittest.TestCase):
    # A card statement's refunds come back as money in under the shop's own
    # name; they belong in the category of the purchase they reverse.  Money in
    # that the description rules recognise, and money in nobody recognises,
    # must not be pulled in by that.
    CASES = [
        ("RETURN LOBLAWS #4861 TORONTO ON", 12.00, "Groceries"),
        ("AMAZON.CA*2K4XY1 AMAZON.CA ON", 23.45, "Miscellaneous"),
        ("LOBLAWS #4861 TORONTO ON", -99.09, "Groceries"),
        ("PAYROLL DEPOSIT NORTHWIND", 2483.18, "Employment Income"),
        ("TRANSFER FROM SAVINGS 1234", 200.0, "Internal Transfer"),
        ("E-TRANSFER RECEIVED FROM JANE", 60.0, "Interac e-Transfer Received"),
        ("XYZZY UNKNOWN MERCHANT", 19.99, None),
    ]

    def test_a_refund_follows_the_purchase_into_its_category(self):
        lines = ["Sub Run()", "    Dim rules As Collection", "    Dim rule As clsRule",
                 "    Set rules = modRules.ByPriority(SheetRules())"]
        for index, (description, amount, _) in enumerate(self.CASES, start=1):
            literal = vbahost.basic_string(description)
            lines += [
                f"    Set rule = modRules.RuleFor(rules, "
                f"modRules.CleanMerchant({literal}), {literal}, \"Card\", {amount})",
                "    If rule Is Nothing Then",
                f"        Emit {index}, \"(none)\"",
                "    Else",
                f"        Emit {index}, rule.Category",
                "    End If",
            ]
        lines.append("End Sub")

        got = [row[1] for row in _run("\n".join(lines))]
        want = [category or "(none)" for _, _, category in self.CASES]
        self.assertEqual(
            [(case[0], expected, actual)
             for case, expected, actual in zip(self.CASES, want, got)
             if expected != actual],
            [])

    def test_only_merchant_rules_that_expect_money_out_count_as_purchases(self):
        body = '''
Sub Run()
    Dim rule As clsRule
    Set rule = New clsRule
    rule.LookIn = "Merchant": rule.Flow = "Money out"
    Emit "merchant out", Flag(rule.IsMerchantPaymentRule())
    rule.Flow = "Any"
    Emit "merchant any", Flag(rule.IsMerchantPaymentRule())
    rule.Flow = "Money in"
    Emit "merchant in", Flag(rule.IsMerchantPaymentRule())
    rule.LookIn = "Description": rule.Flow = "Money out"
    Emit "description out", Flag(rule.IsMerchantPaymentRule())
    rule.LookIn = "Any"
    Emit "any out", Flag(rule.IsMerchantPaymentRule())
    rule.LookIn = ""
    Emit "blank out", Flag(rule.IsMerchantPaymentRule())
End Sub
'''
        self.assertEqual(dict(_run(body)), {
            "merchant out": "yes", "merchant any": "no", "merchant in": "no",
            "description out": "no", "any out": "no", "blank out": "no"})


if __name__ == "__main__":
    unittest.main()
