"""Runs the shipped VBA against the awkward input real banks produce.

Every assertion here executes the module text from ``vba/`` inside LibreOffice
Basic, so what is being checked is the code that ships in the workbook rather
than a Python restatement of it.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests import libreoffice, vbahost  # noqa: E402
from tools import sample  # noqa: E402

TODAY = date(2026, 9, 1)


def setUpModule():
    try:
        libreoffice.context()
    except libreoffice.Unavailable as exc:  # pragma: no cover - environment
        raise unittest.SkipTest(str(exc))


class AmountTests(unittest.TestCase):
    """None of this is CDbl: CDbl follows the machine's regional settings."""

    CASES = [
        ("-45.00", -45.0),
        ("45.00", 45.0),
        ("$1,234.56", 1234.56),
        ("1,234.56", 1234.56),
        ("(45.00)", -45.0),          # accounting negative
        ("45.00-", -45.0),           # trailing sign
        ("1 234,56 $", 1234.56),     # French-Canadian
        ("1.234,56", 1234.56),       # dot thousands, comma decimal
        ("1234,56", 1234.56),
        ("1,234", 1234.0),
        ("CAD 12.50", 12.5),
        ("0.00", 0.0),
        ("\u00a01\u00a0999,99", 1999.99),   # non-breaking spaces
        ("", None),
        ("   ", None),
        ("Amount", None),
        ("--", None),
    ]

    def test_parses_every_layout(self):
        body = _probe([f"modParse.ParseAmount({vbahost.basic_string(text)}, ok)"
                       for text, _ in self.CASES],
                      kind="Double", emit="Money(result), Flag(ok)",
                      declare="Dim ok As Boolean")
        rows = vbahost.rows(vbahost.run(body))
        self.assertEqual(len(rows), len(self.CASES))
        for (text, expected), (_, value, ok) in zip(self.CASES, rows):
            if expected is None:
                self.assertEqual(ok, "no", f"{text!r} should not parse")
            else:
                self.assertEqual(ok, "yes", f"{text!r} should parse")
                self.assertAlmostEqual(float(value), expected, places=2, msg=text)


class DateTests(unittest.TestCase):
    CASES = [
        ("03/11/2026", "MM/DD/YYYY", "2026-03-11"),
        ("11/03/2026", "DD/MM/YYYY", "2026-03-11"),
        ("2026-03-11", "YYYY-MM-DD", "2026-03-11"),
        ("20260311", "YYYYMMDD", "2026-03-11"),
        ("11-Mar-2026", "DD-MMM-YYYY", "2026-03-11"),
        ("Mar 11, 2026", "MMM-DD-YYYY", "2026-03-11"),
        ("2026-03-11 14:32:07", "YYYY-MM-DD", "2026-03-11"),   # trailing time
        ("31/12/26", "DD/MM/YYYY", "2026-12-31"),              # two digit year
        ("2026/03/11", "AUTO", "2026-03-11"),
        ("13/03/2026", "AUTO", "2026-03-13"),                  # day > 12 decides
        ("11 mars 2026", "AUTO", "2026-03-11"),                # French month
        ("11 aout 2026", "AUTO", "2026-08-11"),
        ("Transaction Date", "AUTO", None),                    # a header row
        ("", "AUTO", None),
        ("2026-13-11", "YYYY-MM-DD", None),                    # month 13
        # DateSerial would roll this forward to March 2nd; the parser must not.
        ("2026-02-30", "YYYY-MM-DD", None),
    ]

    def test_reads_each_written_order(self):
        body = _probe(
            [f"modParse.ParseDate({vbahost.basic_string(text)}, "
             f"{vbahost.basic_string(pattern)}, ok)"
             for text, pattern, _ in self.CASES],
            kind="Date", emit="Stamp(result), Flag(ok)",
            declare="Dim ok As Boolean")
        rows = vbahost.rows(vbahost.run(body))
        self.assertEqual(len(rows), len(self.CASES))
        for (text, _, expected), (_, stamp, ok) in zip(self.CASES, rows):
            if expected is None:
                self.assertEqual(ok, "no", f"{text!r} should not parse")
            else:
                self.assertEqual(ok, "yes", f"{text!r} should parse")
                self.assertEqual(stamp, expected, text)


class DelimitedTextTests(unittest.TestCase):
    def test_handles_quotes_commas_and_line_breaks(self):
        text = (
            'Date,Description,Amount\r\n'
            '2026-03-01,"LOBLAWS #123, TORONTO",-48.14\r\n'
            '\r\n'                                        # blank line
            '2026-03-02,"He said ""hi""",-1.00\r\n'
            '2026-03-03,"two\nlines",-2.00\r\n'
            '2026-03-04,plain,-3.00'                      # no trailing newline
        )
        body = f'''
Sub Run()
    Dim rows As Collection
    Dim i As Long
    Set rows = modParse.SplitRows({vbahost.basic_string(text)}, ",")
    Emit "count", rows.Count
    For i = 1 To rows.Count
        Emit i, rows.Item(i).Count, _
             Replace$(modParse.FieldAt(rows.Item(i), 2), Chr$(10), "<lf>"), _
             modParse.FieldAt(rows.Item(i), 3)
    Next i
End Sub
'''
        rows = vbahost.rows(vbahost.run(body))
        self.assertEqual(rows[0], ["count", "5"], "the blank line must be dropped")
        self.assertEqual([row[1] for row in rows[1:]], ["3"] * 5)
        self.assertEqual([row[2] for row in rows[1:]], [
            "Description",
            "LOBLAWS #123, TORONTO",
            'He said "hi"',
            "two<lf>lines",
            "plain",
        ])
        self.assertEqual([row[3] for row in rows[1:]],
                         ["Amount", "-48.14", "-1.00", "-2.00", "-3.00"])

    def test_joins_the_columns_a_description_is_split_over(self):
        # RBC spreads one description over two columns.
        body = '''
Sub Run()
    Dim rows As Collection
    Set rows = modParse.SplitRows("a,b,PAYROLL DEPOSIT NORTHWIND,LOGISTICS,c", ",")
    Emit "joined", modParse.FieldsAt(rows.Item(1), "3,4")
    Emit "single", modParse.FieldsAt(rows.Item(1), "3")
    Emit "missing", modParse.FieldsAt(rows.Item(1), "3,9")
    Emit "beyond", "[" & modParse.FieldAt(rows.Item(1), 12) & "]"
End Sub
'''
        rows = dict((row[0], row[1]) for row in vbahost.rows(vbahost.run(body)))
        self.assertEqual(rows["joined"], "PAYROLL DEPOSIT NORTHWIND LOGISTICS")
        self.assertEqual(rows["single"], "PAYROLL DEPOSIT NORTHWIND")
        self.assertEqual(rows["missing"], "PAYROLL DEPOSIT NORTHWIND")
        self.assertEqual(rows["beyond"], "[]")

    def test_splits_on_semicolons_and_tabs(self):
        body = '''
Sub Run()
    Dim rows As Collection
    Set rows = modParse.SplitRows("a;b;c", ";")
    Emit "semicolon", rows.Item(1).Count, modParse.FieldAt(rows.Item(1), 2)
    Set rows = modParse.SplitRows("a" & Chr$(9) & "b" & Chr$(9) & "c", Chr$(9))
    Emit "tab", rows.Item(1).Count, modParse.FieldAt(rows.Item(1), 2)
End Sub
'''
        rows = vbahost.rows(vbahost.run(body))
        self.assertEqual(rows, [["semicolon", "3", "b"], ["tab", "3", "b"]])


class EncodingTests(unittest.TestCase):
    """Bank downloads arrive in whatever the bank's server felt like."""

    LINE = "ÉPICERIE MÉTRO – MONTRÉAL"
    TEXT = f"Date,Description\n2026-03-01,{LINE}\n"

    def _read(self, raw: bytes) -> str:
        handle, path = tempfile.mkstemp(prefix="cft-encoding-", suffix=".csv")
        with os.fdopen(handle, "wb") as stream:
            stream.write(raw)
        try:
            body = ('Sub Run()\n    Emit "text", Replace$(modParse.ReadTextFile('
                    f'{vbahost.basic_string(path)}), Chr$(10), "<lf>")\nEnd Sub\n')
            row = vbahost.rows(vbahost.run(body))[0]
            return row[1] if len(row) > 1 else ""
        finally:
            os.unlink(path)

    def test_utf8_without_a_bom(self):
        self.assertIn(self.LINE, self._read(self.TEXT.encode("utf-8")))

    def test_utf8_with_a_bom(self):
        raw = b"\xef\xbb\xbf" + self.TEXT.encode("utf-8")
        text = self._read(raw)
        self.assertIn(self.LINE, text)
        self.assertTrue(text.startswith("Date"), "the BOM must not survive")

    def test_utf16_little_endian(self):
        raw = b"\xff\xfe" + self.TEXT.encode("utf-16-le")
        self.assertIn(self.LINE, self._read(raw))

    def test_windows_1252(self):
        self.assertIn(self.LINE, self._read(self.TEXT.encode("cp1252")))

    def test_an_empty_file_reads_as_nothing(self):
        self.assertEqual(self._read(b""), "")


class MerchantTests(unittest.TestCase):
    """The noise Canadian terminals staple onto a description."""

    CASES = [
        ("IDP PURCHASE - 4612 CANADIAN TIRE #2196 ON", "Canadian Tire"),
        ("IDP PURCHASE - 2615 LOBLAWS #3980 TORONTO ON", "Loblaws Toronto"),
        ("INTERAC RETAIL PURCHASE - 1234 NO FRILLS #101 TORONTO ON",
         "No Frills Toronto"),
        ("TIM HORTONS #1246 TORONTO ON", "Tim Hortons Toronto"),
        ("PRESTO FARE/TRANSIT TORONTO ON", "Presto Fare/transit Toronto"),
        ("ROGERS PREAUTHORIZED PAYMENT", "Rogers"),
        ("GOODLIFE CLUBS PREAUTHORIZED DEBIT", "Goodlife Clubs"),
        ("NETFLIX.COM MISC PAYMENT", "Netflix.com"),
        ("AMZ*UBER EATS TORONTO ON", "Amz*uber Eats Toronto"),
        ("PAYMENT - THANK YOU / PAIEMENT - MERCI", "Thank You / Paiement - Merci"),
        ("SQ *THE COFFEE PLACE", "The Coffee Place"),
        ("MONTHLY ACCOUNT FEE", "Monthly Account Fee"),
        ("TFSA CONTRIBUTION TANGERINE INVESTMENT",
         "TFSA Contribution Tangerine Investment"),
        ("ATM WITHDRAWAL 100 QUEEN ST W", "ATM Withdrawal 100 Queen St W"),
        ("", ""),
    ]

    def test_strips_noise_and_store_numbers(self):
        body = _probe([f"modRules.CleanMerchant({vbahost.basic_string(text)})"
                       for text, _ in self.CASES], kind="String", emit="result")
        rows = vbahost.rows(vbahost.run(body))
        got = [row[1] if len(row) > 1 else "" for row in rows]
        self.assertEqual(got, [expected for _, expected in self.CASES])

    def test_the_same_shop_on_four_terminals_is_one_merchant(self):
        descriptions = [
            f"IDP PURCHASE - {ref} LOBLAWS #{store} TORONTO ON"
            for ref, store in (("4612", "2196"), ("1234", "3980"),
                               ("9999", "101"), ("2615", "4"))
        ]
        body = _probe([f"modRules.CleanMerchant({vbahost.basic_string(text)})"
                       for text in descriptions], kind="String", emit="result")
        names = {row[1] for row in vbahost.rows(vbahost.run(body))}
        self.assertEqual(names, {"Loblaws Toronto"})


class TextTests(unittest.TestCase):
    def test_title_case_keeps_canadian_acronyms(self):
        words = ["lcbo toronto", "rrsp contribution", "tim hortons", "ttc presto",
                 "bmo joint chequing", "tfsa top up"]
        body = _probe([f"modUtil.TitleCaseWords({vbahost.basic_string(word)})"
                       for word in words], kind="String", emit="result")
        got = [row[1] for row in vbahost.rows(vbahost.run(body))]
        self.assertEqual(got, ["LCBO Toronto", "RRSP Contribution", "Tim Hortons",
                               "TTC Presto", "BMO Joint Chequing", "TFSA Top Up"])

    def test_condense_spaces_folds_every_kind_of_whitespace(self):
        body = ('Sub Run()\n'
                '    Emit "a", "[" & modUtil.CondenseSpaces("  a" & Chr$(9) & '
                '"b  " & Chr$(160) & " c ") & "]"\n'
                'End Sub\n')
        self.assertEqual(vbahost.rows(vbahost.run(body))[0][1], "[a b c]")

    def test_file_name_is_taken_off_either_separator(self):
        paths = ["C:\\Users\\alex\\Downloads\\rbc.csv",
                 "/home/alex/Downloads/rbc.csv", "rbc.csv"]
        body = _probe([f"modImport.FileNameOnly({vbahost.basic_string(path)})"
                       for path in paths], kind="String", emit="result")
        got = [row[1] for row in vbahost.rows(vbahost.run(body))]
        self.assertEqual(got, ["rbc.csv"] * 3)


class HashTests(unittest.TestCase):
    """The duplicate key has to agree with the one the builder wrote."""

    WORDS = ["", "a", "LOBLAWS", "ÉPICERIE", "x" * 200,
             "RBC CHEQUING (ALEX)|2026-03-11|-48.14|LOBLAWS #123"]

    def test_fnv1a_matches_the_python_implementation(self):
        body = _probe([f"modUtil.HashText({vbahost.basic_string(word)})"
                       for word in self.WORDS], kind="String", emit="result")
        rows = vbahost.rows(vbahost.run(body))
        got = [row[1] if len(row) > 1 else "" for row in rows]
        self.assertEqual(got, [sample.fnv1a(word) for word in self.WORDS])

    def test_match_key_matches_the_ledger_key(self):
        records = sample.build(TODAY)[:40]
        calls = [
            f"modUtil.MatchKey({vbahost.basic_string(txn.account)}, "
            f"DateSerial({txn.when.year}, {txn.when.month}, {txn.when.day}), "
            f"{txn.amount}, {vbahost.basic_string(txn.description)})"
            for txn in records
        ]
        got = [row[1] for row in vbahost.rows(vbahost.run(_probe(
            calls, kind="String", emit="result")))]
        self.assertEqual(got, [
            sample.match_key(txn.account, txn.when, txn.amount, txn.description)
            for txn in records
        ])


class KeyedCollectionTests(unittest.TestCase):
    """The portable stand-in for Scripting.Dictionary."""

    def test_put_get_and_bump(self):
        body = '''
Sub Run()
    Dim bag As Collection
    Set bag = New Collection
    Emit "missing", modUtil.GetVal(bag, "nope", -1)
    Emit "haskey", Flag(modUtil.HasKey(bag, "nope"))
    modUtil.PutVal bag, "a", 1
    modUtil.PutVal bag, "a", 2
    Emit "overwritten", modUtil.GetVal(bag, "a", -1), bag.Count
    modUtil.BumpVal bag, "b"
    modUtil.BumpVal bag, "b"
    Emit "bumped", modUtil.GetVal(bag, "b", -1)
    Emit "haskey2", Flag(modUtil.HasKey(bag, "a"))
End Sub
'''
        rows = dict((row[0], row[1:]) for row in vbahost.rows(vbahost.run(body)))
        self.assertEqual(rows["missing"], ["-1"])
        self.assertEqual(rows["haskey"], ["no"])
        self.assertEqual(rows["overwritten"], ["2", "1"])
        self.assertEqual(rows["bumped"], ["2"])
        self.assertEqual(rows["haskey2"], ["yes"])


def _probe(calls, kind: str, emit: str, declare: str = "") -> str:
    """A ``Sub Run`` that calls each expression in turn and emits the result."""
    lines = [f"    result = {call}\n    Emit {index}, {emit}"
             for index, call in enumerate(calls, start=1)]
    return (f"Sub Run()\n    Dim result As {kind}\n"
            + (f"    {declare}\n" if declare else "")
            + "\n".join(lines) + "\nEnd Sub\n")


if __name__ == "__main__":
    unittest.main()
