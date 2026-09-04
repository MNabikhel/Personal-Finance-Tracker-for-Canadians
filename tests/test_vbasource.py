"""Checks the VBA sources agree with themselves and with the builder.

Excel is the only thing that can compile this code, and it is not here, so
these tests stand in for the compiler: every qualified call has to resolve to a
public procedure that accepts that many arguments, every constant and class
member has to exist, and the sheet, table and column names the macros use have
to be the ones ``tools.workbook`` actually writes.  A rename on either side
breaks the workbook silently, and this is what notices.
"""

from __future__ import annotations

import os
import re
import sys
import unittest
from datetime import date
from typing import Dict, List, Set

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests import vbasource  # noqa: E402
from tools import data, vbaproject, workbook  # noqa: E402

MODULES: Dict[str, vbasource.Module] = vbasource.load()

# ThisWorkbook is both one of our modules and the host object every macro reads
# the workbook through, so it is not a qualifier we can resolve names against.
NAMES: Set[str] = set(MODULES) - vbasource.HOST_OBJECTS


class DeclarationTests(unittest.TestCase):
    def test_every_module_is_named_after_its_file(self):
        for module in MODULES.values():
            with self.subTest(module.path.name):
                self.assertEqual(module.name, module.path.stem)

    def test_every_module_turns_on_option_explicit(self):
        # Without it a mistyped name becomes a silent empty Variant.
        for module in MODULES.values():
            with self.subTest(module.name):
                self.assertRegex(module.text, vbasource.OPTION_EXPLICIT)

    def test_class_files_carry_the_class_preamble(self):
        # vbaproject.build reads this to decide whether a file is a class; the
        # workbook module is registered as a document module instead.
        for module in MODULES.values():
            if not module.is_class or module.name == "ThisWorkbook":
                continue
            with self.subTest(module.name):
                self.assertTrue(module.text.startswith("VERSION 1.0 CLASS"))

    def test_the_builder_places_every_module_deliberately(self):
        # Load order decides the bytes of vbaProject.bin, so a new module must
        # be given a place rather than landing wherever sorting puts it.
        standard = sorted(name for name, module in MODULES.items()
                          if not module.is_class)
        self.assertEqual(sorted(vbaproject.MODULE_ORDER), standard)

    def test_the_project_includes_the_classes_and_the_workbook_module(self):
        classes = {name for name, module in MODULES.items() if module.is_class}
        self.assertIn("ThisWorkbook", classes)
        self.assertEqual(classes - {"ThisWorkbook"},
                         {"clsTxn", "clsProfile", "clsRule"})


class CallTests(unittest.TestCase):
    """Every ``modX.Something`` has to exist, be public, and take the arguments."""

    def setUp(self):
        self.calls: List[vbasource.Reference] = []
        for module in MODULES.values():
            self.calls += vbasource.references(module, NAMES)

    def test_there_are_calls_to_check(self):
        # A parser that silently matched nothing would make the rest of this
        # class pass without checking anything.
        self.assertGreater(len(self.calls), 100)

    def test_every_qualified_call_resolves(self):
        missing = [str(call) for call in self.calls
                   if not MODULES[call.qualifier].member(call.name)]
        self.assertEqual(missing, [])

    def test_no_module_reaches_into_another_modules_privates(self):
        private = []
        for call in self.calls:
            procedure = MODULES[call.qualifier].procedures.get(call.name.lower())
            if procedure and not procedure.public and call.module != call.qualifier:
                private.append(str(call))
        self.assertEqual(private, [])

    def test_every_call_hands_over_the_right_number_of_arguments(self):
        wrong = []
        for call in self.calls:
            if call.arguments is None:
                continue
            procedure = MODULES[call.qualifier].procedures.get(call.name.lower())
            if procedure is None or not procedure.accepts(call.arguments):
                if procedure is not None:
                    wrong.append(
                        f"{call} takes {call.arguments} but "
                        f"{procedure.name} wants {procedure.required}"
                        f"..{procedure.limit}")
        self.assertEqual(wrong, [])

    def test_every_call_within_a_module_hands_over_the_right_number_too(self):
        wrong = []
        checked = 0
        for module in MODULES.values():
            for call in vbasource.local_references(module):
                procedure = module.procedures[call.name.lower()]
                checked += 1
                if not procedure.accepts(call.arguments):
                    wrong.append(f"{call} takes {call.arguments} but "
                                 f"{procedure.name} wants {procedure.required}"
                                 f"..{procedure.limit}")
        self.assertGreater(checked, 50, "the local calls were not being found")
        self.assertEqual(wrong, [])

    def test_no_module_calls_itself_through_its_own_name(self):
        # Harmless, but it reads as though another module were involved.
        self.assertEqual([str(call) for call in self.calls
                          if call.module == call.qualifier], [])


class ButtonTargetTests(unittest.TestCase):
    """What a button or shortcut names as a string is invisible to the compiler
    and only fails when clicked, so it is resolved here instead."""

    TARGET = re.compile(r'"(mod\w+)\.(\w+)"')

    def setUp(self):
        self.targets = self.TARGET.findall(MODULES["modUI"].text)

    def test_there_are_targets_to_check(self):
        # Eight dashboard buttons, nine ledger buttons and four shortcuts,
        # with some macros reachable from more than one place.
        self.assertGreaterEqual(len(self.targets), 21)

    def test_every_button_and_shortcut_runs_a_public_parameterless_sub(self):
        for qualifier, name in self.targets:
            with self.subTest(f"{qualifier}.{name}"):
                self.assertIn(qualifier, MODULES)
                procedure = MODULES[qualifier].procedures.get(name.lower())
                self.assertIsNotNone(procedure, "no such procedure")
                self.assertTrue(procedure.public)
                self.assertEqual(procedure.kind.lower(), "sub")
                self.assertEqual(procedure.required, 0,
                                 "Excel calls a button's macro with no arguments")

    def test_every_planned_button_is_present_with_the_right_action(self):
        expected = {
            "Import statements": "modImport.ImportStatements",
            "Apply rules": "modRules.CategorizeUncategorized",
            "Find transfers": "modTransfers.DetectTransfers",
            "Needs a category": "modReport.ShowUncategorized",
            "Refresh": "modReport.RefreshAll",
            "Setup wizard": "modSetup.RunSetupWizard",
            "Couple mode on/off": "modHousehold.ToggleHouseholdMode",
            "Help": "modSetup.ShowHelp",
            "Undo an import": "modImport.UndoImport",
            "Teach a rule": "modRules.TeachRuleFromSelection",
            "Set owner": "modHousehold.SetOwnerForSelection",
            "Show all rows": "modReport.ClearLedgerFilters",
            "Apply rules to all": "modRules.RecategorizeAll",
            "Rebuild formulas": "modLedger.RepairFormulas",
            "Start fresh": "modSetup.ClearAllTransactions",
            "Back to dashboard": "modReport.GoToDashboard",
        }
        source = MODULES["modUI"].text
        for label, target in expected.items():
            with self.subTest(label):
                self.assertIn(f'"{label}", "{target}"', source)

    def test_the_button_rows_fit_above_the_first_content_row(self):
        # Two rows of buttons hang from B3 on both sheets; the next row down
        # holds the month selector on one and the table header on the other.
        source = MODULES["modUI"].text

        def constant(name):
            return float(re.search(rf"Const {name} As \w+ = ([\d.]+)", source).group(1))

        rows_needed = 2 * constant("BTN_HEIGHT") + constant("BTN_GAP")
        built = workbook.build(date(2026, 9, 1))
        for sheet in (workbook.SH_DASHBOARD, workbook.SH_TXN):
            ws = built[sheet]
            clearance = sum(ws.row_dimensions[r].height or 15 for r in (3, 4, 5))
            with self.subTest(sheet):
                self.assertGreaterEqual(clearance, rows_needed + 6)


class ClassMemberTests(unittest.TestCase):
    """``txn.Amount`` has to be something clsTxn actually has."""

    def setUp(self):
        self.classes = {name: module for name, module in MODULES.items()
                        if name.startswith(vbasource.CLASS_PREFIX)}

    def test_every_member_read_off_a_typed_variable_exists(self):
        missing = []
        checked = 0
        for module in MODULES.values():
            for line, typed in vbasource.typed_locals(module):
                for found in re.finditer(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)", line):
                    holder = typed.get(found.group(1).lower())
                    if holder is None or holder not in self.classes:
                        continue
                    checked += 1
                    if not self.classes[holder].member(found.group(2)):
                        missing.append(f"{module.name}: {holder} has no "
                                       f"{found.group(2)} in {line!r}")
        self.assertGreater(checked, 40, "the variables were not being resolved")
        self.assertEqual(missing, [])

    def test_the_classes_expose_what_the_ledger_writes(self):
        for member in ("TxnDate", "Description", "Merchant", "Amount", "Account",
                       "Owner", "SourceFile", "MatchKey"):
            with self.subTest(member):
                self.assertTrue(MODULES["clsTxn"].member(member))


class ConstantTests(unittest.TestCase):
    def test_every_constant_looking_name_is_declared(self):
        declared = set()
        for module in MODULES.values():
            declared |= set(module.constants)
        # Enum-like names that come from the host rather than from us.
        host = {"VB_NAME", "VB_GLOBALNAMESPACE", "VB_CREATABLE",
                "VB_PREDECLAREDID", "VB_EXPOSED", "VB_BASE", "VB_TEMPLATEDERIVED",
                "VB_CUSTOMIZABLE"}
        unknown = []
        for module in MODULES.values():
            for line in module.lines:
                if line.startswith("Attribute "):
                    continue
                for found in vbasource.UPPER_CONSTANT.finditer(line):
                    token = found.group(1)
                    if token in host or token.lower() in declared:
                        continue
                    unknown.append(f"{module.name}: {token} in {line!r}")
        self.assertEqual(unknown, [])

    def test_the_workbook_constants_are_a_complete_map_of_its_surface(self):
        # These families name every sheet, table and ledger column, whether or
        # not a macro reads them yet, and the agreement tests below check each
        # one against the workbook the builder writes.  What must not happen is
        # a family losing an entry, which would leave that part of the workbook
        # spelled out in a string literal somewhere instead.
        constants = MODULES["modConst"].constants
        for prefix, count in (("sh_", len(workbook.CODE_NAMES)),
                              ("tbl_", 6),
                              ("col_", len(workbook.TXN_HEADERS))):
            with self.subTest(prefix):
                self.assertEqual(
                    len([name for name in constants if name.startswith(prefix)]),
                    count)


class BuilderAgreementTests(unittest.TestCase):
    """The names in modConst have to be the ones the builder writes."""

    def _constants(self, prefix: str) -> Dict[str, str]:
        out = {}
        for line in MODULES["modConst"].text.splitlines():
            found = re.match(rf'Public Const ({prefix}\w+) As String = "(.*)"', line)
            if found:
                out[found.group(1)] = found.group(2)
        return out

    def test_the_sheet_names_match(self):
        self.assertEqual(set(self._constants("SH_").values()),
                         set(workbook.CODE_NAMES))

    def test_the_table_names_match(self):
        built = set()
        wb = workbook.build(workbook.date(2026, 9, 1))
        for sheet in wb.worksheets:
            built |= set(sheet.tables)
        self.assertLessEqual(set(self._constants("TBL_").values()), built)

    def test_the_transaction_column_names_match(self):
        self.assertLessEqual(set(self._constants("COL_").values()),
                             set(workbook.TXN_HEADERS))

    def test_the_import_log_columns_are_all_addressed_by_header(self):
        built = workbook.build(workbook.date(2026, 9, 1))
        table = built[workbook.SH_LOG].tables["tblLog"]
        headers = {column.name for column in table.tableColumns}
        self.assertEqual(set(self._constants("LG_").values()), headers)

    def test_the_importer_fills_every_column_the_user_does_not(self):
        # A ledger column that is neither calculated, nor filled by the import,
        # nor left for the user on purpose would simply come out blank.
        columns = self._constants("COL_")
        written = {columns[found] for found in re.findall(
            r"modLedger\.WriteColumn[^\n]*?(COL_\w+)", MODULES["modImport"].code)}
        calculated = set(workbook.TXN_FORMULAS)
        user_entered = {"Reimbursable", "Notes"}
        self.assertEqual(written | calculated | user_entered,
                         set(workbook.TXN_HEADERS))
        self.assertEqual(written & calculated, set(),
                         "the importer must not fight the calculated columns")

    def test_the_named_ranges_exist_in_the_workbook(self):
        wb = workbook.build(workbook.date(2026, 9, 1))
        wanted = set(self._constants("NR_").values())
        self.assertEqual(wanted - set(wb.defined_names), set())

    def test_the_settings_values_match_the_dropdowns(self):
        constants = self._constants("MODE_")
        self.assertEqual({constants["MODE_SINGLE"], constants["MODE_COUPLE"]},
                         {"Single", "Couple"})
        self.assertLessEqual({constants["MODE_SIGNED"],
                              constants["MODE_SIGNED_FLIP"],
                              constants["MODE_DEBIT_CREDIT"]},
                             set(data.AMOUNT_MODES))

    def test_the_categories_the_macros_rely_on_exist(self):
        categories = {row[0] for row in data.CATEGORIES}
        for key, value in self._constants("CAT_").items():
            with self.subTest(key):
                self.assertIn(value, categories)

class WorkflowInvariantTests(unittest.TestCase):
    """Cross-module ordering that can silently change financial results."""

    def test_undo_rechecks_transfer_pairs_after_deleting_the_batch(self):
        body = MODULES["modImport"].code
        deleted = body.index("modLedger.DeleteRowsWhere")
        rechecked = body.index("modTransfers.DetectTransfers False", deleted)
        self.assertLess(deleted, rechecked)

    def test_reapplying_all_rules_restores_transfer_detection(self):
        body = MODULES["modRules"].code
        self.assertIn("If includeTagged Then modTransfers.DetectTransfers False", body)

    def test_biggest_merchants_nets_refunds_instead_of_adding_them(self):
        body = MODULES["modReport"].code
        self.assertIn("-modUtil.NzNum(views(i, 1))", body)
        self.assertNotIn("Abs(modUtil.NzNum(views(i, 1)))", body)
        self.assertIn("> 0 Then", body)

    def test_fast_mode_off_is_safe_before_fast_mode_was_started(self):
        body = MODULES["modUtil"].code
        self.assertIn("If mFastDepth = 0 Then Exit Sub", body)
        self.assertIn("Do While mFastDepth > 0", body)

    def test_each_import_gets_a_collision_checked_batch_id(self):
        body = MODULES["modImport"].code
        allocated = body.index("batchId = NewBatchId(batchStamp, i)")
        dispatched = min(body.index("modPdf.ImportOnePdf"),
                         body.index("ImportOneFile(CStr(chosen(i))"))
        self.assertLess(allocated, dispatched)
        self.assertIn("Do While BatchIdExists(candidate)", body)
        self.assertIn("Set lo = modUtil.TxnTable()", body)

    def test_setup_removes_only_marked_sample_accounts_when_starting_fresh(self):
        self.assertIn(
            "modAccounts.ClearSampleAccounts",
            MODULES["modSetup"].code,
        )
        accounts = MODULES["modAccounts"].text
        self.assertIn('"Sample account", vbTextCompare', accounts)
        self.assertIn("lo.DataBodyRange.Rows(i).ClearContents", accounts)

    def test_starting_fresh_also_clears_the_macro_written_merchant_list(self):
        source = MODULES["modSetup"].code
        cleared = source.index("modLedger.ClearRowsKeepOne lo")
        refreshed = source.index("modReport.RefreshTopMerchants", cleared)
        self.assertLess(cleared, refreshed)

    def test_dashboard_selectors_refresh_the_sorted_merchant_list(self):
        source = MODULES["ThisWorkbook"].code
        self.assertIn("Case SH_DASHBOARD", source)
        self.assertIn("DashboardChanged Target", source)
        self.assertGreaterEqual(source.count("modReport.RefreshTopMerchants"), 3)

    def test_no_header_exports_reuse_the_account_saved_format(self):
        source = MODULES["modImport"].code
        detected = source.index("Set profile = modProfiles.DetectProfile(rows)")
        saved = source.index("modProfiles.FindProfileByName", detected)
        asked = source.index("modProfiles.AskForProfile(fileName)", saved)
        self.assertLess(detected, saved)
        self.assertLess(saved, asked)

    def test_large_manual_category_pastes_are_not_silently_left_retaggable(self):
        source = MODULES["ThisWorkbook"].code
        self.assertNotIn("touched.Cells.Count >", source)

    def test_workbook_change_event_uses_excels_canonical_signature_and_recovers(self):
        source = MODULES["ThisWorkbook"].text
        self.assertIn(
            "Private Sub Workbook_SheetChange(ByVal Sh As Object, "
            "ByVal Target As Range)",
            source,
        )
        code = MODULES["ThisWorkbook"].code
        self.assertIn("On Error GoTo CleanUp", code)
        self.assertIn("Application.EnableEvents = True", code)


if __name__ == "__main__":
    unittest.main()
