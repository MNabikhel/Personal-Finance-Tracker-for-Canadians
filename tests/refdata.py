"""The workbook's own reference data, restated as Basic that rebuilds it.

The Rules and Bank Formats sheets are the macros' input, so running the real
import path outside Excel means handing the code the real rows.  What is
rendered here is exactly what ``modProfiles.Profile`` and ``modRules.LoadRules``
would have lifted off those sheets, built from ``tools.data`` rather than from a
fixture, so a test cannot quietly pass against rules the workbook does not
actually ship.
"""

from __future__ import annotations

from typing import List, Optional, Sequence

from tests import vbahost
from tools import data

# What LoadRules substitutes for an empty Min/Max Amount cell.
NO_LIMIT = 1e15


def _number(value: Optional[float]) -> str:
    if value is None:
        return "0"
    if float(value) == int(value):
        return str(int(value))
    return repr(float(value))


def _call(name: str, arguments: Sequence[str]) -> str:
    return f"    {name} out, " + ", ".join(arguments)


def profiles_basic() -> str:
    """``AllProfiles``/``ProfileNamed``, holding every shipped bank format."""
    calls: List[str] = []
    for index, row in enumerate(data.BANK_FORMATS, start=1):
        (name, institution, skip, delimiter, date_col, date_format,
         description_cols, amount_mode, amount_col, debit_col, credit_col,
         signature, _notes) = row
        calls.append(_call("AddProfile", [
            str(index),
            vbahost.basic_string(name),
            vbahost.basic_string(institution),
            _number(skip),
            vbahost.basic_string(delimiter),
            _number(date_col),
            vbahost.basic_string(date_format),
            vbahost.basic_string(description_cols),
            vbahost.basic_string(amount_mode),
            _number(amount_col),
            _number(debit_col),
            _number(credit_col),
            vbahost.basic_string(signature),
        ]))

    return PROFILE_SOURCE.replace("'{calls}", "\n".join(calls))


def rules_basic() -> str:
    """``AllRules``, holding every seed rule in the order the sheet lists them."""
    calls: List[str] = []
    for index, row in enumerate(data.seed_rules(), start=1):
        (priority, enabled, look_in, test, pattern, minimum, maximum, flow,
         category, owner, _hits, _notes) = row
        if not pattern or not category or str(enabled).lower() == "no":
            continue        # IsRuleUsable would have dropped it
        calls.append(_call("AddRule", [
            str(index),
            _number(priority),
            vbahost.basic_string(look_in),
            vbahost.basic_string(test),
            vbahost.basic_string(pattern),
            repr(-NO_LIMIT) if minimum is None else _number(minimum),
            repr(NO_LIMIT) if maximum is None else _number(maximum),
            vbahost.basic_string(flow),
            vbahost.basic_string(category),
            vbahost.basic_string(owner),
        ]))

    return RULE_SOURCE.replace("'{calls}", "\n".join(calls))


PROFILE_SOURCE = """Option VBASupport 1

' Every row of the Bank Formats sheet, as modProfiles.Profile would build it.
' Deliberately not called AllProfiles: modProfiles has a function by that name
' which reads the sheet, and an unqualified call would reach that one instead.
Function SheetProfiles() As Collection
    Dim out As Collection
    Set out = New Collection
'{calls}
    Set SheetProfiles = out
End Function

Function ProfileNamed(ByVal wanted As String) As clsProfile
    Dim all As Collection
    Dim i As Long
    Set all = SheetProfiles()
    For i = 1 To all.Count
        If all.Item(i).Name = wanted Then
            Set ProfileNamed = all.Item(i)
            Exit Function
        End If
    Next i
    Err.Raise 5, , "no bank format named " & wanted
End Function

Sub AddProfile(ByVal out As Collection, ByVal rowIndex As Long, _
               ByVal profileName As String, ByVal institution As String, _
               ByVal skipRows As Long, ByVal delimiterName As String, _
               ByVal dateColumn As Long, ByVal dateFormat As String, _
               ByVal descriptionColumns As String, ByVal amountMode As String, _
               ByVal amountColumn As Long, ByVal debitColumn As Long, _
               ByVal creditColumn As Long, ByVal signature As String)
    Dim profile As clsProfile
    Set profile = New clsProfile
    profile.RowIndex = rowIndex
    profile.Name = profileName
    profile.Institution = institution
    profile.SkipRows = skipRows
    profile.DelimiterName = delimiterName
    profile.DateColumn = dateColumn
    profile.DateFormat = dateFormat
    profile.DescriptionColumns = descriptionColumns
    profile.AmountMode = amountMode
    profile.AmountColumn = amountColumn
    profile.DebitColumn = debitColumn
    profile.CreditColumn = creditColumn
    profile.Signature = signature
    out.Add profile
End Sub
"""

RULE_SOURCE = """Option VBASupport 1

' Every usable row of the Rules sheet, as modRules.LoadRules would build it,
' in sheet order - sorting them is modRules.ByPriority's job, not ours.
Function SheetRules() As Collection
    Dim out As Collection
    Set out = New Collection
'{calls}
    Set SheetRules = out
End Function

Sub AddRule(ByVal out As Collection, ByVal rowIndex As Long, _
            ByVal priority As Double, ByVal lookIn As String, _
            ByVal test As String, ByVal pattern As String, _
            ByVal minAmount As Double, ByVal maxAmount As Double, _
            ByVal flow As String, ByVal category As String, _
            ByVal setOwner As String)
    Dim rule As clsRule
    Set rule = New clsRule
    rule.RowIndex = rowIndex
    rule.Priority = priority
    rule.LookIn = lookIn
    rule.Test = test
    rule.Pattern = pattern
    rule.MinAmount = minAmount
    rule.MaxAmount = maxAmount
    rule.Flow = flow
    rule.Category = category
    rule.SetOwner = setOwner
    out.Add rule
End Sub
"""
