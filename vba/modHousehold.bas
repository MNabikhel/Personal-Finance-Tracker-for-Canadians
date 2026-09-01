Attribute VB_Name = "modHousehold"
Option Explicit

'== Couple / household mode =================================================
' Single mode keeps the workbook simple: one owner, no splitting.  Couple mode
' turns on per-person shares, a fair-share settlement and the Household sheet.
'=============================================================================

Public Sub ToggleHouseholdMode()
    Dim wantCouple As Boolean
    On Error GoTo Fail

    wantCouple = Not modUtil.IsCoupleMode()

    If wantCouple Then
        If Len(modUtil.NzStr(modUtil.Setting(NR_PERSON_B))) = 0 Then
            modUtil.SetSetting NR_PERSON_B, Trim$(InputBox( _
                "Name of the second person:", APP_NAME, "Person B"))
        End If
        modUtil.SetSetting NR_MODE, MODE_COUPLE
    Else
        If MsgBox("Switch back to single mode?" & vbCrLf & vbCrLf & _
                  "Per-person columns and the Household sheet get hidden. " & _
                  "Nothing is deleted.", vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
            Exit Sub
        End If
        modUtil.SetSetting NR_MODE, MODE_SINGLE
    End If

    ApplyMode
    MsgBox "Household mode is now: " & modUtil.NzStr(modUtil.Setting(NR_MODE)) & ".", _
           vbInformation, APP_NAME
    Exit Sub
Fail:
    modUtil.ReportError "ToggleHouseholdMode"
End Sub

' Shows or hides everything that only makes sense for two people.
Public Sub ApplyMode()
    Dim couple As Boolean
    Dim lo As ListObject

    On Error Resume Next
    couple = modUtil.IsCoupleMode()

    Set lo = modUtil.TxnTable()
    If Not lo Is Nothing Then
        lo.ListColumns(COL_SPLIT).Range.EntireColumn.Hidden = Not couple
        lo.ListColumns(COL_SHARE_A).Range.EntireColumn.Hidden = Not couple
        lo.ListColumns(COL_SHARE_B).Range.EntireColumn.Hidden = Not couple
        ' Plumbing for the View selector, never worth looking at.
        lo.ListColumns(COL_VIEW).Range.EntireColumn.Hidden = True
    End If

    modUtil.Sh(SH_HOUSEHOLD).Visible = IIf(couple, xlSheetVisible, xlSheetHidden)
    ' Settling up between two people reads as nonsense when there is only one.
    modUtil.Sh(SH_DASHBOARD).Range(NR_COUPLE_BLOCK).EntireRow.Hidden = Not couple

    ' The Dashboard "View" selector only offers a person split for couples.
    With modUtil.Sh(SH_DASHBOARD).Range("ReportView").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=IIf(couple, "=ViewList", "=HouseholdOnly")
    End With
    If Not couple Then modUtil.Sh(SH_DASHBOARD).Range("ReportView").Value = "Household"
    On Error GoTo 0
End Sub

' Bulk-assigns the owner of the selected transaction rows.
Public Sub SetOwnerForSelection()
    Dim lo As ListObject
    Dim area As Range, cell As Range
    Dim owners As Variant
    Dim choice As Long
    Dim column As Long
    Dim touched As Long

    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    Set area = Application.Intersect(Application.Selection.EntireRow, lo.DataBodyRange)
    If area Is Nothing Then
        MsgBox "Select one or more transaction rows first.", vbInformation, APP_NAME
        Exit Sub
    End If

    If modUtil.IsCoupleMode() Then
        owners = Array(modUtil.PersonAName(), modUtil.PersonBName(), OWNER_JOINT)
    Else
        owners = Array(modUtil.PersonAName())
    End If

    choice = modUtil.AskChoice("Set the owner of the selected rows to:", owners)
    If choice = 0 Then Exit Sub

    column = modUtil.ColumnIndex(lo, COL_OWNER)
    modUtil.FastMode True
    For Each cell In Application.Intersect(area, lo.ListColumns(column).DataBodyRange)
        cell.Value = owners(choice - 1)
        touched = touched + 1
    Next cell
    modUtil.FastMode False
    Application.Calculate

    MsgBox touched & " row(s) now belong to " & owners(choice - 1) & ".", _
           vbInformation, APP_NAME
    Exit Sub
Fail:
    modUtil.ReportError "SetOwnerForSelection"
End Sub
