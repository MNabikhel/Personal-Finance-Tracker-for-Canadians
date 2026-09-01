Attribute VB_Name = "modSetup"
Option Explicit

'== First-run setup and housekeeping ========================================

Public Sub RunSetupWizard()
    Dim modeChoice As Long
    Dim nameA As String, nameB As String
    Dim splitText As String
    Dim provinceChoice As Long
    Dim provinces As Variant
    Dim startFresh As VbMsgBoxResult

    On Error GoTo Fail

    modeChoice = modUtil.AskChoice( _
        "Who is this workbook for?", Array("Just me", "A couple / household"), _
        APP_NAME & " setup")
    If modeChoice = 0 Then Exit Sub

    nameA = Trim$(InputBox("First name of the main person:", APP_NAME, _
                           modUtil.NzStr(modUtil.Setting(NR_PERSON_A), "Person A")))
    If Len(nameA) = 0 Then Exit Sub
    modUtil.SetSetting NR_PERSON_A, nameA

    If modeChoice = 2 Then
        nameB = Trim$(InputBox("First name of the second person:", APP_NAME, _
                               modUtil.NzStr(modUtil.Setting(NR_PERSON_B), "Person B")))
        If Len(nameB) = 0 Then nameB = "Person B"
        modUtil.SetSetting NR_PERSON_B, nameB
        modUtil.SetSetting NR_MODE, MODE_COUPLE

        splitText = Trim$(InputBox( _
            "For shared (Joint) expenses, what share belongs to " & nameA & "?" & _
            vbCrLf & vbCrLf & "Enter a percentage, for example 50 for an even split.", _
            APP_NAME, CStr(CLng(modUtil.NzNum(modUtil.Setting(NR_SPLIT), 0.5) * 100))))
        If IsNumeric(splitText) Then
            modUtil.SetSetting NR_SPLIT, _
                Application.WorksheetFunction.Median(0, Val(splitText) / 100, 1)
        End If
    Else
        modUtil.SetSetting NR_MODE, MODE_SINGLE
    End If

    provinces = Array("AB", "BC", "MB", "NB", "NL", "NS", "NT", "NU", "ON", _
                      "PE", "QC", "SK", "YT")
    provinceChoice = modUtil.AskChoice("Which province or territory do you live in?", _
                                       provinces)
    If provinceChoice > 0 Then
        modUtil.SetSetting NR_PROVINCE, provinces(provinceChoice - 1)
    End If

    modUtil.SetSetting NR_CONFIGURED, "Yes"
    modHousehold.ApplyMode
    modUI.EnsureButtons

    If HasSampleData() Then
        startFresh = MsgBox("This workbook still contains the sample transactions." & _
            vbCrLf & vbCrLf & "Delete them and start with an empty ledger?", _
            vbYesNo + vbQuestion, APP_NAME)
        If startFresh = vbYes Then ClearAllTransactions False
    End If

    modUtil.Sh(SH_ACCOUNTS).Activate
    MsgBox "Setup is done." & vbCrLf & vbCrLf & _
           "Next: list your accounts here (one row per bank or card account), then " & _
           "use Import statements on the Dashboard.", vbInformation, APP_NAME
    Exit Sub

Fail:
    modUtil.ReportError "RunSetupWizard"
End Sub

Public Function EnsureConfigured() As Boolean
    If StrComp(modUtil.NzStr(modUtil.Setting(NR_CONFIGURED)), "Yes", vbTextCompare) = 0 Then
        EnsureConfigured = True
        Exit Function
    End If

    If MsgBox("Let's set the workbook up first - it only takes a moment." & vbCrLf & _
              vbCrLf & "Run setup now?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
        RunSetupWizard
        EnsureConfigured = (StrComp(modUtil.NzStr(modUtil.Setting(NR_CONFIGURED)), _
                                    "Yes", vbTextCompare) = 0)
    Else
        ' Nothing here is mandatory; carry on with the defaults.
        EnsureConfigured = True
    End If
End Function

Public Function HasSampleData() As Boolean
    Dim lo As ListObject
    Dim sources As Variant
    Dim i As Long

    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then Exit Function
    sources = modLedger.ReadColumn(lo, COL_SOURCE)
    For i = 1 To UBound(sources, 1)
        If InStr(1, modUtil.NzStr(sources(i, 1)), "sample", vbTextCompare) > 0 Then
            HasSampleData = True
            Exit Function
        End If
    Next i
End Function

Public Sub ClearAllTransactions(Optional ByVal confirm As Boolean = True)
    Dim lo As ListObject
    Dim logTable As ListObject
    Dim rules As ListObject
    Dim i As Long

    On Error GoTo Fail

    If confirm Then
        If MsgBox("Delete every transaction in the ledger?" & vbCrLf & vbCrLf & _
                  "Accounts, categories, rules and settings are kept.", _
                  vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub
        If MsgBox("This cannot be undone. Delete them?", _
                  vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub
    End If

    modUtil.FastMode True
    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) > 1 Then
        lo.DataBodyRange.Offset(1, 0).Resize(modUtil.BodyRows(lo) - 1).Rows.Delete
    End If
    If modUtil.BodyRows(lo) = 1 Then
        lo.DataBodyRange.Rows(1).ClearContents
        modLedger.ApplyTemplates lo, 1, 1
    End If

    On Error Resume Next
    Set logTable = modUtil.Tbl(SH_LOG, TBL_LOG)
    If Not logTable Is Nothing Then
        If modUtil.BodyRows(logTable) > 1 Then
            logTable.DataBodyRange.Offset(1, 0) _
                .Resize(modUtil.BodyRows(logTable) - 1).Rows.Delete
        End If
        logTable.DataBodyRange.Rows(1).ClearContents
    End If

    Set rules = modRules.RulesTable()
    If Not rules Is Nothing Then
        For i = 1 To modUtil.BodyRows(rules)
            rules.DataBodyRange.Cells(i, modUtil.ColumnIndex(rules, RL_HITS)).Value = 0
        Next i
    End If
    On Error GoTo 0

    modUtil.FastMode False
    Application.Calculate
    If confirm Then MsgBox "The ledger is empty and ready for your own data.", _
                           vbInformation, APP_NAME
    Exit Sub

Fail:
    modUtil.ReportError "ClearAllTransactions"
End Sub

Public Sub ShowHelp()
    On Error Resume Next
    modUtil.Sh(SH_HELP).Activate
    modUtil.Sh(SH_HELP).Range("A1").Select
End Sub
