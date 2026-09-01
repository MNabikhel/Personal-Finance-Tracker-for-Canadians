Attribute VB_Name = "modImport"
Option Explicit

'== Statement import =========================================================
' Reads one or more bank/credit-card CSV exports, normalises them into the
' ledger's shape, skips rows that are already present and logs every batch.
'
' Sign convention in the ledger: money leaving an account is negative, money
' arriving is positive - regardless of how the bank chose to express it.
'=============================================================================

Public Type TxnRecord
    TxnDate As Date
    Description As String
    Merchant As String
    Amount As Double
    Account As String
    Owner As String
    DupKey As String
    SourceFile As String
End Type

Public Sub ImportStatements()
    Dim chosen As Variant
    Dim i As Long
    Dim totalImported As Long, totalSkipped As Long, filesDone As Long
    Dim summary As String
    Dim batchStamp As String
    Dim line As String

    On Error GoTo Fail

    If Not modSetup.EnsureConfigured() Then Exit Sub

    chosen = Application.GetOpenFilename( _
        FileFilter:="Bank exports (*.csv;*.txt),*.csv;*.txt,All files (*.*),*.*", _
        Title:="Select one or more bank or credit card exports", _
        MultiSelect:=True)

    If VarType(chosen) = vbBoolean Then Exit Sub   ' user cancelled

    batchStamp = Format$(Now, "yyyymmdd-hhnnss")

    For i = LBound(chosen) To UBound(chosen)
        line = ImportOneFile(CStr(chosen(i)), "B" & batchStamp & "-" & i, _
                             totalImported, totalSkipped)
        If Len(line) > 0 Then
            filesDone = filesDone + 1
            summary = summary & line & vbCrLf
        End If
    Next i

    If filesDone = 0 Then
        MsgBox "Nothing was imported.", vbInformation, APP_NAME
        Exit Sub
    End If

    If totalImported > 0 Then
        modUtil.FastMode True
        modRules.CategorizeUncategorized False
        modTransfers.DetectTransfers False
        modUtil.FastMode False
        Application.Calculate
        modReport.RefreshAll
    End If

    MsgBox summary & vbCrLf & _
           "Added " & totalImported & " transaction(s) and skipped " & totalSkipped & _
           " duplicate or unreadable row(s)." & vbCrLf & vbCrLf & _
           "Anything the rules could not place is left as """ & CAT_UNCATEGORIZED & _
           """ on the Transactions sheet.", vbInformation, APP_NAME
    Exit Sub

Fail:
    modUtil.ReportError "ImportStatements"
End Sub

Private Function ImportOneFile(ByVal path As String, ByVal batchId As String, _
                               ByRef totalImported As Long, _
                               ByRef totalSkipped As Long) As String
    Dim text As String
    Dim rows As Collection
    Dim profileRow As Long
    Dim accountName As String
    Dim fileName As String
    Dim answer As VbMsgBoxResult
    Dim records() As TxnRecord
    Dim recordCount As Long
    Dim readCount As Long, dupeCount As Long, badCount As Long

    fileName = FileNameOnly(path)
    text = modParse.ReadTextFile(path)
    If Len(text) = 0 Then
        MsgBox fileName & " is empty.", vbExclamation, APP_NAME
        Exit Function
    End If

    Set rows = modParse.SplitRows(text, ",")
    If rows.Count = 0 Then
        MsgBox fileName & " has no data rows.", vbExclamation, APP_NAME
        Exit Function
    End If
    profileRow = modProfiles.DetectProfile(rows)

    Do
        If profileRow = 0 Then
            profileRow = modProfiles.AskForProfile(fileName)
            If profileRow = 0 Then Exit Function
        End If

        Set rows = modParse.SplitRows(text, modProfiles.DelimiterOf(profileRow))
        answer = MsgBox("Check the first rows of " & fileName & ":" & vbCrLf & vbCrLf & _
                        modProfiles.PreviewText(rows, profileRow) & vbCrLf & _
                        "Yes = import using this format" & vbCrLf & _
                        "No = pick a different format" & vbCrLf & _
                        "Cancel = skip this file", _
                        vbYesNoCancel + vbQuestion, APP_NAME)
        If answer = vbCancel Then Exit Function
        If answer = vbYes Then Exit Do
        profileRow = 0
    Loop

    accountName = modAccounts.ResolveAccount(fileName, profileRow)
    If Len(accountName) = 0 Then Exit Function

    modUtil.FastMode True
    modUtil.Status "reading " & fileName & " ..."
    ParseRows rows, profileRow, accountName, fileName, records, recordCount, _
              readCount, badCount

    If recordCount > 0 Then
        recordCount = DropDuplicates(records, recordCount, dupeCount)
        If recordCount > 0 Then AppendRecords records, recordCount, batchId
    End If
    modUtil.FastMode False

    If recordCount = 0 And dupeCount = 0 Then
        MsgBox "No usable transactions were found in " & fileName & "." & vbCrLf & _
               "Check the Date/Amount column numbers for this format on the " & _
               "Bank Formats sheet.", vbExclamation, APP_NAME
        Exit Function
    End If

    totalImported = totalImported + recordCount
    totalSkipped = totalSkipped + dupeCount + badCount

    modAccounts.LogBatch batchId, fileName, _
                         modProfiles.ProfileValue(profileRow, PF_NAME), accountName, _
                         readCount, recordCount, dupeCount, badCount

    ImportOneFile = fileName & ": " & recordCount & " added, " & dupeCount & _
                    " duplicate(s), " & badCount & " unreadable row(s)."
End Function

'--- Row level parsing ------------------------------------------------------

Public Function BuildRecord(ByRef values() As String, ByVal profileRow As Long, _
                            ByRef ok As Boolean) As TxnRecord
    Dim record As TxnRecord
    Dim dateOk As Boolean, amountOk As Boolean
    Dim amount As Double, debit As Double, credit As Double
    Dim mode As String

    ok = False
    record.TxnDate = modParse.ParseDate( _
        modParse.FieldAt(values, modProfiles.ProfileNumber(profileRow, PF_DATE_COL, 1)), _
        modProfiles.ProfileValue(profileRow, PF_DATE_FMT), dateOk)
    If Not dateOk Then
        BuildRecord = record
        Exit Function
    End If

    record.Description = modParse.FieldsAt(values, _
        modProfiles.ProfileValue(profileRow, PF_DESC_COLS))

    mode = modProfiles.ProfileValue(profileRow, PF_AMOUNT_MODE)
    If StrComp(mode, MODE_DEBIT_CREDIT, vbTextCompare) = 0 Then
        debit = modParse.ParseAmount(modParse.FieldAt(values, _
            modProfiles.ProfileNumber(profileRow, PF_DEBIT_COL, 0)), amountOk)
        If Not amountOk Then debit = 0
        credit = modParse.ParseAmount(modParse.FieldAt(values, _
            modProfiles.ProfileNumber(profileRow, PF_CREDIT_COL, 0)), amountOk)
        If Not amountOk Then credit = 0
        amount = Abs(credit) - Abs(debit)
        amountOk = (debit <> 0 Or credit <> 0)
    Else
        amount = modParse.ParseAmount(modParse.FieldAt(values, _
            modProfiles.ProfileNumber(profileRow, PF_AMOUNT_COL, 0)), amountOk)
        If StrComp(mode, MODE_SIGNED_FLIP, vbTextCompare) = 0 Then amount = -amount
    End If

    If Not amountOk Then
        BuildRecord = record
        Exit Function
    End If

    record.Amount = amount
    record.Merchant = modRules.CleanMerchant(record.Description)
    ok = True
    BuildRecord = record
End Function

Private Sub ParseRows(ByVal rows As Collection, ByVal profileRow As Long, _
                      ByVal accountName As String, ByVal fileName As String, _
                      ByRef records() As TxnRecord, ByRef recordCount As Long, _
                      ByRef readCount As Long, ByRef badCount As Long)
    Dim values() As String
    Dim record As TxnRecord
    Dim ok As Boolean
    Dim i As Long, skipRows As Long, capacity As Long
    Dim ownerName As String

    skipRows = modProfiles.ProfileNumber(profileRow, PF_SKIP, 0)
    ownerName = modAccounts.OwnerOfAccount(accountName)
    capacity = rows.Count
    If capacity < 1 Then capacity = 1
    ReDim records(1 To capacity)
    recordCount = 0

    For i = skipRows + 1 To rows.Count
        values = rows.Item(i)
        readCount = readCount + 1
        record = BuildRecord(values, profileRow, ok)
        If ok Then
            record.Account = accountName
            record.Owner = ownerName
            record.SourceFile = fileName
            record.DupKey = modUtil.MatchKey(accountName, record.TxnDate, _
                                             record.Amount, record.Description)
            recordCount = recordCount + 1
            records(recordCount) = record
        Else
            badCount = badCount + 1
        End If
    Next i
End Sub

'--- Duplicate handling -----------------------------------------------------

' Keeps only the rows that go beyond what the ledger already holds for a given
' key, so re-downloading an overlapping statement adds nothing while a genuine
' pair of identical same-day purchases still both land.
Private Function DropDuplicates(ByRef records() As TxnRecord, ByVal count As Long, _
                                ByRef dupeCount As Long) As Long
    Dim existing As Collection, seen As Collection
    Dim i As Long, kept As Long
    Dim key As String

    If StrComp(modUtil.NzStr(modUtil.Setting(NR_DUPES), "Yes"), "Yes", vbTextCompare) <> 0 Then
        DropDuplicates = count
        Exit Function
    End If

    Set existing = ExistingKeyCounts()
    Set seen = New Collection

    For i = 1 To count
        key = records(i).DupKey
        modUtil.BumpVal seen, key
        If modUtil.GetVal(seen, key, 0) <= modUtil.GetVal(existing, key, 0) Then
            dupeCount = dupeCount + 1
        Else
            kept = kept + 1
            If kept <> i Then records(kept) = records(i)
        End If
    Next i

    DropDuplicates = kept
End Function

Private Function ExistingKeyCounts() As Collection
    Dim lo As ListObject
    Dim keys As Variant
    Dim i As Long
    Dim counts As Collection

    Set counts = New Collection
    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then
        Set ExistingKeyCounts = counts
        Exit Function
    End If

    keys = lo.ListColumns(COL_KEY).DataBodyRange.Value
    If Not IsArray(keys) Then
        If Len(Trim$(CStr(keys))) > 0 Then modUtil.BumpVal counts, CStr(keys)
    Else
        For i = 1 To UBound(keys, 1)
            If Len(Trim$(CStr(keys(i, 1)))) > 0 Then
                modUtil.BumpVal counts, CStr(keys(i, 1))
            End If
        Next i
    End If
    Set ExistingKeyCounts = counts
End Function

'--- Writing to the ledger --------------------------------------------------

Private Sub AppendRecords(ByRef records() As TxnRecord, ByVal count As Long, _
                          ByVal batchId As String)
    Dim lo As ListObject
    Dim firstRow As Long
    Dim i As Long
    Dim ids() As Variant, dates() As Variant, accounts() As Variant
    Dim owners() As Variant, descriptions() As Variant, merchants() As Variant
    Dim amounts() As Variant, categories() As Variant, sources() As Variant
    Dim batches() As Variant, keys() As Variant, tags() As Variant
    Dim nextId As Long

    Set lo = modUtil.TxnTable()
    nextId = modUtil.BodyRows(lo) + 1
    firstRow = modLedger.AddRows(lo, count)

    ReDim ids(1 To count, 1 To 1)
    ReDim dates(1 To count, 1 To 1)
    ReDim accounts(1 To count, 1 To 1)
    ReDim owners(1 To count, 1 To 1)
    ReDim descriptions(1 To count, 1 To 1)
    ReDim merchants(1 To count, 1 To 1)
    ReDim amounts(1 To count, 1 To 1)
    ReDim categories(1 To count, 1 To 1)
    ReDim sources(1 To count, 1 To 1)
    ReDim batches(1 To count, 1 To 1)
    ReDim keys(1 To count, 1 To 1)
    ReDim tags(1 To count, 1 To 1)

    For i = 1 To count
        ids(i, 1) = "T" & Format$(nextId + i - 1, "000000")
        dates(i, 1) = CDbl(records(i).TxnDate)
        accounts(i, 1) = records(i).Account
        owners(i, 1) = records(i).Owner
        descriptions(i, 1) = records(i).Description
        merchants(i, 1) = records(i).Merchant
        amounts(i, 1) = records(i).Amount
        categories(i, 1) = CAT_UNCATEGORIZED
        sources(i, 1) = records(i).SourceFile
        batches(i, 1) = batchId
        keys(i, 1) = records(i).DupKey
        tags(i, 1) = TAG_IMPORT
    Next i

    modLedger.WriteColumn lo, firstRow, count, COL_ID, ids
    modLedger.WriteColumn lo, firstRow, count, COL_DATE, dates
    modLedger.WriteColumn lo, firstRow, count, COL_ACCOUNT, accounts
    modLedger.WriteColumn lo, firstRow, count, COL_OWNER, owners
    modLedger.WriteColumn lo, firstRow, count, COL_DESC, descriptions
    modLedger.WriteColumn lo, firstRow, count, COL_MERCHANT, merchants
    modLedger.WriteColumn lo, firstRow, count, COL_AMOUNT, amounts
    modLedger.WriteColumn lo, firstRow, count, COL_CATEGORY, categories
    modLedger.WriteColumn lo, firstRow, count, COL_SOURCE, sources
    modLedger.WriteColumn lo, firstRow, count, COL_BATCH, batches
    modLedger.WriteColumn lo, firstRow, count, COL_KEY, keys
    modLedger.WriteColumn lo, firstRow, count, COL_TAGGEDBY, tags
End Sub

Public Function FileNameOnly(ByVal path As String) As String
    Dim separatorAt As Long
    separatorAt = InStrRev(path, Application.PathSeparator)
    If separatorAt = 0 Then separatorAt = InStrRev(path, "\")
    If separatorAt = 0 Then separatorAt = InStrRev(path, "/")
    If separatorAt = 0 Then
        FileNameOnly = path
    Else
        FileNameOnly = Mid$(path, separatorAt + 1)
    End If
End Function
