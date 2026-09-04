Attribute VB_Name = "modImport"
Option Explicit

'== Statement import =========================================================
' Reads one or more bank/credit-card statements - CSV exports here, PDFs by
' way of modPdf - normalises them into the ledger's shape, skips rows that are
' already present and logs every batch.
'
' Sign convention in the ledger: money leaving an account is negative, money
' arriving is positive - regardless of how the bank chose to express it.
'=============================================================================

Public Sub ImportStatements()
    Dim chosen As Variant
    Dim i As Long
    Dim totalImported As Long, totalSkipped As Long, filesDone As Long
    Dim summary As String
    Dim batchStamp As String
    Dim batchId As String
    Dim line As String

    On Error GoTo Fail

    If Not modSetup.EnsureConfigured() Then Exit Sub

    chosen = Application.GetOpenFilename( _
        FileFilter:="Bank statements (*.csv;*.txt;*.pdf),*.csv;*.txt;*.pdf," & _
                    "CSV exports (*.csv;*.txt),*.csv;*.txt," & _
                    "PDF statements (*.pdf),*.pdf,All files (*.*),*.*", _
        Title:="Select one or more bank or credit card statements", _
        MultiSelect:=True)

    If VarType(chosen) = vbBoolean Then Exit Sub   ' user cancelled

    batchStamp = Format$(Now, "yyyymmdd-hhnnss")

    For i = LBound(chosen) To UBound(chosen)
        batchId = NewBatchId(batchStamp, i)
        If modPdf.IsPdf(CStr(chosen(i))) Then
            line = modPdf.ImportOnePdf(CStr(chosen(i)), batchId, _
                                       totalImported, totalSkipped)
        Else
            line = ImportOneFile(CStr(chosen(i)), batchId, _
                                 totalImported, totalSkipped)
        End If
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

' A user can finish one import and start another within the same second.  A
' timestamp plus file index would then repeat, and undoing the second log row
' would also delete the first import.  Existing log and ledger rows settle the
' collision before anything is appended.
Private Function NewBatchId(ByVal stamp As String, ByVal fileIndex As Long) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long

    baseName = "B" & stamp & "-" & fileIndex
    candidate = baseName
    suffix = 1
    Do While BatchIdExists(candidate)
        suffix = suffix + 1
        candidate = baseName & "-" & suffix
    Loop
    NewBatchId = candidate
End Function

Private Function BatchIdExists(ByVal batchId As String) As Boolean
    Dim lo As ListObject
    Dim values As Variant
    Dim i As Long

    On Error Resume Next
    Set lo = modUtil.Tbl(SH_LOG, TBL_LOG)
    On Error GoTo 0
    If Not lo Is Nothing Then
        values = modLedger.ReadColumn(lo, modAccounts.LG_BATCH)
        If IsArray(values) Then
            For i = 1 To UBound(values, 1)
                If StrComp(modUtil.NzStr(values(i, 1)), batchId, vbTextCompare) = 0 Then
                    BatchIdExists = True
                    Exit Function
                End If
            Next i
        End If
    End If

    ' A previous append may have succeeded immediately before a logging error.
    Set lo = Nothing
    On Error Resume Next
    Set lo = modUtil.TxnTable()
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    values = modLedger.ReadColumn(lo, COL_BATCH)
    If Not IsArray(values) Then Exit Function
    For i = 1 To UBound(values, 1)
        If StrComp(modUtil.NzStr(values(i, 1)), batchId, vbTextCompare) = 0 Then
            BatchIdExists = True
            Exit Function
        End If
    Next i
End Function

Private Function ImportOneFile(ByVal path As String, ByVal batchId As String, _
                               ByRef totalImported As Long, _
                               ByRef totalSkipped As Long) As String
    Dim text As String
    Dim rows As Collection
    Dim profile As clsProfile
    Dim accountName As String
    Dim fileName As String
    Dim answer As VbMsgBoxResult
    Dim records As Collection
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
    Set profile = modProfiles.DetectProfile(rows)
    If profile Is Nothing Then
        Set profile = modProfiles.FindProfileByName( _
            modAccounts.FormatForFileName(fileName))
    End If

    Do
        If profile Is Nothing Then
            Set profile = modProfiles.AskForProfile(fileName)
            If profile Is Nothing Then Exit Function
        End If

        Set rows = modParse.SplitRows(text, profile.Delimiter())
        answer = MsgBox("Check the first rows of " & fileName & ":" & vbCrLf & vbCrLf & _
                        modProfiles.PreviewText(rows, profile) & vbCrLf & _
                        "Yes = import using this format" & vbCrLf & _
                        "No = pick a different format" & vbCrLf & _
                        "Cancel = skip this file", _
                        vbYesNoCancel + vbQuestion, APP_NAME)
        If answer = vbCancel Then Exit Function
        If answer = vbYes Then Exit Do
        Set profile = Nothing
    Loop

    accountName = modAccounts.ResolveAccount(fileName, profile)
    If Len(accountName) = 0 Then Exit Function

    modUtil.FastMode True
    modUtil.Status "reading " & fileName & " ..."
    Set records = ReadRecords(rows, profile, accountName, _
                              modAccounts.OwnerOfAccount(accountName), _
                              fileName, readCount, badCount)
    Set records = WithoutDuplicates(records, ExistingKeyCounts(), dupeCount)
    If records.Count > 0 Then AppendRecords records, batchId
    modUtil.FastMode False

    If records.Count = 0 And dupeCount = 0 Then
        MsgBox "No usable transactions were found in " & fileName & "." & vbCrLf & _
               "Check the Date/Amount column numbers for this format on the " & _
               "Bank Formats sheet.", vbExclamation, APP_NAME
        Exit Function
    End If

    totalImported = totalImported + records.Count
    totalSkipped = totalSkipped + dupeCount + badCount

    modAccounts.LogBatch batchId, fileName, profile.Name, accountName, _
                         readCount, records.Count, dupeCount, badCount

    ImportOneFile = fileName & ": " & records.Count & " added, " & dupeCount & _
                    " duplicate(s), " & badCount & " unreadable row(s)."
End Function

'--- Reading ----------------------------------------------------------------

' Every readable row of the file.  Header lines and bank notices do not parse
' as a transaction and are counted as unreadable rather than stopping the
' import: a BMO export, for one, opens with three lines of prose.
Public Function ReadRecords(ByVal rows As Collection, ByVal profile As clsProfile, _
                            ByVal accountName As String, ByVal ownerName As String, _
                            ByVal fileName As String, _
                            ByRef readCount As Long, ByRef badCount As Long) As Collection
    Dim out As Collection
    Dim txn As clsTxn
    Dim i As Long

    Set out = New Collection
    Set ReadRecords = out

    For i = profile.SkipRows + 1 To rows.Count
        readCount = readCount + 1
        Set txn = profile.ReadRow(rows.Item(i))
        If txn Is Nothing Then
            badCount = badCount + 1
        Else
            txn.Account = accountName
            txn.Owner = ownerName
            txn.SourceFile = fileName
            out.Add txn
        End If
    Next i
End Function

'--- Duplicate handling -----------------------------------------------------

' Keeps only the rows that go beyond what the ledger already holds for a given
' key, so re-downloading an overlapping statement adds nothing while a genuine
' pair of identical same-day purchases still both land.
Public Function WithoutDuplicates(ByVal records As Collection, _
                                  ByVal existing As Collection, _
                                  ByRef dupeCount As Long) As Collection
    Dim out As Collection
    Dim seen As Collection
    Dim i As Long
    Dim key As String

    Set out = New Collection
    Set WithoutDuplicates = out
    Set seen = New Collection

    For i = 1 To records.Count
        key = records.Item(i).MatchKey()
        modUtil.BumpVal seen, key
        If modUtil.GetVal(seen, key, 0) <= modUtil.GetVal(existing, key, 0) Then
            dupeCount = dupeCount + 1
        Else
            out.Add records.Item(i)
        End If
    Next i
End Function

' How many times each match key already appears in the ledger.  Empty when the
' user has turned duplicate skipping off, which lets everything through.
Public Function ExistingKeyCounts() As Collection
    Dim lo As ListObject
    Dim keys As Variant
    Dim i As Long
    Dim counts As Collection

    Set counts = New Collection
    Set ExistingKeyCounts = counts

    If StrComp(modUtil.NzStr(modUtil.Setting(NR_DUPES), "Yes"), "Yes", _
               vbTextCompare) <> 0 Then Exit Function

    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then Exit Function

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
End Function

'--- Writing to the ledger --------------------------------------------------

' Adds the records to the ledger as one batch, uncategorised and tagged as
' imported; the rules run over them afterwards.  Shared with the PDF import.
Public Sub AppendRecords(ByVal records As Collection, ByVal batchId As String)
    Dim lo As ListObject
    Dim txn As clsTxn
    Dim firstRow As Long, count As Long
    Dim i As Long
    Dim ids() As Variant, dates() As Variant, accounts() As Variant
    Dim owners() As Variant, descriptions() As Variant, merchants() As Variant
    Dim amounts() As Variant, categories() As Variant, sources() As Variant
    Dim batches() As Variant, keys() As Variant, tags() As Variant
    Dim nextId As Long

    count = records.Count
    Set lo = modUtil.TxnTable()
    nextId = modLedger.NextTxnNumber(lo)
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
        Set txn = records.Item(i)
        ids(i, 1) = "T" & Format$(nextId + i - 1, "000000")
        dates(i, 1) = CDbl(txn.TxnDate)
        accounts(i, 1) = txn.Account
        owners(i, 1) = txn.Owner
        descriptions(i, 1) = txn.Description
        merchants(i, 1) = txn.Merchant
        amounts(i, 1) = txn.Amount
        categories(i, 1) = CAT_UNCATEGORIZED
        sources(i, 1) = txn.SourceFile
        batches(i, 1) = batchId
        keys(i, 1) = txn.MatchKey()
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

'--- Taking an import back ----------------------------------------------------

' Deletes everything one batch added and marks it in the Import Log.  For a
' statement read against the wrong account, or a PDF whose signs came out
' wrong: one press instead of a hunt through the ledger.
Public Sub UndoImport()
    Dim logTable As ListObject
    Dim options() As String, ids() As String
    Dim logRows() As Long
    Dim i As Long, count As Long, choice As Long
    Dim statusColumn As Long
    Dim batchId As String
    Dim removed As Long

    On Error GoTo Fail
    Set logTable = modUtil.Tbl(SH_LOG, TBL_LOG)
    statusColumn = modUtil.ColumnIndex(logTable, modAccounts.LG_STATUS)

    ' The most recent batches first, leaving out any already undone; the menu
    ' shows at most nine.
    count = 0
    For i = modUtil.BodyRows(logTable) To 1 Step -1
        If IsUndoable(logTable, i, statusColumn) Then count = count + 1
        If count = 9 Then Exit For
    Next i
    If count = 0 Then
        MsgBox "There is no import to undo.", vbInformation, APP_NAME
        Exit Sub
    End If

    ReDim options(0 To count - 1)
    ReDim ids(0 To count - 1)
    ReDim logRows(0 To count - 1)
    count = 0
    For i = modUtil.BodyRows(logTable) To 1 Step -1
        If IsUndoable(logTable, i, statusColumn) Then
            batchId = modUtil.NzStr(logTable.DataBodyRange.Cells(i, _
                          modUtil.ColumnIndex(logTable, modAccounts.LG_BATCH)).Value)
            options(count) = LogLabel(logTable, i)
            ids(count) = batchId
            logRows(count) = i
            count = count + 1
            If count > UBound(options) Then Exit For
        End If
    Next i

    choice = modUtil.AskChoice("Which import should be taken back?" & vbCrLf & _
                               "Every transaction it added is deleted from the ledger.", _
                               options)
    If choice = 0 Then Exit Sub

    If MsgBox("Delete the transactions imported from " & vbCrLf & _
              options(choice - 1) & "?" & vbCrLf & vbCrLf & _
              "Categories you set on them by hand go with them.", _
              vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub

    modUtil.FastMode True
    removed = modLedger.DeleteRowsWhere(modUtil.TxnTable(), COL_BATCH, ids(choice - 1))
    logTable.DataBodyRange.Cells(logRows(choice - 1), statusColumn).Value = _
        "Undone " & Format$(Now, "yyyy-mm-dd hh:nn")
    modUtil.FastMode False

    ' If one side of an automatically detected transfer was in this batch, the
    ' other side must stop being excluded from the reports.  Detection first
    ' clears and rechecks its old automatic pairs.
    modTransfers.DetectTransfers False
    Application.Calculate
    modReport.RefreshAll

    MsgBox removed & " transaction(s) removed. Any transfer pairs affected by " & _
           "the removal were checked again.", vbInformation, APP_NAME
    Exit Sub

Fail:
    modUtil.ReportError "UndoImport"
End Sub

' A log row that names a batch and has not been undone already.
Private Function IsUndoable(ByVal logTable As ListObject, ByVal rowIndex As Long, _
                            ByVal statusColumn As Long) As Boolean
    If Len(modUtil.NzStr(logTable.DataBodyRange.Cells(rowIndex, _
           modUtil.ColumnIndex(logTable, modAccounts.LG_BATCH)).Value)) = 0 Then Exit Function
    If modUtil.NzNum(logTable.DataBodyRange.Cells(rowIndex, _
           modUtil.ColumnIndex(logTable, modAccounts.LG_IMPORTED)).Value) <= 0 Then Exit Function
    IsUndoable = (Len(modUtil.NzStr(logTable.DataBodyRange.Cells(rowIndex, _
                                    statusColumn).Value)) = 0)
End Function

' How a batch is described in the undo menu: its file, count and date.
Private Function LogLabel(ByVal logTable As ListObject, ByVal rowIndex As Long) As String
    Dim whenText As String
    Dim whenValue As Variant
    whenValue = logTable.DataBodyRange.Cells(rowIndex, _
                    modUtil.ColumnIndex(logTable, modAccounts.LG_WHEN)).Value
    If IsDate(whenValue) Then whenText = " on " & Format$(CDate(whenValue), "yyyy-mm-dd")
    LogLabel = modUtil.NzStr(logTable.DataBodyRange.Cells(rowIndex, _
                   modUtil.ColumnIndex(logTable, modAccounts.LG_FILE)).Value, "(no file)") & _
               " - " & modUtil.NzStr(logTable.DataBodyRange.Cells(rowIndex, _
                   modUtil.ColumnIndex(logTable, modAccounts.LG_IMPORTED)).Value, "0") & _
               " added" & whenText
End Function

Public Function FileNameOnly(ByVal path As String) As String
    Dim separatorAt As Long
    separatorAt = InStrRev(path, "\")
    If InStrRev(path, "/") > separatorAt Then separatorAt = InStrRev(path, "/")
    If separatorAt = 0 Then
        FileNameOnly = path
    Else
        FileNameOnly = Mid$(path, separatorAt + 1)
    End If
End Function
