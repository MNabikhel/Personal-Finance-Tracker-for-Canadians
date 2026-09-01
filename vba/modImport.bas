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
        If modPdf.IsPdf(CStr(chosen(i))) Then
            line = modPdf.ImportOnePdf(CStr(chosen(i)), "B" & batchStamp & "-" & i, _
                                       totalImported, totalSkipped)
        Else
            line = ImportOneFile(CStr(chosen(i)), "B" & batchStamp & "-" & i, _
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
