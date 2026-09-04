Attribute VB_Name = "modPdf"
Option Explicit

'== PDF statements ===========================================================
' A PDF is a picture of a page, so before anything can be imported the text
' has to be got back out of it.  Excel has one reader for that - Power Query's
' PDF connector, in Excel for Windows under Microsoft 365 - and it is driven
' here from VBA: a query is added, loaded onto a scratch sheet, read, and both
' are removed again.  Where that reader is missing, Word is asked to convert
' the file instead, which any Word since 2013 can do.  Either way the result is
' a list of lines, and modPdfText turns those into transactions.
'
' Nothing in here is reachable from the test harness: neither LibreOffice nor
' this repository can run Power Query or Word.  It is kept to plumbing for
' that reason, with the parsing kept out of it.
'=============================================================================

' What the Import Log and the Accounts sheet call statements read this way.
Public Const PDF_FORMAT_NAME As String = "PDF statement"

' The query and scratch objects are named so they can always be found and
' removed, even after a failure part way through.
Private Const PQ_QUERY_NAME As String = "cftPdfStatement"

Public Function IsPdf(ByVal path As String) As Boolean
    IsPdf = (LCase$(Right$(path, 4)) = ".pdf")
End Function

' Imports one PDF statement.  Mirrors modImport.ImportOneFile: returns a line
' for the summary, or "" when nothing was imported.
Public Function ImportOnePdf(ByVal path As String, ByVal batchId As String, _
                             ByRef totalImported As Long, _
                             ByRef totalSkipped As Long) As String
    Dim fileName As String, how As String, kind As String
    Dim lines As Collection, records As Collection
    Dim anchorYear As Long, anchorMonth As Long
    Dim readCount As Long, badCount As Long, dupeCount As Long
    Dim answer As VbMsgBoxResult
    Dim yearText As String
    Dim profile As clsProfile
    Dim accountName As String
    Dim triedBoth As Boolean
    Dim i As Long
    Dim txn As clsTxn

    On Error GoTo Fail

    fileName = modImport.FileNameOnly(path)
    modUtil.Status "reading " & fileName & " ..."
    Set lines = ExtractLines(path, how)
    modUtil.Status ""

    If lines Is Nothing Then
        MsgBox "Excel could not read " & fileName & "." & vbCrLf & vbCrLf & _
               "Reading a PDF needs Excel for Windows with Power Query's PDF " & _
               "reader (Microsoft 365), or Word installed to convert the file. " & _
               "Neither worked here." & vbCrLf & vbCrLf & _
               "Download the statement as CSV instead - every bank offers it under " & _
               "the same download button - and import that.", vbExclamation, APP_NAME
        Exit Function
    End If
    If lines.Count = 0 Then
        MsgBox fileName & " has no text in it." & vbCrLf & vbCrLf & _
               "A scanned statement is a picture of the page. Download the statement " & _
               "itself from the bank, as PDF or CSV.", vbExclamation, APP_NAME
        Exit Function
    End If

    anchorYear = modPdfText.StatementAnchor(lines, anchorMonth)
    If anchorYear = 0 Then
        yearText = Trim$(InputBox("Which year is " & fileName & " for?" & vbCrLf & vbCrLf & _
                                  "The statement does not print one next to its dates.", _
                                  APP_NAME, CStr(Year(Date))))
        If Not IsNumeric(yearText) Then Exit Function
        anchorYear = CLng(Val(yearText))
    End If

    kind = modPdfText.DetectKind(lines, anchorYear, anchorMonth)
    Do
        Set records = modPdfText.ReadStatement(lines, kind, anchorYear, anchorMonth, _
                                               readCount, badCount)
        If records.Count = 0 Then
            If triedBoth Then
                If answer = vbNo Then
                    MsgBox "Read as a " & LCase$(kind) & " statement, " & fileName & _
                           " has no lines that look like transactions, so it was " & _
                           "skipped. Import it again and accept the first reading, or " & _
                           "use the CSV version of the statement.", vbExclamation, APP_NAME
                Else
                    MsgBox "No transactions could be read from " & fileName & "." & _
                           vbCrLf & vbCrLf & "The lines on the page do not start with " & _
                           "a date and end with an amount. Import the CSV version of " & _
                           "this statement instead.", vbExclamation, APP_NAME
                End If
                Exit Function
            End If
            triedBoth = True
            kind = OtherKind(kind)
        Else
            answer = MsgBox(fileName & ", read with " & how & " as a " & LCase$(kind) & _
                            " statement. Money out is negative:" & vbCrLf & vbCrLf & _
                            modPdfText.PreviewText(records, 6) & vbCrLf & _
                            records.Count & " transaction(s) found." & vbCrLf & vbCrLf & _
                            "Yes = import these" & vbCrLf & _
                            "No = read it as a " & LCase$(OtherKind(kind)) & _
                            " statement instead" & vbCrLf & _
                            "Cancel = skip this file", vbYesNoCancel + vbQuestion, APP_NAME)
            If answer = vbCancel Then Exit Function
            If answer = vbYes Then Exit Do
            triedBoth = True
            kind = OtherKind(kind)
        End If
    Loop

    ' The account resolver wants a bank format to go by; a PDF has none, so a
    ' stand-in carries the name the Accounts sheet can match on.
    Set profile = New clsProfile
    profile.Name = PDF_FORMAT_NAME
    accountName = modAccounts.ResolveAccount(fileName, profile)
    If Len(accountName) = 0 Then Exit Function

    modUtil.FastMode True
    For i = 1 To records.Count
        Set txn = records.Item(i)
        txn.Account = accountName
        txn.Owner = modAccounts.OwnerOfAccount(accountName)
        txn.SourceFile = fileName
    Next i
    Set records = modImport.WithoutDuplicates(records, modImport.ExistingKeyCounts(), _
                                              dupeCount)
    If records.Count > 0 Then modImport.AppendRecords records, batchId
    modUtil.FastMode False

    totalImported = totalImported + records.Count
    totalSkipped = totalSkipped + dupeCount + badCount

    modAccounts.LogBatch batchId, fileName, PDF_FORMAT_NAME & ": " & kind & " (" & how & ")", _
                         accountName, readCount, records.Count, dupeCount, badCount

    ImportOnePdf = fileName & ": " & records.Count & " added, " & dupeCount & _
                   " duplicate(s), " & badCount & " unreadable line(s)."
    Exit Function

Fail:
    modUtil.ReportError "ImportOnePdf"
End Function

Private Function OtherKind(ByVal kind As String) As String
    If StrComp(kind, modPdfText.KIND_CARD, vbTextCompare) = 0 Then
        OtherKind = modPdfText.KIND_ACCOUNT
    Else
        OtherKind = modPdfText.KIND_CARD
    End If
End Function

'--- Getting the text out ---------------------------------------------------

' The lines of the PDF, or Nothing when nothing on this computer can read it.
' how says which reader managed it.
Private Function ExtractLines(ByVal path As String, ByRef how As String) As Collection
    Dim lines As Collection
    Dim fallback As Collection

    modUtil.FastMode True
    Set lines = LinesViaPowerQuery(path)
    If lines Is Nothing Then
        Set lines = LinesViaWord(path)
        If Not lines Is Nothing Then how = "Word"
    ElseIf lines.Count = 0 Then
        ' The connector can load a PDF successfully yet recover no page rows.
        ' Word's converter is independent and sometimes still gets the text.
        Set fallback = LinesViaWord(path)
        If Not fallback Is Nothing Then
            If fallback.Count > 0 Then
                Set lines = fallback
                how = "Word"
            End If
        End If
    Else
        how = "Excel's PDF reader"
    End If
    If Len(how) = 0 And Not lines Is Nothing Then how = "Excel's PDF reader"
    modUtil.FastMode False
    Set ExtractLines = lines
End Function

' Power Query's Pdf.Tables, loaded onto a scratch sheet and read back.  The
' workbook is addressed as Object where Queries is involved so that the
' project still compiles in an Excel too old to have them.
Private Function LinesViaPowerQuery(ByVal path As String) As Collection
    Dim host As Object
    Dim previous As Object
    Dim scratch As Worksheet
    Dim lo As ListObject
    Dim values As Variant
    Dim out As Collection
    Dim i As Long
    Dim oldAlerts As Boolean

    On Error GoTo Fail
    Set host = ThisWorkbook
    Set previous = ActiveSheet
    oldAlerts = Application.DisplayAlerts
    DropQuery
    host.Queries.Add PQ_QUERY_NAME, PdfQueryFormula(path)

    Set scratch = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Set lo = scratch.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
                "Location=" & PQ_QUERY_NAME & ";Extended Properties=""""", _
        Destination:=scratch.Range("A1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & PQ_QUERY_NAME & "]")
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With

    Set out = New Collection
    If Not lo.DataBodyRange Is Nothing Then
        values = lo.DataBodyRange.Value
        If IsArray(values) Then
            For i = 1 To UBound(values, 1)
                AddLine out, values(i, 1)
            Next i
        Else
            AddLine out, values
        End If
    End If
    Set LinesViaPowerQuery = out

Cleanup:
    On Error Resume Next
    If Not scratch Is Nothing Then
        Application.DisplayAlerts = False
        scratch.Delete
        Application.DisplayAlerts = oldAlerts
    End If
    DropQuery
    If Not previous Is Nothing Then previous.Activate
    Exit Function

Fail:
    Set LinesViaPowerQuery = Nothing
    Resume Cleanup
End Function

' The M code: every page of the PDF as lines of text, its cells joined with
' two spaces.  Older readers return only the tables they found, without the
' "Page" entries; those are read whole instead.
Private Function PdfQueryFormula(ByVal path As String) As String
    Dim q As String
    q = """"
    PdfQueryFormula = _
        "let" & vbLf & _
        "    Source = Pdf.Tables(File.Contents(" & q & Replace$(path, q, q & q) & q & "))," & vbLf & _
        "    Pages = Table.SelectRows(Source, each [Kind] = " & q & "Page" & q & ")," & vbLf & _
        "    Chosen = if Table.IsEmpty(Pages) then Source else Pages," & vbLf & _
        "    PageLines = List.Transform(Chosen[Data], (page) => List.Transform(" & _
        "Table.ToRows(page), (row) => Text.Combine(List.Transform(row, " & _
        "each if _ = null then " & q & q & " else Text.From(_)), " & q & "  " & q & ")))," & vbLf & _
        "    Lines = List.Combine(PageLines)," & vbLf & _
        "    Result = Table.FromList(Lines, Splitter.SplitByNothing(), {" & q & "Line" & q & "}, " & _
        "null, ExtraValues.Ignore)" & vbLf & _
        "in" & vbLf & _
        "    Result"
End Function

Private Sub DropQuery()
    Dim host As Object
    Dim i As Long
    On Error Resume Next
    Set host = ThisWorkbook
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        If InStr(1, ThisWorkbook.Connections(i).Name, PQ_QUERY_NAME, vbTextCompare) > 0 Then
            ThisWorkbook.Connections(i).Delete
        End If
    Next i
    host.Queries(PQ_QUERY_NAME).Delete
    On Error GoTo 0
End Sub

' Word opens a PDF by converting it to a document.  Its text comes back with
' each table cell closed by CR+BEL and each row by a second CR+BEL, which is
' enough to put a row's cells back on one line.
Private Function LinesViaWord(ByVal path As String) As Collection
    Dim wordApp As Object
    Dim doc As Object
    Dim text As String

    On Error GoTo Fail
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    Set doc = wordApp.Documents.Open(path, False, True, False)
    text = doc.Content.Text
    doc.Close 0
    Set doc = Nothing
    wordApp.Quit
    Set wordApp = Nothing

    text = Replace$(text, vbCr & Chr$(7) & vbCr & Chr$(7), vbLf)
    text = Replace$(text, vbCr & Chr$(7), "  ")
    Set LinesViaWord = modPdfText.SplitLines(text)
    Exit Function

Fail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close 0
    If Not wordApp Is Nothing Then wordApp.Quit
    Set LinesViaWord = Nothing
End Function

Private Sub AddLine(ByVal target As Collection, ByVal value As Variant)
    Dim line As String
    line = modUtil.CondenseSpaces(modUtil.NzStr(value))
    If Len(line) > 0 Then target.Add line
End Sub
