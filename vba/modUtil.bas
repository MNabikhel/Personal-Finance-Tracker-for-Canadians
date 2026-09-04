Attribute VB_Name = "modUtil"
Option Explicit

'== Shared helpers ===========================================================
' Sheet/table lookups, settings access, keyed collections and speed toggles.
' Nothing here uses Windows-only libraries, so the workbook also runs in
' Excel for Mac.
'=============================================================================

Private mCalcMode As XlCalculation
Private mFastDepth As Long

'--- Workbook objects --------------------------------------------------------

Public Function Sh(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set Sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Sh Is Nothing Then
        Err.Raise vbObjectError + 512, "Sh", _
            "Worksheet '" & sheetName & "' is missing. Restore it from a fresh copy " & _
            "of the workbook."
    End If
End Function

Public Function Tbl(ByVal sheetName As String, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set Tbl = Sh(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If Tbl Is Nothing Then
        Err.Raise vbObjectError + 513, "Tbl", _
            "Table '" & tableName & "' is missing from '" & sheetName & "'."
    End If
End Function

Public Function TxnTable() As ListObject
    Set TxnTable = Tbl(SH_TXN, TBL_TXN)
End Function

Public Function ColumnIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(CStr(lc.Name)), header, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    Err.Raise vbObjectError + 514, "ColumnIndex", _
        "Column '" & header & "' was not found in table '" & lo.Name & "'."
End Function

Public Function BodyRows(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        BodyRows = 0
    Else
        BodyRows = lo.DataBodyRange.Rows.Count
    End If
End Function

Public Function CellIn(ByVal lo As ListObject, ByVal rowIndex As Long, _
                       ByVal header As String) As Range
    Set CellIn = lo.DataBodyRange.Cells(rowIndex, ColumnIndex(lo, header))
End Function

'--- Settings ---------------------------------------------------------------

Public Function Setting(ByVal settingName As String) As Variant
    On Error GoTo Fail
    Setting = ThisWorkbook.Names(settingName).RefersToRange.Value
    Exit Function
Fail:
    Setting = Empty
End Function

Public Sub SetSetting(ByVal settingName As String, ByVal value As Variant)
    On Error Resume Next
    ThisWorkbook.Names(settingName).RefersToRange.Value = value
    On Error GoTo 0
End Sub

Public Function IsCoupleMode() As Boolean
    IsCoupleMode = (StrComp(NzStr(Setting(NR_MODE)), MODE_COUPLE, vbTextCompare) = 0)
End Function

Public Function PersonAName() As String
    PersonAName = NzStr(Setting(NR_PERSON_A), "Person A")
End Function

Public Function PersonBName() As String
    PersonBName = NzStr(Setting(NR_PERSON_B), "Person B")
End Function

'--- Conversion helpers ------------------------------------------------------

Public Function NzStr(ByVal value As Variant, Optional ByVal fallback As String = "") As String
    If IsError(value) Then
        NzStr = fallback
    ElseIf IsEmpty(value) Then
        NzStr = fallback
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        NzStr = fallback
    Else
        NzStr = Trim$(CStr(value))
    End If
End Function

Public Function NzNum(ByVal value As Variant, Optional ByVal fallback As Double = 0) As Double
    If IsError(value) Then
        NzNum = fallback
    ElseIf VarType(value) = vbDate Then
        ' A date-formatted cell reads back as a Date, and IsNumeric says False
        ' to those; the day serial is the number wanted.
        NzNum = CDbl(value)
    ElseIf IsNumeric(value) Then
        NzNum = CDbl(value)
    Else
        NzNum = fallback
    End If
End Function

Public Function MonthKey(ByVal dateValue As Date) As String
    MonthKey = Format$(dateValue, "yyyy-mm")
End Function

'--- Keyed collections (a portable stand-in for Scripting.Dictionary) --------

Public Function HasKey(ByVal source As Collection, ByVal key As String) As Boolean
    Dim ignored As Variant
    On Error Resume Next
    ignored = source.Item(key)
    HasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function GetVal(ByVal source As Collection, ByVal key As String, _
                       Optional ByVal fallback As Variant = 0) As Variant
    On Error Resume Next
    GetVal = source.Item(key)
    If Err.Number <> 0 Then
        Err.Clear
        GetVal = fallback
    End If
    On Error GoTo 0
End Function

Public Sub PutVal(ByVal target As Collection, ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    target.Remove key
    Err.Clear
    On Error GoTo 0
    target.Add value, key
End Sub

Public Sub BumpVal(ByVal target As Collection, ByVal key As String)
    PutVal target, key, GetVal(target, key, 0) + 1
End Sub

'--- Hashing ----------------------------------------------------------------

' FNV-1a, used to build a stable duplicate-detection key for each transaction.
Public Function HashText(ByVal text As String) As String
    Dim hashHigh As Long, hashLow As Long
    Dim i As Long, ch As Long
    Dim productLow As Long, productHigh As Long

    hashHigh = &H811C&
    hashLow = &H9DC5&

    For i = 1 To Len(text)
        ch = AscW(Mid$(text, i, 1)) And &HFFFF&
        hashLow = hashLow Xor ch
        ' Multiply the 32-bit value by the FNV prime 16777619 using 16-bit halves.
        productLow = (hashLow And &HFFFF&) * &H193&
        productHigh = (hashHigh And &HFFFF&) * &H193& + ((hashLow And &HFFFF&) * &H100&)
        hashLow = productLow And &HFFFF&
        hashHigh = (productHigh + (productLow \ &H10000)) And &HFFFF&
    Next i

    HashText = Right$("0000" & Hex$(hashHigh), 4) & Right$("0000" & Hex$(hashLow), 4)
End Function

Public Function MatchKey(ByVal accountName As String, ByVal dateValue As Date, _
                         ByVal amount As Double, ByVal description As String) As String
    MatchKey = HashText(UCase$(accountName) & "|" & Format$(dateValue, "yyyy-mm-dd") & _
                        "|" & Format$(amount, "0.00") & "|" & UCase$(CondenseSpaces(description)))
End Function

'--- Text -------------------------------------------------------------------

Public Function CondenseSpaces(ByVal text As String) As String
    Dim result As String
    result = Replace$(Replace$(Replace$(text, vbTab, " "), vbCr, " "), vbLf, " ")
    result = Replace$(result, Chr$(160), " ")
    Do While InStr(result, "  ") > 0
        result = Replace$(result, "  ", " ")
    Loop
    CondenseSpaces = Trim$(result)
End Function

Public Function TitleCaseWords(ByVal text As String) As String
    Dim parts() As String, i As Long, word As String
    parts = Split(LCase$(CondenseSpaces(text)), " ")
    For i = LBound(parts) To UBound(parts)
        word = parts(i)
        If Len(word) > 0 Then
            If IsAcronym(word) Then
                parts(i) = UCase$(word)
            Else
                parts(i) = UCase$(Left$(word, 1)) & Mid$(word, 2)
            End If
        End If
    Next i
    TitleCaseWords = Join(parts, " ")
End Function

Private Function IsAcronym(ByVal word As String) As Boolean
    Select Case UCase$(word)
        Case "ATM", "POS", "GST", "HST", "CRA", "EI", "CPP", "OAS", "TFSA", "RRSP", _
             "FHSA", "RESP", "LCBO", "SAQ", "BMO", "RBC", "TD", "CIBC", "PC", "GO", _
             "TTC", "STM", "AMEX", "VISA", "MC", "ON", "BC", "AB", "QC", "NS", "NB", _
             "MB", "SK", "PE", "NL", "YT", "NT", "NU"
            IsAcronym = True
    End Select
End Function

'--- Speed / status ---------------------------------------------------------

Public Sub FastMode(ByVal switchOn As Boolean)
    If switchOn Then
        If mFastDepth = 0 Then
            mCalcMode = Application.Calculation
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
        End If
        mFastDepth = mFastDepth + 1
    Else
        ' A failure can happen before its caller reached FastMode True.
        ' Restoring an uninitialised mCalcMode (zero) then raises a second
        ' Excel error and hides the useful first one.
        If mFastDepth = 0 Then Exit Sub
        mFastDepth = mFastDepth - 1
        If mFastDepth = 0 Then
            Application.Calculation = mCalcMode
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Application.StatusBar = False
        End If
    End If
End Sub

Public Sub Status(ByVal message As String)
    On Error Resume Next
    If Len(message) = 0 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = APP_NAME & ": " & message
    End If
    On Error GoTo 0
End Sub

Public Sub ReportError(ByVal source As String)
    Dim number As Long
    Dim description As String
    number = Err.Number
    description = Err.Description
    FastMode False
    MsgBox "Something went wrong in " & source & "." & vbCrLf & vbCrLf & _
           description & " (error " & number & ")", _
           vbExclamation, APP_NAME
End Sub

Public Function AskChoice(ByVal prompt As String, ByVal options As Variant, _
                          Optional ByVal title As String = "") As Long
    ' Returns a 1-based index into options, or 0 when the user cancels.
    Dim menu As String, i As Long, answer As String
    If Len(title) = 0 Then title = APP_NAME
    For i = LBound(options) To UBound(options)
        menu = menu & (i - LBound(options) + 1) & ") " & options(i) & vbCrLf
    Next i
    Do
        answer = InputBox(prompt & vbCrLf & vbCrLf & menu, title, "1")
        If Len(answer) = 0 Then Exit Function
        If IsNumeric(answer) Then
            i = CLng(answer)
            If i >= 1 And i <= (UBound(options) - LBound(options) + 1) Then
                AskChoice = i
                Exit Function
            End If
        End If
        MsgBox "Enter a number between 1 and " & _
               (UBound(options) - LBound(options) + 1) & ".", vbInformation, title
    Loop
End Function
