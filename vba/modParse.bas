Attribute VB_Name = "modParse"
Option Explicit

'== File reading and field parsing ===========================================
' Bank exports arrive in a surprising number of shapes: UTF-8 with a BOM,
' Windows-1252, quoted fields containing commas or line breaks, four different
' date orders and French-Canadian number formatting.  Everything in this module
' is locale independent on purpose - CDate and CDbl are never used on file data
' because they follow the machine's regional settings.
'=============================================================================

'--- Whole-file reading ------------------------------------------------------

Public Function ReadTextFile(ByVal path As String) As String
    Dim handle As Integer
    Dim raw() As Byte
    Dim length As Long

    length = FileLen(path)
    If length = 0 Then Exit Function

    ReDim raw(0 To length - 1)
    handle = FreeFile
    Open path For Binary Access Read As #handle
    Get #handle, , raw
    Close #handle

    ReadTextFile = DecodeBytes(raw)
End Function

Private Function DecodeBytes(ByRef raw() As Byte) As String
    Dim upper As Long
    upper = UBound(raw)

    If upper >= 2 Then
        If raw(0) = &HEF And raw(1) = &HBB And raw(2) = &HBF Then
            DecodeBytes = Utf8ToText(raw, 3)
            Exit Function
        End If
    End If
    If upper >= 1 Then
        If raw(0) = &HFF And raw(1) = &HFE Then
            DecodeBytes = Utf16LeToText(raw, 2)
            Exit Function
        End If
    End If

    If LooksLikeUtf8(raw) Then
        DecodeBytes = Utf8ToText(raw, 0)
    Else
        DecodeBytes = AnsiToText(raw)
    End If
End Function

Private Function LooksLikeUtf8(ByRef raw() As Byte) As Boolean
    ' True when at least one valid multi-byte sequence exists and none is invalid.
    Dim i As Long, upper As Long, extra As Long, j As Long
    Dim sawMultiByte As Boolean

    upper = UBound(raw)
    i = 0
    Do While i <= upper
        If raw(i) < &H80 Then
            extra = 0
        ElseIf (raw(i) And &HE0) = &HC0 Then
            extra = 1
        ElseIf (raw(i) And &HF0) = &HE0 Then
            extra = 2
        ElseIf (raw(i) And &HF8) = &HF0 Then
            extra = 3
        Else
            Exit Function
        End If
        If extra > 0 Then
            If i + extra > upper Then Exit Function
            For j = 1 To extra
                If (raw(i + j) And &HC0) <> &H80 Then Exit Function
            Next j
            sawMultiByte = True
        End If
        i = i + extra + 1
    Loop
    LooksLikeUtf8 = sawMultiByte
End Function

Private Function Utf8ToText(ByRef raw() As Byte, ByVal startAt As Long) As String
    Dim i As Long, upper As Long
    Dim code As Long, extra As Long, j As Long
    Dim out As String, buffer As String, bufferLen As Long

    upper = UBound(raw)
    buffer = Space$(8192)
    bufferLen = 0
    i = startAt

    Do While i <= upper
        If raw(i) < &H80 Then
            code = raw(i)
            extra = 0
        ElseIf (raw(i) And &HE0) = &HC0 Then
            code = raw(i) And &H1F
            extra = 1
        ElseIf (raw(i) And &HF0) = &HE0 Then
            code = raw(i) And &HF
            extra = 2
        ElseIf (raw(i) And &HF8) = &HF0 Then
            code = raw(i) And &H7
            extra = 3
        Else
            code = &HFFFD
            extra = 0
        End If

        For j = 1 To extra
            If i + j > upper Then Exit For
            code = (code * &H40&) + (raw(i + j) And &H3F)
        Next j
        i = i + extra + 1

        If code > &HFFFF& Then
            ' Encode astral characters as a UTF-16 surrogate pair.
            code = code - &H10000
            bufferLen = bufferLen + 1
            Mid$(buffer, bufferLen, 1) = ChrW$(&HD800& + (code \ &H400&))
            bufferLen = bufferLen + 1
            Mid$(buffer, bufferLen, 1) = ChrW$(&HDC00& + (code And &H3FF&))
        Else
            bufferLen = bufferLen + 1
            Mid$(buffer, bufferLen, 1) = ChrW$(code)
        End If

        If bufferLen > Len(buffer) - 4 Then
            out = out & Left$(buffer, bufferLen)
            bufferLen = 0
        End If
    Loop

    Utf8ToText = out & Left$(buffer, bufferLen)
End Function

Private Function Utf16LeToText(ByRef raw() As Byte, ByVal startAt As Long) As String
    Dim i As Long, out As String
    For i = startAt To UBound(raw) - 1 Step 2
        out = out & ChrW$(raw(i) + raw(i + 1) * 256&)
    Next i
    Utf16LeToText = out
End Function

Private Function AnsiToText(ByRef raw() As Byte) As String
    ' Windows-1252 (and Latin-1) map onto Unicode with only a handful of
    ' exceptions in the 0x80-0x9F range, handled explicitly below.
    Dim i As Long, code As Long, out As String, buffer As String, bufferLen As Long
    buffer = Space$(8192)
    For i = LBound(raw) To UBound(raw)
        code = raw(i)
        If code >= &H80 And code <= &H9F Then code = Cp1252High(code)
        bufferLen = bufferLen + 1
        Mid$(buffer, bufferLen, 1) = ChrW$(code)
        If bufferLen = Len(buffer) Then
            out = out & buffer
            bufferLen = 0
        End If
    Next i
    AnsiToText = out & Left$(buffer, bufferLen)
End Function

Private Function Cp1252High(ByVal code As Long) As Long
    Select Case code
        Case &H80: Cp1252High = &H20AC
        Case &H82: Cp1252High = &H201A
        Case &H83: Cp1252High = &H192
        Case &H84: Cp1252High = &H201E
        Case &H85: Cp1252High = &H2026
        Case &H86: Cp1252High = &H2020
        Case &H87: Cp1252High = &H2021
        Case &H88: Cp1252High = &H2C6
        Case &H89: Cp1252High = &H2030
        Case &H8A: Cp1252High = &H160
        Case &H8B: Cp1252High = &H2039
        Case &H8C: Cp1252High = &H152
        Case &H8E: Cp1252High = &H17D
        Case &H91: Cp1252High = &H2018
        Case &H92: Cp1252High = &H2019
        Case &H93: Cp1252High = &H201C
        Case &H94: Cp1252High = &H201D
        Case &H95: Cp1252High = &H2022
        Case &H96: Cp1252High = &H2013
        Case &H97: Cp1252High = &H2014
        Case &H98: Cp1252High = &H2DC
        Case &H99: Cp1252High = &H2122
        Case &H9A: Cp1252High = &H161
        Case &H9B: Cp1252High = &H203A
        Case &H9C: Cp1252High = &H153
        Case &H9E: Cp1252High = &H17E
        Case &H9F: Cp1252High = &H178
        Case Else: Cp1252High = code
    End Select
End Function

'--- Delimited text parsing -------------------------------------------------

Public Function SplitRows(ByVal text As String, ByVal delimiter As String) As Collection
    ' Returns a Collection of records; each record is itself a Collection of
    ' field strings, indexed from 1.  Quoted fields may contain the delimiter
    ' and line breaks.
    Dim rows As Collection
    Dim fields As Collection
    Dim i As Long, length As Long
    Dim ch As String, field As String
    Dim inQuotes As Boolean
    Dim sep As String

    Set rows = New Collection
    Set fields = New Collection
    sep = Left$(delimiter & ",", 1)
    length = Len(text)

    For i = 1 To length
        ch = Mid$(text, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < length And Mid$(text, i + 1, 1) = """" Then
                    field = field & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                field = field & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuotes = True
                Case sep
                    fields.Add field
                    field = ""
                Case vbCr
                    ' Handled with the following vbLf, or on its own for old Macs.
                    If i = length Or Mid$(text, i + 1, 1) <> vbLf Then
                        FlushRow rows, fields, field
                        Set fields = New Collection
                        field = ""
                    End If
                Case vbLf
                    FlushRow rows, fields, field
                    Set fields = New Collection
                    field = ""
                Case Else
                    field = field & ch
            End Select
        End If
    Next i

    If Len(field) > 0 Or fields.Count > 0 Then FlushRow rows, fields, field
    Set SplitRows = rows
End Function

Private Sub FlushRow(ByVal rows As Collection, ByVal fields As Collection, _
                     ByVal lastField As String)
    Dim trimmed As Collection
    Dim i As Long
    Dim value As String
    Dim isBlank As Boolean

    fields.Add lastField
    Set trimmed = New Collection
    isBlank = True
    For i = 1 To fields.Count
        value = Trim$(CStr(fields.Item(i)))
        If Len(value) > 0 Then isBlank = False
        trimmed.Add value
    Next i
    If Not isBlank Then rows.Add trimmed
End Sub

Public Function FieldAt(ByVal fields As Collection, ByVal position As Long) As String
    ' 1-based column position; blank when the column does not exist.
    If position < 1 Or position > fields.Count Then Exit Function
    FieldAt = fields.Item(position)
End Function

Public Function FieldsAt(ByVal fields As Collection, ByVal positions As String) As String
    ' positions is a comma separated list such as "5,6"; the pieces are joined
    ' with a single space so multi-part descriptions read naturally.
    Dim parts() As String, i As Long, piece As String, out As String
    parts = Split(positions, ",")
    For i = LBound(parts) To UBound(parts)
        If IsNumeric(Trim$(parts(i))) Then
            piece = FieldAt(fields, CLng(Trim$(parts(i))))
            If Len(piece) > 0 Then
                If Len(out) > 0 Then out = out & " "
                out = out & piece
            End If
        End If
    Next i
    FieldsAt = modUtil.CondenseSpaces(out)
End Function

Public Function JoinFields(ByVal fields As Collection, ByVal separator As String) As String
    Dim i As Long, out As String
    For i = 1 To fields.Count
        If i > 1 Then out = out & separator
        out = out & fields.Item(i)
    Next i
    JoinFields = out
End Function

'--- Numbers ----------------------------------------------------------------

Public Function ParseAmount(ByVal text As String, ByRef ok As Boolean) As Double
    Dim cleaned As String, i As Long, ch As String
    Dim negative As Boolean
    Dim lastComma As Long, lastDot As Long

    ok = False
    cleaned = Trim$(Replace$(Replace$(text, Chr$(160), " "), " ", ""))
    If Len(cleaned) = 0 Then Exit Function

    If Left$(cleaned, 1) = "(" And Right$(cleaned, 1) = ")" Then
        negative = True
        cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
    End If
    If Right$(cleaned, 1) = "-" Then
        negative = True
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    End If

    ' Keep digits, separators and a leading sign only.
    Dim kept As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "#" Or ch = "." Or ch = "," Then
            kept = kept & ch
        ElseIf ch = "-" And Len(kept) = 0 Then
            negative = Not negative
        End If
        ' Anything else - "$", "CAD", stray letters - is currency noise.
    Next i
    If Len(kept) = 0 Then Exit Function

    lastComma = InStrRev(kept, ",")
    lastDot = InStrRev(kept, ".")

    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            ' 1.234,56 - European/French-Canadian layout.
            kept = Replace$(kept, ".", "")
            kept = Replace$(kept, ",", ".")
        Else
            kept = Replace$(kept, ",", "")
        End If
    ElseIf lastComma > 0 Then
        If Len(kept) - lastComma = 2 Then
            kept = Replace$(kept, ",", ".")   ' 1234,56
        Else
            kept = Replace$(kept, ",", "")    ' 1,234
        End If
    End If

    If Not IsNumeric(kept) Then Exit Function
    ParseAmount = Val(kept)                     ' Val is always dot-decimal.
    If negative Then ParseAmount = -ParseAmount
    ok = True
End Function

'--- Dates ------------------------------------------------------------------

Public Function ParseDate(ByVal text As String, ByVal pattern As String, _
                          ByRef ok As Boolean) As Date
    Dim cleaned As String
    Dim tokens As Variant
    Dim y As Long, m As Long, d As Long
    Dim order As String

    ok = False
    cleaned = Trim$(text)
    If Len(cleaned) = 0 Then Exit Function

    order = UCase$(Trim$(pattern))
    If Len(order) = 0 Then order = "AUTO"

    ' A trailing time stamp is harmless: only the first three tokens are read.
    If order = "YYYYMMDD" Then
        cleaned = OnlyDigits(cleaned)
        If Len(cleaned) < 8 Then Exit Function
        cleaned = Left$(cleaned, 8)
        y = CLng(Left$(cleaned, 4))
        m = CLng(Mid$(cleaned, 5, 2))
        d = CLng(Right$(cleaned, 2))
    Else
        tokens = SplitDateTokens(cleaned)
        If UBound(tokens) < 2 Then Exit Function
        Select Case order
            Case "YYYY-MM-DD", "YYYY/MM/DD"
                y = ToYear(tokens(0)): m = ToMonth(tokens(1)): d = ToNum(tokens(2))
            Case "DD/MM/YYYY", "DD-MM-YYYY"
                d = ToNum(tokens(0)): m = ToMonth(tokens(1)): y = ToYear(tokens(2))
            Case "MM/DD/YYYY", "MM-DD-YYYY"
                m = ToMonth(tokens(0)): d = ToNum(tokens(1)): y = ToYear(tokens(2))
            Case "DD-MMM-YYYY", "DD MMM YYYY"
                d = ToNum(tokens(0)): m = ToMonth(tokens(1)): y = ToYear(tokens(2))
            Case "MMM-DD-YYYY", "MMM DD YYYY"
                m = ToMonth(tokens(0)): d = ToNum(tokens(1)): y = ToYear(tokens(2))
            Case Else
                ' AUTO: infer from the token shapes.
                If Len(tokens(0)) = 4 Then
                    y = ToYear(tokens(0)): m = ToMonth(tokens(1)): d = ToNum(tokens(2))
                ElseIf Not IsNumeric(tokens(0)) Then
                    m = ToMonth(tokens(0)): d = ToNum(tokens(1)): y = ToYear(tokens(2))
                ElseIf Not IsNumeric(tokens(1)) Then
                    d = ToNum(tokens(0)): m = ToMonth(tokens(1)): y = ToYear(tokens(2))
                ElseIf ToNum(tokens(0)) > 12 Then
                    d = ToNum(tokens(0)): m = ToMonth(tokens(1)): y = ToYear(tokens(2))
                Else
                    m = ToMonth(tokens(0)): d = ToNum(tokens(1)): y = ToYear(tokens(2))
                End If
        End Select
    End If

    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or y < 1900 Or y > 2200 Then Exit Function

    On Error GoTo Fail
    Dim result As Date
    result = DateSerial(y, m, d)
    ' DateSerial quietly rolls February 30th forward into March; a date that
    ' does not survive the round trip was never a date.
    If Year(result) <> y Or Month(result) <> m Or Day(result) <> d Then Exit Function
    ParseDate = result
    ok = True
    Exit Function
Fail:
    ok = False
End Function

' Returns a Variant rather than String(): an array-typed return value is not
' portable, and every caller treats the result as a plain array anyway.
Private Function SplitDateTokens(ByVal text As String) As Variant
    Dim scrubbed As String, i As Long, ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "[0-9A-Za-z]" Then
            scrubbed = scrubbed & ch
        Else
            scrubbed = scrubbed & "|"
        End If
    Next i
    Do While InStr(scrubbed, "||") > 0
        scrubbed = Replace$(scrubbed, "||", "|")
    Loop
    SplitDateTokens = Split(Trim2(scrubbed, "|"), "|")
End Function

Private Function Trim2(ByVal text As String, ByVal ch As String) As String
    Do While Left$(text, 1) = ch And Len(text) > 0
        text = Mid$(text, 2)
    Loop
    Do While Right$(text, 1) = ch And Len(text) > 0
        text = Left$(text, Len(text) - 1)
    Loop
    Trim2 = text
End Function

Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "#" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function

Private Function ToNum(ByVal token As String) As Long
    If IsNumeric(token) Then ToNum = CLng(Val(token))
End Function

Private Function ToYear(ByVal token As String) As Long
    Dim value As Long
    value = ToNum(token)
    If value < 100 Then
        If value < 70 Then value = 2000 + value Else value = 1900 + value
    End If
    ToYear = value
End Function

Private Function ToMonth(ByVal token As String) As Long
    Dim names As Variant, i As Long
    If IsNumeric(token) Then
        ToMonth = CLng(Val(token))
        Exit Function
    End If
    names = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", _
                  "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    For i = 0 To 11
        If UCase$(Left$(token, 3)) = names(i) Then
            ToMonth = i + 1
            Exit Function
        End If
    Next i
    ' French month abbreviations used by Quebec institutions.
    Select Case UCase$(Left$(token, 4))
        Case "JANV": ToMonth = 1
        Case "FEVR", "FÉVR": ToMonth = 2
        Case "MARS": ToMonth = 3
        Case "AVRI": ToMonth = 4
        Case "JUIN": ToMonth = 6
        Case "JUIL": ToMonth = 7
        Case "AOUT", "AOÛT": ToMonth = 8
        Case "SEPT": ToMonth = 9
        Case "OCTO": ToMonth = 10
        Case "NOVE": ToMonth = 11
        Case "DECE", "DÉCE": ToMonth = 12
        Case Else
            If UCase$(Left$(token, 3)) = "MAI" Then ToMonth = 5
    End Select
End Function
