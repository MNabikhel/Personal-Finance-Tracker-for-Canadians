Attribute VB_Name = "modProfiles"
Option Explicit

'== Bank format profiles =====================================================
' Every row of the "Bank Formats" sheet describes how to read one institution's
' CSV export.  Users can edit a profile or add their own without touching code.
'=============================================================================

Public Const PF_NAME As String = "Profile"
Public Const PF_INSTITUTION As String = "Institution"
Public Const PF_SKIP As String = "Skip Rows"
Public Const PF_DELIM As String = "Delimiter"
Public Const PF_DATE_COL As String = "Date Col"
Public Const PF_DATE_FMT As String = "Date Format"
Public Const PF_DESC_COLS As String = "Description Cols"
Public Const PF_AMOUNT_MODE As String = "Amount Mode"
Public Const PF_AMOUNT_COL As String = "Amount Col"
Public Const PF_DEBIT_COL As String = "Debit Col"
Public Const PF_CREDIT_COL As String = "Credit Col"
Public Const PF_SIGNATURE As String = "Header Contains"
Public Const PF_NOTES As String = "Notes"

Public Function ProfilesTable() As ListObject
    Set ProfilesTable = modUtil.Tbl(SH_FORMATS, TBL_FORMATS)
End Function

Public Function ProfileValue(ByVal rowIndex As Long, ByVal header As String) As String
    Dim lo As ListObject
    Set lo = ProfilesTable()
    ProfileValue = modUtil.NzStr(lo.DataBodyRange.Cells(rowIndex, _
                   modUtil.ColumnIndex(lo, header)).Value)
End Function

Public Function ProfileNumber(ByVal rowIndex As Long, ByVal header As String, _
                              Optional ByVal fallback As Long = 0) As Long
    Dim text As String
    text = ProfileValue(rowIndex, header)
    If IsNumeric(text) Then ProfileNumber = CLng(Val(text)) Else ProfileNumber = fallback
End Function

Public Function ProfileNames() As Variant
    Dim lo As ListObject, i As Long, names() As String
    Set lo = ProfilesTable()
    ReDim names(0 To modUtil.BodyRows(lo) - 1)
    For i = 1 To modUtil.BodyRows(lo)
        names(i - 1) = ProfileValue(i, PF_NAME) & " (" & ProfileValue(i, PF_INSTITUTION) & ")"
    Next i
    ProfileNames = names
End Function

Public Function DelimiterOf(ByVal rowIndex As Long) As String
    Select Case UCase$(ProfileValue(rowIndex, PF_DELIM))
        Case "TAB": DelimiterOf = vbTab
        Case "SEMICOLON": DelimiterOf = ";"
        Case "PIPE": DelimiterOf = "|"
        Case Else: DelimiterOf = ","
    End Select
End Function

' Returns the profile row whose signature matches the first few lines of the
' file, or 0 when nothing matches confidently.
Public Function DetectProfile(ByVal rows As Collection) As Long
    Dim lo As ListObject
    Dim probe As String
    Dim i As Long, best As Long, bestScore As Long, score As Long

    Set lo = ProfilesTable()
    probe = ProbeText(rows)
    If Len(probe) = 0 Then Exit Function

    For i = 1 To modUtil.BodyRows(lo)
        score = SignatureScore(ProfileValue(i, PF_SIGNATURE), probe)
        If score > bestScore Then
            bestScore = score
            best = i
        End If
    Next i
    If bestScore > 0 Then DetectProfile = best
End Function

Private Function ProbeText(ByVal rows As Collection) As String
    Dim i As Long, values() As String, joined As String
    For i = 1 To rows.Count
        If i > 4 Then Exit For
        values = rows.Item(i)
        joined = joined & Join(values, ",") & vbLf
    Next i
    ProbeText = UCase$(Replace$(Replace$(joined, """", ""), " ", ""))
End Function

' A signature is a semicolon separated list of fragments that must all appear
' in the file's opening lines.  Longer signatures win ties.
Private Function SignatureScore(ByVal signature As String, ByVal probe As String) As Long
    Dim fragments() As String, i As Long, fragment As String, total As Long
    If Len(Trim$(signature)) = 0 Then Exit Function
    fragments = Split(signature, ";")
    For i = LBound(fragments) To UBound(fragments)
        fragment = UCase$(Replace$(Replace$(Trim$(fragments(i)), """", ""), " ", ""))
        If Len(fragment) = 0 Then
            ' ignore empty fragments
        ElseIf InStr(probe, fragment) = 0 Then
            Exit Function
        Else
            total = total + Len(fragment)
        End If
    Next i
    SignatureScore = total
End Function

Public Function AskForProfile(ByVal fileName As String) As Long
    Dim choice As Long
    choice = modUtil.AskChoice("Which format matches " & fileName & "?", ProfileNames())
    AskForProfile = choice
End Function

' Builds a human readable preview so the user can confirm the mapping before
' anything is written to the ledger.
Public Function PreviewText(ByVal rows As Collection, ByVal profileRow As Long) As String
    Dim i As Long, shown As Long
    Dim values() As String
    Dim out As String
    Dim record As TxnRecord
    Dim ok As Boolean

    out = "Format: " & ProfileValue(profileRow, PF_NAME) & vbCrLf & _
          "Institution: " & ProfileValue(profileRow, PF_INSTITUTION) & vbCrLf & vbCrLf
    For i = ProfileNumber(profileRow, PF_SKIP, 0) + 1 To rows.Count
        values = rows.Item(i)
        record = modImport.BuildRecord(values, profileRow, ok)
        If ok Then
            shown = shown + 1
            out = out & Format$(record.TxnDate, "yyyy-mm-dd") & "   " & _
                  Format$(record.Amount, "#,##0.00;-#,##0.00") & "   " & _
                  Left$(record.Description, 42) & vbCrLf
            If shown >= 4 Then Exit For
        End If
    Next i
    If shown = 0 Then out = out & "(no readable rows found)" & vbCrLf
    PreviewText = out
End Function
