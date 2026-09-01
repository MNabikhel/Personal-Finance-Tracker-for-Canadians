Attribute VB_Name = "modProfiles"
Option Explicit

'== Bank format profiles =====================================================
' Every row of the "Bank Formats" sheet describes how to read one institution's
' CSV export.  Users can edit a profile or add their own without touching code.
' This module is the only place that reads that sheet; everything else works
' with the clsProfile objects it hands out.
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

Public Function Profile(ByVal rowIndex As Long) As clsProfile
    Dim out As clsProfile
    Set out = New clsProfile
    out.RowIndex = rowIndex
    out.Name = ProfileValue(rowIndex, PF_NAME)
    out.Institution = ProfileValue(rowIndex, PF_INSTITUTION)
    out.SkipRows = ProfileNumber(rowIndex, PF_SKIP, 0)
    out.DelimiterName = ProfileValue(rowIndex, PF_DELIM)
    out.DateColumn = ProfileNumber(rowIndex, PF_DATE_COL, 1)
    out.DateFormat = ProfileValue(rowIndex, PF_DATE_FMT)
    out.DescriptionColumns = ProfileValue(rowIndex, PF_DESC_COLS)
    out.AmountMode = ProfileValue(rowIndex, PF_AMOUNT_MODE)
    out.AmountColumn = ProfileNumber(rowIndex, PF_AMOUNT_COL, 0)
    out.DebitColumn = ProfileNumber(rowIndex, PF_DEBIT_COL, 0)
    out.CreditColumn = ProfileNumber(rowIndex, PF_CREDIT_COL, 0)
    out.Signature = ProfileValue(rowIndex, PF_SIGNATURE)
    Set Profile = out
End Function

Public Function AllProfiles() As Collection
    Dim lo As ListObject
    Dim out As Collection
    Dim i As Long

    Set out = New Collection
    Set AllProfiles = out
    Set lo = ProfilesTable()
    For i = 1 To modUtil.BodyRows(lo)
        If Len(ProfileValue(i, PF_NAME)) > 0 Then out.Add Profile(i)
    Next i
End Function

Public Function ProfileTitles(ByVal profiles As Collection) As Variant
    Dim names() As String
    Dim i As Long

    If profiles.Count = 0 Then Exit Function
    ReDim names(0 To profiles.Count - 1)
    For i = 1 To profiles.Count
        names(i - 1) = profiles.Item(i).Title()
    Next i
    ProfileTitles = names
End Function

'--- Recognising a file -----------------------------------------------------

' The profile whose signature matches the opening lines of the file, or
' Nothing when none of them does confidently.
Public Function MatchProfile(ByVal profiles As Collection, _
                             ByVal rows As Collection) As clsProfile
    Dim probe As String
    Dim i As Long, bestScore As Long, score As Long

    probe = ProbeText(rows)
    If Len(probe) = 0 Then Exit Function

    For i = 1 To profiles.Count
        score = SignatureScore(profiles.Item(i).Signature, probe)
        If score > bestScore Then
            bestScore = score
            Set MatchProfile = profiles.Item(i)
        End If
    Next i
End Function

Public Function DetectProfile(ByVal rows As Collection) As clsProfile
    Set DetectProfile = MatchProfile(AllProfiles(), rows)
End Function

Private Function ProbeText(ByVal rows As Collection) As String
    Dim i As Long, joined As String
    For i = 1 To rows.Count
        If i > 4 Then Exit For
        joined = joined & modParse.JoinFields(rows.Item(i), ",") & vbLf
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

Public Function AskForProfile(ByVal fileName As String) As clsProfile
    Dim profiles As Collection
    Dim choice As Long

    Set profiles = AllProfiles()
    If profiles.Count = 0 Then Exit Function
    choice = modUtil.AskChoice("Which format matches " & fileName & "?", _
                               ProfileTitles(profiles))
    If choice > 0 Then Set AskForProfile = profiles.Item(choice)
End Function

' Builds a human readable preview so the user can confirm the mapping before
' anything is written to the ledger.
Public Function PreviewText(ByVal rows As Collection, _
                            ByVal profile As clsProfile) As String
    Dim i As Long, shown As Long
    Dim txn As clsTxn
    Dim out As String

    out = "Format: " & profile.Name & vbCrLf & _
          "Institution: " & profile.Institution & vbCrLf & vbCrLf
    For i = profile.SkipRows + 1 To rows.Count
        Set txn = profile.ReadRow(rows.Item(i))
        If Not txn Is Nothing Then
            shown = shown + 1
            out = out & Format$(txn.TxnDate, "yyyy-mm-dd") & "   " & _
                  Format$(txn.Amount, "#,##0.00;-#,##0.00") & "   " & _
                  Left$(txn.Description, 42) & vbCrLf
            If shown >= 4 Then Exit For
        End If
    Next i
    If shown = 0 Then out = out & "(no readable rows found)" & vbCrLf
    PreviewText = out
End Function
