Attribute VB_Name = "modAccounts"
Option Explicit

'== Accounts and the import log =============================================
' Works out which account a downloaded file belongs to and records what each
' import batch did.
'=============================================================================

Public Const AC_NAME As String = "Account"
Public Const AC_INSTITUTION As String = "Institution"
Public Const AC_TYPE As String = "Type"
Public Const AC_OWNER As String = "Owner"
Public Const AC_FORMAT As String = "Bank Format"
Public Const AC_FILEMATCH As String = "File Name Contains"
Public Const AC_NOTES As String = "Notes"

' Import Log column headers
Public Const LG_WHEN As String = "When"
Public Const LG_BATCH As String = "Batch"
Public Const LG_FILE As String = "File"
Public Const LG_IMPORTED As String = "Imported"
Public Const LG_STATUS As String = "Status"

Public Function AccountsTable() As ListObject
    Set AccountsTable = modUtil.Tbl(SH_ACCOUNTS, TBL_ACCOUNTS)
End Function

Public Function AccountValue(ByVal rowIndex As Long, ByVal header As String) As String
    Dim lo As ListObject
    Set lo = AccountsTable()
    AccountValue = modUtil.NzStr(lo.DataBodyRange.Cells(rowIndex, _
                   modUtil.ColumnIndex(lo, header)).Value)
End Function

Public Function AccountRowCount() As Long
    Dim lo As ListObject, i As Long, total As Long
    Set lo = AccountsTable()
    For i = 1 To modUtil.BodyRows(lo)
        If Len(AccountValue(i, AC_NAME)) > 0 Then total = total + 1
    Next i
    AccountRowCount = total
End Function

Public Function AccountNames() As Variant
    Dim lo As ListObject, i As Long, names() As String, found As Long
    Dim total As Long
    total = AccountRowCount()
    If total = 0 Then Exit Function
    Set lo = AccountsTable()
    ReDim names(0 To total - 1)
    For i = 1 To modUtil.BodyRows(lo)
        If Len(AccountValue(i, AC_NAME)) > 0 Then
            names(found) = AccountValue(i, AC_NAME)
            found = found + 1
        End If
    Next i
    AccountNames = names
End Function

Public Function AccountRow(ByVal accountName As String) As Long
    Dim lo As ListObject, i As Long
    Set lo = AccountsTable()
    For i = 1 To modUtil.BodyRows(lo)
        If StrComp(AccountValue(i, AC_NAME), accountName, vbTextCompare) = 0 Then
            AccountRow = i
            Exit Function
        End If
    Next i
End Function

Public Function OwnerOfAccount(ByVal accountName As String) As String
    Dim rowIndex As Long
    Dim owner As String
    rowIndex = AccountRow(accountName)
    If rowIndex > 0 Then owner = AccountValue(rowIndex, AC_OWNER)
    If Len(owner) = 0 Then
        If modUtil.IsCoupleMode() Then
            owner = OWNER_JOINT
        Else
            owner = modUtil.PersonAName()
        End If
    End If
    OwnerOfAccount = owner
End Function

Public Function IsCreditAccount(ByVal accountName As String) As Boolean
    Dim rowIndex As Long
    rowIndex = AccountRow(accountName)
    If rowIndex = 0 Then Exit Function
    Select Case UCase$(AccountValue(rowIndex, AC_TYPE))
        Case "CREDIT CARD", "LINE OF CREDIT", "LOAN", "MORTGAGE"
            IsCreditAccount = True
    End Select
End Function

' Decides which account a file belongs to: first by the "File Name Contains"
' hint, then by the bank format, and finally by asking.
Public Function ResolveAccount(ByVal fileName As String, _
                               ByVal profile As clsProfile) As String
    Dim lo As ListObject
    Dim i As Long
    Dim hint As String
    Dim bestHintLength As Long
    Dim bestHintAccount As String
    Dim bestHintTied As Boolean
    Dim profileName As String
    Dim candidate As String
    Dim candidateCount As Long
    Dim names As Variant
    Dim choice As Long

    Set lo = AccountsTable()
    profileName = profile.Name

    For i = 1 To modUtil.BodyRows(lo)
        hint = AccountValue(i, AC_FILEMATCH)
        If Len(hint) > 0 And Len(AccountValue(i, AC_NAME)) > 0 Then
            If InStr(1, fileName, hint, vbTextCompare) > 0 Then
                If Len(hint) > bestHintLength Then
                    bestHintLength = Len(hint)
                    bestHintAccount = AccountValue(i, AC_NAME)
                    bestHintTied = False
                ElseIf Len(hint) = bestHintLength And _
                       StrComp(bestHintAccount, AccountValue(i, AC_NAME), _
                               vbTextCompare) <> 0 Then
                    bestHintTied = True
                End If
            End If
        End If
    Next i
    ' A more specific snippet ("rbc-visa") wins over a broad one ("rbc").
    ' Equal best matches are ambiguous and fall through to the account picker.
    If bestHintLength > 0 And Not bestHintTied Then
        ResolveAccount = bestHintAccount
        Exit Function
    End If

    ' The format settles it only when exactly one account uses it.
    For i = 1 To modUtil.BodyRows(lo)
        If Len(AccountValue(i, AC_NAME)) > 0 Then
            If StrComp(AccountValue(i, AC_FORMAT), profileName, vbTextCompare) = 0 Then
                candidate = AccountValue(i, AC_NAME)
                candidateCount = candidateCount + 1
            End If
        End If
    Next i
    If candidateCount = 1 Then
        ResolveAccount = candidate
        Exit Function
    End If

    names = AccountNames()
    If IsEmpty(names) Then
        ResolveAccount = CreateAccountInteractively(profile)
        Exit Function
    End If

    choice = modUtil.AskChoice("Which account does " & fileName & " belong to?" & _
                               vbCrLf & "(Tip: fill in ""File Name Contains"" on the " & _
                               "Accounts sheet to skip this question next time.)", names)
    If choice = 0 Then Exit Function
    ResolveAccount = names(choice - 1)
End Function

Public Function CreateAccountInteractively(ByVal profile As clsProfile) As String
    Dim accountName As String
    Dim accountType As Long
    Dim ownerChoice As Long
    Dim owners As Variant
    Dim types As Variant

    accountName = Trim$(InputBox("No accounts are set up yet." & vbCrLf & vbCrLf & _
        "Name this account (for example ""RBC Chequing"" or ""Amex Cobalt""):", _
        APP_NAME))
    If Len(accountName) = 0 Then Exit Function

    types = Array("Chequing", "Savings", "Credit Card", "Line of Credit", "Cash")
    accountType = modUtil.AskChoice("What kind of account is " & accountName & "?", types)
    If accountType = 0 Then accountType = 1

    If modUtil.IsCoupleMode() Then
        owners = Array(modUtil.PersonAName(), modUtil.PersonBName(), OWNER_JOINT)
        ownerChoice = modUtil.AskChoice("Who owns " & accountName & "?", owners)
        If ownerChoice = 0 Then ownerChoice = 3
    Else
        owners = Array(modUtil.PersonAName())
        ownerChoice = 1
    End If

    AddAccount accountName, profile.Institution, _
               CStr(types(accountType - 1)), CStr(owners(ownerChoice - 1)), _
               profile.Name
    CreateAccountInteractively = accountName
End Function

Public Sub AddAccount(ByVal accountName As String, ByVal institution As String, _
                      ByVal accountType As String, ByVal owner As String, _
                      ByVal bankFormat As String)
    Dim lo As ListObject
    Dim target As ListRow
    Dim rowIndex As Long

    Set lo = AccountsTable()
    rowIndex = FirstEmptyRow(lo)
    If rowIndex = 0 Then
        Set target = lo.ListRows.Add
        rowIndex = target.Index
    End If

    lo.DataBodyRange.Cells(rowIndex, modUtil.ColumnIndex(lo, AC_NAME)).Value = accountName
    lo.DataBodyRange.Cells(rowIndex, modUtil.ColumnIndex(lo, AC_INSTITUTION)).Value = institution
    lo.DataBodyRange.Cells(rowIndex, modUtil.ColumnIndex(lo, AC_TYPE)).Value = accountType
    lo.DataBodyRange.Cells(rowIndex, modUtil.ColumnIndex(lo, AC_OWNER)).Value = owner
    lo.DataBodyRange.Cells(rowIndex, modUtil.ColumnIndex(lo, AC_FORMAT)).Value = bankFormat
End Sub

' The setup wizard ships with example accounts so every report has context,
' but a user who accepts "start fresh" needs a genuinely clean account list.
' Only rows still marked as samples are cleared; anything the user added or
' renamed and re-noted is preserved.
Public Function ClearSampleAccounts() As Long
    Dim lo As ListObject
    Dim i As Long
    Dim notesColumn As Long

    Set lo = AccountsTable()
    notesColumn = modUtil.ColumnIndex(lo, AC_NOTES)
    For i = 1 To modUtil.BodyRows(lo)
        If StrComp(modUtil.NzStr(lo.DataBodyRange.Cells(i, notesColumn).Value), _
                   "Sample account", vbTextCompare) = 0 Then
            lo.DataBodyRange.Rows(i).ClearContents
            ClearSampleAccounts = ClearSampleAccounts + 1
        End If
    Next i
End Function

Private Function FirstEmptyRow(ByVal lo As ListObject) As Long
    Dim i As Long
    For i = 1 To modUtil.BodyRows(lo)
        If Len(modUtil.NzStr(lo.DataBodyRange.Cells(i, _
               modUtil.ColumnIndex(lo, AC_NAME)).Value)) = 0 Then
            FirstEmptyRow = i
            Exit Function
        End If
    Next i
End Function

'--- Import log -------------------------------------------------------------

Public Sub LogBatch(ByVal batchId As String, ByVal fileName As String, _
                    ByVal profileName As String, ByVal accountName As String, _
                    ByVal rowsRead As Long, ByVal imported As Long, _
                    ByVal duplicates As Long, ByVal unreadable As Long)
    Dim lo As ListObject
    Dim target As ListRow
    Dim rowIndex As Long

    On Error Resume Next
    Set lo = modUtil.Tbl(SH_LOG, TBL_LOG)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    rowIndex = 0
    If modUtil.BodyRows(lo) = 1 Then
        If Len(modUtil.NzStr(lo.DataBodyRange.Cells(1, 1).Value)) = 0 Then rowIndex = 1
    End If
    If rowIndex = 0 Then
        Set target = lo.ListRows.Add
        rowIndex = target.Index
    End If

    With lo.DataBodyRange
        .Cells(rowIndex, 1).Value = Now
        .Cells(rowIndex, 1).NumberFormat = "yyyy-mm-dd hh:mm"
        .Cells(rowIndex, 2).Value = batchId
        .Cells(rowIndex, 3).Value = fileName
        .Cells(rowIndex, 4).Value = profileName
        .Cells(rowIndex, 5).Value = accountName
        .Cells(rowIndex, 6).Value = rowsRead
        .Cells(rowIndex, 7).Value = imported
        .Cells(rowIndex, 8).Value = duplicates
        .Cells(rowIndex, 9).Value = unreadable
    End With
End Sub
