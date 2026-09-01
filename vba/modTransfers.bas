Attribute VB_Name = "modTransfers"
Option Explicit

'== Internal transfer detection =============================================
' A credit card payment or a move between your own accounts is not income and
' not an expense - it is the same money twice.  This finds those mirrored pairs
' and tags both sides so the reports stay honest.
'=============================================================================

Public Sub DetectTransfers(Optional ByVal interactive As Boolean = True)
    Dim lo As ListObject
    Dim rowCount As Long
    Dim dates As Variant, accounts As Variant, amounts As Variant
    Dim categories As Variant, tags As Variant
    Dim buckets As Collection
    Dim matched() As Boolean
    Dim i As Long, j As Long, pairs As Long
    Dim windowDays As Long
    Dim key As String
    Dim candidates() As String
    Dim k As Long
    Dim categoryName As String

    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    rowCount = modUtil.BodyRows(lo)
    If rowCount < 2 Then
        If interactive Then MsgBox "Not enough transactions to compare.", _
                                   vbInformation, APP_NAME
        Exit Sub
    End If

    windowDays = CLng(modUtil.NzNum(modUtil.Setting(NR_TRANSFER_DAYS), 4))
    modUtil.FastMode True

    dates = modLedger.ReadColumn(lo, COL_DATE)
    accounts = modLedger.ReadColumn(lo, COL_ACCOUNT)
    amounts = modLedger.ReadColumn(lo, COL_AMOUNT)
    categories = modLedger.ReadColumn(lo, COL_CATEGORY)
    tags = modLedger.ReadColumn(lo, COL_TAGGEDBY)

    ReDim matched(1 To rowCount)
    Set buckets = New Collection

    ' Bucket every eligible row by the absolute amount so the search stays linear
    ' in practice instead of comparing every row with every other row.
    For i = 1 To rowCount
        If Eligible(categories, amounts, i) Then
            key = Format$(Abs(modUtil.NzNum(amounts(i, 1))), "0.00")
            modUtil.PutVal buckets, key, _
                CStr(modUtil.GetVal(buckets, key, "")) & "," & i
        End If
    Next i

    For i = 1 To rowCount
        If Not matched(i) Then
            If Eligible(categories, amounts, i) And modUtil.NzNum(amounts(i, 1)) < 0 Then
                key = Format$(Abs(modUtil.NzNum(amounts(i, 1))), "0.00")
                candidates = Split(CStr(modUtil.GetVal(buckets, key, "")), ",")
                For k = LBound(candidates) To UBound(candidates)
                    If Len(candidates(k)) > 0 Then
                        j = CLng(candidates(k))
                        If j <> i Then
                            If Not matched(j) Then
                                If IsMirror(dates, accounts, amounts, i, j, windowDays) Then
                                    If modAccounts.IsCreditAccount( _
                                            modUtil.NzStr(accounts(j, 1))) Then
                                        categoryName = CAT_CARD_PAYMENT
                                    Else
                                        categoryName = CAT_TRANSFER
                                    End If
                                    categories(i, 1) = categoryName
                                    categories(j, 1) = categoryName
                                    tags(i, 1) = TAG_TRANSFER
                                    tags(j, 1) = TAG_TRANSFER
                                    matched(i) = True
                                    matched(j) = True
                                    pairs = pairs + 1
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next k
            End If
        End If
    Next i

    If pairs > 0 Then
        modLedger.WriteColumn lo, 1, rowCount, COL_CATEGORY, categories
        modLedger.WriteColumn lo, 1, rowCount, COL_TAGGEDBY, tags
    End If
    modUtil.FastMode False

    If interactive Then
        Application.Calculate
        MsgBox pairs & " transfer pair(s) were tagged." & vbCrLf & vbCrLf & _
               "Transfers are excluded from income and expense totals.", _
               vbInformation, APP_NAME
    End If
    Exit Sub

Fail:
    modUtil.ReportError "DetectTransfers"
End Sub

' Only rows that nobody has deliberately categorised are candidates.
Private Function Eligible(ByRef categories As Variant, ByRef amounts As Variant, _
                          ByVal rowIndex As Long) As Boolean
    Dim category As String
    If Len(modUtil.NzStr(amounts(rowIndex, 1))) = 0 Then Exit Function
    If modUtil.NzNum(amounts(rowIndex, 1)) = 0 Then Exit Function
    category = modUtil.NzStr(categories(rowIndex, 1))
    Select Case True
        Case Len(category) = 0
            Eligible = True
        Case StrComp(category, CAT_UNCATEGORIZED, vbTextCompare) = 0
            Eligible = True
        Case StrComp(category, CAT_TRANSFER, vbTextCompare) = 0
            Eligible = True
        Case StrComp(category, CAT_CARD_PAYMENT, vbTextCompare) = 0
            Eligible = True
    End Select
End Function

Private Function IsMirror(ByRef dates As Variant, ByRef accounts As Variant, _
                          ByRef amounts As Variant, ByVal outRow As Long, _
                          ByVal inRow As Long, ByVal windowDays As Long) As Boolean
    Dim outAmount As Double, inAmount As Double
    Dim gap As Double

    outAmount = modUtil.NzNum(amounts(outRow, 1))
    inAmount = modUtil.NzNum(amounts(inRow, 1))
    If Abs(outAmount + inAmount) > 0.005 Then Exit Function
    If inAmount <= 0 Then Exit Function
    If StrComp(modUtil.NzStr(accounts(outRow, 1)), _
               modUtil.NzStr(accounts(inRow, 1)), vbTextCompare) = 0 Then Exit Function

    gap = Abs(modUtil.NzNum(dates(outRow, 1)) - modUtil.NzNum(dates(inRow, 1)))
    IsMirror = (gap <= windowDays)
End Function
