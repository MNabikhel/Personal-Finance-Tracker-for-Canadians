Attribute VB_Name = "modReport"
Option Explicit

'== Refresh and navigation ==================================================
' Almost every number in this workbook is a live formula, so refreshing mostly
' means recalculating.  The one thing formulas cannot do without dynamic arrays
' is rank merchants, so that list is written here.
'=============================================================================

Public Const TOP_MERCHANT_COUNT As Long = 10

Public Sub RefreshAll()
    On Error GoTo Fail
    modUtil.FastMode True
    RefreshTopMerchants
    modUtil.FastMode False
    Application.CalculateFull
    modUtil.Status ""
    Exit Sub
Fail:
    modUtil.ReportError "RefreshAll"
End Sub

' Fills the "Biggest merchants" block on the Dashboard for the selected month
' and view.  The block starts at the named range TopMerchants (2 columns).
Public Sub RefreshTopMerchants()
    Dim lo As ListObject
    Dim target As Range
    Dim merchants As Variant, months As Variant, views As Variant
    Dim types As Variant
    Dim totals As Collection
    Dim names() As String, values() As Double
    Dim i As Long, count As Long
    Dim monthWanted As String
    Dim merchant As String
    Dim rows As Long

    On Error Resume Next
    Set target = modUtil.Sh(SH_DASHBOARD).Range("TopMerchants")
    On Error GoTo 0
    If target Is Nothing Then Exit Sub

    target.ClearContents
    rows = target.Rows.Count

    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then Exit Sub

    monthWanted = modUtil.NzStr(modUtil.Sh(SH_DASHBOARD).Range("ReportMonth").Value)

    merchants = modLedger.ReadColumn(lo, COL_MERCHANT)
    months = modLedger.ReadColumn(lo, COL_MONTH)
    types = modLedger.ReadColumn(lo, COL_TYPE)
    ' The "View Amount" column already holds the household total or one
    ' person's share, whichever the Dashboard is set to.
    views = modLedger.ReadColumn(lo, COL_VIEW)

    Set totals = New Collection
    For i = 1 To UBound(views, 1)
        If StrComp(modUtil.NzStr(types(i, 1)), "Expense", vbTextCompare) = 0 Then
            If Len(monthWanted) = 0 Or _
               StrComp(modUtil.NzStr(months(i, 1)), monthWanted, vbTextCompare) = 0 Then
                merchant = modUtil.NzStr(merchants(i, 1), "(no merchant)")
                modUtil.PutVal totals, merchant, _
                    CDbl(modUtil.GetVal(totals, merchant, 0)) + _
                    -modUtil.NzNum(views(i, 1))
            End If
        End If
    Next i

    If totals.Count = 0 Then Exit Sub

    ReDim names(1 To totals.Count)
    ReDim values(1 To totals.Count)
    ' Collections cannot be enumerated by key, so rebuild the pairs by walking
    ' the ledger once more in order.
    count = 0
    For i = 1 To UBound(merchants, 1)
        merchant = modUtil.NzStr(merchants(i, 1), "(no merchant)")
        If modUtil.HasKey(totals, merchant) Then
            ' Refunds are positive expense rows and must reduce what was
            ' spent at the merchant.  A fully refunded (or net-credit)
            ' merchant is not a "biggest spend".
            If CDbl(modUtil.GetVal(totals, merchant, 0)) > 0 Then
                count = count + 1
                names(count) = merchant
                values(count) = CDbl(modUtil.GetVal(totals, merchant, 0))
            End If
            totals.Remove merchant
        End If
    Next i

    SortDescending names, values, count
    If count > rows Then count = rows
    If count > TOP_MERCHANT_COUNT Then count = TOP_MERCHANT_COUNT

    For i = 1 To count
        target.Cells(i, 1).Value = names(i)
        target.Cells(i, 2).Value = values(i)
    Next i
End Sub

Private Sub SortDescending(ByRef names() As String, ByRef values() As Double, _
                           ByVal count As Long)
    Dim i As Long, j As Long
    Dim keyName As String
    Dim keyValue As Double

    For i = 2 To count
        keyName = names(i)
        keyValue = values(i)
        j = i - 1
        Do While j >= 1
            If values(j) >= keyValue Then Exit Do
            names(j + 1) = names(j)
            values(j + 1) = values(j)
            j = j - 1
        Loop
        names(j + 1) = keyName
        values(j + 1) = keyValue
    Next i
End Sub

'--- Navigation -------------------------------------------------------------

Public Sub GoToDashboard()
    On Error Resume Next
    modUtil.Sh(SH_DASHBOARD).Activate
    modUtil.Sh(SH_DASHBOARD).Range("A1").Select
End Sub

Public Sub ShowUncategorized()
    Dim lo As ListObject
    Dim column As Long

    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    modUtil.Sh(SH_TXN).Activate
    If modUtil.BodyRows(lo) = 0 Then
        MsgBox "There are no transactions yet.", vbInformation, APP_NAME
        Exit Sub
    End If

    column = modUtil.ColumnIndex(lo, COL_CATEGORY)
    If lo.AutoFilter Is Nothing Then lo.Range.AutoFilter
    lo.Range.AutoFilter Field:=column, Criteria1:=Array(CAT_UNCATEGORIZED, "="), _
                        Operator:=xlFilterValues
    lo.Range.Cells(2, 1).Select

    MsgBox modRules.CountUncategorized() & " transaction(s) need a category." & vbCrLf & _
           vbCrLf & "Pick one from the Category drop-down, or select a row and use " & _
           """Teach a rule"" so future imports know what to do.", _
           vbInformation, APP_NAME
    Exit Sub
Fail:
    modUtil.ReportError "ShowUncategorized"
End Sub

Public Sub ClearLedgerFilters()
    Dim lo As ListObject
    On Error Resume Next
    Set lo = modUtil.TxnTable()
    If lo.ShowAutoFilter Then lo.AutoFilter.ShowAllData
    modUtil.Sh(SH_TXN).Activate
End Sub
