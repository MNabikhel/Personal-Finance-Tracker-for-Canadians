Attribute VB_Name = "modLedger"
Option Explicit

'== Ledger plumbing =========================================================
' Growing the transactions table, re-applying the calculated-column formulas
' and writing whole columns in one shot.
'
' The formulas themselves are not hard-coded here: they live in the "Engine"
' sheet (tblTemplates) so the workbook stays the single source of truth.
'=============================================================================

Public Const TBL_TEMPLATES As String = "tblTemplates"

Public Function AddRows(ByVal lo As ListObject, ByVal count As Long) As Long
    Dim existing As Long
    Dim firstNew As Long

    If count < 1 Then Exit Function
    existing = modUtil.BodyRows(lo)

    If existing = 0 Then
        lo.Resize lo.Range.Resize(count + lo.HeaderRowRange.Rows.Count)
        firstNew = 1
    ElseIf existing = 1 And IsRowBlank(lo, 1) Then
        firstNew = 1
        If count > 1 Then
            lo.Resize lo.Range.Resize(lo.Range.Rows.Count + count - 1)
        End If
    Else
        firstNew = existing + 1
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + count)
    End If

    ApplyTemplates lo, firstNew, count
    ApplyFormats lo, firstNew, count
    SyncPrintArea lo
    AddRows = firstNew
End Function

' The ledger is formatted thousands of rows past the ones in use, so with no
' print area Excel offers to print all of it.  The area is set when the
' workbook is built, out to the Type column, and has to follow the table down
' from there as rows arrive.
Public Sub SyncPrintArea(ByVal lo As ListObject)
    Dim corner As Range
    On Error Resume Next
    Set corner = lo.Range.Cells(lo.Range.Rows.Count, modUtil.ColumnIndex(lo, COL_TYPE))
    lo.Parent.PageSetup.PrintArea = "$B$1:" & corner.Address
    On Error GoTo 0
End Sub

Public Function IsRowBlank(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim dateCell As Range, amountCell As Range
    Set dateCell = modUtil.CellIn(lo, rowIndex, COL_DATE)
    Set amountCell = modUtil.CellIn(lo, rowIndex, COL_AMOUNT)
    IsRowBlank = (Len(modUtil.NzStr(dateCell.Value)) = 0 And _
                  Len(modUtil.NzStr(amountCell.Value)) = 0)
End Function

Public Sub WriteColumn(ByVal lo As ListObject, ByVal firstRow As Long, _
                       ByVal count As Long, ByVal header As String, _
                       ByRef values As Variant)
    Dim target As Range
    Set target = lo.DataBodyRange.Cells(firstRow, modUtil.ColumnIndex(lo, header))
    target.Resize(count, 1).Value = values
End Sub

' Writes the calculated-column formulas over a block of rows.  The templates are
' stored in R1C1 ("RC3" means "column 3 of this row"), which is independent of
' where the block starts, so one string is correct for every row in it.
Public Sub ApplyTemplates(ByVal lo As ListObject, ByVal firstRow As Long, _
                          ByVal count As Long)
    Dim templates As ListObject
    Dim i As Long
    Dim header As String, formula As String
    Dim target As Range

    On Error Resume Next
    Set templates = modUtil.Tbl(SH_ENGINE, TBL_TEMPLATES)
    On Error GoTo 0
    If templates Is Nothing Then Exit Sub
    If modUtil.BodyRows(templates) = 0 Then Exit Sub

    For i = 1 To modUtil.BodyRows(templates)
        header = modUtil.NzStr(templates.DataBodyRange.Cells(i, 1).Value)
        formula = modUtil.NzStr(templates.DataBodyRange.Cells(i, 2).Value)
        If Len(header) > 0 And Len(formula) > 0 Then
            On Error Resume Next
            Set target = Nothing
            Set target = lo.DataBodyRange.Cells(firstRow, _
                            modUtil.ColumnIndex(lo, header)).Resize(count, 1)
            On Error GoTo 0
            If Not target Is Nothing Then target.FormulaR1C1 = "=" & formula
        End If
    Next i
End Sub

Public Sub ApplyFormats(ByVal lo As ListObject, ByVal firstRow As Long, _
                        ByVal count As Long)
    FormatBlock lo, firstRow, count, COL_DATE, "yyyy-mm-dd"
    FormatBlock lo, firstRow, count, COL_AMOUNT, "#,##0.00;[Red]-#,##0.00"
    FormatBlock lo, firstRow, count, COL_SHARE_A, "#,##0.00;[Red]-#,##0.00"
    FormatBlock lo, firstRow, count, COL_SHARE_B, "#,##0.00;[Red]-#,##0.00"
    FormatBlock lo, firstRow, count, COL_VIEW, "#,##0.00;[Red]-#,##0.00"
    FormatBlock lo, firstRow, count, COL_SPLIT, "0%"
End Sub

Private Sub FormatBlock(ByVal lo As ListObject, ByVal firstRow As Long, _
                        ByVal count As Long, ByVal header As String, _
                        ByVal numberFormat As String)
    On Error Resume Next
    lo.DataBodyRange.Cells(firstRow, modUtil.ColumnIndex(lo, header)) _
        .Resize(count, 1).NumberFormat = numberFormat
    On Error GoTo 0
End Sub

' Menu action: rebuild every calculated column, for instance after a formula
' was overwritten by accident.
Public Sub RepairFormulas()
    Dim lo As ListObject
    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then
        MsgBox "There are no transactions yet.", vbInformation, APP_NAME
        Exit Sub
    End If
    modUtil.FastMode True
    ApplyTemplates lo, 1, modUtil.BodyRows(lo)
    ApplyFormats lo, 1, modUtil.BodyRows(lo)
    modUtil.FastMode False
    Application.Calculate
    MsgBox "Calculated columns were rebuilt for " & modUtil.BodyRows(lo) & " row(s).", _
           vbInformation, APP_NAME
    Exit Sub
Fail:
    modUtil.ReportError "RepairFormulas"
End Sub

' The number the next transaction ID should carry: one past the highest in the
' ledger, not the row count, so that IDs stay unique after rows are removed.
Public Function NextTxnNumber(ByVal lo As ListObject) As Long
    Dim ids As Variant
    Dim i As Long
    Dim text As String
    Dim highest As Long

    ids = ReadColumn(lo, COL_ID)
    If IsArray(ids) Then
        For i = 1 To UBound(ids, 1)
            text = modUtil.NzStr(ids(i, 1))
            If Left$(text, 1) = "T" And IsNumeric(Mid$(text, 2)) Then
                If CLng(Val(Mid$(text, 2))) > highest Then highest = CLng(Val(Mid$(text, 2)))
            End If
        Next i
    End If
    NextTxnNumber = highest + 1
End Function

' Deletes every row whose value in the column equals the given text, and says
' how many went.  A table that would be left empty keeps one blank row with
' its formulas, which is how the rest of the code expects an empty ledger.
Public Function DeleteRowsWhere(ByVal lo As ListObject, ByVal header As String, _
                                ByVal wanted As String) As Long
    Dim values As Variant
    Dim i As Long
    Dim removed As Long

    values = ReadColumn(lo, header)
    If Not IsArray(values) Then Exit Function

    For i = UBound(values, 1) To 1 Step -1
        If StrComp(modUtil.NzStr(values(i, 1)), wanted, vbTextCompare) = 0 Then
            If modUtil.BodyRows(lo) > 1 Then
                lo.ListRows(i).Delete
            Else
                lo.DataBodyRange.Rows(1).ClearContents
                ApplyTemplates lo, 1, 1
            End If
            removed = removed + 1
        End If
    Next i
    If removed > 0 Then SyncPrintArea lo
    DeleteRowsWhere = removed
End Function

' Reads one column of the ledger into a 1-based array, or an empty Variant when
' the ledger is empty.
Public Function ReadColumn(ByVal lo As ListObject, ByVal header As String) As Variant
    Dim count As Long
    count = modUtil.BodyRows(lo)
    If count = 0 Then Exit Function
    If count = 1 Then
        ' A single-cell range returns a scalar, so wrap it to keep callers simple.
        Dim oneCell(1 To 1, 1 To 1) As Variant
        oneCell(1, 1) = lo.DataBodyRange.Cells(1, modUtil.ColumnIndex(lo, header)).Value
        ReadColumn = oneCell
    Else
        ReadColumn = lo.ListColumns(modUtil.ColumnIndex(lo, header)).DataBodyRange.Value
    End If
End Function
