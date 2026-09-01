Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = True
Attribute VB_Customizable = True
Option Explicit

'== Workbook events =========================================================

Private Sub Workbook_Open()
    On Error Resume Next

    modUI.EnsureButtons
    modUI.EnableShortcuts
    modHousehold.ApplyMode
    modReport.RefreshTopMerchants
    Application.Goto modUtil.Sh(SH_DASHBOARD).Range("A1"), True

    If StrComp(modUtil.NzStr(modUtil.Setting(NR_CONFIGURED)), "Yes", vbTextCompare) <> 0 Then
        MsgBox APP_NAME & " " & APP_VERSION & vbCrLf & vbCrLf & _
               "The sample data is here so you can look around. When you are ready, " & _
               "press ""Setup wizard"" to enter your own details - it will offer to " & _
               "clear the samples.", vbInformation, APP_NAME
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modUI.DisableShortcuts
End Sub

' A category typed in by hand outranks the rules: tag the row as Manual so that
' "Apply rules to all" leaves it alone.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim lo As ListObject
    Dim touched As Range
    Dim cell As Range
    Dim taggedColumn As Long

    If Sh.Name <> SH_TXN Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub

    On Error Resume Next
    Set lo = modUtil.TxnTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set touched = Application.Intersect(Target, lo.ListColumns(COL_CATEGORY).DataBodyRange)
    If touched Is Nothing Then Exit Sub
    If touched.Cells.Count > 500 Then Exit Sub

    taggedColumn = modUtil.ColumnIndex(lo, COL_TAGGEDBY) - _
                   modUtil.ColumnIndex(lo, COL_CATEGORY)

    Application.EnableEvents = False
    For Each cell In touched.Cells
        If Len(modUtil.NzStr(cell.Value)) > 0 Then
            cell.Offset(0, taggedColumn).Value = TAG_MANUAL
        End If
    Next cell
    Application.EnableEvents = True
End Sub
