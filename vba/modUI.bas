Attribute VB_Name = "modUI"
Option Explicit

'== Buttons and shortcuts ===================================================
' Excel cannot store a button that runs a macro inside a file built outside of
' Excel, so the buttons are drawn on first open instead.  They are rebuilt from
' scratch every time, which also repairs a workbook where someone deleted one.
'=============================================================================

Private Const BTN_PREFIX As String = "cftBtn_"
Private Const BTN_WIDTH As Single = 132
Private Const BTN_HEIGHT As Single = 26
Private Const BTN_GAP As Single = 6
' Two rows of buttons fit above each table's header; the ledger's bar is
' wider because the ledger is.
Private Const BTN_PER_ROW As Long = 4
Private Const BTN_PER_LEDGER_ROW As Long = 5

Public Sub EnsureButtons()
    On Error Resume Next
    DrawBar modUtil.Sh(SH_DASHBOARD), "ButtonAnchor", DashboardButtons(), BTN_PER_ROW
    DrawBar modUtil.Sh(SH_TXN), "TxnButtonAnchor", LedgerButtons(), BTN_PER_LEDGER_ROW
    On Error GoTo 0
End Sub

Private Function DashboardButtons() As Variant
    DashboardButtons = Array( _
        "Import statements", "modImport.ImportStatements", _
        "Apply rules", "modRules.CategorizeUncategorized", _
        "Find transfers", "modTransfers.DetectTransfers", _
        "Needs a category", "modReport.ShowUncategorized", _
        "Refresh", "modReport.RefreshAll", _
        "Setup wizard", "modSetup.RunSetupWizard", _
        "Couple mode on/off", "modHousehold.ToggleHouseholdMode", _
        "Help", "modSetup.ShowHelp")
End Function

Private Function LedgerButtons() As Variant
    LedgerButtons = Array( _
        "Import statements", "modImport.ImportStatements", _
        "Undo an import", "modImport.UndoImport", _
        "Teach a rule", "modRules.TeachRuleFromSelection", _
        "Set owner", "modHousehold.SetOwnerForSelection", _
        "Show all rows", "modReport.ClearLedgerFilters", _
        "Apply rules to all", "modRules.RecategorizeAll", _
        "Rebuild formulas", "modLedger.RepairFormulas", _
        "Start fresh", "modSetup.ClearAllTransactions", _
        "Back to dashboard", "modReport.GoToDashboard")
End Function

Private Sub DrawBar(ByVal target As Worksheet, ByVal anchorName As String, _
                    ByVal definitions As Variant, ByVal perRow As Long)
    Dim anchor As Range
    Dim i As Long, index As Long
    Dim btn As Shape
    Dim posLeft As Single, posTop As Single

    RemoveButtons target

    On Error Resume Next
    Set anchor = target.Range(anchorName)
    On Error GoTo 0
    If anchor Is Nothing Then Set anchor = target.Range("B3")

    For i = LBound(definitions) To UBound(definitions) - 1 Step 2
        posLeft = anchor.Left + (index Mod perRow) * (BTN_WIDTH + BTN_GAP)
        posTop = anchor.Top + (index \ perRow) * (BTN_HEIGHT + BTN_GAP)

        Set btn = target.Shapes.AddShape(msoShapeRoundedRectangle, posLeft, posTop, _
                                         BTN_WIDTH, BTN_HEIGHT)
        With btn
            .Name = BTN_PREFIX & index
            .OnAction = CStr(definitions(i + 1))
            .Fill.ForeColor.RGB = RGB(196, 30, 58)      ' Canadian red
            .Fill.Solid
            .Line.ForeColor.RGB = RGB(150, 20, 42)
            .Adjustments(1) = 0.16
            With .TextFrame2
                .TextRange.Text = CStr(definitions(i))
                .TextRange.Font.Size = 10
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .VerticalAnchor = msoAnchorMiddle
                .HorizontalAnchor = msoAnchorCenter
                .WordWrap = msoFalse
                .MarginLeft = 2
                .MarginRight = 2
            End With
            .Placement = xlFreeFloating
        End With
        index = index + 1
    Next i
End Sub

Public Sub RemoveButtons(ByVal target As Worksheet)
    Dim i As Long
    For i = target.Shapes.Count To 1 Step -1
        If Left$(target.Shapes(i).Name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            target.Shapes(i).Delete
        End If
    Next i
End Sub

'--- Keyboard shortcuts -----------------------------------------------------

Public Sub EnableShortcuts()
    On Error Resume Next
    Application.OnKey "+^{i}", InThisWorkbook("modImport.ImportStatements")
    Application.OnKey "+^{r}", InThisWorkbook("modRules.CategorizeUncategorized")
    Application.OnKey "+^{t}", InThisWorkbook("modRules.TeachRuleFromSelection")
    Application.OnKey "+^{u}", InThisWorkbook("modReport.ShowUncategorized")
    On Error GoTo 0
End Sub

' A key binding is application-wide, so the macro is named with its workbook
' or Excel looks for it in whichever workbook happens to be active.
Private Function InThisWorkbook(ByVal macro As String) As String
    InThisWorkbook = "'" & ThisWorkbook.Name & "'!" & macro
End Function

Public Sub DisableShortcuts()
    On Error Resume Next
    Application.OnKey "+^{i}"
    Application.OnKey "+^{r}"
    Application.OnKey "+^{t}"
    Application.OnKey "+^{u}"
    On Error GoTo 0
End Sub
