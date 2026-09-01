Attribute VB_Name = "modRules"
Option Explicit

'== Merchant clean-up and the categorisation rules ==========================
' Bank descriptions are noisy ("IDP PURCHASE - 1234 LOBLAWS #1234 TORONTO ON").
' CleanMerchant boils that down to something readable, and the rules on the
' Rules sheet map it to a category.
'=============================================================================

Public Const RL_PRIORITY As String = "Priority"
Public Const RL_ENABLED As String = "Enabled"
Public Const RL_FIELD As String = "Look In"
Public Const RL_TEST As String = "Test"
Public Const RL_PATTERN As String = "Pattern"
Public Const RL_MIN As String = "Min Amount"
Public Const RL_MAX As String = "Max Amount"
Public Const RL_FLOW As String = "Flow"
Public Const RL_CATEGORY As String = "Category"
Public Const RL_OWNER As String = "Set Owner"
Public Const RL_HITS As String = "Hits"
Public Const RL_NOTES As String = "Notes"

Public Const CAT_NAME As String = "Category"
Public Const CAT_OWNER As String = "Default Owner"

' Maps a sorted rule position back to its row on the Rules sheet.
Private mRuleRowMap() As Long

Public Function RulesTable() As ListObject
    Set RulesTable = modUtil.Tbl(SH_RULES, TBL_RULES)
End Function

'--- Description clean-up ---------------------------------------------------

Public Function CleanMerchant(ByVal description As String) As String
    Dim text As String
    Dim noise As Variant
    Dim i As Long

    text = " " & UCase$(modUtil.CondenseSpaces(description)) & " "

    noise = Array( _
        "IDP PURCHASE -", "INTERAC PURCHASE -", "INTERAC RETAIL PURCHASE -", _
        "POINT OF SALE PURCHASE", "POINT OF SALE - INTERAC", "POS PURCHASE", _
        "VISA DEBIT PURCHASE -", "VISA DEBIT RETAIL PURCHASE -", "DEBIT PURCHASE -", _
        "RETAIL PURCHASE", "PREAUTHORIZED DEBIT", "PRE-AUTHORIZED DEBIT", _
        "PREAUTHORIZED PAYMENT", "ELECTRONIC FUNDS TRANSFER", "MISC PAYMENT", _
        "BILL PAYMENT", "ONLINE BANKING PAYMENT", "ONLINE BANKING TRANSFER", _
        "INTERNET BANKING", "WWW PAYMENT", "PAYMENT -", "PURCHASE -", "PURCHASE", _
        "TRANSACTION -", "ACH DEBIT", "ACH CREDIT", "CHQ#", "CHEQUE", "SQ *", "SQ*", _
        "TST*", "TST-", "SP ", "PAYPAL *", "PAYPAL*", "APPLE PAY", "GOOGLE PAY")

    For i = LBound(noise) To UBound(noise)
        text = Replace$(text, " " & noise(i) & " ", " ")
        text = Replace$(text, " " & noise(i), " ")
    Next i

    text = StripReferenceNumbers(text)
    text = StripProvinceTail(text)
    text = modUtil.CondenseSpaces(text)

    If Len(text) = 0 Then text = modUtil.CondenseSpaces(description)
    CleanMerchant = modUtil.TitleCaseWords(Left$(text, 60))
End Function

Private Function StripReferenceNumbers(ByVal text As String) As String
    ' Drops store numbers and long digit runs but keeps names such as "7-Eleven".
    Dim words() As String
    Dim i As Long
    Dim word As String
    Dim kept As String

    words = Split(modUtil.CondenseSpaces(text), " ")
    For i = LBound(words) To UBound(words)
        word = words(i)
        If Len(word) = 0 Then
            ' skip
        ElseIf word Like "#*" And Len(word) >= 4 And IsNumeric(Replace$(word, "#", "")) Then
            ' store number such as "#1234"
        ElseIf IsNumeric(word) And Len(word) >= 4 Then
            ' reference number
        ElseIf word Like "*[0-9][0-9][0-9][0-9][0-9]*" Then
            ' embedded long digit run
        Else
            If Len(kept) > 0 Then kept = kept & " "
            kept = kept & word
        End If
    Next i
    StripReferenceNumbers = kept
End Function

Private Function StripProvinceTail(ByVal text As String) As String
    ' Removes a trailing province/country code, e.g. "LOBLAWS TORONTO ON".
    Dim words() As String
    Dim upper As Long
    Dim tail As String

    words = Split(modUtil.CondenseSpaces(text), " ")
    upper = UBound(words)
    Do While upper >= 1
        tail = UCase$(words(upper))
        Select Case tail
            Case "ON", "QC", "BC", "AB", "MB", "SK", "NS", "NB", "PE", "NL", _
                 "YT", "NT", "NU", "CA", "CAN", "CANADA", "US", "USA"
                upper = upper - 1
            Case Else
                Exit Do
        End Select
    Loop
    StripProvinceTail = Join(SubArray(words, 0, upper), " ")
End Function

Private Function SubArray(ByRef source() As String, ByVal fromIndex As Long, _
                          ByVal toIndex As Long) As Variant
    Dim out() As String
    Dim i As Long
    If toIndex < fromIndex Then
        ReDim out(0 To 0)
        SubArray = out
        Exit Function
    End If
    ReDim out(0 To toIndex - fromIndex)
    For i = fromIndex To toIndex
        out(i - fromIndex) = source(i)
    Next i
    SubArray = out
End Function

'--- Applying rules ---------------------------------------------------------

Public Sub CategorizeUncategorized(Optional ByVal interactive As Boolean = True)
    ApplyRules False, interactive
End Sub

Public Sub RecategorizeAll(Optional ByVal interactive As Boolean = True)
    If interactive Then
        If MsgBox("Re-run the rules over every transaction?" & vbCrLf & vbCrLf & _
                  "Categories you set by hand will be kept; only rows tagged " & _
                  """Rule"", """ & TAG_IMPORT & """ or blank are touched.", _
                  vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub
    End If
    ApplyRules True, interactive
End Sub

Private Sub ApplyRules(ByVal includeTagged As Boolean, ByVal interactive As Boolean)
    Dim lo As ListObject
    Dim rowCount As Long
    Dim merchants As Variant, descriptions As Variant, amounts As Variant
    Dim accounts As Variant, categories As Variant, tags As Variant, owners As Variant
    Dim i As Long, matched As Long, ruleIndex As Long
    Dim ruleCount As Long
    Dim rulePattern() As String, ruleField() As String, ruleTest() As String
    Dim ruleCategory() As String, ruleOwner() As String, ruleFlow() As String
    Dim ruleMin() As Double, ruleMax() As Double
    Dim ruleHits() As Long
    Dim changed As Boolean
    Dim categoryOwners As Collection
    Dim couple As Boolean
    Dim newOwner As String

    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    rowCount = modUtil.BodyRows(lo)
    If rowCount = 0 Then
        If interactive Then MsgBox "There are no transactions yet.", vbInformation, APP_NAME
        Exit Sub
    End If

    ruleCount = LoadRules(rulePattern, ruleField, ruleTest, ruleCategory, ruleOwner, _
                          ruleFlow, ruleMin, ruleMax, ruleHits)
    If ruleCount = 0 Then
        If interactive Then
            MsgBox "There are no enabled rules on the Rules sheet.", vbInformation, APP_NAME
        End If
        Exit Sub
    End If

    couple = modUtil.IsCoupleMode()
    Set categoryOwners = CategoryOwners()

    modUtil.FastMode True
    merchants = modLedger.ReadColumn(lo, COL_MERCHANT)
    descriptions = modLedger.ReadColumn(lo, COL_DESC)
    amounts = modLedger.ReadColumn(lo, COL_AMOUNT)
    accounts = modLedger.ReadColumn(lo, COL_ACCOUNT)
    categories = modLedger.ReadColumn(lo, COL_CATEGORY)
    tags = modLedger.ReadColumn(lo, COL_TAGGEDBY)
    owners = modLedger.ReadColumn(lo, COL_OWNER)

    For i = 1 To rowCount
        If Len(modUtil.NzStr(amounts(i, 1))) > 0 Then
            If Retaggable(modUtil.NzStr(categories(i, 1)), _
                          modUtil.NzStr(tags(i, 1)), includeTagged) Then
                ruleIndex = FirstMatch(rulePattern, ruleField, ruleTest, ruleFlow, _
                                       ruleMin, ruleMax, ruleCount, _
                                       modUtil.NzStr(merchants(i, 1)), _
                                       modUtil.NzStr(descriptions(i, 1)), _
                                       modUtil.NzStr(accounts(i, 1)), _
                                       modUtil.NzNum(amounts(i, 1)))
                If ruleIndex > 0 Then
                    categories(i, 1) = ruleCategory(ruleIndex)
                    tags(i, 1) = TAG_RULE
                    ' Whose card was tapped is not the same question as whose
                    ' expense it is, so shared categories move to the Joint
                    ' owner unless the rule names someone specific.
                    If couple Then
                        newOwner = ruleOwner(ruleIndex)
                        If Len(newOwner) = 0 Then
                            newOwner = modUtil.NzStr(modUtil.GetVal(categoryOwners, _
                                       UCase$(ruleCategory(ruleIndex)), ""))
                        End If
                        If Len(newOwner) > 0 Then owners(i, 1) = newOwner
                    End If
                    ruleHits(ruleIndex) = ruleHits(ruleIndex) + 1
                    matched = matched + 1
                    changed = True
                End If
            End If
        End If
    Next i

    If changed Then
        modLedger.WriteColumn lo, 1, rowCount, COL_CATEGORY, categories
        modLedger.WriteColumn lo, 1, rowCount, COL_TAGGEDBY, tags
        modLedger.WriteColumn lo, 1, rowCount, COL_OWNER, owners
        SaveHits ruleHits, ruleCount
    End If
    modUtil.FastMode False

    If interactive Then
        MsgBox matched & " transaction(s) were categorised." & vbCrLf & _
               CountUncategorized() & " still need a category.", vbInformation, APP_NAME
    End If
    Exit Sub

Fail:
    modUtil.ReportError "Categorize"
End Sub

Private Function Retaggable(ByVal category As String, ByVal taggedBy As String, _
                            ByVal includeTagged As Boolean) As Boolean
    If Len(category) = 0 Then
        Retaggable = True
    ElseIf StrComp(category, CAT_UNCATEGORIZED, vbTextCompare) = 0 Then
        Retaggable = True
    ElseIf includeTagged Then
        Retaggable = (StrComp(taggedBy, TAG_MANUAL, vbTextCompare) <> 0)
    End If
End Function

Private Function LoadRules(ByRef patterns() As String, ByRef fields() As String, _
                           ByRef tests() As String, ByRef categories() As String, _
                           ByRef owners() As String, ByRef flows() As String, _
                           ByRef minimums() As Double, ByRef maximums() As Double, _
                           ByRef hits() As Long) As Long
    Dim lo As ListObject
    Dim rowCount As Long, i As Long, kept As Long
    Dim order() As Long
    Dim priorities() As Double

    Set lo = RulesTable()
    rowCount = modUtil.BodyRows(lo)
    If rowCount = 0 Then Exit Function

    ReDim patterns(1 To rowCount)
    ReDim fields(1 To rowCount)
    ReDim tests(1 To rowCount)
    ReDim categories(1 To rowCount)
    ReDim owners(1 To rowCount)
    ReDim flows(1 To rowCount)
    ReDim minimums(1 To rowCount)
    ReDim maximums(1 To rowCount)
    ReDim hits(1 To rowCount)
    ReDim order(1 To rowCount)
    ReDim priorities(1 To rowCount)

    For i = 1 To rowCount
        If IsRuleUsable(lo, i) Then
            kept = kept + 1
            patterns(kept) = UCase$(RuleValue(lo, i, RL_PATTERN))
            fields(kept) = RuleValue(lo, i, RL_FIELD)
            tests(kept) = RuleValue(lo, i, RL_TEST)
            categories(kept) = RuleValue(lo, i, RL_CATEGORY)
            owners(kept) = RuleValue(lo, i, RL_OWNER)
            flows(kept) = RuleValue(lo, i, RL_FLOW)
            minimums(kept) = modUtil.NzNum(RuleValue(lo, i, RL_MIN), -1E+15)
            maximums(kept) = modUtil.NzNum(RuleValue(lo, i, RL_MAX), 1E+15)
            hits(kept) = 0
            order(kept) = i
            priorities(kept) = modUtil.NzNum(RuleValue(lo, i, RL_PRIORITY), 500)
        End If
    Next i

    If kept = 0 Then Exit Function
    SortRules patterns, fields, tests, categories, owners, flows, minimums, _
              maximums, order, priorities, kept
    mRuleRowMap = order
    LoadRules = kept
End Function

Private Function IsRuleUsable(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    If Len(RuleValue(lo, rowIndex, RL_PATTERN)) = 0 Then Exit Function
    If Len(RuleValue(lo, rowIndex, RL_CATEGORY)) = 0 Then Exit Function
    If StrComp(modUtil.NzStr(RuleValue(lo, rowIndex, RL_ENABLED), "Yes"), _
               "No", vbTextCompare) = 0 Then Exit Function
    IsRuleUsable = True
End Function

Private Function RuleValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                           ByVal header As String) As String
    RuleValue = modUtil.NzStr(lo.DataBodyRange.Cells(rowIndex, _
                modUtil.ColumnIndex(lo, header)).Value)
End Function

' Insertion sort by priority; rule lists are short and this keeps ties stable.
Private Sub SortRules(ByRef patterns() As String, ByRef fields() As String, _
                      ByRef tests() As String, ByRef categories() As String, _
                      ByRef owners() As String, ByRef flows() As String, _
                      ByRef minimums() As Double, ByRef maximums() As Double, _
                      ByRef order() As Long, ByRef priorities() As Double, _
                      ByVal count As Long)
    Dim i As Long, j As Long
    Dim keyPriority As Double
    Dim keyPattern As String, keyField As String, keyTest As String
    Dim keyCategory As String, keyOwner As String, keyFlow As String
    Dim keyMin As Double, keyMax As Double
    Dim keyOrder As Long

    For i = 2 To count
        keyPriority = priorities(i)
        keyPattern = patterns(i): keyField = fields(i): keyTest = tests(i)
        keyCategory = categories(i): keyOwner = owners(i): keyFlow = flows(i)
        keyMin = minimums(i): keyMax = maximums(i): keyOrder = order(i)
        j = i - 1
        Do While j >= 1
            If priorities(j) <= keyPriority Then Exit Do
            priorities(j + 1) = priorities(j)
            patterns(j + 1) = patterns(j): fields(j + 1) = fields(j)
            tests(j + 1) = tests(j): categories(j + 1) = categories(j)
            owners(j + 1) = owners(j): flows(j + 1) = flows(j)
            minimums(j + 1) = minimums(j): maximums(j + 1) = maximums(j)
            order(j + 1) = order(j)
            j = j - 1
        Loop
        priorities(j + 1) = keyPriority
        patterns(j + 1) = keyPattern: fields(j + 1) = keyField
        tests(j + 1) = keyTest: categories(j + 1) = keyCategory
        owners(j + 1) = keyOwner: flows(j + 1) = keyFlow
        minimums(j + 1) = keyMin: maximums(j + 1) = keyMax
        order(j + 1) = keyOrder
    Next i
End Sub

Private Function FirstMatch(ByRef patterns() As String, ByRef fields() As String, _
                            ByRef tests() As String, ByRef flows() As String, _
                            ByRef minimums() As Double, ByRef maximums() As Double, _
                            ByVal ruleCount As Long, ByVal merchant As String, _
                            ByVal description As String, ByVal accountName As String, _
                            ByVal amount As Double) As Long
    Dim i As Long
    Dim target As String

    For i = 1 To ruleCount
        Select Case UCase$(fields(i))
            Case "DESCRIPTION": target = UCase$(description)
            Case "ACCOUNT": target = UCase$(accountName)
            Case "ANY", "": target = UCase$(merchant & " " & description)
            Case Else: target = UCase$(merchant)
        End Select

        If FlowAllows(flows(i), amount) Then
            If Abs(amount) >= minimums(i) - 0.0000001 And _
               Abs(amount) <= maximums(i) + 0.0000001 Then
                If TextMatches(target, patterns(i), tests(i)) Then
                    FirstMatch = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function FlowAllows(ByVal flow As String, ByVal amount As Double) As Boolean
    Select Case UCase$(Trim$(flow))
        Case "MONEY IN", "IN", "CREDIT": FlowAllows = (amount > 0)
        Case "MONEY OUT", "OUT", "DEBIT": FlowAllows = (amount < 0)
        Case Else: FlowAllows = True
    End Select
End Function

Private Function TextMatches(ByVal target As String, ByVal pattern As String, _
                             ByVal test As String) As Boolean
    Select Case UCase$(Trim$(test))
        Case "STARTS WITH": TextMatches = (Left$(target, Len(pattern)) = pattern)
        Case "ENDS WITH": TextMatches = (Right$(target, Len(pattern)) = pattern)
        Case "EQUALS": TextMatches = (target = pattern)
        Case "LIKE": TextMatches = (target Like pattern)
        Case Else: TextMatches = (InStr(target, pattern) > 0)
    End Select
End Function

Private Sub SaveHits(ByRef hits() As Long, ByVal ruleCount As Long)
    Dim lo As ListObject
    Dim i As Long, column As Long
    Dim current As Long

    Set lo = RulesTable()
    column = modUtil.ColumnIndex(lo, RL_HITS)
    For i = 1 To ruleCount
        If hits(i) > 0 Then
            current = CLng(modUtil.NzNum(lo.DataBodyRange.Cells(mRuleRowMap(i), column).Value))
            lo.DataBodyRange.Cells(mRuleRowMap(i), column).Value = current + hits(i)
        End If
    Next i
End Sub

Public Function CountUncategorized() As Long
    Dim lo As ListObject
    Dim categories As Variant, amounts As Variant
    Dim i As Long, total As Long
    Dim category As String

    Set lo = modUtil.TxnTable()
    If modUtil.BodyRows(lo) = 0 Then Exit Function
    categories = modLedger.ReadColumn(lo, COL_CATEGORY)
    amounts = modLedger.ReadColumn(lo, COL_AMOUNT)
    For i = 1 To UBound(categories, 1)
        If Len(modUtil.NzStr(amounts(i, 1))) > 0 Then
            category = modUtil.NzStr(categories(i, 1))
            If Len(category) = 0 Or _
               StrComp(category, CAT_UNCATEGORIZED, vbTextCompare) = 0 Then
                total = total + 1
            End If
        End If
    Next i
    CountUncategorized = total
End Function

'--- Teaching a new rule ----------------------------------------------------

Public Sub TeachRuleFromSelection()
    Dim lo As ListObject
    Dim cell As Range
    Dim rowIndex As Long
    Dim merchant As String, pattern As String, category As String
    Dim rules As ListObject
    Dim newRow As ListRow

    On Error GoTo Fail
    Set lo = modUtil.TxnTable()
    Set cell = Application.ActiveCell

    If Application.Intersect(cell, lo.DataBodyRange) Is Nothing Then
        MsgBox "Select a transaction row first, then use this button.", _
               vbInformation, APP_NAME
        Exit Sub
    End If

    rowIndex = cell.Row - lo.DataBodyRange.Row + 1
    merchant = modUtil.NzStr(modUtil.CellIn(lo, rowIndex, COL_MERCHANT).Value)
    category = modUtil.NzStr(modUtil.CellIn(lo, rowIndex, COL_CATEGORY).Value)

    pattern = Trim$(InputBox( _
        "Create a rule from this transaction." & vbCrLf & vbCrLf & _
        "Description: " & modUtil.NzStr(modUtil.CellIn(lo, rowIndex, COL_DESC).Value) & _
        vbCrLf & vbCrLf & "Text to look for in the merchant name:", APP_NAME, merchant))
    If Len(pattern) = 0 Then Exit Sub

    category = Trim$(InputBox("Which category should that get?" & vbCrLf & _
        "(Type it exactly as it appears on the Categories sheet.)", APP_NAME, category))
    If Len(category) = 0 Then Exit Sub
    If Not CategoryExists(category) Then
        If MsgBox("""" & category & """ is not on the Categories sheet." & vbCrLf & _
                  "Add the rule anyway?", vbYesNo + vbQuestion, APP_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    Set rules = RulesTable()
    Set newRow = rules.ListRows.Add(1)
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_PRIORITY)).Value = 10
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_ENABLED)).Value = "Yes"
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_FIELD)).Value = "Merchant"
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_TEST)).Value = "Contains"
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_PATTERN)).Value = pattern
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_FLOW)).Value = "Any"
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_CATEGORY)).Value = category
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_HITS)).Value = 0
    newRow.Range.Cells(1, modUtil.ColumnIndex(rules, RL_NOTES)).Value = _
        "Learned " & Format$(Date, "yyyy-mm-dd")

    ' Apply the new rule to the row in front of the user right away.
    modUtil.CellIn(lo, rowIndex, COL_CATEGORY).Value = category
    modUtil.CellIn(lo, rowIndex, COL_TAGGEDBY).Value = TAG_RULE

    CategorizeUncategorized False
    Application.Calculate
    MsgBox "Rule added: merchant contains """ & pattern & """ -> " & category & ".", _
           vbInformation, APP_NAME
    Exit Sub

Fail:
    modUtil.ReportError "TeachRuleFromSelection"
End Sub

' category (upper case) -> the owner it belongs to by default, for the
' categories that name one.
Private Function CategoryOwners() As Collection
    Dim lo As ListObject
    Dim map As Collection
    Dim i As Long
    Dim nameColumn As Long, ownerColumn As Long
    Dim category As String, owner As String

    Set map = New Collection
    Set CategoryOwners = map

    On Error Resume Next
    Set lo = modUtil.Tbl(SH_CATEGORIES, TBL_CATEGORIES)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If modUtil.BodyRows(lo) = 0 Then Exit Function

    On Error Resume Next
    ownerColumn = modUtil.ColumnIndex(lo, CAT_OWNER)
    On Error GoTo 0
    If ownerColumn = 0 Then Exit Function
    nameColumn = modUtil.ColumnIndex(lo, CAT_NAME)

    For i = 1 To modUtil.BodyRows(lo)
        category = modUtil.NzStr(lo.DataBodyRange.Cells(i, nameColumn).Value)
        owner = modUtil.NzStr(lo.DataBodyRange.Cells(i, ownerColumn).Value)
        If Len(category) > 0 And Len(owner) > 0 Then
            modUtil.PutVal map, UCase$(category), owner
        End If
    Next i
End Function

Public Function CategoryExists(ByVal category As String) As Boolean
    Dim lo As ListObject
    Dim i As Long
    Set lo = modUtil.Tbl(SH_CATEGORIES, TBL_CATEGORIES)
    For i = 1 To modUtil.BodyRows(lo)
        If StrComp(modUtil.NzStr(lo.DataBodyRange.Cells(i, 1).Value), category, _
                   vbTextCompare) = 0 Then
            CategoryExists = True
            Exit Function
        End If
    Next i
End Function
