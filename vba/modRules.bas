Attribute VB_Name = "modRules"
Option Explicit

'== Merchant clean-up and the categorisation rules ==========================
' Bank descriptions are noisy ("IDP PURCHASE - 1234 LOBLAWS #1234 TORONTO ON").
' CleanMerchant boils that down to something readable, and the rules on the
' Rules sheet map it to a category.
'=============================================================================

' Column headers on the sheet.
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
        ElseIf IsStoreNumber(word) Then
            ' store number such as "#1234"
        ElseIf IsNumeric(word) And Len(word) >= 4 Then
            ' reference number
        ElseIf word Like "*#####*" Then
            ' embedded long digit run
        Else
            If Len(kept) > 0 Then kept = kept & " "
            kept = kept & word
        End If
    Next i
    StripReferenceNumbers = kept
End Function

' "#1234", the store number a terminal appends.  Written out rather than as a
' Like pattern because "#" is the digit wildcard, not a literal, in a pattern.
Private Function IsStoreNumber(ByVal word As String) As Boolean
    If Len(word) < 2 Then Exit Function
    If Left$(word, 1) <> "#" Then Exit Function
    IsStoreNumber = IsNumeric(Mid$(word, 2))
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

'--- Matching --------------------------------------------------------------

' The first rule that claims this transaction, or Nothing.  Rules are expected
' in priority order; LoadRules puts them that way.
Public Function FirstMatch(ByVal rules As Collection, ByVal merchant As String, _
                           ByVal description As String, ByVal accountName As String, _
                           ByVal amount As Double) As clsRule
    Dim i As Long

    For i = 1 To rules.Count
        If rules.Item(i).Matches(merchant, description, accountName, amount) Then
            Set FirstMatch = rules.Item(i)
            Exit Function
        End If
    Next i
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
    Dim i As Long, matched As Long
    Dim rules As Collection
    Dim rule As clsRule
    Dim hits As Collection
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

    Set rules = LoadRules()
    If rules.Count = 0 Then
        If interactive Then
            MsgBox "There are no enabled rules on the Rules sheet.", vbInformation, APP_NAME
        End If
        Exit Sub
    End If
    Set hits = New Collection

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
                Set rule = FirstMatch(rules, modUtil.NzStr(merchants(i, 1)), _
                                      modUtil.NzStr(descriptions(i, 1)), _
                                      modUtil.NzStr(accounts(i, 1)), _
                                      modUtil.NzNum(amounts(i, 1)))
                If Not rule Is Nothing Then
                    categories(i, 1) = rule.Category
                    tags(i, 1) = TAG_RULE
                    ' Whose card was tapped is not the same question as whose
                    ' expense it is, so shared categories move to the Joint
                    ' owner unless the rule names someone specific.
                    If couple Then
                        newOwner = rule.SetOwner
                        If Len(newOwner) = 0 Then
                            newOwner = modUtil.NzStr(modUtil.GetVal(categoryOwners, _
                                       UCase$(rule.Category), ""))
                        End If
                        If Len(newOwner) > 0 Then owners(i, 1) = newOwner
                    End If
                    modUtil.BumpVal hits, CStr(rule.RowIndex)
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
        SaveHits hits
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

' Every usable rule on the sheet, in the order they should be tried.
Public Function LoadRules() As Collection
    Dim lo As ListObject
    Dim loaded As Collection
    Dim rule As clsRule
    Dim i As Long

    Set loaded = New Collection
    Set lo = RulesTable()

    For i = 1 To modUtil.BodyRows(lo)
        If IsRuleUsable(lo, i) Then
            Set rule = New clsRule
            rule.RowIndex = i
            rule.Priority = modUtil.NzNum(RuleValue(lo, i, RL_PRIORITY), 500)
            rule.LookIn = RuleValue(lo, i, RL_FIELD)
            rule.Test = RuleValue(lo, i, RL_TEST)
            rule.Pattern = RuleValue(lo, i, RL_PATTERN)
            rule.MinAmount = modUtil.NzNum(RuleValue(lo, i, RL_MIN), -1E+15)
            rule.MaxAmount = modUtil.NzNum(RuleValue(lo, i, RL_MAX), 1E+15)
            rule.Flow = RuleValue(lo, i, RL_FLOW)
            rule.Category = RuleValue(lo, i, RL_CATEGORY)
            rule.SetOwner = RuleValue(lo, i, RL_OWNER)
            loaded.Add rule
        End If
    Next i

    Set LoadRules = ByPriority(loaded)
End Function

' Insertion sort; rule lists are short and this keeps rules of equal priority
' in the order the user put them on the sheet.
Public Function ByPriority(ByVal rules As Collection) As Collection
    Dim ordered() As Object
    Dim key As clsRule
    Dim out As Collection
    Dim i As Long, j As Long

    Set out = New Collection
    Set ByPriority = out
    If rules.Count = 0 Then Exit Function

    ReDim ordered(1 To rules.Count)
    For i = 1 To rules.Count
        Set ordered(i) = rules.Item(i)
    Next i

    For i = 2 To rules.Count
        Set key = ordered(i)
        j = i - 1
        Do While j >= 1
            If ordered(j).Priority <= key.Priority Then Exit Do
            Set ordered(j + 1) = ordered(j)
            j = j - 1
        Loop
        Set ordered(j + 1) = key
    Next i

    For i = 1 To rules.Count
        out.Add ordered(i)
    Next i
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

' hits is keyed by the rule's row on the sheet.
Private Sub SaveHits(ByVal hits As Collection)
    Dim lo As ListObject
    Dim i As Long, column As Long
    Dim added As Long, current As Long

    Set lo = RulesTable()
    column = modUtil.ColumnIndex(lo, RL_HITS)
    For i = 1 To modUtil.BodyRows(lo)
        added = CLng(modUtil.GetVal(hits, CStr(i), 0))
        If added > 0 Then
            current = CLng(modUtil.NzNum(lo.DataBodyRange.Cells(i, column).Value))
            lo.DataBodyRange.Cells(i, column).Value = current + added
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
