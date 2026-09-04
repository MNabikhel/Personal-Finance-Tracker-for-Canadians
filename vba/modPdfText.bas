Attribute VB_Name = "modPdfText"
Option Explicit

'== Statement text to transactions ===========================================
' A PDF statement is a printed page, so once its text has been recovered
' (modPdf does that) what is left is a list of lines.  This module turns those
' lines into transactions without knowing which bank printed them:
'
'   MAR 03  MAR 04  TIM HORTONS #3324 TORONTO ON        5.80
'   03 Mar          Payroll Deposit NORTHWIND    2,483.18   5,120.33
'
' A transaction line starts with a date, may repeat it as a posting date,
' carries its description in the middle and ends with money: the amount, or
' the amount followed by the running balance.  Everything else on the page -
' addresses, totals, notices, column headings - fails that shape and is
' ignored.  Nothing here touches Excel, so it all runs in the test harness.
'=============================================================================

Public Const KIND_CARD As String = "Credit card"
Public Const KIND_ACCOUNT As String = "Bank account"

' Money in a running balance that moves by exactly the amount printed beside
' it, allowing for the cents.
Private Const BALANCE_TOLERANCE As Double = 0.011

'--- Lines ------------------------------------------------------------------

' The text as trimmed, non-empty lines, whichever line ending it arrived with.
Public Function SplitLines(ByVal text As String) As Collection
    Dim out As Collection
    Dim parts() As String
    Dim i As Long
    Dim line As String

    Set out = New Collection
    Set SplitLines = out
    text = Replace$(Replace$(text, vbCrLf, vbLf), vbCr, vbLf)
    parts = Split(text, vbLf)
    For i = LBound(parts) To UBound(parts)
        line = modUtil.CondenseSpaces(parts(i))
        If Len(line) > 0 Then out.Add line
    Next i
End Function

'--- The statement's own date ----------------------------------------------

' Transaction lines rarely carry a year, so it is taken from the statement:
' the latest full date printed anywhere on it (statement date, period end,
' payment due date - all of them sit at or after the transactions).  Returns
' the year, and the month through anchorMonth; 0 when no full date was found.
Public Function StatementAnchor(ByVal lines As Collection, ByRef anchorMonth As Long) As Long
    Dim i As Long, t As Long
    Dim tokens As Variant
    Dim found As Date, best As Date
    Dim used As Long

    anchorMonth = 0
    For i = 1 To lines.Count
        tokens = Split(lines.Item(i), " ")
        For t = LBound(tokens) To UBound(tokens)
            found = FullDateAt(tokens, t, used)
            If used > 0 Then
                If found > best Then best = found
                t = t + used - 1
            End If
        Next t
    Next i

    If best > 0 Then
        StatementAnchor = Year(best)
        anchorMonth = Month(best)
    End If
End Function

' A date with a year in it, starting at tokens(start): "March 19, 2026",
' "19 March 2026", "2026-03-19", "03/19/2026", "March 2026".  used says how
' many tokens it took, 0 when there is none.
Private Function FullDateAt(ByRef tokens As Variant, ByVal start As Long, _
                            ByRef used As Long) As Date
    Dim m As Long, d As Long, y As Long
    Dim ok As Boolean
    Dim token As String

    used = 0
    token = Strip(CStr(tokens(start)))

    m = modParse.MonthFromName(token)
    If m > 0 Then
        If start + 2 <= UBound(tokens) Then
            d = DayNumber(CStr(tokens(start + 1)))
            y = YearNumber(CStr(tokens(start + 2)))
            If d > 0 And y > 0 Then
                FullDateAt = SafeDate(y, m, d, ok)
                If ok Then used = 3
            End If
        End If
        ' "March 2026" on its own is not taken: a card's expiry month would
        ' outrank every real date on the page.
        Exit Function
    End If

    d = DayNumber(token)
    If d > 0 And start + 2 <= UBound(tokens) Then
        m = modParse.MonthFromName(CStr(tokens(start + 1)))
        y = YearNumber(CStr(tokens(start + 2)))
        If m > 0 And y > 0 Then
            FullDateAt = SafeDate(y, m, d, ok)
            If ok Then used = 3
            Exit Function
        End If
    End If

    If IsNumericDate(token) Then
        FullDateAt = modParse.ParseDate(token, "AUTO", ok)
        If ok Then
            If Year(FullDateAt) >= 2000 And Year(FullDateAt) <= 2100 And _
               InStr(token, Right$(CStr(Year(FullDateAt)), 4)) > 0 Then used = 1
        End If
    End If
End Function

'--- Kind of statement -------------------------------------------------------

' Whether the text reads as a credit card statement or a bank account
' statement.  The words on the page settle it when they can; otherwise the
' shape of the lines does: an account statement carries a running balance
' that moves by each amount, and a card statement does not.
Public Function DetectKind(ByVal lines As Collection, ByVal anchorYear As Long, _
                           ByVal anchorMonth As Long) As String
    Dim text As String
    Dim i As Long
    Dim readCount As Long, badCount As Long, consistent As Long

    For i = 1 To lines.Count
        text = text & LCase$(lines.Item(i)) & vbLf
    Next i

    If InStr(text, "credit limit") > 0 Or InStr(text, "minimum payment") > 0 Or _
       InStr(text, "payment due") > 0 Or InStr(text, "paiement minimum") > 0 Then
        DetectKind = KIND_CARD
        Exit Function
    End If
    If InStr(text, "withdrawal") > 0 Or InStr(text, "deposits") > 0 Or _
       InStr(text, "opening balance") > 0 Or InStr(text, "closing balance") > 0 Or _
       InStr(text, "retraits") > 0 Then
        DetectKind = KIND_ACCOUNT
        Exit Function
    End If

    ReadLines lines, KIND_ACCOUNT, anchorYear, anchorMonth, readCount, badCount, consistent
    If consistent >= 2 Then
        DetectKind = KIND_ACCOUNT
    Else
        DetectKind = KIND_CARD
    End If
End Function

'--- Reading ----------------------------------------------------------------

' Every transaction on the statement, as clsTxn records without an account.
' readCount is the number of lines that began with a date, badCount those of
' them that carried money but could not be read as a transaction.
Public Function ReadStatement(ByVal lines As Collection, ByVal kind As String, _
                              ByVal anchorYear As Long, ByVal anchorMonth As Long, _
                              ByRef readCount As Long, ByRef badCount As Long) As Collection
    Dim consistent As Long
    Set ReadStatement = ReadLines(lines, kind, anchorYear, anchorMonth, _
                                  readCount, badCount, consistent)
End Function

Private Function ReadLines(ByVal lines As Collection, ByVal kind As String, _
                           ByVal anchorYear As Long, ByVal anchorMonth As Long, _
                           ByRef readCount As Long, ByRef badCount As Long, _
                           ByRef consistent As Long) As Collection
    Dim out As Collection
    Dim i As Long
    Dim txn As clsTxn
    Dim lastDate As Date
    Dim balance As Double
    Dim haveBalance As Boolean
    Dim dated As Boolean, moneyed As Boolean, matched As Boolean

    Set out = New Collection
    Set ReadLines = out
    readCount = 0
    badCount = 0
    consistent = 0

    For i = 1 To lines.Count
        Set txn = ParseLine(lines.Item(i), kind, anchorYear, anchorMonth, lastDate, _
                            balance, haveBalance, dated, moneyed, matched)
        If dated Then readCount = readCount + 1
        If matched Then consistent = consistent + 1
        If txn Is Nothing Then
            If dated And moneyed Then badCount = badCount + 1
        Else
            out.Add txn
        End If
    Next i
End Function

' One line.  lastDate, balance and haveBalance carry from line to line: the
' date because some banks print it once for a day's transactions, the balance
' because it is what tells a withdrawal from a deposit.  dated and moneyed
' report what the line had, matched whether its balance moved by its amount.
Private Function ParseLine(ByVal line As String, ByVal kind As String, _
                           ByVal anchorYear As Long, ByVal anchorMonth As Long, _
                           ByRef lastDate As Date, ByRef balance As Double, _
                           ByRef haveBalance As Boolean, ByRef dated As Boolean, _
                           ByRef moneyed As Boolean, ByRef matched As Boolean) As clsTxn
    Dim tokens As Variant
    Dim first As Long, last As Long
    Dim used As Long
    Dim when As Date
    Dim amounts() As Double
    Dim moneyCount As Long
    Dim creditMark As Boolean
    Dim description As String
    Dim amount As Double, newBalance As Double
    Dim txn As clsTxn
    Dim ok As Boolean

    dated = False
    moneyed = False
    matched = False

    line = modUtil.CondenseSpaces(line)
    If Len(line) = 0 Then Exit Function
    tokens = Split(line, " ")
    first = LBound(tokens)
    last = UBound(tokens)

    ' The date, and a posting date after it if the bank prints one.  Some
    ' banks put a reference number before the date; it is stepped over.
    when = TxnDateAt(tokens, first, anchorYear, anchorMonth, used)
    If used = 0 And first < last Then
        If IsReferenceNumber(CStr(tokens(first))) Then
            when = TxnDateAt(tokens, first + 1, anchorYear, anchorMonth, used)
            If used > 0 Then first = first + 1
        End If
    End If
    If used > 0 Then
        dated = True
        first = first + used
        TxnDateAt tokens, first, anchorYear, anchorMonth, used
        If used > 0 Then first = first + used
    End If

    ' The money at the end of the line, reading backwards.
    ReDim amounts(1 To 3)
    Do While last >= first And moneyCount < 3
        If IsCreditMark(CStr(tokens(last))) Then
            If moneyCount = 0 Then creditMark = True
            last = last - 1
        ElseIf tokens(last) = "$" Then
            last = last - 1
        ElseIf IsMoneyToken(CStr(tokens(last))) Then
            moneyCount = moneyCount + 1
            amounts(moneyCount) = modParse.ParseAmount(CStr(tokens(last)), ok)
            last = last - 1
            ' "1 234,56": a French statement's thousands, split off by the
            ' space.  Only the comma-decimal shape is joined back up.
            If last >= first And InStr(tokens(last + 1), ",") = Len(tokens(last + 1)) - 2 Then
                If IsThousands(CStr(tokens(last))) Then
                    amounts(moneyCount) = modParse.ParseAmount( _
                        tokens(last) & tokens(last + 1), ok)
                    last = last - 1
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    moneyed = (moneyCount > 0)
    If Not moneyed Then Exit Function

    description = JoinTokens(tokens, first, last)
    ' A line the bank filled in for us rather than a transaction: balances
    ' carried forward, totals, limits.  A balance line still seeds the running
    ' balance, which is what lets the first transaction after it be signed.
    If IsSummary(description) Then
        If IsBalanceLine(description) And moneyCount >= 1 Then
            balance = amounts(1)
            haveBalance = True
        End If
        Exit Function
    End If

    If Not dated Then
        ' Only a bank statement with a running balance may leave the date off
        ' a line, and then the balance has to prove the line is a transaction.
        If StrComp(kind, KIND_ACCOUNT, vbTextCompare) <> 0 Then Exit Function
        If lastDate = 0 Or moneyCount < 2 Or Not haveBalance Then Exit Function
        If Abs(Abs(amounts(1) - balance) - Abs(amounts(2))) > BALANCE_TOLERANCE Then Exit Function
        when = lastDate
    End If
    If Len(description) = 0 Then Exit Function

    If StrComp(kind, KIND_ACCOUNT, vbTextCompare) = 0 Then
        If moneyCount >= 2 Then
            ' amounts(1) is the balance after the line; amounts(2) the amount.
            newBalance = amounts(1)
            amount = Abs(amounts(2))
            If haveBalance And Abs(Abs(newBalance - balance) - amount) <= BALANCE_TOLERANCE Then
                If newBalance < balance Then amount = -amount
                matched = True
            ElseIf Not LooksLikeMoneyIn(description) Then
                amount = -amount
            End If
            balance = newBalance
            haveBalance = True
        Else
            amount = Abs(amounts(1))
            If Not LooksLikeMoneyIn(description) Then amount = -amount
        End If
    Else
        ' On a card statement a plain figure is a charge and a credit is
        ' marked, with a minus sign or a CR after it. If an account-shaped line
        ' is being tried as a card, the earlier figure is the transaction and
        ' the final one is its running balance.
        amount = amounts(moneyCount)
        If creditMark Or amount < 0 Then
            amount = Abs(amount)
        Else
            amount = -Abs(amount)
        End If
    End If
    If amount = 0 Then Exit Function

    lastDate = when
    Set txn = New clsTxn
    txn.TxnDate = when
    txn.Description = description
    txn.Merchant = modRules.CleanMerchant(description)
    txn.Amount = amount
    Set ParseLine = txn
End Function

'--- Dates on transaction lines ---------------------------------------------

' The date a transaction line opens with: "MAR 03", "Mar. 3", "03 Mar",
' "03/03", "2026-03-03", "03/03/2026", "Mar 3, 2026".  A date printed without
' a year is put in the statement's year, or the year before when the month
' would otherwise fall after the statement (a December purchase on a January
' statement).  used is the number of tokens taken, 0 when there is no date.
Public Function TxnDateAt(ByRef tokens As Variant, ByVal start As Long, _
                          ByVal anchorYear As Long, ByVal anchorMonth As Long, _
                          ByRef used As Long) As Date
    Dim m As Long, d As Long, y As Long
    Dim token As String
    Dim ok As Boolean
    Dim parts As Variant

    used = 0
    If start > UBound(tokens) Then Exit Function
    token = Strip(CStr(tokens(start)))

    m = modParse.MonthFromName(token)
    If m > 0 Then
        If start + 1 > UBound(tokens) Then Exit Function
        d = DayNumber(CStr(tokens(start + 1)))
        If d = 0 Then Exit Function
        used = 2
        If start + 2 <= UBound(tokens) Then
            y = YearNumber(CStr(tokens(start + 2)))
            If y > 0 Then used = 3
        End If
    Else
        d = DayNumber(token)
        If d > 0 And start + 1 <= UBound(tokens) Then
            m = modParse.MonthFromName(CStr(tokens(start + 1)))
            If m = 0 Then Exit Function
            used = 2
            If start + 2 <= UBound(tokens) Then
                y = YearNumber(CStr(tokens(start + 2)))
                If y > 0 Then used = 3
            End If
        ElseIf IsNumericDate(token) Then
            parts = Split(Replace$(Replace$(token, "-", "/"), ".", "/"), "/")
            If UBound(parts) = 2 Then
                TxnDateAt = modParse.ParseDate(token, "AUTO", ok)
                If ok Then used = 1
                Exit Function
            ElseIf UBound(parts) = 1 And InStr(token, "/") > 0 Then
                ' Month and day only, as 03/03.  Month first unless the first
                ' part cannot be a month.  (1.25 is money, not a date.)
                m = DayNumber(CStr(parts(0)))
                d = DayNumber(CStr(parts(1)))
                If m > 12 And d <= 12 Then
                    y = m: m = d: d = y: y = 0
                End If
                If m = 0 Or d = 0 Or m > 12 Then Exit Function
                used = 1
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    If y = 0 Then
        y = anchorYear
        If y = 0 Then y = Year(Date)
        If anchorMonth > 0 And m > anchorMonth + 1 Then y = y - 1
    End If

    TxnDateAt = SafeDate(y, m, d, ok)
    If Not ok Then used = 0
End Function

'--- Token shapes -----------------------------------------------------------

' 1,234.56  $1,234.56  -1,234.56  1,234.56-  (1,234.56)  1234,56 - with the
' cents always present, which is what keeps reference numbers, dates and
' store numbers from reading as money.
Public Function IsMoneyToken(ByVal token As String) As Boolean
    Dim body As String
    Dim i As Long, ch As String
    Dim digits As Long, separators As Long

    body = Trim$(token)
    If Right$(body, 1) = "%" Then Exit Function
    If Left$(body, 1) = "(" And Right$(body, 1) = ")" Then body = Mid$(body, 2, Len(body) - 2)
    If Left$(body, 1) = "-" Or Left$(body, 1) = "+" Then body = Mid$(body, 2)
    If Left$(body, 1) = "$" Then body = Mid$(body, 2)
    If Left$(body, 1) = "-" Then body = Mid$(body, 2)
    If Right$(body, 1) = "-" Then body = Left$(body, Len(body) - 1)
    If Right$(body, 1) = "$" Then body = Left$(body, Len(body) - 1)
    If Len(body) < 4 Then Exit Function

    ' Exactly two digits after the last separator.
    ch = Mid$(body, Len(body) - 2, 1)
    If ch <> "." And ch <> "," Then Exit Function
    For i = 1 To Len(body)
        ch = Mid$(body, i, 1)
        If ch Like "#" Then
            digits = digits + 1
        ElseIf ch = "." Or ch = "," Then
            separators = separators + 1
            If i = 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    IsMoneyToken = (digits >= 3 And separators >= 1)
End Function

' A run of digits with no separator: the reference number some card
' statements print ahead of the date.  Too long to be a day, too short to be
' an account number.
Private Function IsReferenceNumber(ByVal token As String) As Boolean
    Dim i As Long
    If Len(token) < 3 Or Len(token) > 6 Then Exit Function
    For i = 1 To Len(token)
        If Not (Mid$(token, i, 1) Like "#") Then Exit Function
    Next i
    IsReferenceNumber = True
End Function

' The thousands of a French-formatted amount, standing alone: 1 to 3 digits.
Private Function IsThousands(ByVal token As String) As Boolean
    Dim i As Long
    If Len(token) < 1 Or Len(token) > 3 Then Exit Function
    For i = 1 To Len(token)
        If Not (Mid$(token, i, 1) Like "#") Then Exit Function
    Next i
    IsThousands = True
End Function

' CR after an amount marks a credit on most Canadian card statements.
Private Function IsCreditMark(ByVal token As String) As Boolean
    Select Case UCase$(Strip(token))
        Case "CR", "-CR", "CR-": IsCreditMark = True
    End Select
End Function

Private Function IsNumericDate(ByVal token As String) As Boolean
    Dim i As Long, ch As String, digits As Long, marks As Long
    For i = 1 To Len(token)
        ch = Mid$(token, i, 1)
        If ch Like "#" Then
            digits = digits + 1
        ElseIf ch = "/" Or ch = "-" Or ch = "." Then
            marks = marks + 1
        Else
            Exit Function
        End If
    Next i
    IsNumericDate = (digits >= 3 And marks >= 1 And marks <= 2)
End Function

Private Function DayNumber(ByVal token As String) As Long
    Dim cleaned As String
    cleaned = Strip(token)
    If Len(cleaned) = 0 Or Len(cleaned) > 2 Then Exit Function
    If Not (cleaned Like "#") And Not (cleaned Like "##") Then Exit Function
    DayNumber = CLng(Val(cleaned))
    If DayNumber > 31 Then DayNumber = 0
End Function

Private Function YearNumber(ByVal token As String) As Long
    Dim cleaned As String
    cleaned = Strip(token)
    If Not (cleaned Like "####") Then Exit Function
    YearNumber = CLng(Val(cleaned))
    If YearNumber < 2000 Or YearNumber > 2100 Then YearNumber = 0
End Function

' Trailing punctuation a date token is often printed with.
Private Function Strip(ByVal token As String) As String
    Strip = Trim$(token)
    Do While Len(Strip) > 0
        Select Case Right$(Strip, 1)
            Case ",", ".", ":", ";"
                Strip = Left$(Strip, Len(Strip) - 1)
            Case Else
                Exit Do
        End Select
    Loop
End Function

Private Function SafeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long, _
                          ByRef ok As Boolean) As Date
    ok = False
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Or y < 1900 Or y > 2200 Then Exit Function
    SafeDate = DateSerial(y, m, d)
    ok = (Day(SafeDate) = d)
End Function

Private Function JoinTokens(ByRef tokens As Variant, ByVal first As Long, _
                            ByVal last As Long) As String
    Dim i As Long, out As String
    For i = first To last
        If Len(out) > 0 Then out = out & " "
        out = out & tokens(i)
    Next i
    JoinTokens = out
End Function

'--- Words ------------------------------------------------------------------

' Lines that carry money but are the statement talking, not a transaction.
Public Function IsSummary(ByVal description As String) As Boolean
    Dim lower As String
    Dim starts As Variant
    Dim i As Long
    lower = LCase$(Trim$(description))
    If Len(lower) = 0 Then
        IsSummary = True
        Exit Function
    End If
    If IsBalanceLine(lower) Then
        IsSummary = True
        Exit Function
    End If
    starts = Array("minimum payment", "payment due", "credit limit", "available credit", _
                   "total", "sub-total", "subtotal", "statement", "interest rate", _
                   "annual interest", "amount due", "paiement minimum", _
                   "limite de credit", "limite de crédit")
    For i = LBound(starts) To UBound(starts)
        If Left$(lower, Len(starts(i))) = starts(i) Then
            IsSummary = True
            Exit Function
        End If
    Next i
End Function

' A line that states the balance rather than changing it.  On a bank
' statement it is also where the running balance starts from.
Public Function IsBalanceLine(ByVal description As String) As Boolean
    Dim lower As String
    Dim starts As Variant
    Dim i As Long
    lower = LCase$(Trim$(description))
    starts = Array("previous balance", "previous statement balance", "new balance", _
                   "opening balance", "closing balance", "balance forward", _
                   "balance brought forward", "balance carried forward", _
                   "beginning balance", "ending balance", "solde")
    For i = LBound(starts) To UBound(starts)
        If Left$(lower, Len(starts(i))) = starts(i) Then
            IsBalanceLine = True
            Exit Function
        End If
    Next i
End Function

' For an account statement with no usable running balance: whether the
' wording says money arrived.  Wrong now and then, which is what the preview
' and the ledger's editable Amount are for.
Public Function LooksLikeMoneyIn(ByVal description As String) As Boolean
    Dim lower As String
    Dim moneyOut As Variant
    Dim words As Variant
    Dim i As Long
    lower = " " & LCase$(description) & " "

    ' These contain words that otherwise look credit-like.  Apple Pay is a
    ' purchase, and overdraft/debit interest is a charge.
    moneyOut = Array(" apple pay", " google pay", " interest charge", _
                     " overdraft interest", " debit interest", " interest debit")
    For i = LBound(moneyOut) To UBound(moneyOut)
        If InStr(lower, moneyOut(i)) > 0 Then Exit Function
    Next i

    words = Array(" deposit", " payroll", " refund", " credit memo", _
                  " interest", " rebate", " transfer from", " received", " reversal", _
                  " cashback", " cash back", " canada fed", " canada pro", _
                  " canada rit", " canada child", " cra ", " gst", " hst", _
                  " child benefit", " dividend", " depot", " dépôt", " remboursement", _
                  " paie ")
    For i = LBound(words) To UBound(words)
        If InStr(lower, words(i)) > 0 Then
            LooksLikeMoneyIn = True
            Exit Function
        End If
    Next i
End Function

'--- Preview -----------------------------------------------------------------

' The first few transactions, for the user to check before they are imported.
Public Function PreviewText(ByVal records As Collection, ByVal limit As Long) As String
    Dim i As Long
    Dim txn As clsTxn
    Dim out As String
    For i = 1 To records.Count
        If i > limit Then
            out = out & "... and " & (records.Count - limit) & " more" & vbCrLf
            Exit For
        End If
        Set txn = records.Item(i)
        out = out & Format$(txn.TxnDate, "yyyy-mm-dd") & "  " & _
              Format$(txn.Amount, "0.00") & "  " & _
              Left$(txn.Description, 40) & vbCrLf
    Next i
    PreviewText = out
End Function
