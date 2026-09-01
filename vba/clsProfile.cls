VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'== One row of the "Bank Formats" sheet ======================================
' How to read a particular institution's CSV export.  A profile is lifted off
' the sheet once per import and then does the reading itself, so nothing below
' this point needs to know where the settings came from.
'=============================================================================

Public RowIndex As Long
Public Name As String
Public Institution As String
Public SkipRows As Long
Public DelimiterName As String
Public DateColumn As Long
Public DateFormat As String
Public DescriptionColumns As String
Public AmountMode As String
Public AmountColumn As Long
Public DebitColumn As Long
Public CreditColumn As Long
Public Signature As String

Public Function Delimiter() As String
    Select Case UCase$(Trim$(DelimiterName))
        Case "TAB": Delimiter = vbTab
        Case "SEMICOLON": Delimiter = ";"
        Case "PIPE": Delimiter = "|"
        Case Else: Delimiter = ","
    End Select
End Function

Public Function Title() As String
    Title = Name & " (" & Institution & ")"
End Function

' Reads one record of the file.  Returns Nothing when the row is not a
' transaction at all - a header, a bank notice, a running total - which is how
' those get skipped without the user having to say so.
Public Function ReadRow(ByVal fields As Collection) As clsTxn
    Dim txn As clsTxn
    Dim ok As Boolean
    Dim when As Date
    Dim amount As Double

    when = modParse.ParseDate(modParse.FieldAt(fields, DateColumn), DateFormat, ok)
    If Not ok Then Exit Function

    amount = ReadAmount(fields, ok)
    If Not ok Then Exit Function

    Set txn = New clsTxn
    txn.TxnDate = when
    txn.Description = modParse.FieldsAt(fields, DescriptionColumns)
    txn.Merchant = modRules.CleanMerchant(txn.Description)
    txn.Amount = amount
    Set ReadRow = txn
End Function

' Money out is negative and money in is positive, whichever way the bank chose
' to write it: one signed column, one signed column with the sign inverted
' (most credit cards), or separate debit and credit columns.
Private Function ReadAmount(ByVal fields As Collection, ByRef ok As Boolean) As Double
    Dim debit As Double, credit As Double
    Dim parsed As Boolean

    If StrComp(AmountMode, MODE_DEBIT_CREDIT, vbTextCompare) = 0 Then
        debit = modParse.ParseAmount(modParse.FieldAt(fields, DebitColumn), parsed)
        If Not parsed Then debit = 0
        credit = modParse.ParseAmount(modParse.FieldAt(fields, CreditColumn), parsed)
        If Not parsed Then credit = 0
        ok = (debit <> 0 Or credit <> 0)
        ReadAmount = Abs(credit) - Abs(debit)
    Else
        ReadAmount = modParse.ParseAmount(modParse.FieldAt(fields, AmountColumn), ok)
        If StrComp(AmountMode, MODE_SIGNED_FLIP, vbTextCompare) = 0 Then
            ReadAmount = -ReadAmount
        End If
    End If
End Function
