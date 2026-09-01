VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'== One transaction read out of a bank export ================================
' Everything a statement row carries before it reaches the ledger.  The last
' three fields come from the file it was found in rather than from the row.
'=============================================================================

Public TxnDate As Date
Public Description As String
Public Merchant As String
Public Amount As Double
Public Account As String
Public Owner As String
Public SourceFile As String

' The duplicate-detection key: same account, day, amount and description means
' the same transaction, however many times the statement is downloaded.
Public Function MatchKey() As String
    MatchKey = modUtil.MatchKey(Account, TxnDate, Amount, Description)
End Function
