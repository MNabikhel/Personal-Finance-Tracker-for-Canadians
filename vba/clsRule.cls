VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'== One row of the Rules sheet ===============================================
' "When this text shows up in that field, and the money is going that way, call
' it this category."  Rules are tried in Priority order and the first one that
' claims a transaction wins.
'=============================================================================

Public RowIndex As Long
Public Priority As Double
Public LookIn As String
Public Test As String
Public Pattern As String
Public MinAmount As Double
Public MaxAmount As Double
Public Flow As String
Public Category As String
Public SetOwner As String

Public Function Matches(ByVal merchant As String, ByVal description As String, _
                        ByVal accountName As String, ByVal amount As Double) As Boolean
    If Not FlowAllows(amount) Then Exit Function
    ' A tolerance, because the limits are money typed in by a person.
    If Abs(amount) < MinAmount - 0.0000001 Then Exit Function
    If Abs(amount) > MaxAmount + 0.0000001 Then Exit Function
    Matches = TextMatches(Target(merchant, description, accountName))
End Function

' A rule that names a merchant and expects money to go to it - the kind a
' refund from that merchant should follow back into the same category.
Public Function IsMerchantPaymentRule() As Boolean
    Select Case UCase$(Trim$(LookIn))
        Case "DESCRIPTION", "ACCOUNT", "ANY", "": Exit Function
    End Select
    Select Case UCase$(Trim$(Flow))
        Case "MONEY OUT", "OUT", "DEBIT": IsMerchantPaymentRule = True
    End Select
End Function

Private Function Target(ByVal merchant As String, ByVal description As String, _
                        ByVal accountName As String) As String
    Select Case UCase$(Trim$(LookIn))
        Case "DESCRIPTION": Target = UCase$(description)
        Case "ACCOUNT": Target = UCase$(accountName)
        Case "ANY", "": Target = UCase$(merchant & " " & description)
        Case Else: Target = UCase$(merchant)
    End Select
End Function

Private Function FlowAllows(ByVal amount As Double) As Boolean
    Select Case UCase$(Trim$(Flow))
        Case "MONEY IN", "IN", "CREDIT": FlowAllows = (amount > 0)
        Case "MONEY OUT", "OUT", "DEBIT": FlowAllows = (amount < 0)
        Case Else: FlowAllows = True
    End Select
End Function

Private Function TextMatches(ByVal target As String) As Boolean
    Dim needle As String
    needle = UCase$(Pattern)
    Select Case UCase$(Trim$(Test))
        Case "STARTS WITH": TextMatches = (Left$(target, Len(needle)) = needle)
        Case "ENDS WITH": TextMatches = (Right$(target, Len(needle)) = needle)
        Case "EQUALS": TextMatches = (target = needle)
        Case "LIKE": TextMatches = (target Like needle)
        Case "WORD": TextMatches = ContainsWord(target, needle)
        Case Else: TextMatches = (InStr(target, needle) > 0)
    End Select
End Function

' "Contains", but only where the pattern stands as a word of its own.  Short
' patterns are otherwise swallowed by longer names that happen to contain them:
' MOBIL is a gas station and also the tail of FREEDOM MOBILE, MEC is a co-op
' and also the start of MECHANIC.
Private Function ContainsWord(ByVal target As String, ByVal needle As String) As Boolean
    Dim at As Long

    If Len(needle) = 0 Then Exit Function
    at = InStr(target, needle)
    Do While at > 0
        If Not IsWordChar(CharBefore(target, at)) Then
            If Not IsWordChar(Mid$(target, at + Len(needle), 1)) Then
                ContainsWord = True
                Exit Function
            End If
        End If
        at = InStr(at + 1, target, needle)
    Loop
End Function

Private Function CharBefore(ByVal target As String, ByVal at As Long) As String
    If at > 1 Then CharBefore = Mid$(target, at - 1, 1)
End Function

Private Function IsWordChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsWordChar = (ch Like "[0-9A-Za-z]")
End Function
