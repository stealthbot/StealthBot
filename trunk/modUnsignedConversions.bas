Attribute VB_Name = "UnsignedConversions"
' From MSDN Knowledge Base: http://support.microsoft.com/kb/q189323/

Option Explicit

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'Function UnsignedToLong(Value As Double) As Long
'  If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
'  If Value <= MAXINT_4 Then
'    UnsignedToLong = Value
'  Else
'    UnsignedToLong = Value - OFFSET_4
'  End If
'End Function

Function LongToUnsigned(Value As Long) As Double
  If Value < 0 Then
    LongToUnsigned = Value + OFFSET_4
  Else
    LongToUnsigned = Value
  End If
End Function

'Function UnsignedToInteger(Value As Long) As Integer
'  If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
'  If Value <= MAXINT_2 Then
'    UnsignedToInteger = Value
'  Else
'    UnsignedToInteger = Value - OFFSET_2
'  End If
'End Function

'Function IntegerToUnsigned(Value As Integer) As Long
'  If Value < 0 Then
'    IntegerToUnsigned = Value + OFFSET_2
'  Else
'    IntegerToUnsigned = Value
'  End If
'End Function
