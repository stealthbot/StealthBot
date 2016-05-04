Option Strict Off
Option Explicit On
Module modUnsignedConversions
	' From MSDN Knowledge Base: http://support.microsoft.com/kb/q189323/
	
	
	Private Const OFFSET_4 As Double = 4294967296#
	Private Const MAXINT_4 As Double = 2147483647
	Private Const OFFSET_2 As Integer = 65536
	Private Const MAXINT_2 As Short = 32767
	
	Function UnsignedToLong(ByRef Value As Double) As Integer
		If Value < 0 Or Value >= OFFSET_4 Then Error(6) ' Overflow
		If Value <= MAXINT_4 Then
			UnsignedToLong = Value
		Else
			UnsignedToLong = Value - OFFSET_4
		End If
	End Function
	
	Function LongToUnsigned(ByRef Value As Integer) As Double
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
End Module