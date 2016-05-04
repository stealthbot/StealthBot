Option Strict Off
Option Explicit On
Friend Class clsBNLSRecvBuffer
	' cBNCSBufferfer.cls
	' cuphead@valhallalegends.com
	
	Private strData As String
	
	Public Sub AddData(ByRef Data As String)
		strData = strData & Data
	End Sub
	
	Public Function FullPacket() As Boolean
		Dim lngPacketLen As Integer
		
		FullPacket = False
		
		If (Len(strData) > 0) Then
			lngPacketLen = StringToWord(Mid(strData, 1, 2))
			
			If (Len(strData) >= lngPacketLen) Then
				FullPacket = True
			End If
		End If
	End Function
	
	Public Function GetPacket() As String
		Dim lngPacketLen As Integer
		
		lngPacketLen = StringToWord(Mid(strData, 1, 2))
		GetPacket = Mid(strData, 1, lngPacketLen)
		strData = Mid(strData, lngPacketLen + 1)
	End Function
	
	Public Sub ClearBuffer()
		strData = vbNullString
	End Sub
	
	'Private Function StringToWord(Data As String) As Long
	'    Dim tmp As String
	'    Dim A As String
	'    Dim b As String
	'
	'    tmp = ToHex(Data)
	'
	'    A = Mid$(tmp, 3, 2)
	'    b = Mid$(tmp, 1, 2)
	'
	'    tmp = A & b
	'
	'    StringToWord = Val("&H" & tmp)
	'End Function
	
	Private Function ToHex(ByRef Data As String) As String
		Dim I As Short
		
		For I = 1 To Len(Data)
			ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, I, 1))), 2)
		Next I
	End Function
End Class