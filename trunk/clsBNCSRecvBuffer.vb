Option Strict Off
Option Explicit On
Friend Class clsBNCSRecvBuffer
	' cBNCSBufferfer.cls
	' cuphead@valhallalegends.com
	'  Includes debugging code by Stealth (WriteLog etc)
	'  Includes a modified FullPacket sub also by Stealth that checks for a valid
	'   Battle.net packet at the front of the buffer (11/8/06)
	
	Private strData As String
	
	Public Sub AddData(ByVal Data As String)
		'    If InStr(Data, Chr(&HFF)) <> 1 Then
		'        If LenB(strData) > 0 Then
		'            strData = strData & Data
		'        Else
		'            If MDebug("showdrops") Then
		'                frmChat.AddChat vbRed, "Packet data dropped due to bad formatting:"
		'                frmChat.AddChat vbRed, DebugOutput(Data)
		'            End If
		'        End If
		'    Else
		strData = strData & Data
		'    End If
		'WriteLog "Data received. Added to buffer:", True
		'WriteLog Data
	End Sub
	
	'Public Sub WriteLog(ByVal s As String, Optional ByVal NoDebug As Boolean = False)
	'    If Dir$(App.Path & "\packetlog.txt") = "" Then
	'        Open App.Path & "\packetlog.txt" For Output As #1
	'        Close #1
	'    End If
	'
	'    Open App.Path & "\packetlog.txt" For Append As #1
	'        If NoDebug Then
	'            Print #1, s
	'        Else
	'            Print #1, DebugOutput(s) & vbCrLf
	'        End If
	'    Close #1
	'End Sub
	
	Public Function GetBuffer() As String
		GetBuffer = strData
	End Function
	
	Public Function FullPacket() As Boolean
		Dim lngPacketLen, L As Integer
		
		FullPacket = False
		
		If Len(strData) > 0 Then
			L = InStr(strData, Chr(&HFF))
			
			If L = 1 Then
				lngPacketLen = StringToWord(Mid(strData, 3, 2))
				
				If (lngPacketLen = 0) Then
					Exit Function
				End If
				
				If (Len(strData) >= lngPacketLen) Then
					If lngPacketLen < 10000 Then
						FullPacket = True
					Else
						frmChat.AddChat(RTBColors.ErrorMessageText, "Error: Packet Length of unusually high Length detected! Packet " & "dropped. Buffer content at this time: " & vbCrLf & DebugOutput(strData))
						
						Call ClearBuffer()
					End If
				End If
			Else
				frmChat.AddChat(RTBColors.ErrorMessageText, "Error: The front of the buffer is not a valid packet!")
				
				If MDebug("showdrops") Then
					frmChat.AddChat(RTBColors.ErrorMessageText, "Error: The front of the buffer is not " & "a valid packet!")
					frmChat.AddChat(RTBColors.ErrorMessageText, "The following data is being purged:")
					
					If L > 0 Then
						frmChat.AddChat(Space(1) & DebugOutput(Mid(strData, 1, L - 1)))
					Else
						frmChat.AddChat(Space(1) & DebugOutput(strData))
					End If
				End If
				
				If L > 0 Then
					strData = Mid(strData, L)
				Else
					strData = ""
				End If
			End If
		End If
	End Function
	
	Public Function GetPacket() As String
		Dim lngPacketLen As Integer
		
		lngPacketLen = StringToWord(Mid(strData, 3, 2))
		
		'WriteLog "Pulling a packet. Length: " & lngPacketLen, True
		
		If lngPacketLen >= 0 Then
			'frmChat.AddChat RTBColors.ErrorMessageText, "-> Warning: Invalid [low] packet Length specified! Flushing the BNCS receive buffer"
			GetPacket = Mid(strData, 1, lngPacketLen)
			
			'WriteLog GetPacket
			strData = Mid(strData, lngPacketLen + 1)
			'WriteLog "Buffer now contains: ", True
			'WriteLog GetBuffer
		End If
	End Function
	
	Public Sub ClearBuffer()
		strData = vbNullString
	End Sub
	
	Public Sub VoidTrimBuffer()
		Dim LastPacketStart As Integer
		
		LastPacketStart = InStrRev(strData, Chr(&HFF) & Chr(&HF),  , CompareMethod.Binary)
		
		If LastPacketStart > 0 Then
			strData = Mid(strData, LastPacketStart)
		End If
	End Sub
	
	'Private Function StringToWord(Data As String) As Long
	'    Dim tmp As String
	'    Dim A As String, b As String
	'
	'    tmp = ToHex(Data)
	'    A = Mid(tmp, 3, 2)
	'    b = Mid(tmp, 1, 2)
	'    tmp = A & b
	'    StringToWord = Val("&H" & tmp)
	'End Function
	
	Private Function ToHex(ByRef Data As String) As String
		Dim i As Short
		
		For i = 1 To Len(Data)
			ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
		Next i
	End Function
	
	
	'Implementation in _DataArrival:
	'    Dim strData As String
	'    sckBNCS.GetData strData
	'
	'    BNCSData.AddData strData
	
	'    While BNCSData.FullPacket = True
	'        Call ParsingRoutine(BNCSData.GetPacket)
	'    Wend
End Class