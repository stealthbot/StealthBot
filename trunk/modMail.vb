Option Strict Off
Option Explicit On
Module modMail
	'modMail - project StealthBot - authored 8/3/04 andy@stealthbot.net
	
	Private CurrentOpenFile As Short
	Private CurrentRecord As Integer
	
	Private MailFile As String
	
	Public Sub AddMail(ByRef tsMsg As udtMail)
		Call OpenMailFile()
		
		'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		FilePut(CurrentOpenFile, tsMsg, CurrentRecord + 1)
		
		Call CloseMailFile()
	End Sub
	
	Public Function GetMailCount(ByVal sUser As String) As Integer
		Dim mTemp As udtMail
		Dim i As Integer
		Dim Count As Integer
		
		Call OpenMailFile()
		
		If (CurrentRecord > 0) Then
			For i = 1 To CurrentRecord
				'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileGet(CurrentOpenFile, mTemp, i)
				
				If (StrComp(sUser, RTrim(mTemp.To_Renamed), CompareMethod.Text) = 0) Then
					Count = Count + 1
				End If
			Next i
			
			GetMailCount = Count
		Else
			GetMailCount = 0
		End If
		
		Call CloseMailFile()
	End Function
	
	Public Sub GetMailMessage(ByVal sUser As String, ByRef theMessage As udtMail)
		Dim msgTemp As udtMail
		Dim i As Integer
		
		Call OpenMailFile()
		
		If (CurrentRecord > 0) Then
			For i = 1 To CurrentRecord
				'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				FileGet(CurrentOpenFile, msgTemp, i)
				
				If (StrComp(sUser, RTrim(msgTemp.To_Renamed), CompareMethod.Text) = 0) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object theMessage. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					theMessage = msgTemp
					
					' Trim off the buffer space from the message.
					theMessage.To_Renamed = Trim(theMessage.To_Renamed)
					theMessage.From = Trim(theMessage.From)
					theMessage.Message = Trim(theMessage.Message)
					
					With msgTemp
						.To_Renamed = vbNullString
					End With
					
					'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FilePut(CurrentOpenFile, msgTemp, i)
					
					Exit For
				End If
			Next i
		Else
			With theMessage
				.To_Renamed = vbNullString
				.From = vbNullString
				.Message = vbNullString
			End With
		End If
		
		Call CloseMailFile()
	End Sub
	
	Public Sub OpenMailFile()
		On Error GoTo ERROR_HANDLER
		
		Dim temp As udtMail
		Dim f As Short
		Dim i As Integer
		
		f = FreeFile
		
		MailFile = GetFilePath(FILE_MAILDB)
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If (LenB(Dir(MailFile)) = 0) Then
			FileOpen(f, MailFile, OpenMode.Output)
			FileClose(f)
		End If
		
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		FileOpen(f, MailFile, OpenMode.Random, , , LenB(temp))
		
		If (LOF(f) > 0) Then
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			i = LOF(f) \ LenB(temp)
			
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If (LOF(f) Mod LenB(temp) <> 0) Then
				i = (i + 1)
			End If
		Else
			i = 0
		End If
		
		CurrentRecord = i
		CurrentOpenFile = f
		
		Exit Sub
		
ERROR_HANDLER: 
		Call frmChat.AddChat(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red), "Error: " & Err.Description & " in " & "OpenMailFile().")
		
		Exit Sub
	End Sub
	
	Public Sub CloseMailFile()
		FileClose(CurrentOpenFile)
	End Sub
	
	Public Sub CleanUpMailFile()
		Dim tMail() As udtMail
		Dim tTemp As udtMail
		Dim i As Integer
		Dim c As Integer
		
		Call OpenMailFile()
		
		If (CurrentRecord > 0) Then
			'UPGRADE_WARNING: Lower bound of array tMail was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim tMail(CurrentRecord)
			
			If (LOF(CurrentOpenFile) > 0) Then
				' mail in the mail file
				' collect valid entries and rewrite it
				For i = 1 To CurrentRecord
					'UPGRADE_WARNING: Get was upgraded to FileGet and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FileGet(CurrentOpenFile, tTemp, i)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object tMail(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					tMail(i) = tTemp
				Next i
			End If
			
			Call CloseMailFile()
			
			' Zap the old file
			Call Kill(MailFile)
			
			' Write a new mail file
			Call OpenMailFile()
			
			c = 1
			
			For i = 1 To UBound(tMail)
				If (Len(Trim(tMail(i).To_Renamed)) > 0) Then
					'UPGRADE_WARNING: Put was upgraded to FilePut and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
					FilePut(CurrentOpenFile, tMail(i), c)
					
					c = (c + 1)
				End If
			Next i
		End If
		
		Call CloseMailFile()
	End Sub
End Module