Option Strict Off
Option Explicit On
Friend Class clsErrorHandler
	'---------------------------------------------------------------------------------------
	' Module    : clsErrorHandler
	' Created   : 8/22/2004 03:04
	' Author    : AndyT (andy@stealthbot.net)
	' Purpose   : Advanced error display
	'---------------------------------------------------------------------------------------
	'
	
	Private miNoProceed As Short
	Private Count10053 As Short
	Private Count11004 As Short
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		miNoProceed = -1
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Function GetErrorString(ByVal lErrNum As Integer, ByVal Source As modEnum.enuErrorSources) As String
		Dim sServerType As String
		
		Select Case (Source)
			Case modEnum.enuErrorSources.BNET : sServerType = "Battle.net"
			Case modEnum.enuErrorSources.BNLS : sServerType = "BNLS"
			Case modEnum.enuErrorSources.MCP : sServerType = "Realm"
		End Select
		
		ExReconTicks = 0
		ExReconMinutes = 0
		
		Select Case (lErrNum)
			Case 10053, 10054
				Count10053 = (Count10053 + 1)
				
				If (Count10053 = 1) Then
					GetErrorString = "The " & sServerType & " server has terminated your connection."
					
					miNoProceed = 0
				Else
					Count10053 = 0
					
					GetErrorString = "You appear to be IPBanned. The bot will attempt to " & "reconnect again in 20 minutes."
					
					If (ExReconnectTimerID) Then
						Call KillTimer(0, ExReconnectTimerID)
					End If
					
					If (ReconnectTimerID) Then
						Call KillTimer(0, ReconnectTimerID)
					End If
					
					ExReconMinutes = (20 * 60)
					
                    ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, 1000, AddressOf ExtendedReconnect_TimerProc)
					
					UserCancelledConnect = False
					
					miNoProceed = 1
				End If
				
			Case 11004, 11001
				Count11004 = Count11004 + 1
				
				If Count11004 = 1 Then
					GetErrorString = "Your computer is unable to contact the " & sServerType & " server."
					
					miNoProceed = 0
				Else
					GetErrorString = "Your computer is having DNS resolution issues. No more " & "reconnection will occur. Please try connecting again in 15-30 minutes, or " & "contact your Internet Service Provider."
					
					miNoProceed = 2
				End If
				
			Case 10060
				miNoProceed = 0
				
				GetErrorString = "The server took too long to respond to your computer. "
				
				Select Case (Source)
					Case modEnum.enuErrorSources.BNET
						GetErrorString = GetErrorString & "Try choosing a different server in the " & "Settings dialog. If you are connecting to a gateway address, such as " & "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
					Case modEnum.enuErrorSources.BNLS
						GetErrorString = GetErrorString & "The BNLS server appears to be unreachable " & "at this time. Please check back in an hour or two, select a different BNLS " & "server, or configure local hashing. (For more information regarding local " & "hashing, visit http://www.stealthbot.net.)"
					Case modEnum.enuErrorSources.MCP
						GetErrorString = GetErrorString & "The Realm server is not responding. Please " & "try connecting again in a couple hours, or disabling Realm logins."
				End Select
				
			Case 10061, 10065
				miNoProceed = 0
				
				GetErrorString = "The server you're connecting to is currently unavailable. "
				
				Select Case Source
					Case modEnum.enuErrorSources.BNET
						GetErrorString = GetErrorString & "Try choosing a different server in the " & "Settings dialog. If you are connecting to a gateway address, such as " & "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
					Case modEnum.enuErrorSources.BNLS
						GetErrorString = GetErrorString & "The BNLS server appears to be unavailable at " & "this time. The bot will keep trying to connect to it; if you continue to " & "receive this error message, wait an hour or so and try again."
					Case modEnum.enuErrorSources.MCP
						GetErrorString = GetErrorString & "The Diablo II Realm server is down."
						
						miNoProceed = 2
				End Select
				
			Case Else
				miNoProceed = 0
		End Select
		
		If (miNoProceed > 0) Then
			UserCancelledConnect = False
			
			If (miNoProceed > 1) Then
				UserCancelledConnect = True
			End If
		Else
			ExReconMinutes = (BotVars.ReconnectDelay / 1000)
			
            ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, 1000, AddressOf ExtendedReconnect_TimerProc)
			
			UserCancelledConnect = False
		End If
	End Function
	
	'UPGRADE_NOTE: Reset was upgraded to Reset_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Reset_Renamed()
		Count10053 = 0
		Count11004 = 0
		
		miNoProceed = -1
	End Sub
	
	Public Function OKToProceed() As Boolean
		OKToProceed = (miNoProceed = 0)
		
		miNoProceed = -1
	End Function
End Class