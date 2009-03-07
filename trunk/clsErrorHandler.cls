VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsErrorHandler
' Created   : 8/22/2004 03:04
' Author    : AndyT (andy@stealthbot.net)
' Purpose   : Advanced error display
'---------------------------------------------------------------------------------------
'
Option Explicit

Private miNoProceed As Integer
Private Count10053  As Integer
Private Count11004  As Integer

Private Sub Class_Initialize()
    miNoProceed = -1
End Sub

Public Function GetErrorString(ByVal lErrNum As Long, ByVal Source As enuErrorSources) As String
    Dim sServerType As String
    
    Select Case (Source)
        Case BNET: sServerType = "Battle.net"
        Case BNLS: sServerType = "BNLS"
        Case MCP:  sServerType = "Realm"
    End Select
    
    ExReconTicks = 0
    ExReconMinutes = 0
    
    Select Case (lErrNum)
        Case 10053, 10054
            Count10053 = (Count10053 + 1)
            
            If (Count10053 = 1) Then
                GetErrorString = "The " & sServerType & _
                    " server has terminated your connection."
                
                miNoProceed = 0
            Else
                Count10053 = 0
                
                GetErrorString = "You appear to be IPBanned. The bot will attempt to " & _
                    "reconnect again in 20 minutes."
                
                If (ExReconnectTimerID) Then
                    Call KillTimer(0, ExReconnectTimerID)
                End If
                
                If (ReconnectTimerID) Then
                    Call KillTimer(0, ReconnectTimerID)
                End If
                
                ExReconMinutes = (20 * 60)
                
                miNoProceed = 1
            End If
        
        Case 11004, 11001
            Count11004 = Count11004 + 1
            
            If Count11004 = 1 Then
                GetErrorString = "Your computer is unable to contact the " & _
                    sServerType & " server."
                    
                miNoProceed = 0
            Else
                GetErrorString = "Your computer is having DNS resolution issues. No more " & _
                    "reconnection will occur. Please try connecting again in 15-30 minutes, or " & _
                        "contact your Internet Service Provider."
                        
                miNoProceed = 2
            End If
            
        Case 10060
            miNoProceed = 0
            
            GetErrorString = "The server took too long to respond to your computer. "
            
            Select Case (Source)
                Case BNET
                    GetErrorString = GetErrorString & "Try choosing a different server in the " & _
                        "Settings dialog. If you are connecting to a gateway address, such as " & _
                            "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                Case BNLS
                    GetErrorString = GetErrorString & "The BNLS server appears to be unreachable " & _
                        "at this time. Please check back in an hour or two, select a different BNLS " & _
                            "server, or configure local hashing. (For more information regarding local " & _
                                "hashing, visit http://www.stealthbot.net.)"
                Case MCP
                    GetErrorString = GetErrorString & "The Realm server is not responding. Please " & _
                        "try connecting again in a couple hours, or disabling Realm logins."
            End Select
    
        Case 10061, 10065
            miNoProceed = 0
            
            GetErrorString = "The server you're connecting to is currently unavailable. "
            
            Select Case Source
                Case BNET
                    GetErrorString = GetErrorString & "Try choosing a different server in the " & _
                        "Settings dialog. If you are connecting to a gateway address, such as " & _
                            "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                Case BNLS
                    GetErrorString = GetErrorString & "The BNLS server appears to be unavailable at " & _
                        "this time. The bot will keep trying to connect to it; if you continue to " & _
                            "receive this error message, wait an hour or so and try again."
                Case MCP
                    GetErrorString = GetErrorString & "The Diablo II Realm server is down."
                    
                    miNoProceed = 2
            End Select
            
        Case Else
            miNoProceed = 0
    End Select
    
    If (miNoProceed > 0) Then
        If (miNoProceed > 1) Then
            UserCancelledConnect = True
        End If
    Else
        ExReconMinutes = (BotVars.ReconnectDelay / 1000)
    
        ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, 1000, _
            AddressOf ExtendedReconnect_TimerProc)
    
        UserCancelledConnect = False
    End If
End Function

Public Sub Reset()
    Count10053 = 0
    Count11004 = 0
    
    miNoProceed = -1
End Sub

Public Function OKToProceed() As Boolean
    OKToProceed = (miNoProceed = 0)
    
    miNoProceed = -1
End Function
