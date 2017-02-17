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

Public Function GetErrorString(ByVal lErrNum As Long, ByVal IsProxy As Boolean, ByVal source As enuPL_ServerTypes) As String
    Dim sServerType As String
    
    Select Case (source)
        Case stBNCS: sServerType = "Battle.net"
        Case stBNLS: sServerType = "BNLS"
        Case stMCP:  sServerType = "Diablo II Realm"
    End Select
    
    If IsProxy Then sServerType = "proxy"
    
    ExReconTicks = 0
    ExReconMinutes = 0
    
    Select Case (lErrNum)
        Case 10053, 10054
            Count10053 = (Count10053 + 1)
            
            If (Count10053 = 1) Or IsProxy Then
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
                
                ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, 1000, _
                    AddressOf ExtendedReconnect_TimerProc)
    
                UserCancelledConnect = False
                
                miNoProceed = 1
            End If
        
        Case 11004, 11001
            Count11004 = Count11004 + 1
            
            If Count11004 = 1 Or IsProxy Then
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
            
            GetErrorString = "The " & sServerType & " server took too long to respond to your computer. "
            
            If IsProxy Then
                GetErrorString = GetErrorString & "Please check your proxy configuration."
            Else
                Select Case (source)
                    Case stBNCS
                        GetErrorString = GetErrorString & "Try choosing a different server in the " & _
                            "Settings dialog. If you are connecting to a gateway address, such as " & _
                                "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                    Case stBNLS
                        GetErrorString = GetErrorString & "The server appears to be unreachable " & _
                            "at this time. Please check back in an hour or two, select a different BNLS " & _
                                "server, or configure local hashing. (For more information regarding local " & _
                                    "hashing, visit http://www.stealthbot.net.)"
                    Case stMCP
                        GetErrorString = GetErrorString & "The server is not responding. Please " & _
                            "try connecting again in a couple hours, or disabling Diablo II Realms."
                End Select
            End If
    
        Case 10061, 10065
            miNoProceed = 0
            
            GetErrorString = "The " & sServerType & " server you're connecting to is currently unavailable. "
            
            If IsProxy Then
                GetErrorString = GetErrorString & "Please check your proxy configuration."
            Else
                Select Case source
                    Case stBNCS
                        GetErrorString = GetErrorString & "Try choosing a different server in the " & _
                            "Settings dialog. If you are connecting to a gateway address, such as " & _
                                "ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                    Case stBNLS
                        GetErrorString = GetErrorString & "The server appears to be unavailable at " & _
                            "this time. The bot will keep trying to connect to it; if you continue to " & _
                                "receive this error message, wait an hour or so and try again."
                    Case stMCP
                        GetErrorString = GetErrorString & "The server is down."
                        
                        miNoProceed = 2
                End Select
            End If
            
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
