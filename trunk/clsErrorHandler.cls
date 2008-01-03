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
Private Count10053 As Integer
Private Count11004 As Integer

Private Sub Class_Initialize()
    miNoProceed = -1
End Sub

Public Function GetErrorString(ByVal lErrNum As Long, ByVal Source As enuErrorSources) As String
    Dim sServerType As String
    
    Select Case Source
        Case BNET
            sServerType = "Battle.net"
        Case BNLS
            sServerType = "BNLS"
        Case mcp
            sServerType = "Realm"
    End Select
    
    Select Case lErrNum
        Case 10053, 10054
            Count10053 = Count10053 + 1
            
            If Count10053 = 1 Then
                GetErrorString = "The " & sServerType & " server has terminated your connection."
                miNoProceed = 0
            Else
                Count10053 = 0
                GetErrorString = "You appear to be IPBanned. The bot will attempt to reconnect again in 20 minutes."
                
                If ExReconnectTimerID > 0 Then
                    MsgBox "!"
                
                    KillTimer 0, ExReconnectTimerID
                End If
                
                ExReconTicks = 0
                ExReconMinutes = 20
                'ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, _
                '    ((ExReconMinutes * 60) * 1000), AddressOf ExtendedReconnect_TimerProc)
                
                miNoProceed = 1
            End If
        
        Case 11004, 11001
            Count11004 = Count11004 + 1
            
            If Count11004 = 1 Then
                GetErrorString = "Your computer is unable to contact the " & sServerType & " servers."
                miNoProceed = 0
            Else
                GetErrorString = "Your computer is having DNS resolution issues. No more reconnection will occur. Please try connecting again in 15-30 minutes, or contact your Internet Service Provider."
                miNoProceed = 2
            End If
            
        Case 10060
            miNoProceed = 0
            GetErrorString = "The server took too long to respond to your computer. "
            
            Select Case Source
                Case BNET
                    GetErrorString = GetErrorString & "Try choosing a different server in the Settings dialog. If you are connecting to a gateway address, such as ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                Case BNLS
                    GetErrorString = GetErrorString & "The BNLS server appears to be unreachable at this time. Please check back in an hour or two, select a different BNLS server, or configure local hashing. (For more information regarding local hashing, visit http://www.stealthbot.net.)"
                Case mcp
                    GetErrorString = GetErrorString & "The Realm server is not responding. Please try connecting again in a couple hours, or disabling Realm logins."
            End Select
    
        Case 10061, 10065
            miNoProceed = 0
            GetErrorString = "The server you're connecting to is currently unavailable. "
            
            Select Case Source
                Case BNET
                    GetErrorString = GetErrorString & "Try choosing a different server in the Settings dialog. If you are connecting to a gateway address, such as ÿcbuseast.battle.netÿcb, try using one of the IP addresses listed below it."
                Case BNLS
                    GetErrorString = GetErrorString & "The BNLS server appears to be unavailable at this time. The bot will keep trying to connect to it; if you continue to receive this error message, wait an hour or so and try again."
                Case mcp
                    GetErrorString = GetErrorString & "The Diablo II Realm server is down."
                    miNoProceed = 2
            End Select
            
        Case Else
            miNoProceed = 0
            
    End Select
    
    If (miNoProceed) Then
        UserCancelledConnect = True
        'Debug.Print "Set here.."
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
