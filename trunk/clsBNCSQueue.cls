VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBNCSQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2009-07-16 - clsBNCSQueue - StealthBot Chat Queue 2
'  Originally by iago - project JavaOp
'#########################################################################################
' CHANGES:
' 2009-07-16 initial port -andy
' 2009-08-11 added BanDelay(), formerly private in frmChat -andy
'#########################################################################################

' (original)
'/*
' * Created on Dec 14, 2004
' * By iago
' */

' SETTINGS
'        p.put("debug", "This will show the current delay and the current number of credits each message, in case you want to find-tune it.");
'        p.put("prevent flooding", "It's a very bad idea to turn this off -- if you do, it won't try to stop you from flooding.");
'        p.put("cost - packet", "This is the amount of credits 'paid' for each sent packet.");
'        p.put("cost - byte", "WARNING: I don't recommend changing ANY of the settings for anti-flood.  But if you want to tweak, you can.  This is the number of credits 'paid' for each byte.");
'        p.put("cost - byte over threshold", "This is the amount of credits 'paid' for each byte after the threshold is reached.");
'        p.put("starting credits", "This is the number of credits you start with.");
'        p.put("threshold bytes", "This is the length of a message that triggers the treshold (extra delay).");
'        p.put("max credits", "This is the maximum number of credits that the bot can have.");
'        p.put("credit rate", "This is the amount of time (in milliseconds) it takes to earn one credit.");

' PRIVATE VARIABLES
Private LastSent As Long
Private Credits As Long

'    /** Note on implementation: This will assume that all previous packets have already been sent.  Don't call this multiple
'     * time in a row and hope to get a good result! */
'    public long getDelay(String text, Object data)
'
' Returns the amount of time you should wait before sending the next message in the queue
Public Function GetDelay(ByVal sText As String) As Long
    Dim ThisByteDelay As Long, ThisPacketCost As Long
    Dim byteCost As Long, RequiredDelay As Long
    
    byteCost = BotVars.QueueCostPerByte
    
    If Credits < BotVars.QueueMaxCredits Then
        ' Adjust credits up
        Credits = Credits + ((GetTickCount() - LastSent) / BotVars.QueueCreditRate)
        
        If Credits > BotVars.QueueMaxCredits Then
            Credits = BotVars.QueueMaxCredits
        End If
    End If
    
    LastSent = GetTickCount()
    '        int thisByteDelay = byteCost;
    ThisByteDelay = byteCost
    
    If (Len(sText) > BotVars.QueueThreshholdBytes) Then
        byteCost = BotVars.QueueCostPerByteOverThreshhold
    End If
    
    ThisPacketCost = BotVars.QueueCostPerPacket + (ThisByteDelay * Len(sText))
    Debug.Print "Cost for this packet: " & ThisPacketCost
    
    '// Check how long this packet will have to wait
    If (Credits < 0) Then
        RequiredDelay = (-1 * Credits) * CreditRate
    End If
    
    If (ThisPacketCost > Credits) Then
        RequiredDelay = (-1 * (Credits - ThisPacketCost)) * ThisByteDelay
    End If
    
    Credits = Credits - ThisPacketCost
    
    Debug.Print "Remaining credits: " & Credits & "; delay: " & RequiredDelay
    GetDelay = RequiredDelay
End Function

Public Sub ClearQueue()
    Credits = 0
    LastSent = GetTickCount()
    ' TODO: Clear the actual queue
End Sub

'//            out.putLocalSetting(getName(), "cost - packet", "190");
'//            out.putLocalSetting(getName(), "cost - byte", "12");
'//            out.putLocalSetting(getName(), "cost - byte over threshold", "15");
'//            out.putLocalSetting(getName(), "starting credits", "750");
'//            out.putLocalSetting(getName(), "threshold bytes", "65");
'//            out.putLocalSetting(getName(), "max credits", "750");
'//            out.putLocalSetting(getName(), "credit rate", "8");
'Private Sub Class_Initialize()
'    ' Set defaults for public variables
'    BotVars.QueueStartingCredits = DefaultStartingCredits                       '750
'    BotVars.QueueThreshholdBytes = DefaultThreshholdBytes                       '65
'    BotVars.QueueCostPerByte = DefaultCostPerByte                               '12
'    BotVars.QueueCostPerPacket = DefaultCostPerPacket                           '190
'    BotVars.QueueCostPerByteOverThreshhold = DefaultCostPerByteOverThreshhold   '15
'    BotVars.QueueMaxCredits = DefaultMaxCredits                                 '750
'    BotVars.QueueCreditRate = DefaultCreditRate                                 '8
'End Sub

Public Property Get DefaultStartingCredits()
    DefaultStartingCredits = 200
End Property

Public Property Get DefaultThreshholdBytes()
    DefaultThreshholdBytes = 200
End Property

Public Property Get DefaultCostPerByte()
    DefaultCostPerByte = 7
End Property

Public Property Get DefaultCostPerPacket()
    DefaultCostPerPacket = 200
End Property

Public Property Get DefaultCostPerByteOverThreshhold()
    DefaultCostPerByteOverThreshhold = 8
End Property

Public Property Get DefaultMaxCredits()
    DefaultMaxCredits = 600
End Property

Public Property Get DefaultCreditRate()
    DefaultCreditRate = 7
End Property

'#########################################################################################
'#########################################################################################
' PROPERTIES
'#########################################################################################
'#########################################################################################

Public Function BanDelay() As Long
    ' define default error handler
    On Error GoTo ERROR_HANDLER
    
    ' base ban delay
    ' The base delay serves two functions: it prevents likely ineffectual attempts at
    ' banning fast floodbots & it provides a small window for bots without similar ban
    ' delay functions to do banning without incurring the high risk of double bans.
    ' The base delay prevents banning at any lower interval than what is specified.
    BanDelay = 100

    ' do we have ops?
    If (g_Channel.Self.IsOperator) Then
        Dim OpCount As Integer
        Dim j       As Integer
        
        ' loop through users in channel
        For j = 1 To g_Channel.Users.Count
            ' is user an operator?
            If (g_Channel.Users(j).IsOperator) Then
                OpCount = (OpCount + 1)
            End If
        Next j
        
        ' do we have more than one op?
        If (OpCount > 1) Then
            ' seed rnd function
            Randomize
            
            ' set random ban delay based primarily on op count
            BanDelay = _
                (BanDelay + ((1 + Rnd * (OpCount * 2)) * (1 + Rnd * 125)))
        End If
    End If
    
    ' exit procedure
    Exit Function

' default error handler
ERROR_HANDLER:
    ' display error message
    Call frmChat.AddChat(vbRed, "Error: " & Err.description & " in BanDelay().")

    ' exit procedure
    Exit Function

End Function
'#########################################################################################
'#########################################################################################

