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
Private m_lMaxCredits As Long
Private m_lCostPerPacket As Long
Private m_lCostPerByte As Long
Private m_lCostPerByteOverThreshhold As Long
Private m_lStartingCredits As Long
Private m_lThreshholdBytes As Long
Private m_lCreditRate As Long

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
    
    byteCost = m_lCostPerByte
    
    If Credits < m_lMaxCredits Then
        Credits = Credits + ((GetTickCount() - LastSent) / CreditRate)
        
        If Credits > m_lMaxCredits Then
            Credits = m_lMaxCredits
        End If
    End If
    
    LastSent = GetTickCount()
    '        int thisByteDelay = byteCost;
    ThisByteDelay = byteCost
    
    If (Len(sText) > m_lThreshholdBytes) Then
        byteCost = Me.CostPerByteOverThreshhold
    End If
    
    ThisPacketCost = Me.CostPerPacket + (ThisByteDelay * Len(sText))
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
Private Sub Class_Initialize()
    ' Set defaults for public variables
    Call Me.FullReset
End Sub

Public Sub FullReset()
    LastSent = GetTickCount()
    Me.StartingCredits = 750
    Me.ThreshholdBytes = 65
    Me.CostPerByte = 12
    Me.CostPerByteOverThreshhold = 15
    Me.CostPerPacket = 190
    Me.MaxCredits = 750
    Me.CreditRate = 8
End Sub


'#########################################################################################
'#########################################################################################
' PROPERTIES
'#########################################################################################
'#########################################################################################
Public Property Get MaxCredits() As Long

    MaxCredits = m_lMaxCredits

End Property

Public Property Let MaxCredits(ByVal lMaxCredits As Long)

    m_lMaxCredits = lMaxCredits

End Property

Public Property Get CostPerPacket() As Long

    CostPerPacket = m_lCostPerPacket

End Property

Public Property Let CostPerPacket(ByVal lCostPerPacket As Long)

    m_lCostPerPacket = lCostPerPacket

End Property

Public Property Get CostPerByte() As Long

    CostPerByte = m_lCostPerByte

End Property

Public Property Let CostPerByte(ByVal lCostPerByte As Long)

    m_lCostPerByte = lCostPerByte

End Property

Public Property Get CostPerByteOverThreshhold() As Long

    CostPerByteOverThreshhold = m_lCostPerByteOverThreshhold

End Property

Public Property Let CostPerByteOverThreshhold(ByVal lCostPerByteOverThreshhold As Long)

    m_lCostPerByteOverThreshhold = lCostPerByteOverThreshhold

End Property

Public Property Get StartingCredits() As Long

    StartingCredits = m_lStartingCredits

End Property

Public Property Let StartingCredits(ByVal lStartingCredits As Long)

    m_lStartingCredits = lStartingCredits

End Property

Public Property Get ThreshholdBytes() As Long

    ThreshholdBytes = m_lThreshholdBytes

End Property

Public Property Let ThreshholdBytes(ByVal lThreshholdBytes As Long)

    m_lThreshholdBytes = lThreshholdBytes

End Property

Public Property Get CreditRate() As Long

    CreditRate = m_lCreditRate

End Property

Public Property Let CreditRate(ByVal lCreditRate As Long)

    m_lCreditRate = lCreditRate

End Property
'#########################################################################################
'#########################################################################################

