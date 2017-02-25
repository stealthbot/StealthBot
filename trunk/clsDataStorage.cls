VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_lNLSHandle As Long

Private m_ServerToken   As Long
Private m_ClientToken   As Long
Private m_lLogonType    As Long
Private m_UDPValue      As Long
Private m_CRevFileTime  As String
Private m_CRevFileName  As String
Private m_CRevSeed      As String
Private m_CRevVersion   As Long
Private m_CRevChecksum  As Long
Private m_CRevResult    As String
Private m_ServerSig     As String
Private m_EmailRegDelay As Boolean
Private m_NLS           As clsNLS
Private m_MCPHandler    As clsMCPHandler
Private m_FirstTimeChat As Boolean
Private m_AccountEntry  As Boolean

Public Sub List()
    With frmChat
        .AddChat RTBColors.ErrorMessageText, StringFormat("Logon Type:   0x{0}", ZeroOffset(LogonType, 8))
        .AddChat RTBColors.ErrorMessageText, StringFormat("Server Token: 0x{0}", ZeroOffset(ServerToken, 8))
        .AddChat RTBColors.ErrorMessageText, StringFormat("Client Token: 0x{0}", ZeroOffset(ClientToken, 8))
        .AddChat RTBColors.ErrorMessageText, StringFormat("UDP Value:    0x{0}", ZeroOffset(UDPValue, 8))
        .AddChat RTBColors.ErrorMessageText, "CRev Info: "
        .AddChat RTBColors.ErrorMessageText, StringFormat("  FileTime: {0}", CRevFileTime)
        .AddChat RTBColors.ErrorMessageText, StringFormat("  FileName: {0}", CRevFileName)
        .AddChat RTBColors.ErrorMessageText, StringFormat("  Seed:     {0}", IIf(InStr(1, CRevSeed, "A=", vbTextCompare) = 0, StrToHex(CRevSeed), CRevSeed))
        .AddChat RTBColors.ErrorMessageText, StringFormat("  Version:  0x{0}", ZeroOffset(CRevVersion, 8))
        .AddChat RTBColors.ErrorMessageText, StringFormat("  Checksum: 0x{0}", ZeroOffset(CRevChecksum, 8))
        .AddChat RTBColors.ErrorMessageText, StringFormat("  Result:   {0}", IIf(InStr(1, CRevSeed, "A=", vbTextCompare) = 0, StrToHex(CRevResult), CRevResult))
        '.AddChat RTBColors.ErrorMessageText, StringFormat("MCP Data:{0}{1}", vbNewLine, DebugOutput(m_MCPData))
    End With
End Sub

Public Sub Reset()
    m_ServerToken = 0
    m_ClientToken = 0
    m_lLogonType = 0
    m_UDPValue = 0
    m_CRevFileTime = vbNullString
    m_CRevFileName = vbNullString
    m_CRevSeed = vbNullString
    m_CRevVersion = 0
    m_CRevChecksum = 0
    m_CRevResult = vbNullString
    m_ServerSig = vbNullString
    m_EmailRegDelay = False
    Set m_MCPHandler = Nothing
    Set m_NLS = Nothing
    m_FirstTimeChat = False
    m_AccountEntry = False
End Sub

Public Property Get LogonType() As Long
    LogonType = m_lLogonType
End Property
Public Property Let LogonType(ByVal lLogonType As Long)
    m_lLogonType = lLogonType
End Property

Public Property Let ServerToken(lData As Long)
    m_ServerToken = lData
End Property
Public Property Get ServerToken() As Long
    ServerToken = m_ServerToken
End Property

Public Property Let ClientToken(lData As Long)
    m_ClientToken = lData
End Property
Public Property Get ClientToken() As Long
    If (m_ClientToken = 0) Then m_ClientToken = GetTickCount
    ClientToken = m_ClientToken
End Property

Public Property Let UDPValue(lData As Long)
    m_UDPValue = lData
End Property
Public Property Get UDPValue() As Long
    UDPValue = m_ServerToken
End Property

Public Property Let CRevFileTime(sData As String)
    m_CRevFileTime = sData
End Property
Public Property Get CRevFileTimeRaw() As String
    CRevFileTimeRaw = m_CRevFileTime
End Property
Public Property Get CRevFileTime() As String
    Dim ft As FILETIME
    Dim st As SYSTEMTIME
    
    If (Not Len(m_CRevFileTime) = 8) Then
        m_CRevFileTime = Left$(StringFormat("{0}{1}", m_CRevFileTime, String$(8, Chr$(0))), 8)
    End If
    CopyMemory ft, ByVal m_CRevFileTime, 8
    
    FileTimeToSystemTime ft, st
    With st
        CRevFileTime = StringFormat("{0}/{1}/{2} {3}:{4}:{5}", _
            .wMonth, .wDay, .wYear, .wHour, .wMinute, .wSecond)
    End With
End Property

Public Property Let CRevFileName(sData As String)
    m_CRevFileName = sData
End Property
Public Property Get CRevFileName() As String
    CRevFileName = m_CRevFileName
End Property

Public Property Let CRevSeed(sData As String)
    m_CRevSeed = sData
End Property
Public Property Get CRevSeed() As String
    CRevSeed = m_CRevSeed
End Property

Public Property Let CRevResult(sData As String)
    m_CRevResult = sData
End Property
Public Property Get CRevResult() As String
    CRevResult = m_CRevResult
End Property

Public Property Let CRevVersion(lData As Long)
    m_CRevVersion = lData
End Property
Public Property Get CRevVersion() As Long
    CRevVersion = m_CRevVersion
End Property

Public Property Let CRevChecksum(lData As Long)
    m_CRevChecksum = lData
End Property
Public Property Get CRevChecksum() As Long
    CRevChecksum = m_CRevChecksum
End Property

Public Property Let ServerSig(sData As String)
    m_ServerSig = sData
End Property
Public Property Get ServerSig() As String
    ServerSig = m_ServerSig
End Property

Public Property Let WaitingForEmail(bData As Boolean)
    m_EmailRegDelay = bData
End Property
Public Property Get WaitingForEmail() As Boolean
    WaitingForEmail = m_EmailRegDelay
End Property

Public Property Get NLS() As clsNLS
    If (m_NLS Is Nothing) Then Set m_NLS = New clsNLS
    Set NLS = m_NLS
End Property

Public Property Set MCPHandler(oData As clsMCPHandler)
    Set m_MCPHandler = oData
End Property
Public Property Get MCPHandler() As clsMCPHandler
    Set MCPHandler = m_MCPHandler
End Property

Public Property Let EnteredChatFirstTime(bData As Boolean)
    m_FirstTimeChat = bData
End Property
Public Property Get EnteredChatFirstTime() As Boolean
    EnteredChatFirstTime = m_FirstTimeChat
End Property

Public Property Let AccountEntry(bData As Boolean)
    m_AccountEntry = bData
End Property
Public Property Get AccountEntry() As Boolean
    AccountEntry = m_AccountEntry
End Property

Private Sub Class_Terminate()
    Reset
End Sub
