VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mVersion As Long
Private mHashCmd As String
Private mKey1Hash As String
Private mKey2Hash As String
Private mGTC As Long
Private mChecksum As Long
Private mEXEInfo As String
Private mServerToken As Long
Private mW3ServerSignature As String * 128
Private mMPQ As Long
Private mVerByte As Long
Private m_lLogonType As Long
Private m_lNLSHandle As Long

Public Sub List()
    With frmChat
        .AddChat vbRed, "Version: " & mVersion
        .AddChat vbRed, "HashCmd: " & mHashCmd
        .AddChat vbRed, "Key1Hash: " & DebugOutput(mKey1Hash)
        .AddChat vbRed, "Key1HashLen: " & Len(mKey1Hash)
        .AddChat vbRed, "Key2Hash: " & DebugOutput(mKey2Hash)
        .AddChat vbRed, "Key2HashLen: " & Len(mKey2Hash)
        .AddChat vbRed, "GTC: " & mGTC
        .AddChat vbRed, "CheckSum: " & mChecksum
        .AddChat vbRed, "EXEInfo: " & mEXEInfo
        .AddChat vbRed, "ServerToken: " & mServerToken
        .AddChat vbRed, "ServerSig: " & DebugOutput(mW3ServerSignature)
    End With
End Sub

Public Sub SetMPQRev(ByVal l As Long)
    mMPQ = l
End Sub

Public Function GetMPQRev() As Long
    GetMPQRev = mMPQ
End Function

Public Sub SetVerByte(ByVal l As Long)
    mVerByte = l
End Sub

Public Function GetVerByte() As Long
    GetVerByte = mVerByte
End Function

Public Sub SetVersion(ByVal Version As Long)
    mVersion = Version
End Sub

Public Function GetVersion() As Long
    GetVersion = mVersion
End Function

Public Sub SetKeyHash(ByVal KeyHash As String, ByVal Key As Byte)
    Select Case Key
        Case 1: mKey1Hash = KeyHash 'KillNull(KeyHash)
        Case 2: mKey2Hash = KeyHash 'KillNull(KeyHash)
    End Select
End Sub

Public Function GetKeyHash(ByVal Key As Byte) As String
    Select Case Key
        Case 1: GetKeyHash = mKey1Hash
        Case 2: GetKeyHash = mKey2Hash
    End Select
End Function

Public Sub SetHashCmd(ByVal HashCmd As String)
    mHashCmd = HashCmd
End Sub

Public Function GetHashCmd() As String
    GetHashCmd = mHashCmd
End Function

Public Sub SetGTC(ByVal GTC As Long)
    mGTC = GTC
End Sub

Public Function GetGTC() As Long
    GetGTC = mGTC
End Function

Public Sub SetCheckSum(ByVal Checksum As Long)
    mChecksum = Checksum
End Sub

Public Function GetChecksum() As Long
    GetChecksum = mChecksum
End Function

Public Sub SetExeInfo(ByVal EXEInfo As String)
    mEXEInfo = EXEInfo
End Sub

Public Function getExeInfo() As String
    getExeInfo = mEXEInfo
End Function

Public Sub SetServerToken(ByVal ServToken As Long)
    mServerToken = ServToken
End Sub

Public Function GetServerToken() As Long
    GetServerToken = mServerToken
End Function

Public Sub SetW3SS(ByVal W3SS As String)
    mW3ServerSignature = W3SS
End Sub

Public Function GetW3SS() As String
    GetW3SS = mW3ServerSignature
End Function


Public Property Get LogonType() As Long
    LogonType = m_lLogonType
End Property

Public Property Let LogonType(ByVal lLogonType As Long)
    m_lLogonType = lLogonType
End Property

Public Property Get NLSHandle() As Long
    NLSHandle = m_lNLSHandle
End Property

Public Property Let NLSHandle(ByVal lNLSHandle As Long)
    m_lNLSHandle = lNLSHandle
End Property
