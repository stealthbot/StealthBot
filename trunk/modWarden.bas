Attribute VB_Name = "modWarden"
Option Explicit

Private Enum SHA1Versions
  SHA1 = 0
  BrokenSHA1 = 1
  LockdownSHA1 = 2
  WardenSHA1 = 3
  Max = &HFFFFFFFF
End Enum

Private Type SHA1Context
  IntermediateHash(0 To 4) As Long
  LengthLow As Long
  LengthHigh As Long
  MessageBlockIndex As Integer
  MessageBlock(0 To 63) As Byte
  Computed As Byte
  Corrupted As Byte
  Version As SHA1Versions
End Type

Private Type MD5Context
  IntermediateHash(0 To 3) As Long
  LengthLow As Long
  LengthHigh As Long
  MessageBlockIndex As Integer
  MessageBlock(0 To 63) As Byte
  Computed As Byte
  Corrupted As Byte
End Type

Private Type MedivRandomContext
  Index As Long
  data(0 To 19) As Byte
  Source1(0 To 19) As Byte
  Source2(0 To 19) As Byte
End Type

Private Declare Sub rc4_init Lib "Warden.dll" (ByVal key As String, ByVal Base As String, ByVal length As Long)
Private Declare Sub rc4_crypt Lib "Warden.dll" (ByVal key As String, ByVal data As String, ByVal length As Long)
Private Declare Sub rc4_crypt_data Lib "Warden.dll" (ByVal data As String, ByVal DataLength As Long, ByVal Base As String, ByVal BaseLength As Long)

Private Declare Function sha1_reset Lib "Warden.dll" (ByRef Context As SHA1Context) As Long
Private Declare Function sha1_input Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal data As String, ByVal length As Long) As Long
Private Declare Function sha1_digest Lib "Warden.dll" (ByRef Context As SHA1Context, ByVal digest As String) As Long
Private Declare Function sha1_checksum Lib "Warden.dll" (ByVal data As String, ByVal length As Long, ByVal Version As Long) As Long

Private Declare Function md5_reset Lib "Warden.dll" (ByRef Context As MD5Context) As Long
Private Declare Function md5_input Lib "Warden.dll" (ByRef Context As MD5Context, ByVal data As String, ByVal length As Long) As Long
Private Declare Function md5_digest Lib "Warden.dll" (ByRef Context As MD5Context, ByVal digest As String) As Long
Private Declare Function md5_verify_data Lib "Warden.dll" (ByVal data As String, ByVal length As Long, ByVal CorrectMD5 As String) As Boolean

Private Declare Sub mediv_random_init Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal seed As String, ByVal length As Long)
Private Declare Sub mediv_random_get_bytes Lib "Warden.dll" (ByRef Context As MedivRandomContext, ByVal buffer As String, ByVal length As Long)

Private Declare Function module_prep Lib "Warden.dll" (ByVal Source As String, ByVal Callback As Long) As Long
Private Declare Function module_init Lib "Warden.dll" (ByVal address As Long, ByVal Callbacks As Long) As Long
Private Declare Sub module_init_rc4 Lib "Warden.dll" (ByVal Callback As Long, ByVal InitData As Long, ByVal data As String, ByVal length As Long)
Private Declare Function module_handle_packet Lib "Warden.dll" (ByVal InitData As Long, ByVal data As String, ByVal length As Long) As Long

Private Declare Function warden_init Lib "Warden.dll" (ByVal SocketHandle As Long) As Long
Private Declare Function warden_data Lib "Warden.dll" (ByVal Instance As Long, ByVal Direction As Long, ByVal PacketID As Long, ByVal data As String, ByVal length As Long) As Long
Private Declare Function warden_cleanup Lib "Warden.dll" (ByVal Instance As Long) As Long


Private Const WARDEN_SEND              As Long = &H0
Private Const WARDEN_RECV              As Long = &H1
Private Const WARDEN_BNCS              As Long = &H2

Public Sub WardenCleanup(Instance As Long)
  Call warden_cleanup(Instance)
End Sub

Public Function WardenInitilize(ByVal SocketHandle As Long) As Long
  WardenInitilize = warden_init(SocketHandle)
End Function

Public Function WardenServerData(Instance As Long, sData As String) As Boolean
  Dim ID As Integer
  ID = Asc(Mid(sData, 2, 1))
  WardenServerData = IIf(warden_data(Instance, WARDEN_BNCS Or WARDEN_RECV, ID, Mid$(sData, 5), Len(sData) - 4) = 0, False, True)
End Function

Public Function WardenClientData(Instance As Long, sData As String) As Boolean
  Dim ID As Long
  ID = Asc(Mid(sData, 2, 1))
  WardenClientData = IIf(warden_data(Instance, WARDEN_BNCS Or WARDEN_SEND, ID, Mid$(sData, 5), Len(sData) - 4) = 0, False, True)
End Function

