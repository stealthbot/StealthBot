VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'clsUserInfo - project StealthBot

Private msUsername As String
Private mlFlags As Long
Private mlPing As Long
Private msStatstring As String
Private msProduct As String
Private mbyInternalFlags As Byte
Private mlLastAction As Long
Private mbSafelisted As Boolean
Private msClan As String
Private mlJoinTime As Long
Private mbIsSelf As Boolean

Private Sub Class_Initialize()
    mlLastAction = GetTickCount()
    mlJoinTime = mlLastAction
End Sub

Public Property Get Username() As String

    Username = msUsername

End Property

Public Property Let Username(ByVal sUsername As String)

    msUsername = sUsername

End Property

Public Property Get Flags() As Long

    Flags = mlFlags

End Property

Public Property Let Flags(ByVal lFlags As Long)

    mlFlags = lFlags

End Property

Public Property Get Ping() As Long

    Ping = mlPing

End Property

Public Property Let Ping(ByVal lPing As Long)

    mlPing = lPing

End Property

Public Property Get Statstring() As String

    Statstring = msStatstring

End Property

Public Property Let Statstring(ByVal sStatstring As String)

    msStatstring = sStatstring

End Property

Public Property Get Product() As String

    Product = msProduct

End Property

Public Property Let Product(ByVal sProduct As String)

    msProduct = sProduct

End Property

Public Property Get InternalFlags() As Byte

    InternalFlags = mbyInternalFlags

End Property

Public Property Let InternalFlags(ByVal byInternalFlags As Byte)

    mbyInternalFlags = byInternalFlags

End Property

Public Function TimeSinceTalk() As Long

    TimeSinceTalk = (GetTickCount() - mlLastAction) / 1000

End Function

Public Function TimeInChannel() As Long
    
    TimeInChannel = (GetTickCount() - mlJoinTime) / 1000
    
End Function
'
'Public Property Let TimeSinceTalk(ByVal lTimeSinceTalk As Long)
'
'    mlTimeSinceTalk = lTimeSinceTalk
'
'End Property

Public Property Get Safelisted() As Boolean

    Safelisted = mbSafelisted

End Property

Public Property Let Safelisted(ByVal bSafelisted As Boolean)

    mbSafelisted = bSafelisted

End Property

Public Property Get Clan() As String

    Clan = msClan

End Property

Public Property Let Clan(ByVal sClan As String)

    msClan = sClan

End Property

Public Property Get JoinTime() As Long

    JoinTime = mlJoinTime

End Property

Public Property Let JoinTime(ByVal lJoinTime As Long)

    mlJoinTime = lJoinTime
    Acts

End Property

Public Property Get IsSelf() As Boolean

    IsSelf = mbIsSelf

End Property

Public Property Let IsSelf(ByVal bIsSelf As Boolean)

    mbIsSelf = bIsSelf

End Property

Public Sub Acts()
    mlLastAction = GetTickCount()
End Sub
