VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFriend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFriend
' DateTime  : 3/15/2004 18:58
' Author    : Stealth
' Purpose   : stores friendlist data
'---------------------------------------------------------------------------------------

Option Explicit

Private msUsername As String
Private mbyStatus As Byte
Private mbyLocation As Byte
Private msProduct As String
Private msChannel As String

Public Property Get Channel() As String

    Channel = msChannel

End Property

Public Property Let Channel(ByVal sChannel As String)

    msChannel = sChannel

End Property

Public Property Get Product() As String

    Product = msProduct

End Property

Public Property Let Product(ByVal sProduct As String)

    msProduct = KillNull(sProduct)

End Property

Public Property Get Location() As Byte

    Location = mbyLocation

End Property

Public Property Let Location(ByVal byLocation As Byte)

    mbyLocation = byLocation

End Property

Public Property Get Status() As Byte

    Status = mbyStatus

End Property

Public Property Let Status(ByVal byStatus As Byte)

    mbyStatus = byStatus

End Property

Public Property Get Username() As String

    Username = msUsername

End Property

Public Property Let Username(ByVal sUsername As String)

    msUsername = KillNull(sUsername)

End Property
