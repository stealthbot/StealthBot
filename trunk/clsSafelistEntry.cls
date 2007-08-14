VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSafelistEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsSafelistEntry
' June 17th, 2006
'  by Andy
'  Written in Iowa on I-80 East

' Container class for safelist entries
'   currently holds a safelisted Name and the person who added them
Private m_sName As String
Private m_sAddedBy As String



Public Property Get Name() As String

    Name = m_sName

End Property

Public Property Let Name(ByVal sName As String)

    m_sName = sName

End Property

Public Property Get AddedBy() As String

    AddedBy = m_sAddedBy

End Property

Public Property Let AddedBy(ByVal sAddedBy As String)

    m_sAddedBy = sAddedBy

End Property
