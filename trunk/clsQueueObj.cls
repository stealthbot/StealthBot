VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueOBj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_message  As String
Private m_priority As Integer
Private m_tag      As String

Public Property Let Message(ByVal Msg As String)
    m_message = Msg
End Property

Public Property Get Message() As String
    Message = m_message
End Property

Public Property Let Priority(ByVal pri As Integer)
    m_priority = pri
End Property

Public Property Get Priority() As Integer
    Priority = m_priority
End Property

Public Property Let Tag(ByVal Tag As String)
    m_tag = Tag
End Property

Public Property Get Tag() As String
    Tag = m_tag
End Property
