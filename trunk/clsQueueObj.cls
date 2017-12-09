VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ObjID    As Double
Private m_Message  As String
Private m_Priority As Integer
Private m_Response As String
Private m_Tag      As String

Public Property Let ID(ByVal i As Double)

    If (m_ObjID > 0) Then
        Exit Property
    End If
    
    m_ObjID = i
    
End Property

Public Property Get ID() As Double

    ID = m_ObjID
    
End Property

Public Property Let Message(ByVal Msg As String)

    m_Message = Msg
    
End Property

Public Property Get Message() As String

    Message = m_Message
    
End Property

Public Property Let Priority(ByVal pri As Integer)

    m_Priority = pri
    
End Property

Public Property Get Priority() As Integer

    Priority = m_Priority
    
End Property

Public Property Let Tag(ByVal Tag As String)

    m_Tag = Tag

End Property

Public Property Get Tag() As String

    Tag = m_Tag

End Property

Public Property Let ResponseTo(ByVal sUser As String)

    m_Response = sUser

End Property

Public Property Get ResponseTo() As String

    ResponseTo = m_Response

End Property


