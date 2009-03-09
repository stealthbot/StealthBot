VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScriptHighResTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tmrObj   As Object
Private m_interval As Integer
Private m_counter  As Integer

Public Property Get tmr() As Object

    Set tmr = m_tmrObj

End Property

Public Property Let tmr(ByRef o As Object)

    Set m_tmrObj = o

End Property

Public Property Get Parent() As Form

    Set Parent = m_tmrObj.Parent

End Property

Public Property Get Index() As Integer

    Index = m_tmrObj.Index

End Property

Public Property Get Name() As String

    Name = m_tmrObj.Name

End Property

Public Property Get Tag() As String

    Tag = m_tmrObj.Name

End Property

Public Property Let Tag(ByVal str As String)

    m_tmrObj.Tag = str

End Property

Public Property Get Interval() As Integer

    Interval = m_interval

End Property

Public Property Let Interval(ByVal I As Integer)

    m_interval = I

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_tmrObj.Enabled

End Property

Public Property Let Enabled(ByVal bln As Boolean)

    m_tmrObj.Enabled = bln

End Property

Public Property Get Counter() As Integer

    Counter = m_counter

End Property

Public Property Let Counter(ByVal I As Integer)

    m_counter = I
    
End Property
