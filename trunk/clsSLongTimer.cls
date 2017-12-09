VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSLongTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TmrObj   As Object
Private m_Interval As Integer
Private m_Counter  As Double

Public Property Get tmr() As Object

    Set tmr = m_TmrObj

End Property

Public Property Let tmr(ByRef o As Object)

    Set m_TmrObj = o

End Property

Public Property Get Parent() As Form

    Set Parent = m_TmrObj.Parent

End Property

Public Property Get Index() As Integer

    Index = m_TmrObj.Index

End Property

Public Property Get Name() As String

    Name = m_TmrObj.Name

End Property

Public Property Get Tag() As String

    Tag = m_TmrObj.Tag

End Property

Public Property Let Tag(ByVal str As String)

    m_TmrObj.Tag = str

End Property

Public Property Get Interval() As Integer

    Interval = m_Interval

End Property

Public Property Let Interval(ByVal i As Integer)

    m_Interval = i

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_TmrObj.Enabled

End Property

Public Property Let Enabled(ByVal bln As Boolean)

    m_TmrObj.Enabled = bln

End Property

Public Property Get Counter() As Double

    Counter = m_Counter

End Property

Public Property Let Counter(ByVal i As Double)

    m_Counter = i
    
End Property
