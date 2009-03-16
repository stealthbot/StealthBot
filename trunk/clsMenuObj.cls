VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd        As Long
Private m_name        As String
Private m_index       As Integer
Private m_caption     As String
Private m_parent      As Object
Private m_enabled     As Boolean
Private m_visible     As Boolean
Private m_help_id     As Long
Private m_checked     As Boolean
Private m_tag         As String
Private m_window_list As Boolean

Private Sub Class_Initialize()

    AppendMenu frmChat.mnuScripts.Parent.hWnd, MF_STRING, 0, ""

End Sub

Public Property Get Caption() As String

    ' ...

End Property

Public Property Let Caption(ByVal str As String)

    ' ...

End Property

Public Property Get Name() As String

    ' ...

End Property

Public Property Get Index() As Integer

    ' ...

End Property

Public Property Get Parent() As Object

    ' ...

End Property

Public Property Get Enabled() As Boolean

    ' ...

End Property

Public Property Let Enabled(ByVal bln As Boolean)

    ' ...

End Property

Public Property Get Checked() As Boolean

    ' ...

End Property

Public Property Let Checked(ByVal bln As Boolean)

    ' ...

End Property

Public Property Get Visible() As Boolean

    ' ...

End Property

Public Property Let Visible(ByVal bln As Boolean)

    ' ...

End Property

Public Property Get HelpContextID() As Long

    ' ...

End Property

Public Property Let HelpContextID(ByVal bln As Long)

    ' ...

End Property

Public Property Get Tag() As String

    ' ...

End Property

Public Property Let Tag(ByVal str As String)

    ' ...

End Property

Public Property Get WindowList() As Boolean

    ' ...

End Property

Public Function CreateSub(ByVal MenuName As String)

    ' ...
    
End Function
