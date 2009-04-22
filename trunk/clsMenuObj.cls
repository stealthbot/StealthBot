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

Private m_hWnd         As Long
Private m_command_id   As Long
Private m_name         As String
Private m_parent       As Object
Private m_enabled      As Boolean
Private m_visible      As Boolean
Private m_help_id      As Long
Private m_checked      As Boolean
Private m_tag          As String
Private m_window_list  As Boolean
Private m_subs         As Collection
Private m_has_children As Boolean

Private Sub Class_Initialize()

    ' ...
    m_enabled = True
    m_visible = True

End Sub

Public Sub Class_Terminate()

    ' ...
    DeleteMenu

End Sub

Public Property Get Name() As String

    Name = m_name

End Property

Public Property Let Name(ByVal str As String)

    m_name = str

End Property

Public Property Get hWnd() As Long

    hWnd = m_hWnd

End Property

Public Property Let hWnd(ByVal lng As Long)

    If (m_hWnd > 0) Then
        DeleteMenu
    End If

    m_hWnd = lng
    
    CreateMenu

End Property

Public Property Get HasChildren() As Boolean

    HasChildren = m_has_children

End Property

Public Property Let HasChildren(ByVal bln As Boolean)

    ' ...
    m_has_children = bln
    
    ' ...
    Enabled = Enabled

End Property

Public Property Get Parent() As Object

    Set Parent = m_parent

End Property

Public Property Let Parent(ByRef obj As Object)

    If (Not (m_parent Is Nothing)) Then
        DeleteMenu
    End If

    Set m_parent = obj
    
    If (TypeOf obj Is Form) Then
        m_hWnd = GetMenu(obj.hWnd)
    ElseIf (TypeOf obj Is clsMenuObj) Then
        m_hWnd = obj.ID
        
        obj.HasChildren = True
    End If
    
    CreateMenu

End Property

Public Property Get ID() As Long

    ID = m_command_id

End Property

Public Property Let ID(ByVal lng As Long)

    If (m_command_id = 0) Then
        m_command_id = lng
    End If

End Property

Public Property Get Caption() As String

    Dim str As String ' ...
    Dim lng As Long   ' ...
    
    str = String$(256, Chr$(0))

    lng = _
        GetMenuString(m_hWnd, m_command_id, str, Len(str), MF_BYCOMMAND)

    Caption = Left$(str, lng)

End Property

Public Property Let Caption(ByVal str As String)

    If (str = "-") Then
        ModifyMenu m_hWnd, m_command_id, MF_BYCOMMAND Or MF_SEPARATOR, m_command_id, str
    Else
        ModifyMenu m_hWnd, m_command_id, MF_BYCOMMAND, m_command_id, str
    End If
    
    DrawMenuBar m_hWnd

End Property

Public Property Get Enabled() As Boolean

    Enabled = m_enabled

End Property

Public Property Let Enabled(ByVal bln As Boolean)

    If (bln = True) Then
        If (ModifyMenu(m_hWnd, _
                m_command_id, MF_BYCOMMAND Or IIf(HasChildren, MF_POPUP, MF_STRING), _
                    m_command_id, Caption()) <> -1) Then
                
            m_enabled = True
        End If
    Else
        If (ModifyMenu(m_hWnd, _
                m_command_id, MF_BYCOMMAND Or MF_GRAYED, m_command_id, Caption()) <> -1) Then
                
            m_enabled = False
        End If
    End If
    
    DrawMenuBar m_hWnd

End Property

Public Property Get Checked() As Boolean

    Checked = m_checked

End Property

Public Property Let Checked(ByVal bln As Boolean)

    If (bln = True) Then
        If (CheckMenuItem(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_CHECKED) <> -1) Then
            m_checked = True
        End If
    Else
        If (CheckMenuItem(m_hWnd, m_command_id, MF_BYCOMMAND Or MF_UNCHECKED) <> -1) Then
            m_checked = False
        End If
    End If
    
    DrawMenuBar m_hWnd

End Property

Public Property Get Visible() As Boolean

    Visible = m_visible

End Property

Public Property Let Visible(ByVal bln As Boolean)

    If (bln = True) Then
        CreateMenu
    Else
        DeleteMenu
    End If

End Property

Public Property Get HelpContextID() As Long

    ' ...

End Property

Public Property Let HelpContextID(ByVal bln As Long)

    ' ...

End Property

Public Property Get Tag() As String

    Tag = m_tag

End Property

Public Property Let Tag(ByVal str As String)

    m_tag = str

End Property

Public Property Get WindowList() As Boolean

    ' ...

End Property

Private Function CreateMenu() As Long

    If (m_command_id = 0) Then
    
        m_command_id = CreatePopupMenu()
    
        AppendMenu m_hWnd, IIf(HasChildren, MF_POPUP, MF_STRING), m_command_id, ""
        
        DrawMenuBar m_hWnd
        
        CreateMenu = m_command_id
        
    End If

End Function

Private Function DeleteMenu() As Long

    If (m_command_id <> 0) Then
    
        RemoveMenu m_hWnd, m_command_id, MF_BYCOMMAND
            
        m_command_id = 0

    End If
    
End Function
