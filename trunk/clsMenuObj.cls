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
Private m_CommandID   As Long
Private m_Name        As String
Private m_Parent      As Object
Private m_Enabled     As Boolean
Private m_Visible     As Boolean
Private m_HelpID      As Long
Private m_Checked     As Boolean
Private m_Tag         As String
Private m_WindowList  As Boolean
Private m_Subs        As Collection
Private m_HasChildren As Boolean

Private Sub Class_Initialize()
    m_Enabled = True
    m_Visible = True
End Sub

Public Sub Class_Terminate()
    DeleteMenu
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal str As String)
    m_Name = str
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
    HasChildren = m_HasChildren
End Property

Public Property Let HasChildren(ByVal bln As Boolean)
    m_HasChildren = bln
    
    Enabled = Enabled
End Property

Public Property Get Parent() As Object
    Set Parent = m_Parent
End Property

Public Property Let Parent(ByRef obj As Object)
    If (Not (m_Parent Is Nothing)) Then
        DeleteMenu
    End If

    Set m_Parent = obj
    
    If (TypeOf obj Is Form) Then
        m_hWnd = GetMenu(obj.hWnd)
    ElseIf (TypeOf obj Is clsMenuObj) Then
        m_hWnd = obj.ID
        
        obj.HasChildren = True
    End If
    
    CreateMenu
End Property

Public Property Get ID() As Long
    ID = m_CommandID
End Property

Public Property Let ID(ByVal lng As Long)
    If (m_CommandID = 0) Then
        m_CommandID = lng
    End If
End Property

Public Property Get Caption() As String
    Dim str As String
    Dim lng As Long
    
    str = String$(256, Chr$(0))

    lng = GetMenuString(m_hWnd, m_CommandID, str, Len(str), MF_BYCOMMAND)

    Caption = Left$(str, lng)
End Property

Public Property Let Caption(ByVal str As String)
    If (str = "-") Then
        ModifyMenu m_hWnd, m_CommandID, MF_BYCOMMAND Or MF_SEPARATOR, m_CommandID, str
    Else
        ModifyMenu m_hWnd, m_CommandID, MF_BYCOMMAND, m_CommandID, str
    End If
    
    DrawMenuBar m_hWnd
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal bln As Boolean)
    If (bln = True) Then
        If (ModifyMenu(m_hWnd, _
                m_CommandID, MF_BYCOMMAND Or IIf(HasChildren, MF_POPUP, MF_STRING), _
                    m_CommandID, Caption()) <> -1) Then
                
            m_Enabled = True
        End If
    Else
        If (ModifyMenu(m_hWnd, _
                m_CommandID, MF_BYCOMMAND Or MF_GRAYED, m_CommandID, Caption()) <> -1) Then
                
            m_Enabled = False
        End If
    End If
    
    DrawMenuBar m_hWnd
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal bln As Boolean)
    If (bln = True) Then
        If (CheckMenuItem(m_hWnd, m_CommandID, MF_BYCOMMAND Or MF_CHECKED) <> -1) Then
            m_Checked = True
        End If
    Else
        If (CheckMenuItem(m_hWnd, m_CommandID, MF_BYCOMMAND Or MF_UNCHECKED) <> -1) Then
            m_Checked = False
        End If
    End If
    
    DrawMenuBar m_hWnd
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal bln As Boolean)
    If (bln = True) Then
        CreateMenu
    Else
        DeleteMenu
    End If
End Property

'To-Do: Complete this Function
Public Property Get HelpContextID() As Long
End Property

'To-Do: Complete this Function
Public Property Let HelpContextID(ByVal bln As Long)
End Property

Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal str As String)
    m_Tag = str
End Property

'To-Do: Complete this Function
Public Property Get WindowList() As Boolean
End Property

Private Function CreateMenu() As Long
    If (m_CommandID = 0) Then
    
        m_CommandID = CreatePopupMenu()
    
        AppendMenu m_hWnd, IIf(HasChildren, MF_POPUP, MF_STRING), m_CommandID, ""
        
        DrawMenuBar m_hWnd
        
        CreateMenu = m_CommandID
        
    End If
End Function

Private Function DeleteMenu() As Long
    If (m_CommandID <> 0) Then
    
        RemoveMenu m_hWnd, m_CommandID, MF_BYCOMMAND
            
        m_CommandID = 0

    End If
End Function
