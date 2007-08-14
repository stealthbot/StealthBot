VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' This is not my code
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
'Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
'Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
'Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hWnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'local variable(s) to hold property value(s)
Private mvarBackColor As Long
Private mvarTitle As String
Private mvarForeColor As Long
Private mvarIcon As ttIconType
Private mvarCentered As Boolean
Private mvarStyle As ttStyleEnum
Private mvarTipText As String
Private mvarVisibleTime As Long
Private mvarDelayTime As Long

'private data
Private m_lTTHwnd As Long ' hwnd of the tooltip
Private m_lParentHwnd As Long ' hwnd of the window the tooltip attached to
Private ti As TOOLINFO

Public Property Let Style(ByVal vData As ttStyleEnum)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.Style = 5
   mvarStyle = vData
End Property

Public Property Get Style() As ttStyleEnum
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.Style
   Style = mvarStyle
End Property

Public Property Let Centered(ByVal vData As Boolean)
   'used when assigning a value to the property, on the left side of an assignment.
   'Syntax: X.Centered = 5
   mvarCentered = vData
End Property

Public Property Get Centered() As Boolean
   'used when retrieving value of a property, on the right side of an assignment.
   'Syntax: Debug.Print X.Centered
   Centered = mvarCentered
End Property

Public Function Create(ByVal ParentHwnd As Long, X As Long, Y As Long) As Boolean
   Dim lWinStyle As Long
   
   If m_lTTHwnd <> 0 Then
      DestroyWindow m_lTTHwnd
   End If
   
   m_lParentHwnd = ParentHwnd
   
   lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
   
   ''create baloon style if desired
   If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
   
    
   m_lTTHwnd = CreateWindowEx(0&, _
      TOOLTIPS_CLASSA, _
      "SBToolTipWindow", _
      lWinStyle, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      CW_USEDEFAULT, _
      0&, _
      0&, _
      App.hInstance, _
      0&)
               
   ''now set our tooltip info structure
   With ti
      ''if we want it centered, then set that flag
      If mvarCentered Then
         .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
      Else
         .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
      End If
       
      ''set the hwnd prop to our parent control's hwnd
      .hWnd = m_lParentHwnd
      .lId = m_lParentHwnd '0
      .hInstance = App.hInstance
      '.lpstr = ALREADY SET
      '.lpRect = lpRect
      .lSize = Len(ti)
   End With
   
   ''add the tooltip structure
   SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti

   ''if we want a title or we want an icon
   If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If

   If mvarForeColor <> Empty Then
      SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
   End If

   If mvarBackColor <> Empty Then
      SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
   End If
   
   SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime
   SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime
   
   Create = CBool(m_lTTHwnd)
End Function

Public Property Let Icon(ByVal vData As ttIconType)
   mvarIcon = vData
   If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property

Public Property Get Icon() As ttIconType
   Icon = mvarIcon
End Property

Public Property Let ForeColor(ByVal vData As Long)
   mvarForeColor = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
   End If
End Property

Public Property Get ForeColor() As Long
   ForeColor = mvarForeColor
End Property

Public Property Let Title(ByVal vData As String)
   mvarTitle = vData
   If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
      SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
   End If
End Property

Public Property Get Title() As String
   Title = ti.lpStr
End Property

Public Property Let BackColor(ByVal vData As Long)
   mvarBackColor = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
   End If
End Property

Public Property Get BackColor() As Long
   BackColor = mvarBackColor
End Property

Public Property Let TipText(ByVal vData As String)
   mvarTipText = vData
   ti.lpStr = vData
   If m_lTTHwnd <> 0 Then
      SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
   End If
End Property

Public Property Get TipText() As String
   TipText = mvarTipText
End Property

Private Sub Class_Initialize()
   InitCommonControls
   mvarDelayTime = 200
   mvarVisibleTime = 5000
End Sub

Private Sub Class_Terminate()
   Destroy
End Sub

Public Sub Destroy()
   If m_lTTHwnd <> 0 Then
      DestroyWindow m_lTTHwnd
   End If
End Sub

Public Property Get VisibleTime() As Long
   VisibleTime = mvarVisibleTime
End Property

Public Property Let VisibleTime(ByVal lData As Long)
   mvarVisibleTime = lData
End Property

Public Property Get DelayTime() As Long
   DelayTime = mvarDelayTime
End Property

Public Property Let DelayTime(ByVal lData As Long)
   mvarDelayTime = lData
End Property

