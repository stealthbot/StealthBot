VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyManager 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage CDKeys"
   ClientHeight    =   2685
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetKey 
      Caption         =   "Set Key"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   28
      ImageHeight     =   14
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":02C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":04EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":0726
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":0A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":0FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKeyManager.frx":14B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvKeys 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   16777215
      BackColor       =   10040064
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Key"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete Selected"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Selected"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtActiveKey 
      BackColor       =   &H00993300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3375
   End
End
Attribute VB_Name = "frmKeyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private KeyProducts As Dictionary

Private m_editing As String

Private Sub Form_Load()
    Me.Icon = frmChat.Icon
    
    ' Set default button states
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSetKey.Enabled = False
    
    ' Create a lookup for product codes to icon indexes
    Set KeyProducts = New Dictionary
    KeyProducts.Add &H1, 3  ' STAR
    KeyProducts.Add &H2, 3  ' STAR
    KeyProducts.Add &H4, 4  ' W2BN
    KeyProducts.Add &H5, 1  ' D2DV (beta, defunct)
    KeyProducts.Add &H6, 5  ' D2DV
    KeyProducts.Add &H7, 5  ' D2DV
    KeyProducts.Add &H9, 1  ' D2DV (stress test, defunct)
    KeyProducts.Add &HA, 6  ' D2XP
    KeyProducts.Add &HC, 6  ' D2XP
    KeyProducts.Add &HD, 1  ' WAR3 (beta, defunct)
    KeyProducts.Add &HE, 7  ' WAR3
    KeyProducts.Add &HF, 7  ' WAR3
    KeyProducts.Add &H11, 1 ' W3XP (beta, defunct)
    KeyProducts.Add &H12, 8 ' W3XP
    KeyProducts.Add &H13, 1 ' W3XP (retail, disabled)
    KeyProducts.Add &H17, 3 ' STAR (online upgrade)
    KeyProducts.Add &H18, 5 ' D2DV (online upgrade)
    KeyProducts.Add &H19, 6 ' D2XP (online upgrade)
    
    Call Local_LoadCDKeys
    
    If lvKeys.ListItems.Count > 0 Then
        Set lvKeys.SelectedItem = lvKeys.ListItems(1)
        lvKeys_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmChat.SettingsForm.Show
    frmChat.SettingsForm.SetFocus
End Sub

Private Sub cmdAdd_Click()
    ProcessKey txtActiveKey.Text
    
    m_editing = vbNullString
    
    txtActiveKey.Text = vbNullString
    lvKeys.SetFocus
End Sub

Private Sub cmdDelete_Click()
    If Not (lvKeys.SelectedItem Is Nothing) Then
        lvKeys.ListItems.Remove lvKeys.SelectedItem.Index
    End If
End Sub

Private Sub cmdDone_Click()
    If LenB(m_editing) > 0 Then
        ProcessKey m_editing
        m_editing = vbNullString
    End If
    
    Call Local_WriteCDKeys
    
    Unload Me
End Sub

Private Sub cmdSetKey_Click()
    If ((Not (frmChat.SettingsForm Is Nothing)) And (Not (lvKeys.SelectedItem Is Nothing))) Then
        If ((lvKeys.SelectedItem.SmallIcon = 6) Or (lvKeys.SelectedItem.SmallIcon = 8)) Then
            frmChat.SettingsForm.txtExpKey.Text = lvKeys.SelectedItem.Tag
        Else
            frmChat.SettingsForm.txtCDKey.Text = lvKeys.SelectedItem.Tag
        End If
        
        Call cmdDone_Click
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not (lvKeys.SelectedItem Is Nothing) Then
        m_editing = lvKeys.SelectedItem.Tag
        lvKeys.ListItems.Remove lvKeys.SelectedItem.Index
        With txtActiveKey
            .Text = m_editing
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub lvKeys_Click()
    Dim Value As Boolean
    Value = (Not (lvKeys.SelectedItem Is Nothing))
    
    cmdEdit.Enabled = Value
    cmdDelete.Enabled = Value
    cmdSetKey.Enabled = Value
End Sub

Private Sub lvKeys_DblClick()
    Call cmdSetKey_Click
End Sub

Private Sub lvKeys_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSetKey_Click
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdDone_Click
    End If
End Sub

Private Sub txtActiveKey_GotFocus()
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSetKey.Enabled = False
End Sub

Private Sub txtActiveKey_Change()
    cmdAdd.Enabled = (Len(txtActiveKey.Text) > 0)
End Sub

Private Sub txtActiveKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cmdAdd.Enabled Then
        Call cmdAdd_Click
    ElseIf KeyAscii = vbKeyEscape Then
        If LenB(m_editing) > 0 Then
            ProcessKey m_editing
    
            m_editing = vbNullString
            
            txtActiveKey.Text = vbNullString
            lvKeys.SetFocus
        ElseIf LenB(txtActiveKey.Text) > 0 Then
            txtActiveKey.Text = vbNullString
        Else
            Call cmdDone_Click
        End If
    End If
End Sub

Private Sub ProcessKey(ByVal sKey As String)
    Dim oKey As New clsKeyDecoder
    Dim KeyProduct As Long
    
    oKey.Initialize sKey
    If Not oKey.IsValid Then
        KeyProduct = -1
    Else
        KeyProduct = oKey.ProductValue
    End If
    
    AddUnique oKey.GetKeyForDisplay(), GetImageIndex(KeyProduct), oKey.Key
    
    Set oKey = Nothing
End Sub

' Adds the specified text and image while checking for duplicates
Private Sub AddUnique(ByVal strNewValue As String, ByVal image As Integer, ByVal Tag As String)
    Dim Item As ListItem
    
    For Each Item In lvKeys.ListItems
        If StrComp(Item.Tag, Tag, vbTextCompare) = 0 Then Exit Sub
    Next Item
    
    AddItem strNewValue, image, Tag
End Sub

' Adds the specified text and image regardless of duplicates
Private Sub AddItem(ByVal Text As String, ByVal image As Integer, ByVal Tag As String)
    With lvKeys.ListItems.Add(, , Text, , image)
        .Tag = Tag
    End With
End Sub

' Returns the image used to identify the key.
Private Function GetImageIndex(ByVal productCode As Long) As Integer
    If productCode = -1 Then GetImageIndex = 1: Exit Function    ' invalid
    
    If KeyProducts.Exists(productCode) Then
        GetImageIndex = KeyProducts.Item(productCode)
    Else
        GetImageIndex = 2   ' unrecognized
    End If
End Function


Private Sub Local_LoadCDKeys()
    Dim keys As Collection
    Dim sKey As Variant
    Set keys = ListFileLoad(GetFilePath(FILE_KEY_LIST))
    
    For Each sKey In keys
        sKey = CStr(Trim(sKey))
        If Len(sKey) > 0 Then ProcessKey sKey
    Next sKey
End Sub

Private Sub Local_WriteCDKeys()
    Dim keys As Collection
    Dim Item As ListItem
    
    Set keys = New Collection
    
    For Each Item In lvKeys.ListItems
        keys.Add Item.Tag
    Next Item
    
    ListFileSave GetFilePath(FILE_KEY_LIST), keys
    
    Set keys = Nothing
End Sub
