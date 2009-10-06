VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLauncher 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StealthBot Launcher v0.0.000"
   ClientHeight    =   5145
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAutoClose 
      BackColor       =   &H00000000&
      Caption         =   "Automatically close this launcher after loading the profile"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Leaving the launcher open will allow you to create and launch additional profiles."
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton cmdRemoveProfile 
      Caption         =   "Remove Profile"
      Enabled         =   0   'False
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateProfile 
      Caption         =   "Create Profile"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdLaunchThis 
      Caption         =   "Launch Selected Profile"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdCreateShortcut 
      Caption         =   "Create a Shortcut"
      Enabled         =   0   'False
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin MSComctlLib.ListView lstProfiles 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   10040064
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "to this profile on your Desktop"
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1800
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblProfiles 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "List of available profiles:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2955
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuCreateProfile 
         Caption         =   "Create New Profile"
      End
      Begin VB.Menu mnuLaunchProfile 
         Caption         =   "Launch Profile"
      End
      Begin VB.Menu mnuCreateShortcut 
         Caption         =   "Create Shortcut"
      End
      Begin VB.Menu mnuRenameProfile 
         Caption         =   "Rename Profile"
      End
      Begin VB.Menu mnuRemoveProfile 
         Caption         =   "Remove Profile"
      End
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is a basic profile launcher for StealthBot that allows us to get around the UAC.

Private Const OBJECT_NAME As String = "frmLauncher"

Private Sub Form_Load()
On Error GoTo ERROR_HANDLER
    Me.Caption = StringFormat("StealthBot Launcher v{0}.{1}.{2}", App.Major, App.Minor, App.Revision)
    
    'CheckForUpdates
    
    If (LenB(Command()) > 0) Then
        If (SetCommandLine(Command())) Then
            Unload Me
            Exit Sub
        End If
    End If
    
    #If COMPILE_CRC = 1 Then
        Dim crc As New clsCRC32
        If (Not crc.ValidateExecutable) Then
            MsgBox "This application has been tampered with, Please download a new version at http://www.StealthBot.net/", vbOKOnly + vbCritical
            Unload Me
            Exit Sub
        End If
        Set crc = Nothing
    #End If
    
    Set cConfig = New clsConfig
    
    ' UI: columns
    lstProfiles.ColumnHeaders.Add , , "center", 0, lvwColumnCenter
    SetupColumns lvwColumnLeft

    ' profiles: load
    'modLauncher.LoadXMLDocument
    LoadProfiles
    
    If (lstProfiles.ListItems.Count > 0) Then ' UI: if count > 0 then enable btns
        EnableButtons
    Else ' UI: if count = 0 then show informative item
        SetupColumns lvwColumnCenter
        'AddChat vbGreen, "To get started using StealthBot, Use the ""Create Profile"" Button."
    End If
    
    ' Settings
    chkAutoClose.Value = IIf(cConfig.AutoClose, 1, 0)
    
    'CheckForUpdates
    
    bIsClosing = False
    
    'Load frmStatus
    'With frmStatus
    '    .Top = Me.Top
    '    .Left = Me.Left + Me.Width + 100
    '    .Show
    'End With
    
    'HookWindowProc Me.hWnd
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERROR_HANDLER

    bIsClosing = True
    'UnHookAllProcs
    
    Unload frmNameDialog
    'Unload frmConfig
    'Unload frmstatus
    
    If (Not cConfig Is Nothing) Then cConfig.SaveConfig
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "Form_Unload"
End Sub

Private Sub lstProfiles_DblClick()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            LaunchProfile lstProfiles.SelectedItem.Text
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "lstProfiles_DblClick"
End Sub

Private Sub lstProfiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERROR_HANDLER

    Dim bProfileMenus As Boolean
    bProfileMenus = (Not lstProfiles.SelectedItem Is Nothing)
    If bProfileMenus Then bProfileMenus = (Not lstProfiles.SelectedItem.Ghosted)
    
    mnuLaunchProfile.Visible = bProfileMenus
    mnuCreateShortcut.Visible = bProfileMenus
    mnuRenameProfile.Visible = bProfileMenus
    mnuRemoveProfile.Visible = bProfileMenus

    If (Button = 2) Then
        PopupMenu mnuRightClick
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "lstProfiles_MouseDown"
End Sub

Private Sub mnuCreateProfile_Click()
On Error GoTo ERROR_HANDLER
    Load frmNameDialog
    frmNameDialog.Show
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "lstProfiles_MouseDown"
End Sub

Private Sub cmdCreateProfile_Click()
On Error GoTo ERROR_HANDLER
    Load frmNameDialog
    frmNameDialog.Show
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdCreateProfile_Click"
End Sub

Private Sub mnuCreateShortcut_Click()
On Error GoTo ERROR_HANDLER:

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            CreateShortcut lstProfiles.SelectedItem.Text
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "mnuCreateShortcut_Click"
End Sub

Private Sub cmdCreateShortcut_Click()
On Error GoTo ERROR_HANDLER:

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            CreateShortcut lstProfiles.SelectedItem.Text
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdCreateShortcut_Click"
End Sub
'
'Private Sub mnuInformation_Click()
'On Error GoTo ERROR_HANDLER:
'
'    If (frmStatus.Visible) Then
'        frmStatus.Hide
'    Else
'        frmStatus.Show
'    End If
'
'    Exit Sub
'ERROR_HANDLER:
'    ErrorHandler Err.Number, OBJECT_NAME, "mnuInformation_Click"
'End Sub

Private Sub mnuLaunchProfile_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            LaunchProfile lstProfiles.SelectedItem.Text
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "mnuLaunchProfile_Click"
End Sub

Private Sub cmdLaunchThis_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            LaunchProfile lstProfiles.SelectedItem.Text
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdLaunchThis_Click"
End Sub

Private Sub mnuRenameProfile_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        ' TODO: impl rename profile (Name currfoldername As newname)
        ' use name dialog?
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "mnuRenameProfile_Click"
End Sub

' TODO: this function has no button! change UI to include rename button?
Private Sub cmdRenameProfile_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        ' TODO: impl rename profile (Name currfoldername As newname)
        ' use name dialog?
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdRenameProfile_Click"
End Sub

Private Sub mnuRemoveProfile_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            RemoveProfile lstProfiles.SelectedItem
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "mnuRemoveProfile_Click"
End Sub

Private Sub cmdRemoveProfile_Click()
On Error GoTo ERROR_HANDLER

    If (Not lstProfiles.SelectedItem Is Nothing) Then
        If (modLauncher.ProfileExists(lstProfiles.SelectedItem.Text)) Then
            RemoveProfile lstProfiles.SelectedItem
        End If
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "cmdRemoveProfile_Click"
End Sub

Private Sub LoadProfiles()
On Error GoTo ERROR_HANDLER
    Dim sFolder As String
    Dim sFile   As String
    sFolder = modLauncher.ReplaceEnvironmentVars("%APPDATA%\StealthBot\")
    
    If (LenB(Dir$(sFolder, vbDirectory)) = 0) Then
        Call modLauncher.MakeDirectory(sFolder)
        Exit Sub
    End If
    
    Do While True
        sFile = Dir$
        If (LenB(sFile) = 0) Then Exit Do
        If ((Not sFile = "..") And _
            ((GetFileAttributes(sFolder & sFile) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)) Then
            
            lstProfiles.ListItems.Add , , sFile
        End If
    Loop
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "LoadProfiles"
End Sub


' enables the buttons that are enabled only when an item is selected
Private Sub EnableButtons()
On Error GoTo ERROR_HANDLER
    cmdCreateShortcut.Enabled = True
    cmdRemoveProfile.Enabled = True
    cmdLaunchThis.Enabled = True
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "EnableButtons"
End Sub

' disables the buttons that are enabled only when an item is selected
Private Sub DisableButtons()
On Error GoTo ERROR_HANDLER
    cmdCreateShortcut.Enabled = False
    cmdRemoveProfile.Enabled = False
    cmdLaunchThis.Enabled = False
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "EnableButtons"
End Sub



' adds profile item to list view,
' deals with informational item if exists
Public Sub ListProfile(ByVal Text As String)
On Error GoTo ERROR_HANDLER
    If lstProfiles.ListItems.Count = 3 Then
        If lstProfiles.ListItems(1).Ghosted Then
            EnableButtons
            SetupColumns lvwColumnLeft
        End If
    End If
    lstProfiles.ListItems.Add , , Text
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "ListProfile"
End Sub

' removes profile item from list view,
' deals with informational item if needs
' to be recreated (all profiles removed)
Public Sub UnlistProfile(Index As Integer)
On Error GoTo ERROR_HANDLER
    lstProfiles.ListItems.Remove Index
    If lstProfiles.ListItems.Count = 0 Then
        DisableButtons
        SetupColumns lvwColumnCenter
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "UnlistProfile"
End Sub

' this will set up columns in one of two fasions:
' lvwColumnLeft-
'  uses the first column, say if there are profiles to show
'  alignment will be to left
' lvwColumnCenter-
'  uses the second column due to this error:
'  "The first column in a ListView control must be left aligned in frmLauncher.SetupColumn()"
'  so that the info text in AddInformationalItem shows up centered!
'  use .SubItem(1) = text to set text in this case...
' items will be cleared and AddInformationalItem() called,
' so neither has to be done anywhere else
Private Sub SetupColumns(ByVal Alignment As ListColumnAlignmentConstants)
On Error GoTo ERROR_HANDLER:
    Const COLUMN_WIDTH_PX As Integer = 170

    lstProfiles.ListItems.Clear
    Select Case Alignment
        Case lvwColumnLeft:
            lstProfiles.ColumnHeaders(1).Width = COLUMN_WIDTH_PX
            lstProfiles.ColumnHeaders(2).Width = 0
        Case lvwColumnCenter:
            lstProfiles.ColumnHeaders(1).Width = 0
            lstProfiles.ColumnHeaders(2).Width = COLUMN_WIDTH_PX
            
            AddInformationalItem
    End Select
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "SetupColumns"
End Sub

' adds two informational items,
' "You have no profiles. Create a profile/
' and launch it to begin using StealthBot!"
' should be using SetupColumns(center) to call this
Private Sub AddInformationalItem()
On Error GoTo ERROR_HANDLER
    With lstProfiles
        With .ListItems.Add()
            .Ghosted = True
            .SubItems(1) = "You have no profiles."
        End With
        With .ListItems.Add()
            .Ghosted = True
            .SubItems(1) = "Create and launch one"
        End With
        With .ListItems.Add()
            .Ghosted = True
            .SubItems(1) = "to begin using StealthBot!"
        End With
    End With
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "AddInformationalItem"
End Sub
'
'Private Sub mnuSettings_Click()
'On Error GoTo ERROR_HANDLER:
'    Load frmConfig
'    frmConfig.Show
'    Exit Sub
'ERROR_HANDLER:
'    ErrorHandler Err.Number, OBJECT_NAME, "mnuSettings_Click"
'End Sub
