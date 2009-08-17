VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5295
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF6633&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   6735
   End
   Begin VB.Image Logo 
      Height          =   4500
      Left            =   120
      Picture         =   "frmSplash.frx":0000
      Top             =   120
      Width           =   6750
   End
   Begin VB.Image BDay 
      Height          =   4740
      Left            =   1320
      Picture         =   "frmSplash.frx":22A21
      Top             =   120
      Visible         =   0   'False
      Width           =   4500
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bday_Click()
    frmChat.Show
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmChat.Show
    Unload Me
End Sub

Private Sub Image_click()
    frmChat.Show
    Unload Me
End Sub

Private Sub IsBirthday(sName As String, iBorn As Integer)
    On Error Resume Next 'this isn't important so screw the errors
    Dim iAge As Integer
    Dim sAppend As String
    
    Logo.Visible = False
    BDay.Visible = True
    
    sAppend = "th"
    iAge = Year(Now()) - iBorn
    If (iAge <= 10 Or iAge >= 20) Then
        If (iAge Mod 10 = 1) Then
            sAppend = "st"
        ElseIf (iAge Mod 10 = 2) Then
            sAppend = "nd"
        ElseIf (iAge Mod 10 = 3) Then
            sAppend = "rd"
        End If
    End If
    
    Label1.Caption = "Happy " & iAge & sAppend & " Birthday " & sName & "!!!"
End Sub

Private Sub Form_Load()
    On Error Resume Next 'this isn't important so screw the errors
    Dim sDate As String
    Me.Icon = frmChat.Icon
    
    sDate = LCase(Month(Now()) & "/" & Day(Now()))
    
    Select Case sDate
      Case "2/21":  IsBirthday "Ribose", 1992
      Case "3/18":  IsBirthday "Pyro", 1992
      Case "5/9":   IsBirthday "Snap", 1987
      Case "4/3":   IsBirthday "Stealth", 1987
      Case "4/22":  IsBirthday "52", 1982
      Case "9/22":  IsBirthday "Eric", 1987
      Case "11/26": IsBirthday "Hdx", 1989
      Case Else: Label1.Caption = "[ " & CVERSION & " ]"
    End Select
End Sub

Private Sub Logo_Click()
    frmChat.Show
    Unload Me
End Sub
