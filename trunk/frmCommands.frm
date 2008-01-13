VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCommands 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Manager"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView trvCommands 
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7435
      _Version        =   393217
      Indentation     =   575
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
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
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Apply and Cl&ose"
      Height          =   300
      Index           =   0
      Left            =   4080
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "kick"
      Height          =   4335
      Left            =   2400
      TabIndex        =   1
      Top             =   23
      Width           =   3015
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   292
         Index           =   1
         Left            =   2160
         TabIndex        =   14
         Top             =   548
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         ItemData        =   "frmCommands.frx":0000
         Left            =   240
         List            =   "frmCommands.frx":0002
         TabIndex        =   12
         Top             =   560
         Width           =   1295
      End
      Begin VB.TextBox txtRank 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         MaxLength       =   25
         TabIndex        =   9
         Top             =   1215
         Width           =   1215
      End
      Begin VB.TextBox txtFlags 
         BackColor       =   &H00993300&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   8
         Top             =   1215
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Disable Command"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00993300&
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   292
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   548
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Custom aliases:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   310
         Width           =   1695
      End
      Begin VB.Label lblRank 
         Caption         =   "Rank (1 - 200):"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label lblFlags 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   975
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Special notes:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   300
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CommandsDoc As MSXML2.DOMDocument

' ...
Private Sub Form_Load()
    Set m_CommandsDoc = New MSXML2.DOMDocument
    
    ' ...
    If (Dir$(App.Path & "\commands.xml") = vbNullString) Then
        Exit Sub
    End If
    
    ' ...
    Call m_CommandsDoc.load(App.Path & "\commands.xml")
    
    ' ...
    Call PopulateTreeView
End Sub

' ...
Private Sub Form_Unload(Cancel As Integer)
    Set m_CommandsDoc = Nothing
End Sub

' ...
Private Sub PopulateTreeView()
    Dim xmlCommand        As MSXML2.IXMLDOMNode
    Dim xmlArgs           As MSXML2.IXMLDOMNodeList
    Dim xmlArgRestricions As MSXML2.IXMLDOMNodeList

    Dim nCommand          As Node ' ...
    Dim nArg              As Node ' ...
    Dim nArgRestriction   As Node ' ...
    
    Dim i As Integer ' ...

    ' ...
    For Each xmlCommand In m_CommandsDoc.documentElement.childNodes
        ' ...
        Set nCommand = _
            trvCommands.Nodes.Add(, , , _
                xmlCommand.Attributes.getNamedItem("name").text)
        
        ' ...
        Set xmlArgs = _
            xmlCommand.selectNodes("arguments/argument")
        
        ' ...
        If (Not (xmlArgs Is Nothing)) Then
            Dim j As Integer ' ...
        
            ' ...
            For i = 0 To (xmlArgs.length - 1)
                ' ...
                Set nArg = _
                    trvCommands.Nodes.Add(nCommand, tvwChild, , _
                        xmlArgs(i).Attributes.getNamedItem("name").text)
                        
                ' ...
                Set xmlArgRestricions = _
                    xmlArgs(i).selectNodes("restriction")
                    
                ' ...
                For j = 0 To (xmlArgRestricions.length - 1)
                    ' ...
                    Set nArgRestriction = _
                        trvCommands.Nodes.Add(nArg, tvwChild, , _
                            xmlArgRestricions(j).Attributes.getNamedItem("name").text)
                Next j
            Next i
        End If
    Next
End Sub
