VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H00000000&
   Caption         =   ":: StealthBot &version :: Disconnected ::"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvChannel 
      Height          =   6375
      Left            =   8880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   10079232
      BackColor       =   0
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
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4145
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1252
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   582
      EndProperty
   End
   Begin VB.Timer tmrAccountLock 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   6240
      Top             =   4560
   End
   Begin MSScriptControlCtl.ScriptControl SControl 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSScriptControlCtl.ScriptControl SCRestricted 
      Left            =   5880
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer tmrScriptLong 
      Enabled         =   0   'False
      Index           =   0
      Left            =   1200
      Top             =   120
   End
   Begin InetCtlsObjects.Inet itcScript 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   7
   End
   Begin MSWinsockLib.Winsock sckScript 
      Index           =   0
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer ChatQueueTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5280
      Top             =   4560
   End
   Begin VB.Timer cacheTimer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   5280
      Top             =   4080
   End
   Begin MSComctlLib.ImageList imlClan 
      Left            =   6480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   37
      ImageHeight     =   23
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2048
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2302
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":25AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2A8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4C5D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   28
      ImageHeight     =   18
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   166
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":6DD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":7071
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":731B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":7955
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":7BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":7F99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":834B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":8985
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":8FBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":915D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":9417
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":9A5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":9BFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":9DA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":A05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":A1E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":A37B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":A609
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":A8DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":ABA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":AE8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":B159
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":B520
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":B896
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":BC1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":BFB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":C35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":C70E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":CAD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":CE55
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":D22B
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":D551
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":D96B
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":DDCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":E407
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":EA41
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":F02F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":F669
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":FCA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":FF14
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1017E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1038F
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":105C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1082D
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":10A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":10CD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":10F3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":11160
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1138D
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":115A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":117E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":11A2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":11C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":11E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":120D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":12300
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":127A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":12A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":12C53
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":12EA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":130E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":13351
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":135B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1381A
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":13A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":13CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":13F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":14189
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":143A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1460A
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":14844
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":14A6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":14CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":14EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1514F
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":153C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1562A
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":15882
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":15AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":15CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":15F1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":16187
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":163C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1660C
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1685A
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":16AA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":16D0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":16F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":17183
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":173BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1762B
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":17895
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":17ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":17D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":17F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":18177
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":183E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":185E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":187E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":18A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":18C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":18F43
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":19289
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1985A
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":19E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1A3D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1AA6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1B0FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1B78E
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1BCF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1C262
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1C7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1CD7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1D320
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1D8CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2360A
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":292CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2EEE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3497A
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3A5AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":401C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":406A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":40C5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4131F
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":41960
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":41F8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":425A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":42AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":430F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4372D
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":43D67
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":443A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":449DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":45015
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4564F
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":45C89
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":462C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":468FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":46F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":470D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":47713
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":47D4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":48387
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":489C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":48FFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":49635
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":49C6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4A2A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4A8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4AF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4B557
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4BB91
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4C1CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4C485
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4C757
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4CA29
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4CCD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4CFC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4D2CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4D5D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4D883
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4DB55
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4DE3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4E129
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4E3D3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrScript 
      Enabled         =   0   'False
      Index           =   0
      Left            =   720
      Top             =   600
   End
   Begin VB.Timer tmrSilentChannel 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   30000
      Left            =   5760
      Top             =   4560
   End
   Begin VB.Timer tmrSilentChannel 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   6240
      Top             =   4080
   End
   Begin VB.ComboBox cboSend 
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   6600
      Width           =   7695
   End
   Begin VB.TextBox txtPost 
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   8280
      TabIndex        =   9
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox txtPre 
      BackColor       =   &H00000000&
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
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   6600
      Width           =   615
   End
   Begin TabDlg.SSTab ListviewTabs 
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   6600
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   661
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Channel  "
      TabPicture(0)   =   "frmChat.frx":4E661
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Friends  "
      TabPicture(1)   =   "frmChat.frx":4E67D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Clan  "
      TabPicture(2)   =   "frmChat.frx":4E699
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Timer tmrIdleTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4560
   End
   Begin VB.CommandButton cmdShowHide 
      Caption         =   " ^^^^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12360
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6600
      Width           =   245
   End
   Begin MSWinsockLib.Winsock sckMCP 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer scTimer 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   600
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   5280
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   7
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   5760
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9367
   End
   Begin MSWinsockLib.Winsock sckBNet 
      Left            =   6240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6112
   End
   Begin VB.Timer UpTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   4080
   End
   Begin MSComctlLib.ListView lvClanList 
      Height          =   6375
      Left            =   8880
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   11245
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      SmallIcons      =   "imlClan"
      ForeColor       =   10079232
      BackColor       =   0
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
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvFriendList 
      Height          =   6375
      Left            =   8880
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   11245
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "imlIcons"
      SmallIcons      =   "imlIcons"
      ForeColor       =   10079232
      BackColor       =   0
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
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Object.Width           =   88
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   0
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbWhispers 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6960
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2990
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":4E6B5
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
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   6615
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   11668
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmChat.frx":4E746
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCurrentChannel 
      Alignment       =   2  'Center
      BackColor       =   &H00CC3300&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   0
      Width           =   3698
   End
   Begin VB.Menu mnuBot 
      Caption         =   "&Bot"
      Begin VB.Menu mnuConnect2 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect2 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuSepT 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountManager 
         Caption         =   "&Account Manager..."
      End
      Begin VB.Menu mnuUserDBManager 
         Caption         =   "&User Database Manager..."
      End
      Begin VB.Menu mnuCommandManager 
         Caption         =   "Command &Manager..."
      End
      Begin VB.Menu mnuScriptManager 
         Caption         =   "&Script Settings Manager..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKeyManager 
         Caption         =   "CD-&Key Manager..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Edit &Profile..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRealmSwitch 
         Caption         =   "Switch &Realm Character..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClanCreate 
         Caption         =   "Create C&lan..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQCTop 
         Caption         =   "Channel &List"
         Begin VB.Menu mnuHomeChannel 
            Caption         =   "&Bot Home"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLastChannel 
            Caption         =   "&Last Channel"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuQCDash 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuQCHeader 
            Caption         =   "- QuickChannels -"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   1
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   2
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   3
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   4
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   5
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   6
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   7
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuCustomChannels 
            Caption         =   ""
            Index           =   8
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuCustomChannelAdd 
            Caption         =   "&Add QuickChannel"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCustomChannelEdit 
            Caption         =   "&Edit QuickChannels..."
         End
         Begin VB.Menu mnuPCDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPCHeader 
            Caption         =   "- Public Channels -"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPublicChannels 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "View &Files"
         Begin VB.Menu mnuOpenBotFolder 
            Caption         =   "Open Bot &Folder"
         End
         Begin VB.Menu mnuSepA 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClearedTxt 
            Caption         =   "Current Text &Log"
         End
         Begin VB.Menu mnuWhisperCleared 
            Caption         =   "&Whisper Window Text Log"
         End
         Begin VB.Menu mnuSepB 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditCaught 
            Caption         =   "Caught &Phrases"
         End
      End
      Begin VB.Menu mnuSettingsRepair 
         Caption         =   "&Tools"
         Begin VB.Menu mnuToolsMenuWarning 
            Caption         =   "- Use Carefully -"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRepairDataFiles 
            Caption         =   "Delete &Data Files"
         End
         Begin VB.Menu mnuUpdateVerbytes 
            Caption         =   "&Update Version Bytes"
         End
         Begin VB.Menu mnuRepairVerbytes 
            Caption         =   "Restore Default &Version Bytes"
         End
         Begin VB.Menu mnuRepairCleanMail 
            Caption         =   "Clean Up &Mail Database"
         End
         Begin VB.Menu mnuPacketLog 
            Caption         =   "Log StealthBot &Packet Traffic"
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetTop 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSetup 
         Caption         =   "&Bot Settings..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuUTF8 
         Caption         =   "Use &UTF-8 in Chat"
      End
      Begin VB.Menu mnuLogging 
         Caption         =   "&Logging Settings"
         Begin VB.Menu mnuLog0 
            Caption         =   "Full Text Logging"
         End
         Begin VB.Menu mnuLog1 
            Caption         =   "Temporary Logging"
         End
         Begin VB.Menu mnuLog2 
            Caption         =   "No Logging"
         End
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "Edit Chat &Filters..."
      End
      Begin VB.Menu mnuCatchPhrases 
         Caption         =   "Edit &Catch Phrases..."
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload Config"
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuDisconnect 
      Caption         =   "&Disconnect"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuToggle 
         Caption         =   "Hide &Join/Leave Messages"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuHideBans 
         Caption         =   "Hide &Ban Messages"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Loc&k Chat Window"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuToggleFilters 
         Caption         =   "Use Chat &Filtering"
         Checked         =   -1  'True
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleWWUse 
         Caption         =   "Use Individual &Whisper Windows"
      End
      Begin VB.Menu mnuSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleShowOutgoing 
         Caption         =   "Show &Outgoing Whispers in Whisper Box"
      End
      Begin VB.Menu mnuHideWhispersInrtbChat 
         Caption         =   "&Hide Whispers in Main Window"
      End
      Begin VB.Menu mnuSepC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Chat Window"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuClearWW 
         Caption         =   "Cl&ear Whisper Window"
      End
      Begin VB.Menu mnuSepD 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlash 
         Caption         =   "Fl&ash Window on Events"
      End
      Begin VB.Menu mnuDisableVoidView 
         Caption         =   "Disable &Silent Channel View"
      End
      Begin VB.Menu mnuRecordWindowPos 
         Caption         =   "&Record Current Position"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "systray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayCaption 
         Caption         =   ""
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopWhisper 
         Caption         =   "W&hisper"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy Name to Clipboard"
      End
      Begin VB.Menu mnuPopAddLeft 
         Caption         =   "Add to &Left Send Box"
      End
      Begin VB.Menu mnuPopAddToFList 
         Caption         =   "Add to &Friends List"
      End
      Begin VB.Menu mnuPopInvite 
         Caption         =   "&Invite to Warcraft III Clan"
      End
      Begin VB.Menu mnuPopShitlist 
         Caption         =   "Shi&tlist"
      End
      Begin VB.Menu mnuPopSafelist 
         Caption         =   "S&afelist"
      End
      Begin VB.Menu mnuPopSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUserlistWhois 
         Caption         =   "&Userlist Whois"
      End
      Begin VB.Menu mnuPopWhois 
         Caption         =   "Battle.net &Whois"
      End
      Begin VB.Menu mnuPopStats 
         Caption         =   "Battle.net &Stats"
      End
      Begin VB.Menu mnuPopProfile 
         Caption         =   "Battle.net &Profile"
      End
      Begin VB.Menu mnuPopWebProfile 
         Caption         =   "W&eb Profile"
      End
      Begin VB.Menu mnuPopSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopKick 
         Caption         =   "&Kick"
      End
      Begin VB.Menu mnuPopBan 
         Caption         =   "&Ban"
      End
      Begin VB.Menu mnuPopSquelch 
         Caption         =   "S&quelch"
      End
      Begin VB.Menu mnuPopUnsquelch 
         Caption         =   "U&nsquelch"
      End
      Begin VB.Menu mnuPopDes 
         Caption         =   "&Designate"
      End
   End
   Begin VB.Menu mnuScripting 
      Caption         =   "Sc&ripting"
      Begin VB.Menu mnuReloadScripts 
         Caption         =   "Reload Scripts"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuOpenScriptFolder 
         Caption         =   "Open Script Folder"
      End
      Begin VB.Menu mnuScriptingDash 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuGetNews 
         Caption         =   "Get &News and Check for Updates"
      End
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&Wiki"
      End
      Begin VB.Menu mnuHelpWebsite 
         Caption         =   "&Forum"
      End
      Begin VB.Menu mnuTerms 
         Caption         =   "&End-User License Agreement"
      End
      Begin VB.Menu mnuChangeLog 
         Caption         =   "&Change Log"
      End
   End
   Begin VB.Menu mnuShortcuts 
      Caption         =   "invisibleMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuItalic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuUnderline 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuListViewButton 
         Caption         =   "invisible CTRL+A Channel"
         Index           =   0
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewButton 
         Caption         =   "invisible CTRL+S Friends"
         Index           =   1
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu mnuListViewButton 
         Caption         =   "invisible CTRL+D Clan"
         Index           =   2
         Shortcut        =   ^D
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPopClanList 
      Caption         =   "clanlist popup menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopClanWhisper 
         Caption         =   "W&hisper"
      End
      Begin VB.Menu mnuPopClanCopy 
         Caption         =   "&Copy Name to Clipboard"
      End
      Begin VB.Menu mnuPopClanAddLeft 
         Caption         =   "Add to &Left Send Box"
      End
      Begin VB.Menu mnuPopClanAddToFList 
         Caption         =   "Add to &Friends List"
      End
      Begin VB.Menu mnuPopClanSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopClanUserlistWhois 
         Caption         =   "&Userlist Whois"
      End
      Begin VB.Menu mnuPopClanWhois 
         Caption         =   "Battle.net &Whois"
      End
      Begin VB.Menu mnuPopClanStats 
         Caption         =   "Battle.net &Stats"
         Begin VB.Menu mnuPopClanStatsWAR3 
            Caption         =   "&Reign of Chaos"
         End
         Begin VB.Menu mnuPopClanStatsW3XP 
            Caption         =   "The &Frozen Throne"
         End
      End
      Begin VB.Menu mnuPopClanProfile 
         Caption         =   "&Battle.net Profile"
      End
      Begin VB.Menu mnuPopClanWebProfile 
         Caption         =   "W&eb Profile"
         Begin VB.Menu mnuPopClanWebProfileWAR3 
            Caption         =   "&Reign of Chaos"
         End
         Begin VB.Menu mnuPopClanWebProfileW3XP 
            Caption         =   "The &Frozen Throne"
         End
      End
      Begin VB.Menu mnuPopClanSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopClanPromote 
         Caption         =   "&Promote"
      End
      Begin VB.Menu mnuPopClanDemote 
         Caption         =   "&Demote"
      End
      Begin VB.Menu mnuPopClanRemove 
         Caption         =   "Remove from Clan"
      End
      Begin VB.Menu mnuPopClanLeave 
         Caption         =   "Leave Clan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopClanMakeChief 
         Caption         =   "Make Chieftain"
      End
      Begin VB.Menu mnuPopClanDisband 
         Caption         =   "Disband Clan"
      End
   End
   Begin VB.Menu mnuPopFList 
      Caption         =   "flistpopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopFLWhisper 
         Caption         =   "W&hisper"
      End
      Begin VB.Menu mnuPopFLCopy 
         Caption         =   "&Copy Name to Clipboard"
      End
      Begin VB.Menu mnuPopFLAddLeft 
         Caption         =   "Add to &Left Send Box"
      End
      Begin VB.Menu mnuPopFLInvite 
         Caption         =   "&Invite to Warcraft III Clan"
      End
      Begin VB.Menu mnuPopFLSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopFLUserlistWhois 
         Caption         =   "&Userlist Whois"
      End
      Begin VB.Menu mnuPopFLWhois 
         Caption         =   "Battle.net &Whois"
      End
      Begin VB.Menu mnuPopFLStats 
         Caption         =   "Battle.net &Stats"
      End
      Begin VB.Menu mnuPopFLProfile 
         Caption         =   "&Battle.net Profile"
      End
      Begin VB.Menu mnuPopFLWebProfile 
         Caption         =   "W&eb Profile"
      End
      Begin VB.Menu mnuPopFLSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopFLPromote 
         Caption         =   "&Promote"
      End
      Begin VB.Menu mnuPopFLDemote 
         Caption         =   "&Demote"
      End
      Begin VB.Menu mnuPopFLRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuPopFLSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopFLRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StealthBot 11/5/02-Present
'Source Code Version: 2.7RC2
Option Explicit

'REVISION #1337!

'Classes
Public WithEvents ClanHandler As clsClanPacketHandler
Attribute ClanHandler.VB_VarHelpID = -1
Public WithEvents FriendListHandler As clsFriendlistHandler
Attribute FriendListHandler.VB_VarHelpID = -1

'Variables
Private m_lCurItemIndex As Long
Private MultiLinePaste As Boolean
Private doAuth As Boolean
Private AUTH_CHECKED As Boolean

'Forms
Public SettingsForm As frmSettings
Public ChatNoScroll As Boolean

Private Const SB_INET_UNSET As String = vbNullString
Private Const SB_INET_NEWS1 As String = "SBNEWS_AUTOCONNECT"
Private Const SB_INET_NEWS  As String = "SBNEWS"
Private Const SB_INET_BNLS1 As String = "BNLSFINDER"
Private Const SB_INET_BNLS2 As String = "BNLSFINDER_DEFAULT"
Private Const SB_INET_VBYTE As String = "VERBYTE"
Private Const SB_INET_BETA  As String = "AUTHBETA"

' LET IT BEGIN
Private Sub Form_Load()
    Dim s As String
    Dim f As Integer
    Dim HeightVal As Long
    Dim FrmSplashInUse As Boolean
    Dim strBeta As String
    Dim sStr() As String
    
    ' COMPILER FLAGS
    #If (BETA = 1) Then
        CVERSION = StringFormat("StealthBot Beta v{0}.{1} - Build {2}", App.Major, App.Minor, App.REVISION)
    #Else
        CVERSION = StringFormat("StealthBot v{0}.{1}.{2}", App.Major, App.Minor, App.REVISION)
    #End If
    
    #If (COMPILE_DEBUG = 1) Then
        CVERSION = StringFormat("{0} - DEBUG", CVERSION)
    #End If
    
    #If (COMPILE_CRC = 1) Then
        Dim crc As clsCRC32
        Set crc = New clsCRC32
        If (Not crc.ValidateExecutable) Then
            MsgBox GetHexProtectionMessage, vbOKOnly + vbCritical
            Call Form_Unload(0)
            Unload frmChat
            Exit Sub
        End If
        Set crc = Nothing
    #End If

    #If (COMPILE_DEBUG = 0) Then
        HookWindowProc frmChat.hWnd
    #End If
    
    SendMessage frmChat.cboSend.hWnd, CB_LIMITTEXT, 0, 0
    
    Set colWhisperWindows = New Collection
    Set colLastSeen = New Collection
    Set GErrorHandler = New clsErrorHandler
    Set BotVars = New clsBotVars
    Set g_Channel = New clsChannelObj
    Set g_Clan = New clsClanObj
    Set g_Friends = New Collection
    
    sStr = SetCommandLine(Command())
    
    ' EVERYTHING ELSE
    rtbWhispers.Visible = False 'default
    rtbWhispersVisible = False

    'Set dictTimerInterval = New Dictionary
    'Set dictTimerEnabled = New Dictionary
    'Set dictTimerCount = New Dictionary
    'dictTimerInterval.CompareMode = TextCompare
    'dictTimerEnabled.CompareMode = TextCompare
    'dictTimerCount.CompareMode = TextCompare
    
    With mnuTrayCaption
        .Caption = CVERSION
        .Enabled = False
    End With
    
    mail = True
    f = FreeFile
    
    With rtbChat
        .Font.Size = 8
        .SelTabCount = 1
        .SelTabs(0) = 15 * Screen.TwipsPerPixelX
        .SelHangingIndent = .SelTabs(0)
    End With
    
    With rtbWhispers
        .Font.Size = 8
        .SelTabCount = 1
        .SelTabs(0) = 15 * Screen.TwipsPerPixelX
        .SelHangingIndent = .SelTabs(0)
    End With
    
    lvChannel.View = lvwReport
    lvChannel.Icons = imlIcons
    lvClanList.View = lvwReport
    lvClanList.Icons = imlIcons
    
    ReDim Phrases(0)
    ReDim Catch(0)
    ReDim gBans(0)
    ReDim gOutFilters(0)
    ReDim gFilters(0)
    ReDim ServerRequests(0)
    ReDim g_Blocklist(0)
    
    Call BuildProductInfo
    
    Set Config = New clsConfig
    Config.Load GetConfigFilePath()
    
    ' SPLASH SCREEN
    If Config.ShowSplashScreen Then
        frmSplash.Show
        FrmSplashInUse = True
    End If
    
    If Config.ShowWhisperBox Then
        If Not rtbWhispersVisible Then Call cmdShowHide_Click
    Else
        If rtbWhispersVisible Then Call cmdShowHide_Click
    End If
    
    If Config.PositionHeight > 0 Then
        HeightVal = (IIf(CLng(Config.PositionHeight) < 200, 200, CLng(Config.PositionHeight)) * Screen.TwipsPerPixelY)
        
        If (rtbWhispersVisible) Then
            HeightVal = HeightVal - (rtbWhispers.Height / Screen.TwipsPerPixelY)
        End If
        
        Me.Height = HeightVal
    End If
    
    If Config.PositionWidth > 0 Then
        Me.Width = (IIf(CLng(Config.PositionWidth) < 300, 300, CLng(Config.PositionWidth)) * Screen.TwipsPerPixelX)
    End If

    'Set window position
    Me.Left = CLng(Config.PositionLeft) * Screen.TwipsPerPixelX
    Me.Top = CLng(Config.PositionTop) * Screen.TwipsPerPixelY

    'Make sure the window is on the screen
    If Config.EnforceScreenBounds Then
        If Config.MonitorCount <> GetMonitorCount Then
            If (Me.Left > (Screen.Width - Me.Width)) Then
                Me.Left = (Screen.Width - Me.Width)
            End If
    
            If (Me.Top > (Screen.Height - Me.Height)) Then
                Me.Top = (Screen.Height - Me.Height)
            End If
            
            Config.MonitorCount = GetMonitorCount
        End If
    End If
    
    'Support for recording maxmized position. - FrOzeN
    If Config.IsMaximized Then
        Me.WindowState = vbMaximized
    End If

    Set ClanHandler = New clsClanPacketHandler
    Set FriendListHandler = New clsFriendlistHandler
    Set ListToolTip = New clsCTooltip
    
    Set ReceiveBuffer(stBNCS) = New clsDataBuffer
    Set ReceiveBuffer(stBNLS) = New clsDataBuffer
    Set ReceiveBuffer(stMCP) = New clsDataBuffer
    
    Call ReloadConfig
    
    Call Form_Resize
    
    Call GetColorLists
    Call InitListviewTabs
    Call DisableListviewTabs
    
    ListviewTabs.Tab = 0
    
    With ListToolTip
        .Style = TTStandard
        .Icon = TTNoIcon
        .DelayTime = 100
    End With
    
    Call ClearChannel
    
    SetTitle "Disconnected"
    
    frmChat.UpdateTrayTooltip
    
    UserCancelledConnect = True
    
    Me.Show
    Me.Refresh
    
    AddChat RTBColors.ConsoleText, "-> Welcome to " & CVERSION & ", by Stealth."
    AddChat RTBColors.ConsoleText, "-> If you enjoy StealthBot, consider supporting its development at http://donate.stealthbot.net"

    Dim x As Integer
    For x = LBound(sStr) To UBound(sStr)
        If (LenB(sStr(x)) > 0) Then
            AddChat RTBColors.InformationText, sStr(x)
        End If
    Next x
    
    On Error Resume Next
    
    VoteDuration = -1
    
    If (LenB(Dir$(GetConfigFilePath())) = 0) Then
        AddChat RTBColors.ServerInfoText, "If you're new to bots, start by choosing 'Bot Settings' " & _
            "under the 'Settings' menu above."
        AddChat RTBColors.ServerInfoText, "For more help, click the 'Step-By-Step Configuration' " & _
            "button inside Settings."
        AddChat RTBColors.ServerInfoText, "For more information and a list of commands, see the " & _
            "Readme by clicking 'Readme' under the 'Help' menu."
        AddChat RTBColors.ServerInfoText, "Please note that any usage of this program is subject to " & _
            "the terms of the End-User License Agreement available at http://eula.stealthbot.net."
    End If
    
    
    Randomize
 
    ID_TASKBARICON = (Rnd * 100)
    
    TASKBARCREATED_MSGID = RegisterWindowMessage("TaskbarCreated")
    
    cboSend.SetFocus
    
    LoadQuickChannels
    PrepareHomeChannelMenu
    PreparePublicChannelMenu
    
    InitScriptControl SControl
    
    On Error Resume Next
    'News call and scripting events
    
    LoadScripts
    'InitMenus
    InitScripts
    
    If FrmSplashInUse Then frmSplash.SetFocus
    
    If Not MDebug("debug") Then
        mnuRecordWindowPos.Visible = False
    End If
    
    'Update the config if it's an old version
    If Config.Version < CONFIG_VERSION Then
        Call Config.Save
    End If
    
    #If COMPILE_DEBUG = 0 Then
        If Config.MinimizeOnStartup Then
            frmChat.WindowState = vbMinimized
            Call Form_Resize
        End If
    #End If
    
    If Not Config.DisableNews Then
        Call RequestInetPage(GetNewsURL(), SB_INET_NEWS1, True)
    ElseIf Config.AutoConnect Then
        Call DoConnect
    End If
    
    'Now loads scripts when the bot opens, instead of after connecting. - FrOzeN
    'RunInAll "Event_Load"
    
    'Dim I As Integer
    'Dim tmp As String
    'Dim str As String
    
    'str = "flood"
    
    'For I = 1 To Len(str)
    '   tmp = tmp & Hex(Asc(Mid(str, I, 1)))
    'Next I

'    BotVars.UseProxy = True
'    BotVars.ProxyIP = "213.210.194.139"
'    BotVars.ProxyPort = 1080
    'BotVars.ProxyIsSocks5 = True
End Sub

Public Sub cacheTimer_Timer()
    ' this code updated 7/23/05 in Chihuahua, Chihuahua, MX
    If (Caching) Then ' time to retrieve stored information and squelch or ban a channel
        Dim strArray() As String
        Dim ret As String
        Dim lPos As Long
        Dim y As String
        Dim c As Integer, n As Integer
        
        Caching = False
        
        ' Changed 08-18-09 - Hdx - Uses the new Channel cache function, Eventually to beremoved to script
        'ret = CacheChannelList(vbNullString, 0, Y)
        ret = CacheChannelList(enRetrieve, y)
        
        lPos = InStr(1, ret, ", ", vbBinaryCompare)
        
        If (lPos) Then
            strArray() = Split(ret, ", ")
        Else
            ReDim Preserve strArray(0)
            strArray(0) = ret
        End If
        
        For c = 0 To UBound(strArray)
            ' [CHANNELOP]  -  [*CHANNELOP]  -  [CHARACTER@USEast (*CHANNELOP)]
            If StrComp(UCase(strArray(c)), strArray(c), vbBinaryCompare) = 0 Then
                If Left$(strArray(c), 1) = "[" And Right$(strArray(c), 1) = "]" Then
                    strArray(c) = Mid(strArray(c), 2, Len(strArray(c)) - 2)
                End If
            End If
            
            strArray(c) = ConvertUsername(CleanUsername(strArray(c)))
            
            'AddChat vbRed, strArray(C)
            
            If Len(strArray(c)) > 1 Then
                If InStr(y, "ban") Then
                    If (g_Channel.Self.IsOperator) Then
                        Ban strArray(c), (AutoModSafelistValue - 1), 0
                    End If
                Else
                    If (GetSafelist(strArray(c)) = False) Then
                        AddQ "/squelch " & strArray(c)
                    End If
                End If
            End If
        Next c
    End If
    
    cacheTimer.Enabled = False
End Sub

Private Sub ChatQueueTimer_Timer()
    modChatQueue.ChatQueueTimerProc
End Sub

Private Sub Form_GotFocus()
    On Error GoTo ERROR_HANDLER

    If (cboSendHadFocus) Then
        cboSend.SetFocus
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Form_GotFocus()."

    Exit Sub
End Sub

' asynchronous Inet
Private Function RequestInetPage(ByVal URL As String, ByVal Request As String, ByVal CancelStillExecuting As Boolean) As Boolean
    On Error GoTo ERROR_HANDLER:

    Dim ret As String
    With Inet
        If .StillExecuting Then
            If CancelStillExecuting Then
                .Cancel
            Else
                RequestInetPage = False
                
                Exit Function
            End If
        End If
        
        .RequestTimeout = 5
        .Tag = Request
        .Execute URL
        
        RequestInetPage = True
    End With
    
    Exit Function
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in RequestInetPage()."
    
    RequestInetPage = False
    
    Exit Function
End Function

' asynchronous Inet response
Private Sub Inet_StateChanged(ByVal State As Integer)
    Dim strData As String
    Dim Buffer As String
    
    Select Case State
        Case icResponseCompleted, icError
            If Inet.ResponseCode >= 1000 Then
                Buffer = "Inet Error #" & Inet.ResponseCode & ": " & Inet.ResponseInfo
            ElseIf Inet.ResponseCode <> 0 Then
                Buffer = "HTTP Error " & Inet.ResponseCode & " " & Inet.ResponseInfo
            Else
                Do
                    strData = Inet.GetChunk(1024, icString)
                    If Len(strData) = 0 Then Exit Do
                    Buffer = Buffer & strData
                Loop
                
                If LenB(Buffer) = 0 Then
                    Buffer = "Empty response"
                End If
            End If
            
            Select Case Inet.Tag
                Case SB_INET_NEWS
                    Call HandleNews(Buffer, Inet.ResponseCode)
                Case SB_INET_NEWS1
                    Call HandleNews(Buffer, Inet.ResponseCode)
                    If Config.AutoConnect And Not g_Connected Then
                        Call DoConnect
                    End If
                Case SB_INET_VBYTE
                    Call HandleUpdateVerbyte(Buffer, Inet.ResponseCode)
                Case SB_INET_BNLS1
                    Call HandleFindBNLSServerListResult(Buffer, Inet.ResponseCode, True)
                Case SB_INET_BNLS2
                    Call HandleFindBNLSServerListResult(Buffer, Inet.ResponseCode, False)
            End Select
            
            Inet.Tag = SB_INET_UNSET
            Inet.Cancel
    End Select
End Sub

Private Sub HandleUpdateVerbyte(ByVal Buffer As String, ByVal ResponseCode As Long)
    On Error Resume Next
    
    Dim ary() As String
    Dim i As Integer
    
    If Inet.ResponseCode <> 0 Then
        AddChat RTBColors.ErrorMessageText, Buffer & ". Error retrieving version bytes from http://www.stealthbot.net. Please visit it for instructions."
    ElseIf Len(Buffer) <> 11 Then
        AddChat RTBColors.ErrorMessageText, "Format not understood. Error retrieving version bytes from http://www.stealthbot.net. Please visit it for instructions."
    Else
        'W2 SC D2 W3
        Dim Keys(3) As String
        
        Keys(0) = "W2"
        Keys(1) = "SC"
        Keys(2) = "D2"
        Keys(3) = "W3"
        
        ary() = Split(Buffer, " ")
        
        For i = 0 To 3
            Config.SetVersionByte Keys(i), CLng(Val("&H" & ary(i)))
        Next i
        Config.SetVersionByte "D2X", CLng(Val("&H" & ary(2)))
        
        Call Config.Save
        
        AddChat RTBColors.SuccessText, "Your config.ini file has been loaded with current version bytes."
    End If
End Sub


'RTB ADDCHAT SUBROUTINE - originally written by Grok[vL] - modified to support
'                         logging and timestamps and color decoding
' Updated 7/23/05 to remove many bulky calls to Len()
' Updated 9/01/05 to remove the changes made on 7/23/05 *smack forehead*
' Updated 9/25/05-10/25/05 to add HTML logging
' Updated 1/3/06 to remove HTML logging
' Updated 8/4/06 to add scrollbar locking (thanks FrOzeN)
' Updated 11/8/06 to log incoming text immediately
' Updated 4/17/07 to not flash the desktop when the scrollbar is held up
' Updated 8/07/07 with greater precision
Sub AddChat(ParamArray saElements() As Variant)
    Dim arr() As Variant
    Dim i     As Integer

    arr() = saElements
    Call DisplayRichText(frmChat.rtbChat, arr)
End Sub

Sub AddWhisper(ParamArray saElements() As Variant)
    Dim arr() As Variant
    
    arr() = saElements
    
    Call DisplayRichText(frmChat.rtbWhispers, arr)
    Exit Sub
    
    
    Dim s As String
    Dim l As Long
    Dim i As Integer
    
    If Not BotVars.LockChat Then
        'If ((BotVars.MaxBacklogSize) And (Len(rtbWhispers.text) >= BotVars.MaxBacklogSize)) Then
            If BotVars.Logging < 2 Then
                Close #1
                Open (StringFormat("{0}{1}-WHISPERS.txt", GetFolderPath("Logs"), Format(Date, "YYYY-MM-DD"))) For Append As #1
            End If
            
            With rtbWhispers
                .Visible = False
                .SelStart = 0
                .SelLength = InStr(1, .Text, vbLf, vbBinaryCompare)
                If BotVars.Logging < 2 Then Print #1, Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                .SelText = vbNullString
                .Visible = True
            End With

            Close #1
        'End If
        
        Select Case BotVars.TSSetting
            Case 0: s = " [" & Time & "] "
            Case 1: s = " [" & Format(Time, "HH:MM:SS") & "] "
            Case 2: s = " [" & Format(Time, "HH:MM:SS") & "." & GetCurrentMS & "] "
            Case 3: s = vbNullString
        End Select
        
        With rtbWhispers
            .SelStart = Len(.Text)
            .SelLength = 0
            .SelColor = RTBColors.TimeStamps
            If .SelBold = True Then .SelBold = False
            If .SelItalic = True Then .SelItalic = False
            .SelText = s
            .SelStart = Len(.Text)
        End With
        
        For i = LBound(saElements) To UBound(saElements) Step 2
            If InStr(1, saElements(i), vbNullChar, vbBinaryCompare) > 0 Then _
                KillNull saElements(i)
            
            If Len(saElements(i + 1)) > 0 Then
                With rtbWhispers
                    .SelStart = Len(.Text)
                    l = .SelStart
                    .SelLength = 0
                    .SelColor = saElements(i)
                    .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                    .SelStart = Len(.Text)
                End With
            End If
        Next i
        
        Call ColorModify(rtbWhispers, l)
    End If
End Sub


'BNLS EVENTS
Sub Event_BNetConnected()
    If (ProxyConnInfo(stBNCS).IsUsingProxy) Then
        AddChat RTBColors.SuccessText, "[BNCS] [PROXY] Connected!"
    Else
        AddChat RTBColors.SuccessText, "[BNCS] Connected!"
    End If
    
    Call SetNagelStatus(sckBNet.SocketHandle, False)
End Sub

Sub Event_BNetConnecting()
    If (ProxyConnInfo(stBNCS).IsUsingProxy) Then
        AddChat RTBColors.InformationText, "[BNCS] [PROXY] Connecting to the SOCKS" & ProxyConnInfo(stBNCS).Version & " proxy server at " & ProxyConnInfo(stBNCS).ProxyIP & ":" & ProxyConnInfo(stBNCS).ProxyPort & "..."
    Else
        AddChat RTBColors.InformationText, "[BNCS] Connecting to the Battle.net server at " & BotVars.Server & "..."
    End If
End Sub

Sub Event_BNetDisconnected()
    tmrIdleTimer.Enabled = False
    UpTimer.Enabled = False
    ConnectionTickCount = 0@
    BotVars.JoinWatch = 0
    
    AddChat RTBColors.ErrorMessageText, "[BNCS] Disconnected."
    
    Call DoDisconnect(True)
    
    SetTitle "Disconnected"
    
    UpdateTrayTooltip
    
    g_Online = False
    
    Call ClearChannel
    
    ProxyConnInfo(stBNCS).Status = psNotConnected
    ProxyConnInfo(stBNLS).Status = psNotConnected
    ProxyConnInfo(stMCP).Status = psNotConnected
    'AddChat RTBColors.ErrorMessageText, "[BNCS] Attempting to reconnect, please wait..."
    'AddChat RTBColors.SuccessText, "Connection initialized."
    
    If sckBNet.State <> sckClosed Then sckBNet.Close
    If sckBNLS.State <> sckClosed Then sckBNLS.Close
    
    BNLSAuthorized = False
    
    'If Not UserCancelledConnect Then
    '    ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
    '        AddressOf Reconnect_TimerProc)
    'End If
End Sub

Sub Event_BNetError(ErrorNumber As Integer, Description As String)
    Dim s As String
    Dim IsProxyConnecting As Boolean
    
    If ProxyConnInfo(stBNCS).IsUsingProxy And ProxyConnInfo(stBNCS).Status <> psOnline Then
        s = "[BNCS] [PROXY] "
    Else
        s = "[BNCS] "
    End If
    
    AddChat RTBColors.ErrorMessageText, s & ErrorNumber & " -- " & Description
    AddChat RTBColors.ErrorMessageText, s & "Disconnected."
    
    If (sckBNet.State <> sckClosed) Then
        Call sckBNet.Close
    End If
    
    If (sckBNLS.State <> sckClosed) Then
        Call sckBNLS.Close
    End If
    
    If (sckMCP.State <> sckClosed) Then
        Call sckMCP.Close
    End If
    
    g_Connected = False
    
    UserCancelledConnect = False
    
    DoDisconnect 1, True
    
    SetTitle "Disconnected"
    
    frmChat.UpdateTrayTooltip
    
    Call ClearChannel
    lvClanList.ListItems.Clear
    lvFriendList.ListItems.Clear
    
    Call UpdateListviewTabs
    
    ' NOV 18 04 Change here should fix the attention-grabbing on errors
    'If Me.WindowState <> vbMinimized Then cboSend.SetFocus
    
    IsProxyConnecting = ProxyConnInfo(stBNCS).IsUsingProxy And ProxyConnInfo(stBNCS).Status <> psOnline
    If DisplayError(ErrorNumber, IsProxyConnecting, stBNCS) = True Then
        AddChat RTBColors.ErrorMessageText, _
            "[BNCS] Attempting to reconnect in " & (BotVars.ReconnectDelay / 1000) & _
                IIf(((BotVars.ReconnectDelay / 1000) <> 1), " seconds", " second") & _
                        "..."
        
        UserCancelledConnect = False 'this should fix the beta reconnect problems
        
        'ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
        '    AddressOf Reconnect_TimerProc)
        
        'ExReconTicks = 0
        'ExReconMinutes = BotVars.ReconnectDelay / 1000
        'ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, _
        '    1000, AddressOf ExtendedReconnect_TimerProc)
    End If
End Sub

Sub Event_BNLSAuthEvent(Success As Boolean)
    If Success = True Then
        AddChat RTBColors.SuccessText, "[BNLS] Authorized!"
    Else
        AddChat RTBColors.ErrorMessageText, "[BNLS] Authorization failed! Please download the latest version of StealthBot from http://www.stealthbot.net."
        Call DoDisconnect
    End If
End Sub

Sub Event_BNLSConnected()
    If (ProxyConnInfo(stBNLS).IsUsingProxy) Then
        AddChat RTBColors.SuccessText, "[BNLS] [PROXY] Connected!"
    Else
        AddChat RTBColors.SuccessText, "[BNLS] Connected!"
    End If
End Sub

Sub Event_BNLSConnecting()
    If (ProxyConnInfo(stBNLS).IsUsingProxy) Then
        AddChat RTBColors.InformationText, "[BNLS] [PROXY] Connecting to the SOCKS" & ProxyConnInfo(stBNLS).Version & " proxy server at " & ProxyConnInfo(stBNLS).ProxyIP & ":" & ProxyConnInfo(stBNLS).ProxyPort & "..."
    Else
        AddChat RTBColors.InformationText, "[BNLS] Connecting to the BNLS server at " & BotVars.BNLSServer & "..."
    End If
End Sub

Sub Event_BNLSDataError(Message As Byte)
    If Message = 0 Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] Your CD-Key was rejected. It may be invalid. Try connecting again."
    ElseIf Message = 1 Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] Error! Your CD-Key is bad."
    ElseIf Message = 2 Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] Error! BNLS has failed CheckRevision. Please check your bot's settings and try again."
        AddChat RTBColors.ErrorMessageText, "[BNLS] Product: " & StrReverse(BotVars.Product) & "."
    ElseIf Message = 3 Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] Error! Bad NLS revision."
    End If
End Sub

Private Sub Event_BNLSError(ErrorNumber As Integer, Description As String)
    If ProxyConnInfo(stBNLS).IsUsingProxy And ProxyConnInfo(stBNLS).Status <> psOnline Then
        DisplayError ErrorNumber, True, stBNLS
    ElseIf Not HandleBnlsError("[BNLS] Error " & ErrorNumber & ": " & Description) Then
        ' if we aren't using the finder display the error
        DisplayError ErrorNumber, False, stBNLS
    End If
End Sub


' this function will return whether we are going to use the finder
Public Function HandleBnlsError(ByVal ErrorMessage As String) As Boolean
    HandleBnlsError = False
    
    sckBNet.Close
    
    ' Is the BNLS server finder enabled?
    If Config.BNLSFinder Then
        Call RotateBnlsServer
    Else
        AddChat RTBColors.ErrorMessageText, ErrorMessage
        
        UserCancelledConnect = False
        DoDisconnect 1, True
    End If
    
    ' return the BotVars
    HandleBnlsError = Config.BNLSFinder
End Function

' Moves the connection to the next available BNLS server
Public Sub RotateBnlsServer()
    'Close the current BNLS connection
    sckBNLS.Close
    
    'Notify user the current BNLS server failed
    AddChat RTBColors.ErrorMessageText, "[BNLS] Connection to " & BotVars.BNLSServer & " failed."
    
    'Notify user other BNLS servers are being located
    AddChat RTBColors.InformationText, "[BNLS] Locating other BNLS servers..."
    
    Call FindBNLSServer
End Sub

Public Sub HandleFindBNLSServerListResult(ByVal strReturn As String, ByVal Result As Integer, ByVal ConfigListSource As Boolean)
    ' convert to LF
    strReturn = Replace(strReturn, vbCr, vbLf)
    strReturn = Replace(strReturn, vbLf & vbLf, vbLf)
    
    If (Inet.ResponseCode <> 0) Or (Right$(strReturn, 1) <> vbLf) Then
        If ConfigListSource Then
            If Not RequestInetPage(BNLS_DEFAULT_SOURCE, SB_INET_BNLS2, True) Then
                Call HandleFindBNLSServerListResult("Inet is busy", -1, False)
            End If
        Else
            AddChat RTBColors.ErrorMessageText, "[BNLS] " & strReturn & ". Unable to use BNLS server finder."
            AddChat RTBColors.ErrorMessageText, "[BNLS] An error occured while trying to locate an alternative BNLS server."
            AddChat RTBColors.ErrorMessageText, "[BNLS]   You may not be connected to the internet or may be having DNS resolution issues."
            AddChat RTBColors.ErrorMessageText, "[BNLS]   Visit http://www.stealthbot.net/ and check the Technical Support forum for more information."
            DoDisconnect
    
            ' ensure that we update our listing on following connection(s)
            BNLSFinderGotList = False
            
            ' ensure checker starts at 0 again on following connection(s)
            BNLSFinderIndex = 0
        End If
        
        Exit Sub
    Else
        ' Split the page up into an array of servers.
        BNLSFinderEntries() = Split(strReturn, vbLf)
    End If
    
    'Mark GotBNLSList as True so it's no longer downloaded for each attempt
    BNLSFinderGotList = True
    
    Call FindBNLSServerEntry
End Sub

'Locates alternative BNLS servers for the bot to use if the current one fails
Public Sub FindBNLSServer()
    'Error handler
    On Error GoTo ERROR_HANDLER
    
    BNLSFinderIndex = BNLSFinderIndex + 1
    
    'Check if the BNLS list has been downloaded
    If (BNLSFinderGotList = False) Then
        'Reset the counter
        BNLSFinderIndex = 0
        
        ' store first bnls server used so that we can avoid connecting to it again
        BNLSFinderLatest = BotVars.BNLSServer
        
        'Get the servers as a list from http://stealthbot.net/p/bnls.php
        If (LenB(Config.BNLSFinderSource) > 0) Then
            If Not RequestInetPage(Config.BNLSFinderSource, SB_INET_BNLS1, True) Then
                Call HandleFindBNLSServerListResult("Inet is busy", -1, False)
            End If
        Else
            If Not RequestInetPage(BNLS_DEFAULT_SOURCE, SB_INET_BNLS2, True) Then
                Call HandleFindBNLSServerListResult("Inet is busy", -1, False)
            End If
        End If
        
        Exit Sub
    End If
    
    Call FindBNLSServerEntry
    
    Exit Sub
    
ERROR_HANDLER:

    'Display the error message to the user
    If Err.Number = ERROR_FINDBNLSSERVER Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] " & Err.Description
        AddChat RTBColors.ErrorMessageText, "[BNLS]   Visit http://www.stealthbot.net/ and check the Technical Support forum for more information."
        DoDisconnect
        
        ' ensure that we update our listing on following connection(s)
        BNLSFinderGotList = False
        
        ' ensure checker starts at 0 again on following connection(s)
        BNLSFinderIndex = 0
    
    Else
        Resume Next
    End If

    Exit Sub
End Sub

Sub FindBNLSServerEntry()
    If BNLSFinderIndex > UBound(BNLSFinderEntries) Then
        'All BNLS servers have been tried and failed
        Err.Raise ERROR_FINDBNLSSERVER, , "All the BNLS servers have failed."
    End If
    
    ' keep increasing counter until we find a server that is valid and isn't the same as the first one
    Do While (StrComp(BNLSFinderEntries(BNLSFinderIndex), BNLSFinderLatest, vbTextCompare) = 0) Or (LenB(BNLSFinderEntries(BNLSFinderIndex)) = 0)
        BNLSFinderIndex = BNLSFinderIndex + 1
        
        If BNLSFinderIndex > UBound(BNLSFinderEntries) Then
            'All BNLS servers have been tried and failed
            Err.Raise ERROR_FINDBNLSSERVER, , "All the BNLS servers have failed."
            Exit Do
        End If
    Loop
    
    BotVars.BNLSServer = BNLSFinderEntries(BNLSFinderIndex)
    
    ConnectBNLS
End Sub

' Updated 8/8/07 to support new prefix/suffix box feature
Public Sub Form_Resize()
    On Error Resume Next
    
    Dim lblHeight As Integer
    Static WasMaximized As Boolean
    Static DoMaximize As Boolean
    
    If Me.WindowState = vbMinimized Then
        If Not BotVars.NoTray Then
            #If (COMPILE_DEBUG <> 1) Then
                Me.Hide
                
                With nid
                    .cbSize = LenB(nid)
                    .hWnd = frmChat.hWnd
                    .uId = ID_TASKBARICON
                    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                    .uCallBackMessage = WM_ICONNOTIFY
                    .hIcon = frmChat.Icon.handle
                    .szTip = GenerateTooltip()
                End With
                
                Shell_NotifyIcon NIM_ADD, nid
            #End If
        End If
    Else
        Shell_NotifyIcon NIM_DELETE, nid
        cboSend.SetFocus
        
        If txtPre.Visible Then
            txtPre.Height = cboSend.Height
            txtPre.Width = txtPost.Width
        End If
        
        If txtPost.Visible Then
            txtPost.Height = cboSend.Height
        End If
        
        'sizing + positioning
        
        'Added IsWindowsVista() call within an IIf() statement.
        'This shrinks the size of the entire layout by a further 80 twips.
        'This will act as a fix for Vista's screen cut off issues.
        '   - FrOzeN
        'This issue only occured under Aero, so the fix breaks the GUI for classic themes
        'on vista. This should be fixed now. ~Pyro
        With lvChannel
            'rtbChat.Width = Me.Width - .Width - IIf(g_OSVersion.IsWindowsVista, 200, 120)
        
            rtbChat.Width = Me.ScaleWidth - .Width
        
        '    .Width = (Me.Width / 4) - 120 'magic number?
        '    If .Width > (.ColumnHeaders.Item(1).Width + 700) Then
        '        .Width = .ColumnHeaders.Item(1).Width + 700
        '
        '        rtbChat.Width = Me.Width - .Width - 120
        '    Else
        '        rtbChat.Width = ((Me.Width / 4) * 3)
        '    End If
        '
        '    .ColumnHeaders.Item(1).Width = (.Width / 3) * 2.5
        End With
        
        lblCurrentChannel.Width = lvChannel.Width
        lvFriendList.Width = lvChannel.Width
        lvClanList.Width = lvChannel.Width
        cboSend.Width = rtbChat.Width
        
        With cmdShowHide
            If rtbWhispersVisible Then
                'Debug.Print "-> " & rtbWhispers.Height
                .Height = rtbWhispers.Height + 285
                .Caption = CAP_HIDE
                .ToolTipText = TIP_HIDE
            Else
                .Height = txtPre.Height - Screen.TwipsPerPixelY
                .Caption = CAP_SHOW
                .ToolTipText = TIP_SHOW
            End If
            
            .ZOrder vbBringToFront
        End With
        
        rtbWhispers.Visible = rtbWhispersVisible
        
        'height is based on rtbchat.height + cmdshowhide.height
        If rtbWhispersVisible Then
            rtbChat.Height = ((Me.ScaleHeight / Screen.TwipsPerPixelY) - (txtPre.Height / _
                Screen.TwipsPerPixelY) - (rtbWhispers.Height / Screen.TwipsPerPixelY)) * _
                    (Screen.TwipsPerPixelY)
            rtbWhispers.Move rtbChat.Left, cboSend.Top + cboSend.Height
        Else
            rtbChat.Height = ((Me.ScaleHeight / Screen.TwipsPerPixelY) - (txtPre.Height / _
                Screen.TwipsPerPixelY)) * (Screen.TwipsPerPixelY)
        End If
        
        lvChannel.Move rtbChat.Left + rtbChat.Width, lblCurrentChannel.Top + lblCurrentChannel.Height, lvChannel.Width, rtbChat.Height - lblCurrentChannel.Height
        lvFriendList.Move lvChannel.Left, lvChannel.Top, lvChannel.Width, rtbChat.Height - lblCurrentChannel.Height
        lvClanList.Move lvChannel.Left, lvChannel.Top, lvChannel.Width, rtbChat.Height - lblCurrentChannel.Height
        lblCurrentChannel.Move lvChannel.Left, rtbChat.Top
        
        If txtPre.Visible Then
            txtPre.Move rtbChat.Left, rtbChat.Top + rtbChat.Height + (Screen.TwipsPerPixelY / 3)
            cboSend.Move txtPre.Left + txtPre.Width, txtPre.Top, rtbChat.Width - txtPre.Width
        Else
            cboSend.Move rtbChat.Left, rtbChat.Top + rtbChat.Height + (Screen.TwipsPerPixelY / 3)
        End If
        
        If txtPost.Visible Then
            cboSend.Width = cboSend.Width - txtPost.Width
            txtPost.Move cboSend.Left + cboSend.Width, cboSend.Top
        End If
        
        lvChannel.Height = rtbChat.Height - lblCurrentChannel.Height
        lvFriendList.Height = lvChannel.Height
        lvClanList.Height = lvChannel.Height
        
        'Minus 80 twips from rtbWhispers.Width if using Vista to fix width issue
        'the issue is not with Vista, but with Aero.
        With rtbWhispers
            If .Visible Then
                .Move rtbChat.Left, cboSend.Top + cboSend.Height, (Me.ScaleWidth - cmdShowHide.Width - Screen.TwipsPerPixelX)
            End If
        End With
        
        ListviewTabs.Height = cboSend.Height
        ListviewTabs.Move lvChannel.Left, cboSend.Top + Screen.TwipsPerPixelY, lvChannel.Width - _
            cmdShowHide.Width - Screen.TwipsPerPixelX, cboSend.Height '+ 2 * Screen.TwipsPerPixelY
        
        If rtbWhispersVisible Then
            cmdShowHide.Move (((rtbWhispers.Left + rtbWhispers.Width) / Screen.TwipsPerPixelX) + 1) * _
                Screen.TwipsPerPixelX, lvChannel.Top + lvChannel.Height + Screen.TwipsPerPixelY
        Else
            cmdShowHide.Move ListviewTabs.Left + ListviewTabs.Width, lvChannel.Top + lvChannel.Height
        End If
        
        With lvChannel
            If Config.HideClanDisplay Then
                .ColumnHeaders(1).Width = (.Width \ 4) * 3 + 200
                .ColumnHeaders(2).Width = 0
            Else
                .ColumnHeaders(1).Width = (.Width \ 4) * 3 - 500
                .ColumnHeaders(2).Width = 700
            End If
            If Config.HidePingDisplay Then
                .ColumnHeaders(3).Width = 0
            Else
                .ColumnHeaders(3).Width = imlIcons.ImageWidth
            End If
        End With
        
        With lvFriendList
            .ColumnHeaders(1).Width = (.Width \ 4) * 3 + 200
            .ColumnHeaders(2).Width = imlIcons.ImageWidth '.Width \ 4 + 200
        End With
        
        With lvClanList
            .ColumnHeaders(1).Width = (.Width \ 4) * 3 - 200
            .ColumnHeaders(2).Width = imlClan.ImageWidth '.Width \ 4 + 200
        End With
    End If
    
    If Me.WindowState = vbMaximized Then
        WasMaximized = True
        Call RecordWindowPosition(True)
    ElseIf Me.WindowState = vbMinimized Then
        If WasMaximized Then
            WasMaximized = False
            DoMaximize = True
        End If
    Else
        WasMaximized = False
        
        If DoMaximize Then
            DoMaximize = False
            
            Me.WindowState = vbMaximized
        End If
    End If
    
    Call rtbChat.Refresh
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in Form_Resize()."
End Sub

Function GenerateTooltip() As String
    GenerateTooltip = String(64, vbNullChar)
    GenerateTooltip = IIf(LenB(GetCurrentUsername) > 0, GetCurrentUsername, "offline") & " @ " & BotVars.Server & " (" & StrReverse(BotVars.Product) & ")" & Chr$(0)
End Function

Sub UpdateTrayTooltip()
    On Error Resume Next

    If Me.WindowState = vbMinimized Then
        With nid
            .cbSize = LenB(nid)
            .hWnd = frmChat.hWnd
            .uId = ID_TASKBARICON
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_ICONNOTIFY
            .hIcon = frmChat.Icon.handle
            .szTip = GenerateTooltip()
        End With
        
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Sub

Private Sub ClanHandler_CandidateListReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue, ByRef Users() As String)
    Dim i As Long
    Dim oRequest As udtServerRequest

    Call FindServerRequest(oRequest, Cookie)

    'Valid Status codes:
    '   0x00: Successfully found candidate(s)
    '   0x01: Clan tag already taken
    '   0x08: Already in clan
    '   0x0a: Invalid clan tag specified
    
    If MDebug("debug") Then
        AddChat RTBColors.ErrorMessageText, "CandidateList received. Status code [0x" & ZeroOffset(Result, 2) & "]."
        If UBound(Users) > -1 Then
            AddChat RTBColors.InformationText, "Potential clan members:"
            
            For i = 0 To UBound(Users)
                AddChat RTBColors.InformationText, Users(i)
            Next i
        End If
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanCandidateList", Result, ConvertStringArray(Users)
    End If
End Sub

Private Sub Clanhandler_InviteMultipleReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue, ByRef Users() As String)
    Dim oRequest As udtServerRequest

    Call FindServerRequest(oRequest, Cookie)

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanInviteMultiple", Result, ConvertStringArray(Users)
    End If
End Sub

Private Sub ClanHandler_MemberLeaves(ByVal Member As String)

    If Not g_Clan.InClan Then Exit Sub

    AddChat RTBColors.JoinText, "[CLAN] ", RTBColors.JoinUsername, Member, RTBColors.JoinText, " has left the clan."

    Dim x   As ListItem
    Dim pos As Integer

    pos = g_Clan.GetUserIndexEx(Member)

    If (pos > 0) Then
        g_Clan.Members.Remove pos
    End If

    Member = ConvertUsername(Member)

    Set x = lvClanList.FindItem(Member)

    If (Not (x Is Nothing)) Then
        lvClanList.ListItems.Remove x.Index
        
        lvClanList.Refresh
        
        Set x = Nothing
    End If

    On Error Resume Next

    RunInAll "Event_ClanMemberLeaves", Member
End Sub

Private Sub ClanHandler_RemovedFromClan(ByVal Status As Byte)
    Dim oRequest As udtServerRequest

    If Not g_Clan.InClan Then Exit Sub

    If Status = 1 Then
        If Not FindServerRequest(oRequest, -1, SID_CLANREMOVEMEMBER, , False) Then
            ' no pending SID_CLANREMOVEMEMBER (self leaving), mention it
            AddChat RTBColors.ErrorMessageText, "[CLAN] You have been removed from the clan, or it has been disbanded."
        End If

        Set g_Clan = New clsClanObj

        ListviewTabs.TabEnabled(LVW_BUTTON_CLAN) = False
        lvClanList.ListItems.Clear
        ListviewTabs.Tab = LVW_BUTTON_CHANNEL
        Call UpdateListviewTabs

        On Error Resume Next
        RunInAll "Event_BotRemovedFromClan"
    End If
End Sub

Private Sub ClanHandler_MyRankChange(ByVal OldRank As enuClanRank, ByVal NewRank As enuClanRank, ByVal Initiator As String)

    If Not g_Clan.InClan Then Exit Sub

    If (g_Clan.Self.Rank < NewRank) Then
        AddChat RTBColors.JoinText, "[CLAN] You have been promoted by ", _
                RTBColors.JoinUsername, Initiator, RTBColors.JoinText, ". Your new rank is ", _
                RTBColors.JoinUsername, ClanHandler.GetRankName(NewRank), RTBColors.JoinText, "."
    Else
        AddChat RTBColors.JoinText, "[CLAN] You have been demoted by ", _
                RTBColors.JoinUsername, Initiator, RTBColors.JoinText, ". Your new rank is ", _
                RTBColors.JoinUsername, ClanHandler.GetRankName(NewRank), RTBColors.JoinText, "."
    End If

    g_Clan.Self.Rank = NewRank

    On Error Resume Next

    RunInAll "Event_BotClanRankChanged", NewRank
End Sub

Private Sub ClanHandler_Info(ByVal ClanTag As String, ByVal Rank As enuClanRank)
    Dim oRequest As udtServerRequest

    Set g_Clan = New clsClanObj

    With g_Clan
        .InClan = True
        .PendingInvitation = False
        .PendingClanMOTD = False
        .Name = ClanTag
    End With

    Call InitListviewTabs
    
    'If g_Clan.Self.Rank = 0 Then g_Clan.Self.Rank = 1
    
    BotVars.Clan = ClanTag

    If FindServerRequest(oRequest, -1, SID_CLANINVITATION, , False) Then
        AddChat RTBColors.JoinText, "[CLAN] You are now a member of ", RTBColors.JoinUsername, "Clan " & ClanTag, RTBColors.JoinText, "!"
            
        RunInAll "Event_BotJoinedClan", ClanTag
    ElseIf FindServerRequest(oRequest, -1, SID_CLANCREATIONINVITATION, , False) Then
        AddChat RTBColors.JoinText, "[CLAN] You are now a member of the newly created ", RTBColors.JoinUsername, "Clan " & ClanTag, RTBColors.JoinText, "!"
            
        RunInAll "Event_BotJoinedClan", ClanTag
    Else
        AddChat RTBColors.JoinText, "[CLAN] You are a ", RTBColors.JoinUsername, ClanHandler.GetRankName(Rank), RTBColors.JoinText, " in ", _
                RTBColors.JoinUsername, "Clan " & ClanTag, RTBColors.JoinText, "."
        
        RunInAll "Event_BotClanInfo", ClanTag, Rank
    End If
    
    Call ClanHandler.RequestClanList(reqInternal)
    Call ClanHandler.RequestClanMOTD(reqInternal)
    
    Call UpdateListviewTabs
End Sub

Private Sub ClanHandler_InvitationReceived(ByVal Cookie As Long, ByVal ClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal IsNewClan As Boolean, ByRef Users() As String)
    Dim oRequest As udtServerRequest

    If Not Config.IgnoreClanInvites Then
        With oRequest
            .ResponseReceived = False
            .HandlerType = reqInternal
            Set .Command = Nothing
            .PacketID = SID_CLANINVITATIONRESPONSE
            .PacketCommand = 0
            .Tag = Array(Cookie, ClanTag, ClanName, InvitedBy, IsNewClan)
        End With

        Set g_Clan = New clsClanObj
        g_Clan.PendingInvitation = True
        g_Clan.PendingInvitationCookie = SaveServerRequest(oRequest)

        With RTBColors
            AddChat .SuccessText, "[CLAN] ", .InformationText, ConvertUsername(InvitedBy), _
                    .SuccessText, " has invited you to join ", .InformationText, "Clan " & ClanTag, _
                    .SuccessText, ": ", .InformationText, ClanName
        End With

        frmClanInvite.Show
    End If

    RunInAll "Event_ClanInvitation", Cookie, ClanTag, ClanName, InvitedBy, IsNewClan
End Sub

Private Sub ClanHandler_MemberUpdate(ByVal Member As clsClanMemberObj)
    Dim ListItem As ListItem
    Dim pos      As Integer
    Dim OldRank  As enuClanRank
    Dim Username As String
    Dim Index    As Integer

    If Not g_Clan.InClan Then Exit Sub

    pos = g_Clan.GetUserIndexEx(Member.Name)
    
    If (pos > 0) Then
        OldRank = g_Clan.Members(pos).Rank

        With Member
            g_Clan.Members(pos).Name = .Name
            g_Clan.Members(pos).Rank = .Rank
            g_Clan.Members(pos).Status = .Status
            g_Clan.Members(pos).Location = .Location
        End With
    Else
        g_Clan.Members.Add Member

        ' we didn't have a record of this user, assume this isn't a rank change...
        OldRank = Member.Rank
    End If

    If StrComp(Member.Name, BotVars.Username, vbTextCompare) = 0 Then
        g_Clan.Self.Rank = Member.Rank
    End If

    If Member.Rank <> OldRank Then
        With RTBColors
            AddChat .JoinText, "[CLAN] Member update: ", .JoinUsername, Member.DisplayName, _
                    .JoinText, " is now a ", .JoinUsername, Member.RankName, .JoinText, "."
        End With
    End If

    Set ListItem = lvClanList.FindItem(Member.Name)
    If Not (ListItem Is Nothing) Then
        ' set the icon and status in place
        SetClanMember ListItem, Member.DisplayName, Member.Rank, Member.Status
        Set ListItem = Nothing
    Else
        ' wasn't found...
        AddClanMember Member.Name, Member.DisplayName, Member.Rank, Member.Status
    End If

    ' re-sort
    lvClanList.Sorted = True

    Call UpdateListviewTabs
    
    On Error Resume Next
    RunInAll "Event_ClanMemberUpdate", Member.Name, Member.Rank, Member.IsOnline
End Sub

Private Sub ClanHandler_GetMOTD(ByVal Cookie As Long, ByVal Message As String)
    Dim oRequest As udtServerRequest

    Call FindServerRequest(oRequest, Cookie)

    g_Clan.MOTD = Message

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanMOTD", Message
    End If
End Sub

Private Sub ClanHandler_GetMemberList(ByVal Cookie As Long, ByVal Members As Collection)
    Dim i As Long
    Dim oRequest As udtServerRequest
    Dim Member As clsClanMemberObj

    Call FindServerRequest(oRequest, Cookie)

    For i = 1 To Members.Count
        Set Member = Members.Item(i)
        g_Clan.Members.Add Member
        AddClanMember Member.Name, Member.DisplayName, Member.Rank, Member.Status

        If oRequest.HandlerType = reqScriptingCall Then
            On Error Resume Next
            RunInAll "Event_ClanMemberList", Member.DisplayName, Member.Rank, Member.Status
        End If
        Set Member = Nothing
    Next i

    ' re-sort
    lvClanList.Sorted = True
    
    Call UpdateListviewTabs
End Sub

Private Sub ClanHandler_GetMemberInfo(ByVal Cookie As Long, ByVal Result As enuClanResponseValue, ByVal ClanName As String, ByVal Rank As enuClanRank, ByVal JoinDate As Date)
    Dim i        As Long
    Dim oRequest As udtServerRequest
    Dim Username As String
    Dim ClanTag  As String
    Dim Member   As clsClanMemberObj
    Dim RespText As String

    Call FindServerRequest(oRequest, Cookie)

    Username = oRequest.Tag(0)
    ClanTag = oRequest.Tag(1)

    If (g_Clan.InClan And StrComp(g_Clan.Name, ClanTag, vbTextCompare) = 0) Then
        g_Clan.FullName = ClanName
        Set Member = g_Clan.GetMember(Username)
        If Not Member Is Nothing Then
            Member.JoinTime = JoinDate
            Member.Rank = Rank
        End If
        Set Member = Nothing

        Call UpdateListviewTabs
    End If

    If Result = clresSuccess Then
        RespText = StringFormat("{0} is a {1} in Clan {2}: {3} since {4}.", Username, ClanHandler.GetRankName(Rank), ClanTag, ClanName, JoinDate)
    Else
        RespText = StringFormat("Error: Get {0} member info failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanMemberInfo", Result, Username, ClanTag, ClanName, Rank, JoinDate
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond RespText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, RespText
    End If
End Sub

Private Sub ClanHandler_DemoteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String
    Dim Username As String

    Call FindServerRequest(oRequest, Cookie)
    Username = oRequest.Tag(0)

    If Result = clresSuccess Then
        ResponseText = StringFormat("{0} demoted successfully.", Username)
    Else
        ResponseText = StringFormat("Error: {0} demotion failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanDemoteUserReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_PromoteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String
    Dim Username As String

    Call FindServerRequest(oRequest, Cookie)
    Username = oRequest.Tag(0)

    If Result = clresSuccess Then
        ResponseText = StringFormat("{0} promoted successfully.", Username)
    Else
        ResponseText = StringFormat("Error: {0} promotion failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanPromoteUserReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_RemoveMemberReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String
    Dim Username As String

    Call FindServerRequest(oRequest, Cookie)
    Username = oRequest.Tag(0)

    If CBool(oRequest.Tag(1)) Then
        If Result = clresSuccess Then
            ResponseText = "Left the clan successfully."
        Else
            ResponseText = StringFormat("Error: Clan leave failed - {0}.", ClanHandler.GetClanResponseText(Result))
        End If
    Else
        If Result = clresSuccess Then
            ResponseText = StringFormat("{0} removed successfully.", Username)
        Else
            ResponseText = StringFormat("Error: {0} removal failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
        End If
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanRemoveUserReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_DisbandClanReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String

    Call FindServerRequest(oRequest, Cookie)

    If Result = clresSuccess Then
        ResponseText = "Clan disbanded successfully."
    Else
        ResponseText = StringFormat("Error: Disband clan failed - {0}.", ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanDisbandReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_MakeChieftainReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String
    Dim Username As String

    Call FindServerRequest(oRequest, Cookie)
    Username = CStr(oRequest.Tag)

    If Result = clresSuccess Then
        ResponseText = StringFormat("{0} is now the chieftain.", Username)
    Else
        ResponseText = StringFormat("Error: Promotion of {0} to chieftain failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanMakeChieftainReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_InviteUserReply(ByVal Cookie As Long, ByVal Result As enuClanResponseValue)
    Dim oRequest As udtServerRequest
    Dim ResponseText As String
    Dim Username As String

    Call FindServerRequest(oRequest, Cookie)
    Username = CStr(oRequest.Tag)

    If Result = clresSuccess Then
        ResponseText = StringFormat("{0} accepted the invitation.", Username)
    Else
        ResponseText = StringFormat("Error: {0} invitation failed - {1}.", Username, ClanHandler.GetClanResponseText(Result))
    End If

    If oRequest.HandlerType = reqScriptingCall Then
        On Error Resume Next
        RunInAll "Event_ClanInviteUserReply", Result
    ElseIf oRequest.HandlerType = reqUserCommand Then
        oRequest.Command.Respond ResponseText
        oRequest.Command.SendResponse
    ElseIf oRequest.HandlerType = reqUserInterface Then
        AddChat RTBColors.ConsoleText, ResponseText
    End If
End Sub

Private Sub ClanHandler_UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)
    If MDebug("debug") Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[CLAN] Unknown clan event [0x" & ZeroOffset(PacketID, 2) & "]. Data is as follows:"
        frmChat.AddChat RTBColors.ErrorMessageText, Data
    End If
End Sub

Public Function GetLogFilePath() As String

    Dim Path As String
    Dim f    As Integer
    
    f = FreeFile
    
    Path = StringFormat("{0}{1}.txt", GetFolderPath("Logs"), Format(Date, "YYYY-MM-DD"))

    If (Dir$(Path) = vbNullString) Then
        Open Path For Output As #f
        Close #1
    End If
    
    GetLogFilePath = Path

End Function

Sub Form_Unload(Cancel As Integer)
    Inet.Cancel
    
    AddChat RTBColors.ErrorMessageText, "Shutting down..."
    
    If Config.FileExists Then
        If Me.WindowState <> vbMinimized Then
            Call RecordWindowPosition(CBool(Me.WindowState = vbMaximized))
        End If
        
        Call Config.Save
    End If

    Call DoDisconnect(1)

    Shell_NotifyIcon NIM_DELETE, nid
    
    On Error Resume Next
    
    RunInAll "Event_LoggedOff"
    RunInAll "Event_Close"
    RunInAll "Event_Shutdown"
    
    DestroyObjs
    
    On Error GoTo 0
    
    If ExReconnectTimerID > 0 Then
        KillTimer 0, ExReconnectTimerID
    End If

    Call modWarden.WardenCleanup(WardenInstance)

    DisableURLDetect frmChat.rtbChat.hWnd
    UnhookWindowProc frmChat.hWnd
    
    DestroyAllWWs
    
    Set g_Logger = Nothing
    Set BotVars = Nothing
    Set ClanHandler = Nothing
    Set ListToolTip = Nothing
    Set GErrorHandler = Nothing
    Set FriendListHandler = Nothing
    Set colWhisperWindows = Nothing
    Set colLastSeen = Nothing
    Set SharedScriptSupport = Nothing
    Set ds = Nothing
    Set ReceiveBuffer(stBNCS) = Nothing
    Set ReceiveBuffer(stBNLS) = Nothing
    Set ReceiveBuffer(stMCP) = Nothing
    
    'Set dictTimerInterval = Nothing
    'Set dictTimerCount = Nothing
    'Set dictTimerEnabled = Nothing
    
    ' Updated to match current form list 2009-02-09 Andy
    Unload frmAbout
    Unload frmCatch
    Unload frmCommandManager
    Unload frmClanInvite
    Unload frmCustomInputBox
    Unload frmDBType
    Unload frmEMailReg
    Unload frmFilters
    Unload frmDBGameSelection
    Unload frmDBNameEntry
    Unload frmDBManager
    Unload frmAccountManager
    Unload frmKeyManager
    Unload frmProfile
    Unload frmQuickChannel
    Unload frmRealm
    Unload frmScript
    Unload frmSettings
    Unload frmSplash
    Unload frmWhisperWindow
    
    ' Added this instead of End to try and fix some system tray crashes 2009-0211-andy
    '  It was used in some capacity before since the API was already declared
    '   in modAPI...
    ' added preprocessor check; the bot was ending the VB6 IDE's process too! - ribose
    ' if it was compiled with the debugger, we don't allow minimizing to tray anyway
    #If (COMPILE_DEBUG <> 1) Then
        Call ExitProcess(0)
    #Else
        End
    #End If
End Sub

Private Sub FriendListHandler_FriendsListReply(ByVal Friends As Collection)
    Dim FriendObj   As clsFriendObj
    Dim i           As Integer
    Dim EntryNumber As Integer

    Set g_Friends = Friends

    If Config.FriendsListTab Then
        lvFriendList.ListItems.Clear
        For i = 1 To Friends.Count
            EntryNumber = i - 1
            Set FriendObj = Friends.Item(i)
            If FriendObj.IsOnline Or Config.ShowOfflineFriends Then
                AddFriendItem FriendObj.DisplayName, FriendObj.Game, FriendObj.Status, FriendObj.LocationID, EntryNumber
            End If
            Set FriendObj = Nothing
        Next i

        ' re-sort
        lvFriendList.Sorted = True

        Call UpdateListviewTabs
    End If
End Sub

Private Sub FriendListHandler_FriendsUpdate(ByVal EntryNumber As Byte, ByVal FriendObj As clsFriendObj)
    Dim ListItem As ListItem
    Dim oRequest As udtServerRequest

    If FriendObj.LocationID <> FRL_OFFLINE Then
        If Not FindServerRequest(oRequest, -1, SID_FRIENDSUPDATE, EntryNumber) Then
            ' NOTE: There is a server bug here where, when this packet is sent automaticlaly
            '   (not requested), the fields contains your own information instead when logged on.
            '   Because of this, we resend the request (if there isn't one already).
            '   (see: https://bnetdocs.org/packet/384/sid-friendsupdate)
            Call FriendListHandler.RequestFriendItem(EntryNumber, reqInternal)
            Exit Sub
        End If
    End If

    If g_Friends.Count > EntryNumber Then
        With g_Friends.Item(EntryNumber + 1)
            .Status = FriendObj.Status
            .LocationID = FriendObj.LocationID
            .Game = FriendObj.Game
            .Location = FriendObj.Location
        End With
    End If

    If Config.FriendsListTab Then
        Set ListItem = GetFriendItem(EntryNumber)
        If Not (ListItem Is Nothing) Then
            ' set the icon and status in place
            SetFriendItem ListItem, EntryNumber, True, FriendObj.Game, FriendObj.Status, FriendObj.LocationID
            Set ListItem = Nothing
        Else
            ' wasn't found...
            If FriendObj.IsOnline Or Config.ShowOfflineFriends Then
                AddFriendItem FriendObj.DisplayName, FriendObj.Game, FriendObj.Status, FriendObj.LocationID, EntryNumber
            End If
        End If

        ' re-sort
        lvFriendList.Sorted = True

        Call UpdateListviewTabs
    End If
End Sub

Private Sub FriendListHandler_FriendsAdd(ByVal FriendObj As clsFriendObj)
    Dim EntryNumber As Integer

    EntryNumber = g_Friends.Count
    g_Friends.Add FriendObj

    If Config.FriendsListTab Then
        If FriendObj.IsOnline Or Config.ShowOfflineFriends Then
            AddFriendItem FriendObj.DisplayName, FriendObj.Game, FriendObj.Status, FriendObj.LocationID, EntryNumber
        End If

        ' re-sort
        lvFriendList.Sorted = True

        Call UpdateListviewTabs
    End If
End Sub

Private Sub FriendListHandler_FriendsRemove(ByVal EntryNumber As Byte)
    Dim ListItem As ListItem
    Dim i        As Integer

    If g_Friends.Count > EntryNumber Then
        g_Friends.Remove EntryNumber + 1
    End If

    If Config.FriendsListTab Then
        Set ListItem = GetFriendItem(EntryNumber)
        If Not (ListItem Is Nothing) Then
            lvFriendList.ListItems.Remove ListItem.Index
            Set ListItem = Nothing
            For i = EntryNumber + 1 To g_Friends.Count + 1
                Set ListItem = GetFriendItem(i)
                If Not (ListItem Is Nothing) Then
                    SetFriendItem ListItem, i - 1
                    Set ListItem = Nothing
                End If
            Next i
        End If

        lvFriendList.Refresh

        Call UpdateListviewTabs
    End If
End Sub

Private Sub FriendListHandler_FriendsPosition(ByVal EntryNumber As Byte, ByVal NewPosition As Byte)
    Dim FriendObj     As clsFriendObj
    Dim ListItem      As ListItem
    Dim ListItemShift As ListItem
    Dim i             As Integer

    If g_Friends.Count > EntryNumber Then
        Set FriendObj = g_Friends.Item(EntryNumber + 1)
        g_Friends.Remove EntryNumber + 1
        If g_Friends.Count > NewPosition Then
            g_Friends.Add FriendObj, , NewPosition + 1
        Else
            g_Friends.Add FriendObj
        End If
        Set FriendObj = Nothing
    End If

    If Config.FriendsListTab Then
        Set ListItem = GetFriendItem(EntryNumber)
        If Not (ListItem Is Nothing) Then
            If EntryNumber < NewPosition Then
                ' f demote
                For i = EntryNumber + 1 To NewPosition
                    Set ListItemShift = GetFriendItem(i)
                    If Not (ListItemShift Is Nothing) Then
                        SetFriendItem ListItemShift, i - 1
                        Set ListItemShift = Nothing
                    End If
                Next i
                SetFriendItem ListItem, NewPosition
            ElseIf EntryNumber > NewPosition Then
                ' f promote
                For i = EntryNumber - 1 To NewPosition Step -1
                    Set ListItemShift = GetFriendItem(i)
                    If Not (ListItemShift Is Nothing) Then
                        SetFriendItem ListItemShift, i + 1
                        Set ListItemShift = Nothing
                    End If
                Next i
                SetFriendItem ListItem, NewPosition
            End If
            Set ListItem = Nothing
        End If

        ' re-sort
        lvFriendList.Sorted = True

        Call UpdateListviewTabs
    End If
End Sub

Private Sub lblCurrentChannel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    PopupMenu mnuQCTop
End Sub

Private Sub ListviewTabs_Click(PreviousTab As Integer)
    Dim CurrentTab As Integer
    Dim PrevListView As ListView
    Dim CurrListView As ListView

    CurrentTab = ListviewTabs.Tab
    If CurrentTab <> PreviousTab Then
        Select Case CurrentTab
            Case LVW_BUTTON_CHANNEL: Set CurrListView = lvChannel
            Case LVW_BUTTON_FRIENDS: Set CurrListView = lvFriendList
            Case LVW_BUTTON_CLAN:    Set CurrListView = lvClanList
        End Select
        CurrListView.Visible = True
        CurrListView.HideSelection = True
        CurrListView.Refresh
    
        Select Case PreviousTab
            Case LVW_BUTTON_CHANNEL: Set PrevListView = lvChannel
            Case LVW_BUTTON_FRIENDS: Set PrevListView = lvFriendList
            Case LVW_BUTTON_CLAN:    Set PrevListView = lvClanList
        End Select
        PrevListView.Visible = False
    End If

    With lblCurrentChannel
        If Not g_Online Then
            .Caption = vbNullString
            .ToolTipText = "Currently offline."
        Else
            Select Case ListviewTabs.Tab
                Case LVW_BUTTON_CHANNEL:
                    If LenB(g_Channel.Name) = 0 Then
                        .Caption = BotVars.Gateway
                        .ToolTipText = StringFormat("Currently online on {0}.", BotVars.Gateway)
                    Else
                        .Caption = StringFormat("{0} ({1})", _
                                g_Channel.Name, g_Channel.Users.Count)
                        .ToolTipText = StringFormat("Currently in {2} channel {0} ({1}).", _
                                g_Channel.Name, g_Channel.Users.Count, g_Channel.sType())
                    End If
                Case LVW_BUTTON_FRIENDS
                    .Caption = StringFormat("Your Friends ({0})", lvFriendList.ListItems.Count)
                    .ToolTipText = StringFormat("Currently viewing {0} friends.", lvFriendList.ListItems.Count)
                Case LVW_BUTTON_CLAN
                    .Caption = StringFormat("Clan {0} ({1} members)", g_Clan.Name, lvClanList.ListItems.Count)
                    .ToolTipText = StringFormat("Currently viewing {1} members of Clan {0}.", g_Clan.Name, lvClanList.ListItems.Count)
            End Select
        End If
    End With
End Sub

Public Sub UpdateListviewTabs()
    Call ListviewTabs_Click(ListviewTabs.Tab)
End Sub

' This procedure relies on code in RecordcboSendSelInfo() that sets global variables
'  cboSendSelLength and cboSendSelStart
' These two properties are zeroed out as the control loses focus and inaccessible
'  (zeroed) at both access time in this method AND in the _LostFocus sub
Private Sub lvChannel_dblClick()
    Dim Value As String

    Value = GetSelectedUser
    If LenB(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
    Value = Value & Space$(1)

    If (Len(Value) > 0) Then
        With cboSend
            .SelStart = cboSendSelStart
            .SelLength = cboSendSelLength
            .SelText = Value
            
            cboSendSelStart = cboSendSelStart + Len(Value)
            cboSendSelLength = 0
            
            .SetFocus
        End With
    End If
End Sub

Private Sub lvChannel_KeyUp(KeyCode As Integer, Shift As Integer)
    Const S_ALT = 4
    
    If (KeyCode = 93) Then
        lvChannel_MouseUp 2, Shift, 0, 0
    ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then
        Dim sStart As Integer
        
        With lvChannel
            If Not (.SelectedItem Is Nothing) Then
                cboSend.SelStart = Len(cboSend.Text)
                cboSend.SelText = .SelectedItem.Text
    
                KeyCode = 0
                Shift = 0
                
                Exit Sub
            End If
        End With
    End If
End Sub

Private Sub lvFriendList_KeyUp(KeyCode As Integer, Shift As Integer)
    Const S_ALT = 4
    
    If (KeyCode = 93) Then
        lvFriendList_MouseUp 2, Shift, 0, 0
    ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then
        Dim sStart As Integer
        
        With lvFriendList
            If Not (.SelectedItem Is Nothing) Then
                cboSend.SelStart = Len(cboSend.Text)
                cboSend.SelText = .SelectedItem.Text
    
                KeyCode = 0
                Shift = 0
                
                Exit Sub
            End If
        End With
    End If
End Sub

Private Sub lvClanList_KeyUp(KeyCode As Integer, Shift As Integer)
    Const S_ALT = 4
    
    If (KeyCode = 93) Then
        lvClanList_MouseUp 2, Shift, 0, 0
    ElseIf (KeyCode = KEY_ALTN And Shift = S_ALT) Then
        Dim sStart As Integer
        
        With lvClanList
            If Not (.SelectedItem Is Nothing) Then
                cboSend.SelStart = Len(cboSend.Text)
                cboSend.SelText = .SelectedItem.Text
    
                KeyCode = 0
                Shift = 0
                
                Exit Sub
            End If
        End With
    End If
End Sub

Private Sub lvFriendList_dblClick()
    Dim Value As String

    Value = GetFriendsSelectedUser
    If LenB(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
    Value = Value & Space$(1)

    If (Len(Value) > 0) Then
        With cboSend
            .SelStart = cboSendSelStart
            .SelLength = cboSendSelLength
            .SelText = Value
            
            cboSendSelStart = cboSendSelStart + Len(Value)
            cboSendSelLength = 0
            
            .SetFocus
        End With
    End If
End Sub

Private Sub lvClanList_dblClick()
    Dim Value As String

    Value = GetClanSelectedUser
    If LenB(cboSend.Text) = 0 Then Value = Value & BotVars.AutoCompletePostfix
    Value = Value & Space$(1)

    If (Len(Value) > 0) Then
        With cboSend
            .SelStart = cboSendSelStart
            .SelLength = cboSendSelLength
            .SelText = Value
            
            cboSendSelStart = cboSendSelStart + Len(Value)
            cboSendSelLength = 0
            
            .SetFocus
        End With
    End If
End Sub

Private Sub lvChannel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim aInx As Integer
    Dim sProd As String * 4
    Dim HasOps As Boolean
    Dim UserIsW3 As Boolean
    Dim UserHasStats As Boolean
    
    If (lvChannel.SelectedItem Is Nothing) Then
        Exit Sub
    End If

    If Button = vbRightButton Then
        If Not (lvChannel.SelectedItem Is Nothing) Then
            aInx = g_Channel.GetUserIndex(GetSelectedUser)
            
            If aInx > 0 Then
                sProd = g_Channel.Users(aInx).Game
                UserIsW3 = (sProd = PRODUCT_W3XP Or sProd = PRODUCT_WAR3)
                Select Case sProd
                    Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
                        UserHasStats = True
                End Select
                HasOps = g_Channel.Self.IsOperator()
                
                mnuPopInvite.Enabled = UserIsW3 And LenB(g_Channel.Users(aInx).Clan) = 0 And InStr(1, GetSelectedUser, "#") = 0 And g_Clan.Self.Rank >= 3
                
                mnuPopStats.Enabled = UserHasStats
                mnuPopWebProfile.Enabled = UserIsW3
                mnuPopKick.Enabled = HasOps
                mnuPopDes.Enabled = HasOps
                mnuPopBan.Enabled = HasOps
            End If
        End If
        
        mnuPopup.Tag = lvChannel.SelectedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
        
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lvFriendList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim aInx As Integer
    Dim sProd As String * 4
    Dim bIsOn As Boolean
    Dim bIsMutual As Boolean
    Dim UserIsW3 As Boolean
    Dim UserHasStats As Boolean
    Dim SelfHasStats As Boolean
    
    If (lvFriendList.SelectedItem Is Nothing) Then
        Exit Sub
    End If

    If Button = vbRightButton Then
        If Not (lvFriendList.SelectedItem Is Nothing) Then
            aInx = lvFriendList.SelectedItem.Index
            
            If aInx > 0 Then
                sProd = g_Friends(aInx).Game
                bIsOn = g_Friends(aInx).IsOnline
                bIsMutual = g_Friends(aInx).IsMutual
                UserIsW3 = (sProd = PRODUCT_W3XP Or sProd = PRODUCT_WAR3)
                Select Case sProd
                    Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
                        UserHasStats = True
                End Select
                Select Case StrReverse(BotVars.Product)
                    Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
                        SelfHasStats = True
                End Select
                
                mnuPopFLWhisper.Enabled = bIsOn
                mnuPopFLInvite.Enabled = UserIsW3 And bIsMutual And g_Clan.Self.Rank >= 3

                mnuPopFLStats.Enabled = UserHasStats Or SelfHasStats
                mnuPopFLWebProfile.Enabled = UserIsW3
            End If
        End If
        
        mnuPopFList.Tag = lvFriendList.SelectedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
        
        PopupMenu mnuPopFList
    End If
End Sub

Private Sub lvChannel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim sOutBuf As String
    Dim sTemp As String
    Dim UserAccess As udtUserAccess
    Dim Clan As String
   
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessageAny(lvChannel.hWnd, LVM_HITTEST, -1, lvhti) + 1
 
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            ListToolTip.Destroy
        Else
            'UserAccess = GetCumulativeAccess(lvChannel.ListItems(m_lCurItemIndex).text, "USER")
        
            ListToolTip.Title = _
                "Information for " & lvChannel.ListItems(m_lCurItemIndex).Text
                
            'If (UserAccess.Name <> vbNullString) Then
            '    sTemp = sTemp & "["
            '
            '    If (UserAccess.Access > 0) Then
            '        sTemp = sTemp & "rank: " & UserAccess.Access
            '    End If
            '
            '    If ((UserAccess.Flags <> "%") And (UserAccess.Flags <> vbNullString)) Then
            '        If (UserAccess.Access > 0) Then
            '            sTemp = sTemp & ", "
            '        End If
            '
            '        sTemp = sTemp & "flags: " & UserAccess.Flags
            '    End If
            '
            '    sTemp = sTemp & "]" & vbCrLf
            'End If
                
            
            lItemIndex = g_Channel.GetUserIndex(lvChannel.ListItems(m_lCurItemIndex).Text)
            
            If (lItemIndex > 0) Then
                With g_Channel.Users(lItemIndex)
                    'ParseStatstring .Statstring, sOutBuf, Clan
            
                    'sTemp = sTemp & vbCrLf
                    sTemp = sTemp & "Ping at logon: " & .Ping & "ms" & vbCrLf
                    sTemp = sTemp & "Flags: " & GetFlagDescription(.Flags, True) & vbCrLf
                    sTemp = sTemp & vbCrLf
                    sTemp = sTemp & .Stats.ToString
                
                    ListToolTip.TipText = sTemp
                    
                End With
                
                Call ListToolTip.Create(lvChannel.hWnd, CLng(x), CLng(y))
            End If
        End If
    End If
End Sub

Private Sub lvFriendList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
   
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessageAny(lvFriendList.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            ListToolTip.Destroy
        Else
            ListToolTip.Title = "Information for " & lvFriendList.ListItems(m_lCurItemIndex).Text
            
            Dim sTemp As String
            
            If ((lItemIndex > 0) And (g_Friends.Count > 0)) Then
                lItemIndex = FriendListHandler.UsernameToFLIndex(lvFriendList.ListItems(m_lCurItemIndex).Text)
            
                With g_Friends.Item(lItemIndex)
'                    Private Const FRL_OFFLINE& = &H0
'                    Private Const FRL_NOTINCHAT& = &H1
'                    Private Const FRL_INCHAT& = &H2
'                    Private Const FRL_PUBLICGAME& = &H3
'                    Private Const FRL_PRIVATEGAME& = &H5
                    If .IsOnline Then
                        sTemp = sTemp & "Using " & ProductCodeToFullName(.Game) & " "
                    End If
                    
                    Select Case .LocationID
                        Case FRL_OFFLINE
                            sTemp = sTemp & "This person is offline."
                        Case FRL_NOTINCHAT
                            sTemp = sTemp & "outside of chat."
                        Case FRL_INCHAT
                            sTemp = sTemp & "in a chat channel."
                        Case FRL_PUBLICGAME
                            sTemp = sTemp & "in public game " & .Location & "."
                        Case FRL_PRIVATEGAME
                            sTemp = sTemp & "in a private game."
                    End Select
                    
'                    Private Const FRS_NONE& = &H0
'                    Private Const FRS_MUTUAL& = &H1
'                    Private Const FRS_DND& = &H2
'                    Private Const FRS_AWAY& = &H4

                    If (.Status And FRS_MUTUAL) = FRS_MUTUAL Then
                        sTemp = sTemp & vbCrLf & "Mutual friend"
                        
                        Select Case (.LocationID)
                            Case FRL_INCHAT
                                sTemp = sTemp & ", in channel " & .Location & "."
                            Case FRL_PUBLICGAME
                                sTemp = sTemp & ", in public game " & .Location & "."
                            Case FRL_PRIVATEGAME_MUTUAL
                                sTemp = sTemp & ", in private game " & .Location & "."
                            Case Else
                                sTemp = sTemp & "."
                        End Select
                    End If
                    
                    If (.Status And FRS_AWAY) = FRS_AWAY Then
                        sTemp = sTemp & vbCrLf & "Currently marked as away."
                    End If
                    
                    If (.Status And FRS_DND) = FRS_DND Then
                        sTemp = sTemp & vbCrLf & "Currently marked as Do Not Disturb."
                    End If
                    
                    ListToolTip.TipText = sTemp
                End With
                
                Call ListToolTip.Create(lvFriendList.hWnd, CLng(x), CLng(y))
            End If
        End If
    End If
End Sub

Private Sub lvClanList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Integer
    Dim ClanMember As clsClanMemberObj
   
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.y = y / Screen.TwipsPerPixelY
    lItemIndex = SendMessageAny(lvClanList.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            ListToolTip.Destroy
        Else
            ListToolTip.Title = "Information for " & lvClanList.ListItems(m_lCurItemIndex).Text
            
            Dim sTemp As String
            
            If ((lItemIndex > 0) And (g_Clan.Members.Count > 0)) Then
                Set ClanMember = g_Clan.GetMember(lvClanList.ListItems(m_lCurItemIndex).Text)
                
                If (Not ClanMember Is Nothing) Then
                    With ClanMember
                        If (.Rank = clrankChieftain) Then
                            sTemp = sTemp & "The "
                        Else
                            sTemp = sTemp & "This "
                        End If
                        
                        sTemp = sTemp & .RankName & " is currently "
                        
                        If (.IsOnline) Then
                            sTemp = sTemp & "online"
                            
                            If (LenB(.Location) > 0) Then
                                sTemp = sTemp & " in " & .Location
                            End If
                        Else
                            sTemp = sTemp & "offline"
                        End If
                        
                        sTemp = sTemp & "."
                        
                        ListToolTip.TipText = sTemp
                    End With
                    Set ClanMember = Nothing
                End If
                
                Call ListToolTip.Create(lvClanList.hWnd, CLng(x), CLng(y))
            End If
        End If
    End If
End Sub

Private Sub mnuAccountManager_Click()
    frmAccountManager.ShowMode ACCOUNT_MODE_LOGON
End Sub

Private Sub mnuCatchPhrases_Click()
    frmCatch.Show
End Sub

Private Sub mnuChangeLog_Click()
    ShellOpenURL "http://www.stealthbot.net/wiki/changelog", "the StealthBot Changelog"
End Sub

Private Sub mnuOpenScriptFolder_Click()
    Dim sPath As String
    sPath = GetFolderPath("Scripts")
    
    ' Does the script folder exist?
    If (LenB(Dir$(sPath, vbDirectory)) > 0) Then
        Shell StringFormat("explorer.exe {0}", sPath), vbNormalFocus
    Else
        ' Try and create it
        MkDir sPath
        
        ' Did it work?
        If (LenB(Dir$(sPath, vbDirectory)) > 0) Then
            Shell StringFormat("explorer.exe {0}", sPath), vbNormalFocus
        Else
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "Your script folder does not exist, and could not be created.")
            Call frmChat.AddChat(RTBColors.ErrorMessageText, "Script folder path: " & sPath)
            Exit Sub
        End If
    End If
End Sub

Private Sub mnuPopClanAddLeft_Click()
    On Error Resume Next
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If txtPre.Enabled Then 'fix for topic 25290 -a
        txtPre.Text = StringFormat("/w {0} ", GetClanSelectedUser)
        
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.Text)
    End If
End Sub

Private Sub mnuPopClanAddToFList_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If Not (lvClanList.SelectedItem Is Nothing) Then
        AddQ "/f a " & CleanUsername(GetClanSelectedUser), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopClanCopy_Click()
    On Error Resume Next
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Clipboard.Clear
    
    Clipboard.SetText GetClanSelectedUser
End Sub

Private Sub mnuPopClanDemote_Click()
    Dim Rank As enuClanRank

    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If MsgBox("Are you sure you want to demote " & GetClanSelectedUser & "?", vbYesNo, "StealthBot") = vbYes Then
        Rank = lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon - 1
        Call ClanHandler.DemoteMember(GetClanSelectedUser, Rank, reqUserInterface)
    End If
End Sub

Private Sub mnuPopClanDisband_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If MsgBox("Are you sure you want to disband this clan?", vbYesNo Or vbCritical, "StealthBot") = vbYes Then
        Call ClanHandler.DisbandClan(reqUserInterface)
    End If
End Sub

Private Sub mnuPopClanLeave_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If MsgBox("Are you sure you want to leave the clan?", vbYesNo, "StealthBot") = vbYes Then
        Call ClanHandler.RemoveMember(GetCurrentUsername, True, reqUserInterface)
    End If
End Sub

Private Sub mnuPopClanMakeChief_Click()
    Dim pBuf As clsDataBuffer

    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If MsgBox("Are you sure you want to make " & GetClanSelectedUser & " the new Chieftain?", vbYesNo Or vbCritical, "StealthBot") = vbYes Then
        Call ClanHandler.MakeMemberChieftain(GetClanSelectedUser, reqUserInterface)
    End If
End Sub

Private Sub mnuPopClanProfile_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Call RequestProfile(GetClanSelectedUser, reqUserInterface)
    
    frmProfile.PrepareForProfile GetClanSelectedUser, False
End Sub

Private Sub mnuPopClanPromote_Click()
    Dim Rank As enuClanRank

    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If MsgBox("Are you sure you want to promote " & GetClanSelectedUser & "?", vbYesNo, "StealthBot") = vbYes Then
        Rank = lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon + 1
        Call ClanHandler.PromoteMember(GetClanSelectedUser, Rank, reqUserInterface)
    End If
End Sub

Private Sub mnuPopClanRemove_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    Dim LastRemoval As Long
    LastRemoval = ClanHandler.TimeSinceLastRemoval

    If LastRemoval < 30 Then
        AddChat RTBColors.ErrorMessageText, "You must wait " & 30 - LastRemoval & " more seconds before you " & _
                "can remove another user from your clan."
    Else
        If MsgBox("Are you sure you want to remove this user from the clan?", vbExclamation + vbYesNo, _
                "StealthBot") = vbYes Then
            If lvClanList.SelectedItem.Index > 0 Then
                Call ClanHandler.RemoveMember(GetClanSelectedUser, False, reqUserInterface)
            End If

            ClanHandler.LastRemoval = GetTickCountMS()
        End If
    End If
End Sub

Private Sub mnuPopClanStatsWAR3_Click()
    Dim sProd As String
    
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    sProd = PRODUCT_WAR3
    
    If (StrComp(sProd, StrReverse$(BotVars.Product), vbBinaryCompare) = 0) Then
        sProd = vbNullString
    Else
        sProd = Space$(1) & sProd
    End If
    
    AddQ "/stats " & CleanUsername(GetClanSelectedUser) & sProd, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopClanStatsW3XP_Click()
    Dim sProd As String
    
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    sProd = PRODUCT_W3XP
    
    If (StrComp(sProd, StrReverse$(BotVars.Product), vbBinaryCompare) = 0) Then
        sProd = vbNullString
    Else
        sProd = Space$(1) & sProd
    End If
    
    AddQ "/stats " & CleanUsername(GetClanSelectedUser) & sProd, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopClanUserlistWhois_Click()
    On Error Resume Next
    
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Dim temp As udtUserAccess
    Dim s As String
    
    s = GetClanSelectedUser
    
    temp = Database.GetUserAccess(s)
    
    With RTBColors
        If temp.Rank > -1 Then
            If temp.Rank > 0 Then
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & "."
                End If
            Else
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "User not found."
                End If
            End If
        Else
            AddChat .ConsoleText, "User not found."
        End If
    End With
End Sub

Private Sub mnuPopClanWebProfileWAR3_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    GetW3LadderProfile GetClanSelectedUser, WAR3
End Sub

Private Sub mnuPopClanWebProfileW3XP_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    GetW3LadderProfile GetClanSelectedUser, W3XP
End Sub

Private Sub mnuPopClanWhisper_Click()
    On Error Resume Next
    
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If cboSend.Text <> vbNullString Then
        AddQ "/w " & CleanUsername(GetClanSelectedUser, True) & Space(1) & _
                cboSend.Text, PRIORITY.CONSOLE_MESSAGE
        
        cboSend.AddItem cboSend.Text, 0
        cboSend.Text = vbNullString
        cboSend.SetFocus
    End If
End Sub

Private Sub mnuPopClanWhois_Click()
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If Not (lvClanList.SelectedItem Is Nothing) Then
        AddQ "/whois " & lvClanList.SelectedItem.Text, PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopFLAddLeft_Click()
    On Error Resume Next
    If Not PopupMenuCLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim Index As Integer
    Dim s As String
    
    If txtPre.Enabled Then 'fix for topic 25290 -a
        s = vbNullString
        If Dii Then s = "*"
        s = StringFormat("/w {0}{1} ", s, GetFriendsSelectedUser)
        txtPre.Text = s
        
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.Text)
    End If
End Sub

Private Sub mnuPopFLCopy_Click()
    On Error Resume Next
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Clipboard.Clear
    
    Clipboard.SetText GetFriendsSelectedUser
End Sub

'Will move the selected user one spot down on the friends list.
Private Sub mnuPopFLDemote_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        With lvFriendList.SelectedItem
            If (.Index < lvFriendList.ListItems.Count) Then
                AddQ "/f d " & GetFriendsSelectedUser, PRIORITY.CONSOLE_MESSAGE
                'MoveFriend .index, .index + 1
            End If
        End With
    End If
End Sub

Private Sub mnuPopFLInvite_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim sPlayer As String
    
    If Not lvFriendList.SelectedItem Is Nothing Then
        sPlayer = GetFriendsSelectedUser
    End If
    
    If LenB(sPlayer) > 0 Then
        If g_Clan.Self.Rank >= 3 Then
            Call ClanHandler.InviteToClan(ReverseConvertUsernameGateway(sPlayer), reqUserInterface)
            AddChat RTBColors.InformationText, "[CLAN] Invitation sent to " & GetFriendsSelectedUser & ", awaiting reply."
        End If
    End If
End Sub

Private Sub mnuPopFLProfile_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.

    If Not lvFriendList.SelectedItem Is Nothing Then
        RequestProfile CleanUsername(lvFriendList.SelectedItem.Text), reqUserInterface

        frmProfile.PrepareForProfile CleanUsername(lvFriendList.SelectedItem.Text), False
    End If
End Sub

'Will move the selected user one spot up on the friends list.
Private Sub mnuPopFLPromote_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        With lvFriendList.SelectedItem
            If (.Index > 1) Then
                AddQ "/f p " & GetFriendsSelectedUser, PRIORITY.CONSOLE_MESSAGE
                'MoveFriend .index, .index - 1
            End If
        End With
    End If
End Sub

Private Sub mnuPopFLRefresh_Click()
    Call FriendListHandler.RequestFriendsList
End Sub

Private Sub mnuPopFLRemove_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        AddQ "/f r " & GetFriendsSelectedUser, PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopFLStats_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Dim aInx As Integer
    Dim sProd As String
    
    aInx = lvFriendList.SelectedItem.Index
    sProd = g_Friends(aInx).Game
    Select Case sProd
        Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
            ' get stats for user on their current product
        Case Else
            ' their current product does not have stats, or they are offline
            Select Case StrReverse$(BotVars.Product)
                Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
                    ' get stats for user on the bot's product
                    sProd = StrReverse$(BotVars.Product)
                Case Else
                    ' unspecified product
                    AddChat RTBColors.ConsoleText, "This bot and the specified friend are not on a game that stores stats viewable via the Battle.net /stats command. " & _
                                                   "Type /stats " & CleanUsername(GetFriendsSelectedUser) & " <desired product code> to get this user's stats for another game."
                    Exit Sub
            End Select
    End Select
    
    If (StrComp(sProd, StrReverse$(BotVars.Product), vbBinaryCompare) = 0) Then
        sProd = vbNullString
    Else
        sProd = Space$(1) & sProd
    End If
    
    AddQ "/stats " & CleanUsername(GetFriendsSelectedUser) & sProd, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopFLUserlistWhois_Click()
    On Error Resume Next
    
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Dim temp As udtUserAccess
    Dim s As String
    
    s = GetFriendsSelectedUser
    
    temp = Database.GetUserAccess(s)
    
    With RTBColors
        If temp.Rank > -1 Then
            If temp.Rank > 0 Then
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & "."
                End If
            Else
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "User not found."
                End If
            End If
        Else
            AddChat .ConsoleText, "User not found."
        End If
    End With
End Sub

Private Sub mnuPopFLWebProfile_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Dim aInx As Integer
    Dim sProd As String
    Dim webProd As enuWebProfileTypes
    
    aInx = lvFriendList.SelectedItem.Index
    sProd = g_Friends(aInx).Game
    Select Case sProd
        ' get web profile for user on their current product
        Case PRODUCT_WAR3
            webProd = WAR3
        Case PRODUCT_W3XP
            webProd = W3XP
        Case Else
            Select Case StrReverse$(BotVars.Product)
                ' get web profile for user on the bot's product
                Case PRODUCT_WAR3
                    webProd = WAR3
                Case PRODUCT_W3XP
                    webProd = W3XP
                Case Else
                    ' their current product does not have stats, or they are offline
                    AddChat RTBColors.ConsoleText, "The specified friend must be online to decide which web profile to view."
                    Exit Sub
            End Select
    End Select
    
    GetW3LadderProfile CleanUsername(GetFriendsSelectedUser), webProd
End Sub

Private Sub mnuPopFLWhisper_Click()
    Dim Value As String
    
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    Value = cboSend.Text
    
    If LenB(Value) > 0 Then
        Value = "/w " & CleanUsername(GetFriendsSelectedUser, True) & Space$(1) & Value
        
        AddQ Value, PRIORITY.CONSOLE_MESSAGE
        
        cboSend.AddItem Value, 0
        cboSend.Text = vbNullString
        
        On Error Resume Next
        cboSend.SetFocus
    End If
End Sub

Private Sub mnuPopInvite_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim sPlayer As String
    
    If Not lvChannel.SelectedItem Is Nothing Then
        sPlayer = GetSelectedUser
    End If
    
    If LenB(sPlayer) > 0 Then
        If g_Clan.Self.Rank >= 3 Then
            Call ClanHandler.InviteToClan(ReverseConvertUsernameGateway(sPlayer), reqUserInterface)
            AddChat RTBColors.InformationText, "[CLAN] Invitation sent to " & GetSelectedUser & ", awaiting reply."
        End If
    End If
End Sub

Private Sub mnuPopProfile_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim sUser As String
    sUser = StripAccountNumber(CleanUsername(GetSelectedUser))
    
    Call RequestProfile(sUser, reqUserInterface)
    
    frmProfile.PrepareForProfile sUser, False
End Sub

Private Sub mnuPopStats_Click()
    Dim aInx As Integer
    Dim sProd As String
    
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    aInx = g_Channel.GetUserIndex(GetSelectedUser)
    sProd = g_Channel.Users(aInx).Game
    Select Case sProd
        Case PRODUCT_STAR, PRODUCT_SEXP, PRODUCT_W2BN, PRODUCT_WAR3, PRODUCT_W3XP, PRODUCT_JSTR, PRODUCT_SSHR
            ' get stats for user on their current product
        Case Else
            ' unspecified product
            AddChat RTBColors.ConsoleText, "The specified user is not on a game that stores stats viewable via the Battle.net /stats command. " & _
                                           "Type /stats " & CleanUsername(GetSelectedUser) & " <desired product code> to get this user's stats for another game."
            Exit Sub
    End Select
    
    If (StrComp(sProd, StrReverse$(BotVars.Product), vbBinaryCompare) = 0) Then
        sProd = vbNullString
    Else
        sProd = Space$(1) & sProd
    End If
    
    AddQ "/stats " & StripAccountNumber(CleanUsername(GetSelectedUser)) & sProd, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopUserlistWhois_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim temp As udtUserAccess
    Dim s As String
    
    s = GetSelectedUser
    
    temp = Database.GetUserAccess(s)
    
    With RTBColors
        If temp.Rank > -1 Then
            If temp.Rank > 0 Then
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & " and flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "Found user " & s & ", with rank " & temp.Rank & "."
                End If
            Else
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "User not found."
                End If
            End If
        Else
            AddChat .ConsoleText, "User not found."
        End If
    End With
End Sub

Private Sub mnuHomeChannel_Click()
    If (LenB(Config.HomeChannel) = 0) Then
        ' do product home join instead
        Call DoChannelJoinProductHome
    Else
        ' go home
        Call FullJoin(Config.HomeChannel, 2)
    End If
End Sub

Private Sub mnuLastChannel_Click()
    If (LenB(BotVars.LastChannel) = 0) Then
        ' do product home join instead
        Call DoChannelJoinProductHome
    Else
        ' go to last
        Call FullJoin(BotVars.LastChannel, 2)
    End If
End Sub

Private Sub mnuRealmSwitch_Click()
    If Dii Then
        If ds.MCPHandler Is Nothing Then
            Set ds.MCPHandler = New clsMCPHandler
            
            SEND_SID_QUERYREALMS2
        Else
            Call ds.MCPHandler.HandleQueryRealmServersResponse
        End If
    End If
End Sub

Private Sub mnuPublicChannels_Click(Index As Integer)
    ' some public channels are redirects
    'If (StrComp(mnuPublicChannels(Index).Caption, g_Channel.Name, vbTextCompare) = 0) Then
    '    Exit Sub
    'End If
    
    If Not BotVars.PublicChannels Is Nothing Then
        If BotVars.PublicChannels.Count > Index Then
            Select Case Config.AutoCreateChannels
                Case "ALERT", "NEVER"
                    Call FullJoin(BotVars.PublicChannels.Item(Index + 1), 0)
                Case Else ' "ALWAYS"
                    Call FullJoin(BotVars.PublicChannels.Item(Index + 1), 2)
            End Select
        End If
        'AddQ "/join " & PublicChannels.Item(Index + 1), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuCustomChannels_Click(Index As Integer)
    If (StrComp(QC(Index + 1), g_Channel.Name, vbTextCompare) = 0) Then
        Exit Sub
    End If

    Select Case Config.AutoCreateChannels
        Case "ALERT", "NEVER"
            Call FullJoin(QC(Index + 1), 0)
        Case Else ' "ALWAYS"
            Call FullJoin(QC(Index + 1), 2)
    End Select
    
    'AddQ "/join " & QC(Index + 1), PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuCommandManager_Click()
    frmCommandManager.Show
End Sub

Private Sub mnuConnect2_Click()
    Call DoConnect
End Sub

Private Sub mnuDisableVoidView_Click()
    mnuDisableVoidView.Checked = Not (mnuDisableVoidView.Checked)
    Config.VoidView = Not CBool(mnuDisableVoidView.Checked)
    Call Config.Save
    
    If Config.VoidView And g_Channel.IsSilent Then
        tmrSilentChannel(1).Enabled = Config.VoidView
        AddQ "/unignore " & GetCurrentUsername
    End If
End Sub

Private Sub mnuDisconnect2_Click()
'    Dim Key As String, L As Long
'    Key = GetProductKey()
    
'    If AttemptedNewVerbyte Then
'        AttemptedNewVerbyte = False
'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
'    End If
    
    GErrorHandler.Reset
    Call DoDisconnect
End Sub

Private Sub mnuEditCaught_Click()
    If Dir$(GetFilePath(FILE_CAUGHT_PHRASES)) = vbNullString Then
        MsgBox "The bot has not caught any phrases yet."
        Exit Sub
    Else
        ShellOpenURL GetFilePath(FILE_CAUGHT_PHRASES), , False
    End If
End Sub

Private Sub mnuFlash_Click()
    mnuFlash.Checked = (Not mnuFlash.Checked)
    Config.FlashOnEvents = CBool(mnuFlash.Checked)
    Call Config.Save
End Sub

'Moves a person in the friends list view.
Private Sub MoveFriend(startPos As Integer, endPos As Integer)
    With lvFriendList.ListItems
        If (startPos > endPos) Then
            .Add endPos, , .Item(startPos).Text, , .Item(startPos).SmallIcon
            .Item(endPos).ListSubItems.Add , , , .Item(startPos + 1).ListSubItems.Item(1).ReportIcon
            .Remove startPos + 1
        Else
            .Add endPos + 1, , .Item(startPos).Text, .Item(startPos).Icon, .Item(startPos).SmallIcon
            .Item(endPos + 1).ListSubItems.Add , , , .Item(startPos).ListSubItems.Item(1).ReportIcon
            .Remove startPos
        End If
    End With
End Sub

Private Sub mnuGetNews_Click()
    If Not RequestInetPage(GetNewsURL(), SB_INET_NEWS, False) Then
        Call HandleNews("Inet is busy", -1)
    End If
End Sub

Sub mnuHelpReadme_Click()
    OpenReadme
End Sub

Sub mnuHelpWebsite_Click()
    ShellOpenURL "http://www.stealthbot.net", "the StealthBot Forum"
End Sub

Private Sub mnuHideBans_Click()
    mnuHideBans.Checked = (Not mnuHideBans.Checked)
    Config.HideBanMessages = CBool(mnuHideBans.Checked)
    Call Config.Save
    AddChat RTBColors.InformationText, "Ban messages " & IIf(mnuHideBans.Checked, "disabled", "enabled") & "."
End Sub

Private Sub mnuHideWhispersInrtbChat_Click()
    mnuHideWhispersInrtbChat.Checked = (Not mnuHideWhispersInrtbChat.Checked)
    Config.HideWhispersInMain = CBool(mnuHideWhispersInrtbChat.Checked)
    Call Config.Save
End Sub

Private Sub mnuLog0_Click()
    BotVars.Logging = 2
    Config.LoggingMode = BotVars.Logging
    Call Config.Save
    
    AddChat RTBColors.InformationText, "Full text logging enabled."
    mnuLog1.Checked = False
    mnuLog0.Checked = True
    mnuLog2.Checked = False
    'mnuLog3.Checked = False
    
    'MakeLoggingDirectory
End Sub

Private Sub mnuLog1_Click()
    BotVars.Logging = 1
    Config.LoggingMode = BotVars.Logging
    Call Config.Save
    
    AddChat RTBColors.InformationText, "Partial text logging enabled."
    mnuLog1.Checked = True
    mnuLog0.Checked = False
    mnuLog2.Checked = False
    'mnuLog3.Checked = False
    
    'MakeLoggingDirectory
End Sub

Private Sub mnuLog2_Click()
    BotVars.Logging = 0
    Config.LoggingMode = BotVars.Logging
    Call Config.Save
    
    AddChat RTBColors.InformationText, "Logging disabled."
    mnuLog1.Checked = False
    mnuLog0.Checked = False
    mnuLog2.Checked = True
    'mnuLog3.Checked = False
End Sub

Private Sub mnuOpenBotFolder_Click()
    Shell StringFormat("explorer.exe {0}", CurDir$()), vbNormalFocus
End Sub


Private Sub mnuPacketLog_Click()
    Dim f As Integer
    
    If mnuPacketLog.Checked Then
        ' turning this feature off
        AddChat RTBColors.SuccessText, "StealthBot packet traffic will no longer be logged."
    Else
        ' turning it on
        AddChat RTBColors.SuccessText, "StealthBot packet traffic will be logged in the bot's folder, in a file named " & Format(Date, "yyyy-MM-dd") & "-PacketLog.txt."
        AddChat RTBColors.SuccessText, "--"
        AddChat RTBColors.SuccessText, "Log packets at your own risk! Please read the note below:"
        AddChat RTBColors.ErrorMessageText, "*** CAUTION: THIS LOG MAY CONTAIN PRIVATE INFORMATION."
        AddChat RTBColors.ErrorMessageText, "*** CAUTION: DO NOT DISTRIBUTE it in public posts on StealthBot.net or on any other website!"
        AddChat RTBColors.ErrorMessageText, "*** CAUTION: Only produce a packet log if you're specifically instructed to by"
        AddChat RTBColors.ErrorMessageText, "*** CAUTION: a StealthBot.net tech, or you know what you're doing!"
        AddChat RTBColors.SuccessText, "If you wish to stop logging packets, uncheck the menu item or restart your bot."
        AddChat RTBColors.SuccessText, "This feature only logs StealthBot traffic. It is not a system-wide packet capture utility."
    End If
    
    mnuPacketLog.Checked = Not mnuPacketLog.Checked
    LogPacketTraffic = mnuPacketLog.Checked
End Sub

Private Sub mnuPopAddLeft_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim Index As Integer
    Dim s As String
    
    If txtPre.Enabled Then 'fix for topic 25290 -a
        s = vbNullString
        If Dii Then s = "*"
        s = StringFormat("/w {0}{1} ", s, GetSelectedUser)
        txtPre.Text = s
        
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.Text)
    End If
End Sub

Private Sub mnuPopAddToFList_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If Not (lvChannel.SelectedItem Is Nothing) Then
        AddQ "/f a " & StripAccountNumber(CleanUsername(GetSelectedUser)), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopDes_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN

    AddQ "/designate " & CleanUsername(GetSelectedUser, True), PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopFLWhois_Click()
    If Not PopupMenuFLUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on.
    
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        AddQ "/whois " & lvFriendList.SelectedItem.Text, PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopSafelist_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN

    Dim toSafe As String
    
    On Error Resume Next
    
    toSafe = GetSelectedUser
    
    Call ProcessCommand(GetCurrentUsername, "/safeadd " & toSafe, True, False)
End Sub

Private Sub mnuPopShitlist_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN

    Dim toBan As String
    
    On Error Resume Next
    
    toBan = GetSelectedUser
    
    Call ProcessCommand(GetCurrentUsername, "/shitadd " & toBan, True, False)
End Sub

Private Sub mnuPopSquelch_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/squelch " & GetSelectedUser, PRIORITY.CONSOLE_MESSAGE, PRIORITY.CONSOLE_MESSAGE
End Sub


Private Sub mnuPopUnsquelch_Click()
    'On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/unsquelch " & GetSelectedUser, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopWhisper_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If cboSend.Text <> vbNullString Then
        AddQ "/w " & CleanUsername(GetSelectedUser, True) & Space(1) & _
                cboSend.Text, PRIORITY.CONSOLE_MESSAGE
        
        cboSend.AddItem cboSend.Text, 0
        cboSend.Text = vbNullString
        cboSend.SetFocus
    End If
End Sub

Private Sub mnuClear_Click()
    Call ClearChatScreen(3)
End Sub

Private Sub mnuClearWW_Click()
    Call ClearChatScreen(2)
End Sub

' clear the chat screen:
' 1 - clear chat window only
' 2 - clear whisper window only
' 3 (default) - clear chat and whisper
Sub ClearChatScreen(Optional ByVal ClearOption As Integer = 3)
    ' if they passed 0 (False), change to 3
    If ClearOption = 0 Then ClearOption = 3
    ' if they passed -1 (True), change to 1 (old behavior: DoNotClearWhispers)
    If ClearOption = -1 Then ClearOption = 1
    ' check for 2 (or 3) and clear whispers
    If ClearOption And 2 Then
        rtbWhispers.Text = vbNullString
        ' add cleared message
        AddWhisper RTBColors.ConsoleText, ">> Whisper window cleared."
    End If
    
    ' check for 1 (or 3) and clear chats
    If ClearOption And 1 Then
        rtbChat.Text = vbNullString
        rtbChatLength = 0
        ' add a sensical cleared message
        If ClearOption And 2 Then
            AddChat RTBColors.InformationText, "Chat window cleared."
        Else
            AddChat RTBColors.InformationText, "Chat and whisper windows cleared."
        End If
    End If
    
    ' set focus to send box
    On Error Resume Next
    cboSend.SetFocus
End Sub

Private Sub mnuPopWhois_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/whois " & CleanUsername(GetSelectedUser, True), PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopWebProfile_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim aInx As Integer
    Dim sProd As String
    Dim webProd As enuWebProfileTypes
    
    aInx = lvChannel.SelectedItem.Index
    sProd = g_Channel.Users(aInx).Game
    Select Case sProd
        ' get web profile for user on their current product
        Case PRODUCT_WAR3
            webProd = WAR3
        Case PRODUCT_W3XP
            webProd = W3XP
        Case Else
            ' their current product does not have a web profile
            AddChat RTBColors.ConsoleText, "The specified user must be on WarCraft III to view their web profile."
            Exit Sub
    End Select
    
    GetW3LadderProfile StripAccountNumber(CleanUsername(GetSelectedUser)), webProd
End Sub

Private Sub mnuClearedTxt_Click()
    Dim sPath As String
    sPath = StringFormat("{0}{1}.txt", g_Logger.LogPath, Format(Date, "yyyy-MM-dd"))
    
    If LenB(Dir$(sPath)) = 0 Then
        AddChat RTBColors.ErrorMessageText, "The log file for today is empty."
    Else
        ShellOpenURL sPath, , False
    End If
End Sub

Private Sub mnuRecordWindowPos_Click()
    RecordWindowPosition
End Sub

Private Sub mnuRepairCleanMail_Click()
    CleanUpMailFile
    frmChat.AddChat RTBColors.SuccessText, "Delivered and invalid pieces of mail have been removed from your mail.dat file."
End Sub

Private Sub mnuRepairDataFiles_Click()
    If MsgBox("Are you sure? This action will delete your mail.dat (Bot mail database) file.", vbYesNo, "Repair data files") = vbYes Then
        On Error Resume Next
        Kill GetFilePath(FILE_MAILDB)
        AddChat RTBColors.SuccessText, "The bot's DAT data files have been removed."
    End If
End Sub

Private Sub mnuRepairVerbytes_Click()
    Dim Index As Integer
    
    For Index = 0 To UBound(ProductList)
        If ProductList(Index).BNLS_ID > 0 Then
            Config.SetVersionByte ProductList(Index).ShortCode, GetVerByte(ProductList(Index).Code, 1)
        End If
    Next
    
    Call Config.Save
    
    frmChat.AddChat RTBColors.SuccessText, "The version bytes stored in config.ini have been restored to their defaults."
End Sub

Private Sub mnuScripts_Click()
'do nothing
End Sub

Private Sub mnuToggleShowOutgoing_Click()
    mnuToggleShowOutgoing.Checked = (Not mnuToggleShowOutgoing.Checked)
    Config.ShowOutgoingWhispers = CBool(mnuToggleShowOutgoing.Checked)
    Call Config.Save
End Sub

Private Sub mnuToggleWWUse_Click()
    mnuToggleWWUse.Checked = (Not mnuToggleWWUse.Checked)
    
    Config.WhisperWindows = CBool(mnuToggleWWUse.Checked)
    Call Config.Save
    
    If Not mnuToggleWWUse.Checked Then
        DestroyAllWWs
    End If
End Sub

Private Sub mnuUpdateVerbytes_Click()
    If Not RequestInetPage(VERBYTE_SOURCE, SB_INET_VBYTE, False) Then
        Call HandleUpdateVerbyte("Inet is busy", -1)
    End If
End Sub

Private Sub mnuWhisperCleared_Click()
    Dim sPath As String
    sPath = StringFormat("{0}{1}-WHISPERS.txt", g_Logger.LogPath, Format(Date, "yyyy-MM-dd"))
    
    If LenB(Dir$(sPath)) = 0 Then
        AddChat RTBColors.ErrorMessageText, "The whisper log file for today is empty."
    Else
        ShellOpenURL sPath, , False
    End If
End Sub

Private Sub mnuEditUsers_Click()
    ShellOpenURL GetFilePath(FILE_USERDB), , False
End Sub

Sub mnuReloadScripts_Click()
    
    On Error GoTo ERROR_HANDLER

    'RunInAll "Event_LoggedOff"
    RunInAll "Event_Close"

    InitScriptControl SControl
    LoadScripts

    InitScripts
    
    Exit Sub

ERROR_HANDLER:

    ' Cannot call this method while the script is executing
    If (Err.Number = -2147467259) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Error: Script is still executing."
        
        Exit Sub
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in mnuReloadScripts_Click()."
    
End Sub

Private Sub mnuSetTop_Click()
    mnuLog0.Checked = False
    mnuLog1.Checked = False
    mnuLog2.Checked = False

    Select Case BotVars.Logging
        Case 2: mnuLog0.Checked = True
        Case 1: mnuLog1.Checked = True
        Case 0: mnuLog2.Checked = True
    End Select
End Sub

Private Sub mnuTerms_Click()
    ShellOpenURL "http://eula.stealthbot.net", "the StealthBot EULA"
End Sub

Private Sub mnuFilters_Click()
    frmFilters.Show
End Sub

Private Sub mnuPopCopy_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Clipboard.Clear
    
    Clipboard.SetText GetSelectedUser
End Sub

Private Sub mnuProfile_Click()
    frmProfile.PrepareForProfile vbNullString, True
End Sub

Private Sub mnuCustomChannelAdd_Click()
    Dim i As Integer
    
    If LenB(g_Channel.Name) > 0 Then
    
        For i = LBound(QC) To UBound(QC)
            If LenB(Trim$(QC(i))) = 0 Then
                QC(i) = g_Channel.Name
                PrepareQuickChannelMenu
                
                Exit Sub
            End If
        Next i
    
    End If
End Sub

Private Sub mnuCustomChannelEdit_Click()
    frmQuickChannel.Show
End Sub

Private Sub mnuReload_Click()
    On Error Resume Next
    Call ReloadConfig(1)
    AddChat RTBColors.SuccessText, "Configuration file loaded."
End Sub

Private Sub mnuUTF8_Click()
    If mnuUTF8.Checked Then
        mnuUTF8.Checked = False
        AddChat RTBColors.ConsoleText, "Messages will no longer be UTF-8-decoded."
    Else
        mnuUTF8.Checked = True
        AddChat RTBColors.ConsoleText, "Messages will now be UTF-8-decoded."
    End If
    
    Config.UseUTF8 = CBool(mnuUTF8.Checked)
    Call Config.Save
End Sub

Private Sub rtbChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And ((KeyCode = vbKeyL) Or (KeyCode = vbKeyE) Or (KeyCode = vbKeyR)) Then
        'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
        Select Case KeyCode
            Case vbKeyL
                Call mnuLock_Click
            Case vbKeyR
                Call mnuReloadScripts_Click
        End Select
        
        'Disable Ctrl+L, Ctrl+E, and Ctrl+R
        KeyCode = 0
    ElseIf (Shift = vbShiftMask) And (KeyCode = vbKeyDelete) Then
        'Call Shift+DEL keyboard shortcut since it doens't work with RTB focus.
        Call mnuClear_Click
    End If
End Sub

Private Sub rtbChat_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 32) Then
        Exit Sub
    End If

    cboSend.SetFocus
    cboSend.SelText = Chr$(KeyAscii)
End Sub

Private Sub rtbWhispers_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And ((KeyCode = vbKeyL) Or (KeyCode = vbKeyE) Or (KeyCode = vbKeyR)) Then
        'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
        Select Case KeyCode
            Case vbKeyL
                Call mnuLock_Click
            Case vbKeyR
                Call mnuReloadScripts_Click
        End Select
        
        'Disable Ctrl+L, Ctrl+E, and Ctrl+R
        KeyCode = 0
    ElseIf (Shift = vbShiftMask) And (KeyCode = vbKeyDelete) Then
        'Call Shift+DEL keyboard shortcut since it doens't work with RTB focus.
        Call mnuClear_Click
    End If
End Sub

Private Sub rtbWhispers_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 32) Then
        Exit Sub
    End If

    cboSend.SetFocus
    cboSend.SelText = Chr$(KeyAscii)
End Sub

Private Sub rtbChat_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 And Len(rtbChat.SelText) > 0 Then
        If Not BotVars.NoRTBAutomaticCopy Then
            Clipboard.Clear
            Clipboard.SetText rtbChat.SelText, vbCFText
        End If
    End If
End Sub

Private Sub rtbWhispers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 1 And Len(rtbWhispers.SelText) > 0 Then
        Clipboard.Clear
        Clipboard.SetText rtbWhispers.SelText, vbCFText
    End If
End Sub

Private Sub mnuToggleFilters_Click()
    mnuToggleFilters.Checked = (Not (mnuToggleFilters.Checked))
    
    If Filters Then
        Filters = False
        AddChat RTBColors.InformationText, "Chat filtering disabled."
    Else
        Filters = True
        AddChat RTBColors.InformationText, "Chat filtering enabled."
    End If
    
    Config.ChatFilters = Filters
    Call Config.Save
End Sub

Private Sub mnuConnect_Click()
    GErrorHandler.Reset
    Call DoConnect
End Sub

Private Sub mnuPopKick_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If MyFlags = 2 Or MyFlags = 18 Then
        AddQ "/kick " & CleanUsername(GetSelectedUser, True), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopBan_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If MyFlags = 2 Or MyFlags = 18 Then
        AddQ "/ban " & CleanUsername(GetSelectedUser, True), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuTrayExit_click()
    Dim Result As VbMsgBoxResult

    Result = MsgBox("Are you sure you want to quit?", vbYesNo, "StealthBot")
    
    If (Result = vbYes) Then
        'frmChat.Show
    
        'UnhookWindowProc
        'RESTORE FORM
        'Call NewWindowProc(frmChat.hWnd, 0&, ID_TASKBARICON, WM_LBUTTONDOWN)
        'Call Form_Unload(0)
        Unload frmChat
    End If
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = vbNormal
    Me.Show
End Sub

Sub mnuLock_Click()
    mnuLock.Checked = (Not (mnuLock.Checked))
    
    If BotVars.LockChat = False Then
        AddChat RTBColors.InformationText, "Chat window locked."
        AddChat RTBColors.ErrorMessageText, "NO MESSAGES WHATSOEVER WILL BE DISPLAYED UNTIL YOU UNLOCK THE WINDOW."
        AddChat RTBColors.ErrorMessageText, "To return to normal mode, press CTRL+L or use the toggle under the Window menu."
        BotVars.LockChat = True
    Else
        BotVars.LockChat = False
        AddChat RTBColors.SuccessText, "Chat window unlocked."
    End If
End Sub

Sub mnuDisconnect_Click()
'    Dim Key As String, L As Long
'    Key = GetProductKey()
    
'    If AttemptedNewVerbyte Then
'        AttemptedNewVerbyte = False
'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
'    End If
    
    GErrorHandler.Reset
    Call DoDisconnect
End Sub

Private Sub mnuExit_Click()
    'Call Form_Unload(0)
    Unload frmChat
End Sub

Private Sub mnuSetup_Click()
    If (SettingsForm Is Nothing) Then
        Set SettingsForm = New frmSettings
    End If
    
    SettingsForm.Show
End Sub

Private Sub mnuAbout_click()
    frmAbout.Show
End Sub

Private Sub mnuToggle_Click()
    mnuToggle.Checked = (Not (mnuToggle.Checked))
    
    If JoinMessagesOff = False Then
        AddChat RTBColors.InformationText, "Join/Leave messages disabled."
        JoinMessagesOff = True
    Else
        AddChat RTBColors.InformationText, "Join/Leave messages enabled."
        JoinMessagesOff = False
    End If
    
    Config.ShowJoinLeaves = Not JoinMessagesOff
    Call Config.Save
End Sub

Private Sub mnuUserDBManager_Click()
    frmDBManager.Show
End Sub

Private Sub mnuScript_Click(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Menu", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_Click"
End Sub

Private Sub sckScript_Connect(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_Connect"
End Sub

Private Sub sckScript_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_ConnectionRequest", RequestID
End Sub

Private Sub sckScript_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_DataArrival", bytesTotal
End Sub

Private Sub sckScript_SendComplete(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_SendComplete"
End Sub

Private Sub sckScript_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_SendProgress", bytesSent, bytesRemaining
End Sub

Private Sub sckScript_Close(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_Close"
End Sub

Private Sub sckScript_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Winsock", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_Error", Number, Description, sCode, Source, HelpFile, HelpContext, CancelDisplay
End Sub

Private Sub itcScript_StateChanged(Index As Integer, ByVal State As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Inet", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_StateChanged", State
End Sub

Private Sub tmrAccountLock_Timer()
    tmrAccountLock.Enabled = False

    If (Not sckBNet.State = sckConnected) Then 'g_online is set to true AFTER we login... makes this moot, changed to socket being connected.
        Exit Sub
    End If

    ds.AccountEntryPending = False
    Call Event_LogonEvent(tmrAccountLock.Tag, -2&, vbNullString)

    AddChat RTBColors.ErrorMessageText, "[BNCS] Your account appears to be locked, likely due to an excessive number of " & _
        "invalid logons.  Please try that account again in 15-20 minutes."

    Call DoDisconnect
End Sub

Private Sub tmrScript_Timer(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("Timer", Index)
    RunInSingle obj.SCModule, obj.ObjName & "_Timer"
End Sub

Private Sub tmrScriptLong_Timer(Index As Integer)
    On Error Resume Next

    Dim obj As scObj
    
    obj = GetScriptObjByIndex("LongTimer", Index)
    obj.obj.Counter = (obj.obj.Counter + 1)
    
    If (obj.obj.Counter >= obj.obj.Interval) Then
        RunInSingle obj.SCModule, obj.ObjName & "_Timer"
        
        obj.obj.Counter = 0
    End If
End Sub

Private Sub txtPre_GotFocus()
    Call cboSend_GotFocus
End Sub

Private Sub txtPost_GotFocus()
    Call cboSend_GotFocus
End Sub

Private Sub cboSend_GotFocus()
    On Error Resume Next

    Dim i As Integer
    cboSend.SelStart = cboSendSelStart
    cboSend.SelLength = cboSendSelLength

    If (BotVars.NoAutocompletion = False) Then
        For i = 0 To (Controls.Count - 1)
            If (TypeOf Controls(i) Is ListView) Or _
                    (TypeOf Controls(i) Is SSTab) Or _
                        (TypeOf Controls(i) Is RichTextBox) Or _
                            (TypeOf Controls(i) Is TextBox) Then
                If (Controls(i).TabStop = False) Then
                    Controls(i).Tag = "False"
                End If

                Controls(i).TabStop = False
            End If
        Next i
    End If
    
    cboSendHadFocus = True
End Sub

Private Sub txtPre_LostFocus()
    Call cboSend_LostFocus
End Sub

Private Sub txtPost_LostFocus()
    Call cboSend_LostFocus
End Sub

Private Sub cboSend_LostFocus()
    On Error Resume Next

    Dim i As Integer
    
    If (BotVars.NoAutocompletion = False) Then
        For i = 0 To (Controls.Count - 1)
            If (TypeOf Controls(i) Is ListView) Or _
                    (TypeOf Controls(i) Is TabStrip) Or _
                        (TypeOf Controls(i) Is RichTextBox) Or _
                            (TypeOf Controls(i) Is TextBox) Then
                            
                If (Controls(i).Tag <> "False") Then
                    Controls(i).TabStop = True
                End If
            End If
        Next i
    End If
    
    cboSendHadFocus = False
End Sub

Private Sub cboSend_Click()
    RecordcboSendSelInfo
End Sub

Private Sub cboSend_Change()
    'Debug.Print cboSendSelLength & "\" & cboSendSelStart
    RecordcboSendSelInfo
End Sub

Private Sub cboSend_KeyUp(KeyCode As Integer, Shift As Integer)
    RecordcboSendSelInfo
End Sub

Private Sub cboSend_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERROR_HANDLER
    
    Static ACBuffer    As String
    Static ACWordStart As Long
    Static ACWordEnd   As Long
    Static ACUserPLen  As Long
    Static ACUserIndex As Integer
    
    Dim AccessAmount As udtUserAccess
    
    Dim Vetoed As Boolean '
    
    Dim i           As Integer
    Dim j           As Integer
    Dim SavedSelPos As Long
    Dim SavedSelLen As Long
    Dim CurrentTab  As Integer
    Dim CurrentList As ListView
    Dim CurrentUser As Long
    Dim Value       As String
    
    SavedSelPos = cboSend.SelStart
    SavedSelLen = cboSend.SelLength
    
    CurrentTab = ListviewTabs.Tab
    Select Case CurrentTab
        Case LVW_BUTTON_FRIENDS: Set CurrentList = lvFriendList
        Case LVW_BUTTON_CLAN:    Set CurrentList = lvClanList
        Case Else:               Set CurrentList = lvChannel
    End Select
    
    If (Not (CurrentList.SelectedItem Is Nothing)) Then
        CurrentUser = CurrentList.SelectedItem.Index
    End If

    Select Case (KeyCode)
        Case vbKeySpace
            With cboSend
                If (LenB(LastWhisper) > 0) Then
                    If (Len(.Text) >= 2) Then
                        If StrComp(Split(.Text)(0), "/r", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .SelText = "/w " & CleanUsername(LastWhisper, True)
                            .SelStart = Len(.Text)
                        End If
                    End If
                    
                    If (Len(.Text) >= 6) Then
                        If StrComp(Split(.Text)(0), "/reply", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            .SelText = "/w " & CleanUsername(LastWhisper, True)
                            .SelStart = Len(.Text)
                        End If
                    End If
                End If
                
                If (LenB(LastWhisperTo) > 0) Then
                    If (Len(.Text) >= 3) Then
                        If StrComp(Split(.Text)(0), "/rw", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.Text)
                            
                            If StrComp(LastWhisperTo, "%f%") = 0 Then
                                .SelText = "/f m"
                            Else
                                .SelText = "/w " & CleanUsername(LastWhisperTo, True)
                            End If
                            
                            .SelStart = Len(.Text)
                        End If
                    End If
                End If
            End With

        Case vbKeyPageDown 'ALT + PAGEDOWN
            If Shift = vbAltMask Then
                With CurrentList
                    If CurrentUser > 0 And CurrentUser < .ListItems.Count Then
                        .HideSelection = False
                        With .ListItems.Item(CurrentUser + 1)
                            .Selected = True
                            .EnsureVisible
                        End With
                    End If
                End With

                cboSend.SetFocus
                cboSend.SelStart = SavedSelPos
                cboSend.SelLength = SavedSelLen
                Exit Sub
            End If

        Case vbKeyPageUp 'ALT + PAGEUP
            If Shift = vbAltMask Then
                With CurrentList
                    If CurrentUser > 1 Then
                        .HideSelection = False
                        With .ListItems.Item(CurrentUser - 1)
                            .Selected = True
                            .EnsureVisible
                        End With
                    End If
                End With

                cboSend.SetFocus
                cboSend.SelStart = SavedSelPos
                cboSend.SelLength = SavedSelLen
                Exit Sub
            End If

        Case vbKeyHome 'ALT+HOME
            If Shift = vbAltMask Then
                With CurrentList
                    If .ListItems.Count > 0 Then
                        .HideSelection = False
                        With .ListItems.Item(1)
                            .Selected = True
                            .EnsureVisible
                        End With
                    End If
                End With

                cboSend.SetFocus
                cboSend.SelStart = SavedSelPos
                cboSend.SelLength = SavedSelLen
                Exit Sub
            End If

        Case vbKeyEnd 'ALT+END
            If Shift = vbAltMask Then
                With CurrentList
                    If .ListItems.Count > 0 Then
                        .HideSelection = False
                        With .ListItems.Item(.ListItems.Count)
                            .Selected = True
                            .EnsureVisible
                        End With
                    End If
                End With

                cboSend.SetFocus
                cboSend.SelStart = SavedSelPos
                cboSend.SelLength = SavedSelLen
                Exit Sub
            End If

        Case vbKeyN, vbKeyInsert 'ALT + N or ALT + INSERT
            If (Shift = vbAltMask) Then
                If (Not (CurrentList.SelectedItem Is Nothing)) Then
                    Value = CurrentList.SelectedItem.Text
                    If cboSend.SelStart = 0 Then Value = Value & BotVars.AutoCompletePostfix
                    Value = Value & Space$(1)
                    
                    cboSend.SelText = Value
                    cboSend.SelStart = cboSend.SelStart + Len(Value)
                    Exit Sub
                End If
            End If

        Case vbKeyV 'PASTE
            Dim x() As String
            Dim n As Long
            
            If (IsScrolling(rtbChat)) Then
                LockWindowUpdate rtbChat.hWnd
            
                SendMessage rtbChat.hWnd, EM_SCROLL, SB_BOTTOM, &H0
                
                LockWindowUpdate &H0
            End If
        
            If (Shift = vbCtrlMask) Then
                On Error Resume Next
                
                If (InStr(1, Clipboard.GetText, Chr(13), vbTextCompare) <> 0) Then
                    x() = Split(Clipboard.GetText, Chr(10))
                    
                    If UBound(x) > 0 Then
                        For n = LBound(x) To UBound(x)
                            x(n) = Replace(x(n), Chr(13), vbNullString)
                            
                            If (x(n) <> vbNullString) Then
                                If (n <> LBound(x)) Then
                                    AddQ txtPre.Text & x(n) & txtPost.Text, PRIORITY.CONSOLE_MESSAGE
                                    
                                    cboSend.AddItem txtPre.Text & x(n) & txtPost.Text, 0
                                Else
                                    AddQ txtPre.Text & cboSend.Text & x(n) & txtPost.Text, _
                                        PRIORITY.CONSOLE_MESSAGE
                                    
                                    cboSend.AddItem txtPre.Text & cboSend.Text & x(n) & txtPost.Text, 0
                                End If
                            End If
                        Next n
                        
                        cboSend.Text = vbNullString
                        
                        MultiLinePaste = True
                    End If
                End If
            End If
            
        Case vbKeyA
            If (Shift = vbCtrlMask) Then
                If (CurrentTab <> LVW_BUTTON_CHANNEL) Then
                    ListviewTabs.Tab = LVW_BUTTON_CHANNEL
                    Call UpdateListviewTabs
                Else
                    cboSend.SelStart = 0
                    cboSend.SelLength = Len(cboSend.Text)
                End If
            End If
            
        Case vbKeyS
            If (Shift = vbCtrlMask) Then
                If (CurrentTab <> LVW_BUTTON_FRIENDS) And (ListviewTabs.TabEnabled(LVW_BUTTON_FRIENDS)) Then
                    ListviewTabs.Tab = LVW_BUTTON_FRIENDS
                    Call UpdateListviewTabs
                End If
            End If
            
        Case vbKeyD
            If (Shift = vbCtrlMask) Then
                If (CurrentTab <> LVW_BUTTON_CLAN) And (ListviewTabs.TabEnabled(LVW_BUTTON_CLAN)) Then
                    ListviewTabs.Tab = LVW_BUTTON_CLAN
                    Call UpdateListviewTabs
                End If
            End If
            
        Case vbKeyB
            If (Shift = vbCtrlMask) Then
                cboSend.SelText = "ÿcb"
            End If
            
        'Case vbKeyJ
        '    If (Shift = vbCtrlMask) Then
        '        Call mnuToggle_Click
        '    End If
            
        Case vbKeyU
            If (Shift = vbCtrlMask) Then
                cboSend.SelText = "ÿcu"
            End If
            
        Case vbKeyI
            If (Shift = vbCtrlMask) Then
                cboSend.SelText = "ÿci"
            End If
            
        Case vbKeyDelete
            ACUserIndex = 0
            
        Case vbKeyTab
            Dim ACList As Collection
            Dim ACUser As String
            Dim ACText As String
            
            If (Shift <> 0) Then
                Call cboSend_LostFocus
                
                If (txtPre.Visible = True) Then
                    Call txtPre.SetFocus
                Else
                    Call ListviewTabs.SetFocus
                    Call UpdateListviewTabs
                End If
            Else
            
                With cboSend
                    ' check if auto-complete active
                    If ACUserIndex = 0 Then
                        ' reset the static variables and auto-complete the current word
                        ACBuffer = .Text
                        
                        If .SelStart = 0 Then
                            ACWordStart = 1
                        Else
                            ACWordStart = InStrRev(ACBuffer, Space$(1), .SelStart, vbBinaryCompare) + 1
                        End If
                        ACWordEnd = InStr(.SelStart + 1, ACBuffer, Space$(1), vbBinaryCompare)
                        If ACWordEnd = 0 Then ACWordEnd = Len(ACBuffer) + 1
                        
                        ACUserPLen = ACWordEnd - ACWordStart
                        If ACUserPLen < 0 Then ACUserPLen = 0
                        
                        'AddChat vbWhite, ACWordEnd - ACWordStart
                        'AddChat vbWhite, Mid$(ACBuffer, ACWordStart, ACUserPLen)
                        
                        ACUserIndex = 1
                    Else
                        ' advance last auto-complete
                        ACUserIndex = ACUserIndex + 1
                    End If
                    
                    ' repopulate autocomplete candidates
                    ' sort as we get them
                    Set ACList = New Collection
                    'If ACUserPLen > 0 Then
                    For i = 1 To CurrentList.ListItems.Count
                        If StrComp(Left$(CurrentList.ListItems.Item(i).Text, ACUserPLen), Mid$(ACBuffer, ACWordStart, ACUserPLen), vbTextCompare) = 0 Then
                            For j = 1 To ACList.Count
                                If StrComp(CurrentList.ListItems.Item(i).Text, ACList.Item(j), vbTextCompare) < 0 Then
                                    Exit For
                                End If
                            Next j
                            
                            'AddChat vbYellow, j & " - " & CurrentList.ListItems.Item(i).Text
                            
                            ' add at found position
                            If j - 1 = ACList.Count Then
                                ACList.Add CurrentList.ListItems.Item(i).Text
                            Else
                                ACList.Add CurrentList.ListItems.Item(i).Text, , j
                            End If
                        End If
                    Next i
                    'End If
                    
                    ' set the user to the next candidate
                    ACUser = vbNullString
                    If ACList.Count > 0 Then
                        If ACList.Count < ACUserIndex Then
                            ACUserIndex = 1
                        End If
                        
                        ACUser = ACList.Item(ACUserIndex)
                    End If
                    Set ACList = Nothing
                    
                    ' save to the text box
                    If (LenB(ACUser) > 0) Then
                        If (ACWordStart > 1) Then
                            ACText = Left$(ACBuffer, ACWordStart - 1)
                        ElseIf (ACWordStart = 1) Then
                            ' only includes the postfix for the first word
                            ACUser = ACUser & BotVars.AutoCompletePostfix
                        End If
                        
                        ACText = ACText & ACUser & Space$(1)
                        
                        SavedSelPos = Len(ACText)
                        
                        If (ACWordEnd > 0) Then
                            ACText = ACText & Mid$(ACBuffer, ACWordEnd + 1)
                        End If
                        
                        .Text = ACText
                        .SelStart = SavedSelPos
                    Else
                        ACUserIndex = 0
                    End If
                End With
            End If
            
        Case vbKeyReturn
            Dim DoRunCommands     As Boolean
            Dim NoProcs()         As String
            Dim StartOutfilterPos As Long
            
            If (IsScrolling(rtbChat)) Then
                LockWindowUpdate rtbChat.hWnd
            
                SendMessage rtbChat.hWnd, EM_SCROLL, SB_BOTTOM, &H0
                
                LockWindowUpdate &H0
            End If
            
            DoRunCommands = True
        
            Select Case (Shift)
                Case vbShiftMask 'CTRL+ENTER - rewhisper
                    If LenB(cboSend.Text) > 0 Then
                        Value = "/w " & LastWhisperTo & Space(1)
                        DoRunCommands = False
                    End If
                    
                Case vbShiftMask Or vbCtrlMask 'CTRL+SHIFT+ENTER - reply
                    If LenB(cboSend.Text) > 0 Then
                        Value = "/w " & LastWhisper & Space(1)
                        DoRunCommands = False
                    End If
            
                Case Else 'normal ENTER - old rules apply
                    If LenB(cboSend.Text) > 0 Then
                        Value = vbNullString
                    End If
            End Select
            
            ' prefix box
            If txtPre.Visible Then Value = Value & txtPre.Text
            ' sendbox
            Value = Value & cboSend.Text
            ' suffix box
            If txtPost.Visible Then Value = Value & txtPost.Text
            
            If DoRunCommands Then
                On Error Resume Next
                DoRunCommands = Not RunInAll("Event_PressedEnter", Value)
                On Error GoTo 0
            End If
            
            If (Left$(Value, 6) = "/tell ") Then
                Value = "/w " & Mid$(Value, 7)
            End If
                
            If DoRunCommands Then
                If (Left$(Value, 1) = "/") Then
                    If (LenB(Config.ServerCommandList) > 0) Then
                        Dim Args() As String
                        Dim UserArg As String
                        Dim CommandList As String
                        
                        CommandList = Config.ServerCommandList
                        
                        ' please note: only commands with "%" as the first argument (or no "%" arguments)
                        ' (i.e. "/w %", "/ban %") are supported to be correctly no-processed
                        ' complexify here if that is needed
                        Args() = Split(Value, Space(1), 3)
                        If UBound(Args) > 0 Then UserArg = Args(1)
                        
                        CommandList = Replace(CommandList, "%", UserArg)
                        
                        NoProcs() = Split(CommandList, ",")
                    Else
                        ReDim NoProcs(0)
                    End If
                    
                    For i = LBound(NoProcs) To UBound(NoProcs)
                        If (LenB(NoProcs(i)) > 0) And (StrComp(Left$(Value, Len(NoProcs(i)) + 2), "/" & NoProcs(i) & Space$(1), vbTextCompare) = 0) Then
                            DoRunCommands = False
                            StartOutfilterPos = Len(NoProcs(i)) + 3
                            Exit For
                        End If
                    Next i
                    
                    If DoRunCommands Then
                        ProcessCommand GetCurrentUsername, Value, True, False
                    Else
                        DoRunCommands = False
                    End If
                Else
                    DoRunCommands = False
                End If
            End If
            
            If Not DoRunCommands Then
                ' Don't do replacements for a command unless it involves text that will be seen by someone else
                '  and don't replace text in the command itself or the target username
                Value = Left$(Value, StartOutfilterPos) & OutFilterMsg(Mid$(Value, StartOutfilterPos + 1))
                
                Call AddQ(Value, PRIORITY.CONSOLE_MESSAGE)
            End If
            
            'Ignore rest of code as the bot is closing
            If BotIsClosing Then
                Exit Sub
            End If
            
            cboSend.AddItem Value, 0
            
            cboSend.Text = vbNullString
            
            If Me.WindowState <> vbMinimized Then
                On Error Resume Next
                cboSend.SetFocus
            End If
    End Select
    
    CurrentList.HideSelection = True
    
    If (KeyCode <> vbKeyTab) Then
        ACUserIndex = 0
    End If

    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error " & Err.Number & " (" & Err.Description & ") " & _
        "in procedure cboSend_KeyDown"
        
    Exit Sub
End Sub

Private Sub cboSend_KeyPress(KeyAscii As Integer)
    Dim oldSelStart As Integer
    Dim sClosest As String
    
    'AddChat vbBlue, KeyAscii

    Select Case KeyAscii
        Case 1, 19, 4, 2, 9, 21, 13, 10
            KeyAscii = 0
        Case 22
            If MultiLinePaste Then
                KeyAscii = 0
                MultiLinePaste = False
            End If
            
    End Select
    
    With cboSend
        If (KeyAscii > 0) Then
            If (.ListCount > 15) Then
                .RemoveItem 15
            End If
        End If
    End With
    
    If Len(cboSend.Text) > 223 Then
        cboSend.ForeColor = vbRed
    Else
        cboSend.ForeColor = vbWhite
    End If
End Sub

Public Sub SControl_Error()
    Call modScripting.SC_Error
End Sub

Private Sub sckBNet_Close()
    sckBNet.Close
    If sckBNLS.State <> sckClosed Then sckBNLS.Close
    
    Call Event_BNetDisconnected
    
    ds.ClientToken = 0
    g_Connected = False
End Sub

Private Sub sckBNet_Connect()
    On Error Resume Next
    
    Call Event_BNetConnected
    
    If MDebug("all") Then
        AddChat COLOR_BLUE, "BNET CONNECT"
    End If
    
    Call modWarden.WardenCleanup(WardenInstance)
    WardenInstance = modWarden.WardenInitilize(sckBNet.SocketHandle)
    ds.Reset
    
    If (ProxyConnInfo(stBNCS).IsUsingProxy) Then
        modProxySupport.InitProxyConnection sckBNet, ProxyConnInfo(stBNCS), BotVars.Server, 6112
    Else
        InitBNetConnection
    End If
    
End Sub

Sub InitBNetConnection()
    Dim buf(0) As Byte
    buf(0) = BNCS_PROTOCOL_BNCS

    g_Connected = True

    Call modPacketBuffer.SendData(buf, 1, False, , sckBNet, stBNCS, phtNONE)

    If BotVars.BNLS Then
        modBNLS.SEND_BNLS_REQUESTVERSIONBYTE
    Else
        Select Case modBNCS.GetLogonSystem()
            Case modBNCS.BNCS_NLS:
                modBNCS.SEND_SID_AUTH_INFO
            Case modBNCS.BNCS_OLS:
                modBNCS.SEND_SID_CLIENTID2
                modBNCS.SEND_SID_LOCALEINFO
                modBNCS.SEND_SID_STARTVERSIONING
            Case modBNCS.BNCS_LLS:
                modBNCS.SEND_SID_CLIENTID
                modBNCS.SEND_SID_STARTVERSIONING
            Case Else:
                AddChat RTBColors.ErrorMessageText, StringFormat("Unknown Logon System Type: {0}", modBNCS.GetLogonSystem())
                AddChat RTBColors.ErrorMessageText, "Please visit http://www.stealthbot.net/sb/issues/?unknownLogonType for information regarding this error."
                DoDisconnect
        End Select
    End If
End Sub

Private Sub sckBNet_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Event_BNetError(Number, Description)
End Sub

Private Sub sckMCP_Close()
    AddChat RTBColors.ErrorMessageText, "[REALM] Connection closed."
    
    If Not ds.MCPHandler Is Nothing Then
        ds.MCPHandler.IsRealmError = True
        
        Call DoDisconnect
    End If
End Sub

Private Sub sckMCP_Connect()
    On Error GoTo ERROR_HANDLER
    
    Dim sIP   As String
    Dim lPort As Long
    
    If MDebug("all") Then
        AddChat COLOR_BLUE, "MCP CONNECT"
    End If
    
    If ProxyConnInfo(stMCP).IsUsingProxy Then
        AddChat RTBColors.SuccessText, "[REALM] [PROXY] Connected!"
        
        sIP = ds.MCPHandler.RealmSelectedServerIP
        lPort = ds.MCPHandler.RealmSelectedServerPort
        modProxySupport.InitProxyConnection sckMCP, ProxyConnInfo(stMCP), sIP, lPort
    Else
        AddChat RTBColors.SuccessText, "[REALM] Connected!"
        
        InitMCPConnection
    End If
    
    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in sckMCP_Connect()."

    Exit Sub
End Sub

Sub InitMCPConnection()
    Dim buf(0) As Byte
    buf(0) = BNCS_PROTOCOL_BNCS

    g_Connected = True

    Call modPacketBuffer.SendData(buf, 1, False, , sckMCP, stMCP, phtNONE)

    ds.MCPHandler.SEND_MCP_STARTUP
End Sub

Private Sub sckMCP_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ERROR_HANDLER
    
    Dim buf() As Byte
    Dim pBuff As clsDataBuffer
    
    If bytesTotal = 0 Then Exit Sub
    
    ' read buffer as Byte()
    sckMCP.GetData buf(), vbArray + vbByte, bytesTotal
    ' add data to buffer
    ReceiveBuffer(stMCP).InsertByteArr buf()

    If ProxyConnInfo(stMCP).IsUsingProxy And ProxyConnInfo(stMCP).Status <> psOnline Then
        Call modProxySupport.ProxyRecvPacket(sckMCP, ProxyConnInfo(stMCP), ReceiveBuffer(stMCP))
    Else
        Do While ReceiveBuffer(stMCP).IsFullPacket(stMCP)
            ' retrieve MCP packet
            Set pBuff = ReceiveBuffer(stMCP).TakePacket(stMCP)
            ' if MCP handler exists, parse
            If Not ds.MCPHandler Is Nothing Then
                Call ds.MCPHandler.MCPRecvPacket(pBuff)
            End If
            ' clean up
            Set pBuff = Nothing
        Loop
    End If
    
    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in sckMCP_DataArrival()."

    Exit Sub
End Sub

Private Sub sckMCP_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not g_Online Then
        ' This message is ignored if we've entered chat
        AddChat RTBColors.ErrorMessageText, "[REALM] Server error " & Number & ": " & Description
        
        If ProxyConnInfo(stMCP).IsUsingProxy And ProxyConnInfo(stMCP).Status <> psOnline Then
            DisplayError Number, True, stMCP
        ElseIf Not ds.MCPHandler Is Nothing Then
            If ds.MCPHandler.FormActive Then
                frmRealm.UnloadRealmError
            End If
        End If
    End If
End Sub


'// Written by Swent. Executes plugin timer subs
Private Sub scTimer_Timer()
    On Error Resume Next

    RunInAll "scTimer_Timer"
End Sub

' centralized "idle" events:
' IDLE MESSAGE (.5*IDLEMESSAGEDELAY minutes - user setting)
' PROFILE AMP (30 seconds)
' BNCS.SID_NULL (2 minutes - keep alive)
' BNCS.SID_CLANMOTD (10 minutes - may change)
' BNCS.SID_FRIENDSLIST (5 minutes - for D1,W2,D2 [no update], SC,W3 [bug in SID_FRIENDSUPDATE])
Private Sub tmrIdleTimer_Timer()
    On Error GoTo ERROR_HANDLER

    ' long-counter
    Static lCounter As Long
    Dim pBuf As clsDataBuffer

    lCounter = lCounter + 1
    If lCounter >= &H3C000000 Then lCounter = 0

    If sckBNet.State = sckConnected And (g_Online Or ds.AccountEntry) Then
        ' bot idle (30 second interval (x config value), offset 0 seconds)
        If g_Online And Config.IdleMessage Then
            Dim IdleWait As Long
            
            IdleWait = Config.IdleMessageDelay * 30
            
            If (lCounter Mod IdleWait) = 0 Then
                'AddQ "IDLE"
                Call tmrIdleTimer_Timer_IdleMsg
            End If
        End If

        ' bot ProfileAmp (30 second interval, offset 5 seconds)
        If g_Online And Config.ProfileAmp Then
            If (lCounter Mod 30&) = 5& Then
                'AddQ "PROFILE AMP"
                Call UpdateProfile
            End If
        End If

        ' BNCS keepalive... (1 minute interval; offset 15 seconds)
        ' if in clan, then (10 minute interval; offset 15 seconds before 4th minute)
        If g_Online And g_Clan.InClan And (lCounter Mod 600&) = 225& Then
            Call ClanHandler.RequestClanMOTD
        ' if friends list enabled, then (10 minute interval; offset 15 seconds before 9th minute)
        ElseIf g_Online And Config.FriendsListTab And (lCounter Mod 600&) = 525& Then
            ' request friendlist instead of FL
            'AddQ "FRIENDS"
            If (lvFriendList.ListItems.Count > 0) Then
                Call FriendListHandler.RequestFriendsList
            Else
                Set pBuf = New clsDataBuffer
                pBuf.SendPacket SID_NULL
                Set pBuf = Nothing
            End If
        ' else standard null (1 minute interval; offset 15 seconds before each minute)
        ElseIf (lCounter Mod 60&) = 45& Then
            'AddQ "NULL"
            Set pBuf = New clsDataBuffer
            pBuf.SendPacket SID_NULL
            Set pBuf = Nothing
        End If
    End If
    
    Exit Sub

ERROR_HANDLER:

    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in tmrIdleTimer_Timer()."
    
    Exit Sub
    
End Sub

Private Sub tmrIdleTimer_Timer_IdleMsg()
    On Error GoTo ERROR_HANDLER

    Dim U As String, IdleMsg As String, s() As String
    Dim IdleWaitS As String, IdleType As String
    Dim IdleWait As Integer, UDP As Byte
    Dim IsError As Boolean

    BotVars.JoinWatch = 0
    
    If Not Config.IdleMessage Then Exit Sub
    
    IdleMsg = Config.IdleMessageText
    IdleWait = Config.IdleMessageDelay
    IdleType = Config.IdleMessageType
    
    If IdleWait < 2 Then Exit Sub
    
    'If iCounter >= IdleWait Then
    'iCounter = 0
    'on error resume next
    If IdleType = "msg" Or IdleType = vbNullString Then
    
        If StrComp(IdleMsg, "null", vbTextCompare) = 0 Or IdleMsg = vbNullString Then
            Exit Sub
        End If
        
        IdleMsg = DoReplacements(IdleMsg)
        
        If (IdleMsg = vbNullString) Then
            IsError = True
        End If
        
    ElseIf IdleType = "uptime" Then
        IdleMsg = "/me -: System Uptime: " & ConvertTimeInterval(GetTickCountMS()) & " :: Connection Uptime: " & ConvertTimeInterval(GetConnectionUptime()) & " :: " & CVERSION & " :-"
        
    ElseIf IdleType = "mp3" Then
        Dim WindowTitle As String
        
        WindowTitle = MediaPlayer.TrackName
        
        If WindowTitle = vbNullString Then
            IsError = True
        ElseIf (MediaPlayer.IsPaused) Then
            IdleMsg = "/me -: Now Playing: " & WindowTitle & " (paused) :: " & CVERSION & " :-"
        Else
            IdleMsg = "/me -: Now Playing: " & WindowTitle & " :: " & CVERSION & " :-"
        End If

    ElseIf IdleType = "quote" Then
        U = g_Quotes.GetRandomQuote
        
        IdleMsg = "/me : " & U
        
        If Len(U) > 217 Then
            IsError = True
        End If
        
    End If
            
    If (IsError) Then
        IdleMsg = "/me -- " & CVERSION
    End If
    
    If sckBNet.State = sckConnected Then
        If InStr(1, IdleMsg, "& ", vbTextCompare) And IdleType = "msg" Then
            s = Split(IdleMsg, "& ")
            
            For IdleWait = LBound(s) To UBound(s)
                If Len(s(IdleWait)) > 215 Then
                    s(IdleWait) = Left$(s(IdleWait), 215)
                End If
                AddQ s(IdleWait)
            Next
        Else
            If Len(IdleMsg) > 215 Then
                IdleMsg = Left$(IdleMsg, 215)
            End If
            
            frmChat.AddQ IdleMsg
        End If
    End If
    'End If
    
    Exit Sub

ERROR_HANDLER:

    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in tmrIdleTimer_Timer_IdleMsg()."
    
    Exit Sub
    
End Sub

Private Sub tmrSilentChannel_Timer(Index As Integer)
    On Error GoTo ERROR_HANDLER

    Dim User    As clsUserObj
    Dim Item    As ListItem
    
    Dim i       As Integer
    Dim j       As Integer
    Dim found   As Boolean
    Dim WasZero As Boolean
    
    If (g_Channel.IsSilent = False) Then
        Exit Sub
    End If

    If (Index = 0) Then
        If (frmChat.mnuDisableVoidView.Checked = False) Then
            'For i = 1 To g_Channel.Users.Count
            '    ' with our doevents, we can miss our cue indicating that we
            '    ' need to stop silent channel processing and cause an rte.
            '    If (i > g_Channel.Users.Count) Then
            '        Exit For
            '    End If
            '
            '     Set user = g_Channel.Users(i)
            '
            '    If (lvChannel.FindItem(user.DisplayName) Is Nothing) Then
            '        Dim Stats As String
            '        Dim Clan  As String
            '
            '        ParseStatstring user.Statstring, Stats, Clan
            '
            '        AddName user.DisplayName, user.Game, user.Flags, user.Ping, user.Clan
            '    End If
            'Next i
            
            Call LockWindowUpdate(&H0)
            
            Call UpdateListviewTabs
        End If
    
        tmrSilentChannel(0).Enabled = False
    ElseIf (Index = 1) Then
        If (mnuDisableVoidView.Checked = False) Then
            If (g_Channel.IsSilent) Then
                Call g_Channel.ClearUsers
                
                frmChat.lvChannel.ListItems.Clear
            End If
        
            Call AddQ("/unsquelch " & GetCurrentUsername, PRIORITY.SPECIAL_MESSAGE)
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error: " & Err.Description & " in tmrSilentChannel_Timer(" & Index & ")."
    
    Exit Sub
End Sub

Private Sub txtPre_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cboSend_KeyPress(KeyAscii)
    End If
End Sub

Private Sub txtPost_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cboSend_KeyPress(KeyAscii)
    End If
End Sub

Sub ConnectBNLS()
    ' Don't try and connect if we don't have a server to connect to.
    If Len(BotVars.BNLSServer) = 0 Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] A working BNLS server could not be found."
        AddChat RTBColors.ErrorMessageText, "[BNLS]   Go to Settings -> Bot Settings -> Connection Settings -> Advanced and either set a server or use the automatic server finder."
        Call DoDisconnect
        Exit Sub
    End If
    
    Call Event_BNLSConnecting
    
    With sckBNLS
        If .State <> sckClosed Then .Close
        
        If ProxyConnInfo(stBNLS).IsUsingProxy Then
            .RemoteHost = ProxyConnInfo(stBNLS).ProxyIP
            .RemotePort = ProxyConnInfo(stBNLS).ProxyPort
        Else
            .RemoteHost = BotVars.BNLSServer
            .RemotePort = 9367
        End If
        .Connect
    End With
End Sub

Sub Connect()
    Dim NotEnoughInfo As Boolean
    Dim MissingInfo As String
    Dim i As Integer
    
    'g_username = BotVars.Username
    
    If sckBNet.State = sckClosed And sckBNLS.State = sckClosed Then
    
        'Vars
        NotEnoughInfo = False
        MissingInfo = "Information required to connect: "
        If BotVars.Username = vbNullString Then
            MissingInfo = MissingInfo & "Username, "
            NotEnoughInfo = True
        End If
        If BotVars.Password = vbNullString Then
            MissingInfo = MissingInfo & "Password, "
            NotEnoughInfo = True
        End If
        If BotVars.Server = vbNullString Then
            MissingInfo = MissingInfo & "Server, "
            NotEnoughInfo = True
        End If
        If BotVars.Product = vbNullString Then
            MissingInfo = MissingInfo & "your choice of Client, "
            NotEnoughInfo = True
        Else
            Select Case GetProductInfo(BotVars.Product).KeyCount
                Case 0
                Case 1
                    If BotVars.CDKey = vbNullString Then
                        MissingInfo = MissingInfo & "CDKey, "
                        NotEnoughInfo = True
                    End If
                Case 2
                    If BotVars.CDKey = vbNullString Then
                        MissingInfo = MissingInfo & "CDKey, "
                        NotEnoughInfo = True
                    End If
                    If BotVars.ExpKey = vbNullString Then
                        MissingInfo = MissingInfo & "expansion CDKey, "
                        NotEnoughInfo = True
                    End If
            End Select
        End If
        
        If NotEnoughInfo Then
            MsgBox "You haven't provided enough information to connect! " & _
                "Please edit your connection settings by choosing Bot Settings under the Settings menu." & _
                vbNewLine & Left$(MissingInfo, Len(MissingInfo) - 2) & ".", vbInformation
            
            Call DoDisconnect(1)
            
            Exit Sub
        End If
        
        For i = 0 To 2
            ProxyConnInfo(i).IsUsingProxy = ProxyConnInfo(i).UseProxy
            If ProxyConnInfo(i).IsUsingProxy And (ProxyConnInfo(i).ProxyPort = 0 Or LenB(ProxyConnInfo(i).ProxyIP) = 0) Then
                MsgBox "You have selected to use a proxy for one or more connections, but no proxy is configured. Please set one up in the Advanced " & _
                    " section of Bot Settings or disable Use Proxy.", vbInformation
                    
                Call DoDisconnect(1)
                
                Exit Sub
            End If
        Next i
        
        SetTitle "Connecting..."
        
        If ((StrComp(BotVars.Product, "PX2D", vbTextCompare) = 0) Or _
            (StrComp(BotVars.Product, "VD2D", vbTextCompare) = 0)) Then
            
            Dii = True
        Else
            Dii = False
        End If

        'Changed 10-07-2009 - Hdx - Are we going to have private betas anymore?
        #If (BETA = 2) Then
            Call AddChat(RTBColors.InformationText, "Authorizing your private-release, please wait...")
            
            If (GetAuth(BotVars.Username)) Then
                Call AddChat(RTBColors.SuccessText, "Private usage authorized, connecting your bot...")
                
                ' was auth function bypassed?
                If (AUTH_CHECKED = False) Then BotVars.Password = Chr$(0)
            Else
                Call AddChat(RTBColors.ErrorMessageText, "- - - - - YOU ARE NOT AUTHORIZED TO USE THIS PROGRAM - - - - -")
                
                Call DoDisconnect
                UserCancelledConnect = False
                Exit Sub
            End If
        #Else
            AddChat RTBColors.InformationText, "Connecting your bot..."
        #End If
        
        If BotVars.BNLS Then
            If Len(BotVars.BNLSServer) = 0 Then
                If BotVars.UseAltBnls Then
                    Call FindBNLSServer
                    
                    Exit Sub
                End If
            End If
            
            Call ConnectBNLS
        Else
            Call Event_BNetConnecting
        End If
        
        With sckBNet
            If .State <> sckClosed Then
                AddChat RTBColors.ErrorMessageText, "Already connected."
                Exit Sub
            End If
    
            If ProxyConnInfo(stBNCS).IsUsingProxy Then
                .RemoteHost = ProxyConnInfo(stBNCS).ProxyIP
                .RemotePort = ProxyConnInfo(stBNCS).ProxyPort
            Else
                .RemoteHost = BotVars.Server
                .RemotePort = 6112
            End If

            If Not BotVars.BNLS Then .Connect

        End With
        
    End If
    
    Exit Sub
    
Error:
    MsgBox "Configuration file error. Please re-write your configuration file " & _
        "using the Setup dialog.", vbCritical + vbOKOnly, "Error"
    
    Call SetTitle("Disconnected")
    
    Exit Sub
End Sub

Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True)
    Dim i As Integer
    If AllowEvents Then
        For i = 0 To (1000 * fSeconds) \ 100
            Sleep 100
            DoEvents
        Next i
    Else
        Sleep fSeconds * 1000
    End If
End Sub

'/* Fires every second */
Private Sub UpTimer_Timer()

    On Error GoTo ERROR_HANDLER

    Dim newColor  As Long
    Dim i         As Integer
    Dim pos       As Integer
    Dim doCheck   As Boolean
    
    If (floodCap > 2) Then
        floodCap = floodCap - 3
    End If
    
    If (VoteDuration > 0) Then
        VoteDuration = VoteDuration - 1
        
        If (VoteDuration = 0) Then
            Dim s As String
            
            s = Voting(BVT_VOTE_END)
            
            If (Len(s) > 1) Then
                AddQ s
            End If
        End If
    End If
    
    If (g_Queue.Count > 0) Then
        Ban vbNullString, 0, 3
    End If

    If (g_Channel.IsSilent = False) Then
        doCheck = True
    
        For i = 1 To g_Channel.Users.Count
            With g_Channel.Users(i)
                If (g_Channel.Self.IsOperator) Then
                    If (.IsOperator = False) Then
                        ' channel password
                        If ((BotVars.ChannelPasswordDelay > 0) And (Len(BotVars.ChannelPassword) > 0)) Then
                            If (.PassedChannelAuth = False) Then
                                If (.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
                                    If (GetSafelist(.DisplayName) = False) Then
                                        Ban .DisplayName & " Password time is up", (AutoModSafelistValue - 1)
                                         
                                        doCheck = False
                                    End If
                                End If
                            End If
                        End If
                        
                        ' idle bans
                        If ((doCheck) And ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0))) Then
                            If (.TimeSinceTalk() > BotVars.IB_Wait) Then
                                If (GetSafelist(.DisplayName) = False) Then
                                    Ban .DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", _
                                        (AutoModSafelistValue - 1), IIf(BotVars.IB_Kick, 1, 0)
                                        
                                    doCheck = False
                                End If
                            End If
                        End If
                    End If
                End If
                
                If (Not BotVars.NoColoring) Then
                    pos = frmChat.GetChannelItemIndex(.Name)
                
                    If (pos > 0) Then
                        newColor = GetNameColor(.Flags, .TimeSinceTalk, StrComp(.DisplayName, _
                            GetCurrentUsername, vbBinaryCompare) = 0)
                        
                        If (lvChannel.ListItems(pos).ForeColor <> newColor) Then
                            lvChannel.ListItems(pos).ForeColor = newColor
                        End If
                    End If
                End If
            End With
            
            doCheck = True
        Next i
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in UpTimer_Timer()."

    Exit Sub
    
End Sub

'StealthLock (c) 2003 Stealth, Please do not remove this header
Private Function GetAuth(ByVal Username As String) As Long
    On Error GoTo ERROR_HANDLER

    Static lastAuth     As Long
    Static lastAuthName As String
    
    Dim clsCRC32 As clsCRC32
    Dim hostFile As String
    Dim hostPath As String
    Dim f        As Integer
    Dim tmp      As String
    Dim Result   As Integer      ' string variable for storing beta authorization result
                                 ' 0  == unauthorized
                                 ' >0 == authorized
    Set clsCRC32 = New clsCRC32
    If (lastAuth) Then
        If (StrComp(Username, lastAuthName, vbTextCompare) = 0) Then
            GetAuth = lastAuth
            
            Exit Function
        End If
    End If
    
    f = FreeFile
    
    If (g_OSVersion.IsWindowsNT) Then
        hostPath = _
            GetRegistryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\", _
                "DatabasePath")
        
        If (Len(hostPath) = 0) Then
            hostPath = "%SystemRoot%\system32\drivers\etc\"
        End If
    Else
        hostPath = "%WinDir%"
    End If
    
    hostPath = ReplaceEnvironmentVars(hostPath & "\hosts")
 
    If (LenB(Dir$(hostPath)) > 0) Then
        Open hostPath For Input As #f
            Do While (EOF(f) = False)
                Line Input #f, tmp
                
                hostFile = hostFile & tmp
            Loop
        Close #f
    End If
    
    If (clsCRC32.CRC32(StringToByteArr(BETA_AUTH_URL)) = BETA_AUTH_URL_CRC32) Then
        If (InStr(1, hostFile, Split(BETA_AUTH_URL, ".")(1), vbTextCompare) = 0) Then
            Result = CInt(Val(Inet.OpenURL(BETA_AUTH_URL & Username)))
        End If
    End If
    
    Do While Inet.StillExecuting
        DoEvents
    Loop
    
    If (Result = 1) Then
        lastAuth = Result
        lastAuthName = Username
    
        GetAuth = True
        
        AUTH_CHECKED = True
    End If
    
    Set clsCRC32 = Nothing

    Exit Function

ERROR_HANDLER:

    AddChat RTBColors.ErrorMessageText, "Beta Auth Error: #", vbRed, Err.Number, vbRed, ": ", vbRed, Err.Description
    Set clsCRC32 = Nothing
    GetAuth = False
    Exit Function
    
End Function

' http://www.go4expert.com/forums/showthread.php?t=208
Private Function ReplaceEnvironmentVars(ByVal str As String) As String

    Dim i     As Integer
    Dim Name  As String
    Dim Value As String
    Dim tmp   As String
    
    tmp = str
    
    i = 1

    While (Environ$(i) <> "")
        Name = Mid$(Environ$(i), 1, InStr(1, Environ$(i), "=") - 1)

        Value = Mid$(Environ$(i), InStr(1, Environ$(i), "=") + 1)

        tmp = Replace(tmp, "%" & Name & "%", Value)
    
        i = i + 1
    Wend

    ReplaceEnvironmentVars = tmp

End Function

Function AddQ(ByVal Message As String, Optional msg_priority As Integer = -1, Optional ByVal User As String = _
    vbNullString, Optional ByVal Tag As String = vbNullString, Optional OversizeDelimiter As String = " ") As Integer

    On Error GoTo ERROR_HANDLER
    
    Dim Splt()         As String
    Dim strTmp         As String
    Dim i              As Long
    Dim currChar       As Long
    Dim Send           As String
    Dim Command        As String
    Dim GTC            As Double
    Dim Q              As clsQueueObj
    Dim delay          As Long
    Dim Index          As Long
    Dim s              As String      ' temp string for settings
    Dim MaxLength      As Integer     ' stores max length for split (with override)
    
    Static LastGTC  As Double
    Static BanCount As Integer

    strTmp = Message
    
    ' cap priority at 100
    If (msg_priority > 100) Then
        msg_priority = 100
    End If
    
    If (g_Queue.Count = 0) Then
        BanCount = 0
    End If
    
    If (strTmp <> vbNullString) Then
        ReDim Splt(0)
    
        ' check for tabs and replace with spaces (2005-09-23)
        If (InStr(1, strTmp, Chr$(9), vbBinaryCompare) <> 0) Then
            strTmp = Replace$(strTmp, Chr$(9), Space$(4))
        End If
        
        ' check for invalid characters in the message
        For i = 1 To Len(strTmp)
            currChar = Asc(Mid$(strTmp, i, 1))
        
            If (currChar < 32) Then
                Exit Function
            End If
        Next i
        
        ' is this an internal or battle.net command?
        If (StrComp(Left$(strTmp, 1), "/", vbBinaryCompare) = 0) Then
            ' if so, we have extra work to do
            For i = 2 To Len(strTmp)
                currChar = Asc(Mid$(strTmp, i, 1))
            
                ' find the first non-space after the /
                If (Not currChar = Asc(Space(1))) Then
                    Exit For
                End If
            Next i
            
            ' if we found a non-space, strip everything
            If (i >= 2) Then
                strTmp = StringFormat("/{0}", Mid$(strTmp, i))
            End If

            ' Find the next instance of a space (the end of the command word)
            Index = InStr(1, strTmp, Space(1), vbBinaryCompare)
            
            ' is it a valid command word?
            If (Index > 2) Then
                ' extract the command word
                Command = LCase$(Mid$(strTmp, 2, (Index - 2)))

                ' test it for being battle.net commands we need to process now
                If ((Command = "w") Or _
                    (Command = "whisper") Or _
                    (Command = "m") Or _
                    (Command = "msg") Or _
                    (Command = "message") Or _
                    (Command = "whois") Or _
                    (Command = "where") Or _
                    (Command = "whereis") Or _
                    (Command = "squelch") Or _
                    (Command = "unsquelch") Or _
                    (Command = "ignore") Or _
                    (Command = "unignore") Or _
                    (Command = "ban") Or _
                    (Command = "unban") Or _
                    (Command = "kick") Or _
                    (Command = "designate")) Then
        
                    Splt() = Split(strTmp, Space$(1), 3)
                    
                    If (UBound(Splt) > 0) Then
                        Command = StringFormat("{0} {1}", Splt(0), ReverseConvertUsernameGateway(Splt(1)))

                        If ((g_Channel.IsSilent) And (frmChat.mnuDisableVoidView.Checked = False)) Then
                            If ((LCase$(Splt(0)) = "/unignore") Or (LCase$(Splt(0)) = "/unsquelch")) Then
                                If (StrComp(Splt(1), GetCurrentUsername, vbTextCompare) = 0) Then
                                    lvChannel.ListItems.Clear
                                End If
                            End If
                        End If
                        
                        If (UBound(Splt) > 1) Then
                            ReDim Preserve Splt(0 To UBound(Splt) - 1)
                        End If
                    End If
                    
                ElseIf ((Command = "f") Or _
                        (Command = "friends")) Then
                    
                    Splt() = Split(strTmp, Space$(1), 3)
                    
                    Command = Splt(0)
                    
                    If (UBound(Splt) >= 1) Then
                        Command = StringFormat("{0} {1}", Command, Splt(1))
                        
                        If (UBound(Splt) >= 2) Then
                            Select Case (LCase$(Splt(1)))
                                Case "m", "msg"
                                    ReDim Preserve Splt(0 To UBound(Splt) - 1)

                                Case Else
                                    Splt() = Split(strTmp, Space$(1), 4)
                                
                                    If ((StrReverse$(BotVars.Product) = PRODUCT_WAR3) Or _
                                        (StrReverse$(BotVars.Product) = PRODUCT_W3XP)) Then
                                        
                                        Command = StringFormat("{0} {1}", Command, ReverseConvertUsernameGateway(Splt(2)))
                                    Else
                                        Command = StringFormat("{0} {1}", Command, Splt(2))
                                    End If
                                    
                                    If (UBound(Splt) >= 3) Then
                                        Command = StringFormat("{0} {1}", Command, Splt(3))
                                    End If
                            End Select
                        End If
                    End If
                Else
                    Command = StringFormat("/{0}", Command)
                    strTmp = Mid$(strTmp, Len(Command) + 2)
                End If
                
                If (Len(Command) >= BNET_MSG_LENGTH) Then
                    Exit Function
                End If

                If (UBound(Splt) > 0) Then
                    strTmp = Mid$(strTmp, _
                        (Len(Join(Splt(), Space$(1))) + (Len(Space$(1))) + 1))
                End If
            End If
        End If
        
        If (msg_priority < 0) Then
            Dim cmdName    As String
            Dim spaceIndex As Long

            spaceIndex = InStr(1, Message, Space$(1), vbBinaryCompare)
            
            If (spaceIndex >= 2) Then
                cmdName = LCase$(Left$(Mid$(Message, 2), spaceIndex - 2))
            Else
                cmdName = LCase$(Mid$(Message, 2))
            End If
            
            Select Case (cmdName)
                Case "designate": msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "resign":    msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "who":       msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "unban":     msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "clan", "c": msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "ban":       msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE
                Case "kick":      msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE
                Case Else:        msg_priority = PRIORITY.MESSAGE_DEFAULT
            End Select
        End If
        
        MaxLength = Config.MaxMessageLength
        
        Call SplitByLen(strTmp, (MaxLength - Len(Command)), Splt(), vbNullString, , OversizeDelimiter)

        ReDim Preserve Splt(0 To UBound(Splt))

        ' add to the queue!
        For i = LBound(Splt) To UBound(Splt)
            ' store current tick
            GTC = GetTickCount()
            
            ' store working copy
            Send = Splt(i)
            If (LenB(Command) > 0) Then
                If (LenB(Send) > 0) Then
                    Send = StringFormat("{0} {1}", Command, Send)
                Else
                    Send = Command
                End If
            ElseIf (Left(Send, 1) = "/" And i > LBound(Splt)) Then
                Send = StringFormat(" {0}", Send)
            End If
            
            ' create the queue object
            Set Q = New clsQueueObj
            
            With Q
                .Message = Send
                .PRIORITY = msg_priority
                .Tag = Tag
            End With

            ' add it
            g_Queue.Push Q
            
            ' should we subject this message to the typical delay,
            ' or can we get it out of here a bit faster?  If we
            ' want it out of here quick, we need an empty queue
            ' and have had at least 10 seconds elapse since the
            ' previous message.
            If (g_Queue.Count() = 1) Then
                If (GTC - LastGTC >= 10000) Then
                    ' set default message delay when queue is empty (in ms)
                    delay = 10
                    
                    ' are we issuing a ban or kick command?
                    If (msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE) Then
                        delay = g_BNCSQueue.BanDelay()
                    End If
                End If
            End If
            
            ' If queueTimerID is 0 the timer is idle right now, so reset it
            If QueueTimerID = 0 Then
                If delay = 0 Then
                    delay = g_BNCSQueue.GetDelay(g_Queue.Peek.Message)
                End If
            
                ' set the delay before our next queue cycle
                QueueTimerID = SetTimer(0&, 0&, delay, AddressOf QueueTimerProc)
            End If
        Next i
        
        AddQ = UBound(Splt) + 1
        
        ' store our tick for future reference
        LastGTC = GTC
    End If
        
    Exit Function
    
ERROR_HANDLER:
    Call AddChat(vbRed, "Error: " & Err.Description & " in AddQ().")

    Exit Function
End Function


Sub ClearChannel()
    ' reset channel object
    Set g_Channel = New clsChannelObj
    
    ' clear channel UI elements
    lvChannel.ListItems.Clear
    lblCurrentChannel.Caption = vbNullString
    
    ' be ready to ignore processing EID_INFO Clan MOTD
    If g_Clan.InClan Then
        g_Clan.PendingClanMOTD = True
    End If
End Sub


Sub ReloadConfig(Optional Mode As Byte = 0)
    On Error GoTo ERROR_HANDLER

    Dim default_group_access As clsDBEntryObj
    Dim s                    As String
    Dim i                    As Integer
    Dim f                    As Integer
    Dim Index                As Integer
    Dim bln                  As Boolean
    Dim doConvert            As Boolean
    Dim command_output()     As String
    
    Dim oCommandGenerator    As clsCommandGeneratorObj
    
    If Mode <> 0 Then
        Config.Load GetConfigFilePath()
    End If
    
    BotVars.TSSetting = Config.TimestampMode

    ' Client settings
    If LenB(BotVars.Username) > 0 And StrComp(BotVars.Username, Config.Username, vbTextCompare) <> 0 Then
        AddChat RTBColors.ConsoleText, "Username set to " & Config.Username & "."
    End If
    BotVars.Username = Config.Username
    BotVars.Password = Config.Password
    BotVars.CDKey = Config.CDKey
    BotVars.ExpKey = Config.ExpKey
    BotVars.Product = StrReverse$(GetProductInfo(Config.Game).Code)
    BotVars.Server = Config.Server
    BotVars.HomeChannel = Config.HomeChannel
    BotVars.BotOwner = Config.BotOwner
    
    BotVars.Trigger = Config.Trigger
    'If (BotVars.TriggerLong = vbNullString) Then
    '    BotVars.Trigger = "."
    'End If
    
    BotVars.BNLSServer = Config.BNLSServer
    BotVars.ReconnectDelay = Config.ReconnectDelay
    
    ' Load database and commands
    Call Database.Load(GetFilePath(FILE_USERDB))
    'Call oCommandGenerator.GenerateCommands
    
    ' Set UI fonts
    If Mode <> 1 Then
        Dim ResizeChatElements As Boolean
        Dim ResizeChannelElements As Boolean
        
        s = Config.ChatFont
        If s <> vbNullString And s <> rtbChat.Font.Name Then
            'rtbChat.Font.Name = s
            cboSend.Font.Name = s
            'rtbWhispers.Font.Name = s
            txtPre.Font.Name = s
            txtPost.Font.Name = s
            
            ResizeChatElements = True
        End If
        
        s = Config.ChannelListFont
        If s <> vbNullString And s <> lvChannel.Font.Name Then
            lvChannel.Font.Name = s
            lvClanList.Font.Name = s
            lvFriendList.Font.Name = s
            lblCurrentChannel.Font.Name = s
            ListviewTabs.Font.Name = s
            
            ResizeChatElements = True
        End If
        
        s = Config.ChatFontSize
        If StrictIsNumeric(s) Then
            If CInt(s) <> rtbChat.Font.Size Then
                'rtbChat.Font.Size = s
                cboSend.Font.Size = s
                'rtbWhispers.Font.Size = s
                txtPre.Font.Size = s
                txtPost.Font.Size = s
                
                ResizeChannelElements = True
            End If
        End If
        
        s = Config.ChannelListFontSize
        If StrictIsNumeric(s) Then
            If CInt(s) <> lvChannel.Font.Size Then
                lvChannel.Font.Size = s
                lvClanList.Font.Size = s
                lvFriendList.Font.Size = s
                lblCurrentChannel.Font.Size = s
                ListviewTabs.Font.Size = s
                
                ResizeChannelElements = True
            End If
        End If
    
        If ResizeChannelElements Then
            Dim lblHeight As Single
            
            lblCurrentChannel.AutoSize = True
            lblHeight = lblCurrentChannel.Height + 40
            lblCurrentChannel.AutoSize = False
            lblCurrentChannel.Height = lblHeight
            
            ResizeChatElements = True
        End If
        
        If ResizeChatElements Then
            Call ChangeRTBFont(rtbChat, Config.ChatFont, Config.ChannelListFontSize)
            Call ChangeRTBFont(rtbWhispers, Config.ChatFont, Config.ChannelListFontSize)
            
            Form_Resize
        End If
    End If
    
    Filters = Config.ChatFilters
    mnuToggleFilters.Checked = Filters
    If (Not Filters) Then
        BotVars.JoinWatch = 0
    End If
    
    BotVars.AutofilterMS = 0
    
    AutoModSafelistValue = Config.AutoSafelistLevel
    BotVars.ShowOfflineFriends = Config.ShowOfflineFriends
    
    BotVars.RetainOldBans = Config.RetainOldBans
    BotVars.StoreAllBans = Config.StoreAllBans
    
    BotVars.GatewayConventions = Config.NamespaceConvention
    BotVars.UseD2Naming = Config.UseD2Naming
    BotVars.D2NamingFormat = Config.D2NamingFormat
    
    BotVars.ShowStatsIcons = Config.ShowStatsIcons
    BotVars.ShowFlagsIcons = Config.ShowFlagIcons
    
    JoinMessagesOff = Not Config.ShowJoinLeaves
    mnuToggle.Checked = JoinMessagesOff

    mail = Config.BotMail
    
    BotVars.BanEvasion = Config.BanEvasion
    BotVars.Logging = Config.LoggingMode
    
    mnuToggleWWUse.Checked = Config.WhisperWindows
    BotVars.WhisperCmds = Config.WhisperCommands
    PhraseBans = Config.PhraseBans
    BotVars.CaseSensitiveFlags = Config.CaseSensitiveDBFlags
    BotVars.AutoCompletePostfix = Config.AutoCompletePostfix
    BotVars.BNLS = Config.UseBNLS
    BotVars.LogDBActions = Config.LogDBActions
    BotVars.LogCommands = Config.LogCommands

    '/* time to idle: defaults to 600 seconds / 10 minutes idle */
    BotVars.SecondsToIdle = Config.SecondsToIdle
    
    BotVars.BanUnderLevel = Config.LevelBanW3
    BotVars.BanUnderLevelMsg = Config.LevelBanMessage
    BotVars.BanPeons = Config.PeonBan
    
    BotVars.KickOnYell = Config.KickOnYell
    
    ' Capped at 32767, topic=29986 -Andy
    BotVars.IB_Wait = Config.IdleBanDelay

    ' Set and create special groups
    BotVars.DefaultShitlistGroup = Config.ShitlistGroup
    BotVars.DefaultTagbansGroup = Config.TagbanGroup
    BotVars.DefaultSafelistGroup = Config.SafelistGroup
    Call Database.CreateSpecialGroups
    
    BotVars.DisableMP3Commands = Not Config.Mp3Commands
    
    BotVars.MaxBacklogSize = Config.MaxBacklogSize
    BotVars.MaxLogFileSize = Config.MaxLogFileSize
    
    If Config.UrlDetection Then
        EnableURLDetect rtbChat.hWnd
    Else
        DisableURLDetect rtbChat.hWnd
    End If
    
    ' reload quotes file
    Set g_Quotes = New clsQuotesObj
    
    BotVars.UseBackupChan = Config.UseBackupChannel
    BotVars.BackupChan = Config.BackupChannel

    mnuUTF8.Checked = Config.UseUTF8
    
    mnuToggleShowOutgoing.Checked = Config.ShowOutgoingWhispers
    mnuHideWhispersInrtbChat.Checked = Config.HideWhispersInMain
    
    'LoadSafelist
    LoadArray LOAD_PHRASES, Phrases()
    LoadArray LOAD_FILTERS, gFilters()
    LoadArray LOAD_BLOCKLIST, g_Blocklist()
    
    ProtectMsg = Config.ChannelProtectionMessage

    Call LoadOutFilters
    
    BotVars.IB_On = Config.IdleBan
    BotVars.IB_Kick = Config.IdleBanKick
    BotVars.IB_Wait = Config.IdleBanDelay

    BotVars.Spoof = Config.PingSpoofing
    
    Protect = Config.ChannelProtection
    BotVars.UseUDP = Config.UseUDP
    BotVars.IPBans = Config.IPBans
    BotVars.UseAltBnls = Config.BNLSFinder
    BotVars.QuietTime = Config.QuietTime
    
    mnuFlash.Checked = Config.FlashOnEvents
    
    BotVars.NoTray = Not Config.MinimizeToTray
    BotVars.NoAutocompletion = Not Config.NameAutoComplete
    BotVars.NoColoring = Not Config.NameColoring
    
    mnuDisableVoidView.Checked = Not Config.VoidView
    
    BotVars.MediaPlayer = Config.MediaPlayer
    
    
    ' Load some queue stuff, reluctantly
    BotVars.QueueMaxCredits = Config.QueueMaxCredits
    BotVars.QueueCostPerPacket = Config.QueueCostPerPacket
    BotVars.QueueCostPerByte = Config.QueueCostPerByte
    BotVars.QueueCostPerByteOverThreshhold = Config.QueueCostPerByteOver
    BotVars.QueueStartingCredits = Config.QueueStartingCredits
    BotVars.QueueThreshholdBytes = Config.QueueThresholdBytes
    BotVars.QueueCreditRate = Config.QueueCreditRate

    BotVars.UseRealm = Config.UseD2Realms

    txtPre.Text = ""
    txtPost.Text = ""
    
    txtPre.Visible = Not Config.DisablePrefixBox
    mnuPopAddLeft.Enabled = Not Config.DisablePrefixBox
    mnuPopFLAddLeft.Enabled = Not Config.DisablePrefixBox
    mnuPopClanAddLeft.Enabled = Not Config.DisablePrefixBox

    txtPost.Visible = Not Config.DisableSuffixBox
    
    '[Other] MathAllowUI - Will allow People to use MessageBox/InputBox or other UI related commands in the .eval/.math commands ~Hdx 09-25-07
    SCRestricted.AllowUI = Config.MathAllowUI
    BotVars.NoRTBAutomaticCopy = Config.DisableRTBAutoCopy
    
    BotVars.UseGreet = Config.GreetMessage
    BotVars.GreetMsg = Config.GreetMessageText
    BotVars.WhisperGreet = Config.WhisperGreet
    
    BotVars.ChatDelay = Config.ChatDelay
    
    s = Config.GetFilePath("Logs")
    If (Not (s = vbNullString)) Then
        g_Logger.LogPath = s
    End If
    
    Call Form_Resize
    
    If (g_Online) Then
        Dim found       As ListItem
        Dim outbuf      As String
        Dim ChannelUser As clsUserObj
        Dim FriendObj   As clsFriendObj
        Dim Member      As clsClanMemberObj

        SetTitle GetCurrentUsername & ", online in channel " & g_Channel.Name
        
        frmChat.UpdateTrayTooltip

        lvChannel.ListItems.Clear
        For i = 1 To g_Channel.Users.Count
            Set ChannelUser = g_Channel.Users(i)

            AddName ChannelUser.DisplayName, ChannelUser.Name, ChannelUser.Game, ChannelUser.Flags, ChannelUser.Ping, _
                    ChannelUser.Stats.IconCode, ChannelUser.Clan
        Next i

        lvFriendList.ListItems.Clear
        If Config.FriendsListTab Then
            For i = 1 To g_Friends.Count
                Set FriendObj = g_Friends.Item(i)
                If FriendObj.IsOnline Or Config.ShowOfflineFriends Then
                    AddFriendItem FriendObj.DisplayName, FriendObj.Game, FriendObj.Status, FriendObj.LocationID, i - 1
                End If
                Set FriendObj = Nothing
            Next i

            ' re-sort
            lvFriendList.Sorted = True
        End If

        lvClanList.ListItems.Clear
        If g_Clan.InClan Then
            For i = 1 To g_Clan.Members.Count
                Set Member = g_Clan.Members.Item(i)
                AddClanMember Member.Name, Member.DisplayName, Member.Rank, Member.Status
                Set Member = Nothing
            Next i

            ' re-sort
            lvClanList.Sorted = True
        End If

        Call UpdateListviewTabs
    End If
    
    For i = 0 To 2
        With ProxyConnInfo(i)
            .ServerType = i
            Select Case i
                Case stBNCS: .UseProxy = Config.UseProxy
                Case stBNLS: .UseProxy = Config.UseProxy And Config.ProxyBNLS
                Case stMCP:  .UseProxy = Config.UseProxy And Config.ProxyMCP
                Case Else:   .UseProxy = False
            End Select
            
            If .UseProxy Then
                ' set these values so that next connection attempt uses them--
                ' they may not be accurate for the current connection so
                ' use the values on the socket to get current IP/Port
                .ProxyIP = Config.ProxyIP
                .ProxyPort = Config.ProxyPort
                If CBool(StrComp(Config.ProxyType, PROXY_SETTING_SOCKS5, vbTextCompare) = 0) Then
                    .Version = 5
                Else
                    .Version = 4
                End If
                .Username = Config.ProxyUsername
                .Password = Config.ProxyPassword
                .RemoteResolveHost = Config.ProxyServerResolve
                
                ' do not set RemoteIP, RemotePort, RemoteHostName, Status, IsUsingProxy;
                ' those are set on proxy connect and shouldn't be touched by the config
            End If
        End With
    Next i

    Call ChatQueue_Initialize

    ' I reluctantly add the queue variables here.
    
    If (g_Online) Then
        Call g_Channel.CheckUsers
    Else
        Err.Clear
     
        '//Removed 10/29/09 - Hdx - I'll add in this feature later properly, does not work as is.
        'If (ReadCfg(OV, "LocalIP") <> vbNullString) Then
        '    If (Err.Number = 0) Then: sckBNet.bind , ReadCfg(OV, "LocalIP")
        '    If (Err.Number = 0) Then: sckBNLS.bind , ReadCfg(OV, "LocalIP")
        '    If (Err.Number = 0) Then: sckMCP.bind , ReadCfg(OV, "LocalIP")
        'End If
    End If
    
    ' disable the script system if override is set
    modScripting.SetScriptSystemDisabled Config.DisableScripting
    
    Set oCommandGenerator = Nothing
    
    Exit Sub

ERROR_HANDLER:
    If (Err.Number = 10049) Then
        AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in ReloadConfig()."
    End If
    
    Resume Next
End Sub

Private Sub ChangeRTBFont(rtb As RichTextBox, ByVal NewFont As String, ByVal NewSize As Integer)
    Dim tmpBuffer As String
    
    With rtb
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = NewSize
        .SelFontName = NewFont
        tmpBuffer = .TextRTF
        .Text = vbNullString
        .Font.Name = NewFont
        .Font.Size = NewSize
        .TextRTF = tmpBuffer
        .SelStart = Len(.Text)
    End With
End Sub

'returns OK to Proceed
Function DisplayError(ByVal ErrorNumber As Integer, ByVal IsProxyConnecting As Boolean, ByVal Source As enuServerTypes) As Boolean
    
    Dim s As String
    
    s = GErrorHandler.GetErrorString(ErrorNumber, IsProxyConnecting, Source)
    
    If (LenB(s) > 0) Then
        If (IsProxyConnecting) Then
            s = "[PROXY] " & s
        End If
        
        Select Case (Source)
            Case stBNLS: s = "[BNLS] " & s
            Case stBNCS: s = "[BNCS] " & s
            Case stMCP: s = "[REALM] " & s
        End Select
        
        AddChat RTBColors.ErrorMessageText, s
    End If
    
    DisplayError = GErrorHandler.OKToProceed()
End Function

Sub LoadOutFilters()
    Const o As String = "Outgoing"
    Const f As String = FILE_FILTERS
    
    Dim s   As String
    Dim i   As Integer
    
    ReDim gOutFilters(1 To 1)
    ReDim Catch(0)
    
    Catch(0) = vbNullString
    
    s = ReadINI(o, "Total", f)
    
    If (Not (StrictIsNumeric(s))) Then
        Exit Sub
    End If
    
    For i = 1 To Val(s)
        gOutFilters(i).ofFind = Replace(LCase(ReadINI(o, "Find" & i, f)), "¦", " ")
        gOutFilters(i).ofReplace = Replace(ReadINI(o, "Replace" & i, f), "¦", " ")
        
        If (i <> Val(s)) Then
            ReDim Preserve gOutFilters(1 To i + 1)
        End If
    Next i
    
    If (Dir$(GetFilePath(FILE_CATCH_PHRASES)) <> vbNullString) Then
        i = FreeFile
        
        Open GetFilePath(FILE_CATCH_PHRASES) For Input As #i
        
            If (LOF(i) < 2) Then
                Close #i
                
                Exit Sub
            End If
            
            Do While Not EOF(i)
                Line Input #i, s
                
                If ((s <> vbNullString) And (s <> " ")) Then
                    Catch(UBound(Catch)) = LCase$(s)
                    
                    ReDim Preserve Catch(0 To UBound(Catch) + 1)
                End If
            Loop
            
            'Note: Why did this happen?
            'If Catch(0) = vbNullString Then Catch(0) = "¯"
            
        Close #i
    End If
End Sub

Function OutFilterMsg(ByVal strOut As String) As String
    Dim i As Integer
    
    If (UBound(gOutFilters) > 0) Then
        For i = LBound(gOutFilters) To UBound(gOutFilters)
            strOut = Replace(strOut, gOutFilters(i).ofFind, _
                gOutFilters(i).ofReplace)
        Next i
    End If
    
    OutFilterMsg = strOut
End Function

Private Sub sckBNet_DataArrival(ByVal bytesTotal As Long)
    #If (COMPILE_DEBUG <> 1) Then
        On Error GoTo ERROR_HANDLER
    #End If
    
    Dim buf() As Byte
    Dim pBuff As clsDataBuffer
    
    If bytesTotal = 0 Then Exit Sub
    
    ' read buffer as Byte()
    sckBNet.GetData buf(), vbArray + vbByte, bytesTotal
    ' add data to buffer
    ReceiveBuffer(stBNCS).InsertByteArr buf()
    
    If ProxyConnInfo(stBNCS).IsUsingProxy And ProxyConnInfo(stBNCS).Status <> psOnline Then
        Call modProxySupport.ProxyRecvPacket(sckBNet, ProxyConnInfo(stBNCS), ReceiveBuffer(stBNCS))
    Else
        Do While ReceiveBuffer(stBNCS).IsFullPacket(stBNCS)
            ' retrieve BNLS packet
            Set pBuff = ReceiveBuffer(stBNCS).TakePacket(stBNCS)
            ' parse
            Call modBNCS.BNCSRecvPacket(pBuff)
            ' clean up
            Set pBuff = Nothing
        Loop
    End If
    
    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in sckBNet_DataArrival()."

    Exit Sub
End Sub

Sub LoadArray(ByVal Mode As Byte, ByRef tArray() As String)
    Dim iFileHandle     As Integer      ' Source file number
    Dim sFilePath       As String       ' Full path to the source file
    Dim sTemp           As String       ' Temporary string
    Dim iFilterCount    As Integer      ' Total number of filters saved
    Dim sFilterSection  As String       ' The section of filters to read (message vs users)
    Dim i               As Integer      ' Counter

    ' Determine where to read from.
    Select Case Mode
        Case LOAD_FILTERS, LOAD_BLOCKLIST
            sFilePath = GetFilePath(FILE_FILTERS)
            sFilterSection = IIf(Mode = LOAD_FILTERS, "TextFilters", "BlockList")
        Case LOAD_PHRASES
            sFilePath = GetFilePath(FILE_PHRASE_BANS)
    End Select
    
    If Dir(sFilePath) <> vbNullString Then
        ' Empty the turn array.
        ReDim tArray(0)
        
        Select Case Mode
            Case LOAD_FILTERS, LOAD_BLOCKLIST
                ' Get the total number of filters
                sTemp = ReadINI(sFilterSection, "Total", FILE_FILTERS)
                If ((LenB(sTemp) > 0) And (CInt(sTemp) > 0)) Then
                    iFilterCount = Int(sTemp)
                    
                    ' Read each filter into a row of the array.
                    For i = 1 To iFilterCount
                        sTemp = ReadINI(sFilterSection, "Filter" & i, FILE_FILTERS)
                        If LenB(sTemp) > 0 Then
                            tArray(UBound(tArray)) = sTemp
                            ReDim Preserve tArray(UBound(tArray) + 1)
                        End If
                    Next i
                End If
            Case Else
                ' Read each line of the file into a row in the array.
                Open sFilePath For Input As #iFileHandle
                Do
                    Line Input #iFileHandle, sTemp
                    If LenB(sTemp) > 0 Then
                        tArray(UBound(tArray)) = sTemp
                        ReDim Preserve tArray(UBound(tArray) + 1)
                    End If
                Loop While Not EOF(iFileHandle)
                Close #iFileHandle
        End Select
        
        ' Remove the last unused row.
        If UBound(tArray) > 0 Then
            ReDim Preserve tArray(UBound(tArray) - 1)
        End If
    End If
End Sub

Private Sub sckBNLS_Close()
    If MDebug("all") Then
        AddChat COLOR_BLUE, "BNLS CLOSE"
    End If
    If (Not BNLSAuthorized) Then
        AddChat RTBColors.ErrorMessageText, "You have been disconnected from the BNLS server. You may be IPbanned from the server, it may be having issues, or there is something blocking your connection."
        AddChat RTBColors.ErrorMessageText, "Try using another BNLS server to connect, and check your firewall settings."
    End If
End Sub

Private Sub sckBNLS_Connect()
    If MDebug("all") Then
        AddChat COLOR_BLUE, "BNLS CONNECT"
    End If
    
    Call Event_BNLSConnected
    
    If (ProxyConnInfo(stBNLS).IsUsingProxy) Then
        modProxySupport.InitProxyConnection sckBNLS, ProxyConnInfo(stBNLS), BotVars.BNLSServer, 9367
    Else
        Call InitBNLSConnection
    End If
End Sub

Sub InitBNLSConnection()
    Call SetNagelStatus(sckBNLS.SocketHandle, False)
    
    modBNLS.SEND_BNLS_AUTHORIZE
End Sub


Private Sub sckBNLS_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ERROR_HANDLER
    
    Dim buf() As Byte
    Dim pBuff As clsDataBuffer
    
    If bytesTotal = 0 Then Exit Sub
    
    ' read buffer as Byte()
    sckBNLS.GetData buf(), vbArray + vbByte, bytesTotal
    ' add data to buffer
    ReceiveBuffer(stBNLS).InsertByteArr buf()

    If ProxyConnInfo(stBNLS).IsUsingProxy And ProxyConnInfo(stBNLS).Status <> psOnline Then
        Call modProxySupport.ProxyRecvPacket(sckBNLS, ProxyConnInfo(stBNLS), ReceiveBuffer(stBNLS))
    Else
        Do While ReceiveBuffer(stBNLS).IsFullPacket(stBNLS)
            ' retrieve BNLS packet
            Set pBuff = ReceiveBuffer(stBNLS).TakePacket(stBNLS)
            ' parse
            Call modBNLS.BNLSRecvPacket(pBuff)
            ' clean up
            Set pBuff = Nothing
        Loop
    End If
    
    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in sckBNLS_DataArrival()."

    Exit Sub
End Sub

Private Sub sckBNLS_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Event_BNLSError(Number, Description)
End Sub

'This function checks if the user selected when right-clicked is the same one when they click on the menu option. - FrOzeN
Private Function PopupMenuUserCheck() As Boolean
    If Not (lvChannel.SelectedItem Is Nothing) Then
        If mnuPopup.Tag <> lvChannel.SelectedItem.Text Then
            PopupMenuUserCheck = False
            Exit Function
        End If
    End If
    
    PopupMenuUserCheck = True
End Function

Private Function PopupMenuFLUserCheck() As Boolean
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        If mnuPopFList.Tag <> lvFriendList.SelectedItem.Text Then
            PopupMenuFLUserCheck = False
            Exit Function
        End If
    End If
    
    PopupMenuFLUserCheck = True
End Function

Private Function PopupMenuCLUserCheck() As Boolean
    If Not (lvClanList.SelectedItem Is Nothing) Then
        If mnuPopClanList.Tag <> lvClanList.SelectedItem.Text Then
            PopupMenuCLUserCheck = False
            Exit Function
        End If
    End If
    
    PopupMenuCLUserCheck = True
End Function

Function GetSelectedUsers() As Collection
    Dim i As Integer

    Set GetSelectedUsers = New Collection
    
    For i = 1 To lvChannel.ListItems.Count
        If (lvChannel.ListItems(i).Selected) Then
            Call GetSelectedUsers.Add(lvChannel.ListItems(i).Text)
        End If
    Next i
End Function

Function GetSelectedUser() As String
    If (lvChannel.SelectedItem Is Nothing) Then
        GetSelectedUser = vbNullString
    
        Exit Function
    End If
    
    GetSelectedUser = lvChannel.SelectedItem.Tag
End Function

Function GetFriendsSelectedUser() As String
    If (lvFriendList.SelectedItem Is Nothing) Then
        GetFriendsSelectedUser = vbNullString
    
        Exit Function
    End If
    
    GetFriendsSelectedUser = CleanUsername(ReverseConvertUsernameGateway(lvFriendList.SelectedItem.Text))
End Function

Function MatchClosest(ByVal toMatch As String, Optional startIndex As Long = 1) As String
    Dim lstView     As ListView

    Dim i           As Integer
    Dim CurrentName As String
    Dim atChar      As Integer
    Dim Index       As Integer
    Dim Loops       As Integer

    i = InStr(1, toMatch, " ", vbBinaryCompare)
    
    If (i > 0) Then
        toMatch = Mid$(toMatch, i + 1)
    End If
    
    Select Case (ListviewTabs.Tab)
        Case 0:
            Set lstView = lvChannel
        Case 1:
            Set lstView = lvFriendList
        Case 2:
            Set lstView = lvClanList
    End Select
    
    With lstView.ListItems
        If (.Count > 0) Then
            Dim c As Integer
            
            If (startIndex > .Count) Then
                Index = 1
            Else
                Index = startIndex
            End If
        
            While (Loops < 2)
                For i = Index To .Count 'for each user
                    CurrentName = .Item(i).Text
                
                    If (Len(CurrentName) >= Len(toMatch)) Then
                        For c = 1 To Len(toMatch) 'for each letter in their name
                            If (StrComp(Mid$(toMatch, c, 1), Mid$(CurrentName, c, 1), _
                                vbTextCompare) <> 0) Then
                                
                                Exit For
                            End If
                        Next c
                        
                        If (c >= (Len(toMatch) + 1)) Then
                            MatchClosest = .Item(i).Text
                            
                            MatchIndex = i
                            
                            Exit Function
                        End If
                    End If
                Next i
                
                Index = 1
                
                Loops = (Loops + 1)
            Wend
            
            Loops = 0
        End If
    End With
    
    atChar = InStr(1, toMatch, Config.GatewayDelimiter, vbBinaryCompare)
    
    If (atChar <> 0) Then
        Dim tmp      As String
        Dim Gateways(5, 2) As String
        Dim OtherGateway As String
        
        ' populate list
        Gateways(0, 0) = "USWest"
        Gateways(0, 1) = "Lordaeron"
        Gateways(1, 0) = "USEast"
        Gateways(1, 1) = "Azeroth"
        Gateways(2, 0) = "Asia"
        Gateways(2, 1) = "Kalimdor"
        Gateways(3, 0) = "Europe"
        Gateways(3, 1) = "Northrend"
        Gateways(4, 0) = "Beta"
        Gateways(4, 1) = "Westfall"
        
        Dim CurrentGateway As Integer
        CurrentGateway = -1
        If (LenB(BotVars.Gateway) > 0) Then
            For i = 0 To UBound(Gateways, 1)
                If (StrComp(BotVars.Gateway, Gateways(i, 0)) = 0) Then
                    CurrentGateway = i
                    OtherGateway = Gateways(i, 1)
                    Exit For
                End If
                If (StrComp(BotVars.Gateway, Gateways(i, 1)) = 0) Then
                    CurrentGateway = i
                    OtherGateway = Gateways(i, 0)
                    Exit For
                End If
            Next i
            If (CurrentGateway = -1) Then ' BotVars.Gateway not known, @[tab]=@BotVars.Gateway
                OtherGateway = BotVars.Gateway
                CurrentGateway = 0
            End If
        Else ' BotVars.Gateway is nothing, @[tab]
            MatchClosest = vbNullString
            
            MatchIndex = 1
            
            Exit Function
        End If
        
    
        If (startIndex > UBound(Gateways, 2)) Then
            Index = 0
        Else
            Index = (startIndex - 1)
        End If
        
        If (Len(toMatch) >= (atChar + 1)) Then
            tmp = Mid$(toMatch, atChar + 1)

            While (Loops < 2)
                If (Len(OtherGateway) >= Len(tmp)) Then
                    If (StrComp(Left$(OtherGateway, Len(tmp)), tmp, _
                        vbTextCompare) = 0) Then
                        
                        Dim j As Integer
                    
                        MatchClosest = Left$(toMatch, atChar) & Gateways(CurrentGateway, i)
                        
                        MatchIndex = (i + 1)
                        
                        Exit Function
                    End If
                End If
                
                Index = 0
                
                Loops = (Loops + 1)
            Wend
        Else
            If (tmp = vbNullString) Then
                MatchClosest = Left$(toMatch, atChar) & OtherGateway
                    
                MatchIndex = (Index + 1)
                    
                Exit Function
            End If
        End If
    End If
    
    MatchClosest = vbNullString
    
    MatchIndex = 1
End Function

'this is a fucking mess. It reads:
'"This copy of StealthBot has been tampered with. Please get a new copy of StealthBot at http://www.stealthbot.net.
'Additionally, please report the website at which you downloaded StealthBot in an e-mail to abuse@stealthbot.net. Thanks!"

Function GetHexProtectionMessage() As String
    GetHexProtectionMessage = _
Chr(Asc("T")) & Chr(Asc("h")) & Chr(Asc("i")) & Chr(Asc("s")) & Chr(Asc(" ")) & Chr(Asc("c")) & _
Chr(Asc("o")) & Chr(Asc("p")) & Chr(Asc("y")) & Chr(Asc(" ")) & Chr(Asc("o")) & Chr(Asc("f")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("h")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc(" ")) & Chr(Asc("b")) & Chr(Asc("e")) & Chr(Asc("e")) & Chr(Asc("n")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("a")) & Chr(Asc("m")) & Chr(Asc("p")) & Chr(Asc("e")) & Chr(Asc("r")) & Chr(Asc("e")) & Chr(Asc("d")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("P")) & Chr(Asc("l")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("g")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc(" ")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("w")) & _
 Chr(Asc(" ")) & Chr(Asc("c")) & Chr(Asc("o")) & Chr(Asc("p")) & Chr(Asc("y")) & Chr(Asc(" ")) & Chr(Asc("o")) & Chr(Asc("f")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("h")) & Chr(Asc("t")) & Chr(Asc("t")) & Chr(Asc("p")) & Chr(Asc(":")) & Chr(Asc("/")) & Chr(Asc("/")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc("w")) & Chr(Asc(".")) & Chr(Asc("s")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("b")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("A")) & Chr(Asc("d")) & Chr(Asc("d")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("i")) & Chr(Asc("o")) & Chr(Asc("n")) & _
    Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("l")) & Chr(Asc("y")) & Chr(Asc(",")) & Chr(Asc(" ")) & _
 Chr(Asc("p")) & Chr(Asc("l")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("r")) & Chr(Asc("e")) & Chr(Asc("p")) & Chr(Asc("o")) & Chr(Asc("r")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("e")) & Chr(Asc("b")) & Chr(Asc("s")) & Chr(Asc("i")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("w")) & Chr(Asc("h")) & Chr(Asc("i")) & Chr(Asc("c")) & Chr(Asc("h")) & Chr(Asc(" ")) & Chr(Asc("y")) & Chr(Asc("o")) & Chr(Asc("u")) & Chr(Asc(" ")) & Chr(Asc("d")) & Chr(Asc("o")) & Chr(Asc("w")) & Chr(Asc("n")) & Chr(Asc("l")) & Chr(Asc("o")) & Chr(Asc("a")) & Chr(Asc("d")) & Chr(Asc("e")) & Chr(Asc("d")) & Chr(Asc(" ")) & Chr(Asc("S")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("B")) & _
 Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(" ")) & Chr(Asc("i")) & Chr(Asc("n")) & Chr(Asc(" ")) & _
 Chr(Asc("a")) & Chr(Asc("n")) & Chr(Asc(" ")) & Chr(Asc("e")) & Chr(Asc("-")) & Chr(Asc("m")) & Chr(Asc("a")) & Chr(Asc("i")) & Chr(Asc("l")) & Chr(Asc(" ")) & Chr(Asc("t")) & Chr(Asc("o")) & Chr(Asc(" ")) & Chr(Asc("a")) & Chr(Asc("b")) & Chr(Asc("u")) & Chr(Asc("s")) & Chr(Asc("e")) & Chr(Asc("@")) & Chr(Asc("s")) & Chr(Asc("t")) & Chr(Asc("e")) & Chr(Asc("a")) & Chr(Asc("l")) & Chr(Asc("t")) & Chr(Asc("h")) & Chr(Asc("b")) & Chr(Asc("o")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc("n")) & Chr(Asc("e")) & Chr(Asc("t")) & Chr(Asc(".")) & Chr(Asc(" ")) & Chr(Asc("T")) & Chr(Asc("h")) & Chr(Asc("a")) & Chr(Asc("n")) & Chr(Asc("k")) & Chr(Asc("s")) & Chr(Asc("!"))
End Function

Sub DeconstructSettings()
    If Not (SettingsForm Is Nothing) Then
        Set SettingsForm = Nothing
    End If
End Sub

'SHOW/HIDE STUFF
Public Sub cmdShowHide_Click()
    rtbWhispersVisible = (StrComp(cmdShowHide.Caption, CAP_HIDE))
    rtbWhispers.Visible = rtbWhispersVisible
    Config.WhisperWindows = CBool(rtbWhispers.Visible)
    Call Config.Save
    
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        If rtbWhispersVisible Then
            Me.Height = Me.Height + rtbWhispers.Height - Screen.TwipsPerPixelY
        Else
            Me.Height = Me.Height - rtbWhispers.Height + Screen.TwipsPerPixelY
        End If
    End If
        
    Call Form_Resize
End Sub

'// to be called on every successful login
Public Sub InitListviewTabs()
    ListviewTabs.TabEnabled(LVW_BUTTON_FRIENDS) = Config.FriendsListTab
    ListviewTabs.TabEnabled(LVW_BUTTON_CLAN) = g_Clan.InClan
End Sub

'// to be called at disconnect time
Public Sub DisableListviewTabs()
    ListviewTabs.TabEnabled(LVW_BUTTON_FRIENDS) = False
    ListviewTabs.TabEnabled(LVW_BUTTON_CLAN) = False
End Sub

Public Function GetSmallIcon(ByVal sProduct As String, ByVal Flags As Long, IconCode As Integer) As Long
    Dim i As Long
    
    If (BotVars.ShowFlagsIcons = False) Then
        i = IconCode ' disable any of the below flags-based icons
    ElseIf (Flags And USER_BLIZZREP) = USER_BLIZZREP Then 'Flags = 1: blizzard rep
        i = ICBLIZZ
    ElseIf (Flags And USER_SYSOP) = USER_SYSOP Then 'Flags = 8: battle.net sysop
        i = ICSYSOP
    ElseIf (Flags And USER_CHANNELOP) = USER_CHANNELOP Then 'op
        i = ICGAVEL
    ElseIf (Flags And USER_GUEST) = USER_GUEST Then 'guest
        i = ICSPECS
    ElseIf (Flags And USER_SPEAKER) = USER_SPEAKER Then 'speaker
        i = ICSPEAKER
    ElseIf (Flags And USER_GFPLAYER) = USER_GFPLAYER Then 'GF player
        i = IC_GF_PLAYER
    ElseIf (Flags And USER_GFOFFICIAL) = USER_GFOFFICIAL Then 'GF official
        i = IC_GF_OFFICIAL
    ElseIf (Flags And USER_SQUELCHED) = USER_SQUELCHED Then 'squelched
        i = ICSQUELCH
    Else
        i = IconCode
    'Else
    '    Select Case (UCase$(sProduct))
    '        Case Is =  PRODUCT_STAR: I = ICSTAR
    '        Case Is = PRODUCT_SEXP: I = ICSEXP
    '        Case Is = PRODUCT_D2DV: I = ICD2DV
    '        Case Is = PRODUCT_D2XP: I = ICD2XP
    '        Case Is = PRODUCT_W2BN: I = ICW2BN
    '        Case Is = PRODUCT_CHAT: I = ICCHAT
    '        Case Is = PRODUCT_DRTL: I = ICDIABLO
    '        Case Is = PRODUCT_DSHR: I = ICDIABLOSW
    '        Case Is = PRODUCT_JSTR: I = ICJSTR
    '        Case Is = PRODUCT_SSHR: I = ICSCSW
    '        Case Is = PRODUCT_WAR3: I = ICWAR3
    '        Case Is = PRODUCT_W3XP: I = ICWAR3X
    '
    '        '*** Special icons for WCG added 6/24/07 ***
    '        Case Is = "WCRF": I = IC_WCRF
    '        Case Is = "WCPL": I = IC_WCPL
    '        Case Is = "WCGO": I = IC_WCGO
    '        Case Is = "WCSI": I = IC_WCSI
    '        Case Is = "WCBR": I = IC_WCBR
    '        Case Is = "WCPG": I = IC_WCPG
    '
    '        '*** Special icons for PGTour ***
    '        Case Is = "__A+": I = IC_PGT_A + 1
    '        Case Is = "___A": I = IC_PGT_A
    '        Case Is = "__A-": I = IC_PGT_A - 1
    '        Case Is = "__B+": I = IC_PGT_B + 1
    '        Case Is = "___B": I = IC_PGT_B
    '        Case Is = "__B-": I = IC_PGT_B - 1
    '        Case Is = "__C+": I = IC_PGT_C + 1
    '        Case Is = "___C": I = IC_PGT_C
    '        Case Is = "__C-": I = IC_PGT_C - 1
    '        Case Is = "__D+": I = IC_PGT_D + 1
    '        Case Is = "___D": I = IC_PGT_D
    '        Case Is = "__D-": I = IC_PGT_D - 1
    '
    '        Case Else: I = ICUNKNOWN
    '    End Select
    End If
    
    GetSmallIcon = i
End Function

Private Function GetLagIcon(ByVal Ping As Long, ByVal Flags As Long) As Long
    Select Case (Ping)
        Case 0
            GetLagIcon = 0
        Case 1 To 199
            GetLagIcon = LAG_1
        Case 200 To 299
            GetLagIcon = LAG_2
        Case 300 To 399
            GetLagIcon = LAG_3
        Case 400 To 499
            GetLagIcon = LAG_4
        Case 500 To 599
            GetLagIcon = LAG_5
        Case Is >= 600 Or -1
            GetLagIcon = LAG_6
        Case Else
            GetLagIcon = ICUNKNOWN
    End Select
    
    If ((Flags And USER_NOUDP) = USER_NOUDP) Then
        GetLagIcon = LAG_PLUG
    End If
End Function

Private Function GetNameColor(ByVal Flags As Long, ByVal IdleTime As Long, ByVal IsSelf As Boolean) As Long
    '/* Self */
    If (IsSelf) Then
        'Debug.Print "Assigned color IsSelf"
        GetNameColor = FormColors.ChannelListSelf
        
        Exit Function
    End If
    
    '/* Squelched */
    If ((Flags And USER_SQUELCHED&) = USER_SQUELCHED&) Then
        'Debug.Print "Assigned color SQUELCH"
        GetNameColor = FormColors.ChannelListSquelched
        
        Exit Function
    End If
    
    '/* Blizzard */
    If (((Flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or _
        ((Flags And USER_SYSOP&) = USER_SYSOP&)) Then
       
        GetNameColor = COLOR_BLUE
        
        Exit Function
    End If
    
    '/* Operator */
    If ((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Then
        'Debug.Print "Assigned color OP"
        GetNameColor = FormColors.ChannelListOps
        Exit Function
    End If
    
    '/* Idle */
    If (IdleTime > BotVars.SecondsToIdle) Then
        'Debug.Print "Assigned color IDLE"
        GetNameColor = FormColors.ChannelListIdle
        Exit Function
    End If
    
    '/* Default */
    'Debug.Print "Assigned color NORMAL"
    GetNameColor = FormColors.ChannelListText
End Function

Public Function GetFlagDescription(ByVal Flags As Long, ByVal ShowAll As Boolean) As String
    Dim sOut As String
    Dim sSep As String
    
    sOut = vbNullString
    sSep = vbNullString
    
    If (Flags And USER_SQUELCHED) = USER_SQUELCHED And ShowAll Then
        sOut = sOut & sSep & "squelched"
        sSep = ", "
    End If
    
    If (Flags And USER_CHANNELOP) = USER_CHANNELOP Then
        sOut = sOut & sSep & "channel operator"
        sSep = ", "
    End If
    
    If (Flags And USER_BLIZZREP) = USER_BLIZZREP Then
        sOut = sOut & sSep & "Blizzard representative"
        sSep = ", "
    End If
    
    If (Flags And USER_SYSOP) = USER_SYSOP Then
        sOut = sOut & sSep & "Battle.net system operator"
        sSep = ", "
    End If
    
    If (Flags And USER_NOUDP) = USER_NOUDP And ShowAll Then
        sOut = sOut & sSep & "UDP plug"
        sSep = ", "
    End If
    
    If (Flags And USER_BEEPENABLED) = USER_BEEPENABLED And ShowAll Then
        sOut = sOut & sSep & "beep enabled"
        sSep = ", "
    End If
    
    If (Flags And USER_GUEST) = USER_GUEST Then
        sOut = sOut & sSep & "guest"
        sSep = ", "
    End If
    
    If (Flags And USER_SPEAKER) = USER_SPEAKER Then
        sOut = sOut & sSep & "speaker"
        sSep = ", "
    End If
    
    If (Flags And USER_GFOFFICIAL) = USER_GFOFFICIAL Then
        sOut = sOut & sSep & "GF official"
        sSep = ", "
    End If
    
    If (Flags And USER_GFPLAYER) = USER_GFPLAYER Then
        sOut = sOut & sSep & "GF player"
        sSep = ", "
    End If
    
    If (LenB(sOut) = 0) And ShowAll Then
        If (Flags = &H0&) Then
            sOut = "normal"
        Else
            sOut = "unknown"
        End If
    End If
    
    GetFlagDescription = sOut
    
    If ShowAll Then
        GetFlagDescription = GetFlagDescription & " [0x" & Right$("00000000" & Hex(Flags), 8) & "]"
    End If
End Function

Public Sub AddName(ByVal Username As String, ByVal AccountName As String, ByVal Product As String, ByVal Flags As Long, ByVal Ping As Long, IconCode As Integer, Optional Clan As String, Optional ForcePosition As Integer)
On Error GoTo ERROR_HANDLER
    Dim i          As Integer
    Dim LagIcon    As Integer
    Dim isPriority As Integer
    Dim IsSelf     As Boolean
    
    If (StrComp(Username, GetCurrentUsername, vbTextCompare) = 0) Then
        MyFlags = Flags
        
        SharedScriptSupport.BotFlags = MyFlags
        
        IsSelf = True
    End If
    
    'If (GetChannelItemIndex(Username) > 0) Then
    '    Exit Sub
    'End If
    
    isPriority = (frmChat.lvChannel.ListItems.Count + 1)
    
    i = GetSmallIcon(Product, Flags, IconCode)
    
    'Special Cases
    'If i = ICSQUELCH Then
    '    'Debug.Print "Returned a SQUELCH icon"
    '    If ForcePosition > 0 Then isPriority = ForcePosition
    '
    If (((Flags And USER_BLIZZREP&) = USER_BLIZZREP&) Or _
            ((Flags And USER_CHANNELOP&) = USER_CHANNELOP&) Or _
            ((Flags And USER_SYSOP&) = USER_SYSOP&)) Then
        
        If (ForcePosition = 0) Then
            isPriority = 1
        Else
            isPriority = ForcePosition
        End If

    Else
        If (ForcePosition > 0) Then
            isPriority = ForcePosition
        End If
    End If
    
    If (i > frmChat.imlIcons.ListImages.Count) Then
        i = frmChat.imlIcons.ListImages.Count
    End If
        
    With frmChat.lvChannel
        .Enabled = False
        
        .ListItems.Add isPriority, , Username, , i
        
        ' store account name here so popup menus work
        .ListItems.Item(isPriority).Tag = AccountName
        
        .ListItems.Item(isPriority).ListSubItems.Add , , Clan
        
        If (.ColumnHeaders(3).Width = 0) Then
            LagIcon = 0
        Else
            LagIcon = GetLagIcon(Ping, Flags)
        End If
        .ListItems.Item(isPriority).ListSubItems.Add , , , LagIcon
        
        If (Not BotVars.NoColoring) Then
            .ListItems.Item(isPriority).ForeColor = GetNameColor(Flags, 0, IsSelf)
        End If
        
        .Enabled = True
        
        '.Refresh
    End With
    
    If IsSelf Then
        Call frmChat.UpdateListviewTabs
    End If

    Exit Sub
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.AddName", Err.Number, Err.Description)
End Sub


Public Function GetChannelItemIndex(ByVal NameToFind As String) As Integer
    Dim lvItem As ListItem

    Set lvItem = frmChat.lvChannel.FindItem(NameToFind)

    If (lvItem Is Nothing) Then
        If BotVars.UseD2Naming Then
            GetChannelItemIndex = 0
            Dim i As Integer
            For i = 1 To lvChannel.ListItems.Count
                If (StrComp(lvChannel.ListItems(i).Tag, CleanUsername(ReverseConvertUsernameGateway(NameToFind)), vbTextCompare) = 0) Then
                    GetChannelItemIndex = i
                    Exit For
                End If
            Next i
        Else
            GetChannelItemIndex = 0
        End If
    Else
        GetChannelItemIndex = lvItem.Index
    End If
End Function

Private Sub AddFriendItem(ByVal Name As String, ByVal Game As String, _
        ByVal Status As Byte, ByVal LocationID As Byte, ByVal EntryNumber As Integer)
On Error GoTo ERROR_HANDLER
    Dim ListItem As ListItem

    Set ListItem = lvFriendList.ListItems.Add(lvFriendList.ListItems.Count + 1, , Name)
    ListItem.ListSubItems.Add , , , ICUNKNOWN
    ListItem.ListSubItems.Add , , vbNullString
    SetFriendItem ListItem, EntryNumber, True, Game, Status, LocationID
    Set ListItem = Nothing

    Exit Sub
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.AddFriendItem", Err.Number, Err.Description)
End Sub

Private Sub SetFriendItem(ByVal ListItem As ListItem, ByVal EntryNumber As Integer, _
        Optional ByVal SettingFields As Boolean = False, Optional ByVal Game As String, _
        Optional ByVal Status As Byte, Optional ByVal LocationID As Byte)
On Error GoTo ERROR_HANDLER
    Dim OnlineIcon As Integer
    Dim GameIcon   As Integer

    If SettingFields Then
        Select Case Game
            Case PRODUCT_STAR: GameIcon = ICSTAR
            Case PRODUCT_SEXP: GameIcon = ICSEXP
            Case PRODUCT_D2DV: GameIcon = ICD2DV
            Case PRODUCT_D2XP: GameIcon = ICD2XP
            Case PRODUCT_W2BN: GameIcon = ICW2BN
            Case PRODUCT_WAR3: GameIcon = ICWAR3
            Case PRODUCT_W3XP: GameIcon = ICWAR3X
            Case PRODUCT_CHAT: GameIcon = ICCHAT
            Case PRODUCT_DRTL: GameIcon = ICDIABLO
            Case PRODUCT_DSHR: GameIcon = ICDIABLOSW
            Case PRODUCT_JSTR: GameIcon = ICJSTR
            Case PRODUCT_SSHR: GameIcon = ICSCSW
            Case Else:         GameIcon = ICUNKNOWN
        End Select

        If LocationID = FRL_PRIVATEGAME_MUTUAL Then LocationID = FRL_PRIVATEGAME
        If LocationID > FRL_PRIVATEGAME_MUTUAL Or LocationID < FRL_OFFLINE Then LocationID = FRL_NOTINCHAT
        If (Status And FRS_MUTUAL) = FRS_MUTUAL Then
            OnlineIcon = IC_FRIEND_MUTUAL_START + LocationID
        Else
            OnlineIcon = IC_FRIEND_START + LocationID
        End If

        If (Not BotVars.NoColoring) Then
            If (Status And FRS_AWAY) = FRS_AWAY Then
                ListItem.ForeColor = FormColors.ChannelListOps
            ElseIf (Status And FRS_DND) = FRS_DND Then
                ListItem.ForeColor = FormColors.ChannelListSquelched
            ElseIf LocationID <> FRL_OFFLINE Then
                ListItem.ForeColor = FormColors.ChannelListText
            Else
                ListItem.ForeColor = FormColors.ChannelListIdle
            End If
        End If

        ListItem.SmallIcon = GameIcon
        ListItem.ListSubItems.Item(1).ReportIcon = OnlineIcon
    End If

    ListItem.Tag = CInt(EntryNumber)
    ListItem.ListSubItems.Item(2).Text = CStr(1000 + EntryNumber)

    Set ListItem = Nothing

    Exit Sub
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.SetFriendItem", Err.Number, Err.Description)
End Sub

Private Function GetFriendItem(ByVal EntryNumber As Integer) As ListItem
    Dim i As Integer

    Set GetFriendItem = Nothing

    For i = 1 To lvFriendList.ListItems.Count
        If (CInt(lvFriendList.ListItems.Item(i).Tag) = EntryNumber) Then
            Set GetFriendItem = lvFriendList.ListItems.Item(i)
            Exit Function
        End If
    Next i
End Function

Private Sub AddClanMember(ByVal Name As String, ByVal DisplayName As String, ByVal Rank As Integer, ByVal Status As Integer)
On Error GoTo ERROR_HANDLER:
    Dim ListItem As ListItem

    Set ListItem = lvClanList.ListItems.Add(lvClanList.ListItems.Count + 1, , DisplayName)
    If (Not BotVars.NoColoring) Then
        If (StrComp(BotVars.Username, Name, vbTextCompare) = 0) Then
            ListItem.ForeColor = FormColors.ChannelListSelf
        End If
    End If
    ListItem.ListSubItems.Add , , , IC_CLAN_UNKNOWN
    ListItem.ListSubItems.Add , , vbNullString
    ListItem.Tag = CStr(Name)
    SetClanMember ListItem, DisplayName, Rank, Status
    Set ListItem = Nothing

    Exit Sub
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.AddClanMember", Err.Number, Err.Description)
End Sub

Private Sub SetClanMember(ByVal ListItem As ListItem, ByVal DisplayName As String, ByVal Rank As Integer, ByVal Status As Integer)
On Error GoTo ERROR_HANDLER
    Dim RankIcon   As Integer
    Dim OnlineIcon As Integer
    RankIcon = Rank

    If (RankIcon = clrankRecruit) Then RankIcon = IC_CLAN_PEON ' peon probation rank -> normal peon icon
    If (RankIcon < clrankRecruit Or RankIcon > clrankChieftain) Then RankIcon = IC_CLAN_UNKNOWN '// handle bad ranks

    If Status <> 0 Then
        OnlineIcon = IC_CLAN_STATUS_ONLINE
    Else
        OnlineIcon = IC_CLAN_STATUS_OFFLINE
    End If

    ListItem.Text = DisplayName
    ListItem.SmallIcon = RankIcon
    ListItem.ListSubItems.Item(1).ReportIcon = OnlineIcon
    ListItem.ListSubItems.Item(2).Text = CStr(1000 * RankIcon + ListItem.Index)

    Exit Sub
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, StringFormat("Error: #{0}: {1} in frmChat.SetClanMember", Err.Number, Err.Description)
End Sub

Private Function GetClanSelectedUser() As String
    With lvClanList
        If Not (.SelectedItem Is Nothing) Then
            If .SelectedItem.Index < 1 Then
                GetClanSelectedUser = vbNullString: Exit Function
            Else
                GetClanSelectedUser = CleanUsername(ReverseConvertUsernameGateway(.SelectedItem.Text))
            End If
        End If
    End With
End Function

Private Sub lvClanList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim aInx As Integer
    Dim bIsOn As Boolean
    Dim MyRank As Long
    Dim TheirRank As Long
    Dim CanMoveUp As Boolean
    Dim CanMoveDown As Boolean
    Dim CanRemove As Boolean
    Dim CanMakeChief As Boolean
    Dim CanDisband As Boolean
    Dim IsSelf As Boolean
    Dim CanLeave As Boolean
    
    If (lvClanList.SelectedItem Is Nothing) Then
        Exit Sub
    End If

    If Button = vbRightButton Then
        If Not (lvClanList.SelectedItem Is Nothing) Then
            aInx = lvClanList.SelectedItem.Index
            
            If aInx > 0 Then
                MyRank = g_Clan.Self.Rank
                TheirRank = g_Clan.GetUser(GetClanSelectedUser).Rank
                IsSelf = StrComp(g_Clan.Self.Name, GetClanSelectedUser, vbBinaryCompare) = 0
                
                CanMoveUp = Not IsSelf And (MyRank > (TheirRank + 1)) And (TheirRank > 0)
                CanMoveDown = Not IsSelf And (MyRank > TheirRank) And (TheirRank > 1)
                CanRemove = Not IsSelf And (CanMoveUp Or CanMoveDown)
                CanMakeChief = Not IsSelf And (MyRank = 4) And (TheirRank > 1)
                CanLeave = IsSelf And (MyRank > 0 And MyRank < 4)
                CanDisband = IsSelf And (MyRank = 4)
                
                bIsOn = g_Clan.GetUser(GetClanSelectedUser).IsOnline
                
                mnuPopClanWhisper.Enabled = bIsOn
                
                mnuPopClanPromote.Enabled = CanMoveUp
                mnuPopClanDemote.Enabled = CanMoveDown
                mnuPopClanRemove.Enabled = CanRemove
                mnuPopClanLeave.Enabled = CanLeave
                mnuPopClanDisband.Enabled = CanDisband
                mnuPopClanMakeChief.Enabled = CanMakeChief
                
                mnuPopClanPromote.Visible = Not IsSelf
                mnuPopClanDemote.Visible = Not IsSelf
                mnuPopClanRemove.Visible = Not IsSelf
                mnuPopClanLeave.Visible = IsSelf And (MyRank <> 4)
                mnuPopClanDisband.Visible = IsSelf And (MyRank = 4)
                mnuPopClanMakeChief.Visible = Not IsSelf And (MyRank = 4)
            End If
        End If
        
        mnuPopClanList.Tag = lvClanList.SelectedItem.Text 'Record which user is selected at time of right-clicking. - FrOzeN
        
        PopupMenu mnuPopClanList
    End If
End Sub

Sub DoConnect()

    If ((sckBNLS.State <> sckClosed) Or (sckBNet.State <> sckClosed)) Then
        Call DoDisconnect
    End If
    
    ConnectionTickCount = 0@
    
    UserCancelledConnect = False
    
    'Reset the BNLS auto-locator list
    BNLSFinderGotList = False
    
    'If Not IsValidIPAddress(BotVars.Server) And BotVars.UseProxy Then
        'AddChat RTBColors.ErrorMessageText, "[PROXY] Proxied connections must use a direct server IP address, such as those listed below your desired gateway in the Connection Settings menu, to connect."
        'AddChat RTBColors.ErrorMessageText, "[PROXY] Please change servers and try connecting again."
    'Else
        Call Connect
    'End If
End Sub

Sub DoDisconnect(Optional ByVal DoNotShow As Boolean = False, Optional ByVal LeaveUCCAlone As Boolean = False)
    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    If (Not (UserCancelledConnect)) Then
        tmrAccountLock.Enabled = False
        tmrIdleTimer.Enabled = False

        SetTitle "Disconnected"
        
        frmChat.UpdateTrayTooltip
        
        Call CloseAllConnections(Not DoNotShow)
        
        Set g_Channel = New clsChannelObj
        Set g_Clan = New clsClanObj
        Set g_Friends = New Collection
        
        ReDim ServerRequests(0)
        
        BotVars.Gateway = vbNullString
        
        CurrentUsername = vbNullString
        
        ListviewTabs.Tab = 0
        
        Call g_Queue.Clear
        
        If Not LeaveUCCAlone Then
            UserCancelledConnect = True
        End If
        
        If (UserCancelledConnect) Then
            'AddChat vbRed, "DISC!"
        
            If ReconnectTimerID > 0 Then
                KillTimer 0, ReconnectTimerID
                ReconnectTimerID = 0
            End If
            
            If ExReconnectTimerID > 0 Then
                KillTimer 0, ExReconnectTimerID
                ExReconnectTimerID = 0
            End If
            
            If SCReloadTimerID > 0 Then
                KillTimer 0, SCReloadTimerID
                SCReloadTimerID = 0
            End If
        Else
            'ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
            '    AddressOf Reconnect_TimerProc)
            '
            'ExReconnectTimerID = SetTimer(0, ExReconnectTimerID, _
            '    BotVars.ReconnectDelay, AddressOf ExtendedReconnect_TimerProc)
        End If
        
        DisableListviewTabs
        
        ProxyConnInfo(stBNLS).Status = psNotConnected
        ProxyConnInfo(stBNCS).Status = psNotConnected
        ProxyConnInfo(stMCP).Status = psNotConnected
        
        lvClanList.ListItems.Clear
        
        ReceiveBuffer(stBNLS).Clear
        ReceiveBuffer(stBNCS).Clear
        ReceiveBuffer(stMCP).Clear
        
        g_Connected = False
        g_Online = False
        ds.EnteredChatFirstTime = False
        ds.ClientToken = 0
        ds.AccountEntry = False

        If frmAccountManager.Visible Then
            frmAccountManager.LeftAccountEntryMode
        End If
        
        Call ClearChannel
        lvClanList.ListItems.Clear
        lvFriendList.ListItems.Clear
        
        'tmrSilentChannel(0).Enabled = False
        
        Call g_Queue.Clear
    
        BNLSAuthorized = False
        ConnectionTickCount = 0@
        
        mnuProfile.Enabled = False
        mnuClanCreate.Visible = False
        mnuRealmSwitch.Visible = False
        
        BotVars.LastChannel = vbNullString
        PrepareHomeChannelMenu
        PrepareQuickChannelMenu
        
        Set BotVars.PublicChannels = Nothing
        PreparePublicChannelMenu
        
        If ((Me.WindowState = vbNormal) And _
            (DoNotShow = 0)) Then
            
            'This SetFocus() call causes an error if any script have InputBoxes open.
            'This is the best fix I could come up with. :( -Pyro
            On Error Resume Next
            Call cboSend.SetFocus
            On Error GoTo ERROR_HANDLER
        End If
        
        ' clean up realms
        If Not ds.MCPHandler Is Nothing Then
            If ds.MCPHandler.FormActive Then
                frmRealm.UnloadAfterBNCSClose
            End If
            
            Set ds.MCPHandler = Nothing
        End If
        
        ' unload clan invitation popup
        Unload frmClanInvite
        
        ' clean up email reg
        Unload frmEMailReg
        
        ' close any pending Inet
        Inet.Tag = SB_INET_UNSET
        Inet.Cancel
        
        ' reset BNLS finder
        BNLSFinderGotList = False
        BNLSFinderIndex = 0
        
        On Error Resume Next
        RunInAll "Event_LoggedOff"
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & " in DoDisconnect()."
    
    Exit Sub
End Sub

Public Sub ParseFriendsPacket(ByVal PacketID As Long, ByVal pBuff As clsDataBuffer)
    FriendListHandler.ParsePacket PacketID, pBuff
End Sub

Public Sub ParseClanPacket(ByVal PacketID As Long, ByVal pBuff As clsDataBuffer)
    ClanHandler.ParseClanPacket PacketID, pBuff
End Sub

Public Sub RecordWindowPosition(Optional Maximized As Boolean = False)
    'Don't record other position information if maximized, otherwise when they unmaximize it will be fullscreen width and height. - FrOzeN
    If Not Maximized Then
        Config.PositionLeft = Int(Me.Left / Screen.TwipsPerPixelX)
        Config.PositionTop = Int(Me.Top / Screen.TwipsPerPixelY)
        Config.PositionHeight = Int(Me.Height / Screen.TwipsPerPixelY)
        Config.PositionWidth = Int(Me.Width / Screen.TwipsPerPixelX)
    End If
    
    Config.IsMaximized = Maximized
    Call Config.Save
End Sub

Public Sub MakeLoggingDirectory()
    On Error Resume Next
    MkDir GetFolderPath("Logs")
End Sub

' Called from several points to keep accurate tabs on the user's prior selection
'  in the send combo
Public Sub RecordcboSendSelInfo()
    'Debug.Print "SelStart: " & cboSend.SelStart & ", SelLength: " & cboSend.SelLength
    cboSendSelLength = cboSend.SelLength
    cboSendSelStart = cboSend.SelStart
End Sub
