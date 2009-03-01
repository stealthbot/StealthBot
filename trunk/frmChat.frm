VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmChat 
   BackColor       =   &H00000000&
   Caption         =   ":: StealthBot &version :: Disconnected ::"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12585
   ForeColor       =   &H00000000&
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ChatQueueTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5280
      Top             =   5160
   End
   Begin VB.Timer cacheTimer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   5280
      Top             =   4680
   End
   Begin MSComctlLib.ImageList imlClan 
      Left            =   8280
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   37
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1F59
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":249A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2744
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4DF5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8280
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   109
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":6F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":9FB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":D11E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1034A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":12952
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":15E1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":18600
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1B4E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1B855
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1E938
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":21AE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":24B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":27BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2AB9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2E0E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":30A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":30E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":311B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":31538
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":318D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":31C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":32028
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":323ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":32541
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":345FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":376B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3791E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":37B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":37D69
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":37FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38228
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38473
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":386DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38900
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38D49
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":38F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":391CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39435
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39638
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39871
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":39F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3A1B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3A3F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3A647
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3A889
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3AAF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3AD50
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3AFBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3B208
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3B487
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3B6BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3B929
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3BB40
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3BDAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3BFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3C20D
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3C460
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3C699
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3C8EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3CB60
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3CDCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3D022
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3D247
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3D49A
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3D6BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3D927
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3DB67
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3DDAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3DFFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3E241
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3E4AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3E6DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3E923
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3EB5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3EDCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3F035
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3F27C
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3F4C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3F6F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3F917
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3FB81
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3FD86
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3FF81
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":401C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":403D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":40643
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":40954
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":40C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4126B
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4183B
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":41DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4247B
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":42B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4319F
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":43706
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":43C73
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":441E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4478C
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":44D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":452DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":482C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4DFFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":53CC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":598DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":5F36E
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":64F9F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrScript 
      Index           =   0
      Left            =   5760
      Top             =   5160
   End
   Begin VB.Timer tmrClanUpdate 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5760
      Top             =   4680
   End
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
   Begin VB.Timer tmrSilentChannel 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   30000
      Left            =   6240
      Top             =   5160
   End
   Begin VB.Timer tmrSilentChannel 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   6240
      Top             =   4680
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
      TabPicture(0)   =   "frmChat.frx":6ABBD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Friends  "
      TabPicture(1)   =   "frmChat.frx":6ABD9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Clan  "
      TabPicture(2)   =   "frmChat.frx":6ABF5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSScriptControlCtl.ScriptControl SCRestricted 
      Left            =   6960
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin VB.Timer quLower 
      Interval        =   5360
      Left            =   7200
      Top             =   5160
   End
   Begin MSScriptControlCtl.ScriptControl SControl 
      Left            =   5520
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer tmrFriendlistUpdate 
      Interval        =   10000
      Left            =   7680
      Top             =   4680
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
      Left            =   6600
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer scTimer 
      Enabled         =   0   'False
      Left            =   7680
      Top             =   5160
   End
   Begin InetCtlsObjects.Inet INet 
      Left            =   6240
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBNLS 
      Left            =   6000
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   9367
   End
   Begin MSWinsockLib.Winsock sckBNet 
      Left            =   5520
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6112
   End
   Begin VB.Timer QueueTimer 
      Interval        =   1175
      Left            =   6720
      Top             =   5160
   End
   Begin VB.Timer UpTimer 
      Left            =   6720
      Top             =   4680
   End
   Begin VB.Timer Timer 
      Left            =   7200
      Top             =   4680
   End
   Begin MSComctlLib.ListView lvClanList 
      Height          =   6375
      Left            =   8880
      TabIndex        =   5
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
      NumItems        =   4
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   88
      EndProperty
   End
   Begin MSComctlLib.ListView lvFriendList 
      Height          =   6375
      Left            =   8880
      TabIndex        =   2
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Object.Width           =   88
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
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":6AC11
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
      TextRTF         =   $"frmChat.frx":6AC8C
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
      Begin VB.Menu mnuUsers 
         Caption         =   "&Database Manager"
      End
      Begin VB.Menu mnuCommandManager 
         Caption         =   "Command Manager"
      End
      Begin VB.Menu mnuSepTabcd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "User &Monitor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepXY 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGetNews 
         Caption         =   "Get &News and Check for Updates"
      End
      Begin VB.Menu mnuUpdateVerbytes 
         Caption         =   "Update &version bytes from StealthBot.net"
      End
      Begin VB.Menu mnuSepZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIgnoreInvites 
         Caption         =   "&Ignore Clan Invitations"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQCTop 
         Caption         =   "&QuickChannels"
         Begin VB.Menu mnuQCEdit 
            Caption         =   "&Edit QuickChannels"
         End
         Begin VB.Menu mnuBlank 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPublicChannels 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDash 
            Caption         =   "-"
            Visible         =   0   'False
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
         Caption         =   "&Bot Settings"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuUTF8 
         Caption         =   "Use &UTF-8 in Chat"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Edit &Profile"
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "&Edit Chat Filters"
      End
      Begin VB.Menu mnuCatchPhrases 
         Caption         =   "Edit &Catch Phrases"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCaught 
         Caption         =   "View Caught P&hrases"
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Edit &Files..."
         Begin VB.Menu mnuOpenBotFolder 
            Caption         =   "Open Bot &Folder"
         End
         Begin VB.Menu mnuOpenPluginsFolder 
            Caption         =   "Open &Plugins Folder"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSepA 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditConfig 
            Caption         =   "Config File"
         End
         Begin VB.Menu mnuEditChangelog 
            Caption         =   "Changelog"
         End
         Begin VB.Menu mnuEditQuotes 
            Caption         =   "Quotes List"
         End
         Begin VB.Menu mnuEditFilters 
            Caption         =   "Chat Filters"
         End
         Begin VB.Menu mnuEditDefinitions 
            Caption         =   "Definitions"
         End
         Begin VB.Menu mnuClearedTxt 
            Caption         =   "Current Text Log"
         End
         Begin VB.Menu mnuWhisperCleared 
            Caption         =   "Whisper Window Text Log"
         End
         Begin VB.Menu mnuEditUsers 
            Caption         =   "Userlist"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuEditMonitor 
            Caption         =   "Monitor"
         End
         Begin VB.Menu mnuEditPhrasebans 
            Caption         =   "Phrasebans"
         End
         Begin VB.Menu mnuEditScript 
            Caption         =   "Script File"
         End
         Begin VB.Menu mnuEditAccessFlags 
            Caption         =   "Access/Flags"
         End
         Begin VB.Menu mnuEditQCini 
            Caption         =   "QuickChannels"
         End
      End
      Begin VB.Menu mnuSettingsRepair 
         Caption         =   "&Tools..."
         Begin VB.Menu mnuToolsMenuWarning 
            Caption         =   "Use carefully!"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSep0607 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRepairDataFiles 
            Caption         =   "Delete &data files"
         End
         Begin VB.Menu mnuRepairVerbytes 
            Caption         =   "Restore default &version bytes"
         End
         Begin VB.Menu mnuRepairCleanMail 
            Caption         =   "Clean up &mail database"
         End
         Begin VB.Menu mnuPacketLog 
            Caption         =   "Log StealthBot &packet traffic"
         End
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnureload 
         Caption         =   "&Reload Config"
      End
      Begin VB.Menu mnuReloadScript 
         Caption         =   "Reload &Script"
         Shortcut        =   ^R
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
         Caption         =   "&Toggle Join/Leave Messages"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuHideBans 
         Caption         =   "Hide& Ban Messages"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuLock 
         Caption         =   "To&ggle Chat Window Lock"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuToggleFilters 
         Caption         =   "Toggle Chat &Filters"
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
      Begin VB.Menu mnuSP5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogging 
         Caption         =   "&Logging Settings"
         Begin VB.Menu mnuLog0 
            Caption         =   "Full Text Logging"
         End
         Begin VB.Menu mnuLog1 
            Caption         =   "Partial Text Logging"
         End
         Begin VB.Menu mnuLog2 
            Caption         =   "No Logging"
         End
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
         Caption         =   "&Flash Window on Events"
      End
      Begin VB.Menu mnuDisableVoidView 
         Caption         =   "Disable Silent Channel View"
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
      Begin VB.Menu mnuPopCopy 
         Caption         =   "Copy Name To Clipboard"
      End
      Begin VB.Menu mnuPopAddLeft 
         Caption         =   "Add to Left Send Box"
      End
      Begin VB.Menu mnuPopWhisper 
         Caption         =   "Whisper"
      End
      Begin VB.Menu mnuPopAddToFList 
         Caption         =   "Add to &Friends List"
      End
      Begin VB.Menu mnuPopSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopWhois 
         Caption         =   "Battle.Net Whois"
      End
      Begin VB.Menu mnuUserlistWhois 
         Caption         =   "Userlist Whois"
      End
      Begin VB.Menu mnuPopSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopKick 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuPopBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuPopShitlist 
         Caption         =   "Shitlist"
      End
      Begin VB.Menu mnuPopSquelch 
         Caption         =   "Squelch"
      End
      Begin VB.Menu mnuPopUnsquelch 
         Caption         =   "Unsquelch"
      End
      Begin VB.Menu mnuPopDes 
         Caption         =   "Designate"
      End
      Begin VB.Menu mnuPopSafelist 
         Caption         =   "Safelist"
      End
      Begin VB.Menu mnuSep0F 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopInvite 
         Caption         =   "Invite to Warcraft III Clan"
      End
      Begin VB.Menu mnuPopStats 
         Caption         =   "Stats"
         Begin VB.Menu mnuStatsW2 
            Caption         =   "Warcraft II BNE"
         End
         Begin VB.Menu mnuStatsSC 
            Caption         =   "Starcraft"
         End
         Begin VB.Menu mnuStatsBW 
            Caption         =   "Brood War"
         End
         Begin VB.Menu mnuStatsW3 
            Caption         =   "Warcraft III"
         End
         Begin VB.Menu mnuStatsFT 
            Caption         =   "The Frozen Throne"
         End
      End
      Begin VB.Menu mnuPopWebProfile 
         Caption         =   "Web Profile Lookup"
         Begin VB.Menu mnuPopWebProfileWAR3 
            Caption         =   "Reign of Chaos"
         End
         Begin VB.Menu mnuPopWebProfileW3XP 
            Caption         =   "The Frozen Throne"
         End
      End
      Begin VB.Menu mnuPopPLookup 
         Caption         =   "Profile Lookup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&Readme and Command List"
      End
      Begin VB.Menu mnuHelpWebsite 
         Caption         =   "&Homepage"
      End
      Begin VB.Menu mnuTerms 
         Caption         =   "&End-User License Agreement"
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
      Begin VB.Menu mnuPopClanWhois 
         Caption         =   "Whois"
      End
      Begin VB.Menu mnuPopBNProfile 
         Caption         =   "Battle.net Profile"
      End
      Begin VB.Menu mnuPopBNWebProfile 
         Caption         =   "Web Profile"
         Begin VB.Menu mnuPopBNWebProfileWAR3 
            Caption         =   "Reign of Chaos"
         End
         Begin VB.Menu mnuPopBNWebProfileW3XP 
            Caption         =   "The Frozen Throne"
         End
      End
      Begin VB.Menu mnuSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopPro 
         Caption         =   "Promote"
      End
      Begin VB.Menu mnuPopDem 
         Caption         =   "Demote"
      End
      Begin VB.Menu mnuSP2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPopRem 
         Caption         =   "Remove from Clan"
      End
      Begin VB.Menu mnuPopLeaveClan 
         Caption         =   "Leave Clan"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuFLpop 
      Caption         =   "flistpopup"
      Visible         =   0   'False
      Begin VB.Menu mnuFLPopWhisper 
         Caption         =   "&Whisper"
      End
      Begin VB.Menu mnuPopFLWhois 
         Caption         =   "W&hois"
      End
      Begin VB.Menu mnuPopFLProfile 
         Caption         =   "&Profile"
      End
      Begin VB.Menu mnuPSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLpopPromote 
         Caption         =   "&Promote"
      End
      Begin VB.Menu mnuFLpopDemote 
         Caption         =   "&Demote"
      End
      Begin VB.Menu mnuFLPopRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuPSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFListRefresh 
         Caption         =   "R&efresh and Reorder"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StealthBot 11/5/02-Present
'Source Code Version: 2.6R3+
Option Explicit

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

Public Sub cacheTimer_Timer()
    ' this code updated 7/23/05 in Chihuahua, Chihuahua, MX
    If (Caching) Then ' time to retrieve stored information and squelch or ban a channel
        Dim strArray() As String
        Dim ret As String
        Dim lPos As Long
        Dim Y As String
        Dim C As Integer, n As Integer
        
        Caching = False
        
        ' ...
        ret = cache(vbNullString, 0, Y)
        
        ' ...
        lPos = InStr(1, ret, ", ", vbBinaryCompare)
        
        ' ...
        If (lPos) Then
            ' ...
            strArray() = Split(ret, ", ")
        Else
            ' ...
            ReDim Preserve strArray(0)
            
            ' ...
            strArray(0) = ret
        End If
        
        For C = 0 To UBound(strArray)
            ' [CHANNELOP]  -  [*CHANNELOP]  -  [CHARACTER@USEast (*CHANNELOP)]
            If StrComp(UCase(strArray(C)), strArray(C), vbTextCompare) = 0 Then
                If Left$(strArray(C), 1) = "[" And Right$(strArray(C), 1) = "]" Then
                    strArray(C) = Mid(strArray(C), 2, Len(strArray(C)) - 2)
                End If
            End If
        
            'n = InStr(strArray(C), "(*")
            
            'If n > 0 Then
            '    ' This covers Character@USeast (*Username)
            '
            '    strArray(C) = Mid$(strArray(C), n + 2)
            '    strArray(C) = Left$(strArray(C), Len(strArray(C)) - 1)
            'Else
            '    n = InStr(strArray(C), "*")
            '
            '    ' This covers *Username
            '
            '    If n > 0 Then
            '        strArray(C) = Mid$(strArray(C), n + 1)
            '    End If
            'End If
            
            strArray(C) = convertUsername(CleanUsername(strArray(C)))
            
            'AddChat vbRed, strArray(C)
            
            If Len(strArray(C)) > 1 Then
                If InStr(Y, "ban") Then
                    If (g_Channel.Self.IsOperator) Then
                        Ban strArray(C), (AutoModSafelistValue - 1), 0
                    End If
                Else
                    If (GetSafelist(strArray(C)) = False) Then
                        AddQ "/squelch " & strArray(C)
                    End If
                End If
            End If
        Next C
    End If
    
    cacheTimer.Enabled = False
End Sub

Private Sub ChatQueueTimer_Timer()
    modChatQueue.ChatQueueTimerProc
End Sub

' LET IT BEGIN
Private Sub Form_Load()
    Dim s As String
    Dim f As Integer ', i As Integer
    Dim L As Long
    Dim FrmSplashInUse As Boolean
    
    ' COMPILER FLAGS
    #If (BETA = 1) Then
        #If (DEV_RELEASE = 1) Then
            CVERSION = "StealthBot Beta v" & App.Major & "." & App.Minor & _
                ZeroOffsetEx(App.REVISION, 3) & " Development Release " & "34"
        #Else
            CVERSION = "StealthBot Beta v" & App.Major & "." & App.Minor & _
                ZeroOffsetEx(App.REVISION, 3) & " Stable"
        #End If
    #Else
        CVERSION = "StealthBot v" & App.Major & "." & App.Minor & " build " & _
            ZeroOffsetEx(App.REVISION, 3) & IIf(Len(REVISION) > 0, " Revision " & REVISION, "")
    #End If
    
    #If (COMPILE_RELEASE = 1) Then
        If (Not ValidateExecutable) Then
            MsgBox GetHexProtectionMessage, vbOKOnly + vbCritical
            'Call Form_Unload(0)
            Unload frmChat
            Exit Sub
        End If
    #End If
    
    #If (COMPILE_DEBUG = 0) Then
        HookWindowProc frmChat.hWnd
        'HookSendBoxWindowProc frmChat.cboSend.hWnd
    #End If
    
    ' 4/10/06:
    ' CHECK FOR CONFIG.INI PATH HACK
    L = InStr(command(), "-cpath ")
    
    If L > 0 And Len(command()) > (L + 7) Then
        ConfigOverride = Mid$(command(), L + 7)
        
        If InStr(ConfigOverride, " ") > 0 Then
            ConfigOverride = Split(ConfigOverride, " ")(0)
        End If
        
        If LenB(GetConfigFilePath()) = 0 Then
            ConfigOverride = ""
        End If
    End If
        
    ' SPLASH SCREEN
    If ReadCfg("Main", "ShowSplash") <> "N" Then
        frmSplash.Show
        FrmSplashInUse = True
    End If
    
    ' EVERYTHING ELSE
    rtbWhispers.Visible = False 'default
    rtbWhispersVisible = False
    
    Set colWhisperWindows = New Collection
    Set colLastSeen = New Collection
    Set GErrorHandler = New clsErrorHandler
    Set BotVars = New clsBotVars
    Set colQueue = New Collection
    Set colSafelist = New Collection
    Set dctCallbacks = New Dictionary
        dctCallbacks.CompareMode = BinaryCompare
        
    Set colDynamicMenus = New Collection
    
    Set dictTimerInterval = New Dictionary
    Set dictTimerEnabled = New Dictionary
    Set dictTimerCount = New Dictionary
    dictTimerInterval.CompareMode = TextCompare
    dictTimerEnabled.CompareMode = TextCompare
    dictTimerCount.CompareMode = TextCompare
    
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
    lvChannel.icons = imlIcons
    lvClanList.View = lvwReport
    lvClanList.icons = imlIcons
    
    ReDim Phrases(0)
    ReDim ClientBans(0)
    ReDim Catch(0)
    ReDim gBans(0)
    ReDim gOutFilters(0)
    ReDim gFilters(0)
    
    s = ReadCfg("Main", "ShowWhisperWindow")
    If s = "Y" Then
        If Not rtbWhispersVisible Then Call cmdShowHide_Click
    Else
        If rtbWhispersVisible Then Call cmdShowHide_Click
    End If
    
    CfgVersion = Val(ReadCfg("Main", "ConfigVersion"))
'
'    If CfgVersion < 3 Then
'
'        s = ReadCFG("Position", "Left")
'        If LenB(s) > 0 And StrictIsNumeric(s) Then Me.Left = IIf(CLng(s) < 100, 100, CLng(s))
'
'        s = ReadCFG("Position", "Top")
'        If LenB(s) > 0 And StrictIsNumeric(s) Then Me.Top = IIf(CLng(s) < 100, 100, CLng(s))
'
'    Else
'
'        s = ReadCFG("Position", "Left")
'        If LenB(s) > 0 And StrictIsNumeric(s) Then Me.Left = (IIf(CLng(s) < 100, 100, CLng(s)) * Screen.TwipsPerPixelX)
'
'        s = ReadCFG("Position", "Top")
'        If LenB(s) > 0 And StrictIsNumeric(s) Then Me.Top = (IIf(CLng(s) < 100, 100, CLng(s)) * Screen.TwipsPerPixelY)
'
'    End If

    s = ReadCfg("Position", "Height")
    If LenB(s) > 0 And StrictIsNumeric(s) Then
        L = (IIf(CLng(s) < 200, 200, CLng(s)) * Screen.TwipsPerPixelY)
        
        If (rtbWhispersVisible) Then
            L = L - (rtbWhispers.Height / Screen.TwipsPerPixelY)
        End If
        
        Me.Height = L
    End If
    
    s = ReadCfg("Position", "Width")
    If LenB(s) > 0 And StrictIsNumeric(s) Then
        Me.Width = (IIf(CLng(s) < 300, 300, CLng(s)) * Screen.TwipsPerPixelX)
    End If

    ' 2.7 must discard old position data
    If CfgVersion >= 4 Then
        s = ReadCfg("Position", "Left")
        If LenB(s) > 0 And StrictIsNumeric(s) Then
            Me.Left = CLng(s) * Screen.TwipsPerPixelX
        End If
        
        s = ReadCfg("Position", "Top")
        If LenB(s) > 0 And StrictIsNumeric(s) Then
            Me.Top = CLng(s) * Screen.TwipsPerPixelY
        End If
    Else
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
    End If
    
    'Support for recording maxmized position. - FrOzeN
    s = ReadCfg("Position", "Maximized")
    
    If s = "True" Then
        Me.WindowState = vbMaximized
    End If
    
    Set ClanHandler = New clsClanPacketHandler
    Set FriendListHandler = New clsFriendlistHandler
    Set ListToolTip = New CTooltip
    
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
    
    With lvClanList
        .View = lvwReport
        .SmallIcons = imlClan
        .ColumnHeaders(1).Width = (.Width \ 4) * 3 - 150
        .ColumnHeaders(2).Width = .Width \ 4 + 200
        .ColumnHeaders(3).Width = 0
    End With
    
    frmChat.KeyPreview = True
    SetTitle "Disconnected"
    
    Me.Show
    Me.Refresh
    Me.AutoRedraw = True
    
    AddChat RTBColors.ConsoleText, "-> Welcome to " & CVERSION & ", by Stealth."
    AddChat RTBColors.ConsoleText, "-> If you enjoy StealthBot, consider supporting its development at http://support.stealthbot.net"

    On Error Resume Next

    'If BotVars.Logging < 2 Then
    '    MakeLoggingDirectory
    '
    '    If (dir$(GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") = vbNullString) Then
    '        Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt" For Output As #1
    '        Close #1
    '    End If
    '
    '    If (dir$(GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt") = vbNullString) Then
    '        Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt" For Output As #1
    '        Close #1
    '    End If
    'End If
    
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
    
    CommandLine = command()
    
    If MDebug("debug") Then _
        AddChat RTBColors.ServerInfoText, " * Program executed in debug mode; unhandled packet " & _
            "information will be displayed."
    
    Randomize
    
    ID_TASKBARICON = (Rnd * 100)
    
    TASKBARCREATED_MSGID = RegisterWindowMessage("TaskbarCreated")
    
    'BNCSUtil Version Check
'    If Not bncsutil_checkVersion(BNCSUTIL_VERSION) Then
'        AddChat RTBColors.ErrorMessageText, "Warning: You are using an outdated copy of BNCSUtil.dll. Some features may not function correctly."
'    End If
    
    cboSend.SetFocus
    
    LoadQuickChannels
    LoadQuotes
    
    LoadPluginSystem SControl

    'Dim Logger As New clsLogger
   
    'Dim I As Integer
   '
   ' For I = 1 To 5
   '     g_Logger.WriteError vbNullString, "oops"
   ' Next I
    
    On Error Resume Next
    'News call and scripting events
    
    s = ReadCfg("Override", "DisableSBNews")
    
    If (LenB(s) = 0) Then DisplayNews
    
    If FrmSplashInUse Then frmSplash.SetFocus
    
    If Not MDebug("debug") Then
        mnuRecordWindowPos.Visible = False
    End If
    
    '#If BETA = 0 Then
        If ReadCfg("Main", "ConnectOnStartup") = "Y" Then
            Call DoConnect
        End If
    '#End If
    
    #If COMPILE_DEBUG = 0 Then
        If ReadCfg("Main", "MinimizeOnStartup") = "Y" Then
            frmChat.WindowState = vbMinimized
            Call Form_Resize
        End If
    #End If
    
    'Now loads scripts when the bot opens, instead of after connecting. - FrOzeN
    SControl.Run "Event_FirstRun"
    SControl.Run "Event_Load"
    
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

Private Sub Form_GotFocus()
    On Error GoTo ERROR_HANDLER

    If (cboSendHadFocus) Then
        cboSend.SetFocus
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in Form_GotFocus()."

    Exit Sub
End Sub

Private Sub DisplayNews()

    Dim ret As String ' ...
    
    ' ...
    ret = INet.OpenURL(GetNewsURL())
    
    ' ...
    HandleNews ret

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
    Dim arr() As Variant ' ...
    Dim I     As Integer ' ...

    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmChat.rtbChat, arr)
End Sub

Sub AddChatFont(ParamArray saElements() As Variant)
    Dim arr() As Variant ' ...
    Dim I     As Integer ' ...

    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmChat.rtbChat, arr)
End Sub


'RTB ADDCHATFONT SUBROUTINE - originally written by Grok[vL] - modified to support
'                         logging and timestamps and color decoding
' AddChatFont allows you to specify FONT, COLOR, MESSAGE and have the message
'   displayed in that font
' Created 2/19/2007 based on a suggestion from Imhotep[Nu]
'Sub AddChatFont(ParamArray saElements() As Variant)
'    On Error Resume Next
'
'    Dim s As String
'    Dim l As Long, lngVerticalPos As Long
'    Dim I As Integer, intRange As Integer, f As Integer
'    Dim blUnlock As Boolean, LogThis As Boolean
'
'    If Not BotVars.LockChat Then
'
'        If g_OSVersion.IsWin2000Plus() Then
'            GetScrollRange rtbChat.hWnd, SB_VERT, 0, intRange
'            lngVerticalPos = SendMessage(rtbChat.hWnd, EM_GETTHUMB, 0&, 0&)
'
'            'Debug.Print "ScrollRange: " & intRange & " ; VerticalPos: " & lngVerticalPos & " ; rtbChatHeight " & rtbChat.Height & " ; pix " & rtbChat.Height / Screen.TwipsPerPixelY
'
'            If (lngVerticalPos + (rtbChat.Height / Screen.TwipsPerPixelY)) <= intRange Then
'                'LockWindowUpdate rtbChat.hWnd
'                rtbChat.Visible = False
'                blUnlock = True
'            End If
'        End If
'
'        LogThis = (BotVars.Logging < 2)
'
'        If ((BotVars.MaxBacklogSize) And (rtbChatLength >= BotVars.MaxBacklogSize)) Then
'            With rtbChat
'                .Visible = False
'                .SelStart = 0
'                .SelLength = InStr(1, .text, vbLf, vbBinaryCompare)
'
'                rtbChatLength = rtbChatLength - .SelLength
'
'                If BotVars.Logging < 2 And LOF(i) < BotVars.MaxLogfileSize Then
'                    i = FreeFile
'                    Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #i
'                        Print #i, Left$(.SelText, Len(.SelText) - 2)
'                    Close #i
'                End If
'
'                .SelText = ""
'                .Visible = True
'            End With
'        End If
'
'        s = GetTimeStamp
'
'        With rtbChat
'            .SelStart = Len(.text)
'            .SelLength = 0
'            .SelColor = RTBColors.TimeStamps
'            If .SelBold = True Then .SelBold = False
'            If .SelItalic = True Then .SelItalic = False
'            If .SelUnderline = True Then .SelUnderline = False
'            .SelText = s
'            .SelStart = Len(.text)
'        End With
'
'        If LogThis Then
'            f = FreeFile
'
'            Open (GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") For Append As #f
'
'            If (LOF(f) >= BotVars.MaxLogFileSize) Then
'                LogThis = False
'                Close #f
'            End If
'        End If
'
'        For I = LBound(saElements) To UBound(saElements) Step 3
'
'            If InStr(1, saElements(I + 1), Chr(0), vbBinaryCompare) > 0 Then _
'                KillNull saElements(I + 1)
'
'            If InStr(1, saElements(I + 2), Chr(0), vbBinaryCompare) > 0 Then _
'                KillNull saElements(I + 2)
'
'            If Len(saElements(I + 2)) > 0 Then
'                l = InStr(1, saElements(I + 1), "{\rtf", vbTextCompare)
'
'                While (l > 0)
'                    Mid$(saElements(I + 2), l + 2, 1) = "/"
'
'                    l = InStr(1, saElements(I + 2), "{\rtf", vbTextCompare)
'                Wend
'
'                With rtbChat
'                    .SelStart = Len(.text)
'                    l = .SelStart
'                    .SelLength = 0
'                    .SelFontName = saElements(I)
'                    .SelColor = saElements(I + 1)
'                    .SelText = saElements(I + 2) & Left$(vbCrLf, -2 * CLng((I + 2) = UBound(saElements)))
'
'                    rtbChatLength = rtbChatLength + Len(s) + Len(saElements(I + 2)) + Len(Left$(vbCrLf, -2 * CLng((I + 2) = UBound(saElements))))
'
'                    .SelStart = Len(.text)
'                End With
'
'                ' Fixed 11/21/06 to properly log timestamps
'                If LogThis Then
'                    Print #f, s & saElements(I + 2) & Left$(vbCrLf, -2 * CLng((I + 2) = UBound(saElements)));
'                End If
'            End If
'
'        Next I
'
'        Call ColorModify(rtbChat, l)
'
'        If LogThis Then
'            Close #f
'            LogThis = False
'        End If
'
'        If blUnlock Then
'            SendMessage rtbChat.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * lngVerticalPos, 0&
'            'LockWindowUpdate 0&
'            rtbChat.Visible = True
'        End If
'
'    End If
'End Sub


Sub AddWhisper(ParamArray saElements() As Variant)
    
    Dim arr() As Variant ' ...
    
    ' ...
    arr() = saElements
    
    ' ...
    Call DisplayRichText(frmChat.rtbWhispers, arr)

    ' ...
    Exit Sub
    
    
    Dim s As String
    Dim L As Long
    Dim I As Integer
    
    If Not BotVars.LockChat Then
        'If ((BotVars.MaxBacklogSize) And (Len(rtbWhispers.text) >= BotVars.MaxBacklogSize)) Then
            If BotVars.Logging < 2 Then
                Close #1
                Open (GetProfilePath() & "\Logs\" & Format(Date, "YYYY-MM-DD") & "-WHISPERS.txt") For Append As #1
            End If
            
            With rtbWhispers
                .Visible = False
                .SelStart = 0
                .SelLength = InStr(1, .text, vbLf, vbBinaryCompare)
                If BotVars.Logging < 2 Then Print #1, Left$(vbCrLf, -2 * CLng((I + 1) = UBound(saElements)))
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
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = RTBColors.TimeStamps
            If .SelBold = True Then .SelBold = False
            If .SelItalic = True Then .SelItalic = False
            .SelText = s
            .SelStart = Len(.text)
        End With
        
        For I = LBound(saElements) To UBound(saElements) Step 2
            If InStr(1, saElements(I), Chr(0), vbBinaryCompare) > 0 Then _
                KillNull saElements(I)
            
            If Len(saElements(I + 1)) > 0 Then
                With rtbWhispers
                    .SelStart = Len(.text)
                    L = .SelStart
                    .SelLength = 0
                    .SelColor = saElements(I)
                    .SelText = saElements(I + 1) & Left$(vbCrLf, -2 * CLng((I + 1) = UBound(saElements)))
                    .SelStart = Len(.text)
                End With
            End If
        Next I
        
        Call ColorModify(rtbWhispers, L)
    End If
End Sub


'BNLS EVENTS
Sub Event_BNetConnected()
    If (BotVars.UseProxy) Then
        AddChat RTBColors.SuccessText, "[PROXY] Connected!"
    Else
        AddChat RTBColors.SuccessText, "[BNET] Connected!"
    End If
    
    Call SetNagelStatus(sckBNet.SocketHandle, False)
End Sub

Sub Event_BNetConnecting()
    If BotVars.UseProxy Then
        AddChat RTBColors.InformationText, "[PROXY] Connecting to the Battle.net server at " & BotVars.Server & "..."
    Else
        AddChat RTBColors.InformationText, "[BNET] Connecting to the Battle.net server at " & BotVars.Server & "..."
    End If
End Sub

Sub Event_BNetDisconnected()
    Timer.Interval = 0
    UpTimer.Interval = 0
    BotVars.JoinWatch = 0
    
    AddChat RTBColors.ErrorMessageText, IIf(BotVars.UseProxy And BotVars.ProxyStatus <> psOnline, "[PROXY] ", "[BNET] ") & "Disconnected."
    
    DoDisconnect (1)
    
    SetTitle "Disconnected"
    g_Online = False
    
    Call ClearChannel
    
    UpdateProxyStatus psNotConnected
    'AddChat RTBColors.ErrorMessageText, "[BNET] Attempting to reconnect, please wait..."
    'AddChat RTBColors.SuccessText, "Connection initialized."
    
    If sckBNet.State <> 0 Then sckBNet.Close
    If sckBNLS.State <> 0 Then sckBNLS.Close
    
    Passed0x0F = 0
    
    Call UpdateTrayTooltip
    
    'If Not UserCancelledConnect Then
    '    ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
    '        AddressOf Reconnect_TimerProc)
    'End If
End Sub

Sub Event_BNetError(ErrorNumber As Integer, description As String)
    Dim s As String
    
    If BotVars.UseProxy And BotVars.ProxyStatus <> psOnline Then
        s = "[PROXY] "
    Else
        s = "[BNET] "
    End If
    
    AddChat RTBColors.ErrorMessageText, s & ErrorNumber & " -- " & description
    AddChat RTBColors.ErrorMessageText, s & "Disconnected."
    
    If (sckBNet.State <> 0) Then
        Call sckBNet.Close
    End If
    
    If (sckBNLS.State <> 0) Then
        Call sckBNLS.Close
    End If
    
    If (sckMCP.State <> 0) Then
        Call sckMCP.Close
    End If
    
    g_Connected = False
    
    UserCancelledConnect = False
    
    DoDisconnect 1, True
    SetTitle "Disconnected"
    
    Call ClearChannel
    lvClanList.ListItems.Clear
    lvFriendList.ListItems.Clear
    
    lblCurrentChannel.Caption = GetChannelString
    
    ' NOV 18 04 Change here should fix the attention-grabbing on errors
    'If Me.WindowState <> vbMinimized Then cboSend.SetFocus
    
    
    If DisplayError(ErrorNumber, IIf(BotVars.UseProxy And BotVars.ProxyStatus <> psOnline, 2, 1), BNET) = True Then
        AddChat RTBColors.ErrorMessageText, _
            IIf(BotVars.UseProxy And BotVars.ProxyStatus <> psOnline, "[PROXY] ", "[BNET] ") & _
                "Attempting to reconnect in " & (BotVars.ReconnectDelay / 1000) & _
                    IIf(((BotVars.ReconnectDelay / 1000) > 1), " seconds", " second") & _
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
    AddChat RTBColors.SuccessText, "[BNLS] Connected!"
    
    Call SetNagelStatus(sckBNLS.SocketHandle, False)
End Sub

Sub Event_BNLSConnecting()
    AddChat RTBColors.InformationText, "[BNLS] Connecting to the BNLS server at " & BotVars.BNLSServer & "..."
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

Sub Event_BNLSError(ErrorNumber As Integer, description As String)
    If sckBNet.State <> 7 Then
    
        sckBNet.Close
        
        'Check the user has using BNLS server finder enabled
        If BotVars.UseAltBnls = True Then
            Call FindAltBNLS
        ElseIf BotVars.UseAltBnls = False Then
            AddChat RTBColors.ErrorMessageText, "[BNLS] Error " & ErrorNumber & ": " & description
            
            UserCancelledConnect = False
            
            DoDisconnect 1, True
            
            If (askedBnls = False) Then
                'Ask the user if they would like to enable the BNLS Automatic Server finder
                Dim msgResult As VbMsgBoxResult
                
                msgResult = MsgBox("BNLS Server Error." & vbCrLf & vbCrLf & _
                                   "Would you like to enable the BNLS Automatic Server Finder?", _
                                   vbYesNo, "BNLS Error")
                
                'Save their answer to the config, and the call this procedure again to reevaluate what to do
                WriteINI "Main", "UseAltBNLS", IIf(msgResult = vbYes, "Y", "N")
                
                If (msgResult = vbYes) Then
                    BotVars.UseAltBnls = True
                    
                    Call Event_BNLSError(ErrorNumber, description)
                Else
                    BotVars.UseAltBnls = False
                End If
                
                askedBnls = True
            End If
            
            If (BotVars.UseAltBnls = False) Then
                DisplayError ErrorNumber, 0, BNLS
            End If
            
            'If DisplayError(ErrorNumber, 0, BNLS) Then
            '    UserCancelledConnect = False
            '    Call DoDisconnect(1, True)
            'Else
            '    Call DoDisconnect
            '    SetTitle "Disconnected"
            'End If
            
            'If Not UserCancelledConnect Then
            '    AddChat vbRed, BotVars.ReconnectDelay
            '
            '
            '    ReconnectTimerID = SetTimer(0, 0, BotVars.ReconnectDelay, _
            '        AddressOf Reconnect_TimerProc)
            'End If
        End If
    End If
End Sub

'Locates alternative BNLS servers for the bot to use if the current one fails
'Added by FrOzeN on 2/sep/09
'Last updated by FrOzeN on 4/sep/09
Public Sub FindAltBNLS()
    'Error handler
    On Error GoTo BNLS_Alt_Finder_Error
    
    Static strBNLS()   As String
    Static intCounter  As Integer
    Static firstServer As String
    
    Const FIND_ALT_BNLS_ERROR As Integer = 12345
        
    intCounter = intCounter + 1
    
    'Close the current BNLS connection
    sckBNLS.Close
    
    'Notify user the current BNLS server failed
    'If (intCounter > 1) Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] Connection to " & BotVars.BNLSServer & " failed."
    'End If
    
    'Notify user other BNLS servers are being located
    AddChat RTBColors.InformationText, "[BNLS] Locating other BNLS servers..."
    
    'Check if the BNLS list has been downloaded
    If (GotBNLSList = False) Then
        Dim strReturn As String
        
        'Reset the counter
        intCounter = 1
                
        If (INet.StillExecuting = False) Then
            ' store first bnls server used so that we can avoid connecting to it again
            firstServer = BotVars.BNLSServer
        
            'Get the servers as a list from http://stealthbot.net/p/bnls.php
            strReturn = INet.OpenURL("http://www.stealthbot.net/p/bnls.php")
            
            ' ...
            If ((strReturn = vbNullString) Or (Left(strReturn, 1) <> vbLf)) Then
                ' ...
                AddChat RTBColors.ErrorMessageText, "[BNLS] An error occured when trying to locate an alternative BNLS server. " & _
                    "Visit http://stealthbot.net/ and check the Technical Support forum for more information."
            
                ' ...
                Call DoDisconnect
            
                ' ...
                GotBNLSList = False
            
                ' ...
                Exit Sub
            End If
            
            ' ...
            strBNLS() = Split(strReturn, vbLf)
            
            'Mark GotBNLSList as True so it's no longer downloaded for each attempt
            GotBNLSList = True
        Else
            'The Inet control seems to still be running
            Err.Raise FIND_ALT_BNLS_ERROR, , "Unable to use BNLS server finder. Visit http://stealthbot.net/ " & _
                "and check the Technical Support forum for more information."
        End If
    End If
    
    If intCounter > UBound(strBNLS) Then
        'All BNLS servers have been tried and failed
        Err.Raise FIND_ALT_BNLS_ERROR, , "All the BNLS servers have failed. Visit http://stealthbot.net/ " & _
            "and check the Technical Support forum for more information."
    End If
    
    ' ...
    If (StrComp(strBNLS(intCounter), firstServer, vbTextCompare) = 0) Then
        intCounter = intCounter + 1
        
        If intCounter > UBound(strBNLS) Then
            'All BNLS servers have been tried and failed
            Err.Raise FIND_ALT_BNLS_ERROR, , "All the BNLS servers have failed. Visit http://stealthbot.net/ " & _
                "and check the Technical Support forum for more information."
        End If
    End If
    
    ' ...
    BotVars.BNLSServer = strBNLS(intCounter)
    
    'Reconnect BNLS using the newly located BNLS server
    With sckBNLS
        .RemoteHost = BotVars.BNLSServer
        .Connect
    End With
    
    ' ...
    AddChat RTBColors.InformationText, "[BNLS] Connecting to " & BotVars.BNLSServer & "..."

    Exit Sub
    
BNLS_Alt_Finder_Error:

    'Display the error message to the user
    If Err.Number = FIND_ALT_BNLS_ERROR Then
        AddChat RTBColors.ErrorMessageText, "[BNLS] " & Err.description
        
        ' ensure that we update our listing on following connection(s)
        GotBNLSList = False
    Else
        AddChat RTBColors.ErrorMessageText, _
            "Error (#" & Err.Number & "): " & Err.description & " in FindAltBNLS()"
    End If
    
    'Disconnect the bot
    Call DoDisconnect

    Exit Sub
End Sub

' This code commented out 10/18/06 - what's it for? I dunno. It's old.
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call Form_Resize
'End Sub

' Updated 8/8/07 to support new prefix/suffix box feature
Sub Form_Resize()
    On Error Resume Next
    
    Static WasMaximized As Boolean
    Static DoMaximize As Boolean
    
    If Me.WindowState = vbMinimized Then
        If Not BotVars.NoTray Then
            #If Not COMPILE_DEBUG = 1 Then
                Me.Hide
                
                With nid
                    .cbSize = LenB(nid)
                    .hWnd = frmChat.hWnd
                    .uId = ID_TASKBARICON
                    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                    .uCallBackMessage = WM_ICONNOTIFY
                    .hIcon = frmChat.Icon.Handle
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
        With lvChannel
            rtbChat.Width = Me.Width - .Width - IIf(g_OSVersion.IsWindowsVista, 200, 120)
        
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
        
        If txtPre.Visible Then
            cboSend.Width = cboSend.Width - txtPre.Width
        End If
        
        If txtPost.Visible Then
            cboSend.Width = cboSend.Width - txtPost.Width
        End If
        
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
        
        lvChannel.Move rtbChat.Left + rtbChat.Width, lblCurrentChannel.Top + lblCurrentChannel.Height
        lvFriendList.Move lvChannel.Left, lvChannel.Top
        lvClanList.Move lvChannel.Left, lvChannel.Top
        lblCurrentChannel.Move lvChannel.Left, rtbChat.Top
        
        If txtPre.Visible Then
            txtPre.Move rtbChat.Left, rtbChat.Top + rtbChat.Height + (Screen.TwipsPerPixelY / 3)
            cboSend.Move txtPre.Left + txtPre.Width, txtPre.Top
        Else
            cboSend.Move rtbChat.Left, rtbChat.Top + rtbChat.Height + (Screen.TwipsPerPixelY / 3)
        End If
        
        If txtPost.Visible Then
            txtPost.Move cboSend.Left + cboSend.Width, cboSend.Top
        End If
        
        lvChannel.Height = rtbChat.Height - lblCurrentChannel.Height
        lvFriendList.Height = lvChannel.Height
        lvClanList.Height = lvChannel.Height
        
        'Minus 80 twips from rtbWhispers.Width if using Vista to fix width issue
        With rtbWhispers
            If .Visible Then
                .Move rtbChat.Left, cboSend.Top + cboSend.Height, (Me.Width - cmdShowHide.Width - 10 * _
                    Screen.TwipsPerPixelX) - IIf(g_OSVersion.IsWindowsVista, 80, 0)
            End If
        End With
        
        ListviewTabs.Height = cboSend.Height
        ListviewTabs.Move lvChannel.Left, cboSend.Top - (Screen.TwipsPerPixelY), lvChannel.Width - _
            cmdShowHide.Width - Screen.TwipsPerPixelX, cboSend.Height '+ 2 * Screen.TwipsPerPixelY
        
        If rtbWhispersVisible Then
            cmdShowHide.Move (((rtbWhispers.Left + rtbWhispers.Width) / Screen.TwipsPerPixelX) + 1) * _
                Screen.TwipsPerPixelX, lvChannel.Top + lvChannel.Height + Screen.TwipsPerPixelY
        Else
            cmdShowHide.Move ListviewTabs.Left + ListviewTabs.Width, lvChannel.Top + lvChannel.Height
        End If
        
        With lvClanList
            .ColumnHeaders(1).Width = (.Width \ 4) * 3 - 150
            .ColumnHeaders(2).Width = .Width \ 4 + 200
            .ColumnHeaders(3).Width = 0
        End With
        
        With lvFriendList
            .ColumnHeaders(1).Width = (.Width \ 4) * 3
            .ColumnHeaders(2).Width = .Width \ 4 + 200
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
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in Form_Resize()."
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
            .hIcon = frmChat.Icon.Handle
            .szTip = GenerateTooltip()
        End With
        
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Sub

Private Sub ClanHandler_CandidateList(ByVal Status As Byte, Users() As String)
    Dim I As Long
    
    'Valid Status codes:
    '   0x00: Successfully found candidate(s)
    '   0x01: Clan tag already taken
    '   0x08: Already in clan
    '   0x0a: Invalid clan tag specified
    
    If MDebug("debug") Then
        AddChat RTBColors.ErrorMessageText, "CandidateList received. Status code [0x" & Hex(Status) & "]."
        If UBound(Users) > -1 Then
            AddChat RTBColors.InformationText, "Potential clan members:"
            
            For I = 0 To UBound(Users)
                AddChat RTBColors.InformationText, Users(I)
            Next I
        End If
    End If
End Sub

Private Sub ClanHandler_MemberLeaves(ByVal Member As String)
    AddChat vbYellow, "[CLAN] " & Member & " has left the clan."
    
    Dim X   As ListItem
    Dim Pos As Integer
    
    Pos = g_Clan.GetUserIndexEx(Member)
    
    ' ...
    If (Pos > 0) Then
        g_Clan.Members.Remove Pos
    End If
    
    Member = convertUsername(Member)

    Set X = lvClanList.FindItem(Member)
    
    If (Not (X Is Nothing)) Then
        lvClanList.ListItems.Remove X.Index
        
        lvClanList.Refresh
        
        Set X = Nothing
    End If
    
    On Error Resume Next

    SControl.Run "Event_ClanMemberLeaves", Member
End Sub

Private Sub ClanHandler_RemovedFromClan(ByVal Status As Byte)
    If Status = 1 Then
        If (AwaitingSelfRemoval = 0) Then
            Set g_Clan = New clsClanObj
        
            Clan.isUsed = False
        
            ListviewTabs.TabEnabled(2) = False
            lvClanList.ListItems.Clear
            ListviewTabs.Tab = 0
            Call ListviewTabs_Click(2)
        
            AddChat RTBColors.ErrorMessageText, "[CLAN] You have been removed from the clan, or it has been disbanded."
        
            On Error Resume Next
            SControl.Run "Event_BotRemovedFromClan"
        End If
        
        On Error Resume Next
        SControl.Run "Event_BotRemovedFromClan"
    End If
End Sub

Private Sub ClanHandler_MyRankChange(ByVal NewRank As Byte)
    If (g_Clan.Self.Rank < NewRank) Then
        AddChat RTBColors.SuccessText, "[CLAN] You have been promoted. Your new rank is ", _
                RTBColors.InformationText, GetRank(NewRank), RTBColors.SuccessText, "."
    ElseIf (g_Clan.Self.Rank > NewRank) Then
        AddChat RTBColors.SuccessText, "[CLAN] You have been demoted. Your new rank is ", _
                RTBColors.InformationText, GetRank(NewRank), RTBColors.SuccessText, "."
    Else
        AddChat RTBColors.SuccessText, "[CLAN] Your new rank is ", RTBColors.InformationText, _
                GetRank(NewRank), RTBColors.SuccessText, "."
    End If

    g_Clan.Self.Rank = NewRank
    
    On Error Resume Next
    
    SControl.Run "Event_BotClanRankChanged", NewRank
End Sub

Private Sub ClanHandler_ClanInfo(ByVal ClanTag As String, ByVal RawClanTag As String, ByVal Rank As Byte)
    Set g_Clan = New clsClanObj
    
    With Clan
        .Name = ClanTag
        .DWName = RawClanTag
        .MyRank = Rank
        .isUsed = True
    End With
    
    With g_Clan
        .Name = ClanTag
    End With
    
    Call InitListviewTabs
    
    'If g_Clan.Self.Rank = 0 Then g_Clan.Self.Rank = 1
    On Error Resume Next
    
    ClanTag = KillNull(ClanTag)
    
    BotVars.Clan = ClanTag
    
    If AwaitingClanMembership = 1 Then
        AddChat RTBColors.SuccessText, "[CLAN] You are now a member of ", RTBColors.InformationText, "Clan " & ClanTag, RTBColors.SuccessText, "!"
        AwaitingClanMembership = 0
            
        SControl.Run "Event_BotJoinedClan", ClanTag
    Else
        AddChat RTBColors.SuccessText, "[CLAN] You are a ", RTBColors.InformationText, GetRank(Rank), RTBColors.SuccessText, " in ", RTBColors.InformationText, "Clan " & ClanTag, RTBColors.SuccessText, "."
        
        SControl.Run "Event_BotClanInfo", ClanTag, Rank
    End If
    
    RequestClanList
    RequestClanMOTD
    
    'frmChat.ClanHandler.RequestClanMotd 1
    
    ' ...
    frmChat.ListviewTabs_Click 0
End Sub

Private Sub ClanHandler_ClanInvitation(ByVal Token As String, ByVal ClanTag As String, ByVal RawClanTag As String, ByVal ClanName As String, ByVal InvitedBy As String, ByVal NewClan As Boolean)
    If Not mnuIgnoreInvites.Checked And IsW3 Then
        Clan.Token = Token
        Clan.DWName = RawClanTag
        Clan.Creator = InvitedBy
        Clan.Name = ClanName
        If NewClan Then Clan.isNew = 1
        
        With RTBColors
            AddChat .SuccessText, "[CLAN] ", .InformationText, convertUsername(InvitedBy), .SuccessText, " has invited you to join a clan: ", .InformationText, ClanName, .SuccessText, " [", .InformationText, ClanTag, .SuccessText, "]"
        End With
        
        frmClanInvite.Show
    End If
End Sub

Private Sub ClanHandler_ClanMemberList(Members() As String)
    Dim ClanMember As clsClanMemberObj
    Dim I          As Long
    
    ' ...
    If AwaitingClanList = 1 Then
        ' ...
        g_Clan.Clear
    
        ' ...
        lvClanList.ListItems.Clear
    
        ' ...
        For I = 0 To UBound(Members) Step 4
            ' ...
            Set ClanMember = New clsClanMemberObj
            
            ' ...
            With ClanMember
                .Name = Members(I)
                .Rank = Val(Members(I + 1))
                .Status = Val(Members(I + 2))
                .Location = Members(I + 3)
            End With

            ' ...
            g_Clan.Members.Add ClanMember
        
            ' ...
            If ((Len(Members(I)) > 0) And (UBound(Members) >= I + 1)) Then
                ' ...
                AddClanMember ClanMember.DisplayName, Val(Members(I + 1)), Val(Members(I + 2))
                
                ' ...
                On Error Resume Next
                
                ' ...
                SControl.Run "Event_ClanMemberList", ClanMember.DisplayName, Val(Members(I + 1)), _
                    Val(Members(I + 2))
            End If
        Next I
    End If
    
    ' ...
    lblCurrentChannel.Caption = GetChannelString()
    
    ' ...
    frmChat.ListviewTabs_Click 0
End Sub

Private Sub ClanHandler_ClanMemberUpdate(ByVal Username As String, ByVal Rank As Byte, ByVal IsOnline As Byte, ByVal Location As String)
    Dim X   As ListItem
    Dim Pos As Integer
    
    Pos = g_Clan.GetUserIndexEx(Username)
    
    If (Pos > 0) Then
        With g_Clan.Members(Pos)
            .Rank = Rank
            .Status = IsOnline
            .Location = Location
        End With
    Else
        Dim ClanMember As clsClanMemberObj
        
        ' ...
        Set ClanMember = New clsClanMemberObj
        
        ' ...
        With ClanMember
            .Rank = Rank
            .Status = IsOnline
            .Location = Location
        End With
        
        ' ...
        g_Clan.Members.Add ClanMember
    End If
    
    Username = convertUsername(Username)
    
    Set X = lvClanList.FindItem(Username)

    If StrComp(Username, CurrentUsername, vbTextCompare) = 0 Then
        g_Clan.Self.Rank = IIf(Rank = 0, Rank + 1, Rank)
        AwaitingClanInfo = 1
    End If
    
    If AwaitingClanInfo = 1 Then
        AwaitingClanInfo = 0
        AddChat RTBColors.SuccessText, "[CLAN] Member update: ", RTBColors.InformationText, Username, RTBColors.SuccessText, " is now a " & GetRank(Rank) & "."
    End If
    
    If Not (X Is Nothing) Then
        lvClanList.ListItems.Remove X.Index
        Set X = Nothing
    End If
    
    AddClanMember Username, CInt(Rank), CInt(IsOnline)
    
    On Error Resume Next
    SControl.Run "Event_ClanMemberUpdate", Username, Rank, IsOnline
End Sub

Private Sub ClanHandler_ClanMOTD(ByVal cookie As Long, ByVal Message As String)
    ' ...
    g_Clan.MOTD = Message
    
    ' ...
    If (cookie = 1) Then
        PassedClanMotdCheck = True
        
        ' ...
        'If (g_Clan.MOTD <> vbNullString) Then
        '    frmChat.AddChat vbBlue, Message
        'End If
    End If
    
    On Error Resume Next
    
    SControl.Run "Event_ClanMOTD", Message
End Sub

Private Sub ClanHandler_DemoteUserReply(ByVal Success As Boolean)
    If Success Then
        AddChat RTBColors.SuccessText, "[CLAN] User demoted successfully."
    Else
        AddChat RTBColors.ErrorMessageText, "[CLAN] User demotion failed."
    End If
    
    lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub ClanHandler_DisbandClanReply(ByVal Success As Boolean)
    If MDebug("debug") Then
        AddChat RTBColors.ConsoleText, "DisbandClanReply: " & Success
    End If
End Sub

Private Sub ClanHandler_InviteUserReply(ByVal Status As Byte)
    '0x00: Invitation accepted
    '0x04: Invitation declined
    '0x05: Failed to invite user
    '0x09: Clan is full
    
    Select Case Status
        Case 0, 1: AddChat RTBColors.SuccessText, "[CLAN] The invitation was accepted."
        Case 4: AddChat RTBColors.ErrorMessageText, "[CLAN] The invitation was declined."
        Case 5: AddChat RTBColors.ErrorMessageText, "[CLAN] The invitation failed."
        Case 9: AddChat RTBColors.ErrorMessageText, "[CLAN] The invitation failed: Your clan is full."
        Case Else: AddChat RTBColors.ErrorMessageText, "[CLAN] Unknown invitation status: " & Status
    End Select
End Sub

Private Sub ClanHandler_PromoteUserReply(ByVal Success As Boolean)
    If Success Then
        AddChat RTBColors.SuccessText, "[CLAN] User promoted successfully."
    Else
        AddChat RTBColors.ErrorMessageText, "[CLAN] User promotion failed."
    End If
    
    lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub ClanHandler_RemoveUserReply(ByVal result As Byte)

'    0x00: Successfully removed user from clan
'    0x02: Too soon to remove user
'    0x03: Not enough members to remove this user
'    0x07: Not authorized to remove the user
'    0x08: User is not in your clan
    
    'Debug.Print "Removed successfully!"
    
    Select Case result
        Case 0
            If AwaitingSelfRemoval = 1 Then
                AwaitingSelfRemoval = 0
                Clan.isUsed = False
                
                ListviewTabs.TabEnabled(2) = False
                lvClanList.ListItems.Clear
                
                ListviewTabs.TabIndex = 0
                Call ListviewTabs_Click(2)
                
                Set g_Clan = New clsClanObj
                
                AddChat RTBColors.SuccessText, "[CLAN] You have successfully left the clan."
            Else
                AddChat RTBColors.SuccessText, "[CLAN] User removed successfully."
                lvClanList.ListItems.Clear
                RequestClanList
            End If
            
        Case 2
            AddChat RTBColors.ErrorMessageText, "[CLAN] That user is currently on probation."
        
        Case 3
            AddChat RTBColors.ErrorMessageText, "[CLAN] There are not enough members for you to remove that user."
        
        Case 7
            AddChat RTBColors.ErrorMessageText, "[CLAN] You are not authorized to remove that user."
            
        Case 8
            AddChat RTBColors.ErrorMessageText, "[CLAN] You are not allowed to remove that user."
            
        Case Else
            AddChat RTBColors.InformationText, "[CLAN] 0x78 Response code: 0x" & Hex(result)
            AddChat RTBColors.InformationText, "[CLAN] You failed to remove that user from the clan."
    End Select
    
    lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub ClanHandler_UnknownClanEvent(ByVal PacketID As Byte, ByVal Data As String)
    If MDebug("debug") Then
        frmChat.AddChat RTBColors.ErrorMessageText, "[CLAN] Unknown clan event [0x" & Hex(PacketID) & "]. Data is as follows:"
        frmChat.AddChat RTBColors.ErrorMessageText, Data
    End If
End Sub

Public Function GetLogFilePath() As String

    Dim Path As String  ' ...
    Dim f    As Integer ' ...
    
    ' ...
    f = FreeFile
    
    ' ...
    Path = _
        GetProfilePath() & "\Logs\" & Format(Date, "YYYY-MM-DD") & ".txt"

    If (Dir$(Path) = vbNullString) Then
        Open Path For Output As #f
        Close #1
    End If
    
    GetLogFilePath = Path

End Function

Sub Form_Unload(Cancel As Integer)
    Dim Key As String, L As Long
    
    'Me.WindowState = vbNormal
    'Me.Show

    'UserCancelledConnect = False

    'Cancel = 1
    
    'scTimer.Enabled = False
    'SControl.Reset
    
    INet.Cancel
    
    AddChat RTBColors.ErrorMessageText, "Shutting down..."
    
    If LenB(Dir$(GetConfigFilePath())) > 0 Then
        If Me.WindowState <> vbMinimized Then
            Call RecordWindowPosition(CBool(Me.WindowState = vbMaximized))
        End If
        
        WriteINI "Main", "ConfigVersion", CONFIG_VERSION
    End If
    
    'With frmChat.INet
    '    If .StillExecuting Then .Cancel
    'End With

    Call DoDisconnect(1)

    Shell_NotifyIcon NIM_DELETE, nid
    
    On Error Resume Next
    
    SControl.Run "Event_Close"
    
    On Error GoTo 0
    
    If BotVars.Logging = 1 Then
        'Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt" For Append As #1
        '    If LOF(1) < BotVars.MaxLogFileSize Then
        '        Print #1, "Bot application closed, dumping chat screen."
        '        Print #1, rtbChat.text
        '    End If
        'Close #1
        
        'Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt" For Append As #1
        '    If LOF(1) < BotVars.MaxLogFileSize Then
        '        Print #1, "Bot application closed, dumping whisper screen."
        '        Print #1, rtbWhispers.text
        '    End If
        'Close #1
    ElseIf BotVars.Logging = 2 Then
        'Kill GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt"
        'Kill GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt"
    End If
    
    'If ReconnectTimerID > 0 Then
    '    KillTimer 0, ReconnectTimerID
    'End If
    
    If ExReconnectTimerID > 0 Then
        KillTimer 0, ExReconnectTimerID
    End If
    
'    If AttemptedNewVerbyte Then
'        AttemptedNewVerbyte = False
'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
'    End If

    ' COMMENTED OUT 2009-28-01 Eric[nK]
    '  BREAKS COMPILATION
    'Call modWarden.WardenCleanUp

    'Call ChatQueue_Terminate

    DisableURLDetect
    UnhookWindowProc
    UnhookSendBoxWindowProc
    
    'Call SharedScriptSupport.Dispose 'Explicit call the Class_Terminate sub in the ScriptSupportClass to destroy all the forms. - FrOzeN
    
    'DeconstructSettings
    'DeconstructMonitor
    DestroyAllWWs
    
    Set g_Logger = Nothing
    Set BotVars = Nothing
    Set ClanHandler = Nothing
    Set ListToolTip = Nothing
    Set GErrorHandler = Nothing
    Set FriendListHandler = Nothing
    Set colQueue = Nothing
    Set colWhisperWindows = Nothing
    Set colLastSeen = Nothing
    Set colSafelist = Nothing
    Set dctCallbacks = Nothing
    Set colDynamicMenus = Nothing
    Set dictMenuIDs = Nothing
    Set dictItemIDs = Nothing
    Set SharedScriptSupport = Nothing
    Set ds = Nothing
    Set NLogin = Nothing
    
    'Set dictTimerInterval = Nothing
    'Set dictTimerCount = Nothing
    'Set dictTimerEnabled = Nothing
    
    ' Updated to match current form list 2009-02-09 Andy
    Unload frmAbout
    Unload frmCatch
    Unload frmCommands
    Unload frmClanInvite
    Unload frmCustomInputBox
    Unload frmDBManager
    Unload frmDBType
    Unload frmEMailReg
    Unload frmFilters
    Unload frmGameSelection
    Unload frmManageKeys
    'Unload frmMonitor
    Unload frmProfile
    'Unload frmProfileManager
    Unload frmQuickChannel
    Unload frmRealm
    Unload frmScriptUI
    Unload frmSettings
    Unload frmSplash
    'Unload frmUserManager
    Unload frmWhisperWindow
    Unload frmWriteProfile
    
    ' Added this instead of End to try and fix some system tray crashes 2009-0211-andy
    '  It was used in some capacity before since the API was already declared
    '   in modAPI...
    Call ExitProcess(0)
End Sub


Public Sub AddFriend(ByVal Username As String, ByVal Product As String, IsOnline As Boolean)
    Dim I As Integer, OnlineIcon As Integer
    Dim f As ListItem
    
    Const ICONLINE = 23
    Const ICOFFLINE = 24
    
    If IsOnline Then OnlineIcon = ICONLINE Else OnlineIcon = ICOFFLINE
    
    'Everybody Else
    Select Case Product
        Case Is = "STAR"
            I = ICSTAR
        Case Is = "SEXP"
            I = ICSEXP
        Case Is = "D2DV"
            I = ICD2DV
        Case Is = "D2XP"
            I = ICD2XP
        Case Is = "W2BN"
            I = ICW2BN
        Case Is = "WAR3"
            I = ICWAR3
        Case Is = "W3XP"
            I = ICWAR3X
        Case Is = "CHAT"
            I = ICCHAT
        Case Is = "DRTL"
            I = ICDIABLO
        Case Is = "DSHR"
            I = ICDIABLOSW
        Case Is = "JSTR"
            I = ICJSTR
        Case Is = "SSHR"
            I = ICSCSW
        Case Else
            I = ICUNKNOWN
    End Select
    
    Set f = lvFriendList.FindItem(Username)
    
    If (f Is Nothing) Then
        With lvFriendList.ListItems
            .Add , , Username, , I
            .Item(.Count).ListSubItems.Add , , , OnlineIcon
        End With
    Else
        f.SmallIcon = I
        f.ListSubItems.Item(1).ReportIcon = OnlineIcon
        
        Set f = Nothing
    End If
    
    ' ...
    lblCurrentChannel.Caption = GetChannelString()
    
    ' ...
    frmChat.ListviewTabs_Click 0
End Sub

Private Sub FriendListHandler_FriendAdded(ByVal Username As String, ByVal Product As String, ByVal Location As Byte, ByVal Status As Byte, ByVal Channel As String)
    'AddFriend Username, Product, (Location > 0)
    'lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub FriendListHandler_FriendListEntry(ByVal Username As String, ByVal Product As String, ByVal Channel As String, ByVal Status As Byte, ByVal Location As Byte)
    AddFriend Username, Product, (Location > 0)
    lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub FriendListHandler_FriendMoved()
    'lvFriendList.ListItems.Clear
    Call FriendListHandler.ResetList
    Call FriendListHandler.RequestFriendsList(PBuffer)
End Sub

Private Sub FriendListHandler_FriendRemoved(ByVal Username As String)
    'Dim X As ListItem
    
    'Set X = lvFriendList.FindItem(Username)
   
    'If (Not (X Is Nothing)) Then
    '    lvFriendList.ListItems.Remove X.index
    '
    '    Set X = Nothing
    'End If
    
    'lblCurrentChannel.Caption = GetChannelString
End Sub

Private Sub FriendListHandler_FriendUpdate(ByVal Username As String, ByVal FLIndex As Byte)
    On Error GoTo ERROR_HANDLER

    Dim X As ListItem
    Dim I As Integer
    Const ICONLINE = 23
    Const ICOFFLINE = 24
    
    Set X = lvFriendList.FindItem(Username)
    
    If Not (X Is Nothing) Then
        With g_Friends.Item(FLIndex)
            Select Case .LocationID
                Case FRL_OFFLINE
                    X.SmallIcon = ICUNKNOWN
                    
                    X.ListSubItems.Item(1).ReportIcon = ICOFFLINE
                    
                Case Else
                    If X.ListSubItems.Item(1).ReportIcon = ICOFFLINE Then
                        'Friend is now online - notify user?
                    End If
                    
                    X.ListSubItems.Item(1).ReportIcon = ICONLINE
                    
                    Select Case .game
                        Case Is = "STAR": I = ICSTAR
                        Case Is = "SEXP": I = ICSEXP
                        Case Is = "D2DV": I = ICD2DV
                        Case Is = "D2XP": I = ICD2XP
                        Case Is = "W2BN": I = ICW2BN
                        Case Is = "WAR3": I = ICWAR3
                        Case Is = "W3XP": I = ICWAR3X
                        Case Is = "CHAT": I = ICCHAT
                        Case Is = "DRTL": I = ICDIABLO
                        Case Is = "DSHR": I = ICDIABLOSW
                        Case Is = "JSTR": I = ICJSTR
                        Case Is = "SSHR": I = ICSCSW
                        Case Else: I = ICUNKNOWN
                    End Select
                    
                    X.SmallIcon = I
            End Select
        End With
        
    End If
    
    Set X = Nothing
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in FriendUpdate()."
    
    Exit Sub
End Sub

Private Sub INet_StateChanged(ByVal State As Integer)
    On Error GoTo ERROR_HANDLER

    'If (State = icResponseCompleted) Then
    '    Call HandleNews(INet.GetChunk(1024, icString))
        
    'The code below is useless, icError doesn't mean it's an issue with the news as it could also be a script _
        calling the the INet control. - FrOzeN
    
    'Else
    '    If (State = icError) Then
            
            'Call AddChat(RTBColors.ErrorMessageText, "Error: There was an error " & _
                "loading the news.")
    '    End If
    'End If
    
    If (Not (BotLoaded)) Then
        'SControl.Run "Event_FirstRun"
        'SControl.Run "Event_Load"
            
        BotLoaded = True
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    Exit Sub
End Sub

Private Sub lblCurrentChannel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer ' ...
 
    ' ...
    If ((mnuPublicChannels(0).Caption <> vbNullString) Or _
            (mnuCustomChannels(0).Caption <> vbNullString)) Then
    
        ' ...
        If ((mnuPublicChannels(0).Caption <> vbNullString) And _
                (mnuCustomChannels(0).Caption <> vbNullString)) Then
            
            ' ...
            frmChat.mnuDash.Visible = True
        Else
            ' ...
            frmChat.mnuDash.Visible = False
        End If
        
        ' ...
        For I = 0 To mnuPublicChannels.Count - 1
            ' ...
            If (mnuPublicChannels(I).Caption <> vbNullString) Then
                mnuPublicChannels(I).Visible = True
            End If
        Next I
        
        ' ...
        mnuQCEdit.Visible = False
        mnuBlank.Visible = False
        
        ' ...
        PopupMenu mnuQCTop
    End If
End Sub

Public Sub ListviewTabs_Click(PreviousTab As Integer)
    Dim CurrentTab As Integer
    
    CurrentTab = ListviewTabs.Tab
    
    'If PreviousTab <> CurrentTab And ListviewTabs.TabEnabled(CurrentTab) Then
        Select Case CurrentTab
            Case LVW_BUTTON_CHANNEL ' = 0 = Channel button clicked
                ' ...
                lblCurrentChannel.ToolTipText = "Currently in " & g_Channel.SType() & _
                    " channel " & g_Channel.Name & " (" & g_Channel.Users.Count & ")"
                
                ' ...
                lvChannel.ZOrder vbBringToFront
                
            Case LVW_BUTTON_FRIENDS ' = 1 = Friends button clicked
                ' ...
                lblCurrentChannel.ToolTipText = "Currently viewing " & g_Friends.Count & " friends"
            
                ' ...
                lvFriendList.ZOrder vbBringToFront
                
            Case LVW_BUTTON_CLAN ' = 2 = Clan button clicked
                ' ...
                lblCurrentChannel.ToolTipText = "Currently viewing " & _
                    g_Clan.Members.Count & " members of clan " & Clan.Name
            
                ' ...
                lvClanList.ZOrder vbBringToFront
                
        End Select
    'End If
    
    lblCurrentChannel.Caption = GetChannelString
End Sub

' This procedure relies on code in RecordcboSendSelInfo() that sets global variables
'  cboSendSelLength and cboSendSelStart
' These two properties are zeroed out as the control loses focus and inaccessible
'  (zeroed) at both access time in this method AND in the _LostFocus sub
Private Sub lvChannel_dblClick()
    Dim s           As String
    Dim T           As String
    Dim oldSelStart As Long
    
    s = GetSelectedUser
    oldSelStart = cboSendSelStart

    If (Len(s) > 0) Then
        With cboSend
            .SelStart = cboSendSelStart 'IIf(cboSendSelStart > 0, cboSendSelStart, 0)
            .SelLength = cboSendSelLength 'IIf(cboSendSelLength > 0, cboSendSelLength + 1, 0)
            .SelText = s
            
            ' This is correct - sets the cursor properly
            cboSendSelStart = Len(.text)
            cboSendSelLength = 0
            
            .SetFocus
        End With
    End If
End Sub

Private Sub lvChannel_KeyUp(KeyCode As Integer, Shift As Integer)
    Const S_ALT = 4
   
    If Shift = S_ALT And KeyCode = KEY_ALTN Then
        Dim sStart As Integer
        
        With lvChannel
            If Not (.SelectedItem Is Nothing) Then
                cboSend.SelStart = Len(cboSend.text)
                cboSend.SelText = .SelectedItem.text
    
                KeyCode = 0
                Shift = 0
                
                Exit Sub
            End If
        End With
    End If
End Sub

Private Sub lvFriendList_dblClick()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        cboSend.text = cboSend.text & lvFriendList.SelectedItem.text
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.text)
    End If
End Sub

Private Sub lvClanList_dblClick()
    If Not (lvClanList.SelectedItem Is Nothing) And Len(cboSend.text) < 200 Then
        cboSend.text = cboSend.text & lvClanList.SelectedItem.text
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.text)
    End If
End Sub

Private Sub lvChannel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim aInx As Integer
    Dim sProd As String * 4
    
    ' ...
    If (lvChannel.SelectedItem Is Nothing) Then
        Exit Sub
    End If

    If Button = 2 Then
        If Not (lvChannel.SelectedItem Is Nothing) Then
            aInx = g_Channel.GetUserIndex(GetSelectedUser)
            
            If aInx > 0 Then
                sProd = g_Channel.Users(aInx).game

                mnuPopWebProfile.Enabled = (sProd = "W3XP" Or sProd = "WAR3")
                mnuPopInvite.Enabled = (mnuPopWebProfile.Enabled And g_Clan.Self.Rank >= 3)
                mnuPopKick.Enabled = (MyFlags = 2 Or MyFlags = 18)
                mnuPopDes.Enabled = (MyFlags = 2 Or MyFlags = 18)
                mnuPopBan.Enabled = (MyFlags = 2 Or MyFlags = 18)
            End If
        Else
            mnuPopWebProfile.Enabled = False
        End If
        
        mnuPopup.Tag = lvChannel.SelectedItem.text 'Record which user is selected at time of right-clicking. - FrOzeN
        
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lvFriendList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuFLpop
    End If
End Sub

Private Sub lvFriendList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
   
    lvhti.pt.X = X / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessageAny(lvFriendList.hWnd, LVM_HITTEST, 0, lvhti) + 1
   
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            ListToolTip.Destroy
        Else
            ListToolTip.Title = "Information for " & lvFriendList.ListItems(m_lCurItemIndex).text
            
            Dim sTemp As String
            
            If ((lItemIndex > 0) And (g_Friends.Count > 0)) Then
                lItemIndex = FriendListHandler.UsernameToFLIndex(lvFriendList.ListItems(m_lCurItemIndex).text)
            
                With g_Friends.Item(lItemIndex)
'                    Private Const FRL_OFFLINE& = &H0
'                    Private Const FRL_NOTINCHAT& = &H1
'                    Private Const FRL_INCHAT& = &H2
'                    Private Const FRL_PUBLICGAME& = &H3
'                    Private Const FRL_PRIVATEGAME& = &H5
                    If .IsOnline Then
                        sTemp = sTemp & "Using " & ProductCodeToFullName(.game) & " "
                    End If
                    
                    Select Case .LocationID
                        Case FRL_OFFLINE
                            sTemp = sTemp & "This person is offline."
                        Case FRL_NOTINCHAT
                            sTemp = sTemp & "in limbo. (not yet in chat)"
                        Case FRL_INCHAT
                            sTemp = sTemp & "in a Battle.net channel."
                        Case FRL_PUBLICGAME
                            sTemp = sTemp & "in a public game."
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
                            Case FRL_PRIVATEGAME
                                sTemp = sTemp & ", in game " & .Location & "."
                            Case Else
                                sTemp = sTemp & "."
                        End Select
                    End If
                    
                    If (.LocationID = FRL_PUBLICGAME) Then
                        sTemp = sTemp & vbCrLf & "Currently in the public game '" & .Location & "'."
                    End If
                    
                    ListToolTip.TipText = sTemp
                End With
                
                Call ListToolTip.Create(lvFriendList.hWnd, CLng(X), CLng(Y))
            End If
        End If
    End If
End Sub

Private Sub lvChannel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    Dim sOutBuf As String
    Dim sTemp As String
    Dim UserAccess As udtGetAccessResponse
    Dim Clan As String
   
    lvhti.pt.X = X / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessageAny(lvChannel.hWnd, LVM_HITTEST, -1, lvhti) + 1
 
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        
        If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
            ListToolTip.Destroy
        Else
            ' ...
            'UserAccess = GetCumulativeAccess(lvChannel.ListItems(m_lCurItemIndex).text, "USER")
        
            ' ...
            ListToolTip.Title = _
                "Information for " & lvChannel.ListItems(m_lCurItemIndex).text
                
            ' ...
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
                
            
            lItemIndex = g_Channel.GetUserIndex(lvChannel.ListItems(m_lCurItemIndex).text)
            
            If (lItemIndex > 0) Then
                With g_Channel.Users(lItemIndex)
                    'ParseStatstring .Statstring, sOutBuf, Clan
            
                    'sTemp = sTemp & vbCrLf
                    sTemp = sTemp & "Ping at login: " & .Ping & "ms" & vbCrLf
                    sTemp = sTemp & "Flags: " & FlagDescription(.Flags) & vbCrLf
                    sTemp = sTemp & vbCrLf
                    sTemp = sTemp & .Stats.ToString
                
                    ListToolTip.TipText = sTemp
                    
                End With
                
                Call ListToolTip.Create(lvChannel.hWnd, CLng(X), CLng(Y))
            End If
        End If
    End If
End Sub

Private Sub mnuBot_Click()
    Dim I As Integer

    If IsW3 And g_Connected Then
        mnuIgnoreInvites.Enabled = True
    Else
        mnuIgnoreInvites.Enabled = False
    End If
    
    ' ...
    mnuDash.Visible = False

    ' ...
    For I = 0 To mnuPublicChannels.Count - 1
        mnuPublicChannels(I).Visible = False
    Next I
    
    ' ...
    mnuQCEdit.Visible = True
    mnuBlank.Visible = True
End Sub

Private Sub mnuCatchPhrases_Click()
    frmCatch.Show
End Sub

Private Sub mnuPublicChannels_Click(Index As Integer)
    ' ...
    If (StrComp(mnuPublicChannels(Index).Caption, g_Channel.Name, vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    'FullJoin mnuChannels(Index).Caption
    
    ' ...
    AddQ "/join " & mnuPublicChannels(Index).Caption, PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuCustomChannels_Click(Index As Integer)
    ' ...
    If (StrComp(mnuCustomChannels(Index).Caption, g_Channel.Name, vbTextCompare) = 0) Then
        Exit Sub
    End If

    'FullJoin mnuChannels(Index).Caption
    
    ' ...
    AddQ "/join " & mnuCustomChannels(Index).Caption, PRIORITY.CONSOLE_MESSAGE
End Sub

'Private Sub mnuCCEditor_Click()
'    frmCCEditor.Show
'End Sub

Sub mnuClearWW_Click()
    rtbWhispers.text = ""
    AddWhisper RTBColors.ConsoleText, ">> Whisper window cleared."
End Sub

Private Sub mnuCommandManager_Click()
    frmCommands.Show vbModal, Me
End Sub

Private Sub mnuConnect2_Click()
    Call DoConnect
End Sub

Private Sub mnuDisableVoidView_Click()
    mnuDisableVoidView.Checked = Not (mnuDisableVoidView.Checked)
    
    WriteINI "Other", "DisableVoidView", IIf(mnuDisableVoidView.Checked, "Y", "N")
End Sub

Private Sub mnuDisconnect2_Click()
    Dim Key As String, L As Long
    Key = GetProductKey()
    
'    If AttemptedNewVerbyte Then
'        AttemptedNewVerbyte = False
'        l = CLng(Val("&H" & ReadCFG("Main", Key & "VerByte")))
'        WriteINI "Main", Key & "VerByte", Hex(l - 1)
'    End If
    
    GErrorHandler.Reset
    Call DoDisconnect
End Sub

Private Sub mnuEditAccessFlags_Click()
    Shell "notepad " & App.Path & "\commands.xml", vbNormalFocus
End Sub

Private Sub mnuEditCaught_Click()
    If Dir$(GetFilePath("caughtphrases.htm")) = vbNullString Then
        MsgBox "The bot has not caught any phrases yet."
        Exit Sub
    Else
        ShellExecute Me.hWnd, "Open", GetFilePath("caughtphrases.htm"), 0&, 0&, 0&
    End If
End Sub

Private Sub mnuEditChangelog_Click()
    Shell "notepad " & App.Path & "\Changelog.txt", vbNormalFocus
End Sub

Private Sub mnuEditPhrasebans_Click()
    Shell "notepad " & GetFilePath("phrasebans.txt"), vbNormalFocus
End Sub

Private Sub mnuEditQCini_Click()
    Shell "notepad " & GetFilePath("quickchannels.ini"), vbNormalFocus
End Sub

Private Sub mnuEditSafelist_Click()
    Shell "notepad " & GetFilePath("safelist.txt"), vbNormalFocus
End Sub

Private Sub mnuEditScript_Click()
    Shell "notepad " & GetFilePath("script.txt"), vbNormalFocus
End Sub

Private Sub mnuEditShitlist_Click()
    Shell "notepad " & GetFilePath("autobans.txt"), vbNormalFocus
End Sub

Private Sub mnuEditTagbans_Click()
    Shell "notepad " & GetFilePath("tagbans.txt"), vbNormalFocus
End Sub

Private Sub mnuFlash_Click()
    If mnuFlash.Checked Then
        mnuFlash.Checked = False
        WriteINI "Other", "FlashWindow", "N"
    Else
        mnuFlash.Checked = True
        WriteINI "Other", "FlashWindow", "Y"
    End If
End Sub

'Will move the selected user one spot down on the friends list.
Private Sub mnuFLpopDemote_Click()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        With lvFriendList.SelectedItem
            If (.Index < lvFriendList.ListItems.Count) Then
              AddQ "/f d " & .text, PRIORITY.CONSOLE_MESSAGE
              'MoveFriend .index, .index + 1
            End If
        End With
    End If
End Sub

'Will move the selected user one spot up on the friends list.
Private Sub mnuFLpopPromote_Click()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        With lvFriendList.SelectedItem
            If (.Index > 1) Then
              AddQ "/f p " & .text, PRIORITY.CONSOLE_MESSAGE
              'MoveFriend .index, .index - 1
            End If
        End With
    End If
End Sub

'Moves a person in the friends list view.
Private Sub MoveFriend(startPos As Integer, endPos As Integer)
    With lvFriendList.ListItems
        If (startPos > endPos) Then
            .Add endPos, , .Item(startPos).text, , .Item(startPos).SmallIcon
            .Item(endPos).ListSubItems.Add , , , .Item(startPos + 1).ListSubItems.Item(1).ReportIcon
            .Remove startPos + 1
        Else
            .Add endPos + 1, , .Item(startPos).text, .Item(startPos).Icon, .Item(startPos).SmallIcon
            .Item(endPos + 1).ListSubItems.Add , , , .Item(startPos).ListSubItems.Item(1).ReportIcon
            .Remove startPos
        End If
    End With
End Sub

Private Sub mnuFLPopRemove_Click()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        AddQ "/f r " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(lvFriendList.SelectedItem.text), _
            PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuFLPopWhisper_Click()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        AddQ "/w " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(lvFriendList.SelectedItem.text) & _
            Space(1) & cboSend.text, PRIORITY.CONSOLE_MESSAGE
            
        cboSend.text = ""
    End If
End Sub

Private Sub mnuGetNews_Click()
    On Error Resume Next
    
    INet.Cancel
    INet.Execute GetNewsURL(), "GET"
End Sub

Sub mnuHelpReadme_Click()
    OpenReadme
End Sub

Sub mnuHelpWebsite_Click()
    ShellExecute Me.hWnd, "Open", "http://www.stealthbot.net", 0&, 0&, 0&
End Sub

Private Sub mnuHideBans_Click()
    mnuHideBans.Checked = (Not mnuHideBans.Checked)

    WriteINI "Main", "HideBanMessages", IIf(mnuHideBans.Checked, "Y", "N")
    AddChat vbYellow, "Ban messages " & IIf(mnuHideBans.Checked, "disabled", "enabled") & "."
End Sub

Private Sub mnuHideWhispersInrtbChat_Click()
    mnuHideWhispersInrtbChat.Checked = (Not mnuHideWhispersInrtbChat.Checked)
    
    WriteINI "Main", "HideWhispersInMain", IIf(mnuHideWhispersInrtbChat.Checked, "Y", "N")
End Sub

Private Sub mnuIgnoreInvites_Click()
    If mnuIgnoreInvites.Checked Then
        mnuIgnoreInvites.Checked = False
        WriteINI "Main", "IgnoreClanInvitations", "N"
    Else
        mnuIgnoreInvites.Checked = True
        WriteINI "Main", "IgnoreClanInvitations", "Y"
    End If
End Sub

Private Sub mnuLog0_Click()
    BotVars.Logging = 2
    WriteINI "Other", "Logging", "2"
    AddChat RTBColors.InformationText, "Full text logging enabled."
    mnuLog1.Checked = False
    mnuLog0.Checked = True
    mnuLog2.Checked = False
    'mnuLog3.Checked = False
    
    'MakeLoggingDirectory
End Sub

Private Sub mnuLog1_Click()
    BotVars.Logging = 1
    WriteINI "Other", "Logging", "1"
    AddChat RTBColors.InformationText, "Partial text logging enabled."
    mnuLog1.Checked = True
    mnuLog0.Checked = False
    mnuLog2.Checked = False
    'mnuLog3.Checked = False
    
    'MakeLoggingDirectory
End Sub

Private Sub mnuLog2_Click()
    BotVars.Logging = 0
    WriteINI "Other", "Logging", "0"
    AddChat RTBColors.InformationText, "Logging disabled."
    mnuLog1.Checked = False
    mnuLog0.Checked = False
    mnuLog2.Checked = True
    'mnuLog3.Checked = False
End Sub

'Private Sub mnuLog3_Click()
'    BotVars.Logging = 3
'    WriteINI "Other", "Logging", "3"
'    AddChat RTBColors.InformationText, "Logging disabled."
'    mnuLog1.Checked = False
'    mnuLog0.Checked = False
'    mnuLog2.Checked = False
'    mnuLog3.Checked = True
'End Sub

'Private Sub mnuMonitor_Click()
'    If Not MonitorExists Then
'        InitMonitor
'    End If
'
'    MonitorForm.Show
'End Sub

Private Sub mnuOpenBotFolder_Click()
    Shell "explorer.exe " & App.Path, vbNormalFocus
End Sub


Private Sub mnuPacketLog_Click()
    Dim f As Integer
    
    If mnuPacketLog.Checked Then
        ' turning this feature off
        AddChat RTBColors.SuccessText, "StealthBot packet traffic will no longer be logged."
        PacketLogFilePath = ""
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
    
    If txtPre.Enabled Then 'fix for topic 25290 -a
        If Dii Then txtPre.text = "/w *" Else txtPre.text = "/w "
        
        txtPre.text = txtPre.text & GetSelectedUser & " "
        
        cboSend.SetFocus
        cboSend.SelStart = Len(cboSend.text)
    End If
End Sub

Private Sub mnuPopAddToFList_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If Not (lvChannel.SelectedItem Is Nothing) Then
        AddQ "/f a " & _
            IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopClanWhois_Click()
    If Not (lvClanList.SelectedItem Is Nothing) Then
        AddQ "/whois " & lvClanList.SelectedItem.text, PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopDes_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN

    AddQ "/designate " & _
        IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser), PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopFLProfile_Click()
    If Not lvFriendList.SelectedItem Is Nothing Then
        RequestProfile CleanUsername(lvFriendList.SelectedItem.text)
    
        frmProfile.lblUsername.Caption = lvFriendList.SelectedItem.text
        frmProfile.Show
    End If
End Sub

Private Sub mnuPopFLWhois_Click()
    If Not (lvFriendList.SelectedItem Is Nothing) Then
        AddQ "/whois " & lvFriendList.SelectedItem.text, PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopSafelist_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim gAcc As udtGetAccessResponse
    Dim toSafe As String
    
    On Error Resume Next
    
    toSafe = GetSelectedUser
    
    gAcc.Access = 1000
    
    Call ProcessCommand(GetCurrentUsername, "/safeadd " & toSafe, True, False)
End Sub

Private Sub mnuPopShitlist_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim gAcc As udtGetAccessResponse
    Dim toBan As String
    
    On Error Resume Next
    
    toBan = GetSelectedUser
    
    gAcc.Access = 1000
    
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
    
    If cboSend.text <> vbNullString Then
        AddQ "/w " & _
            IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser) & Space(1) & _
                cboSend.text, PRIORITY.CONSOLE_MESSAGE
        
        cboSend.AddItem cboSend.text, 0
        cboSend.text = vbNullString
        cboSend.SetFocus
    End If
End Sub

Sub mnuClear_Click()
    rtbChat.text = vbNullString
    rtbChatLength = 0
    rtbWhispers.text = vbNullString
    AddChat RTBColors.InformationText, "Chat and whisper windows cleared."
    cboSend.SetFocus
End Sub

Private Sub mnuPopWhois_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/whois " & _
        IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser), PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuPopInvite_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim sPlayer As String
    
    If Not lvChannel.SelectedItem Is Nothing Then
        sPlayer = GetSelectedUser
    End If
    
    If LenB(sPlayer) > 0 Then
        If g_Clan.Self.Rank >= 3 Then
            InviteToClan (reverseUsername(sPlayer))
            AddChat RTBColors.InformationText, "[CLAN] Invitation sent to " & GetSelectedUser & ", awaiting reply."
        End If
    End If
End Sub

Private Sub mnuPopWebProfileWAR3_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim sPlayer As String
    
    sPlayer = CleanUsername(GetSelectedUser)

    Call GetW3LadderProfile(sPlayer, WAR3)
End Sub

Private Sub mnuPopWebProfileW3XP_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    Dim sPlayer As String
    
    sPlayer = CleanUsername(GetSelectedUser)
    
    Call GetW3LadderProfile(sPlayer, W3XP)
End Sub

Private Sub mnuClearedTxt_Click()
    Shell "notepad " & GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt", vbNormalFocus
End Sub


Private Sub mnuQC_Click(Index As Integer)
    If Len(QC(Index)) > 0 Then
        AddQ "/join " & QC(Index), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuRecordWindowPos_Click()
    RecordWindowPosition
End Sub

Private Sub mnuRepairCleanMail_Click()
    CleanUpMailFile
    frmChat.AddChat RTBColors.SuccessText, "Delivered and invalid pieces of mail have been removed from your mail.dat file."
End Sub

'Private Sub mnuRepairConfig_Click()
'    Dim i As Integer
'
'    If MsgBox("Are you sure?" & vbCrLf & _
'            "This action will remove any overrides or ""hacks"" that you have placed in your config.ini file, including custom version bytes." & vbCrLf & _
'            "It cannot be reversed.", vbYesNo, "Remove config.ini overrides") = vbYes Then
'        Const m = "Main"
'        Const o = "Other"
'
'        Dim Fields(16) As String
'
'        Fields(0) = "ForceDefaultLocaleID"
'        Fields(1) = "UppercasePassword"
'        Fields(2) = "DisableSBNews"
'        Fields(3) = "OwnerName"
'        Fields(4) = "PeonBanMsg"
'        Fields(8) = "D2VerByte"
'        Fields(9) = "SCVerByte"
'        Fields(10) = "W2VerByte"
'        Fields(11) = ""
'        Fields(12) = "SetBNLSServer"
'        Fields(13) = "0x51UnknownOverride"
'        Fields(14) = "DisableSBNews"
'        Fields(15) = "W3VerByte"
'        Fields(16) = "WinampPath"
'
'        For i = 0 To UBound(Fields)
'            If LenB(ReadCFG(m, Fields(i))) > 0 Then
'                If i < 16 Then
'                    WriteINI m, Fields(i), " "
'                Else
'                    WriteINI o, Fields(i), " "
'                End If
'            End If
'        Next i
'    End If
'End Sub

Private Sub mnuRepairDataFiles_Click()
    If MsgBox("Are you sure? This action will delete your mail.dat (Bot mail database) and commands.dat (custom command database) files.", vbYesNo, "Repair data files") = vbYes Then
        On Error Resume Next
        Kill GetFilePath("mail.dat")
        Kill GetFilePath("commands.dat")
        AddChat RTBColors.SuccessText, "The bot's DAT data files have been removed."
    End If
End Sub

Private Sub mnuRepairVerbytes_Click()
    WriteINI "Main", "W2VerByte", Hex(GetVerByte("NB2W", 1))
    WriteINI "Main", "W3VerByte", Hex(GetVerByte("3RAW", 1))
    WriteINI "Main", "W2VerByte", Hex(GetVerByte("RATS", 1))
    WriteINI "Main", "W2VerByte", Hex(GetVerByte("VD2D", 1))
    
    frmChat.AddChat RTBColors.SuccessText, "The version bytes stored in config.ini have been restored to their defaults."
End Sub

Private Sub mnuScripts_Click()
'do nothing
End Sub

Private Sub mnuToggleShowOutgoing_Click()
    mnuToggleShowOutgoing.Checked = (Not mnuToggleShowOutgoing.Checked)
    
    WriteINI "Main", "ShowOutgoingWhispers", IIf(mnuToggleShowOutgoing.Checked, "Y", "N")
End Sub

Private Sub mnuToggleWWUse_Click()
    mnuToggleWWUse.Checked = (Not mnuToggleWWUse.Checked)
    
    WriteINI "Main", "UseWWs", IIf(mnuToggleWWUse.Checked, "Y", "N")
    
    If Not mnuToggleWWUse.Checked Then
        DestroyAllWWs
    End If
End Sub

Private Sub mnuUpdateVerbytes_Click()
    Dim s As String, ary() As String
    Dim I As Integer
    
    Dim Keys(3) As String
    
    Keys(0) = "W2"
    Keys(1) = "SC"
    Keys(2) = "D2"
    Keys(3) = "W3"
    
    If Not INet.StillExecuting Then
        s = INet.OpenURL("http://www.stealthbot.net/verbytes/versionbytes.txt")
        
        If Len(s) = 11 Then
            'W2 SC D2 W3
            ary() = Split(s, " ")
            
            For I = 0 To 3
                WriteINI "Override", Keys(I) & "VerByte", ary(I)
            Next I
            
            AddChat RTBColors.SuccessText, "Your config.ini file has been loaded with current version bytes."
        Else
            AddChat RTBColors.ErrorMessageText, "Error retrieving version bytes from http://www.stealthbot.net! Please visit it for instructions."
        End If
    End If
End Sub

Private Sub mnuWhisperCleared_Click()
    Shell "notepad " & GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt", vbNormalFocus
End Sub

Private Sub mnuEditConfig_Click()
    Shell "notepad " & GetConfigFilePath(), vbNormalFocus
End Sub

Private Sub mnuEditDefinitions_Click()
    Shell "notepad " & GetFilePath("definitions.ini"), vbNormalFocus
End Sub

Private Sub mnuEditFilters_Click()
    Shell "notepad " & GetFilePath("filters.ini"), vbNormalFocus
End Sub

Private Sub mnuEditQuotes_Click()
    Shell "notepad " & GetFilePath("quotes.txt"), vbNormalFocus
End Sub

Private Sub mnuEditUsers_Click()
    Shell "notepad " & GetFilePath("users.txt"), vbNormalFocus
End Sub

Private Sub mnuFListRefresh_Click()
    lvFriendList.ListItems.Clear
    Call FriendListHandler.RequestFriendsList(PBuffer)
End Sub

Sub mnuReloadScript_Click()

    'ReInitScriptControl SControl

    'Exit Sub

    Dim I As Integer, lMenu As Long
    'Dim Message As String
    
    On Error GoTo mnuReloadScript_Click_Error
    
    SControl.Run "Event_Close"
    SControl.Reset
        
MRS_Continue:
    ' Clear callback list
    dctCallbacks.RemoveAll
    
    ' Reset script custom menu
    ' First, go through the collection and delete each submenu
    lMenu = GetMenu(frmChat.hWnd)
    
    ' Remove each of the subitems for all of the script menus
    While colDynamicMenus.Count > 0
        I = GetMenuItemCount(colDynamicMenus.Item(1))
    
        For I = 0 To I
            DeleteMenuItem ScriptMenu_ParentID, colDynamicMenus.Item(1), 0
        Next I
        
        colDynamicMenus.Remove 1
    Wend
    
    ' Remove each of the script menus
    I = GetMenuItemCount(ScriptMenu_ParentID)
    
    For I = 0 To I
        DeleteMenuItem lMenu, ScriptMenu_ParentID, 0
    Next I

    ' Finally, remove the script menu itself and reset its ID
    DeleteMenuItem lMenu, ScriptMenu_ParentID, 5
    ScriptMenu_ParentID = 0

    LoadPluginSystem SControl

    ReInitScriptControl SControl
    
    DrawMenuBar frmChat.hWnd

MRS_Exit:
    On Error GoTo 0
    Exit Sub

mnuReloadScript_Click_Error: ' No code is present
    'Debug.Print Err.Number & ": " & Err.Description
    If Err.Number = 438 Then
        Resume MRS_Continue
    Else
        Debug.Print "Unhandled error in mnuReloadScript_Click()"
        Debug.Print Err.Number & ": " & Err.description
        Resume MRS_Exit
    End If
End Sub

'Private Sub mnuSetTop_Click()
'    mnuLog0.Checked = False
'    mnuLog1.Checked = False
'    mnuLog2.Checked = False
'
'    Select Case BotVars.Logging
'        Case 0: mnuLog0.Checked = True
'        Case 1: mnuLog1.Checked = True
'        Case 2: mnuLog2.Checked = True
'    End Select
'End Sub

Private Sub mnuStatsBW_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/stats " & CleanUsername(GetSelectedUser) & " SEXP", PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuStatsFT_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/stats " & CleanUsername(GetSelectedUser) & " W3XP", PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuStatsSC_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/stats " & CleanUsername(GetSelectedUser) & " STAR", PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuStatsW2_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/stats " & CleanUsername(GetSelectedUser) & " W2BN", PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuStatsW3_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    AddQ "/stats " & CleanUsername(GetSelectedUser) & " WAR3", PRIORITY.CONSOLE_MESSAGE
End Sub

Private Sub mnuTerms_Click()
    Shell "notepad " & App.Path & "\eula.txt", vbNormalFocus
End Sub

Private Sub mnuFilters_Click()
    frmFilters.Show
End Sub

Private Sub mnuPopPLookup_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    RequestProfile CleanUsername(GetSelectedUser)
    
    frmProfile.lblUsername.Caption = GetSelectedUser
    frmProfile.Show
End Sub

Private Sub mnuPopCopy_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Clipboard.Clear
    
    Clipboard.SetText GetSelectedUser
End Sub

Private Sub mnuProfile_Click()
    frmWriteProfile.Show
End Sub

Private Sub mnuQCEdit_Click()
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
        WriteINI "Main", "UTF8", "N"
        AddChat RTBColors.ConsoleText, "Messages will no longer be UTF-8-decoded."
    Else
        mnuUTF8.Checked = True
        WriteINI "Main", "UTF8", "Y"
        AddChat RTBColors.ConsoleText, "Messages will now be UTF-8-decoded."
        
    End If
End Sub

Private Sub mnuWindow_Click()
    mnuLog0.Checked = False
    mnuLog1.Checked = False
    mnuLog2.Checked = False
    'mnuLog3.Checked = False
    
    Select Case BotVars.Logging
        Case 2: mnuLog0.Checked = True
        Case 1: mnuLog1.Checked = True
        Case 0: mnuLog2.Checked = True
        'Case 3: mnuLog3.Checked = True
    End Select
End Sub

Private Sub rtbChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And ((KeyCode = vbKeyL) Or (KeyCode = vbKeyE) Or (KeyCode = vbKeyR)) Then
        'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
        Select Case KeyCode
            Case vbKeyL
                Call mnuLock_Click
            Case vbKeyR
                Call mnuReloadScript_Click
        End Select
        
        'Disable Ctrl+L, Ctrl+E, and Ctrl+R
        KeyCode = 0
    End If
End Sub

Private Sub rtbChat_KeyPress(KeyAscii As Integer)
    ' ...
    If (KeyAscii < 32) Then
        Exit Sub
    End If

    ' ...
    cboSend.SetFocus
    
    ' ...
    cboSend.SelText = Chr$(KeyAscii)
End Sub

Private Sub rtbWhispers_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift = vbCtrlMask) And ((KeyCode = vbKeyL) Or (KeyCode = vbKeyE) Or (KeyCode = vbKeyR)) Then
        'Call Ctrl+L and Ctrl+R keyboard shortcuts as they code to automatically handle them will be canceled out below
        Select Case KeyCode
            Case vbKeyL
                Call mnuLock_Click
            Case vbKeyR
                Call mnuReloadScript_Click
        End Select
        
        'Disable Ctrl+L, Ctrl+E, and Ctrl+R
        KeyCode = 0
    End If
End Sub

Private Sub rtbWhispers_KeyPress(KeyAscii As Integer)
    ' ...
    If (KeyAscii < 32) Then
        Exit Sub
    End If

    ' ...
    cboSend.SetFocus
    
    ' ...
    cboSend.SelText = Chr$(KeyAscii)
End Sub

Private Sub rtbChat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 And Len(rtbChat.SelText) > 0 Then
        If Not BotVars.NoRTBAutomaticCopy Then
            Clipboard.Clear
            Clipboard.SetText rtbChat.SelText, vbCFText
        End If
    End If
End Sub

Private Sub rtbWhispers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 And Len(rtbWhispers.SelText) > 0 Then
        Clipboard.Clear
        Clipboard.SetText rtbWhispers.SelText, vbCFText
    End If
End Sub

Private Sub mnuToggleFilters_Click()
    If Filters Then
        Filters = False
        AddChat RTBColors.InformationText, "Chat filtering disabled."
        WriteINI "Other", "Filters", "N"
    Else
        Filters = True
        AddChat RTBColors.InformationText, "Chat filtering enabled."
        WriteINI "Other", "Filters", "Y"
    End If
End Sub

Private Sub mnuUserlistWhois_Click()
    On Error Resume Next
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    Dim temp As udtGetAccessResponse
    Dim s As String
    
    s = GetSelectedUser
    
    temp = GetAccess(s)
    
    With RTBColors
        If temp.Access > -1 Then
            If temp.Access > 0 Then
                If temp.Flags <> vbNullString Then
                    AddChat .ConsoleText, "Found user " & s & ", with access " & temp.Access & " and flags " & temp.Flags & "."
                Else
                    AddChat .ConsoleText, "Found user " & s & ", with access " & temp.Access & "."
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

Private Sub mnuConnect_Click()
    GErrorHandler.Reset
    Call DoConnect
End Sub

Private Sub mnuPopKick_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If MyFlags = 2 Or MyFlags = 18 Then
        AddQ "/kick " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuPopBan_Click()
    If Not PopupMenuUserCheck Then Exit Sub 'Check user selected is the same one that was right-clicked on. - FrOzeN
    
    If MyFlags = 2 Or MyFlags = 18 Then
        AddQ "/ban " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(GetSelectedUser), PRIORITY.CONSOLE_MESSAGE
    End If
End Sub

Private Sub mnuTrayExit_click()
    Dim result As VbMsgBoxResult

    result = MsgBox("Are you sure you want to quit?", vbYesNo, "StealthBot")
    
    If (result = vbYes) Then
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
    Dim Key As String, L As Long
    Key = GetProductKey()
    
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
    If JoinMessagesOff = False Then
        AddChat RTBColors.InformationText, "Join/Leave messages disabled."
        JoinMessagesOff = True
        WriteINI "Other", "JoinLeaves", "N"
    Else
        AddChat RTBColors.InformationText, "Join/Leave messages enabled."
        JoinMessagesOff = False
        WriteINI "Other", "JoinLeaves", "Y"
        If ForcedJoinsOn = 2 Then ForcedJoinsOn = 1
    End If
End Sub

Private Sub mnuUsers_Click()
    frmDBManager.Show
End Sub

Private Sub txtPre_GotFocus()
    Call cboSend_GotFocus
End Sub

Private Sub txtPost_GotFocus()
    Call cboSend_GotFocus
End Sub

Private Sub cboSend_GotFocus()
    ' ...
    On Error Resume Next

    Dim I As Integer ' ...

    cboSend.SelStart = cboSendSelStart
    cboSend.SelLength = cboSendSelLength

    If (BotVars.NoAutocompletion = False) Then
        ' ..
        For I = 0 To (Controls.Count - 1)
            ' ...
            If (TypeOf Controls(I) Is ListView) Or _
                    (TypeOf Controls(I) Is SSTab) Or _
                        (TypeOf Controls(I) Is RichTextBox) Or _
                            (TypeOf Controls(I) Is TextBox) Then
                            
                ' ...
                If (Controls(I).TabStop = False) Then
                    Controls(I).Tag = "False"
                End If

                ' ...
                Controls(I).TabStop = False
            End If
        Next I
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
    ' ...
    On Error Resume Next

    Dim I As Integer ' ...
    
    If (BotVars.NoAutocompletion = False) Then
        ' ...
        For I = 0 To (Controls.Count - 1)
            ' ...
            If (TypeOf Controls(I) Is ListView) Or _
                    (TypeOf Controls(I) Is TabStrip) Or _
                        (TypeOf Controls(I) Is RichTextBox) Or _
                            (TypeOf Controls(I) Is TextBox) Then
                            
                ' ...
                If (Controls(I).Tag <> "False") Then
                    Controls(I).TabStop = True
                End If
            End If
        Next I
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
    
    Select Case (KeyCode)
        Case KEY_SPACE
            With cboSend
                If (LenB(LastWhisper) > 0) Then
                    If (Len(.text) >= 3) Then
                        If StrComp(Left$(.text, 3), "/r ", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.text)
                            .SelText = _
                                "/w " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(LastWhisper) & " "
                            .SelStart = Len(.text)
                        End If
                    End If
                    
                    If (Len(.text) >= 7) Then
                        If StrComp(Left$(.text, 7), "/reply ", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.text)
                            .SelText = _
                                "/w " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(LastWhisper) & " "
                            .SelStart = Len(.text)
                        End If
                    End If
                End If
                
                If (LenB(LastWhisperTo) > 0) Then
                    If (Len(.text) >= 4) Then
                        If StrComp(Left$(.text, 4), "/rw ", vbTextCompare) = 0 Then
                            .SelStart = 0
                            .SelLength = Len(.text)
                            
                            If StrComp(LastWhisperTo, "%f%") = 0 Then
                                .SelText = "/f m "
                            Else
                                .SelText = _
                                    "/w " & IIf(Dii And usingGameConventions, "*", "") & CleanUsername(LastWhisperTo) & " "
                            End If
                            
                            .SelStart = Len(.text)
                        End If
                    End If
                End If
            End With
    End Select
End Sub

Private Sub cboSend_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERROR_HANDLER
    
    Static strbuf        As String ' ...
    Static user          As String ' ...
    Static spaceIndex(2) As Long   ' ...

    Dim temp As udtGetAccessResponse
    
    Dim I As Long
    Dim L As Long
    Dim n As Integer
    Dim C As Integer ',oldSelStart As Integer
    Dim X() As String
    Dim m As String
    Dim s As String ',sClosest As String
    Dim Vetoed As Boolean
    
    Const S_SHIFT = 1
    Const S_CTRL = 2
    Const S_ALT = 4
    Const S_CTRLALT = 6
    Const S_CTRLSHIFT = 3
    Const S_CTRLSHIFTALT = 7
    Const L_ARROW = 38
    Const R_ARROW = 39

    Const K_END = 35
    
    'AddChat vbRed, "KeyCode: " & KeyCode
    'AddChat vbRed, "Shift: " & Shift


    L = cboSend.SelStart

    With lvChannel

        If (Not (.SelectedItem Is Nothing)) Then
            I = .SelectedItem.Index
        End If
        
        'MsgBox KeyCode

        Select Case (KeyCode)
            Case KEY_PGDN 'ALT + PAGEDOWN
                If Shift = S_ALT Then
                    If I < .ListItems.Count Then
                        .ListItems.Item(I + 1).Selected = True
                        .ListItems.Item(I).Ghosted = False
                        .ListItems.Item(I + 1).Ghosted = True
                    End If

                    cboSend.SetFocus
                    cboSend.SelStart = L
                    Exit Sub
                End If

            Case KEY_PGUP 'ALT + PAGEUP
                If Shift = S_ALT Then
                    If I > 1 Then
                        .ListItems.Item(I - 1).Selected = True
                        .ListItems.Item(I).Ghosted = False
                        .ListItems.Item(I - 1).Ghosted = True
                    End If

                    cboSend.SetFocus
                    cboSend.SelStart = L
                    Exit Sub
                End If

            Case KEY_ALTN, KEY_INSERT 'ALT + N or ALT + INSERT
                If (Shift = S_ALT) Then
                    's = NameWithoutRealm(GetSelectedUser)
                    'c = .SelectedItem.Index
                    'Unfinished business - suggestion from Engel
                                            
                    If (Not (.SelectedItem Is Nothing)) Then
                        cboSend.SelText = .SelectedItem.text
                        cboSend.SelStart = cboSend.SelStart + Len(.SelectedItem.text)
                    End If
                End If

            Case KEY_HOME 'ALT+HOME
                If Shift = S_ALT Then
                    If (I > 0) Then
                        .ListItems.Item(1).Selected = True
                        
                        For C = 1 To .ListItems.Count
                            .ListItems.Item(C).Ghosted = False
                        Next C
                        
                        .ListItems.Item(1).Ghosted = True
    
                        cboSend.SetFocus
                        cboSend.SelStart = L
                    Else
                        If .ListItems.Count > 0 Then
                            .ListItems(1).Selected = True
                            .ListItems(1).Ghosted = True
                            cboSend.SetFocus
                            cboSend.SelStart = L
                        End If
                    End If
                End If

            Case KEY_END 'ALT+END
                If Shift = S_ALT Then
                    If (.ListItems.Count > 0) Then
                        .ListItems.Item(.ListItems.Count).Selected = True
                        .ListItems.Item(I).Ghosted = False
                        .ListItems.Item(.ListItems.Count).Ghosted = True
    
                        cboSend.SetFocus
                        cboSend.SelLength = L
                    End If
                End If
                
            Case KEY_V 'PASTE
                If (Shift = S_CTRL) Then
                    On Error Resume Next
                    
                    If (InStr(1, Clipboard.GetText, Chr(13), vbTextCompare) <> 0) Then
                        X() = Split(Clipboard.GetText, Chr(10))
                        
                        If UBound(X) > 0 Then
                            For n = LBound(X) To UBound(X)
                                X(n) = Replace(X(n), Chr(13), vbNullString)
                                
                                If (X(n) <> vbNullString) Then
                                    If (n <> LBound(X)) Then
                                        AddQ txtPre.text & X(n) & txtPost.text, PRIORITY.CONSOLE_MESSAGE
                                        
                                        cboSend.AddItem txtPre.text & X(n) & txtPost.text, 0
                                    Else
                                        AddQ txtPre.text & cboSend.text & X(n) & txtPost.text, _
                                            PRIORITY.CONSOLE_MESSAGE
                                        
                                        cboSend.AddItem txtPre.text & cboSend.text & X(n) & txtPost.text, 0
                                    End If
                                End If
                            Next n
                            
                            cboSend.text = vbNullString
                            
                            MultiLinePaste = True
                        End If
                    End If
                End If
                
            Case KEY_A
                If (Shift = S_CTRL) Then
                    C = ListviewTabs.TabIndex
                    ListviewTabs.TabIndex = 0
                    Call ListviewTabs_Click(C)
                End If
                
            Case KEY_S
                If (Shift = S_CTRL) Then
                    C = ListviewTabs.TabIndex
                    ListviewTabs.TabIndex = 1
                    Call ListviewTabs_Click(C)
                End If
                
            Case KEY_D
                If (Shift = S_CTRL) Then
                    C = ListviewTabs.TabIndex
                    
                    ListviewTabs.TabIndex = 2
                    
                    Call ListviewTabs_Click(C)
                End If
                
            Case KEY_B
                If (Shift = S_CTRL) Then
                    With cboSend
                        .SelText = "ÿcb"
                        .SelLength = 0
                        .SelStart = .SelStart + 3
                    End With
                End If
                
            'Case KEY_J
            '    If (Shift = S_CTRL) Then
            '        Call mnuToggle_Click
            '    End If
                
            Case KEY_U
                If (Shift = S_CTRL) Then
                    With cboSend
                        .SelText = "ÿcu"
                        .SelLength = 0
                        .SelStart = .SelStart + 3
                    End With
                End If
                
            Case KEY_I
                If (Shift = S_CTRL) Then
                    With cboSend
                        .SelText = "ÿci"
                        .SelLength = 0
                        .SelStart = .SelStart + 3
                    End With
                End If
                
            Case KEY_DELETE
                strbuf = vbNullString
                
            Case vbKeyTab
                Dim prevStart As Long   ' ...
                Dim tmpStr    As String ' ...
                Dim res       As String ' ...
            
                If (Shift) Then
                    Call cboSend_LostFocus
                    
                    If (txtPre.Visible = True) Then
                        Call txtPre.SetFocus
                    Else
                        Call ListviewTabs.SetFocus
                    End If
                Else
                    With cboSend
                        ' ...
                        If (user = vbNullString) Then
                            ' ...
                            strbuf = .text
                            
                            ' ...
                            If (.SelStart > 0) Then
                                ' grab space before cursor
                                spaceIndex(0) = _
                                    InStrRev(strbuf, Space$(1), .SelStart, vbBinaryCompare)
                                
                                ' grab space after cursor
                                spaceIndex(1) = _
                                    InStr(.SelStart, strbuf, Space$(1), vbBinaryCompare)
                            End If
                            
                            ' ...
                            If (spaceIndex(0) > 0) Then
                                user = Mid$(strbuf, spaceIndex(0), _
                                    IIf(spaceIndex(1), spaceIndex(1) - spaceIndex(0), Len(.text)))
                            Else
                                user = Mid$(strbuf, 1, IIf(spaceIndex(1), spaceIndex(1) - 1, Len(.text)))
                            End If
                            
                            ' ...
                            MatchIndex = 1
                        Else
                            MatchIndex = MatchIndex + 1
                        End If
                        
                        ' ...
                        If (user <> vbNullString) Then
                            res = MatchClosest(user, IIf(MatchIndex, MatchIndex, 1))

                            ' final check
                            If (res <> vbNullString) Then
                                Dim SelStart As Long   ' ...
                                Dim tmp      As String ' ...

                                ' ...
                                If (spaceIndex(0) > 0) Then
                                    tmp = Left$(strbuf, spaceIndex(0))
                                End If
                                
                                ' ...
                                tmp = tmp & res
                                
                                ' ...
                                SelStart = Len(tmp)
                                
                                ' ...
                                If (spaceIndex(1) > 0) Then
                                    tmp = tmp & Mid$(strbuf, spaceIndex(1))
                                End If
                                
                                ' ...
                                .text = tmp
                                
                                ' ...
                                .SelStart = SelStart
                            End If
                        End If
                    End With
                End If
                
            Case KEY_ENTER
                'n = UsernameToIndex(CurrentUsername)
                
                'Debug.Print n
                
                'If (n > 0) Then
                '    With colUsersInChannel
                '        .Item(n).Acts
                '    End With
                'End If
            
                Select Case (Shift)
                    Case S_CTRL 'CTRL+ENTER - rewhisper
                        If LenB(cboSend.text) > 0 Then
                            AddQ "/w " & LastWhisperTo & Space(1) & cboSend.text, _
                                PRIORITY.CONSOLE_MESSAGE
                                
                            cboSend.text = vbNullString
                        End If
                        
                    Case S_CTRLSHIFT 'CTRL+SHIFT+ENTER - reply
                        If LenB(cboSend.text) > 0 Then
                            AddQ "/w " & LastWhisper & Space(1) & cboSend.text, _
                                PRIORITY.CONSOLE_MESSAGE
                            cboSend.text = vbNullString
                        End If
                
                    Case Else 'normal ENTER - old rules apply
                        'If (Highlighted) Then
                        '    cboSend.SelText = vbNullString
                        '
                        '    Highlighted = False
                        'End If
                    
                        If (LenB(cboSend.text) > 0) Then
                            On Error Resume Next
                            
                            'If (g_Channel.IsSilent) And Not mnuDisableVoidView.Checked Then
                            '    BNCSBuffer.VoidTrimBuffer
                            'End If
                            
                            SetVeto False
                            
                            SControl.Run "Event_PressedEnter", cboSend.text
                            
                            Vetoed = GetVeto
                            
                            If (Not (Vetoed)) Then
                                s = txtPre.text & cboSend.text & txtPost.text
                            
                                If (Left$(s, 6) = "/tell ") Then
                                    s = "/w " & Mid$(s, 7)
                                    
                                    Call AddQ(OutFilterMsg(s), PRIORITY.CONSOLE_MESSAGE)
                                    
                                    GoTo theEnd
                                    
                                'ElseIf (LCase$(s) = "/fl" And MDebug("debug")) Then
                                '    For n = 1 To g_Friends.Count
                                '        AddChat vbMagenta, g_Friends.Item(n).Name & _
                                '            " - " & g_Friends.Item(n).game
                                '    Next n
                                
                                'ElseIf (LCase$(s) = "/accountinfo") Then
                                '    RequestSystemKeys
                                '
                                '    GoTo theEnd
                                    
                                'ElseIf (LCase$(s) = "/lvchandims") Then
                                '    Dim j As Integer ' ...
                                '
                                '    AddChat vbRed, "lvChannel:"
                                '    AddChat vbRed, " Height: " & lvChannel.Height
                                '    AddChat vbRed, " Width: " & lvChannel.Width
                                '
                                '    For j = 1 To lvChannel.ColumnHeaders.Count
                                '        AddChat vbRed, " Column " & j & ":"
                                '        AddChat vbRed, "  Width:" & lvChannel.ColumnHeaders(j).Width
                                '    Next j
                                '
                                'ElseIf (LCase$(s) = "/cls") Then
                                '    Call mnuClear_Click
                                '
                                '    GoTo theEnd
                                '
                                'ElseIf (LCase$(s) = "/ds_list") Then
                                '    Call ds.List
                                '
                                '    GoTo theEnd
                                '
                                'ElseIf (Left$(LCase$(s), 7) = "/setcl ") Then
                                '    CommandLine = Mid$(s, 8)
                                '    frmChat.AddChat RTBColors.SuccessText, _
                                '            "The command line for this instance has been changed."
                                '
                                '    GoTo theEnd
                                '
                                'ElseIf ((s = "/force") And (MDebug("debug"))) Then
                                '    MyFlags = 2
                                '    SharedScriptSupport.BotFlags = MyFlags
                                '    AddChat RTBColors.ConsoleText, "Flags forced to 2."
                                '
                                'ElseIf ((s = "/flags") And (MDebug("debug"))) Then
                                '    For n = 1 To g_Channel.Users.Count
                                '        With g_Channel.Users(n)
                                '            AddChat RTBColors.ConsoleText, .Name & Space$(4) & .Flags
                                '        End With
                                '    Next n
                                '
                                '    n = 0
                                '
                                '    GoTo theEnd
                                '
                                'ElseIf LCase(Left$(s, 7)) = "/watch " Then
                                '    WatchUser = LCase(Right(s, Len(s) - 7))
                                '    AddChat RTBColors.ConsoleText, "Watching " & Right(s, Len(s) - 7)
                                '
                                '    GoTo theEnd
                                '
                                'ElseIf LCase$(s) = "/watchoff" Then
                                '    WatchUser = vbNullString
                                '    AddChat RTBColors.ConsoleText, "Watch off."
                                '    GoTo theEnd
                                'ElseIf (LCase(Left$(s, 7)) = "/reply ") Then
                                '
                                '   m = Right(s, (Len(s) - 7))
                                'ElseIf (LCase(Left$(s, 7)) = "/reply ") Then
                                '    m = Right(s, (Len(s) - 7))
                                '
                                '    AddQ "/w " & LastWhisper & Space(1) & OutFilterMsg(m), _
                                '        PRIORITY.CONSOLE_MESSAGE
                                    
                                'ElseIf (LCase(Left$(s, 9)) = "/profile ") Then
                                '    If (sckBNet.State = 7) Then
                                '        RequestProfile Right(s, Len(s) - 9)
                                '    End If
                                '
                                '    frmProfile.lblUsername.Caption = Right(s, Len(s) - 9)
                                '    frmProfile.Show
                                
                                ElseIf (LCase(Left$(s, 1)) = "/") Then
                                    Dim commandResult As Boolean ' ..
                                
                                    'If ((Left$(s, 3) = "/w ") Or (Left$(s, 3) = "/m ")) Then
                                    '    If (Dii) Then
                                    '        If (StrComp(Mid$(s, 4, 1), "*") <> 0) Then
                                    '            s = Mid$(s, 1, 3) & "*" & _
                                    '                Mid$(s, 4)
                                    '        End If
                                    '    End If
                                    'End If
                                    
                                    temp.Access = 201
                                    temp.Flags = "A"
                                    
                                    m = OutFilterMsg(s)
                                    
                                    commandResult = ProcessCommand(GetCurrentUsername, m, _
                                        True, False)
                                Else
                                    Call AddQ(OutFilterMsg(s), PRIORITY.CONSOLE_MESSAGE)
                                End If
                                
                                'Ignore rest of code as the bot is closing
                                If BotIsClosing Then
                                    Exit Sub
                                End If
                                
                            End If
theEnd:
                            cboSend.AddItem cboSend.text, 0
                            
                            cboSend.text = vbNullString
                            
                            If Me.WindowState <> vbMinimized Then
                                cboSend.SetFocus
                            End If
                        End If
                    'case...
                End Select
                
                '########## end ENTER cases
            
            
        End Select
    End With
    
    ' ...
    If (KeyCode <> vbKeyTab) Then
        ' ...
        user = vbNullString
        
        ' ...
        strbuf = vbNullString
        
        ' ...
        spaceIndex(0) = 0
        spaceIndex(1) = 0
    End If

    Exit Sub

ERROR_HANDLER:
    AddChat RTBColors.ErrorMessageText, "Error " & Err.Number & " (" & Err.description & ") " & _
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
            
            'If ((OKToDoAutocompletion(.text, KeyAscii)) And _
            '    (KeyAscii <> 8)) Then
            '
            '    If (Highlighted) Then
            '        .SelText = ""
            '        Highlighted = False
            '    End If
            '
            '    If (.SelStart = Len(.text)) Then
            '        If (MatchIndex > 0) Then
            '            sClosest = MatchClosest(.text & Chr(KeyAscii), _
            '                MatchIndex)
            '        End If
            '
            '        If (Len(sClosest) = 0) Then
            '            sClosest = MatchClosest(.text & Chr(KeyAscii), 1)
            '        End If
            '
            '        oldSelStart = Len(.text) + 1 'text is "b" = 1 = 2
            '
            '        If (LenB(sClosest) > 0) Then
            '            .SelStart = oldSelStart
            '            .SelLength = 0
            '            .SelText = Chr(KeyAscii) & Mid$(sClosest, (oldSelStart + 1) - InStr(.text, " "))
            '            .SelStart = oldSelStart
            '            .SelLength = Len(.text)
            '
            '            KeyAscii = 0
            '            Highlighted = True
            '        End If
            '    End If
            'Else
            '    Highlighted = False
            'End If
        End If
    End With
    
    If Len(cboSend.text) > 223 Then
        cboSend.ForeColor = vbRed
    Else
        cboSend.ForeColor = vbWhite
    End If
End Sub

Private Sub quLower_Timer()
    On Error GoTo ERROR_HANDLER

    If QueueLoad > 0 Then QueueLoad = QueueLoad - 1
    If QueueMaster > 0 Then QueueMaster = QueueMaster - 2
    If QueueMaster < 0 Then QueueMaster = 0
    Dim gA As udtGetAccessResponse
    
    If Unsquelching Then Unsquelching = False
    
    Exit Sub
    
ERROR_HANDLER:

    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in quLower_Timer()."

    Exit Sub
End Sub


Private Sub QueueTimer_Timer()
    On Error GoTo ERROR_HANDLER

    Static delay As Integer
    Static Count As Integer

    Dim Message  As String
    Dim Tag      As String
    Dim Sent     As Byte
    Dim I        As Integer
    Dim override As Integer
    Dim pri      As Integer
    Dim id       As Integer
    
    If ((g_Queue.Count) And (g_Online)) Then
        With g_Queue.Peek
            Message = .Message
            Tag = .Tag
            pri = .PRIORITY
            id = .id
        End With
        
        ' ...
        'If (pri = PRIORITY.CHANNEL_MODERATION_MESSAGE) Then
        '    ' ...
        '    If (Left$(Message, 5) = "/ban ") Then
        '        Dim user As String ' ...
        '
        '        ' ...
        '        If (Len(Message) >= 6) Then
        '            user = Mid$(Message, 6)
        '
        '            If (InStr(user, " ") <> 0) Then
        '                user = Mid$(user, 1, InStr(user, " ") - 1)
        '            End If
        '        End If
        '
        '        ' ...
        '        If ((g_Channel.IsOnRecentBanList(user)) Or _
        '                (g_Channel.IsOnRecentBanList(StripRealm(user)))) Then
        '
        '            ' ...
        '            'AddChat vbRed, "FOUND ONE!"
        '
        '            ' ...
        '            Call g_Queue.Pop
        '
        '            ' ...
        '            Call QueueTimer_Timer
        '
        '            ' ...
        '            Exit Sub
        '        End If
        '    End If
        'End If
        
        'frmChat.AddChat vbYellow, Message
        
        If (StrComp(Message, "%%%%%blankqueuemessage%%%%%", vbBinaryCompare) = 0) Then
            '// This is a dummy queue message faking a 70-character queue entry
            QueueLoad = (QueueLoad + 1)
            QueueMaster = (QueueMaster + 3)
            
            ' ...
            Call g_Queue.Pop
        Else
            If ((StrComp(Left$(Message, 11), "/unsquelch ", vbTextCompare) = 0) Or _
                (StrComp(Left$(Message, 10), "/unignore ", vbTextCompare) = 0)) Then
                
                ' ...
                Unsquelching = True
            End If

            If ((QueueLoad < 3) And (QueueMaster < 16)) Then
                If (Len(Message) <= 70) Then
                    QueueLoad = (QueueLoad + 1)
                    QueueMaster = (QueueMaster + 3)
                ElseIf (Len(Message) <= 130) Then
                    QueueLoad = (QueueLoad + 2)
                    QueueMaster = (QueueMaster + 5)
                ElseIf (Len(Message) <= 170) Then
                    QueueLoad = (QueueLoad + 3)
                    QueueMaster = (QueueMaster + 7)
                Else
                    QueueLoad = (QueueLoad + 4)
                    QueueMaster = (QueueMaster + 9)
                End If
                
                Sent = 1
                
                'frmChat.AddChat vbRed, Message
                
                Call bnetSend(Message, Tag, id)
            End If
        End If
        
        If ((QueueMaster >= 15) And ((QueueTimer.Interval - delay) <> 2400)) Then
            QueueTimer.Interval = 2400 ' 2400
        ElseIf ((QueueMaster < 15) And ((QueueTimer.Interval - delay) = 2400)) Then
            QueueTimer.Interval = 1175 ' 1175
        Else
            QueueTimer.Interval = 1175 ' 1175
        End If
        
        If (Sent = 1) Then
            ' ...
            Call g_Queue.Pop
            
            ' are we issuing a ban or kick command?
            If (g_Queue.Peek.PRIORITY = PRIORITY.CHANNEL_MODERATION_MESSAGE) Then
                ' ...
                delay = BanDelay()
                
                ' ...
                QueueTimer.Interval = (QueueTimer.Interval + delay)
            End If
        End If

    End If
    
    Exit Sub
    
ERROR_HANDLER:

    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in QueueTimer_Timer()."

    Exit Sub

End Sub


Public Sub SControl_Error()
    AddChat RTBColors.ErrorMessageText, "Scripting runtime error " & Chr(39) & SControl.Error.Number & Chr(39) & ": (line " & SControl.Error.line & "; column " & SControl.Error.Column & ")"
    AddChat RTBColors.ErrorMessageText, SControl.Error.description & "."
    AddChat RTBColors.ErrorMessageText, "Offending line: >> " & SControl.Error.text
    
    SControl.Error.Clear
    INet.Cancel
End Sub

Private Sub sckBNet_Close()
    sckBNet.Close
    If sckBNLS.State <> 0 Then sckBNLS.Close
    
    'If it's locating another BNLS then don't message the user about the disconnection to Battle.net
    If LocatingAltBNLS Then
        LocatingAltBNLS = False
    Else
        Call Event_BNetDisconnected
    End If
    
    g_Connected = False
End Sub

Private Sub sckBNet_Connect()
    On Error Resume Next
    
    Call Event_BNetConnected
    
    If MDebug("all") Then
        AddChat COLOR_BLUE, "BNET CONNECT"
    End If
        
    If (Not (BotVars.UseProxy)) Then
        InitBNetConnection
    Else
        LogonToProxy sckBNet, BotVars.Server, 6112, BotVars.ProxyIsSocks5
    End If
End Sub

Sub InitBNetConnection()
    g_Connected = True
    
    'sckBNet.SendData ChrW(1)
    Call Send(sckBNet.SocketHandle, ChrW(1), 1, 0)
    
    If BotVars.BNLS Then
        NLogin.Send_0x10 BotVars.Product
    Else
        Call Send0x50
    End If
End Sub

Private Sub sckBNet_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Event_BNetError(Number, description)
End Sub

Private Sub sckMCP_Close()
    AddChat RTBColors.ErrorMessageText, "[REALM] Connection closed."
    RealmError = True
    Call DoDisconnect
End Sub

Private Sub sckMCP_Connect()
    On Error Resume Next
    
    If MDebug("all") Then
        AddChat COLOR_BLUE, "MCP CONNECT"
    End If
    
    AddChat RTBColors.SuccessText, "[REALM] Connection established!"
    
    'sckMCP.SendData ChrW(1)
    Call Send(sckMCP.SocketHandle, ChrW(1), 1, 0)
    
    frmRealm.MCPHandler.SendStartup
End Sub

Private Sub sckMCP_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ERROR_HANDLER

    Dim Data As String
    
    sckMCP.GetData Data, vbString
    frmRealm.MCPHandler.ParseMCPPacket Data
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in sckMCP_DataArrival()."

    Exit Sub
End Sub

Private Sub sckMCP_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not g_Online Then
        ' This message is ignored if we've been online for awhile.
        AddChat RTBColors.ErrorMessageText, "[REALM] Server error " & Number & ": " & description
        RealmError = True
        Unload frmRealm
    End If
End Sub


'// Written by Swent. Executes plugin timer subs
Private Sub scTimer_Timer()
    On Error GoTo ERROR_HANDLER

    If modScripting.boolOverride Then
        On Error Resume Next
        SControl.Run "scTimer_Timer"
        Exit Sub
    End If

    '// Are plugins enabled?
    If Not CBool(SharedScriptSupport.GetSetting("ps", "enabled")) Then Exit Sub
    
    Dim strKeys() As String, strKey() As String, I As Integer
    
    On Error Resume Next
    SControl.Error.Clear
    strKeys = Split(modScripting.GetPTKeys)

    '// Execute all existing plugin timer subs at the appropriate intervals
    For I = 0 To modScripting.dictTimerEnabled.Count - 1
        strKey = Split(strKeys(I), ":")
    
        '// Is this timer enabled?
        If modScripting.GetPTEnabled(strKey(0), strKey(1)) Then
    
            '// Is the plugin that this timer belongs to enabled?
            If CBool(SharedScriptSupport.GetSetting(strKey(0), "enabled")) Then
    
                '// Has this timer reached the end of its interval countdown?
                If modScripting.GetPTLeft(strKey(0), strKey(1)) = 1 Then
    
                    '// Execute this timer sub
                    frmChat.SControl.Run strKey(0) & "_" & strKey(1) & "_Timer"
    
                    '// Handle errors
                    If SControl.Error.Number <> 0 Then
                        AddChat vbYellow, "The """ & strKey(1) & """ timer in your """ & strKey(0) & """ plugin has been disabled due to an error."
                        modScripting.SetPTEnabled strKey(0), strKey(1), False
                        SControl.Error.Clear
                    End If
    
                    '// Reset this timer's countdown
                    modScripting.SetPTCount strKey(0), strKey(1), modScripting.GetPTInterval(strKey(0), strKey(1))
                Else
                    '// Subtract one second from this timer's countdown
                    modScripting.SetPTCount strKey(0), strKey(1), modScripting.GetPTLeft(strKey(0), strKey(1)) - 1
                End If
            End If
        End If
    Next
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error: " & Err.description & " in scTimer_Timer()."
    
    Exit Sub
End Sub


Private Sub Timer_Timer()
    On Error GoTo ERROR_HANDLER

    Dim U As String, IdleMsg As String, s() As String
    Dim IdleWaitS As String, IdleType As String
    Dim f As Integer, IdleWait As Integer
    Static iCounter As Integer, UDP As Byte
     
    If iCounter >= 32760 Then iCounter = 0
    
    'If LenB(Dir$(GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt")) = 0 Then
    '    'On Error Resume Next
    '    Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt" For Output As #1
    '    Close #1
    'End If
    
    If ReadCfg("Other", "ProfileAmp") = "Y" And g_Online Then Call UpdateProfile
    
    
    BotVars.JoinWatch = 0
    
    If (AutoChatFilter) Then
        'If ((GetTickCount() - AutoChatFilter) >= 180000) Then
        '    frmChat.AddChat RTBColors.TalkBotUsername, _
        '        "Chat filters have been deactivated; " & _
        '            "activate them by pressing CTRL + F."
        '
        '    Call WriteINI("Other", "Filters", "N")
        '
        '    Filters = False
        '
        '    AutoChatFilter = 0
        'End If
    End If
    
    iCounter = iCounter + 1
    
    If sckBNet.State = 7 And Not IsW3 Then
        If iCounter Mod 4 = 0 Then
            PBuffer.SendPacket &H0
        End If
    End If
    
    IdleMsg = ReadCfg("Main", "Idles")
    
    If IdleMsg <> "Y" Then Exit Sub
    
    If ((StrComp(g_Channel.Name, "Clan SBs", vbTextCompare) = 0) And _
        (IsStealthBotTech = False)) Then Exit Sub
    
    IdleMsg = ReadCfg("Main", "IdleMsg")
    IdleWaitS = ReadCfg("Main", "IdleWait")
    IdleType = ReadCfg("Main", "IdleType")
    
    If StrictIsNumeric(IdleWaitS) Then IdleWait = IdleWaitS

    If IdleWait < 2 Then Exit Sub
    
    If iCounter >= IdleWait And StrComp(LCase(g_Channel.Name), "op [vl]", vbTextCompare) <> 0 Then
        iCounter = 0
        'on error resume next
        If IdleType = "msg" Or IdleType = vbNullString Then
        
            If StrComp(IdleMsg, "null", vbTextCompare) = 0 Or IdleMsg = vbNullString Then
                Exit Sub
            End If
            
            IdleMsg = Replace(IdleMsg, "%cpuup", ConvertTime(GetUptimeMS))
            IdleMsg = Replace(IdleMsg, "%chan", g_Channel.Name)
            IdleMsg = Replace(IdleMsg, "%c", g_Channel.Name)
            IdleMsg = Replace(IdleMsg, "%me", GetCurrentUsername)
            IdleMsg = Replace(IdleMsg, "%v", CVERSION)
            IdleMsg = Replace(IdleMsg, "%ver", CVERSION)
            IdleMsg = Replace(IdleMsg, "%bc", BanCount)
            IdleMsg = Replace(IdleMsg, "%botup", ConvertTime(uTicks))
            IdleMsg = Replace(IdleMsg, "%mp3", Replace(MediaPlayer.TrackName, "&", "+"))
            IdleMsg = Replace(IdleMsg, "%quote", GetRandomQuote)
            IdleMsg = Replace(IdleMsg, "%rnd", GetRandomPerson)
            IdleMsg = Replace(IdleMsg, "%t", Time$)
            
            If (IdleMsg = vbNullString) Then
                GoTo Error
            End If
            
        ElseIf IdleType = "uptime" Then
            IdleMsg = "/me -: System Uptime: " & ConvertTime(GetUptimeMS()) & " :: Connection Uptime: " & ConvertTime(uTicks) & " :: " & CVERSION & " :-"
            
        ElseIf IdleType = "mp3" Then
            Dim WindowTitle As String
            
            WindowTitle = MediaPlayer.TrackName
            
            If WindowTitle = vbNullString Then
                IdleMsg = "/me .: " & CVERSION & " :: anti-idle :."
                GoTo Send
            End If
            
            If (MediaPlayer.IsPaused) Then
                IdleMsg = "/me -: Now Playing: " & WindowTitle & " (paused) :: " & CVERSION & " :-"
            Else
                IdleMsg = "/me -: Now Playing: " & WindowTitle & " :: " & CVERSION & " :-"
            End If
 
        ElseIf IdleType = "quote" Then
            U = GetRandomQuote
            If Len(U) > 217 Then GoTo Error
            IdleMsg = "/me : " & U
            
        End If
        GoTo Send
Error:
        IdleMsg = "/me -- " & CVERSION
Send:
        If sckBNet.State = 7 Then
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
        
        Close #f
    End If
    
    Exit Sub

ERROR_HANDLER:

    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in Timer_Timer()."
    
    Exit Sub
    
End Sub

Private Sub tmrClanUpdate_Timer()
    ' ...
    If (g_Channel.Self.Clan <> vbNullString) Then
        RequestClanMOTD
    End If
End Sub

Private Sub tmrFriendlistUpdate_Timer()
    If (g_Online) Then
        If (BotVars.UsingDirectFList) Then
            If (lvFriendList.ListItems.Count > 0) Then
                Call FriendListHandler.RequestFriendsList(PBuffer)
            End If
        End If
    End If
End Sub

Private Sub tmrSilentChannel_Timer(Index As Integer)
    On Error GoTo ERROR_HANDLER

    Dim user    As clsUserObj
    Dim Item    As ListItem
    
    Dim I       As Integer ' ...
    Dim j       As Integer ' ...
    Dim found   As Boolean ' ...
    Dim WasZero As Boolean ' ...
    
    ' ...
    If (g_Channel.IsSilent = False) Then
        Exit Sub
    End If

    ' ...
    If (Index = 0) Then
        ' ...
        If (frmChat.mnuDisableVoidView.Checked = False) Then
            ' ...
            'For i = 1 To g_Channel.Users.Count
            '    ' with our doevents, we can miss our cue indicating that we
            '    ' need to stop silent channel processing and cause an rte.
            '    If (i > g_Channel.Users.Count) Then
            '        Exit For
            '    End If
           '
           '     ' ...
           '     Set user = g_Channel.Users(i)
           '
            '    ' ...
            '    If (lvChannel.FindItem(user.DisplayName) Is Nothing) Then
            '        Dim Stats As String ' ...
            '        Dim Clan  As String ' ...
            '
            '        ' ...
            '        ParseStatstring user.Statstring, Stats, Clan
            ''
            '        ' ...
            '        AddName user.DisplayName, user.Game, user.Flags, user.Ping, user.Clan
            '    End If
            'Next i
            
            ' ...
            Call LockWindowUpdate(&H0)
            
            ' ...
            lblCurrentChannel.Caption = GetChannelString()
        End If
    
        ' ...
        tmrSilentChannel(0).Enabled = False
    ElseIf (Index = 1) Then
        ' ...
        If (mnuDisableVoidView.Checked = False) Then
            ' ...
            If (g_Channel.IsSilent) Then
                ' ...
                Call g_Channel.ClearUsers
                
                ' ...
                frmChat.lvChannel.ListItems.Clear
            End If
        
            ' ...
            Call AddQ("/unsquelch " & GetCurrentUsername, PRIORITY.SPECIAL_MESSAGE)
        End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error: " & Err.description & " in tmrSilentChannel_Timer(" & Index & ")."
    
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

Sub Connect()
    Dim BNLS As Byte
    
    ' ...
    g_username = BotVars.Username
    
    If sckBNet.State = 0 And sckBNLS.State = 0 Then
    
        Const f As String = "Main" ', p As String = "config.ini"
        'Vars
        If BotVars.Username = vbNullString Or _
            BotVars.Password = vbNullString Or _
                BotVars.CDKey = vbNullString Or _
                    BotVars.Server = vbNullString Or _
                        BotVars.HomeChannel = vbNullString _
                            Or BotVars.Product = vbNullString Then
                            
                MsgBox "You haven't provided enough information to connect! " & _
                    "Please edit your connection settings by choosing Bot Settings under the Settings menu." & _
                    vbNewLine & "Required information to connect: Username, Password, CDKey, Home channel, " & _
                    "Server, your choice of Client and Trigger."
                
                Call DoDisconnect(1)
                
                Exit Sub
        End If
        
        SetTitle "Connecting..."
        AddChat RTBColors.InformationText, "Connecting your bot..."
        
        If ((StrComp(BotVars.Product, "PX2D", vbTextCompare) = 0) Or _
            (StrComp(BotVars.Product, "VD2D", vbTextCompare) = 0)) Then
            
            Dii = True
        Else
            Dii = False
        End If
        'StealthLock Check
        'NOT 'Disabled due to public version!

        #If (BETA = 1) Then
            Call AddChat(RTBColors.InformationText, "Authorizing your private-release " & _
                "bot, please wait.")
            
            ' ...
            If (GetAuth(BotVars.Username)) Then
                Call AddChat(RTBColors.SuccessText, _
                    "Private usage authorized, connecting to Battle.Net.")
                
                ' was auth function bypassed?
                If (AUTH_CHECKED = False) Then
                    BotVars.Password = Chr$(0)
                End If
            Else
                ' ...
                Call AddChat(RTBColors.ErrorMessageText, _
                    "- - - - - YOU ARE NOT AUTHORIZED TO USE THIS PROGRAM - - - - -")
                
                ' ...
                Call DoDisconnect
                
                ' ...
                UserCancelledConnect = False
                
                ' ...
                Exit Sub
            End If
        #End If
        
        
        If BotVars.BNLS Then
            Call Event_BNLSConnecting
            
            With sckBNLS
                If .State <> 0 Then .Close
                
                .RemoteHost = BotVars.BNLSServer
                .RemotePort = 9367
                .Connect
            End With
            BNLS = 1
        Else
            Call Event_BNetConnecting
        End If
        
        With sckBNet
            If .State <> 0 Then
                AddChat RTBColors.ErrorMessageText, "Already connected."
                Exit Sub
            End If
            
            If Not BotVars.UseProxy Then
                .RemoteHost = BotVars.Server
                .RemotePort = 6112
            Else
                '// PROXY
                If BotVars.ProxyPort > 0 And LenB(BotVars.ProxyIP) > 0 Then
                    .RemoteHost = BotVars.ProxyIP
                    .RemotePort = BotVars.ProxyPort
                Else
                    MsgBox "You have selected to use proxies, but no proxy is configured. Please set one up in the Advanced " & _
                        " section of Bot Settings."
                        
                    DoDisconnect
                    
                    Exit Sub
                End If
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

'OLD GETACCESS
'Function GetAccess(ByVal Username As String) As Integer
'    On Error Resume Next
'    Dim strCompare As String
'    Dim Acc As Integer, f As Integer
'    Username = LCase(Username)
'    If Dir$(App.Path & "\users.txt") = vbNullString Then Exit Function
'
'    f = FreeFile
'    Open (App.Path & "\users.txt") For Input As #f
'    If LOF(f) < 2 Then Exit Function
'    Do
'        Input #f, strCompare
'        On Error Resume Next
'        Acc = Right(strCompare, 3)
'        strCompare = LCase(Left$(strCompare, Len(strCompare) - 3))
'        strCompare = Replace(strCompare, " ", vbNullString)
'        If StrComp(Username, strCompare, vbTextCompare) = 0 Then
 '           GetAccess = Acc
 '           Close #f
'            Exit Function
'        End If
'    Loop While Not EOF(f)
'    Close #f
'End Function

'Function Evaluate(ByVal Username As String, ByVal Flag As String, ByVal Access As Integer, ByVal Mode As Byte) As Boolean
'
'End Function


Public Sub Pause(ByVal fSeconds As Single, Optional ByVal AllowEvents As Boolean = True)
    Dim I As Integer
    If AllowEvents Then
        For I = 0 To (1000 * fSeconds) \ 100
            Sleep 100
            DoEvents
        Next I
    Else
        Sleep fSeconds * 1000
    End If
End Sub

'/* Fires every second */
Private Sub UpTimer_Timer()

    On Error GoTo ERROR_HANDLER

    Dim newColor  As Long
    Dim I         As Integer
    Dim Pos       As Integer
    Dim doCheck   As Boolean

    uTicks = (uTicks + 1000)
    
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

    ' ...
    If (g_Channel.IsSilent = False) Then
        ' ...
        doCheck = True
    
        ' ...
        For I = 1 To g_Channel.Users.Count
            ' ...
            With g_Channel.Users(I)
                ' ...
                If (g_Channel.Self.IsOperator) Then
                    ' ...
                    If (.IsOperator = False) Then
                        ' channel password
                        If ((BotVars.ChannelPasswordDelay > 0) And (Len(BotVars.ChannelPassword) > 0)) Then
                            ' ...
                            If (.PassedChannelAuth = False) Then
                                ' ...
                                If (.TimeInChannel() > BotVars.ChannelPasswordDelay) Then
                                    ' ...
                                    If (GetSafelist(.DisplayName) = False) Then
                                        ' ...
                                        Ban .DisplayName & " Password time is up", (AutoModSafelistValue - 1)
                                         
                                        ' ...
                                        doCheck = False
                                    End If
                                End If
                            End If
                        End If
                        
                        ' idle bans
                        If ((doCheck) And ((BotVars.IB_On = BTRUE) And (BotVars.IB_Wait > 0))) Then
                            ' ...
                            If (.TimeSinceTalk() > BotVars.IB_Wait) Then
                                ' ...
                                If (GetSafelist(.DisplayName) = False) Then
                                    ' ...
                                    Ban .DisplayName & " Idle for " & BotVars.IB_Wait & "+ seconds", _
                                        (AutoModSafelistValue - 1), IIf(BotVars.IB_Kick, 1, 0)
                                        
                                    ' ...
                                    doCheck = False
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' ...
                If (BotVars.NoColoring = False) Then
                    ' ...
                    Pos = checkChannel(.DisplayName)
                
                    ' ...
                    If (Pos > 0) Then
                        ' ...
                        newColor = GetNameColor(.Flags, .TimeSinceTalk, StrComp(.DisplayName, _
                            GetCurrentUsername, vbBinaryCompare) = 0)
                        
                        ' ...
                        If (lvChannel.ListItems(Pos).ForeColor <> newColor) Then
                            lvChannel.ListItems(Pos).ForeColor = newColor
                        End If
                    End If
                End If
            End With
            
            ' ...
            doCheck = True
        Next I
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in UpTimer_Timer()."

    Exit Sub
    
End Sub

'StealthLock (c) 2003 Stealth, Please do not remove this header
Private Function GetAuth(ByVal Username As String) As Long
    On Error GoTo ERROR_HANDLER

    Static lastAuth     As Long    ' ...
    Static lastAuthName As String  ' ...
    
    Dim clsCRC32 As clsCRC32 ' ...
    Dim result   As Integer  ' string variable for storing beta authorization result
                             ' 0  == unauthorized
                             ' >0 == authorized
                                                   
    ' ...
    Set clsCRC32 = New clsCRC32
                      
    ' ...
    If (lastAuth) Then
        ' ...
        If (StrComp(Username, lastAuthName, vbTextCompare) = 0) Then
            ' ...
            GetAuth = lastAuth
            
            ' ...
            Exit Function
        End If
    End If
    
    ' ...
    If (clsCRC32.GenerateCRC32(BETA_AUTH_URL) = BETA_AUTH_URL_CRC32) Then
        result = _
            CInt(Val(INet.OpenURL(BETA_AUTH_URL & Username)))
    End If

    ' ...
    Do While INet.StillExecuting
        DoEvents
    Loop
    
    ' ...
    If (result > 0) Then
        ' ...
        lastAuth = result
        lastAuthName = Username
        
        ' ...
        GetAuth = lastAuth
        
        ' ...
        AUTH_CHECKED = True
    End If
    
    ' ...
    Set clsCRC32 = Nothing

    ' ...
    Exit Function

ERROR_HANDLER:

    ' ...
    Set clsCRC32 = Nothing

    ' ...
    GetAuth = False
    
    ' ...
    Exit Function
    
End Function

' ...
Sub AddQ(ByVal Message As String, Optional msg_priority As Integer = -1, Optional ByVal user As String = _
    vbNullString, Optional ByVal Tag As String = vbNullString, Optional OversizeDelimiter As String = " ")
    
    ' ...
    On Error GoTo ERROR_HANDLER
    
    ' maximum size of battle.net messages
    Const BNET_MSG_LENGTH = 223

    ' ...
    Static LastGTC  As Double
    Static BanCount As Integer
    
    ' ...
    Dim strTmp As String
    
    ' ...
    strTmp = Message
    
    ' ...
    If (msg_priority > 100) Then
        msg_priority = 100
    End If
    
    ' ...
    If (g_Queue.Count = 0) Then
        BanCount = 0
    End If
    
    ' ...
    If (strTmp <> vbNullString) Then
        Dim Splt()         As String      ' ...
        Dim I              As Long        ' ...
        Dim currChar       As Long        ' ...
        Dim Send           As String      ' ...
        Dim command        As String      ' ...
        Dim GTC            As Double      ' ...
        Dim strUser        As String      ' ...
        Dim nameConversion As Boolean     ' ...
        Dim Q              As clsQueueOBj ' ...
        Dim j              As Integer     ' ...
        Dim delay          As Integer     ' ...
        
        ' ...
        ReDim Splt(0)
    
        ' check for tabs and replace with spaces (2005-09-23)
        If (InStr(1, strTmp, Chr$(9), vbBinaryCompare) <> 0) Then
            strTmp = Replace$(strTmp, Chr$(9), Space(4))
        End If
        
        ' ...
        For I = 1 To Len(strTmp)
            ' ...
            currChar = Asc(Mid$(strTmp, I, 1))
        
            ' ...
            If (currChar < 32) Then
                Exit Sub
            End If
        Next I
        
        ' ...
        If (StrComp(Left$(strTmp, 1), "/", vbBinaryCompare) = 0) Then
            Dim Index As Long ' ...
            
            ' ...
            For I = 2 To Len(strTmp)
                ' ...
                currChar = Asc(Mid$(strTmp, I, 1))
            
                ' ...
                If (currChar <> Asc(Space(1))) Then
                    Exit For
                End If
            Next I
            
            ' ...
            If (I > 2) Then
                strTmp = "/" & Mid$(strTmp, I)
            End If

            ' ...
            Index = InStr(1, strTmp, Space(1), vbBinaryCompare)
            
            ' ...
            If (Index > 2) Then
                ' ...
                command = Mid$(strTmp, 2, (Index - 2))

                ' ...
                If ((command = "w") Or _
                    (command = "whisper") Or _
                    (command = "m") Or _
                    (command = "msg") Or _
                    (command = "message") Or _
                    (command = "whois") Or _
                    (command = "where") Or _
                    (command = "whereis") Or _
                    (command = "squelch") Or _
                    (command = "unsquelch") Or _
                    (command = "ignore") Or _
                    (command = "unignore") Or _
                    (command = "ban") Or _
                    (command = "unban") Or _
                    (command = "kick") Or _
                    (command = "designate")) Then
        
                    ' ...
                    Splt() = Split(strTmp, Space$(1), 3)
                    
                    ' ...
                    If (UBound(Splt) > 0) Then
                        ' ...
                        command = Splt(0) & Space$(1) & reverseUsername(Splt(1)) & _
                            Space$(1)

                        ' ...
                        If ((g_Channel.IsSilent) And (frmChat.mnuDisableVoidView.Checked = False)) Then
                            ' ...
                            If ((LCase$(Splt(0)) = "/unignore") Or (LCase$(Splt(0)) = "/unsquelch")) Then
                                ' ...
                                If (StrComp(Splt(1), GetCurrentUsername, vbTextCompare) = 0) Then
                                    ' ...
                                    lvChannel.ListItems.Clear
                                End If
                            End If
                        End If
                        
                        ' ...
                        If (UBound(Splt) > 1) Then
                            ReDim Preserve Splt(0 To UBound(Splt) - 1)
                        End If
                    End If
                ElseIf ((command = "f") Or _
                        (command = "friends")) Then
                    
                    ' ...
                    Splt() = Split(strTmp, Space$(1), 3)
                    
                    ' ...
                    command = Splt(0) & Space$(1)
                    
                    ' ...
                    If (UBound(Splt) >= 1) Then
                        ' ...
                        command = command & Splt(1) & Space$(1)
                    
                        ' ...
                        If (UBound(Splt) >= 2) Then
                            ' ...
                            Select Case (LCase$(Splt(1)))
                                Case "m", "msg"
                                    ' ...
                                    ReDim Preserve Splt(0 To UBound(Splt) - 1)

                                Case Else
                                    ' ...
                                    Splt() = Split(strTmp, Space$(1), 4)
                                
                                    ' ...
                                    If ((StrReverse$(BotVars.Product) = "WAR3") Or _
                                        (StrReverse$(BotVars.Product) = "W3XP")) Then
                                        
                                        ' ...
                                        command = command & reverseUsername(Splt(2)) & _
                                            Space$(1)
                                    Else
                                        ' ...
                                        command = command & Splt(2) & Space$(1)
                                    End If
                                    
                                    ' ...
                                    If (UBound(Splt) >= 3) Then
                                        command = command & Splt(3)
                                    End If
                            End Select
                        End If
                    End If
                Else
                    ' ...
                    command = "/" & command & Space$(1)
                    
                    ' ...
                    strTmp = Mid$(strTmp, Len(command) + 1)
                End If
                
                ' ...
                If (Len(command) >= BNET_MSG_LENGTH) Then
                    Exit Sub
                End If

                ' ...
                If (UBound(Splt) > 0) Then
                    ' ...
                    strTmp = Mid$(strTmp, _
                        (Len(Join(Splt(), Space$(1))) + (Len(Space$(1))) + 1))
                End If
            End If
        End If
        
        ' ...
        If (msg_priority < 0) Then
            Dim cmdName    As String ' ...
            Dim spaceIndex As Long   ' ...
            
            ' ...
            If (Len(Message) >= 2) Then
                ' ...
                spaceIndex = InStr(1, Message, Space$(1), vbBinaryCompare)
                
                ' ...
                If (spaceIndex) Then
                    cmdName = LCase$(Left$(Mid$(Message, 2), spaceIndex - 2))
                Else
                    cmdName = LCase$(Mid$(Message, 2))
                End If
            End If
            
            ' ...
            Select Case (cmdName)
                Case "designate": msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "resign":    msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "who":       msg_priority = PRIORITY.SPECIAL_MESSAGE
                Case "ban":       msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE
                Case "unban":     msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE
                Case "kick":      msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE
                Case Else:        msg_priority = PRIORITY.MESSAGE_DEFAULT
            End Select
        End If
        
        ' ...
        Call SplitByLen(strTmp, (BNET_MSG_LENGTH - Len(command)), Splt(), vbNullString, _
            " [more]", OversizeDelimiter)

        ' ...
        ReDim Preserve Splt(0 To UBound(Splt))

        ' ...
        For I = LBound(Splt) To UBound(Splt)
            ' store current tick
            GTC = GetTickCount()
            
            ' store working copy
            Send = _
                command & Splt(I)
            
            ' ...
            Set Q = New clsQueueOBj
            
            ' ...
            With Q
                .Message = Send
                .PRIORITY = msg_priority
                .ResponseTo = vbNullString
                .Tag = Tag
            End With

            ' ...
            Call g_Queue.Push(Q)
            
            ' should we subject this message to the typical delay,
            ' or can we get it out of here a bit faster?  If we
            ' want it out of here quick, we need an empty queue
            ' and have had at least 10 seconds elapse since the
            ' previous message.
            If (g_Queue.Count() = 0) Then
                If (GTC - LastGTC >= 10000) Then
                    ' set default message delay when queue is empty (in ms)
                    delay = 10
                    
                    ' are we issuing a ban or kick command?
                    If (msg_priority = PRIORITY.CHANNEL_MODERATION_MESSAGE) Then
                        delay = BanDelay()
                    End If
                    
                    ' set the delay before our next queue cycle
                    With frmChat.QueueTimer
                        .Enabled = False
                        .Interval = delay
                        .Enabled = True
                    End With
                End If
            End If
        Next I
        
        ' store our tick for future reference
        LastGTC = GTC
    End If
        
    Exit Sub
    
ERROR_HANDLER:
    Call AddChat(vbRed, "Error: " & Err.description & " in AddQ().")

    Exit Sub
End Sub

Private Function BanDelay() As Integer

    ' define default error handler
    On Error GoTo ERROR_HANDLER
    
    ' base ban delay
    ' The base delay serves two functions: it prevents likely ineffectual attempts at
    ' banning fast floodbots & it provides a small window for bots without similar ban
    ' delay functions to do banning without incurring the high risk of double bans.
    ' The base delay prevents banning at any lower interval than what is specified.
    BanDelay = 100

    ' do we have ops?
    If (g_Channel.Self.IsOperator) Then
        Dim OpCount As Integer ' ...
        Dim j       As Integer ' ...
        
        ' loop through users in channel
        For j = 1 To g_Channel.Users.Count
            ' is user an operator?
            If (g_Channel.Users(j).IsOperator) Then
                OpCount = (OpCount + 1)
            End If
        Next j
        
        ' do we have more than one op?
        If (OpCount > 1) Then
            ' seed rnd function
            Randomize
        
            ' set random ban delay based primarily on op count
            BanDelay = _
                (BanDelay + ((1 + Rnd * (OpCount * 2)) * (1 + Rnd * 200)))
        End If
    End If
    
    ' exit procedure
    Exit Function

' default error handler
ERROR_HANDLER:
    ' display error message
    Call AddChat(vbRed, "Error: " & Err.description & " in BanDelay().")

    ' exit procedure
    Exit Function

End Function

Sub ClearChannel()
    ' ...
    Set g_Channel = Nothing
    
    ' ...
    Set g_Channel = New clsChannelObj

    ' ...
    lvChannel.ListItems.Clear
    
    ' ...
    lblCurrentChannel.Caption = vbNullString
    
    ' ...
    PassedClanMotdCheck = False
End Sub

Sub ReloadConfig(Optional Mode As Byte = 0)
    On Error GoTo ERROR_HANDLER

    Const MN                 As String = "Main"
    Const OT                 As String = "Other"

    Dim default_group_access As udtGetAccessResponse
    Dim s                    As String
    Dim I                    As Integer
    Dim f                    As Integer
    Dim Index                As Integer
    Dim D2GameConventions    As String
    Dim W3GameConventions    As String
    Dim gameConventions      As String
    Dim bln                  As Boolean
    Dim doConvert            As Boolean
    Dim command_output()     As String
        
    s = ReadCfg(OT, "Timestamp")
    If StrictIsNumeric(s) And Val(s) < 4 Then BotVars.TSSetting = CInt(s) Else BotVars.TSSetting = 0
    
    s = BotVars.Username
    
    BotVars.Username = ReadCfg(MN, "Username")
    
    If LenB(s) > 0 Then
        If StrComp(BotVars.Username, s, vbTextCompare) <> 0 Then
            AddChat RTBColors.ServerInfoText, "Username set to " & BotVars.Username & "."
        End If
    End If
    
    BotVars.Password = ReadCfg(MN, "Password")
    BotVars.CDKey = UCase$(ReadCfg(MN, "CDKey"))
    BotVars.ExpKey = UCase$(ReadCfg(MN, "ExpKey"))
    
    If BotVars.ExpKey = "" Then
        BotVars.ExpKey = UCase$(ReadCfg(MN, "LODKey"))
    End If
    
    BotVars.Product = ReadCfg(MN, "Product")
    BotVars.Server = ReadCfg(MN, "Server")
    BotVars.BanUnderLevel = Val(ReadCfg(OT, "BanUnderLevel"))
    BotVars.BanD2UnderLevel = Val(ReadCfg(OT, "BanD2UnderLevel"))
    BotVars.HomeChannel = ReadCfg(MN, "HomeChan")
    BotVars.BotOwner = ReadCfg(MN, "Owner")
    BotVars.Trigger = ReadCfg(MN, "Trigger")
    
    If (BotVars.TriggerLong = vbNullString) Then
        BotVars.Trigger = "."
    End If
    
    Call LoadDatabase
    
    If Mode <> 1 Then
        s = ReadCfg(OT, "ChatFont")
        If s <> vbNullString And s <> rtbChat.Font.Name Then
            rtbChat.Font.Name = s
        End If
        
        s = ReadCfg(OT, "ChanFont")
        If s <> vbNullString And s <> lvChannel.Font.Name Then
            lvChannel.Font.Name = s
        End If
        
        s = ReadCfg(OT, "ChatSize")
        If StrictIsNumeric(s) Then
            If CInt(s) <> rtbChat.Font.Size Then
                rtbChat.Font.Size = s
            End If
        End If
        
        s = ReadCfg(OT, "ChanSize")
        If StrictIsNumeric(s) Then
            If CInt(s) <> lvChannel.Font.Size Then
                lvChannel.Font.Size = s
            End If
        End If
    End If
    
    s = ReadCfg(OT, "Filters")
    
    If s = "Y" Then
        Filters = True
    Else
        Filters = False
        
        BotVars.JoinWatch = 0
    End If
    
    BotVars.AutofilterMS = 300 'default
    s = ReadCfg(MN, "AutofilterMS")
    
    If LenB(s) > 0 Then
        If StrictIsNumeric(s) Then
            BotVars.AutofilterMS = Val(s)
        End If
    End If
    
    s = ReadCfg("Override", "AutoModerationSafelistValue")
    If Val(s) > 0 And Val(s) < 1001 Then
        AutoModSafelistValue = Val(s)
    Else
        AutoModSafelistValue = 20
    End If
    
    s = ReadCfg(MN, "BNLSServer")
    If LenB(s) = 0 Then
        s = "bnls.valhallalegends.com"
    End If
    
    BotVars.BNLSServer = s
        
    s = ReadCfg(MN, "ShowOfflineFriends")
    
    If s = "Y" Then
        BotVars.ShowOfflineFriends = True
    Else
        BotVars.ShowOfflineFriends = False
    End If
    
    s = ReadCfg(OT, "HideClanDisplay")
    If (s = "Y") Then
        With lvChannel
            .Width = (.Width - .ColumnHeaders(2).Width)
            .ColumnHeaders(2).Width = 0
        End With
    End If
    
    s = ReadCfg(OT, "HidePingDisplay")
    If (s = "Y") Then
        With lvChannel
            .Width = (.Width - .ColumnHeaders(3).Width)
            .ColumnHeaders(3).Width = 0
        End With
    End If
    
    s = ReadCfg(OT, "RetainOldBans")
    If (s = "Y") Then
        BotVars.RetainOldBans = True
    Else
        BotVars.RetainOldBans = False
    End If
    
    s = ReadCfg(OT, "StoreAllBans")
    If (s = "Y") Then
        BotVars.StoreAllBans = True
    Else
        BotVars.StoreAllBans = False
    End If
    
    gameConventions = ReadCfg(OT, "UseGameConventions")
    D2GameConventions = ReadCfg(OT, "UseD2GameConventions")
    W3GameConventions = ReadCfg(OT, "Usew3GameConventions")
    
    If (gameConventions = vbNullString) Then
        gameConventions = "Y"
    End If
    
    If (W3GameConventions = vbNullString) Then
        W3GameConventions = "Y"
    End If
    
    With BotVars
        .UseGameConventions = (gameConventions = "Y")
        .UseD2GameConventions = (D2GameConventions = "Y")
        .UseW3GameConventions = (W3GameConventions = "Y")
    End With

    ' ...
    If (g_Online) Then
        Dim found       As ListItem ' ...
        Dim CurrentUser As Object
        Dim outbuf      As String

        ' ...
        SetTitle GetCurrentUsername & ", online in channel " & g_Channel.Name
        
        ' ...
        lvChannel.ListItems.Clear
        
        ' ...
        For I = 1 To g_Channel.Users.Count
            ' ...
            Set CurrentUser = g_Channel.Users(I)
        
            ' ...
            AddName CurrentUser.DisplayName, CurrentUser.game, CurrentUser.Flags, CurrentUser.Ping, _
                CurrentUser.Stats.IconCode, CurrentUser.Clan
        Next I
        
        ' ...
        frmChat.lvFriendList.ListItems.Clear
        
        ' ...
        For I = 1 To g_Friends.Count
            ' ...
            Set CurrentUser = g_Friends(I)
        
            ' ...
            AddFriend CurrentUser.DisplayName, CurrentUser.game, CurrentUser.Status
        Next I
    End If
    
    s = ReadCfg(OT, "JoinLeaves")
    If s = "Y" Then JoinMessagesOff = False Else JoinMessagesOff = True
    
    s = ReadCfg(OT, "ShowStatsIcons")
    
    If s = "N" Then
        BotVars.ShowStatsIcons = False
    Else
        BotVars.ShowStatsIcons = True
    End If
    
    s = ReadCfg(OT, "Mail")
    If s = "N" Then mail = False Else mail = True
    
'    s = ReadCFG(OT, "DisableMonitor")
'    If s = "Y" Then DisableMonitor = True Else DisableMonitor = False
'
    s = ReadCfg(OT, "BanEvasion")
    If s = "N" Then BotVars.BanEvasion = False Else BotVars.BanEvasion = True

    
    s = ReadCfg(OT, "Logging")
    If StrictIsNumeric(s) Then BotVars.Logging = Val(s) Else BotVars.Logging = 1
        
    mnuToggleWWUse.Checked = (ReadCfg("Main", "UseWWs") = "Y")
    
    s = ReadCfg(MN, "WhisperBack")
    If s = "N" Then BotVars.WhisperCmds = False Else BotVars.WhisperCmds = True
    
    s = ReadCfg(OT, "Phrasebans")
    If s = "Y" Then PhraseBans = True Else PhraseBans = False
    
    s = ReadCfg(OT, "CaseSensitiveFlags")
    If s = "Y" Then BotVars.CaseSensitiveFlags = True Else BotVars.CaseSensitiveFlags = False
    
    s = ReadCfg(OT, "AutoCompletePostfix")
    BotVars.AutoCompletePostfix = s
    
    s = ReadCfg(MN, "UseBNLS")
    If s = "N" Then BotVars.BNLS = False Else BotVars.BNLS = True
    
    s = ReadCfg(MN, "LogDBActions")
    If s = "Y" Then BotVars.LogDBActions = True Else BotVars.LogDBActions = False
    
    s = ReadCfg(MN, "LogCommands")
    If s = "Y" Then BotVars.LogCommands = True Else BotVars.LogCommands = False
    
    '/* time to idle: defaults to 600 seconds / 10 minutes idle */
    BotVars.SecondsToIdle = 600
    
    s = ReadCfg("Override", "SecondsToIdle")
    If LenB(s) > 0 Then
        If StrictIsNumeric(s) Then
            If Val(s) < 1000000 Then
                BotVars.SecondsToIdle = Val(s)
            End If
        End If
    End If
    
    s = ReadCfg(OT, "PeonBans")
    If s = "1" Then BotVars.BanPeons = 1 Else BotVars.BanPeons = 0
    
    s = ReadCfg(OT, "KickOnYell")
    If s = "Y" Then BotVars.KickOnYell = 1 Else BotVars.KickOnYell = 0
    
    ' Capped at 32767, topic=29986 -Andy
    s = ReadCfg(OT, "IdleBanDelay")
    If StrictIsNumeric(s) Then
        If Val(s) < 32767 Then
            BotVars.IB_Wait = CInt(s)
        Else
            BotVars.IB_Wait = 32767
        End If
    End If
    
    s = ReadCfg(OT, "DefaultShitlistGroup")
    BotVars.DefaultShitlistGroup = s
    
    ' ...
    If (BotVars.DefaultShitlistGroup <> vbNullString) Then
        ' ...
        default_group_access = _
                GetAccess(BotVars.DefaultShitlistGroup, "GROUP")
        
        ' ...
        If (default_group_access.Username = vbNullString) Then
            ' ...
            Call ProcessCommand(GetCurrentUsername, "/add " & BotVars.DefaultShitlistGroup & _
                    " B --type group --banmsg Shitlisted", True, False, False)
        End If
    End If
    
    s = ReadCfg(OT, "DefaultTagbansGroup")
    BotVars.DefaultTagbansGroup = s
    
    ' ...
    If (BotVars.DefaultTagbansGroup <> vbNullString) Then
        ' ...
        default_group_access = _
                GetAccess(BotVars.DefaultTagbansGroup, "GROUP")
        
        ' ...
        If (default_group_access.Username = vbNullString) Then
            ' ...
            Call ProcessCommand(CurrentUsername, "/add " & BotVars.DefaultTagbansGroup & _
                    " B --type group --banmsg Tagbanned", True, False, False)
        End If
    End If
    
    s = ReadCfg(OT, "DefaultSafelistGroup")
    BotVars.DefaultSafelistGroup = s
    
    ' ...
    If (BotVars.DefaultSafelistGroup <> vbNullString) Then
        ' ...
        default_group_access = _
                GetAccess(BotVars.DefaultSafelistGroup, "GROUP")
        
        ' ...
        If (default_group_access.Username = vbNullString) Then
            ' ...
            Call ProcessCommand(GetCurrentUsername, "/add " & BotVars.DefaultSafelistGroup & _
                    " S --type group", True, False, False)
        End If
    End If
    
    s = ReadCfg(OT, "AllowMP3")
    If s = "N" Then BotVars.DisableMP3Commands = True Else BotVars.DisableMP3Commands = False
    
    s = ReadCfg(MN, "MaxBacklogSize")
    If ((s = vbNullString) Or (StrictIsNumeric(s) = False)) Then
        BotVars.MaxBacklogSize = 10000
    Else
        BotVars.MaxBacklogSize = Val(s)
    End If
    
    s = ReadCfg(MN, "MaxLogFileSize")
    If ((s = vbNullString) Or (StrictIsNumeric(s) = False)) Then
        BotVars.MaxLogFileSize = 50000000
    Else
        BotVars.MaxLogFileSize = Val(s)
    End If
    
    s = ReadCfg(MN, "DoNotUseDirectFList")
    If s = "Y" Then
        BotVars.UsingDirectFList = False
    Else
        BotVars.UsingDirectFList = True
    End If
    
    s = ReadCfg(MN, "URLDetect")
    If s = "Y" Then
        EnableURLDetect rtbChat.hWnd
    Else
        DisableURLDetect
    End If
    
    If BotVars.MaxBacklogSize < 0 Then BotVars.MaxBacklogSize = 10000
    If BotVars.MaxLogFileSize < 0 Then BotVars.MaxLogFileSize = 50000000
    
    '// this section must read _absolutely correctly_ or the SetTimer API call will fail
    s = ReadCfg(MN, "ReconnectDelay")
    If LenB(s) > 0 Then
        If StrictIsNumeric(s) Then
            If Val(s) < 60000 Then
                If Val(s) > 0 Then
                    BotVars.ReconnectDelay = Val(s)
                Else
                    BotVars.ReconnectDelay = 1000
                End If
            Else
                BotVars.ReconnectDelay = 60000
            End If
        Else
            BotVars.ReconnectDelay = 1000
        End If
    Else
        BotVars.ReconnectDelay = 1000
    End If
    
    s = ReadCfg(MN, "UseBackupChan")
    If s = "Y" Then BotVars.UseBackupChan = True Else BotVars.UseBackupChan = False
    
    BotVars.BackupChan = ReadCfg(MN, "BackupChan")
    
    s = ReadCfg("Main", "UTF8")
    If s = "N" Then mnuUTF8.Checked = False Else mnuUTF8.Checked = True
    
    mnuToggleShowOutgoing.Checked = (ReadCfg("Main", "ShowOutgoingWhispers") = "Y")
    mnuHideWhispersInrtbChat.Checked = (ReadCfg("Main", "HideWhispersInMain") = "Y")
    mnuIgnoreInvites.Checked = (ReadCfg("Main", "IgnoreClanInvitations") = "Y")
    
    'LoadSafelist
    LoadArray LOAD_PHRASES, Phrases()
    LoadArray LOAD_FILTERS, gFilters()
    
    's = readcfg(ot, "SendDelay")
    'If strictisnumeric(s) Then
    '    quSend.Interval = s
    'Else
    '     quSend.Interval = 1150
    'End If
    
    's = readcfg(ot, "LowerDelay")
    'If strictisnumeric(s) Then
    '    quLower.Interval = s
    'Else
         'quLower.Interval = 2500
    'End If
    
    ProtectMsg = ReadCfg("Other", "ProtectMsg")
    If ProtectMsg = vbNullString Then ProtectMsg = "Channel Protection"
    
    Call LoadOutFilters
    
    s = ReadCfg(OT, "IdleBans")
    If s = "Y" Then BotVars.IB_On = BTRUE Else BotVars.IB_On = BFALSE
    
    s = ReadCfg(OT, "KickIdle")
    If s = "Y" Then BotVars.IB_Kick = True Else BotVars.IB_Kick = False
    
    s = ReadCfg(OT, "IdleBanDelay")
    If (StrictIsNumeric(s) And s <> vbNullString) Then BotVars.IB_Wait = CInt(s) Else BotVars.IB_Wait = 0
    
    s = ReadCfg(MN, "Spoof")
    If StrictIsNumeric(Left$(s, 1)) Then
        Select Case Left$(s, 1)
            Case "0": BotVars.Spoof = 0
            Case "1": BotVars.Spoof = 1
            Case "2": BotVars.Spoof = 2
            Case Else: BotVars.Spoof = 0
        End Select
    Else
        BotVars.Spoof = 0
    End If
    
    s = ReadCfg(MN, "Protect")
    If s = "Y" Then Protect = True Else Protect = False
    
    s = ReadCfg(MN, "UDP")
    If s = "Y" Then BotVars.UseUDP = True Else BotVars.UseUDP = False
    
    s = ReadCfg(OT, "IPBans")
    If s = "Y" Then BotVars.IPBans = True Else BotVars.IPBans = False
    
    s = ReadCfg(MN, "UseAltBnls")
    If s = "Y" Then BotVars.UseAltBnls = True Else BotVars.UseAltBnls = False
    
    's = ReadCFG(OT, "ClientBansOn")
    'If s = "Y" Then BotVars.ClientBans = True Else BotVars.ClientBans = False
    
    's = ReadCFG(OT, "ClientBans")
    'ClientBans() = Split(s, " ")
    'If UBound(ClientBans) = -1 Then ReDim ClientBans(0)
    
    s = ReadCfg(MN, "QuietTime")
    If s = "Y" Then BotVars.QuietTime = True Else BotVars.QuietTime = False
    
    s = ReadCfg(OT, "FlashWindow")
    If s = "Y" Then mnuFlash.Checked = True Else mnuFlash.Checked = False
    
    s = ReadCfg(MN, "UseProxy")
    If s = "Y" Then
        BotVars.UseProxy = True 'Added comment so the thing would let me commit. To shut Swent up
        If (sckBNet.State = sckConnected) Then BotVars.ProxyStatus = psOnline
    Else
        BotVars.UseProxy = False
    End If
    
    s = ReadCfg(MN, "ProxyPort")
    If StrictIsNumeric(s) Then
        If Val(s) < 65536 Then BotVars.ProxyPort = CLng(s) Else BotVars.ProxyPort = 0
    Else
        BotVars.ProxyPort = 0
    End If
    
    s = ReadCfg(MN, "ProxyIsSocks5")
    'Debug.Print s
    If s = "Y" Then BotVars.ProxyIsSocks5 = True Else BotVars.ProxyIsSocks5 = False
    'Debug.Print BotVars.ProxyIsSocks5
    
    s = ReadCfg(OT, "NoTray")
    If s = "Y" Then BotVars.NoTray = True Else BotVars.NoTray = False
    
    s = ReadCfg(OT, "NoAutocomplete")
    If s = "Y" Then BotVars.NoAutocompletion = True Else BotVars.NoAutocompletion = False
    
    s = ReadCfg(OT, "NoColoring")
    If s = "Y" Then BotVars.NoColoring = True Else BotVars.NoColoring = False
    
    s = ReadCfg(OT, "DisableVoidView")
    If s = "Y" Then mnuDisableVoidView.Checked = True Else mnuDisableVoidView.Checked = False
    
    s = ReadCfg(OT, "MediaPlayer")
    If s <> vbNullString Then
        BotVars.MediaPlayer = s
    Else
        BotVars.MediaPlayer = "Winamp"
    End If

    s = ReadCfg(MN, "UseRealm")
    If s = "Y" Then BotVars.UseRealm = True Else BotVars.UseRealm = False
    
    txtPre.text = ""
    txtPost.text = ""
    
    s = ReadCfg(OT, "DisablePrefix")
    If s = "Y" Then txtPre.Visible = False Else txtPre.Visible = True
    
    s = ReadCfg(OT, "DisableSuffix")
    If s = "Y" Then txtPost.Visible = False Else txtPost.Visible = True
    
    '[Other] MathAllowUI - Will allow People to use MessageBox/InputBox or other UI related commands in the .eval/.math commands ~Hdx 09-25-07
    s = ReadCfg(OT, "MathAllowUI")
    If s = "Y" Then SCRestricted.AllowUI = True Else SCRestricted.AllowUI = False
    
    BotVars.NoRTBAutomaticCopy = (ReadCfg("Override", "NoRTBAutomaticCopy") = "Y")
    
    BotVars.GreetMsg = ReadCfg(OT, "GreetMsg")
    BotVars.UseGreet = (ReadCfg(OT, "UseGreets") = "Y")
    BotVars.WhisperGreet = (ReadCfg(OT, "WhisperGreet") = "Y")
    
    BotVars.ProxyIP = ReadCfg(MN, "ProxyIP")
    
    s = ReadCfg(OT, "ChatDelay")
    If (s = vbNullString) Then
        BotVars.ChatDelay = 500
    Else
        BotVars.ChatDelay = CLng(Val(s))
    End If

    Call ChatQueue_Initialize

    'If BotVars.Logging < 2 Then
    '    MakeLoggingDirectory
    '
    '    If dir$(GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt") = vbNullString Then
    '        Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & ".txt" For Output As #1
    '        Close #1
    '    End If
    '
    '    If dir$(GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt") = vbNullString Then
    '        Open GetProfilePath() & "\Logs\" & Format(Date, "yyyy-MM-dd") & "-WHISPERS.txt" For Output As #1
    '        Close #1
    '    End If
    'End If
    
    If (g_Online) Then
        Call g_Channel.CheckUsers
    Else
        Err.Clear
    
        If (ReadCfg(MN, "LocalIP") <> vbNullString) Then
            If (Err.Number = 0) Then: sckBNet.Bind , ReadCfg(MN, "LocalIP")
            If (Err.Number = 0) Then: sckBNLS.Bind , ReadCfg(MN, "LocalIP")
            If (Err.Number = 0) Then: sckMCP.Bind , ReadCfg(MN, "LocalIP")
        End If
    End If
    
    Exit Sub

ERROR_HANDLER:
    If (Err.Number = 10049) Then
        AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in ReloadConfig()."
    End If
    
    Resume Next
End Sub

'returns OK to Proceed
Function DisplayError(ByVal ErrorNumber As Integer, bytType As Byte, _
    ByVal Source As enuErrorSources) As Boolean
    
    Dim s As String
    
    s = GErrorHandler.GetErrorString(ErrorNumber, Source)
    
    If (LenB(s) > 0) Then
        Select Case (bytType)
            Case 0: s = "[BNLS] " & s
            Case 1: s = "[BNET] " & s
            Case 2: s = "[PROXY] " & s
        End Select
        
        AddChat RTBColors.ErrorMessageText, s
    End If
    
    DisplayError = GErrorHandler.OKToProceed()
End Function

Sub LoadOutFilters()
    Const O As String = "Outgoing"
    Const f As String = "filters.ini"
    
    Dim s   As String
    Dim I   As Integer
    
    ReDim gOutFilters(1 To 1)
    ReDim Catch(0)
    
    Catch(0) = vbNullString
    
    s = ReadINI(O, "Total", f)
    
    If (Not (StrictIsNumeric(s))) Then
        Exit Sub
    End If
    
    For I = 1 To Val(s)
        gOutFilters(I).ofFind = Replace(LCase(ReadINI(O, "Find" & I, f)), "¦", " ")
        gOutFilters(I).ofReplace = Replace(ReadINI(O, "Replace" & I, f), "¦", " ")
        
        If (I <> Val(s)) Then
            ReDim Preserve gOutFilters(1 To I + 1)
        End If
    Next I
    
    If (Dir$(GetFilePath("catchphrases.txt")) <> vbNullString) Then
        I = FreeFile
        
        Open GetFilePath("catchphrases.txt") For Input As #I
        
            If (LOF(I) < 2) Then
                Close #I
                
                Exit Sub
            End If
            
            Do While Not EOF(I)
                Line Input #I, s
                
                If ((s <> vbNullString) And (s <> " ")) Then
                    Catch(UBound(Catch)) = LCase$(s)
                    
                    ReDim Preserve Catch(0 To UBound(Catch) + 1)
                End If
            Loop
            
            'Note: Why did this happen?
            'If Catch(0) = vbNullString Then Catch(0) = "¯"
            
        Close #I
    End If
End Sub

Function OutFilterMsg(ByVal strOut As String) As String
    Dim I As Integer
    
    If (UBound(gOutFilters) > 0) Then
        For I = LBound(gOutFilters) To UBound(gOutFilters)
            strOut = Replace(strOut, gOutFilters(I).ofFind, _
                gOutFilters(I).ofReplace)
        Next I
    End If
    
    OutFilterMsg = strOut
End Function

Sub SetFloodbotMode(ByVal Mode As Byte)
    Dim I   As Integer
    Dim Add As Byte

    Select Case (Mode)
        Case 0      'OFF
            bFlood = False
            
            SetNagelStatus frmChat.sckBNet.SocketHandle, True
                        
            Call g_Queue.Clear
            
            SetProcessPriority 0, frmChat.hWnd, ppNormal
            
            AddChat RTBColors.TalkBotUsername, "The channel list is most likely not accurate. Please " & _
                "rejoin the bot to correct this."
            
            ReDim gFloodSafelist(0)
            
        Case 1      'ON
            bFlood = True
            
            SetNagelStatus frmChat.sckBNet.SocketHandle, False
            
            Call g_Queue.Clear
            
            AddChat RTBColors.TalkBotUsername, "You have enabled Emergency Floodbot Protection:"
            AddChat RTBColors.InformationText, "- All message-queue actions have been suspended."
            AddChat RTBColors.InformationText, "- No greet messages or command responses will be displayed."
            AddChat RTBColors.InformationText, "- You can still use any commands from the channel or " & _
                "in-bot. You may not see their results."
            AddChat RTBColors.InformationText, "- Any user that joins and IS NOT SAFELISTED will be banned."
            AddChat RTBColors.InformationText, "- You can add users to the safelist using the safelist " & _
                "command in-bot or in-channel."
            AddChat RTBColors.TalkBotUsername, "Type '/efp off' to return to normal."
            
            SetProcessPriority 0, frmChat.hWnd, ppHigh
            
            ReDim gFloodSafelist(0)
            
            For I = 1 To colSafelist.Count
                If (Not (GetSafelist(colSafelist.Item(I).Name))) Then
                    gFloodSafelist(UBound(gFloodSafelist)) = _
                        Replace(PrepareCheck(colSafelist.Item(I).Name), Space(1), _
                            vbNullString)
                    
                    ReDim Preserve gFloodSafelist(UBound(gFloodSafelist) + 1)
                End If
            Next I
            
            For I = LBound(DB) To UBound(DB)
                With DB(I)
                    If (GetShitlist(DB(I).Username)) Then
                        Add = 1
                    End If
                End With
                
                If (GetSafelist(DB(I).Username)) Then
                    Add = 1
                End If
                
                If (Add = 0) Then
                    gFloodSafelist(UBound(gFloodSafelist)) = _
                        DB(I).Username
                    
                    ReDim Preserve gFloodSafelist(UBound(gFloodSafelist) + 1)
                End If
            Next I
    End Select
End Sub

Private Sub sckBNet_DataArrival(ByVal bytesTotal As Long)
    'On Error GoTo ERROR_HANDLER

    Dim strTemp     As String
    Dim fTemp       As String
    Dim BufferLimit As Long
    Dim interations As Integer
    
    sckBNet.GetData strTemp, vbString
    
'    Debug.Print "--> socket received a packet"
'    Debug.Print DebugOutput(strTemp)
    
    If Not BotVars.UseProxy Or BotVars.ProxyStatus = psOnline Then
        'Debug.Print String(50, "-")
        BNCSBuffer.AddData strTemp
    
        While BNCSBuffer.FullPacket And BufferLimit < 20
            
            strTemp = BNCSBuffer.GetPacket
            
            'BNCSBuffer.WriteLog "Parsing the following packet:", True
            'BNCSBuffer.WriteLog strTemp
            
            ' EFP System now running under the parsing layer for extra-crispy efficiency
            'If (bFlood) Then
            '    If (Asc(Mid$(strTemp, 2, 1)) = &HF) Then
            '        If (Conv(Mid$(strTemp, 5, 4)) = ID_JOIN) Then
            '            fTemp = KillNull(Mid$(strTemp, 29))
            '
            '            If (StrComp(flood, fTemp, vbBinaryCompare) <> 0) Then
            '                If (Not (GetSafelist(fTemp))) Then
            '                    If (floodCap < 45) Then
            '                        Call APISend("/ban " & fTemp)
            '
            '                        floodCap = (floodCap + 30)
            '
            '                        flood = fTemp
            '
            '                        Exit Sub
            '                    End If
            '                End If
            '            End If
            '        End If
            '    End If
            'End If
            
            Call BNCSParsePacket(strTemp)
            
            ' ...
            'interations = (interations + 1)
           
            'If (interations >= 2000) Then
            '    MsgBox "ahhhh!"
            '
            '    Exit Sub
            'End If
            
            ' Why do we need this?  Anyway, it's causing topic id #26093
            ' (The Void issue).
            'BufferLimit = (BufferLimit + 1) 'DebugOutput Left$(strBuffer, lngLen)
        Wend
    Else
        'proxy is ON and NOT CONNECTED
        'parse incoming data
        ParseProxyPacket strTemp
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in sckBNet_DataArrival()."
    
    Exit Sub
End Sub

Sub LoadArray(ByVal Mode As Byte, ByRef tArray() As String)
    Dim f As Integer
    Dim Path As String
    Dim temp As String
    Dim I As Integer
    Dim C As Integer
    
    f = FreeFile
    
    Const FI As String = "TextFilters"
    
    Select Case Mode
        Case LOAD_FILTERS
            Path = GetFilePath("filters.ini")
        Case LOAD_PHRASES
            Path = GetFilePath("phrasebans.txt")
        Case LOAD_DB
            Path = GetFilePath("users.txt")
            Exit Sub
    End Select
    
    If Dir(Path) <> vbNullString Then
        Open Path For Input As #f
        If LOF(f) > 2 Then
            ReDim tArray(0)
            If Mode <> LOAD_FILTERS Then
                Do
                    Line Input #f, temp
                    If Len(temp) > 0 Then
                        ' removed for 2.5 - why am I PCing it ?
                        'If Mode = LOAD_SAFELIST Then temp = PrepareCheck(temp)
                        tArray(UBound(tArray)) = LCase(temp)
                        ReDim Preserve tArray(UBound(tArray) + 1)
                    End If
                Loop While Not EOF(f)
            Else
                temp = ReadINI(FI, "Total", "filters.ini")
                If temp <> vbNullString And CInt(temp) > -1 Then
                    C = Int(temp)
                    For I = 1 To C
                        temp = ReadINI(FI, "Filter" & I, "filters.ini")
                        If temp <> vbNullString Then
                            tArray(UBound(tArray)) = LCase(temp)
                            If I <> C Then ReDim Preserve tArray(UBound(tArray) + 1)
                        End If
                    Next I
                End If
            End If
        End If
        Close #f
    End If
End Sub

Private Sub sckBNLS_Close()
    If Passed0x0F < 1 Then
        AddChat RTBColors.ErrorMessageText, StrReverse(StrReverse("This version of StealthBot has been disabled."))
        AddChat RTBColors.ErrorMessageText, StrReverse(StrReverse("Please visit http://www.stealthbot.net for more information and to download an updated version."))
        AddChat RTBColors.ErrorMessageText, StrReverse(StrReverse("If you are receiving this message in error, please check your firewall settings, especially programs such as Norton Internet Security. NIS is notorious for blocking the connection to the BNLS server."))
    End If
End Sub

Private Sub sckBNLS_Connect()
    If MDebug("all") Then
        AddChat COLOR_BLUE, "BNLS CONNECT"
    End If
    
    Call Event_BNLSConnected
    
    With PBuffer
        .InsertNTString "stealth"
        .vLSendPacket &HE
    End With
    
    SetNagelStatus sckBNLS.SocketHandle, False
    
    'frmChat.sckBNet.Connect 'BNLS is authorized, proceed to initiate BNet connection.
End Sub

'Private Function DebugFormat(ByVal sIn As String) As String
'    Dim buf As String
'    Dim i As Integer
'
'    For i = 1 To Len(sIn)
'        buf = buf &
'    Next i
'
'    DebugFormat = buf
'End Function

Private Sub sckBNLS_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ERROR_HANDLER

    Dim strTemp As String
    
    sckBNLS.GetData strTemp, vbString
    
    If BotVars.UseProxy And (BotVars.ProxyStatus = psConnecting Or BotVars.ProxyStatus = psLoggingIn) Then
    
'        Debug.Print "prox input: " & DebugOutput(strTemp)
    
        Select Case BotVars.ProxyStatus
            Case psConnecting
                'chr(5) or chr(4) depending on version & method
                'do an instr search to find the method number.
                'For public proxys you are looking for chr(0)
                '<macyui>
                
                If InStr(1, strTemp, Chr(5)) > 0 Or InStr(1, strTemp, Chr(4)) > 0 Then
                    If InStr(1, strTemp, Chr(0)) > 0 Then
                        UpdateProxyStatus psLoggingIn, PROXY_LOGGING_IN
                        
                        LogonToProxy sckBNLS, BotVars.BNLSServer, 9367, False
                    Else
                        UpdateProxyStatus psNotConnected, PROXY_IS_NOT_PUBLIC
                        sckBNLS.Close
                    End If
                Else
                    UpdateProxyStatus psNotConnected, PROXY_IS_NOT_PUBLIC
                    sckBNLS.Close
                End If
                
            Case psLoggingIn
                'Then when it sends back chr(5) & chr(0) indicating
                'connection success, login and proceed as usual.
                
                If InStr(1, strTemp, Chr(5)) > 0 And InStr(1, strTemp, Chr(0)) > 0 Then
                    UpdateProxyStatus psOnline, PROXY_LOGIN_SUCCESS
                Else
                    UpdateProxyStatus psNotConnected, PROXY_LOGIN_FAILED
                    sckBNLS.Close
                End If
                
        End Select
    Else
        BNLSBuffer.AddData strTemp
            
        While BNLSBuffer.FullPacket
            NLogin.ParsePacket BNLSBuffer.GetPacket
        Wend
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in sckBNLS_DataArrival()."

    Exit Sub
End Sub

Private Sub sckBNLS_Error(ByVal Number As Integer, description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Event_BNLSError(Number, description)
End Sub

'This function checks if the user selected when right-clicked is the same one when they click on the menu option. - FrOzeN
Private Function PopupMenuUserCheck() As Boolean
    If Not (lvChannel.SelectedItem Is Nothing) Then
        If mnuPopup.Tag <> lvChannel.SelectedItem.text Then
            PopupMenuUserCheck = False
            Exit Function
        End If
    End If
    
    PopupMenuUserCheck = True
End Function

Function GetSelectedUsers() As Collection
    Dim I As Integer ' ...

    Set GetSelectedUsers = New Collection
    
    For I = 1 To lvChannel.ListItems.Count
        If (lvChannel.ListItems(I).Selected) Then
            Call GetSelectedUsers.Add(lvChannel.ListItems(I).text)
        End If
    Next I
End Function

Function GetSelectedUser() As String
    If (lvChannel.SelectedItem Is Nothing) Then
        GetSelectedUser = vbNullString
    
        Exit Function
    End If
    
    GetSelectedUser = lvChannel.SelectedItem.text
End Function

Function GetRandomPerson() As String
    Dim I As Integer ' ...
    
    ' ...
    If (g_Channel.Users.Count > 0) Then
        ' ...
        Randomize
        
        ' ...
        I = Int(g_Channel.Users.Count * Rnd + 1)

        ' ...
        GetRandomPerson = g_Channel.Users(I).DisplayName
    End If
End Function

Function MatchClosest(ByVal toMatch As String, Optional startIndex As Long = 1) As String
    Dim lstView     As ListView

    Dim I           As Integer ' ...
    Dim CurrentName As String  ' ...
    Dim atChar      As Integer ' ...
    Dim Index       As Integer ' ...
    Dim Loops       As Integer ' ...

    I = InStr(1, toMatch, " ", vbBinaryCompare)
    
    If (I > 0) Then
        toMatch = Mid$(toMatch, I + 1)
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
            Dim C As Integer ' ...
            
            If (startIndex > .Count) Then
                Index = 1
            Else
                Index = startIndex
            End If
        
            While (Loops < 2)
                For I = Index To .Count 'for each user
                    CurrentName = .Item(I).text
                
                    If (Len(CurrentName) >= Len(toMatch)) Then
                        For C = 1 To Len(toMatch) 'for each letter in their name
                            If (StrComp(Mid$(toMatch, C, 1), Mid$(CurrentName, C, 1), _
                                vbTextCompare) <> 0) Then
                                
                                Exit For
                            End If
                        Next C
                        
                        If (C >= (Len(toMatch) + 1)) Then
                            MatchClosest = _
                                    .Item(I).text & BotVars.AutoCompletePostfix
                            
                            MatchIndex = I
                            
                            Exit Function
                        End If
                    End If
                Next I
                
                ' ...
                Index = 1
                
                ' ...
                Loops = (Loops + 1)
            Wend
            
            ' ...
            Loops = 0
        End If
    End With
    
    atChar = InStr(1, toMatch, "@", vbBinaryCompare)
    
    If (atChar <> 0) Then
        Dim tmp      As String  ' ...
        Dim realms() As String  ' ...
        
        ' ...
        ReDim realms(0 To 8)
        
        ' ...
        realms(0) = "USWest"
        realms(1) = "USEast"
        realms(2) = "Asia"
        realms(3) = "Europe"
        realms(4) = "Lordaeron"
        realms(5) = "Azeroth"
        realms(6) = "Kalimdor"
        realms(7) = "Northrend"
        
        If (startIndex > UBound(realms)) Then
            Index = 0
        Else
            Index = (startIndex - 1)
        End If
        
        ' ...
        If (Len(toMatch) >= (atChar + 1)) Then
            tmp = Mid$(toMatch, atChar + 1)

            While (Loops < 2)
                ' ...
                For I = Index To UBound(realms)
                    ' ...
                    If (Len(realms(I)) >= Len(tmp)) Then
                        ' ...
                        If (StrComp(Left$(realms(I), Len(tmp)), tmp, _
                            vbTextCompare) = 0) Then
                            
                            Dim j As Integer ' ...
                        
                            MatchClosest = Left$(toMatch, atChar) & realms(I) & _
                                    BotVars.AutoCompletePostfix
                            
                            MatchIndex = (I + 1)
                            
                            Exit Function
                        End If
                    End If
                Next I
                
                ' ...
                Index = 0
                
                ' ...
                Loops = (Loops + 1)
            Wend
        Else
            If (tmp = vbNullString) Then
                ' ...
                MatchClosest = Left$(toMatch, atChar) & realms(Index) & _
                        BotVars.AutoCompletePostfix
                    
                MatchIndex = (Index + 1)
                    
                Exit Function
            End If
        End If
    End If
    
    ' ...
    MatchClosest = vbNullString
    
    ' ...
    MatchIndex = 1
End Function

Function GetChannelString() As String
    If Not g_Online Then
        GetChannelString = vbNullString
    Else
        Select Case ListviewTabs.Tab
            Case 0: GetChannelString = g_Channel.Name & " (" & lvChannel.ListItems.Count & ")"
            Case 1: GetChannelString = lvFriendList.ListItems.Count & " friends listed"
            Case 2: GetChannelString = "Clan " & g_Clan.Name & ": " & lvClanList.ListItems.Count & " members."
        End Select
    End If
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

'Sub DeconstructMonitor()
'    If Not (MonitorForm Is Nothing) Then
'        Set MonitorForm = Nothing
'    End If
'End Sub

'SHOW/HIDE STUFF
Public Sub cmdShowHide_Click()
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        rtbWhispersVisible = (StrComp(cmdShowHide.Caption, CAP_HIDE))
        
        If rtbWhispersVisible Then
            Me.Height = Me.Height + rtbWhispers.Height - Screen.TwipsPerPixelY
        Else
            Me.Height = Me.Height - rtbWhispers.Height + Screen.TwipsPerPixelY
        End If
        
        rtbWhispers.Visible = rtbWhispersVisible
        
        WriteINI "Main", "ShowWhisperWindow", IIf(rtbWhispersVisible, "Y", "N")
        
        Call Form_Resize
    Else
        AddChat RTBColors.ErrorMessageText, "You cannot open or close the Whisper Window while the form is maximized."
    End If
End Sub

'// to be called on every successful login
Sub InitListviewTabs()
    Dim toSet As Boolean

    If IsW3() Then
        If Clan.isUsed Then
            toSet = True
        Else
            toSet = False
        End If
    Else
        toSet = False
    End If
    
    ListviewTabs.TabEnabled(LVW_BUTTON_CLAN) = toSet
    
    If BotVars.UsingDirectFList Then
        toSet = True
    Else
        toSet = False
    End If
    
    ListviewTabs.TabEnabled(LVW_BUTTON_FRIENDS) = toSet
End Sub

'// to be called at disconnect time
Sub DisableListviewTabs()
    ListviewTabs.TabEnabled(LVW_BUTTON_FRIENDS) = False
    ListviewTabs.TabEnabled(LVW_BUTTON_CLAN) = False
End Sub

Sub AddClanMember(ByVal Name As String, Rank As Integer, Online As Integer)
    
    Dim visible_rank As Integer
    
    visible_rank = Rank
    
    If visible_rank = 0 Then visible_rank = 1
    If visible_rank > 4 Then visible_rank = 5 '// handle bad ranks
    
    '// add user
    
    Name = KillNull(Name)
    
    With lvClanList
        .ListItems.Add .ListItems.Count + 1, , Name, , visible_rank
        .ListItems(.ListItems.Count).ListSubItems.Add , , , Online + 6
        .ListItems(.ListItems.Count).ListSubItems.Add , , visible_rank
        .SortKey = 2
        .SortOrder = lvwDescending
        .Sorted = True
    End With
    
    ' ...
    lblCurrentChannel.Caption = GetChannelString()
    
    ' ...
    frmChat.ListviewTabs_Click 0
    
    On Error Resume Next
    SControl.Run "Event_ClanInfo", Name, Rank, Online
End Sub

Private Function GetClanSelectedUser() As String
    With lvClanList
        If Not (.SelectedItem Is Nothing) Then
            If .SelectedItem.Index < 1 Then
                GetClanSelectedUser = vbNullString: Exit Function
            Else
                GetClanSelectedUser = .SelectedItem.text
            End If
        End If
    End With
End Function

Private Sub lvClanList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        lvClanList.SetFocus
        
        Dim lvhti As LVHITTESTINFO
        Dim lItemIndex As Long, m_lCurItemIndex As Long
        
        lvhti.pt.X = X / Screen.TwipsPerPixelX
        lvhti.pt.Y = Y / Screen.TwipsPerPixelY
        lItemIndex = SendMessageAny(lvClanList.hWnd, LVM_HITTEST, 0, lvhti) + 1
        
        If lItemIndex > 0 Then
            lvClanList.ListItems(lItemIndex).Selected = True
            
            If Not (lvClanList.SelectedItem Is Nothing) Then
                If lvClanList.SelectedItem.Index < 0 Then
                    
                    mnuPopDem.Enabled = False
                    mnuPopPro.Enabled = False
                    mnuPopBNProfile.Enabled = False
                    mnuPopRem.Enabled = False
                    
                Else
                
                    mnuPopRem.Enabled = False
                    mnuPopDem.Enabled = False
                    mnuPopPro.Enabled = False
                    
                    If g_Clan.Self.Rank > 2 Then
                            
                        mnuPopBNProfile.Enabled = True
                        
                        Select Case lvClanList.SelectedItem.SmallIcon
                        
                            Case 4
                                mnuPopDem.Enabled = False
                                mnuPopRem.Enabled = False
                                mnuPopPro.Enabled = False
                                
                            Case 3
                                
                                mnuPopPro.Enabled = False
                                
                                If g_Clan.Self.Rank = 4 Then
                                    
                                    mnuPopDem.Enabled = True
                                    mnuPopRem.Enabled = True
                                    
                                Else
                                    
                                    mnuPopDem.Enabled = False
                                    mnuPopRem.Enabled = False
                                
                                End If
                            
                            Case 2
                                
                                mnuPopDem.Enabled = True
                                mnuPopPro.Enabled = True
                                mnuPopRem.Enabled = True
                                
                            Case 1
                                
                                mnuPopDem.Enabled = False
                                mnuPopPro.Enabled = True
                                mnuPopRem.Enabled = True
                                
                        End Select
                    
                    End If
                End If
            End If
            
            If StrComp(GetClanSelectedUser(), GetCurrentUsername, vbTextCompare) = 0 Then
                If g_Clan.Self.Rank > 0 Then
                    mnuSP2.Visible = True
                    mnuPopLeaveClan.Visible = True
                Else
                    mnuSP2.Visible = False
                    mnuPopLeaveClan.Visible = False
                End If
                
                mnuPopRem.Visible = False
                mnuPopLeaveClan.Visible = True
            Else
                mnuPopRem.Visible = True
                mnuPopLeaveClan.Visible = False
            End If
            
            PopupMenu mnuPopClanList
        End If
    End If
End Sub

Private Sub mnuPopLeaveClan_Click()
    If MsgBox("Are you sure you want to leave the clan?", vbYesNo, "StealthBot") = vbYes Then
        With PBuffer
            .InsertDWord &H1    '//cookie
            .InsertNTString GetCurrentUsername
            .SendPacket &H78
        End With

        AwaitingSelfRemoval = 1
    End If
End Sub

Private Sub mnuPopBNProfile_Click()
    frmProfile.lblUsername.Caption = GetClanSelectedUser
    RequestProfile GetClanSelectedUser
End Sub

Private Sub mnuPopBNWebProfileW3XP_Click()
    GetW3LadderProfile GetClanSelectedUser, W3XP
End Sub

Private Sub mnuPopBNWebProfileWAR3_Click()
    GetW3LadderProfile GetClanSelectedUser, WAR3
End Sub

Private Sub mnuPopDem_Click()
    If MsgBox("Are you sure you want to demote " & GetClanSelectedUser & "?", vbYesNo, "StealthBot") = vbYes Then
        
        With PBuffer
            .InsertDWord &H1
            .InsertNTString GetClanSelectedUser
            .InsertByte lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon - 1
            .SendPacket &H7A
        End With
        
        AwaitingClanInfo = 1
        
    End If
End Sub

Private Sub mnuPopPro_Click()
    If MsgBox("Are you sure you want to promote " & GetClanSelectedUser & "?", vbYesNo, "StealthBot") = vbYes Then
        With PBuffer
            .InsertDWord &H3
            .InsertNTString GetClanSelectedUser
            .InsertByte lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon + 1
            .SendPacket &H7A
        End With
        
        AwaitingClanInfo = 1
    End If
End Sub

Private Sub mnuPopRem_Click()
    Dim L As Long
    L = TimeSinceLastRemoval

    If L < 30 Then
        AddChat RTBColors.ErrorMessageText, "You must wait " & 30 - L & " more seconds before you " & _
                "can remove another user from your clan."
    Else
        If MsgBox("Are you sure you want to remove this user from the clan?", vbExclamation + vbYesNo, _
                "StealthBot") = vbYes Then
                
            With PBuffer
                If lvClanList.SelectedItem.Index > 0 Then
                    .InsertDWord 1 'lvClanList.ListItems(lvClanList.SelectedItem.Index).SmallIcon
                    .InsertNTString GetClanSelectedUser
                    .SendPacket &H78
                End If
                
                AwaitingClanInfo = 1
            End With
            
            LastRemoval = GetTickCount
        End If
    End If
End Sub

Sub DoConnect()
    If sckBNLS.State <> 0 Or sckBNet.State <> 0 Then Call DoDisconnect
    uTicks = 0
    
    UserCancelledConnect = False
    
    'Reset the BNLS auto-locator list
    GotBNLSList = False
    
    If Not IsValidIPAddress(BotVars.Server) And BotVars.UseProxy Then
        AddChat RTBColors.ErrorMessageText, "[PROXY] Proxied connections must use a direct server IP address, such as those listed below your desired gateway in the Connection Settings menu, to connect."
        AddChat RTBColors.ErrorMessageText, "[PROXY] Please change servers and try connecting again."
    Else
        Call Connect
    End If
End Sub

Sub DoDisconnect(Optional ByVal DoNotShow As Byte = 0, Optional ByVal LeaveUCCAlone As Boolean = False)
    On Error GoTo ERROR_HANDLER

    Dim I As Integer
    
    If (Not (UserCancelledConnect)) Then
        SetTitle "Disconnected"
        
        Call NLogin.CloseConnection(DoNotShow)
        
        Set g_Channel = Nothing
        Set g_Clan = Nothing
        Set g_Friends = Nothing
        
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
        
        BotVars.ProxyStatus = psNotConnected
        
        Clan.isUsed = False
        lvClanList.ListItems.Clear
        
        BNLSBuffer.ClearBuffer
        BNCSBuffer.ClearBuffer
        
        g_Connected = False
        g_Online = False
        
        Call ClearChannel
        lvClanList.ListItems.Clear
        lvFriendList.ListItems.Clear
        
        ' ...
        'tmrSilentChannel(0).Enabled = False
        
        Call g_Queue.Clear
    
        Passed0x0F = 0
        uTicks = 0
        
        With mnuPublicChannels(0)
            .Caption = vbNullString
            .Visible = False
        End With
        
        For I = 1 To mnuPublicChannels.Count - 1
            Call Unload(mnuPublicChannels(I))
        Next I
        
        If ((Me.WindowState = vbNormal) And _
            (DoNotShow = 0)) Then
            
            Call cboSend.SetFocus
        End If
        
        DestroyNLSObject
        
        Unload frmRealm
        
        PassedClanMotdCheck = False
        
        On Error Resume Next
        SControl.Run "Event_LoggedOff"
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & " in DoDisconnect()."
    
    Exit Sub
End Sub

Public Sub ParseFriendsPacket(ByVal PacketID As Long, ByVal Contents As String)
    FriendListHandler.ParsePacket PacketID, Contents
End Sub

Public Sub ParseClanPacket(ByVal PacketID As Long, ByVal Contents As String)
    ClanHandler.ParseClanPacket PacketID, Contents
End Sub

Public Sub RecordWindowPosition(Optional Maximized As Boolean = False)
    'Don't record other position information if maximized, otherwise when they unmaximize it will be fullscreen width and height. - FrOzeN
    If Not Maximized Then
        WriteINI "Position", "Left", Int(Me.Left / Screen.TwipsPerPixelX)
        WriteINI "Position", "Top", Int(Me.Top / Screen.TwipsPerPixelY)
        WriteINI "Position", "Height", Int(Me.Height / Screen.TwipsPerPixelY)
        WriteINI "Position", "Width", Int(Me.Width / Screen.TwipsPerPixelX)
    End If
    
    WriteINI "Position", "Maximized", CStr(Maximized)
    WriteINI "Main", "ConfigVersion", CONFIG_VERSION
End Sub

Public Sub MakeLoggingDirectory()
    On Error Resume Next
    MkDir GetProfilePath() & "\Logs\"
End Sub

' Called from several points to keep accurate tabs on the user's prior selection
'  in the send combo
Public Sub RecordcboSendSelInfo()
    'Debug.Print "SelStart: " & cboSend.SelStart & ", SelLength: " & cboSend.SelLength
    cboSendSelLength = cboSend.SelLength
    cboSendSelStart = cboSend.SelStart
End Sub

