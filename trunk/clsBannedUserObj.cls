VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBannedUserObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsBannedUserObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name             As String
Private m_is_banned        As Boolean
Private m_banned_date      As Date
Private m_ban_reason       As String
