VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCommand.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_name         As String
Private m_params       As String
Private m_local        As Boolean
Private m_publicOutput As Boolean

' ...
Public Property Get Name() As String
    Name = m_name
End Property

' ...
Public Property Let Name(strName As String)
    m_name = strName
End Property

' ...
Public Property Get Params() As String
    Params = m_params
End Property

' ...
Public Property Let Params(strParams As String)
    m_params = Params
End Property

' ...
Public Property Get IsLocal() As Boolean
    IsLocal = m_local
End Property

' ...
Public Property Let IsLocal(blnLocal As Boolean)
    m_local = IsLocal
End Property

' ...
Public Property Get PublicOutput() As Boolean
    PublicOutput = m_publicOutput
End Property

' ...
Public Property Let PublicOutput(blnOutput As Boolean)
    m_publicOutput = PublicOutput
End Property
