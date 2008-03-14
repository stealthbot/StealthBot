VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsCommandObj.cls
' Copyright (C) 2008 Eric Evans
' ...

Option Explicit

Private m_command_docs As clsCommandDocObj
Private m_name         As String
Private m_args         As String
Private m_local        As Boolean
Private m_publicOutput As Boolean
Private m_datasrc      As String

' ...
Public Property Get Name() As String
    
    Name = m_name
    
End Property

' ...
Public Property Let Name(strName As String)

    m_name = strName
    
End Property

' ...
Public Property Get datasrc() As String
    
    datasrc = m_datasrc
    
End Property

' ...
Public Property Let datasrc(str As String)

    datasrc = str
    
End Property

' ...
Public Property Get Args() As String

    Args = m_args
    
End Property

' ...
Public Property Let Args(strArgs As String)

    m_args = strArgs
    
End Property

' ...
Public Property Get IsLocal() As Boolean
    
    IsLocal = m_local
    
End Property

' ...
Public Property Let IsLocal(bln As Boolean)

    m_local = bln
    
End Property

' ...
Public Property Get PublicOutput() As Boolean

    PublicOutput = m_publicOutput
    
End Property

' ...
Public Property Let PublicOutput(bln As Boolean)

    m_publicOutput = bln
    
End Property

' ...
Public Property Get docs() As Object

    Set docs = OpenCommand(m_name, datasrc)

End Property
