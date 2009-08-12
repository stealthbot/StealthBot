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

'// This object is a representation of a command instance. A reference to this object
'// is returned to a script module by using the abstract IsCommand() method.


Private m_validCommand As Boolean
Private m_command_docs As clsCommandDocObj
Private m_name         As String
Private m_username     As String
Private m_args         As String
Private m_arguments    As Collection
Private m_local        As Boolean
Private m_publicOutput As Boolean
Private m_xmlarguments As Dictionary

Public Property Get Username() As String
    Username = m_username
End Property
Public Property Let Username(strName As String)
    m_username = strName
End Property
Public Property Get Name() As String
    Name = m_name
End Property
Public Property Let Name(strName As String)
    m_name = strName
End Property
Public Property Get Args() As String
    Args = m_args
End Property
Public Property Let Args(strArgs As String)
    m_args = strArgs
    Call SplitArguments
End Property
Public Property Get Arguments() As Collection
    Set Arguments = m_arguments
End Property
Public Property Let Arguments(col As Collection)
    Set m_arguments = col
End Property
Public Property Get IsLocal() As Boolean
    IsLocal = m_local
End Property
Public Property Let IsLocal(bln As Boolean)
    m_local = bln
End Property
Public Property Get PublicOutput() As Boolean
    PublicOutput = m_publicOutput
End Property
Public Property Let PublicOutput(bln As Boolean)
    m_publicOutput = bln
End Property


Public Property Get IsValidCommand() As Boolean
    IsValidCommand = m_validCommand
End Property
Public Property Let IsValidCommand(bln As Boolean)
    m_validCommand = bln
End Property
Public Property Get HasAccess() As Boolean
    HasAccess = m_command_docs.HasAccess(Me, m_username)
End Property
Public Property Get HasCorrectSyntax() As Boolean
    HasCorrectSyntax = m_command_docs.HasCorrectSyntax(Me)
End Property

Public Sub Class_Initialize()
    
    '// initialize values
    m_validCommand = False
    Set m_command_docs = New clsCommandDocObj
    m_name = vbNullString
    m_username = vbNullString
    m_args = vbNullString
    Set m_arguments = New Collection
    m_local = False
    m_publicOutput = False
    Set m_xmlarguments = New Dictionary
    m_xmlarguments.CompareMode = TextCompare
    
End Sub

Public Sub Class_Terminate()
    
    '// clean up
    Set m_command_docs = Nothing
    Set m_arguments = Nothing
    
End Sub

Public Property Get docs() As clsCommandDocObj
    
    If m_command_docs Is Nothing Then
        '// this command is nothing, lets create it
        Set m_command_docs = New clsCommandDocObj
        Call m_command_docs.OpenCommand(m_name, Chr$(0))
        Set docs = m_command_docs
    Else
        '// this command already has a value, lets make sure its still valid
        If m_command_docs.Name = m_name Then
            '// all good, lets return it
            Set docs = m_command_docs
        Else
            '// ugh, this doc object is for a different command, we need to
            '// destroy it and start all over again
            Set m_command_docs = Nothing
            Set docs = Me.docs
        End If
    End If
End Property
Public Property Let docs(ByRef cmd As clsCommandDocObj)
    Set m_command_docs = cmd
End Property

Private Function StripNumeric(ByRef sString As String) As String
    Dim sTemp As String
    sTemp = StripWord(sString)
    If (IsNumeric(sTemp)) Then
        StripNumeric = sTemp
    ElseIf LenB(sTemp) > 0 Then
        sString = sTemp & Space(1) & sString
    End If
End Function

Private Function StripWord(ByRef sString As String) As String
    Dim I As Integer
    I = InStr(sString, Space(1))
    If (I > 0) Then
        StripWord = Left$(sString, I - 1)
        sString = Mid$(sString, I + 1)
    ElseIf Len(sString) > 0 Then
        StripWord = sString
        sString = vbNullString
    End If
End Function

Private Function StripString(ByRef sString As String) As String
    Dim I As Integer
    If (Left$(sString, 1) = Chr$(34)) Then
        sString = Replace$(sString, "\\", Chr$(1))
        sString = Replace$(sString, "\" & Chr(34), Chr$(2))
        I = InStr(2, sString, Chr$(34))
        If (I > 2) Then
            sString = Replace$(sString, Chr$(1), "\")
            sString = Replace$(sString, Chr$(2), Chr$(34))
            StripString = Left$(Mid$(sString, 2), I - 2)
            sString = Mid$(sString, I + 1)
        Else
            sString = Replace$(sString, Chr$(1), "\")
            sString = Replace$(sString, Chr$(2), Chr$(34))
            StripString = sString
            sString = vbNullString
        End If
    Else
        StripString = sString
        sString = vbNullString
    End If
End Function

Private Sub SplitArguments()
    Dim sArgs       As String
    Dim I           As Integer
    Dim Param       As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj
    Dim sTemp       As String
    sArgs = Me.Args
    
    For Each Param In Me.docs.Parameters
        Select Case LCase(Param.DataType)
            Case "word":    sTemp = StripWord(sArgs)
            Case "numeric": sTemp = StripNumeric(sArgs)
            Case "number":  sTemp = StripNumeric(sArgs)
            Case "string":  sTemp = StripString(sArgs)
        End Select
        If (LenB(sTemp) = 0 And Not Param.IsOptional) Then
            IsValidCommand = False
        End If
        m_xmlarguments(Param.Name) = sTemp
    Next
End Sub

Public Function Argument(sName As String) As String
    If (m_xmlarguments.Exists(sName)) Then
        Argument = m_xmlarguments.Item(sName)
    Else
        Argument = vbNullString
    End If
End Function
