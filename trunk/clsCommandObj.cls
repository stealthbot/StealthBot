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


Private m_valid        As Boolean
Private m_command_docs As clsCommandDocObj
Private m_name         As String
Private m_username     As String
Private m_args         As String
Private m_arguments    As Collection
Private m_local        As Boolean
Private m_publicOutput As Boolean
Private m_xmlarguments As Dictionary
Private m_hasaccess    As Boolean
Private m_splithasrun  As Boolean
Private m_waswhispered As Boolean
Private m_response     As Collection

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
Public Property Get IsValid() As Boolean
    If (Not m_splithasrun) Then SplitArguments
    IsValid = m_valid
End Property
Public Property Let IsValid(bln As Boolean)
    m_valid = bln
End Property
Public Property Get HasAccess() As Boolean
    If (Not m_splithasrun) Then SplitArguments
    HasAccess = m_hasaccess
End Property
Public Property Get WasWhispered() As Boolean
    WasWhispered = m_waswhispered
End Property
Public Property Let WasWhispered(bValue As Boolean)
    m_waswhispered = bValue
End Property

Public Sub Class_Initialize()
    
    '// initialize values
    m_valid = False
    Set m_command_docs = New clsCommandDocObj
    m_name = vbNullString
    m_username = vbNullString
    m_args = vbNullString
    Set m_arguments = New Collection
    m_local = False
    m_publicOutput = False
    Set m_xmlarguments = New Dictionary
    m_xmlarguments.CompareMode = TextCompare
    m_hasaccess = False
    m_splithasrun = False
    m_waswhispered = False
    Set m_response = New Collection
    
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
        If StrComp(m_command_docs.Name, m_name, TextCompare) = 0 Then
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


'Strips (removes) and returns a Numeric argument from the passed string
Private Function StripNumeric(ByRef sString As String) As String
    Dim sTemp As String
    sTemp = StripWord(sString)
    If (StrictIsNumeric(sTemp)) Then
        StripNumeric = sTemp
    ElseIf LenB(sTemp) > 0 Then
        sString = sTemp & Space(1) & sString
    End If
End Function

'Strips (removes) and returns a Word argument from the passed string
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

'Strips (removes) and returns a String argument from the passed string
'EXAs:
'This is a String -> This is a String
'This is a "String" -> This is a "String"
'"This is a String" -> This is a String
'"This is a \"String\"" -> This is a "String"
'"This is a \\\"String\"" -> This is a \"String"
'"This is a \String" -> This is a \String
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

'Will Split up the Arguments for this instance of the command, Based on the XML specs of this command.
'EXA: .Mail Username This is a message!
'Creates a Dictionary as such:
'  Dict("Username") = "Username"
'  Dict("Message")  = "This is a message!"
'This also checks the user's access to use this command in the specific restriction context.
Private Sub SplitArguments()
On Error GoTo ERROR_HANDLER
    Dim sArgs       As String
    Dim I           As Integer
    Dim Param       As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj
    Dim sTemp       As String
    Dim dbAccess    As udtGetAccessResponse
    dbAccess = GetCumulativeAccess(Me.Username)
    sArgs = Me.Args
    
    
    If (IsLocal) Then
      dbAccess.Rank = 201
      dbAccess.Flags = "A"
    End If
    
    If (dbAccess.Rank >= Me.docs.RequiredRank And Me.docs.RequiredRank > -1) Then
        m_hasaccess = True
    End If
    
    If (CheckForAnyFlags(Me.docs.RequiredFlags, dbAccess.Flags)) Then
        m_hasaccess = True
    End If
    
    IsValid = True
    
    For Each Param In Me.docs.Parameters
        Select Case LCase(Param.DataType)
            Case "word":    sTemp = StripWord(sArgs)
            Case "numeric": sTemp = StripNumeric(sArgs)
            Case "number":  sTemp = StripNumeric(sArgs)
            Case "string":  sTemp = StripString(sArgs)
        End Select
        
        For Each Restriction In Param.Restrictions 'Loop Through the Restrictions
            If (LenB(sTemp) > 0) Then
                If (CheckMatch(Restriction.MatchMessage, sTemp)) Then 'If they match (regex)
                    'If Rank = -1 It means it's missing, and it MUST have Flags. Or if Rank > User's Access
                    If (Restriction.RequiredRank = -1 Or Restriction.RequiredRank > dbAccess.Rank) Then
                        If (Not CheckForAnyFlags(Restriction.RequiredFlags, dbAccess.Flags)) Then
                            m_hasaccess = False
                        End If
                    End If
                End If
            End If
        Next
        
        If (LenB(sTemp) = 0 And Not Param.IsOptional) Then
            IsValid = False
        End If
        m_xmlarguments(Param.Name) = sTemp
    Next
    
    If (IsLocal) Then m_hasaccess = True
    m_splithasrun = True

    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.description & " in clsCommandObj.SplitArguments().")
End Sub

Public Function Argument(sName As String) As String
    If (Not m_splithasrun) Then SplitArguments
    If (m_xmlarguments.Exists(sName)) Then
        Argument = m_xmlarguments.Item(sName)
    Else
        Argument = vbNullString
    End If
End Function

Private Function CheckForAnyFlags(sNeeded As String, sHave As String) As Boolean
On Error GoTo ERROR_HANDLER
    Dim I As Integer
    CheckForAnyFlags = False
    
    If (LenB(sHave) = 0) Then Exit Function
    
    For I = 1 To Len(sNeeded)
        If (InStr(1, sHave, Mid(sNeeded, I, 1), vbTextCompare) > 0) Then
            CheckForAnyFlags = True
            Exit Function
        End If
    Next

    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.description & " in clsCommandObj.CheckForAnyFlags().")
End Function

'Adds a line to the response queue
Public Sub Respond(strResponse As String)
    If (LenB(strResponse) > 0) Then m_response.Add strResponse
End Sub

'Cleares the response queue
Public Sub ClearResponse()
    Do While m_response.Count
      m_response.Remove 1
    Loop
End Sub

'Gets the Response Queue
Public Function GetResponse() As Collection
    Set GetResponse = m_response
End Function

'This will respond in the proper style based on how the command was used.
'This will messup emote responses /me if it is not public output or whisper commands is turned on.
'If your response MUST be a specific style, Use AddQ or DSP
Public Function SendResponse()
On Error GoTo ERROR_HANDLER
    Dim I As Integer
    If (IsLocal) Then
        If (PublicOutput) Then
            For I = 1 To m_response.Count
                frmChat.AddQ m_response.Item(I), PRIORITY.CONSOLE_MESSAGE
            Next I
        Else
            For I = 1 To m_response.Count
                frmChat.AddChat RTBColors.ConsoleText, m_response.Item(I)
            Next I
        End If
    Else
        If ((BotVars.WhisperCmds) Or (WasWhispered)) Then
            For I = 1 To m_response.Count
                frmChat.AddQ "/w " & Username & Space$(1) & m_response.Item(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
            Next I
        Else
            For I = 1 To m_response.Count
                frmChat.AddQ m_response.Item(I), PRIORITY.COMMAND_RESPONSE_MESSAGE
            Next I
        End If
    End If
    ClearResponse
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.description & " in clsCommandObj.Respond().")
End Function

Private Function CheckMatch(sExpression As String, sData As String) As Boolean
On Error GoTo ERROR_HANDLER
    Dim mRegExp  As New RegExp
    mRegExp.Global = True
    mRegExp.Pattern = sExpression
    CheckMatch = mRegExp.Test(sData)
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in clsCommandObj.CheckMatch()"
End Function
