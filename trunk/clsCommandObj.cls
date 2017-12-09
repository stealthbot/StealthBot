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

Option Explicit

'// This object is a representation of a command instance. A reference to this object
'// is returned to a script module by using the abstract IsCommand() method.


Private m_Valid        As Boolean
Private m_CommandDocs  As clsCommandDocObj
Private m_Name         As String
Private m_Username     As String
Private m_Args         As String
Private m_Arguments    As Collection
Private m_Local        As Boolean
Private m_PublicOutput As Boolean
Private m_XmlArguments As Dictionary
Private m_HasAccess    As Boolean
Private m_SplitHasRun  As Boolean
Private m_WasWhispered As Boolean
Private m_Response     As Collection
Private m_Restrictions As Dictionary

Public Property Get Username() As String
    Username = m_Username
End Property
Public Property Let Username(strName As String)
    m_Username = strName
End Property
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(strName As String)
    m_Name = strName
End Property
Public Property Get Args() As String
    Args = m_Args
End Property
Public Property Let Args(strArgs As String)
    m_Args = strArgs
End Property
Public Property Get Arguments() As Collection
    Set Arguments = m_Arguments
End Property
Public Property Let Arguments(col As Collection)
    Set m_Arguments = col
End Property
Public Property Get IsLocal() As Boolean
    IsLocal = m_Local
End Property
Public Property Let IsLocal(bln As Boolean)
    m_Local = bln
End Property
Public Property Get PublicOutput() As Boolean
    PublicOutput = m_PublicOutput
End Property
Public Property Let PublicOutput(bln As Boolean)
    m_PublicOutput = bln
End Property
Public Property Get IsValid() As Boolean
    If (Not m_SplitHasRun) Then SplitArguments
    IsValid = m_Valid
End Property
Public Property Let IsValid(bln As Boolean)
    m_Valid = bln
End Property
Public Property Get HasAccess() As Boolean
    If (Not m_SplitHasRun) Then SplitArguments
    HasAccess = m_HasAccess
End Property
Public Property Get WasWhispered() As Boolean
    WasWhispered = m_WasWhispered
End Property
Public Property Let WasWhispered(bValue As Boolean)
    m_WasWhispered = bValue
End Property

Public Property Get Source() As Integer
    If m_Local = True Then
        Source = 4
    Else
        If m_WasWhispered = True Then
            Source = 3
        Else
            Source = 1
        End If
    End If
End Property


Public Sub Class_Initialize()
    '// initialize values
    m_Valid = False
    Set m_CommandDocs = New clsCommandDocObj
    m_Name = vbNullString
    m_Username = vbNullString
    m_Args = vbNullString
    Set m_Arguments = New Collection
    m_Local = False
    m_PublicOutput = False
    Set m_XmlArguments = New Dictionary
    m_XmlArguments.CompareMode = TextCompare
    m_HasAccess = False
    m_SplitHasRun = False
    m_WasWhispered = False
    Set m_Response = New Collection
    Set m_Restrictions = New Dictionary
End Sub

Public Sub Class_Terminate()
    '// clean up
    Set m_CommandDocs = Nothing
    Set m_Arguments = Nothing
End Sub

Public Property Get docs() As clsCommandDocObj
    If m_CommandDocs Is Nothing Then
        '// this command is nothing, lets create it
        Set m_CommandDocs = New clsCommandDocObj
        Call m_CommandDocs.OpenCommand(m_Name, vbNullChar)
        Set docs = m_CommandDocs
    Else
        '// this command already has a value, lets make sure its still valid
        If StrComp(m_CommandDocs.Name, m_Name, TextCompare) = 0 Then
            '// all good, lets return it
            Set docs = m_CommandDocs
        Else
            '// ugh, this doc object is for a different command, we need to
            '// destroy it and start all over again
            Set m_CommandDocs = Nothing
            Set docs = Me.docs
        End If
    End If
End Property

Public Property Let docs(ByRef cmd As clsCommandDocObj)
    Set m_CommandDocs = cmd
End Property

'Strips (removes) and returns a Numeric argument from the passed string
Private Function StripNumeric(ByRef sString As String) As String
    Dim sTemp As String
    sTemp = StripWord(sString)
    If (StrictIsNumeric(sTemp, True)) Then
        StripNumeric = sTemp
    ElseIf LenB(sTemp) > 0 Then
        sString = sTemp & Space(1) & sString
    End If
End Function

'Strips (removes) and returns a Word argument from the passed string
Private Function StripWord(ByRef sString As String) As String
    Dim i As Integer
    i = InStr(sString, Space(1))
    If (i > 0) Then
        StripWord = Left$(sString, i - 1)
        sString = Mid$(sString, i + 1)
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
Private Function StripString(ByRef sString As String, IsLastArgument As Boolean) As String
    Dim i As Integer
    If (IsLastArgument) Then
        StripString = sString
        sString = vbNullString
        Exit Function
    End If
    If (Left$(sString, 1) = Chr$(34)) Then
        sString = Replace$(sString, "\\", Chr$(1))
        sString = Replace$(sString, "\" & Chr(34), Chr$(2))
        i = InStr(2, sString & " ", Chr$(34) & " ")
        If (i > 2) Then
            sString = Replace$(sString, Chr$(1), "\")
            sString = Replace$(sString, Chr$(2), Chr$(34))
            StripString = Left$(Mid$(sString, 2), i - 2)
            sString = Mid$(sString, i + 2)
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
    Dim i           As Integer
    Dim Param       As clsCommandParamsObj
    Dim Restriction As clsCommandRestrictionObj
    Dim sTemp       As String
    Dim sError      As String
    Dim dbAccess    As udtUserAccess
    
    If IsLocal Then
        dbAccess = Database.GetConsoleAccess()
    Else
        dbAccess = Database.GetUserAccess(Me.Username)
    End If
    
    sArgs = Me.Args
    
    
    If (dbAccess.Rank >= Me.docs.RequiredRank And Me.docs.RequiredRank > -1) Then
        m_HasAccess = True
    End If
    
    If (CheckForAnyFlags(Me.docs.RequiredFlags, dbAccess.Flags)) Then
        m_HasAccess = True
    End If
    
    IsValid = True
    
    For Each Param In Me.docs.Parameters
    
        Select Case LCase$(Param.DataType)
            Case "word":    sTemp = StripWord(sArgs)
            Case "numeric": sTemp = StripNumeric(sArgs)
            Case "number":  sTemp = StripNumeric(sArgs)
            Case "string":  sTemp = StripString(sArgs, (Param.Name = Me.docs.Parameters.Item(Me.docs.Parameters.Count).Name))
        End Select
        
        If (LenB(Param.MatchMessage)) Then
            If (Not CheckMatch(Param.MatchMessage, sTemp, Param.MatchCaseSensitive)) Then
                If m_HasAccess Then
                    If LenB(Param.MatchError) > 0 Then
                        sError = Replace$(Param.MatchError, "%Value", sTemp)
                        sError = Replace$(sError, "%Rank", dbAccess.Rank)
                        sError = Replace$(sError, "%Flags", dbAccess.Flags)

                        Respond sError
                    Else
                        Respond "That command is restricted."
                    End If
                    
                    m_SplitHasRun = True
                    IsValid = False
                Else
                    If (LCase$(Param.DataType) = "string") Then
                        sArgs = StringFormat("{0}{1}{0} {2}", Chr$(34), sTemp, sArgs)
                    Else
                        sArgs = StringFormat("{0} {1}", sTemp, sArgs)
                    End If
                End If
            End If
        End If
        
        If (LenB(sTemp) > 0) Then
            For Each Restriction In Param.Restrictions 'Loop Through the Restrictions
                m_Restrictions(Restriction.Name) = True
                If (CheckMatch(Restriction.MatchMessage, sTemp, Restriction.MatchCaseSensitive)) Then 'If they match (regex)
                    'If Rank = -1 It means it's missing, and it MUST have Flags. Or if Rank > User's Access
                    If (Restriction.RequiredRank = -1 Or Restriction.RequiredRank > dbAccess.Rank) Then
                        If (Not CheckForAnyFlags(Restriction.RequiredFlags, dbAccess.Flags)) Then
                            If m_HasAccess Then
                                If LenB(Restriction.MatchError) > 0 Then
                                    sError = Replace$(Restriction.MatchError, "%Value", sTemp)
                                    sError = Replace$(sError, "%Rank", dbAccess.Rank)
                                    sError = Replace$(sError, "%Flags", dbAccess.Flags)

                                    Respond sError
                                ElseIf LenB(Param.RestrictionsSharedError) > 0 Then
                                    sError = Replace$(Param.RestrictionsSharedError, "%Value", sTemp)
                                    sError = Replace$(sError, "%Rank", dbAccess.Rank)
                                    sError = Replace$(sError, "%Flags", dbAccess.Flags)

                                    Respond sError
                                Else
                                    Respond "That command is restricted."
                                End If
                                
                                If (Restriction.Fatal) Then m_HasAccess = False
                                m_Restrictions(Restriction.Name) = False
                            End If
                        End If
                    End If
                End If
            Next Restriction
        End If
        
        If (LenB(sTemp) = 0 And Not Param.IsOptional) Then
            IsValid = False
        End If
        m_XmlArguments(Param.Name) = sTemp
    Next
    
    If (IsLocal) Then m_HasAccess = True
    m_SplitHasRun = True

    Exit Sub
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.SplitArguments().")
End Sub

Public Function Argument(sName As String) As String
    If (Not m_SplitHasRun) Then SplitArguments
    If (m_XmlArguments.Exists(sName)) Then
        Argument = m_XmlArguments.Item(sName)
    Else
        Argument = vbNullString
    End If
End Function

Public Function Restriction(sName As String) As Boolean
    If (Not m_SplitHasRun) Then SplitArguments
    If (m_Restrictions.Exists(sName)) Then
        Restriction = m_Restrictions.Item(sName)
    Else
        Restriction = False
    End If
End Function

Private Function CheckForAnyFlags(sNeeded As String, sHave As String) As Boolean
On Error GoTo ERROR_HANDLER
    Dim i As Integer
    CheckForAnyFlags = False
    
    If (LenB(sHave) = 0) Then Exit Function
    
    For i = 1 To Len(sNeeded)
        If (InStr(1, sHave, Mid(sNeeded, i, 1), vbTextCompare) > 0) Then
            CheckForAnyFlags = True
            Exit Function
        End If
    Next

    Exit Function
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.CheckForAnyFlags().")
End Function

'Adds a line to the response queue
Public Sub Respond(strResponse As Variant)
    If (LenB(strResponse) > 0) Then m_Response.Add CStr(strResponse)
End Sub

'Cleares the response queue
Public Sub ClearResponse()
    Do While m_Response.Count
      m_Response.Remove 1
    Loop
End Sub

'Gets the Response Queue
Public Function GetResponse() As Collection
    Set GetResponse = m_Response
End Function

'This will respond in the proper style based on how the command was used.
'This will messup emote responses /me if it is not public output or whisper commands is turned on.
'If your response MUST be a specific style, Use AddQ or DSP
Public Function SendResponse()
On Error GoTo ERROR_HANDLER
    Dim i As Integer
    If (IsLocal) Then
        If (PublicOutput) Then
            For i = 1 To m_Response.Count
                frmChat.AddQ m_Response.Item(i), enuPriority.CONSOLE_MESSAGE
            Next i
        Else
            For i = 1 To m_Response.Count
                frmChat.AddChat g_Color.ConsoleText, m_Response.Item(i)
            Next i
        End If
    Else
        If ((BotVars.WhisperCmds Or WasWhispered) And (PublicOutput = False)) Then
            For i = 1 To m_Response.Count
                frmChat.AddQ "/w " & Username & Space$(1) & m_Response.Item(i), enuPriority.COMMAND_RESPONSE_MESSAGE
            Next i
        Else
            For i = 1 To m_Response.Count
                frmChat.AddQ m_Response.Item(i), enuPriority.COMMAND_RESPONSE_MESSAGE
            Next i
        End If
    End If
    ClearResponse
    Exit Function
    
ERROR_HANDLER:
    Call frmChat.AddChat(vbRed, "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.Respond().")
End Function

Private Function CheckMatch(sExpression As String, sData As String, Optional CaseSensitive As Boolean = True) As Boolean
On Error GoTo ERROR_HANDLER
    Dim mRegExp As RegExp
    Set mRegExp = New RegExp
    mRegExp.Global = True
    mRegExp.Pattern = sExpression
    mRegExp.IgnoreCase = (Not CaseSensitive)
    CheckMatch = mRegExp.Test(sData)
    Set mRegExp = Nothing
    
    Exit Function
    
ERROR_HANDLER:
    frmChat.AddChat g_Color.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in clsCommandObj.CheckMatch()"
End Function
