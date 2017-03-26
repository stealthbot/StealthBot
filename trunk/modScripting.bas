Attribute VB_Name = "modScripting"
'/* Scripting.bas
' * ~~~~~~~~~~~~~~~~
' * StealthBot VBScript support module
' * ~~~~~~~~~~~~~~~~
' * Modified by Swent 11/06/2007
' * ~~~~~~~~~~~~~~~~
' * Update Ribose/2009-08-15
' */
Option Explicit

Public Type scObj
    SCModule As Module
    ObjName  As String
    ObjType  As String
    obj      As Object
End Type

Public Type scInc
    SCModule  As Module
    IncName   As String
    lineCount As Integer
End Type

Private m_arrObjs()         As scObj
Private m_objCount          As Integer
Private m_arrIncs()         As scInc
Private m_incCount          As Integer
Private m_sc_control        As ScriptControl
Private m_is_reloading      As Boolean
Private m_ExecutingMdl      As Module
Private m_TempMdlName       As String
Private m_IsEventError      As Boolean
Private VetoNextMessage     As Boolean
Private VetoNextPacket      As Boolean
Private m_ScriptObservers   As Collection
Private m_FunctionObservers As Collection
Private m_SystemDisabled    As Boolean
Private m_SCInitialized     As Boolean

Public Sub InitScriptControl(ByVal SC As ScriptControl)

    ' check whether the override is disabling the script system
    If m_SystemDisabled Then Exit Sub
    
    m_is_reloading = True

    SC.Reset

    frmChat.Inet.Cancel
    frmChat.scTimer.Enabled = False
    
    DestroyObjs

    If Config.ScriptingAllowUI Then
        SC.AllowUI = True
    End If

    '// Create scripting objects
    SC.addObject "ssc", SharedScriptSupport, True
    SC.addObject "scTimer", frmChat.scTimer
    SC.addObject "scInet", frmChat.Inet
    SC.addObject "BotVars", BotVars

    Set m_sc_control = SC
    
    Set m_ScriptObservers = New Collection
    Set m_FunctionObservers = New Collection
    
    m_is_reloading = False
    
    ' this will always be true after first successful load
    m_SCInitialized = True

End Sub

Public Sub LoadScripts()

    On Error GoTo ERROR_HANDLER

    Dim CurrentModule As Module
    Dim Paths         As Collection
    Dim strPath       As String
    Dim FileName      As String
    Dim fileExt       As String
    Dim i             As Integer
    Dim j             As Integer
    Dim str           As String
    Dim tmp           As String
    Dim res           As Boolean

    ' check whether the override is disabling the script system
    If m_SystemDisabled Then Exit Sub
    
    Set Paths = New Collection

    ' ********************************
    '      LOAD SCRIPTS
    ' ********************************
    
    ' set script folder path
    strPath = GetFolderPath("Scripts")
       
    ' ensure scripts folder exists
    If (LenB(Dir$(strPath)) > 0) Then
        ' grab initial script file name
        FileName = Dir$(strPath)
        
        ' grab script files
        ' note: if we don't enumerate this list prior to script loading,
        ' scripting errors can kill further script loading.
        Do While (FileName <> vbNullString)
            ' add script file to collection
            Paths.Add FileName
        
            ' grab next script file name
            FileName = Dir$()
        Loop

        ' Cycle through each of the files.
        For i = 1 To Paths.Count
            ' Does the file have the extension for a script?
            If (IsValidFileExtension(GetFileExtension(Paths(i)))) Then
                ' Add a new module to the script control.
                Set CurrentModule = _
                    m_sc_control.Modules.Add(m_sc_control.Modules.Count + 1)
                
                ' store the temporary name of the module for parsing errors
                m_TempMdlName = Paths(i)
                
                ' set executing module reference for parsing errors and module-specific functions
                Set m_ExecutingMdl = CurrentModule
                
                ' Load the file into the module.
                res = FileToModule(CurrentModule, strPath & Paths(i))
                
                ' set executing module to nothing
                Set m_ExecutingMdl = Nothing
                
                ' Does the script have a valid name?
                If (IsScriptNameValid(CurrentModule) = False) Then
                    ' No. Try to fix it.
                    GetScriptDictionary(CurrentModule)("Name") = CleanFileName(m_TempMdlName)
                
                    ' Is it valid now?
                    If (IsScriptNameValid(CurrentModule) = False) Then
                        ' No, disable it.
                        
                        frmChat.AddChat RTBColors.ErrorMessageText, "Scripting error: " & Paths(i) & " has been " & _
                            "disabled due to a naming issue."
                            
                        str = strPath & "\disabled\"
                        
                        MkDir str
                            
                        Kill str & Paths(i)
    
                        Name strPath & Paths(i) As str & Paths(i)
    
                        InitScriptControl m_sc_control
                        
                        LoadScripts
                        
                        Exit Sub
                    End If
                End If
            End If
        Next i
    End If
    
    InitMenus

    frmChat.AddChat RTBColors.SuccessText, "Scripts loaded."
        
    Exit Sub

ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (" & Err.Number & "): " & Err.Description & " in LoadScripts()."

End Sub

Private Function FileToModule(ByRef ScriptModule As Module, ByVal FilePath As String, Optional ByVal defaults As Boolean = True) As Boolean

    On Error GoTo ERROR_HANDLER

    Static strContent    As String
    Static includes      As Collection
    
    Dim strLine          As String
    Dim f                As Integer
    Dim blnCheckOperands As Boolean
    Dim blnKeepLine      As Boolean
    Dim blnScriptData    As Boolean
    Dim i                As Integer
    Dim lineCount        As Integer
    Dim blnIncIsValid    As Boolean
    Dim strInclude       As String

    blnCheckOperands = True
    
    f = FreeFile
    
    If (defaults) Then
        Set includes = New Collection
    End If

    Open FilePath For Input As #f
        Do While (EOF(f) = False)
            Line Input #f, strLine
            
            strLine = Trim$(strLine)
            
            ' default to keep line
            blnKeepLine = True
            
            If (Len(strLine) >= 1) Then
                If ((blnCheckOperands) And (Left$(strLine, 1) = "#")) Then
                    ' this line is a directive, parse and don't keep in code
                    blnKeepLine = False
                    If (InStr(1, strLine, " ") <> 0) Then
                        If (Len(strLine) >= 2) Then
                            Dim strCommand As String
                        
                            strCommand = _
                                LCase$(Mid$(strLine, 2, InStr(1, strLine, " ") - 2))
    
                            If (strCommand = "include") Then
                                If (Len(strLine) >= 12) Then
                                    Dim strPath As String
                                    Dim strFullPath As String
                                    
                                    strPath = _
                                        Mid$(strLine, 11, Len(strLine) - 11)
                                    
                                    If (Left$(strPath, 1) = "\") Then
                                        strFullPath = StringFormat("{0}\Scripts{1}", CurDir$(), strPath)
                                    Else
                                        strFullPath = strPath
                                    End If
                                    
                                    blnIncIsValid = True
                                    
                                    ' check if file exists to include
                                    If LenB(Dir$(strFullPath)) = 0 Then
                                        frmChat.AddChat RTBColors.ErrorMessageText, "Scripting warning! " & Dir$(FilePath) & " is trying to include " & _
                                            "a file that does not exist: " & strPath
                                        blnIncIsValid = False
                                    End If
                                    
                                    ' check if file is already included by this script
                                    For i = 1 To includes.Count
                                        If StrComp(includes(i), strFullPath, vbTextCompare) = 0 Then
                                            frmChat.AddChat RTBColors.ErrorMessageText, "Scripting warning! " & Dir$(FilePath) & " is trying to include " & _
                                                "a file that has already been included: " & strPath
                                            blnIncIsValid = False
                                        End If
                                    Next i
                                    
                                    If blnIncIsValid Then
                                        ' store the file path to an include
                                        includes.Add strFullPath
                                    End If
                                End If
                            End If
                        End If
                    End If
                ElseIf (((LenB(strLine) > 0) And _
                         (StrComp(Left$(strLine, 1), "'") <> 0)) Or _
                        ((LenB(strLine) = 3) And _
                         (StrComp(strLine, "rem", vbTextCompare) <> 0)) Or _
                        ((LenB(strLine) > 3) And _
                         (StrComp(Left$(strLine, 3), "rem", vbTextCompare) <> 0) And _
                         (InStr(1, Mid$(strLine, 4, 1), "abcdefghijklmnopqrstuvwxyz01243567890_", vbTextCompare) = 0))) _
                       Then
                    ' this line is not a comment or blank line, stop checking for #include
                    blnCheckOperands = False
                End If
            End If
            
            ' if this line is not an #include
            If (blnKeepLine) Then
                ' keep it, append it to content
                strContent = _
                    strContent & strLine & vbCrLf
            Else
                ' remove it, but keep line numbers consistant in errors
                strContent = _
                    strContent & vbCrLf
            End If
            
            ' increment counter
            lineCount = lineCount + 1
            
            ' clean up
            strLine = vbNullString
        Loop
    Close #f
  
    ' if we are not loading an include, set it up
    If (defaults) Then
        ' store module-level functions
        ScriptModule.ExecuteStatement GetDefaultModuleProcs()
        
        ' initialize the variables globally
        ScriptModule.ExecuteStatement "Public Script, DataBuffer"
        
        ' store Script dictionary into script
        SetScriptDictionary ScriptModule
        
        ' create default DataBuffer object
        Set ScriptModule.CodeObject.DataBuffer = SharedScriptSupport.DataBufferEx()
        
        ' set the Script("Path") value
        GetScriptDictionary(ScriptModule)("Path") = FilePath
        
        ' add this as the "first" include for this script, with no name so that later it is changed to "scriptname"
        AddInclude ScriptModule, vbNullString, lineCount
        
        ' add includes to end of code, process the same for includes in includes
        i = 1
        Do While i <= includes.Count
            FileToModule ScriptModule, includes(i), False
            i = i + 1
        Loop
        
        ' store the position in the arrInc array where this script's line number information starts
        ' the line number information will be in order of #include directives
        ' the error handler will use this information to continually subtract linecounts until
        ' the line number is in the bounds of the include information where the file actually is
        ' so that we know what include is causing an error
        GetScriptDictionary(ScriptModule)("IncludeLBound") = m_incCount - includes.Count - 1
        
        ' add content
        ScriptModule.AddCode strContent
        
        ' clean up object
        Set includes = Nothing
        
        ' clean up content for next call to this function
        strContent = vbNullString
    Else
        ' this is an #include, store our information
        AddInclude ScriptModule, "#" & Mid$(FilePath, InStrRev(FilePath, "\", -1, vbBinaryCompare) + 1), lineCount
    End If
    
    FileToModule = True
    
    Exit Function
    
ERROR_HANDLER:

    strContent = vbNullString
    
    Set includes = Nothing

    FileToModule = False

End Function

' stores information about an include for more accurate error handling!
Private Sub AddInclude(ByRef SCModule As Module, ByVal Name As String, ByVal Lines As Integer)
    Dim inc As scInc
    
    ' redefine array size
    If (m_incCount) Then
        ReDim Preserve m_arrIncs(0 To m_incCount)
    Else
        ReDim m_arrIncs(0)
    End If
    
    ' store our #include name, path, and length
    inc.IncName = Name
    inc.lineCount = Lines

    ' store our module handle
    Set inc.SCModule = SCModule
    
    ' store our inc
    m_arrIncs(m_incCount) = inc
    
    ' increment include counter
    m_incCount = (m_incCount + 1)
End Sub

Private Function GetDefaultModuleProcs() As String

    Dim str As String ' storage buffer for module code

    ' GetModuleID() module-level function
    str = str & "Function GetModuleID()" & vbNewLine
    str = str & "   GetModuleID = SSC.GetModuleID(Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' GetScriptModule() module-level function
    str = str & "Function GetScriptModule()" & vbNewLine
    str = str & "   Set GetScriptModule = SSC.GetScriptModule(Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' GetWorkingDirectory() module-level function
    str = str & "Function GetWorkingDirectory()" & vbNewLine
    str = str & "   GetWorkingDirectory = SSC.GetWorkingDirectory(Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' CreateObj() module-level function
    str = str & "Function CreateObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Set CreateObj = _ " & vbNewLine
    str = str & "         SSC.CreateObj(ObjType, ObjName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' DestroyObj() module-level function
    str = str & "Sub DestroyObj(ObjName)" & vbNewLine
    str = str & "   SSC.DestroyObj ObjName, Script(""Name"")" & vbNewLine
    str = str & "End Sub" & vbNewLine

    ' GetObjByName() module-level function
    str = str & "Function GetObjByName(ObjName)" & vbNewLine
    str = str & "   Set GetObjByName = _ " & vbNewLine
    str = str & "         SSC.GetObjByName(ObjName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' GetSettingsEntry() module-level function
    str = str & "Function GetSettingsEntry(EntryName)" & vbNewLine
    str = str & "   GetSettingsEntry = SSC.GetSettingsEntry(EntryName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' IsEnabled() module-level function
    str = str & "Function IsEnabled()" & vbNewLine
    str = str & "   IsEnabled = (StrComp(GetSettingsEntry(""Enabled""), ""false"", vbTextCompare) <> 0)" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' WriteSettingsEntry() module-level function
    str = str & "Sub WriteSettingsEntry(EntryName, EntryValue)" & vbNewLine
    str = str & "   SSC.WriteSettingsEntry EntryName, EntryValue, , Script(""Name"")" & vbNewLine
    str = str & "End Sub" & vbNewLine

    ' CreateCommand() module-level function
    str = str & "Function CreateCommand(commandName)" & vbNewLine
    str = str & "   Set CreateCommand = _ " & vbNewLine
    str = str & "         SSC.CreateCommand(commandName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' OpenCommand() module-level function
    str = str & "Function OpenCommand(commandName)" & vbNewLine
    str = str & "   Set OpenCommand = _ " & vbNewLine
    str = str & "         SSC.OpenCommand(commandName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' DeleteCommand() module-level function
    str = str & "Function DeleteCommand(commandName)" & vbNewLine
    str = str & "   Set DeleteCommand = _ " & vbNewLine
    str = str & "         SSC.DeleteCommand(commandName, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' GetCommands() module-level function
    str = str & "Function GetCommands()" & vbNewLine
    str = str & "   Set GetCommands = _ " & vbNewLine
    str = str & "         SSC.GetCommands(Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' store module-level coding
    GetDefaultModuleProcs = str

End Function

Private Function IsScriptNameValid(ByRef CurrentModule As Module) As Boolean

    On Error Resume Next

    Dim j            As Integer
    Dim str          As String
    Dim tmp          As String
    Dim nameDisallow As String
    
    str = GetScriptName(CurrentModule.Name)

    If (str = vbNullString) Then
        IsScriptNameValid = False
                
        Exit Function
    End If
    
    nameDisallow = "\/:*?<>|"""

    For j = 1 To Len(str)
        If (InStr(1, nameDisallow, Mid$(str, j, 1), vbTextCompare) <> 0) Then
            IsScriptNameValid = False
    
            Exit Function
        End If
    Next j

    For j = 2 To m_sc_control.Modules.Count
        If (m_sc_control.Modules(j).Name <> CurrentModule.Name) Then
            tmp = GetScriptName(CStr(j))
                
            If (StrComp(str, tmp, vbTextCompare) = 0) Then
                IsScriptNameValid = False

                Exit Function
            End If
        End If
    Next j
    
    IsScriptNameValid = True

End Function

Public Sub InitScripts()

    On Error Resume Next
    
    Static reloading As Boolean
    
    Dim i   As Integer
    Dim tmp As String
    ' check whether the override is disabling the script system
    If m_SystemDisabled Then Exit Sub
    
    If (reloading = False) Then
        RunInAll "Event_FirstRun"
        
        reloading = True
    End If
    
    For i = 2 To m_sc_control.Modules.Count
        If (i > 1) Then
            tmp = m_sc_control.Modules(i).CodeObject.GetSettingsEntry("Enabled")
        End If

        If (StrComp(tmp, "False", vbTextCompare) <> 0) Then
            InitScript m_sc_control.Modules(i)
        End If
    Next i

End Sub

Public Sub InitScript(ByVal SCModule As Module)

    On Error Resume Next

    Dim i          As Integer
    Dim startTime  As Long
    Dim finishTime As Long

    startTime = GetTickCount()
    
    RunInSingle SCModule, "Event_Load"
    
    finishTime = GetTickCount()
 
    '// 03/27/2009 52 - added default Script property for the load time
    GetScriptDictionary(SCModule)("InitPerf") = (finishTime - startTime)

    'If (g_Online) Then
    '    RunInSingle SCModule, "Event_LoggedOn", GetCurrentUsername, BotVars.Product
    '    RunInSingle SCModule, "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags
    '
    '    If (g_Channel.Users.Count > 0) Then
    '        For i = 1 To g_Channel.Users.Count
    '            With g_Channel.Users(i)
    '                 RunInSingle SCModule, "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
    '                    .Game, False
    '            End With
    '         Next i
    '     End If
    'End If
End Sub

Public Function RunInAll(ParamArray Parameters() As Variant) As Boolean

    On Error Resume Next

    Dim SC      As ScriptControl
    Dim i       As Integer
    Dim arr()   As Variant
    Dim str     As String
    Dim oldVeto As Boolean
    Dim veto    As Boolean
    Dim oldEM   As Module
    Dim obj     As Module
    Dim Proc    As Procedure

    ' check whether the override is disabling the script system
    If m_SystemDisabled Then Exit Function

    Set SC = m_sc_control
    
    If (m_is_reloading) Then
        Exit Function
    End If
    
    oldVeto = GetVeto 'Keeps the old veto, for recursion, and Sets to false.
    veto = False      'Just to be sure
    
    Set oldEM = m_ExecutingMdl ' keep last reference, for recursion

    arr() = Parameters()

    For i = 2 To SC.Modules.Count
        Set obj = SC.Modules(i)
        Set m_ExecutingMdl = obj
        
        str = obj.CodeObject.GetSettingsEntry("Enabled")
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            
            ' check if module has the procedure before calling it!
            'For Each Proc In obj.Procedures
            '    If StrComp(Proc.Name, arr(0), vbTextCompare) = 0 Then
            '        ' it does, count args
            '        If Proc.numArgs = UBound(arr) Then
            '            ' call it
            CallByNameEx obj, "Run", VbMethod, arr()
            '        End If
            '        Exit For
            '    End If
            'Next Proc
            
            veto = veto Or GetVeto 'Did they veto it, or was it already vetod?
        End If
        Set obj = Nothing
        Set m_ExecutingMdl = Nothing
    Next
    
    Set m_ExecutingMdl = oldEM ' return to last reference
    
    SetVeto oldVeto 'Reset the old veto, this is for recursion
    RunInAll = veto 'Was this particular event vetoed?
    
    ' if this is the outermost of a recursion, make sure errors clear after this
    If m_ExecutingMdl Is Nothing Then
        m_sc_control.Error.Clear
    End If
End Function

Public Function RunInSingle(ByRef obj As Module, ParamArray Parameters() As Variant) As Boolean

    On Error Resume Next

    Dim i       As Integer
    Dim x       As Integer
    Dim arr()   As Variant
    Dim str     As String
    Dim oldVeto As Boolean
    Dim oldEM   As Module
    Dim Proc    As Procedure
    Dim sobsers As Collection
    Dim fobsers As Collection
    Dim obser   As Module
    Dim cobser  As Boolean
    Dim mname   As String

    ' check whether the override is disabling the script system
    If m_SystemDisabled Then Exit Function
        
    If (m_is_reloading) Then
        Exit Function
    End If
    
    oldVeto = GetVeto  'Keeps the old veto, for recursion, and Sets to false.
    
    Set oldEM = m_ExecutingMdl ' keep old module reference, for recursion
    
    Set m_ExecutingMdl = obj
    m_IsEventError = (StrComp(arr(0), "Event_Error", vbBinaryCompare) = 0)

    arr() = Parameters()
    
    'If Obj is nothing then we are 'running' an internal event
    'This is so scriptors can observe internal events, like Internal Commands
    If (obj Is Nothing) Then
        str = "True"
        mname = vbNullString
    Else
        str = obj.CodeObject.GetSettingsEntry("Enabled")
        mname = GetScriptName(obj.Name)
    End If
    
    If (Not StrComp(str, "False", vbTextCompare) = 0) Then
        If (Not obj Is Nothing) Then CallByNameEx obj, "Run", VbMethod, arr()
        RunInSingle = GetVeto 'Was this particular event vetoed?
        
        'Call any scripts that are observing this one
        Set sobsers = GetScriptObservers(mname, False)
        
        For i = 1 To sobsers.Count
            Set obser = GetModuleByName(sobsers.Item(i))
            If (Not obser Is Nothing) Then 'Is the script real/loaded?
                str = obser.CodeObject.GetSettingsEntry("Enabled")
                If (Not StrComp(str, "False", vbTextCompare) = 0) Then 'Is it off?
                    Set m_ExecutingMdl = obser
                    CallByNameEx obser, "Run", VbMethod, arr
                End If
            End If
            Set obser = Nothing
        Next i
        
        Set fobsers = GetFunctionObservers(CStr(arr(0)))
        
        For i = 1 To fobsers.Count
            cobser = True
            If (StrComp(fobsers.Item(i), mname, vbTextCompare) = 0) Then 'Dont call itself
                cobser = False
            Else
                For x = 1 To sobsers.Count 'See if we already called it with Script Observers
                    If (StrComp(sobsers.Item(x), fobsers.Item(i), vbTextCompare) = 0) Then
                        cobser = False
                        Exit For
                    End If
                Next x
            End If
            
            If (cobser) Then
                Set obser = GetModuleByName(fobsers.Item(i))
                If (Not obser Is Nothing) Then
                    str = obser.CodeObject.GetSettingsEntry("Enabled")
                    If (Not StrComp(str, "False", vbTextCompare) = 0) Then
                        Set m_ExecutingMdl = obser
                        CallByNameEx obser, "Run", VbMethod, arr
                    End If
                End If
                Set obser = Nothing
            End If
        Next i
    End If
    
    m_IsEventError = False
    Set m_ExecutingMdl = oldEM ' got back to old reference
    SetVeto oldVeto 'Reset the old veto, this is for recursion
    
    ' if this is the outermost of a recursion, make sure errors clear after this
    If m_ExecutingMdl Is Nothing Then
        m_sc_control.Error.Clear
    End If
End Function

Public Sub CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray _
    As Variant)
    
    On Error GoTo ERROR_HANDLER
    
    Dim oTLI    As TLI.TLIApplication
    Dim ProcID  As Long
    Dim numArgs As Long
    Dim i       As Long
    Dim v()     As Variant

    Set oTLI = New TLIApplication

    ProcID = oTLI.InvokeID(obj, ProcName)

    If (IsMissing(vArgsArray)) Then
        Call oTLI.InvokeHook(obj, ProcID, CallType)
    End If
    
    If (IsArray(vArgsArray)) Then
        numArgs = UBound(vArgsArray)
        
        ReDim v(numArgs)
        
        For i = 0 To numArgs
            ' corrected object passing -Ribose/2009-08-10
            If IsObject(vArgsArray(numArgs - i)) Then
                Set v(i) = vArgsArray(numArgs - i)
            Else
                Let v(i) = vArgsArray(numArgs - i)
            End If
        Next i
        
        Call oTLI.InvokeHookArray(obj, ProcID, CallType, v)
    End If
    
    Set oTLI = Nothing

    Exit Sub

ERROR_HANDLER:

    If (frmChat.SControl.Error) Then
        Exit Sub
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in CallByNameEx()."
        
    Set oTLI = Nothing
    
End Sub

Public Function Objects(objIndex As Integer) As scObj

    Objects = m_arrObjs(objIndex)

End Function

Private Function ObjCount(Optional ObjType As String, Optional ByVal SCModule As Module = Nothing) As Integer
    
    Dim i As Integer

    If (ObjType <> vbNullString) Then
        For i = 0 To m_objCount - 1
            If (SCModule Is Nothing) Then
                If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
                    ObjCount = (ObjCount + 1)
                End If
            Else
                If (StrComp(SCModule.Name, m_arrObjs(i).SCModule.Name) = 0) Then
                    If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
                        ObjCount = (ObjCount + 1)
                    End If
                End If
            End If
        Next i
    Else
        ObjCount = m_objCount
    End If

End Function

Public Function CreateObj(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    On Error Resume Next

    Dim obj As scObj
    Dim ScriptName As String
    
    If SCModule Is Nothing Then Exit Function
    
    Set CreateObj = Nothing
    If (Not ValidObjectName(ObjName)) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Scripting error: The variable name provided to CreateObj was not valid."
        Exit Function
    End If
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim i As Integer ' loop counter variable

        For i = 0 To m_objCount - 1
            If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
                If (StrComp(m_arrObjs(i).ObjType, ObjType, vbTextCompare) = 0) Then
                    If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                        Set CreateObj = m_arrObjs(i).obj
                    
                        Exit Function
                    End If
                End If
            End If
        Next i
        
        ReDim Preserve m_arrObjs(0 To m_objCount)
    Else
        ReDim m_arrObjs(0)
    End If

    ' store our module name & type
    obj.ObjName = ObjName
    obj.ObjType = ObjType

    ' store our module handle
    Set obj.SCModule = SCModule
    
    ScriptName = GetScriptName(SCModule.Name)
    
    ' grab/create instance of object
    Select Case (UCase$(ObjType))
        Case "TIMER"
            If (ObjCount(ObjType) > 0) Then
                Load frmChat.tmrScript(ObjCount(ObjType))
            End If
            
            Set obj.obj = _
                    frmChat.tmrScript(ObjCount(ObjType))
                    
        Case "LONGTIMER"
            If (ObjCount(ObjType) > 0) Then
                Load frmChat.tmrScriptLong(ObjCount(ObjType))
            End If
        
            Set obj.obj = New clsSLongTimer
        
            obj.obj.tmr = _
                frmChat.tmrScriptLong(ObjCount(ObjType))
                    
            With obj.obj.tmr
                .Interval = 1000
            End With
            
        Case "WINSOCK"
            If (ObjCount(ObjType) > 0) Then
                Load frmChat.sckScript(ObjCount(ObjType))
            End If
            
            Set obj.obj = _
                    frmChat.sckScript(ObjCount(ObjType))
        
        Case "INET"
            If (ObjCount(ObjType) > 0) Then
                Load frmChat.itcScript(ObjCount(ObjType))
            End If
            
            Set obj.obj = _
                    frmChat.itcScript(ObjCount(ObjType))
        
        Case "FORM"
            Set obj.obj = New frmScript
            
            obj.obj.SetName ObjName
            obj.obj.SetSCModule SCModule
            
            HookWindowProc obj.obj.hWnd
            
        Case "MENU"
            ' check if there are no menus and we're adding one
            If (ObjCount("Menu", SCModule) = 0) Then
                Dim tmp As clsMenuObj
                ' get the dynamic menu dash
                Set tmp = DynamicMenus("dashmnu" & ScriptName)
                ' show it
                tmp.Visible = True
                tmp.Caption = "-"
            End If
        
            Set obj.obj = New clsMenuObj
            
            obj.obj.Name = ObjName
            
            obj.obj.Parent = _
                DynamicMenus("mnu" & ScriptName)
                
            DynamicMenus.Add obj.obj
    End Select

    ' store object
    m_arrObjs(m_objCount) = obj
    
    ' increment object counter
    m_objCount = (m_objCount + 1)
    
    ' create class variable for object
    SCModule.ExecuteStatement "Set " & ObjName & " = GetObjByName(" & Chr$(34) & _
        ObjName & Chr$(34) & ")"
    
    ' unfortunately creating a new form triggers the scripting events
    ' too early, so we have to call them manually here.
    If (UCase$(ObjType) = "FORM") Then
        obj.obj.Initialize
    End If

    ' return object
    Set CreateObj = obj.obj

End Function

Public Sub DestroyObjs(Optional ByVal SCModule As Object = Nothing)

    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    For i = m_objCount - 1 To 0 Step -1
        If (SCModule Is Nothing) Then
            DestroyObj m_arrObjs(i).SCModule, m_arrObjs(i).ObjName
        Else
            If (SCModule.Name = m_arrObjs(i).SCModule.Name) Then
                DestroyObj m_arrObjs(i).SCModule, m_arrObjs(i).ObjName
            End If
        End If
    Next i
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in DestroyObjs()."
        
    Resume Next
    
End Sub

Public Sub DestroyObj(ByVal SCModule As Module, ByVal ObjName As String)

    On Error GoTo ERROR_HANDLER

    Dim i     As Integer
    Dim Index As Integer
    
    If SCModule Is Nothing Then Exit Sub
    
    If (m_objCount = 0) Then
        Exit Sub
    End If
    
    Index = m_objCount
    
    For i = 0 To m_objCount - 1
        If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                Index = i
            
                Exit For
            End If
        End If
    Next i
    
    If (Index >= m_objCount) Then
        Exit Sub
    End If
    
    Select Case (UCase$(m_arrObjs(Index).ObjType))
        Case "TIMER"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload frmChat.tmrScript(m_arrObjs(Index).obj.Index)
            Else
                frmChat.tmrScript(0).Enabled = False
            End If
            
        Case "LONGTIMER"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload frmChat.tmrScriptLong(m_arrObjs(Index).obj.Index)
            Else
                frmChat.tmrScriptLong(0).Enabled = False
            End If
            
        Case "WINSOCK"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload frmChat.sckScript(m_arrObjs(Index).obj.Index)
            Else
                frmChat.sckScript(0).Close
            End If
            
        Case "INET"
            If (m_arrObjs(Index).obj.Index > 0) Then
                Unload frmChat.itcScript(m_arrObjs(Index).obj.Index)
            Else
                frmChat.itcScript(0).Cancel
            End If

        Case "FORM"
            UnhookWindowProc m_arrObjs(Index).obj.hWnd
        
            m_arrObjs(Index).obj.DestroyObjs
            
            Unload m_arrObjs(Index).obj
            
        Case "MENU"
            ' check if there is one menu left and we're destroying it
            If (ObjCount("Menu", SCModule) = 1) Then
                Dim IsForm As Boolean
                Dim tmp As clsMenuObj
                ' default to false
                IsForm = False
                ' get the menu
                Set tmp = m_arrObjs(Index).obj
                ' get root menu
                Do While StrComp(Right$(tmp.Name, 4), "ROOT", vbBinaryCompare) <> 0
                    ' check if this has a window as a parent instead
                    If StrComp(TypeName$(tmp.Parent), "frmScript", vbBinaryCompare) = 0 Then
                        ' yes, get outta here!
                        IsForm = True
                        Exit Do
                    End If
                    Set tmp = tmp.Parent
                Loop
                ' check again so that the dash doesn't get hidden
                If Not IsForm Then
                    ' get dynamic dash from caption of root
                    Set tmp = DynamicMenus("dashmnu" & tmp.Caption)
                    ' show it
                    tmp.Visible = False
                End If
            End If
            
            m_arrObjs(Index).obj.Class_Terminate
    End Select

    Set m_arrObjs(Index).obj = Nothing
    
    If (Index < m_objCount - 1) Then
        For i = Index To ((m_objCount - 1) - 1)
            m_arrObjs(i) = m_arrObjs(i + 1)
        Next i
    End If
    
    If (m_objCount > 1) Then
        ReDim Preserve m_arrObjs(0 To m_objCount - 1)
    Else
        ReDim m_arrObjs(0)
    End If
    
    SCModule.ExecuteStatement "Set " & ObjName & " = Nothing"
    
    m_objCount = (m_objCount - 1)
    
    Exit Sub
    
ERROR_HANDLER:

    ' scripting engine has been reset - likely due to a reload
    If (Err.Number = -2147467259) Then
        Resume Next
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, _
        "Error (#" & Err.Number & "): " & Err.Description & " in DestroyObj()."
        
    Resume Next
    
End Sub

Public Function GetObjByName(ByRef SCModule As Module, ByVal ObjName As String) As Object

    Dim i As Integer
    
    If SCModule Is Nothing Then Exit Function
    
    For i = 0 To m_objCount - 1
        If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                Set GetObjByName = m_arrObjs(i).obj

                Exit Function
            End If
        End If
    Next i

End Function

Public Function GetScriptObjByMenuID(ByVal MenuID As Long) As scObj

    Dim i As Integer
    Dim j As Integer

    For i = 0 To ObjCount() - 1
        If (StrComp("Menu", Objects(i).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(i).obj.ID = MenuID) Then
                GetScriptObjByMenuID = m_arrObjs(i)
                
                Exit Function
            End If
        'ElseIf (StrComp("Form", Objects(I).ObjType, vbTextCompare) = 0) Then
        '    For j = 0 To Objects(I).obj.ObjCount("Menu") - 1
        '        If (StrComp("Menu", Objects(I).obj.Objects(j).ObjectType, vbTextCompare) = 0) Then
        '            If (Objects(I).obj.Objects(j).ID = MenuID) Then
        '                GetScriptObjByMenuID = Objects(I)
        '
        '                Exit Function
        '            End If
        '        End If
        '    Next j
        End If
    Next i

End Function

Public Function GetScriptObjByIndex(ByVal ObjType As String, ByVal Index As Integer) As scObj

    Dim i As Integer

    For i = 0 To ObjCount() - 1
        If (StrComp(ObjType, Objects(i).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(i).obj.Index = Index) Then
                GetScriptObjByIndex = m_arrObjs(i)
                
                Exit For
            End If
        End If
    Next i

End Function

Public Function InitMenus()

    On Error GoTo ERROR_HANDLER

    Dim tmp  As clsMenuObj
    Dim Name As String
    Dim i    As Integer

    ' destroy all the menus and start over
    DestroyMenus
    
    ' for each script add menus
    For i = 2 To frmChat.SControl.Modules.Count
        If (i = 2) Then
            frmChat.mnuScriptingDash(0).Visible = True
        End If
    
        ' name is the script name at this module index
        Name = GetScriptName(CStr(i))
    
        ' new menu
        Set tmp = New clsMenuObj
    
        ' root menu, give name, hwnd, and caption
        tmp.Name = vbNullChar & Name & vbNullChar & "ROOT"
        tmp.hWnd = GetSubMenu(GetMenu(frmChat.hWnd), 5)
        tmp.Caption = Name
            
        ' add with script name as key for adding more menus
        DynamicMenus.Add tmp, "mnu" & Name
            
        ' new menu
        Set tmp = New clsMenuObj
    
        ' enable/disable menu, give name, parent, and caption
        tmp.Name = vbNullChar & Name & vbNullChar & "ENABLE|DISABLE"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "Enabled"
        
        If (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), _
                "False", vbTextCompare) <> 0) Then
                
            tmp.Checked = True
        End If
        
        ' add it
        DynamicMenus.Add tmp
        
        ' new menu
        Set tmp = New clsMenuObj
    
        ' view script menu item, give name, parent, and caption
        tmp.Name = vbNullChar & Name & vbNullChar & "VIEW_SCRIPT"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "View Script"
        
        ' add it
        DynamicMenus.Add tmp
        
        ' new menu
        Set tmp = New clsMenuObj
        
        ' dash, give name, parent, caption, and hide it
        tmp.Name = vbNullChar & Name & vbNullChar & "DASH"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "-"
        tmp.Visible = False
        
        ' give it a key so it can't be confused with another script's main menu or menu dash
        DynamicMenus.Add tmp, "dashmnu" & Name
    Next i
    
    Exit Function
        
ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in InitMenus()."

    Err.Clear
    
    Resume Next

End Function

Public Function DestroyMenus()

    On Error GoTo ERROR_HANDLER

    Dim i As Integer
    
    frmChat.mnuScriptingDash(0).Visible = False
    
    For i = DynamicMenus.Count To 1 Step -1
        
        If (Len(DynamicMenus(i).Name) > 0) Then
            If (Left$(DynamicMenus(i).Name, 1) = vbNullChar) Then
                DynamicMenus(i).Class_Terminate

                DynamicMenus.Remove i
            End If
        End If
        
    Next i
    
    Exit Function
    
ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in DestroyMenus()."

    Err.Clear
    
    Resume Next
    
End Function

Public Function Scripts() As Object

    On Error Resume Next

    Dim i   As Integer
    Dim str As String
    Dim SCModule As Module
    Dim ScriptName As String

    Set Scripts = New Collection

    For i = 2 To frmChat.SControl.Modules.Count
        ScriptName = GetScriptName(CStr(i))
        
        str = GetScriptModule.CodeObject.GetSettingsEntry("Public")
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            Scripts.Add frmChat.SControl.Modules(i).CodeObject, ScriptName
        End If
    Next i

End Function

Public Function GetModuleByName(ByVal ScriptName As String) As Module
    Dim i As Integer
    
    For i = 2 To frmChat.SControl.Modules.Count
        If (StrComp(GetScriptName(CStr(i)), ScriptName, vbTextCompare) = 0) Then
            Set GetModuleByName = frmChat.SControl.Modules(i)
            Exit Function
        End If
    Next i
End Function

Public Sub SetVeto(ByVal b As Boolean)

    VetoNextMessage = b
    
End Sub

Public Function GetVeto() As Boolean

    GetVeto = VetoNextMessage
    
    VetoNextMessage = False
    
End Function

Private Function GetFileExtension(ByVal FileName As String)
        
    Dim arr() As String

    arr = Split(FileName, ".")
    
    If UBound(arr) = 0 Then
        GetFileExtension = ""
    Else
        GetFileExtension = arr(UBound(arr))
    End If

End Function

Private Function IsValidFileExtension(ByVal ext As String) As Boolean

    Dim exts() As String
    Dim i      As Integer

    ReDim exts(0 To 1)
    
    'exts(0) = "dat"
    exts(0) = "txt"
    exts(1) = "vbs"
    
    For i = LBound(exts) To UBound(exts)
        If (StrComp(ext, exts(i), vbTextCompare) = 0) Then
            IsValidFileExtension = True
            
            Exit Function
        End If
    Next i
    
    IsValidFileExtension = False

End Function

Private Function CleanFileName(ByVal FileName As String) As String
    
    On Error Resume Next
    
    If (InStr(1, FileName, ".") > 1) Then
        CleanFileName = _
            Left$(FileName, InStr(1, FileName, ".") - 1)
    End If

End Function

'06/26/09 - Hdx Very crappy function to check if Object names are valid, a-z0-9_ and 1st chr a-z (eventually should be a regexp)
Public Function ValidObjectName(sName As String) As Boolean
    Dim x As Integer
    Dim sValid As String
    
    sValid = "abcdefghijklmnopqrstuvwxyz0123456789_"
    ValidObjectName = False
    
    For x = 1 To Len(sName)
        If (InStr(1, Left$(sValid, IIf(x = 1, 26, 37)), Mid$(sName, x, 1), vbTextCompare) = 0) Then Exit Function
    Next x
  
    ValidObjectName = True
End Function

Public Function ConvertStringArray(sArr() As String) As Variant()
    Dim vArr() As Variant
    Dim x As Integer
    ReDim vArr(LBound(sArr) To UBound(sArr))
    For x = LBound(sArr) To UBound(sArr)
        vArr(x) = CVar(sArr(x))
    Next x
    ConvertStringArray = vArr
End Function

Public Sub SC_Error()
    
    Dim Name        As String
    Dim ErrType     As String
    Dim Number      As Long
    Dim Description As String
    Dim Line        As Long
    Dim Column      As Long
    Dim Source      As String
    Dim Text        As String
    Dim IncIndex    As Integer
    Dim i           As Integer
    Dim tmp         As String
    
    ' check whether the override is disabling the script system
    ' (this function is being called in that case due to /exec)
    If m_SystemDisabled Then Exit Sub
    
    With m_sc_control
        Number = .Error.Number
        Description = .Error.Description
        Line = .Error.Line
        Column = .Error.Column
        Source = .Error.Source
        Text = .Error.Text
    End With
    
    If m_ExecutingMdl Is Nothing Then
        Exit Sub ' exec error handler will handle this
    Else
        ' start at the stored include index
        IncIndex = GetScriptDictionary(m_ExecutingMdl)("IncludeLBound")
        ' loop until we have reached a line number within this include's bounds
        For i = IncIndex To UBound(m_arrIncs)
            If Not m_ExecutingMdl Is m_arrIncs(i).SCModule Then
                ' we are no longer looping through this script's includes-- line number is too large or index is invalid
                Name = GetScriptName(m_ExecutingMdl.Name)
                Exit For
            ElseIf Line <= m_arrIncs(i).lineCount Then
                ' it is this script which is erroring
                Name = GetScriptName(m_ExecutingMdl.Name) & m_arrIncs(i).IncName
                Exit For
            Else
                ' this include did not contain this line
                ' decrease the line number the error was found by this include's line count
                Line = Line - m_arrIncs(i).lineCount
            End If
        Next i
        If LenB(Name) = 0 Or Left$(Name, 1) = "#" Then Name = m_TempMdlName & Name
    End If
    
    ' default to runtime error
    ErrType = "runtime"
    
    ' check if its a parsing error
    If InStr(1, Source, "compilation", vbBinaryCompare) > 0 Then
        ErrType = "parsing"
    End If
    
    ' check if the script is planning to handle errors itself, if Script("HandleErrors") = True, then call event_error
    If ((StrComp(GetScriptDictionary(m_ExecutingMdl)("HandleErrors"), _
                 "True", vbTextCompare) = 0) And _
                 (m_IsEventError = False)) Then
        ' call Event_Error(Number, Description, Line, Column, Text, Source)
        If (RunInSingle(m_ExecutingMdl, "Event_Error", Number, Description, Line, Column, Text, Source) = True) Then
            ' if vetoed, exit
            Exit Sub
        End If
    End If
    
    ' display error if script enabled
    If InStr(Name, "#") > 0 Then
        tmp = Left$(Name, InStr(Name, "#") - 1)
    Else
        tmp = Name
    End If
    tmp = SharedScriptSupport.GetSettingsEntry("Enabled", CleanFileName(tmp))
    
    If (StrComp(tmp, "False", vbTextCompare) <> 0) Then
        frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Scripting {0} error '{1}' in {2}: (line {3}; column {4})", _
            ErrType, Number, Name, Line, Column)
        frmChat.AddChat RTBColors.ErrorMessageText, Description
        If LenB(Trim$(Text)) > 0 Then
            frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Offending line: >> {0}", Text)
        End If
    End If
    
    m_sc_control.Error.Clear
    
End Sub

' get currently executing module (can be nothing!)
Public Function GetScriptModule(Optional ByVal ScriptName As String = vbNullString) As Module

    On Error GoTo ERROR_HANDLER
    
    Dim i As Integer
    
    ' only loop if the scriptname was provided
    If LenB(ScriptName) > 0 Then
        ' loop through modules
        For i = 2 To m_sc_control.Modules.Count
            ' if Script("Name") = ScriptName then
            If StrComp(GetScriptName(CStr(i)), ScriptName, vbTextCompare) = 0 Then
                ' return this module
                Set GetScriptModule = m_sc_control.Modules(i)
                
                Exit Function
            End If
        Next i
    Else
        ' return currently executing
        Set GetScriptModule = m_ExecutingMdl
    End If
    
    Exit Function

ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in GetScriptModule()."

    Err.Clear
    
    Resume Next

End Function

' get the module.Name for the provided script name
' returns the currently executing one if none provided
' if the currently executing one is nothing (/exec), "exec" is returned
Public Function GetModuleID(Optional ByVal ScriptName As String = vbNullString) As String

    On Error GoTo ERROR_HANDLER
    
    Dim i As Integer
    
    ' only loop if the scriptname was provided
    If LenB(ScriptName) > 0 Then
        ' loop through modules
        For i = 2 To m_sc_control.Modules.Count
            ' if Script("Name") = ScriptName then
            If StrComp(GetScriptName(CStr(i)), ScriptName, vbTextCompare) = 0 Then
                ' return this ID
                GetModuleID = m_sc_control.Modules(i).Name
                Exit Function
            End If
        Next i
    End If
    
    ' if this is the execute command, the current module is nothing
    If m_ExecutingMdl Is Nothing Then
        GetModuleID = "exec"
    Else
        ' return currently executing ID
        GetModuleID = m_ExecutingMdl.Name
    End If

    Exit Function

ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in GetModuleID()."

    Err.Clear
    
    Resume Next
    
End Function

' returns the script name from the Script() object for the provided module id
' if not found returns the module ID provided
' if none provided, uses the currently executing module ID
Public Function GetScriptName(Optional ByVal ModuleID As String = vbNullString) As String

    On Error GoTo ERROR_HANDLER
    
    Dim Module As Module
    
    ' if not provided
    If (LenB(ModuleID) = 0) Then
        ' use currently executing module ID
        ModuleID = GetModuleID
    End If
    
    If StrictIsNumeric(ModuleID) = False Then Exit Function
    
    Set Module = m_sc_control.Modules(ModuleID)
    
    If Module Is Nothing Then Exit Function
    
    ' get Script() value "Name"
    GetScriptName = GetScriptDictionary(Module)("Name")
    
    Exit Function

ERROR_HANDLER:

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in GetScriptName()."

    Err.Clear
    
    Resume Next
    
End Function

'Adds a Observer/Observie pair to the ScriptObservers collection.
'Observer\x0Observie
'Checks for duplicates, Ignoring cases of corse.
'Also prevents Observing yourself
Public Sub AddScriptObserver(ByVal ModuleName As String, ByVal sTargetScript As String)
On Error GoTo ERROR_HANDLER
    Dim i As Integer
    Dim observed As Collection
    
    If (StrComp(ModuleName, sTargetScript, vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    Set observed = GetScriptObservers(ModuleName)

    For i = 1 To observed.Count
        If (StrComp(observed.Item(i), sTargetScript, vbTextCompare) = 0) Then
            Set observed = Nothing
            Exit Sub
        End If
    Next i
    Set observed = Nothing
    
    m_ScriptObservers.Add ModuleName & vbNullChar & sTargetScript
    
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modScripting.AddScriptObserver()"
End Sub

'Returns a collection of scripts that the passed script is currently observing, or being observed by
'Needs a better name...
Public Function GetScriptObservers(sScriptName As String, Optional IsObserver As Boolean = True) As Collection
On Error GoTo ERROR_HANDLER

    Dim i As Integer
    Dim sObserver As String
    Dim sObservie As String

    Set GetScriptObservers = New Collection

    For i = 1 To m_ScriptObservers.Count
        If (InStr(m_ScriptObservers.Item(i), vbNullChar)) Then
            sObserver = Split(m_ScriptObservers.Item(i), Chr$(0))(0)
            sObservie = Split(m_ScriptObservers.Item(i), Chr$(0))(1)
            
            
            If (StrComp(sObserver, sScriptName, vbTextCompare) = 0 And IsObserver) Then
                GetScriptObservers.Add sObservie
            ElseIf (StrComp(sObservie, sScriptName, vbTextCompare) = 0 And IsObserver = False) Then
                GetScriptObservers.Add sObserver
            End If
        End If
    Next i

    Exit Function
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, "Error: #" & Err.Number & ": " & Err.Description & " in modScripting.GetScriptObservers()"
End Function

'Adds a Function/Observer pair to the Function Observer collection
'Observer\x00Function
Public Sub AddFunctionObserver(ByVal ModuleName As String, ByVal sTargetFunction As String)
On Error GoTo ERROR_HANDLER
    Dim i     As Integer
    Dim sItem As String
        
    sItem = StringFormat("{0}{1}{2}", ModuleName, Chr$(0), sTargetFunction)
        
    For i = 1 To m_FunctionObservers.Count
        If (StrComp(m_FunctionObservers.Item(i), sItem, vbTextCompare) = 0) Then
            Exit Sub
        End If
    Next i
    
    m_FunctionObservers.Add sItem
    
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in modScripting.AddFunctionObservers()", Err.Number, Err.Description)
End Sub

'Returns a collections of scripts who are observing this event in all scripts.
Public Function GetFunctionObservers(sFunctionName As String) As Collection
On Error GoTo ERROR_HANDLER:
    Dim i As Integer
    Dim sFunction As String
    Dim sObserver As String
    
    Set GetFunctionObservers = New Collection
    
    For i = 1 To m_FunctionObservers.Count
        If (InStr(m_FunctionObservers.Item(i), Chr$(0))) Then
            sObserver = Split(m_FunctionObservers.Item(i), Chr$(0))(0)
            sFunction = Split(m_FunctionObservers.Item(i), Chr$(0))(1)
            
            If (StrComp(sFunctionName, sFunction, vbTextCompare) = 0) Then
                GetFunctionObservers.Add sObserver
            End If
        End If
    Next i
    
    Exit Function
ERROR_HANDLER:
    frmChat.AddChat RTBColors.ErrorMessageText, StringFormat("Error #{0}: {1} in modScripting.GetFunctionObservers()", Err.Number, Err.Description)
End Function

' call this during config reload to enable/disable the system
' (all calls to the script system will exit if true!!)
' if this call is enabling the system, load scripts (init sc, init scripts, load scripts)
' if this call is disabling the system, clean up scripts (close)
Public Sub SetScriptSystemDisabled(ByVal SystemDisabled As Boolean)
    
    On Error GoTo ERROR_HANDLER
    
    If m_SystemDisabled = Not SystemDisabled Then
        If SystemDisabled Then
            ' system is being disabled, close (if open)
            If m_SCInitialized Then
                'RunInAll "Event_LoggedOff"
                RunInAll "Event_Close"
            End If
            ' hide scripting menu
            frmChat.mnuScripting.Visible = False
            ' store the state
            m_SystemDisabled = True
        Else
            ' store the state
            m_SystemDisabled = False
            ' show scripting menu
            frmChat.mnuScripting.Visible = True
            ' system is being enabled, open
            InitScriptControl frmChat.SControl
            LoadScripts
            InitScripts
        End If
    End If
    
    Exit Sub

ERROR_HANDLER:

    ' Cannot call this method while the script is executing
    If (Err.Number = -2147467259) Then
        frmChat.AddChat RTBColors.ErrorMessageText, "Error: Script is still executing."
        
        Exit Sub
    End If

    frmChat.AddChat RTBColors.ErrorMessageText, "Error (#" & Err.Number & "): " & Err.Description & _
        " in SetScriptSystemDisabled()."
End Sub

Public Function GetScriptSystemDisabled() As Boolean

    GetScriptSystemDisabled = m_SystemDisabled
    
End Function

' call this function to get a script module's CodeObject.Script dictionary.
' this will make sure that the CodeObject.Script is of type Dictionary and
' can be accessed as such
' if not, the CodeObject.Script object will be restored (but all data was lost)
' prevents RTE due to CodeObject.Script(key) accesses failing when Script not
' of type Dictionary ~Ribose
Public Function GetScriptDictionary(ByRef mdl As Module) As Dictionary
    If (StrComp(TypeName$(mdl.CodeObject.Script), "Dictionary") <> 0) Then
        SetScriptDictionary mdl
        frmChat.AddChat RTBColors.ErrorMessageText, "Scripting error: A Script object has been reset. " & _
                               "Script module ID: " & mdl.Name
    End If
    Set GetScriptDictionary = mdl.CodeObject.Script
End Function

' call this function to set or reset the contents of the CodeObject.Script
' dictionary for the specified module
Private Function SetScriptDictionary(ByRef mdl As Module)
    ' let's store a dictionary into the specified CodeObject.Script
    Dim Dict As Scripting.Dictionary
    Set Dict = New Scripting.Dictionary
    ' make Script object keys case-insensitive
    Dict.CompareMode = TextCompare
    ' store it
    Set mdl.CodeObject.Script = Dict
End Function

