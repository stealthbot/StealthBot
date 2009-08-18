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

Private m_srootmenu     As New clsMenuObj
Private m_arrObjs()     As scObj
Private m_objCount      As Integer
Private m_sc_control    As ScriptControl
Private m_is_reloading  As Boolean
Private m_ExecutingMdl  As Module
Private m_TempMdlName   As String
Private m_IsEventError  As Boolean
Private VetoNextMessage As Boolean
Private VetoNextPacket  As Boolean
Private m_ScriptObservers  As Collection

Public Sub InitScriptControl(ByVal SC As ScriptControl)

    ' ...
    m_is_reloading = True

    ' ...
    SC.Reset

    ' ...
    frmChat.INet.Cancel
    frmChat.scTimer.Enabled = False
    
    ' ...
    DestroyObjs

    ' ...
    If (ReadINI("Other", "ScriptAllowUI", GetConfigFilePath()) <> "N") Then
        SC.AllowUI = True
    End If

    '// Create scripting objects
    SC.addObject "ssc", SharedScriptSupport, True
    SC.addObject "scTimer", frmChat.scTimer
    SC.addObject "scINet", frmChat.INet
    SC.addObject "BotVars", BotVars

    ' ...
    Set m_sc_control = SC
    
    Set m_ScriptObservers = New Collection
    
    ' ...
    m_is_reloading = False

End Sub

Public Sub LoadScripts()

    ' ...
    On Error GoTo ERROR_HANDLER

    Dim CurrentModule As Module
    Dim Paths         As New Collection
    Dim strPath       As String  ' ...
    Dim filename      As String  ' ...
    Dim fileExt       As String  ' ...
    Dim I             As Integer ' ...
    Dim j             As Integer ' ...
    Dim str           As String  ' ...
    Dim tmp           As String  ' ...
    Dim res           As Boolean ' ...

    ' ********************************
    '      LOAD SCRIPTS
    ' ********************************
    
    ' set script folder path
    strPath = ReadCfg("FilePaths", "Scripts")
    If (strPath = vbNullString) Then
        strPath = App.Path & "\scripts\"
    ElseIf (Not (Right$(strPath, 1) = "\")) Then
        strPath = strPath & "\"
    End If

    
    ' ensure scripts folder exists
    If (Dir(strPath) <> vbNullString) Then
        ' grab initial script file name
        filename = Dir(strPath)
        
        ' grab script files
        ' note: if we don't enumerate this list prior to script loading,
        ' scripting errors can kill further script loading.
        Do While (filename <> vbNullString)
            ' add script file to collection
            Paths.Add filename
        
            ' grab next script file name
            filename = Dir()
        Loop

        ' Cycle through each of the files.
        For I = 1 To Paths.Count
            ' Does the file have the extension for a script?
            If (IsValidFileExtension(GetFileExtension(Paths(I)))) Then
                ' Add a new module to the script control.
                Set CurrentModule = _
                    m_sc_control.Modules.Add(m_sc_control.Modules.Count + 1)
                
                ' store the temporary name of the module for parsing errors
                m_TempMdlName = Paths(I)
                
                ' set executing module reference for parsing errors and module-specific functions
                Set m_ExecutingMdl = CurrentModule
                
                ' Load the file into the module.
                res = FileToModule(CurrentModule, strPath & Paths(I))
                
                ' set executing module to nothing
                Set m_ExecutingMdl = Nothing
                
                ' Does the script have a valid name?
                If (IsScriptNameValid(CurrentModule) = False) Then
                    ' No. Try to fix it.
                    CurrentModule.CodeObject.Script("Name") = CleanFileName(m_TempMdlName)
                
                    ' Is it valid now?
                    If (IsScriptNameValid(CurrentModule) = False) Then
                        ' No, disable it.
                        
                        frmChat.AddChat vbRed, "Scripting error: " & Paths(I) & " has been " & _
                            "disabled due to a naming issue."
                            
                        str = strPath & "\disabled\"
                        
                        MkDir str
                            
                        Kill str & Paths(I)
    
                        Name strPath & Paths(I) As str & Paths(I)
    
                        InitScriptControl m_sc_control
                        
                        LoadScripts
                        
                        Exit Sub
                    End If
                End If
                
                ' ...
                'If (res = False) Then
                '    CurrentModule.AddCode GetDefaultModuleProcs(CurrentModule.Name, _
                '        Paths(I))
                'End If
            End If
        Next I
    End If
    
    ' ...
    InitMenus

    frmChat.AddChat vbGreen, "Scripts loaded."
        
    ' ...
    Exit Sub

' ...
ERROR_HANDLER:

    ' ...
    frmChat.AddChat vbRed, _
        "Error (" & Err.Number & "): " & Err.description & " in LoadScripts()."

End Sub

Private Function FileToModule(ByRef ScriptModule As Module, ByVal filePath As String, Optional ByVal defaults As Boolean = True) As Boolean

    On Error GoTo ERROR_HANDLER

    Static strContent    As String  ' ...

    Dim includes         As New Collection
    Dim strLine          As String  ' ...
    Dim f                As Integer ' ...
    Dim blnCheckOperands As Boolean
    Dim blnKeepLine      As Boolean
    Dim blnScriptData    As Boolean
    Dim I                As Integer

    ' ...
    blnCheckOperands = True
    
    ' ...
    f = FreeFile

    ' ...
    Open filePath For Input As #f
        ' ...
        Do While (EOF(f) = False)
            ' ...
            Line Input #f, strLine
            
            ' ...
            strLine = Trim(strLine)
            
            ' default to keep line
            blnKeepLine = True
            
            ' ...
            If (Len(strLine) >= 1) Then
                If ((blnCheckOperands) And (Left$(strLine, 1) = "#")) Then
                    ' this line is a directive, parse and don't keep in code
                    blnKeepLine = False
                    If (InStr(1, strLine, " ") <> 0) Then
                        If (Len(strLine) >= 2) Then
                            Dim strCommand As String ' ...
                        
                            strCommand = _
                                LCase$(Mid$(strLine, 2, InStr(1, strLine, " ") - 2))
    
                            If (strCommand = "include") Then
                                If (Len(strLine) >= 12) Then
                                    Dim tmp As String ' ...
                                    
                                    ' ...
                                    tmp = _
                                        LCase$(Mid$(strLine, 11, Len(strLine) - 11))
                                    
                                    ' ...
                                    If (Left$(tmp, 1) = "\") Then
                                       tmp = App.Path & "\scripts" & tmp
                                    End If

                                    ' ...
                                    includes.Add tmp
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
                    ' this line is not a comment or blank line, disable #include
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
            
            ' clean up
            strLine = vbNullString
        Loop
    Close #f
  
    ' if we are not loading an include, set it up
    If (defaults) Then
        ' store module-level functions
        ScriptModule.ExecuteStatement GetDefaultModuleProcs()
        
        ' create Script object
        ScriptModule.ExecuteStatement "Set Script = CreateObject(""Scripting.Dictionary"")"
        
        ' make Script object keys case-insensitive
        ScriptModule.CodeObject.Script.CompareMode = Scripting.CompareMethod.TextCompare
        
        ' create default DataBuffer object
        ScriptModule.ExecuteStatement "Set DataBuffer = DataBufferEx()"
        
        ' set the path Script() value
        ScriptModule.CodeObject.Script("Path") = filePath
        
        ' add includes to end of code, process the same for includes in includes
        For I = 1 To includes.Count
            FileToModule ScriptModule, includes(I), False
        Next I
        
        ' add content
        ScriptModule.AddCode strContent
        
        ' clean up object
        Set includes = Nothing
        
        ' clean up content for next call to this function
        strContent = vbNullString
    End If
    
    FileToModule = True
    
    Exit Function
    
ERROR_HANDLER:

    ' ...
    strContent = vbNullString

    FileToModule = False

End Function


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

    ' WriteSettingsEntry() module-level function
    str = str & "Sub WriteSettingsEntry(EntryName, EntryValue)" & vbNewLine
    str = str & "   SSC.WriteSettingsEntry EntryName, EntryValue, Script(""Name"")" & vbNewLine
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

    ' IsCommand() module-level function
    str = str & "Function IsCommand(commandText, username)" & vbNewLine
    str = str & "   Set IsCommand = _ " & vbNewLine
    str = str & "         SSC.IsCommand(commandText, username, Script(""Name""))" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' store module-level coding
    GetDefaultModuleProcs = str

End Function

Private Function IsScriptNameValid(ByRef CurrentModule As Module) As Boolean

    On Error Resume Next

    Dim j            As Integer ' ...
    Dim str          As String  ' ...
    Dim tmp          As String  ' ...
    Dim nameDisallow As String
    Dim I            As Integer
    
    ' ...
    str = GetScriptName(CurrentModule.Name)

    ' ...
    If (str = vbNullString) Then
        IsScriptNameValid = False
                
        Exit Function
    End If
    
    ' ...
    nameDisallow = "\/:*?<>|"""

    ' ...
    For j = 1 To Len(str)
        If (InStr(1, nameDisallow, Mid$(str, j, 1), vbTextCompare) <> 0) Then
            IsScriptNameValid = False
    
            Exit Function
        End If
    Next j

    ' ...
    For j = 2 To m_sc_control.Modules.Count
        ' ...
        If (m_sc_control.Modules(j).Name <> CurrentModule.Name) Then
            tmp = GetScriptName(CStr(j))
                
            If (StrComp(str, tmp, vbTextCompare) = 0) Then
                IsScriptNameValid = False

                Exit Function
            End If
        End If
    Next j
    
    ' ...
    IsScriptNameValid = True

End Function

Public Sub InitScripts()

    On Error Resume Next
    
    Static reloading As Boolean ' ...
    
    Dim I   As Integer ' ...
    Dim tmp As String  ' ...
    
    If (reloading = False) Then
        RunInAll "Event_FirstRun"
        
        reloading = True
    End If
    
    For I = 2 To m_sc_control.Modules.Count
        If (I > 1) Then
            tmp = m_sc_control.Modules(I).CodeObject.GetSettingsEntry("Enabled")
        End If

        If (StrComp(tmp, "False", vbTextCompare) <> 0) Then
            InitScript m_sc_control.Modules(I)
        End If
    Next I

End Sub

Public Sub InitScript(ByVal SCModule As Module)

    On Error Resume Next

    Dim I          As Integer ' ...
    Dim startTime  As Long
    Dim finishTime As Long

    ' ...
    startTime = GetTickCount()
    
    ' ...
    RunInSingle SCModule, "Event_Load"
    
    ' ...
    finishTime = GetTickCount()
 
    '// 03/27/2009 52 - added default Script property for the load time
    SCModule.CodeObject.Script("InitPerf") = (finishTime - startTime)

    If (g_Online) Then
        RunInSingle SCModule, "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        RunInSingle SCModule, "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags
    
        If (g_Channel.Users.Count > 0) Then
            For I = 1 To g_Channel.Users.Count
                With g_Channel.Users(I)
                     RunInSingle SCModule, "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
                        .game, False
                End With
             Next I
         End If
    End If
End Sub

Public Function RunInAll(ParamArray Parameters() As Variant) As Boolean

    On Error Resume Next

    Dim SC      As ScriptControl
    Dim I       As Integer
    Dim arr()   As Variant
    Dim str     As String
    Dim oldVeto As Boolean
    Dim veto    As Boolean
    Dim oldEM   As Module
    Dim obj     As Module
    Dim Proc    As Procedure

    Set SC = m_sc_control
    
    If (m_is_reloading) Then
        Exit Function
    End If
    
    oldVeto = GetVeto 'Keeps the old veto, for recursion, and Sets to false.
    veto = False      'Just to be sure
    
    Set oldEM = m_ExecutingMdl ' keep last reference, for recursion

    arr() = Parameters()

    For I = 2 To SC.Modules.Count
        Set obj = SC.Modules(I)
        Set m_ExecutingMdl = obj
        
        str = obj.CodeObject.GetSettingsEntry("Enabled")
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            
            '' check if module has the procedure before calling it!
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

    Dim I       As Integer
    Dim arr()   As Variant
    Dim str     As String
    Dim oldVeto As Boolean
    Dim oldEM   As Module
    Dim Proc    As Procedure
    Dim obsers  As Collection
    Dim obser   As Module
        
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
    Else
        str = obj.CodeObject.GetSettingsEntry("Enabled")
    End If
    
    If (Not StrComp(str, "False", vbTextCompare) = 0) Then
        If (Not obj Is Nothing) Then CallByNameEx obj, "Run", VbMethod, arr()
        RunInSingle = GetVeto 'Was this particular event vetoed?
        
        'Call any scripts that are observing this one
        If (Not obj Is Nothing) Then
            Set obsers = GetScriptObservers(GetScriptName(obj.Name), False)
        Else
            Set obsers = GetScriptObservers(vbNullString, False)
        End If
        For I = 1 To obsers.Count
            Set obser = GetModuleByName(obsers.Item(I))
            If (Not obser Is Nothing) Then 'Is the script real/loaded?
                str = obser.CodeObject.GetSettingsEntry("Enabled")
                If (StrComp(str, "False", vbTextCompare) <> 0) Then 'Is it off?
                    Set m_ExecutingMdl = obser
                    CallByNameEx obser, "Run", VbMethod, arr
                End If
            End If
            Set obser = Nothing
        Next I
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
    Dim I       As Long
    Dim v()     As Variant

    Set oTLI = New TLIApplication

    ProcID = oTLI.InvokeID(obj, ProcName)

    If (IsMissing(vArgsArray)) Then
        Call oTLI.InvokeHook(obj, ProcID, CallType)
    End If
    
    If (IsArray(vArgsArray)) Then
        numArgs = UBound(vArgsArray)
        
        ReDim v(numArgs)
        
        For I = 0 To numArgs
            ' corrected object passing -Ribose/2009-08-10
            If IsObject(vArgsArray(numArgs - I)) Then
                Set v(I) = vArgsArray(numArgs - I)
            Else
                Let v(I) = vArgsArray(numArgs - I)
            End If
        Next I
        
        Call oTLI.InvokeHookArray(obj, ProcID, CallType, v)
    End If
    
    Set oTLI = Nothing

    Exit Sub

ERROR_HANDLER:

    ' ...
    If (frmChat.SControl.Error) Then
        Exit Sub
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in CallByNameEx()."
        
    Set oTLI = Nothing
    
End Sub

Public Function Objects(objIndex As Integer) As scObj

    Objects = m_arrObjs(objIndex)

End Function

Private Function ObjCount(Optional ObjType As String, Optional ByVal SCModule As Module = Nothing) As Integer
    
    Dim I As Integer ' ...

    If (ObjType <> vbNullString) Then
        For I = 0 To m_objCount - 1
            If (SCModule Is Nothing) Then
                If (StrComp(ObjType, m_arrObjs(I).ObjType, vbTextCompare) = 0) Then
                    ObjCount = (ObjCount + 1)
                End If
            Else
                If (StrComp(SCModule.Name, m_arrObjs(I).SCModule.Name) = 0) Then
                    If (StrComp(ObjType, m_arrObjs(I).ObjType, vbTextCompare) = 0) Then
                        ObjCount = (ObjCount + 1)
                    End If
                End If
            End If
        Next I
    Else
        ObjCount = m_objCount
    End If

End Function

Public Function CreateObj(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    On Error Resume Next

    Dim obj As scObj ' ...
    Dim scriptName As String
    
    If SCModule Is Nothing Then Exit Function
    
    Set CreateObj = Nothing
    If (Not ValidObjectName(ObjName)) Then Exit Function
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim I As Integer ' loop counter variable

        For I = 0 To m_objCount - 1
            If (m_arrObjs(I).SCModule.Name = SCModule.Name) Then
                If (StrComp(m_arrObjs(I).ObjType, ObjType, vbTextCompare) = 0) Then
                    If (StrComp(m_arrObjs(I).ObjName, ObjName, vbTextCompare) = 0) Then
                        Set CreateObj = m_arrObjs(I).obj
                    
                        Exit Function
                    End If
                End If
            End If
        Next I
        
        ReDim Preserve m_arrObjs(0 To m_objCount)
    Else
        ReDim m_arrObjs(0)
    End If

    ' store our module name & type
    obj.ObjName = ObjName
    obj.ObjType = ObjType

    ' store our module handle
    Set obj.SCModule = SCModule
    
    scriptName = GetScriptName(SCModule.Name)
    
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
            
            ' ...
            obj.obj.setName ObjName
            obj.obj.setSCModule SCModule
            
            ' ...
            HookWindowProc obj.obj.hWnd
            
        Case "MENU"
            If (ObjCount("Menu", SCModule) = 0) Then
                Dim tmp As New clsMenuObj ' ...
                
                tmp.Name = _
                    "mnu" & scriptName & "Dash1"
                tmp.Parent = _
                    DynamicMenus("mnu" & scriptName)
                tmp.Caption = "-"
                
                ' ...
                m_arrObjs(m_objCount).ObjName = tmp.Name
                m_arrObjs(m_objCount).ObjType = "Menu"
                
                ' ...
                Set m_arrObjs(m_objCount).SCModule = SCModule
                
                ' ...
                Set m_arrObjs(m_objCount).obj = tmp
                
                ' ...
                m_objCount = (m_objCount + 1)
                
                ' ...
                ReDim Preserve m_arrObjs(0 To m_objCount)
                
                ' ...
                DynamicMenus.Add tmp
            End If
        
            ' ...
            Set obj.obj = New clsMenuObj
            
            ' ...
            obj.obj.Name = ObjName
            
            ' ...
            obj.obj.Parent = _
                DynamicMenus("mnu" & scriptName)
                
            ' ...
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

    Dim I As Integer ' ...
    
    ' ...
    For I = m_objCount - 1 To 0 Step -1
        If (SCModule Is Nothing) Then
            DestroyObj m_arrObjs(I).SCModule, m_arrObjs(I).ObjName
        Else
            If (SCModule.Name = m_arrObjs(I).SCModule.Name) Then
                DestroyObj m_arrObjs(I).SCModule, m_arrObjs(I).ObjName
            End If
        End If
    Next I
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in DestroyObjs()."
        
    Resume Next
    
End Sub

Public Sub DestroyObj(ByVal SCModule As Module, ByVal ObjName As String)

    On Error GoTo ERROR_HANDLER

    Dim I     As Integer ' ...
    Dim Index As Integer ' ...
    
    If SCModule Is Nothing Then Exit Sub
    
    ' ...
    If (m_objCount = 0) Then
        Exit Sub
    End If
    
    ' ...
    Index = m_objCount
    
    ' ...
    For I = 0 To m_objCount - 1
        If (m_arrObjs(I).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(I).ObjName, ObjName, vbTextCompare) = 0) Then
                Index = I
            
                Exit For
            End If
        End If
    Next I
    
    ' ...
    If (Index >= m_objCount) Then
        Exit Sub
    End If
    
    ' ...
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
            m_arrObjs(Index).obj.Class_Terminate
    End Select

    ' ...
    Set m_arrObjs(Index).obj = Nothing
    
    ' ...
    If (Index < m_objCount - 1) Then
        For I = Index To ((m_objCount - 1) - 1)
            m_arrObjs(I) = m_arrObjs(I + 1)
        Next I
    End If
    
    ' ...
    If (m_objCount > 1) Then
        ReDim Preserve m_arrObjs(0 To m_objCount - 1)
    Else
        ReDim m_arrObjs(0)
    End If
    
    ' ...
    SCModule.ExecuteStatement "Set " & ObjName & " = Nothing"
    
    ' ...
    m_objCount = (m_objCount - 1)
    
    ' ...
    Exit Sub
    
ERROR_HANDLER:

    ' scripting engine has been reset - likely due to a reload
    If (Err.Number = -2147467259) Then
        Resume Next
    End If

    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in DestroyObj()."
        
    Resume Next
    
End Sub

Public Function GetObjByName(ByRef SCModule As Module, ByVal ObjName As String) As Object

    Dim I As Integer ' ...
    
    If SCModule Is Nothing Then Exit Function
    
    ' ...
    For I = 0 To m_objCount - 1
        If (m_arrObjs(I).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(I).ObjName, ObjName, vbTextCompare) = 0) Then
                Set GetObjByName = m_arrObjs(I).obj

                Exit Function
            End If
        End If
    Next I

End Function

Public Function GetScriptObjByMenuID(ByVal MenuID As Long) As scObj

    Dim I As Integer ' ...
    Dim j As Integer ' ...

    For I = 0 To ObjCount() - 1
        If (StrComp("Menu", Objects(I).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(I).obj.ID = MenuID) Then
                GetScriptObjByMenuID = m_arrObjs(I)
                
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
    Next I

End Function

Public Function GetScriptObjByIndex(ByVal ObjType As String, ByVal Index As Integer) As scObj

    Dim I As Integer ' ...

    For I = 0 To ObjCount() - 1
        If (StrComp(ObjType, Objects(I).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(I).obj.Index = Index) Then
                GetScriptObjByIndex = m_arrObjs(I)
                
                Exit For
            End If
        End If
    Next I

End Function

Public Function InitMenus()

    On Error GoTo ERROR_HANDLER

    Dim tmp  As clsMenuObj ' ...
    Dim Name As String     ' ...
    Dim I    As Integer    ' ...

    ' ...
    DestroyMenus
    
    ' ...
    For I = 2 To frmChat.SControl.Modules.Count
        If (I = 2) Then
            frmChat.mnuScriptingDash(0).Visible = True
        End If
    
        ' ...
        Name = GetScriptName(CStr(I))
    
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & Chr$(0) & "ROOT"
        tmp.hWnd = GetSubMenu(GetMenu(frmChat.hWnd), 5)
        tmp.Caption = Name
            
        ' ...
        DynamicMenus.Add tmp, "mnu" & Name
            
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & Chr$(0) & "ENABLE|DISABLE"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "Enabled"
        
        If (StrComp(GetModuleByName(Name).CodeObject.GetSettingsEntry("Enabled"), _
                "False", vbTextCompare) <> 0) Then
                
            tmp.Checked = True
        End If
        
        ' ...
        DynamicMenus.Add tmp
        
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & Chr$(0) & "VIEW_SCRIPT"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "View Script"
        
        ' ...
        DynamicMenus.Add tmp
    Next I
    
    Exit Function
        
ERROR_HANDLER:

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in InitMenus()."

    Err.Clear
    
    Resume Next

End Function

Public Function DestroyMenus()

    On Error GoTo ERROR_HANDLER

    Dim I As Integer ' ...
    
    frmChat.mnuScriptingDash(0).Visible = False
    
    For I = DynamicMenus.Count To 1 Step -1
        
        If (Len(DynamicMenus(I).Name) > 0) Then
            If (Left$(DynamicMenus(I).Name, 1) = Chr$(0)) Then
                DynamicMenus(I).Class_Terminate

                DynamicMenus.Remove I
            End If
        End If
        
    Next I
    
    Exit Function
    
ERROR_HANDLER:

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in DestroyMenus()."

    Err.Clear
    
    Resume Next
    
End Function

Public Function Scripts() As Object

    On Error Resume Next

    Dim I   As Integer ' ...
    Dim str As String  ' ...
    Dim SCModule As Module
    Dim scriptName As String

    Set Scripts = New Collection

    For I = 2 To frmChat.SControl.Modules.Count
        scriptName = GetScriptName(CStr(I))
        
        str = GetScriptModule.CodeObject.GetSettingsEntry("Public")
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            Scripts.Add frmChat.SControl.Modules(I).CodeObject, scriptName
        End If
    Next I

End Function

Public Function GetModuleByName(ByVal scriptName As String) As Module
    Dim I As Integer
    
    For I = 2 To frmChat.SControl.Modules.Count
        If (StrComp(GetScriptName(CStr(I)), scriptName, vbTextCompare) = 0) Then
            Set GetModuleByName = frmChat.SControl.Modules(I)
            Exit Function
        End If
    Next I
End Function

Public Sub SetVeto(ByVal b As Boolean)

    VetoNextMessage = b
    
End Sub

Public Function GetVeto() As Boolean

    GetVeto = VetoNextMessage
    
    VetoNextMessage = False
    
End Function

Private Function GetFileExtension(ByVal filename As String)
        
    Dim arr() As String

    arr = Split(filename, ".")
    
    If UBound(arr) = 0 Then
        GetFileExtension = ""
    Else
        GetFileExtension = arr(UBound(arr))
    End If

End Function

Private Function IsValidFileExtension(ByVal ext As String) As Boolean

    Dim exts() As String  ' ...
    Dim I      As Integer ' ...

    ' ...
    ReDim exts(0 To 1)
    
    ' ...
    'exts(0) = "dat"
    exts(0) = "txt"
    exts(1) = "vbs"
    
    ' ...
    For I = LBound(exts) To UBound(exts)
        If (StrComp(ext, exts(I), vbTextCompare) = 0) Then
            IsValidFileExtension = True
            
            Exit Function
        End If
    Next I
    
    IsValidFileExtension = False

End Function

Private Function CleanFileName(ByVal filename As String) As String
    
    On Error Resume Next
    
    ' ...
    If (InStr(1, filename, ".") > 1) Then
        CleanFileName = _
            Left$(filename, InStr(1, filename, ".") - 1)
    End If

End Function

'06/26/09 - Hdx Vary crappy function to check if Object names are valid, a-z0-9_ and 1st chr a-z (eventually should be a regexp)
Public Function ValidObjectName(sName As String) As Boolean
  Dim x As Integer
  Dim sValid As String
  
  sValid = "abcdefghijklmnopqrstuvwxyz0123456789_"
  ValidObjectName = False
  
  For x = 1 To Len(sName)
    If (InStr(1, Left(sValid, IIf(x = 1, 26, 37)), Mid$(sName, x, 1), vbTextCompare) = 0) Then Exit Function
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
    
    Dim Name    As String
    Dim ErrType As String
    
    With m_sc_control
        ErrType = "runtime"
        If m_ExecutingMdl Is Nothing Then
            Exit Sub ' exec error handler will handle this
        Else
            Name = GetScriptName(m_ExecutingMdl.Name)
            If Name = m_ExecutingMdl.Name Then Name = m_TempMdlName
            Name = Name & " script"
        End If
        
        ' check if its a parsing error
        If InStr(1, .Error.source, "compilation", vbBinaryCompare) > 0 Then
            ErrType = "parsing"
        Else
            ' check if the script is planning to handle errors itself, if Script("HandleErrors") = True, then call event_error
            If ((StrComp(m_ExecutingMdl.CodeObject.Script("HandleErrors"), _
                         "True", vbTextCompare) = 0) And _
                         (m_IsEventError = False)) Then
                ' call Event_Error(ErrObj)
                RunInSingle m_ExecutingMdl, "Event_Error", .Error
                
                ' if cleared, exit
                If .Error.Number = 0 Then Exit Sub
            End If
        End If
        
        ' display error
        frmChat.AddChat RTBColors.ErrorMessageText, _
            "Scripting " & ErrType & " error " & Chr(39) & .Error.Number & Chr(39) & _
            " in " & Name & ": (line " & .Error.line & "; column " & .Error.Column & ")"
        frmChat.AddChat RTBColors.ErrorMessageText, .Error.description
        frmChat.AddChat RTBColors.ErrorMessageText, "Offending line: >> " & .Error.Text
        .Error.Clear
    End With
    
End Sub

' get currently executing module (can be nothing!)
Public Function GetScriptModule(Optional ByVal scriptName As String = vbNullString) As Module

    On Error GoTo ERROR_HANDLER
    
    Dim I As Integer
    
    ' only loop if the scriptname was provided
    If LenB(scriptName) > 0 Then
        ' loop through modules
        For I = 2 To m_sc_control.Modules.Count
            ' if Script("Name") = ScriptName then
            If StrComp(GetScriptName(CStr(I)), scriptName, vbTextCompare) = 0 Then
                ' return this module
                Set GetScriptModule = m_sc_control.Modules(I)
                
                Exit Function
            End If
        Next I
    Else
        ' return currently executing
        Set GetScriptModule = m_ExecutingMdl
    End If
    
    Exit Function

ERROR_HANDLER:

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in GetScriptModule()."

    Err.Clear
    
    Resume Next

End Function

' get the module.Name for the provided script name
' returns the currently executing one if none provided
' if the currently executing one is nothing (/exec), "exec" is returned
Public Function GetModuleID(Optional ByVal scriptName As String = vbNullString) As String

    On Error GoTo ERROR_HANDLER
    
    Dim I As Integer
    
    ' only loop if the scriptname was provided
    If LenB(scriptName) > 0 Then
        ' loop through modules
        For I = 2 To m_sc_control.Modules.Count
            ' if Script("Name") = ScriptName then
            If StrComp(GetScriptName(CStr(I)), scriptName, vbTextCompare) = 0 Then
                ' return this ID
                GetModuleID = m_sc_control.Modules(I).Name
                Exit Function
            End If
        Next I
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

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
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
    GetScriptName = Module.CodeObject.Script("Name")
    
    Exit Function

ERROR_HANDLER:

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
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
    Dim I As Integer
    Dim observed As Collection
    
    If (StrComp(ModuleName, sTargetScript, vbTextCompare) = 0) Then
        Exit Sub
    End If
    
    Set observed = GetScriptObservers(ModuleName)

    For I = 1 To observed.Count
        If (StrComp(observed.Item(I), sTargetScript, vbTextCompare) = 0) Then
            Set observed = Nothing
            Exit Sub
        End If
    Next I
    Set observed = Nothing
    
    m_ScriptObservers.Add ModuleName & Chr$(0) & sTargetScript
    
    Exit Sub
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modScripting.AddScriptObserver()"
End Sub

'Returns a collection of scripts that the passed script is currently observing, or being observed by
'Needs a better name...
Public Function GetScriptObservers(sScriptName As String, Optional IsObserver As Boolean = True) As Collection
On Error GoTo ERROR_HANDLER

    Dim I As Integer
    Dim sObserver As String
    Dim sObservie As String

    Set GetScriptObservers = New Collection

    For I = 1 To m_ScriptObservers.Count
        If (InStr(m_ScriptObservers.Item(I), Chr$(0))) Then
            sObserver = Split(m_ScriptObservers.Item(I), Chr$(0))(0)
            sObservie = Split(m_ScriptObservers.Item(I), Chr$(0))(1)
            
            
            If (StrComp(sObserver, sScriptName, vbTextCompare) = 0 And IsObserver) Then
                GetScriptObservers.Add sObservie
            ElseIf (StrComp(sObservie, sScriptName, vbTextCompare) = 0 And IsObserver = False) Then
                GetScriptObservers.Add sObserver
            End If
        End If
    Next I

    Exit Function
ERROR_HANDLER:
    frmChat.AddChat vbRed, "Error: #" & Err.Number & ": " & Err.description & " in modScripting.GetScriptObservers()"
End Function
