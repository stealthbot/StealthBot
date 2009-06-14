Attribute VB_Name = "modScripting"
'/* Scripting.bas
' * ~~~~~~~~~~~~~~~~
' * StealthBot VBScript support module
' * ~~~~~~~~~~~~~~~~
' * Modified by Swent 11/06/2007
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
Private m_sc_control    As Object
Private m_is_reloading  As Boolean
Private VetoNextMessage As Boolean
Private VetoNextPacket  As Boolean

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

        ' ...
        For I = 1 To Paths.Count
            ' ...
            If (IsValidFileExtension(GetFileExtension(Paths(I)))) Then
                ' ...
                Set CurrentModule = _
                    m_sc_control.Modules.Add(m_sc_control.Modules.Count + 1)
            
                ' ...
                res = FileToModule(CurrentModule, strPath & Paths(I))
            
                ' ...
                If (IsScriptNameValid(CurrentModule) = False) Then
                    ' ...
                    CurrentModule.CodeObject.Script("Name") = CleanFileName(Paths(I))
                
                    ' ...
                    If (IsScriptNameValid(CurrentModule) = False) Then
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
                If (res = False) Then
                    CurrentModule.AddCode GetDefaultModuleProcs(CurrentModule.Name, _
                        Paths(I))
                End If
            End If
        Next I
    End If
    
    ' ...
    InitMenus

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
            
            ' ...
            If (Len(strLine) >= 1) Then
                If ((blnCheckOperands) And (Left$(strLine, 1) = "#")) Then
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
                                       tmp = App.Path & "\scripts\" & tmp
                                    End If

                                    ' ...
                                    includes.Add tmp
                                End If
                            End If
                        End If
                    End If
                Else
                    blnCheckOperands = False
                End If
            End If
            
            ' ...
            If (blnCheckOperands = False) Then
                strContent = _
                    strContent & strLine & vbCrLf
            End If
            
            ' ...
            strLine = vbNullString
        Loop
    Close #f
  
    ' ...
    If (defaults) Then
        ' ...
        ScriptModule.ExecuteStatement _
            "Set Script = CreateObject(""Scripting.Dictionary"")"
            
        ' ...
        ScriptModule.ExecuteStatement "Set DataBuffer = DataBufferEx()"
        
        ' ...
        ScriptModule.ExecuteStatement _
            "Script(""Path"") = " & Chr$(34) & filePath & Chr$(34)

        ' ...
        For I = 1 To includes.Count
            FileToModule ScriptModule, includes(I), False
        Next

        ' ...
        strContent = strContent & _
            GetDefaultModuleProcs(ScriptModule.Name, filePath)
    
        ' ...
        ScriptModule.AddCode strContent
        
        ' ...
        Set includes = Nothing
        
        ' ...
        strContent = ""
    End If
    
    FileToModule = True
    
    Exit Function
    
ERROR_HANDLER:

    ' ...
    strContent = ""

    FileToModule = False

End Function

Private Function GetDefaultModuleProcs(ByVal ScriptID As String, ByVal ScriptPath As String) As String

    Dim str As String ' storage buffer for module code

    ' GetModuleID() module-level function
    str = str & "Function GetModuleID()" & vbNewLine
    str = str & "   GetModuleID = " & Chr$(34) & ScriptID & Chr$(34) & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetScriptModule() module-level function
    str = str & "Function GetScriptModule()" & vbNewLine
    str = str & "   Set GetScriptModule = IGetScriptModule(GetModuleID)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetWorkingDirectory() module-level function
    str = str & "Function GetWorkingDirectory()" & vbNewLine
    str = str & "   GetWorkingDirectory = _" & vbNewLine
    str = str & "       BotPath() & ""Scripts\"" & " & "Script(""Name"") & " & """\""" & vbNewLine
    str = str & "   MkDirEx GetWorkingDirectory" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' CreateObj() module-level function
    str = str & "Function CreateObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Set CreateObj = _ " & vbNewLine
    str = str & "         ICreateObj(GetModuleID(), ObjType, ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' DestroyObj() module-level function
    str = str & "Sub DestroyObj(ObjName)" & vbNewLine
    str = str & "   IDestroyObj GetModuleID(), ObjName" & vbNewLine
    str = str & "End Sub" & vbNewLine
    
    ' GetObjByName() module-level function
    str = str & "Function GetObjByName(ObjName)" & vbNewLine
    str = str & "   Set GetObjByName = _ " & vbNewLine
    str = str & "         IGetObjByName(GetModuleID(), ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetSettingsEntry() module-level function
    str = str & "Function GetSettingsEntry(EntryName)" & vbNewLine
    str = str & "   GetSettingsEntry = GetSettingsEntryEx(Script(""Name""), EntryName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' WriteSettingsEntry() module-level function
    str = str & "Sub WriteSettingsEntry(EntryName, EntryValue)" & vbNewLine
    str = str & "   WriteSettingsEntryEx Script(""Name""), EntryName, EntryValue" & vbNewLine
    str = str & "End Sub" & vbNewLine

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
    str = _
        CurrentModule.CodeObject.Script("Name")

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
            tmp = _
                m_sc_control.Modules(j).CodeObject.Script("Name")
                
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
    Dim str As String  ' ...
    
    If (reloading = False) Then
        RunInAll "Event_FirstRun"
        
        reloading = True
    End If
    
    For I = 2 To m_sc_control.Modules.Count
        If (I > 1) Then
            str = _
                m_sc_control.Modules(I).CodeObject.GetSettingsEntry("Enabled")
        End If

        If (StrComp(str, "False", vbTextCompare) <> 0) Then
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

    Set SC = m_sc_control
    
    If (m_is_reloading) Then
        Exit Function
    End If
    
    oldVeto = GetVeto 'Keeps the old veto, for recursion, and Sets to false.
    veto = False      'Just to be sure

    arr() = Parameters()

    For I = 2 To SC.Modules.Count
        str = SC.Modules(I).CodeObject.GetSettingsEntry("Enabled")
        
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            CallByNameEx SC.Modules(I), "Run", VbMethod, arr()
            veto = veto Or GetVeto 'Did they veto it, or was it already vetod?
        End If
    Next
    
    SetVeto oldVeto 'Reset the old veto, this is for recursion
    RunInAll = veto 'Was this particular event vetoed?
End Function

Public Function RunInSingle(ByRef obj As Module, ParamArray Parameters() As Variant) As Boolean

    On Error Resume Next

    Dim I     As Integer
    Dim arr() As Variant
    Dim str   As String
    Dim oldVeto As Boolean
        
    If (m_is_reloading) Then
        Exit Function
    End If
    
    oldVeto = GetVeto  'Keeps the old veto, for recursion, and Sets to false.

    arr() = Parameters()

    str = obj.CodeObject.GetSettingsEntry("Enabled")
    
    If (StrComp(str, "False", vbTextCompare) <> 0) Then
        CallByNameEx obj, "Run", VbMethod, arr()
    End If
    
    RunInSingle = GetVeto 'Was this particular event vetoed?
    SetVeto oldVeto 'Reset the old veto, this is for recursion
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
            v(I) = vArgsArray(numArgs - I)
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
                    "mnu" & SCModule.CodeObject.Script("Name") & "Dash1"
                tmp.Parent = _
                    DynamicMenus("mnu" & SCModule.CodeObject.Script("Name"))
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
                DynamicMenus("mnu" & SCModule.CodeObject.Script("Name"))
                
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
        Name = _
            frmChat.SControl.Modules(I).CodeObject.Script("Name")
    
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & " ROOT"
        tmp.hWnd = GetSubMenu(GetMenu(frmChat.hWnd), 5)
        tmp.Caption = Name
            
        ' ...
        DynamicMenus.Add tmp, "mnu" & Name
            
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & " ENABLE|DISABLE"
        tmp.Parent = DynamicMenus("mnu" & Name)
        tmp.Caption = "Enabled"
        
        If (StrComp(frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Enabled"), _
                "False", vbTextCompare) <> 0) Then
                
            tmp.Checked = True
        End If
        
        ' ...
        DynamicMenus.Add tmp
        
        ' ...
        Set tmp = New clsMenuObj
    
        ' ...
        tmp.Name = Chr$(0) & Name & " VIEW_SCRIPT"
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

    Set Scripts = New Collection

    For I = 2 To frmChat.SControl.Modules.Count
        str = _
            frmChat.SControl.Modules(I).CodeObject.GetSettingsEntry("Public")
                
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            Scripts.Add frmChat.SControl.Modules(I).CodeObject, _
                frmChat.SControl.Modules(I).CodeObject.Script("Name")
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
