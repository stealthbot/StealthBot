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

Public VetoNextMessage   As Boolean
Public boolOverride      As Boolean

Private m_arrObjs()      As scObj
Private m_objCount       As Integer

Public Sub InitScriptControl(ByRef SC As ScriptControl)

    ' ...
    DestroyObjs
    SC.Reset
    
    ' ...
    If (ReadINI("Other", "ScriptAllowUI", GetConfigFilePath()) <> "N") Then
        SC.AllowUI = True
    End If

    '// Create scripting objects
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars

End Sub

Public Sub LoadScripts(ByRef SC As ScriptControl)

    ' ...
    On Error GoTo ERROR_HANDLER

    Dim CurrentModule As Module
    Dim wrkScripts    As New Collection
    
    Dim strPath  As String  ' ...
    Dim filename As String  ' ...
    Dim fileExt  As String  ' ...
    Dim i        As Integer ' ...
    Dim j        As Integer ' ...
    Dim str      As String  ' ...
    Dim tmp      As String  ' ...
    
    ' ********************************
    '      LOAD REGULAR SCRIPTS
    ' ********************************
    
    ' ...
    strPath = App.Path & "\scripts\"
    
    ' ...
    If (Dir(strPath) <> vbNullString) Then
        ' ...
        filename = Dir(strPath)
        
        ' ...
        Do While (filename <> vbNullString)
            ' ...
            wrkScripts.Add filename
            
            ' ...
            filename = Dir()
        Loop
        
        For i = 1 To wrkScripts.Count
            ' ...
            If (IsValidFileExtension(GetFileExtension(wrkScripts(i)))) Then
                ' ...
                Set CurrentModule = SC.Modules.Add(SC.Modules.Count + 1)
                
                ' ...
                FileToModule CurrentModule, strPath & wrkScripts(i)

                ' ...
                If (CurrentModule.CodeObject.Script("Name") = vbNullString) Then
                    CurrentModule.CodeObject.Script("Name") = CleanFileName(wrkScripts(i))
                End If

                ' ...
                If (IsScriptUnique(SC, CurrentModule) = False) Then
                    frmChat.AddChat vbRed, "Scripting error: " & wrkScripts(i) & " has been " & _
                        "disabled due to a naming conflict."
                        
                    str = strPath & "\disabled\"
                    
                    MkDir str
                        
                    Kill str & wrkScripts(i)

                    Name strPath & wrkScripts(i) As str & wrkScripts(i)

                    InitScriptControl SC
                    LoadScripts SC
                    
                    Exit Sub
                End If
            End If
        Next i
    End If
    
    ' ...
    Set wrkScripts = New Collection
    
    ' ********************************
    '      LOAD PLUGIN SYSTEM
    ' ********************************

    ' ...
    If (ReadINI("Override", "DisablePS", GetConfigFilePath()) <> "Y") Then
        ' ...
        strPath = GetFilePath("PluginSystem.dat")
        
        ' ...
        If (LenB(Dir$(strPath)) = 0) Then
            Call frmChat.AddChat(vbRed, "Cannot find PluginSystem.dat. It must exist in order to load plugins!")
            Call frmChat.AddChat(vbYellow, "You may download PluginSystem.dat to your StealthBot folder using the link below.")
            Call frmChat.AddChat(vbWhite, "http://www.stealthbot.net/p/Users/Swent/index.php?file=PluginSystem.dat")
        Else
            FileToModule SC.Modules(1), strPath
        End If
        
        If (SC.Modules(1).CodeObject.Script("Name") = vbNullString) Then
            SC.Modules(1).CodeObject.Script("Name") = "PluginSystem"
        End If
    End If
    
    ' ********************************
    '     SET GLOBAL SCRIPT NAMES
    ' ********************************

    For i = 1 To SC.Modules.Count
        ' ...
        str = _
            SC.Modules(i).CodeObject.Script("Name")
    
        ' ...
        tmp = _
            SC.Modules(i).CodeObject.GetSettingsEntry("Public")
    
        ' ...
        If (StrComp(tmp, "False", vbTextCompare) <> 0) Then
            SC.Modules(1).ExecuteStatement "Set " & str & " = Scripts(" & Chr$(34) & _
                str & Chr$(34) & ")"
        End If
    Next i
    
    UpdateScripts SC
    
    ' ...
    Exit Sub

' ...
ERROR_HANDLER:

    ' object does not support property or method - function missing
    If (Err.Number = 438) Then
        Err.Clear
    
        Resume Next
    End If
    
    ' path/file errors - likely due to duplicate script handling
    If ((Err.Number = 75) Or (Err.Number = 53)) Then
        Err.Clear
    
        Resume Next
    End If

    ' ...
    frmChat.AddChat vbRed, _
        "Error (" & Err.Number & "): " & Err.description & " in LoadScripts()."

    Exit Sub

End Sub

Private Function IsScriptUnique(ByRef SC As ScriptControl, ByRef CurrentModule As Module) As Boolean

    On Error Resume Next

    Dim j   As Integer ' ...
    Dim str As String  ' ...
    Dim tmp As String  ' ...
    
    ' ...
    str = _
        CurrentModule.CodeObject.Script("Name")

    ' ...
    If (str <> vbNullString) Then
        For j = 1 To SC.Modules.Count
            ' ...
            If (SC.Modules(j).Name <> CurrentModule.Name) Then
                tmp = _
                    SC.Modules(j).CodeObject.Script("Name")
                    
                If (StrComp(str, tmp, vbTextCompare) = 0) Then
                    IsScriptUnique = False

                    Exit Function
                End If
            End If
        Next j
    End If
    
    IsScriptUnique = True

End Function

Public Function UpdateScripts(ByRef SC As ScriptControl)

    On Error Resume Next
    
    Dim wrkScripts As New Collection ' ...
    Dim i          As Integer ' ...
    Dim str        As String  ' ...
    Dim tmp        As String  ' ...
    
    ' ...
    If (SC.Modules.Count) Then
        Dim CRC32    As New clsCRC32
        Dim filePath As String

        ' ...
        frmChat.AddChat RTBColors.InformationText, "Checking for script updates..."
        
        ' ...
        For i = 1 To SC.Modules.Count
            str = _
                SC.Modules(i).CodeObject.Script("UpdateLocation")
                
            If (str <> vbNullString) Then
                ' ...
                filePath = App.Path & "\scripts\" & SC.Modules(i).Name
            
                ' ...
                URLDownloadToFile 0, str, filePath & ".tmp", 0, 0
                
                ' ...
                If (CRC32.GetFileCRC32(filePath) <> CRC32.GetFileCRC32(filePath & ".tmp")) Then
                    ' ...
                    wrkScripts.Add SC.Modules(i)
                
                    ' ...
                    Kill filePath
                    
                    ' ...
                    Name filePath & ".tmp" As filePath
                    
                    ' ...
                    FileToModule SC.Modules(i), filePath
                End If
                
                ' ...
                Kill filePath & ".tmp"
            End If
        Next i

        If (wrkScripts.Count) Then
            str = "Successfully updated the following scripts: "
            
            For i = 1 To wrkScripts.Count
                str = str & _
                    wrkScripts(i).CodeObject.GetScriptName & ", "
            Next i
            
            frmChat.AddChat vbGreen, Left$(str, Len(str) - 2)
        Else
            frmChat.AddChat vbGreen, "Scripts are up to date."
        End If
        
        Set CRC32 = Nothing
    End If

End Function

Private Function FileToModule(ByRef ScriptModule As Module, ByVal filePath As String, Optional ByVal defaults As Boolean = True)

    On Error GoTo ERROR_HANDLER

    Dim strLine          As String  ' ...
    Dim strContent       As String  ' ...
    Dim f                As Integer ' ...
    Dim blnCheckOperands As Boolean
    
    ' ...
    f = FreeFile
    
    ' ...
    blnCheckOperands = True

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
                                        filePath = App.Path & "\scripts\" & tmp
                                    Else
                                        filePath = tmp
                                    End If
            
                                    ' ...
                                    FileToModule ScriptModule, filePath, False
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
    ScriptModule.AddCode strContent
    
    ' ...
    If (defaults) Then
        CreateDefautModuleProcs ScriptModule
    End If
    
    ' ...
    Exit Function
    
ERROR_HANDLER:

    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in FileToModule()."
        
    Exit Function

End Function

Private Sub CreateDefautModuleProcs(ByRef ScriptModule As Module)

    On Error GoTo ERROR_HANDLER

    Dim str As String  ' storage buffer for module code

    ' ...
    ScriptModule.ExecuteStatement _
        "Set Script = CreateObject(" & Chr$(34) & "Scripting.Dictionary" & Chr$(34) & ")"

    ' ...
    ScriptModule.Run "Data"
    
    ' ...
    ScriptModule.ExecuteStatement "Set DataBuffer = DataBufferEx()"

    ' GetModuleName() module-level function
    str = str & "Function GetModuleName()" & vbNewLine
    str = str & "   GetModuleName = " & Chr$(34) & ScriptModule.Name & Chr$(34) & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetScriptName() module-level function
    str = str & "Function GetScriptName()" & vbNewLine
    str = str & "   GetScriptName = Script(""Name"")" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' CreateObj() module-level function
    str = str & "Function CreateObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Set CreateObj = _ " & vbNewLine
    str = str & "         CreateObjEx(GetModuleName(), ObjType, ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' DestroyObj() module-level function
    str = str & "Sub DestroyObj(ObjName)" & vbNewLine
    str = str & "   DestroyObjEx GetModuleName(), ObjName" & vbNewLine
    str = str & "End Sub" & vbNewLine
    
    ' GetObjByName() module-level function
    str = str & "Function GetObjByName(ObjName)" & vbNewLine
    str = str & "   Set GetObjByName = _ " & vbNewLine
    str = str & "         GetObjByNameEx(GetModuleName(), ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetSettingsEntry() module-level function
    str = str & "Function GetSettingsEntry(EntryName)" & vbNewLine
    str = str & "   GetSettingsEntry = GetSettingsEntryEx(GetScriptName(), EntryName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' WriteSettingsEntry() module-level function
    str = str & "Sub WriteSettingsEntry(EntryName, EntryValue)" & vbNewLine
    str = str & "   WriteSettingsEntryEx GetScriptName(), EntryName, EntryValue" & vbNewLine
    str = str & "End Sub" & vbNewLine

    ' store module-level coding
    ScriptModule.AddCode str
    
    Exit Sub
    
ERROR_HANDLER:

    ' object does not support property or method - function missing
    If (Err.Number = 438) Then
        Err.Clear
        
        Resume Next
    End If

    ' ...
    frmChat.AddChat vbRed, _
        "Error (" & Err.Number & "): " & Err.description & " in LoadScripts()."

    Exit Sub
    
End Sub

Private Function GetFileExtension(ByVal filename As String)

    On Error Resume Next

    ' ...
    If (InStr(1, filename, ".") <> 0) Then
        GetFileExtension = _
            Mid$(filename, InStr(1, filename, ".") + 1)
    End If

End Function

Private Function IsValidFileExtension(ByVal ext As String) As Boolean

    Dim exts() As String  ' ...
    Dim i      As Integer ' ...

    ' ...
    ReDim exts(0 To 2)
    
    ' ...
    exts(0) = "dat"
    exts(1) = "txt"
    exts(2) = "vbs"
    
    ' ...
    For i = 0 To UBound(exts) - 1
        If (StrComp(ext, exts(i), vbTextCompare) = 0) Then
            IsValidFileExtension = True
            
            Exit Function
        End If
    Next i
    
    IsValidFileExtension = False

End Function

Private Function CleanFileName(ByVal filename As String) As String
    
    CleanFileName = filename
    
    CleanFileName = Replace(CleanFileName, " ", "_")
    CleanFileName = Replace(CleanFileName, ".", "_")
    
End Function

Public Sub InitScripts()
    
    Dim i   As Integer ' ...
    Dim str As String  ' ...
    
    For i = 1 To frmChat.SControl.Modules.Count
        If (i > 1) Then
            str = _
                frmChat.SControl.Modules(i).CodeObject.GetSettingsEntry("Enabled")
        End If
    
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            InitScript frmChat.SControl.Modules(i)
        End If
    Next i

End Sub

Public Sub InitScript(ByRef SCModule As Module)

    On Error GoTo ERROR_HANDLER

    Dim i As Integer ' ...

    ' ...
    SCModule.Run "Event_Load"

    ' ...
    If (g_Online) Then
        SCModule.Run "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        SCModule.Run "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If (g_Channel.Users.Count > 0) Then
            For i = 1 To g_Channel.Users.Count
                With g_Channel.Users(i)
                     SCModule.Run "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
                        .game, False
                End With
             Next i
         End If
    End If
    
    Exit Sub
    
ERROR_HANDLER:

    ' object does not support property or method - function missing
    If (Err.Number = 438) Then
        Err.Clear
    
        Resume Next
    End If
    
    ' path not found - deletion of running scripts?
    If (Err.Number = 76) Then
        Err.Clear
    
        Resume Next
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in InitScript()."
    
    Exit Sub

End Sub

Public Sub RunInAll(ParamArray Parameters() As Variant)

    On Error GoTo ERROR_HANDLER

    Dim SC    As ScriptControl
    Dim i     As Integer ' ...
    Dim arr() As Variant ' ...
    Dim str   As String  ' ...
    
    ' ...
    Set SC = frmChat.SControl
    
    ' ...
    arr() = Parameters()

    ' ...
    For i = 1 To SC.Modules.Count
        If (i > 1) Then
            str = _
                SC.Modules(i).CodeObject.GetSettingsEntry("Enabled")
        End If

        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            CallByNameEx SC.Modules(i), "Run", VbMethod, arr()
        End If
    Next

    Exit Sub
    
ERROR_HANDLER:

    ' object does not support property or method - function missing
    If (Err.Number = 438) Then
        Err.Clear
    
        Resume Next
    End If
    
    ' path not found - deletion of running scripts?
    If (Err.Number = 76) Then
        Err.Clear
    
        Resume Next
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in RunInAll()."
    
    Exit Sub
    
End Sub

Public Function CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray _
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
        CallByNameEx = oTLI.InvokeHook(obj, ProcID, CallType)
    End If
    
    If (IsArray(vArgsArray)) Then
        numArgs = UBound(vArgsArray)
        
        ReDim v(numArgs)
        
        For i = 0 To numArgs
            v(i) = vArgsArray(numArgs - i)
        Next i
        
        CallByNameEx = oTLI.InvokeHookArray(obj, ProcID, CallType, v)
    End If
    
    Set oTLI = Nothing
    
    Exit Function

ERROR_HANDLER:
    ' ...
    If (frmChat.SControl.Error) Then
        Exit Function
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.description & _
        " in CallByNameEx()."
        
    Set oTLI = Nothing
        
    Exit Function
    
End Function

Public Function Objects(objIndex As Integer) As scObj

    Objects = m_arrObjs(objIndex)

End Function

Public Function ObjCount(Optional ObjType As String) As Integer
    
    Dim i As Integer ' ...

    If (ObjType <> vbNullString) Then
        For i = 0 To m_objCount - 1
            If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
                ObjCount = (ObjCount + 1)
            End If
        Next i
    Else
        ObjCount = m_objCount
    End If

End Function

Public Function CreateObjEx(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    On Error Resume Next

    Dim obj As scObj ' ...
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim i As Integer ' loop counter variable

        For i = 0 To m_objCount - 1
            If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
                If (StrComp(m_arrObjs(i).ObjType, ObjType, vbTextCompare) = 0) Then
                    If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                        Set CreateObjEx = m_arrObjs(i).obj
                    
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
            
        ' i don't think menus are going to work :|
        'Case "MENU"
        '    If (ObjCount(ObjType) > 0) Then
        '        Load frmChat.mnuScript(ObjCount(ObjType))
        '    End If
        '
        '    Set obj.obj = _
        '            frmChat.mnuScript(ObjCount(ObjType))
            
    End Select

    ' store object
    m_arrObjs(m_objCount) = obj
    
    ' increment object counter
    m_objCount = (m_objCount + 1)
    
    ' create class variable for object
    SCModule.ExecuteStatement "Set " & ObjName & " = GetObjByName(" & Chr$(34) & _
        ObjName & Chr$(34) & ")"

    ' return object
    Set CreateObjEx = obj.obj

End Function

Public Sub DestroyObjEx(ByVal SCModule As Module, ByVal ObjName As String)

    On Error GoTo ERROR_HANDLER

    Dim i     As Integer ' ...
    Dim Index As Integer ' ...
    
    ' ...
    If (m_objCount = 0) Then
        Exit Sub
    End If
    
    ' ...
    Index = m_objCount
    
    ' ...
    For i = 0 To m_objCount - 1
        If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                Index = i
            
                Exit For
            End If
        End If
    Next i
    
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
            m_arrObjs(Index).obj.DestroyObjs
            
            UnhookWindowProc m_arrObjs(Index).obj.hWnd
            
            Unload m_arrObjs(Index).obj
    End Select

    ' ...
    Set m_arrObjs(Index).obj = Nothing
    
    ' ...
    If (Index < m_objCount) Then
        For i = Index To ((m_objCount - 1) - 1)
            m_arrObjs(i) = m_arrObjs(i + 1)
        Next i
    End If
    
    ' ...
    ReDim Preserve m_arrObjs(0 To m_objCount - 1)
    
    ' ...
    m_objCount = (m_objCount - 1)
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in DestroyObj()."
        
    Resume Next
    
End Sub

Public Function GetObjByNameEx(ByRef SCModule As Module, ByVal ObjName As String) As Object

    Dim i As Integer ' ...
    
    ' ...
    For i = 0 To m_objCount - 1
        If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                Set GetObjByNameEx = m_arrObjs(i).obj

                Exit Function
            End If
        End If
    Next i

End Function

Public Function GetSCObjByIndexEx(ByVal ObjType As String, ByVal Index As Integer) As scObj

    Dim i As Integer ' ...

    For i = 0 To ObjCount() - 1
        If (StrComp(ObjType, Objects(i).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(i).obj.Index = Index) Then
                GetSCObjByIndexEx = m_arrObjs(i)
                
                Exit For
            End If
        End If
    Next i

End Function

Public Sub DestroyObjs(Optional ByVal SCModule As Object = Nothing)

    On Error GoTo ERROR_HANDLER

    Dim i As Integer ' ...
    
    ' ...
    For i = m_objCount - 1 To 0 Step -1
        If (SCModule Is Nothing) Then
            DestroyObjEx m_arrObjs(i).SCModule, m_arrObjs(i).ObjName
        Else
            If (SCModule.Name = m_arrObjs(i).SCModule.Name) Then
                DestroyObjEx m_arrObjs(i).SCModule, m_arrObjs(i).ObjName
            End If
        End If
    Next i
    
    Exit Sub
    
ERROR_HANDLER:
    
    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in DestroyObjs()."
        
    Resume Next
    
End Sub

Public Sub SetVeto(ByVal B As Boolean)

    VetoNextMessage = B
    
End Sub

Public Function GetVeto() As Boolean

    GetVeto = VetoNextMessage
    
    VetoNextMessage = False
    
End Function
