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

    Dim strPath  As String  ' ...
    Dim filename As String  ' ...
    Dim fileExt  As String  ' ...
    Dim I        As Integer ' ...
    
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
            If (IsValidFileExtension(GetFileExtension(filename))) Then
                ' ...
                Set CurrentModule = SC.Modules.Add(filename)
                
                ' ...
                FileToModule CurrentModule, strPath & filename
            End If
            
            ' ...
            filename = Dir()
        Loop
    End If
    
    ' ********************************
    '      LOAD PLUGIN SYSTEM
    ' ********************************

    ' ...
    If (ReadINI("Override", "DisablePS", GetConfigFilePath()) <> "Y") Then
        ' ...
        boolOverride = False
    
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
    Else
        boolOverride = True
    End If
        
    ' ...
    Exit Sub

' ...
ERROR_HANDLER:

    ' ...
    frmChat.AddChat vbRed, "Error: " & Err.description & " in LoadScripts()."

    ' ...
    Exit Sub

End Sub

Private Function FileToModule(ByRef ScriptModule As Module, ByVal filePath As String)

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
                                    FileToModule ScriptModule, filePath
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
    CreateDefautModuleProcs ScriptModule

    ' ...
    ScriptModule.AddCode strContent
    
    Exit Function
    
ERROR_HANDLER:

    frmChat.AddChat vbRed, _
        "Error (#" & Err.Number & "): " & Err.description & " in FileToModule()."
        
    Exit Function

End Function

Private Sub CreateDefautModuleProcs(ByRef ScriptModule As Module)

    Dim str As String ' storage buffer for module code
    
    ' GetModuleName() module-level function
    str = str & "Function GetModuleName()" & vbNewLine
    str = str & "   GetModuleName = " & Chr$(34) & ScriptModule.Name & Chr$(34) & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' GetScriptName() module-level function
    str = str & "Function GetScriptName()" & vbNewLine
    str = str & "   On Error Resume Next" & vbNewLine
    str = str & "   GetScriptName = Name()" & vbNewLine
    str = str & "   If (LenB(GetScriptName) = 0) Then" & vbNewLine
    str = str & "      GetScriptName = GetModuleName()" & vbNewLine
    str = str & "   End If" & vbNewLine
    str = str & "   Err.Clear" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' CreateObj() module-level function
    str = str & "Function CreateObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Set CreateObj = _ " & vbNewLine
    str = str & "         CreateObjEx(GetModuleName(), ObjType, ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine

    ' DeleteObj() module-level function
    str = str & "Sub DeleteObj(ObjType, ObjName)" & vbNewLine
    str = str & "   DeleteObjEx GetModuleName(), ObjType, ObjName" & vbNewLine
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
    Dim I      As Integer ' ...

    ' ...
    ReDim exts(0 To 2)
    
    ' ...
    exts(0) = "dat"
    exts(1) = "txt"
    exts(2) = "vbs"
    
    ' ...
    For I = 0 To UBound(exts) - 1
        If (StrComp(ext, exts(I), vbTextCompare) = 0) Then
            IsValidFileExtension = True
            
            Exit Function
        End If
    Next I
    
    IsValidFileExtension = False

End Function

Public Function InitScripts()

    Dim I As Integer ' ...

    ' ...
    RunInAll "Event_Load"

    ' ...
    If (g_Online) Then
        RunInAll "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        RunInAll "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If (g_Channel.Users.Count > 0) Then
            For I = 1 To g_Channel.Users.Count
                With g_Channel.Users(I)
                     RunInAll "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
                        .game, False
                End With
             Next I
         End If
    End If

End Function

Public Sub RunInAll(ParamArray Parameters() As Variant)

    On Error GoTo ERROR_HANDLER

    Dim SC    As ScriptControl
    Dim I     As Integer ' ...
    Dim arr() As Variant ' ...
    Dim str   As String  ' ...
    
    ' ...
    Set SC = frmChat.SControl
    
    ' ...
    arr() = Parameters()

    ' ...
    For I = 1 To SC.Modules.Count
        If (I > 1) Then
            str = _
                SC.Modules(I).CodeObject.GetSettingsEntry("Enabled")
        End If
            
        If (StrComp(str, "False", vbTextCompare) <> 0) Then
            CallByNameEx SC.Modules(I), "Run", VbMethod, arr()
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
    Dim I       As Long
    Dim v()     As Variant
    
    Set oTLI = New TLIApplication

    ProcID = oTLI.InvokeID(obj, ProcName)

    If (IsMissing(vArgsArray)) Then
        CallByNameEx = oTLI.InvokeHook(obj, ProcID, CallType)
    End If
    
    If (IsArray(vArgsArray)) Then
        numArgs = UBound(vArgsArray)
        
        ReDim v(numArgs)
        
        For I = 0 To numArgs
            v(I) = vArgsArray(numArgs - I)
        Next I
        
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
    
    Dim I As Integer ' ...

    If (ObjType <> vbNullString) Then
        For I = 0 To m_objCount - 1
            If (StrComp(ObjType, m_arrObjs(I).ObjType, vbTextCompare) = 0) Then
                ObjCount = (ObjCount + 1)
            End If
        Next I
    Else
        ObjCount = m_objCount
    End If

End Function

Public Function CreateObjEx(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    Dim obj As scObj ' ...
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim I As Integer ' loop counter variable

        For I = 0 To m_objCount - 1
            If (m_arrObjs(I).SCModule.Name = SCModule.Name) Then
                If (StrComp(m_arrObjs(I).ObjType, ObjType, vbTextCompare) = 0) Then
                    If (StrComp(m_arrObjs(I).ObjName, ObjName, vbTextCompare) = 0) Then
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

Public Sub DeleteObjEx(ByRef SCModule As Module, ByVal TimerName As String)

    
End Sub

Public Function GetObjByNameEx(ByRef SCModule As Module, ByVal ObjName As String) As Object

    Dim I As Integer ' ...
    
    ' ...
    For I = 0 To m_objCount - 1
        If (m_arrObjs(I).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(I).ObjName, ObjName, vbTextCompare) = 0) Then
                Set GetObjByNameEx = m_arrObjs(I).obj

                Exit Function
            End If
        End If
    Next I

End Function

Public Function GetSCObjByIndexEx(ByVal ObjType As String, ByVal Index As Integer) As scObj

    Dim I As Integer ' ...

    For I = 0 To ObjCount() - 1
        If (StrComp(ObjType, Objects(I).ObjType, vbTextCompare) = 0) Then
            If (m_arrObjs(I).obj.Index = Index) Then
                GetSCObjByIndexEx = m_arrObjs(I)
                
                Exit For
            End If
        End If
    Next I

End Function

Private Sub DestroyObjs()

    On Error GoTo ERROR_HANDLER

    Dim I As Integer ' ...
    
    ' ...
    For I = m_objCount - 1 To 0 Step -1
        ' ...
        Select Case (UCase$(m_arrObjs(I).ObjType))
            Case "TIMER"
                If (m_arrObjs(I).obj.Index > 0) Then
                    Unload frmChat.tmrScript(m_arrObjs(I).obj.Index)
                Else
                    frmChat.tmrScript(0).Enabled = False
                End If
                
            Case "LONGTIMER"
                If (m_arrObjs(I).obj.Index > 0) Then
                    Unload frmChat.tmrScriptLong(m_arrObjs(I).obj.Index)
                Else
                    frmChat.tmrScriptLong(0).Enabled = False
                End If
                
            Case "WINSOCK"
                If (m_arrObjs(I).obj.Index > 0) Then
                    Unload frmChat.sckScript(m_arrObjs(I).obj.Index)
                Else
                    frmChat.sckScript(0).Close
                End If
                
            Case "INET"
                If (m_arrObjs(I).obj.Index > 0) Then
                    Unload frmChat.itcScript(m_arrObjs(I).obj.Index)
                Else
                    frmChat.itcScript(0).Cancel
                End If
                
            Case "FORM"
                m_arrObjs(I).obj.DestroyObjs
                
                Unload m_arrObjs(I).obj
                
        End Select

        ' ...
        Set m_arrObjs(I).obj = Nothing
    Next I
    
    m_objCount = 0
    
    ReDim m_arrObjs(m_objCount)
    
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
