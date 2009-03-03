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

Public dictSettings      As Dictionary
Public dictTimerInterval As Dictionary
Public dictTimerEnabled  As Dictionary
Public dictTimerCount    As Dictionary
Public VetoNextMessage   As Boolean
Public boolOverride      As Boolean

Private m_arrObjs()      As scObj
Private m_objCount       As Integer

Public Sub InitScriptControl(ByRef SC As ScriptControl)

    ' ...
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
    Dim i        As Integer ' ...
    
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
            Set CurrentModule = SC.Modules.Add(filename)
            
            ' ...
            FileToModule CurrentModule, strPath & filename
    
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
    frmChat.AddChat vbRed, "Error: " & Err.Description & " in LoadScripts()."

    ' ...
    Exit Sub

End Sub

Private Function FileToModule(ByRef ScriptModule As Module, ByVal filePath As String)

    Dim strLine    As String  ' ...
    Dim strContent As String  ' ...
    Dim f          As Integer ' ...
    
    ' ...
    f = FreeFile

    ' ...
    Open filePath For Input As #f
        ' ...
        Do While (EOF(f) = False)
            ' ...
            Line Input #f, strLine
            
            ' ...
            If (Len(strLine) > 1) Then
                ' ...
                If (StrComp(Left$(Trim(LCase$(strLine)), 8), "#include", vbTextCompare) <> 0) Then
                    strContent = strContent & strLine & vbCrLf
                Else
                    ' ...
                    If (Len(Trim(LCase$(strLine))) > 8) Then
                        FileToModule ScriptModule, Mid$(Trim(LCase$(strLine)), 9)
                    End If
                End If
            End If
            
            ' ...
            strLine = vbNullString
        Loop
    Close #f
    
    ' ...
    CreateDefautModuleProcs ScriptModule

    ' ...
    ScriptModule.AddCode strContent

End Function

Private Sub CreateDefautModuleProcs(ByRef ScriptModule As Module)

    Dim str As String ' storage buffer for module code

    ' GetModuleName() module-level function
    str = str & "Function GetModuleName()" & vbNewLine
    str = str & "   GetModuleName = " & Chr$(34) & ScriptModule.Name & Chr$(34) & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' CreateObj() module-level function
    str = str & "Function CreateObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Set CreateObj = _ " & vbNewLine
    str = str & "         CreateObjEx(GetModuleName(), ObjType, ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' DeleteObj() module-level function
    str = str & "Sub DeleteObj(ObjType, ObjName)" & vbNewLine
    str = str & "   Call DeleteObjEx(GetModuleName(), ObjType, ObjName)" & vbNewLine
    str = str & "End Sub" & vbNewLine
    
    ' GetObjByName() module-level function
    str = str & "Function GetObjByName(ObjType, ObjName)" & vbNewLine
    str = str & "   Set GetObjByName = _ " & vbNewLine
    str = str & "         GetObjByNameEx(GetModuleName(), ObjType, ObjName)" & vbNewLine
    str = str & "End Function" & vbNewLine
    
    ' store module-level coding
    ScriptModule.AddCode str
    
End Sub

Public Function InitScripts()

    Dim i As Integer ' ...

    ' ...
    RunInAll "Event_Load"

    ' ...
    If (g_Online) Then
        RunInAll "Event_LoggedOn", GetCurrentUsername, BotVars.Product
        RunInAll "Event_ChannelJoin", g_Channel.Name, g_Channel.Flags

        If (g_Channel.Users.Count > 0) Then
            For i = 1 To g_Channel.Users.Count
                With g_Channel.Users(i)
                     RunInAll "Event_UserInChannel", .DisplayName, .Flags, .Stats.ToString, .Ping, _
                        .game, False
                End With
             Next i
         End If
    End If

End Function

Public Sub RunInAll(ParamArray Parameters() As Variant)

    On Error GoTo ERROR_HANDLER

    Dim SC    As ScriptControl
    Dim i     As Integer ' ...
    Dim arr() As Variant ' ...
    
    ' ...
    Set SC = frmChat.SControl
    
    ' ...
    arr() = Parameters()

    ' ...
    For i = 1 To SC.Modules.Count
        CallByNameEx SC.Modules(i), "Run", VbMethod, arr()
    Next

    Exit Sub
    
ERROR_HANDLER:
    ' object does not support property or method
    If (Err.Number = 438) Then
        Resume Next
    End If

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.Description & _
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

    frmChat.AddChat vbRed, "Error (#" & Err.Number & "): " & Err.Description & _
        " in CallByNameEx()."
        
    Set oTLI = Nothing
        
    Exit Function
    
End Function

Public Function Objects(objIndex As Integer) As scObj

    Objects = m_arrObjs(objIndex)

End Function

Public Function ObjCount(Optional ObjType As String) As Integer
    
    Dim i As Integer ' ...

    Select Case (UCase$(ObjType))
        Case "TIMER", "WINSOCK", "INET", "FORM", "MENU"
            For i = 0 To m_objCount - 1
                If (StrComp(ObjType, m_arrObjs(i).ObjType, vbTextCompare) = 0) Then
                    ObjCount = (ObjCount + 1)
                End If
            Next i
            
        Case Else
            ObjCount = m_objCount
    End Select

End Function

Public Function CreateObjEx(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    Dim obj As scObj ' ...
    
    ' redefine array size & check for duplicate controls
    If (m_objCount) Then
        Dim i As Integer ' loop counter variable

        For i = 0 To m_objCount - 1
            If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
                If (StrComp(m_arrObjs(i).ObjType, ObjType, vbTextCompare) = 0) Then
                    If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
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
        Case "MENU"
    End Select

    ' store object
    m_arrObjs(m_objCount) = obj
    
    ' increment object counter
    m_objCount = (m_objCount + 1)
    
    ' create class variable for object
    SCModule.ExecuteStatement "Set " & ObjName & " = GetObjByName(" & _
        Chr$(34) & ObjType & Chr$(34) & ", " & Chr$(34) & ObjName & Chr$(34) & ")"

    ' return object
    Set CreateObjEx = obj.obj

End Function

Public Sub DeleteObjEx(ByRef SCModule As Module, ByVal TimerName As String)

    
End Sub

Public Function GetObjByNameEx(ByRef SCModule As Module, ByVal ObjType As String, ByVal ObjName As String) As Object

    Dim i As Integer ' ...
    
    ' ...
    For i = 0 To m_objCount - 1
        If (m_arrObjs(i).SCModule.Name = SCModule.Name) Then
            If (StrComp(m_arrObjs(i).ObjType, ObjType, vbTextCompare) = 0) Then
                If (StrComp(m_arrObjs(i).ObjName, ObjName, vbTextCompare) = 0) Then
                    Set GetObjByNameEx = m_arrObjs(i).obj
    
                    Exit Function
                End If
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

Public Sub SetVeto(ByVal B As Boolean)

    VetoNextMessage = B
    
End Sub

Public Function GetVeto() As Boolean

    GetVeto = VetoNextMessage
    
    VetoNextMessage = False
    
End Function

'// Written by Swent. Sets a plugin timer's interval.
Public Sub SetPTInterval(ByVal strPrefix As String, ByVal strTimerName As String, ByVal intInterval As Integer)
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName

    dictTimerInterval(strKey) = intInterval
    dictTimerCount(strKey) = intInterval
    
    If Not dictTimerEnabled.Exists(strKey) Then
       dictTimerEnabled(strKey) = False
    End If
End Sub


'// Written by Swent. Enables or disables a plugin timer.
Public Sub SetPTEnabled(ByVal strPrefix As String, ByVal strTimerName As String, ByVal boolEnabled As Boolean)
    
    dictTimerEnabled(strPrefix & ":" & strTimerName) = boolEnabled
End Sub


'// Written by Swent. Modifies the count in a running plugin timer.
Public Sub SetPTCount(ByVal strPrefix As String, ByVal strTimerName As String, ByVal intCount As Integer)
    
    dictTimerCount(strPrefix & ":" & strTimerName) = intCount
End Sub


'// Written by Swent. Gets the enabled status of a plugin timer.
Public Function GetPTEnabled(ByVal strPrefix As String, ByVal strTimerName As String)
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerEnabled.Exists(strKey) Then
        GetPTEnabled = dictTimerEnabled(strKey)
    Else
        GetPTEnabled = -1
    End If
End Function


'// Written by Swent. Gets a plugin timer's interval setting.
Public Function GetPTInterval(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerInterval.Exists(strKey) Then
        GetPTInterval = dictTimerInterval(strKey)
    Else
        GetPTInterval = -1
    End If
End Function


'// Written by Swent. Get's the seconds left before a plugin timer sub executes.
Public Function GetPTLeft(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName

    If dictTimerCount.Exists(strKey) Then
        GetPTLeft = dictTimerCount(strKey)
    Else
        GetPTLeft = -1
    End If
End Function

'// Written by Swent. Gets the time since a plugin timer sub was last executed.
Public Function GetPTWaiting(ByVal strPrefix As String, ByVal strTimerName As String) As Integer
    Dim strKey As String
    strKey = strPrefix & ":" & strTimerName
    
    If dictTimerCount.Exists(strKey) Then
        GetPTWaiting = dictTimerInterval(strKey) - dictTimerCount(strKey) + 1
    Else
        GetPTWaiting = -1
    End If
End Function


'// Written by Swent. Gets keys for the timer dictionaries.
Public Function GetPTKeys() As String

    GetPTKeys = Join(dictTimerEnabled.Keys)
End Function
