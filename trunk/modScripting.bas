Attribute VB_Name = "modScripting"
'/* Scripting.bas
' * ~~~~~~~~~~~~~~~~
' * StealthBot VBScript support module
' * ~~~~~~~~~~~~~~~~
' * Modified by Swent 10/8/2007
' */
Option Explicit

Public VetoNextMessage As Boolean

'// Loads the Plugin System
'//   Called from Form_Load() and mnuReloadScript_Click() in frmChat
Public Sub LoadPluginSystem(ByRef SC As ScriptControl)
    Dim Path As String, intFile As Integer, strLine As String, strContent As String

   On Error GoTo LoadPluginSystem_Error

    '// Reset the Script Control
    SC.Reset
    
    '// Allow UI's unless they've been disabled by the user
    If ReadINI("Other", "ScriptAllowUI", GetConfigFilePath()) <> "N" Then SC.AllowUI = True

    '// PluginSystem.dat exists?
    Path = GetFilePath("PluginSystem.dat")
    If LenB(Dir$(Path)) = 0 Then
        AddChat vbRed, "No PluginSystem.dat file is present. It must exist in order to load plugins!"
        AddChat vbGreen, "Developers, please use the PS dev code here: http://stealthbot.net/p/Users/Swent/ps-sbdev.txt"
        AddChat vbGreen, "Save it to a text file named ""PluginSystem.dat"" in your trunk folder."
        Exit Sub
    End If
    
    '// Create scripting objects
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars
    
    '// Load PluginSystem.dat
    intFile = FreeFile
    Open Path For Input As #intFile
    
        Do While Not EOF(intFile)
            strLine = vbNullString
            Line Input #intFile, strLine
            If Len(strLine) > 1 Then strContent = strContent & strLine & vbCrLf
        Loop
    Close #intFile
    SC.AddCode strContent

LoadPluginSystem_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadPluginSystem of Module modScripting"
    Debug.Print "Using variable: " & Path
End Sub

Public Sub SetVeto(ByVal b As Boolean)
    VetoNextMessage = b
End Sub

Public Function GetVeto() As Boolean
    GetVeto = VetoNextMessage
    VetoNextMessage = False
End Function

Public Sub ReInitScriptControl(ByRef SC As ScriptControl)
    Dim i As Integer
    Dim Message As String

    On Error GoTo ReInitScriptControl_Error

    BotLoaded = True
    SC.Run "Event_Load"

    If g_Online Then
        SC.Run "Event_LoggedOn", CurrentUsername, BotVars.Product
        SC.Run "Event_ChannelJoin", gChannel.Current, 0

        If colUsersInChannel.Count > 0 Then
            For i = 1 To colUsersInChannel.Count
                Message = ""

                With colUsersInChannel.Item(i)
                     ParseStatstring .Statstring, Message, .Clan

                     SC.Run "Event_UserInChannel", .Username, .Flags, Message, .Ping, .Product, False
                 End With
             Next i
         End If
    End If

    On Error GoTo 0
    Exit Sub

ReInitScriptControl_Error:

    'Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure ReInitScriptControl of Module modScripting"
End Sub
