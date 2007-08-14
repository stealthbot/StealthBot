Attribute VB_Name = "modScripting"
'/* Scripting.bas
' * ~~~~~~~~~~~~~~~~
' * StealthBot VBScript support module
' * ~~~~~~~~~~~~~~~~
' */
Option Explicit

Private LoadedFilesCount As Long
Private LoadedFilesAtt As Long
Private colFilesToRead As Collection
Public VetoNextMessage As Boolean

'// assumes that the script control is already reset
Public Function LoadScript(ByVal Path As String, ByRef SC As ScriptControl) As String
    
   On Error GoTo LoadScript_Error
    
    Set colFilesToRead = New Collection
    LoadedFilesCount = 0
    LoadedFilesAtt = 0
    
    SC.Reset
    SC.AllowUI = False
    
    If (LenB(ReadCFG("Other", "ScriptAllowUI")) > 0) Then
        SC.AllowUI = True
    End If
    
    If LenB(Dir$(Path)) = 0 Then
        LoadScript = "No script.txt file is present. It must exist, if only to #include other files!"
        Exit Function
    End If
    
    SC.AddObject "ssc", SharedScriptSupport, True
    SC.AddObject "scTimer", frmChat.scTimer
    SC.AddObject "scINet", frmChat.INet
    SC.AddObject "BotVars", BotVars
    
    AddFileToSC Path, SC

LoadScript_Exit:
    LoadScript = "Loaded " & LoadedFilesCount & " of " & LoadedFilesAtt & " script files referenced."
    
    Set colFilesToRead = Nothing
    
    Exit Function

LoadScript_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadScript of Module modScripting"
    Debug.Print "Using variable: " & Path
    Resume LoadScript_Exit
    
End Function

Private Sub AddFileToSC(ByVal Path As String, ByRef SC As ScriptControl)
    Dim f As Integer
    Dim s As String, Temp As String, File As String
    
   On Error GoTo AddFileToSC_Error

    LoadedFilesAtt = LoadedFilesAtt + 1
    
    If LoadedFilesAtt > 300 Then Exit Sub
    
    If Dir$(Path) <> vbNullString And LenB(Path) > 0 Then
        
        f = FreeFile
        
        Open Path For Input As #f
        
        If LOF(f) > 1 Then
            
            LoadedFilesCount = LoadedFilesCount + 1
            
            Do While Not EOF(f)
                Temp = vbNullString
                Line Input #f, Temp
                
                If Len(Temp) > 1 Then
                    
                    '// #include C:\Program Files\Whatever.vbs
                    '// #include scripts\whatever.vbs
                    '// #include whatever.vbs
                
                    If Left$(Temp, 1) = "#" And Len(Temp) > 5 Then '// INCLUDES
                                                
                        File = Mid$(Temp, 10)
                        
                        If InStr(File, ":") = 0 Then
                            File = GetProfilePath() & "\" & File
                        End If
                        
                        colFilesToRead.Add File
                        
                    Else
                    
                        s = s & Temp & vbCrLf
                        'Debug.Print "Added: " & temp
                        
                        If InStr(1, Temp, "End Sub", vbTextCompare) > 0 Then
                            On Error GoTo errorLoadingScript
                            SC.AddCode s
errorLoadingScript:
                            s = vbNullString
                        End If
                        
                    End If
                End If
            Loop
            
        End If
        
        Close #f
        
        If colFilesToRead.Count > 0 Then
            s = colFilesToRead.Item(1)
            colFilesToRead.Remove 1
            AddFileToSC s, SC
        End If
        
    End If

AddFileToSC_Exit:
   Exit Sub

AddFileToSC_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure AddFileToSC of Module modScripting"
    Resume AddFileToSC_Exit
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
