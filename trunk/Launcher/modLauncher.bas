Attribute VB_Name = "modLauncher"
Option Explicit

'Helper modules for the launcher
Private Const OBJECT_NAME As String = "modLauncher"

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Const NORMAL_PRIORITY_CLASS As Long = &H20

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, _
    ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private xml_doc As DOMDocument60
Private CommandLine As String
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Public cConfig    As clsConfig
Public bIsClosing As Boolean

Public Sub ErrorHandler(lError As Long, sObjectName As String, sFunctionName As String)
    Dim sPath As String
    
    
    'AddChat vbRed, StringFormat("Error #{0}: {1} in {2}.{3}()", lError, Error(lError), sObjectName, sFunctionName)
    
    sPath = ReplaceEnvironmentVars("%APPDATA%\StealthBot\LauncherErrors.txt")
    If (LenB(Dir$(sPath)) = 0) Then
        Open sPath For Output As #1: Close #1
    End If
    
    Open sPath For Append As #1
        Print #1, StringFormat("Error #{0}: {1} in {2}.{3}()", lError, Error(lError), sObjectName, sFunctionName)
    Close #1
    
    Err.Clear
End Sub

Public Function StringFormat(source As String, ParamArray params() As Variant)
On Error GoTo ERROR_HANDLER:

    Dim retVal As String, i As Integer
    retVal = source
    For i = LBound(params) To UBound(params)
        retVal = Replace(retVal, "{" & i & "}", CStr(params(i)))
    Next
    StringFormat = retVal
    
    Exit Function
    
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "StringFormat"
    StringFormat = vbNullString
End Function

Public Function ReplaceEnvironmentVars(ByVal str As String) As String
On Error GoTo ERROR_HANDLER:

    Dim i     As Integer
    Dim Name  As String
    Dim Value As String
    Dim tmp   As String
    
    tmp = str
    
    i = 1

    While (LenB(Environ$(i)) > 0)
        Name = Mid$(Environ$(i), 1, InStr(1, Environ$(i), "=") - 1)
        Value = Mid$(Environ$(i), InStr(1, Environ$(i), "=") + 1)
        tmp = Replace(tmp, "%" & Name & "%", Value)
        i = i + 1
    Wend
    ReplaceEnvironmentVars = tmp
    
    Exit Function
    
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "ReplaceEnvironmentVars"
    ReplaceEnvironmentVars = vbNullString
End Function

Public Function MakeDirectory(sPath As String) As Boolean
On Error GoTo ERROR_HANDLER
    MkDir sPath
    MakeDirectory = True
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "MakeDirectory"
    MakeDirectory = False
End Function

Public Sub LoadXMLDocument()
On Error GoTo ERROR_HANDLER:
Exit Sub
    If (Not xml_doc Is Nothing) Then
        Set xml_doc = Nothing
    End If
    
    Set xml_doc = New DOMDocument60
    If (Not xml_doc.Load(App.Path & "\Launcher.xml")) Then
        MsgBox "Failed to load Launcher.xml"
    End If
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "LoadXMLDocument"
End Sub

Public Function CopyProfileFiles(sProfile As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sRootPath As String
    
    sRootPath = StringFormat("{0}\StealthBot\{1}", ReplaceEnvironmentVars("%APPDATA%"), sProfile)
    
    'Copy over the \Default\ Directory if it exists.
    If (LenB(Dir$(StringFormat("{0}\Default\", App.Path), vbDirectory)) > 0) Then
        If (Not CopyFolder(StringFormat("{0}\Default", App.Path), sRootPath)) Then
            'could not copy default folder
        End If
    End If
    
    CopyProfileFiles = True
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "CopyProfileFiles"
End Function

Public Function KillFolder(ByVal FullPath As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim oFso As Object
    Set oFso = CreateObject("Scripting.FileSystemObject")

    If Right(FullPath, 1) = "\" Then FullPath = Left$(FullPath, Len(FullPath) - 1)

    If oFso.FolderExists(FullPath) Then
        Dir App.Path  'Use App.Path because that *should* always exist unless some voodo was performed. But 'C:/' is not garenteed.
        oFso.DeleteFolder FullPath, True
        KillFolder = (Err.Number = 0 And oFso.FolderExists(FullPath) = False)
    Else
        KillFolder = True
    End If

    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "KillFolder"
    KillFolder = False
End Function

Public Function LaunchProfile(sProfile As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim lRet     As Long
    Dim sPath    As String
    Dim security As SECURITY_ATTRIBUTES
    Dim suInfo   As STARTUPINFO
    Dim pInfo    As PROCESS_INFORMATION
    
    If (Not ProfileExists(sProfile)) Then Exit Function
    
    sPath = StringFormat(ReplaceEnvironmentVars("%APPDATA%\StealthBot\{0}\"), sProfile)
    lRet = CreateProcess(StringFormat("{0}\StealthBot v2.7.exe", App.Path), _
      StringFormat("-addpath {0}{1}{0}", Chr$(34), App.Path), _
      security, security, False, _
      NORMAL_PRIORITY_CLASS, _
      ByVal 0&, sPath, suInfo, pInfo)
      
    If (cConfig Is Nothing) Then Set cConfig = New clsConfig
    If (cConfig.AutoClose) Then Unload frmLauncher


    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "LaunchProfile"
    LaunchProfile = False
End Function

Private Function GetDesktopPath() As String
On Error GoTo ERROR_HANDLER:
    Dim oShell As Object
    Dim sPath  As String
    
    Set oShell = CreateObject("WScript.Shell")
    sPath = oShell.SpecialFolders("Desktop")
    If (Not Right$(sPath, 1) = "\") Then sPath = sPath & "\"
    
    
    GetDesktopPath = sPath
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "GetDesktopPath"
End Function

Public Function CreateShortcut(sProfile As String)
On Error GoTo ERROR_HANDLER:
    Dim oShell    As Object
    Dim oShortCut As Object
    Dim sDesktop  As String
    Dim sPath     As String
    
    sDesktop = GetDesktopPath
    If (LenB(sDesktop) = 0) Then
        MsgBox "Failed to get desktop folder."
        Exit Function
    End If
    
    sPath = StringFormat("{0}StealthBot - {1}.lnk", sDesktop, sProfile)
    
    Set oShell = CreateObject("WScript.Shell")
    Set oShortCut = oShell.CreateShortcut(sPath)
    With oShortCut
        .TargetPath = StringFormat("{0}{1}\{2}.exe{0}", Chr$(34), App.Path, App.EXEName)
        .Arguments = StringFormat("-LaunchProfile {0}{1}{0}", Chr$(34), sProfile)
        .save
    End With
    
    MsgBox StringFormat("Created shortcut for profile {0}{1}{0} on your desktop.{2}{3}", Chr$(34), sProfile, vbNewLine, sPath)
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "CreateShortcut"
End Function

Public Function ProfileExists(sProfile As String)
On Error GoTo ERROR_HANDLER:
    Dim sPath As String
    
    If (LenB(sProfile) = 0) Then
        ProfileExists = False
        Exit Function
    End If
    
    sPath = StringFormat(ReplaceEnvironmentVars("%APPDATA%\StealthBot\{0}\"), sProfile)
    
    ProfileExists = LenB(Dir$(sPath, vbDirectory)) > 0
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "ProfileExists"
End Function

Public Function CreateProfile(sName As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sPath As String
    CreateProfile = False
    
    sPath = StringFormat("{0}\StealthBot\{1}", ReplaceEnvironmentVars("%APPDATA%"), sName)
    
    If (ProfileExists(sName)) Then
        MsgBox "Profile already exists!"
        Exit Function
    End If
    
    If (Not MakeDirectory(sPath)) Then
        MsgBox "Error creating profile directory."
        Exit Function
    End If
    
    If (Not CopyProfileFiles(sName)) Then
        MsgBox "Failed to copy profile files over."
        KillFolder sPath
        Exit Function
    End If
    
    CreateProfile = True
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "CreateProfile"
End Function

Public Function RemoveProfile(ByRef Item As ListItem)
On Error GoTo ERROR_HANDLER
    Dim lRet As Long
    Dim sProfile As String
    
    sProfile = Item.Text
    
    lRet = MsgBox(StringFormat("This will delete EVERYTHING in the {0}{1}{0} profile. Are you sure?", Chr$(34), sProfile), vbYesNoCancel + vbQuestion)
    
    If (lRet = vbYes) Then
        If (KillFolder(StringFormat(ReplaceEnvironmentVars("%APPDATA%\StealthBot\{0}"), sProfile))) Then
            frmLauncher.UnlistProfile Item.Index
        Else
            MsgBox "Failed to delete the profile. It may be in use by an application. Try rebooting your computer and deleting it again.", vbInformation + vbOKOnly
            Exit Function
        End If
    End If
    
    RemoveProfile = True
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "RemoveProfile"
End Function

Private Function StripString(ByRef sTemp As String) As String
On Error GoTo ERROR_HANDLER
    Dim sValue   As String
    
    If (Left$(sTemp, 1) = Chr$(34)) Then
        If (InStr(2, sTemp, Chr$(34), vbTextCompare) > 0) Then
            sValue = Mid$(sTemp, 2, InStr(2, sTemp, Chr$(34), vbTextCompare) - 2)
            sTemp = Mid$(sTemp, Len(sValue) + 4)
        Else
            sValue = Mid$(Split(sTemp & " -", " -")(0), 2)
            sTemp = Mid$(sTemp, Len(sValue) + 3)
        End If
    Else
        sValue = Split(sTemp & " -", " -")(0)
        sTemp = Mid$(sTemp, Len(sValue) + 2)
    End If
    StripString = sValue
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "StripString"
End Function

Public Function SetCommandLine(sCommandLine As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sTemp    As String
    Dim sSetting As String
    Dim sValue   As String
    CommandLine = vbNullString
    sTemp = sCommandLine
    
    Do While Left$(sTemp, 1) = "-"
        sSetting = Split(Mid$(sTemp, 2) & Space$(1), Space$(1))(0)
        sTemp = Mid$(sTemp, Len(sSetting) + 3)
        Select Case LCase$(sSetting)
        
            Case "launchprofile":
                sValue = StripString(sTemp)
                If (Not LenB(sValue) = 0) Then
                    If (Not ProfileExists(sValue)) Then
                        MsgBox StringFormat("The Profile {0}{1}{0} does not exist!", Chr$(34), sValue)
                    Else
                        LaunchProfile sValue
                        SetCommandLine = True
                        Exit Function
                    End If
                End If
                
            Case Else:
                CommandLine = StringFormat("{0}-{1} ", CommandLine, sSetting)
        End Select
    Loop
    Exit Function
ERROR_HANDLER:
    SetCommandLine = False
    ErrorHandler Err.Number, OBJECT_NAME, "SetCommandLine"
End Function

Public Function CopyFolder(sSource As String, sDest As String) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sFile     As String
    Dim sSourcePath As String
    Dim sDestPath   As String
    Dim sFiles      As New Collection
    Dim X           As Integer
    
    CopyFolder = False
    
    If (LenB(Dir$(sDest, vbDirectory)) = 0) Then
        If (Not MakeDirectory(sDest)) Then Exit Function
    End If
    
    If (LenB(Dir$(StringFormat("{0}\", sSource), vbDirectory)) = 0) Then Exit Function
    
    Do While True
         sFile = Dir$
         If (LenB(sFile) = 0) Then Exit Do
         If (Not sFile = "..") Then sFiles.Add sFile
    Loop
    
    For X = 1 To sFiles.Count
        sFile = sFiles.Item(X)
        sSourcePath = StringFormat("{0}\{1}", sSource, sFile)
        sDestPath = StringFormat("{0}\{1}", sDest, sFile)
        If ((GetFileAttributes(sSourcePath) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then
            If (Not CopyFolder(sSourcePath, sDestPath)) Then
                KillFolder sDest
                Exit Function
            End If
        Else
            Call FileCopy(sSourcePath, sDestPath)
            If (LenB(Dir$(sDestPath)) = 0) Then
                KillFolder sDest
                Exit Function
            End If
        End If
    Next X
    
    CopyFolder = True
        
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "CopyFolder"
    CopyFolder = False
End Function

'Public Sub AddChat(ParamArray saElements() As Variant)
'On Error GoTo ERROR_HANDLER:
'    Dim i As Integer
'    With frmStatus.rtbStatus
'        If (Len(.Text) > &H4000) Then
'            .SelStart = 0
'            .SelLength = &H100
'            .SelText = vbNullString
'        End If
'
'        .SelStart = Len(.Text)
'        .SelLength = 0
'        .SelColor = vbWhite
'        .SelText = StringFormat("[{0}] ", Time)
'        .SelStart = Len(.Text)
'
'        For i = LBound(saElements) To UBound(saElements) Step 2
'            .SelStart = Len(.Text)
'            .SelLength = 0
'            .SelColor = saElements(i)
'            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
'            .SelStart = Len(.Text)
'        Next i
'    End With
'    Exit Sub
'ERROR_HANDLER:
'    If (Err.Number = 13 Or Err.Number = 91) Then Exit Sub
'    ErrorHandler Err.Number, OBJECT_NAME, "AddChat"
'End Sub

'Public Function GetWebPath()
'    GetWebPath = "http://www.StealthBot.net/sb/Launcher/"
'End Function

Public Function ReplaceVars(sString As String) As String
    sString = Replace$(sString, "{PROFILEPATH}", "%APPDATA\StealthBot")
    sString = ReplaceEnvironmentVars(sString)
    ReplaceVars = sString
End Function

'Public Sub CheckForUpdates()
'On Error GoTo ERROR_HANDLER:
'
'    Dim sTemp As String
'    Dim i     As Integer
'    Dim sCRC  As String
'
'    With frmLauncher.Inet
'
'        sTemp = .OpenURL(StringFormat("{0}?p=lnews", GetWebPath))
'        AddChat vbGreen, StringFormat("Launcher news:{0}{1}", vbNewLine, ReplaceVars(sTemp))
'
'        sTemp = .OpenURL(StringFormat("{0}?p=lupdate", GetWebPath))
'
'        i = InStr(sTemp, Chr$(&HFF))
'        If (i = 0) Then
'            AddChat vbRed, "Failed to get launcer update information."
'            Exit Sub
'        End If
'
'        If (Not StrComp(Left$(sTemp, i - 1), StringFormat("{0}.{1}", App.Major, App.Minor), vbTextCompare) = 0) Then
'            sTemp = .OpenURL(StringFormat("{0}?p=latest_url", GetWebPath))
'            AddChat vbGreen, "New updates avalible: ", vbWhite, sTemp
'            Exit Sub
'        End If
'    End With
'    Exit Sub
'ERROR_HANDLER:
'    ErrorHandler Err.Number, OBJECT_NAME, "CheckForUpdates"
'End Sub
