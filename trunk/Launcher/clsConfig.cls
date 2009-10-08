VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const OBJECT_NAME As String = "clsConfig"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private m_loaded     As Boolean
Private m_configpath As String
Private m_autoclose  As Boolean

Private Sub WriteConfig(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String, ByVal sFile As String)
On Error GoTo ERROR_HANDLER:
    WritePrivateProfileString sSection, sKey, sValue, sFile
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "WriteConfig"
End Sub

Private Function ReadConfig(ByVal sSection As String, ByVal sKey As String, ByVal sFile As String, Optional ByVal sDefault As String = vbNullString) As String
On Error GoTo ERROR_HANDLER:
    Dim sBuffer As String
    Dim lRet    As String
    
    If (LenB(Dir(sFile)) > 0) Then
        sBuffer = String(255, vbNull)
        lRet = GetPrivateProfileString(sSection, sKey, sDefault, sBuffer, 255, sFile)
        ReadConfig = Left$(sBuffer, lRet)
    Else
        ReadConfig = sDefault
    End If
    Exit Function
    
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "ReadConfig"
    ReadConfig = sDefault
End Function

Private Function ReadBoolean(ByVal sSection As String, ByVal sKey As String, ByVal bDefault As Boolean) As Boolean
On Error GoTo ERROR_HANDLER:
    Dim sValue As String
    sValue = ReadConfig(sSection, sKey, m_configpath, IIf(bDefault, "True", "False"))
    
    ReadBoolean = (StrComp(sValue, "True", vbTextCompare) = 0)
    
    Exit Function
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "ReadBoolean"
    ReadBoolean = bDefault
End Function

Public Sub LoadConfig()
On Error GoTo ERROR_HANDLER:
    m_configpath = ReplaceEnvironmentVars("%APPDATA%\StealthBot\Launcher.ini")
    
    m_autoclose = ReadBoolean("Launcher", "AutoClose", False)
    
    m_loaded = True
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "LoadConfig"
End Sub

Public Sub SaveConfig()
On Error GoTo ERROR_HANDLER:
    WriteConfig "Launcher", "AutoClose", IIf(m_autoclose, "True", "False"), m_configpath
    
    Exit Sub
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "SaveConfig"
End Sub


Public Property Let AutoClose(ByVal bData As Boolean)
On Error GoTo ERROR_HANDLER:
    m_autoclose = bData
    
    Exit Property
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "AutoClose Let"
End Property

Public Property Get AutoClose() As Boolean
On Error GoTo ERROR_HANDLER:
    If (Not m_loaded) Then LoadConfig
    AutoClose = m_autoclose
    
    Exit Property
ERROR_HANDLER:
    ErrorHandler Err.Number, OBJECT_NAME, "AutoClose Get"
End Property
