VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'modOSVersion.bas
' project StealthBot
' October 2006 from code at:
'  http://vbnet.mvps.org/index.html?code/helpers/iswinversion.htm

Option Explicit

Private Declare Function GetVersionEx Lib "Kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
  
Private Const VER_NT_WORKSTATION       As Long = 1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = 2
Private Const VER_NT_SERVER            As Long = 3

Private Const VER_PLATFORM_ID_VISTA    As Long = 2
Private Const VER_PLATFORM_WIN32_NT    As Long = 2

Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: string providing arbitrary additional information
End Type

Private m_isCached As Boolean
Private m_osVer    As OSVERSIONINFO

Private Sub Class_Initialize()

    GetVersion

End Sub

Private Sub Class_Terminate()

    m_isCached = False
    
End Sub

Private Sub GetVersion()

    m_osVer.OSVSize = Len(m_osVer)

    If (GetVersionEx(m_osVer) = 1) Then
        m_isCached = True
    End If

End Sub

Public Property Get Name() As String
    
    ' ...
    
End Property

Public Property Get IsWindowsNT() As Boolean

    IsWindowsNT = CBool(m_osVer.PlatformID = VER_PLATFORM_WIN32_NT)
    
End Property

Public Property Get IsWindows95() As Boolean
    
    IsWindows95 = _
        CBool((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 0))
    
End Property

Public Property Get IsWindows98() As Boolean

    If ((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 10)) Then
        If (m_osVer.dwBuildNumber < 2222) Then
            IsWindows98 = True
        End If
    End If
    
End Property

Public Property Get IsWindowsME() As Boolean

    If ((m_osVer.dwVerMajor = 4) And (m_osVer.dwVerMinor = 10)) Then
        If (m_osVer.dwBuildNumber >= 2222) Then
            IsWindowsME = True
        End If
    End If
    
End Property

Public Property Get IsWindows2000() As Boolean

    IsWindows2000 = _
        CBool((m_osVer.dwVerMajor = 5) And (m_osVer.dwVerMinor = 0))
    
End Property


Public Property Get IsWindowsXP() As Boolean

    IsWindowsXP = _
        CBool(((m_osVer.dwVerMajor = 5) And (m_osVer.dwVerMinor = 1)))

    ' check for winxp 64-bit
    If (IsWindowsXP = False) Then
        If ((m_osVer.dwVerMajor = 5) And (m_osVer.dwVerMinor = 2)) Then
            If (m_osVer.PlatformID = VER_NT_WORKSTATION) Then
                IsWindowsXP = True
            End If
        End If
    End If
    
End Property

Public Property Get IsWindows2003() As Boolean

    If (IsWindowsXP = False) Then
        IsWindows2003 = _
            ((m_osVer.dwVerMajor = 5) And (m_osVer.dwVerMinor = 2))
    End If
    
End Property

'Added by FrOzeN on 18th September, 2008.
'Returns true if Vista, false if not.
' updated by eric
Public Property Get IsWindowsVista() As Boolean

    IsWindowsVista = _
        CBool((m_osVer.dwVerMajor = 6) And (m_osVer.PlatformID And VER_NT_WORKSTATION))
    
End Property

Public Property Get IsWindows2008() As Boolean

    If (IsWindowsVista = False) Then
        IsWindows2008 = CBool((m_osVer.dwVerMajor = 6))
    End If
    
End Property

Public Property Get IsWindows7() As Boolean

    ' ...
    
End Property

Public Property Get IsWin2000Plus() As Boolean
    
    If (IsWindowsNT) Then
        IsWin2000Plus = CBool(m_osVer.dwVerMajor >= 5)
    End If
    
End Property
