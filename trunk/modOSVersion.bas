Attribute VB_Name = "modOSVersion"
'modOSVersion.bas
' project StealthBot
' October 2006 from code at:
'  http://vbnet.mvps.org/index.html?code/helpers/iswinversion.htm
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
  
Private Const VER_PLATFORM_WIN32_NT As Long = 2

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

Public Function IsWin2000Plus() As Boolean
  'returns True if running Windows 2000 or later
  'Updated 11/28/06 to cache responses
  'Updated 04/05/08 to change Dims to Statics (joetheodd)
  
    Static CachedResponseAvailable As Boolean
    Static CachedResponse As Boolean
  
    Dim osv As OSVERSIONINFO

    If Not CachedResponseAvailable Then
        osv.OSVSize = Len(osv)
    
        If GetVersionEx(osv) = 1 Then
            CachedResponse = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                          (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
            CachedResponseAvailable = True
        End If
    End If
    
    IsWin2000Plus = CachedResponse
End Function
