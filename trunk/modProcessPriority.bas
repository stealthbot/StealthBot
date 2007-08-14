Attribute VB_Name = "modProcessPriority"
Option Explicit

' Win32 API declarations
Private Declare Function _
   GetWindowThreadProcessId Lib "user32" ( _
   ByVal hWnd As Long, lpdwProcessId As Long) _
   As Long
Private Declare Function OpenProcess Lib _
   "kernel32" (ByVal dwDesiredAccess As Long, _
   ByVal bInheritHandle As Long, _
   ByVal dwProcessID As Long) As Long
Private Declare Function SetPriorityClass Lib _
   "kernel32" (ByVal hProcess As Long, _
   ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib _
   "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function CloseHandle Lib _
   "kernel32" (ByVal hObject As Long) As Long

' Used by the OpenProcess API call
Private Const PROCESS_QUERY_INFORMATION _
   As Long = &H400
Private Const PROCESS_SET_INFORMATION _
   As Long = &H200

' Used by SetPriorityClass
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Public Enum ProcessPriorities
   ppIdle = IDLE_PRIORITY_CLASS
   ppNormal = NORMAL_PRIORITY_CLASS
   ppHigh = HIGH_PRIORITY_CLASS
   ppRealtime = REALTIME_PRIORITY_CLASS
End Enum

Public Function SetProcessPriority(Optional ByVal ProcessID As Long, Optional ByVal hWnd As Long, _
                    Optional ByVal Priority As ProcessPriorities = NORMAL_PRIORITY_CLASS) As Long
                    
   Dim hProc As Long
   Const fdwAccess1 As Long = _
      PROCESS_QUERY_INFORMATION Or _
      PROCESS_SET_INFORMATION
   Const fdwAccess2 As Long = _
      PROCESS_QUERY_INFORMATION

   If ProcessID = 0 Then
      Call GetWindowThreadProcessId(hWnd, _
         ProcessID)
   End If

   hProc = OpenProcess(fdwAccess1, 0&, ProcessID)
   If hProc Then
      Call SetPriorityClass(hProc, Priority)
   Else
      hProc = OpenProcess(fdwAccess2, 0&, _
         ProcessID)
   End If

   SetProcessPriority = GetPriorityClass(hProc)
   Call CloseHandle(hProc)
End Function
