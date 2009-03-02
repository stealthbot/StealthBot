VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//This class is ugly
'//I orginaly wrote it for some fun, so ignore the retardation
'//Ive only added what ive needed/used it for so far, so If you want
'//other stuff, you need to add it.
Private m_ptrMe         As Long 'pointer of this class
Private m_ptrFunc       As Long 'function pointer of old Execute
Private m_Mem           As Long
Private m_Pos           As Long

Private Sub Class_Initialize()
    Call CopyMemory(m_ptrMe, ByVal ObjPtr(Me), 4)
    Call CopyMemory(m_ptrFunc, ByVal m_ptrMe + 28, 4)
    m_Mem = malloc(256) 'don't make to bigger functions with out making the buffer bigger
    Call CopyMemory(ByVal m_ptrMe + 28, m_Mem, 4)
End Sub
Private Sub Class_Terminate()
    Call CopyMemory(m_ptrMe, ByVal ObjPtr(Me), 4)
    Call CopyMemory(ByVal m_ptrMe + 28, m_ptrFunc, 4)
    Call free(m_Mem)
End Sub



Public Function Execute() As Long
    Execute = 1
End Function

Public Sub retn(Optional ByVal Num As Integer = 0)
    If (Num = 0) Then
        Call i08(&HC3)
    Else
        Call i08(&HC2)
        Call i16(Num)
    End If
End Sub
Public Sub pop__eax(): Call i08(&H58): End Sub
Public Sub pop__ecx(): Call i08(&H59): End Sub
Public Sub pop__edx(): Call i08(&H5A): End Sub
Public Sub pop__edp(): Call i08(&H55): End Sub
Public Sub pop__ebp(): Call i08(&H5D): End Sub
Public Sub pop__esi(): Call i08(&H5E): End Sub
Public Sub pop__edi(): Call i08(&H5F): End Sub
Public Sub pop__ebx(): Call i08(&H5B): End Sub

Public Sub push_eax(): Call i08(&H50): End Sub
Public Sub push_ecx(): Call i08(&H51): End Sub
Public Sub push_edx(): Call i08(&H52): End Sub
Public Sub push_esi(): Call i08(&H56): End Sub
Public Sub push_edp(): Call i08(&H55): End Sub
Public Sub push_ebx(): Call i08(&H53): End Sub
Public Sub push_v08(ByVal Num As Byte): Call i08(&H6A): Call i08(Num): End Sub
Public Sub push_v32(ByVal Num As Long): Call i08(&H68): Call i32(Num): End Sub

Public Sub jmp__eax(): Call i16(&HE0FF): End Sub
Public Sub jmp__ecx(): Call i16(&HE1FF): End Sub
Public Sub jmp__edx(): Call i16(&HE2FF): End Sub

Public Sub call_ptr(ByVal Ptr As Long):     Call i16(&H15FF):   Call i32(Ptr): End Sub
Public Sub call_eax(): Call i16(&HD0FF): End Sub
Public Sub call_ecx(): Call i16(&HD1FF): End Sub
Public Sub call_edx(): Call i16(&HD2FF): End Sub

Public Sub xor__eax_eax(): Call i16(&HC033): End Sub
Public Sub xor__ecx_ecx(): Call i16(&HC933): End Sub
Public Sub xor__edx_edx(): Call i16(&HD233): End Sub
Public Sub xor__esi_esi(): Call i16(&HF633): End Sub

Public Sub not__eax(): Call i16(&HD0F7): End Sub
Public Sub not__ecx(): Call i16(&HD1F7): End Sub
Public Sub not__edx(): Call i16(&HD2F7): End Sub

Public Sub shr__eax_v08(ByVal Num As Byte): Call i16(&HE8C1): Call i08(Num): End Sub
Public Sub shl__edx_v08(ByVal Num As Byte): Call i16(&HE2C1): Call i08(Num): End Sub

Public Sub ror__eax_v08(ByVal Num As Byte)
    If (Num = 1) Then
        Call i16(&HC8D1)
    Else
        Call i16(&HC8C1)
        Call i08(Num)
    End If
End Sub

Public Sub rol__eax_v08(ByVal Num As Byte)
    If (Num = 1) Then
        Call i16(&HC0D1)
    Else
        Call i16(&HC0C1)
        Call i08(Num)
    End If
End Sub

Public Sub inc__eax(): Call i08(&H40): End Sub
Public Sub inc__ecx(): Call i08(&H41): End Sub
Public Sub inc__edx(): Call i08(&H42): End Sub

Public Sub dec__eax(): Call i08(&H48): End Sub
Public Sub dec__ecx(): Call i08(&H49): End Sub
Public Sub dec__edx(): Call i08(&H4A): End Sub

Public Sub mov__eax_ptr(ByVal Ptr As Long): Call i08(&HA1):     Call i32(Ptr): End Sub
Public Sub mov__ecx_ptr(ByVal Ptr As Long): Call i16(&HD8B):    Call i32(Ptr): End Sub
Public Sub mov__edx_ptr(ByVal Ptr As Long): Call i16(&H158B):   Call i32(Ptr): End Sub
Public Sub mov__eax_v32(ByVal Num As Long): Call i08(&HB8):     Call i32(Num): End Sub
Public Sub mov__ecx_v32(ByVal Num As Long): Call i08(&HB9):     Call i32(Num): End Sub
Public Sub mov__edx_v32(ByVal Num As Long): Call i08(&HBA):     Call i32(Num): End Sub

Public Sub mov__ptr_eax(ByVal Ptr As Long): Call i08(&HA3):     Call i32(Ptr): End Sub
Public Sub mov__ptr_ecx(ByVal Ptr As Long): Call i16(&HD89):    Call i32(Ptr): End Sub
Public Sub mov__ptr_edx(ByVal Ptr As Long): Call i16(&H1589):   Call i32(Ptr): End Sub

Public Sub mov__ptr_v32(ByRef Ptr As Long, ByVal Num As Long)
    Call i16(&H5C7)
    Call i32(Ptr)
    Call i32(Num)
End Sub

Public Function GetCode() As String
    If (m_Pos < 1) Then Exit Function
    GetCode = Space(m_Pos)
    Call CopyMemory(ByVal GetCode, ByVal m_Mem, m_Pos)
End Function
Private Sub i08(ByRef bValue As Byte): Call CopyMemory(ByVal m_Mem + m_Pos, bValue, 1): m_Pos = m_Pos + 1: End Sub
Private Sub i16(ByRef intValue As Integer): Call CopyMemory(ByVal m_Mem + m_Pos, intValue, 2): m_Pos = m_Pos + 2: End Sub
Private Sub i32(ByRef lngValue As Long): Call CopyMemory(ByVal m_Mem + m_Pos, lngValue, 4): m_Pos = m_Pos + 4: End Sub


