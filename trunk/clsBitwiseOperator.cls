VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitwiseOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' From David Fritts (CodeId=45885)
' ASM corrected by Robert Rayment & David Fritts
' Class recast by Ulli

Private Type tPD
    hMem                As Long
    PtrToOldCode        As Long
End Type

Private ProcDetails()   As tPD

Private VTIndex         As Long
Private code            As Byte
Private CodeSize        As Long
Private PtrToNewCode    As Long
Private PtrToMyself     As Long
Private i               As Long

Private Sub Class_Initialize()
    VTIndex = -1    'initialize index into Virtual Table
    ShiftLeft 0, 0  'this sets up m/c code and modifies the VT
    ShiftRight 0, 0  'this sets up m/c code and modifies the VT
    RotateLeft 0, 0  'this sets up m/c code and modifies the VT
    RotateRight 0, 0  'this sets up m/c code and modifies the VT
End Sub

Public Function ShiftLeft(ByVal lngValue As Long, ByVal lngShift As Long) As Long
    'this is in fact only called once during class initialize
    'subsequent calls are diverted (via the VT) to the m/c code

    DivertTo "8B442408 8B4C240C 8B542410 D3E0 8902 31C0 C21000"
End Function

Public Function ShiftRight(ByVal lngValue As Long, ByVal lngShift As Long) As Long

    'this is in fact only called once during class initialize
    'subsequent calls are diverted (via the VT) to the m/c code

    DivertTo "8B442408 8B4C240C 8B542410 D3E8 8902 31C0 C21000"
End Function

Public Function RotateLeft(ByVal lngValue As Long, ByVal lngRotate As Long) As Long
    'this is in fact only called once during class initialize
    'subsequent calls are diverted (via the VT) to the m/c code

    DivertTo "8B442408 8B4C240C 8B542410 D3C0 8902 31C0 C21000"
End Function

Public Function RotateRight(ByVal lngValue As Long, ByVal lngRotate As Long) As Long
    'this is in fact only called once during class initialize
    'subsequent calls are diverted (via the VT) to the m/c code

    DivertTo "8B442408 8B4C240C 8B542410 D3C8 8902 31C0 C21000"
End Function

Private Sub DivertTo(ByVal HexCode As String)
    VTIndex = VTIndex + 1 'inc index into VT
    
    ReDim Preserve ProcDetails(0 To VTIndex) 'adjust array size
    
    HexCode = Replace$(HexCode, " ", "") 'remove spaces from hex code
    
    CodeSize = Len(HexCode) / 2 'length of the resulting binary code (2 hex chars per byte of code)

    With ProcDetails(VTIndex)
        .hMem = GlobalAlloc(0, CodeSize) 'get memory for m/c code and save handle
        
        PtrToNewCode = GlobalLock(.hMem) 'get far pointer to allocated memory

        For i = 0 To CodeSize - 1
            code = Val("&H" & Mid$(HexCode, i + i + 1, 2)) 'convert hex to binary m/c code
            
            Call CopyMemory(ByVal PtrToNewCode + i, code, 1)  'store it in allocated memory
        Next i

        .PtrToOldCode = VirtualTableEntry 'save old VT entry; VTIndex determines which entry
        
        VirtualTableEntry = PtrToNewCode 'overwrite VT entry; VTIndex determines which entry
        
        GlobalUnlock .hMem 'unlock memory
    End With 'PROCDETAILS(VTINDEX)
End Sub

Private Property Let VirtualTableEntry(ByVal FarPointer As Long)
    Call CopyMemory(PtrToMyself, ByVal ObjPtr(Me), 4)  'get pointer to object (Me)
    Call CopyMemory(ByVal PtrToMyself + &H1C + VTIndex * 4, FarPointer, 4)  'put VT entry
End Property

Private Property Get VirtualTableEntry() As Long
    Call CopyMemory(PtrToMyself, ByVal ObjPtr(Me), 4)  'get pointer to object (Me)
    Call CopyMemory(VirtualTableEntry, ByVal PtrToMyself + &H1C + VTIndex * 4, 4)  'get VT entry
End Property

Private Sub Class_Terminate()
    For VTIndex = VTIndex To 0 Step -1 'VTIndex still points to the last VT entry overwritten
        With ProcDetails(VTIndex)
            VirtualTableEntry = .PtrToOldCode 'restore VT entry; VTIndex determines which entry
            GlobalFree .hMem 'release memory used for m/c code
        End With 'PROCDETAILS(VTINDEX)
    Next VTIndex
End Sub

