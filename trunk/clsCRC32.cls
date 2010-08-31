VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CRC32_POLYNOMIAL As Long = &HEDB88320

Private CRC32Table(0 To 255) As Long

'Public Functions
Private Sub InitCRC32()
    Static CRC32Initialized As Boolean
    
    Dim i      As Long
    Dim j      As Long
    Dim K      As Long
    Dim XorVal As Long
    
    If (CRC32Initialized) Then
        Exit Sub
    End If
    
    CRC32Initialized = True
    
    For i = 0 To 255
        K = i
        
        For j = 1 To 8
            If K And 1 Then XorVal = CRC32_POLYNOMIAL Else XorVal = 0
            If K < 0 Then K = ((K And &H7FFFFFFF) \ 2) Or &H40000000 Else K = K \ 2
            K = K Xor XorVal
        Next
        
        CRC32Table(i) = K
    Next    
End Sub

Public Function CRC32(ByVal Data As String) As Long
    Dim i As Long
    Dim j As Long
    
    Call InitCRC32
    
    CRC32 = &HFFFFFFFF
    
    For i = 1 To Len(Data)
        j = CByte(Asc(Mid(Data, i, 1))) Xor (CRC32 And &HFF&)
        
        If (CRC32 < 0) Then
            CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100&) Or &H800000
        Else
            CRC32 = CRC32 \ &H100&
        End If
            
        CRC32 = (CRC32 Xor CRC32Table(j))
    Next
    
    CRC32 = (Not (CRC32))
End Function

Public Function GetFileCRC32(ByVal filePath As String) As Long
    Dim str As String  
    Dim tmp As String 
    Dim f   As Integer 
    
    f = FreeFile
    
    Open filePath For Input As #f
        Do While (EOF(f) = False)
            Line Input #f, tmp
            
            str = str & tmp
        Loop
    Close #f
    
    GetFileCRC32 = CRC32(str)
End Function

'Modified from code given to me by David Fritts (sneakcharm@yahoo.com)
Public Function ValidateExecutable() As Boolean
    On Error GoTo ValidateExecutable_Error
    
    'Dim CRC32          As clsCRC32
    Dim strFilePath    As String
    Dim intFreeFile    As Integer
    Dim strBuffer      As String
    Dim strFileCRC     As String * 8
    Dim strComputedCRC As String * 8
    Dim lngComputedCRC As Long

    'Set CRC32 = New clsCRC32
    
    strFilePath = App.Path & "/" & App.EXEName & ".exe"
    
    'Generate a CRC for ourselves
    intFreeFile = FreeFile
    
    'read the sections you want to protect
    Open strFilePath For Binary Access Read As #intFreeFile
        strBuffer = String$(LOF(intFreeFile) - 8, vbNullChar)
        
        Get #intFreeFile, 1, strBuffer
    Close #intFreeFile
    
    'Compute the new CRC
    lngComputedCRC = CRC32(strBuffer)
    strComputedCRC = ZeroOffset(lngComputedCRC, 8)
    
    'Read a CRC from ourselves
    intFreeFile = FreeFile
    
    Open strFilePath For Binary As #intFreeFile
        Get #intFreeFile, FileLen(strFilePath) - 7, strFileCRC
    Close #intFreeFile
    
    If (StrComp(strComputedCRC, strFileCRC, vbBinaryCompare) = 0) Then
        ValidateExecutable = True
    Else
        ValidateExecutable = False
    End If
    
    'Set CRC32 = Nothing

ValidateExecutable_Exit:
    Exit Function

ValidateExecutable_Error:
    ValidateExecutable = True
    
    Debug.Print "Error " & Err.Number & " (" & Err.description & ") in procedure " & _
        "ValidateExecutable of Module modCRC32Checksum"
        
    Resume ValidateExecutable_Exit    
End Function
