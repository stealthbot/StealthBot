Attribute VB_Name = "modCRC32Checksum"
Option Explicit

'Modified from code given to me by David Fritts (sneakcharm@yahoo.com)
Public Function ValidateExecutable() As Boolean
    On Error GoTo ValidateExecutable_Error
    
    Dim CRC32          As clsCRC32
    Dim strFilePath    As String
    Dim intFreeFile    As Integer
    Dim strBuffer      As String
    Dim strFileCRC     As String * 8
    Dim strComputedCRC As String * 8
    Dim lngComputedCRC As Long

    Set CRC32 = New clsCRC32
    
    strFilePath = App.Path & "/" & App.EXEName & ".exe"
    
    'Generate a CRC for ourselves
    intFreeFile = FreeFile
    
    'read the sections you want to protect
    Open strFilePath For Binary Access Read As #intFreeFile
        strBuffer = String$(LOF(intFreeFile) - 8, vbNullChar)
        
        Get #intFreeFile, 1, strBuffer
    Close #intFreeFile
    
    'Compute the new CRC
    lngComputedCRC = CRC32.GenerateCRC32(strBuffer)
    strComputedCRC = Hex$(lngComputedCRC)
    
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
    
    Set CRC32 = Nothing

ValidateExecutable_Exit:
    Exit Function

ValidateExecutable_Error:
    ValidateExecutable = True
    
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure " & _
        "ValidateExecutable of Module modCRC32Checksum"
        
    Resume ValidateExecutable_Exit
End Function
