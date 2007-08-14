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
'by David Fritts (aka sneakcharm) (sneakcharm@yahoo.com)
Option Explicit

Private Const POLYNOMIAL As Long = &HEDB88320
Private Const INITIALIZE As Long = &HFFFFFFFF

Private mobjBitwiseOperator As clsBitwiseOperator
Private mlngCRC32Table(255) As Long

Private Sub Class_Initialize()
    Dim lngBytes As Long, lngBits As Long
    Dim lngCRC32 As Long, lngTemporaryCRC32 As Long

    Set mobjBitwiseOperator = New clsBitwiseOperator
    
    For lngBytes = 0 To 255

        lngCRC32 = lngBytes

        For lngBits = 0 To 7
            lngTemporaryCRC32 = mobjBitwiseOperator.ShiftRight(lngCRC32, 1)

            If (lngCRC32 And &H1) Then
                lngCRC32 = lngTemporaryCRC32 Xor POLYNOMIAL
            Else
                lngCRC32 = lngTemporaryCRC32
            End If
        Next lngBits

        mlngCRC32Table(lngBytes) = lngCRC32
    Next lngBytes
End Sub

Public Function GenerateCRC32(ByVal strBuffer As String) As Long
    Dim lngBytes As Long, bytCurrentByte As Byte
    Dim lngActualValue As Long
    Dim lngTableValue As Long, lngIndex As Long
    Dim lngInitializeValue As Long
    
    'strBuffer = UnicodeToAscii(strBuffer)

    lngInitializeValue = INITIALIZE
    
    For lngBytes = 1 To Len(strBuffer)
        
        bytCurrentByte = Asc(Mid$(strBuffer, lngBytes, 1))
        
        lngActualValue = mobjBitwiseOperator.ShiftRight(lngInitializeValue, 8)
        
        lngIndex = lngInitializeValue And &HFF
        lngIndex = lngIndex Xor bytCurrentByte
        lngTableValue = mlngCRC32Table(lngIndex)
        
        lngInitializeValue = lngActualValue Xor lngTableValue
    Next lngBytes
    
    GenerateCRC32 = lngInitializeValue Xor &HFFFFFFFF
End Function

Private Sub Class_Terminate()
    Set mobjBitwiseOperator = Nothing
End Sub

