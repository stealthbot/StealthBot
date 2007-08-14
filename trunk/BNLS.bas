Attribute VB_Name = "modBNLSCode"
'StealthBot BNLS-related Methods Module
'   http://www.valhallalegends.com/yoni/BNLSProtocolSpec.txt
' Updated 11/7/06 moved GetBNLSProductID() into this module as a public method
Option Explicit

'Public Variables
Private CRC32Table(0 To 255) As Long

'Public Classes
Public Packet As New PacketBuffer
Public NLogin As New cBNLS
Public ds As New DataStorage

'Public Definitions

'Public Prototypes

'Public Constants
Private Const CRC32_POLYNOMIAL As Long = &HEDB88320

'Public Functions
Private Sub InitCRC32()
    Dim i As Long, j As Long, K As Long, XorVal As Long
    
    Static CRC32Initialized As Boolean
    If CRC32Initialized Then Exit Sub
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

Private Function CRC32(ByVal Data As String) As Long
    Dim i As Long, j As Long
    
    Call InitCRC32
    
    CRC32 = &HFFFFFFFF
    
    For i = 1 To Len(Data)
        j = CByte(Asc(Mid(Data, i, 1))) Xor (CRC32 And &HFF&)
        If CRC32 < 0 Then CRC32 = ((CRC32 And &H7FFFFFFF) \ &H100&) Or &H800000 Else CRC32 = CRC32 \ &H100&
        CRC32 = CRC32 Xor CRC32Table(j)
    Next
    
    CRC32 = Not CRC32
End Function

Public Function BNLSChecksum(ByVal Password As String, ByVal ServerCode As Long) As Long
    BNLSChecksum = CRC32(Password & Right("0000000" & Hex(ServerCode), 8))
End Function

Public Function GetBNLSProductID(ByVal sProdID As String) As Long 'Needed For BNLS_VERSIONCHECK & BNLS_REQUESTVERSIONBYTE
    Select Case UCase(sProdID)
        Case "RATS": GetBNLSProductID = &H1
        Case "PXES": GetBNLSProductID = &H2
        Case "NB2W": GetBNLSProductID = &H3
        Case "VD2D": GetBNLSProductID = &H4
        Case "PX2D": GetBNLSProductID = &H5
        Case "RTSJ": GetBNLSProductID = &H6
        Case "3RAW": GetBNLSProductID = &H7
        Case "PX3W": GetBNLSProductID = &H8
        Case Else: GetBNLSProductID = &H0
    End Select
End Function
