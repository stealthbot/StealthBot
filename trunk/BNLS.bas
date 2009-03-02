Attribute VB_Name = "modBNLSCode"
'StealthBot BNLS-related Methods Module
'   http://www.valhallalegends.com/yoni/BNLSProtocolSpec.txt
' Updated 11/7/06 moved GetBNLSProductID() into this module as a public method
Option Explicit

'Public Classes
Public Packet As New PacketBuffer
Public NLogin As New cBNLS
Public ds     As New DataStorage

Public Function BNLSChecksum(ByVal Password As String, ByVal ServerCode As Long) As Long

    Dim clsCRC32 As clsCRC32
    
    Set clsCRC32 = New clsCRC32

    BNLSChecksum = _
        clsCRC32.CRC32(Password & Right("0000000" & hex(ServerCode), 8))
        
    Set clsCRC32 = Nothing
    
End Function

'Needed For BNLS_VERSIONCHECK & BNLS_REQUESTVERSIONBYTE
Public Function GetBNLSProductID(ByVal sProdID As String) As Long

    Select Case (UCase$(sProdID))
        Case "RATS": GetBNLSProductID = &H1
        Case "PXES": GetBNLSProductID = &H2
        Case "NB2W": GetBNLSProductID = &H3
        Case "VD2D": GetBNLSProductID = &H4
        Case "PX2D": GetBNLSProductID = &H5
        Case "RTSJ": GetBNLSProductID = &H6
        Case "3RAW": GetBNLSProductID = &H7
        Case "PX3W": GetBNLSProductID = &H8
        Case Else:   GetBNLSProductID = &H0
    End Select
    
End Function
