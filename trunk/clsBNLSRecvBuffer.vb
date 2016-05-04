Option Strict Off
Option Explicit On
Friend Class clsBNLSRecvBuffer

    Private bData() As Byte
    Private DataSize As Integer
	
    Public Sub AddData(ByRef Data() As Byte)
        ReDim Preserve bData(DataSize + Data.Length - 1)
    End Sub

    Public Function FullPacket() As Boolean
        Dim lngPacketLen As Integer

        FullPacket = False

        If (Len(DataSize) > 0) Then
            lngPacketLen = BitConverter.ToInt16(bData, 0)

            If (DataSize >= lngPacketLen) Then
                FullPacket = True
            End If
        End If
    End Function

    Public Function GetPacket() As Byte()
        Dim lngPacketLen As Integer
        Dim i As Integer

        lngPacketLen = BitConverter.ToInt16(bData, 0)
        Buffer.BlockCopy(bData, 0, GetPacket, 0, lngPacketLen)

        If lngPacketLen < DataSize Then
            For i = lngPacketLen To UBound(bData)
                bData(i - lngPacketLen) = bData(i)
            Next

            DataSize = (DataSize - lngPacketLen)
            ReDim Preserve bData(DataSize - 1)
        Else
            ClearBuffer()
        End If
    End Function

    Public Sub ClearBuffer()
        ReDim bData(0)
        bData(0) = 0

        DataSize = 0
    End Sub

    Private Function ToHex(ByRef Data As String) As String
        Dim I As Short

        For I = 1 To Len(Data)
            ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, I, 1))), 2)
        Next I
    End Function
End Class