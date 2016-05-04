Option Strict Off
Option Explicit On
Friend Class clsBNCSRecvBuffer

    Private bData() As Byte
    Private DataSize As Integer

    Public Sub AddData(ByVal Data As String)
        ReDim Preserve bData(DataSize + Data.Length - 1)
    End Sub

    Public Function GetBuffer() As Byte()
        GetBuffer = bData
    End Function

    Public Function FullPacket() As Boolean
        Dim lngPacketLen, L As Integer
        Dim drop() As Byte
        Dim i As Integer

        FullPacket = False

        If DataSize > 0 Then
            L = Array.IndexOf(bData, &HFF)

            If L = 0 Then
                lngPacketLen = BitConverter.ToInt16(bData, 2)

                If (lngPacketLen = 0) Then
                    Exit Function
                End If

                If (DataSize >= lngPacketLen) Then
                    If lngPacketLen < 10000 Then
                        FullPacket = True
                    Else
                        frmChat.AddChat(RTBColors.ErrorMessageText, "Error: Packet Length of unusually high Length detected! Packet " & "dropped. Buffer content at this time: " & vbCrLf & DebugOutput(bData))

                        Call ClearBuffer()
                    End If
                End If
            Else
                frmChat.AddChat(RTBColors.ErrorMessageText, "Error: The front of the buffer is not a valid packet!")

                If MDebug("showdrops") Then
                    frmChat.AddChat(RTBColors.ErrorMessageText, "Error: The front of the buffer is not " & "a valid packet!")
                    frmChat.AddChat(RTBColors.ErrorMessageText, "The following data is being purged:")

                    If L > 0 Then
                        ReDim drop(L - 1)
                        Buffer.BlockCopy(bData, 0, drop, 0, L)

                        frmChat.AddChat(Space(1) & DebugOutput(drop))
                    Else
                        frmChat.AddChat(Space(1) & DebugOutput(bData))
                    End If
                End If

                If L > 0 Then
                    ' Trim off the extra data
                    For i = L To UBound(bData)
                        bData(i - L) = bData(i)
                    Next

                    DataSize = (DataSize - L)
                    ReDim Preserve bData(DataSize - 1)
                Else
                    ClearBuffer()
                End If
            End If
        End If
    End Function

    Public Function GetPacket() As Byte()
        Dim lngPacketLen As Integer
        Dim i As Integer

        lngPacketLen = BitConverter.ToInt16(bData, 2)
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
        Dim i As Short

        For i = 1 To Len(Data)
            ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
        Next i
    End Function

End Class