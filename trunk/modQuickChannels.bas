Attribute VB_Name = "modQuickChannels"
'modQuickChannels - Project StealthBot
'   by Andy
'   June 2006

Public Sub LoadQuickChannels()
    Dim i As Integer
    Dim s As String
    Dim b As Boolean
    
    b = (Len(Dir$(GetFilePath("quickchannels.ini"))) > 0)
    
    For i = 0 To 8
        If b Then
            s = GetQC(i)
        Else
            s = GetDefaultQC(i)
        End If
        
        QC(i) = s
        frmChat.mnuQC(i).Caption = s
    Next i
    
    DoQCMenu
End Sub

Public Function GetQC(ByVal Index As Integer) As String
    GetQC = ReadINI("QuickChannels", Index, "quickchannels.ini")
End Function

Public Function GetDefaultQC(ByVal Index As Integer) As String
    Select Case Index
        Case (0): GetDefaultQC = "Clan SBS"
        Case (1): GetDefaultQC = "Clan DKe"
        Case (2): GetDefaultQC = "Clan TDA"
        Case (3): GetDefaultQC = "Clan BNU"
        Case (4): GetDefaultQC = "Op W@R"
    End Select
End Function

Public Sub SaveQCs()
    Dim i As Integer
    
    For i = 0 To 8
        WriteINI "QuickChannels", i, QC(i), "quickchannels.ini"
    Next i
End Sub

Public Sub DoQCMenu()
    For i = 0 To 8
        If LenB(QC(i)) > 0 Then
            frmChat.mnuQC(i).Visible = True
        Else
            frmChat.mnuQC(i).Visible = False
        End If
    Next i
End Sub
