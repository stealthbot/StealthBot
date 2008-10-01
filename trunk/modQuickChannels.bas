Attribute VB_Name = "modQuickChannels"
'modQuickChannels - Project StealthBot
'   by Andy
'   June 2006

Public Sub LoadQuickChannels()
    Dim I As Integer
    Dim s As String
    Dim B As Boolean
    
    B = (Len(Dir$(GetFilePath("quickchannels.ini"))) > 0)

    For I = 0 To 8
        If B Then
            s = GetQC(I)
        Else
            s = GetDefaultQC(I)
        End If
        
        If (s = vbNullString) Then
            Exit For
        End If

        QC(I) = s
        'frmChat.mnuQC(I).Caption = s
        
        'If (frmChat.mnuCustomChannels(0).Caption <> vbNullString) Then
        '    Call Load(frmChat.mnuCustomChannels(frmChat.mnuCustomChannels.Count))
        'End If
        
        frmChat.mnuCustomChannels(I).Caption = s
    Next I

    DoQCMenu
End Sub

Public Function GetQC(ByVal index As Integer) As String
    GetQC = ReadINI("QuickChannels", index, "quickchannels.ini")
End Function

Public Function GetDefaultQC(ByVal index As Integer) As String
    Select Case index
        Case 0: GetDefaultQC = "Clan SBS"
        Case 1: GetDefaultQC = "Clan BoT"
        Case 2: GetDefaultQC = "Clan DKe"
        Case 3: GetDefaultQC = "Clan TDA"
        Case 4: GetDefaultQC = "Clan BNU"
        Case 5: GetDefaultQC = "Op W@R"
    End Select
End Function

Public Sub SaveQCs()
    Dim I As Integer
    
    For I = 0 To 8
        WriteINI "QuickChannels", I, QC(I), "quickchannels.ini"
    Next I
End Sub

Public Sub DoQCMenu()
    For I = 0 To 8
        If LenB(QC(I)) > 0 Then
            'frmChat.mnuQC(I).Visible = True
            frmChat.mnuCustomChannels(I).Visible = True
        Else
            'frmChat.mnuQC(I).Visible = False
            frmChat.mnuCustomChannels(I).Visible = False
        End If
    Next I
End Sub
