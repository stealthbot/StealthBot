Attribute VB_Name = "modQuickChannels"
'modQuickChannels - Project StealthBot
'   by Andy
'   June 2006
Option Explicit

Public Sub LoadQuickChannels()
    Dim UseDefaultQC As Boolean
    Dim i As Integer
    Dim QCCollection As Collection
    Dim LineText As String
    
    UseDefaultQC = True
    
    ' "upgrade" a QC.ini to a QC.txt, deletes QC.ini
    If (LenB(Dir$(GetFilePath("QuickChannels.ini"))) > 0) Then
    
        UseDefaultQC = False
        
        Call UpgradeQuickChannelsToList
        
    End If
    
    ' read QC.txt
    If LenB(Dir$(GetFilePath(FILE_QUICK_CHANNELS))) > 0 Then
    
        UseDefaultQC = False
        
        Set QCCollection = ListFileLoad(GetFilePath(FILE_QUICK_CHANNELS), 9)
        
        For i = 1 To QCCollection.Count
            
            QC(i) = Trim$(QCCollection.Item(i))
            
        Next i
        
        Set QCCollection = Nothing
        
    End If
    
    ' no QC files, use defaults and save it
    If UseDefaultQC Then
    
        i = 0
        Do
        
            LineText = GetDefaultQC(i)
            
            If LenB(LineText) > 0 Then
        
                QC(i + 1) = LineText
            
                i = i + 1
            
            End If
            
        Loop Until LenB(LineText) = 0
        
        'Call SaveQuickChannels
        
    End If

    PrepareQuickChannelMenu
End Sub

Public Function GetDefaultQC(ByVal Index As Integer) As String

    ' no default QC...
    GetDefaultQC = vbNullString
    'Select Case Index
        'Case 0: GetDefaultQC = "Clan BoT"
        'Case 1: GetDefaultQC = "Public Chat 1"
        'Case 2: GetDefaultQC = "Clan TDA"
        'Case 3: GetDefaultQC = "Clan BNU"
        'Case 4: GetDefaultQC = "Op W@R"
    'End Select

End Function

Public Sub UpgradeQuickChannelsToList()
    Dim sFileINI As String
    Dim i As Integer
    
    sFileINI = GetFilePath("QuickChannels.ini")
    
    ' old QC.ini bounds
    For i = 0 To 8
        QC(i + 1) = Trim$(ReadINI("QuickChannels", CStr(i), sFileINI))
    Next i
    
    Kill sFileINI
    
    Call SaveQuickChannels
End Sub

Public Sub SaveQuickChannels()
    Dim i As Integer
    Dim QCCollection As New Collection
    
    For i = LBound(QC) To UBound(QC)
        If LenB(QC(i)) = 0 Then
            QC(i) = " "
        End If
        QCCollection.Add QC(i)
    Next i
    
    Call ListFileSave(GetFilePath(FILE_QUICK_CHANNELS), QCCollection)
    
    Set QCCollection = Nothing
End Sub

Public Sub PrepareQuickChannelMenu()
    Dim i As Integer
    Dim Caption As String
    Dim ShownAddQC As Boolean
    Dim FoundThisChannel As Boolean
    
    ShownAddQC = False
    FoundThisChannel = False
    
    ' bounds of mnuCustomChannels
    For i = 0 To 8
        Caption = Trim$(QC(i + 1))
        
        With frmChat.mnuCustomChannels(i)
            If LenB(Caption) > 0 Then
                If StrComp(Caption, g_Channel.Name, vbTextCompare) = 0 Then
                    FoundThisChannel = True
                End If
                
                .Visible = True
                .Caption = MakeChannelMenuItemSafe(Caption)
            Else
                If Not ShownAddQC And LenB(g_Channel.Name) > 0 Then
                    frmChat.mnuCustomChannelAdd.Visible = True
                    frmChat.mnuCustomChannelAdd.Caption = StringFormat("&Add {0}{1}{0} as F{2}", Chr$(34), MakeChannelMenuItemSafe(g_Channel.Name, True), CStr(i + 1))
                    
                    ShownAddQC = True
                End If
                
                .Visible = False
                .Caption = vbNullString
            End If
        End With
    Next i
    
    If FoundThisChannel Or Not ShownAddQC Then
        frmChat.mnuCustomChannelAdd.Visible = False
    End If
End Sub

Public Sub PrepareHomeChannelMenu()
    Dim ShowHome As Boolean
    Dim ShowLast As Boolean

    ShowHome = (LenB(Config.HomeChannel) > 0 And StrComp(Config.HomeChannel, g_Channel.Name, vbTextCompare) <> 0)
    With frmChat.mnuHomeChannel
        .Caption = MakeChannelMenuItemSafe(Config.HomeChannel, True) & " (&Home Channel)"
        .Visible = ShowHome
    End With

    ShowLast = (LenB(BotVars.LastChannel) > 0 And StrComp(BotVars.LastChannel, g_Channel.Name, vbTextCompare) <> 0)
    With frmChat.mnuLastChannel
        .Caption = MakeChannelMenuItemSafe(BotVars.LastChannel, True) & " (&Previous Channel)"
        .Visible = ShowLast
    End With

    frmChat.mnuQCDash.Visible = (ShowHome Or ShowLast)
End Sub

Public Sub PreparePublicChannelMenu()
    Dim i As Integer
    Dim AnyVisible As Boolean

    ' unload all menu items
    For i = 1 To frmChat.mnuPublicChannels.Count - 1
        Call Unload(frmChat.mnuPublicChannels(i))
    Next i

    ' hide 0th menu item
    With frmChat.mnuPublicChannels(0)
        .Caption = vbNullString
        .Visible = False
    End With

    AnyVisible = False

    If Not BotVars.PublicChannels Is Nothing Then
        If BotVars.PublicChannels.Count > 0 Then
            For i = 1 To BotVars.PublicChannels.Count
                AnyVisible = True

                If (i > 1) Then
                    ' load a new one
                    Call Load(frmChat.mnuPublicChannels(i - 1))
                End If

                With frmChat.mnuPublicChannels(i - 1)
                    .Caption = MakeChannelMenuItemSafe(BotVars.PublicChannels.Item(i))
                    .Visible = True
                End With
            Next i
        End If
    End If

    ' dash and header visibility
    frmChat.mnuPCDash.Visible = AnyVisible
    frmChat.mnuPCHeader.Visible = AnyVisible
End Sub

Private Function MakeChannelMenuItemSafe(ByVal sChannel As String, Optional ByVal CanBeDash As Boolean = False) As String
    sChannel = Replace(sChannel, "&", "&&", , , vbBinaryCompare)

    If Not CanBeDash And StrComp(sChannel, "-", vbBinaryCompare) = 0 Then
        sChannel = "&-"
    End If

    MakeChannelMenuItemSafe = sChannel
End Function

