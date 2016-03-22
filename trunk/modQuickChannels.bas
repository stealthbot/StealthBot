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
    If LenB(Dir$(GetFilePath("QuickChannels.txt"))) > 0 Then
    
        UseDefaultQC = False
        
        Set QCCollection = ListFileLoad(GetFilePath("QuickChannels.txt"), 9)
        
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

    DoQuickChannelMenu
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
    
    Call ListFileSave(GetFilePath("QuickChannels.txt"), QCCollection)
    
    Set QCCollection = Nothing
End Sub

Public Sub DoQuickChannelMenu()
    Dim i As Integer
    Dim Caption As String
    Dim ShownAddQC As Boolean
    Dim FoundThisChannel As Boolean
    
    ShownAddQC = False
    FoundThisChannel = False
    
    ' bounds of mnuCustomChannels
    For i = 0 To 8
        Caption = Trim$(QC(i + 1))
        
        If LenB(Caption) > 0 Then
            If StrComp(Caption, g_Channel.Name, vbTextCompare) = 0 Then
                FoundThisChannel = True
            End If
            
            frmChat.mnuCustomChannels(i).Visible = True
            
            Caption = Replace(Caption, "&", "&&", , , vbBinaryCompare)
            If StrComp(Caption, "-", vbBinaryCompare) = 0 Then
                Caption = "&-"
            End If
            
            frmChat.mnuCustomChannels(i).Caption = Caption
        Else
            frmChat.mnuCustomChannels(i).Visible = False
            
            frmChat.mnuCustomChannels(i).Caption = vbNullString
            
            If Not ShownAddQC And LenB(g_Channel.Name) > 0 Then
                frmChat.mnuCustomChannelAdd.Visible = True
                frmChat.mnuCustomChannelAdd.Caption = StringFormat("&Add {0}{1}{0} as F{2}", Chr$(34), g_Channel.Name, CStr(i + 1))
                
                ShownAddQC = True
            End If
        End If
    Next i
    
    If FoundThisChannel Or Not ShownAddQC Then
        frmChat.mnuCustomChannelAdd.Visible = False
    End If
End Sub
