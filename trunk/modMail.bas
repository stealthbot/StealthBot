Attribute VB_Name = "modMail"
'modMail - project StealthBot - authored 8/3/04 andy@stealthbot.net
Option Explicit

Private CurrentOpenFile As Integer
Private CurrentRecord   As Long

Public Sub AddMail(ByRef tsMsg As udtMail)
    Call OpenMailFile
    
    tsMsg.To = LCase(tsMsg.To)
    Put #CurrentOpenFile, CurrentRecord + 1, tsMsg
    
    Call CloseMailFile
End Sub

Public Function GetMailCount(ByVal sUser As String) As Long
    Dim mTemp As udtMail
    Dim i     As Long
    Dim Count As Long
    
    Call OpenMailFile
    
    sUser = LCase$(sUser)
    
    If (CurrentRecord > 0) Then
        For i = 1 To CurrentRecord
            Get #CurrentOpenFile, i, mTemp
            
            If (StrComp(sUser, RTrim(mTemp.To), vbTextCompare) = 0) Then
                Count = Count + 1
            End If
        Next i
        
        GetMailCount = Count
    Else
        GetMailCount = 0
    End If
    
    Call CloseMailFile
End Function

Public Sub GetMailMessage(ByVal sUser As String, ByRef theMessage As udtMail)
    Dim msgTemp As udtMail
    Dim i       As Integer
    
    Call OpenMailFile
    
    sUser = LCase$(sUser)
    
    If (CurrentRecord > 0) Then
        For i = 1 To CurrentRecord
            Get #CurrentOpenFile, i, msgTemp
            
            If (StrComp(sUser, RTrim(msgTemp.To), vbTextCompare) = 0) Then
                theMessage = msgTemp
                
                With msgTemp
                    .To = vbNullString
                End With
                
                Put #CurrentOpenFile, i, msgTemp
                
                Exit For
            End If
        Next i
    Else
        With theMessage
            .To = vbNullString
            .From = vbNullString
            .Message = vbNullString
        End With
    End If
    
    Call CloseMailFile
End Sub

Public Sub OpenMailFile()
    Dim Temp As udtMail
    Dim f    As Integer
    Dim i    As Integer
    
    f = FreeFile
    
    If (LenB(Dir$(GetFilePath("mail.dat"))) = 0) Then
        Open GetFilePath("mail.dat") For Output As #f
        Close #f
    End If
    
    Open GetFilePath("mail.dat") For Random As #f Len = LenB(Temp)
    
    If (LOF(f) > 0) Then
        i = LOF(f) \ LenB(Temp)
        
        If (LOF(f) Mod LenB(Temp) <> 0) Then
            i = (i + 1)
        End If
    Else
        i = 0
    End If
    
    CurrentRecord = i
    CurrentOpenFile = f
End Sub

Public Sub CloseMailFile()
    Close #CurrentOpenFile
End Sub

Public Sub CleanUpMailFile()
    Dim tMail() As udtMail
    Dim tTemp   As udtMail
    Dim i       As Long
    Dim c       As Long
    
    Call OpenMailFile
    
    If (CurrentRecord > 0) Then
        ReDim tMail(1 To CurrentRecord)
        
        If (LOF(CurrentOpenFile) > 0) Then
            ' mail in the mail file
            ' collect valid entries and rewrite it
            For i = 1 To CurrentRecord
                Get #CurrentOpenFile, i, tTemp
                
                tMail(i) = tTemp
            Next i
        End If
        
        Call CloseMailFile
        
        ' Zap the old file
        Call Kill(GetFilePath("mail.dat"))
        
        ' Write a new mail file
        Call OpenMailFile
        
        c = 1

        For i = 1 To UBound(tMail)
            If (Len(Trim(tMail(i).To)) > 0) Then
                Put #CurrentOpenFile, c, tMail(i)
                
                ' ...
                c = (c + 1)
            End If
        Next i
    End If

    Call CloseMailFile
End Sub
