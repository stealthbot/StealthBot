Attribute VB_Name = "modMail"
'modMail - project StealthBot - authored 8/3/04 andy@stealthbot.net
Option Explicit

Private CurrentOpenFile As Integer
Private CurrentRecord As Long

Public Sub AddMail(ByRef tsMsg As udtMail)
    Call OpenMailFile
    
    tsMsg.To = LCase(tsMsg.To)
    Put #CurrentOpenFile, CurrentRecord + 1, tsMsg
    
    Call CloseMailFile
End Sub

Public Function GetMailCount(ByVal sUser As String) As Long
    Dim i As Long, Count As Long
    Dim mTemp As udtMail
    
    Call OpenMailFile
    
    sUser = LCase(sUser)
    
    If CurrentRecord > 0 Then
        For i = 1 To CurrentRecord
            Get #CurrentOpenFile, i, mTemp
            
            If StrComp(sUser, RTrim(mTemp.To)) = 0 Then
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
    Dim i As Integer
    
    Call OpenMailFile
    
    sUser = LCase(sUser)
    
    If CurrentRecord > 0 Then
        For i = 1 To CurrentRecord
            Get #CurrentOpenFile, i, msgTemp
            
            If StrComp(sUser, RTrim(msgTemp.To)) = 0 Then
                theMessage = msgTemp
                
                msgTemp.To = vbNullString
                
                Put #CurrentOpenFile, i, msgTemp
                Exit For
            End If
        Next i
    Else
        theMessage.To = ""
        theMessage.From = ""
        theMessage.Message = ""
    End If
    
    CloseMailFile
End Sub

Public Sub OpenMailFile()
    Dim f As Integer, i As Integer
    Dim Temp As udtMail
    
    f = FreeFile
    
    If LenB(Dir$(GetFilePath("mail.dat"))) = 0 Then
        Open GetFilePath("mail.dat") For Output As #f
        Close #f
    End If
    
    Open GetFilePath("mail.dat") For Random As #f Len = LenB(Temp)
    
    If LOF(f) > 0 Then
        i = LOF(f) \ LenB(Temp)
        If LOF(f) Mod LenB(Temp) <> 0 Then i = i + 1
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
    Dim tMail() As udtMail, tTemp As udtMail
    Dim i As Long, c As Long
    
    OpenMailFile
    
    'Debug.Print "Mail file open. CurrentRecord: " & CurrentRecord
    
    If CurrentRecord > 0 Then
        ReDim tMail(1 To CurrentRecord)
        
        'Debug.Print "Sizing array: 1 to " & CurrentRecord
        If LOF(CurrentOpenFile) > 0 Then
            ' mail in the mail file
            ' collect valid entries and rewrite it
            For i = 1 To CurrentRecord
                Get #CurrentOpenFile, i, tTemp
                
                tMail(i) = tTemp
                'Debug.Print " -- Got a message out of the file, for " & Trim(tTemp.To)
            Next i
        End If
        
        CloseMailFile
        
        ' Zap the old file
        Kill GetFilePath("mail.dat")
        
        ' Write a new mail file
        OpenMailFile
        
        c = 1
        
        'Debug.Print "Looping through tMail..."
        
        For i = 1 To UBound(tMail)
'            Debug.Print " -- c: " & c
'            Debug.Print " -- tMail(i).To: (" & Len(Trim(tMail(i).To)) & ") " & Trim(tMail(i).To)
'            Debug.Print " -- tMail(i).From: (" & Len(Trim(tMail(i).From)) & ") " & Trim(tMail(i).From)
'            Debug.Print " -- tMail(i).Message: (" & Len(Trim(tMail(i).Message)) & ") " & Trim(tMail(i).Message)
            
            If Len(Trim(tMail(i).To)) > 0 Then
                Put #CurrentOpenFile, c, tMail(i)
                c = c + 1
                
                'Debug.Print "-- Copied & incremented c!"
            End If
        Next i
    End If
    
    'Debug.Print "Finished!"
    CloseMailFile
End Sub
