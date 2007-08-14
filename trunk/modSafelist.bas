Attribute VB_Name = "modSafelist"
' modSafelist
' Encapsulated safelist code
' Written 6/17/06 on I-80 East in Iowa
Option Explicit

Public Sub LoadSafelist()
    Dim Path As String
    Dim f As Integer
    Dim s As String
    Dim slTemp As clsSafelistEntry
    
    With colSafelist
        While .Count > 0
            .Remove 1
        Wend
    End With
    
    f = FreeFile
    
    Path = GetFilePath("safelist.txt")
    
    If LenB(Dir$(Path)) Then
        Open Path For Input As #f
        
            Do While Not EOF(f)
                Set slTemp = New clsSafelistEntry
                Line Input #f, s
                
                s = Trim(s)
                
                slTemp.Name = PrepareCheck(GetStringChunk(s, 1))
                slTemp.AddedBy = GetStringChunk(s, 2)
                
                If slTemp.AddedBy = "%" Then slTemp.AddedBy = ""
                
                colSafelist.Add slTemp
                
                Set slTemp = Nothing
            Loop
        Close #f
    End If
End Sub

Public Function AddToSafelist(ByVal Username As String, Optional ByVal Speaker As String) As String
    Dim c As Integer
    Dim slTemp As clsSafelistEntry
    
    Username = LCase(Username)
    
    If GetSafelist(Username) Then
        AddToSafelist = "That user is already safelisted, or matches a safelisted tag."
        Exit Function
    End If
    
    For c = LBound(gBans) To UBound(gBans)
        If StrComp(gBans(c).Username, Username) = 0 Then
            frmChat.AddQ "/unban " & IIf(Dii, "*", "") & gBans(c).Username, 1
        End If
    Next c
    
    Set slTemp = New clsSafelistEntry
        slTemp.Name = PrepareCheck(Username)
        slTemp.AddedBy = Speaker
        
        colSafelist.Add slTemp
    Set slTemp = Nothing
    
    Call UpdateSafelistedStatus(Username, True)
    
    WriteSafelist
    
    If bFlood Then
        
        gFloodSafelist(UBound(gFloodSafelist)) = Replace(PrepareCheck(Username), " ", vbNullString)
        ReDim Preserve gFloodSafelist(UBound(gFloodSafelist) + 1)
        
    End If
End Function

Public Sub WriteSafelist()
    Dim Y As String
    Dim c As Integer, f As Integer
    
    Y = GetFilePath("safelist.txt")
    f = FreeFile
    
    Open (Y) For Output As #f
        For c = 1 To colSafelist.Count
            With colSafelist.Item(c)
                If LenB(.Name) > 0 Then
                    Print #f, ReversePrepareCheck(.Name) & Space(1) & IIf(.AddedBy <> "", .AddedBy, "%")
                End If
            End With
        Next c
    Close #f
End Sub

Public Function RemoveFromSafelist(ByVal Username As String) As Boolean
    Dim c As Integer
    
    Username = PrepareCheck(Username)
    
    For c = 1 To colSafelist.Count
        If StrComp(colSafelist.Item(c).Name, Username) = 0 Then
            colSafelist.Remove c
            
            WriteSafelist
            
            RemoveFromSafelist = True
            Exit Function
        End If
    Next c
    
    RemoveFromSafelist = False
End Function
