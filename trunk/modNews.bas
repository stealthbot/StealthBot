Attribute VB_Name = "modNews"
Option Explicit

Public Sub HandleNews(ByVal s As String)

    On Error Resume Next
    
    Dim Splt() As String, SubSplt() As String
    Dim i As Integer
    Dim OldValue As Boolean
    
    OldValue = frmChat.mnuUTF8.Checked ' old value of UTF8 encoding setting
    
    Splt() = Split(s, "|")
    
    'Current ver code | Regular news | Beta news | Regular CVString | Beta CVString
    
    frmChat.mnuUTF8.Checked = False
    
    If UBound(Splt) = 4 Then
        If StrictIsNumeric(Splt(0)) Then
            frmChat.AddChat RTBColors.ServerInfoText, "The current StealthBot version is ÿcb" & Splt(3) & "."
            frmChat.AddChat RTBColors.ServerInfoText, " "
            
            If Val(Splt(0)) > VERCODE Then  '// old version
                frmChat.AddChat RTBColors.ErrorMessageText, "ÿcbYou are running an outdated version of StealthBot."
                frmChat.AddChat RTBColors.ErrorMessageText, "To download an updated version or for more information, visit http://www.stealthbot.net."
                frmChat.AddChat RTBColors.ErrorMessageText, "To disable version checking, add the line " & Chr(34) & "DisableSBNews=Y" & Chr(34) & " under the [Main] section of your config.ini file."
            End If

            If Len(Splt(1)) > 1 Then
                frmChat.AddChat RTBColors.ServerInfoText, ">> ÿcbStealthBot News"
                If InStr(1, Splt(1), "\n") > 0 Then
                    SubSplt() = Split(Splt(1), "\n")
                    
                    For i = 0 To UBound(SubSplt)
                        frmChat.AddChat RTBColors.ServerInfoText, ">> " & SubSplt(i)
                    Next i
                Else
                    frmChat.AddChat RTBColors.ServerInfoText, ">> " & Splt(1)
                End If
            End If
            
            '############# Beta only
            #If BETA Then
                frmChat.AddChat RTBColors.ServerInfoText, "->> "
                frmChat.AddChat RTBColors.ServerInfoText, "->> ÿcbStealthBot Beta News"
                
                If InStr(1, Splt(2), "\n") > 0 Then
                    SubSplt() = Split(Splt(2), "\n")
                    
                    For i = 0 To UBound(SubSplt)
                        frmChat.AddChat RTBColors.ServerInfoText, "->> " & SubSplt(i)
                    Next i
                Else
                    frmChat.AddChat RTBColors.ServerInfoText, "->> " & Splt(2)
                End If
                
                frmChat.AddChat RTBColors.ServerInfoText, " "
                frmChat.AddChat RTBColors.ServerInfoText, "The current StealthBot Beta version is " & Splt(4) & "."
            #End If
            '##############
            
            
        End If
    
        frmChat.mnuUTF8.Checked = OldValue
    End If
    
End Sub
