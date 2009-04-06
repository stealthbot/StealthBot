Attribute VB_Name = "modMenu"
'Using the Menu APIs to Grow or Shrink a Menu During Run-time
'(c) Jon Vote, 2003
'
'Idioma Software Inc.
'jon@idioma-software.com
'www.idioma-software.com
'www.skycoder.com


'Adapted to StealthBot
' 2007-06-10, Andy T

Option Explicit

' ...
Public Sub ProcessMenu(hWnd As Long, lngMenuCommand As Long)
    Dim obj As scObj ' ...

    ' ...
    obj = GetScriptObjByMenuID(lngMenuCommand)
    
    ' is this a dynamic scripting menu?
    If (obj.ObjName <> vbNullString) Then
        On Error Resume Next

        obj.SCModule.Run obj.ObjName & "_Click"
    Else
        Dim I As Integer ' ...
        
        For I = 1 To DynamicMenus.Count
            If (DynamicMenus(I).ID = lngMenuCommand) Then
                ' is this a default scripting menu?
                If (Left$(DynamicMenus(I).Name, 1) = Chr$(0)) Then
                    Dim s_name   As String ' ...
                    Dim sub_name As String ' ...

                    s_name = _
                        Split(Mid$(DynamicMenus(I).Name, 2))(0)
                    sub_name = _
                        Split(Mid$(DynamicMenus(I).Name, 2))(1)
                        
                    If (sub_name = "ENABLE|DISABLE") Then
                        If (DynamicMenus(I).Checked) Then
                            ProcessCommand GetCurrentUsername, "/disable " & s_name, True
                            
                            DynamicMenus(I).Checked = False
                        Else
                            ProcessCommand GetCurrentUsername, "/enable " & s_name, True
                            
                            DynamicMenus(I).Checked = True
                        End If
                    ElseIf (sub_name = "VIEW_SCRIPT") Then
                        Shell "notepad " & Scripts(s_name).Script("Path"), vbNormalFocus
                    End If
                End If
                
                Exit For
            End If
        Next I
    End If
End Sub
