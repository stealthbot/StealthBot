Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmSplash
	Inherits System.Windows.Forms.Form
	
	Private lStartTick As Integer ' The tick count when the form was loaded.
	Private bHasShown As Boolean ' Set to true after 1 second
	
	' Number of milliseconds before the form automatically unloads
	Private Const AUTO_UNLOAD_DELAY As Short = 15000
	
	Private Sub Bday_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Bday.Click
		frmChat.Show()
		Me.Close()
	End Sub
	
	Private Sub frmSplash_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		frmChat.Show()
		Me.Close()
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub Image_click()
		frmChat.Show()
		Me.Close()
	End Sub
	
	
	
	Private Sub IsBirthday(ByRef sName As String, ByRef iBorn As Short)
		On Error Resume Next 'this isn't important so screw the errors
		Dim iAge As Short
		Dim sAppend As String
		
		Logo.Visible = False
		BDay.Visible = True
		
		sAppend = "th"
		iAge = Year(Now) - iBorn
		If (iAge <= 10 Or iAge >= 20) Then
			If (iAge Mod 10 = 1) Then
				sAppend = "st"
			ElseIf (iAge Mod 10 = 2) Then 
				sAppend = "nd"
			ElseIf (iAge Mod 10 = 3) Then 
				sAppend = "rd"
			End If
		End If
		
		Label1.Text = "Happy " & iAge & sAppend & " Birthday " & sName & "!!!"
	End Sub
	
	Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error Resume Next 'this isn't important so screw the errors
		Dim sDate As String
		Me.Icon = frmChat.Icon
		
		sDate = LCase(Month(Now) & "/" & VB.Day(Now))
		
		Select Case sDate
			Case "2/21" : IsBirthday("Ribose", 1992)
			Case "3/18" : IsBirthday("Pyro", 1992)
			Case "5/9" : IsBirthday("Snap", 1987)
			Case "4/3" : IsBirthday("Stealth", 1987)
			Case "4/22" : IsBirthday("52", 1982)
			Case "9/22" : IsBirthday("Eric", 1987)
			Case "11/26" : IsBirthday("Hdx", 1989)
			Case Else : Label1.Text = "[ " & CVERSION & " ]"
		End Select
		
		lStartTick = GetTickCount
	End Sub
	
	Private Sub frmSplash_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		If (bHasShown) Then
			tmrUnload.Enabled = False
			Me.Close()
		End If
	End Sub
	
	Private Sub Logo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Logo.Click
		frmChat.Show()
		Me.Close()
	End Sub
	
	Private Sub tmrUnload_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrUnload.Tick
		On Error GoTo KILL_FORM
		
		If ((GetTickCount - lStartTick) >= AUTO_UNLOAD_DELAY) Then
			GoTo KILL_FORM
		End If
		bHasShown = True
		Exit Sub
		
KILL_FORM: 
		tmrUnload.Enabled = False
		Me.Close()
		Exit Sub
	End Sub
End Class