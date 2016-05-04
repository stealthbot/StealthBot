Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAbout
	Inherits System.Windows.Forms.Form
	
	Private lNumClicks As Integer
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		lblTitle.ForeColor = System.Drawing.Color.White
		txtDescr.ForeColor = System.Drawing.Color.White
		lblTitle.Text = ".: " & CVERSION
		
		txtDescr.ReadOnly = True
		
		AddLine("The list of current StealthBot project contributors can be found at")
		AddLine("-> http://contributors.stealthbot.net <-")
		AddCrLf()
		AddLine("StealthBot is updated and maintained by a team of developers who mostly volunteer their time.")
		AddLine("-> The current list of developers can be found at http://devteam.stealthbot.net.")
		AddLine("-> Donations to the project are split between developers after expenses.")
		AddCrLf()
		AddLine("All Blizzard copyrights are (c) 1996-present Blizzard Entertainment.")
		AddLine("For a detailed listing of Blizzard copyrights, please visit")
		AddLine("  http://www.blizzard.com/copyright.shtml")
		AddLine("For any further legal information regarding StealthBot and the StealthBot website, visit")
		AddLine("  http://www.stealthbot.net/legal/web/")
		AddCrLf()
		AddLine("- Hdx for his continued help and maintenance of the JBLS server at jbls.org")
		AddLine("- The staff and users at StealthBot.net for their continued support")
		AddCrLf()
		AddLine("- Thanks to all of the beta testers and people who suggested features to me -- StealthBot wouldn't be possible without you")
		AddLine("- Thanks to Retain, Jack, Hdx, and Berzerker, for their long lists of suggestions and bug reports")
		AddLine("- Thanks to PhiX for his excellent continued support help on StealthBot.net")
		AddCrLf()
		AddLine("- My website administrators & moderators, for their continued help in managing the stealthbot.net forums")
		AddCrLf()
		AddLine("- And, an extra-special thanks to:")
		AddLine("-> The tech support people on StealthBot.net and in Clan SBS, for giving their time to help others")
		AddLine("-> The people who have donated to the StealthBot project - you can find a current list at http://contributors.stealthbot.net - THANK YOU!")
		
		lblBottom.Text = "(c)2002-2009 Andy T - all rights reserved." & vbCrLf & "Use of this program is subject to the License Agreement found at http://eula.stealthbot.net."
	End Sub
	
	Private Sub AddLine(ByVal sIn As String)
		txtDescr.Text = txtDescr.Text & sIn & vbCrLf
	End Sub
	
	Private Sub AddCrLf()
		txtDescr.Text = txtDescr.Text & vbCrLf
	End Sub
	
	Private Sub frmAbout_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.FromPixelsUserX(eventArgs.X, 0, 8240.18, 585)
		Dim y As Single = VB6.FromPixelsUserY(eventArgs.Y, 0, 4006.71, 387)
		Dim i As Byte
		For i = 0 To 3
			lblURL(i).ForeColor = System.Drawing.Color.White
		Next i
		lblOK.ForeColor = System.Drawing.Color.White
	End Sub
	
	Private Sub lblOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblOK.Click
		lNumClicks = 0
		Me.Close()
	End Sub
	
	Private Sub lblOK_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblOK.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim i As Byte
		For i = 0 To 3
			lblURL(i).ForeColor = System.Drawing.Color.White
		Next i
		lblOK.ForeColor = System.Drawing.Color.Blue
	End Sub
	
	Private Sub lblTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblTitle.Click
		TitleClicked()
	End Sub
	
	Private Sub lblTitle_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblTitle.DoubleClick
		TitleClicked()
	End Sub
	
	Sub TitleClicked()
		lNumClicks = lNumClicks + 1
		
		If lNumClicks Mod 14 = 0 Then
			lblTitle.Text = " hamtaro rocks :P "
		ElseIf lNumClicks Mod 7 = 0 Then 
			lblTitle.Text = "< Think Outside the Bun >"
		Else
			lblTitle.Text = ".: " & CVERSION
		End If
	End Sub
	
	Private Sub lblURL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblURL.Click
		Dim Index As Short = lblURL.GetIndex(eventSender)
		Select Case Index
			Case 0 : ShellOpenURL("http://www.stealthbot.net", "the StealthBot Forum")
			Case 1 : ShellOpenURL("mailto:stealth@stealthbot.net")
			Case 2 : ShellOpenURL("http://contributors.stealthbot.net", "the StealthBot Contributors List")
			Case 3 : ShellOpenURL("http://www.stealthbot.net/wiki/Main_Page", "the StealthBot Wiki")
		End Select
	End Sub
	
	Private Sub lblURL_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblURL.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lblURL.GetIndex(eventSender)
		Dim i As Byte
		For i = 0 To 3
			lblURL(i).ForeColor = System.Drawing.Color.White
		Next i
		lblOK.ForeColor = System.Drawing.Color.White
		lblURL(Index).ForeColor = System.Drawing.Color.Blue
	End Sub
	
	Private Sub txtDescr_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescr.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Or KeyAscii = System.Windows.Forms.Keys.Escape Then
			Me.Close()
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class