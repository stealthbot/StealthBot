Option Strict Off
Option Explicit On
Friend Class frmCatch
	Inherits System.Windows.Forms.Form
	
	'UPGRADE_WARNING: Event chkFlashOnCaughtPhrase.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkFlashOnCaughtPhrase_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkFlashOnCaughtPhrase.CheckStateChanged
		If Config.FlashOnCatchPhrases <> CBool(chkFlashOnCaughtPhrase.CheckState) Then
			Config.FlashOnCatchPhrases = CBool(chkFlashOnCaughtPhrase.CheckState)
			Call Config.Save()
		End If
	End Sub
	
	Private Sub cmdDone_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDone.Click
		Dim i, f As Short
		'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		ReDim Preserve Catch_Renamed(0)
		If lbCatch.Items.Count < 0 Then
			Me.Close()
			Exit Sub
		End If
		
		f = FreeFile
		FileOpen(f, GetFilePath(FILE_CATCH_PHRASES), OpenMode.Output)
		
		'UPGRADE_NOTE: Catch was upgraded to Catch_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		ReDim Preserve Catch_Renamed(UBound(Catch_Renamed) + 1)
		For i = 0 To lbCatch.Items.Count
			Catch_Renamed(i) = VB6.GetItemString(lbCatch, i)
			PrintLine(f, VB6.GetItemString(lbCatch, i))
			If i <> lbCatch.Items.Count Then 
			End If
		Next i
		
		FileClose(f)
		Me.Close()
	End Sub
	
	Private Sub cmdEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEdit.Click
		If lbCatch.SelectedIndex >= 0 Then
			txtModify.Text = lbCatch.Text
			lbCatch.Items.RemoveAt(lbCatch.SelectedIndex)
		End If
	End Sub
	
	Private Sub cmdOutAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOutAdd.Click
		If txtModify.Text <> vbNullString Then
			lbCatch.Items.Add(txtModify.Text)
			txtModify.Text = vbNullString
		End If
	End Sub
	
	Private Sub cmdOutRem_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOutRem.Click
		If lbCatch.SelectedIndex >= 0 Then
			lbCatch.Items.RemoveAt(lbCatch.SelectedIndex)
		End If
	End Sub
	
	Private Sub frmCatch_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		If Config.FlashOnCatchPhrases Then
			chkFlashOnCaughtPhrase.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		
		Dim i As Short
		For i = LBound(Catch_Renamed) To UBound(Catch_Renamed)
			If Catch_Renamed(i) <> vbNullString Then
				lbCatch.Items.Add(Catch_Renamed(i))
			End If
		Next i
	End Sub
	
	Private Sub frmCatch_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call cmdDone_Click(cmdDone, New System.EventArgs())
	End Sub
End Class