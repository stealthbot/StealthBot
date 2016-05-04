Option Strict Off
Option Explicit On
Friend Class frmQuickChannel
	Inherits System.Windows.Forms.Form
	
	Private Sub frmQuickChannel_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Icon = frmChat.Icon
		
		Me.KeyPreview = True
		
		Dim i As Short
		
		' bounds of Channel controls
		For i = 0 To 8
			Channel(i).Text = QC(i + 1)
		Next i
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub frmQuickChannel_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = System.Windows.Forms.Keys.Return Then
			Call cmdDone_Click(cmdDone, New System.EventArgs())
		ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then 
			Call cmdCancel_Click(cmdCancel, New System.EventArgs())
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub cmdDone_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDone.Click
		Dim i As Short
		
		'write the qc list
		' bounds of Channel controls
		For i = 0 To 8
			QC(i + 1) = Channel(i).Text
		Next i
		
		SaveQuickChannels()
		
		PrepareQuickChannelMenu()
		
		Me.Close()
	End Sub
End Class