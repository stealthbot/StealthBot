Option Strict Off
Option Explicit On
Friend Class frmDBType
	Inherits System.Windows.Forms.Form
	
	Private m_path As String
	
	Private Sub frmDBType_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		cbxChoice.SelectedIndex = 0
	End Sub
	
	Private Sub btnDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnDelete.Click
		Me.Close()
	End Sub
	
	Private Sub btnSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click
		Dim Index As Short = btnSave.GetIndex(eventSender)
		Call frmDBManager.ImportDatabase(m_path, (cbxChoice.SelectedIndex))
		
		Me.Close()
	End Sub
	
	Public Sub setFilePath(ByRef strPath As String)
		m_path = strPath
	End Sub
End Class