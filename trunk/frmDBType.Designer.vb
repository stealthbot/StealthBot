<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDBType
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _btnSave_1 As System.Windows.Forms.Button
	Public WithEvents cbxChoice As System.Windows.Forms.ComboBox
	Public WithEvents btnDelete As System.Windows.Forms.Button
	Public WithEvents btnSave As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDBType))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._btnSave_1 = New System.Windows.Forms.Button
		Me.cbxChoice = New System.Windows.Forms.ComboBox
		Me.btnDelete = New System.Windows.Forms.Button
		Me.btnSave = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "Database Type"
		Me.ClientSize = New System.Drawing.Size(144, 60)
		Me.Location = New System.Drawing.Point(3, 21)
		Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmDBType"
		Me._btnSave_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._btnSave_1.Text = "OK"
		Me._btnSave_1.Size = New System.Drawing.Size(65, 20)
		Me._btnSave_1.Location = New System.Drawing.Point(8, 32)
		Me._btnSave_1.TabIndex = 1
		Me._btnSave_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._btnSave_1.BackColor = System.Drawing.SystemColors.Control
		Me._btnSave_1.CausesValidation = True
		Me._btnSave_1.Enabled = True
		Me._btnSave_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._btnSave_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._btnSave_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._btnSave_1.TabStop = True
		Me._btnSave_1.Name = "_btnSave_1"
		Me.cbxChoice.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cbxChoice.ForeColor = System.Drawing.Color.White
		Me.cbxChoice.Size = New System.Drawing.Size(129, 21)
		Me.cbxChoice.Location = New System.Drawing.Point(8, 8)
		Me.cbxChoice.Items.AddRange(New Object(){"Safelist (SB)", "Shitlist (SB)", "Tagbans (SB)"})
		Me.cbxChoice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbxChoice.TabIndex = 0
		Me.cbxChoice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbxChoice.CausesValidation = True
		Me.cbxChoice.Enabled = True
		Me.cbxChoice.IntegralHeight = True
		Me.cbxChoice.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbxChoice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbxChoice.Sorted = False
		Me.cbxChoice.TabStop = True
		Me.cbxChoice.Visible = True
		Me.cbxChoice.Name = "cbxChoice"
		Me.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnDelete.Text = "Cancel"
		Me.btnDelete.Size = New System.Drawing.Size(65, 20)
		Me.btnDelete.Location = New System.Drawing.Point(72, 32)
		Me.btnDelete.TabIndex = 2
		Me.btnDelete.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
		Me.btnDelete.CausesValidation = True
		Me.btnDelete.Enabled = True
		Me.btnDelete.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnDelete.TabStop = True
		Me.btnDelete.Name = "btnDelete"
		Me.Controls.Add(_btnSave_1)
		Me.Controls.Add(cbxChoice)
		Me.Controls.Add(btnDelete)
		Me.btnSave.SetIndex(_btnSave_1, CType(1, Short))
		CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class