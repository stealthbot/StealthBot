<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCustomInputBox
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
	Public WithEvents cboGame As System.Windows.Forms.ComboBox
	Public WithEvents cboServer As System.Windows.Forms.ComboBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdBack As System.Windows.Forms.Button
	Public WithEvents cmdNext As System.Windows.Forms.Button
	Public WithEvents txtInput As System.Windows.Forms.TextBox
	Public WithEvents lblText As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCustomInputBox))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cboGame = New System.Windows.Forms.ComboBox
		Me.cboServer = New System.Windows.Forms.ComboBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdBack = New System.Windows.Forms.Button
		Me.cmdNext = New System.Windows.Forms.Button
		Me.txtInput = New System.Windows.Forms.TextBox
		Me.lblText = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "(caption)"
		Me.ClientSize = New System.Drawing.Size(281, 148)
		Me.Location = New System.Drawing.Point(3, 24)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmCustomInputBox"
		Me.cboGame.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboGame.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboGame.ForeColor = System.Drawing.Color.White
		Me.cboGame.Size = New System.Drawing.Size(265, 21)
		Me.cboGame.Location = New System.Drawing.Point(8, 104)
		Me.cboGame.TabIndex = 6
		Me.cboGame.Text = "Choose One"
		Me.cboGame.CausesValidation = True
		Me.cboGame.Enabled = True
		Me.cboGame.IntegralHeight = True
		Me.cboGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboGame.Sorted = False
		Me.cboGame.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboGame.TabStop = True
		Me.cboGame.Visible = True
		Me.cboGame.Name = "cboGame"
		Me.cboServer.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.cboServer.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboServer.ForeColor = System.Drawing.Color.White
		Me.cboServer.Size = New System.Drawing.Size(265, 21)
		Me.cboServer.Location = New System.Drawing.Point(8, 104)
		Me.cboServer.TabIndex = 5
		Me.cboServer.Text = "Choose One"
		Me.cboServer.CausesValidation = True
		Me.cboServer.Enabled = True
		Me.cboServer.IntegralHeight = True
		Me.cboServer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboServer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboServer.Sorted = False
		Me.cboServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cboServer.TabStop = True
		Me.cboServer.Visible = True
		Me.cboServer.Name = "cboServer"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "X &Cancel"
		Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.Size = New System.Drawing.Size(57, 17)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 128)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdBack.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBack.Text = "<< &Back"
		Me.cmdBack.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBack.Size = New System.Drawing.Size(57, 17)
		Me.cmdBack.Location = New System.Drawing.Point(152, 128)
		Me.cmdBack.TabIndex = 2
		Me.cmdBack.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBack.CausesValidation = True
		Me.cmdBack.Enabled = True
		Me.cmdBack.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBack.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBack.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBack.TabStop = True
		Me.cmdBack.Name = "cmdBack"
		Me.cmdNext.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNext.Text = "&Next >>"
		Me.cmdNext.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNext.Size = New System.Drawing.Size(57, 17)
		Me.cmdNext.Location = New System.Drawing.Point(216, 128)
		Me.cmdNext.TabIndex = 1
		Me.cmdNext.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNext.CausesValidation = True
		Me.cmdNext.Enabled = True
		Me.cmdNext.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNext.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNext.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNext.TabStop = True
		Me.cmdNext.Name = "cmdNext"
		Me.txtInput.AutoSize = False
		Me.txtInput.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtInput.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtInput.ForeColor = System.Drawing.Color.White
		Me.txtInput.Size = New System.Drawing.Size(265, 19)
		Me.txtInput.Location = New System.Drawing.Point(8, 104)
		Me.txtInput.TabIndex = 0
		Me.txtInput.AcceptsReturn = True
		Me.txtInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtInput.CausesValidation = True
		Me.txtInput.Enabled = True
		Me.txtInput.HideSelection = True
		Me.txtInput.ReadOnly = False
		Me.txtInput.Maxlength = 0
		Me.txtInput.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtInput.MultiLine = False
		Me.txtInput.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtInput.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtInput.TabStop = True
		Me.txtInput.Visible = True
		Me.txtInput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtInput.Name = "txtInput"
		Me.lblText.BackColor = System.Drawing.Color.Black
		Me.lblText.Text = "[ message ]"
		Me.lblText.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblText.ForeColor = System.Drawing.Color.White
		Me.lblText.Size = New System.Drawing.Size(265, 97)
		Me.lblText.Location = New System.Drawing.Point(8, 8)
		Me.lblText.TabIndex = 4
		Me.lblText.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblText.Enabled = True
		Me.lblText.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblText.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblText.UseMnemonic = True
		Me.lblText.Visible = True
		Me.lblText.AutoSize = False
		Me.lblText.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblText.Name = "lblText"
		Me.Controls.Add(cboGame)
		Me.Controls.Add(cboServer)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdBack)
		Me.Controls.Add(cmdNext)
		Me.Controls.Add(txtInput)
		Me.Controls.Add(lblText)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class