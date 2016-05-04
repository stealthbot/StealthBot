<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCatch
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
	Public WithEvents chkFlashOnCaughtPhrase As System.Windows.Forms.CheckBox
	Public WithEvents lbCatch As System.Windows.Forms.ListBox
	Public WithEvents cmdDone As System.Windows.Forms.Button
	Public WithEvents cmdOutAdd As System.Windows.Forms.Button
	Public WithEvents txtModify As System.Windows.Forms.TextBox
	Public WithEvents cmdOutRem As System.Windows.Forms.Button
	Public WithEvents cmdEdit As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCatch))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkFlashOnCaughtPhrase = New System.Windows.Forms.CheckBox
		Me.lbCatch = New System.Windows.Forms.ListBox
		Me.cmdDone = New System.Windows.Forms.Button
		Me.cmdOutAdd = New System.Windows.Forms.Button
		Me.txtModify = New System.Windows.Forms.TextBox
		Me.cmdOutRem = New System.Windows.Forms.Button
		Me.cmdEdit = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Catch Phrases"
		Me.ClientSize = New System.Drawing.Size(264, 398)
		Me.Location = New System.Drawing.Point(4, 27)
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
		Me.Name = "frmCatch"
		Me.chkFlashOnCaughtPhrase.BackColor = System.Drawing.Color.Black
		Me.chkFlashOnCaughtPhrase.Text = "Flash window on caught phrases"
		Me.chkFlashOnCaughtPhrase.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFlashOnCaughtPhrase.ForeColor = System.Drawing.Color.White
		Me.chkFlashOnCaughtPhrase.Size = New System.Drawing.Size(185, 17)
		Me.chkFlashOnCaughtPhrase.Location = New System.Drawing.Point(40, 328)
		Me.chkFlashOnCaughtPhrase.TabIndex = 8
		Me.chkFlashOnCaughtPhrase.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFlashOnCaughtPhrase.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFlashOnCaughtPhrase.CausesValidation = True
		Me.chkFlashOnCaughtPhrase.Enabled = True
		Me.chkFlashOnCaughtPhrase.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFlashOnCaughtPhrase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFlashOnCaughtPhrase.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFlashOnCaughtPhrase.TabStop = True
		Me.chkFlashOnCaughtPhrase.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFlashOnCaughtPhrase.Visible = True
		Me.chkFlashOnCaughtPhrase.Name = "chkFlashOnCaughtPhrase"
		Me.lbCatch.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.lbCatch.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbCatch.ForeColor = System.Drawing.Color.White
		Me.lbCatch.Size = New System.Drawing.Size(249, 228)
		Me.lbCatch.Location = New System.Drawing.Point(8, 32)
		Me.lbCatch.TabIndex = 6
		Me.lbCatch.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lbCatch.CausesValidation = True
		Me.lbCatch.Enabled = True
		Me.lbCatch.IntegralHeight = True
		Me.lbCatch.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbCatch.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lbCatch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbCatch.Sorted = False
		Me.lbCatch.TabStop = True
		Me.lbCatch.Visible = True
		Me.lbCatch.MultiColumn = False
		Me.lbCatch.Name = "lbCatch"
		Me.cmdDone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDone.Text = "&Done"
		Me.cmdDone.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDone.Size = New System.Drawing.Size(121, 17)
		Me.cmdDone.Location = New System.Drawing.Point(136, 304)
		Me.cmdDone.TabIndex = 3
		Me.cmdDone.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDone.CausesValidation = True
		Me.cmdDone.Enabled = True
		Me.cmdDone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDone.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDone.TabStop = True
		Me.cmdDone.Name = "cmdDone"
		Me.cmdOutAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOutAdd.Text = "&Add It!"
		Me.cmdOutAdd.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOutAdd.Size = New System.Drawing.Size(121, 17)
		Me.cmdOutAdd.Location = New System.Drawing.Point(136, 288)
		Me.cmdOutAdd.TabIndex = 1
		Me.cmdOutAdd.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOutAdd.CausesValidation = True
		Me.cmdOutAdd.Enabled = True
		Me.cmdOutAdd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOutAdd.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOutAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOutAdd.TabStop = True
		Me.cmdOutAdd.Name = "cmdOutAdd"
		Me.txtModify.AutoSize = False
		Me.txtModify.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtModify.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtModify.ForeColor = System.Drawing.Color.White
		Me.txtModify.Size = New System.Drawing.Size(249, 19)
		Me.txtModify.Location = New System.Drawing.Point(8, 264)
		Me.txtModify.TabIndex = 0
		Me.txtModify.AcceptsReturn = True
		Me.txtModify.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtModify.CausesValidation = True
		Me.txtModify.Enabled = True
		Me.txtModify.HideSelection = True
		Me.txtModify.ReadOnly = False
		Me.txtModify.Maxlength = 0
		Me.txtModify.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtModify.MultiLine = False
		Me.txtModify.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtModify.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtModify.TabStop = True
		Me.txtModify.Visible = True
		Me.txtModify.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtModify.Name = "txtModify"
		Me.cmdOutRem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOutRem.Text = "&Remove Selected Item"
		Me.cmdOutRem.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOutRem.Size = New System.Drawing.Size(129, 17)
		Me.cmdOutRem.Location = New System.Drawing.Point(8, 304)
		Me.cmdOutRem.TabIndex = 2
		Me.cmdOutRem.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOutRem.CausesValidation = True
		Me.cmdOutRem.Enabled = True
		Me.cmdOutRem.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOutRem.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOutRem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOutRem.TabStop = True
		Me.cmdOutRem.Name = "cmdOutRem"
		Me.cmdEdit.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdEdit.Text = "&Edit Selected Item"
		Me.cmdEdit.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdEdit.Size = New System.Drawing.Size(129, 17)
		Me.cmdEdit.Location = New System.Drawing.Point(8, 288)
		Me.cmdEdit.TabIndex = 4
		Me.cmdEdit.BackColor = System.Drawing.SystemColors.Control
		Me.cmdEdit.CausesValidation = True
		Me.cmdEdit.Enabled = True
		Me.cmdEdit.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdEdit.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdEdit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdEdit.TabStop = True
		Me.cmdEdit.Name = "cmdEdit"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.BackColor = System.Drawing.Color.Black
		Me.Label2.Text = "Any phrase that the bot sees that contains a phrase listed here will be recorded in the 'caughtphrases.htm' file in the bot's folder."
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.White
		Me.Label2.Size = New System.Drawing.Size(249, 41)
		Me.Label2.Location = New System.Drawing.Point(8, 352)
		Me.Label2.TabIndex = 7
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Text = "-- StealthBot Catch Phrases --"
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(177, 17)
		Me.Label1.Location = New System.Drawing.Point(40, 8)
		Me.Label1.TabIndex = 5
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(chkFlashOnCaughtPhrase)
		Me.Controls.Add(lbCatch)
		Me.Controls.Add(cmdDone)
		Me.Controls.Add(cmdOutAdd)
		Me.Controls.Add(txtModify)
		Me.Controls.Add(cmdOutRem)
		Me.Controls.Add(cmdEdit)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class