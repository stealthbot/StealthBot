<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmDBNameEntry
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
	Public WithEvents txtEntry As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblEntry As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDBNameEntry))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtEntry = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblEntry = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.SystemColors.MenuText
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "Name Entry"
		Me.ClientSize = New System.Drawing.Size(249, 111)
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
		Me.Name = "frmDBNameEntry"
		Me.txtEntry.AutoSize = False
		Me.txtEntry.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtEntry.ForeColor = System.Drawing.Color.White
		Me.txtEntry.Size = New System.Drawing.Size(217, 19)
		Me.txtEntry.Location = New System.Drawing.Point(16, 56)
		Me.txtEntry.Maxlength = 30
		Me.txtEntry.TabIndex = 0
		Me.txtEntry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEntry.AcceptsReturn = True
		Me.txtEntry.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEntry.CausesValidation = True
		Me.txtEntry.Enabled = True
		Me.txtEntry.HideSelection = True
		Me.txtEntry.ReadOnly = False
		Me.txtEntry.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEntry.MultiLine = False
		Me.txtEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEntry.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEntry.TabStop = True
		Me.txtEntry.Visible = True
		Me.txtEntry.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEntry.Name = "txtEntry"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(57, 17)
		Me.cmdCancel.Location = New System.Drawing.Point(120, 80)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(57, 17)
		Me.cmdOK.Location = New System.Drawing.Point(176, 80)
		Me.cmdOK.TabIndex = 1
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.lblEntry.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblEntry.BackColor = System.Drawing.SystemColors.MenuText
		Me.lblEntry.Text = "Choose a name for the %s entry."
		Me.lblEntry.ForeColor = System.Drawing.Color.White
		Me.lblEntry.Size = New System.Drawing.Size(185, 33)
		Me.lblEntry.Location = New System.Drawing.Point(30, 8)
		Me.lblEntry.TabIndex = 2
		Me.lblEntry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblEntry.Enabled = True
		Me.lblEntry.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblEntry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblEntry.UseMnemonic = True
		Me.lblEntry.Visible = True
		Me.lblEntry.AutoSize = False
		Me.lblEntry.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblEntry.Name = "lblEntry"
		Me.Controls.Add(txtEntry)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblEntry)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class