<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmClanInvite
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
	Public WithEvents cmdDecline As System.Windows.Forms.Button
	Public WithEvents cmdAccept As System.Windows.Forms.Button
	Public WithEvents lblUser As System.Windows.Forms.Label
	Public WithEvents lblClan As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmClanInvite))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdDecline = New System.Windows.Forms.Button
		Me.cmdAccept = New System.Windows.Forms.Button
		Me.lblUser = New System.Windows.Forms.Label
		Me.lblClan = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Warcraft III Clan Invitation"
		Me.ClientSize = New System.Drawing.Size(217, 127)
		Me.Location = New System.Drawing.Point(5, 31)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmClanInvite"
		Me.cmdDecline.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDecline.Text = "&Decline"
		Me.cmdDecline.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDecline.Size = New System.Drawing.Size(89, 25)
		Me.cmdDecline.Location = New System.Drawing.Point(112, 96)
		Me.cmdDecline.TabIndex = 1
		Me.cmdDecline.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDecline.CausesValidation = True
		Me.cmdDecline.Enabled = True
		Me.cmdDecline.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDecline.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDecline.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDecline.TabStop = True
		Me.cmdDecline.Name = "cmdDecline"
		Me.cmdAccept.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAccept.Text = "&Accept"
		Me.cmdAccept.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAccept.Size = New System.Drawing.Size(89, 25)
		Me.cmdAccept.Location = New System.Drawing.Point(16, 96)
		Me.cmdAccept.TabIndex = 0
		Me.cmdAccept.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAccept.CausesValidation = True
		Me.cmdAccept.Enabled = True
		Me.cmdAccept.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAccept.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAccept.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAccept.TabStop = True
		Me.cmdAccept.Name = "cmdAccept"
		Me.lblUser.BackColor = System.Drawing.Color.Black
		Me.lblUser.Text = "A. Random Person"
		Me.lblUser.Font = New System.Drawing.Font("Tahoma", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUser.ForeColor = System.Drawing.Color.White
		Me.lblUser.Size = New System.Drawing.Size(201, 25)
		Me.lblUser.Location = New System.Drawing.Point(8, 8)
		Me.lblUser.TabIndex = 4
		Me.lblUser.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUser.Enabled = True
		Me.lblUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUser.UseMnemonic = True
		Me.lblUser.Visible = True
		Me.lblUser.AutoSize = False
		Me.lblUser.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUser.Name = "lblUser"
		Me.lblClan.BackColor = System.Drawing.Color.Black
		Me.lblClan.Text = "Clan %clan"
		Me.lblClan.Font = New System.Drawing.Font("Tahoma", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblClan.ForeColor = System.Drawing.Color.White
		Me.lblClan.Size = New System.Drawing.Size(201, 25)
		Me.lblClan.Location = New System.Drawing.Point(8, 64)
		Me.lblClan.TabIndex = 3
		Me.lblClan.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblClan.Enabled = True
		Me.lblClan.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblClan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblClan.UseMnemonic = True
		Me.lblClan.Visible = True
		Me.lblClan.AutoSize = False
		Me.lblClan.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblClan.Name = "lblClan"
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Text = "has invited you to join"
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(201, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 40)
		Me.Label1.TabIndex = 2
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdDecline)
		Me.Controls.Add(cmdAccept)
		Me.Controls.Add(lblUser)
		Me.Controls.Add(lblClan)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class