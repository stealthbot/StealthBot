<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEMailReg
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
	Public WithEvents cmdAskLater As System.Windows.Forms.Button
	Public WithEvents cmdIgnore As System.Windows.Forms.Button
	Public WithEvents cmdGo As System.Windows.Forms.Button
	Public WithEvents txtAddress As System.Windows.Forms.TextBox
	Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEMailReg))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.cmdAskLater = New System.Windows.Forms.Button
		Me.cmdIgnore = New System.Windows.Forms.Button
		Me.cmdGo = New System.Windows.Forms.Button
		Me.txtAddress = New System.Windows.Forms.TextBox
		Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "E-Mail Registration"
		Me.ClientSize = New System.Drawing.Size(433, 176)
		Me.Location = New System.Drawing.Point(7, 33)
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
		Me.Name = "frmEMailReg"
		Me.cmdAskLater.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdAskLater
		Me.cmdAskLater.Text = "&Ask Me Later"
		Me.cmdAskLater.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAskLater.Size = New System.Drawing.Size(193, 17)
		Me.cmdAskLater.Location = New System.Drawing.Point(232, 152)
		Me.cmdAskLater.TabIndex = 5
		Me.cmdAskLater.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAskLater.CausesValidation = True
		Me.cmdAskLater.Enabled = True
		Me.cmdAskLater.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAskLater.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAskLater.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAskLater.TabStop = True
		Me.cmdAskLater.Name = "cmdAskLater"
		Me.cmdIgnore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdIgnore.Text = "&Never Ask Again"
		Me.cmdIgnore.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdIgnore.Size = New System.Drawing.Size(97, 17)
		Me.cmdIgnore.Location = New System.Drawing.Point(328, 136)
		Me.cmdIgnore.TabIndex = 4
		Me.cmdIgnore.BackColor = System.Drawing.SystemColors.Control
		Me.cmdIgnore.CausesValidation = True
		Me.cmdIgnore.Enabled = True
		Me.cmdIgnore.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdIgnore.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdIgnore.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdIgnore.TabStop = True
		Me.cmdIgnore.Name = "cmdIgnore"
		Me.cmdGo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdGo.Text = "&OK"
		Me.AcceptButton = Me.cmdGo
		Me.cmdGo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdGo.Size = New System.Drawing.Size(97, 17)
		Me.cmdGo.Location = New System.Drawing.Point(232, 136)
		Me.cmdGo.TabIndex = 3
		Me.cmdGo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdGo.CausesValidation = True
		Me.cmdGo.Enabled = True
		Me.cmdGo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdGo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdGo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdGo.TabStop = True
		Me.cmdGo.Name = "cmdGo"
		Me.txtAddress.AutoSize = False
		Me.txtAddress.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtAddress.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAddress.ForeColor = System.Drawing.Color.White
		Me.txtAddress.Size = New System.Drawing.Size(217, 19)
		Me.txtAddress.Location = New System.Drawing.Point(8, 136)
		Me.txtAddress.Maxlength = 254
		Me.txtAddress.TabIndex = 2
		Me.txtAddress.AcceptsReturn = True
		Me.txtAddress.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAddress.CausesValidation = True
		Me.txtAddress.Enabled = True
		Me.txtAddress.HideSelection = True
		Me.txtAddress.ReadOnly = False
		Me.txtAddress.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAddress.MultiLine = False
		Me.txtAddress.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAddress.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAddress.TabStop = True
		Me.txtAddress.Visible = True
		Me.txtAddress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAddress.Name = "txtAddress"
		Me.Line1.BorderColor = System.Drawing.Color.White
		Me.Line1.X1 = 8
		Me.Line1.X2 = 424
		Me.Line1.Y1 = 128
		Me.Line1.Y2 = 128
		Me.Line1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line1.BorderWidth = 1
		Me.Line1.Visible = True
		Me.Line1.Name = "Line1"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.BackColor = System.Drawing.Color.Black
		Me.Label2.Text = "click here for more information"
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.White
		Me.Label2.Size = New System.Drawing.Size(153, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 104)
		Me.Label2.TabIndex = 1
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(417, 89)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 0
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdAskLater)
		Me.Controls.Add(cmdIgnore)
		Me.Controls.Add(cmdGo)
		Me.Controls.Add(txtAddress)
		Me.ShapeContainer1.Shapes.Add(Line1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(ShapeContainer1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class