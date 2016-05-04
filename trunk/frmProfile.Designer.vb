<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmProfile
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents rtbLocation As System.Windows.Forms.RichTextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents rtbProfile As System.Windows.Forms.RichTextBox
	Public WithEvents rtbSex As System.Windows.Forms.RichTextBox
	Public WithEvents rtbAge As System.Windows.Forms.RichTextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents lblUsername As System.Windows.Forms.Label
	Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProfile))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.cmdOK = New System.Windows.Forms.Button
		Me.rtbLocation = New System.Windows.Forms.RichTextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.rtbProfile = New System.Windows.Forms.RichTextBox
		Me.rtbSex = New System.Windows.Forms.RichTextBox
		Me.rtbAge = New System.Windows.Forms.RichTextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.lblUsername = New System.Windows.Forms.Label
		Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "Profile Viewer"
		Me.ClientSize = New System.Drawing.Size(527, 297)
		Me.Location = New System.Drawing.Point(10, 36)
		Me.Icon = CType(resources.GetObject("frmProfile.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmProfile"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "&Write"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.Size = New System.Drawing.Size(57, 41)
		Me.cmdOK.Location = New System.Drawing.Point(8, 248)
		Me.cmdOK.TabIndex = 3
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.rtbLocation.Size = New System.Drawing.Size(441, 19)
		Me.rtbLocation.Location = New System.Drawing.Point(80, 80)
		Me.rtbLocation.TabIndex = 1
		Me.rtbLocation.BackColor = System.Drawing.Color.Black
		Me.rtbLocation.Enabled = True
		Me.rtbLocation.RTF = resources.GetString("rtbLocation.TextRTF")
		Me.rtbLocation.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbLocation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.rtbLocation.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
		Me.rtbLocation.Name = "rtbLocation"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "&Cancel"
		Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.Size = New System.Drawing.Size(57, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 224)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.rtbProfile.Size = New System.Drawing.Size(441, 185)
		Me.rtbProfile.Location = New System.Drawing.Point(80, 104)
		Me.rtbProfile.TabIndex = 2
		Me.rtbProfile.BackColor = System.Drawing.Color.Black
		Me.rtbProfile.Enabled = True
		Me.rtbProfile.ReadOnly = True
		Me.rtbProfile.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rtbProfile.RTF = resources.GetString("rtbProfile.TextRTF")
		Me.rtbProfile.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbProfile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rtbProfile.Name = "rtbProfile"
		Me.rtbSex.Size = New System.Drawing.Size(441, 19)
		Me.rtbSex.Location = New System.Drawing.Point(80, 56)
		Me.rtbSex.TabIndex = 0
		Me.rtbSex.BackColor = System.Drawing.Color.Black
		Me.rtbSex.Enabled = True
		Me.rtbSex.RTF = resources.GetString("rtbSex.TextRTF")
		Me.rtbSex.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbSex.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.rtbSex.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
		Me.rtbSex.Name = "rtbSex"
		Me.rtbAge.Size = New System.Drawing.Size(441, 19)
		Me.rtbAge.Location = New System.Drawing.Point(80, 32)
		Me.rtbAge.TabIndex = 5
		Me.rtbAge.BackColor = System.Drawing.Color.Black
		Me.rtbAge.Enabled = True
		Me.rtbAge.RTF = resources.GetString("rtbAge.TextRTF")
		Me.rtbAge.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbAge.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.rtbAge.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
		Me.rtbAge.Name = "rtbAge"
		Me.Label5.BackColor = System.Drawing.Color.Black
		Me.Label5.Text = "Sex"
		Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.ForeColor = System.Drawing.Color.White
		Me.Label5.Size = New System.Drawing.Size(57, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 56)
		Me.Label5.TabIndex = 9
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.Enabled = True
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.lblUsername.BackColor = System.Drawing.Color.Black
		Me.lblUsername.Text = "Username"
		Me.lblUsername.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUsername.ForeColor = System.Drawing.Color.White
		Me.lblUsername.Size = New System.Drawing.Size(201, 17)
		Me.lblUsername.Location = New System.Drawing.Point(80, 8)
		Me.lblUsername.TabIndex = 6
		Me.lblUsername.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblUsername.Enabled = True
		Me.lblUsername.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUsername.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUsername.UseMnemonic = True
		Me.lblUsername.Visible = True
		Me.lblUsername.AutoSize = False
		Me.lblUsername.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUsername.Name = "lblUsername"
		Me.Line1.BorderColor = System.Drawing.SystemColors.ActiveCaptionText
		Me.Line1.X1 = 72
		Me.Line1.X2 = 72
		Me.Line1.Y1 = 8
		Me.Line1.Y2 = 288
		Me.Line1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line1.BorderWidth = 1
		Me.Line1.Visible = True
		Me.Line1.Name = "Line1"
		Me.Label4.BackColor = System.Drawing.Color.Black
		Me.Label4.Text = "Description"
		Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.ForeColor = System.Drawing.Color.White
		Me.Label4.Size = New System.Drawing.Size(57, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 104)
		Me.Label4.TabIndex = 11
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.Enabled = True
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.BackColor = System.Drawing.Color.Black
		Me.Label3.Text = "Location"
		Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.ForeColor = System.Drawing.Color.White
		Me.Label3.Size = New System.Drawing.Size(57, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 80)
		Me.Label3.TabIndex = 10
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.Enabled = True
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.BackColor = System.Drawing.Color.Black
		Me.Label2.Text = "Age"
		Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.Color.White
		Me.Label2.Size = New System.Drawing.Size(57, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 32)
		Me.Label2.TabIndex = 8
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Text = "Username"
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.White
		Me.Label1.Size = New System.Drawing.Size(57, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 7
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(rtbLocation)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(rtbProfile)
		Me.Controls.Add(rtbSex)
		Me.Controls.Add(rtbAge)
		Me.Controls.Add(Label5)
		Me.Controls.Add(lblUsername)
		Me.ShapeContainer1.Shapes.Add(Line1)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(ShapeContainer1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class