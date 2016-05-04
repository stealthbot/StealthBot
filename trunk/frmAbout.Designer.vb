<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout
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
	Public WithEvents txtDescr As System.Windows.Forms.TextBox
	Public WithEvents _lblURL_2 As System.Windows.Forms.Label
	Public WithEvents lblSpecialThanks As System.Windows.Forms.Label
	Public WithEvents _lblURL_3 As System.Windows.Forms.Label
	Public WithEvents _lblURL_1 As System.Windows.Forms.Label
	Public WithEvents _lblURL_0 As System.Windows.Forms.Label
	Public WithEvents lblOK As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents _Line1_0 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblBottom As System.Windows.Forms.Label
	Public WithEvents Line1 As LineShapeArray
	Public WithEvents lblURL As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.txtDescr = New System.Windows.Forms.TextBox
		Me._lblURL_2 = New System.Windows.Forms.Label
		Me.lblSpecialThanks = New System.Windows.Forms.Label
		Me._lblURL_3 = New System.Windows.Forms.Label
		Me._lblURL_1 = New System.Windows.Forms.Label
		Me._lblURL_0 = New System.Windows.Forms.Label
		Me.lblOK = New System.Windows.Forms.Label
		Me._Line1_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblTitle = New System.Windows.Forms.Label
		Me._Line1_0 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblBottom = New System.Windows.Forms.Label
		Me.Line1 = New LineShapeArray(components)
		Me.lblURL = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblURL, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "About StealthBot"
		Me.ClientSize = New System.Drawing.Size(585, 387)
		Me.Location = New System.Drawing.Point(183, 156)
		Me.ForeColor = System.Drawing.Color.Black
		Me.Icon = CType(resources.GetObject("frmAbout.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmAbout"
		Me.txtDescr.AutoSize = False
		Me.txtDescr.BackColor = System.Drawing.Color.FromARGB(0, 51, 153)
		Me.txtDescr.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDescr.ForeColor = System.Drawing.Color.White
		Me.txtDescr.Size = New System.Drawing.Size(529, 235)
		Me.txtDescr.Location = New System.Drawing.Point(24, 56)
		Me.txtDescr.MultiLine = True
		Me.txtDescr.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtDescr.TabIndex = 7
		Me.txtDescr.AcceptsReturn = True
		Me.txtDescr.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDescr.CausesValidation = True
		Me.txtDescr.Enabled = True
		Me.txtDescr.HideSelection = True
		Me.txtDescr.ReadOnly = False
		Me.txtDescr.Maxlength = 0
		Me.txtDescr.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDescr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDescr.TabStop = True
		Me.txtDescr.Visible = True
		Me.txtDescr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDescr.Name = "txtDescr"
		Me._lblURL_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblURL_2.BackColor = System.Drawing.Color.Black
		Me._lblURL_2.Text = "StealthBot Contributors"
		Me._lblURL_2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblURL_2.ForeColor = System.Drawing.Color.White
		Me._lblURL_2.Size = New System.Drawing.Size(145, 17)
		Me._lblURL_2.Location = New System.Drawing.Point(304, 312)
		Me._lblURL_2.TabIndex = 8
		Me._lblURL_2.Enabled = True
		Me._lblURL_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblURL_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblURL_2.UseMnemonic = True
		Me._lblURL_2.Visible = True
		Me._lblURL_2.AutoSize = False
		Me._lblURL_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblURL_2.Name = "_lblURL_2"
		Me.lblSpecialThanks.BackColor = System.Drawing.Color.Black
		Me.lblSpecialThanks.Text = "Special thanks to..."
		Me.lblSpecialThanks.Font = New System.Drawing.Font("Tahoma", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSpecialThanks.ForeColor = System.Drawing.Color.White
		Me.lblSpecialThanks.Size = New System.Drawing.Size(537, 25)
		Me.lblSpecialThanks.Location = New System.Drawing.Point(24, 32)
		Me.lblSpecialThanks.TabIndex = 6
		Me.lblSpecialThanks.UseMnemonic = False
		Me.lblSpecialThanks.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblSpecialThanks.Enabled = True
		Me.lblSpecialThanks.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSpecialThanks.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSpecialThanks.Visible = True
		Me.lblSpecialThanks.AutoSize = False
		Me.lblSpecialThanks.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSpecialThanks.Name = "lblSpecialThanks"
		Me._lblURL_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblURL_3.BackColor = System.Drawing.Color.Black
		Me._lblURL_3.Text = "StealthBot Wiki"
		Me._lblURL_3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblURL_3.ForeColor = System.Drawing.Color.White
		Me._lblURL_3.Size = New System.Drawing.Size(81, 17)
		Me._lblURL_3.Location = New System.Drawing.Point(464, 312)
		Me._lblURL_3.TabIndex = 5
		Me._lblURL_3.Enabled = True
		Me._lblURL_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblURL_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblURL_3.UseMnemonic = True
		Me._lblURL_3.Visible = True
		Me._lblURL_3.AutoSize = False
		Me._lblURL_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblURL_3.Name = "_lblURL_3"
		Me._lblURL_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblURL_1.BackColor = System.Drawing.Color.Black
		Me._lblURL_1.Text = "Send Me E-mail"
		Me._lblURL_1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblURL_1.ForeColor = System.Drawing.Color.White
		Me._lblURL_1.Size = New System.Drawing.Size(89, 17)
		Me._lblURL_1.Location = New System.Drawing.Point(192, 312)
		Me._lblURL_1.TabIndex = 4
		Me._lblURL_1.Enabled = True
		Me._lblURL_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblURL_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblURL_1.UseMnemonic = True
		Me._lblURL_1.Visible = True
		Me._lblURL_1.AutoSize = False
		Me._lblURL_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblURL_1.Name = "_lblURL_1"
		Me._lblURL_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblURL_0.BackColor = System.Drawing.Color.Black
		Me._lblURL_0.Text = "The StealthBot Website and Support Forum"
		Me._lblURL_0.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblURL_0.ForeColor = System.Drawing.Color.White
		Me._lblURL_0.Size = New System.Drawing.Size(129, 33)
		Me._lblURL_0.Location = New System.Drawing.Point(32, 312)
		Me._lblURL_0.TabIndex = 3
		Me._lblURL_0.Enabled = True
		Me._lblURL_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblURL_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblURL_0.UseMnemonic = True
		Me._lblURL_0.Visible = True
		Me._lblURL_0.AutoSize = False
		Me._lblURL_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblURL_0.Name = "_lblURL_0"
		Me.lblOK.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblOK.BackColor = System.Drawing.Color.Black
		Me.lblOK.Text = "[ OK ]"
		Me.lblOK.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOK.ForeColor = System.Drawing.Color.White
		Me.lblOK.Size = New System.Drawing.Size(75, 32)
		Me.lblOK.Location = New System.Drawing.Point(480, 344)
		Me.lblOK.TabIndex = 1
		Me.lblOK.Enabled = True
		Me.lblOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOK.UseMnemonic = True
		Me.lblOK.Visible = True
		Me.lblOK.AutoSize = False
		Me.lblOK.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOK.Name = "lblOK"
		Me._Line1_1.BorderColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_1.X1 = 29
		Me._Line1_1.X2 = 511
		Me._Line1_1.Y1 = 207
		Me._Line1_1.Y2 = 207
		Me._Line1_1.BorderWidth = 1
		Me._Line1_1.Visible = True
		Me._Line1_1.Name = "_Line1_1"
		Me.lblTitle.BackColor = System.Drawing.Color.Black
		Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.ForeColor = System.Drawing.Color.White
		Me.lblTitle.Size = New System.Drawing.Size(533, 32)
		Me.lblTitle.Location = New System.Drawing.Point(24, 8)
		Me.lblTitle.TabIndex = 0
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTitle.Enabled = True
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.UseMnemonic = True
		Me.lblTitle.Visible = True
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me._Line1_0.BorderColor = System.Drawing.Color.White
		Me._Line1_0.BorderWidth = 2
		Me._Line1_0.X1 = 30
		Me._Line1_0.X2 = 511
		Me._Line1_0.Y1 = 208
		Me._Line1_0.Y2 = 208
		Me._Line1_0.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_0.Visible = True
		Me._Line1_0.Name = "_Line1_0"
		Me.lblBottom.BackColor = System.Drawing.Color.Black
		Me.lblBottom.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBottom.ForeColor = System.Drawing.Color.White
		Me.lblBottom.Size = New System.Drawing.Size(457, 33)
		Me.lblBottom.Location = New System.Drawing.Point(24, 352)
		Me.lblBottom.TabIndex = 2
		Me.lblBottom.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBottom.Enabled = True
		Me.lblBottom.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBottom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBottom.UseMnemonic = True
		Me.lblBottom.Visible = True
		Me.lblBottom.AutoSize = False
		Me.lblBottom.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBottom.Name = "lblBottom"
		Me.Controls.Add(txtDescr)
		Me.Controls.Add(_lblURL_2)
		Me.Controls.Add(lblSpecialThanks)
		Me.Controls.Add(_lblURL_3)
		Me.Controls.Add(_lblURL_1)
		Me.Controls.Add(_lblURL_0)
		Me.Controls.Add(lblOK)
		Me.ShapeContainer1.Shapes.Add(_Line1_1)
		Me.Controls.Add(lblTitle)
		Me.ShapeContainer1.Shapes.Add(_Line1_0)
		Me.Controls.Add(lblBottom)
		Me.Controls.Add(ShapeContainer1)
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		Me.Line1.SetIndex(_Line1_0, CType(0, Short))
		Me.lblURL.SetIndex(_lblURL_2, CType(2, Short))
		Me.lblURL.SetIndex(_lblURL_3, CType(3, Short))
		Me.lblURL.SetIndex(_lblURL_1, CType(1, Short))
		Me.lblURL.SetIndex(_lblURL_0, CType(0, Short))
		CType(Me.lblURL, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class