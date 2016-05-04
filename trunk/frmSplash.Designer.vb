<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSplash
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
	Public WithEvents tmrUnload As System.Windows.Forms.Timer
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Logo As System.Windows.Forms.PictureBox
	Public WithEvents BDay As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSplash))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.tmrUnload = New System.Windows.Forms.Timer(components)
		Me.Label1 = New System.Windows.Forms.Label
		Me.Logo = New System.Windows.Forms.PictureBox
		Me.BDay = New System.Windows.Forms.PictureBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.BackColor = System.Drawing.Color.Black
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(464, 353)
		Me.Location = New System.Drawing.Point(17, 94)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSplash"
		Me.tmrUnload.Interval = 1000
		Me.tmrUnload.Enabled = True
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.BackColor = System.Drawing.Color.Black
		Me.Label1.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.FromARGB(51, 102, 255)
		Me.Label1.Size = New System.Drawing.Size(449, 33)
		Me.Label1.Location = New System.Drawing.Point(8, 312)
		Me.Label1.TabIndex = 0
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Logo.Size = New System.Drawing.Size(450, 300)
		Me.Logo.Location = New System.Drawing.Point(8, 8)
		Me.Logo.Image = CType(resources.GetObject("Logo.Image"), System.Drawing.Image)
		Me.Logo.Enabled = True
		Me.Logo.Cursor = System.Windows.Forms.Cursors.Default
		Me.Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Logo.Visible = True
		Me.Logo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Logo.Name = "Logo"
		Me.BDay.Size = New System.Drawing.Size(300, 316)
		Me.BDay.Location = New System.Drawing.Point(88, 8)
		Me.BDay.Image = CType(resources.GetObject("BDay.Image"), System.Drawing.Image)
		Me.BDay.Visible = False
		Me.BDay.Enabled = True
		Me.BDay.Cursor = System.Windows.Forms.Cursors.Default
		Me.BDay.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.BDay.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.BDay.Name = "BDay"
		Me.Controls.Add(Label1)
		Me.Controls.Add(Logo)
		Me.Controls.Add(BDay)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class