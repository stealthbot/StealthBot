<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWhisperWindow
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
	Public WithEvents mnuSave As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep2 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuClose As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuIgnoreAndClose As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuSep As System.Windows.Forms.ToolStripSeparator
	Public WithEvents mnuHide As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public cdlSave As System.Windows.Forms.SaveFileDialog
	Public WithEvents txtSend As System.Windows.Forms.TextBox
	Public WithEvents rtbWhispers As System.Windows.Forms.RichTextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWhisperWindow))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSave = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep2 = New System.Windows.Forms.ToolStripSeparator
		Me.mnuClose = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuIgnoreAndClose = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuSep = New System.Windows.Forms.ToolStripSeparator
		Me.mnuHide = New System.Windows.Forms.ToolStripMenuItem
		Me.cdlSave = New System.Windows.Forms.SaveFileDialog
		Me.txtSend = New System.Windows.Forms.TextBox
		Me.rtbWhispers = New System.Windows.Forms.RichTextBox
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.BackColor = System.Drawing.Color.Black
		Me.Text = "< account name >"
		Me.ClientSize = New System.Drawing.Size(313, 242)
		Me.Location = New System.Drawing.Point(11, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmWhisperWindow"
		Me.mnuOptions.Name = "mnuOptions"
		Me.mnuOptions.Text = "&Options"
		Me.mnuOptions.Checked = False
		Me.mnuOptions.Enabled = True
		Me.mnuOptions.Visible = True
		Me.mnuSave.Name = "mnuSave"
		Me.mnuSave.Text = "&Save Conversation"
		Me.mnuSave.Checked = False
		Me.mnuSave.Enabled = True
		Me.mnuSave.Visible = True
		Me.mnuSep2.Enabled = True
		Me.mnuSep2.Visible = True
		Me.mnuSep2.Name = "mnuSep2"
		Me.mnuClose.Name = "mnuClose"
		Me.mnuClose.Text = "&Close"
		Me.mnuClose.Checked = False
		Me.mnuClose.Enabled = True
		Me.mnuClose.Visible = True
		Me.mnuIgnoreAndClose.Name = "mnuIgnoreAndClose"
		Me.mnuIgnoreAndClose.Text = "&Ignore and Close"
		Me.mnuIgnoreAndClose.Checked = False
		Me.mnuIgnoreAndClose.Enabled = True
		Me.mnuIgnoreAndClose.Visible = True
		Me.mnuSep.Enabled = True
		Me.mnuSep.Visible = True
		Me.mnuSep.Name = "mnuSep"
		Me.mnuHide.Name = "mnuHide"
		Me.mnuHide.Text = "&Hide"
		Me.mnuHide.Checked = False
		Me.mnuHide.Enabled = True
		Me.mnuHide.Visible = True
		Me.txtSend.AutoSize = False
		Me.txtSend.BackColor = System.Drawing.Color.Black
		Me.txtSend.ForeColor = System.Drawing.Color.White
		Me.txtSend.Size = New System.Drawing.Size(297, 19)
		Me.txtSend.Location = New System.Drawing.Point(8, 216)
		Me.txtSend.TabIndex = 1
		Me.txtSend.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSend.AcceptsReturn = True
		Me.txtSend.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSend.CausesValidation = True
		Me.txtSend.Enabled = True
		Me.txtSend.HideSelection = True
		Me.txtSend.ReadOnly = False
		Me.txtSend.Maxlength = 0
		Me.txtSend.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSend.MultiLine = False
		Me.txtSend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSend.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSend.TabStop = True
		Me.txtSend.Visible = True
		Me.txtSend.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSend.Name = "txtSend"
		Me.rtbWhispers.Size = New System.Drawing.Size(297, 177)
		Me.rtbWhispers.Location = New System.Drawing.Point(8, 32)
		Me.rtbWhispers.TabIndex = 0
		Me.rtbWhispers.BackColor = System.Drawing.Color.Black
		Me.rtbWhispers.Enabled = True
		Me.rtbWhispers.ReadOnly = True
		Me.rtbWhispers.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rtbWhispers.RTF = resources.GetString("rtbWhispers.TextRTF")
		Me.rtbWhispers.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbWhispers.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rtbWhispers.Name = "rtbWhispers"
		Me.Controls.Add(txtSend)
		Me.Controls.Add(rtbWhispers)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuOptions})
		mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuSave, Me.mnuSep2, Me.mnuClose, Me.mnuIgnoreAndClose, Me.mnuSep, Me.mnuHide})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class