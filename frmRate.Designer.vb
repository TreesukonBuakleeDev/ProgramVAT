<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmRate
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
	Public WithEvents txtRate As System.Windows.Forms.TextBox
	Public WithEvents txtTo As System.Windows.Forms.TextBox
	Public WithEvents txtFrom As System.Windows.Forms.TextBox
    Public WithEvents cmdUpdate As BaseClass.MBGlassButton
    Public WithEvents cmdCancel As BaseClass.MBGlassButton
    Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_7 As System.Windows.Forms.Label
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRate))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.txtTo = New System.Windows.Forms.TextBox
        Me.txtFrom = New System.Windows.Forms.TextBox
        Me.cmdUpdate = New BaseClass.MBGlassButton
        Me.cmdCancel = New BaseClass.MBGlassButton
        Me._lblLabels_1 = New System.Windows.Forms.Label
        Me._lblLabels_0 = New System.Windows.Forms.Label
        Me._lblLabels_7 = New System.Windows.Forms.Label
        Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtRate
        '
        Me.txtRate.AcceptsReturn = True
        Me.txtRate.BackColor = System.Drawing.SystemColors.Window
        Me.txtRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRate.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtRate.Location = New System.Drawing.Point(103, 54)
        Me.txtRate.MaxLength = 0
        Me.txtRate.Name = "txtRate"
        Me.txtRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRate.Size = New System.Drawing.Size(61, 21)
        Me.txtRate.TabIndex = 2
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTo
        '
        Me.txtTo.AcceptsReturn = True
        Me.txtTo.BackColor = System.Drawing.SystemColors.Window
        Me.txtTo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtTo.Location = New System.Drawing.Point(215, 20)
        Me.txtTo.MaxLength = 0
        Me.txtTo.Name = "txtTo"
        Me.txtTo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTo.Size = New System.Drawing.Size(62, 21)
        Me.txtTo.TabIndex = 1
        Me.txtTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtFrom
        '
        Me.txtFrom.AcceptsReturn = True
        Me.txtFrom.BackColor = System.Drawing.SystemColors.Window
        Me.txtFrom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFrom.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFrom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFrom.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtFrom.Location = New System.Drawing.Point(102, 21)
        Me.txtFrom.MaxLength = 0
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFrom.Size = New System.Drawing.Size(62, 21)
        Me.txtFrom.TabIndex = 0
        Me.txtFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdUpdate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdUpdate.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdUpdate.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdUpdate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdUpdate.FlatAppearance.BorderSize = 0
        Me.cmdUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdUpdate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdUpdate.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdUpdate.Image = Global.ProgramVAT.My.Resources.Resources.edit_validated_icon
        Me.cmdUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUpdate.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdUpdate.Location = New System.Drawing.Point(60, 90)
        Me.cmdUpdate.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmdUpdate.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmdUpdate.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdUpdate.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdUpdate.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmdUpdate.Size = New System.Drawing.Size(88, 28)
        Me.cmdUpdate.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmdUpdate.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmdUpdate.TabIndex = 3
        Me.cmdUpdate.Text = "       Update"
        Me.cmdUpdate.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdCancel.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdCancel.Image = Global.ProgramVAT.My.Resources.Resources.close__2_
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdCancel.Location = New System.Drawing.Point(166, 90)
        Me.cmdCancel.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmdCancel.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmdCancel.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdCancel.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmdCancel.Size = New System.Drawing.Size(88, 28)
        Me.cmdCancel.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmdCancel.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "       Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(3, 54)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(90, 18)
        Me._lblLabels_1.TabIndex = 7
        Me._lblLabels_1.Text = "Change To"
        Me._lblLabels_1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(184, 23)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(25, 18)
        Me._lblLabels_0.TabIndex = 5
        Me._lblLabels_0.Text = "To"
        '
        '_lblLabels_7
        '
        Me._lblLabels_7.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_7.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_7, CType(7, Short))
        Me._lblLabels_7.Location = New System.Drawing.Point(21, 21)
        Me._lblLabels_7.Name = "_lblLabels_7"
        Me._lblLabels_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_7.Size = New System.Drawing.Size(72, 18)
        Me._lblLabels_7.TabIndex = 3
        Me._lblLabels_7.Text = "From Rate"
        Me._lblLabels_7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmRate
        '
        Me.AcceptButton = Me.cmdCancel
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(304, 138)
        Me.Controls.Add(Me.txtRate)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me._lblLabels_1)
        Me.Controls.Add(Me._lblLabels_0)
        Me.Controls.Add(Me._lblLabels_7)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRate"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmRate"
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class