<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCompany
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
	Public WithEvents txtTax As System.Windows.Forms.TextBox
    Public WithEvents cmdCancel As BaseClass.MBGlassButton
    Public WithEvents cmdSave As BaseClass.MBGlassButton
    Public WithEvents _lblAddress_0 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_1 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_2 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_3 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_4 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_5 As System.Windows.Forms.Label
    Public WithEvents _lblAddress_6 As System.Windows.Forms.Label
    Public WithEvents lblTax As System.Windows.Forms.Label
    Public WithEvents lblAddress As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCompany))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.cmdCancel = New BaseClass.MBGlassButton
        Me.cmdSave = New BaseClass.MBGlassButton
        Me._lblAddress_0 = New System.Windows.Forms.Label
        Me._lblAddress_1 = New System.Windows.Forms.Label
        Me._lblAddress_2 = New System.Windows.Forms.Label
        Me._lblAddress_3 = New System.Windows.Forms.Label
        Me._lblAddress_4 = New System.Windows.Forms.Label
        Me._lblAddress_5 = New System.Windows.Forms.Label
        Me._lblAddress_6 = New System.Windows.Forms.Label
        Me.lblTax = New System.Windows.Forms.Label
        Me.lblAddress = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        CType(Me.lblAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtTax
        '
        Me.txtTax.AcceptsReturn = True
        Me.txtTax.BackColor = System.Drawing.SystemColors.Window
        Me.txtTax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTax.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTax.Location = New System.Drawing.Point(107, 201)
        Me.txtTax.MaxLength = 10
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTax.Size = New System.Drawing.Size(191, 21)
        Me.txtTax.TabIndex = 0
        Me.txtTax.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdCancel.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdCancel.Image = Global.ProgramVAT.My.Resources.Resources.close__2_
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdCancel.Location = New System.Drawing.Point(217, 15)
        Me.cmdCancel.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmdCancel.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmdCancel.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdCancel.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmdCancel.Size = New System.Drawing.Size(90, 28)
        Me.cmdCancel.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmdCancel.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmdCancel.TabIndex = 2
        Me.cmdCancel.Text = "       Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        Me.cmdCancel.Visible = False
        '
        'cmdSave
        '
        Me.cmdSave.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdSave.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.FlatAppearance.BorderSize = 0
        Me.cmdSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdSave.Image = Global.ProgramVAT.My.Resources.Resources.Save
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdSave.Location = New System.Drawing.Point(121, 15)
        Me.cmdSave.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmdSave.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmdSave.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdSave.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmdSave.Size = New System.Drawing.Size(90, 28)
        Me.cmdSave.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmdSave.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmdSave.TabIndex = 1
        Me.cmdSave.Text = "       Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        Me.cmdSave.Visible = False
        '
        '_lblAddress_0
        '
        Me._lblAddress_0.AutoSize = True
        Me._lblAddress_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_0.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_0, CType(0, Short))
        Me._lblAddress_0.Location = New System.Drawing.Point(15, 15)
        Me._lblAddress_0.Name = "_lblAddress_0"
        Me._lblAddress_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_0.Size = New System.Drawing.Size(44, 15)
        Me._lblAddress_0.TabIndex = 10
        Me._lblAddress_0.Text = "Name "
        '
        '_lblAddress_1
        '
        Me._lblAddress_1.AutoSize = True
        Me._lblAddress_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_1, CType(1, Short))
        Me._lblAddress_1.Location = New System.Drawing.Point(15, 42)
        Me._lblAddress_1.Name = "_lblAddress_1"
        Me._lblAddress_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_1.Size = New System.Drawing.Size(60, 15)
        Me._lblAddress_1.TabIndex = 9
        Me._lblAddress_1.Text = "Address1"
        '
        '_lblAddress_2
        '
        Me._lblAddress_2.AutoSize = True
        Me._lblAddress_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_2, CType(2, Short))
        Me._lblAddress_2.Location = New System.Drawing.Point(15, 68)
        Me._lblAddress_2.Name = "_lblAddress_2"
        Me._lblAddress_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_2.Size = New System.Drawing.Size(60, 15)
        Me._lblAddress_2.TabIndex = 8
        Me._lblAddress_2.Text = "Address2"
        '
        '_lblAddress_3
        '
        Me._lblAddress_3.AutoSize = True
        Me._lblAddress_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_3, CType(3, Short))
        Me._lblAddress_3.Location = New System.Drawing.Point(15, 95)
        Me._lblAddress_3.Name = "_lblAddress_3"
        Me._lblAddress_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_3.Size = New System.Drawing.Size(60, 15)
        Me._lblAddress_3.TabIndex = 7
        Me._lblAddress_3.Text = "Address3"
        '
        '_lblAddress_4
        '
        Me._lblAddress_4.AutoSize = True
        Me._lblAddress_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_4, CType(4, Short))
        Me._lblAddress_4.Location = New System.Drawing.Point(15, 121)
        Me._lblAddress_4.Name = "_lblAddress_4"
        Me._lblAddress_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_4.Size = New System.Drawing.Size(60, 15)
        Me._lblAddress_4.TabIndex = 6
        Me._lblAddress_4.Text = "Address4"
        '
        '_lblAddress_5
        '
        Me._lblAddress_5.AutoSize = True
        Me._lblAddress_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_5, CType(5, Short))
        Me._lblAddress_5.Location = New System.Drawing.Point(15, 148)
        Me._lblAddress_5.Name = "_lblAddress_5"
        Me._lblAddress_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_5.Size = New System.Drawing.Size(84, 15)
        Me._lblAddress_5.TabIndex = 5
        Me._lblAddress_5.Text = "City - Province"
        '
        '_lblAddress_6
        '
        Me._lblAddress_6.AutoSize = True
        Me._lblAddress_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblAddress_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblAddress_6.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblAddress_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAddress.SetIndex(Me._lblAddress_6, CType(6, Short))
        Me._lblAddress_6.Location = New System.Drawing.Point(15, 174)
        Me._lblAddress_6.Name = "_lblAddress_6"
        Me._lblAddress_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblAddress_6.Size = New System.Drawing.Size(96, 15)
        Me._lblAddress_6.TabIndex = 4
        Me._lblAddress_6.Text = "Contry - Zipcode"
        '
        'lblTax
        '
        Me.lblTax.AutoSize = True
        Me.lblTax.BackColor = System.Drawing.SystemColors.Control
        Me.lblTax.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTax.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTax.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTax.Location = New System.Drawing.Point(15, 201)
        Me.lblTax.Name = "lblTax"
        Me.lblTax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTax.Size = New System.Drawing.Size(73, 15)
        Me.lblTax.TabIndex = 3
        Me.lblTax.Text = "Tax Number"
        '
        'frmCompany
        '
        Me.AcceptButton = Me.cmdCancel
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(312, 237)
        Me.Controls.Add(Me.txtTax)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me._lblAddress_0)
        Me.Controls.Add(Me._lblAddress_1)
        Me.Controls.Add(Me._lblAddress_2)
        Me.Controls.Add(Me._lblAddress_3)
        Me.Controls.Add(Me._lblAddress_4)
        Me.Controls.Add(Me._lblAddress_5)
        Me.Controls.Add(Me._lblAddress_6)
        Me.Controls.Add(Me.lblTax)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCompany"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Form1"
        CType(Me.lblAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class