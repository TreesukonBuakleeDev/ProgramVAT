<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOpen
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
	Public WithEvents cboCompany As System.Windows.Forms.ComboBox
    Public WithEvents btnOpen As BaseClass.MBGlassButton
    Public WithEvents lblName As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOpen))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cboCompany = New System.Windows.Forms.ComboBox
        Me.btnOpen = New BaseClass.MBGlassButton
        Me.lblName = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New BaseClass.MBGlassButton
        Me.SuspendLayout()
        '
        'cboCompany
        '
        Me.cboCompany.BackColor = System.Drawing.SystemColors.Window
        Me.cboCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCompany.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCompany.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCompany.Location = New System.Drawing.Point(9, 27)
        Me.cboCompany.Name = "cboCompany"
        Me.cboCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCompany.Size = New System.Drawing.Size(322, 23)
        Me.cboCompany.TabIndex = 0
        '
        'btnOpen
        '
        Me.btnOpen.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnOpen.BackColor = System.Drawing.SystemColors.Control
        Me.btnOpen.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnOpen.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOpen.FlatAppearance.BorderSize = 0
        Me.btnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpen.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOpen.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnOpen.Location = New System.Drawing.Point(77, 57)
        Me.btnOpen.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnOpen.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnOpen.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnOpen.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOpen.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnOpen.Size = New System.Drawing.Size(88, 32)
        Me.btnOpen.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnOpen.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnOpen.TabIndex = 1
        Me.btnOpen.Text = "      Open"
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.SystemColors.Control
        Me.lblName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblName.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblName.Location = New System.Drawing.Point(16, 72)
        Me.lblName.Name = "lblName"
        Me.lblName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblName.Size = New System.Drawing.Size(33, 17)
        Me.lblName.TabIndex = 3
        Me.lblName.Text = "Label2"
        Me.lblName.Visible = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(9, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(322, 19)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Company list"
        '
        'btnClose
        '
        Me.btnClose.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnClose.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnClose.FlatAppearance.BorderSize = 0
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnClose.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnClose.Image = Global.ProgramVAT.My.Resources.Resources.close__2_
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Location = New System.Drawing.Point(181, 57)
        Me.btnClose.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnClose.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnClose.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnClose.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnClose.Size = New System.Drawing.Size(88, 32)
        Me.btnClose.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnClose.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "      Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'frmOpen
        '
        Me.AcceptButton = Me.btnOpen
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(343, 100)
        Me.Controls.Add(Me.cboCompany)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmOpen"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmOpen"
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents btnClose As BaseClass.MBGlassButton
#End Region
End Class