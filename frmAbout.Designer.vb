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
    Public WithEvents btnClose As BaseClass.MBGlassButton
    Public WithEvents Image1 As System.Windows.Forms.PictureBox
    Public WithEvents _Label2_1 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblCompany As System.Windows.Forms.Label
    Public WithEvents _Line1_1 As System.Windows.Forms.Label
    Public WithEvents lblDescription As System.Windows.Forms.Label
    Public WithEvents lblTitle As System.Windows.Forms.Label
    Public WithEvents _Line1_0 As System.Windows.Forms.Label
    Public WithEvents lblVersion As System.Windows.Forms.Label
    Public WithEvents lblDisclaimer As System.Windows.Forms.Label
    Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents Line1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAbout))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnClose = New BaseClass.MBGlassButton
        Me.Image1 = New System.Windows.Forms.PictureBox
        Me._Label2_1 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblCompany = New System.Windows.Forms.Label
        Me._Line1_1 = New System.Windows.Forms.Label
        Me.lblDescription = New System.Windows.Forms.Label
        Me.lblTitle = New System.Windows.Forms.Label
        Me._Line1_0 = New System.Windows.Forms.Label
        Me.lblVersion = New System.Windows.Forms.Label
        Me.lblDisclaimer = New System.Windows.Forms.Label
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Line1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClose
        '
        Me.btnClose.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.btnClose.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnClose.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.FlatAppearance.BorderSize = 0
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnClose.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnClose.Image = Global.ProgramVAT.My.Resources.Resources.close__2_
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnClose.Location = New System.Drawing.Point(264, 154)
        Me.btnClose.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnClose.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnClose.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnClose.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnClose.Size = New System.Drawing.Size(88, 28)
        Me.btnClose.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnClose.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "       Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'Image1
        '
        Me.Image1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Image1.BackgroundImage = CType(resources.GetObject("Image1.BackgroundImage"), System.Drawing.Image)
        Me.Image1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Image1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Image1.Location = New System.Drawing.Point(3, 11)
        Me.Image1.Name = "Image1"
        Me.Image1.Size = New System.Drawing.Size(117, 64)
        Me.Image1.TabIndex = 7
        Me.Image1.TabStop = False
        '
        '_Label2_1
        '
        Me._Label2_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label2_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label2_1.ForeColor = System.Drawing.Color.Black
        Me.Label2.SetIndex(Me._Label2_1, CType(1, Short))
        Me._Label2_1.Location = New System.Drawing.Point(123, 27)
        Me._Label2_1.Name = "_Label2_1"
        Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label2_1.Size = New System.Drawing.Size(55, 15)
        Me._Label2_1.TabIndex = 7
        Me._Label2_1.Text = "Version"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(13, 155)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(245, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Contact :"
        '
        'lblCompany
        '
        Me.lblCompany.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lblCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompany.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.ForeColor = System.Drawing.Color.Black
        Me.lblCompany.Location = New System.Drawing.Point(13, 170)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompany.Size = New System.Drawing.Size(245, 18)
        Me.lblCompany.TabIndex = 4
        Me.lblCompany.Text = "Application Title"
        '
        '_Line1_1
        '
        Me._Line1_1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me._Line1_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Line1.SetIndex(Me._Line1_1, CType(1, Short))
        Me._Line1_1.Location = New System.Drawing.Point(13, 154)
        Me._Line1_1.Name = "_Line1_1"
        Me._Line1_1.Size = New System.Drawing.Size(320, 1)
        Me._Line1_1.TabIndex = 8
        '
        'lblDescription
        '
        Me.lblDescription.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDescription.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.ForeColor = System.Drawing.Color.Black
        Me.lblDescription.Location = New System.Drawing.Point(126, 42)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDescription.Size = New System.Drawing.Size(226, 110)
        Me.lblDescription.TabIndex = 0
        Me.lblDescription.Text = "App Description"
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Black
        Me.lblTitle.Location = New System.Drawing.Point(123, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(229, 15)
        Me.lblTitle.TabIndex = 2
        Me.lblTitle.Text = "Application Title"
        '
        '_Line1_0
        '
        Me._Line1_0.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me._Line1_0.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Line1.SetIndex(Me._Line1_0, CType(0, Short))
        Me._Line1_0.Location = New System.Drawing.Point(13, 155)
        Me._Line1_0.Name = "_Line1_0"
        Me._Line1_0.Size = New System.Drawing.Size(296, 1)
        Me._Line1_0.TabIndex = 9
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVersion.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVersion.Location = New System.Drawing.Point(184, 27)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVersion.Size = New System.Drawing.Size(168, 24)
        Me.lblVersion.TabIndex = 3
        Me.lblVersion.Text = "Version"
        '
        'lblDisclaimer
        '
        Me.lblDisclaimer.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDisclaimer.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
        Me.lblDisclaimer.Location = New System.Drawing.Point(12, 189)
        Me.lblDisclaimer.Name = "lblDisclaimer"
        Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDisclaimer.Size = New System.Drawing.Size(340, 112)
        Me.lblDisclaimer.TabIndex = 1
        Me.lblDisclaimer.Text = "Warning: ......."
        '
        'frmAbout
        '
        Me.AcceptButton = Me.btnClose
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(360, 319)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Image1)
        Me.Controls.Add(Me._Label2_1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblCompany)
        Me.Controls.Add(Me._Line1_1)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me._Line1_0)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.lblDisclaimer)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(156, 129)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAbout"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "About MyApp"
        CType(Me.Image1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class