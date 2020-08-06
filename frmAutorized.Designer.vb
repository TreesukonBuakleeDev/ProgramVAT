<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAutorized
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
    Public WithEvents cmdCancel As BaseClass.MBGlassButton
    Public WithEvents cmdOpen As BaseClass.MBGlassButton
    Public WithEvents txtPass As System.Windows.Forms.TextBox
    Public WithEvents txtUser As System.Windows.Forms.TextBox
    Public WithEvents lblRequest As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAutorized))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New BaseClass.MBGlassButton
        Me.cmdOpen = New BaseClass.MBGlassButton
        Me.txtPass = New System.Windows.Forms.TextBox
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.lblRequest = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdCancel.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdCancel.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.FlatAppearance.BorderSize = 0
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdCancel.Location = New System.Drawing.Point(163, 108)
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
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "   Clear"
        Me.cmdCancel.UseVisualStyleBackColor = False
        Me.cmdCancel.UseWaitCursor = True
        Me.cmdCancel.Visible = False
        '
        'cmdOpen
        '
        Me.cmdOpen.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmdOpen.BackColor = System.Drawing.Color.LightGreen
        Me.cmdOpen.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmdOpen.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOpen.FlatAppearance.BorderSize = 0
        Me.cmdOpen.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdOpen.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOpen.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmdOpen.Image = CType(resources.GetObject("cmdOpen.Image"), System.Drawing.Image)
        Me.cmdOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOpen.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmdOpen.Location = New System.Drawing.Point(54, 108)
        Me.cmdOpen.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmdOpen.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmdOpen.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdOpen.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOpen.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmdOpen.Size = New System.Drawing.Size(88, 28)
        Me.cmdOpen.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmdOpen.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmdOpen.TabIndex = 2
        Me.cmdOpen.Text = "   OK"
        Me.cmdOpen.UseVisualStyleBackColor = False
        '
        'txtPass
        '
        Me.txtPass.AcceptsReturn = True
        Me.txtPass.BackColor = System.Drawing.SystemColors.Window
        Me.txtPass.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPass.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPass.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPass.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtPass.Location = New System.Drawing.Point(86, 72)
        Me.txtPass.MaxLength = 0
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPass.Size = New System.Drawing.Size(152, 21)
        Me.txtPass.TabIndex = 1
        '
        'txtUser
        '
        Me.txtUser.AcceptsReturn = True
        Me.txtUser.BackColor = System.Drawing.SystemColors.Window
        Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUser.Location = New System.Drawing.Point(86, 39)
        Me.txtUser.MaxLength = 0
        Me.txtUser.Name = "txtUser"
        Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUser.Size = New System.Drawing.Size(152, 21)
        Me.txtUser.TabIndex = 0
        '
        'lblRequest
        '
        Me.lblRequest.BackColor = System.Drawing.Color.Transparent
        Me.lblRequest.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRequest.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRequest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRequest.Location = New System.Drawing.Point(175, 9)
        Me.lblRequest.Name = "lblRequest"
        Me.lblRequest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRequest.Size = New System.Drawing.Size(103, 19)
        Me.lblRequest.TabIndex = 7
        Me.lblRequest.Text = "Request"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(157, 19)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Request Autorized for "
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(39, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(41, 19)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "User :"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(12, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(77, 19)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Password :"
        '
        'frmAutorized
        '
        Me.AcceptButton = Me.cmdOpen
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(288, 145)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOpen)
        Me.Controls.Add(Me.txtPass)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.lblRequest)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Leelawadee", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAutorized"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class