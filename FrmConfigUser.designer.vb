<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConfigUser
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConfigUser))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtusername = New System.Windows.Forms.TextBox
        Me.lblServername = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtconfirm = New System.Windows.Forms.TextBox
        Me.txtpass = New System.Windows.Forms.TextBox
        Me.lblUsername = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblDBname = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnP = New System.Windows.Forms.Button
        Me.btnN = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.txtusername)
        Me.GroupBox1.Controls.Add(Me.lblServername)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtconfirm)
        Me.GroupBox1.Controls.Add(Me.txtpass)
        Me.GroupBox1.Controls.Add(Me.lblUsername)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(12, 34)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(390, 120)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Sing Up"
        '
        'txtusername
        '
        Me.txtusername.Enabled = False
        Me.txtusername.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtusername.Location = New System.Drawing.Point(197, 20)
        Me.txtusername.Name = "txtusername"
        Me.txtusername.Size = New System.Drawing.Size(186, 21)
        Me.txtusername.TabIndex = 0
        '
        'lblServername
        '
        Me.lblServername.AutoSize = True
        Me.lblServername.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServername.Location = New System.Drawing.Point(114, 22)
        Me.lblServername.Name = "lblServername"
        Me.lblServername.Size = New System.Drawing.Size(77, 15)
        Me.lblServername.TabIndex = 8
        Me.lblServername.Text = "User Name :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(122, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 15)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Password :"
        '
        'txtconfirm
        '
        Me.txtconfirm.Enabled = False
        Me.txtconfirm.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtconfirm.Location = New System.Drawing.Point(197, 73)
        Me.txtconfirm.Name = "txtconfirm"
        Me.txtconfirm.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtconfirm.Size = New System.Drawing.Size(186, 21)
        Me.txtconfirm.TabIndex = 2
        '
        'txtpass
        '
        Me.txtpass.Enabled = False
        Me.txtpass.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtpass.Location = New System.Drawing.Point(197, 46)
        Me.txtpass.Name = "txtpass"
        Me.txtpass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtpass.Size = New System.Drawing.Size(186, 21)
        Me.txtpass.TabIndex = 1
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsername.Location = New System.Drawing.Point(75, 75)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(116, 15)
        Me.lblUsername.TabIndex = 13
        Me.lblUsername.Text = "Confirm Password :"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 20)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(69, 70)
        Me.PictureBox1.TabIndex = 14
        Me.PictureBox1.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Image = Global.ProgramVAT.My.Resources.Resources.close__2_
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.Location = New System.Drawing.Point(327, 160)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 31)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "     &Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Image = Global.ProgramVAT.My.Resources.Resources.Save
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.Location = New System.Drawing.Point(237, 160)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 31)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "    &Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'lblDBname
        '
        Me.lblDBname.AutoSize = True
        Me.lblDBname.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBname.Location = New System.Drawing.Point(15, 9)
        Me.lblDBname.Name = "lblDBname"
        Me.lblDBname.Size = New System.Drawing.Size(41, 15)
        Me.lblDBname.TabIndex = 50
        Me.lblDBname.Text = "Name"
        '
        'txtName
        '
        Me.txtName.Enabled = False
        Me.txtName.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(59, 6)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(255, 21)
        Me.txtName.TabIndex = 0
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.Color.Transparent
        Me.btnNew.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Image = Global.ProgramVAT.My.Resources.Resources.bt_nmip_new
        Me.btnNew.Location = New System.Drawing.Point(359, 6)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(21, 21)
        Me.btnNew.TabIndex = 6
        Me.btnNew.UseVisualStyleBackColor = False
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.Transparent
        Me.btnDelete.BackgroundImage = Global.ProgramVAT.My.Resources.Resources.DeleteRed
        Me.btnDelete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(381, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(21, 21)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete Database")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnP
        '
        Me.btnP.Enabled = False
        Me.btnP.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnP.Location = New System.Drawing.Point(320, 6)
        Me.btnP.Name = "btnP"
        Me.btnP.Size = New System.Drawing.Size(19, 21)
        Me.btnP.TabIndex = 4
        Me.btnP.Text = "<"
        Me.btnP.UseVisualStyleBackColor = True
        '
        'btnN
        '
        Me.btnN.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnN.Location = New System.Drawing.Point(338, 6)
        Me.btnN.Name = "btnN"
        Me.btnN.Size = New System.Drawing.Size(19, 21)
        Me.btnN.TabIndex = 5
        Me.btnN.Text = ">"
        Me.btnN.UseVisualStyleBackColor = True
        '
        'FrmConfigUser
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(407, 196)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnN)
        Me.Controls.Add(Me.btnP)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblDBname)
        Me.Controls.Add(Me.txtName)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(423, 234)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(423, 234)
        Me.Name = "FrmConfigUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Config User"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents txtusername As System.Windows.Forms.TextBox
    Friend WithEvents lblServername As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtconfirm As System.Windows.Forms.TextBox
    Friend WithEvents txtpass As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents lblDBname As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnP As System.Windows.Forms.Button
    Friend WithEvents btnN As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
