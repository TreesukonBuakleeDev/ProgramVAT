<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLocation
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLocation))
        Me.txtComment = New System.Windows.Forms.TextBox
        Me.label5 = New System.Windows.Forms.Label
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.label4 = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.label3 = New System.Windows.Forms.Label
        Me.txtPrefix = New System.Windows.Forms.TextBox
        Me.label2 = New System.Windows.Forms.Label
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.label1 = New System.Windows.Forms.Label
        Me.btnDelete = New BaseClass.MBGlassButton
        Me.btnUpdate = New BaseClass.MBGlassButton
        Me.dataGridView1 = New BaseClass.BaseGridview
        Me.Type = New BaseClass.DataGridViewTextImageColumnMergeCell
        Me.AccountCode = New BaseClass.DataGridViewTextImageColumnMergeCell
        Me.btnNew = New BaseClass.MBGlassButton
        Me.cmbN = New BaseClass.MBGlassButton
        Me.cmbP = New BaseClass.MBGlassButton
        CType(Me.dataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(81, 87)
        Me.txtComment.MaxLength = 100
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(320, 21)
        Me.txtComment.TabIndex = 4
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Location = New System.Drawing.Point(13, 91)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(62, 15)
        Me.label5.TabIndex = 29
        Me.label5.Text = "Comment"
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(81, 60)
        Me.txtAddress.MaxLength = 200
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(320, 21)
        Me.txtAddress.TabIndex = 3
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(13, 64)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(53, 15)
        Me.label4.TabIndex = 27
        Me.label4.Text = "Address"
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(81, 33)
        Me.txtName.MaxLength = 50
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(320, 21)
        Me.txtName.TabIndex = 2
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(13, 37)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(41, 15)
        Me.label3.TabIndex = 25
        Me.label3.Text = "Name"
        '
        'txtPrefix
        '
        Me.txtPrefix.Location = New System.Drawing.Point(290, 5)
        Me.txtPrefix.MaxLength = 4
        Me.txtPrefix.Name = "txtPrefix"
        Me.txtPrefix.Size = New System.Drawing.Size(111, 21)
        Me.txtPrefix.TabIndex = 1
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(247, 9)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(37, 15)
        Me.label2.TabIndex = 23
        Me.label2.Text = "Prefix"
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(111, 6)
        Me.txtLocation.MaxLength = 6
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(100, 21)
        Me.txtLocation.TabIndex = 0
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(13, 9)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(54, 15)
        Me.label1.TabIndex = 20
        Me.label1.Text = "Location"
        '
        'btnDelete
        '
        Me.btnDelete.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnDelete.BackColor = System.Drawing.Color.Transparent
        Me.btnDelete.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnDelete.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnDelete.FlatAppearance.BorderSize = 0
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnDelete.Image = Global.ProgramVAT.My.Resources.Resources.DeleteRed
        Me.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDelete.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnDelete.Location = New System.Drawing.Point(271, 250)
        Me.btnDelete.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnDelete.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnDelete.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDelete.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDelete.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnDelete.Size = New System.Drawing.Size(100, 29)
        Me.btnDelete.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnDelete.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnDelete.TabIndex = 8
        Me.btnDelete.Text = "    Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnUpdate.BackColor = System.Drawing.Color.Transparent
        Me.btnUpdate.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnUpdate.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnUpdate.FlatAppearance.BorderSize = 0
        Me.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpdate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdate.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnUpdate.Image = Global.ProgramVAT.My.Resources.Resources.edit_validated_icon
        Me.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUpdate.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnUpdate.Location = New System.Drawing.Point(156, 250)
        Me.btnUpdate.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnUpdate.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnUpdate.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUpdate.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUpdate.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnUpdate.Size = New System.Drawing.Size(100, 29)
        Me.btnUpdate.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnUpdate.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnUpdate.TabIndex = 6
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnUpdate.UseVisualStyleBackColor = False
        '
        'dataGridView1
        '
        Me.dataGridView1.AllowUserToAddRows = False
        Me.dataGridView1.AllowUserToDeleteRows = False
        Me.dataGridView1.AllowUserToResizeRows = False
        Me.dataGridView1.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.dataGridView1.ColIndex = -1
        Me.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Type, Me.AccountCode})
        Me.dataGridView1.DGEditMode = True
        Me.dataGridView1.Location = New System.Drawing.Point(13, 114)
        Me.dataGridView1.Name = "dataGridView1"
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(51, Byte), Integer), CType(CType(153, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.dataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dataGridView1.ShowCellToolTips = False
        Me.dataGridView1.Size = New System.Drawing.Size(389, 130)
        Me.dataGridView1.TabIndex = 5
        Me.dataGridView1.UseRowColor = True
        Me.dataGridView1.UseRowCount = False
        Me.dataGridView1.UseVisualStyleRendererCell = True
        Me.dataGridView1.UseVisualStyleRendererRow = True
        '
        'Type
        '
        Me.Type.HeaderText = "Type"
        Me.Type.Name = "Type"
        Me.Type.ReadOnly = True
        Me.Type.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'AccountCode
        '
        Me.AccountCode.HeaderText = "AccountCode"
        Me.AccountCode.Name = "AccountCode"
        Me.AccountCode.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'btnNew
        '
        Me.btnNew.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnNew.BackColor = System.Drawing.Color.Transparent
        Me.btnNew.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnNew.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnNew.FlatAppearance.BorderSize = 0
        Me.btnNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNew.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnNew.Image = Global.ProgramVAT.My.Resources.Resources.file_new
        Me.btnNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnNew.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnNew.Location = New System.Drawing.Point(41, 250)
        Me.btnNew.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnNew.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnNew.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnNew.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnNew.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnNew.Size = New System.Drawing.Size(100, 29)
        Me.btnNew.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnNew.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnNew.TabIndex = 7
        Me.btnNew.Text = "   New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'cmbN
        '
        Me.cmbN.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmbN.BackColor = System.Drawing.Color.Transparent
        Me.cmbN.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmbN.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmbN.FlatAppearance.BorderSize = 0
        Me.cmbN.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbN.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbN.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmbN.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmbN.Location = New System.Drawing.Point(217, 5)
        Me.cmbN.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmbN.Name = "cmbN"
        Me.cmbN.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmbN.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmbN.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmbN.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmbN.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmbN.Size = New System.Drawing.Size(24, 23)
        Me.cmbN.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmbN.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmbN.TabIndex = 9
        Me.cmbN.Text = ">"
        Me.cmbN.UseVisualStyleBackColor = True
        '
        'cmbP
        '
        Me.cmbP.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.cmbP.BackColor = System.Drawing.Color.Transparent
        Me.cmbP.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.cmbP.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmbP.FlatAppearance.BorderSize = 0
        Me.cmbP.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmbP.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbP.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.cmbP.ImageSize = New System.Drawing.Size(24, 24)
        Me.cmbP.Location = New System.Drawing.Point(81, 5)
        Me.cmbP.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.cmbP.Name = "cmbP"
        Me.cmbP.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.cmbP.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.cmbP.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmbP.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmbP.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.cmbP.Size = New System.Drawing.Size(24, 23)
        Me.cmbP.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.cmbP.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.cmbP.TabIndex = 10
        Me.cmbP.Text = "<"
        Me.cmbP.UseVisualStyleBackColor = True
        '
        'FrmLocation
        '
        Me.AcceptButton = Me.btnUpdate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(414, 288)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.dataGridView1)
        Me.Controls.Add(Me.txtComment)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.txtPrefix)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.txtLocation)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.cmbN)
        Me.Controls.Add(Me.cmbP)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(420, 316)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(420, 316)
        Me.Name = "FrmLocation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Location"
        CType(Me.dataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents btnDelete As BaseClass.MBGlassButton
    Private WithEvents btnUpdate As BaseClass.MBGlassButton
    Private WithEvents dataGridView1 As BaseClass.BaseGridview 'System.Windows.Forms.DataGridView
    Private WithEvents txtComment As System.Windows.Forms.TextBox
    Private WithEvents label5 As System.Windows.Forms.Label
    Private WithEvents txtAddress As System.Windows.Forms.TextBox
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents txtName As System.Windows.Forms.TextBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents txtPrefix As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents txtLocation As System.Windows.Forms.TextBox
    Private WithEvents btnNew As BaseClass.MBGlassButton
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents cmbN As BaseClass.MBGlassButton
    Private WithEvents cmbP As BaseClass.MBGlassButton
    Friend WithEvents Type As BaseClass.DataGridViewTextImageColumnMergeCell
    Friend WithEvents AccountCode As BaseClass.DataGridViewTextImageColumnMergeCell
End Class
