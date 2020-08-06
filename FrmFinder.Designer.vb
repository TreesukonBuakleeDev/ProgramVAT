<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmFinder
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmFinder))
        Me.btnFind = New BaseClass.MBGlassButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnFind
        '
        Me.btnFind.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnFind.BackColor = System.Drawing.Color.Transparent
        Me.btnFind.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnFind.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnFind.FlatAppearance.BorderSize = 0
        Me.btnFind.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFind.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFind.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnFind.Image = CType(resources.GetObject("btnFind.Image"), System.Drawing.Image)
        Me.btnFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFind.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnFind.Location = New System.Drawing.Point(446, 18)
        Me.btnFind.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnFind.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnFind.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnFind.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnFind.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnFind.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnFind.TabIndex = 1
        Me.btnFind.Text = "     Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Find Text"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(73, 18)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(367, 26)
        Me.TextBox1.TabIndex = 0
        '
        'FrmFinder
        '
        Me.AcceptButton = Me.btnFind
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(538, 69)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnFind)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(554, 107)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(554, 107)
        Me.Name = "FrmFinder"
        Me.Text = "Finder..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFind As BaseClass.MBGlassButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
