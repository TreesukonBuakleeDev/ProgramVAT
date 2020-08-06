<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DeleteVat
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DeleteVat))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Button1 = New System.Windows.Forms.Button
        Me.CMYear = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lstVu = New BaseClass.BaseGridview
        Me.ID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Month = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Total_Entry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Action = New System.Windows.Forms.DataGridViewLinkColumn
        Me.MonthNum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Button2 = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.lstVu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(345, 10)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 25)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "   Select"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'CMYear
        '
        Me.CMYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CMYear.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CMYear.FormattingEnabled = True
        Me.CMYear.Location = New System.Drawing.Point(85, 12)
        Me.CMYear.Name = "CMYear"
        Me.CMYear.Size = New System.Drawing.Size(242, 23)
        Me.CMYear.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(49, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 15)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Year"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.CMYear)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(425, 44)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel2.Controls.Add(Me.lstVu)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 44)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Padding = New System.Windows.Forms.Padding(5)
        Me.Panel2.Size = New System.Drawing.Size(425, 253)
        Me.Panel2.TabIndex = 2
        '
        'lstVu
        '
        Me.lstVu.AllowUserToAddRows = False
        Me.lstVu.AllowUserToDeleteRows = False
        Me.lstVu.AllowUserToOrderColumns = True
        Me.lstVu.AllowUserToResizeRows = False
        Me.lstVu.BackgroundColor = System.Drawing.Color.Gainsboro
        Me.lstVu.ColIndex = 0
        Me.lstVu.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ID, Me.Month, Me.Total_Entry, Me.Action, Me.MonthNum})
        Me.lstVu.DGEditMode = False
        Me.lstVu.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstVu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lstVu.GridColor = System.Drawing.SystemColors.HotTrack
        Me.lstVu.Location = New System.Drawing.Point(5, 5)
        Me.lstVu.MultiSelect = False
        Me.lstVu.Name = "lstVu"
        Me.lstVu.ReadOnly = True
        Me.lstVu.RowHeadersWidth = 50
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        Me.lstVu.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.lstVu.Size = New System.Drawing.Size(415, 243)
        Me.lstVu.TabIndex = 0
        Me.lstVu.UseRowColor = True
        Me.lstVu.UseRowCount = True
        Me.lstVu.UseVisualStyleRendererCell = True
        Me.lstVu.UseVisualStyleRendererRow = True
        '
        'ID
        '
        Me.ID.HeaderText = "ID"
        Me.ID.Name = "ID"
        Me.ID.ReadOnly = True
        Me.ID.Visible = False
        Me.ID.Width = 40
        '
        'Month
        '
        Me.Month.HeaderText = "Month"
        Me.Month.Name = "Month"
        Me.Month.ReadOnly = True
        Me.Month.Width = 140
        '
        'Total_Entry
        '
        Me.Total_Entry.HeaderText = "Total Entry"
        Me.Total_Entry.Name = "Total_Entry"
        Me.Total_Entry.ReadOnly = True
        Me.Total_Entry.Width = 80
        '
        'Action
        '
        Me.Action.HeaderText = "Action"
        Me.Action.Name = "Action"
        Me.Action.ReadOnly = True
        Me.Action.Width = 50
        '
        'MonthNum
        '
        Me.MonthNum.HeaderText = "MonthNum"
        Me.MonthNum.Name = "MonthNum"
        Me.MonthNum.ReadOnly = True
        Me.MonthNum.Visible = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel3.Controls.Add(Me.Panel4)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 297)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(425, 43)
        Me.Panel3.TabIndex = 1
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Button2)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel4.Location = New System.Drawing.Point(336, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(89, 43)
        Me.Panel4.TabIndex = 0
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.SystemColors.Control
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.Location = New System.Drawing.Point(9, 6)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 25)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "   Close"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'DeleteVat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.RoyalBlue
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(425, 340)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Font = New System.Drawing.Font("Leelawadee", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DeleteVat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DeleteVat"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        CType(Me.lstVu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CMYear As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Public WithEvents lstVu As BaseClass.BaseGridview
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Month As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Total_Entry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Action As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents MonthNum As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
