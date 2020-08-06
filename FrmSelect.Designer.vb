<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSelect
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmSelect))
        Me.tableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.panel1 = New System.Windows.Forms.Panel
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtEntryT = New System.Windows.Forms.TextBox
        Me.txtEntryF = New System.Windows.Forms.TextBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboModule = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnSelect = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.txtBatchT = New System.Windows.Forms.TextBox
        Me.txtBatchF = New System.Windows.Forms.TextBox
        Me.dgvShow = New System.Windows.Forms.DataGridView
        Me.tableLayoutPanel1.SuspendLayout()
        Me.panel1.SuspendLayout()
        CType(Me.dgvShow, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tableLayoutPanel1
        '
        Me.tableLayoutPanel1.ColumnCount = 3
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tableLayoutPanel1.Controls.Add(Me.panel1, 1, 0)
        Me.tableLayoutPanel1.Controls.Add(Me.dgvShow, 1, 1)
        Me.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.tableLayoutPanel1.Name = "tableLayoutPanel1"
        Me.tableLayoutPanel1.RowCount = 3
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.tableLayoutPanel1.Size = New System.Drawing.Size(791, 504)
        Me.tableLayoutPanel1.TabIndex = 2
        '
        'panel1
        '
        Me.panel1.Controls.Add(Me.Label6)
        Me.panel1.Controls.Add(Me.Label5)
        Me.panel1.Controls.Add(Me.txtEntryT)
        Me.panel1.Controls.Add(Me.txtEntryF)
        Me.panel1.Controls.Add(Me.ComboBox2)
        Me.panel1.Controls.Add(Me.Label4)
        Me.panel1.Controls.Add(Me.cboModule)
        Me.panel1.Controls.Add(Me.Label3)
        Me.panel1.Controls.Add(Me.btnSelect)
        Me.panel1.Controls.Add(Me.Label2)
        Me.panel1.Controls.Add(Me.label1)
        Me.panel1.Controls.Add(Me.txtBatchT)
        Me.panel1.Controls.Add(Me.txtBatchF)
        Me.panel1.Location = New System.Drawing.Point(23, 3)
        Me.panel1.Name = "panel1"
        Me.panel1.Size = New System.Drawing.Size(745, 94)
        Me.panel1.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(321, 62)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 15)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Entry To"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(34, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 15)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Entry From"
        '
        'txtEntryT
        '
        Me.txtEntryT.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEntryT.Location = New System.Drawing.Point(377, 59)
        Me.txtEntryT.Name = "txtEntryT"
        Me.txtEntryT.Size = New System.Drawing.Size(148, 21)
        Me.txtEntryT.TabIndex = 12
        '
        'txtEntryF
        '
        Me.txtEntryF.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEntryF.Location = New System.Drawing.Point(106, 59)
        Me.txtEntryF.Name = "txtEntryF"
        Me.txtEntryF.Size = New System.Drawing.Size(149, 21)
        Me.txtEntryF.TabIndex = 11
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(377, 6)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(148, 23)
        Me.ComboBox2.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(299, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 15)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Transection"
        '
        'cboModule
        '
        Me.cboModule.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboModule.FormattingEnabled = True
        Me.cboModule.Items.AddRange(New Object() {"AP", "AR", "GL", "CB"})
        Me.cboModule.Location = New System.Drawing.Point(106, 6)
        Me.cboModule.Name = "cboModule"
        Me.cboModule.Size = New System.Drawing.Size(149, 23)
        Me.cboModule.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(53, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 15)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Module"
        '
        'btnSelect
        '
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Image = Global.ProgramVAT.My.Resources.Resources.cursor__2_
        Me.btnSelect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSelect.Location = New System.Drawing.Point(566, 55)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(91, 29)
        Me.btnSelect.TabIndex = 6
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(317, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Batch To"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(30, 36)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(70, 15)
        Me.label1.TabIndex = 4
        Me.label1.Text = "Batch From"
        '
        'txtBatchT
        '
        Me.txtBatchT.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBatchT.Location = New System.Drawing.Point(377, 33)
        Me.txtBatchT.Name = "txtBatchT"
        Me.txtBatchT.Size = New System.Drawing.Size(148, 21)
        Me.txtBatchT.TabIndex = 0
        '
        'txtBatchF
        '
        Me.txtBatchF.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBatchF.Location = New System.Drawing.Point(106, 33)
        Me.txtBatchF.Name = "txtBatchF"
        Me.txtBatchF.Size = New System.Drawing.Size(149, 21)
        Me.txtBatchF.TabIndex = 0
        '
        'dgvShow
        '
        Me.dgvShow.AllowUserToAddRows = False
        Me.dgvShow.AllowUserToDeleteRows = False
        Me.dgvShow.AllowUserToResizeRows = False
        Me.dgvShow.BackgroundColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvShow.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvShow.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvShow.Location = New System.Drawing.Point(23, 103)
        Me.dgvShow.MultiSelect = False
        Me.dgvShow.Name = "dgvShow"
        Me.dgvShow.ReadOnly = True
        Me.dgvShow.RowHeadersVisible = False
        Me.dgvShow.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvShow.Size = New System.Drawing.Size(745, 378)
        Me.dgvShow.TabIndex = 1
        '
        'FrmSelect
        '
        Me.AcceptButton = Me.btnSelect
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(791, 504)
        Me.Controls.Add(Me.tableLayoutPanel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmSelect"
        Me.Text = "FrmSelect"
        Me.tableLayoutPanel1.ResumeLayout(False)
        Me.panel1.ResumeLayout(False)
        Me.panel1.PerformLayout()
        CType(Me.dgvShow, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents tableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Private WithEvents panel1 As System.Windows.Forms.Panel
    Private WithEvents dgvShow As System.Windows.Forms.DataGridView
    Friend WithEvents txtBatchT As System.Windows.Forms.TextBox
    Friend WithEvents txtBatchF As System.Windows.Forms.TextBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Private WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboModule As System.Windows.Forms.ComboBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtEntryT As System.Windows.Forms.TextBox
    Friend WithEvents txtEntryF As System.Windows.Forms.TextBox
End Class
