<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmFrindBatch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmFrindBatch))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.chkAuto = New System.Windows.Forms.CheckBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.chkShow = New System.Windows.Forms.CheckBox
        Me.tableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.panel1 = New System.Windows.Forms.Panel
        Me.cboStatement = New System.Windows.Forms.ComboBox
        Me.cboFindBy = New System.Windows.Forms.ComboBox
        Me.btnFindNow = New System.Windows.Forms.Button
        Me.btnSelect = New System.Windows.Forms.Button
        Me.label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.txtFilter = New System.Windows.Forms.TextBox
        Me.dgvFrindBath = New System.Windows.Forms.DataGridView
        Me.tableLayoutPanel1.SuspendLayout()
        Me.panel1.SuspendLayout()
        CType(Me.dgvFrindBath, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkAuto
        '
        Me.chkAuto.AutoSize = True
        Me.chkAuto.Checked = True
        Me.chkAuto.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAuto.Location = New System.Drawing.Point(282, 43)
        Me.chkAuto.Name = "chkAuto"
        Me.chkAuto.Size = New System.Drawing.Size(92, 19)
        Me.chkAuto.TabIndex = 5
        Me.chkAuto.Text = "Auto Search"
        Me.chkAuto.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "dd/MM/yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(76, 76)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(195, 21)
        Me.dtpDate.TabIndex = 2
        '
        'chkShow
        '
        Me.chkShow.AutoSize = True
        Me.chkShow.Checked = True
        Me.chkShow.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShow.Location = New System.Drawing.Point(282, 66)
        Me.chkShow.Name = "chkShow"
        Me.chkShow.Size = New System.Drawing.Size(203, 19)
        Me.chkShow.TabIndex = 6
        Me.chkShow.Text = "Show Posted and Delete Batchs"
        Me.chkShow.UseVisualStyleBackColor = True
        '
        'tableLayoutPanel1
        '
        Me.tableLayoutPanel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.tableLayoutPanel1.ColumnCount = 3
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.tableLayoutPanel1.Controls.Add(Me.panel1, 1, 0)
        Me.tableLayoutPanel1.Controls.Add(Me.dgvFrindBath, 1, 1)
        Me.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.tableLayoutPanel1.Name = "tableLayoutPanel1"
        Me.tableLayoutPanel1.RowCount = 3
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 115.0!))
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.tableLayoutPanel1.Size = New System.Drawing.Size(1136, 648)
        Me.tableLayoutPanel1.TabIndex = 1
        '
        'panel1
        '
        Me.panel1.Controls.Add(Me.chkAuto)
        Me.panel1.Controls.Add(Me.chkShow)
        Me.panel1.Controls.Add(Me.dtpDate)
        Me.panel1.Controls.Add(Me.cboStatement)
        Me.panel1.Controls.Add(Me.cboFindBy)
        Me.panel1.Controls.Add(Me.btnFindNow)
        Me.panel1.Controls.Add(Me.btnSelect)
        Me.panel1.Controls.Add(Me.label2)
        Me.panel1.Controls.Add(Me.label1)
        Me.panel1.Controls.Add(Me.txtFilter)
        Me.panel1.Location = New System.Drawing.Point(26, 3)
        Me.panel1.Name = "panel1"
        Me.panel1.Size = New System.Drawing.Size(1083, 108)
        Me.panel1.TabIndex = 0
        '
        'cboStatement
        '
        Me.cboStatement.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStatement.FormattingEnabled = True
        Me.cboStatement.Location = New System.Drawing.Point(76, 40)
        Me.cboStatement.Name = "cboStatement"
        Me.cboStatement.Size = New System.Drawing.Size(195, 23)
        Me.cboStatement.TabIndex = 1
        '
        'cboFindBy
        '
        Me.cboFindBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFindBy.FormattingEnabled = True
        Me.cboFindBy.Items.AddRange(New Object() {"Show all records", "Batch Number", "Description", "BatchDate"})
        Me.cboFindBy.Location = New System.Drawing.Point(76, 9)
        Me.cboFindBy.Name = "cboFindBy"
        Me.cboFindBy.Size = New System.Drawing.Size(195, 23)
        Me.cboFindBy.TabIndex = 0
        '
        'btnFindNow
        '
        Me.btnFindNow.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindNow.Image = CType(resources.GetObject("btnFindNow.Image"), System.Drawing.Image)
        Me.btnFindNow.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnFindNow.Location = New System.Drawing.Point(281, 9)
        Me.btnFindNow.Name = "btnFindNow"
        Me.btnFindNow.Size = New System.Drawing.Size(87, 27)
        Me.btnFindNow.TabIndex = 3
        Me.btnFindNow.Text = "     Find Now"
        Me.btnFindNow.UseVisualStyleBackColor = True
        '
        'btnSelect
        '
        Me.btnSelect.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Image = CType(resources.GetObject("btnSelect.Image"), System.Drawing.Image)
        Me.btnSelect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSelect.Location = New System.Drawing.Point(376, 9)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(87, 27)
        Me.btnSelect.TabIndex = 4
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(17, 77)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(40, 15)
        Me.label2.TabIndex = 0
        Me.label2.Text = "Filter :"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(17, 13)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(53, 15)
        Me.label1.TabIndex = 0
        Me.label1.Text = "Find By :"
        '
        'txtFilter
        '
        Me.txtFilter.Location = New System.Drawing.Point(97, 76)
        Me.txtFilter.Name = "txtFilter"
        Me.txtFilter.Size = New System.Drawing.Size(174, 21)
        Me.txtFilter.TabIndex = 5
        '
        'dgvFrindBath
        '
        Me.dgvFrindBath.AllowUserToAddRows = False
        Me.dgvFrindBath.AllowUserToDeleteRows = False
        Me.dgvFrindBath.AllowUserToResizeRows = False
        Me.dgvFrindBath.BackgroundColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvFrindBath.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvFrindBath.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFrindBath.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvFrindBath.Location = New System.Drawing.Point(26, 118)
        Me.dgvFrindBath.MultiSelect = False
        Me.dgvFrindBath.Name = "dgvFrindBath"
        Me.dgvFrindBath.ReadOnly = True
        Me.dgvFrindBath.RowHeadersVisible = False
        Me.dgvFrindBath.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvFrindBath.Size = New System.Drawing.Size(1084, 504)
        Me.dgvFrindBath.TabIndex = 1
        '
        'FrmFrindBatch
        '
        Me.AcceptButton = Me.btnFindNow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1136, 648)
        Me.Controls.Add(Me.tableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmFrindBatch"
        Me.Text = "FrmSelect"
        Me.tableLayoutPanel1.ResumeLayout(False)
        Me.panel1.ResumeLayout(False)
        Me.panel1.PerformLayout()
        CType(Me.dgvFrindBath, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents chkAuto As System.Windows.Forms.CheckBox
    Private WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Private WithEvents chkShow As System.Windows.Forms.CheckBox
    Private WithEvents tableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Private WithEvents panel1 As System.Windows.Forms.Panel
    Private WithEvents cboStatement As System.Windows.Forms.ComboBox
    Private WithEvents cboFindBy As System.Windows.Forms.ComboBox
    Private WithEvents btnFindNow As System.Windows.Forms.Button
    Private WithEvents btnSelect As System.Windows.Forms.Button
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Public WithEvents txtFilter As System.Windows.Forms.TextBox
    Private WithEvents dgvFrindBath As System.Windows.Forms.DataGridView
End Class
