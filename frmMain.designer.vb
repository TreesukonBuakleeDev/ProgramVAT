<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
    ' Inherits BaseClass.BaseFrom
    Inherits Form
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
    Public dlgFileOpen As System.Windows.Forms.OpenFileDialog
    'BaseClass.MBGlassButton
    Public WithEvents cboLoc As System.Windows.Forms.ComboBox
    Public WithEvents txtRate As System.Windows.Forms.TextBox
    Public WithEvents cboRate As System.Windows.Forms.ComboBox
    Public WithEvents cboView As System.Windows.Forms.ComboBox
    Public WithEvents lblRptParam As System.Windows.Forms.Label
    Public WithEvents LbLocation As System.Windows.Forms.Label
    Public WithEvents lblRate As System.Windows.Forms.Label
    Public WithEvents LbRate As System.Windows.Forms.Label
    Public WithEvents LbView As System.Windows.Forms.Label
    Public WithEvents LbFrom As System.Windows.Forms.Label
    Public WithEvents LbTo As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.dlgFileOpen = New System.Windows.Forms.OpenFileDialog()
        Me.cboLoc = New System.Windows.Forms.ComboBox()
        Me.txtRate = New System.Windows.Forms.TextBox()
        Me.cboRate = New System.Windows.Forms.ComboBox()
        Me.cboView = New System.Windows.Forms.ComboBox()
        Me.lblRptParam = New System.Windows.Forms.Label()
        Me.LbLocation = New System.Windows.Forms.Label()
        Me.lblRate = New System.Windows.Forms.Label()
        Me.LbRate = New System.Windows.Forms.Label()
        Me.LbView = New System.Windows.Forms.Label()
        Me.LbFrom = New System.Windows.Forms.Label()
        Me.LbTo = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dlgFile = New System.Windows.Forms.OpenFileDialog()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Status = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.txtTo = New System.Windows.Forms.DateTimePicker()
        Me.txtFrom = New System.Windows.Forms.DateTimePicker()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.btnToButtom = New BaseClass.MBGlassButton()
        Me.btnToNex = New BaseClass.MBGlassButton()
        Me.btnToPre = New BaseClass.MBGlassButton()
        Me.cboPage = New System.Windows.Forms.ComboBox()
        Me.btnToTop = New BaseClass.MBGlassButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BaseLineSeparate2 = New BaseClass.BaseLineSeparate()
        Me.btnAbout = New BaseClass.MBGlassButton()
        Me.TaskBar = New wyDay.Controls.Windows7ProgressBar()
        Me.lblCompany = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lstVu = New BaseClass.BaseGridview()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnDBSet = New BaseClass.MBGlassButton()
        Me.btnUserSet = New BaseClass.MBGlassButton()
        Me.btnOpen = New BaseClass.MBGlassButton()
        Me.btnSetTaxNo = New BaseClass.MBGlassButton()
        Me.btnSetLocation = New BaseClass.MBGlassButton()
        Me.btnGetVat = New BaseClass.MBGlassButton()
        Me.btnPrint = New BaseClass.MBGlassButton()
        Me.btnAddVat = New BaseClass.MBGlassButton()
        Me.btnEditVat = New BaseClass.MBGlassButton()
        Me.btnRunning = New BaseClass.MBGlassButton()
        Me.btnDeleteVat = New BaseClass.MBGlassButton()
        Me.btnRate = New BaseClass.MBGlassButton()
        Me.btnHelp = New BaseClass.MBGlassButton()
        Me.btnExit = New BaseClass.MBGlassButton()
        Me.btnRefreh = New BaseClass.MBGlassButton()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.BaseLineSeparate2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.lstVu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboLoc
        '
        Me.cboLoc.BackColor = System.Drawing.SystemColors.Window
        Me.cboLoc.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLoc.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLoc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLoc.Location = New System.Drawing.Point(348, 4)
        Me.cboLoc.Name = "cboLoc"
        Me.cboLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLoc.Size = New System.Drawing.Size(98, 23)
        Me.cboLoc.TabIndex = 2
        '
        'txtRate
        '
        Me.txtRate.AcceptsReturn = True
        Me.txtRate.BackColor = System.Drawing.SystemColors.Window
        Me.txtRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRate.Location = New System.Drawing.Point(718, 4)
        Me.txtRate.MaxLength = 5
        Me.txtRate.Name = "txtRate"
        Me.txtRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRate.Size = New System.Drawing.Size(36, 21)
        Me.txtRate.TabIndex = 5
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRate
        '
        Me.cboRate.BackColor = System.Drawing.SystemColors.Window
        Me.cboRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRate.Location = New System.Drawing.Point(637, 4)
        Me.cboRate.Name = "cboRate"
        Me.cboRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRate.Size = New System.Drawing.Size(75, 23)
        Me.cboRate.TabIndex = 4
        '
        'cboView
        '
        Me.cboView.BackColor = System.Drawing.SystemColors.Window
        Me.cboView.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboView.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboView.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboView.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboView.Location = New System.Drawing.Point(495, 4)
        Me.cboView.Name = "cboView"
        Me.cboView.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboView.Size = New System.Drawing.Size(93, 23)
        Me.cboView.TabIndex = 3
        '
        'lblRptParam
        '
        Me.lblRptParam.BackColor = System.Drawing.SystemColors.Control
        Me.lblRptParam.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRptParam.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRptParam.Location = New System.Drawing.Point(730, 488)
        Me.lblRptParam.Name = "lblRptParam"
        Me.lblRptParam.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRptParam.Size = New System.Drawing.Size(57, 22)
        Me.lblRptParam.TabIndex = 29
        Me.lblRptParam.Text = "Label8"
        Me.lblRptParam.Visible = False
        '
        'LbLocation
        '
        Me.LbLocation.AutoSize = True
        Me.LbLocation.BackColor = System.Drawing.Color.Transparent
        Me.LbLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.LbLocation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbLocation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LbLocation.Location = New System.Drawing.Point(283, 8)
        Me.LbLocation.Name = "LbLocation"
        Me.LbLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LbLocation.Size = New System.Drawing.Size(54, 15)
        Me.LbLocation.TabIndex = 26
        Me.LbLocation.Text = "Location"
        '
        'lblRate
        '
        Me.lblRate.AutoSize = True
        Me.lblRate.BackColor = System.Drawing.Color.Transparent
        Me.lblRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lblRate.Location = New System.Drawing.Point(756, 8)
        Me.lblRate.Name = "lblRate"
        Me.lblRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblRate.Size = New System.Drawing.Size(18, 15)
        Me.lblRate.TabIndex = 25
        Me.lblRate.Text = "%"
        '
        'LbRate
        '
        Me.LbRate.AutoSize = True
        Me.LbRate.BackColor = System.Drawing.Color.Transparent
        Me.LbRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.LbRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LbRate.Location = New System.Drawing.Point(598, 7)
        Me.LbRate.Name = "LbRate"
        Me.LbRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LbRate.Size = New System.Drawing.Size(33, 15)
        Me.LbRate.TabIndex = 24
        Me.LbRate.Text = "Rate"
        Me.LbRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LbView
        '
        Me.LbView.AutoSize = True
        Me.LbView.BackColor = System.Drawing.Color.Transparent
        Me.LbView.Cursor = System.Windows.Forms.Cursors.Default
        Me.LbView.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbView.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LbView.Location = New System.Drawing.Point(454, 8)
        Me.LbView.Name = "LbView"
        Me.LbView.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LbView.Size = New System.Drawing.Size(36, 15)
        Me.LbView.TabIndex = 23
        Me.LbView.Text = "View "
        '
        'LbFrom
        '
        Me.LbFrom.AutoSize = True
        Me.LbFrom.BackColor = System.Drawing.Color.Transparent
        Me.LbFrom.Cursor = System.Windows.Forms.Cursors.Default
        Me.LbFrom.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbFrom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LbFrom.Location = New System.Drawing.Point(3, 8)
        Me.LbFrom.Name = "LbFrom"
        Me.LbFrom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LbFrom.Size = New System.Drawing.Size(36, 15)
        Me.LbFrom.TabIndex = 22
        Me.LbFrom.Text = "From"
        '
        'LbTo
        '
        Me.LbTo.AutoSize = True
        Me.LbTo.BackColor = System.Drawing.Color.Transparent
        Me.LbTo.Cursor = System.Windows.Forms.Cursors.Default
        Me.LbTo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbTo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LbTo.Location = New System.Drawing.Point(148, 8)
        Me.LbTo.Name = "LbTo"
        Me.LbTo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LbTo.Size = New System.Drawing.Size(20, 15)
        Me.LbTo.TabIndex = 21
        Me.LbTo.Text = "To"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(1068, -25)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(10, 10)
        Me.Label2.TabIndex = 18
        Me.Label2.Visible = False
        '
        'dlgFile
        '
        Me.dlgFile.FileName = "OpenFileDialog1"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.Status, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 672)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1008, 22)
        Me.StatusStrip1.TabIndex = 36
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(120, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Status
        '
        Me.Status.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Status.Name = "Status"
        Me.Status.Size = New System.Drawing.Size(873, 17)
        Me.Status.Spring = True
        Me.Status.Text = "123"
        Me.Status.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        Me.ToolStripProgressBar1.Visible = False
        '
        'txtTo
        '
        Me.txtTo.CustomFormat = "dd/MM/yyyy"
        Me.txtTo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtTo.Location = New System.Drawing.Point(174, 4)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(99, 21)
        Me.txtTo.TabIndex = 1
        '
        'txtFrom
        '
        Me.txtFrom.CustomFormat = "dd/MM/yyyy"
        Me.txtFrom.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtFrom.Location = New System.Drawing.Point(46, 4)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(99, 21)
        Me.txtFrom.TabIndex = 0
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoSize = True
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.FlowLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(1008, 0)
        Me.FlowLayoutPanel1.TabIndex = 37
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.Panel2.Controls.Add(Me.txtFrom)
        Me.Panel2.Controls.Add(Me.txtTo)
        Me.Panel2.Controls.Add(Me.LbTo)
        Me.Panel2.Controls.Add(Me.LbFrom)
        Me.Panel2.Controls.Add(Me.LbView)
        Me.Panel2.Controls.Add(Me.LbRate)
        Me.Panel2.Controls.Add(Me.cboLoc)
        Me.Panel2.Controls.Add(Me.lblRate)
        Me.Panel2.Controls.Add(Me.txtRate)
        Me.Panel2.Controls.Add(Me.LbLocation)
        Me.Panel2.Controls.Add(Me.cboRate)
        Me.Panel2.Controls.Add(Me.cboView)
        Me.Panel2.Controls.Add(Me.Panel5)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 41)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1008, 30)
        Me.Panel2.TabIndex = 0
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.btnToButtom)
        Me.Panel5.Controls.Add(Me.btnToNex)
        Me.Panel5.Controls.Add(Me.btnToPre)
        Me.Panel5.Controls.Add(Me.cboPage)
        Me.Panel5.Controls.Add(Me.btnToTop)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel5.Location = New System.Drawing.Point(781, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(227, 30)
        Me.Panel5.TabIndex = 27
        '
        'btnToButtom
        '
        Me.btnToButtom.BackColor = System.Drawing.Color.Transparent
        Me.btnToButtom.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnToButtom.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnToButtom.FlatAppearance.BorderSize = 0
        Me.btnToButtom.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnToButtom.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnToButtom.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnToButtom.Location = New System.Drawing.Point(173, 3)
        Me.btnToButtom.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnToButtom.Name = "btnToButtom"
        Me.btnToButtom.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnToButtom.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnToButtom.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToButtom.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToButtom.Size = New System.Drawing.Size(48, 24)
        Me.btnToButtom.TabIndex = 4
        Me.btnToButtom.Text = ">>"
        Me.btnToButtom.UseVisualStyleBackColor = True
        '
        'btnToNex
        '
        Me.btnToNex.BackColor = System.Drawing.Color.Transparent
        Me.btnToNex.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnToNex.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnToNex.FlatAppearance.BorderSize = 0
        Me.btnToNex.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnToNex.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnToNex.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnToNex.Location = New System.Drawing.Point(144, 3)
        Me.btnToNex.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnToNex.Name = "btnToNex"
        Me.btnToNex.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnToNex.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnToNex.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToNex.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToNex.Size = New System.Drawing.Size(28, 24)
        Me.btnToNex.TabIndex = 3
        Me.btnToNex.Text = ">>"
        Me.btnToNex.UseVisualStyleBackColor = True
        '
        'btnToPre
        '
        Me.btnToPre.BackColor = System.Drawing.Color.Transparent
        Me.btnToPre.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnToPre.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnToPre.FlatAppearance.BorderSize = 0
        Me.btnToPre.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnToPre.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnToPre.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnToPre.Location = New System.Drawing.Point(52, 3)
        Me.btnToPre.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnToPre.Name = "btnToPre"
        Me.btnToPre.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnToPre.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnToPre.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToPre.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToPre.Size = New System.Drawing.Size(28, 24)
        Me.btnToPre.TabIndex = 1
        Me.btnToPre.Text = "<"
        Me.btnToPre.UseVisualStyleBackColor = True
        '
        'cboPage
        '
        Me.cboPage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPage.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboPage.FormattingEnabled = True
        Me.cboPage.Location = New System.Drawing.Point(81, 3)
        Me.cboPage.Name = "cboPage"
        Me.cboPage.Size = New System.Drawing.Size(62, 23)
        Me.cboPage.TabIndex = 2
        '
        'btnToTop
        '
        Me.btnToTop.BackColor = System.Drawing.Color.Transparent
        Me.btnToTop.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnToTop.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnToTop.FlatAppearance.BorderSize = 0
        Me.btnToTop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnToTop.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnToTop.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnToTop.Location = New System.Drawing.Point(3, 3)
        Me.btnToTop.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnToTop.Name = "btnToTop"
        Me.btnToTop.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnToTop.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnToTop.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToTop.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnToTop.Size = New System.Drawing.Size(48, 24)
        Me.btnToTop.TabIndex = 0
        Me.btnToTop.Text = "<<"
        Me.btnToTop.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.Panel1.Controls.Add(Me.BaseLineSeparate2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1008, 41)
        Me.Panel1.TabIndex = 38
        '
        'BaseLineSeparate2
        '
        Me.BaseLineSeparate2.BackColor = System.Drawing.Color.Transparent
        Me.BaseLineSeparate2.BackColorStyle = BaseClass.BaseLineSeparate.EnumBackgrounFillStyle.Gradient
        Me.BaseLineSeparate2.Controls.Add(Me.btnAbout)
        Me.BaseLineSeparate2.Controls.Add(Me.TaskBar)
        Me.BaseLineSeparate2.Controls.Add(Me.lblCompany)
        Me.BaseLineSeparate2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BaseLineSeparate2.FillColor = System.Drawing.Color.FromArgb(CType(CType(20, Byte), Integer), CType(CType(20, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.BaseLineSeparate2.FillColorGradient = System.Drawing.Color.FromArgb(CType(CType(134, Byte), Integer), CType(CType(206, Byte), Integer), CType(CType(225, Byte), Integer))
        Me.BaseLineSeparate2.FillColorStyle = BaseClass.BaseLineSeparate.EnumBackgrounFillStyle.Solid
        Me.BaseLineSeparate2.GradientColor1 = System.Drawing.Color.FromArgb(CType(CType(40, Byte), Integer), CType(CType(40, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.BaseLineSeparate2.GradientColor2 = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.BaseLineSeparate2.HeaderAngle = "38"
        Me.BaseLineSeparate2.HeaderHeight = 38
        Me.BaseLineSeparate2.HeaderStyle = BaseClass.BaseLineSeparate.Style.TopRight
        Me.BaseLineSeparate2.HeaderWidth = 100
        Me.BaseLineSeparate2.LineAngle = "6"
        Me.BaseLineSeparate2.LineColor = System.Drawing.Color.Gray
        Me.BaseLineSeparate2.LineHeight = 10
        Me.BaseLineSeparate2.Location = New System.Drawing.Point(0, 0)
        Me.BaseLineSeparate2.Name = "BaseLineSeparate2"
        Me.BaseLineSeparate2.ShowPosition = BaseClass.BaseLineSeparate.Position.Top
        Me.BaseLineSeparate2.Size = New System.Drawing.Size(1008, 41)
        Me.BaseLineSeparate2.TabIndex = 45
        Me.BaseLineSeparate2.Txt = ""
        Me.BaseLineSeparate2.txtDetail = ""
        '
        'btnAbout
        '
        Me.btnAbout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAbout.BackColor = System.Drawing.Color.FromArgb(CType(CType(215, Byte), Integer), CType(CType(228, Byte), Integer), CType(CType(242, Byte), Integer))
        Me.btnAbout.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnAbout.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAbout.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAbout.FlatAppearance.BorderSize = 0
        Me.btnAbout.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAbout.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAbout.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAbout.Image = Global.ProgramVAT.My.Resources.Resources.talking_about_money
        Me.btnAbout.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAbout.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnAbout.Location = New System.Drawing.Point(919, 10)
        Me.btnAbout.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnAbout.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnAbout.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnAbout.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnAbout.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAbout.Size = New System.Drawing.Size(77, 25)
        Me.btnAbout.TabIndex = 0
        Me.btnAbout.Text = "About"
        Me.btnAbout.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAbout.UseVisualStyleBackColor = False
        '
        'TaskBar
        '
        Me.TaskBar.ContainerControl = Me
        Me.TaskBar.Location = New System.Drawing.Point(718, 13)
        Me.TaskBar.Name = "TaskBar"
        Me.TaskBar.Size = New System.Drawing.Size(142, 22)
        Me.TaskBar.Step = 0
        Me.TaskBar.TabIndex = 41
        Me.TaskBar.Visible = False
        '
        'lblCompany
        '
        Me.lblCompany.AutoSize = True
        Me.lblCompany.BackColor = System.Drawing.Color.Transparent
        Me.lblCompany.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCompany.Font = New System.Drawing.Font("Tahoma", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lblCompany.ForeColor = System.Drawing.Color.White
        Me.lblCompany.Location = New System.Drawing.Point(12, 14)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCompany.Size = New System.Drawing.Size(99, 23)
        Me.lblCompany.TabIndex = 42
        Me.lblCompany.Text = "Company"
        Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.Panel3.Controls.Add(Me.lstVu)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(111, 71)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Padding = New System.Windows.Forms.Padding(1, 5, 5, 5)
        Me.Panel3.Size = New System.Drawing.Size(897, 601)
        Me.Panel3.TabIndex = 39
        '
        'lstVu
        '
        Me.lstVu.AllowUserToOrderColumns = True
        Me.lstVu.AllowUserToResizeRows = False
        Me.lstVu.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.lstVu.ColIndex = 0
        Me.lstVu.DGEditMode = False
        Me.lstVu.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstVu.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lstVu.Location = New System.Drawing.Point(1, 5)
        Me.lstVu.Name = "lstVu"
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        Me.lstVu.RowsDefaultCellStyle = DataGridViewCellStyle1
        Me.lstVu.Size = New System.Drawing.Size(891, 591)
        Me.lstVu.TabIndex = 0
        Me.lstVu.UseRowColor = True
        Me.lstVu.UseRowCount = True
        Me.lstVu.UseVisualStyleRendererCell = False
        Me.lstVu.UseVisualStyleRendererRow = False
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.Panel4.Controls.Add(Me.GroupBox1)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel4.Location = New System.Drawing.Point(0, 71)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Padding = New System.Windows.Forms.Padding(5, 0, 3, 5)
        Me.Panel4.Size = New System.Drawing.Size(111, 601)
        Me.Panel4.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.btnDBSet)
        Me.GroupBox1.Controls.Add(Me.btnUserSet)
        Me.GroupBox1.Controls.Add(Me.btnOpen)
        Me.GroupBox1.Controls.Add(Me.btnSetTaxNo)
        Me.GroupBox1.Controls.Add(Me.btnSetLocation)
        Me.GroupBox1.Controls.Add(Me.btnGetVat)
        Me.GroupBox1.Controls.Add(Me.btnPrint)
        Me.GroupBox1.Controls.Add(Me.btnAddVat)
        Me.GroupBox1.Controls.Add(Me.btnEditVat)
        Me.GroupBox1.Controls.Add(Me.btnRunning)
        Me.GroupBox1.Controls.Add(Me.btnDeleteVat)
        Me.GroupBox1.Controls.Add(Me.btnRate)
        Me.GroupBox1.Controls.Add(Me.btnHelp)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnRefreh)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(5, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(103, 596)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnDBSet
        '
        Me.btnDBSet.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnDBSet.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnDBSet.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnDBSet.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDBSet.FlatAppearance.BorderSize = 0
        Me.btnDBSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDBSet.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDBSet.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDBSet.Image = CType(resources.GetObject("btnDBSet.Image"), System.Drawing.Image)
        Me.btnDBSet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDBSet.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnDBSet.Location = New System.Drawing.Point(8, 413)
        Me.btnDBSet.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnDBSet.Name = "btnDBSet"
        Me.btnDBSet.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnDBSet.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnDBSet.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDBSet.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDBSet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDBSet.Size = New System.Drawing.Size(89, 30)
        Me.btnDBSet.TabIndex = 11
        Me.btnDBSet.Text = "DB Set"
        Me.btnDBSet.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDBSet.UseVisualStyleBackColor = False
        '
        'btnUserSet
        '
        Me.btnUserSet.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnUserSet.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnUserSet.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnUserSet.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnUserSet.FlatAppearance.BorderSize = 0
        Me.btnUserSet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUserSet.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUserSet.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnUserSet.Image = CType(resources.GetObject("btnUserSet.Image"), System.Drawing.Image)
        Me.btnUserSet.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnUserSet.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnUserSet.Location = New System.Drawing.Point(8, 449)
        Me.btnUserSet.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnUserSet.Name = "btnUserSet"
        Me.btnUserSet.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnUserSet.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnUserSet.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUserSet.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUserSet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnUserSet.Size = New System.Drawing.Size(89, 30)
        Me.btnUserSet.TabIndex = 12
        Me.btnUserSet.Text = "User Set"
        Me.btnUserSet.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnUserSet.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnOpen.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnOpen.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnOpen.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnOpen.FlatAppearance.BorderSize = 0
        Me.btnOpen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpen.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnOpen.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnOpen.Location = New System.Drawing.Point(8, 17)
        Me.btnOpen.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnOpen.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnOpen.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnOpen.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnOpen.Size = New System.Drawing.Size(89, 30)
        Me.btnOpen.TabIndex = 0
        Me.btnOpen.TabStop = False
        Me.btnOpen.Text = "Open"
        Me.btnOpen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnSetTaxNo
        '
        Me.btnSetTaxNo.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnSetTaxNo.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnSetTaxNo.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnSetTaxNo.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSetTaxNo.FlatAppearance.BorderSize = 0
        Me.btnSetTaxNo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSetTaxNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSetTaxNo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSetTaxNo.Image = CType(resources.GetObject("btnSetTaxNo.Image"), System.Drawing.Image)
        Me.btnSetTaxNo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSetTaxNo.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnSetTaxNo.Location = New System.Drawing.Point(8, 53)
        Me.btnSetTaxNo.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnSetTaxNo.Name = "btnSetTaxNo"
        Me.btnSetTaxNo.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnSetTaxNo.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnSetTaxNo.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSetTaxNo.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSetTaxNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSetTaxNo.Size = New System.Drawing.Size(89, 30)
        Me.btnSetTaxNo.TabIndex = 1
        Me.btnSetTaxNo.Text = "Tax No"
        Me.btnSetTaxNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSetTaxNo.UseVisualStyleBackColor = False
        '
        'btnSetLocation
        '
        Me.btnSetLocation.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnSetLocation.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnSetLocation.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnSetLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSetLocation.FlatAppearance.BorderSize = 0
        Me.btnSetLocation.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSetLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSetLocation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSetLocation.Image = Global.ProgramVAT.My.Resources.Resources.building
        Me.btnSetLocation.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSetLocation.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnSetLocation.Location = New System.Drawing.Point(8, 89)
        Me.btnSetLocation.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnSetLocation.Name = "btnSetLocation"
        Me.btnSetLocation.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnSetLocation.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnSetLocation.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSetLocation.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSetLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSetLocation.Size = New System.Drawing.Size(89, 30)
        Me.btnSetLocation.TabIndex = 2
        Me.btnSetLocation.Text = "Location"
        Me.btnSetLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSetLocation.UseVisualStyleBackColor = False
        '
        'btnGetVat
        '
        Me.btnGetVat.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnGetVat.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnGetVat.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnGetVat.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnGetVat.FlatAppearance.BorderSize = 0
        Me.btnGetVat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnGetVat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetVat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnGetVat.Image = Global.ProgramVAT.My.Resources.Resources.profit__1_
        Me.btnGetVat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnGetVat.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnGetVat.Location = New System.Drawing.Point(8, 161)
        Me.btnGetVat.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnGetVat.Name = "btnGetVat"
        Me.btnGetVat.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnGetVat.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnGetVat.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnGetVat.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnGetVat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnGetVat.Size = New System.Drawing.Size(89, 30)
        Me.btnGetVat.TabIndex = 4
        Me.btnGetVat.Text = "Get Vat"
        Me.btnGetVat.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnGetVat.UseVisualStyleBackColor = False
        '
        'btnPrint
        '
        Me.btnPrint.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnPrint.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnPrint.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnPrint.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnPrint.FlatAppearance.BorderSize = 0
        Me.btnPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnPrint.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnPrint.Location = New System.Drawing.Point(8, 377)
        Me.btnPrint.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnPrint.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnPrint.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnPrint.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnPrint.Size = New System.Drawing.Size(89, 30)
        Me.btnPrint.TabIndex = 10
        Me.btnPrint.Text = "Print"
        Me.btnPrint.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnPrint.UseVisualStyleBackColor = False
        '
        'btnAddVat
        '
        Me.btnAddVat.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnAddVat.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnAddVat.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnAddVat.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnAddVat.FlatAppearance.BorderSize = 0
        Me.btnAddVat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAddVat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddVat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAddVat.Image = Global.ProgramVAT.My.Resources.Resources.add41
        Me.btnAddVat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAddVat.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnAddVat.Location = New System.Drawing.Point(8, 197)
        Me.btnAddVat.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnAddVat.Name = "btnAddVat"
        Me.btnAddVat.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnAddVat.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnAddVat.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnAddVat.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnAddVat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnAddVat.Size = New System.Drawing.Size(89, 30)
        Me.btnAddVat.TabIndex = 5
        Me.btnAddVat.Text = "Add Vat"
        Me.btnAddVat.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAddVat.UseVisualStyleBackColor = False
        '
        'btnEditVat
        '
        Me.btnEditVat.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnEditVat.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnEditVat.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnEditVat.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnEditVat.FlatAppearance.BorderSize = 0
        Me.btnEditVat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEditVat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnEditVat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnEditVat.Image = Global.ProgramVAT.My.Resources.Resources.edit
        Me.btnEditVat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnEditVat.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnEditVat.Location = New System.Drawing.Point(8, 233)
        Me.btnEditVat.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnEditVat.Name = "btnEditVat"
        Me.btnEditVat.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnEditVat.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnEditVat.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnEditVat.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnEditVat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnEditVat.Size = New System.Drawing.Size(89, 30)
        Me.btnEditVat.TabIndex = 6
        Me.btnEditVat.Text = "Edit Vat"
        Me.btnEditVat.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnEditVat.UseVisualStyleBackColor = False
        '
        'btnRunning
        '
        Me.btnRunning.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnRunning.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnRunning.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnRunning.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRunning.FlatAppearance.BorderSize = 0
        Me.btnRunning.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRunning.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRunning.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRunning.Image = Global.ProgramVAT.My.Resources.Resources.media_play_symbol
        Me.btnRunning.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRunning.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnRunning.Location = New System.Drawing.Point(8, 341)
        Me.btnRunning.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnRunning.Name = "btnRunning"
        Me.btnRunning.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnRunning.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnRunning.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRunning.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRunning.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRunning.Size = New System.Drawing.Size(89, 30)
        Me.btnRunning.TabIndex = 9
        Me.btnRunning.Text = "Running"
        Me.btnRunning.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnRunning.UseVisualStyleBackColor = False
        '
        'btnDeleteVat
        '
        Me.btnDeleteVat.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnDeleteVat.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnDeleteVat.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnDeleteVat.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnDeleteVat.FlatAppearance.BorderSize = 0
        Me.btnDeleteVat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDeleteVat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteVat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDeleteVat.Image = Global.ProgramVAT.My.Resources.Resources.bin1
        Me.btnDeleteVat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDeleteVat.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnDeleteVat.Location = New System.Drawing.Point(8, 269)
        Me.btnDeleteVat.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnDeleteVat.Name = "btnDeleteVat"
        Me.btnDeleteVat.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnDeleteVat.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnDeleteVat.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDeleteVat.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDeleteVat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnDeleteVat.Size = New System.Drawing.Size(89, 30)
        Me.btnDeleteVat.TabIndex = 7
        Me.btnDeleteVat.Text = "Delete"
        Me.btnDeleteVat.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDeleteVat.UseVisualStyleBackColor = False
        '
        'btnRate
        '
        Me.btnRate.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnRate.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnRate.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnRate.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRate.FlatAppearance.BorderSize = 0
        Me.btnRate.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRate.Image = Global.ProgramVAT.My.Resources.Resources.controls
        Me.btnRate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRate.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnRate.Location = New System.Drawing.Point(8, 305)
        Me.btnRate.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnRate.Name = "btnRate"
        Me.btnRate.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnRate.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnRate.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRate.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRate.Size = New System.Drawing.Size(89, 30)
        Me.btnRate.TabIndex = 8
        Me.btnRate.Text = "Adj.Rate"
        Me.btnRate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnRate.UseVisualStyleBackColor = False
        '
        'btnHelp
        '
        Me.btnHelp.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnHelp.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnHelp.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnHelp.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnHelp.FlatAppearance.BorderSize = 0
        Me.btnHelp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnHelp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHelp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnHelp.Image = Global.ProgramVAT.My.Resources.Resources.info
        Me.btnHelp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnHelp.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnHelp.Location = New System.Drawing.Point(8, 485)
        Me.btnHelp.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnHelp.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnHelp.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnHelp.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnHelp.Size = New System.Drawing.Size(89, 30)
        Me.btnHelp.TabIndex = 13
        Me.btnHelp.Text = "Help"
        Me.btnHelp.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnHelp.UseVisualStyleBackColor = False
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnExit.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnExit.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.FlatAppearance.BorderSize = 0
        Me.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnExit.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnExit.Location = New System.Drawing.Point(8, 521)
        Me.btnExit.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnExit.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnExit.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnExit.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnExit.Size = New System.Drawing.Size(89, 30)
        Me.btnExit.TabIndex = 13
        Me.btnExit.Text = "Exit"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnRefreh
        '
        Me.btnRefreh.BackColor = System.Drawing.Color.FromArgb(CType(CType(153, Byte), Integer), CType(CType(180, Byte), Integer), CType(CType(209, Byte), Integer))
        Me.btnRefreh.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnRefreh.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnRefreh.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnRefreh.FlatAppearance.BorderSize = 0
        Me.btnRefreh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRefreh.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefreh.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnRefreh.Image = CType(resources.GetObject("btnRefreh.Image"), System.Drawing.Image)
        Me.btnRefreh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRefreh.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnRefreh.Location = New System.Drawing.Point(8, 125)
        Me.btnRefreh.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnRefreh.Name = "btnRefreh"
        Me.btnRefreh.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnRefreh.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnRefreh.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRefreh.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnRefreh.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnRefreh.Size = New System.Drawing.Size(89, 30)
        Me.btnRefreh.TabIndex = 3
        Me.btnRefreh.Text = "Refresh"
        Me.btnRefreh.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnRefreh.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(1008, 694)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.lblRptParam)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1022, 688)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmMain"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.BaseLineSeparate2.ResumeLayout(False)
        Me.BaseLineSeparate2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        CType(Me.lstVu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dlgFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    'Public WithEvents TS_DESC As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtFrom As System.Windows.Forms.DateTimePicker
    Public WithEvents lstVu As BaseClass.BaseGridview
    Public WithEvents btnAbout As BaseClass.MBGlassButton
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents cboPage As System.Windows.Forms.ComboBox
    Friend WithEvents BaseLineSeparate2 As BaseClass.BaseLineSeparate
    Public WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents btnToButtom As BaseClass.MBGlassButton
    Friend WithEvents btnToNex As BaseClass.MBGlassButton
    Friend WithEvents btnToPre As BaseClass.MBGlassButton
    Friend WithEvents btnToTop As BaseClass.MBGlassButton
    Public WithEvents btnExit As BaseClass.MBGlassButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents btnDBSet As BaseClass.MBGlassButton
    Public WithEvents btnUserSet As BaseClass.MBGlassButton
    Public WithEvents btnOpen As BaseClass.MBGlassButton
    Public WithEvents btnSetTaxNo As BaseClass.MBGlassButton
    Public WithEvents btnSetLocation As BaseClass.MBGlassButton
    Public WithEvents btnGetVat As BaseClass.MBGlassButton
    Public WithEvents btnPrint As BaseClass.MBGlassButton
    Public WithEvents btnAddVat As BaseClass.MBGlassButton
    Public WithEvents btnEditVat As BaseClass.MBGlassButton
    Public WithEvents btnRunning As BaseClass.MBGlassButton
    Public WithEvents btnDeleteVat As BaseClass.MBGlassButton
    Public WithEvents btnRate As BaseClass.MBGlassButton
    Public WithEvents btnHelp As BaseClass.MBGlassButton
    Public WithEvents btnRefreh As BaseClass.MBGlassButton
    Public WithEvents TaskBar As wyDay.Controls.Windows7ProgressBar
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Status As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
#End Region
End Class