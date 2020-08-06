<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmConfigDB
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmConfigDB))
        Me.lblDBname = New System.Windows.Forms.Label()
        Me.txtCompany = New System.Windows.Forms.TextBox()
        Me.btnNew = New System.Windows.Forms.Button()
        Me.btnP = New System.Windows.Forms.Button()
        Me.btnN = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnACCClose = New System.Windows.Forms.Button()
        Me.btnACCSave = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Ch_AR = New System.Windows.Forms.CheckBox()
        Me.Ch_AP = New System.Windows.Forms.CheckBox()
        Me.Ch_CB = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Ch_Doc = New System.Windows.Forms.RadioButton()
        Me.Ch_PoST = New System.Windows.Forms.RadioButton()
        Me.ch_loginN = New System.Windows.Forms.RadioButton()
        Me.ch_loginY = New System.Windows.Forms.RadioButton()
        Me.Ch_runFormat = New System.Windows.Forms.ComboBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Rd_Affter = New System.Windows.Forms.RadioButton()
        Me.Rd_Before = New System.Windows.Forms.RadioButton()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.CB_PAGE = New System.Windows.Forms.ComboBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.CB_AppY = New System.Windows.Forms.RadioButton()
        Me.CB_AppN = New System.Windows.Forms.RadioButton()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.PO_YES = New System.Windows.Forms.RadioButton()
        Me.PO_NO = New System.Windows.Forms.RadioButton()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.Ch_CBReverse = New System.Windows.Forms.CheckBox()
        Me.Ch_GLReverse = New System.Windows.Forms.CheckBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.cb_OPT_AP = New System.Windows.Forms.CheckBox()
        Me.cb_OPT_CB = New System.Windows.Forms.CheckBox()
        Me.cb_OPT_GL = New System.Windows.Forms.CheckBox()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.cb_UPRUN_Y = New System.Windows.Forms.RadioButton()
        Me.cb_UPRUN_N = New System.Windows.Forms.RadioButton()
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.cb_RERUN_Y = New System.Windows.Forms.RadioButton()
        Me.cb_RERUN_N = New System.Windows.Forms.RadioButton()
        Me.GroupBox14 = New System.Windows.Forms.GroupBox()
        Me.Rb_AutoRunning = New System.Windows.Forms.RadioButton()
        Me.Rb_ManualRunning = New System.Windows.Forms.RadioButton()
        Me.GroupBox15 = New System.Windows.Forms.GroupBox()
        Me.CB_OE = New System.Windows.Forms.CheckBox()
        Me.CB_CB = New System.Windows.Forms.CheckBox()
        Me.CB_AP = New System.Windows.Forms.CheckBox()
        Me.CB_PO = New System.Windows.Forms.CheckBox()
        Me.CB_GL = New System.Windows.Forms.CheckBox()
        Me.CB_AR = New System.Windows.Forms.CheckBox()
        Me.GroupBox16 = New System.Windows.Forms.GroupBox()
        Me.txFormatRunnig = New System.Windows.Forms.MaskedTextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.TAX_CB = New System.Windows.Forms.GroupBox()
        Me.Cb_CB_TAX = New System.Windows.Forms.ComboBox()
        Me.TAX_AP = New System.Windows.Forms.GroupBox()
        Me.Cb_AP_TAX = New System.Windows.Forms.ComboBox()
        Me.TAX_AR = New System.Windows.Forms.GroupBox()
        Me.Cb_AR_TAX = New System.Windows.Forms.ComboBox()
        Me.TAX_GL = New System.Windows.Forms.GroupBox()
        Me.Cb_GL_TAX = New System.Windows.Forms.ComboBox()
        Me.BRANCH_CB = New System.Windows.Forms.GroupBox()
        Me.Cb_CB_BRANCH = New System.Windows.Forms.ComboBox()
        Me.BRANCH_AP = New System.Windows.Forms.GroupBox()
        Me.Cb_AP_BRANCH = New System.Windows.Forms.ComboBox()
        Me.BRANCH_AR = New System.Windows.Forms.GroupBox()
        Me.Cb_AR_BRANCH = New System.Windows.Forms.ComboBox()
        Me.BRANCH_GL = New System.Windows.Forms.GroupBox()
        Me.Cb_GL_BRANCH = New System.Windows.Forms.ComboBox()
        Me.Ch_mS = New System.Windows.Forms.RadioButton()
        Me.Ch_OOr = New System.Windows.Forms.RadioButton()
        Me.CH_Per = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox17 = New System.Windows.Forms.GroupBox()
        Me.chk_docAR = New System.Windows.Forms.CheckBox()
        Me.chk_docAP = New System.Windows.Forms.CheckBox()
        Me.chk_docCB = New System.Windows.Forms.CheckBox()
        Me.chk_docGL = New System.Windows.Forms.CheckBox()
        Me.PnConTax = New System.Windows.Forms.Panel()
        Me.PnDBset = New System.Windows.Forms.Panel()
        Me.VAT = New BaseClass.BaseGroupBox()
        Me.txtVATDBName = New System.Windows.Forms.ComboBox()
        Me.txtVATServerName = New System.Windows.Forms.TextBox()
        Me.txtVATPassword = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtVATUsername = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ACCPAC = New BaseClass.BaseGroupBox()
        Me.txtACCDBName = New System.Windows.Forms.ComboBox()
        Me.txtACCServerName = New System.Windows.Forms.TextBox()
        Me.txtACCPassword = New System.Windows.Forms.TextBox()
        Me.lblServername = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtACCUsername = New System.Windows.Forms.TextBox()
        Me.lblUsername = New System.Windows.Forms.Label()
        Me.PnConfig = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox14.SuspendLayout()
        Me.GroupBox15.SuspendLayout()
        Me.GroupBox16.SuspendLayout()
        Me.TAX_CB.SuspendLayout()
        Me.TAX_AP.SuspendLayout()
        Me.TAX_AR.SuspendLayout()
        Me.TAX_GL.SuspendLayout()
        Me.BRANCH_CB.SuspendLayout()
        Me.BRANCH_AP.SuspendLayout()
        Me.BRANCH_AR.SuspendLayout()
        Me.BRANCH_GL.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox17.SuspendLayout()
        Me.PnConTax.SuspendLayout()
        Me.PnDBset.SuspendLayout()
        Me.VAT.SuspendLayout()
        Me.ACCPAC.SuspendLayout()
        Me.PnConfig.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDBname
        '
        Me.lblDBname.AutoSize = True
        Me.lblDBname.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBname.Location = New System.Drawing.Point(120, 9)
        Me.lblDBname.Name = "lblDBname"
        Me.lblDBname.Size = New System.Drawing.Size(88, 14)
        Me.lblDBname.TabIndex = 50
        Me.lblDBname.Text = "Company Name  "
        '
        'txtCompany
        '
        Me.txtCompany.Enabled = False
        Me.txtCompany.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCompany.Location = New System.Drawing.Point(213, 5)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.Size = New System.Drawing.Size(336, 20)
        Me.txtCompany.TabIndex = 0
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.Color.Transparent
        Me.btnNew.BackgroundImage = Global.ProgramVAT.My.Resources.Resources.bt_nmip_new
        Me.btnNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNew.Location = New System.Drawing.Point(593, 5)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(21, 21)
        Me.btnNew.TabIndex = 3
        Me.btnNew.UseVisualStyleBackColor = False
        '
        'btnP
        '
        Me.btnP.Enabled = False
        Me.btnP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnP.Location = New System.Drawing.Point(555, 5)
        Me.btnP.Name = "btnP"
        Me.btnP.Size = New System.Drawing.Size(19, 21)
        Me.btnP.TabIndex = 1
        Me.btnP.Text = "<"
        Me.btnP.UseVisualStyleBackColor = True
        '
        'btnN
        '
        Me.btnN.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnN.Location = New System.Drawing.Point(574, 5)
        Me.btnN.Name = "btnN"
        Me.btnN.Size = New System.Drawing.Size(19, 21)
        Me.btnN.TabIndex = 2
        Me.btnN.Text = ">"
        Me.btnN.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.Transparent
        Me.btnDelete.BackgroundImage = Global.ProgramVAT.My.Resources.Resources.DeleteRed
        Me.btnDelete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(614, 5)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(21, 21)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.ToolTip1.SetToolTip(Me.btnDelete, "Delete")
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnACCClose
        '
        Me.btnACCClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnACCClose.Image = Global.ProgramVAT.My.Resources.Resources.refresh_button
        Me.btnACCClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnACCClose.Location = New System.Drawing.Point(706, 10)
        Me.btnACCClose.Name = "btnACCClose"
        Me.btnACCClose.Size = New System.Drawing.Size(75, 23)
        Me.btnACCClose.TabIndex = 1
        Me.btnACCClose.Text = "        &Close"
        Me.btnACCClose.UseVisualStyleBackColor = True
        '
        'btnACCSave
        '
        Me.btnACCSave.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.btnACCSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnACCSave.Image = Global.ProgramVAT.My.Resources.Resources.add4
        Me.btnACCSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnACCSave.Location = New System.Drawing.Point(613, 10)
        Me.btnACCSave.Name = "btnACCSave"
        Me.btnACCSave.Size = New System.Drawing.Size(75, 23)
        Me.btnACCSave.TabIndex = 0
        Me.btnACCSave.Text = "Add"
        Me.btnACCSave.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnACCClose)
        Me.Panel1.Controls.Add(Me.btnACCSave)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 523)
        Me.Panel1.MaximumSize = New System.Drawing.Size(809, 38)
        Me.Panel1.MinimumSize = New System.Drawing.Size(809, 38)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(809, 38)
        Me.Panel1.TabIndex = 8
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Ch_AR)
        Me.GroupBox2.Controls.Add(Me.Ch_AP)
        Me.GroupBox2.Controls.Add(Me.Ch_CB)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(11, 1)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(183, 49)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Use Legal Name In Module"
        '
        'Ch_AR
        '
        Me.Ch_AR.AutoSize = True
        Me.Ch_AR.Location = New System.Drawing.Point(113, 21)
        Me.Ch_AR.Name = "Ch_AR"
        Me.Ch_AR.Size = New System.Drawing.Size(44, 18)
        Me.Ch_AR.TabIndex = 2
        Me.Ch_AR.Text = "A/R"
        Me.Ch_AR.UseVisualStyleBackColor = True
        '
        'Ch_AP
        '
        Me.Ch_AP.AutoSize = True
        Me.Ch_AP.Location = New System.Drawing.Point(63, 21)
        Me.Ch_AP.Name = "Ch_AP"
        Me.Ch_AP.Size = New System.Drawing.Size(43, 18)
        Me.Ch_AP.TabIndex = 1
        Me.Ch_AP.Text = "A/P"
        Me.Ch_AP.UseVisualStyleBackColor = True
        '
        'Ch_CB
        '
        Me.Ch_CB.AutoSize = True
        Me.Ch_CB.Location = New System.Drawing.Point(14, 21)
        Me.Ch_CB.Name = "Ch_CB"
        Me.Ch_CB.Size = New System.Drawing.Size(43, 18)
        Me.Ch_CB.TabIndex = 0
        Me.Ch_CB.Text = "C/B"
        Me.Ch_CB.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Ch_Doc)
        Me.GroupBox3.Controls.Add(Me.Ch_PoST)
        Me.GroupBox3.Enabled = False
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(208, 1)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(183, 49)
        Me.GroupBox3.TabIndex = 6
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "DateMode"
        '
        'Ch_Doc
        '
        Me.Ch_Doc.AutoSize = True
        Me.Ch_Doc.Location = New System.Drawing.Point(101, 19)
        Me.Ch_Doc.Name = "Ch_Doc"
        Me.Ch_Doc.Size = New System.Drawing.Size(69, 18)
        Me.Ch_Doc.TabIndex = 1
        Me.Ch_Doc.Text = "Doc Date"
        Me.Ch_Doc.UseVisualStyleBackColor = True
        '
        'Ch_PoST
        '
        Me.Ch_PoST.AutoSize = True
        Me.Ch_PoST.Checked = True
        Me.Ch_PoST.Location = New System.Drawing.Point(16, 20)
        Me.Ch_PoST.Name = "Ch_PoST"
        Me.Ch_PoST.Size = New System.Drawing.Size(71, 18)
        Me.Ch_PoST.TabIndex = 0
        Me.Ch_PoST.TabStop = True
        Me.Ch_PoST.Text = "Post Date"
        Me.Ch_PoST.UseVisualStyleBackColor = True
        '
        'ch_loginN
        '
        Me.ch_loginN.AutoSize = True
        Me.ch_loginN.Checked = True
        Me.ch_loginN.Location = New System.Drawing.Point(66, 19)
        Me.ch_loginN.Name = "ch_loginN"
        Me.ch_loginN.Size = New System.Drawing.Size(38, 18)
        Me.ch_loginN.TabIndex = 1
        Me.ch_loginN.TabStop = True
        Me.ch_loginN.Text = "No"
        Me.ch_loginN.UseVisualStyleBackColor = True
        '
        'ch_loginY
        '
        Me.ch_loginY.AutoSize = True
        Me.ch_loginY.Location = New System.Drawing.Point(13, 19)
        Me.ch_loginY.Name = "ch_loginY"
        Me.ch_loginY.Size = New System.Drawing.Size(44, 18)
        Me.ch_loginY.TabIndex = 0
        Me.ch_loginY.Text = "Yes"
        Me.ch_loginY.UseVisualStyleBackColor = True
        '
        'Ch_runFormat
        '
        Me.Ch_runFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Ch_runFormat.FormattingEnabled = True
        Me.Ch_runFormat.Items.AddRange(New Object() {"0", "00", "000", "0000", "00000", "000000", "0000000"})
        Me.Ch_runFormat.Location = New System.Drawing.Point(16, 18)
        Me.Ch_runFormat.Name = "Ch_runFormat"
        Me.Ch_runFormat.Size = New System.Drawing.Size(105, 22)
        Me.Ch_runFormat.TabIndex = 0
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.ch_loginY)
        Me.GroupBox6.Controls.Add(Me.ch_loginN)
        Me.GroupBox6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(265, 254)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(126, 45)
        Me.GroupBox6.TabIndex = 19
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Security  Login"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Ch_runFormat)
        Me.GroupBox7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox7.Location = New System.Drawing.Point(143, 56)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(139, 45)
        Me.GroupBox7.TabIndex = 8
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Format Running Number"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Rd_Affter)
        Me.GroupBox4.Controls.Add(Me.Rd_Before)
        Me.GroupBox4.Enabled = False
        Me.GroupBox4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(11, 107)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(185, 45)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Get Type"
        '
        'Rd_Affter
        '
        Me.Rd_Affter.AutoSize = True
        Me.Rd_Affter.Checked = True
        Me.Rd_Affter.Location = New System.Drawing.Point(100, 19)
        Me.Rd_Affter.Name = "Rd_Affter"
        Me.Rd_Affter.Size = New System.Drawing.Size(78, 18)
        Me.Rd_Affter.TabIndex = 1
        Me.Rd_Affter.TabStop = True
        Me.Rd_Affter.Text = "Affter Post"
        Me.Rd_Affter.UseVisualStyleBackColor = True
        '
        'Rd_Before
        '
        Me.Rd_Before.AutoSize = True
        Me.Rd_Before.Location = New System.Drawing.Point(14, 19)
        Me.Rd_Before.Name = "Rd_Before"
        Me.Rd_Before.Size = New System.Drawing.Size(82, 18)
        Me.Rd_Before.TabIndex = 0
        Me.Rd_Before.Text = "Before Post"
        Me.Rd_Before.UseVisualStyleBackColor = True
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.CB_PAGE)
        Me.GroupBox9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox9.Location = New System.Drawing.Point(288, 56)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(103, 45)
        Me.GroupBox9.TabIndex = 9
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Row Display"
        '
        'CB_PAGE
        '
        Me.CB_PAGE.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_PAGE.FormattingEnabled = True
        Me.CB_PAGE.Items.AddRange(New Object() {"10", "20", "30", "40", "50", "60", "70", "80", "90", "100", "150", "200", "250", "300", "All"})
        Me.CB_PAGE.Location = New System.Drawing.Point(16, 18)
        Me.CB_PAGE.Name = "CB_PAGE"
        Me.CB_PAGE.Size = New System.Drawing.Size(69, 22)
        Me.CB_PAGE.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.CB_AppY)
        Me.GroupBox5.Controls.Add(Me.CB_AppN)
        Me.GroupBox5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(138, 158)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(121, 45)
        Me.GroupBox5.TabIndex = 13
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Approved"
        '
        'CB_AppY
        '
        Me.CB_AppY.AutoSize = True
        Me.CB_AppY.Location = New System.Drawing.Point(14, 19)
        Me.CB_AppY.Name = "CB_AppY"
        Me.CB_AppY.Size = New System.Drawing.Size(44, 18)
        Me.CB_AppY.TabIndex = 0
        Me.CB_AppY.Text = "Yes"
        Me.CB_AppY.UseVisualStyleBackColor = True
        '
        'CB_AppN
        '
        Me.CB_AppN.AutoSize = True
        Me.CB_AppN.Checked = True
        Me.CB_AppN.Location = New System.Drawing.Point(70, 19)
        Me.CB_AppN.Name = "CB_AppN"
        Me.CB_AppN.Size = New System.Drawing.Size(38, 18)
        Me.CB_AppN.TabIndex = 1
        Me.CB_AppN.TabStop = True
        Me.CB_AppN.Text = "No"
        Me.CB_AppN.UseVisualStyleBackColor = True
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.PO_YES)
        Me.GroupBox10.Controls.Add(Me.PO_NO)
        Me.GroupBox10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox10.Location = New System.Drawing.Point(11, 158)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(121, 45)
        Me.GroupBox10.TabIndex = 12
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "P/O Receipt TAX"
        '
        'PO_YES
        '
        Me.PO_YES.AutoSize = True
        Me.PO_YES.Location = New System.Drawing.Point(14, 20)
        Me.PO_YES.Name = "PO_YES"
        Me.PO_YES.Size = New System.Drawing.Size(44, 18)
        Me.PO_YES.TabIndex = 0
        Me.PO_YES.Text = "Yes"
        Me.PO_YES.UseVisualStyleBackColor = True
        '
        'PO_NO
        '
        Me.PO_NO.AutoSize = True
        Me.PO_NO.Checked = True
        Me.PO_NO.Location = New System.Drawing.Point(69, 19)
        Me.PO_NO.Name = "PO_NO"
        Me.PO_NO.Size = New System.Drawing.Size(38, 18)
        Me.PO_NO.TabIndex = 1
        Me.PO_NO.TabStop = True
        Me.PO_NO.Text = "No"
        Me.PO_NO.UseVisualStyleBackColor = True
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.Ch_CBReverse)
        Me.GroupBox11.Controls.Add(Me.Ch_GLReverse)
        Me.GroupBox11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox11.Location = New System.Drawing.Point(265, 158)
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.Size = New System.Drawing.Size(126, 45)
        Me.GroupBox11.TabIndex = 14
        Me.GroupBox11.TabStop = False
        Me.GroupBox11.Text = "Case Reverse"
        '
        'Ch_CBReverse
        '
        Me.Ch_CBReverse.AutoSize = True
        Me.Ch_CBReverse.Location = New System.Drawing.Point(66, 21)
        Me.Ch_CBReverse.Name = "Ch_CBReverse"
        Me.Ch_CBReverse.Size = New System.Drawing.Size(43, 18)
        Me.Ch_CBReverse.TabIndex = 1
        Me.Ch_CBReverse.Text = "C/B"
        Me.Ch_CBReverse.UseVisualStyleBackColor = True
        '
        'Ch_GLReverse
        '
        Me.Ch_GLReverse.AutoSize = True
        Me.Ch_GLReverse.Location = New System.Drawing.Point(14, 20)
        Me.Ch_GLReverse.Name = "Ch_GLReverse"
        Me.Ch_GLReverse.Size = New System.Drawing.Size(43, 18)
        Me.Ch_GLReverse.TabIndex = 0
        Me.Ch_GLReverse.Text = "G/L"
        Me.Ch_GLReverse.UseVisualStyleBackColor = True
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.cb_OPT_AP)
        Me.GroupBox12.Controls.Add(Me.cb_OPT_CB)
        Me.GroupBox12.Controls.Add(Me.cb_OPT_GL)
        Me.GroupBox12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox12.Location = New System.Drawing.Point(208, 107)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(183, 45)
        Me.GroupBox12.TabIndex = 11
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Optional Field VAT"
        '
        'cb_OPT_AP
        '
        Me.cb_OPT_AP.AutoSize = True
        Me.cb_OPT_AP.Location = New System.Drawing.Point(16, 19)
        Me.cb_OPT_AP.Name = "cb_OPT_AP"
        Me.cb_OPT_AP.Size = New System.Drawing.Size(43, 18)
        Me.cb_OPT_AP.TabIndex = 0
        Me.cb_OPT_AP.Text = "A/P"
        Me.cb_OPT_AP.UseVisualStyleBackColor = True
        '
        'cb_OPT_CB
        '
        Me.cb_OPT_CB.AutoSize = True
        Me.cb_OPT_CB.Location = New System.Drawing.Point(70, 19)
        Me.cb_OPT_CB.Name = "cb_OPT_CB"
        Me.cb_OPT_CB.Size = New System.Drawing.Size(43, 18)
        Me.cb_OPT_CB.TabIndex = 1
        Me.cb_OPT_CB.Text = "C/B"
        Me.cb_OPT_CB.UseVisualStyleBackColor = True
        '
        'cb_OPT_GL
        '
        Me.cb_OPT_GL.AutoSize = True
        Me.cb_OPT_GL.Location = New System.Drawing.Point(123, 19)
        Me.cb_OPT_GL.Name = "cb_OPT_GL"
        Me.cb_OPT_GL.Size = New System.Drawing.Size(43, 18)
        Me.cb_OPT_GL.TabIndex = 2
        Me.cb_OPT_GL.Text = "G/L"
        Me.cb_OPT_GL.UseVisualStyleBackColor = True
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.cb_UPRUN_Y)
        Me.GroupBox8.Controls.Add(Me.cb_UPRUN_N)
        Me.GroupBox8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox8.Location = New System.Drawing.Point(138, 209)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(121, 45)
        Me.GroupBox8.TabIndex = 16
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Update Running"
        '
        'cb_UPRUN_Y
        '
        Me.cb_UPRUN_Y.AutoSize = True
        Me.cb_UPRUN_Y.Location = New System.Drawing.Point(14, 20)
        Me.cb_UPRUN_Y.Name = "cb_UPRUN_Y"
        Me.cb_UPRUN_Y.Size = New System.Drawing.Size(44, 18)
        Me.cb_UPRUN_Y.TabIndex = 0
        Me.cb_UPRUN_Y.Text = "Yes"
        Me.cb_UPRUN_Y.UseVisualStyleBackColor = True
        '
        'cb_UPRUN_N
        '
        Me.cb_UPRUN_N.AutoSize = True
        Me.cb_UPRUN_N.Checked = True
        Me.cb_UPRUN_N.Location = New System.Drawing.Point(70, 20)
        Me.cb_UPRUN_N.Name = "cb_UPRUN_N"
        Me.cb_UPRUN_N.Size = New System.Drawing.Size(38, 18)
        Me.cb_UPRUN_N.TabIndex = 1
        Me.cb_UPRUN_N.TabStop = True
        Me.cb_UPRUN_N.Text = "No"
        Me.cb_UPRUN_N.UseVisualStyleBackColor = True
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.cb_RERUN_Y)
        Me.GroupBox13.Controls.Add(Me.cb_RERUN_N)
        Me.GroupBox13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox13.Location = New System.Drawing.Point(11, 209)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Size = New System.Drawing.Size(121, 45)
        Me.GroupBox13.TabIndex = 15
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Remember Running"
        '
        'cb_RERUN_Y
        '
        Me.cb_RERUN_Y.AutoSize = True
        Me.cb_RERUN_Y.Location = New System.Drawing.Point(14, 20)
        Me.cb_RERUN_Y.Name = "cb_RERUN_Y"
        Me.cb_RERUN_Y.Size = New System.Drawing.Size(44, 18)
        Me.cb_RERUN_Y.TabIndex = 0
        Me.cb_RERUN_Y.Text = "Yes"
        Me.cb_RERUN_Y.UseVisualStyleBackColor = True
        '
        'cb_RERUN_N
        '
        Me.cb_RERUN_N.AutoSize = True
        Me.cb_RERUN_N.Checked = True
        Me.cb_RERUN_N.Location = New System.Drawing.Point(69, 20)
        Me.cb_RERUN_N.Name = "cb_RERUN_N"
        Me.cb_RERUN_N.Size = New System.Drawing.Size(38, 18)
        Me.cb_RERUN_N.TabIndex = 1
        Me.cb_RERUN_N.TabStop = True
        Me.cb_RERUN_N.Text = "No"
        Me.cb_RERUN_N.UseVisualStyleBackColor = True
        '
        'GroupBox14
        '
        Me.GroupBox14.Controls.Add(Me.Rb_AutoRunning)
        Me.GroupBox14.Controls.Add(Me.Rb_ManualRunning)
        Me.GroupBox14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox14.Location = New System.Drawing.Point(265, 209)
        Me.GroupBox14.Name = "GroupBox14"
        Me.GroupBox14.Size = New System.Drawing.Size(126, 45)
        Me.GroupBox14.TabIndex = 17
        Me.GroupBox14.TabStop = False
        Me.GroupBox14.Text = "Running"
        '
        'Rb_AutoRunning
        '
        Me.Rb_AutoRunning.AutoSize = True
        Me.Rb_AutoRunning.Checked = True
        Me.Rb_AutoRunning.Location = New System.Drawing.Point(14, 20)
        Me.Rb_AutoRunning.Name = "Rb_AutoRunning"
        Me.Rb_AutoRunning.Size = New System.Drawing.Size(48, 18)
        Me.Rb_AutoRunning.TabIndex = 0
        Me.Rb_AutoRunning.TabStop = True
        Me.Rb_AutoRunning.Text = "Auto"
        Me.Rb_AutoRunning.UseVisualStyleBackColor = True
        '
        'Rb_ManualRunning
        '
        Me.Rb_ManualRunning.AutoSize = True
        Me.Rb_ManualRunning.Location = New System.Drawing.Point(66, 20)
        Me.Rb_ManualRunning.Name = "Rb_ManualRunning"
        Me.Rb_ManualRunning.Size = New System.Drawing.Size(59, 18)
        Me.Rb_ManualRunning.TabIndex = 1
        Me.Rb_ManualRunning.Text = "Manual"
        Me.Rb_ManualRunning.UseVisualStyleBackColor = True
        '
        'GroupBox15
        '
        Me.GroupBox15.Controls.Add(Me.CB_OE)
        Me.GroupBox15.Controls.Add(Me.CB_CB)
        Me.GroupBox15.Controls.Add(Me.CB_AP)
        Me.GroupBox15.Controls.Add(Me.CB_PO)
        Me.GroupBox15.Controls.Add(Me.CB_GL)
        Me.GroupBox15.Controls.Add(Me.CB_AR)
        Me.GroupBox15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox15.Location = New System.Drawing.Point(11, 302)
        Me.GroupBox15.Name = "GroupBox15"
        Me.GroupBox15.Size = New System.Drawing.Size(380, 77)
        Me.GroupBox15.TabIndex = 20
        Me.GroupBox15.TabStop = False
        Me.GroupBox15.Text = "Show Module"
        '
        'CB_OE
        '
        Me.CB_OE.AutoSize = True
        Me.CB_OE.Location = New System.Drawing.Point(213, 36)
        Me.CB_OE.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.CB_OE.Name = "CB_OE"
        Me.CB_OE.Size = New System.Drawing.Size(110, 18)
        Me.CB_OE.TabIndex = 3
        Me.CB_OE.Text = "Order Entry (O/E)"
        Me.CB_OE.UseVisualStyleBackColor = True
        '
        'CB_CB
        '
        Me.CB_CB.AutoSize = True
        Me.CB_CB.Location = New System.Drawing.Point(213, 55)
        Me.CB_CB.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.CB_CB.Name = "CB_CB"
        Me.CB_CB.Size = New System.Drawing.Size(103, 18)
        Me.CB_CB.TabIndex = 5
        Me.CB_CB.Text = "CashBook (C/B)"
        Me.CB_CB.UseVisualStyleBackColor = True
        '
        'CB_AP
        '
        Me.CB_AP.AutoSize = True
        Me.CB_AP.Location = New System.Drawing.Point(17, 17)
        Me.CB_AP.Margin = New System.Windows.Forms.Padding(1, 1, 1, 0)
        Me.CB_AP.Name = "CB_AP"
        Me.CB_AP.Size = New System.Drawing.Size(142, 18)
        Me.CB_AP.TabIndex = 0
        Me.CB_AP.Text = "Accounts Payable (A/P)"
        Me.CB_AP.UseVisualStyleBackColor = True
        '
        'CB_PO
        '
        Me.CB_PO.AutoSize = True
        Me.CB_PO.Location = New System.Drawing.Point(213, 17)
        Me.CB_PO.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.CB_PO.Name = "CB_PO"
        Me.CB_PO.Size = New System.Drawing.Size(137, 18)
        Me.CB_PO.TabIndex = 1
        Me.CB_PO.Text = "Purchase Orders (P/O)"
        Me.CB_PO.UseVisualStyleBackColor = True
        '
        'CB_GL
        '
        Me.CB_GL.AutoSize = True
        Me.CB_GL.Location = New System.Drawing.Point(17, 55)
        Me.CB_GL.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.CB_GL.Name = "CB_GL"
        Me.CB_GL.Size = New System.Drawing.Size(129, 18)
        Me.CB_GL.TabIndex = 4
        Me.CB_GL.Text = "General Ledger (G/L)"
        Me.CB_GL.UseVisualStyleBackColor = True
        '
        'CB_AR
        '
        Me.CB_AR.AutoSize = True
        Me.CB_AR.Location = New System.Drawing.Point(17, 36)
        Me.CB_AR.Margin = New System.Windows.Forms.Padding(1, 0, 1, 0)
        Me.CB_AR.Name = "CB_AR"
        Me.CB_AR.Size = New System.Drawing.Size(158, 18)
        Me.CB_AR.TabIndex = 2
        Me.CB_AR.Text = "Accounts Receivable (A/R)"
        Me.CB_AR.UseVisualStyleBackColor = True
        '
        'GroupBox16
        '
        Me.GroupBox16.Controls.Add(Me.txFormatRunnig)
        Me.GroupBox16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox16.Location = New System.Drawing.Point(11, 56)
        Me.GroupBox16.Name = "GroupBox16"
        Me.GroupBox16.Size = New System.Drawing.Size(126, 45)
        Me.GroupBox16.TabIndex = 7
        Me.GroupBox16.TabStop = False
        Me.GroupBox16.Text = "Format Running"
        '
        'txFormatRunnig
        '
        Me.txFormatRunnig.Location = New System.Drawing.Point(17, 17)
        Me.txFormatRunnig.Name = "txFormatRunnig"
        Me.txFormatRunnig.Size = New System.Drawing.Size(90, 20)
        Me.txFormatRunnig.TabIndex = 0
        Me.txFormatRunnig.Text = "yy/MM/"
        '
        'btnLogin
        '
        Me.btnLogin.BackgroundImage = Global.ProgramVAT.My.Resources.Resources.log_in
        Me.btnLogin.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLogin.Location = New System.Drawing.Point(5, 5)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(21, 21)
        Me.btnLogin.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.btnLogin, "Log in for FMS")
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'TAX_CB
        '
        Me.TAX_CB.Controls.Add(Me.Cb_CB_TAX)
        Me.TAX_CB.Location = New System.Drawing.Point(13, 8)
        Me.TAX_CB.Name = "TAX_CB"
        Me.TAX_CB.Size = New System.Drawing.Size(140, 45)
        Me.TAX_CB.TabIndex = 21
        Me.TAX_CB.TabStop = False
        Me.TAX_CB.Text = "Cash Book TaxID From"
        '
        'Cb_CB_TAX
        '
        Me.Cb_CB_TAX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_CB_TAX.FormattingEnabled = True
        Me.Cb_CB_TAX.Items.AddRange(New Object() {"URL", "E-mail", "State/Province", "Country", "Business Reg. No"})
        Me.Cb_CB_TAX.Location = New System.Drawing.Point(7, 16)
        Me.Cb_CB_TAX.Name = "Cb_CB_TAX"
        Me.Cb_CB_TAX.Size = New System.Drawing.Size(125, 22)
        Me.Cb_CB_TAX.TabIndex = 0
        '
        'TAX_AP
        '
        Me.TAX_AP.Controls.Add(Me.Cb_AP_TAX)
        Me.TAX_AP.Location = New System.Drawing.Point(168, 8)
        Me.TAX_AP.Name = "TAX_AP"
        Me.TAX_AP.Size = New System.Drawing.Size(140, 45)
        Me.TAX_AP.TabIndex = 23
        Me.TAX_AP.TabStop = False
        Me.TAX_AP.Text = "AP TaxID From"
        '
        'Cb_AP_TAX
        '
        Me.Cb_AP_TAX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_AP_TAX.FormattingEnabled = True
        Me.Cb_AP_TAX.Items.AddRange(New Object() {"Web Site", "E-mail", "Optional Field", "Registration Number", "Business Reg. No"})
        Me.Cb_AP_TAX.Location = New System.Drawing.Point(7, 16)
        Me.Cb_AP_TAX.Name = "Cb_AP_TAX"
        Me.Cb_AP_TAX.Size = New System.Drawing.Size(125, 22)
        Me.Cb_AP_TAX.TabIndex = 0
        '
        'TAX_AR
        '
        Me.TAX_AR.Controls.Add(Me.Cb_AR_TAX)
        Me.TAX_AR.Location = New System.Drawing.Point(324, 8)
        Me.TAX_AR.Name = "TAX_AR"
        Me.TAX_AR.Size = New System.Drawing.Size(140, 45)
        Me.TAX_AR.TabIndex = 25
        Me.TAX_AR.TabStop = False
        Me.TAX_AR.Text = "AR TaxID From"
        '
        'Cb_AR_TAX
        '
        Me.Cb_AR_TAX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_AR_TAX.FormattingEnabled = True
        Me.Cb_AR_TAX.Items.AddRange(New Object() {"Web Site", "E-mail", "Optional Field", "Registration Number", "Business Reg. No"})
        Me.Cb_AR_TAX.Location = New System.Drawing.Point(7, 16)
        Me.Cb_AR_TAX.Name = "Cb_AR_TAX"
        Me.Cb_AR_TAX.Size = New System.Drawing.Size(125, 22)
        Me.Cb_AR_TAX.TabIndex = 0
        '
        'TAX_GL
        '
        Me.TAX_GL.Controls.Add(Me.Cb_GL_TAX)
        Me.TAX_GL.Location = New System.Drawing.Point(480, 8)
        Me.TAX_GL.Name = "TAX_GL"
        Me.TAX_GL.Size = New System.Drawing.Size(187, 45)
        Me.TAX_GL.TabIndex = 27
        Me.TAX_GL.TabStop = False
        Me.TAX_GL.Text = "GL TaxID From"
        '
        'Cb_GL_TAX
        '
        Me.Cb_GL_TAX.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_GL_TAX.FormattingEnabled = True
        Me.Cb_GL_TAX.Items.AddRange(New Object() {"A/P Vendor ,A/R Customer", "Optional Field", "Comment Comma 4", "Comment Comma 5"})
        Me.Cb_GL_TAX.Location = New System.Drawing.Point(10, 16)
        Me.Cb_GL_TAX.Name = "Cb_GL_TAX"
        Me.Cb_GL_TAX.Size = New System.Drawing.Size(171, 22)
        Me.Cb_GL_TAX.TabIndex = 0
        '
        'BRANCH_CB
        '
        Me.BRANCH_CB.Controls.Add(Me.Cb_CB_BRANCH)
        Me.BRANCH_CB.Location = New System.Drawing.Point(13, 59)
        Me.BRANCH_CB.Name = "BRANCH_CB"
        Me.BRANCH_CB.Size = New System.Drawing.Size(140, 45)
        Me.BRANCH_CB.TabIndex = 22
        Me.BRANCH_CB.TabStop = False
        Me.BRANCH_CB.Text = "Cash Book Branch From"
        '
        'Cb_CB_BRANCH
        '
        Me.Cb_CB_BRANCH.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_CB_BRANCH.FormattingEnabled = True
        Me.Cb_CB_BRANCH.Items.AddRange(New Object() {"URL", "E-mail", "State/Province", "Country"})
        Me.Cb_CB_BRANCH.Location = New System.Drawing.Point(7, 16)
        Me.Cb_CB_BRANCH.Name = "Cb_CB_BRANCH"
        Me.Cb_CB_BRANCH.Size = New System.Drawing.Size(125, 22)
        Me.Cb_CB_BRANCH.TabIndex = 0
        '
        'BRANCH_AP
        '
        Me.BRANCH_AP.Controls.Add(Me.Cb_AP_BRANCH)
        Me.BRANCH_AP.Location = New System.Drawing.Point(168, 59)
        Me.BRANCH_AP.Name = "BRANCH_AP"
        Me.BRANCH_AP.Size = New System.Drawing.Size(140, 45)
        Me.BRANCH_AP.TabIndex = 24
        Me.BRANCH_AP.TabStop = False
        Me.BRANCH_AP.Text = "AP Branch From"
        '
        'Cb_AP_BRANCH
        '
        Me.Cb_AP_BRANCH.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_AP_BRANCH.FormattingEnabled = True
        Me.Cb_AP_BRANCH.Items.AddRange(New Object() {"Web Site", "E-mail", "Optional Field", "Registration Number", "Business Reg. No"})
        Me.Cb_AP_BRANCH.Location = New System.Drawing.Point(7, 16)
        Me.Cb_AP_BRANCH.Name = "Cb_AP_BRANCH"
        Me.Cb_AP_BRANCH.Size = New System.Drawing.Size(125, 22)
        Me.Cb_AP_BRANCH.TabIndex = 0
        '
        'BRANCH_AR
        '
        Me.BRANCH_AR.Controls.Add(Me.Cb_AR_BRANCH)
        Me.BRANCH_AR.Location = New System.Drawing.Point(324, 59)
        Me.BRANCH_AR.Name = "BRANCH_AR"
        Me.BRANCH_AR.Size = New System.Drawing.Size(140, 45)
        Me.BRANCH_AR.TabIndex = 26
        Me.BRANCH_AR.TabStop = False
        Me.BRANCH_AR.Text = "AR Branch From"
        '
        'Cb_AR_BRANCH
        '
        Me.Cb_AR_BRANCH.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_AR_BRANCH.FormattingEnabled = True
        Me.Cb_AR_BRANCH.Items.AddRange(New Object() {"Web Site", "E-mail", "Optional Field", "Registration Number", "Business Reg. No"})
        Me.Cb_AR_BRANCH.Location = New System.Drawing.Point(7, 16)
        Me.Cb_AR_BRANCH.Name = "Cb_AR_BRANCH"
        Me.Cb_AR_BRANCH.Size = New System.Drawing.Size(125, 22)
        Me.Cb_AR_BRANCH.TabIndex = 0
        '
        'BRANCH_GL
        '
        Me.BRANCH_GL.Controls.Add(Me.Cb_GL_BRANCH)
        Me.BRANCH_GL.Location = New System.Drawing.Point(480, 59)
        Me.BRANCH_GL.Name = "BRANCH_GL"
        Me.BRANCH_GL.Size = New System.Drawing.Size(187, 45)
        Me.BRANCH_GL.TabIndex = 28
        Me.BRANCH_GL.TabStop = False
        Me.BRANCH_GL.Text = "GL Branch From"
        '
        'Cb_GL_BRANCH
        '
        Me.Cb_GL_BRANCH.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Cb_GL_BRANCH.FormattingEnabled = True
        Me.Cb_GL_BRANCH.Items.AddRange(New Object() {"A/P Vendor ,A/R Customer", "Optional Field", "Comment Comma 4", "Comment Comma 5"})
        Me.Cb_GL_BRANCH.Location = New System.Drawing.Point(10, 16)
        Me.Cb_GL_BRANCH.Name = "Cb_GL_BRANCH"
        Me.Cb_GL_BRANCH.Size = New System.Drawing.Size(171, 22)
        Me.Cb_GL_BRANCH.TabIndex = 0
        '
        'Ch_mS
        '
        Me.Ch_mS.AutoSize = True
        Me.Ch_mS.Checked = True
        Me.Ch_mS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ch_mS.Location = New System.Drawing.Point(193, 20)
        Me.Ch_mS.Name = "Ch_mS"
        Me.Ch_mS.Size = New System.Drawing.Size(64, 18)
        Me.Ch_mS.TabIndex = 1
        Me.Ch_mS.TabStop = True
        Me.Ch_mS.Text = "MS SQL"
        Me.Ch_mS.UseVisualStyleBackColor = True
        '
        'Ch_OOr
        '
        Me.Ch_OOr.AutoSize = True
        Me.Ch_OOr.Enabled = False
        Me.Ch_OOr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ch_OOr.Location = New System.Drawing.Point(318, 19)
        Me.Ch_OOr.Name = "Ch_OOr"
        Me.Ch_OOr.Size = New System.Drawing.Size(57, 18)
        Me.Ch_OOr.TabIndex = 2
        Me.Ch_OOr.Text = "Oracle"
        Me.Ch_OOr.UseVisualStyleBackColor = True
        Me.Ch_OOr.Visible = False
        '
        'CH_Per
        '
        Me.CH_Per.AutoSize = True
        Me.CH_Per.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CH_Per.Location = New System.Drawing.Point(55, 19)
        Me.CH_Per.Name = "CH_Per"
        Me.CH_Per.Size = New System.Drawing.Size(73, 18)
        Me.CH_Per.TabIndex = 0
        Me.CH_Per.Text = "Pervasive"
        Me.CH_Per.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CH_Per)
        Me.GroupBox1.Controls.Add(Me.Ch_OOr)
        Me.GroupBox1.Controls.Add(Me.Ch_mS)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(380, 49)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Database Type"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.ProgramVAT.My.Resources.Resources.FMS_Logo3
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.Location = New System.Drawing.Point(430, 129)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(328, 151)
        Me.PictureBox1.TabIndex = 51
        Me.PictureBox1.TabStop = False
        '
        'GroupBox17
        '
        Me.GroupBox17.Controls.Add(Me.chk_docAR)
        Me.GroupBox17.Controls.Add(Me.chk_docAP)
        Me.GroupBox17.Controls.Add(Me.chk_docCB)
        Me.GroupBox17.Controls.Add(Me.chk_docGL)
        Me.GroupBox17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox17.Location = New System.Drawing.Point(11, 254)
        Me.GroupBox17.Name = "GroupBox17"
        Me.GroupBox17.Size = New System.Drawing.Size(248, 45)
        Me.GroupBox17.TabIndex = 18
        Me.GroupBox17.TabStop = False
        Me.GroupBox17.Text = "Doc No"
        '
        'chk_docAR
        '
        Me.chk_docAR.AutoSize = True
        Me.chk_docAR.Location = New System.Drawing.Point(69, 19)
        Me.chk_docAR.Name = "chk_docAR"
        Me.chk_docAR.Size = New System.Drawing.Size(44, 18)
        Me.chk_docAR.TabIndex = 1
        Me.chk_docAR.Text = "A/R"
        Me.chk_docAR.UseVisualStyleBackColor = True
        '
        'chk_docAP
        '
        Me.chk_docAP.AutoSize = True
        Me.chk_docAP.Location = New System.Drawing.Point(14, 19)
        Me.chk_docAP.Name = "chk_docAP"
        Me.chk_docAP.Size = New System.Drawing.Size(43, 18)
        Me.chk_docAP.TabIndex = 0
        Me.chk_docAP.Text = "A/P"
        Me.chk_docAP.UseVisualStyleBackColor = True
        '
        'chk_docCB
        '
        Me.chk_docCB.AutoSize = True
        Me.chk_docCB.Location = New System.Drawing.Point(127, 19)
        Me.chk_docCB.Name = "chk_docCB"
        Me.chk_docCB.Size = New System.Drawing.Size(43, 18)
        Me.chk_docCB.TabIndex = 2
        Me.chk_docCB.Text = "C/B"
        Me.chk_docCB.UseVisualStyleBackColor = True
        '
        'chk_docGL
        '
        Me.chk_docGL.AutoSize = True
        Me.chk_docGL.Location = New System.Drawing.Point(181, 19)
        Me.chk_docGL.Name = "chk_docGL"
        Me.chk_docGL.Size = New System.Drawing.Size(43, 18)
        Me.chk_docGL.TabIndex = 3
        Me.chk_docGL.Text = "G/L"
        Me.chk_docGL.UseVisualStyleBackColor = True
        '
        'PnConTax
        '
        Me.PnConTax.Controls.Add(Me.BRANCH_GL)
        Me.PnConTax.Controls.Add(Me.TAX_GL)
        Me.PnConTax.Controls.Add(Me.BRANCH_AR)
        Me.PnConTax.Controls.Add(Me.TAX_AR)
        Me.PnConTax.Controls.Add(Me.BRANCH_AP)
        Me.PnConTax.Controls.Add(Me.TAX_AP)
        Me.PnConTax.Controls.Add(Me.BRANCH_CB)
        Me.PnConTax.Controls.Add(Me.TAX_CB)
        Me.PnConTax.Location = New System.Drawing.Point(3, 409)
        Me.PnConTax.Name = "PnConTax"
        Me.PnConTax.Size = New System.Drawing.Size(806, 128)
        Me.PnConTax.TabIndex = 6
        '
        'PnDBset
        '
        Me.PnDBset.Controls.Add(Me.GroupBox1)
        Me.PnDBset.Controls.Add(Me.VAT)
        Me.PnDBset.Controls.Add(Me.ACCPAC)
        Me.PnDBset.Location = New System.Drawing.Point(405, 32)
        Me.PnDBset.Name = "PnDBset"
        Me.PnDBset.Size = New System.Drawing.Size(404, 385)
        Me.PnDBset.TabIndex = 7
        '
        'VAT
        '
        Me.VAT.BackColor = System.Drawing.Color.Transparent
        Me.VAT.Controls.Add(Me.txtVATDBName)
        Me.VAT.Controls.Add(Me.txtVATServerName)
        Me.VAT.Controls.Add(Me.txtVATPassword)
        Me.VAT.Controls.Add(Me.Label7)
        Me.VAT.Controls.Add(Me.Label8)
        Me.VAT.Controls.Add(Me.Label9)
        Me.VAT.Controls.Add(Me.txtVATUsername)
        Me.VAT.Controls.Add(Me.Label10)
        Me.VAT.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.VAT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VAT.ForeColor = System.Drawing.Color.Black
        Me.VAT.Location = New System.Drawing.Point(9, 211)
        Me.VAT.Name = "VAT"
        Me.VAT.Size = New System.Drawing.Size(380, 170)
        Me.VAT.TabIndex = 32
        Me.VAT.TabStop = False
        Me.VAT.Text = "Database(VAT)"
        '
        'txtVATDBName
        '
        Me.txtVATDBName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtVATDBName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.txtVATDBName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVATDBName.FormattingEnabled = True
        Me.txtVATDBName.Location = New System.Drawing.Point(132, 52)
        Me.txtVATDBName.Name = "txtVATDBName"
        Me.txtVATDBName.Size = New System.Drawing.Size(168, 22)
        Me.txtVATDBName.TabIndex = 3
        '
        'txtVATServerName
        '
        Me.txtVATServerName.Enabled = False
        Me.txtVATServerName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVATServerName.Location = New System.Drawing.Point(132, 27)
        Me.txtVATServerName.Name = "txtVATServerName"
        Me.txtVATServerName.Size = New System.Drawing.Size(168, 20)
        Me.txtVATServerName.TabIndex = 0
        '
        'txtVATPassword
        '
        Me.txtVATPassword.AccessibleName = ""
        Me.txtVATPassword.Enabled = False
        Me.txtVATPassword.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVATPassword.Location = New System.Drawing.Point(132, 107)
        Me.txtVATPassword.Name = "txtVATPassword"
        Me.txtVATPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtVATPassword.Size = New System.Drawing.Size(168, 20)
        Me.txtVATPassword.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(52, 30)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 14)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Server "
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(52, 110)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 14)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "Password "
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(52, 56)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 14)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "Database "
        '
        'txtVATUsername
        '
        Me.txtVATUsername.Enabled = False
        Me.txtVATUsername.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVATUsername.Location = New System.Drawing.Point(132, 80)
        Me.txtVATUsername.Name = "txtVATUsername"
        Me.txtVATUsername.Size = New System.Drawing.Size(168, 20)
        Me.txtVATUsername.TabIndex = 1
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(52, 83)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(63, 14)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "User Name "
        '
        'ACCPAC
        '
        Me.ACCPAC.BackColor = System.Drawing.Color.Transparent
        Me.ACCPAC.Controls.Add(Me.txtACCDBName)
        Me.ACCPAC.Controls.Add(Me.txtACCServerName)
        Me.ACCPAC.Controls.Add(Me.txtACCPassword)
        Me.ACCPAC.Controls.Add(Me.lblServername)
        Me.ACCPAC.Controls.Add(Me.Label5)
        Me.ACCPAC.Controls.Add(Me.Label4)
        Me.ACCPAC.Controls.Add(Me.txtACCUsername)
        Me.ACCPAC.Controls.Add(Me.lblUsername)
        Me.ACCPAC.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ACCPAC.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ACCPAC.ForeColor = System.Drawing.Color.Black
        Me.ACCPAC.Location = New System.Drawing.Point(9, 58)
        Me.ACCPAC.MaximumSize = New System.Drawing.Size(380, 160)
        Me.ACCPAC.MinimumSize = New System.Drawing.Size(380, 25)
        Me.ACCPAC.Name = "ACCPAC"
        Me.ACCPAC.Size = New System.Drawing.Size(380, 147)
        Me.ACCPAC.TabIndex = 31
        Me.ACCPAC.TabStop = False
        Me.ACCPAC.Text = "Database(Accpac)"
        '
        'txtACCDBName
        '
        Me.txtACCDBName.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtACCDBName.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.txtACCDBName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACCDBName.FormattingEnabled = True
        Me.txtACCDBName.Location = New System.Drawing.Point(132, 52)
        Me.txtACCDBName.Name = "txtACCDBName"
        Me.txtACCDBName.Size = New System.Drawing.Size(168, 22)
        Me.txtACCDBName.TabIndex = 3
        '
        'txtACCServerName
        '
        Me.txtACCServerName.Enabled = False
        Me.txtACCServerName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACCServerName.Location = New System.Drawing.Point(132, 27)
        Me.txtACCServerName.Name = "txtACCServerName"
        Me.txtACCServerName.Size = New System.Drawing.Size(168, 20)
        Me.txtACCServerName.TabIndex = 0
        '
        'txtACCPassword
        '
        Me.txtACCPassword.AccessibleName = ""
        Me.txtACCPassword.Enabled = False
        Me.txtACCPassword.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACCPassword.Location = New System.Drawing.Point(132, 107)
        Me.txtACCPassword.Name = "txtACCPassword"
        Me.txtACCPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtACCPassword.Size = New System.Drawing.Size(168, 20)
        Me.txtACCPassword.TabIndex = 2
        '
        'lblServername
        '
        Me.lblServername.AutoSize = True
        Me.lblServername.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServername.Location = New System.Drawing.Point(52, 30)
        Me.lblServername.Name = "lblServername"
        Me.lblServername.Size = New System.Drawing.Size(43, 14)
        Me.lblServername.TabIndex = 8
        Me.lblServername.Text = "Server "
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(52, 110)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 14)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Password "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(52, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 14)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Database "
        '
        'txtACCUsername
        '
        Me.txtACCUsername.Enabled = False
        Me.txtACCUsername.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACCUsername.Location = New System.Drawing.Point(132, 80)
        Me.txtACCUsername.Name = "txtACCUsername"
        Me.txtACCUsername.Size = New System.Drawing.Size(168, 20)
        Me.txtACCUsername.TabIndex = 1
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUsername.Location = New System.Drawing.Point(52, 83)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(63, 14)
        Me.lblUsername.TabIndex = 13
        Me.lblUsername.Text = "User Name "
        '
        'PnConfig
        '
        Me.PnConfig.Controls.Add(Me.GroupBox17)
        Me.PnConfig.Controls.Add(Me.GroupBox16)
        Me.PnConfig.Controls.Add(Me.GroupBox15)
        Me.PnConfig.Controls.Add(Me.GroupBox14)
        Me.PnConfig.Controls.Add(Me.GroupBox13)
        Me.PnConfig.Controls.Add(Me.GroupBox8)
        Me.PnConfig.Controls.Add(Me.GroupBox6)
        Me.PnConfig.Controls.Add(Me.GroupBox12)
        Me.PnConfig.Controls.Add(Me.GroupBox4)
        Me.PnConfig.Controls.Add(Me.GroupBox9)
        Me.PnConfig.Controls.Add(Me.GroupBox10)
        Me.PnConfig.Controls.Add(Me.GroupBox5)
        Me.PnConfig.Controls.Add(Me.GroupBox7)
        Me.PnConfig.Controls.Add(Me.GroupBox3)
        Me.PnConfig.Controls.Add(Me.GroupBox11)
        Me.PnConfig.Controls.Add(Me.GroupBox2)
        Me.PnConfig.Location = New System.Drawing.Point(5, 32)
        Me.PnConfig.Name = "PnConfig"
        Me.PnConfig.Size = New System.Drawing.Size(399, 385)
        Me.PnConfig.TabIndex = 5
        '
        'FrmConfigDB
        '
        Me.AcceptButton = Me.btnACCSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(799, 561)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnN)
        Me.Controls.Add(Me.btnP)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.lblDBname)
        Me.Controls.Add(Me.txtCompany)
        Me.Controls.Add(Me.PnDBset)
        Me.Controls.Add(Me.PnConfig)
        Me.Controls.Add(Me.PnConTax)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(815, 600)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(815, 600)
        Me.Name = "FrmConfigDB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Database  Setup"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox10.PerformLayout()
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox12.PerformLayout()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox13.PerformLayout()
        Me.GroupBox14.ResumeLayout(False)
        Me.GroupBox14.PerformLayout()
        Me.GroupBox15.ResumeLayout(False)
        Me.GroupBox15.PerformLayout()
        Me.GroupBox16.ResumeLayout(False)
        Me.GroupBox16.PerformLayout()
        Me.TAX_CB.ResumeLayout(False)
        Me.TAX_AP.ResumeLayout(False)
        Me.TAX_AR.ResumeLayout(False)
        Me.TAX_GL.ResumeLayout(False)
        Me.BRANCH_CB.ResumeLayout(False)
        Me.BRANCH_AP.ResumeLayout(False)
        Me.BRANCH_AR.ResumeLayout(False)
        Me.BRANCH_GL.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox17.ResumeLayout(False)
        Me.GroupBox17.PerformLayout()
        Me.PnConTax.ResumeLayout(False)
        Me.PnDBset.ResumeLayout(False)
        Me.VAT.ResumeLayout(False)
        Me.VAT.PerformLayout()
        Me.ACCPAC.ResumeLayout(False)
        Me.ACCPAC.PerformLayout()
        Me.PnConfig.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ACCPAC As BaseClass.BaseGroupBox
    Friend WithEvents btnACCClose As System.Windows.Forms.Button
    Friend WithEvents txtACCServerName As System.Windows.Forms.TextBox
    Friend WithEvents txtACCPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblServername As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtACCUsername As System.Windows.Forms.TextBox
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents lblDBname As System.Windows.Forms.Label
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnP As System.Windows.Forms.Button
    Friend WithEvents btnN As System.Windows.Forms.Button
    Friend WithEvents VAT As BaseClass.BaseGroupBox
    Friend WithEvents txtVATServerName As System.Windows.Forms.TextBox
    Friend WithEvents txtVATPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtVATUsername As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnACCSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Ch_AP As System.Windows.Forms.CheckBox
    Friend WithEvents Ch_CB As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Ch_Doc As System.Windows.Forms.RadioButton
    Friend WithEvents Ch_PoST As System.Windows.Forms.RadioButton
    Friend WithEvents Ch_runFormat As System.Windows.Forms.ComboBox
    Friend WithEvents ch_loginN As System.Windows.Forms.RadioButton
    Friend WithEvents ch_loginY As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Rd_Affter As System.Windows.Forms.RadioButton
    Friend WithEvents Rd_Before As System.Windows.Forms.RadioButton
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents CB_PAGE As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents CB_AppY As System.Windows.Forms.RadioButton
    Friend WithEvents CB_AppN As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents PO_YES As System.Windows.Forms.RadioButton
    Friend WithEvents PO_NO As System.Windows.Forms.RadioButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox11 As System.Windows.Forms.GroupBox
    Friend WithEvents Ch_CBReverse As System.Windows.Forms.CheckBox
    Friend WithEvents Ch_GLReverse As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox12 As System.Windows.Forms.GroupBox
    Friend WithEvents cb_OPT_CB As System.Windows.Forms.CheckBox
    Friend WithEvents cb_OPT_GL As System.Windows.Forms.CheckBox
    Friend WithEvents cb_OPT_AP As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents cb_UPRUN_Y As System.Windows.Forms.RadioButton
    Friend WithEvents cb_UPRUN_N As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox13 As System.Windows.Forms.GroupBox
    Friend WithEvents cb_RERUN_Y As System.Windows.Forms.RadioButton
    Friend WithEvents cb_RERUN_N As System.Windows.Forms.RadioButton
    Friend WithEvents txtVATDBName As System.Windows.Forms.ComboBox
    Friend WithEvents txtACCDBName As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox14 As System.Windows.Forms.GroupBox
    Friend WithEvents Rb_AutoRunning As System.Windows.Forms.RadioButton
    Friend WithEvents Rb_ManualRunning As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox15 As System.Windows.Forms.GroupBox
    Friend WithEvents CB_CB As System.Windows.Forms.CheckBox
    Friend WithEvents CB_AP As System.Windows.Forms.CheckBox
    Friend WithEvents CB_GL As System.Windows.Forms.CheckBox
    Friend WithEvents CB_AR As System.Windows.Forms.CheckBox
    Friend WithEvents CB_OE As System.Windows.Forms.CheckBox
    Friend WithEvents CB_PO As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox16 As System.Windows.Forms.GroupBox
    Friend WithEvents txFormatRunnig As System.Windows.Forms.MaskedTextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TAX_CB As System.Windows.Forms.GroupBox
    Friend WithEvents TAX_AP As System.Windows.Forms.GroupBox
    Friend WithEvents TAX_AR As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_CB_TAX As System.Windows.Forms.ComboBox
    Friend WithEvents Cb_AP_TAX As System.Windows.Forms.ComboBox
    Friend WithEvents Cb_AR_TAX As System.Windows.Forms.ComboBox
    Friend WithEvents TAX_GL As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_GL_TAX As System.Windows.Forms.ComboBox
    Friend WithEvents BRANCH_CB As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_CB_BRANCH As System.Windows.Forms.ComboBox
    Friend WithEvents BRANCH_AP As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_AP_BRANCH As System.Windows.Forms.ComboBox
    Friend WithEvents BRANCH_AR As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_AR_BRANCH As System.Windows.Forms.ComboBox
    Friend WithEvents BRANCH_GL As System.Windows.Forms.GroupBox
    Friend WithEvents Cb_GL_BRANCH As System.Windows.Forms.ComboBox
    Friend WithEvents Ch_mS As System.Windows.Forms.RadioButton
    Friend WithEvents Ch_OOr As System.Windows.Forms.RadioButton
    Friend WithEvents CH_Per As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox17 As System.Windows.Forms.GroupBox
    Friend WithEvents chk_docAP As System.Windows.Forms.CheckBox
    Friend WithEvents chk_docCB As System.Windows.Forms.CheckBox
    Friend WithEvents chk_docGL As System.Windows.Forms.CheckBox
    Friend WithEvents chk_docAR As System.Windows.Forms.CheckBox
    Friend WithEvents PnConTax As System.Windows.Forms.Panel
    Friend WithEvents PnDBset As System.Windows.Forms.Panel
    Friend WithEvents PnConfig As System.Windows.Forms.Panel
    Friend WithEvents btnLogin As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Ch_AR As System.Windows.Forms.CheckBox
End Class
