Public Class FrmConfigDB
    'Dim h As New BaseFILEIO(Application.StartupPath)
    Dim DT As DataTable
    Dim Index As Integer = 0
    Dim MODE, ComSetting, CMD As String
    Dim iCount, Use_datemode As String
    Dim inCB, inAP, inAR, TGLREVERSE, TCBREVERSE, tOPT_AP, tOPT_CB, tOPT_GL, tRERUN, tUPRUN, tVATAutoRun, tS_AP, tS_AR, tS_PO, ts_OE, tS_GL, tS_CB, tVNO_AP, tVNO_CB, tVNO_GL, tVNO_AR As Boolean
    Dim ts_TAX_AP, ts_TAX_CB, ts_TAX_AR, ts_BRANCH_AP, ts_BRANCH_CB, ts_BRANCH_AR As TAXFROM
    Dim ts_TAX_GL, ts_BRANCH_GL As TAXFROMGL

    Private Sub FrnConfigDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '''''''''''''
        If CMD = "CanNotFindServer" Then
            ReadDB(ComID)
            btnNew.Enabled = False
            btnN.Enabled = False
            btnP.Enabled = False
            btnDelete.Enabled = False
        ElseIf CMD = "New" Then
            btnNew_Click(Nothing, Nothing)
            btnNew.Enabled = False
            btnN.Enabled = False
            btnP.Enabled = False
            btnDelete.Enabled = False
        ElseIf CMD = "" Then
            Index = ComID
            ReadDB(ComID)
        End If

        'Me.AcceptButton = btnACCSave


        ACCPAC.BTValue = True
        GACC_BT_CLICK(True)
        VAT.BTValue = True
        GVAT_BT_CLICK(True)

        If UserName <> "admin" Then
            GroupBox6.Enabled = False

        End If

        If fmsadmin <> "FMSADMIN" Then
            PnDBset.Visible = False
            btnNew.Visible = False
            btnDelete.Visible = False

        End If
        'Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Public Sub ReadDB(ByVal Order As String)
        inCB = GetPrivateProfile("COM" & Order, "Use_LEGALNAME_inCB", "FALSE")
        inAP = GetPrivateProfile("COM" & Order, "Use_LEGALNAME_inAP", "FALSE")
        inAR = GetPrivateProfile("COM" & Order, "Use_LEGALNAME_inAR", "FALSE")
        TGLREVERSE = GetPrivateProfile("COM" & Order, "GLREVERSE", "FALSE")
        TCBREVERSE = GetPrivateProfile("COM" & Order, "CBREVERSE", "FALSE")
        Use_datemode = GetPrivateProfile("COM" & Order, "DATEMODE", "POST")
        Ch_runFormat.Text = GetPrivateProfile("COM" & Order, "FORMAT", "00000")
        CB_PAGE.SelectedIndex = CInt(GetPrivateProfile("COM" & Order, "DISPLAYPAGE", 5) - 1)
        txFormatRunnig.Text = GetPrivateProfile("COM" & Order, "FormatRunnig", "yy+MM+/")

        tOPT_AP = GetPrivateProfile("COM" & Order, "OPT_AP", "FALSE")
        tOPT_CB = GetPrivateProfile("COM" & Order, "OPT_CB", "FALSE")
        tOPT_GL = GetPrivateProfile("COM" & Order, "OPT_GL", "FALSE")

        tRERUN = GetPrivateProfile("COM" & Order, "RERUN", "FALSE")
        tUPRUN = GetPrivateProfile("COM" & Order, "UPRUN", "FALSE")

        tS_AP = GetPrivateProfile("COM" & Order, "SHOWAP", "TRUE")
        tS_AR = GetPrivateProfile("COM" & Order, "SHOWAR", "TRUE")
        tS_PO = GetPrivateProfile("COM" & Order, "SHOWPO", "TRUE")
        ts_OE = GetPrivateProfile("COM" & Order, "SHOWOE", "TRUE")
        tS_GL = GetPrivateProfile("COM" & Order, "SHOWGL", "TRUE")
        tS_CB = GetPrivateProfile("COM" & Order, "SHOWCB", "TRUE")

        tVATAutoRun = IIf(GetPrivateProfile("COM" & Order, "VATRUNMANUAL", "TRUE") = "N", False, True)
        ts_TAX_AP = GetPrivateProfile("COM" & Order, "Use_TAX_AP", "0")
        ts_TAX_AR = GetPrivateProfile("COM" & Order, "Use_TAX_AR", "0")
        ts_TAX_CB = GetPrivateProfile("COM" & Order, "Use_TAX_CB", "0")
        ts_TAX_GL = GetPrivateProfile("COM" & Order, "Use_TAX_GL", "0")

        ts_BRANCH_AP = GetPrivateProfile("COM" & Order, "Use_BRANCH_AP", "0")
        ts_BRANCH_AR = GetPrivateProfile("COM" & Order, "Use_BRANCH_AR", "0")
        ts_BRANCH_CB = GetPrivateProfile("COM" & Order, "Use_BRANCH_CB", "0")
        ts_BRANCH_GL = GetPrivateProfile("COM" & Order, "Use_BRANCH_GL", "0")

        tVNO_AP = GetPrivateProfile("COM" & Order, "VNO_AP", "FALSE")
        tVNO_AR = GetPrivateProfile("COM" & Order, "VNO_AR", "FALSE")
        tVNO_CB = GetPrivateProfile("COM" & Order, "VNO_CB", "FALSE")
        tVNO_GL = GetPrivateProfile("COM" & Order, "VNO_GL", "FALSE")

        Cb_AP_TAX.SelectedIndex = ts_TAX_AP
        Cb_AR_TAX.SelectedIndex = ts_TAX_AR
        Cb_CB_TAX.SelectedIndex = ts_TAX_CB
        Cb_GL_TAX.SelectedIndex = ts_TAX_GL

        Cb_AP_BRANCH.SelectedIndex = ts_BRANCH_AP
        Cb_AR_BRANCH.SelectedIndex = ts_BRANCH_AR
        Cb_CB_BRANCH.SelectedIndex = ts_BRANCH_CB
        Cb_GL_BRANCH.SelectedIndex = ts_BRANCH_GL

        If tRERUN = True Then
            cb_RERUN_Y.Checked = True
        Else
            cb_RERUN_N.Checked = True
        End If
        If tUPRUN = True Then
            cb_UPRUN_Y.Checked = True
        Else
            cb_UPRUN_N.Checked = True
        End If

        If tVATAutoRun = True Then
            Rb_AutoRunning.Checked = True
        Else
            Rb_ManualRunning.Checked = True
        End If

        If CBool(GetPrivateProfile("COM" & Order, "Verify", 0)) Then
            CB_AppY.Checked = True
        Else
            CB_AppN.Checked = True
        End If
        If Ch_runFormat.Text.Trim = "" Then
            Ch_runFormat.Text = "00000"
        End If
        If Use_datemode = "POST" Then
            Ch_PoST.Checked = True
        Else
            Ch_Doc.Checked = True
        End If

        If GetPrivateProfile("COM" & Order, "INVFROMPO", "0") = "0" Then
            PO_NO.Checked = True
        Else
            PO_YES.Checked = True
        End If

        cb_OPT_AP.Checked = tOPT_AP
        cb_OPT_CB.Checked = tOPT_CB
        cb_OPT_GL.Checked = tOPT_GL

        Ch_AP.Checked = inAP
        Ch_AR.Checked = inAR
        Ch_CB.Checked = inCB

        Ch_GLReverse.Checked = TGLREVERSE
        Ch_CBReverse.Checked = TCBREVERSE

        CB_AP.Checked = tS_AP
        CB_AR.Checked = tS_AR
        CB_PO.Checked = tS_PO
        CB_OE.Checked = ts_OE
        CB_GL.Checked = tS_GL
        CB_CB.Checked = tS_CB

        chk_docAP.Checked = tVNO_AP
        chk_docAR.Checked = tVNO_AR
        chk_docCB.Checked = tVNO_CB
        chk_docGL.Checked = tVNO_GL

        ToolTip1.SetToolTip(btnNew, "New")
        ToolTip1.SetToolTip(btnP, "Previous")
        ToolTip1.SetToolTip(btnN, "Next")
        EnabledControl(False)
        ' If CMD <> "" Then
        If Order Is Nothing Then
            btnACCSave.Text = "Add"
            btnACCClose.Text = "Cancel"
        Else
            btnACCSave.Text = "Save"
            btnACCClose.Text = "Reset"
        End If
        EnabledControl(True)
        btnP.Enabled = False
        btnN.Enabled = False
        btnDelete.Enabled = True
        Select Case GetPrivateProfile("COM" & Order, "DATABASE", "").Split(";")(0)
            Case "PVSW"
                CH_Per.Checked = True
            Case "MSSQL"
                Ch_mS.Checked = True
            Case "ORCL"
                Ch_OOr.Checked = True
        End Select
        CH_Per_CheckedChanged(Nothing, Nothing)

        If GetPrivateProfile("COM" & Order, "LOGIN", "0").Split("")(0) = "1" Then
            ch_loginY.Checked = True
        Else
            ch_loginN.Checked = True
        End If


        If GetPrivateProfile("COM" & Order, "GetType", "1").Split("")(0) = "1" Then
            Rd_Affter.Checked = True
        Else
            Rd_Before.Checked = True
        End If

        txtCompany.Text = GetPrivateProfile("COMPANY", "COM" & Order, "").Split(";")(0)
        Dim Dr_VAT() As String = GetPrivateProfile("COM" & Order, "VAT", "").Split(";")
        Dim Dr_ACC() As String = GetPrivateProfile("COM" & Order, "ACCPAC", "").Split(";")
        If Dr_ACC.Length > 0 Then
            txtACCServerName.Text = Dr_ACC(0)
            If Dr_ACC.Length > 2 Then
                txtACCDBName.Text = Dr_ACC(1)
            Else
                txtACCDBName.Text = ""
            End If

            If Dr_ACC.Length > 3 Then
                txtACCUsername.Text = Dr_ACC(2)
            Else
                txtACCUsername.Text = ""
            End If

            If Dr_ACC.Length > 4 Then
                txtACCPassword.Text = Dr_ACC(3)
            Else
                txtACCPassword.Text = ""
            End If

        End If
        If Dr_VAT.Length > 0 Then
            txtVATServerName.Text = Dr_VAT(0)
            If Dr_VAT.Length > 2 Then
                txtVATDBName.Text = Dr_VAT(1)
            Else
                txtVATDBName.Text = ""
            End If

            If Dr_VAT.Length > 3 Then
                txtVATUsername.Text = Dr_VAT(2)
            Else
                txtVATUsername.Text = ""
            End If

            If Dr_VAT.Length > 4 Then
                txtVATPassword.Text = Dr_VAT(3)
            Else
                txtVATPassword.Text = ""
            End If

        End If
        'End If
        Dim max As Integer = GetPrivateProfile("COMPANY", "List", "0").Replace(";", "")
        If Index + 1 > max Then
            btnN.Enabled = False
        Else
            btnN.Enabled = True
        End If

        If Index > 1 Then
            btnP.Enabled = True
        Else
            btnP.Enabled = False
        End If
    End Sub

    Public Sub EnabledControl(ByVal Value As Boolean)
        txtCompany.Enabled = Value
        txtACCServerName.Enabled = Value
        txtACCDBName.Enabled = Value
        txtACCUsername.Enabled = Value
        txtACCPassword.Enabled = Value

        txtVATServerName.Enabled = Value
        txtVATDBName.Enabled = Value
        txtVATUsername.Enabled = Value
        txtVATPassword.Enabled = Value
    End Sub

    Private Sub btnP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnP.Click
        ClearValuecontrol()
        Index = Index - 1
        ReadDB(Index)
    End Sub

    Private Sub btnN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnN.Click
        ClearValuecontrol()
        Index = Index + 1
        ReadDB(Index)
    End Sub

    Public Sub ClearValuecontrol()
        EnabledControl(True)
        txtACCServerName.Text = ""
        txtACCDBName.Text = ""
        txtACCUsername.Text = ""
        txtACCPassword.Text = ""

        txtVATServerName.Text = ""
        txtVATDBName.Text = ""
        txtVATUsername.Text = ""
        txtVATPassword.Text = ""

        txtACCServerName.BackColor = Color.White
        txtACCDBName.BackColor = Color.White
        txtACCUsername.BackColor = Color.White
        txtACCPassword.BackColor = Color.White

        txtVATServerName.BackColor = Color.White
        txtVATDBName.BackColor = Color.White
        txtVATUsername.BackColor = Color.White
        txtVATPassword.BackColor = Color.White

        txtCompany.Text = ""
        ACCPAC.SetDefuleColor()
        VAT.SetDefuleColor()

    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        btnP.Enabled = False
        btnN.Enabled = False
        txtACCServerName.Text = ""
        txtACCDBName.Text = ""
        txtACCUsername.Text = ""
        txtACCPassword.Text = ""

        txtVATServerName.Text = ""
        txtVATDBName.Text = ""
        txtVATUsername.Text = ""
        txtVATPassword.Text = ""

        txtCompany.Text = ""
        btnACCSave.Text = "Add"
        EnabledControl(True)
        Ch_mS.Checked = True
        Ch_Doc.Checked = True
        Ch_AP.Checked = False
        Ch_AR.Checked = False
        Ch_CB.Checked = False
        Rb_AutoRunning.Checked = True


        cb_OPT_AP.Checked = True
        cb_OPT_CB.Checked = True
        cb_OPT_GL.Checked = True
        Ch_AP.Checked = False
        Ch_AR.Checked = False
        Ch_CB.Checked = False
        Ch_GLReverse.Checked = False
        Ch_CBReverse.Checked = False

        CB_AP.Checked = True
        CB_AR.Checked = True
        CB_PO.Checked = True
        CB_OE.Checked = True
        CB_GL.Checked = True
        CB_CB.Checked = True

        Cb_AP_TAX.SelectedIndex = 0
        Cb_AR_TAX.SelectedIndex = 0
        Cb_CB_TAX.SelectedIndex = 0
        Cb_GL_TAX.SelectedIndex = 0

        Cb_AP_BRANCH.SelectedIndex = 0
        Cb_AR_BRANCH.SelectedIndex = 0
        Cb_CB_BRANCH.SelectedIndex = 0
        Cb_GL_BRANCH.SelectedIndex = 0

        PO_NO.Checked = True
        CB_AppY.Checked = False
        Rd_Before.Checked = True
        ACCPAC.SetDefuleColor()
        VAT.SetDefuleColor()
        ReadDB(Nothing)
        btnDelete.Enabled = False


    End Sub

    Private Function BtnSave(ByRef DB As BaseClass.BaseGroupBox, ByRef txtServerName As TextBox, ByRef txtDBName As ComboBox, ByRef txtUsername As TextBox, ByRef txtPassword As TextBox, ByRef StrCon As String) As Boolean
        Try
            Dim cn As ADODB.Connection
            Dim rs As BaseClass.BaseODBCIO
            Dim Connstring As String
            Dim strsplit As Object
            cn = New ADODB.Connection
            rs = New BaseClass.BaseODBCIO
            If CH_Per.Checked Then
                ComSetting = txtServerName.Text & ";"
            ElseIf Ch_mS.Checked Then
                ComSetting = txtServerName.Text & ";" & txtDBName.Text & ";" & txtUsername.Text & ";" & txtPassword.Text & ";"
            ElseIf Ch_OOr.Checked Then
                ComSetting = txtServerName.Text & ";" & txtDBName.Text & ";" & txtUsername.Text & ";" & txtPassword.Text & ";"
            End If

            If CMD = "" Then
                If DB.Name = "ACCPAC" Then
                    ComACCPAC = ComSetting
                ElseIf DB.Name = "VAT" Then
                    ComVAT = ComSetting
                End If
            End If

            strsplit = Split(ComSetting, ";")
            Connstring = ""
            For i As Integer = LBound(strsplit) To UBound(strsplit)
                Select Case i
                    Case Is = 0
                        If CH_Per.Checked Or (Ch_OOr.Checked And MODE = 2) Then
                            Connstring = "Provider=MSDASQL.1;Persist Security Info=False;" & "Data Source=" & strsplit(i)
                        ElseIf Ch_mS.Checked Then

                            ' Connstring = "driver={SQL Server};server=" & strsplit(i)
                            Connstring = "Provider=SQLOLEDB.1;server=" & strsplit(i)

                        ElseIf Ch_OOr.Checked Then
                            Connstring = "Provider=MSDAORA.1;Data Source=" & strsplit(i) & ";Persist Security Info=False"
                        End If
                        ConnDB = "Data Source=" & strsplit(i)
                    Case Is = 1
                        If CH_Per.Checked Then
                            Connstring = Connstring
                        ElseIf Ch_mS.Checked Then
                            Connstring = Connstring & ";database=" & strsplit(i)
                        End If
                        ConnDB &= " ;Initial Catalog=" & strsplit(i)
                    Case Is = 2
                        If Ch_mS.Checked Then
                            Connstring = Connstring & ";uid=" & strsplit(i)
                        ElseIf Ch_OOr.Checked Then
                            Connstring = Connstring & ";USER ID=" & strsplit(i)
                        End If
                        'ConnDB &= ";Persist Security Info=True;User ID=" & strsplit(i)

                        ConnDB &= ";User ID=" & strsplit(i)
                    Case Is = 3
                        If Ch_mS.Checked Then
                            Connstring = Connstring & ";pwd=" & strsplit(i)
                        ElseIf Ch_OOr.Checked Then
                            Connstring = Connstring & ";PASSWORD=" & strsplit(i)
                        End If
                        ConnDB &= " ;Password=" & strsplit(i)
                End Select
            Next

            Try
                cn.Open(Connstring & ";Connect Timeout=3;")
                StrCon = Connstring
            Catch ex As Exception
                txtServerName.BackColor = Color.Red
                If CH_Per.Checked = False Then
                    txtDBName.BackColor = Color.Red
                    txtUsername.BackColor = Color.Red
                    txtPassword.BackColor = Color.Red
                End If
                Return False
            End Try

            txtServerName.BackColor = Color.White
            txtDBName.BackColor = Color.White
            txtUsername.BackColor = Color.White
            txtPassword.BackColor = Color.White

        Catch ex As Exception
            ShowToolTip(DB, ToolTipIcon.Warning, "Setup configuration Incomplete." & Environment.NewLine & "Please try again")
            Return False
        End Try
        Return True
    End Function

    Public Sub New(ByVal _CMD As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        CMD = _CMD
    End Sub

    Private Sub GACC_BT_CLICK(ByVal Value As Boolean) Handles ACCPAC.BT_CLICK
        'If Value = True Then
        '    ACCPAC.Height = ACCPAC.MaximumSize.Height
        '    VAT.Location = New Point(VAT.Location.X, ACCPAC.Location.Y + ACCPAC.Height + 6)
        '    Me.Height += ACCPAC.MaximumSize.Height + 6
        'Else
        '    ACCPAC.Height = ACCPAC.MinimumSize.Height
        '    VAT.Location = New Point(VAT.Location.X, ACCPAC.Location.Y + ACCPAC.Height + 6)
        '    Me.Height = Me.Height - ACCPAC.MaximumSize.Height + 25
        'End If
        'Application.DoEvents()
    End Sub

    Private Sub GVAT_BT_CLICK(ByVal Value As Boolean) Handles VAT.BT_CLICK
        'If Value = True Then
        '    VAT.Height = VAT.MaximumSize.Height
        '    Me.Height += VAT.MaximumSize.Height + 6
        '    'GCB.Location = New Point(GCB.Location.X, GCB.Location.Y + GACC.MaximumSize.Height - 25)
        '    'GVAT.Location = New Point(GVAT.Location.X, GVAT.Location.Y + GCB.MaximumSize.Height - 25)
        'Else
        '    VAT.Height = VAT.MinimumSize.Height
        '    Me.Height = Me.Height - VAT.MaximumSize.Height + 25 + 6
        '    'GCB.Location = New Point(GCB.Location.X, GCB.Location.Y - GACC.MaximumSize.Height + 25)
        '    'GVAT.Location = New Point(GVAT.Location.X, GVAT.Location.Y - GCB.MaximumSize.Height + 25)
        'End If
        'Application.DoEvents()
    End Sub

    Private Sub FrmConfigDB_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.Height = VAT.Location.Y + VAT.Height + Panel1.Height + 50
        Application.DoEvents()
    End Sub

    Private Sub CH_Per_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CH_Per.CheckedChanged
        txtACCDBName.Enabled = Not CH_Per.Checked
        txtACCPassword.Enabled = Not CH_Per.Checked
        txtACCUsername.Enabled = Not CH_Per.Checked
        txtVATDBName.Enabled = Not CH_Per.Checked
        txtVATPassword.Enabled = Not CH_Per.Checked
        txtVATUsername.Enabled = Not CH_Per.Checked
        If CH_Per.Checked = True Then
            txtACCDBName.Text = ""
            txtACCPassword.Text = ""
            txtACCUsername.Text = ""

            txtVATDBName.Text = ""
            txtVATPassword.Text = ""
            txtVATUsername.Text = ""
        End If
    End Sub

    Private Sub Ch_mS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ch_OOr.CheckedChanged, Ch_mS.CheckedChanged
        txtACCDBName.Enabled = Not CH_Per.Checked
        txtACCPassword.Enabled = Not CH_Per.Checked
        txtACCUsername.Enabled = Not CH_Per.Checked

        txtVATDBName.Enabled = Not CH_Per.Checked
        txtVATPassword.Enabled = Not CH_Per.Checked
        txtVATUsername.Enabled = Not CH_Per.Checked

    End Sub

    Private Sub BtnACCSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnACCSave.Click
        Dim VACC, VVAT As Boolean
        iCount = (CInt(GetPrivateProfile("COMPANY", "List", "0").Replace(";", "")) + 1).ToString()
        MODE = 1
        Dim StrCon As String = ""
        If BtnSave(ACCPAC, txtACCServerName, txtACCDBName, txtACCUsername, txtACCPassword, StrCon) Then
            If FindApp("PO", StrCon) = False Then PO_NO.Checked = True
            GroupBox10.Enabled = FindApp("PO", StrCon)
            VACC = True
            ACCPAC.SetBTColor(Color.Green)
        Else
            VACC = False
        End If

        MODE = 0
        If BtnSave(VAT, txtVATServerName, txtVATDBName, txtVATUsername, txtVATPassword, StrCon) Then
            VVAT = True
            VAT.SetBTColor(Color.Green)
        Else
            VVAT = False
        End If
        If VACC = False Then
            If ACCPAC.BTValue = False Then
                ACCPAC.BTValue = True
                GACC_BT_CLICK(True)
            End If
            ACCPAC.SetBTColor(Color.Red)
            ShowToolTip(ACCPAC.Controls("BT"), ToolTipIcon.Warning, "Setup configuration Incomplete." & Environment.NewLine & "Please try again")
        End If
        If VVAT = False Then
            If VAT.BTValue = False Then
                VAT.BTValue = True
                GVAT_BT_CLICK(True)
            End If
            VAT.SetBTColor(Color.Red)
            ShowToolTip(VAT.Controls("BT"), ToolTipIcon.Warning, "Setup configuration Incomplete." & Environment.NewLine & "Please try again")
        End If

        If VACC And VVAT Then
            Dim i As String
            If CMD = "New" Then
                If btnACCSave.Text = "Add" Then
                    i = iCount
                Else
                    i = Index
                End If
            ElseIf CMD = "CanNotFindServer" Then
                i = ComID
            ElseIf btnACCSave.Text = "Add" Then
                i = iCount
            Else
                i = Index
            End If
            Index = CInt(i)

            If btnACCSave.Text = "Add" Then
                WritePrivateProfileString("COMPANY", "LIST", i, Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "VATPAR", "Y", Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "SHOWNAMEONLY", "Y", Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "Owner", "dbo", Application.StartupPath & "/FMSVAT.ini")
            End If

            WritePrivateProfileString("COM" & i, "VATRUNMANUAL", IIf(Rb_AutoRunning.Checked, "Y", "N"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "OPT_AP", IIf(cb_OPT_AP.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "OPT_CB", IIf(cb_OPT_CB.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "OPT_GL", IIf(cb_OPT_GL.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "RERUN", IIf(cb_RERUN_Y.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "UPRUN", IIf(cb_UPRUN_Y.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "Use_LEGALNAME_inCB", IIf(Ch_CB.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "Use_LEGALNAME_inAP", IIf(Ch_AP.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "Use_LEGALNAME_inAR", IIf(Ch_AR.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "CBREVERSE", IIf(Ch_CBReverse.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "GLREVERSE", IIf(Ch_GLReverse.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "Verify", IIf(CB_AppY.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "GetType", IIf(Rd_Before.Checked, "0", "1"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "INVFROMPO", IIf(PO_YES.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWAP", IIf(CB_AP.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWAR", IIf(CB_AR.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWPO", IIf(CB_PO.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWOE", IIf(CB_OE.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWCB", IIf(CB_CB.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "SHOWGL", IIf(CB_GL.Checked, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "FormatRunnig", txFormatRunnig.Text, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_TAX_AP", Cb_AP_TAX.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_TAX_AR", Cb_AR_TAX.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_TAX_CB", Cb_CB_TAX.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_TAX_GL", Cb_GL_TAX.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_BRANCH_AP", Cb_AP_BRANCH.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_BRANCH_AR", Cb_AR_BRANCH.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_BRANCH_CB", Cb_CB_BRANCH.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "USE_BRANCH_GL", Cb_GL_BRANCH.SelectedIndex, Application.StartupPath & "/FMSVAT.ini")

          
            WritePrivateProfileString("COM" & i, "VNO_AP", IIf(chk_docAP.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "VNO_AR", IIf(chk_docAR.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "VNO_CB", IIf(chk_docCB.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "VNO_GL", IIf(chk_docGL.Checked, "TRUE", "FALSE"), Application.StartupPath & "/FMSVAT.ini")


            If Ch_PoST.Checked = True Then
                Use_datemode = "POST"
                WritePrivateProfileString("COM" & i, "DATEMODE", Use_datemode, Application.StartupPath & "/FMSVAT.ini")
            Else
                Use_datemode = "DOCU"
                WritePrivateProfileString("COM" & i, "DATEMODE", Use_datemode, Application.StartupPath & "/FMSVAT.ini")
            End If

            If Me.Owner.Name = "frmMain" And btnACCSave.Text = "Save" And ComID = i Then
                Call BuildConnection(lAccpac, txtACCServerName.Text & ";" & txtACCDBName.Text & ";" & txtACCUsername.Text & ";" & txtACCPassword.Text, Me)
                Call BuildConnection(lVat, txtVATServerName.Text & ";" & txtVATDBName.Text & ";" & txtVATUsername.Text & ";" & txtVATPassword.Text, Me)
                DATEMODE = Use_datemode
                Use_LEGALNAME_inAP = Ch_AP.Checked
                Use_LEGALNAME_inAR = Ch_AR.Checked
                Use_POST = Rd_Affter.Checked
                Use_LEGALNAME_inCB = Ch_CB.Checked
                Use_DisplayPage = CB_PAGE.SelectedIndex + 1
                Use_INVFROMPO = PO_YES.Checked
                GLREVERSE = Ch_GLReverse.Checked
                CBREVERSE = Ch_CBReverse.Checked

                OPT_AP = cb_OPT_AP.Checked
                OPT_CB = cb_OPT_CB.Checked
                OPT_GL = cb_OPT_GL.Checked

                RERUN = cb_RERUN_Y.Checked
                UPRUN = cb_UPRUN_Y.Checked
                FormatRunnig = txFormatRunnig.Text
                VATAutoRun = Rb_AutoRunning.Checked
                VATFormat = Ch_runFormat.Text

                Use_TAX_AP = Cb_AP_TAX.SelectedIndex
                Use_TAX_AR = Cb_AR_TAX.SelectedIndex
                Use_TAX_CB = Cb_CB_TAX.SelectedIndex
                Use_TAX_GL = Cb_GL_TAX.SelectedIndex

                Use_BRANCH_AP = Cb_AP_BRANCH.SelectedIndex
                Use_BRANCH_AR = Cb_AR_BRANCH.SelectedIndex
                Use_BRANCH_CB = Cb_CB_BRANCH.SelectedIndex
                Use_BRANCH_GL = Cb_GL_BRANCH.SelectedIndex

                VNO_AP = chk_docAP.Checked
                VNO_CB = chk_docCB.Checked
                VNO_AR = chk_docAR.Checked
                VNO_GL = chk_docGL.Checked

                Dim cboView As ComboBox = CType(Me.Owner, frmMain).Controls("Panel2").Controls("cboView")
                Dim lstVu As DataGridView = CType(Me.Owner, frmMain).Controls("Panel3").Controls("lstVu")
                Use_Approved = CB_AppY.Checked
                lstVu.Columns(1).Visible = Use_Approved
                '''''''''''''''
                S_AP = CB_AP.Checked
                S_AR = CB_AR.Checked
                S_GL = CB_GL.Checked
                S_CB = CB_CB.Checked
                S_PO = CB_PO.Checked
                s_OE = CB_OE.Checked
                ''''''''''''''   
                Call LoadDBScript()
                Call SetInitial()
                Call PrepareVatSet()
                Call PrepareVatUpdateNew()
                Call PrepareType()
                Call PrepareLocation()
                Call PrepareAccount()
                Call PrepareVatTable()
                Call PrepareVatInsertTable()
                Call PrepareRUNNING()


                CType(Me.Owner, frmMain).HeaderCheckBox.Visible = Use_Approved
                CType(Me.Owner, frmMain).btnRefreh_Click(Nothing, Nothing)
                If cboView.SelectedIndex + 1 = 2 Then
                    'lstVu.Columns.Item(5).Width = VB6.TwipsToPixelsX(2650.39 + 1400.31)
                    'lstVu.Columns.Item(4).Width = 0
                    lstVu.Columns.Item(4).Width = VB6.TwipsToPixelsX(2650.39 + 1400.31)
                    lstVu.Columns.Item(3).Width = 0
                    lstVu.Columns.Item(3).Visible = False
                Else
                    lstVu.Columns.Item(4).Width = VB6.TwipsToPixelsX(2650.39)
                    lstVu.Columns.Item(3).Width = VB6.TwipsToPixelsX(1400.31)
                End If
            End If
            WritePrivateProfileString("COM" & i, "LOGIN", IIf(ch_loginY.Checked = True, "1", "0"), Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "FORMAT", Ch_runFormat.Text, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COM" & i, "DISPLAYPAGE", CB_PAGE.SelectedIndex + 1, Application.StartupPath & "/FMSVAT.ini")

            WritePrivateProfileString("COMPANY", "COM" & i, txtCompany.Text & ";", Application.StartupPath & "/FMSVAT.ini")
            If CH_Per.Checked Then
                WritePrivateProfileString("COM" & i, "DATABASE", "PVSW;", Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "ACCPAC", BaseClass.AES_Encrypt(txtACCServerName.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "VAT", BaseClass.AES_Encrypt(txtVATServerName.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
            ElseIf Ch_mS.Checked Then
                WritePrivateProfileString("COM" & i, "DATABASE", "MSSQL;", Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "ACCPAC", BaseClass.AES_Encrypt(txtACCServerName.Text & ";" & txtACCDBName.Text & ";" & txtACCUsername.Text & ";" & txtACCPassword.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "VAT", BaseClass.AES_Encrypt(txtVATServerName.Text & ";" & txtVATDBName.Text & ";" & txtVATUsername.Text & ";" & txtVATPassword.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
            ElseIf Ch_OOr.Checked Then
                WritePrivateProfileString("COM" & i, "DATABASE", "MSSQL;", Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "ACCPAC", BaseClass.AES_Encrypt(txtACCServerName.Text & ";" & txtACCDBName.Text & ";" & txtACCUsername.Text & ";" & txtACCPassword.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
                WritePrivateProfileString("COM" & i, "VAT", BaseClass.AES_Encrypt(txtVATServerName.Text & ";" & txtVATDBName.Text & ";" & txtVATUsername.Text & ";" & txtVATPassword.Text & ";", "ABCDEFG"), Application.StartupPath & "/FMSVAT.ini")
            End If
            btnACCSave.Text = "Save"
            btnACCClose.Text = "Reset"
            Dim max As Integer = GetPrivateProfile("COMPANY", "List", "0")
            If max > i Then
                btnN.Enabled = True
            Else
                btnN.Enabled = False
            End If
            If i = 1 Then
                btnP.Enabled = False
            Else
                btnP.Enabled = True
            End If

            If CMD = "New" Or CMD = "CanNotFindServer" Then
                Me.DialogResult = Windows.Forms.DialogResult.Yes
                Me.Close()
            End If


            'If ch_loginY.Checked = True Then

            '    Dim frm As frmOpen = New frmOpen
            '    frm.ShowDialog()
            '    If frm.DialogResult <> Windows.Forms.DialogResult.Yes Then
            '        frm.btnClose.Enabled = False

            '    End If

            'End If





        End If


    End Sub

    Private Sub BtnACCClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnACCClose.Click
        If btnACCSave.Text = "Add" Then
            Dim max As Integer = GetPrivateProfile("COMPANY", "List", "0")
            If max > 0 Then
                Index = 0
                btnN_Click(Nothing, Nothing)
            End If
        Else
            If CMD = "CanNotFindServer" Then
                ReadDB(ComID)
            Else
                ReadDB(Index)
            End If

        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Me.Owner.Name = "frmMain" And btnACCSave.Text = "Save" And ComID = Index Then
            MessageBox.Show("Company already to Use Cannot Delete!", "Warning!!", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If

        If MessageBox.Show("Confirm Delete " & txtCompany.Text & "?", "Confirmation...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> Windows.Forms.DialogResult.Yes Then Exit Sub
        Dim max As Integer = GetPrivateProfile("COMPANY", "List", "0")
        Dim VAR() As String = {"VATRUNMANUAL", "OPT_AP", "OPT_CB", "OPT_GL", "RERUN", "UPRUN", "USE_LEGALNAME_INCB", "USE_LEGALNAME_INAP", "CBREVERSE", "GLREVERSE", "VERIFY", "GETTYPE", "INVFROMPO", "SHOWAP", "SHOWAR", "SHOWPO", "SHOWOE", "SHOWCB", "SHOWGL", "DATEMODE", "LOGIN", "FORMAT", "DISPLAYPAGE", "DATABASE", "ACCPAC", "VAT", "VATPAR", "SHOWNAMEONLY", "OWNER", "USE_TAX_AP", "USE_TAX_AR", "USE_TAX_CB", "USE_TAX_GL", "USE_BRANCH_AP", "USE_BRANCH_AR", "USE_BRANCH_CB", "USE_BRANCH_GL"}
        For Each VALUE As String In VAR
            WritePrivateProfileString("COM" & Index, VALUE, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        Next
        WritePrivateProfileString("COM" & Index, vbNullString, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        WritePrivateProfileString("COMPANY", "COM" & Index, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        Dim countID As Integer = (CInt(GetPrivateProfile("COMPANY", "List", "0")) - 1).ToString()
        WritePrivateProfileString("COMPANY", "List", countID.ToString, Application.StartupPath & "/FMSVAT.ini")

        For i As Integer = Index + 1 To max
            WritePrivateProfileString("COMPANY", "COM" & i - 1, vbNullString, Application.StartupPath & "/FMSVAT.ini")
            WritePrivateProfileString("COMPANY", "COM" & i - 1, GetPrivateProfile("COMPANY", "COM" & i, ""), Application.StartupPath & "/FMSVAT.ini")
            For Each VALUE As String In VAR
                WritePrivateProfileString("COM" & i - 1, VALUE, GetPrivateProfileString("COM" & i, VALUE, ""), Application.StartupPath & "/FMSVAT.ini")
            Next
        Next

        WritePrivateProfileString("COMPANY", "COM" & max, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        For Each VALUE As String In VAR
            WritePrivateProfileString("COM" & max, VALUE, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        Next
        WritePrivateProfileString("COM" & max, vbNullString, vbNullString, Application.StartupPath & "/FMSVAT.ini")
        If countID < Index Then
            ReadDB(countID)
            ComID = countID
            Index = countID
        Else
            ReadDB(IIf(Index - 1 = 0, 1, Index - 1))
            ComID = IIf(ComID - 1 = 0, 1, ComID - 1)
            Index = IIf(ComID - 1 = 0, 1, ComID - 1)
        End If

    End Sub

    Private Sub txtACCServerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtACCServerName.TextChanged
        txtVATServerName.Text = txtACCServerName.Text
    End Sub

    Private Sub txtACCUsername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtACCUsername.TextChanged
        txtVATUsername.Text = txtACCUsername.Text
    End Sub

    Private Sub txtACCPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtACCPassword.TextChanged
        txtVATPassword.Text = txtACCPassword.Text
    End Sub

    Private Sub BtnACCSave_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnACCSave.KeyDown
        If e.KeyCode = Keys.F5 Then
            Dim frm As New BaseClass.Query()
            frm.Scon = txtACCServerName.Text
            frm.DBcon = txtACCDBName.Text
            frm.USerCon = txtACCUsername.Text
            frm.PassCon = txtACCPassword.Text
            frm.Show()
        End If
    End Sub

    Private Sub txtACCDBName_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtACCDBName.DropDown
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim Connstring As String
        Dim strsplit As Object
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        If CH_Per.Checked Then
            ComSetting = txtACCServerName.Text & ";"
        ElseIf Ch_mS.Checked Then
            ComSetting = txtACCServerName.Text & ";master;" & txtACCUsername.Text & ";" & txtACCPassword.Text & ";"
        ElseIf Ch_OOr.Checked Then
            ComSetting = txtACCServerName.Text & ";master;" & txtACCUsername.Text & ";" & txtACCPassword.Text & ";"
        End If

        strsplit = Split(ComSetting, ";")
        Connstring = ""
        For i As Integer = LBound(strsplit) To UBound(strsplit)
            Select Case i
                Case Is = 0
                    If CH_Per.Checked Or (Ch_OOr.Checked And MODE = 2) Then
                        Connstring = "Provider=MSDASQL.1;Persist Security Info=False;" & "Data Source=" & strsplit(i)
                    ElseIf Ch_mS.Checked Then

                        ' Connstring = "driver={SQL Server};server=" & strsplit(i)
                        Connstring = "Provider=SQLOLEDB.1;server=" & strsplit(i)

                    ElseIf Ch_OOr.Checked Then
                        Connstring = "Provider=MSDAORA.1;Data Source=" & strsplit(i) & ";Persist Security Info=False"
                    End If
                    ConnDB = "Data Source=" & strsplit(i)
                Case Is = 1
                    If CH_Per.Checked Then
                        Connstring = Connstring
                    ElseIf Ch_mS.Checked Then
                        Connstring = Connstring & ";database=" & strsplit(i)
                    End If
                    ConnDB &= " ;Initial Catalog=" & strsplit(i)
                Case Is = 2
                    If Ch_mS.Checked Then
                        Connstring = Connstring & ";uid=" & strsplit(i)
                    ElseIf Ch_OOr.Checked Then
                        Connstring = Connstring & ";USER ID=" & strsplit(i)
                    End If
                    'ConnDB &= ";Persist Security Info=True;User ID=" & strsplit(i)

                    ConnDB &= ";User ID=" & strsplit(i)
                Case Is = 3
                    If Ch_mS.Checked Then
                        Connstring = Connstring & ";pwd=" & strsplit(i)
                    ElseIf Ch_OOr.Checked Then
                        Connstring = Connstring & ";PASSWORD=" & strsplit(i)
                    End If
                    ConnDB &= " ;Password=" & strsplit(i)
            End Select
        Next

        Dim Sqlstr As String = "select name from sys.databases order by name"
        Try
            cn.Open(Connstring & ";Connect Timeout=3;")
            rs.Open(Sqlstr, cn)
            For i As Integer = 0 To txtACCDBName.Items.Count - 1
                txtACCDBName.Items.RemoveAt(0)
            Next
            For i As Integer = 0 To rs.options.QueryDT.Rows.Count - 1
                txtACCDBName.Items.Add(rs.options.QueryDT.Rows(i)(0).ToString)
            Next
            cn.Close()
        Catch ex As Exception
            txtACCDBName.Items.Clear()
        End Try
        rs = Nothing
    End Sub

    Private Sub txtVATDBName_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVATDBName.DropDown
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim Connstring As String
        Dim strsplit As Object
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        If CH_Per.Checked Then
            ComSetting = txtVATServerName.Text & ";"
        ElseIf Ch_mS.Checked Then
            ComSetting = txtVATServerName.Text & ";master;" & txtVATUsername.Text & ";" & txtVATPassword.Text & ";"
        ElseIf Ch_OOr.Checked Then
            ComSetting = txtVATServerName.Text & ";master;" & txtVATUsername.Text & ";" & txtVATPassword.Text & ";"
        End If

        strsplit = Split(ComSetting, ";")
        Connstring = ""
        For i As Integer = LBound(strsplit) To UBound(strsplit)
            Select Case i
                Case Is = 0
                    If CH_Per.Checked Or (Ch_OOr.Checked And MODE = 2) Then
                        Connstring = "Provider=MSDASQL.1;Persist Security Info=False;" & "Data Source=" & strsplit(i)
                    ElseIf Ch_mS.Checked Then

                        ' Connstring = "driver={SQL Server};server=" & strsplit(i)
                        Connstring = "Provider=SQLOLEDB.1;server=" & strsplit(i)

                    ElseIf Ch_OOr.Checked Then
                        Connstring = "Provider=MSDAORA.1;Data Source=" & strsplit(i) & ";Persist Security Info=False"
                    End If
                    ConnDB = "Data Source=" & strsplit(i)
                Case Is = 1
                    If CH_Per.Checked Then
                        Connstring = Connstring
                    ElseIf Ch_mS.Checked Then
                        Connstring = Connstring & ";database=" & strsplit(i)
                    End If
                    ConnDB &= " ;Initial Catalog=" & strsplit(i)
                Case Is = 2
                    If Ch_mS.Checked Then
                        Connstring = Connstring & ";uid=" & strsplit(i)
                    ElseIf Ch_OOr.Checked Then
                        Connstring = Connstring & ";USER ID=" & strsplit(i)
                    End If
                    'ConnDB &= ";Persist Security Info=True;User ID=" & strsplit(i)

                    ConnDB &= ";User ID=" & strsplit(i)
                Case Is = 3
                    If Ch_mS.Checked Then
                        Connstring = Connstring & ";pwd=" & strsplit(i)
                    ElseIf Ch_OOr.Checked Then
                        Connstring = Connstring & ";PASSWORD=" & strsplit(i)
                    End If
                    ConnDB &= " ;Password=" & strsplit(i)
            End Select
        Next

        Dim Sqlstr As String = "select name from sys.databases order by name"
        Try
            cn.Open(Connstring & ";Connect Timeout=3;")
            rs.Open(Sqlstr, cn)
            For i As Integer = 0 To txtVATDBName.Items.Count - 1
                txtVATDBName.Items.RemoveAt(0)
            Next
            For i As Integer = 0 To rs.options.QueryDT.Rows.Count - 1
                txtVATDBName.Items.Add(rs.options.QueryDT.Rows(i)(0).ToString)
            Next
            cn.Close()
        Catch ex As Exception
            txtVATDBName.Items.Clear()
        End Try
        rs = Nothing
    End Sub


    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        Dim frmData As FrmLoginFMS = New FrmLoginFMS(CMD)
        frmData.ShowDialog()
        'If frmData.DialogResult <> Windows.Forms.DialogResult.Yes Then
        '    Application.Exit()
        '    Exit Sub
        'End If
    End Sub

End Class