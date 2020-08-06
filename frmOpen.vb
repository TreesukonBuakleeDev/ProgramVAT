Option Strict Off
Option Explicit On
Friend Class frmOpen
    Inherits System.Windows.Forms.Form
    Public ClickSelect As Boolean

    Private Sub cboCompany_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCompany.SelectedIndexChanged
        CheckChange()
    End Sub

    Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOpen.Click
        Dim strsplit() As String
        Dim strVatPartial As String
        Dim strVatName As String
        Dim strVatAutoRun As String
        Dim sql As String = ""
        Dim ConACCPAC As String = ""
        Dim rs As BaseClass.BaseODBCIO
        Dim cnACCPAC As ADODB.Connection


        'Get INI COM Setting
        ComID = CStr(cboCompany.SelectedIndex)
        ComDatabase = GetPrivateProfile("COM" & ComID, "DATABASE", "")
        Use_LEGALNAME_inCB = GetPrivateProfile("COM" & ComID, "Use_LEGALNAME_inCB", "0")
        Use_LEGALNAME_inAP = GetPrivateProfile("COM" & ComID, "Use_LEGALNAME_inAP", "0")
        Use_LEGALNAME_inAR = GetPrivateProfile("COM" & ComID, "Use_LEGALNAME_inAR", "0")
        Use_POST = GetPrivateProfile("COM" & ComID, "GetType", "1")
        Use_DisplayPage = GetPrivateProfile("COM" & ComID, "DISPLAYPAGE", "5")
        Use_Approved = GetPrivateProfile("COM" & ComID, "Verify", "0")
        DATEMODE = GetPrivateProfile("COM" & ComID, "DATEMODE", "DOCU")
        Use_INVFROMPO = GetPrivateProfile("COM" & ComID, "INVFROMPO", "0")
        strsplit = Split(ComDatabase, ";")
        ComDatabase = strsplit(0)
        GLREVERSE = GetPrivateProfile("COM" & ComID, "GLREVERSE", "FALSE")
        CBREVERSE = GetPrivateProfile("COM" & ComID, "CBREVERSE", "FALSE")

        Use_TAX_AP = GetPrivateProfile("COM" & ComID, "Use_TAX_AP", "0")
        Use_TAX_AR = GetPrivateProfile("COM" & ComID, "Use_TAX_AR", "0")
        Use_TAX_CB = GetPrivateProfile("COM" & ComID, "Use_TAX_CB", "0")
        Use_TAX_GL = GetPrivateProfile("COM" & ComID, "Use_TAX_GL", "0")

        Use_BRANCH_AP = GetPrivateProfile("COM" & ComID, "Use_BRANCH_AP", "0")
        Use_BRANCH_AR = GetPrivateProfile("COM" & ComID, "Use_BRANCH_AR", "0")
        Use_BRANCH_CB = GetPrivateProfile("COM" & ComID, "Use_BRANCH_CB", "0")
        Use_BRANCH_GL = GetPrivateProfile("COM" & ComID, "Use_BRANCH_GL", "0")

        S_AP = GetPrivateProfile("COM" & ComID, "SHOWAP", "TRUE")
        S_AR = GetPrivateProfile("COM" & ComID, "SHOWAR", "TRUE")
        S_GL = GetPrivateProfile("COM" & ComID, "SHOWGL", "TRUE")
        S_CB = GetPrivateProfile("COM" & ComID, "SHOWCB", "TRUE")
        S_PO = GetPrivateProfile("COM" & ComID, "SHOWPO", "TRUE")
        s_OE = GetPrivateProfile("COM" & ComID, "SHOWOE", "TRUE")

        strVatPartial = GetPrivateProfile("COM" & ComID, "VATPAR", "Y")
        strsplit = Split(strVatPartial, ";")
        ComVatPartial = IIf(strsplit(0) = "Y", True, False)

        'Config Vat Autorunno
        strVatAutoRun = GetPrivateProfile("COM" & ComID, "VATRUNMANUAL", "N")
        strsplit = Split(strVatAutoRun, ";")
        VATAutoRun = IIf(strsplit(0) = "Y", True, False)


        strVatName = GetPrivateProfile("COM" & ComID, "SHOWNAMEONLY", "Y")
        strsplit = Split(CStr(strVatName), ";")
        ComVatName = IIf(strsplit(0) = "Y", True, False)

        lblName.Text = GetPrivateProfile("COM" & ComID, "ACCPAC", "")
        ComACCPAC = IIf(Len(lblName.Text) > 0, lblName.Text, "")

        lblName.Text = GetPrivateProfile("COM" & ComID, "VAT", "")
        ComVAT = IIf(Len(lblName.Text) > 0, lblName.Text, "")

        OPT_AP = GetPrivateProfile("COM" & ComID, "OPT_AP", "FALSE")
        OPT_CB = GetPrivateProfile("COM" & ComID, "OPT_CB", "FALSE")
        OPT_GL = GetPrivateProfile("COM" & ComID, "OPT_GL", "FALSE")

        VNO_AP = GetPrivateProfile("COM" & ComID, "VNO_AP", "FALSE")
        VNO_AR = GetPrivateProfile("COM" & ComID, "VNO_AR", "FALSE")
        VNO_CB = GetPrivateProfile("COM" & ComID, "VNO_CB", "FALSE")
        VNO_GL = GetPrivateProfile("COM" & ComID, "VNO_GL", "FALSE")

        RERUN = GetPrivateProfile("COM" & ComID, "RERUN", "FALSE")
        UPRUN = GetPrivateProfile("COM" & ComID, "UPRUN", "FALSE")

        FormatRunnig = GetPrivateProfile("COM" & ComID, "FormatRunnig", "yy+MM+/")
        VATFormat = GetPrivateProfile("COM" & ComID, "FORMAT", "00000")

        ClickSelect = True
        ClickSelect1 = True
        DialogResult = Windows.Forms.DialogResult.Yes

        Dim conacc() As String = ComACCPAC.Split(";")

        'ConACCPAC = "Provider=SQLOLEDB.1;server=" & conacc(0) & ";database=" & conacc(1) & ";uid=" & conacc(2) & ";pwd=" & conacc(3) & ""
        'cnACCPAC = New ADODB.Connection
        'cnACCPAC.ConnectionTimeout = 15
        'cnACCPAC.Open(ConACCPAC)
        'cnACCPAC.CommandTimeout = 15



        Me.Close()
    End Sub

    Private Sub frmOpen_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim i As Short
        Dim ComList As Short
        Dim ComName As String
        'Get INI list
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        ComList = GetPrivateProfile("COMPANY", "LIST", "0").ToString.Replace(";", "")
        If ComList = 0 Then
            Dim frm As New FrmConfigDB("New")
            fmsadmin = "FMSADMIN"
            frm.ShowDialog(Me)
            If frm.DialogResult <> Windows.Forms.DialogResult.Yes Then
                'Application.Exit()
                Me.Close()
                Exit Sub
            End If
            ComList = CShort(GetPrivateProfile("COMPANY", "LIST", "0"))
            fmsadmin = ""
        End If

        If Val(CStr(ComList)) > 0 Then
            For i = 0 To Val(CStr(ComList))
                ComName = GetPrivateProfile("COMPANY", "COM" & CStr(i), "(Please select)")
                cboCompany.Items.Add(New VB6.ListBoxItem(Replace(ComName, ";", ""), i))
            Next
            cboCompany.SelectedIndex = 0
        Else
            Me.Close()
        End If

        ClickSelect = False
        ClickSelect1 = False
        Me.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
        CheckChange()
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub CheckChange()
        If cboCompany.SelectedIndex > 0 Then
            btnOpen.Enabled = True
        Else
            btnOpen.Enabled = False
        End If
    End Sub

    'Private Sub BtNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
    '    'Dim frm As New FrmConfigDB("New")
    '    'frm.ShowDialog(Me)
    '    '' If frm.DialogResult = Windows.Forms.DialogResult.Yes Then
    '    'cboCompany.Items.Clear()
    '    'frmOpen_Load(Nothing, Nothing)
    '    '' End If
    'End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()

    End Sub
End Class