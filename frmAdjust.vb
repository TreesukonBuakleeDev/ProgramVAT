Option Strict Off
Option Explicit On

Imports System.Globalization
Imports System.Runtime.InteropServices

Friend Class frmAdjust
    Inherits System.Windows.Forms.Form
    Public VATID As Integer
    Public VDate As String
    Public INVDATE As String
    Public IDINVC As String
    Public DOCNO As String
    Public RunNo As String
    Public NEWDOCNO As String
    Public INVNAME As String
    Public INVAMT As Double
    Public TaxAMT As Double
    'UPGRADE_NOTE: RATE was upgraded to RATE_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public RATE_Renamed As Double
    Public LOCID As String
    Public VTYPE As Integer
    Public TTYPE As Integer
    Public VATCOMMENT As String
    Public Source As String
    Public Batch As String
    Public Entry As String
    Public CBRef As String
    Public ClickSelect As Boolean
    Public DETAILNO As String
    Dim lbLoad As Boolean
    Dim FRM As frmMain
    Dim CMD_TYPE As String
    Public TypeofPU As String
    Public TranNo As String
    Public CIF As Decimal
    Public TaxCIF As Decimal
    Public Verify As Boolean
    Public TAXID As String
    Public BRANCH As String
    Public Currency As String
    Public SourceCurr As Double
    Public ExchangRate As Double
    Public Code As String
    'Public Code1 As String
    Dim fi As New DateTimeFormatInfo
    Dim myDateTran As Date '= DateTime.ParseExact(y & mm & "01", "yyyyMMdd", fi)
    Dim myDate As Date '= DateTime.ParseExact(y & mm & dd, "yyyyMMdd", fi)

    Private Sub checkAdjust()
        If lbLoad Then
            'cmdOK.Enabled = Len(txtDate.Value.ToString("yyyyMMdd")) > 0 And Len(txtInvNo.Text) > 0 And Len(txtName.Text) > 0 And Len(txtAmount.Text) > 0 And Len(txtTax.Text) > 0 And Len(Tx_Cif.Text) > 0 And Len(Tx_TaxCIF.Text) > 0
            If cboType.SelectedIndex = 0 Then
                Lb_TranNo.Text = "Type Of PU"
                Tx_Cif.Visible = False
                Tx_TaxCIF.Visible = False
                Lb_Cif.Visible = False
                Lb_taxCIF.Visible = False
                Tx_Cif.Value = 0
                txtSourceCurr.Enabled = False
                txtCurr.Enabled = False
                txtExchang.Enabled = False
                Label15.Visible = False
                Me.MinimumSize = New Size(497, 650)
                Me.MaximumSize = New Size(497, 650)
            Else
                Lb_TranNo.Text = "Transport No."
                Tx_Cif.Visible = True
                Tx_TaxCIF.Visible = True
                Lb_Cif.Visible = True
                Lb_taxCIF.Visible = True
                txtSourceCurr.Enabled = True
                txtCurr.Enabled = True
                txtExchang.Enabled = True
                Label15.Visible = True
                Me.MinimumSize = New Size(497, 650)
                Me.MaximumSize = New Size(497, 650)
            End If
        End If
    End Sub

    Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged
        checkAdjust()
    End Sub

    Private Sub chkNone_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkNone.CheckStateChanged
        checkAdjust()
    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCancel.Click
        ClickSelect = False
        Me.Hide()
    End Sub

    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSave.Click
        'cmdOK.Enabled = False
        If Trim(TxtTrans.Value.ToString("yyyyMMdd")) = "" Then
            ShowToolTip(TxtTrans, ToolTipIcon.Warning, "Trans. Date cannot be blank")
            Exit Sub
        Else
            INVDATE = TxtTrans.Value.ToString("yyyyMMdd")
        End If
        If Trim(txtDate.Value.ToString("yyyyMMdd")) = "" Then
            ShowToolTip(txtDate, ToolTipIcon.Warning, "Doc Date cannot be blank")
            Exit Sub
        Else
            VDate = txtDate.Value.ToString("yyyyMMdd")
        End If

        If Trim(txtInvNo.Text) = "" Then
            ShowToolTip(txtInvNo, ToolTipIcon.Warning, "Inv No. cannot be blank")
            Exit Sub
        Else
            IDINVC = txtInvNo.Text
        End If

        If Trim(txtName.Text) = "" Then
            ShowToolTip(txtName, ToolTipIcon.Warning, "Name Inv. cannot be blank")
            Exit Sub
        Else
            INVNAME = txtName.Text
        End If

        If Not (IsNumeric(txtAmount.Text)) Then
            ShowToolTip(txtAmount, ToolTipIcon.Warning, "Amount Not Is Numeric")
            Exit Sub
        End If

        If Not (IsNumeric(txtTax.Text)) Then
            ShowToolTip(txtTax, ToolTipIcon.Warning, "Amount Tax Not Is Numeric")
            Exit Sub
        End If

        If Not (IsNumeric(txtRate.Text)) Then
            ShowToolTip(txtRate, ToolTipIcon.Warning, "Rate Not Is Numeric")
            Exit Sub
        End If


        RunNo = txtRunno.Text
        DOCNO = txtDocNo.Text
        NEWDOCNO = txtNewDoc.Text

        INVAMT = CDbl(txtAmount.Text)
        TaxAMT = CDbl(txtTax.Text)
        If cboType.SelectedIndex = 0 Then
            TypeofPU = Tx_tranNo.Text
        Else
            TranNo = Tx_tranNo.Text
        End If
        CIF = Tx_Cif.Text
        TaxCIF = Tx_TaxCIF.Text
        If Trim(cboLocation.Text) = "" Then
            ShowToolTip(cboLocation, ToolTipIcon.Warning, "Location cannot be blank")
            Exit Sub
        Else
            LOCID = cboLocation.Text
        End If
        VTYPE = IIf(chkNone.CheckState = 0, 1, 0)
        TTYPE = VB6.GetItemData(cboType, cboType.SelectedIndex)
        TTYPE = cboType.SelectedIndex + 1
        RATE_Renamed = txtRate.Text
        VATCOMMENT = txtComment.Text
        CBRef = txtRef.Text
        ClickSelect = True
        Verify = Cb_Verify.Checked
        TAXID = TxTaxID.Text.Trim
        BRANCH = TxBranch.Text.Trim
        Currency = txtCurr.Text.Trim
        SourceCurr = CDbl(txtSourceCurr.Text)
        ExchangRate = txtExchang.Text


        Dim cn As ADODB.Connection
        Dim sqlstr As String
        Dim rs As BaseClass.BaseODBCIO
        cn = New ADODB.Connection
        If CMD_TYPE = "EDITVAT" Then
            'Mouse(MON)
            Dim loVATID, i As Integer
            loVATID = CDbl(FRM.lstVu.CurrentRow.Cells("PK").Value)
            i = CDbl(FRM.lstVu.CurrentRow.Cells(0).Value) - 1
            System.Windows.Forms.Cursor.Current = Cursors.AppStarting


            With Me
                '*********************************************************************************************
                '                              Update รายการ VAT ที่มีการแก้ไข
                '*********************************************************************************************
                'sqlstr = UpVat(CShort(Mid(loVATID, 4, Len(loVATID) - 3)), .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .RATE_Renamed, .VTYPE, .TTYPE, .VATCOMMENT, .CBRef, .INVDATE)


                sqlstr = FRM.UpVat(CDec(.VATID), .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .RATE_Renamed, .VTYPE, .TTYPE, .VATCOMMENT, .CBRef, .INVDATE, .RunNo, TypeofPU, TranNo, CIF, TaxCIF, IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH)

                cn.Open(ConnVAT)
                cn.Execute(sqlstr)
                '   cn.Close

                '*********************************************************************************************
                '               บันทึกรายการ VAT ที่มีการแก้ไข เพื่อใช้กรณี Print ย้อนหลัง
                '*********************************************************************************************
                cn = New ADODB.Connection
                rs = New BaseClass.BaseODBCIO

                sqlstr = " Select * from FMSVATINSERT  WHERE (SOURCE = '" & .Source & "')  "
                sqlstr = sqlstr & " AND (BATCH ='" & .Batch & "') AND (ENTRY = '" & .Entry & "') AND (LOCID = '" & .LOCID & "') AND (TRANSNBR='" & .DETAILNO & "')"

                cn.Open(ConnVAT)
                rs.Open(sqlstr, cn)
                '    cn.Close
                If rs.options.Rs.RecordCount > 0 And Cb_Verify.Checked = False Then
                    If rs.options.QueryDT.Rows(0).Item("verify").ToString.Trim = "1" Then
                        sqlstr = "delete from FMSVATINSERT  WHERE (SOURCE = '" & .Source & "')  "
                        sqlstr = sqlstr & " AND (BATCH ='" & .Batch & "') AND (ENTRY = '" & .Entry & "') AND (LOCID = '" & .LOCID & "') AND (TRANSNBR='" & .DETAILNO & "')"

                        sqlstr &= " Update FMSVAT SET VERIFY=0 WHERE (SOURCE = '" & .Source & "')  "
                        sqlstr = sqlstr & " AND (BATCH ='" & .Batch & "') AND (ENTRY = '" & .Entry & "') AND (LOCID = '" & .LOCID & "') AND (TRANSNBR='" & .DETAILNO & "')"
                    Else
                        sqlstr = FRM.UpEditVat(CDec(.VATID), .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .RATE_Renamed, .VTYPE, .TTYPE, .VATCOMMENT, .CBRef, .INVDATE, .Batch, .Entry, .Source, .DETAILNO, .RunNo, TypeofPU, TranNo, CIF, TaxCIF, IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH)
                    End If
                ElseIf rs.options.Rs.RecordCount > 0 Then
                    'sqlstr = UpEditVat(CShort(Mid(loVATID, 4, Len(loVATID) - 3)), .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .RATE_Renamed, .VTYPE, .TTYPE, .VATCOMMENT, .CBRef, .INVDATE, .Batch, .Entry, .Source, .DETAILNO)
                    sqlstr = FRM.UpEditVat(CDec(.VATID), .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .RATE_Renamed, .VTYPE, .TTYPE, .VATCOMMENT, .CBRef, .INVDATE, .Batch, .Entry, .Source, .DETAILNO, .RunNo, TypeofPU, TranNo, CIF, TaxCIF, IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH)
                Else
                    sqlstr = InsNewVat(CDec(.VATID), .INVDATE, .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, "", .Source, .Batch, .Entry, "", .VATCOMMENT, .CBRef, "Edit", .DETAILNO, .RunNo, TypeofPU, TranNo, CIF, TaxCIF, "", "", IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH, Currency, ExchangRate, SourceCurr)
                End If

                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(sqlstr)
                cn.Close()

                ' WriteLog
                WriteLog("Save Data Success")
                MessageBox.Show("Save Data Success", "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                mocolVatDisPlay.Item(loVATID).TXDATE = CDbl(Me.VDate)
                mocolVatDisPlay.Item(loVATID).INVDATE = CDbl(Me.INVDATE)
                mocolVatDisPlay.Item(loVATID).IDINVC = .IDINVC
                mocolVatDisPlay.Item(loVATID).DOCNO = .DOCNO
                mocolVatDisPlay.Item(loVATID).NEWDOCNO = .NEWDOCNO
                mocolVatDisPlay.Item(loVATID).INVNAME = .INVNAME
                mocolVatDisPlay.Item(loVATID).INVAMT = .INVAMT
                mocolVatDisPlay.Item(loVATID).INVTAX = .TaxAMT
                mocolVatDisPlay.Item(loVATID).LOCID = .LOCID
                mocolVatDisPlay.Item(loVATID).VTYPE = .VTYPE
                mocolVatDisPlay.Item(loVATID).TTYPE = .TTYPE
                mocolVatDisPlay.Item(loVATID).RATE_Renamed = .RATE_Renamed
                mocolVatDisPlay.Item(loVATID).VATCOMMENT = .VATCOMMENT
                mocolVatDisPlay.Item(loVATID).Source = .Source
                mocolVatDisPlay.Item(loVATID).Batch = .Batch
                mocolVatDisPlay.Item(loVATID).Entry = .Entry
                mocolVatDisPlay.Item(loVATID).CBRef = .CBRef
                mocolVatDisPlay.Item(loVATID).TypeOfPU = .TypeofPU
                mocolVatDisPlay.Item(loVATID).TranNo = .TranNo
                mocolVatDisPlay.Item(loVATID).CIF = .CIF
                mocolVatDisPlay.Item(loVATID).TaxCIF = .TaxCIF
                mocolVatDisPlay.Item(loVATID).Verify = Verify
                mocolVatDisPlay.Item(loVATID).TAXID = TAXID
                mocolVatDisPlay.Item(loVATID).Branch = BRANCH

                Dim myDate As Date
                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                fi.ShortDatePattern = "yyyyMMdd"
                myDate = DateTime.ParseExact(CStr(mocolVatDisPlay.Item(loVATID).INVDATE).Trim, "yyyyMMdd", fi)
                FRM.lstVu.CurrentRow.Cells("TXDATE").Value = myDate.ToString("dd/MM/yyyy")
                FRM.lstVu.CurrentRow.Cells("IDINVC").Value = mocolVatDisPlay.Item(loVATID).IDINVC
                FRM.lstVu.CurrentRow.Cells("DOCNO").Value = mocolVatDisPlay.Item(loVATID).DOCNO
                FRM.lstVu.CurrentRow.Cells("NEWDOCNO").Value = mocolVatDisPlay.Item(loVATID).NEWDOCNO
                FRM.lstVu.CurrentRow.Cells("INVNAME").Value = mocolVatDisPlay.Item(loVATID).INVNAME
                FRM.lstVu.CurrentRow.Cells("INVAMT").Value = mocolVatDisPlay.Item(loVATID).INVAMT.ToString("#,##0.00")
                FRM.lstVu.CurrentRow.Cells("INVTAX").Value = mocolVatDisPlay.Item(loVATID).INVTAX.ToString("#,##0.00")
                FRM.lstVu.CurrentRow.Cells("LOCID").Value = mocolVatDisPlay.Item(loVATID).LOCID
                ' FRM.lstVu.CurrentRow.Cells("VTYPE").Value = mocolVatDisPlay.Item(loVATID).VTYPE
                'FRM.lstVu.CurrentRow.Cells("TTYPE").Value = mocolVatDisPlay.Item(loVATID).TTYPE
                FRM.lstVu.CurrentRow.Cells("RATE").Value = mocolVatDisPlay.Item(loVATID).RATE_Renamed
                'FRM.lstVu.CurrentRow.Cells("COMMENT").Value = mocolVatDisPlay.Item(loVATID).VATCOMMENT
                FRM.lstVu.CurrentRow.Cells("Source").Value = mocolVatDisPlay.Item(loVATID).Source
                FRM.lstVu.CurrentRow.Cells("Batch").Value = mocolVatDisPlay.Item(loVATID).Batch
                FRM.lstVu.CurrentRow.Cells("Entry").Value = mocolVatDisPlay.Item(loVATID).Entry
                FRM.lstVu.CurrentRow.Cells("CBRef").Value = mocolVatDisPlay.Item(loVATID).CBRef
                FRM.lstVu.CurrentRow.Cells("TypeOfPU").Value = mocolVatDisPlay.Item(loVATID).TypeOfPU
                FRM.lstVu.CurrentRow.Cells("TranNo").Value = mocolVatDisPlay.Item(loVATID).TranNo
                FRM.lstVu.CurrentRow.Cells("CIF").Value = mocolVatDisPlay.Item(loVATID).CIF
                FRM.lstVu.CurrentRow.Cells("TaxCIF").Value = mocolVatDisPlay.Item(loVATID).TaxCIF
                FRM.lstVu.CurrentRow.Cells("Verify").Value = mocolVatDisPlay.Item(loVATID).Verify
                FRM.lstVu.CurrentRow.Cells("TAXID").Value = mocolVatDisPlay.Item(loVATID).TAXID
                FRM.lstVu.CurrentRow.Cells("BRANCH").Value = mocolVatDisPlay.Item(loVATID).Branch

                FRM.lstVu.CurrentRow.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 130)
            End With

        ElseIf CMD_TYPE = "ADDVAT" Then
            With Me
                cn = New ADODB.Connection
                rs = New BaseClass.BaseODBCIO

                If ComDatabase = "ORCL" Then
                    sqlstr = "SELECT  MAX(cast(TRANSNBR as number))+1  AS DETAILNO From FMSVATINSERT WHERE     (SOURCE = 'US') AND (STATUS='Add')"
                Else
                    sqlstr = "SELECT  MAX(cast(TRANSNBR as INT))+1  AS DETAILNO From FMSVATINSERT WHERE     (SOURCE = 'US') AND (STATUS='Add')"
                End If
                cn.Open(ConnVAT)
                rs.Open(sqlstr, cn)

                If rs.options.Rs.RecordCount = 0 Then DETAILNO = "1"
                Do While Not rs.options.Rs.EOF
                    If IsDBNull(rs.options.Rs.Fields("DETAILNO").Value) OrElse rs.options.Rs.Fields("DETAILNO").Value = 0 Then
                        DETAILNO = "1"
                    Else
                        DETAILNO = Trim(rs.options.Rs.Fields("DETAILNO").Value)
                    End If
                    rs.options.Rs.MoveNext()
                Loop

                '    cn.Close
                cn = New ADODB.Connection
                sqlstr = InsVat(1, .INVDATE, .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, "", "US", "999999", "99999", "", .VATCOMMENT, .CBRef, DETAILNO, .RunNo, "", "", TypeofPU, TranNo, CIF, TaxCIF, IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH, Code, Currency, ExchangRate, SourceCurr, 0)
                'sqlstr = StrSS

                cn.Open(ConnVAT)
                cn.Execute(sqlstr)
                '*********************************************************************************************
                '               บันทึกรายการ VAT ที่มีการ เพิ่ม รายการ
                '*********************************************************************************************

                sqlstr = InsNewVat(.VATID, .INVDATE, .VDate, .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .TaxAMT, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, "", .Source, "999999", CStr(99999), "", .VATCOMMENT, .CBRef, "Add", DETAILNO, .RunNo, TypeofPU, TranNo, CIF, TaxCIF, "", "", IIf(Cb_Verify.Checked, "1", "0"), TAXID, BRANCH, Currency, ExchangRate, SourceCurr)

                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(sqlstr)
                cn.Close()
                DETAILNO = ""
            End With

            WriteLog("Save Data Success")
            MessageBox.Show("Save Data Success", "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
            FRM.GetVatlist()
            FRM.ListShow()
        End If
        System.Windows.Forms.Cursor.Current = Cursors.Default

        'UPGRADE_NOTE: Object frmData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'UPGRADE_NOTE: Object cn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        cn = Nothing
        'If UseVisualStyleRendererCell = False Or UseVisualStyleRendererRow = False Then
        '    Exit Sub
        'End If
        'lstVu.UseVisualStyleRendererRow = False
        'lstVu.UseVisualStyleRendererCell = False
        Dim curentrow As Integer = 0
        If FRM.lstVu.CurrentRow IsNot Nothing Then
            curentrow = FRM.lstVu.CurrentRow.Index
            FRM.lstVu.CurrentCell = FRM.lstVu.Rows(curentrow).Cells(2)
            FRM.lstVu.CurrentCell.Selected = True
        End If
        'Call FRM.cmdRefreh_Click(Nothing, Nothing)

        'Me.Hide()
    End Sub

    Private Sub frmAdjust_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim i As Integer

        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        fi.ShortDatePattern = "yyyyMMdd"
        Dim myDateTran As Date
        Dim myDate As Date
        Try
            myDateTran = DateTime.ParseExact(INVDATE, "yyyyMMdd", fi)
        Catch ex As Exception
            myDateTran = Now
        End Try
        Try
            myDate = DateTime.ParseExact(VDate, "yyyyMMdd", fi)
        Catch ex As Exception
            myDate = Now
        End Try

        Cb_Verify.Visible = Use_Approved
        lbLoad = False
        TxtTrans.Value = myDateTran
        txtDate.Value = myDate
        txtInvNo.Text = IDINVC.Trim
        If Len(IDINVC) = 0 Then
            Me.Text = "Add vat record"
        Else
            Me.Text = "Edit vat record"
        End If
        Cb_Verify.Checked = Verify
        txtRunno.Text = IIf(RunNo Is Nothing, "", RunNo).Trim
        txtDocNo.Text = IIf(DOCNO Is Nothing, "", DOCNO).Trim
        txtNewDoc.Text = IIf(NEWDOCNO Is Nothing, "", NEWDOCNO).Trim
        txtName.Text = IIf(INVNAME Is Nothing, "", INVNAME).Trim
        txtAmount.Text = IIf(Format(INVAMT, "#,##0.000") Is Nothing, "", Format(INVAMT, "#,##0.000")).Trim
        txtTax.Text = IIf(Format(TaxAMT, "#,##0.000") Is Nothing, "", Format(TaxAMT, "#,##0.000")).Trim
        txtRate.Text = IIf(Format(RATE_Renamed, "#,##0.000") Is Nothing, "", Format(RATE_Renamed, "#,##0.000")).Trim
        txtComment.Text = IIf(VATCOMMENT Is Nothing, "", VATCOMMENT).Trim
        Label1.Text = IIf(DETAILNO Is Nothing, "", DETAILNO).Trim
        txtRef.Text = IIf(CBRef Is Nothing, "", CBRef).Trim
        chkNone.CheckState = IIf(VTYPE = 0, 1, 0)
        lblSource.Text = "Source : " & IIf(Source Is Nothing, "", Source).Trim
        lblBatch.Text = "Batch : " & IIf(Batch Is Nothing, "", Batch).Trim & "  Entry : " & IIf(Entry Is Nothing, "", Entry).Trim
        TxTaxID.Text = TAXID
        TxBranch.Text = BRANCH
        'txtCurr.Text = Code

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        mocolLocation = New colLocation
        sqlstr = "SELECT LOCID FROM FMSVLOC ORDER BY LOCID"
        cn.Open(ConnVAT)
        rs.Open(sqlstr, cn)

        cboLocation.Items.Clear()
        Do While Not rs.options.Rs.EOF
            cboLocation.Items.Add(IIf(iffs.ISNULL(rs.options.Rs, 0), "", (rs.options.Rs.Collect(0))))
            rs.options.Rs.MoveNext()
        Loop
        For I As Integer = 0 To cboLocation.Items.Count - 1
            If cboLocation.Items(I).ToString.Trim = LOCID.Trim Then
                cboLocation.SelectedIndex = I
                Exit For
            End If
        Next

        If cboLocation.SelectedIndex = -1 Then
            cboLocation.SelectedIndex = 0 ' SendMessage(cboLocation.Handle.ToInt32, CB_FINDSTRING, 0, LOCID)
            'Call RefreshVatView(cboType)
            'cboType.SelectedIndex = TTYPE - 1
            cboType.SelectedIndex = FRM.cboView.SelectedIndex
        End If

        If cboType.SelectedIndex = 0 Then
            Tx_tranNo.Text = TypeofPU
        Else
            Tx_tranNo.Text = TranNo
        End If
        Tx_TaxCIF.Text = TaxCIF
        Tx_Cif.Text = CIF
        rs = Nothing
        cn = Nothing
        lbLoad = True
        checkAdjust()
        If FRM.lstVu.CurrentRow IsNot Nothing Then
            If FRM.lstVu.CurrentRow.Index = 0 Then
                BT_Previous.Enabled = False
            Else
                BT_Previous.Enabled = True
            End If
            If FRM.lstVu.CurrentRow.Index + 1 = FRM.lstVu.RowCount Then
                BT_Next.Enabled = False
            Else
                BT_Next.Enabled = True
            End If
        Else
            BT_Next.Enabled = False
            BT_Previous.Enabled = False
        End If
    End Sub

    Private Sub CallRate()
        If lbLoad Then
            If CDbl(txtTax.Text) <> 0 And CDbl(txtAmount.Text) <> 0 Then
                txtRate.Text = CStr((CDbl(txtTax.Text) / CDbl(txtAmount.Text)) * 100)
            Else
                txtRate.Text = CStr(0)
            End If
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.000")
        End If
    End Sub

    Public Sub TextboxNumberic(ByRef Obj As NumericUpDown, ByRef e As System.Windows.Forms.KeyPressEventArgs)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 45 To 46, 48 To 57, System.Windows.Forms.Keys.Back
                KeyAscii = KeyAscii
            Case Else
                Beep()
                KeyAscii = 0
        End Select
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDate_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtDate.MouseWheel
        If txtDate.Focused Then
            If e.Delta > 0 Then
                SendKeys.Send("{UP}")
            Else
                SendKeys.Send("{DOWN}")
            End If
        End If
    End Sub

    Private Sub TxtTrans_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TxtTrans.MouseWheel
        If TxtTrans.Focused Then
            If e.Delta > 0 Then
                SendKeys.Send("{UP}")
            Else
                SendKeys.Send("{DOWN}")
            End If
        End If
    End Sub

    Private Sub txtDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDate.KeyPress
        If Not (IsNumeric(e.KeyChar)) Then e.Handled = True
    End Sub

    Private Sub txtDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDate.ValueChanged
        checkAdjust()
    End Sub

    Private Sub txtDocNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocNo.TextChanged
        checkAdjust()
    End Sub


    Private Sub txtInvNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvNo.TextChanged
        checkAdjust()
    End Sub

    Private Sub txtName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtName.TextChanged, TxTaxID.TextChanged, TxBranch.TextChanged
        checkAdjust()
    End Sub

    Private Sub txtNewDoc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNewDoc.TextChanged
        checkAdjust()
    End Sub

    Private Sub txtRate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.TextChanged
        checkAdjust()
    End Sub

    Private Sub txtRate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.Leave
        If Not (IsNumeric(txtRate.Text)) Then
            ShowToolTip(txtRate, ToolTipIcon.Warning, "Rate Not Is Numeric")
            Exit Sub
        End If
        If CheckMoneyValue(txtRate.Text) = True Then
            txtRate.Text = Format(CDbl(txtRate.Text), "#,##0.000")
            Exit Sub
        Else
            txtRate.Text = CStr(0)
            ShowToolTip(txtRate, ToolTipIcon.Warning, "Wrong format Amount")
        End If


        'Dim Txt As NumericUpDown = CType(eventSender, NumericUpDown)
        'If Not (IsNumeric(Txt.Text)) Then
        '    ShowToolTip(Txt, ToolTipIcon.Warning, "Rate Not Is Numeric")
        '    Exit Sub
        'End If
        'If CheckMoneyValue(Txt.Text) = True Then
        '    Txt.Text = Format(CDbl(Txt.Text), "#,##0.000")
        '    Call CallRate()
        '    Exit Sub
        'Else
        '    Txt.Text = CStr(0)
        '    ShowToolTip(Txt, ToolTipIcon.Warning, "Wrong format Amount")
        'End If

    End Sub
    Private Sub txtExchang_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExchang.Leave
        If Not (IsNumeric(txtExchang.Text)) Then
            ShowToolTip(txtExchang, ToolTipIcon.Warning, "Exchange Rate Not Is Numeric")
            Exit Sub
        End If
        If CheckMoneyValue(txtExchang.Text) = True Then
            txtExchang.Text = Format(CDbl(txtExchang.Text), "#,##0.0000000")
            Exit Sub
        Else
            txtExchang.Text = CStr(0)
            ShowToolTip(txtExchang, ToolTipIcon.Warning, "Wrong format Amount")
        End If

    End Sub
   

    Function CheckMoneyValue(ByRef myValue As Double) As Boolean
        On Error GoTo ErrHandler
        myValue = Format(CDbl(myValue), "#,##0.0000000")
        CheckMoneyValue = True
        Exit Function
ErrHandler:
        CheckMoneyValue = False
    End Function

    Private Sub cboLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLocation.SelectedIndexChanged
        LOCID = cboLocation.Items(cboLocation.SelectedIndex)
        Call RefreshVatView(cboType)
        cboType.SelectedIndex = TTYPE - 1
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTax.KeyPress, txtAmount.KeyPress, Tx_TaxCIF.KeyPress, Tx_Cif.KeyPress, txtSourceCurr.KeyPress, txtExchang.KeyPress
        TextboxNumberic(sender, e)
    End Sub

    Private Sub txtAmount_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTax.Leave, txtAmount.Leave, Tx_TaxCIF.Leave, Tx_Cif.Leave, txtSourceCurr.Leave
        Dim Tx As NumericUpDown = CType(sender, NumericUpDown)
        If Not (IsNumeric(Tx.Text)) Then
            ShowToolTip(Tx, ToolTipIcon.Warning, "Amount Not Is Numeric")
            Exit Sub
        End If
        If CheckMoneyValue(Tx.Text) = True Then
            Tx.Text = Format(CDbl(Tx.Text), "#,##0.000")
            Call CallRate()
            Exit Sub
        Else
            Tx.Text = CStr(0)
            ShowToolTip(Tx, ToolTipIcon.Warning, "Wrong format Amount")
        End If
    End Sub

    Private Sub Tx_Cif_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTax.ValueChanged, txtAmount.ValueChanged, Tx_TaxCIF.ValueChanged, Tx_Cif.ValueChanged
        checkAdjust()
        If CType(sender, NumericUpDown).Name = "Tx_Cif" Then
            Tx_TaxCIF.Value = Tx_Cif.Value * 7 / 100
        End If
        txtExchang.Text = FormatNumber((txtAmount.Text / txtSourceCurr.Text), 7, TriState.UseDefault)

    End Sub

    Public Sub ReadData(ByVal i As Integer)
        With Me
            .VATID = mocolVatDisPlay.Item(i).VATID
            .VDate = CStr(mocolVatDisPlay.Item(i).TXDATE)
            .INVDATE = CStr(mocolVatDisPlay.Item(i).INVDATE)
            .IDINVC = mocolVatDisPlay.Item(i).IDINVC
            .DOCNO = mocolVatDisPlay.Item(i).DOCNO
            .NEWDOCNO = mocolVatDisPlay.Item(i).NEWDOCNO
            .INVNAME = mocolVatDisPlay.Item(i).INVNAME
            .INVAMT = mocolVatDisPlay.Item(i).INVAMT
            .TaxAMT = mocolVatDisPlay.Item(i).INVTAX
            .LOCID = mocolVatDisPlay.Item(i).LOCID
            .VTYPE = mocolVatDisPlay.Item(i).VTYPE
            .TTYPE = mocolVatDisPlay.Item(i).TTYPE
            .RATE_Renamed = mocolVatDisPlay.Item(i).RATE_Renamed
            .VATCOMMENT = Trim(mocolVatDisPlay.Item(i).VATCOMMENT)
            .Source = mocolVatDisPlay.Item(i).Source
            .Batch = mocolVatDisPlay.Item(i).Batch
            .Entry = mocolVatDisPlay.Item(i).Entry
            .CBRef = mocolVatDisPlay.Item(i).CBRef
            .DETAILNO = mocolVatDisPlay.Item(i).TRANSNBR
            .RunNo = mocolVatDisPlay.Item(i).Runno
            .TypeofPU = mocolVatDisPlay.Item(i).TypeOfPU
            .TranNo = mocolVatDisPlay.Item(i).TranNo
            .CIF = mocolVatDisPlay.Item(i).CIF
            .TaxCIF = mocolVatDisPlay.Item(i).TaxCIF
            .Verify = mocolVatDisPlay.Item(i).Verify
            .TAXID = mocolVatDisPlay.Item(i).TAXID
            .BRANCH = mocolVatDisPlay.Item(i).Branch
        End With
        frmAdjust_Load(Nothing, Nothing)
    End Sub
    Private Sub BT_Previous_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Previous.Click
        FRM.lstVu.CurrentCell = FRM.lstVu.Rows(FRM.lstVu.CurrentRow.Index - 1).Cells(2)
        FRM.lstVu.CurrentCell.Selected = True
        ReadData(FRM.lstVu.CurrentRow.Cells("PK").Value)

        If FRM.lstVu.CurrentRow.Index = 0 Then
            BT_Previous.Enabled = False
        Else
            BT_Previous.Enabled = True
        End If
        If FRM.lstVu.CurrentRow.Index + 1 = FRM.lstVu.RowCount Then
            BT_Next.Enabled = False
        Else
            BT_Next.Enabled = True
        End If
    End Sub

    Private Sub BT_Next_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Next.Click
        FRM.lstVu.CurrentCell = FRM.lstVu.Rows(FRM.lstVu.CurrentRow.Index + 1).Cells(2)
        FRM.lstVu.CurrentCell.Selected = True
        ReadData(FRM.lstVu.CurrentRow.Cells("PK").Value)

        If FRM.lstVu.CurrentRow.Index = 0 Then
            BT_Previous.Enabled = False
        Else
            BT_Previous.Enabled = True
        End If
        If FRM.lstVu.CurrentRow.Index + 1 = FRM.lstVu.RowCount Then
            BT_Next.Enabled = False
        Else
            BT_Next.Enabled = True
        End If
    End Sub

    Private Sub txtSourceCurr_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSourceCurr.ValueChanged
        txtExchang.Text = FormatNumber((txtAmount.Text / txtSourceCurr.Text), 7, TriState.UseDefault)
    End Sub

End Class