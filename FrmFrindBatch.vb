Imports System.Data.SqlClient
Public Class FrmFrindBatch
    Dim ds As DataSet = New DataSet
    Dim tmpBatchid As String = ""
    Dim BATCHSTAT As String = ""
    Public _id As String

    Public Property b_id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            _id = value
        End Set
    End Property

    Function Checksyntax() As Boolean
        Dim bool As Boolean
        If cboFindBy.SelectedIndex = 1 Or cboFindBy.SelectedIndex = 2 Then

            If Me.txtFilter.Text = "" Then

                bool = False
            Else
                bool = True
            End If
        Else
            bool = True
        End If

        Return bool

    End Function
    Public Sub New(ByVal tmpBatchid_ As String)

        InitializeComponent()
        tmpBatchid = tmpBatchid_

    End Sub

    Private Sub FrmFrindBatch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "Search Batch"
        ShowData()
        cboFindBy.SelectedIndex = 0
        dgvFrindBath.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells)
        dgvFrindBath.Columns("Batch Number").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Number of Entries").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Bank Exchange Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Posting Sequence No.").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Batch Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Func. Batch Total").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvFrindBath.Columns("Batch Number").DefaultCellStyle.Format = "#,##0"
        dgvFrindBath.Columns("Batch Total").DefaultCellStyle.Format = "#,##0.000"
        dgvFrindBath.Columns("Func. Batch Total").DefaultCellStyle.Format = "#,##0.000"
        dgvFrindBath.Columns("Bank Exchange Rate").DefaultCellStyle.Format = "#,##0.0000000"

        For i As Integer = 0 To dgvFrindBath.Rows.Count - 1

            If Convert.ToDecimal(dgvFrindBath.Rows(i).Cells("Batch Number").Value.ToString()).ToString("###0") = tmpBatchid Then

                dgvFrindBath.FirstDisplayedScrollingRowIndex = i
                dgvFrindBath.CurrentCell = dgvFrindBath.Rows(i).Cells("Batch Number")
                dgvFrindBath.Rows(i).Selected = True

            End If
        Next


    End Sub


    Private Sub ShowData()

        Dim strSql As String = ""
        'Cons.GetConString()
        'Cons.SetData()
        If dgvFrindBath.Rows.Count > 0 Then
            If ds.Tables("T").Rows.Count > 0 Then

                ds.Tables("T").Clear()
            End If
        End If


        If cboFindBy.SelectedIndex = 1 Then


    
            strSql = "SELECT    GLPOST.BATCHNBR , GLPOST.ENTRYNBR, GLPOST.JRNLDATE,GLPJC.COMMENT," '0-3
            strSql &= "         GLPOST.JNLDTLREF,GLPOST.JNLDTLDESC,GLPOST.TRANSQTY," '4-6
            strSql &= "         GLPOST.TRANSAMT,GLPOST.TRANSNBR," '7-9
            strSql &= "         GLPOST.POSTINGSEQ,GLPOST.CNTDETAIL,GLPOST.FISCALYR,GLPOST.FISCALPERD" '10-13
            strSql &= " FROM    ((GLPOST "
            strSql &= "         LEFT OUTER JOIN GLPJC ON (GLPOST.TRANSNBR = GLPJC.TRANSNBR) AND (GLPOST.ENTRYNBR = GLPJC.ENTRYNBR) AND (GLPOST.BATCHNBR = GLPJC.BATCHNBR)) "

            If ComDatabase = "PVSW" Then
                strSql &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV')"
                strSql &= "         AND (GLPOST.BATCHNBR >= '" & DATEFROM & "' AND GLPOST.BATCHNBR <= '" & DATETO & "')"

            Else
                strSql &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV') "
                strSql &= "         AND (GLPOST.BATCHNBR >= '" & DATEFROM & "' AND GLPOST.BATCHNBR <= '" & DATETO & "')"

            End If

           

            'strSql &= " GROUP BY GLPOST.BATCHNBR, GLPOST.ENTRYNBR, GLPOST.JRNLDATE, GLPJC.COMMENT, GLPOST.JNLDTLREF, GLPOST.JNLDTLDESC, GLPOST.TRANSQTY, " & _
            '          " GLPOST.TRANSAMT, GLAMF.ACCTFMTTD, GLPOST.TRANSNBR, GLPOST.POSTINGSEQ, GLPOST.CNTDETAIL, GLPOST.FISCALYR,  " & _
            '          " GLPOST.FISCALPERD, GLJEDO.VALUE,FMSVLACC.VTYPE"




            strSql = "  select CNTBTCH as [Batch Number],BATCHDESC as[Description],convert(date ,convert(varchar(20),AUDTDATE)) as[Batch Date] "
            strSql &= " ,case BATCHTYPE when 3 then 'Generated' when 1 then 'Entered' when 2 then 'Imported' when 4 then 'System' when 5 then 'External' end  as[Batch Type] "
            strSql &= " ,case BATCHSTAT when 1 then 'Open' when 3 then 'Posted' when 4 then 'Deleted' when 7 then 'Ready To Post' when 8 then 'Check Creation In Progress' else 'Post In Progress' end as [Batch Status] "
            strSql &= " ,SRCEAPPL as [Source Application] "
            strSql &= " ,CNTENTER as[Number of Entries],AMTENTER as [Batch Total],IDBANK as [Bank Code] "
            strSql &= " ,CODECURN as [Bank Currency Code],convert(date,convert(varchar(20),DATERATE )) as [Bank Rate Date],RATEEXCHHC as [Bank Exchange Rate] "
            strSql &= " ,case SWPRINTED when 1 then 'Yes' when 0 then 'No' end as [Batch Printed Flag] "
            strSql &= " ,POSTSEQNBR as[Posting Sequence No.],FUNCAMOUNT as [Func. Batch Total] "
            strSql &= " from APBTA "
            strSql &= " where cast(CNTBTCH as VARCHAR(10)) " + cboStatement.Text + " '" + txtFilter.Text + "'"
            strSql &= " and (ltrim(rtrim(PAYMTYPE)) = 'PY') " + BATCHSTAT + " Order by CNTBTCH,AUDTDATE"

            If ds.Tables("T").Rows.Count <> 0 Then
                ds.Tables.Remove("T")
            End If

            ds.Tables.Add(ExQuery(strSql, "T"))

        ElseIf cboFindBy.SelectedIndex = 2 And txtFilter.Text <> "" Then

            If cboStatement.SelectedIndex = 0 Then

                strSql = "  select CNTBTCH as [Batch Number],BATCHDESC as[Description],convert(date ,convert(varchar(20),AUDTDATE)) as[Batch Date] "
                strSql &= " ,case BATCHTYPE when 3 then 'Generated' when 1 then 'Entered' when 2 then 'Imported' when 4 then 'System' when 5 then 'External' end  as[Batch Type] "
                strSql &= " ,case BATCHSTAT when 1 then 'Open' when 3 then 'Posted' when 4 then 'Deleted' when 7 then 'Ready To Post' when 8 then 'Check Creation In Progress' else 'Post In Progress' end as [Batch Status] "
                strSql &= " ,SRCEAPPL as [Source Application] "
                strSql &= " ,CNTENTER as[Number of Entries],AMTENTER as [Batch Total],IDBANK as [Bank Code] "
                strSql &= " ,CODECURN as [Bank Currency Code],convert(date,convert(varchar(20),DATERATE )) as [Bank Rate Date],RATEEXCHHC as [Bank Exchange Rate] "
                strSql &= " ,case SWPRINTED when 1 then 'Yes' when 0 then 'No' end as [Batch Printed Flag] "
                strSql &= " ,POSTSEQNBR as[Posting Sequence No.],FUNCAMOUNT as [Func. Batch Total] "
                strSql &= " from APBTA "
                strSql &= " where BATCHDESC like '" + txtFilter.Text + "%'"
                strSql &= " and (ltrim(rtrim(PAYMTYPE)) = 'PY') " + BATCHSTAT + " Order by CNTBTCH,AUDTDATE"

                If ds.Tables("T").Rows.Count <> 0 Then

                    ds.Tables.Remove("T")
                End If

                ds.Tables.Add(ExQuery(strSql, "T"))

            Else

                strSql = "  select CNTBTCH as [Batch Number],BATCHDESC as[Description],convert(date ,convert(varchar(20),AUDTDATE)) as[Batch Date] "
                strSql &= " ,case BATCHTYPE when 3 then 'Generated' when 1 then 'Entered' when 2 then 'Imported' when 4 then 'System' when 5 then 'External' end  as[Batch Type] "
                strSql &= " ,case BATCHSTAT when 1 then 'Open' when 3 then 'Posted' when 4 then 'Deleted' when 7 then 'Ready To Post' when 8 then 'Check Creation In Progress' else 'Post In Progress' end as [Batch Status] "
                strSql &= " ,SRCEAPPL as [Source Application] "
                strSql &= " ,CNTENTER as[Number of Entries],AMTENTER as [Batch Total],IDBANK as [Bank Code] "
                strSql &= " ,CODECURN as [Bank Currency Code],convert(date,convert(varchar(20),DATERATE )) as [Bank Rate Date],RATEEXCHHC as [Bank Exchange Rate] "
                strSql &= " ,case SWPRINTED when 1 then 'Yes' when 0 then 'No' end as [Batch Printed Flag] "
                strSql &= " ,POSTSEQNBR as[Posting Sequence No.],FUNCAMOUNT as [Func. Batch Total] "
                strSql &= " from APBTA "
                strSql &= " where BATCHDESC like '%" + txtFilter.Text + "%'"
                strSql &= " and (ltrim(rtrim(PAYMTYPE)) = 'PY') " + BATCHSTAT + " Order by CNTBTCH,AUDTDATE"

                If ds.Tables("T").Rows.Count <> 0 Then
                    ds.Tables.Remove("T")
                End If

                ds.Tables.Add(ExQuery(strSql, "T"))

            End If

        ElseIf cboFindBy.SelectedIndex = 3 Then

            strSql = "  select CNTBTCH as [Batch Number],BATCHDESC as[Description],convert(date ,convert(varchar(20),AUDTDATE)) as[Batch Date] "
            strSql &= " ,case BATCHTYPE when 3 then 'Generated' when 1 then 'Entered' when 2 then 'Imported' when 4 then 'System' when 5 then 'External' end  as[Batch Type] "
            strSql &= " ,case BATCHSTAT when 1 then 'Open' when 3 then 'Posted' when 4 then 'Deleted' when 7 then 'Ready To Post' when 8 then 'Check Creation In Progress' else 'Post In Progress' end as [Batch Status] "
            strSql &= " ,SRCEAPPL as [Source Application] "
            strSql &= " ,CNTENTER as[Number of Entries],AMTENTER as [Batch Total],IDBANK as [Bank Code] "
            strSql &= " ,CODECURN as [Bank Currency Code],convert(date,convert(varchar(20),DATERATE )) as [Bank Rate Date],RATEEXCHHC as [Bank Exchange Rate] "
            strSql &= " ,case SWPRINTED when 1 then 'Yes' when 0 then 'No' end as [Batch Printed Flag] "
            strSql &= " ,POSTSEQNBR as[Posting Sequence No.],FUNCAMOUNT as [Func. Batch Total] "
            strSql &= " from APBTA "
            strSql &= " where AUDTDATE " + cboStatement.Text + " '" + dtpDate.Value.ToString("yyyyMMdd") + "'"
            strSql &= " and (ltrim(rtrim(PAYMTYPE)) = 'PY') " + BATCHSTAT + " Order by CNTBTCH,AUDTDATE"

            If ds.Tables("T").Rows.Count <> 0 Then
                ds.Tables.Remove("T")
            End If
            ds.Tables.Add(ExQuery(strSql, "T"))

        Else

            strSql = "  select CNTBTCH as [Batch Number],BATCHDESC as[Description],convert(date ,convert(varchar(20),AUDTDATE)) as[Batch Date] "
            strSql &= " ,case BATCHTYPE when 3 then 'Generated' when 1 then 'Entered' when 2 then 'Imported' when 4 then 'System' when 5 then 'External' end  as[Batch Type] "
            strSql &= " ,case BATCHSTAT when 1 then 'Open' when 3 then 'Posted' when 4 then 'Deleted' when 7 then 'Ready To Post' when 8 then 'Check Creation In Progress' else 'Post In Progress' end as [Batch Status] "
            strSql &= " ,SRCEAPPL as [Source Application] "
            strSql &= " ,CNTENTER as[Number of Entries],AMTENTER as [Batch Total],IDBANK as [Bank Code] "
            strSql &= " ,CODECURN as [Bank Currency Code],convert(date,convert(varchar(20),DATERATE )) as [Bank Rate Date],RATEEXCHHC as [Bank Exchange Rate] "
            strSql &= " ,case SWPRINTED when 1 then 'Yes' when 0 then 'No' end as [Batch Printed Flag] "
            strSql &= " ,POSTSEQNBR as[Posting Sequence No.],FUNCAMOUNT as [Func. Batch Total] "
            strSql &= " from APBTA "
            strSql &= " where (ltrim(rtrim(PAYMTYPE)) = 'PY') " + BATCHSTAT + " Order by CNTBTCH,AUDTDATE"

            If ds.Tables("T").Rows.Count <> 0 Then
                ds.Tables.Remove("T")
            End If
            ds.Tables.Add(ExQuery(strSql, "T"))

        End If


        Me.dgvFrindBath.DataSource = ds.Tables("T")

        If dgvFrindBath.Rows.Count > 0 Then

            If ISNULL(dgvFrindBath.Rows(0).Cells(1).Value, "") = "" Then

                txtFilter.Text = ""
                MessageBox.Show("Exception", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtFilter.Focus()

            End If
        End If

    End Sub

    Public Function ExQuery(ByVal sql As String, ByVal name As String) As DataTable

        Try

            Call Connecttion("ACC")
            Dim com As SqlCommand = New SqlCommand(sql, ConA)
            Dim dt As DataTable = New DataTable(name)
            Dim da As SqlDataAdapter = New SqlDataAdapter(com)

            da.Fill(dt)
            com.Clone()
            com.Dispose()
            da.Dispose()
            Return dt

        Catch ex As Exception

            Call Connecttion("ACC")
            Dim com As SqlCommand = New SqlCommand(sql, ConA)
            Dim dt As DataTable = New DataTable(name)
            Dim da As SqlDataAdapter = New SqlDataAdapter(com)
            da.Fill(dt)
            com.Clone()
            com.Dispose()
            da.Dispose()
            Return dt

        End Try

    End Function

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click

        If Checksyntax() And dgvFrindBath.Rows.Count > 0 Then
            _id = Convert.ToDecimal(dgvFrindBath.Rows(dgvFrindBath.CurrentCell.RowIndex).Cells("Batch Number").Value.ToString()).ToString("###0")
            Me.Close()

        Else

            MessageBox.Show("Please select at least one.", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.txtFilter.Focus()

        End If
    End Sub

    Private Sub txtFilter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFilter.KeyPress

        If cboFindBy.SelectedIndex <> 1 Then
            Return
        End If

        If (e.KeyChar.Equals(Convert.ToChar(Keys.Back))) Then

        Else

            Dim isNumber As Integer = 0
            If Not Integer.TryParse(e.KeyChar.ToString(), isNumber) Then 'TryParseแปลงตัวแปลต่างๆให้เป็นตัวเลข int 

                e.Handled = True

            Else

                e.Handled = False
                _id = txtFilter.Text
            End If

        End If

    End Sub

    Private Sub txtFilter_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilter.TextChanged
        If chkAuto.Checked Then
            ShowData()
        End If
    End Sub



    Private Sub cboFindBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFindBy.SelectedIndexChanged
        If cboFindBy.SelectedIndex = 1 Or cboFindBy.SelectedIndex = 3 Then

            cboStatement.Items.Clear()
            cboStatement.Items.AddRange(New Object() {"=", ">", "<", ">=", "<=", "!="})

        Else

            cboStatement.Items.Clear()
            cboStatement.Items.AddRange(New Object() {"Starts with", "Countains"})

        End If

        If cboFindBy.SelectedIndex = 1 Then

            dtpDate.Visible = False
            txtFilter.Visible = True
            cboStatement.Visible = True
            cboStatement.SelectedIndex = 0
            label2.Visible = True
            txtFilter.Text = "0"

        ElseIf cboFindBy.SelectedIndex = 2 Then

            dtpDate.Visible = False
            txtFilter.Visible = True
            txtFilter.Text = ""
            label2.Visible = True
            cboStatement.Visible = True
            cboStatement.SelectedIndex = 0

        ElseIf cboFindBy.SelectedIndex = 3 Then

            dtpDate.Visible = True
            txtFilter.Visible = False
            label2.Visible = True
            cboStatement.Visible = True
            cboStatement.SelectedIndex = 0

        Else

            dtpDate.Visible = False
            txtFilter.Visible = False
            label2.Visible = False
            cboStatement.Visible = False
            If (chkAuto.Checked) Then
                ShowData()
            End If
        End If

    End Sub

    Private Sub dtpDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDate.ValueChanged
        If (chkAuto.Checked) Then
            ShowData()
        End If

    End Sub

    Private Sub cboStatement_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStatement.SelectedIndexChanged
        If (chkAuto.Checked) Then
            ShowData()
        End If
    End Sub

    Private Sub chkAuto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAuto.CheckedChanged

        btnFindNow.Enabled = Not chkAuto.Checked
        ShowData()

    End Sub

    Private Sub btnFindNow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindNow.Click
        ShowData()
    End Sub

    Private Sub chkShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShow.CheckedChanged
        If chkShow.Checked = False Then

            BATCHSTAT = "and (BATCHSTAT <> 3 and BATCHSTAT <> 4)"
        Else

            BATCHSTAT = ""
        End If

        If chkAuto.Checked Then

            ShowData()
        End If
    End Sub

    Private Sub dgvFrindBath_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFrindBath.CellDoubleClick
        If e.RowIndex = -1 Then
            Return

            b_id = Convert.ToDecimal(dgvFrindBath.Rows(dgvFrindBath.CurrentCell.RowIndex).Cells("Batch Number").Value.ToString()).ToString("###0")

            txtFilter.Text = _id
            Me.Close()
        End If
    End Sub


End Class