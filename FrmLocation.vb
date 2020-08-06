Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Data.SqlClient

Public Class FrmLocation
    Private constr As String
    Private locDS As BaseClass.BaseODBCIO
    Private position As Integer = 0
    Private AccDS As BaseClass.BaseODBCIO
    Private cnVAT As ADODB.Connection
    Private Count As Integer = 0

    Sub locationinfo()
        Try
            Dim sqlstr As String
            'Dim i As Double
            Dim F As Boolean = True
            sqlstr = " SELECT LOCID, LOCNAME, LOCADD, LOCPREFIX, LOCDESC "
            sqlstr &= " FROM   FMSVLOC "
            sqlstr &= " ORDER BY LOCID "
            cnVAT = New ADODB.Connection
            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            cnVAT.CommandTimeout = 3600
            locDS = New BaseClass.BaseODBCIO
            locDS.Open(sqlstr, cnVAT)
            position = 0
            Do While (locDS.options.Rs.EOF = False And F = True)
                position += 1
                txtAddress.Text = IIf(IsDBNull(locDS.options.Rs.Collect(2)), "", (locDS.options.Rs.Collect(2))).Trim
                txtComment.Text = IIf(IsDBNull(locDS.options.Rs.Collect(4)), "", (locDS.options.Rs.Collect(4))).Trim
                txtLocation.Text = IIf(IsDBNull(locDS.options.Rs.Collect(0)), "", (locDS.options.Rs.Collect(0))).Trim
                txtName.Text = IIf(IsDBNull(locDS.options.Rs.Collect(1)), "", (locDS.options.Rs.Collect(1))).Trim
                txtPrefix.Text = IIf(IsDBNull(locDS.options.Rs.Collect(3)), "", (locDS.options.Rs.Collect(3))).Trim
                'locDS.MoveNext()
                Count = locDS.options.Rs.RecordCount
                F = False
            Loop
            If locDS.options.Rs.RecordCount = 0 Then
                position = 0
                txtAddress.Text = ""
                txtComment.Text = ""
                txtLocation.Text = ""
                txtName.Text = ""
                txtPrefix.Text = ""
                'locDS.MoveNext()
                Count = 0
            End If
            'cnVAT.Close()
            'If locDS.Tables.Count > 0 Then
            '    If locDS.Tables(0).Rows.Count > 0 Then
            '        position = 0

            '        txtAddress.Text = locDS.Tables(0).Rows(position)("LOCADD").ToString().Trim().Trim()
            '        txtComment.Text = locDS.Tables(0).Rows(position)("LOCDESC").ToString().Trim().Trim()
            '        txtLocation.Text = locDS.Tables(0).Rows(position)("LOCID").ToString().Trim().Trim()
            '        txtName.Text = locDS.Tables(0).Rows(position)("LOCNAME").ToString().Trim().Trim()
            '        txtPrefix.Text = locDS.Tables(0).Rows(position)("LOCPREFIX").ToString().Trim().Trim()
            '    End If
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub cregrid()
        Try
            Dim sqlstr As String
            'Dim i As Integer
            cnVAT = New ADODB.Connection
            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            cnVAT.CommandTimeout = 3600
            AccDS = New BaseClass.BaseODBCIO

            'Dim con As SqlConnection = New SqlConnection(ConnDB)
            sqlstr = " SELECT RTRIM(VTYPE) AS TYPE ,RTRIM(ACCTVAT) AS Account"
            sqlstr &= " FROM  FMSVLACC "
            sqlstr &= " WHERE (LOCID = '" & txtLocation.Text.Trim() & "') "
            sqlstr &= " ORDER BY VTYPE "
            'con.Open()
            'AccDS.Tables.Clear()
            'AccDS.Clear()
            'Dim adap As SqlDataAdapter = New SqlDataAdapter(sqlstr, con)
            'adap.Fill(AccDS)
            AccDS.Open(sqlstr, cnVAT)

            dataGridView1.Rows.Clear()
            dataGridView1.Rows.Add(3)
            dataGridView1.Rows(0).Cells(0).Value = "PURCHASE"
            dataGridView1.Rows(1).Cells(0).Value = "SALE"
            dataGridView1.Rows(2).Cells(0).Value = "EXPENSE"
            If AccDS.options.Rs.EOF = True Then
                btnUpdate.Text = "Save"
                btnUpdate.Image = Global.ProgramVAT.My.Resources.Resources.Save_icon_1
            Else
                btnUpdate.Text = "Update"
                btnUpdate.Image = Global.ProgramVAT.My.Resources.Resources.edit_validated_icon

            End If

            Do While AccDS.options.Rs.EOF = False
                If (IIf(IsDBNull(AccDS.options.Rs.Collect(0)), "", (AccDS.options.Rs.Collect(0))) = "1") Then 'PURCHASE
                    dataGridView1.Rows(0).Cells(1).Value = IIf(IsDBNull(AccDS.options.Rs.Collect(1)), "", (AccDS.options.Rs.Collect(1))).Trim
                ElseIf (IIf(IsDBNull(AccDS.options.Rs.Collect(0)), "", (AccDS.options.Rs.Collect(0))) = "2") Then 'SALE
                    dataGridView1.Rows(1).Cells(1).Value = IIf(IsDBNull(AccDS.options.Rs.Collect(1)), "", (AccDS.options.Rs.Collect(1))).Trim
                ElseIf (IIf(IsDBNull(AccDS.options.Rs.Collect(0)), "", (AccDS.options.Rs.Collect(0))) = "3") Then 'EXPENSE
                    dataGridView1.Rows(2).Cells(1).Value = IIf(IsDBNull(AccDS.options.Rs.Collect(1)), "", (AccDS.options.Rs.Collect(1))).Trim
                End If
                AccDS.options.Rs.MoveNext()
            Loop
           
            dataGridView1.ClearPictureEdit(0, True)
            If AccDS.options.Rs.RecordCount = 0 Then
                dataGridView1.Rows(0).Cells(1).Value = ""
                dataGridView1.Rows(1).Cells(1).Value = ""
                dataGridView1.Rows(2).Cells(1).Value = ""
            End If
            'For i = 0 To AccDS.Tables(0).Rows.Count - 1
            '    If (AccDS.Tables(0).Rows(i)("Type").ToString().Trim().Trim() = "1") Then 'PURCHASE

            '        dataGridView1.Rows(0).Cells(1).Value = AccDS.Tables(0).Rows(i)("Account").ToString().Trim()

            '    ElseIf (AccDS.Tables(0).Rows(i)("Type").ToString().Trim().Trim() = "2") Then 'SALE

            '        dataGridView1.Rows(1).Cells(1).Value = AccDS.Tables(0).Rows(i)("Account").ToString().Trim()

            '    ElseIf (AccDS.Tables(0).Rows(i)("Type").ToString().Trim().Trim() = "3") Then 'EXPENSE

            '        dataGridView1.Rows(2).Cells(1).Value = AccDS.Tables(0).Rows(i)("Account").ToString().Trim()
            '    End If
            'Next
            'con.Close()
            'con.Dispose()
            'cnVAT.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmLocation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        locationinfo()
        cregrid()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        txtAddress.Text = ""
        txtComment.Text = ""
        txtLocation.Text = ""
        txtName.Text = ""
        txtPrefix.Text = ""
        cregrid()

    End Sub

    Private Sub cmbP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbP.Click
        If (Count > 0) And (position <> 1) Then
            position -= 1
            locDS.options.Rs.MovePrevious()
            If Not (locDS.options.Rs.EOF) Then
                txtAddress.Text = IIf(IsDBNull(locDS.options.Rs.Collect(2)), "", (locDS.options.Rs.Collect(2))).Trim
                txtComment.Text = IIf(IsDBNull(locDS.options.Rs.Collect(4)), "", (locDS.options.Rs.Collect(4))).Trim
                txtLocation.Text = IIf(IsDBNull(locDS.options.Rs.Collect(0)), "", (locDS.options.Rs.Collect(0))).Trim
                txtName.Text = IIf(IsDBNull(locDS.options.Rs.Collect(1)), "", (locDS.options.Rs.Collect(1))).Trim
                txtPrefix.Text = IIf(IsDBNull(locDS.options.Rs.Collect(3)), "", (locDS.options.Rs.Collect(3))).Trim
                'position -= 1
                'txtAddress.Text = locDS.Tables(0).Rows(position)("LOCADD").ToString().Trim().Trim()
                'txtComment.Text = locDS.Tables(0).Rows(position)("LOCDESC").ToString().Trim().Trim()
                'txtLocation.Text = locDS.Tables(0).Rows(position)("LOCID").ToString().Trim().Trim()
                'txtName.Text = locDS.Tables(0).Rows(position)("LOCNAME").ToString().Trim().Trim()
                'txtPrefix.Text = locDS.Tables(0).Rows(position)("LOCPREFIX").ToString().Trim().Trim()
                cregrid()
            End If
        End If
    End Sub

    Private Sub cmbN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbN.Click
        If (Count > 0) And (position <> Count) Then
            position += 1
            locDS.options.Rs.MoveNext()
            If Not (locDS.options.Rs.EOF) Then
                txtAddress.Text = IIf(IsDBNull(locDS.options.Rs.Collect(2)), "", (locDS.options.Rs.Collect(2))).Trim
                txtComment.Text = IIf(IsDBNull(locDS.options.Rs.Collect(4)), "", (locDS.options.Rs.Collect(4))).Trim
                txtLocation.Text = IIf(IsDBNull(locDS.options.Rs.Collect(0)), "", (locDS.options.Rs.Collect(0))).Trim
                txtName.Text = IIf(IsDBNull(locDS.options.Rs.Collect(1)), "", (locDS.options.Rs.Collect(1))).Trim
                txtPrefix.Text = IIf(IsDBNull(locDS.options.Rs.Collect(3)), "", (locDS.options.Rs.Collect(3))).Trim
                'txtAddress.Text = locDS.Tables(0).Rows(position)("LOCADD").ToString().Trim().Trim()
                'txtComment.Text = locDS.Tables(0).Rows(position)("LOCDESC").ToString().Trim().Trim()
                'txtLocation.Text = locDS.Tables(0).Rows(position)("LOCID").ToString().Trim().Trim()
                'txtName.Text = locDS.Tables(0).Rows(position)("LOCNAME").ToString().Trim().Trim()
                'txtPrefix.Text = locDS.Tables(0).Rows(position)("LOCPREFIX").ToString().Trim().Trim()
                cregrid()
            End If
        End If
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            Dim sqlstr As String
            Dim i As Integer
            Dim DS As BaseClass.BaseODBCIO
            cnVAT = New ADODB.Connection
            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            cnVAT.CommandTimeout = 3600

            'Dim con As SqlConnection = New SqlConnection(ConnDB)
            'con.Open()
            'Dim com As SqlCommand = New SqlCommand()
            'com.Connection = con

            Dim rs As New BaseClass.BaseODBCIO
            sqlstr = "select * from FMSVLACC where ACCTVAT ='" & Trim(dataGridView1.Rows(0).Cells(1).Value) & "' and ltrim(rtrim(LOCID))<>'" & txtLocation.Text.Trim & "'"
            rs.Open(sqlstr, cnVAT)
            If rs.options.Rs.RecordCount > 0 Then
                dataGridView1.CurrentCell = dataGridView1.Rows(0).Cells(1)
                Dim wid As Integer = 0
                Dim he As Integer = (dataGridView1.CurrentCell.Size.Height / 2) + dataGridView1.ColumnHeadersHeight + dataGridView1.Location.Y - 40
                For Each co As DataGridViewColumn In dataGridView1.Columns
                    wid += co.Width
                Next
                ShowToolTip(Me, ToolTipIcon.Warning, "Account Code is Duplicate", wid, he)
                dataGridView1.BeginEdit(True)
                Exit Sub
            End If
            rs = New BaseClass.BaseODBCIO
            sqlstr = "select * from FMSVLACC where ACCTVAT ='" & Trim(dataGridView1.Rows(1).Cells(1).Value) & "' and ltrim(rtrim(LOCID)) <>'" & txtLocation.Text.Trim & "'"
            rs.Open(sqlstr, cnVAT)
            If rs.options.Rs.RecordCount > 0 Then
                dataGridView1.CurrentCell = dataGridView1.Rows(1).Cells(1)
                Dim wid As Integer = 0
                Dim he As Integer = (dataGridView1.CurrentCell.Size.Height / 2) + dataGridView1.ColumnHeadersHeight + dataGridView1.Location.Y - 40
                For Each co As DataGridViewColumn In dataGridView1.Columns
                    wid += co.Width
                Next
                ShowToolTip(Me, ToolTipIcon.Warning, "Account Code is Duplicate", wid, he)
                dataGridView1.BeginEdit(True)
                Exit Sub
            End If

            sqlstr = " SELECT     LOCID  FROM         FMSVLOC "
            sqlstr &= " WHERE     (LOCID = '" + txtLocation.Text.Trim() & "')"
            'Dim ds As New DataSet
            'Dim adap As SqlDataAdapter = New SqlDataAdapter(sqlstr, con)
            'adap.Fill(ds)
            DS = New BaseClass.BaseODBCIO
            DS.Open(sqlstr, cnVAT)
            If DS.options.Rs.RecordCount = 0 Then
                ' insert
                sqlstr = " INSERT INTO FMSVLOC "
                sqlstr &= " (LOCID, LOCNAME, LOCADD, LOCPREFIX, LOCDESC) "
                sqlstr &= "  VALUES     ('" & txtLocation.Text.Trim().Replace("'", "''")
                sqlstr &= "    ','" & txtName.Text.Trim().Replace("'", "''")
                sqlstr &= "     ','" & txtAddress.Text.Trim().Replace("'", "''")
                sqlstr &= "     ','" & txtPrefix.Text.Trim().Replace("'", "''")
                sqlstr &= "     ','" & txtComment.Text.Trim().Replace("'", "''") & "') "
                'com.CommandText = sqlstr
                'com.ExecuteNonQuery()
            Else
                ' update
                sqlstr = " UPDATE    FMSVLOC "
                sqlstr &= " SET              LOCNAME ='" & txtName.Text.Trim().Replace("'", "''")
                sqlstr &= " ', LOCADD ='" & txtAddress.Text.Trim().Replace("'", "''")
                sqlstr &= " ', LOCPREFIX ='" & txtPrefix.Text.Trim().Replace("'", "''")
                sqlstr &= " ', LOCDESC = '" & txtComment.Text.Trim().Replace("'", "''") & "'"
                sqlstr &= "  WHERE     (LOCID = '" & txtLocation.Text.Trim().Replace("'", "''") & "')"
                'com.CommandText = sqlstr
                'com.ExecuteNonQuery()
            End If
            cnVAT.Execute(sqlstr)

            'add Vat Account
            sqlstr = "DELETE FROM FMSVLACC WHERE (LOCID = '" & txtLocation.Text.Trim().Replace("'", "''") & "')"
            cnVAT.Execute(sqlstr)


            For i = 0 To dataGridView1.Rows.Count - 1
                If Not IsDBNull(dataGridView1.Rows(i).Cells(1).Value) Then
                    sqlstr = ""
                    sqlstr = "INSERT INTO FMSVLACC(ACCTVAT, LOCID, VTYPE)VALUES ('" & Trim(dataGridView1.Rows(i).Cells(1).Value) & "','" & Trim(txtLocation.Text) & "','" & CStr(i + 1) + "')"
                    If dataGridView1.Rows(i).Cells(1).Value.ToString().Trim() <> "" Then
                        cnVAT.Execute(sqlstr)
                    End If
                End If
            Next
            locationinfo()
            cregrid()
            'con.Close()
            'con.Dispose()
            'cnVAT.Close()

            MessageBox.Show("Save Success", "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Text = "Update"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            'Dim con As SqlConnection = New SqlConnection(ConnDB)
            'Dim com As New SqlCommand
            Dim sqlstr As String
            cnVAT = New ADODB.Connection
            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            cnVAT.CommandTimeout = 3600
            If MsgBox("Delete selected record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.ProductName) <> MsgBoxResult.Yes Then
                Exit Sub
            End If

            sqlstr = " DELETE FMSVLOC WHERE (LOCID = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            sqlstr = " DELETE FMSVLACC WHERE (LOCID = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            sqlstr = " DELETE FMSTAXGL WHERE (AUTHORITY = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            sqlstr = " DELETE FMSVAT WHERE (LOCID = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            sqlstr = " DELETE FMSVATINSERT WHERE (LOCID = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            sqlstr = " DELETE FMSVATTEMP WHERE (LOCID = '" & txtLocation.Text.Trim() & "');"
            Try
                cnVAT.Execute(sqlstr)
            Catch ex As Exception
            End Try
            'com.ExecuteNonQuery()
            'con.Close()
            'con.Dispose()
            MessageBox.Show("Delete Success", "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
            locationinfo()
            cregrid()
            'cnVAT.Close()
            CType(Me.Owner, frmMain).btnRefreh_Click(Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    'Private Sub dataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellClick
    '    selectgridrows = e.RowIndex
    '    If (e.RowIndex > -1) Then
    '        dataGridView1.Rows(e.RowIndex).Selected = True
    '        selectgridrows = e.RowIndex
    '    End If
    'End Sub

    'Private Sub dataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dataGridView1.CellContentClick
    '    If e.RowIndex <> -1 Then
    '        dataGridView1.Rows(e.RowIndex).Selected = True
    '    End If
    'End Sub

    Private selectgridrows As Integer = -1

End Class