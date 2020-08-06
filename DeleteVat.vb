Imports ADOR
Imports System.Globalization

Public Class DeleteVat
    Dim cnVAT As New ADODB.Connection
    Dim rs As New BaseClass.BaseODBCIO
    Dim sqlstr As String = ""

    Private Sub DeleteVat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        sqlstr = "select substring(TXDATE,0,5) as vYear from FMSVATINSERT  group by substring(TXDATE,0,5) order by substring(TXDATE,0,5)"
        rs.Open(sqlstr, cnVAT)
        For i As Integer = 0 To rs.options.QueryDT.Rows.Count - 1
            With rs.options.QueryDT.Rows(i)
                CMYear.Items.Add(.Item(0).ToString)
            End With
        Next
        If CMYear.Items.Count > 0 Then
            CMYear.SelectedIndex = CMYear.Items.Count - 1
            Button1_MouseClick(Nothing, Nothing)
        End If
    End Sub

    Private Sub Button1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseClick
        lstVu.Rows.Clear()
        cnVAT = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        sqlstr = "select substring(TXDATE,0,5) as vYear, substring(TXDATE,5,2)as vMonth ,COUNT(*)as NumEntry    from FMSVATINSERT where substring(TXDATE,0,5)='" & CMYear.Text & "'   group by substring(TXDATE,5,2),substring(TXDATE,0,5) order by substring(TXDATE,0,5) , substring(TXDATE,5,2)"
        rs.Open(sqlstr, cnVAT)
        For i As Integer = 0 To rs.options.QueryDT.Rows.Count - 1
            With rs.options.QueryDT.Rows(i)
                Dim myDate As Date
                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                fi.ShortDatePattern = "yyyyMMdd"
                myDate = DateTime.ParseExact(CStr(.Item(0) & .Item(1) & "01").Trim, "yyyyMMdd", fi)
                lstVu.Rows.Add(New String() {(i + 1).ToString(), .Item("vMonth").ToString & "-" & myDate.ToString("MMMM"), .Item("NumEntry"), "Delete", .Item("vMonth").ToString})
            End With
        Next
    End Sub

    Private Sub lstVu_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles lstVu.CellContentClick
        If e.ColumnIndex = 3 And e.RowIndex > -1 Then
            cnVAT = New ADODB.Connection
            rs = New BaseClass.BaseODBCIO
            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            If MessageBox.Show(" Delete Vat " & lstVu.Rows(e.RowIndex).Cells(2).Value & " Entry ,Please Confirm !", "Confirm Delete", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
                sqlstr = "delete FMSVATINSERT where substring(TXDATE,0,7) ='" & CMYear.Text & lstVu.Rows(e.RowIndex).Cells(4).Value.ToString & "'"
                Dim record As New Object
                cnVAT.Execute(sqlstr, record)
                If record > 0 Then
                    MessageBox.Show("Delete " & CMYear.Text & lstVu.Rows(e.RowIndex).Cells(4).Value.ToString & " Sucess " & CInt(record).ToString("#,##0") & " Record !", "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                Button1_MouseClick(Nothing, Nothing)
                CType(Me.Owner, frmMain).btnRefreh_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub Button2_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseClick
        Me.Close()
    End Sub
End Class