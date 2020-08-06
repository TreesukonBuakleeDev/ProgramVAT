Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class PriviewReport
    Dim myLogin As CrystalDecisions.Shared.TableLogOnInfo
    Dim myTable As CrystalDecisions.CrystalReports.Engine.Table
    Public Server, DataBase, Userid, Password, sSelection, sReplace As String
    Public ODBCVAT(3) As String
    Public ODBCACC(3) As String
    Public Rate, Loc, txtFrom, txtTo, Company, TAXNO, Vtype, VTType As String
    Public TypeReport As Integer
    Private Sub PriviewReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim rpt As New ReportDocument
            rpt = New ReportDocument()

            Dim strFullReport As String = ""
            'Dim strReportName As String = ReportBudget
            Dim STR As String = Application.StartupPath

            strFullReport = ReportChk
            rpt.Load(strFullReport, OpenReportMethod.OpenReportByDefault)
            If ComDatabase = "MSSQL" Then
                'Call LoginDatabase(rpt)
                Call LoginDatabaseODBC(rpt)
            Else
                Call LoginDatabasePSW(rpt)
            End If
            Dim RecordSelectionFormula As String = rpt.RecordSelectionFormula
            If sSelection.Trim <> "" And sReplace.Trim <> "" Then
                RecordSelectionFormula = RecordSelectionFormula.Replace(sSelection, sReplace)
                rpt.RecordSelectionFormula = RecordSelectionFormula
            End If
            Dim sqlstr As String = ""
            If S_AP Or S_AR Or S_GL Or S_CB Or s_OE Or S_PO Then
                sqlstr &= " AND {FMSVAT.SOURCE} in [""US"","
                If S_AP Then sqlstr &= """AP"","
                If S_AR Then sqlstr &= """AR"","
                If S_GL Then sqlstr &= """GL"","
                If S_CB Then sqlstr &= """CB"","
                If S_PO Then sqlstr &= """PO"","
                If s_OE Then sqlstr &= """OE"","
                sqlstr = sqlstr.Substring(0, sqlstr.Length - 1) & "]"
            Else
                sqlstr &= " AND {FMSVAT.SOURCE} in [""US""]"
            End If
            rpt.RecordSelectionFormula &= sqlstr

            rpt.SetParameterValue("Rate", Rate) '
            rpt.SetParameterValue("Location", Loc) '
            rpt.SetParameterValue("DateFrom", txtFrom) '
            rpt.SetParameterValue("DateTo", txtTo) '
            rpt.SetParameterValue("Company", Company) '
            rpt.SetParameterValue("TaxNo", TAXNO) '
            rpt.SetParameterValue("Vtype", Vtype) '
            rpt.SetParameterValue("VTType", VTType) '
            CrystalReportViewer1.ReportSource = rpt
            Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            Call writefile(Application.StartupPath.ToString() & "\\", "log" & DateTime.Today.Year & DateTime.Today.Month & DateTime.Today.Day & ".txt", System.DateTime.Now.ToString() & "Report " & ex.Source & ":" & ex.Message)
        End Try
    End Sub

    Private Sub LoginDatabaseODBC(ByVal Obj As Object)
        ODBCACC = ComACCPAC.Split(";")
        ODBCVAT = ComVAT.Split(";")
        Dim abc As ReportDocument = Obj
        Try
            If Password Is Nothing = True Then
                Password = ""
            End If
            For Each myTable In Obj.Database.Tables
                If myTable.Name = "FMSVAT" Or myTable.Name = "FMSVLOC" Then
                    myLogin = myTable.LogOnInfo
                    myLogin.ConnectionInfo.ServerName = ODBCVAT(0)
                    myLogin.ConnectionInfo.DatabaseName = ODBCVAT(1)
                    myLogin.ConnectionInfo.UserID = ODBCVAT(2)
                    myLogin.ConnectionInfo.Password = ODBCVAT(3)
                    myTable.ApplyLogOnInfo(myLogin)
                Else
                    myLogin = myTable.LogOnInfo
                    myLogin.ConnectionInfo.ServerName = ODBCACC(0)
                    myLogin.ConnectionInfo.DatabaseName = ODBCACC(1)
                    myLogin.ConnectionInfo.UserID = ODBCACC(2)
                    myLogin.ConnectionInfo.Password = ODBCACC(3)
                    myTable.ApplyLogOnInfo(myLogin)
                End If
            Next
            Dim s As ReportDocument = Obj
            Dim i As Integer
            For i = 0 To s.Subreports.Count - 1
                For Each myTable In s.Subreports(i).Database.Tables
                    If myTable.Name = "FMSVAT" Or myTable.Name = "FMSVLOC" Then
                        myLogin = myTable.LogOnInfo
                        myLogin.ConnectionInfo.ServerName = ODBCVAT(0)
                        myLogin.ConnectionInfo.DatabaseName = ODBCVAT(1)
                        myLogin.ConnectionInfo.UserID = ODBCVAT(2)
                        myLogin.ConnectionInfo.Password = ODBCVAT(3)
                        myTable.ApplyLogOnInfo(myLogin)
                    Else
                        myLogin = myTable.LogOnInfo
                        myLogin.ConnectionInfo.ServerName = ODBCACC(0)
                        myLogin.ConnectionInfo.DatabaseName = ODBCACC(1)
                        myLogin.ConnectionInfo.UserID = ODBCACC(2)
                        myLogin.ConnectionInfo.Password = ODBCACC(3)
                        myTable.ApplyLogOnInfo(myLogin)
                    End If
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoginDatabase(ByVal obj As Object)
        Try
            If Password Is Nothing = True Then
                Password = ""
            End If
            For Each myTable In obj.Database.Tables
                myLogin = myTable.LogOnInfo
                myLogin.ConnectionInfo.ServerName = Server
                myLogin.ConnectionInfo.DatabaseName = DataBase
                myLogin.ConnectionInfo.UserID = Userid
                myLogin.ConnectionInfo.Password = Password.Trim
                myTable.ApplyLogOnInfo(myLogin)
            Next
            Dim s As ReportDocument = obj
            Dim i As Integer
            For i = 0 To s.Subreports.Count - 1
                For Each myTable In s.Subreports(i).Database.Tables
                    myLogin = myTable.LogOnInfo
                    myLogin.ConnectionInfo.ServerName = Server
                    myLogin.ConnectionInfo.DatabaseName = DataBase
                    myLogin.ConnectionInfo.UserID = Userid
                    myLogin.ConnectionInfo.Password = Password.Trim
                    myTable.ApplyLogOnInfo(myLogin)
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub LoginDatabasePSW(ByVal obj As Object)
        Try
            If Password Is Nothing = True Then
                Password = ""
            End If
            For Each myTable In obj.Database.Tables
                myLogin = myTable.LogOnInfo
                myLogin.ConnectionInfo.ServerName = Server
                'myLogin.ConnectionInfo.DatabaseName = DataBase
                'myLogin.ConnectionInfo.UserID = Userid
                'myLogin.ConnectionInfo.Password = Password.Trim
                myTable.ApplyLogOnInfo(myLogin)
            Next
            Dim s As ReportDocument = obj
            Dim i As Integer
            For i = 0 To s.Subreports.Count - 1
                For Each myTable In s.Subreports(i).Database.Tables
                    myLogin = myTable.LogOnInfo
                    myLogin.ConnectionInfo.ServerName = Server
                    'myLogin.ConnectionInfo.DatabaseName = DataBase
                    'myLogin.ConnectionInfo.UserID = Userid
                    'myLogin.ConnectionInfo.Password = Password.Trim
                    myTable.ApplyLogOnInfo(myLogin)
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function writefile(ByVal path As String, ByVal filename As String, ByVal message As String) As Boolean
        Try
            Dim sr As StreamWriter
            If (File.Exists(path & filename)) Then
                sr = File.AppendText(path & filename)
                sr.WriteLine(message)
                sr.Close()
                sr.Dispose()
            Else
                sr = File.CreateText(path & filename)
                sr.WriteLine(message)
                sr.Dispose()
            End If
            Return True
        Catch ex As Exception
            Console.WriteLine(CStr(ex.Data.ToString()) & ": " & ex.Source & ": " & ex.Message)
            Return False
        End Try
    End Function
End Class
