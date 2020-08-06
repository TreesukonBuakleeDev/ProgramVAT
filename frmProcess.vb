Option Strict Off
Option Explicit On
Friend Class frmProcess
	Inherits System.Windows.Forms.Form
    Public Process As Integer
    Public WithEvents Timer1 As New Timer
    Public swTime As Integer

    Private Sub frmProcess_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Timer1.Stop()
        Timer1.Enabled = False
    End Sub

	Private Sub frmProcess_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Select Case Process
			Case 0
				lblProcess.Text = "Setting Process"
			Case 1
				lblProcess.Text = "On Process"
		End Select
		timProcess.Enabled = True
        Me.Text = "Process"
        swTime = 0
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Timer1.Start()
    End Sub
    Public Sub ProcessShowInsert(ByVal Showstring As String)

        lblInsert.Text = Showstring.Trim
        lblInsert.Refresh()

    End Sub
	
    Private Sub ProcessShowTxt(ByRef ProcessValue As Integer, ByRef ShowText As String)

        Try
            CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).ShowInTaskbar = True
            CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Maximum = pgbProcess.Maximum
            CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Minimum = pgbProcess.Minimum
            CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Value = ProcessValue

            Dim cn As New ADODB.Connection
            Dim rs As New BaseClass.BaseODBCIO
            cn.Open(ConnVAT)
            rs.Open("select object_id  from sys.tables where name ='fmsvset'", cn)
            If rs.options.QueryDT.Rows.Count > 0 Then
                rs.Open("select * from fmsvset where INGETVAT =1", cn)
                If rs.options.QueryDT.Rows.Count > 0 Then
                    If rs.options.QueryDT.Rows(0).Item("USERGET").ToString = UserLogin.Trim Then
                        cn.Execute("Update FMSVSET SET STATUS='" & ProcessValue & "'")
                    End If
                End If
            End If
            rs = Nothing
            cn.Close()
        Catch ex As Exception
        End Try
        pgbProcess.Value = ProcessValue
        lblProcess.Text = ShowText & vbCrLf & "( " & ProcessValue & " / " & pgbProcess.Maximum & " )"
        lblProcess.Refresh()

    End Sub


    Private Sub timProcess_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles timProcess.Tick

        timProcess.Enabled = False
        pgbProcess.Minimum = 0
        pgbProcess.Value = 1
        Select Case Process
            Case 0 'OPEN
                'Mouse(MON)
                System.Windows.Forms.Cursor.Current = Cursors.AppStarting

                pgbProcess.Maximum = 6
                '************************
                'ResetTable
                '************************
                Call ProcessShowTxt((pgbProcess.Value), "Initial Setting")
                Call SetInitial() 'เช็ค Database
                Call LoadDBScript() 'กรณีเพิ่มฟิลด์
                Call PrepareVatSet() 'Get Company Name
                Call PrepareVatUpdateNew()
                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Type")
                Call PrepareType()
                Call PrepareRUNNING()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Pre-Location")
                Call PrepareLocation()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Pre-Account")
                Call PrepareAccount()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat Data")
                Call PrepareVatTable()
                Call PrepareUserPassTable()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Loading Data")
                Try
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).ShowInTaskbar = False
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Maximum = 0
                Catch ex As Exception
                End Try
                Me.Close()
                'Mouse(MOFF)
                System.Windows.Forms.Cursor.Current = Cursors.Default

                WriteLog("<Process = 0 (Open)> Complete")

            Case 1
                '************************
                'Process GET VAT
                '************************
                'Mouse(MON)
                System.Windows.Forms.Cursor.Current = Cursors.AppStarting

                'StartTime = TimeOfDay
                pgbProcess.Maximum = 11

                Call ProcessShowTxt((pgbProcess.Value), "Initial Setting")
                Call SetInitial()
                Call PrepareRUNNING()
                Call ProcessSetDate()
                Call PrepareTempTable()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Clear Old Data")
                Call DropVatTable()
                Call PrepareVatTable()
                Call LoadDBScript()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Tax")
                Call PrepareTax()
                Call ProcessTax()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Tax")
                Call PrepareTaxGL()
                Call ProcessTaxGL()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat From Module GL")
                Call ProcessGL()

                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat From Module AP-PO")
                Call PrepareTempTableAP()
                Call PrepareTempTable()

                If FindApp("PO") = True Then
                    Call ProcessAPPJHTEMP()
                    Call ProcessAPOBLTEMP()
                    'หาก Use_INVFROMPO = TRUE ให้ดึงข้อมูลมาทีละ Module โดยเริ่มจาก AP,PO ตามลำดับ
                    'หากเป็น False ให้ดึงรวมกันเลยจาก Sub ProcessAP_PO
                    If Use_INVFROMPO = True Then
                        Call ProcessAP()
                        Call PrepareTempTable()
                        Call ProcessPO()
                    Else
                        Call ProcessAP_PO()
                    End If

                Else
                    If FindApp("AP") = True Then
                        Call ProcessAPPJHTEMP()
                        Call ProcessAPOBLTEMP()
                        Call ProcessAP()
                    End If
                End If

                '/**********************************************************/
                '    Start Find Vat Case Input AP Payment Vat
                '/**********************************************************/----

                Call PrepareTempTable()
                If FindApp("AP") = True Then
                    Call ProcessAPPay()
                End If
                '/**********************************************************/
                '     End  Find Vat Case Input AP Payment Vat
                '/**********************************************************/

                '/*********************************************************************/
                '     Start Find Vat Case Input AP Payment Misc Payment
                '/*********************************************************************/
                Call PrepareTempTable()
                If FindApp("AP") = True Then
                    Call ProcessAPMisc()
                End If
                '/*********************************************************************/
                '     End Find Vat Case Input AP Payment Misc Payment
                '/*********************************************************************/

                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat From Module AR-OE")
                Call PrepareTempTable()
                Call PrepareTempTableAR()
                Dim a() As String = ("").Split(",")

                If FindApp("OE") = True Then
                    Call ProcessARPJHTEMP()
                    Call ProcessAROBLTEMP()
                    Call ProcessOE()
                Else
                    If FindApp("AR") = True Then
                        Call ProcessARPJHTEMP()
                        Call ProcessAROBLTEMP()
                        Call ProcessAR()
                    End If
                End If


                '/*********************************************************************/
                '     Start Find Vat Case Input AR Receipt Vat
                '/*********************************************************************/
                Call PrepareTempTable()
                If FindApp("AR") = True Then
                    Call ProcessARRec()
                End If
                '/*********************************************************************/
                '     End Find Vat Case Input AR Receipt Vat
                '/*********************************************************************/

                '/*********************************************************************/
                '     Start Find Vat Case Input AP Receipt Misc Recept
                '/*********************************************************************/
                Call PrepareTempTable()
                If FindApp("AR") = True Then
                    Call ProcessARMisc()
                End If
                '/*********************************************************************/
                '     End Find Vat Case Input AP Receipt Misc Recept
                '/*********************************************************************/


                'Call ProcessShowTxt(pgbProcess.Value + 1, "NoneVat From Module AR-OE")
                'Call PrepareTempTable()
                'If FindApp("OE") = True Then
                '    Call ProcessOENOVAT()
                'Else
                '    If FindApp("AR") = True Then
                '        Call ProcessARNOVAT()
                '    End If
                'End If


                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat From Module CB")
                If FindApp("CB") = True Then
                    Call PrepareCB()
                    Call ProcessCB()
                End If

                Call ProcessShowTxt(pgbProcess.Value + 1, "Vat From Module BK")
                Call PrepareTempTable()
                Call ProcessBK()
                Call ProcessShowTxt(pgbProcess.Value + 1, "Complete")

                'Mouse(MOFF)
                System.Windows.Forms.Cursor.Current = Cursors.Default

                Try
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).ShowInTaskbar = False
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Maximum = 0
                Catch ex As Exception
                End Try
                Me.Close()
                WriteLog("<Process = 1 (Process GET VAT)> Complete")
            Case 2 'Update Running
                'Mouse(MON)
                System.Windows.Forms.Cursor.Current = Cursors.AppStarting

                pgbProcess.Maximum = 3
                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting New Document Number")
                If VATAutoRun = False Then
                    '/***************************************/
                    '     Manual Running Number
                    '/**************************************/
                    Call ProcessManulaNewDocNo()
                Else
                    '/***************************************/
                    '     VAT App Auto Running Number
                    '/**************************************/
                    Call ProcessNewDocNoCustom()
                End If
                Call ProcessShowTxt(pgbProcess.Value + 1, "Setting Complete")
                Try
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).ShowInTaskbar = False
                    CType(Owner.Controls("Panel1").Controls("BaseLineSeparate2").Controls("TaskBar"), wyDay.Controls.Windows7ProgressBar).Maximum = 0
                Catch ex As Exception
                End Try
                Me.Close()
                'Mouse(MOFF)
                System.Windows.Forms.Cursor.Current = Cursors.Default
        End Select
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Timmer.Visible = True
        swTime += 1
        Timmer.Text = Math.Floor(swTime / 3600).ToString("00") & ":" & Math.Floor(swTime / 60).ToString("00") & ":" & (swTime Mod 60).ToString("00")
    End Sub
End Class