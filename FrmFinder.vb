Imports System.Globalization

Public Class FrmFinder

    Dim CurColumn As String = ""
    Dim CurI As Integer = 0
    Dim OnSearch As Boolean = False
    Private Sub MbGlassButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click
        If OnSearch = True Then Exit Sub
        TextBox1.Focus()
        TextBox1.SelectAll()
        OnSearch = True
        Dim Grid As BaseClass.BaseGridview = CType(Owner, frmMain).lstVu
        If CurI = 0 Then CurI = 1
        For i As Integer = CurI To mocolVatDisPlay.Count
            If i > CurI Then
                CurColumn = ""
            End If

            Dim INVDATE As String
            Try
                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                fi.ShortDatePattern = "yyyyMMdd"
                INVDATE = DateTime.ParseExact(CStr(mocolVatDisPlay(i).INVDATE).Trim, "yyyyMMdd", fi).ToString("dd/MM/yyyy")
            Catch ex As Exception
                INVDATE = ""
            End Try

            If mocolVatDisPlay(i).DOCNO.Trim = TextBox1.Text.Trim And CurColumn <> "DOCNO" And CurI <> i Then
                CurColumn = "DOCNO"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("DOCNO")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf (INVDATE = TextBox1.Text.Trim) And CurColumn <> "TXDATE" And CurI <> i Then
                CurColumn = "TXDATE"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("TXDATE")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVNAME.Trim = TextBox1.Text.Trim And CurColumn <> "INVNAME" And CurI <> i Then
                CurColumn = "INVNAME"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVNAME")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).Batch.Trim = TextBox1.Text.Trim And CurColumn <> "BATCH" And CurI <> i Then
                CurColumn = "BATCH"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("BATCH")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVTAX.ToString("#,##0.00") = TextBox1.Text.Trim And CurColumn <> "INVTAX" And CurI <> i Then
                CurColumn = "INVTAX"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVTAX")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVAMT.ToString("#,##0.00") = TextBox1.Text.Trim And CurColumn <> "INVAMT" And CurI <> i Then
                CurColumn = "INVAMT"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVAMT")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).IDINVC.Trim = TextBox1.Text.Trim And CurColumn <> "IDINVC" And CurI <> i Then
                CurColumn = "IDINVC"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("IDINVC")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).NEWDOCNO.Trim = TextBox1.Text.Trim And CurColumn <> "NEWDOCNO" And CurI <> i Then
                CurColumn = "NEWDOCNO"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                Grid.CurrentCell = Grid.Rows(rowID).Cells("NEWDOCNO")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub

            End If
        Next

        CurColumn = ""
        CurI = 0
        For i As Integer = 1 To mocolVatDisPlay.Count
            If i > CurI Then
                CurColumn = ""
            End If

            Dim TXDATE, INVDATE As String
            Try
                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                fi.ShortDatePattern = "yyyyMMdd"
                TXDATE = DateTime.ParseExact(CStr(mocolVatDisPlay(i).TXDATE).Trim, "yyyyMMdd", fi).ToString("dd/MM/yyyy")
                INVDATE = DateTime.ParseExact(CStr(mocolVatDisPlay(i).INVDATE).Trim, "yyyyMMdd", fi).ToString("dd/MM/yyyy")
            Catch ex As Exception
                TXDATE = ""
                INVDATE = ""
            End Try

            If mocolVatDisPlay(i).DOCNO.Trim = TextBox1.Text.Trim And CurColumn <> "DOCNO" And CurI <> i Then
                CurColumn = "DOCNO"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("DOCNO") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("DOCNO")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf (TXDATE = TextBox1.Text.Trim Or INVDATE = TextBox1.Text.Trim) And CurColumn <> "TXDATE" And CurI <> i Then
                CurColumn = "TXDATE"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("TXDATE") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("TXDATE")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVNAME.Trim = TextBox1.Text.Trim And CurColumn <> "INVNAME" And CurI <> i Then
                CurColumn = "INVNAME"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("INVNAME") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVNAME")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).Batch.Trim = TextBox1.Text.Trim And CurColumn <> "BATCH" And CurI <> i Then
                CurColumn = "BATCH"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("BATCH") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("BATCH")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVTAX.ToString("#,##0.00") = TextBox1.Text.Trim And CurColumn <> "INVTAX" And CurI <> i Then
                CurColumn = "INVTAX"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("INVTAX") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVTAX")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).INVAMT.ToString("#,##0.00") = TextBox1.Text.Trim And CurColumn <> "INVAMT" And CurI <> i Then
                CurColumn = "INVAMT"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("INVAMT") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("INVAMT")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).IDINVC.Trim = TextBox1.Text.Trim And CurColumn <> "IDINVC" And CurI <> i Then
                CurColumn = "IDINVC"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("IDINVC") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("IDINVC")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub
            ElseIf mocolVatDisPlay(i).NEWDOCNO.Trim = TextBox1.Text.Trim And CurColumn <> "NEWDOCNO" And CurI <> i Then
                CurColumn = "NEWDOCNO"
                CurI = i
                Dim pageto As Integer = 0
                Dim rowID As Integer = 0
                If Use_DisplayPage = 15 Then
                    pageto = 1
                    rowID = i - 1
                Else
                    pageto = Math.Ceiling(i / (Use_DisplayPage * 10))
                    rowID = IIf(((i Mod (Use_DisplayPage * 10)) - 1) = -1, (Use_DisplayPage * 10) - 1, (i Mod (Use_DisplayPage * 10)) - 1)
                End If
                CType(Owner, frmMain).cboPage.Text = pageto
                If Grid.CurrentCell Is Grid.Rows(rowID).Cells("NEWDOCNO") Then
                    OnSearch = False
                    Exit For
                End If
                Grid.CurrentCell = Grid.Rows(rowID).Cells("NEWDOCNO")
                Grid.CurrentRow.Selected = True
                OnSearch = False
                Exit Sub

            End If
        Next
        Me.Show()
        OnSearch = False
        ShowToolTip(TextBox1, ToolTipIcon.Info, "Find not found!!")
    End Sub

    Private Sub FrmFinder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            CType(Owner, frmMain).FrmFinderText = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmFinder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            MbGlassButton1_Click(Nothing, Nothing)
        End If
    End Sub
End Class