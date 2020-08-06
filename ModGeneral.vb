Option Strict Off
Option Explicit On
Module ModGeneral

    Public CheckTTypeRunning As Integer
    Public CheckRate As String
    Public Const MON As Integer = 0 'MousePointer
    Public Const MOFF As Integer = 1 'MousePointer
    Public Const TLEFT As Integer = 0 'Text Left Justify
    Public Const TRIGHT As Integer = 1 'Text Left Justify
    'Public Report As New CRAXDRT.Report
    Public ConnDB As String
	
	Public Sub ToggleSort(ByRef a_cntr As System.Windows.Forms.ListView)
		'UPGRADE_ISSUE: MSComctlLib.ListView property a_cntr.Sorted was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'If Not (a_cntr.Sorted) Or a_cntr.Sorting = System.Windows.Forms.SortOrder.Descending Then
        If Not (a_cntr.Sorting) Or a_cntr.Sorting = System.Windows.Forms.SortOrder.Descending Then
            a_cntr.Sorting = System.Windows.Forms.SortOrder.Ascending
        Else
            a_cntr.Sorting = System.Windows.Forms.SortOrder.Descending
        End If
	End Sub
    Public Sub Mouse(ByRef ai_job As Integer)
        Select Case ai_job
            Case Is = MON
                'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.AppStarting
                System.Windows.Forms.Cursor.Current = Cursors.AppStarting
            Case Is = MOFF
                'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Windows.Forms.Cursor.Current = Cursors.Default
        End Select
    End Sub
	
    Public Sub HighlightBox(ByRef ao_Box As TextBox)
        With ao_Box
            .SelectionStart = 0
            .SelectionLength = Len(ao_Box.Text)
        End With
    End Sub
	
	Public Function ReQuote(ByRef strtext As String) As String
		ReQuote = Replace(strtext, "'", "''")
	End Function
	
	Public Function OnlyNumericKey(ByVal aKey As Short) As Short
		
		Select Case aKey
			
			Case 48 To 57
				OnlyNumericKey = aKey
			Case System.Windows.Forms.Keys.Back
                OnlyNumericKey = aKey
                'Case 62
                '    OnlyNumericKey = aKey
                'Case 60
                '    OnlyNumericKey = aKey
            Case Else
                Beep()
                OnlyNumericKey = 0
        End Select
		
	End Function
	
    Public Function limitStr(ByRef mytext As String, ByRef mylength As Integer, Optional ByRef mystyle As Integer = 0) As String
        limitStr = ""
        Select Case mystyle

            Case Is = 0
                If Len(mytext) < mylength Then
                    limitStr = mytext & Space(mylength - Len(mytext))
                Else
                    limitStr = Mid(mytext, 1, mylength)
                End If
            Case Is = 1
                If Len(mytext) < mylength Then
                    limitStr = Space(mylength - Len(mytext)) & mytext
                Else
                    limitStr = Mid(mytext, 1, mylength)
                End If
        End Select
    End Function
	
    Public Function splitref(ByRef chaine As String, ByRef separ As String, ByRef tableau() As String, ByRef nb_elem As Integer) As Integer
        On Error GoTo erreur
        Dim pos_act, pos_occur As Integer

        If Right(chaine, 1) <> separ Then chaine = chaine & separ
        Do
            pos_act = pos_occur + Len(separ)
            pos_occur = InStr(pos_act, chaine, separ)
            If pos_occur <> 0 Then
                ReDim Preserve tableau(nb_elem)
                tableau(nb_elem) = Mid(chaine, pos_act, pos_occur - pos_act)
                nb_elem = nb_elem + 1
            End If
        Loop Until pos_occur = 0
        splitref = 0
        Exit Function
erreur:
        splitref = Err.Number
    End Function
	'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    'Public Function DateToDB(ByRef str_Renamed As String) As String

    '       Dim tableau_a_remplir() As String
    '	Dim mm As String
    '	Dim dd As String
    '	Dim yyyy As String
    '       Dim nombre_elements, rep As Integer
    '	DateToDB = ""
    '	dd = ""
    '	mm = ""
    '       yyyy = ""
    '	If str_Renamed <> "" Then
    '		rep = splitref(Trim(str_Renamed), "/", tableau_a_remplir, nombre_elements)
    '           If rep = 0 Then

    '               If Len(tableau_a_remplir(0)) <> 2 Then
    '                   dd = "0" & tableau_a_remplir(0)
    '               Else
    '                   dd = tableau_a_remplir(0)
    '               End If
    '               If Len(tableau_a_remplir(1)) <> 2 Then
    '                   mm = "0" & tableau_a_remplir(1)
    '               Else
    '                   mm = tableau_a_remplir(1)
    '               End If
    '               If Len(tableau_a_remplir(2)) <> 4 Then
    '                   yyyy = "20" & tableau_a_remplir(2)
    '               Else
    '                   yyyy = tableau_a_remplir(2)
    '               End If
    '               DateToDB = yyyy & mm & dd
    '           Else
    '               DateToDB = CStr(0)
    '           End If
    '	End If
    '   End Function

    Public Function sZero(ByVal len As Integer) As String
        Dim Result As String = ""
        For i As Integer = 0 To len - 1
            Result &= "0"
        Next
        Return Result
    End Function

    Public Function ZeroDate(ByVal Value As String, ByVal Format1 As String, ByVal Format2 As String) As String
        Dim INVDATE As String = ""
        Try
            Dim TM() As String = Value.Split("/")
            Dim LN() As String = Format1.Split("/")
            Value = ""
            If TM.Length > 0 Then
                If IsNumeric(TM(0)) Then
                    TM(0) = CInt(TM(0)).ToString(sZero(LN(0).Trim.Length))
                    Value &= TM(0) & "/"
                End If
            End If
            If TM.Length > 1 Then
                If IsNumeric(TM(1)) Then
                    TM(1) = CInt(TM(1)).ToString(sZero(LN(1).Trim.Length))
                    Value &= TM(1) & "/"
                End If
            End If
            If TM.Length > 2 Then
                If IsNumeric(TM(2)) Then
                    TM(2) = CInt(TM(2)).ToString(sZero(LN(2).Trim.Length))
                    Value &= TM(2)
                End If
            End If
        Catch ex As Exception
            INVDATE = "0"
        End Try
        Return INVDATE
    End Function
    Public Function TryDate2(ByVal Value As String, ByVal Format1 As String, ByVal Format2 As String) As String
        Dim INVDATE As String
        Try
            Dim fi As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo
            fi.ShortDatePattern = Format1
            INVDATE = DateTime.ParseExact(Value.Trim, Format1, fi).ToString(Format2)
        Catch ex1 As Exception
            INVDATE = "0"
            Dim F As String = ""
            Try
                Dim S() As String = Value.Split("/")
                If S.Length > 0 Then
                    For N As Integer = 0 To S(0).Length - 1
                        F &= "d"
                    Next
                End If
                If S.Length > 1 Then
                    F &= "/"
                    For N As Integer = 0 To S(1).Length - 1
                        F &= "M"
                    Next
                End If
                If S.Length > 2 Then
                    F &= "/"
                    For N As Integer = 0 To S(1).Length - 1
                        F &= "y"
                    Next
                End If
                Dim fi As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo
                fi.ShortDatePattern = Format1
                INVDATE = DateTime.ParseExact(Value.Trim, F, fi).ToString(Format2)
            Catch ex2 As Exception
                INVDATE = "0"
            End Try
        End Try
        Return INVDATE
    End Function

    'Public Function DateToD(ByRef str_Renamed As String, ByVal YY As String) As String

    '    Dim tableau_a_remplir() As String
    '    Dim mm As String
    '    Dim dd As String
    '    Dim yyyy, Tmp As String
    '    Dim nombre_elements, rep, i As Integer
    '    DateToD = ""
    '    dd = ""
    '    mm = ""
    '    yyyy = ""
    '    If str_Renamed <> "" Then
    '        rep = splitref(Trim(str_Renamed), "/", tableau_a_remplir, nombre_elements)
    '        If rep = 0 And InStr(str_Renamed, "/") > 0 Then
    '            If Len(tableau_a_remplir(0)) <> 2 Then
    '                If Len(tableau_a_remplir(0)) > 2 Then
    '                    dd = "01"
    '                Else
    '                    dd = "0" & tableau_a_remplir(0)
    '                End If
    '            Else
    '                dd = tableau_a_remplir(0)
    '            End If
    '            If Len(tableau_a_remplir(1)) <> 2 Then
    '                mm = "0" & tableau_a_remplir(1)
    '            Else
    '                mm = tableau_a_remplir(1)
    '            End If
    '            If nombre_elements = 3 Then
    '                If Len(tableau_a_remplir(2)) <> 4 Then
    '                    yyyy = "20" & tableau_a_remplir(2)
    '                Else
    '                    yyyy = tableau_a_remplir(2)
    '                End If
    '            Else
    '                yyyy = Right(YY, 4)
    '            End If
    '            DateToD = yyyy & mm & dd
    '        Else
    '            yyyy = Right(Trim(YY), 4)
    '            i = InStr(YY, "/")
    '            Tmp = YY
    '            If i > 0 Then
    '                dd = Mid(YY, 1, i - 1)
    '                Tmp = Mid(YY, i + 1)
    '                If Len(Trim(dd)) = 1 Then
    '                    dd = "0" & dd
    '                End If
    '            End If
    '            i = InStr(Tmp, "/")
    '            If i > 0 Then
    '                mm = Mid(Tmp, 1, i - 1)
    '                If Len(Trim(mm)) = 1 Then
    '                    mm = "0" & mm
    '                End If
    '            End If
    '            DateToD = yyyy & mm & dd
    '        End If
    '    Else
    '        Return 0
    '    End If
    'End Function

    Public Function ISNULL(ByVal Data As ADODB.Recordset, ByVal i As Integer) As Boolean
        If IsDBNull(Data.Collect(i)) Then
            ISNULL = True
        Else
            ISNULL = False
        End If
    End Function
    Public Function ISNULL(ByVal Data As ADODB.Recordset, ByVal i As String) As Boolean
        If IsDBNull(Data.Collect(i)) Then
            ISNULL = True
        Else
            ISNULL = False
        End If
    End Function
End Module