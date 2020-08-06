Public Class clsFuntion
    Private Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" ( _
    ByVal process As IntPtr, _
    ByVal minimumWorkingSetSize As Integer, _
    ByVal maximumWorkingSetSize As Integer) As Integer

    Public Shared Sub FlushMemory()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        If (Environment.OSVersion.Platform = PlatformID.Win32NT) Then
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
        End If
    End Sub

    Public Function ISNULL(ByVal Data As ADODB.Recordset, ByVal i As Integer) As Boolean
        If IsDBNull(Data.Collect(i)) Then
            ISNULL = True
        Else
            ISNULL = False
        End If
    End Function

    Public Function ChangeMonthToText(ByVal MM As String) As String
        Dim TextMonth As String = String.Empty
        MM = MonthToMM(MM)
        Select Case System.Threading.Thread.CurrentThread.CurrentCulture.Name
            Case "th-TH"
                Select Case MM
                    Case "01"
                        TextMonth = "���Ҥ�"
                    Case "02"
                        TextMonth = "����Ҿѹ��"
                    Case "03"
                        TextMonth = "�չҤ�"
                    Case "04"
                        TextMonth = "����¹"
                    Case "05"
                        TextMonth = "����Ҥ�"
                    Case "06"
                        TextMonth = "�Զع�¹"
                    Case "07"
                        TextMonth = "�á�Ҥ�"
                    Case "08"
                        TextMonth = "�ԧ�Ҥ�"
                    Case "09"
                        TextMonth = "�ѹ��¹"
                    Case "10"
                        TextMonth = "���Ҥ�"
                    Case "11"
                        TextMonth = "��Ȩԡ�¹"
                    Case "12"
                        TextMonth = "�ѹ�Ҥ�"
                End Select

            Case Else
                Select Case MM
                    Case "01"
                        TextMonth = "January"
                    Case "02"
                        TextMonth = "February"
                    Case "03"
                        TextMonth = "March"
                    Case "04"
                        TextMonth = "April"
                    Case "05"
                        TextMonth = "May"
                    Case "06"
                        TextMonth = "June"
                    Case "07"
                        TextMonth = "July"
                    Case "08"
                        TextMonth = "August"
                    Case "09"
                        TextMonth = "September"
                    Case "10"
                        TextMonth = "October"
                    Case "11"
                        TextMonth = "November"
                    Case "12"
                        TextMonth = "December"

                End Select
        End Select

        Return TextMonth
    End Function

    Public Function MonthToMM(ByVal mStr As String) As String
        MonthToMM = ""
        Select Case mStr.Length.ToString
            Case "1"
                MonthToMM = "0" & mStr
            Case "2"
                MonthToMM = mStr
        End Select
        Return MonthToMM
    End Function

    Public Function CultureYear(ByVal YYYY As String) As String
        CultureYear = ""
        Select Case System.Threading.Thread.CurrentThread.CurrentCulture.Name
            Case "th-TH"
                CultureYear = CInt(YYYY.Trim) - 543
            Case Else
                CultureYear = YYYY.Trim
        End Select
    End Function

    '���ѹ�á�ͧ��͹ �ҡ�ѹ�Ѩ�غѹ
    Function GetFirstDayOfMonth(ByVal CurrentDate As DateTime) As DateTime
        Return (New DateTime(CurrentDate.Year, CurrentDate.Month, 1))
    End Function



    '���ѹ�ش���¢ͧ��͹ �ҡ�ѹ�Ѩ�غѹ
    Function GetLastDayOfMonth(ByVal CurrentDate As DateTime) As DateTime
        With CurrentDate
            Return (New DateTime(.Year, .Month, Date.DaysInMonth(.Year, .Month)))
        End With
    End Function


    Function GetDayInMonth(ByVal ai_year As Integer, ByVal ai_month As Integer) As Integer
        Return DateTime.DaysInMonth(ai_year, ai_month)
    End Function

End Class
