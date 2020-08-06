Public Module BaseVariable

    Public Use_LEGALNAME_inCB As Boolean = False
    Public Use_LEGALNAME_inAP As Boolean = False
    Public Use_LEGALNAME_inAR As Boolean = False
    Public Use_DisplayPage As Integer = 5
    Public Use_POST As Boolean = True
    Public Use_TAX_AP As TAXFROM = TAXFROM.URL
    Public Use_TAX_AR As TAXFROM = TAXFROM.URL
    Public Use_TAX_CB As TAXFROM = TAXFROM.URL
    Public Use_TAX_GL As TAXFROMGL = TAXFROMGL.VEN_CUS

    Public Use_BRANCH_AP As TAXFROM = TAXFROM.URL
    Public Use_BRANCH_AR As TAXFROM = TAXFROM.URL
    Public Use_BRANCH_CB As TAXFROM = TAXFROM.URL
    Public Use_BRANCH_GL As TAXFROMGL = TAXFROMGL.VEN_CUS

    Public UserLogin As String = String.Empty
    Public UserName As String = ""
    Public Use_Approved As Boolean = False
    Public Use_INVFROMPO As Boolean = True
    Public GLREVERSE As Boolean = False
    Public CBREVERSE As Boolean = False

    Public S_AP As Boolean = True
    Public S_AR As Boolean = True
    Public S_GL As Boolean = True
    Public S_CB As Boolean = True
    Public S_PO As Boolean = True
    Public s_OE As Boolean = True

    Public Enum TAXFROM
        URL = 0
        EMAIL = 1
        OPF = 2
        RegisNumber = 3
        BRN = 4
        State = 2
        Country = 3

    End Enum

    Public Enum TAXFROMGL
        VEN_CUS = 0
        OPF = 1
        COMMENT4 = 2
        COMMENT5 = 3
    End Enum

    Public Sub ShowToolTip(ByRef _obj As Object, ByVal pMessageTitle As ToolTipIcon, ByVal pMessageWarning As String, Optional ByVal vPOINTx As Integer = 0, Optional ByVal vPOINTy As Integer = 0)
        Dim _TToolTip As New ToolTip

        Try
            _TToolTip.IsBalloon = True
            _TToolTip.ShowAlways = True
            _TToolTip.ToolTipIcon = pMessageTitle
            _TToolTip.ToolTipTitle = pMessageTitle.ToString & "!"  '"Warning"

            Dim TopForm As Form = _obj.TopLevelControl
            Dim parent As Control = _obj.Parent
            Dim parentRectangle As System.Drawing.Rectangle
            Dim ToolTipTime As Integer

            ToolTipTime = Len(pMessageWarning) * 75
            If ToolTipTime < 1500 Then
                ToolTipTime = 1500
            ElseIf ToolTipTime > 4000 Then
                ToolTipTime = 4000
            End If
            If TopForm.IsMdiContainer Then
                TopForm = TopForm.ActiveMdiChild
                parentRectangle = TopForm.Parent.DisplayRectangle
            Else
                parentRectangle = Screen.PrimaryScreen.WorkingArea
            End If
            If _obj.Location.X + (_obj.Width / 2) > TopForm.ClientSize.Width Then
                TopForm.Width = TopForm.Width + ((_obj.Location.X + (_obj.Width)) - TopForm.ClientSize.Width) + (TopForm.Width - TopForm.ClientSize.Width)
                If (TopForm.Location.X + TopForm.Width) > parentRectangle.Right Then
                    TopForm.Location = New System.Drawing.Point(parentRectangle.Right - TopForm.Width, TopForm.Location.Y)
                End If
            End If

            _TToolTip.Show(String.Empty, _obj, 0)
            If vPOINTx = 0 And vPOINTy = 0 Then
                _TToolTip.Show(pMessageWarning, _obj, New Point(_obj.Width / 2, _obj.height / 2), ToolTipTime)
            Else
                _TToolTip.Show(pMessageWarning, _obj, New Point(vPOINTx, vPOINTy), ToolTipTime)
            End If
            _obj.Focus()
            Application.DoEvents()
        Catch ex As Exception

        End Try

    End Sub
End Module
