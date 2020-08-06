Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.IO

Public Class InstallUtil

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    Private Sub InstallUtil_BeforeInstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles Me.BeforeInstall
        Try
            Dim paths() As String = Split(Environment.GetEnvironmentVariable("PATH"), ";")
            Dim fileInPath As String = "install_flash.exe"
            Dim path As String = ""
            For Each pathItem As String In paths
                If File.Exists(pathItem & "\" & fileInPath) Then
                    path = pathItem & "\" & fileInPath
                End If
            Next

            Dim p As New Process
            p.StartInfo.FileName = path
            p.Start()
            p.WaitForExit()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
