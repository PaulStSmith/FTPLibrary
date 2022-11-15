Public Class frmFTP

    Dim ftpRequest As System.Net.FtpWebRequest = Nothing

    Private Sub btnFTPClient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPClient.Click

        Me.txtLog.Text = "Creating FTP Client..." & vbCrLf

        Dim clientFTP As New FTPLibrary.ftpClient(Me.txtServer.Text, CInt(Me.txtPort.Text), Me.txtUser.Text, Me.txtPass.Text)

        Try
            Me.txtLog.Text &= "Refreshing Server Info..." & vbCrLf
            clientFTP.Async = Me.chkAsync.Checked
            Call clientFTP.Refresh()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & clientFTP.StatusDescription, MsgBoxStyle.Exclamation)
            Exit Sub
        End Try

        Dim sb As New System.Text.StringBuilder
        sb.Append("File List..." & vbCrLf)
        For Each f As FTPLibrary.ftpFile In clientFTP.FileList
            sb.Append(IIf(f.IsDirectory, "<DIR>", "     ") & f.Size.ToString.PadLeft(10) & " " & f.Name & vbCrLf)
        Next
        Me.txtLog.Text = sb.ToString()

    End Sub

End Class
