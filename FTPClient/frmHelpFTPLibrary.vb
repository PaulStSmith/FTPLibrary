Imports System.Windows.Forms

Friend Class frmHelpFTPLibrary

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Font = System.Drawing.SystemFonts.DialogFont

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Public Sub PrepareDialog(ByVal dirEntrySample As String)

        Dim strHTML As String = ""
        Dim isConnected As Boolean = False

        '*
        '* Are we connected?
        '*
        strHTML = ""
        Try
            isConnected = My.Computer.Network.Ping("geocities.com")
        Catch ex As Exception
        End Try

        If (isConnected) Then
            strHTML &= "<html>"
            strHTML &= "<head>"
            strHTML &= "<title>FTP Client</title>"
            strHTML &= "</head>"
            strHTML &= "<frameset rows=""*"">"
            strHTML &= "<frame name=""main"" src=""http://geocities.com/pjondevelopment/ooops.htm?" & dirEntrySample & """>"
            strHTML &= "</frameset>"
            strHTML &= "</html>"
        Else
            strHTML &= "<html>"

            strHTML &= "<head>"
            strHTML &= "<title>Ooops</title>"
            strHTML &= "<style type=""text/css"">"
            strHTML &= "<!--"
            strHTML &= "h1   { color: #FF0000 }"
            strHTML &= "body { font-family: Verdana; font-size: 10pt }"
            strHTML &= "-->"
            strHTML &= "</style>"
            strHTML &= "</head>"

            strHTML &= "<body>"

            strHTML &= "<h1>Ooops!</h1>"
            strHTML &= "<p>Looks like you connected to an FTP server that is unknown to this library.</p>"
            strHTML &= "<p>Would you like to help building a better FTP library?</p>"
            strHTML &= "<p>Please, connect to the Internet and click on the following link.</p>"
            strHTML &= "<p><a target=""_blank"" href=""http://geocities.com/pjondevelopment/ooops.htm?" & dirEntrySample & """>Report an Unknown FTP Server.</a></p>"

            strHTML &= "</body>"

            strHTML &= "</html>"
        End If

        Dim tempFile As String = System.IO.Path.GetTempFileName() & ".htm"

        With (New System.IO.StreamWriter(tempFile, False, System.Text.Encoding.UTF8))
            .Write(strHTML)
            .Close()
        End With

        Me.WebBrowser1.Navigate("file://" & tempFile)

    End Sub

End Class
