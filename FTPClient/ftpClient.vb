'**************************************************
' FILE      : ftpClient.vb
' AUTHOR    : Paulo.Santos
' CREATION  : 2005.DEC.23
' COPYRIGHT : Copyright © 2005-2007
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       Simple FTP Client.
'
' Change log:
' 0.1   2005.DEC.23
'       Paulo.Santos
'       Created.
'***************************************************

''' <summary>
''' A mask of bits that represents the permission for each file and folder on the FTP server.
''' </summary>
Public Enum AttributeFlags As UShort
    ''' <summary>
    ''' Is a Directory (Folder)
    ''' </summary>
    Directory = &H800
    ''' <summary>
    ''' Owner can Read
    ''' </summary>
    OwnerRead = &H100
    ''' <summary>
    ''' Owner can Write
    ''' </summary>
    OwnerWrite = &H80
    ''' <summary>
    ''' Owner can Execute
    ''' </summary>
    OwnerExecute = &H40
    ''' <summary>
    ''' Group can Read
    ''' </summary>
    GroupRead = &H20
    ''' <summary>
    ''' Group can Write
    ''' </summary>
    GroupWrite = &H10
    ''' <summary>
    ''' Group can Execute
    ''' </summary>
    GroupExecute = &H8
    ''' <summary>
    ''' Public can Read
    ''' </summary>
    PublicRead = &H4
    ''' <summary>
    ''' Public can Write
    ''' </summary>
    PublicWrite = &H2
    ''' <summary>
    ''' Public can Execute
    ''' </summary>
    PublicExecute = &H1
    ''' <summary>
    ''' Added for completition. Hardly any file will have no attributes.
    ''' </summary>
    None = 0
End Enum

''' <summary>
''' Represents a file or folder on the FTP server.
''' </summary>
Public Class ftpFile
    ''' <summary>
    ''' The name of the file or folder.
    ''' </summary>
    Public Name As String
    ''' <summary>
    ''' The size of the file in bytes. Folders always have size 0.
    ''' </summary>
    Public Size As ULong
    ''' <summary>
    ''' The owner of the file or folder. If available.
    ''' </summary>
    Public Owner As String
    ''' <summary>
    ''' The timestamp of the file.
    ''' </summary>
    Public [Date] As DateTime

    Private __Flags As AttributeFlags

    ''' <summary>
    ''' A map of bits that represents the attributes of the file or folder.
    ''' </summary>
    Public Property Flags() As AttributeFlags
        Get
            Return __Flags
        End Get
        Set(ByVal value As AttributeFlags)
            __Flags = value
        End Set
    End Property

    ''' <summary>
    ''' Returns a System.Boolean that indicates if this instance of ftpFile Class is a folder (or directory).
    ''' </summary>
    Public Function IsDirectory() As Boolean
        Return CBool(__Flags And AttributeFlags.Directory)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the owner can read the file.
    ''' </summary>
    Public Function CanOwnerRead() As Boolean
        Return CBool(__Flags And AttributeFlags.OwnerRead)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the owner can write the file.
    ''' </summary>
    Public Function CanOwnerWrite() As Boolean
        Return CBool(__Flags And AttributeFlags.OwnerWrite)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the owner can execute the file.
    ''' </summary>
    Public Function CanOwnerExecute() As Boolean
        Return CBool(__Flags And AttributeFlags.OwnerExecute)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the group can read the file.
    ''' </summary>
    Public Function CanGroupRead() As Boolean
        Return CBool(__Flags And AttributeFlags.GroupRead)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the group can write the file.
    ''' </summary>
    Public Function CanGroupWrite() As Boolean
        Return CBool(__Flags And AttributeFlags.GroupWrite)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the group can execute the file.
    ''' </summary>
    Public Function CanGroupExecute() As Boolean
        Return CBool(__Flags And AttributeFlags.GroupExecute)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the public can read the file.
    ''' </summary>
    Public Function CanPublicRead() As Boolean
        Return CBool(__Flags And AttributeFlags.PublicRead)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the public can write the file.
    ''' </summary>
    Public Function CanPublicWrite() As Boolean
        Return CBool(__Flags And AttributeFlags.PublicWrite)
    End Function

    ''' <summary>
    ''' Returns a System.Boolean that indicates if the public can execute the file.
    ''' </summary>
    Public Function CanPublicExecute() As Boolean
        Return CBool(__Flags And AttributeFlags.PublicExecute)
    End Function

End Class

''' <summary>
''' Represents the current path on the FTP server.
''' </summary>
Public Class ftpDirectory
    Inherits ArrayList

    ''' <summary>
    ''' Returns an ftpFile object.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the element to get or set.</param>
    Public Shadows Property Item(ByVal Index As Integer) As ftpFile
        Get
            Return MyBase.Item(Index)
        End Get
        Set(ByVal value As ftpFile)
            MyBase.Item(Index) = value
        End Set
    End Property

    ''' <summary>
    ''' Add a new ftpFile object that represents a file or folder on the FTP server
    ''' </summary>
    ''' <param name="file">An ftpFile object that represents a file or folder on the FTP server.</param>
    Friend Shadows Sub Add(ByVal file As ftpFile)
        MyBase.Add(file)
    End Sub

End Class

''' <summary>
''' An FTP client build on top of the System.Net.FtpWebRequest class.
''' </summary>
Public Class ftpClient

#Region " Enumerators "

    ''' <summary>
    ''' Enumerate the known FTP servers.
    ''' </summary>
    ''' <remarks>
    ''' This needs attention because there are dozens of FTP servers 
    ''' out there and each one of them has a different approach.
    ''' <p/>
    ''' We detect the differences between them based on the way 
    ''' they send the directory information, the response 
    ''' to the LIST command.
    ''' <p/>
    ''' So far we have three known servers:
    ''' <ul>
    ''' <li>Microsoft</li>
    ''' <li>Unix A</li>
    ''' <li>Unix B</li>
    ''' </ul>
    ''' The Unix'es A and B are because we couldn't determine 
    ''' the exact FTP server that repplied to our tests.
    ''' <p/>
    ''' As soon we manage to know their proper names this 
    ''' enumerator will be updated accordingly, however when this 
    ''' happen, we probably will mantain the UnixA and UnixB 
    ''' enumerators further for backwards compatibility.
    ''' </remarks>
    Public Enum HostType As UShort
        ''' <summary>
        ''' An Unix FTP server.
        ''' </summary>
        UnixA = 0
        ''' <summary>
        ''' An Unix FTP server.
        ''' </summary>
        UnixB
        ''' <summary>
        ''' A Microsoft FTP server.
        ''' </summary>
        Microsoft = &H4000
        ''' <summary>
        ''' Added for completition. We cannot process commands on an Unknown FTP server.
        ''' </summary>
        Unknown = &H8000
    End Enum

#End Region

#Region " EventArgs "

    ''' <summary>
    ''' Stores information on asynchornous operations.
    ''' </summary>
    Public Class RequestProgressEventArgs
        Inherits EventArgs

        Private __BytesSent As Integer
        Private __BytesTotal As Integer
        Private __BytesReceived As Integer

        ''' <summary>
        ''' Initializes a new instance of the RequestProgressEventArgs
        ''' </summary>
        ''' <param name="BytesSent">The number of the bytes sent on this request.</param>
        ''' <param name="BytesReceived">The number of the bytes received on this request.</param>
        ''' <param name="BytesTotal">The total amount of bytes expected to be sent or recieved.</param>
        Protected Friend Sub New(ByVal BytesSent As Integer, ByVal BytesReceived As Integer, ByVal BytesTotal As Integer)
            __BytesSent = BytesSent
            __BytesTotal = BytesTotal
            __BytesReceived = BytesReceived
        End Sub

        ''' <summary>
        ''' The number of the bytes sent on this request.
        ''' </summary>
        Public ReadOnly Property BytesSent() As Integer
            Get
                Return __BytesSent
            End Get
        End Property

        ''' <summary>
        ''' The number of the bytes received on this request.
        ''' </summary>
        Public ReadOnly Property BytesReceived() As Integer
            Get
                Return __BytesReceived
            End Get
        End Property

        ''' <summary>
        ''' The total amount of bytes expected to be sent or recieved.
        ''' </summary>
        Public ReadOnly Property BytesTotal() As Integer
            Get
                Return __BytesTotal
            End Get
        End Property

    End Class

#End Region

#Region " Public Events "

    ''' <summary>
    ''' Called before any request is made.
    ''' </summary>
    Public Event onBeforeRequest(ByVal sender As Object, ByVal e As EventArgs)
    ''' <summary>
    ''' Called during request to inform the main thread of the progress.
    ''' </summary>
    Public Event onRequestProgress(ByVal sender As Object, ByVal e As RequestProgressEventArgs)
    ''' <summary>
    ''' Called on the completition of a request.
    ''' </summary>
    Public Event onRequestComplete(ByVal sender As Object, ByVal e As EventArgs)

#End Region

#Region " Internal Fields "

    Private __ftpRequest As System.Net.FtpWebRequest
    Private __ftpResponse As System.Net.FtpWebResponse
    Private __ftpRequestStream As System.IO.Stream

    Private __Busy As Boolean = False
    Private __Path As String = ""
    Private __Async As Boolean = False
    Private __Files As New ftpDirectory
    Private __UseSSL As Boolean = False
    Private __Server As System.Uri = Nothing
    Private __HostType As HostType = HostType.Unknown
    Private __BufferSize As Integer = 1024
    Private __CurrentPath As String
    Private __LocalStream As System.IO.Stream
    Private __UserCredentials As System.Net.NetworkCredential = Nothing
    Private __AllowUserInput As Boolean = True

    ''' <summary>
    ''' The object that sinchronize all the asynchornous transfers.
    ''' </summary>
    Private __OperationComplete As New System.Threading.ManualResetEvent(False)

    ''' <summary>
    ''' The parser that will be used to decode the response from the FTP server to the LIST command.
    ''' </summary>
    Private __Parser As System.Text.RegularExpressions.Regex = Nothing

    ''' <summary>
    ''' A list of known server and its respective regular expressions.
    ''' </summary>
    ''' <remarks>
    ''' __Parsers is an two-dimensional array with the following structure:
    ''' <p/>
    ''' <code>{
    '''   { HostType.identifier, "Parser Regular Expression" },
    '''   { HostType.identifier, "Parser Regular Expression" }, ...
    ''' }</code>
    ''' <p/>
    ''' The Parser Regular Expression MUST be coded with named capture groups. 
    ''' This are the capture groups expected for the parser routine:
    ''' <p/>
    ''' <table border="0" cellpadding="2" cellspacing="0">
    ''' <tr>
    ''' <td>name</td><td>The name of the file or folder</td>
    ''' <td>size</td><td>the size of the file</td>
    ''' <td>timestamp</td><td>The modification date of the file or folder</td>
    ''' <td>owner</td><td>the owner of the file or folder</td>
    ''' <td>permOwner</td><td>file Owner's permissions</td>
    ''' <td>permGroup</td><td>file Groups's permissions</td>
    ''' <td>permPublic</td><td>The name of the file or folder</td>
    ''' <td>dir</td><td>file Public's permissions</td>
    ''' </tr>
    ''' </table>
    ''' <p/>
    ''' At the very least the Regular Expression MUST include the 
    ''' groups dir, name, size and timestamp.
    ''' </remarks>
    Private __Parsers As Object(,) _
                = { _
                    {HostType.UnixA, "(?<dir>[-d])(?<permOwner>[-r][-w][-x])(?<permGroup>[-r][-w][-x])(?<permPublic>[-r][-w][-x])\s+\d+\s+(?<user>\w+)\s+(?<group>\w+)\s+(?<size>\d+)\s+(?<timestamp>\w+\s+(\s|\d)\d\s+\d\d:\d\d)\s+(?<name>.+)"}, _
                    {HostType.UnixB, "(?<dir>[-d])(?<permOwner>[-r][-w][-x])(?<permGroup>[-r][-w][-x])(?<permPublic>[-r][-w][-x])\s+\d+\s+(?<user>\w+)\s+(?<group>\w+)\s+(?<size>\d+)\s+(?<timestamp>\w+\s+(\s|\d)\d\s+\d{4})\s+(?<name>.+)"}, _
                    {HostType.Microsoft, "(?<timestamp>\d\d-\d\d-\d{2,4}\s+\d\d:\d\d([aApP][mM]))\s+(?<dir><DIR>)?\s+(?<size>\d+)?\s+(?<name>.+)"} _
                   }

#End Region

#Region " Constructors "

    ''' <summary>
    ''' Initializes a new instance of the ftpClient class.
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ftpClient class.
    ''' </summary>
    ''' <param name="Server">A <see cref="T:System.Uri"/> object that represent the FTP server.</param>
    Public Sub New(ByVal Server As System.Uri)
        Call MyClass.New()
        __Server = Server
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ftpClient class.
    ''' </summary>
    ''' <param name="Server">A <see cref="T:System.Uri"/> object that represent the FTP server.</param>
    ''' <param name="UserCredentials">A <see cref="T:System.Net.NetworkCredential"/> object that holds the user name and password to be used against the FTP server.</param>
    Public Sub New(ByVal Server As System.Uri, ByVal UserCredentials As System.Net.NetworkCredential)
        Call MyClass.New(Server)
        __UserCredentials = UserCredentials
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ftpClient class.
    ''' </summary>
    ''' <param name="Server">The Full Qualified Domain Name of the FTP server.</param>
    ''' <param name="Port">The Port number where the FTP client will be connected. Default 21.</param>
    Public Sub New(ByVal Server As String, ByVal Port As Integer)
        Call MyClass.New(New System.Uri("ftp://" & Server & ":" & Port.ToString))
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the ftpClient class.
    ''' </summary>
    ''' <remarks>This constructor request a password. Unless the UseSSL property is set to True the password will be transmited as plain text. Anyone monitoring the network at that moment may see it.</remarks>
    ''' <param name="Server">The Full Qualified Domain Name of the FTP server.</param>
    ''' <param name="Port">The Port number where the FTP client will be connected. Default 21.</param>
    ''' <param name="UserName">The user name for the FTP server.</param>
    ''' <param name="Password">The password for the user.</param>
    Public Sub New(ByVal Server As String, ByVal Port As Integer, ByVal UserName As String, ByVal Password As String)
        Call MyClass.New(Server, Port)
        __UserCredentials = New System.Net.NetworkCredential(UserName, Password)
    End Sub

#End Region

#Region " Public Properties "

    ''' <summary>
    ''' Gets of sets if the data transfer will be conducted asynchronously.
    ''' </summary>
    Public Property Async() As Boolean
        Get
            Return __Async
        End Get
        Set(ByVal value As Boolean)
            If (__Busy) AndAlso (__Async <> value) Then
                Throw New InvalidOperationException("Cannot change switch asynchronous mode while an operation is still in progress.")
            End If
            __Async = value
        End Set
    End Property

    ''' <summary>
    ''' The size, in bytes, of the buffer used during data transfers. Default 1024 bytes.
    ''' </summary>
    Public Property BufferSize() As Integer
        Get
            Return __BufferSize
        End Get
        Set(ByVal value As Integer)
            __BufferSize = value
        End Set
    End Property

    ''' <summary>
    ''' Indicates if the instance of the ftpClient class is transfering data.
    ''' </summary>
    Public ReadOnly Property Busy() As Boolean
        Get
            Return __Busy
        End Get
    End Property

    ''' <summary>
    ''' An ftpDirectory instance that stores the list of files and folders of the current Path on the FTP server.
    ''' </summary>
    Public ReadOnly Property FileList() As ftpDirectory
        Get
            Return __Files
        End Get
    End Property

    ''' <summary>
    ''' The current path on the server
    ''' </summary>
    Public Property Path() As String
        Get
            Return __Path
        End Get
        Set(ByVal value As String)
            If (__Busy) Then
                Throw New InvalidOperationException("Cannot change path while an operation is still in progress.")
            End If
            __Path = value
            Refresh()
        End Set
    End Property

    ''' <summary>
    ''' A <see cref="T:System.Uri"/> instance that represents the FTP server.
    ''' </summary>
    Public Property Server() As System.Uri
        Get
            Return __Server
        End Get
        Set(ByVal value As System.Uri)
            If (__Busy) Then
                Throw New InvalidOperationException("Cannot change server while an operation is still in progress.")
            End If
            __Server = value
            Call Me.Refresh()
        End Set
    End Property

    ''' <summary>
    ''' A numeric code indicating the last response from the server.
    ''' </summary>
    ''' <remarks>
    ''' More information can be obtained in the RFC959.
    ''' </remarks>
    Public ReadOnly Property StatusCode() As System.Net.FtpStatusCode
        Get
            If (__ftpResponse Is Nothing) Then
                Return Net.FtpStatusCode.Undefined
            Else
                Return __ftpResponse.StatusCode
            End If
        End Get
    End Property

    ''' <summary>
    ''' The human readable description of the StatusCode.
    ''' </summary>
    Public ReadOnly Property StatusDescription() As String
        Get
            If (__ftpResponse Is Nothing) Then
                Return "Undefined"
            Else
                Return __ftpResponse.StatusDescription
            End If
        End Get
    End Property

    ''' <summary>
    ''' An instance of the <see cref="T:System.Net.NetworkCredential"/> class that represents the FTP User Account used to log in the remote server.
    ''' </summary>
    Public Property UserCredentials() As System.Net.NetworkCredential
        Get
            Return __UserCredentials
        End Get
        Set(ByVal value As System.Net.NetworkCredential)
            If (__Busy) Then
                Throw New InvalidOperationException("Cannot change user while an operation is still in progress.")
            End If
            __UserCredentials = value
            Call Me.Refresh()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a <see cref="T:System.Boolean"/> that specifies that an SSL connection should be used.
    ''' </summary>
    Public Property UseSSL() As Boolean
        Get
            Return __UseSSL
        End Get
        Set(ByVal value As Boolean)
            __UseSSL = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a <see cref="T:System.Boolean"/> that specifies that in case of unkown FTP servers a window will be displayed to the user to request information in order to improve the FTP Library.
    ''' </summary>
    Public Property AllowUserInput() As Boolean
        Get
            Return __AllowUserInput
        End Get
        Set(ByVal value As Boolean)
            __AllowUserInput = value
        End Set
    End Property

#End Region

#Region " Public Methods "

    ''' <summary>
    ''' Deletes a file on the Server. 
    ''' </summary>
    ''' <param name="FileName">The name of the file to be deleted.</param>
    ''' <remarks>
    ''' After calling this method is necessary to call the <see cref="M:ftpClient.Refresh"/> method to ensure that the <see cref="P:ftpClient.FileList"/> property is up to date.
    ''' </remarks>
    Public Overloads Sub DeleteFile(ByVal FileName As String)
        If (__Busy) Then
            Throw New InvalidOperationException("Cannot delete a file while an operation is still in progress.")
        End If
        Call ExecCommand(System.Net.WebRequestMethods.Ftp.DeleteFile, FileName)
    End Sub

    ''' <summary>
    ''' Deletes a file on the Server. 
    ''' </summary>
    ''' <param name="File">An ftpFile object that represents the file to be deleted.</param>
    ''' <remarks>
    ''' After calling this method is necessary to call the <see cref="M:ftpClient.Refresh"/> method to ensure that the <see cref="P:ftpClient.FileList"/> property is up to date.
    ''' </remarks>
    Public Overloads Sub DeleteFile(ByVal File As ftpFile)
        Call DeleteFile(File.Name)
    End Sub

    ''' <summary>
    ''' Downloads a file from the FTP server.
    ''' </summary>
    ''' <param name="RemoteFileName">The name of the file to be downloaded.</param>
    ''' <param name="LocalFileName">The full path of the local file where the file will be saved.</param>
    Public Overloads Sub DownloadFile(ByVal RemoteFileName As String, Optional ByVal LocalFileName As String = "")
        If (String.IsNullOrEmpty(LocalFileName)) Then
            LocalFileName = RemoteFileName.Substring(RemoteFileName.LastIndexOf("/") + 1)
        End If

        Dim iosLocal As New System.IO.FileStream(LocalFileName, IO.FileMode.OpenOrCreate, IO.FileAccess.Write, IO.FileShare.None, 1024, IO.FileOptions.SequentialScan)
        Call DownloadFile(RemoteFileName, iosLocal)

        '*
        '* Closes the local file
        '*
        iosLocal.Close()
    End Sub

    ''' <summary>
    ''' Downloads a file from the FTP server.
    ''' </summary>
    ''' <param name="RemoteFileName">The name of the file to be downloaded.</param>
    ''' <param name="LocalStream">A System.IO.Stream that will hold the data from the file.</param>
    Public Overloads Sub DownloadFile(ByVal RemoteFileName As String, ByVal LocalStream As System.IO.Stream)
        If (__Busy) Then
            Throw New InvalidOperationException("Cannot download a file while an operation is still in progress.")
        End If

        __Busy = True

        __LocalStream = LocalStream
        Call ExecCommand(System.Net.WebRequestMethods.Ftp.DownloadFile, RemoteFileName)

    End Sub

    ''' <summary>
    ''' Downloads a file from the FTP server.
    ''' </summary>
    ''' <param name="RemoteFile">An ftpFile object that represents the file to  be downloaded.</param>
    ''' <param name="LocalFileName">The full path of the local file where the file will be saved.</param>
    Public Overloads Sub DownloadFile(ByVal RemoteFile As ftpFile, Optional ByVal LocalFileName As String = "")
        Call DownloadFile(RemoteFile.Name, LocalFileName)
    End Sub

    ''' <summary>
    ''' Downloads a file from the FTP server.
    ''' </summary>
    ''' <param name="RemoteFile">An ftpFile object that represents the file to  be downloaded.</param>
    ''' <param name="LocalStream">A System.IO.Stream that will hold the data from the file.</param>
    Public Overloads Sub DownloadFile(ByVal RemoteFile As ftpFile, ByVal LocalStream As System.IO.Stream)
        Call DownloadFile(RemoteFile.Name, LocalStream)
    End Sub

    ''' <summary>
    ''' Updates the FileList with fresh information from the FTP server.
    ''' </summary>
    Public Sub Refresh()

        If (__Server Is Nothing) Then
            Exit Sub
        End If

        If (__Busy) Then
            Throw New InvalidOperationException("Cannot refresh server while an operation is still in progress.")
        End If

        Call ExecCommand(System.Net.WebRequestMethods.Ftp.ListDirectoryDetails)

    End Sub

    ''' <summary>
    ''' Uploads a stream to the FTP server.
    ''' </summary>
    ''' <param name="LocalStream">A <see cref="T:System.IO.Stream"/> that holds the data to be transfered to the FTP server.</param>
    ''' <param name="RemoteFileName">The name of the file on the FTP server.</param>
    Public Sub UploadStream(ByVal LocalStream As System.IO.Stream, ByVal RemoteFileName As String)
        If (__Busy) Then
            Throw New InvalidOperationException("Cannot upload a file while an operation is still in progress.")
        End If

        __Busy = True

        __LocalStream = LocalStream
        Call initializeFTPRequest(System.Net.WebRequestMethods.Ftp.UploadFile, RemoteFileName)
        If (__Async) Then
            __ftpRequest.BeginGetRequestStream(AddressOf processAsyncRequest, Nothing)
            __OperationComplete.WaitOne()
        Else
            __ftpRequestStream = __ftpRequest.GetRequestStream
            Call FillRequestStream()
            Call GetResponse()
            Call processResponse()
        End If

    End Sub

    ''' <summary>
    ''' Uploads a file to the FTP server.
    ''' </summary>
    ''' <param name="LocalFileName">The full path of the local file.</param>
    ''' <param name="RemoteFileName">The name of the file on the FTP server. If not informed it is assumed the name of the local file.</param>
    Public Sub UploadFile(ByVal LocalFileName As String, Optional ByVal RemoteFileName As String = "")
        If (__Busy) Then
            Throw New InvalidOperationException("Cannot upload a file while an operation is still in progress.")
        End If

        If (RemoteFileName = "") Then
            RemoteFileName = LocalFileName.Substring(LocalFileName.LastIndexOf("\") + 1)
        End If

        Dim iosLocal As New System.IO.FileStream(LocalFileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read, 1024, IO.FileOptions.SequentialScan)
        Call UploadStream(iosLocal, RemoteFileName)
        iosLocal.Close()
    End Sub

#End Region

#Region " Private Functions "

    ''' <summary>
    ''' Initializes the underlieing FTP connection.
    ''' </summary>
    ''' <param name="cmd">The command to be executed.</param>
    ''' <param name="FileName">An optional file name on which the command will be executed.</param>
    Private Sub initializeFTPRequest(ByVal cmd As String, Optional ByVal FileName As String = "")

        If (__Async) Then
            RaiseEvent onBeforeRequest(Me, EventArgs.Empty)
        End If

        If (__Path = "/") Then
            __ftpRequest = System.Net.WebRequest.Create(__Server.ToString & FileName)
        ElseIf (__Path.EndsWith("/")) Then
            __ftpRequest = System.Net.WebRequest.Create(__Server.ToString & __Path & FileName)
        Else
            __ftpRequest = System.Net.WebRequest.Create(__Server.ToString & __Path & "/" & FileName)
        End If

        __ftpRequest.Credentials = __UserCredentials
        __ftpRequest.EnableSsl = __UseSSL
        __ftpRequest.Method = cmd

    End Sub

    ''' <summary>
    ''' Executes an FTP command on the server.
    ''' </summary>
    ''' <param name="cmd">The command to be executed.</param>
    ''' <param name="FileName">An optional file name which will be used while processing the command.</param>
    Private Sub ExecCommand(ByVal cmd As String, Optional ByVal FileName As String = "")

        If (__Server Is Nothing) Then
            Throw New InvalidOperationException("Server not configured.")
        End If

        __Busy = True
        Call initializeFTPRequest(cmd, FileName)
        Call GetResponse()

    End Sub

    ''' <summary>
    ''' When uploading a file, fills the Request Stream with the data from the local file.
    ''' </summary>
    Private Sub FillRequestStream()

        Try
            Dim byteCount As Integer = 0
            Dim readBytes As Integer = 0
            Dim Buffer(__BufferSize - 1) As Byte

            Do
                readBytes = __LocalStream.Read(Buffer, 0, __BufferSize)
                __ftpRequestStream.Write(Buffer, 0, readBytes)
                byteCount += readBytes
                If (__Async) Then
                    RaiseEvent onRequestProgress(Me, New RequestProgressEventArgs(byteCount, 0, __LocalStream.Length))
                End If
            Loop While (readBytes <> 0)
            __ftpRequestStream.Close()
        Catch ex As Exception
            Throw
        Finally
            Call GetResponse()
        End Try

    End Sub

    ''' <summary>
    ''' Retrieves information from the undelining FTP connection.
    ''' </summary>
    Private Sub GetResponse()

        Try
            If (__Async) Then
                __ftpRequest.BeginGetResponse(AddressOf processAsyncResponse, Nothing)
                __OperationComplete.WaitOne()
            Else
                __ftpResponse = __ftpRequest.GetResponse()
                Call processResponse()
            End If
        Catch ex As Exception
            Throw
        End Try

    End Sub

    ''' <summary>
    ''' When accessing the FTP server asynchronously process every request to the server.
    ''' </summary>
    Private Sub processAsyncRequest(ByVal arState As IAsyncResult)
        '*
        '* Prepare the Request
        '*
        __ftpRequestStream = __ftpRequest.EndGetRequestStream(arState)
        Call FillRequestStream()
        ' Call GetResponse()
    End Sub

    ''' <summary>
    ''' When accessing the FTP server asynchronously process every response from the server.
    ''' </summary>
    Private Sub processAsyncResponse(ByVal arState As IAsyncResult)
        '*
        '* Read the Response
        '*
        __ftpResponse = __ftpRequest.EndGetResponse(arState)
        Call processResponse()
    End Sub

    ''' <summary>
    ''' Parses and decoded the response from the FTP server.
    ''' </summary>
    Private Sub processResponse()

        '*
        '* Convert the Response to a byte array
        '*
        Dim ios As System.IO.Stream = __ftpResponse.GetResponseStream()
        Dim ms As New System.IO.MemoryStream
        Dim Buffer(__BufferSize - 1) As Byte
        Dim abData As Byte()
        Dim bytesRead As Integer = 0
        Dim byteCount As Integer = 0

        Do
            bytesRead = ios.Read(Buffer, 0, __BufferSize)
            If (__ftpRequest.Method = System.Net.WebRequestMethods.Ftp.DownloadFile) Then
                '*
                '* Save the data to local file
                '*
                __LocalStream.Write(Buffer, 0, bytesRead)
            Else
                '*
                '* Save the data for memory processing
                '*
                ms.Write(Buffer, 0, bytesRead)
            End If

            byteCount += bytesRead
            If (__Async) Then
                RaiseEvent onRequestProgress(Me, New RequestProgressEventArgs(0, byteCount, ios.Length))
            End If

            '*
            '* Sleep for 10ms to give the processor a chance to breathe
            '*
            System.Threading.Thread.Sleep(10)
        Loop While (bytesRead <> 0)

        abData = ms.ToArray()
        ms = Nothing

        Select Case __ftpRequest.Method
            Case System.Net.WebRequestMethods.Ftp.ListDirectoryDetails
                '*
                '* split the data into lines
                '*
                Dim strDir As String() = System.Text.Encoding.UTF8.GetString(abData).Replace(vbLf, "").Split(vbCr)

                '*
                '* process each file
                '*
                __Files = New ftpDirectory
                For Each strFile As String In strDir
                    '*
                    '* THE most amazing bug I ever encountered
                    '*
                    '* Looks like that if the last string after a split is empty
                    '* the split doesn't return an empty string. Instead it returns
                    '* a string which contains a "character" Nothing in its first
                    '* position. That's the only way I found to catch it.
                    '*
                    If (strFile = "") OrElse (strFile.Chars(0) = Nothing) Then
                        Exit For
                    End If

                    '*
                    '* Try to guess the server type
                    '*
                    If (__HostType = HostType.Unknown) Then
                        __HostType = getHostType(strFile)
                        If (__HostType = HostType.Unknown) Then
                            If (__AllowUserInput) Then
                                Dim d As New frmHelpFTPLibrary
                                d.PrepareDialog(strFile)
                                d.ShowDialog()
                            End If
                            Throw New InvalidOperationException("Unknown FTP Server")
                        End If
                    End If

                    '*
                    '* Decode the file info
                    '*
                    Dim oFile As New ftpFile

                    With __Parser.Match(strFile)
                        '*
                        '* Basic Information
                        '*
                        oFile.Name = .Groups("name").Value
                        oFile.Size = CULng("0" & .Groups("size").Value)
                        Try
                            oFile.Date = CDate(.Groups("timestamp").Value)
                        Catch ex As Exception
                        End Try

                        oFile.Owner = IIf(.Groups("owner") Is Nothing, "", .Groups("owner").Value)

                        '*
                        '* Is a directory
                        '*
                        oFile.Flags = AttributeFlags.None
                        oFile.Flags = oFile.Flags Or IIf(.Groups("dir").Success And .Groups("dir").Value <> "-", AttributeFlags.Directory, AttributeFlags.None)

                        '*
                        '* File Permission Attributes
                        '*
                        If (.Groups("permOwner").Success) Then
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permOwner").Value.ToLower.Chars(0) = "r", AttributeFlags.OwnerRead, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permOwner").Value.ToLower.Chars(1) = "w", AttributeFlags.OwnerWrite, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permOwner").Value.ToLower.Chars(2) = "x", AttributeFlags.OwnerExecute, AttributeFlags.None)
                        End If

                        If (.Groups("permGroup").Success) Then
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permGroup").Value.ToLower.Chars(0) = "r", AttributeFlags.GroupRead, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permGroup").Value.ToLower.Chars(1) = "w", AttributeFlags.GroupWrite, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permGroup").Value.ToLower.Chars(2) = "x", AttributeFlags.GroupExecute, AttributeFlags.None)
                        End If

                        If (.Groups("permPublic").Success) Then
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permPublic").Value.ToLower.Chars(0) = "r", AttributeFlags.PublicRead, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permPublic").Value.ToLower.Chars(1) = "w", AttributeFlags.PublicWrite, AttributeFlags.None)
                            oFile.Flags = oFile.Flags Or IIf(.Groups("permPublic").Value.ToLower.Chars(2) = "x", AttributeFlags.PublicExecute, AttributeFlags.None)
                        End If
                    End With

                    __Files.Add(oFile)
                Next
        End Select

        If (__Path = "") Then
            __Path = __Server.AbsolutePath
        End If

        __Busy = False
        If (__Async) Then
            RaiseEvent onRequestComplete(Me, EventArgs.Empty)
            __OperationComplete.Set()
        End If

    End Sub

    ''' <summary>
    ''' Using  a directory entry from the server tries to discover the type of Host (Windows, Unix, etc).
    ''' </summary>
    ''' <param name="fileInfo">The directory entry that will be matched against known server replies.</param>
    Private Function getHostType(ByVal fileInfo As String) As HostType

        Dim r As System.Text.RegularExpressions.Regex

        For i As Integer = 0 To __Parsers.GetUpperBound(0)
            r = New System.Text.RegularExpressions.Regex(__Parsers(i, 1))
            If (r.Match(fileInfo).Success) Then
                __Parser = r
                Return __Parsers(i, 0)
            End If
        Next

        Return HostType.Unknown

    End Function

#End Region

End Class
