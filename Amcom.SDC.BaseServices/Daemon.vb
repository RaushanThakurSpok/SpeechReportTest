Imports System
Imports System.Collections
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Authentication
Imports System.IO
Imports System.Net.Security
Namespace TCP
    Public Class Daemon
        ReadOnly mPort As Integer
        ReadOnly mCapacity As Integer
        Private mListener As Socket = Nothing
        Private mClientConnections As New Dictionary(Of String, ClientConnection)
        Private mRunning As Boolean = False
        Private mStopping As Boolean = False
        Private _sslSelfSignPassword As String
        Private mBindAddress As String
        Public Event DataReceived As EventHandler(Of TCPDataReceivedEventArgs)
        Public Event Connected As EventHandler(Of TCPConnectionStateEventArgs)
        Public Event Disconnected As EventHandler(Of TCPConnectionStateEventArgs)
        Public Event ErrorOccurred As EventHandler(Of TCPErrorEventArgs)
        Public Event Stopped As EventHandler
        Private _sslProtocols As SslProtocols = SslProtocols.Default
        Public Enum DaemonTransport As Integer
            ''' <summary>
            ''' Non-secure standard TCP
            ''' </summary>
            ''' <remarks>Non-secure standard TCP</remarks>
            tcp
            ''' <summary>
            ''' SSL via Certificate Chain
            ''' </summary>
            ''' <remarks>SSL via Certificate Chain</remarks>
            ssl
            ''' <summary>
            ''' SSL via Self-Signed Certificates (P2P)
            ''' </summary>
            ''' <remarks>SSL via Self-Signed Certificates (P2P)</remarks>
            sslSelfSignedP2P
        End Enum
        Private _transport As DaemonTransport
        Private _certFile As String = String.Empty
        Private _certX509 As X509Certificate2
        Private _sslListener As TcpListener
        Private _selfsignedCert As Boolean
#Region "Constructor"
        ''' <summary>
        ''' Constructor for non-secure standard TCP
        ''' </summary>
        ''' <param name="port">Port to listen on</param>
        ''' <param name="capacity">Number of possible connections</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal port As Integer, ByVal capacity As Integer)
            mPort = port
            mCapacity = capacity
            mBindAddress = ""
            _selfsignedCert = False
            _transport = DaemonTransport.tcp
        End Sub

        ''' <summary>
        ''' Constructor for an SSL transport using a self-signed certificate.  The certificate will be generated for the server the code is running on.
        ''' </summary>
        ''' <param name="port">Port to listen on</param>
        ''' <param name="capacity">Number of possible connections</param>
        ''' <param name="sslSelfSignStartDate">Start date for self-signed certificate</param>
        ''' <param name="sslSelfSignEndDate">End date for self-signed certificate</param>
        ''' <param name="sslSelfSignPassword">Password for the certificate</param>
        ''' <remarks>The certificate will be save to the file system as [InstallPath]\Files\[AppName]-sslcert.pfx</remarks>
        Public Sub New(ByVal port As Integer, ByVal capacity As Integer, ByVal serverName As String, ByVal sslSelfSignStartDate As Date, ByVal sslSelfSignEndDate As Date, ByVal sslSelfSignPassword As String)
            mPort = port
            mCapacity = capacity
            mBindAddress = ""
            _sslSelfSignPassword = sslSelfSignPassword
            _transport = DaemonTransport.sslSelfSignedP2P
            If String.IsNullOrEmpty(serverName) Then serverName = My.Computer.Name
            'in this constructor, we will build a self-signed certificate.
            _selfsignedCert = True
            'certificate is valid for 30 days.
            Dim c As Byte() = Nothing
            Try
                c = Certificate.CreateSelfSignCertificatePfx("CN=" & serverName, sslSelfSignStartDate, sslSelfSignEndDate, sslSelfSignPassword)
            Catch EX As Exception
            End Try
            _certFile = Path.Combine(App.FilePath, App.Name & "-" & "sslcert.pfx")
            Using binWriter As New BinaryWriter(File.Open(_certFile, FileMode.Create))
                binWriter.Write(c)
            End Using


        End Sub
        ''' <summary>
        ''' Constructor for an SSL transport using a supplied certificate file. 
        ''' </summary>
        ''' <param name="port">Port the daemon will listen on</param>
        ''' <param name="capacity">total possible number of connections</param>
        ''' <param name="certificateFile">Path to the certificate file</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal port As Integer, ByVal capacity As Integer, ByVal certificateFile As String)
            mPort = port
            mCapacity = capacity
            mBindAddress = ""
            _transport = DaemonTransport.ssl
            _certFile = certificateFile
            _selfsignedCert = False
        End Sub
        ''' <summary>
        ''' Constructor for an SSL transport using a supplied X509 certificate. 
        ''' </summary>
        ''' <param name="port">Port the daemon will listen on</param>
        ''' <param name="capacity">total possible number of connections</param>
        ''' <param name="certificate">X509 Certificate</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal port As Integer, ByVal capacity As Integer, ByVal certificate As X509Certificate2)
            mPort = port
            mCapacity = capacity
            mBindAddress = ""
            _transport = DaemonTransport.ssl
            _certFile = ""
            _certX509 = certificate
            _selfsignedCert = False
        End Sub
#End Region
#Region "Properties"
        Public Property SslProtocol() As SslProtocols
            Get
                Return _sslProtocols
            End Get
            Set(ByVal value As SslProtocols)
                _sslProtocols = value
            End Set
        End Property
        Public ReadOnly Property Transport() As DaemonTransport
            Get
                Return _transport
            End Get
        End Property
        Public Property BindAddress() As String
            Get
                Return mBindAddress
            End Get
            Set(ByVal value As String)
                mBindAddress = value
            End Set
        End Property
        Public ReadOnly Property ClientConnections() As Dictionary(Of String, ClientConnection)
            Get
                Return mClientConnections
            End Get
        End Property
#End Region
#Region "Methods"
        Public Sub CloseConnection(ByVal connectid As String)
            mClientConnections(connectid).Socket.Close()
            mClientConnections.Remove(connectid)
        End Sub
        Private Sub RaiseDisconnectedEvent(ByVal connection As ClientConnection)
            If connection IsNot Nothing Then
                mClientConnections.Remove(connection.EndPoint.ToString)
            End If
            If connection IsNot Nothing Then
                RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(connection.EndPoint))
            End If
        End Sub
        Private Sub RaiseErrorEvent(ByVal connection As ClientConnection, ByVal [error] As SocketException)
            If (connection IsNot Nothing) Then
                RaiseEvent ErrorOccurred(Me, New TCPErrorEventArgs(connection.EndPoint, [error]))
            End If
        End Sub

        Private Sub HandleConnectionData(ByVal connection As ClientConnection, ByVal parameter As IAsyncResult)
            Dim read As Integer = 0
            If connection.SslStream IsNot Nothing Then
                read = connection.SslStream.EndRead(parameter)
            Else
                read = connection.Socket.EndReceive(parameter)
            End If
            If read = 0 Then
                RaiseDisconnectedEvent(connection)
            Else
                Dim received As Byte() = New Byte(read - 1) {}
                Array.Copy(connection.Buffer, 0, received, 0, read)
                RaiseEvent DataReceived(Me, New TCPDataReceivedEventArgs(connection.EndPoint, received))
                StartWaitingForData(connection)
            End If
        End Sub
        Private Sub HandleIncomingData(ByVal parameter As IAsyncResult)
            Dim connection As ClientConnection = DirectCast(parameter.AsyncState, ClientConnection)
            Try
                HandleConnectionData(connection, parameter)
            Catch ex As ObjectDisposedException
                RaiseDisconnectedEvent(connection)
            Catch ex As SocketException
                mClientConnections.Remove(connection.EndPoint.ToString)
                RaiseErrorEvent(connection, ex)
                RaiseDisconnectedEvent(connection)
            End Try
        End Sub
        Private Sub StartWaitingForData(ByVal connection As ClientConnection)
            If connection.SslStream IsNot Nothing Then
                connection.SslStream.BeginRead(connection.Buffer, 0, connection.Buffer.Length, New AsyncCallback(AddressOf HandleIncomingData), connection)
            Else
                connection.Socket.BeginReceive(connection.Buffer, 0, connection.Buffer.Length, SocketFlags.None, New AsyncCallback(AddressOf HandleIncomingData), connection)
            End If
        End Sub
        Private Function ProcessValidation(ByVal client As TcpClient) As ClientConnection
            'A client has connected. Create the 
            'SslStream using the client's network stream.
            Dim sslStream As New SslStream(client.GetStream(), True)
            'Authenticate the server but don't require the client to authenticate.
            Try
                Dim serverCertificate As X509Certificate

                If _certFile.Length > 0 Then
                    serverCertificate = New X509Certificate2(_certFile, _sslSelfSignPassword)
                Else
                    serverCertificate = _certX509
                End If

                sslStream.AuthenticateAsServer(serverCertificate, False, _sslProtocols, True)
                'Set timeouts for the read and write to 5 seconds.
                sslStream.ReadTimeout = 5000
                sslStream.WriteTimeout = 5000
                Dim result As ClientConnection = New ClientConnection(client, sslStream)
                If mClientConnections.ContainsKey(result.EndPoint.ToString) Then
                    result.Socket.Close()
                    Return Nothing
                End If
                Return result
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Private Function CreateClientConnection(ByVal parameter As IAsyncResult) As ClientConnection
            Dim result As ClientConnection = Nothing
            Try
                If _transport = DaemonTransport.ssl OrElse _transport = DaemonTransport.sslSelfSignedP2P Then
                    Dim client As TcpClient = _sslListener.EndAcceptTcpClient(parameter)
                    Dim temp As ClientConnection = ProcessValidation(client)
                    If temp IsNot Nothing Then
                        result = temp
                    Else
                        client.Close()
                        Return Nothing
                    End If
                Else
                    Dim socket As Socket = mListener.EndAccept(parameter)
                    socket.LingerState = (New System.Net.Sockets.LingerOption(False, 0))
                    result = New ClientConnection(socket)
                End If
            Catch ex As AmcomException
            Catch ex As SocketException
                App.ExceptionLog(New AmcomException("Failed creating socket connection"))
            End Try
            If mClientConnections.ContainsKey(result.EndPoint.ToString) Then
                result.Socket.Close()
                Return Nothing
            End If
            mClientConnections.Add(result.EndPoint.ToString, result)
            RaiseEvent Connected(Me, New TCPConnectionStateEventArgs(result.EndPoint))
            StartWaitingForData(result)
            Return result
        End Function
        Private Sub HandleSslConnection(ByVal parameter As IAsyncResult)
            Dim connection As ClientConnection = Nothing

            If mStopping Then
                RaiseDisconnectedEvent(connection)
                Return
            End If

            Try
                connection = CreateClientConnection(parameter)
            Catch generatedExceptionName As ObjectDisposedException
                RaiseDisconnectedEvent(connection)
            Catch ex As SocketException
                If (Not connection Is Nothing) Then
                    mClientConnections.Remove(connection.EndPoint.ToString)
                    RaiseErrorEvent(connection, ex)
                    RaiseDisconnectedEvent(connection)
                End If
            Catch ex As Exception
                App.ExceptionLog(New AmcomException("Unexpected TCP failure", ex))
                If (connection IsNot Nothing) Then
                    mClientConnections.Remove(connection.EndPoint.ToString)
                    RaiseErrorEvent(connection, ex)
                    RaiseDisconnectedEvent(connection)
                End If
            End Try
            WaitForSslClient()
        End Sub
        Private Sub HandleConnection(ByVal parameter As IAsyncResult)
            Dim connection As ClientConnection = Nothing
            Try
                connection = CreateClientConnection(parameter)
            Catch generatedExceptionName As ObjectDisposedException
                RaiseDisconnectedEvent(connection)
            Catch ex As SocketException
                If (Not connection Is Nothing) Then
                    mClientConnections.Remove(connection.EndPoint.ToString)
                    RaiseErrorEvent(connection, ex)
                    RaiseDisconnectedEvent(connection)
                End If
            Catch ex As Exception
                App.ExceptionLog(New AmcomException("Unexpected TCP failure", ex))
                If (connection IsNot Nothing) Then
                    mClientConnections.Remove(connection.EndPoint.ToString)
                    RaiseErrorEvent(connection, ex)
                    RaiseDisconnectedEvent(connection)
                End If
            End Try
            WaitForClient()
        End Sub
        Private Sub WaitForSslClient()
            Try
                If Not mStopping Then _sslListener.BeginAcceptTcpClient(New AsyncCallback(AddressOf HandleSslConnection), Nothing)
            Catch ex As ObjectDisposedException
                'dont care
            End Try
        End Sub
        Private Sub WaitForClient()
            Try
                mListener.BeginAccept(New AsyncCallback(AddressOf HandleConnection), Nothing)
            Catch ex As ObjectDisposedException
                'dont care
            End Try
        End Sub
        Private Sub HandleSendFinished(ByVal parameter As IAsyncResult)
            Dim connection As ClientConnection = DirectCast(parameter.AsyncState, ClientConnection)
            Try
                If connection.SslStream IsNot Nothing Then
                    connection.SslStream.EndWrite(parameter)
                Else
                    connection.Socket.EndSend(parameter)
                End If
            Catch ex As SocketException
                mClientConnections.Remove(connection.EndPoint.ToString)

                RaiseErrorEvent(connection, ex)
                RaiseDisconnectedEvent(connection)
            End Try
        End Sub
        Public Sub Start()
            mStopping = False
            If _transport = DaemonTransport.ssl OrElse _transport = DaemonTransport.sslSelfSignedP2P Then

                'Create a TCP/IP (IPv4) socket and listen for incoming connections.
                If mBindAddress.Length > 0 Then
                    _sslListener = New TcpListener(IPAddress.Parse(mBindAddress), mPort)
                Else
                    _sslListener = New TcpListener(IPAddress.Any, mPort)
                End If
                _sslListener.Start()
                WaitForSslClient()
                Exit Sub
            End If
            Try
                mListener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                If mBindAddress.Length > 0 Then
                    mListener.Bind(New IPEndPoint(IPAddress.Parse(mBindAddress), mPort))
                Else
                    mListener.Bind(New IPEndPoint(IPAddress.Any, mPort))
                End If
            Catch ex As IO.IOException
                Throw New AmcomException("failed to start TCP Daemon", ex)
            Catch ex As ArgumentException
                Throw New AmcomException("failed to start TCP Daemon", ex)
            Catch ex As SocketException
                Throw New AmcomException("Address is in use.  Is the service already running?")
            End Try
            mRunning = True
            mListener.Listen(mCapacity)
            WaitForClient()
        End Sub
        Public Sub [Stop]()
            If _sslListener IsNot Nothing Then
                mStopping = True
                _sslListener.Stop()
                For Each client As ClientConnection In mClientConnections.Values
                    client.SslStream.Close()
                Next
            Else
                mStopping = True
                mListener.Close()
                For Each client As ClientConnection In mClientConnections.Values
                    client.Socket.Close()
                Next
            End If
            mClientConnections.Clear()
            RaiseEvent Stopped(Me, New System.EventArgs)
            mRunning = False
            mStopping = False
        End Sub
        Public Sub Send(ByVal client As IPEndPoint, ByVal buffer As String)
            Dim ascii As New ASCIIEncoding
            Dim b() As Byte = ascii.GetBytes(buffer)
            Send(client, b)
        End Sub
        Public Sub Send(ByVal client As String, ByVal buffer As String)
            For Each cc As ClientConnection In mClientConnections.Values
                If cc IsNot Nothing Then
                    If cc.EndPoint.ToString = client Then
                        Send(cc.EndPoint, buffer)
                        Exit Sub
                    End If
                End If
            Next
        End Sub
        Public Sub Send(ByVal client As String, ByVal buffer() As Byte)
            For Each cc As ClientConnection In mClientConnections.Values
                If cc IsNot Nothing Then
                    If cc.EndPoint.ToString = client Then
                        Send(cc.EndPoint, buffer)
                        Exit Sub
                    End If
                End If
            Next
        End Sub
        Public Sub Send(ByVal client As IPEndPoint, ByVal buffer As Byte())
            If Not mClientConnections.ContainsKey(client.ToString) Then
                Throw New ArgumentException("Client not connected.", "client")
            End If
            Dim connection As ClientConnection = mClientConnections(client.ToString)
            Try
                If connection.SslStream IsNot Nothing Then
                    connection.SslStream.BeginWrite(buffer, 0, buffer.Length, New AsyncCallback(AddressOf HandleSendFinished), connection)
                Else
                    connection.Socket.BeginSend(buffer, 0, buffer.Length, SocketFlags.None, New AsyncCallback(AddressOf HandleSendFinished), connection)
                End If
            Catch generatedExceptionName As ObjectDisposedException
                RaiseDisconnectedEvent(connection)
            Catch ex As SocketException
                mClientConnections.Remove(client.ToString)
                RaiseErrorEvent(connection, ex)
                RaiseDisconnectedEvent(connection)

            End Try
        End Sub


#End Region


        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
End Namespace

