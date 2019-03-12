Imports System
Imports System.Collections
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.Net.Security
Imports System.Security.Authentication
Imports System.Security.Cryptography.X509Certificates
Imports System.IO

Namespace TCP
    Public Class Client
        Implements IDisposable

        Private mOwner As Object
        Private mServer As String
        Private mPort As Integer
        Private mClient As TcpClient
        Private mStream As NetworkStream
        Private mBuffer As Byte()
        Private mSendBuffer As Byte()
        Private mSend As Object = Nothing
        Private mReceive As Object = Nothing
        Private disposedValue As Boolean = False        ' To detect redundant calls
        Private mConnected As Boolean
        Private _certificateErrors As Hashtable = New Hashtable()
        Private _transport As TransportType = TransportType.tcp
        Private _sslServerName As String
        Private _sslPort = 443
        Private _sslStream As SslStream
        Private _ignoreSSLChainErrors As Boolean = False
        Private _authenticateClient As Boolean = False
        Private _clientCertificateFile As String = String.Empty
        Private _clientCertificateX509 As X509Certificate
        Private _sslProtocol As SslProtocols = SslProtocols.Default
        Public Enum TransportType
            tcp
            ssl
        End Enum
#Region "Constructor"
        ''' <summary>
        ''' Create a standard, Non-secure TCP Client
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Initialize()
        End Sub
        ''' <summary>
        ''' Create a standard, Non-secure TCP Client
        ''' </summary>
        ''' <param name="uiObject">User interface object associated with this connection</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal uiObject As Object)
            mOwner = uiObject
            Initialize()
        End Sub
        ''' <summary>
        ''' Creates an SSL Client
        ''' </summary>
        ''' <param name="sslServerName">Name which will be on the server certificate</param>
        ''' <param name="sslPort">Port server is running on</param>
        ''' <remarks>Server Validation Only</remarks>
        Public Sub New(ByVal sslServerName As String, ByVal sslPort As Integer)
            _transport = TransportType.ssl
            _sslServerName = sslServerName
            _sslPort = sslPort
            Initialize()
        End Sub
        ''' <summary>
        ''' Creates an SSL Client
        ''' </summary>
        ''' <param name="sslServerName">Name which will be on the server certificate</param>
        ''' <param name="sslPort">Port server is running on</param>
        ''' <param name="clientCertificateFile">Path to the client certificate.</param>
        ''' <remarks>Server and Client Validation</remarks>
        Public Sub New(ByVal sslServerName As String, ByVal sslPort As Integer, ByVal clientCertificateFile As String)
            _transport = TransportType.ssl
            _sslServerName = sslServerName
            _sslPort = sslPort
            _clientCertificateFile = clientCertificateFile
            _authenticateClient = True
            Initialize()
        End Sub
        ''' <summary>
        ''' Creates an SSL Client
        ''' </summary>
        ''' <param name="sslServerName">Name which will be on the server certificate</param>
        ''' <param name="sslPort">Port server is running on</param>
        ''' <param name="clientCertificate">Path to the client certificate.</param>
        ''' <remarks>Server and Client Validation</remarks>
        Public Sub New(ByVal sslServerName As String, ByVal sslPort As Integer, ByVal clientCertificate As X509Certificate)
            _transport = TransportType.ssl
            _sslServerName = sslServerName
            _sslPort = sslPort
            _clientCertificateFile = ""
            _clientCertificateX509 = clientCertificate
            _authenticateClient = True
            Initialize()
        End Sub
#End Region
#Region "Properties"
        Public Property SslProtocol() As SslProtocols
            Get
                Return _sslProtocol
            End Get
            Set(ByVal value As SslProtocols)
                _sslProtocol = value
            End Set
        End Property
        Public Property IgnoreSSLCertificateChainErrors() As Boolean
            Get
                Return _ignoreSSLChainErrors
            End Get
            Set(ByVal value As Boolean)
                _ignoreSSLChainErrors = value
            End Set
        End Property
        Public ReadOnly Property Transport() As TransportType
            Get
                Return _transport
            End Get
        End Property
        Public ReadOnly Property IsConnected() As Boolean
            Get
                Return mConnected
            End Get
        End Property
        Public Property ServerAddress() As String
            Get
                Return mServer
            End Get
            Set(ByVal value As String)
                mServer = value
            End Set
        End Property
        Public Property ServerPort() As Integer
            Get
                Return mPort
            End Get
            Set(ByVal value As Integer)
                mPort = value
            End Set
        End Property
        Public Property UIDelegate() As Object
            Get
                Return mOwner
            End Get
            Set(ByVal value As Object)
                mOwner = value
            End Set
        End Property
#End Region
#Region "Delegates"
        Private Delegate Sub RaiseDataReceivedEvent(ByVal sender As Object, ByVal e As ClientDataReceivedEventArgs)
        Private Delegate Sub RaiseConnectedEvent(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
        Private Delegate Sub RaiseDisconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
#End Region
#Region "Event Definitions"
        Public Event DataReceived(ByVal sender As Object, ByVal e As ClientDataReceivedEventArgs)
        Public Event Connected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
        Public Event Disconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
#End Region
#Region "Methods"
        Public Sub Connect()
            If mServer.Length = 0 Then
                Throw New ArgumentException("No server address has been supplied")
            End If
            If _transport = TransportType.ssl Then
                If _sslPort = 0 Then
                    Throw New ArgumentException("No server port has been specified")
                End If
            Else
                If mPort = 0 Then
                    Throw New ArgumentException("No server port has been specified")
                End If
            End If
            If mConnected Then
                Throw (New InvalidOperationException("already connected to server"))
            End If
            If _transport = TransportType.ssl Then
                If _sslServerName.Length = 0 Then
                    Throw New InvalidOperationException("No SSL ServerName supplied")
                End If

            End If
            If _transport = TransportType.ssl Then
                ConnectSSL(mServer, _sslServerName)
            Else
                Connect(mServer, mPort)
            End If
        End Sub
        Public Sub Connect(ByVal server As String, ByVal port As String)

            If mConnected Then
                Throw (New AmcomException("Server already connected", New InvalidOperationException("already connected to server")))
            End If
            If _transport = TransportType.ssl Then
                mServer = server
                _sslPort = port
                ConnectSSL(mServer, _sslServerName)
                Exit Sub
            End If
            Do
                Try
                    mClient.Connect(server, port)
                    mServer = server
                    mPort = port
                    Exit Do
                Catch ex As ObjectDisposedException
                    mClient = New TcpClient
                Catch ex As SocketException
                    If ex.ErrorCode = 10056 Then
                        mClient.Close()
                        mClient = New TcpClient
                    Else
                        Throw New AmcomException("Failed to connect to " & server & ", port " & port & ": " & ex.Message, ex)
                    End If
                End Try
            Loop
            mConnected = True
            GenerateConnectEvent()
            WaitForRead()
        End Sub
        Public Sub ConnectAsync()
            If mConnected Then
                Throw (New AmcomException("Server already connected", New InvalidOperationException("already connected to server")))
            End If
            If mServer.Length = 0 Then
                Throw New ArgumentException("No server address has been supplied")
            End If
            If mPort = 0 Then
                Throw New ArgumentException("No server port has been specified")
            End If
            If _transport = TransportType.ssl Then
                ConnectSSL(mServer, _sslServerName)
                Exit Sub
            End If
            ConnectAsync(mServer, mPort)

        End Sub
        Public Sub ConnectAsync(ByVal server As String, ByVal port As String)
            If _transport = TransportType.ssl Then
                mServer = server
                _sslPort = port
                ConnectSSL(mServer, _sslServerName)
                Exit Sub
            End If
            Try
                mClient.BeginConnect(server, port, AddressOf HandleTCPConnect, server)
            Catch ex As SocketException
                Throw New AmcomException("Failed to connect to " & server & ", port " & port, ex)
            End Try
        End Sub
        Public Sub Disconnect()
            Try
                If _sslStream IsNot Nothing Then
                    _sslStream.Close()
                    mClient.Close()
                Else
                    mClient.Client.Shutdown(SocketShutdown.Both)
                    mClient.Close()
                End If
            Catch ex As Exception
            End Try
            mConnected = False
            RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(Nothing))
        End Sub
        Private Sub GenerateConnectEvent()
            If mOwner IsNot Nothing Then
                Dim d As New RaiseConnectedEvent(AddressOf OnConnected)
                mOwner.Invoke(d, Me, New TCPConnectionStateEventArgs(mClient.Client.RemoteEndPoint))
            Else
                RaiseEvent Connected(Me, New TCPConnectionStateEventArgs(mClient.Client.RemoteEndPoint))
            End If

        End Sub
        Private Sub GenerateDisconnectEvent()
            GenerateDisconnectEvent("")
        End Sub
        Private Sub GenerateDisconnectEvent(ByVal reason As String)
            '   If Not mConnected Then Exit Sub
            'make sure client is disconnected.
            If mOwner IsNot Nothing Then
                Dim d As New RaiseDisconnected(AddressOf OnDisconnected)
                mOwner.Invoke(d, Me, New TCPConnectionStateEventArgs(mClient.Client.RemoteEndPoint))
            Else
                Try
                    RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(mClient.Client.RemoteEndPoint, reason))
                    mClient.Client.Disconnect(True)
                Catch ex As NullReferenceException
                    RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(Nothing, reason))
                Catch ex As SocketException
                    RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(Nothing, reason))
                Catch ex As ObjectDisposedException
                    RaiseEvent Disconnected(Me, New TCPConnectionStateEventArgs(Nothing, reason))
                End Try
                mConnected = False
            End If
        End Sub
        Private Sub HandleTCPConnect(ByVal ar As System.IAsyncResult)
            Try
                mClient.Client.EndConnect(ar)
            Catch ex As NullReferenceException
                GenerateDisconnectEvent()
                Exit Sub
            Catch ex As SocketException
                GenerateDisconnectEvent(ex.Message)
                Exit Sub
            End Try
            mConnected = True
            GenerateConnectEvent()
            WaitForRead()
        End Sub
        Private Sub HandleTCPRead(ByVal ar As System.IAsyncResult)
            If Not mConnected Then Exit Sub 'A forced Disconnect was called
            Dim numbytes As Integer = 0
            Try
                If _sslStream IsNot Nothing Then
                    numbytes = _sslStream.EndRead(ar)
                Else
                    numbytes = mClient.Client.EndReceive(ar)
                End If
            Catch ex As ArgumentException
                GenerateDisconnectEvent()
                Exit Sub

            Catch ex As SocketException
                If Not mClient.Connected Then
                    GenerateDisconnectEvent()
                    Exit Sub
                End If
            Catch ex As ObjectDisposedException
                GenerateDisconnectEvent()
                Exit Sub

            End Try
            If numbytes > 0 Then
                If ar.IsCompleted Then
                    Dim temp(numbytes) As Byte
                    Array.Copy(mBuffer, 0, temp, 0, numbytes)
                    If mOwner IsNot Nothing Then
                        Dim d As New RaiseDataReceivedEvent(AddressOf OnReceive)
                        mOwner.Invoke(d, Me, New ClientDataReceivedEventArgs(temp))
                    Else
                        RaiseEvent DataReceived(Me, New ClientDataReceivedEventArgs(temp))
                    End If
                End If
            End If
            Try
                If _sslStream IsNot Nothing Then
                    _sslStream.BeginRead(mBuffer, 0, mClient.ReceiveBufferSize, AddressOf HandleTCPRead, mBuffer)
                Else
                    mStream.BeginRead(mBuffer, 0, mClient.ReceiveBufferSize, AddressOf HandleTCPRead, mBuffer)
                End If
            Catch ex As SocketException
                GenerateDisconnectEvent()
            Catch ex As IO.IOException
                GenerateDisconnectEvent()
            Catch ex As ObjectDisposedException
                GenerateDisconnectEvent()
            End Try
        End Sub

        Private Sub HandleTCPWrite(ByVal ar As System.IAsyncResult)
            Try
                If _sslStream IsNot Nothing Then
                    _sslStream.EndWrite(ar)
                Else
                    mClient.Client.EndSend(ar)
                End If
            Catch ex As ArgumentException
                GenerateDisconnectEvent()
            Catch ex As SocketException
                GenerateDisconnectEvent()
            Catch ex As ObjectDisposedException
                GenerateDisconnectEvent()
            End Try
        End Sub

        Private Sub Initialize()
            mServer = ""
            mPort = 0
            mClient = New TcpClient
            ReDim mBuffer(mClient.ReceiveBufferSize)
            ReDim mSendBuffer(mClient.SendBufferSize)
        End Sub

        Public Sub Send(ByVal data As String)
            Dim t As New Text.ASCIIEncoding
            Dim b() As Byte = t.GetBytes(data)
            Do
                Try
                    If _sslStream IsNot Nothing Then
                        _sslStream.BeginWrite(b, 0, b.Length, AddressOf HandleTCPWrite, b)
                    Else
                        mStream.BeginWrite(b, 0, b.Length, AddressOf HandleTCPWrite, b)
                    End If
                    Exit Do
                Catch ex As SocketException
                    Select Case ex.SocketErrorCode
                        Case SocketError.WouldBlock, SocketError.AlreadyInProgress
                            Thread.Sleep(50)
                        Case Else
                            App.ExceptionLog(New AmcomException("SocketException in TCPClient:" & ex.Message, ex))
                            GenerateDisconnectEvent()
                            Exit Do
                    End Select
                Catch ex As IO.IOException
                    App.ExceptionLog(New AmcomException("IOException in TCPClient:" & ex.Message, ex))
                    GenerateDisconnectEvent()
                    Exit Do
                Catch ex As ObjectDisposedException
                    GenerateDisconnectEvent()
                    Exit Do
                End Try
            Loop
        End Sub
        Protected Sub OnReceive(ByVal sender As Object, ByVal e As ClientDataReceivedEventArgs)
            RaiseEvent DataReceived(sender, e)
        End Sub

        Protected Sub OnConnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            RaiseEvent Connected(sender, e)
        End Sub

        Protected Overridable Sub OnDisconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            RaiseEvent Disconnected(sender, e)
        End Sub

        Public Sub Send(ByVal data() As Byte)
            mStream.BeginWrite(mSendBuffer, 0, data.Length, AddressOf HandleTCPWrite, mSendBuffer)
        End Sub

        Private Sub WaitForRead()
            mStream = mClient.GetStream()
            mStream.BeginRead(mBuffer, 0, mClient.ReceiveBufferSize, AddressOf HandleTCPRead, mBuffer)
        End Sub
#End Region

#Region " IDisposable Support "
        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    If mStream IsNot Nothing Then mStream.Dispose()
                End If
            End If
            Me.disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

#Region "SSL Support Functions"
        Private Function ValidateServerCertificate(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal sslPolicyErrors1 As SslPolicyErrors) As Boolean
            If sslPolicyErrors1 = SslPolicyErrors.None Then
                Return True
            Else
                If sslPolicyErrors1 = SslPolicyErrors.RemoteCertificateChainErrors Then
                    'Server is who we think it is, but not from certified authority (a la self-signed)
                    If _ignoreSSLChainErrors Then
                        Return True
                    End If
                End If
            End If
            ' Do not allow this client to communicate with unauthenticated servers.
            Console.WriteLine("Certificate error: {0}", sslPolicyErrors1)
            Return False
        End Function
        Private Sub ConnectSSL(ByVal machineName As String, ByVal serverOnCertName As String)
            ' Create a TCP/IP client socket.
            ' machineName is the host running the server application.
            mClient = New TcpClient(machineName, _sslPort)
            Console.WriteLine("Client connected.")
            ' Create an SSL stream that will close the client's stream.
            _sslStream = New SslStream(mClient.GetStream(), True, New RemoteCertificateValidationCallback(AddressOf ValidateServerCertificate), Nothing)
            Try
                If _authenticateClient Then
                    Dim x509Coll As New X509CertificateCollection
                    If _clientCertificateFile.Length > 0 Then
                        x509Coll.Add(New X509Certificate(_clientCertificateFile))
                    Else
                        x509Coll.Add(_clientCertificateX509)
                    End If

                    ' The server name must match the name on the server certificate.
                    _sslStream.AuthenticateAsClient(serverOnCertName, x509Coll, _sslProtocol, False)
                Else
                    _sslStream.AuthenticateAsClient(serverOnCertName)
                End If


            Catch e As AuthenticationException
                Console.WriteLine("Exception: {0}", e.Message)
                If e.InnerException IsNot Nothing Then
                    Console.WriteLine("Inner exception: {0}", e.InnerException.Message)
                End If
                Throw New AmcomException("Authentication failed")
            End Try

            mConnected = True
            GenerateConnectEvent()
            WaitForSSLRead()

        End Sub

        Private Sub WaitForSSLRead()
            _sslStream.BeginRead(mBuffer, 0, mClient.ReceiveBufferSize, AddressOf HandleTCPRead, mBuffer)
        End Sub

        Private Shared Function ReadMessage(ByVal sslStream As SslStream) As String
            ' Read the  message sent by the server.
            ' The end of the message is signaled using the
            ' "<EOF>" marker.
            Dim buffer As Byte() = New Byte(2047) {}
            Dim messageData As New StringBuilder()
            Dim bytes As Integer = -1
            Do
                bytes = sslStream.Read(buffer, 0, buffer.Length)

                ' Use Decoder class to convert from bytes to UTF8
                ' in case a character spans two buffers.
                Dim decoder As Decoder = Encoding.UTF8.GetDecoder()
                Dim chars As Char() = New Char(decoder.GetCharCount(buffer, 0, bytes) - 1) {}
                decoder.GetChars(buffer, 0, bytes, chars, 0)
                messageData.Append(chars)
                ' Check for EOF.
                If messageData.ToString().IndexOf("<EOF>") <> -1 Then
                    Exit Do
                End If
            Loop While bytes <> 0

            Return messageData.ToString()
        End Function
#End Region
    End Class
End Namespace

