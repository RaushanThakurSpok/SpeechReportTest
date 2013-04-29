Imports System.Xml
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Runtime.CompilerServices
Imports System.Security.Cryptography.X509Certificates

Namespace TCP
    Public Class MessagingDaemon

        Private Enum ReceiveModes
            text
            binary
        End Enum

        Private mEnabled As Boolean
        Private mPort As String
        Private mRemoteServer As Daemon
        Private mClients As List(Of String)
        Private mReceiveMode As ReceiveModes
        Private mMessages As MessagingDaemonMessages
        Private myLock As New Object
        Private _BufferData As ArrayList
        Private _BufferAddresses As ArrayList
        Private m As MessagingDaemonMessage
        Private _Stop As Boolean

        Private _CertFile As String = Nothing
        Private _CertX509 As X509Certificate = Nothing
        Private _SelfSignStartDate As Date = Nothing
        Private _SelfSignEndDate As Date = Nothing
        Private _SelfSignPassword As String = Nothing
        Private _ServerName As String = Nothing

        Private _transport As Daemon.DaemonTransport

#Region "Properties"
        Public ReadOnly Property Daemon() As Daemon
            Get
                Return mRemoteServer
            End Get
        End Property

        Public Property Messages() As MessagingDaemonMessages
            Get
                Return mMessages
            End Get
            Set(ByVal value As MessagingDaemonMessages)
                mMessages = value
            End Set
        End Property
#End Region

#Region "Events"
        Public Event CommandReceived As EventHandler(Of MessagingDaemonMessageArgs)
        Public Event ErrorOccurred As EventHandler(Of MessagingDaemonErrorArgs)
        Public Event DaemonClientDisconnected As EventHandler(Of TCPConnectionStateEventArgs)
        Public Event DaemonClientConnected As EventHandler(Of TCPConnectionStateEventArgs)
#End Region

#Region "Constructor"
        Public Sub New(ByVal Port As Integer)
            _Stop = False

            _transport = TCP.Daemon.DaemonTransport.tcp

            mPort = Port
            Initialize()
            '
            Dim t As New Thread(AddressOf ParseBuffer)
            t.Name = "ParseBuffer"
            t.IsBackground = True
            t.Start()
        End Sub

        Public Sub New(ByVal Port As Integer, ByVal certFile As String)
            _Stop = False

            _transport = TCP.Daemon.DaemonTransport.ssl
            _CertFile = certFile

            mPort = Port
            Initialize()
            '
            Dim t As New Thread(AddressOf ParseBuffer)
            t.Name = "ParseBuffer"
            t.IsBackground = True
            t.Start()
        End Sub

        Public Sub New(ByVal Port As Integer, ByVal certX509 As X509Certificate)
            _Stop = False

            _transport = TCP.Daemon.DaemonTransport.ssl
            _CertFile = ""
            _CertX509 = certX509

            mPort = Port
            Initialize()
            '
            Dim t As New Thread(AddressOf ParseBuffer)
            t.Name = "ParseBuffer"
            t.IsBackground = True
            t.Start()
        End Sub

        Public Sub New(ByVal Port As Integer, ByVal serverName As String, ByVal sslSelfSignStartDate As Date, ByVal sslSelfSignEndDate As Date, ByVal sslSelfSignPassword As String)
            _Stop = False

            _transport = TCP.Daemon.DaemonTransport.sslSelfSignedP2P
            _SelfSignStartDate = sslSelfSignStartDate
            _SelfSignEndDate = sslSelfSignEndDate
            _SelfSignPassword = sslSelfSignPassword
            _ServerName = serverName

            mPort = Port
            Initialize()
            '
            Dim t As New Thread(AddressOf ParseBuffer)
            t.Name = "ParseBuffer"
            t.IsBackground = True
            t.Start()
        End Sub
#End Region

#Region "Methods"
        Private Sub ParseBuffer()
            Do
                Try
                    If (_Stop) Then Exit Do
                    '
                    CheckForCommand()
                    '
                    Thread.Sleep(10)
                Catch ex As ThreadAbortException
                    _Stop = True
                    Exit Do
                End Try
            Loop
        End Sub

        Private Sub CheckForCommand()
            SyncLock myLock
                For i As Integer = 0 To _BufferAddresses.Count - 1
                    Dim sBuffer As String = _BufferData(i)
                    Dim startCmd As Integer = sBuffer.IndexOf(Chr(2))
                    Dim endCmd As Integer = sBuffer.IndexOf(Chr(3))
                    Dim sAddress = _BufferAddresses(i)
                    '
                    If startCmd > -1 AndAlso endCmd > -1 AndAlso endCmd > startCmd Then
                        Dim s As String = sBuffer.Substring(startCmd + 1, endCmd - 1)
                        sBuffer = sBuffer.Substring(endCmd + 1)
                        _BufferData(i) = sBuffer
                        Dim doc As New XmlDocument
                        doc.LoadXml(s)
                        m = New MessagingDaemonMessage
                        m.Decode(doc)
                        m.QueryAddress = sAddress
                        RaiseEvent CommandReceived(Me, New MessagingDaemonMessageArgs(m))
                    ElseIf (endCmd > -1 AndAlso startCmd = -1) OrElse (endCmd > -1 AndAlso (endCmd > startCmd)) Then
                        sBuffer = sBuffer.Substring(endCmd + 1)
                        _BufferData(i) = sBuffer
                    End If
                Next
            End SyncLock
        End Sub

        Private Sub ClientConnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            mClients.Add(e.EndPoint.ToString)
            App.TraceLog(TraceLevel.Info, e.EndPoint.Address.ToString & " has connected")
            RaiseEvent DaemonClientConnected(Me, e)
        End Sub

        Private Sub ClientDisconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            If (mClients.Contains(e.EndPoint.ToString)) Then
                Try
                    mClients.Remove(e.EndPoint.ToString)
                Catch ex As Exception
                    App.TraceLog(TraceLevel.Error, "There was a problem removing " & e.EndPoint.ToString & " from the mClients collection.")
                End Try
            End If
            App.TraceLog(TraceLevel.Info, e.EndPoint.Address.ToString & " has disconnected")
            RaiseEvent DaemonClientDisconnected(Me, e)
        End Sub

        Private Sub ClientDataReceived(ByVal sender As Object, ByVal e As TCPDataReceivedEventArgs)
            Select Case mReceiveMode
                Case ReceiveModes.text
                    SyncLock myLock
                        Dim iIndex As Integer
                        Dim sAddress As String = e.EndPoint.ToString
                        '            
                        If (e.DataString.Length > 0) Then
                            iIndex = _BufferAddresses.IndexOf(sAddress)
                            '
                            If (iIndex = -1) Then
                                _BufferAddresses.Add(sAddress)
                                _BufferData.Add(e.DataString)
                            Else
                                _BufferData(iIndex) &= e.DataString
                            End If
                        End If
                    End SyncLock
                Case ReceiveModes.binary
            End Select
        End Sub

        Public Sub ClientErrorOccurred(ByVal sender As Object, ByVal e As TCPErrorEventArgs)
            RaiseEvent ErrorOccurred(Me, New MessagingDaemonErrorArgs(e.EndPoint.Address.ToString, e.Error.Message))
        End Sub

        Public Sub Initialize()
            mClients = New List(Of String)
            mMessages = New MessagingDaemonMessages
            mEnabled = True 'App.Config.GetBoolean(ConfigSection.control, "remoteEnabled")
            'mPort = mPort '"8888" 'App.Config.GetInteger(ConfigSection.control, "remotePort")
            App.TraceLog(TraceLevel.Info, "Messaging Daemon port set to port " & mPort & " and is " & IIf(mEnabled, "Enabled", "Disabled"))
            mReceiveMode = ReceiveModes.text
            _BufferAddresses = New ArrayList
            _BufferData = New ArrayList
        End Sub

        Public Sub Send(ByVal message As MessagingDaemonMessage)
            Dim sb As New Text.StringBuilder
            If Not mClients.Contains(message.QueryAddress) Then
                If Daemon.ClientConnections.ContainsKey(message.QueryAddress) Then
                    App.TraceLog(TraceLevel.Verbose, "Cannot send remote command to '" & message.QueryAddress & "', it has disconnected.")
                End If
                Exit Sub
            End If
            sb.Append(Chr(2))
            sb.Append(message.ResponseDocument.OuterXml)
            sb.Append(Chr(3))
            Try
                mRemoteServer.Send(message.QueryAddress, sb.ToString)
            Catch ex As SocketException
                App.TraceLog(TraceLevel.Error, "Cannot send remote command to '" & message.QueryAddress & "', error: " & ex.Message)
            Catch ex As IO.IOException
                App.TraceLog(TraceLevel.Error, "Cannot send remote command to '" & message.QueryAddress & "', error: " & ex.Message)
            End Try
        End Sub
        Public Sub Start()
3:

            App.TraceLog(TraceLevel.Info, "Messaging Daemon Starting...")
            If mEnabled Then
                Try
                    _Stop = False
                    '
                    Dim port As Integer = mPort ' App.Config.GetInteger(ConfigSection.configuration, "control/remotePort")
                    '
                    If (mRemoteServer IsNot Nothing) Then
                        Try
                            mRemoteServer.ClientConnections.Clear()
                        Catch ex As Exception
                            App.TraceLog(TraceLevel.Error, ex.ToString)
                        Finally
                            mRemoteServer = Nothing
                        End Try
                    End If

                    Select Case _transport
                        Case TCP.Daemon.DaemonTransport.ssl
                            If _CertFile.Length > 0 Then
                                mRemoteServer = New Daemon(port, 1000, _CertFile)
                            Else
                                mRemoteServer = New Daemon(port, 1000, _CertX509)
                            End If

                        Case TCP.Daemon.DaemonTransport.sslSelfSignedP2P
                            mRemoteServer = New Daemon(port, 1000, _ServerName, _SelfSignStartDate, _SelfSignEndDate, _SelfSignPassword)
                        Case Else
                            mRemoteServer = New Daemon(port, 1000)
                    End Select
                    '
                    AddHandler mRemoteServer.Connected, AddressOf ClientConnected
                    AddHandler mRemoteServer.Disconnected, AddressOf ClientDisconnected
                    AddHandler mRemoteServer.DataReceived, AddressOf ClientDataReceived
                    AddHandler mRemoteServer.ErrorOccurred, AddressOf ClientErrorOccurred
                    mRemoteServer.Start()
                    App.TraceLog(TraceLevel.Info, "Messaging Daemon now listening on port " & mPort)
                Catch ex As Exception
                    App.ExceptionLog(New AmcomException("Failed to start TCPDaemon for Messaging Daemon", ex))
                End Try
            End If
        End Sub
        Public Sub [Stop]()
            _Stop = True
            If mRemoteServer Is Nothing Then Exit Sub
            App.TraceLog(TraceLevel.Info, "Messaging Daemon Stopping...")
            RemoveHandler mRemoteServer.Connected, AddressOf ClientConnected
            RemoveHandler mRemoteServer.Disconnected, AddressOf ClientDisconnected
            RemoveHandler mRemoteServer.DataReceived, AddressOf ClientDataReceived
            RemoveHandler mRemoteServer.ErrorOccurred, AddressOf ClientErrorOccurred
            mRemoteServer.Stop()
        End Sub
#End Region

    End Class
End Namespace
