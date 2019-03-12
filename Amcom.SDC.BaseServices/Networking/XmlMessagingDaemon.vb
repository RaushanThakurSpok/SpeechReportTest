Imports System.Xml
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Runtime.CompilerServices
Namespace TCP
    Public Class XmlMessagingDaemon

        Private Enum ReceiveModes
            text
            binary
        End Enum
        Private Class Client
            Public Key As String = String.Empty
            Public Buffer As String = String.Empty
            Public Sub New(ByVal clientkey As String)
                Key = clientkey
            End Sub
        End Class
        Private mEnabled As Boolean
        Private mPort As String
        Private mRemoteServer As Daemon
        Private mClients As Dictionary(Of String, XmlMessagingDaemon.Client)
        Private mReceiveMode As ReceiveModes
        Private mMessages As MessagingDaemonMessages
        Private myLock As New Object
        'Private _BufferData As ArrayList
        'Private _BufferAddresses As ArrayList
        Private m As MessagingDaemonMessage
        Private _Stop As Boolean
        Private _commandStart As String
        Private _commandStop As String
        Private _includeDelimiters As Boolean

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
        Public Event CommandReceived As EventHandler(Of XmlMessageArgs)
        Public Event ErrorOccurred As EventHandler(Of MessagingDaemonErrorArgs)
        Public Event DaemonClientDisconnected As EventHandler(Of TCPConnectionStateEventArgs)
        Public Event DaemonClientConnected As EventHandler(Of TCPConnectionStateEventArgs)
#End Region

#Region "Constructor"
        Public Sub New(ByVal Port As Integer)
            Me.New(Port, Chr(2), Chr(3), False)

        End Sub
        Public Sub New(ByVal port As Integer, ByVal startCommand As String, ByVal endCommand As String, ByVal includeDelimiters As Boolean)
            _commandStart = startCommand
            _commandStop = endCommand
            mClients = New Dictionary(Of String, Client)
            _includeDelimiters = includeDelimiters
            _Stop = False
            mPort = port
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
                For Each c As XmlMessagingDaemon.Client In mClients.Values
                    If c.Buffer.Length = 0 Then Continue For
                    Dim sBuffer As String = c.Buffer
                    'Dim startCmd As Integer = sBuffer.IndexOf(Chr(2))
                    'Dim endCmd As Integer = sBuffer.IndexOf(Chr(3))

                    Dim startCmd As Integer = sBuffer.IndexOf(_commandStart)
                    Dim endCmd As Integer = sBuffer.IndexOf(_commandStop)

                    Dim sAddress = c.Key
                    '
                    If startCmd > -1 AndAlso endCmd > -1 AndAlso endCmd > startCmd Then
                        Dim s As String
                        Try

                            s = sBuffer.Substring(startCmd + _commandStart.Length, endCmd - (startCmd + _commandStart.Length))
                            's = sBuffer.Substring(startCmd + _commandStart.Length, endCmd - (_commandStart.Length + 1))
                            sBuffer = sBuffer.Substring(endCmd + (_commandStop.Length))
                            c.Buffer = sBuffer

                            If _includeDelimiters Then
                                s = String.Format("{0}{1}{2}", _commandStart, s, _commandStop)
                            End If

                            Try
                                Dim doc As New XmlDocument
                                s = s.Replace(vbCrLf, Char.ToString(Chr(10)))
                                doc.LoadXml(s)
                                'm = New MessagingDaemonMessage
                                '
                                RaiseEvent CommandReceived(Me, New XmlMessageArgs(doc))
                            Catch ex As Exception
                                App.TraceLog(TraceLevel.Error, "Could not parse XML message: {0}", ex.Message)
                                App.TraceLog(TraceLevel.Error, "TCP Message Output: {0}", s)
                                RaiseEvent ErrorOccurred(Me, New MessagingDaemonErrorArgs("Error parsing XML.", String.Format("Error parsing XML document: {0}", ex.Message)))
                            End Try

                        Catch ex As Exception
                            App.TraceLog(TraceLevel.Error, "XML Parse Buffer Error - Substring issue: {0}", ex.Message)
                            App.TraceLog(TraceLevel.Error, "Buffer Content: {0}", sBuffer)
                            RaiseEvent ErrorOccurred(Me, New MessagingDaemonErrorArgs("XML Parse Buffer Error", String.Format("Error parsing XML document: {0}", ex.Message)))
                            c.Buffer = sBuffer.Substring(endCmd + _commandStop.Length)
                        End Try

                    ElseIf (endCmd > -1 AndAlso startCmd = -1) OrElse (endCmd > -1 AndAlso (endCmd > startCmd)) Then
                        Try
                            sBuffer = sBuffer.Substring(endCmd + _commandStop.Length + 1)
                            c.Buffer = sBuffer
                        Catch ex As Exception
                            App.TraceLog(TraceLevel.Error, "XML Parse Buffer Error - Substring issue: {0}", ex.Message)
                            App.TraceLog(TraceLevel.Error, "Buffer Content: {0}", sBuffer)
                            RaiseEvent ErrorOccurred(Me, New MessagingDaemonErrorArgs("XML Parse Buffer Error", String.Format("Error parsing XML document: {0}", ex.Message)))
                        End Try
                    End If
                Next
            End SyncLock
        End Sub

        Private Sub ClientConnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            Dim addr As String = e.EndPoint.ToString
            SyncLock myLock
                mClients.Add(addr, New Client(addr))
            End SyncLock
            App.TraceLog(TraceLevel.Info, e.EndPoint.Address.ToString & " has connected")
            RaiseEvent DaemonClientConnected(Me, e)
        End Sub

        Private Sub ClientDisconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            If (mClients.ContainsKey(e.EndPoint.ToString)) Then
                Try
                    SyncLock myLock
                        mClients.Remove(e.EndPoint.ToString)
                    End SyncLock
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
                        Dim sAddress As String = e.EndPoint.ToString
                        '            
                        If (e.UTF8.Length > 0) Then
                            If Not mClients.ContainsKey(sAddress) Then
                                App.TraceLog(TraceLevel.Error, "Received data for {0}, but no buffer found.", sAddress)
                                Exit Sub
                            End If
                            mClients(sAddress).Buffer &= e.UTF8
                        End If
                    End SyncLock
                Case ReceiveModes.binary
            End Select
        End Sub

        Public Sub ClientErrorOccurred(ByVal sender As Object, ByVal e As TCPErrorEventArgs)
            RaiseEvent ErrorOccurred(Me, New MessagingDaemonErrorArgs(e.EndPoint.Address.ToString, e.Error.Message))
        End Sub


        Public Sub Initialize()
            mClients = New Dictionary(Of String, XmlMessagingDaemon.Client)
            mMessages = New MessagingDaemonMessages
            mEnabled = True 'App.Config.GetBoolean(ConfigSection.control, "remoteEnabled")
            'mPort = mPort '"8888" 'App.Config.GetInteger(ConfigSection.control, "remotePort")
            App.TraceLog(TraceLevel.Info, "Messaging Daemon port set to port " & mPort & " and is " & IIf(mEnabled, "Enabled", "Disabled"))
            mReceiveMode = ReceiveModes.text


        End Sub
        Public Sub Send(ByVal message As MessagingDaemonMessage)
            Dim sb As New Text.StringBuilder
            If Not mClients.ContainsKey(message.QueryAddress) Then
                If Daemon.ClientConnections.ContainsKey(message.QueryAddress) Then
                    App.TraceLog(TraceLevel.Error, "Cannot send remote command to '" & message.QueryAddress & "', it has disconnected.")
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
                    mRemoteServer = New Daemon(port, 1000)
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

			For Each client As KeyValuePair(Of String, XmlMessagingDaemon.Client) In mClients
				Dim endpoint As String = client.Key
				Dim split() As String = endpoint.Split(":".ToCharArray())
				Dim ipAddress As IPAddress
				If Not ipAddress.TryParse(split(0), ipAddress) Then
					Continue For
				End If

				Dim port As Integer
				If Not Int32.TryParse(split(1), port) Then
					Continue For
				End If

				Dim ep As New IPEndPoint(ipAddress, port)
				Dim tcsea As New TCPConnectionStateEventArgs(ep, "Plugin shutting down")

				RaiseEvent DaemonClientDisconnected(Me, tcsea)
			Next

			RemoveHandler mRemoteServer.Connected, AddressOf ClientConnected
			RemoveHandler mRemoteServer.Disconnected, AddressOf ClientDisconnected
			RemoveHandler mRemoteServer.DataReceived, AddressOf ClientDataReceived
			RemoveHandler mRemoteServer.ErrorOccurred, AddressOf ClientErrorOccurred
			mRemoteServer.Stop()
		End Sub
#End Region

    End Class
End Namespace
