Imports System.Timers
Imports System.Threading
Imports System.Xml
Imports System.Text

Namespace TCP
    Public Class MessagingDaemonClient

        Private _TCPClient As Client
        Private _PrimaryServerName As String
        Private _ServerPort As Integer
        Private _Connected As Boolean = False
        Private _Initialized As Boolean = False
        Private _Messages As MessagingDaemonMessages
        Private _Stop As Boolean = False
        Private _Buffer As String
        Private _Busy As Boolean
        Private _ConnectionAttemptsBeginEventRaised As Boolean
        Private _myLogLevel As TraceLevel
        Private _mdcThread As Thread = Nothing
        Private _ignoreSSLChainErrors As Boolean = False

        Private myLock As New Object

        Private _encoding As MDMEncoding = MDMEncoding.ASCII7

        Public Enum sendType As Integer
            Normal
            iDesk
        End Enum

        Public Enum TransportType
            tcp
            ssl
        End Enum

        Public Enum MDMEncoding
            ASCII
            ASCII7
            UTF8
        End Enum

        Public Event CommandReceived As EventHandler(Of MessagingDaemonMessageArgs)
        Public Event DaemonConnected As EventHandler
        Public Event DaemonDisconnected As EventHandler
        Public Event FailedConnectionAttempt As EventHandler
        Public Event ConnectionAttemptsBegin As EventHandler

        Public Property Encoding() As MDMEncoding
            Get
                Return _encoding
            End Get
            Set(ByVal value As MDMEncoding)
                _encoding = value
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

        Public Property MyLogLevel() As TraceLevel
            Get
                Return _myLogLevel
            End Get
            Set(ByVal value As TraceLevel)
                _myLogLevel = value
            End Set
        End Property

        Public ReadOnly Property Messages() As MessagingDaemonMessages
            Get
                Return _Messages
            End Get
        End Property

        Public ReadOnly Property Connected() As Boolean
            Get
                Return _Connected
            End Get
        End Property

        Public Property DaemonServerName() As String
            Get
                Return _PrimaryServerName
            End Get
            Set(ByVal value As String)
                _PrimaryServerName = value
            End Set
        End Property

        Public Sub New(ByVal daemonName As String, ByVal daemonPort As Integer, ByVal messageList As MessagingDaemonMessages)
            _ServerPort = daemonPort
            _PrimaryServerName = daemonName
            _Connected = False
            _Messages = messageList
            _Buffer = ""
            _Busy = False
            _ConnectionAttemptsBeginEventRaised = False
            '
            _TCPClient = New Client()
            _TCPClient.IgnoreSSLCertificateChainErrors = _ignoreSSLChainErrors
            '
            AddHandler _TCPClient.Connected, AddressOf HandleConnected
            AddHandler _TCPClient.DataReceived, AddressOf HandleDataReceived
            AddHandler _TCPClient.Disconnected, AddressOf HandleDisconnected
        End Sub

        Public Sub New(ByVal daemonName As String, ByVal daemonPort As Integer, ByVal messageList As MessagingDaemonMessages, ByVal transport As Client.TransportType)
            _ServerPort = daemonPort
            _PrimaryServerName = daemonName
            _Connected = False
            _Messages = messageList
            _Buffer = ""
            _Busy = False
            _ConnectionAttemptsBeginEventRaised = False
            '
            If transport = Client.TransportType.tcp Then
                _TCPClient = New Client()
            Else
                _TCPClient = New Client(_PrimaryServerName, _ServerPort)
            End If
            '
            AddHandler _TCPClient.Connected, AddressOf HandleConnected
            AddHandler _TCPClient.DataReceived, AddressOf HandleDataReceived
            AddHandler _TCPClient.Disconnected, AddressOf HandleDisconnected
        End Sub

        Public Sub Start()
            _Stop = False
            _Connected = False
            _TCPClient.IgnoreSSLCertificateChainErrors = _ignoreSSLChainErrors
            '
            Dim _mdcThread As New Thread(AddressOf SMThread)
            _mdcThread.Name = "MDC"
            _mdcThread.Start()
        End Sub

        Private Sub SMThread()
            BeginConnection()
            '
            Do
                Try
                    If (_Stop) Then Exit Do
                    Thread.Sleep(100)
                Catch ex As ThreadAbortException
                    _Stop = True
                    Exit Do
                End Try
            Loop
            '
            If (_Connected) Then
                _TCPClient.Disconnect()
                _TCPClient.Dispose()
            End If
        End Sub

        Public Sub [Stop]()
            _Stop = True
            If _mdcThread IsNot Nothing Then _mdcThread.Abort()
        End Sub

        Public Sub Clear()
            _Initialized = False
        End Sub

        Private Sub CheckForCommand()
            SyncLock myLock
                Dim startCmd As Integer = _Buffer.IndexOf(Chr(2))
                Dim endCmd As Integer = _Buffer.IndexOf(Chr(3))
                Dim m As MessagingDaemonMessage
                Dim s As String

                Try
                    If startCmd > -1 AndAlso endCmd > -1 AndAlso endCmd > startCmd Then
                        s = _Buffer.Substring(startCmd + 1, (endCmd - (startCmd + 1))).Replace(Chr(0), "")
                        _Buffer = _Buffer.Substring(endCmd + 1)
                        Dim doc As New XmlDocument
                        doc.LoadXml(s)
                        m = New MessagingDaemonMessage
                        m.Decode(doc)
                        RaiseEvent CommandReceived(Me, New MessagingDaemonMessageArgs(m))
                    ElseIf (endCmd > -1 AndAlso startCmd = -1) OrElse (endCmd > -1 AndAlso (endCmd > startCmd)) Then
                        _Buffer = _Buffer.Substring(endCmd + 1)
                    End If
                Catch ex As Exception
                    App.TraceLog(TraceLevel.Error, "CheckForCommand() Exception - {0}", ex.Message)
                End Try

            End SyncLock
        End Sub

        Public Sub BeginConnection()
            Do
                If (_Stop) Then Exit Do
                '
                If (Not _Connected) Then
                    Try
                        If (Not _ConnectionAttemptsBeginEventRaised) Then
                            RaiseEvent ConnectionAttemptsBegin(Me, New EventArgs)
                            _ConnectionAttemptsBeginEventRaised = True
                        End If
                        '
                        _TCPClient.Connect(_PrimaryServerName, _ServerPort)
                    Catch ex As AmcomException
                        App.TraceLog(_myLogLevel, _ServerPort)
                    Catch ex As Exception
                        App.TraceLog(_myLogLevel, "Error in MessageDaemonClient.BeginConnection: " & ex.ToString)
                    End Try
                    '
                    If (Not _TCPClient.IsConnected) Then
                        RaiseEvent FailedConnectionAttempt(Me, New EventArgs)
                    End If
                End If

                If (_Stop) Then Exit Do

                If _Buffer.Length > 0 Then
                    CheckForCommand()
                End If

                If _Connected Then
                    Thread.Sleep(20)
                Else
                    Thread.Sleep(1000)
                End If
            Loop
        End Sub

        Public Sub Send(ByVal msg As MessagingDaemonMessage)
            Send(msg, sendType.Normal)
        End Sub

        Public Sub Send(ByVal msg As MessagingDaemonMessage, ByVal sendingType As sendType)

            Dim sb As New Text.StringBuilder
            '
            Try
                sb.Append(Chr(2))
                If sendingType = sendType.Normal Then
                    sb.Append(msg.QueryDocument.OuterXml)
                Else
                    sb.Append(msg.ResponseDocument.OuterXml)
                End If
                sb.Append(Chr(3))
                _TCPClient.Send(sb.ToString)
                ' App.TraceLog(_myLogLevel, "Sending message [" & msg.Command & "]")
            Catch ex As Exception
                App.TraceLog(_myLogLevel, ex.ToString)
            End Try
        End Sub

        Private Sub HandleConnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            _Connected = True
            _ConnectionAttemptsBeginEventRaised = False
            App.TraceLog(_myLogLevel, "Connected to " & e.EndPoint.ToString)
            RaiseEvent DaemonConnected(Me, New EventArgs)
        End Sub

        Private Sub HandleDisconnected(ByVal sender As Object, ByVal e As TCPConnectionStateEventArgs)
            If _Stop Then Exit Sub
            _ConnectionAttemptsBeginEventRaised = False
            _Connected = False
            _Initialized = False
            App.TraceLog(_myLogLevel, "Disconnected...")
            RaiseEvent DaemonDisconnected(Me, New EventArgs)
        End Sub

        Private Sub HandleDataReceived(ByVal sender As Object, ByVal e As ClientDataReceivedEventArgs)
            Dim sData As String

            SyncLock myLock
                Select Case _encoding
                    Case MDMEncoding.ASCII
                        sData = e.Data.ToString()
                    Case MDMEncoding.ASCII7
                        sData = e.DataString()
                    Case MDMEncoding.UTF8
                        sData = e.UTF8
                    Case Else
                        sData = e.DataString()
                End Select

                If (sData.Length > 0) Then
					_Buffer = _Buffer & sData
                    _Buffer.Replace(Chr(0), "").Trim()
                End If
            End SyncLock

        End Sub

        Private Sub HandleCommandReceived(ByVal RemoteMessageReceived As MessagingDaemonMessage)
            ' App.TraceLog(_myLogLevel, "command is: " & RemoteMessageReceived.Command)
            RaiseEvent CommandReceived(Me, New MessagingDaemonMessageArgs(RemoteMessageReceived))
        End Sub

        Public Sub Disconnect()

            _TCPClient.Disconnect()
            _TCPClient.Dispose()
        End Sub

        Public Sub Dispose()
            Try
                _TCPClient.Disconnect()
            Catch ex As Exception
                'NOP
            Finally
                Try
                    _TCPClient.Dispose()
                Catch ex As Exception
                    'NOP
                End Try
            End Try
        End Sub
    End Class
End Namespace
