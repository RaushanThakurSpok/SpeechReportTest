Imports System
Imports System.Collections
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Net.Security

Namespace TCP
    Public Class ClientConnection
		Public Const cBufferSize As Integer = 16384
		Private mSocket As Socket = Nothing
        Private mClient As TcpClient = Nothing
        Private mEndPoint As IPEndPoint
		Private _sslStream As SslStream
		Private mBufferIndex As Integer
        Public Sub New(ByVal client As TcpClient)
            mClient = client
            mEndPoint = DirectCast(client.Client.RemoteEndPoint, IPEndPoint)
            mSocket = client.Client
		End Sub
        Public Sub New(ByVal client As TcpClient, ByVal secureStream As SslStream)
            mClient = client
            mEndPoint = DirectCast(client.Client.RemoteEndPoint, IPEndPoint)
            mSocket = client.Client
            _sslStream = secureStream
		End Sub
        Public Sub New(ByVal socket As Socket)
            mSocket = socket
            mEndPoint = DirectCast(socket.RemoteEndPoint, IPEndPoint)
		End Sub
        Public Sub New(ByVal socket As Socket, ByVal secureStream As SslStream)
            mSocket = socket
            mEndPoint = DirectCast(socket.RemoteEndPoint, IPEndPoint)
			_sslStream = secureStream
        End Sub
        Public ReadOnly Property Socket() As Socket
            Get
                Return mSocket
            End Get
        End Property
        Public ReadOnly Property EndPoint() As IPEndPoint
            Get
                Return mEndPoint
            End Get
		End Property
		Public ReadOnly Property SslStream() As SslStream
			Get
				Return _sslStream
			End Get
		End Property
		Public Property BufferIndex() As Integer
			Get
				Return mBufferIndex
			End Get
			Set(value As Integer)
				mBufferIndex = value
			End Set
		End Property
	End Class
End Namespace
