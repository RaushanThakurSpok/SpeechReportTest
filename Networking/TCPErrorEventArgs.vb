Imports System.net
Imports System.Net.Sockets
Namespace TCP
    Public Class TCPErrorEventArgs
        Inherits System.EventArgs
        Private mEndpoint As IPEndPoint
        Private mError As SocketException
        Public Sub New(ByVal ep As IPEndPoint, ByVal [error] As SocketException)
            mEndpoint = ep
            mError = [error]
        End Sub
        Public ReadOnly Property EndPoint() As IPEndPoint
            Get
                Return mEndpoint
            End Get
        End Property
        Public ReadOnly Property [Error]() As SocketException
            Get
                Return mError
            End Get
        End Property

    End Class
End Namespace