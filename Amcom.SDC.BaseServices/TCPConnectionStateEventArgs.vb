Imports System.net
Namespace TCP
    Public Class TCPConnectionStateEventArgs
        Inherits System.EventArgs
        Private mEndpoint As IPEndPoint
        Private _Reason As String
        Public Sub New(ByVal ep As IPEndPoint, ByVal reason As String)
            mEndpoint = ep
            _Reason = reason
        End Sub

        Public Sub New(ByVal ep As IPEndPoint)
            mEndpoint = ep
            _Reason = ""
        End Sub
        Public ReadOnly Property EndPoint() As IPEndPoint
            Get
                Return mEndpoint
            End Get
        End Property
        Public ReadOnly Property Reason() As String
            Get
                Return _Reason
            End Get
        End Property
    End Class
End Namespace