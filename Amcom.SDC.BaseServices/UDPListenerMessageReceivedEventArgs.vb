Namespace UDP
    Public Class UDPListenerMessageReceivedEventArgs
        Inherits EventArgs

        Private _Message As String

        Public Sub New(ByVal Message As String)
            _Message = Message
        End Sub

        Public ReadOnly Property Message() As String
            Get
                Return _Message
            End Get
        End Property
    End Class
End Namespace
