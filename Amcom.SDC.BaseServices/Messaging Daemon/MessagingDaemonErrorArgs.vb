Namespace TCP
    Public Class MessagingDaemonErrorArgs
        Inherits EventArgs
        Private mKey As String
        Private mMessage As String
        Public Sub New(ByVal key As String, ByVal message As String)
            mKey = key
            mMessage = message
        End Sub
        Public ReadOnly Property Key() As String
            Get
                Return mKey
            End Get
        End Property

        Public ReadOnly Property Message() As String
            Get
                Return mMessage
            End Get
        End Property

    End Class
End Namespace
