Namespace TCP
    Public Class MessagingDaemonMessageArgs
        Inherits EventArgs
        Private mMsg As MessagingDaemonMessage
        Public Sub New()
            mMsg = Nothing
        End Sub
        Public Sub New(ByVal msg As MessagingDaemonMessage)
            mMsg = msg
        End Sub
        Public Property Message() As MessagingDaemonMessage
            Get
                Return mMsg
            End Get
            Set(ByVal value As MessagingDaemonMessage)
                mMsg = value
            End Set
        End Property
    End Class
End Namespace
