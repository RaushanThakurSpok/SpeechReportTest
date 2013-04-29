Namespace Notifications
    Public Class NotificationCompletedEventArgs
        Inherits EventArgs
        Public _success As Boolean
        Public Sub New(ByVal success As Boolean)
            _success = success
        End Sub
        Public ReadOnly Property Success() As Boolean
            Get
                Return _success
            End Get
        End Property
    End Class
End Namespace