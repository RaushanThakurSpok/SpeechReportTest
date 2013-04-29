Namespace Notifications
    Public Class NotifierFailedException
        Inherits AmcomException
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class
End Namespace