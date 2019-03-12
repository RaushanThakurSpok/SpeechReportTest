Public Class AppAlreadyRunningException
    Inherits Exception
    Public Sub New()
        MyBase.new()
    End Sub
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub
End Class
