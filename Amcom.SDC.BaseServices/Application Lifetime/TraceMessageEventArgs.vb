Public Class TraceMessageEventArgs
    Inherits EventArgs
    Private mLevel As Integer
    Private mMessage As String
    Public Sub New(ByVal level As Integer, ByVal message As String)
        mLevel = level
        mMessage = message
    End Sub
    Public ReadOnly Property Level() As Integer
        Get
            Return mLevel
        End Get
    End Property
    Public ReadOnly Property Message() As String
        Get
            Return mMessage
        End Get
    End Property
End Class
