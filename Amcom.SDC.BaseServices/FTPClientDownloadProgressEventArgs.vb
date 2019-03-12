Public Class FTPClientDownloadProgressEventArgs
    Inherits EventArgs
    Private _filename As String
    Private _percentage As Integer

    Public Sub New(ByVal fileName As String, ByVal percentage As Integer)
        _fileName = fileName
        _percentage = percentage
    End Sub

    Public ReadOnly Property Percentage() As Integer
        Get
            Return _percentage
        End Get
    End Property

    Public ReadOnly Property FileName() As String
        Get
            Return _filename
        End Get
    End Property

End Class
