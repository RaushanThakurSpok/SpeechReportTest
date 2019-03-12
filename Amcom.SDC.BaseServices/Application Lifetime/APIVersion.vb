Public Class APIVersion
    Private _major As Integer
    Private _minor As Integer

    Sub New(ByVal major As Integer, ByVal minor As Integer)
        _major = major
        _minor = minor
    End Sub

    Public ReadOnly Property MajorVersion() As Integer
        Get
            Return _major
        End Get
    End Property

    Public ReadOnly Property MinorVersion() As Integer
        Get
            Return _minor
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return _major & "." & _minor
    End Function

    Public Function IsEqualTo(ByVal majorVersion As Integer, ByVal minorVersion As Integer) As Boolean
        Return (_major = majorVersion) AndAlso (minorVersion = _minor)
    End Function
End Class
