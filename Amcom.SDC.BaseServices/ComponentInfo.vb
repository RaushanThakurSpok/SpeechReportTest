Public Class ComponentInfo
    Private _APIVersion As String = ""
    Private _FileVersion As String = ""
    Private _Path As String = ""

    Public Sub New(ByVal fullPath As String, ByVal apiVersion As String)

        Dim fv As FileVersionInfo = FileVersionInfo.GetVersionInfo(fullPath)
        _FileVersion = fv.FileVersion
        _APIVersion = apiVersion
        _Path = fullPath
    End Sub
    Public Sub New(ByVal fullpath As String)
        Dim apiVersion As String = ""
        Dim fileVersion As String = ""
        Dim fv As FileVersionInfo = FileVersionInfo.GetVersionInfo(fullpath)
        fileVersion = fv.FileVersion
        _APIVersion = apiVersion
        _FileVersion = fileVersion
        _Path = fullpath
    End Sub
    Public ReadOnly Property FileVersion() As String
        Get
            Return _FileVersion
        End Get
    End Property
    Public ReadOnly Property APIVersion() As String
        Get
            Return _APIVersion
        End Get
    End Property
    Public ReadOnly Property FullPath() As String
        Get
            Return _Path
        End Get
    End Property
    Public ReadOnly Property FileName() As String
        Get
            Return IO.Path.GetFileName(_Path)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim buffer As New Text.StringBuilder
        buffer.Append(Me.FileName)
        buffer.Append(" ")
        buffer.Append("FileVer: ")
        buffer.Append(_FileVersion)
        If _APIVersion.Length > 0 Then
            buffer.Append("ApiVer: ")
            buffer.Append(_APIVersion)
        End If
        Return buffer.ToString
    End Function
End Class
