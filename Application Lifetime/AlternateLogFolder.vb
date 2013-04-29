Public Class AlternateLogFolder
    Private _folderName As String
    Private _fileNamePrefix As String
    ''' <summary>
    ''' Specify an alternate folder name
    ''' </summary>
    ''' <param name="folderName"></param>
    ''' <remarks>Name cannot contain a backslash, or an InvalidArgument exception will be thrown</remarks>
    Public Sub New(ByVal folderName As String)
        If folderName.Contains("\") Then
            Throw New ArgumentException("alternate folder name may not contain backslash (\)")
        End If
        _folderName = folderName
        _fileNamePrefix = String.Empty

    End Sub

    ''' <summary>
    ''' Specify an alternate folder and/or file name
    ''' </summary>
    ''' <param name="folderName"></param>
    ''' <param name="fileNamePrefix"></param>
    ''' <remarks>Names cannot contain a backslash, or an InvalidArgument exception will be thrown</remarks>
    Public Sub New(folderName As String, fileNamePrefix As String)
        If folderName.Contains("\") Then
            Throw New ArgumentException("alternate folder name may not contain backslash (\)")
        ElseIf fileNamePrefix.Contains("\") Then
            Throw New ArgumentException("alternate file name may not contain backslash (\)")
        End If
        _folderName = folderName
        _fileNamePrefix = fileNamePrefix
    End Sub

    ''' <summary>
    ''' Returns the name of an alternate log folder
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FolderName() As String
        Get
            Return _folderName
        End Get
    End Property

    ''' <summary>
    ''' Returns the an alternate file name prefix
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileNamePrefix As String
        Get
            Return _fileNamePrefix
        End Get
    End Property
End Class
