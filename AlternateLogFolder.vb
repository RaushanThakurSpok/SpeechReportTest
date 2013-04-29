Public Class AlternateLogFolder
    Private _folderName As String
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="folderName"></param>
    ''' <remarks>Name cannot contain a backslash, or an InvalidArgument exception will be thrown</remarks>
    Public Sub New(ByVal folderName As String)
        If folderName.Contains("\") Then
            Throw New ArgumentException("alternate folder name may not contain backslash (\)")
        End If
        _folderName = folderName
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
End Class
