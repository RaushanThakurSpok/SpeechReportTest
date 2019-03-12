Imports System.IO
Imports System.Text

''' <summary>
''' Class to read a text file backwards.
''' </summary>
''' <remarks></remarks>
Public Class BackwardReader
    Implements IDisposable

    Private _path As String = ""
    Private _fs As FileStream = Nothing

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="path"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String)
        _path = path
        _fs = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        _fs.Seek(0, SeekOrigin.[End])
    End Sub

    ''' <summary>
    ''' Read's the next line in.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadLine() As String
        Dim line As Byte()
        Dim text As Byte() = New Byte(1) {}
        Dim position As Long = 0
        Dim eotPos As Long = 0
        Dim count As Integer = 0

        _fs.Seek(0, SeekOrigin.Current)

        position = _fs.Position

        If _fs.Length > 1 Then
            Dim testBytes As Byte() = New Byte(1) {}
            _fs.Seek(-2, SeekOrigin.Current)
            _fs.Read(testBytes, 0, 2)

            Dim testString As String = ASCIIEncoding.ASCII.GetString(testBytes)

            If testString.Equals(vbCrLf) Then
                eotPos = position - 2
                _fs.Seek(-2, SeekOrigin.Current)
            End If

            While _fs.Position > 0
                _fs.Seek(-2, SeekOrigin.Current)
                _fs.Read(testBytes, 0, 2)
                testString = UTF8Encoding.UTF8.GetString(testBytes)

                If testString.Equals(vbCrLf) Then
                    count = (eotPos - _fs.Position)

                    'Get the entire line
                    line = New Byte(count) {}
                    _fs.Read(line, 0, count)

                    _fs.Seek(-(count), SeekOrigin.Current)

                    Return UTF8Encoding.UTF8.GetString(line)
                Else
                    _fs.Seek(-2, SeekOrigin.Current)
                End If
            End While
        End If

        Return String.Empty

    End Function

    ''' <summary>
    ''' Whether or not the start of file has been reached.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SOF() As Boolean
        Get
            Return _fs.Position = 0
        End Get
    End Property

    ''' <summary>
    ''' Closes the FileStream and disposes of it's resources.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close()
        _fs.Close() : _fs.Dispose()
    End Sub

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
                Me.Close()
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class