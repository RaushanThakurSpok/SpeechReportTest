Imports System.IO
Imports System.Threading

Public Class FileLock

    Public Shared Function WaitForExclusiveAccess(ByVal fileName As String) As Boolean
        Return WaitForExclusiveAccess(fileName, 250, 5)
    End Function

    Public Shared Function WaitForExclusiveAccess(ByVal fileName As String, ByVal waitms As Integer) As Boolean
        Return WaitForExclusiveAccess(fileName, waitms, 5)
    End Function

    Public Shared Function WaitForExclusiveAccess(ByVal fileName As String, ByVal waitms As Integer, ByVal tries As Integer) As Boolean
        Dim count As Integer = 1
        Dim ok As Boolean = False
        Do
            Dim myFileStream As System.IO.FileStream = Nothing
            Try   ' attempt to open the file  
                myFileStream = New IO.FileStream(FileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
            Catch ex As IOException
                'dont care--we errored
            Finally
                If Not myFileStream Is Nothing Then
                    ok = myFileStream.CanRead
                    myFileStream.Close()
                End If
            End Try
            If ok Then Exit Do
            count = count + 1
            If count > Tries Then Exit Do
            Try
                Thread.Sleep(Waitms)
            Catch ex As Exception
            End Try
        Loop
        Return ok
    End Function

End Class
