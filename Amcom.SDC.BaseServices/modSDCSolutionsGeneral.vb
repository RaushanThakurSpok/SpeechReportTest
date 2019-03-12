Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Runtime.CompilerServices
Imports System.IO
Imports System.Reflection

Public Module modSDCSolutionsGeneral

    Private mApp As Application = Application.GetInstance
    Public _AppMutex As Mutex = Nothing
    'Generic delegates
    Public Delegate Function Funk(Of TReturn)() As TReturn
    Public Delegate Function Funk(Of TReturn, TParam1)(ByVal param As TParam1) As TReturn
    Public Delegate Function Funk(Of TReturn, TParam1, TParam2)(ByVal param1 As TParam1, ByVal param2 As TParam2) As TReturn
    Public Delegate Function Funk(Of TReturn, TParam1, TParam2, TParam3)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3) As TReturn
    Public Delegate Function Funk(Of TReturn, TParam1, TParam2, TParam3, TParam4)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3, ByVal param4 As TParam4) As TReturn
    Public Delegate Function Funk(Of TReturn, TParam1, TParam2, TParam3, TParam4, TParam5)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3, ByVal param4 As TParam4, ByVal param5 As TParam5) As TReturn
    Public Delegate Sub SubProk()
    Public Delegate Sub SubProk(Of TParam1)(ByVal param As TParam1)
    Public Delegate Sub SubProk(Of TParam1, TParam2)(ByVal param1 As TParam1, ByVal param2 As TParam2)
    Public Delegate Sub SubProk(Of TParam1, TParam2, TParam3)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3)
    Public Delegate Sub SubProk(Of TParam1, TParam2, TParam3, TParam4)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3, ByVal param4 As TParam4)
    Public Delegate Sub SubProk(Of TParam1, TParam2, TParam3, TParam4, TParam5)(ByVal param1 As TParam1, ByVal param2 As TParam2, ByVal param3 As TParam3, ByVal param4 As TParam4, ByVal param5 As TParam5)

#Region "Enumerations"
    Public Enum FormatDataStyles
        [decimal]
        hex
        nonPrint
    End Enum
#End Region
#Region "Properties"

    Public ReadOnly Property App() As Application
        Get
            Return mApp
        End Get
    End Property

#End Region

#Region "Methods"
    ''' <summary>
    ''' Reads value from sql db and handles for null values
    ''' </summary>
    ''' <param name="reader">SQL Data Reader object</param>
    ''' <param name="field">Name of DB field</param>
    ''' <param name="default">Default value to return if null</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As String) As String
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetString(index)
        End If
    End Function
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As Short) As Short
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetInt16(index)
        End If
    End Function
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As Integer) As Integer
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetInt32(index)
        End If
    End Function
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As Int64) As Int64
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetInt64(index)
        End If
    End Function
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As Boolean) As Boolean
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetBoolean(index)
        End If
    End Function
    Public Function CorrectNull(ByVal reader As System.Data.SqlClient.SqlDataReader, ByVal field As String, ByVal [default] As Double) As Double
        Dim index As Integer = 0
        Try
            index = reader.GetOrdinal(field)
        Catch ex As Exception
            Throw New ArgumentException("Field is not found.")
        End Try
        If IsDBNull(reader(index)) Then
            Return [default]
        Else
            Return reader.GetDouble(index)
        End If
    End Function
    Public Function CreateGUID() As String
        Return "{" & System.Guid.NewGuid().ToString & "}"
    End Function
    Public Function CreateFileName(ByVal extension As String) As String
        Return Format(Now, "s").Replace(":", "-") & "." & extension
    End Function
    Public Function ToMixedCase(ByVal text As String) As String
        Dim result As String = Regex.Replace(text, "\w+", AddressOf ToMixedCase)
        Return result
    End Function

    Private Function ToMixedCase(ByVal m As Match) As String
        Dim x As String = m.ToString().ToLower
        If x.Length = 0 Then Return ""
        Return Char.ToUpper(x.Chars(0)) & x.Substring(1, x.Length - 1)
    End Function

    Public Function FormatData(ByVal data As String, ByVal style As FormatDataStyles)
        Dim sOut As New Text.StringBuilder
        For Each s As Char In data
            Select Case style
                Case FormatDataStyles.decimal
                    sOut.Append(Format(Asc(s), "00"))
                Case FormatDataStyles.hex
                    Dim h As String = Hex(Asc(s))
                    If h.Length = 1 Then
                        sOut.Append("0")
                    End If
                    sOut.Append(h)
                Case FormatDataStyles.nonPrint
                    Select Case Asc(s)
                        Case 0
                            sOut.Append("<NUL>")
                        Case 1
                            sOut.Append("<SOH>")
                        Case 2
                            sOut.Append("<STX>")
                        Case 3
                            sOut.Append("<ETX>")
                        Case 4
                            sOut.Append("<EOT>")
                        Case 5
                            sOut.Append("<ENQ>")
                        Case 6
                            sOut.Append("<ACK>")
                        Case 7
                            sOut.Append("<BEL>")
                        Case 8
                            sOut.Append("<BS>")
                        Case 9
                            sOut.Append("<TAB>")
                        Case 10
                            sOut.Append("<LF>")
                        Case 11
                            sOut.Append("<VT>")
                        Case 12
                            sOut.Append("<FF>")
                        Case 13
                            sOut.Append("<CR>")
                        Case 14
                            sOut.Append("<SO>")
                        Case 15
                            sOut.Append("<SI>")
                        Case 16
                            sOut.Append("<DLE>")
                        Case 17
                            sOut.Append("<DC1>")
                        Case 18
                            sOut.Append("<DC2>")
                        Case 19
                            sOut.Append("<DC3>")
                        Case 20
                            sOut.Append("<DC4>")
                        Case 21
                            sOut.Append("<NAK>")
                        Case 22
                            sOut.Append("<SYN>")
                        Case 23
                            sOut.Append("<ETB>")
                        Case 24
                            sOut.Append("<CAN>")
                        Case 25
                            sOut.Append("<EM>")
                        Case 26
                            sOut.Append("<SUB>")
                        Case 27
                            sOut.Append("<ESC>")
                        Case 28
                            sOut.Append("<FS>")
                        Case 29
                            sOut.Append("<GS>")
                        Case 30
                            sOut.Append("<RS>")
                        Case 31
                            sOut.Append("<US>")
                        Case Else
                            sOut.Append(s)
                    End Select
            End Select
        Next
        sOut.Append(" ")
        Return sOut.ToString.Trim
    End Function


    Public Function SqlEscape(ByVal s As String) As String
        Return s.Replace("'", "''")
    End Function


    <Extension()> _
       Public Sub ReflectToDebug(ByVal O As Object)
        Dim stWrite As New System.IO.StringWriter
        Reflect(O, stWrite)
        stWrite.Flush()
        Debug.Print(stWrite.ToString)
    End Sub
    <Extension()> _
       Public Function ReflectToString(ByVal O As Object) As String
        Dim stWrite As New System.IO.StringWriter
        Reflect(O, stWrite)
        stWrite.Flush()
        Return stWrite.ToString
    End Function

    <Extension()> _
    Public Sub Reflect(ByVal O As Object)
        Reflect(O, Console.Out)
    End Sub

    <Extension()> _
    Public Sub Reflect(ByVal O As Object, ByVal writer As TextWriter)

        If (TypeOf O Is IEnumerable) Then DirectCast(O, IEnumerable).Reflect(writer)

        Dim info As PropertyInfo() = O.GetType().GetProperties()

        For Each item As PropertyInfo In info
            Try
                writer.WriteLine("{0} : {1}", item.Name, item.GetValue(O, Nothing))
            Catch
                writer.WriteLine("{0} : {1}", item.Name, item.ToString())
            End Try
        Next

        writer.WriteLine(O)
    End Sub

    <Extension()> _
    Public Sub Reflect(ByVal O As IList, ByVal writer As TextWriter)

        For Each elem As Object In O
            Reflect(elem, writer)
        Next

    End Sub

    <Extension()> _
    Public Sub Reflect(ByVal O As IEnumerable, ByVal writer As TextWriter)

        For Each elem As Object In O

            Reflect(elem, writer)
        Next

    End Sub

#End Region

End Module
