Imports System.Text

Namespace UISupport
    Public Class PickListKeyEncoder

        Private _propertyName As String = ""
        Private _server As String = ""
        Private _database As String = ""
        Private _sql As String = ""
        Private _caption As String = ""
        Private _displayField As String = ""
        Private _selectField As String = ""
        Public Sub New(ByVal propertyName As String, ByVal caption As String, ByVal databaseServer As String, ByVal database As String, ByVal sql As String, ByVal displayField As String, ByVal selectField As String)
            _server = databaseServer
            _database = database
            _sql = sql
            _caption = caption
            _displayField = displayField
            _selectField = selectField
            _propertyName = propertyName
        End Sub
        Public ReadOnly Property PropertyName() As String
            Get
                Return _propertyName
            End Get
        End Property
        Public ReadOnly Property DisplayField() As String
            Get
                Return _displayField
            End Get
        End Property
        Public ReadOnly Property SelectField() As String
            Get
                Return _selectField
            End Get
        End Property
        Public ReadOnly Property Caption() As String
            Get
                Return _caption
            End Get
        End Property
        Public ReadOnly Property Server() As String
            Get
                Return _server
            End Get
        End Property
        Public ReadOnly Property Database() As String
            Get
                Return _database
            End Get
        End Property
        Public ReadOnly Property Sql() As String
            Get
                Return _sql
            End Get
        End Property
        Public Overrides Function ToString() As String
            Dim encoding As New System.Text.ASCIIEncoding
            Return "PICKL|" & _propertyName & "|" & _caption & "|" & _server & "|" & _database & "|" & _sql & "|" & _displayField & "|" & _selectField
        End Function
        Public Shared Function FromString(ByVal s As String) As PickListKeyEncoder
            If Not s.Contains("|") Then Return Nothing
            If Not s.Contains("PICKL") Then Return Nothing
            Dim i() As String = s.Split("|")
            If i.Count <> 8 Then Return Nothing
            Return New PickListKeyEncoder(i(1), i(2), i(3), i(4), i(5), i(6), i(7))
        End Function
    End Class
End Namespace