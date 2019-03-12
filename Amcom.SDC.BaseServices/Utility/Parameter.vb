Public Class Parameter
    Implements IDisposable

    Private mName As String
    Private mValue As Object
    Private mPositionID As Integer
    Private mIsOptional As Boolean
    Friend Event ValueChanged As EventHandler

    Public Sub New(ByVal name As String)
        mName = name
    End Sub
    Public Sub New(ByVal name As String, ByVal value As Object)
        mName = name
        mValue = value
    End Sub
    Public Sub New(ByVal name As String, ByVal value As String)
        mName = name
        mValue = value
    End Sub
    Public Sub New(ByVal name As String, ByVal positionId As Integer)
        mName = name
        mValue = Value
        mPositionID = positionId
        mIsOptional = False
    End Sub
    Public Sub New(ByVal name As String, ByVal positionId As Integer, ByVal isOptional As Boolean)
        mName = name
        mValue = Value
        mPositionID = positionId
        mIsOptional = isOptional
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Return mName
        End Get
    End Property
    Public ReadOnly Property DisplayName() As String
        Get
            Return Utility.InitCap(Utility.CamelPascalToWords(mName))
        End Get
    End Property
    Public Property ObjectValue() As Object
        Get
            Return mValue
        End Get
        Set(ByVal value As Object)
            mValue = value
            RaiseEvent ValueChanged(Me, New System.EventArgs)
        End Set
    End Property
    Public Overridable Property Value() As String
        Get
            Select Case TypeName(mValue)
                Case "String"
                    Return mValue
                Case "Char"
                    Return mValue.ToString
                Case "Integer"
                    Return mValue.ToString
                Case "Long"
                    Return mValue.ToString
                Case "Short"
                    Return mValue.ToString
                Case "ULong"
                    Return mValue.ToString
                Case "UInteger"
                    Return mValue.ToString
                Case Else
                    Return String.Empty
            End Select
        End Get
        Set(ByVal value As String)
            mValue = value
            RaiseEvent ValueChanged(Me, New System.EventArgs)
        End Set
    End Property
    Public ReadOnly Property ParameterValue() As Parameter
        Get
            If TypeOf mValue Is Parameter Then
                Return mValue
            Else
                Return Nothing
            End If
        End Get

    End Property
    Public ReadOnly Property ParameterList() As Parameters
        Get
            If TypeOf mValue Is Parameters Then
                Return mValue
            Else
                Return Nothing
            End If
        End Get

    End Property
    Public ReadOnly Property PositionId() As Integer
        Get
            Return mPositionID
        End Get
    End Property

    Public Function Copy() As Parameter
        Return New Parameter(mName, mPositionID)
    End Function
    Public Function Copy(ByVal positionId As Integer) As Parameter
        Return New Parameter(mName, positionId)
    End Function

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If TypeOf mValue Is Parameter Then
                    mValue.Disppose()
                ElseIf TypeOf mValue Is Parameters Then
                    mValue.Dispose()
                End If
                mValue = Nothing
                Console.WriteLine("Parameter Disposed")
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
