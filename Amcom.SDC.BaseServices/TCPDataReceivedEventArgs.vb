Imports System.net
Imports System.text
Namespace TCP
    Public Class TCPDataReceivedEventArgs
        Inherits System.EventArgs
        Private mData As Byte()
        Private mEndPoint As IPEndPoint
        Public Sub New(ByVal endPoint As IPEndPoint, ByVal data As Byte())
            mData = data
            mEndPoint = endPoint
        End Sub
        Public ReadOnly Property Data() As Byte()
            Get
                Return mData
            End Get
        End Property
        Public ReadOnly Property DataString() As String
            Get
                Dim ascii As New Text.ASCIIEncoding
                Dim s As String = ascii.GetString(mData, 0, mData.Length)

                Return s
            End Get
        End Property


        Public ReadOnly Property EndPoint() As IPEndPoint
            Get
                Return mEndPoint
            End Get
        End Property

    End Class
End Namespace