Imports System.Xml
Public Class XmlMessageArgs
    Inherits EventArgs
    Private mMsg As XmlDocument
    Public Sub New()
        mMsg = Nothing
    End Sub
    Public Sub New(ByVal msg As XmlDocument)
        mMsg = msg
    End Sub
    Public Property Message() As XmlDocument
        Get
            Return mMsg
        End Get
        Set(ByVal value As XmlDocument)
            mMsg = value
        End Set
    End Property
End Class
