Imports System.Net.Mail

Public Class SMTPSender
    Public Sub Send(ByVal SendTo As String, ByVal SendFrom As String, ByVal Subject As String, ByVal Body As String)
        Send(SendTo, SendFrom, Subject, Body, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    End Sub

    Public Sub Send(ByVal SendTo As String, ByVal SendFrom As String, ByVal Subject As String, ByVal Body As String, ByVal Attachments As StreamAttachments)
        Send(SendTo, SendFrom, Subject, Body, Attachments, Nothing, Nothing, Nothing, Nothing, Nothing)
    End Sub

    Public Sub Send(ByVal SendTo As String, ByVal SendFrom As String, ByVal Subject As String, ByVal Body As String, ByVal Attachments As StreamAttachments, ByVal Server As String, ByVal Port As Integer, ByVal UseSSL As Boolean)
        Send(SendTo, SendFrom, Subject, Body, Attachments, Server, Port, UseSSL, Nothing, Nothing)
    End Sub

    Public Sub Send(ByVal SendTo As String, ByVal SendFrom As String, ByVal Subject As String, ByVal Body As String, ByVal Attachments As StreamAttachments, ByVal Server As String, ByVal Port As Integer, ByVal UseSSL As Boolean, ByVal UserName As String, ByVal Password As String)
        'Start by creating a mail message object
        Dim MyMailMessage As New MailMessage()
        Dim aSendTo() As String

        If SendFrom = Nothing Then
            MyMailMessage.From = New MailAddress("amcom@amcomsoft.com")
        Else
            MyMailMessage.From = New MailAddress(SendFrom)
        End If

        If SendTo IsNot Nothing Then
            aSendTo = Split(SendTo, ";")
            'To is a collection of MailAddress types
            For Each addr As String In aSendTo
                MyMailMessage.To.Add(addr)
            Next
        End If

        MyMailMessage.Subject = Subject
        MyMailMessage.Body = Body

        If Attachments IsNot Nothing Then
            'To is a collection of MailAddress types
            For Each attach As StreamAttachment In Attachments
                Dim MyAttachment As New Attachment(attach.Stream, attach.Name)
                MyMailMessage.Attachments.Add(MyAttachment)
            Next
        End If

        'Create the SMTPClient object and specify the SMTP GMail server
        Dim SMTPServer As New SmtpClient
        If Server Is Nothing Then
            SMTPServer.Host = "127.0.0.1"
        Else
            SMTPServer.Host = Server
        End If

        If Port = 0 Then
            SMTPServer.Port = 25
        Else
            SMTPServer.Port = Port
        End If

        If UserName Is Nothing Then
            SMTPServer.UseDefaultCredentials = False
        Else
            SMTPServer.Credentials = New System.Net.NetworkCredential(UserName, Password)
        End If

        SMTPServer.EnableSsl = UseSSL

        Try
            SMTPServer.Send(MyMailMessage)
            'MessageBox.Show("Email Sent")
        Catch ex As SmtpException
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

End Class

Public Class StreamAttachments
    Inherits List(Of StreamAttachment)

    Public Sub New()

    End Sub

End Class

Public Class StreamAttachment
    Private _name As String
    Private _stream As IO.MemoryStream

    Public Sub New(ByVal name As String, ByVal stream As IO.MemoryStream)
        _name = name
        _stream = stream
    End Sub

    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property Stream() As IO.MemoryStream
        Get
            Return _stream
        End Get
        Set(ByVal value As IO.MemoryStream)
            _stream = value
        End Set
    End Property
End Class