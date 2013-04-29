Imports System.Net.Mail
Namespace Notifications
    Public Class EmailNotification
        Public Shared Sub SendEmail(ByVal toAddress As String, ByVal subject As String, ByVal body As String, ByVal fromAddress As String, ByVal ParamArray Attachments() As String)
            Dim configs As New Configuration()
            If Not configs.HasEmailProvider Then
                Throw New AmcomException("No email providers have been configured in App.xml.")
            End If
            App.TraceLog(TraceLevel.Verbose, "Notification:SendMail... emerating providers...")
            For Each pc As Configuration.ProviderConfiguration In configs.EmailProviders
                Select Case pc.Item("type")
                    Case "amcom"
                    Case "gmail"
                        '<email priority="1" type="gmail">
                        '  <config>
                        '    <smtpServer>smtp.gmail.com</smtpServer>
                        '    <port>587</port>
                        '    <SslTls>true</SslTls>
                        '    <user>ENC(CJVpsdbmsg/i1JEyais7um/9GkcW3cdl5aYZXMtdVvc=)</user>
                        '    <password>ENC(66YTa4RPn57J5rAsOqknhg==)</password>
                        '  </config>
                        '</email>

                        App.TraceLog(TraceLevel.Verbose, "Notification:SendMail...Found GMail provider.")
                        Try
                            Dim server As String = pc.Item("smtpServer")
                            Dim port As Integer = pc.Item("port")
                            Dim user As String = App.Config.DecryptConfigString(pc.Item("user"))
                            Dim pass As String = App.Config.DecryptConfigString(pc.Item("password"))
                            Dim mailer As New MailIt()
                            App.TraceLog(TraceLevel.Verbose, "Notification:SendMail...Calling SendMail()")
                            mailer.SendMail(toAddress, fromAddress, subject, body, Nothing, server, port, True, user, pass)
                        Catch ex As AmcomException
                            Throw New AmcomException(ex.Message)
                        Catch ex As Exception
                            Throw New AmcomException("Failed to send email -- invalid 'Intellidesk' configuration in 'notificationProviders' section of App.xml configuration file.")
                        End Try
                    Case Else
                        Throw New AmcomException("No email services have been configured in the 'modules/notificationsProviders' section of App.xml.")
                End Select
            Next
        End Sub

        Private Class MailIt
            Public Sub SendMail(ByVal sendTo As String, ByVal sendFrom As String, ByVal subject As String, ByVal body As String, ByVal attachments As StreamAttachments, ByVal server As String, ByVal port As Integer, ByVal useSSL As Boolean, ByVal userName As String, ByVal password As String)
                'Start by creating a mail message object
                Dim MyMailMessage As New MailMessage(sendFrom, sendTo)

                MyMailMessage.Subject = Subject
                MyMailMessage.Body = Body
                If attachments IsNot Nothing Then
                    'To is a collection of MailAddress types
                    For Each attach As StreamAttachment In attachments
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
                    App.TraceLog(TraceLevel.Verbose, "Notification:SendMail...Sending Mail....")

                    SMTPServer.Send(MyMailMessage)
                    'MessageBox.Show("Email Sent")
                    App.TraceLog(TraceLevel.Verbose, "Notification:SendMail...Mail sent successfully")
                Catch ex As SmtpException
                    App.TraceLog(TraceLevel.Verbose, "Notification:SendMail...Failed to send mail: " & ex.Message)
                    Throw New AmcomException(ex.Message)
                End Try
            End Sub

        End Class
    End Class
End Namespace
