Imports System.Net.Mail
Imports System.Reflection
Imports Utilities.Code.CodeUtilities
Imports System.Threading
Imports Utilities

Namespace Notifications
    Public NotInheritable Class EmailNotification
        Public Shared Event NotificationCompleted As EventHandler(Of NotificationCompletedEventArgs)
        Public Shared Sub SendEmailAsync(ByVal toAddress As String, ByVal subject As String, ByVal body As String, ByVal fromAddress As String, ByVal ParamArray attachments() As String)
            Dim s() As String = {toAddress, subject, body, fromAddress}
            Dim t As Tuple(Of Object, Object) = New Tuple(Of Object, Object)(s, attachments)
            Dim tr As New Thread(AddressOf SendMailThread)
            tr.IsBackground = True
            tr.Start(t)
        End Sub
        Private Shared Sub SendMailThread(ByVal o As Object)
            Dim t As Tuple(Of Object, Object) = DirectCast(o, Tuple(Of Object, Object))
            Dim params() As String = t.Item1
            Dim attachments() As String = t.Item2
            Try
                If EmailNotification.SendEmail(params(0), params(1), params(2), params(3), attachments) Then
                    RaiseEvent NotificationCompleted(Nothing, New NotificationCompletedEventArgs(True))
                Else
                    RaiseEvent NotificationCompleted(Nothing, New NotificationCompletedEventArgs(False))
                End If
            Catch ex As Exception
            End Try
        End Sub
        ''' <summary>
        ''' Sends an email message
        ''' </summary>
        ''' <param name="toAddress">email addressee.  Seperate multiple email addresses with a semicolon (;)</param>
        ''' <param name="subject">subject of message</param>
        ''' <param name="body">body of message</param>
        ''' <param name="fromAddress">from address (may not display on delivery)</param>
        ''' <param name="attachments">One or more files with full file path</param>
        ''' <returns></returns>
        ''' <remarks>If fails, will throw a NotifierFailedException</remarks>
        Public Shared Function SendEmail(ByVal toAddress As String, ByVal subject As String, ByVal body As String, ByVal fromAddress As String, ByVal ParamArray attachments() As String) As Boolean
            Dim mailSent As Boolean = False
            Dim configs As New Configuration()
            If Not configs.HasEmailProvider Then
                Throw New NotifierFailedException("No email providers have been configured in App.xml.")
            End If
            App.TraceLog(TraceLevel.Verbose, "Notification:SendMail... enumerating providers...")
            For count As Integer = 0 To configs.EmailProviders.Count - 1
                Dim email As IMailProvider = Nothing
                Dim pc As Configuration.ProviderConfiguration = configs.EmailProviders(count)
                'see if we can load the provider plugin.
                Dim ass As Assembly = Nothing
                Try
                    ass = Assembly.LoadFrom(IO.Path.Combine(App.PluginPath, pc("implementation")))
                Catch ex As Exception
                    Throw New NotifierFailedException("Failed to load plugin " & pc("implementation"))
                End Try
                email = ReflectionUtilities.FindInstanceByInterface(Of IMailProvider)(ass)
                If email Is Nothing Then
                    Continue For
                End If
                App.TraceLog(TraceLevel.Verbose, "Loaded priority {0} email provider '{1}'.", pc("priority"), email.Name)
                Dim params As New Dictionary(Of String, String)
                For Each key As String In pc.Keys
                    If key = "implementation" Then Continue For
                    params.Add(key, pc(key))
                Next
                If email Is Nothing Then
                    Throw New NotifierFailedException("Failed to find email interface in provided plugins.  Check installation/configuration")
                End If

                email.SetConfiguration(params)
                Try
                    email.SendMail(fromAddress, toAddress, subject, body, attachments)
                    mailSent = True
                Catch ex As ArgumentException
                    App.ExceptionLog(New AmcomException(ex.Message, ex))
                Catch ex As NotifierFailedException
                    App.ExceptionLog(ex)
                End Try
                If mailSent Then
                    App.TraceLog(TraceLevel.Verbose, "Notification email sent successfully")
                    Return True
                End If
            Next
            Return mailSent
        End Function

    End Class
End Namespace
