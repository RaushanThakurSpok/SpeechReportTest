Imports Amcom.SDC.BaseServices.TCP
Imports System.Threading
Imports System.Reflection
Imports Utilities

Namespace Notifications
    Public NotInheritable Class PageNotification
        ''' <summary>
        ''' Sends an alphanumeric page to the specified telephone number
        ''' </summary>
        ''' <param name="telNumber">Telephone number in PSTN format</param>
        ''' <param name="text">Page text (1000 chars max)</param>
        ''' <remarks>Throws a NotifierFailedException if error.  Page Server configurations are set in App.xml of host</remarks>
        Public Shared Function SendPage(ByVal telNumber As String, ByVal text As String)
            Dim pageSent As Boolean = False
            Dim configs As New Configuration()
            If Not configs.HasPagerProvider Then
                Throw New NotifierFailedException("No paging providers have been configured in App.xml.")
            End If
            App.TraceLog(TraceLevel.Verbose, "Notification:Paging enumerating providers...")
            For count As Integer = 0 To configs.PagingProviders.Count - 1
                Dim pager As IPagerProvider = Nothing
                Dim pc As Configuration.ProviderConfiguration = configs.PagingProviders(count)
                'see if we can load the provider plugin.
                Dim ass As Assembly = Nothing
                Try
                    ass = Assembly.LoadFrom(IO.Path.Combine(App.PluginPath, pc("implementation")))
                Catch ex As Exception
                    Throw New NotifierFailedException("Failed to load plugin " & pc("implementation"))
                End Try
                pager = ReflectionUtilities.FindInstanceByInterface(Of IPagerProvider)(ass)
                If pager Is Nothing Then
                    Continue For
                End If
                App.TraceLog(TraceLevel.Verbose, "Loaded priority {0} pager provider '{1}'.", pc("priority"), pager.Name)
                Dim params As New Dictionary(Of String, String)
                For Each key As String In pc.Keys
                    If key = "implementation" Then Continue For
                    params.Add(key, pc(key))
                Next
                If pager Is Nothing Then
                    Throw New NotifierFailedException("Failed to find pager interface in provided plugins.  Check installation/configuration")
                End If

                pager.SetConfiguration(params)
                Try
                    pager.SendPage(telNumber, text)
                    pageSent = True
                Catch ex As ArgumentException
                    App.TraceLog(TraceLevel.Error, "EXCEPTION: " & ex.Message)
                Catch ex As NotifierFailedException
                    App.TraceLog(TraceLevel.Error, "EXCEPTION: " & ex.Message)
                End Try
                If pageSent Then
                    App.TraceLog(TraceLevel.Verbose, "Notification page sent successfully")
                    Return True
                End If
            Next
            Return pageSent
        End Function

        Private Class SendIntellideskPage
            Private WithEvents _client As Client
            Private _sig As AutoResetEvent
            Private _result As String = String.Empty
            Public Sub SendPage(ByVal ipdns As String, ByVal port As String, ByVal service As String, ByVal telNumber As String, ByVal text As String)
                _client = New Client()
                Try
                    _client.Connect(ipdns, port)
                Catch ex As Exception
                    Throw New AmcomException(String.Format("Failed to connect to server '{0}' on port {1}", ipdns, port))
                End Try
                Dim message As String = "PAGENOREFNUM|" & telNumber & "|" & service & "|" & text & "|A|" & vbCrLf
                _sig = New AutoResetEvent(False)
                _result = String.Empty
                _client.Send(message)
                _sig.WaitOne(5000)
                _client.Disconnect()
                _client.Dispose()
                If _result.Contains("ACCEPT") Then
                    Exit Sub
                ElseIf _result = String.Empty Then
                    Throw New AmcomException("Paging server failed to respond within 5 seconds.")
                ElseIf _result.Contains("FAILED") Then
                    Throw New AmcomException(String.Format("Failed to send page: '{0}'", _result))
                End If
            End Sub

            Private Sub _client_DataReceived(ByVal sender As Object, ByVal e As TCP.ClientDataReceivedEventArgs) Handles _client.DataReceived
                _result = e.DataString
                _sig.Set()
            End Sub
        End Class
    End Class
End Namespace
