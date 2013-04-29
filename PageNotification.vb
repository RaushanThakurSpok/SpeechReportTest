Imports Amcom.SDC.BaseServices.TCP
Imports System.Threading
Namespace Notifications
    Public NotInheritable Class PageNotification
        ''' <summary>
        ''' Sends an alphanumeric page to the specified telephone number
        ''' </summary>
        ''' <param name="telNumber">Telephone number in PSTN format</param>
        ''' <param name="text">Page text (1000 chars max)</param>
        ''' <remarks>Throws an AcomException if error.  Page Server configurations are set in App.xml of host</remarks>
        Public Shared Sub SendPage(ByVal telNumber As String, ByVal text As String)
            Dim configs As New Configuration()
            If Not configs.HasPagerProvider Then
                Throw New AmcomException("No paging providers have been configured in App.xml.")
            End If
            For Each pc As Configuration.ProviderConfiguration In configs.PagingProviders
                Select Case pc.Item("type")
                    Case "intellidesk"
                        Try
                            Dim server As String = pc.Item("ipOrDnsName")
                            Dim port As Integer = pc.Item("port")
                            Dim service As String = pc.Item("service")
                            Dim page As New SendIntellideskPage()
                            page.SendPage(server, port, service, telNumber, text)
                        Catch ex As AmcomException
                            Throw New AmcomException(ex.Message)
                        Catch ex As Exception
                            Throw New AmcomException("Failed to send page -- invalid 'Intellidesk' configuration in 'notificationProviders' section of App.xml configuration file.")
                        End Try
                    Case Else
                        Throw New AmcomException("No paging services have been configured in the 'modules/notificationsProviders' section of App.xml.")
                End Select
            Next
        End Sub

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
