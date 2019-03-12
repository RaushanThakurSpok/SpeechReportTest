Imports System
Imports System.Collections
Imports System.Net.Sockets
Imports System.Text
Imports System.Net

Public Class SyslogSender
    Public Enum PriorityType
        Emergency
        Alert
        Critical
        [Error]
        Warning
        Notice
        Informational
        Debug
    End Enum
    Private Shared udp As UdpClient
    Private Shared ascii As New ASCIIEncoding()
    Private Shared machine As String = System.Net.Dns.GetHostName() & " "
    Private Shared m_sysLogIpAddress As String
    Private Sub New()
        SyslogSender.SysLogIpAddress = App.Config.SysLogServer
    End Sub
    Public Shared Property SysLogIpAddress() As String
        Get
            Return m_sysLogIpAddress
        End Get

        Set(ByVal value As String)
            m_sysLogIpAddress = value
        End Set
    End Property

    Public Shared Sub Send(ByVal ipAddress As String, ByVal body As String)
        If ipAddress Is Nothing OrElse (ipAddress.Length < 5) Then
            ipAddress = Dns.GetHostEntry(Dns.GetHostName()).AddressList(0).ToString()
        Else
            ipAddress = Dns.GetHostEntry(ipAddress).AddressList(0).ToString()
        End If
        Send(ipAddress, SyslogSender.PriorityType.Warning, DateTime.Now, body)
    End Sub
    Public Shared Sub Send(ByVal ipAddress As String, ByVal traceLevel As TraceLevel, ByVal time As DateTime, ByVal body As String)
        Dim priority As PriorityType
        Select Case traceLevel
            Case Diagnostics.TraceLevel.Off
                Exit Sub
            Case Diagnostics.TraceLevel.Info
                priority = PriorityType.Informational
            Case Diagnostics.TraceLevel.Error
                priority = PriorityType.Error
            Case Diagnostics.TraceLevel.Warning
                priority = PriorityType.Warning
            Case Diagnostics.TraceLevel.Verbose
                priority = PriorityType.Debug
            Case Else
                Exit Sub
        End Select
        If ipAddress Is Nothing OrElse (ipAddress.Length < 5) Then
            ipAddress = Dns.GetHostEntry(Dns.GetHostName()).AddressList(0).ToString()
        Else
            Dim lookup As Boolean = Not ipAddress.Contains(".")
            If Not lookup Then lookup = (From item In ipAddress.Split(".") Where Not IsNumeric(item)).Count > 0
            If lookup Then
                Try
                    If ipAddress.ToLower = "localhost" Then
                        ipAddress = Net.IPAddress.Loopback.ToString
                    Else
                        ipAddress = Dns.GetHostEntry(ipAddress).AddressList(0).ToString
                    End If
                Catch
                End Try
            End If
        End If
        udp = New UdpClient(ipAddress, 514)
        Dim rawMsg As Byte()
        Dim strParams As String() = {priority.ToString() & ": ", time.ToString("yyyy-MM-dd_HH:mm:ss.fff"), machine, body}
        rawMsg = ascii.GetBytes(String.Concat(strParams))
        udp.Send(rawMsg, rawMsg.Length)
        udp.Close()
        udp = Nothing
    End Sub

    Public Shared Sub Send(ByVal ipAddress As String, ByVal priority As PriorityType, ByVal time As DateTime, ByVal body As String)
        If ipAddress Is Nothing OrElse (ipAddress.Length < 5) Then
            ipAddress = Dns.GetHostEntry(Dns.GetHostName()).AddressList(0).ToString()
        End If
        udp = New UdpClient(ipAddress, 514)
        Dim rawMsg As Byte()
        Dim strParams As String() = {priority.ToString() & ": ", time.ToString("YYYY-MM-DD_HH:MM:SS.MMM"), machine, body}
        rawMsg = ascii.GetBytes(String.Concat(strParams))
        udp.Send(rawMsg, rawMsg.Length)
        udp.Close()
        udp = Nothing
    End Sub
End Class
