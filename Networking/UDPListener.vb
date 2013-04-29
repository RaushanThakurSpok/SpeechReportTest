Imports System.Net
Imports System.Net.Sockets

Namespace UDP
    Public Class UDPListener

        Private _Port As Integer
        Private _ListenerSocket As Socket

        Public Event MessageReceived(ByVal sender As Object, ByVal e As UDPListenerMessageReceivedEventArgs)

        Public Sub New(ByVal Port As Integer)
            _Port = Port
            '
            _ListenerSocket = New Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp)
            _ListenerSocket.EnableBroadcast = True
            _ListenerSocket.MulticastLoopback = True
            _ListenerSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
            _ListenerSocket.Bind(New IPEndPoint(System.Net.IPAddress.Any, _Port))
        End Sub

        Public Sub Start()
            StartReceive()
        End Sub

        Public Sub [Stop]()
            Try
                _ListenerSocket.Disconnect(True)
            Catch
                'NOP
            Finally
                Try
                    _ListenerSocket.Close()
                Catch
                    'NOP
                End Try
            End Try
        End Sub

        Private Sub StartReceive()
            Dim buff(65535) As Byte
            '
            Try
                _ListenerSocket.BeginReceive(buff, 0, buff.Length, SocketFlags.None, New AsyncCallback(AddressOf DataReceived), buff)
            Catch ex As Exception
                Try
                    Throw New Exception(ex.ToString)
                Catch ex2 As Exception
                    'NOP
                End Try
            End Try
        End Sub

        Private Sub DataReceived(ByVal ar As IAsyncResult)
            Dim buff() As Byte = CType(ar.AsyncState, Byte())
            '
            If (buff IsNot Nothing) Then
                RaiseEvent MessageReceived(Me, New UDPListenerMessageReceivedEventArgs(AdjustResponse(ToStringFromByteArray(buff))))
            End If
            '
            StartReceive()
        End Sub

        Private Function AdjustResponse(ByVal data As String) As String
            Dim sbReturn As New Text.StringBuilder
            '
            For Each c As Char In data
                If (c <> Nothing) Then
                    sbReturn.Append(c)
                Else
                    Exit For
                End If
            Next
            '
            Return sbReturn.ToString
        End Function

        Private Function ToStringFromByteArray(ByVal data() As Byte) As String
            Dim encoding As New System.Text.ASCIIEncoding
            Return encoding.GetString(data)
        End Function
    End Class
End Namespace
