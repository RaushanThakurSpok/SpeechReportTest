Imports System.Net.Sockets
Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Public Class ProtocolNotSupportedException
    Inherits System.Exception

    Public Sub New()
        MyBase.New("The protocol selected is not supported.")
    End Sub

    Public Sub New(ByVal Message As String)
        MyBase.New(Message)
    End Sub
End Class

Public Class UDPLogger
    


    ' Default property values
    Private mprotocol As ProtocolType = ProtocolType.Udp
    Private mClientAddress As IPAddress = IPAddress.Any
    Private mClientPort As Integer = 0
    Private mData As Byte() = New Byte() {}
    Private mEncode As New ASCIIEncoding
    Private mStack As Stack(Of String)
    Private mAbort As Boolean
    Public Sub New()
        Me.ClientAddress = IPAddress.Parse("127.0.0.1")
        Me.ClientPort = CType(5711, Integer)  '(S*D)+C
        mStack = New Stack(Of String)
        mAbort = False
    End Sub

   
    Public Property Client() As IPEndPoint
        Get
            Return New IPEndPoint(mClientAddress, mClientPort)
        End Get
        Set(ByVal Value As IPEndPoint)
            mClientAddress = Value.Address
            mClientPort = Value.Port
        End Set
    End Property
    Public Property ClientAddress() As IPAddress
        Get
            Return mClientAddress
        End Get
        Set(ByVal Value As IPAddress)
            mClientAddress = Value
        End Set
    End Property

    Public Property ClientPort() As Integer
        Get
            Return mClientPort
        End Get
        Set(ByVal Value As Integer)
            mClientPort = Value
        End Set
    End Property
    Public Sub Run()

        Do
            Thread.Sleep(50)
            Try
            Catch ex As ThreadAbortException
                mAbort = True
            End Try
            If mAbort Then Exit Do
            If mStack.Count Then
                Debug.Print("Stack pop, count is " & mStack.Count)
                Dim s As String = mStack.Pop
                SendTraceMessage(s)
            End If
        Loop
    End Sub
    Public Sub [Stop]()
        mAbort = True
    End Sub
    Private Sub SendTraceMessage(ByVal message As String)
        ' Encode message per settings
        mData = mEncode.GetBytes(message)
        ' Send the message
        SendUDPMessage(mData)

    End Sub
    Public Sub SendMessage(ByVal message As String)
        mStack.Push(message)
        Debug.Print("Stack push: Count is " & mStack.Count)
    End Sub
    Private Function SendUDPMessage(ByVal _data As Byte()) As Integer
        ' Create a UDP Server and send the message, then clean up
        Dim _UDPServer As UdpClient = Nothing
        Dim ReturnCode As Integer
        Try
            _UDPServer = New UdpClient
            ReturnCode = 0
            _UDPServer.Connect(Client)
            ReturnCode = _UDPServer.Send(_data, _data.Length)
        Catch ex As Exception
            Throw ex
        Finally
            If Not (_UDPServer Is Nothing) Then
                _UDPServer.Close()
            End If
        End Try
        Return ReturnCode
    End Function

End Class

 