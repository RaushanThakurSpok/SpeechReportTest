Imports System.IO
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Soap
Imports System.Threading
Imports System.Security.Permissions
<Serializable()> _
    Public Class AmcomException
    Inherits Exception

    Private mAppDomainName As String = AppDomain.CurrentDomain.FriendlyName
    Private mAssemblyName As String = System.Reflection.Assembly.GetExecutingAssembly.FullName
    Private mCallStack As String = Environment.StackTrace
    Private mDateTime As Date = Date.Now
    Private mExceptionId As String = Guid.NewGuid.ToString
    Private mMachineName As String = Environment.MachineName
    Private mThreadId As Integer = Thread.CurrentThread.ManagedThreadId
    Private mThreadName As String = Thread.CurrentThread.Name
    Private mThreadUser As String = Thread.CurrentPrincipal.Identity.Name

    Public Sub New()
    End Sub

    Public Sub New(ByVal message As String)
        MyBase.new(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.new(message, innerException)
    End Sub

    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.new(info, context)
        mAppDomainName = info.GetString("AppDomainName")
        mAssemblyName = info.GetString("AssemblyName")
        mCallStack = info.GetString("CallStack")
        mDateTime = info.GetDateTime("DateTime")
        mExceptionId = info.GetString("Id")
        mMachineName = info.GetString("MachineName")
        mThreadId = info.GetInt32("ThreadId")
        mThreadName = info.GetString("ThreadName")
        mThreadUser = info.GetString("ThreadUser")
    End Sub

    <SecurityPermissionAttribute(SecurityAction.Demand, SerializationFormatter:=True)> _
    Public Overrides Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        If info Is Nothing Then Throw New ArgumentNullException("info")
        info.AddValue("AppDomainName", mAppDomainName, GetType(String))
        info.AddValue("AssemblyName", mAssemblyName, GetType(String))
        info.AddValue("CallStack", mCallStack, GetType(String))
        info.AddValue("DateTime", mDateTime, GetType(Date))
        info.AddValue("Id", mExceptionId, GetType(String))
        info.AddValue("MachineName", mMachineName, GetType(String))
        info.AddValue("ThreadId", mThreadId, GetType(Integer))
        info.AddValue("ThreadName", mThreadName, GetType(String))
        info.AddValue("ThreadUser", mThreadUser, GetType(String))
        MyBase.GetObjectData(info, context)
    End Sub

    Public Overrides Function ToString() As String
        Using ms As MemoryStream = New MemoryStream()
            Dim context As StreamingContext = New StreamingContext(StreamingContextStates.File)
            Dim formatter As SoapFormatter = New SoapFormatter(Nothing, context)
            formatter.Serialize(ms, Me)
            ms.Position = 0
            Using sr As New StreamReader(ms)
                Return sr.ReadToEnd
            End Using
        End Using
    End Function

    Public ReadOnly Property AppDomainName() As String
        Get
            Return mAppDomainName
        End Get
    End Property

    Public ReadOnly Property AssemblyName() As String
        Get
            Return mAssemblyName
        End Get
    End Property

    Public ReadOnly Property CallStack() As String
        Get
            Return mCallStack
        End Get
    End Property

    Public ReadOnly Property DateTime() As Date
        Get
            Return mDateTime
        End Get
    End Property

    Public ReadOnly Property Id() As String
        Get
            Return mExceptionId
        End Get
    End Property

    Public ReadOnly Property MachineName() As String
        Get
            Return mMachineName
        End Get
    End Property

    Public ReadOnly Property ThreadId() As Integer
        Get
            Return mThreadId
        End Get
    End Property

    Public ReadOnly Property ThreadName() As String
        Get
            Return mThreadName
        End Get
    End Property

    Public ReadOnly Property ThreadUser() As String
        Get
            Return mThreadUser
        End Get
    End Property

End Class
