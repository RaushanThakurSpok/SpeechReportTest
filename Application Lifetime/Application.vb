Imports System.Configuration
Imports System.Deployment
Imports System.Deployment.Application
Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Soap
Imports System.Security.Permissions
Imports System.Threading
Imports System.Xml
Imports Microsoft.Win32
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Utilities
'Imports NetFwTypeLib
Imports Amcom.SDC.CodeUtilities
Imports System.Globalization
Public Enum Product
    [default]
    none
End Enum

Public Enum ApplicationType
    Windows = 0
    Console = 1
    WindowsService = 2
    NoConfigConsole = 3
    NoConfigComponent = 4
End Enum

Public Enum ApplicationInstanceType
    singleInstanceOnly
    multiInstance
End Enum
''' <summary>
''' The Application class manages an application lifetime,from startup through termination.  
''' </summary>
''' <remarks></remarks>
Public Class Application

    Implements IStatistics


    'used to implement the singleton pattern so that only one instance 
    'of this object is ever alive in the application
    Private Shared mInstance As Application
    Private Shared mMutex As New Mutex

    Private mApplicationType As ApplicationType           'the type of application
    Private ReadOnly mCompany As String = "Amcom Software"  'the name of our company
    Private mConfig As ConfigFile                         'the applications configuration file
    Private mIsInitialized As Boolean                     'has the application been initialized
    Private mName As String                               'name of te application
    Private mRootPath As String
    Private mSettings As ApplicationSettingsBase      'dll's can use
    Private mUserId As Integer
    Private mUserIsAdministrator As Boolean
    Private mUserName As String
    Private mVersion As String 'version
    Private mVisualStudioHosted As Boolean            'is the application being debugged by visual studio
    Private mUDPLogger As UDPLogger
    Private mSendUDPLogging As Boolean
    Private mShowExceptionDialogInVStudio As Boolean
    Public Event FirstRun As EventHandler(Of EventArgs)
    Public Event TraceMessage As EventHandler(Of TraceMessageEventArgs)
    Private ReadOnly mKey As String = "474AC74F-EE99-4bbf-ADF3-7BE96F9749E6"
    Private ReadOnly mKey2 As String = "5DF8E51C-A699-4a5d-990C-1BB5C9DAF712"
    Private mApplicationStartTime As Date = Nothing  'used to measure the running time of any application.  See Uptime()
    Private mLastLogCheck As Object = Nothing 'checking for old logs
    Private _OldLogDayMax As Integer = 0
    Private _OldExceptionLogDayMax As Integer = 0
    Private _settingsLoaded As Boolean = False
    Private _componentInfoCache As Dictionary(Of String, ComponentInfo)
    Private _alternameSettingsDomain As String = ""
    Private _syslogServer As String = ""
    Private _syslogEnabled As Boolean = False
    Private _fullApplicationPath As String = ""
    Private _cachedLogName As String = ""
    Private _maxLogSize As Integer = 0
    Private _maxWordWrapLength As Integer = -1
    Private _currentLogFilename As String = String.Empty
    Private _currentLogDate As DateTime = DateTime.MaxValue
    Private _loggedExceptionCountToday As Integer = 0
    'FireWall COM interface Constants
    Private Const CLSID_FIREWALL_MANAGER As String = "{304CE942-6E39-40D8-943A-B913C40C9CD4}"
    ' ProgID for the AuthorizedApplication object
    Private Const PROGID_AUTHORIZED_APPLICATION As String = "HNetCfg.FwAuthorizedApplication"
    Private _appName As String
    Private _securityLog As EventLog

    Private Event _SystemTimeAdjusted As EventHandler(Of SystemTimeAdjustedArgs)
    Public Delegate Sub SystemTimeAdjusted(ByVal sender As Object, ByVal e As SystemTimeAdjustedArgs)
    Private _staHandlers As List(Of SystemTimeAdjusted)
    Private _usingAppPathAsBase = False
    Private _traceLogLock As New Object

#Region "Properties"
    Public Property ApplicationName() As String
        Get
            Return _appName
        End Get
        Set(ByVal value As String)
            _appName = value
        End Set
    End Property
    Public ReadOnly Property ComponentInfoCache() As IDictionary(Of String, ComponentInfo)
        Get
            Return _componentInfoCache
        End Get
    End Property
    ''' <summary>
    ''' returns the Application type enum value for this application (console, windows, service,web)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ApplicationType() As ApplicationType
        Get
            Return mApplicationType
        End Get
    End Property
    ''' <summary>
    ''' Returns the ConfigFile object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Config() As ConfigFile
        Get
            Return mConfig
        End Get
    End Property
    ''' <summary>
    ''' Returns the dotted ip address of the host computer in IP4 format.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ComputerAddress() As String
        Get
            Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Environment.MachineName)
            Dim IpAddr As IPAddress() = ipEntry.AddressList
            Return IpAddr(0).ToString()
        End Get
    End Property
    ''' <summary>
    ''' Returns DNS hostname of the computer
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ComputerName() As String
        Get
            Return Dns.GetHostName
        End Get
    End Property
    ''' <summary>
    ''' Returns the Internet IP Address of the facility where there code is running.  
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>This will be the address outside any firewall or router.  If Internet not available, returns an empty string</remarks>
    Public Function InternetAddress() As String
        Try
            Dim request As HttpWebRequest = WebRequest.Create("http://checkip.dyndns.org/")
            Using response As HttpWebResponse = request.GetResponse()
                Using reader As New IO.StreamReader(response.GetResponseStream())
                    Dim sb As New Text.StringBuilder
                    Dim str As String = reader.ReadLine()
                    Do While str IsNot Nothing
                        sb.Append(str)
                        str = reader.ReadLine()
                    Loop
                    str = sb.ToString
                    str = str.Substring(InStr(str, ":")).Trim
                    str = str.Substring(0, InStr(str, "<") - 1)
                    Return str
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function
    ''' <summary>
    ''' Returns the City based on the WAN IP of this system
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Public API; decent accuracy</remarks>
    Public Function CityFromWANIP() As String
        Return CityFromIP("")
    End Function
    ''' <summary>
    ''' Returns the City based on the passed ip
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Public API; decent accuracy</remarks>
Public Function CityFromIP(ByVal ip As String) As String
        Try
            Dim request As HttpWebRequest = WebRequest.Create(IIf(ip.Length = 0, "http://ipinfodb.com/ip_query.php", "http://ipinfodb.com/ip_query.php?ip=" & ip))
            Using response As HttpWebResponse = request.GetResponse()
                Using reader As New IO.StreamReader(response.GetResponseStream())
                    Dim sb As New Text.StringBuilder
                    Dim str As String = reader.ReadToEnd
                    Dim doc As New XmlDocument
                    doc.LoadXml(str)
                    If doc.SelectSingleNode("Response/Status").InnerText <> "OK" Then Return ""
                    Return doc.SelectSingleNode("Response/City").InnerText & "," & doc.SelectSingleNode("Response/RegionName").InnerText & ", " & doc.SelectSingleNode("Response/CountryCode").InnerText
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Returns the path to the Plug-in Directory for the application using Base Services.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PluginPath() As String
        Get
            Return Path.Combine(mRootPath, "Plugins")
        End Get
    End Property
    ''' <summary>
    ''' Returns the Configuration Path for Applications using Base Services
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ConfigurationPath() As String
        Get
            Return Path.Combine(mRootPath, "Config")
        End Get
    End Property
    ''' <summary>
    ''' Returns the Path where serialized exceptions are placed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ExceptionPath() As String
        Get
            Return Path.Combine(mRootPath, "Exceptions")
        End Get
    End Property

    ''' <summary>
    ''' Returns the path where Log files are stored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LogPath() As String
        Get
            Return Path.Combine(mRootPath, "Logs")
        End Get
    End Property
    ''' <summary>
    ''' Returns the path where generic files are stored
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FilePath() As String
        Get
            Return Path.Combine(mRootPath, "Files")
        End Get
    End Property
    ''' <summary>
    ''' returns the base name of the application
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name() As String
        Get
            Return mName
        End Get
    End Property
    ''' <summary>
    ''' Returns full path to the Application/Service
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FullApplicationPath() As String
        Get
            Return _fullApplicationPath
        End Get
    End Property
    ''' <summary>
    ''' returns the root path of the application, excluding the name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RootPath() As String
        Get
            Return mRootPath
        End Get
    End Property
    ''' <summary>
    ''' Obsolete.  Syslog enabled through App.xml, sent by tracelog
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Do not use</remarks>
    Public Property SendUDPLogPackets() As Boolean
        Get
            Return False
        End Get
        Set(ByVal value As Boolean)
            'mSendUDPLogging = value
            'If value = False Then
            '    If mUDPLogger IsNot Nothing Then
            '        mUDPLogger.Stop()
            '    End If
            'Else
            '    mUDPLogger = New UDPLogger
            '    Dim t As New Thread(AddressOf StartUDPLogger)
            '    t.IsBackground = True
            '    t.Start()
            'End If
        End Set
    End Property
    ''' <summary>
    ''' returns the Base .Net Application settings object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Settings() As ApplicationSettingsBase
        Get
            Return mSettings
        End Get
    End Property
    ''' <summary>
    ''' When set true, enables showing the end-user exception dialog while in VS.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ShowExceptionDialogInVisualStudio() As Boolean
        Get
            Return mShowExceptionDialogInVStudio
        End Get
        Set(ByVal value As Boolean)
            mShowExceptionDialogInVStudio = value
        End Set
    End Property
    ''' <summary>
    ''' Returns the application version
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Version() As String
        Get
            Return mVersion
        End Get
    End Property
    ''' <summary>
    ''' Returns true if application is hosted in Visual Studio
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VisualStudioHosted() As Boolean
        Get
            Return mVisualStudioHosted
        End Get
    End Property

    'Public ReadOnly Property FirewallEnabled() As Boolean
    '    Get
    '        Dim manager As INetFwMgr = GetFirewallManager()
    '        Return manager.LocalPolicy.CurrentProfile.FirewallEnabled
    '    End Get
    'End Property
#End Region

#Region "Methods"



    '--- TO DO FIX FIREWALL
    'Private Function GetFirewallManager() As NetFwTypeLib.INetFwMgr
    '    App.TraceLog(TraceLevel.Error, "FireWall Management disabled in this version")
    '    Dim objectType As Type = Type.GetTypeFromCLSID(New Guid(CLSID_FIREWALL_MANAGER))
    '    Return (TryCast(Activator.CreateInstance(objectType), NetFwTypeLib.INetFwMgr))
    'End Function



    'Public Function AuthorizeApplication(ByVal title As String, ByVal applicationPath As String, ByVal scope As NET_FW_SCOPE_, ByVal ipVersion As NET_FW_IP_VERSION_) As Boolean
    '    ' Create the type from prog id
    '    Dim type__1 As Type = Type.GetTypeFromProgID(PROGID_AUTHORIZED_APPLICATION)
    '    Dim auth As INetFwAuthorizedApplication = TryCast(Activator.CreateInstance(type__1), INetFwAuthorizedApplication)
    '    Try
    '        auth.Name = title
    '        auth.ProcessImageFileName = applicationPath
    '        auth.Scope = scope
    '        auth.IpVersion = ipVersion
    '        auth.Enabled = True
    '    Catch ex As Exception
    '        'fail for any number of reasons... platform issues?
    '        Return False
    '    End Try

    '    Dim manager As NetFwTypeLib.INetFwMgr = GetFirewallManager()
    '    Try
    '        manager.LocalPolicy.CurrentProfile.AuthorizedApplications.Add(auth)
    '    Catch ex As Exception
    '        Return False
    '    End Try
    '    Return True

    'End Function
    ''' <summary>
    ''' Returns a database connection string based on information in App.xml
    ''' </summary>
    ''' <param name="applicationDatabase">pass a database section name here.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConnectionString(ByVal applicationDatabase As String, ByVal allowMultiRecordSets As WinConnectionInfo.RecordSetShareOptionType) As String
        Select Case ApplicationType
            Case ApplicationType.Console
                Return WinConnectionInfo.ConnectionString(applicationDatabase, allowMultiRecordSets)
            Case ApplicationType.Windows
                Return WinConnectionInfo.ConnectionString(applicationDatabase)
            Case ApplicationType.WindowsService
                Return WinConnectionInfo.ConnectionString(applicationDatabase)
        End Select
        Return ""
    End Function
    ''' <summary>
    ''' Returns a database connection string based on information in App.xml
    ''' </summary>
    ''' <param name="applicationDatabase">pass a database section name here.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConnectionString(ByVal applicationDatabase As String) As String
        Select Case ApplicationType
            Case ApplicationType.Console
                Return WinConnectionInfo.ConnectionString(applicationDatabase)
            Case ApplicationType.Windows
                Return WinConnectionInfo.ConnectionString(applicationDatabase)
            Case ApplicationType.WindowsService
                Return WinConnectionInfo.ConnectionString(applicationDatabase)
        End Select
        Return ""
    End Function
    ''' <summary>
    ''' Returns the default database connection string based on information in App.xml
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConnectionString() As String
        Select Case ApplicationType
            Case ApplicationType.Console
                Return WinConnectionInfo.ConnectionString
            Case ApplicationType.Windows
                Return WinConnectionInfo.ConnectionString
            Case ApplicationType.WindowsService
                Return WinConnectionInfo.ConnectionString
        End Select
        Return ""
    End Function
    ''' <summary>
    ''' Returns the default database connection string based on information in App.xml
    ''' </summary>
    ''' <param name="allowMultiRecordSets">allows multiple recordsets on same connection (SQL Server MARS)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConnectionString(ByVal allowMultiRecordSets As WinConnectionInfo.RecordSetShareOptionType) As String
        Select Case ApplicationType
            Case ApplicationType.Console
                Return WinConnectionInfo.ConnectionString(allowMultiRecordSets)
            Case ApplicationType.Windows
                Return WinConnectionInfo.ConnectionString(allowMultiRecordSets)
            Case ApplicationType.WindowsService
                Return WinConnectionInfo.ConnectionString(allowMultiRecordSets)
        End Select
        Return ""
    End Function
    ''' <summary>
    ''' Provides notification of system time changes
    ''' </summary>
    ''' <param name="eventDelegate">A handler for the event.  Use the AddressOf operator on the passed sub.</param>
    ''' <remarks>Throws NotSupportedException on WinXP, Win2K3 Server and other non-supported OS</remarks>
    Public Sub AddHandlerSystemTime(ByVal eventDelegate As SystemTimeAdjusted)
        'If this is the first time someone is hooking the event, enable looking at the security log
        Dim ver As Single = Single.Parse(My.Computer.Info.OSVersion.Substring(0, 3))
        If ver < 6.0 Then
            Throw New NotSupportedException(String.Format("System Time Change Event not supported on Windows Operating Syatems less than 6.0  This is {0} Vertsion {1}", My.Computer.Info.OSFullName, My.Computer.Info.OSVersion))
        End If

        If _staHandlers.Count = 0 Then
            _securityLog = New EventLog("Security", ".")
            AddHandler _securityLog.EntryWritten, AddressOf OnSecurityLogWrite
            _securityLog.EnableRaisingEvents = True
        End If
        'If Not _staHandlers.Contains(eventDelegate) Then
        _staHandlers.Add(eventDelegate)
        'End If
    End Sub
    Public Sub RemoveHandlerSystemTime(ByVal eventDelegate As SystemTimeAdjusted)
        'If this is the first time someone is hooking the event, enable looking at the security log
        If _staHandlers.Count = 0 Then Exit Sub
        For Each x As SystemTimeAdjusted In _staHandlers
            If x.GetHashCode = eventDelegate.GetHashCode Then
                _staHandlers.Remove(eventDelegate)
                Exit For
            End If
        Next
        If _staHandlers.Count = 0 Then
            _securityLog.EnableRaisingEvents = False
            RemoveHandler _securityLog.EntryWritten, AddressOf OnSecurityLogWrite
        End If
    End Sub
    Private Sub CheckForOldLogs(alternateLogFolder As AlternateLogFolder)
        Dim delList As New List(Of String)
        Dim process As Boolean = False
        If Not _settingsLoaded Then Exit Sub
        If _OldLogDayMax = 0 Then
            _OldLogDayMax = Me.Config.GetInteger(ConfigSection.diagnostics, "deleteLogsAfterDays", 0)
            If _OldLogDayMax = 0 Then
                _OldLogDayMax = Me.Config.GetInteger(ConfigSection.diagnostics, "logfile/deleteLogsAfterDays", 15)
            End If
        End If
        If _OldExceptionLogDayMax = 0 Then
            _OldExceptionLogDayMax = Me.Config.GetInteger(ConfigSection.diagnostics, "deleteExceptionLogsAfterDays", 0)
            If _OldExceptionLogDayMax = 0 Then
                _OldExceptionLogDayMax = Me.Config.GetInteger(ConfigSection.diagnostics, "logfile/deleteExceptionLogsAfterDays", 15)
            End If
        End If
        If _OldLogDayMax < 1 AndAlso _OldExceptionLogDayMax <= 1 Then Exit Sub
        If _OldLogDayMax < 1 AndAlso _OldExceptionLogDayMax <= 1 Then Exit Sub
        If mLastLogCheck Is Nothing Then
            mLastLogCheck = Now
            process = True
        ElseIf Math.Abs(DateDiff(DateInterval.Day, mLastLogCheck, Now)) > 0 Then
            mLastLogCheck = Now
            process = True
        End If
        If Not process Then Exit Sub
        If _OldExceptionLogDayMax > 0 Then
            For Each file As String In My.Computer.FileSystem.GetFiles(Me.ExceptionPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
                If DateDiff(DateInterval.Day, My.Computer.FileSystem.GetFileInfo(file).LastWriteTime, Now) > _OldExceptionLogDayMax Then
                    delList.Add(file)
                End If
            Next
        End If
        If _OldLogDayMax > 0 Then

            Dim logFolderPath = Me.LogPath

            If Not alternateLogFolder Is Nothing Then
                'if using alternate folders, clean those up
                If Not String.IsNullOrEmpty(alternateLogFolder.FolderName) Then
                    logFolderPath = Path.Combine(Me.LogPath, alternateLogFolder.FolderName)
                End If
            End If

            For Each file As String In My.Computer.FileSystem.GetFiles(logFolderPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
                If DateDiff(DateInterval.Day, My.Computer.FileSystem.GetFileInfo(file).LastWriteTime, Now) > _OldLogDayMax Then
                    delList.Add(file)
                End If
            Next
        End If
        If delList.Count = 0 Then Exit Sub
        For Each File As String In delList
            Try
                My.Computer.FileSystem.DeleteFile(File, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                If File.ToLower.Contains(".txt") Then
                    Me.TraceLog(TraceLevel.Verbose, "Log File " & File & " deleted as it is over " & _OldLogDayMax & " days old.")
                ElseIf File.ToLower.Contains(".xml") Then
                    Me.TraceLog(TraceLevel.Verbose, "Exception Log File " & File & " deleted as it is over " & _OldLogDayMax & " days old.")
                End If
            Catch ex As Exception
                Me.TraceLog(TraceLevel.Error, "failed to delete log file: " & File & " Error: " & ex.Message)
            End Try
        Next
    End Sub
    Public Sub Close()
        'do any cleanup here!
    End Sub
    ''' <summary>
    ''' Creates a standard directory structure underneath application.  The folders are Exceptions, Logs, Files and Plugins.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateDirectoryStructure()
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Exceptions"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Logs"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Files"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Plugins"))

    End Sub
    ''' <summary>
    ''' 'returns current time/date in format "Thu Jan 21 15:01:55 EST 2010"
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FileShortDateStamp() As String
        Dim d As Date = Now
        Dim u As TzTimeZone = TzTimeZone.CurrentTimeZone
        Dim t As Date = Now
        Dim sb As New Text.StringBuilder
        sb.Append(Format(d, "ddd"))
        sb.Append(" ")
        sb.Append(Format(d, "MMM"))
        sb.Append(" ")
        sb.Append(Format(d, "dd"))
        sb.Append(" ")
        sb.Append(Format(d, "HH:MM:ss"))
        sb.Append(" ")
        sb.Append(u.GetAbbreviation)
        sb.Append(" ")
        sb.Append(Format(d, "yyyy"))
        Return sb.ToString
    End Function
    ''' <summary>
    ''' Checks for and creates the standard Application directories under installpath: Exceptions, Logs, Files, Plugins
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CheckDirectoryStructure()
        Dim path As String

        If mApplicationType = BaseServices.ApplicationType.NoConfigConsole Then
            path = IO.Path.Combine(mRootPath, "Exceptions")
            If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
            path = IO.Path.Combine(mRootPath, "Logs")
            If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
            path = IO.Path.Combine(mRootPath, "Files")
            If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
            path = IO.Path.Combine(mRootPath, "PlugIns")
            If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
            Exit Sub
        End If
        path = IO.Path.Combine(mRootPath, "Exceptions")
        If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
        path = IO.Path.Combine(mRootPath, "Logs")
        If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
        path = IO.Path.Combine(mRootPath, "Files")
        If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)
        path = IO.Path.Combine(mRootPath, "PlugIns")
        If Not My.Computer.FileSystem.DirectoryExists(path) Then My.Computer.FileSystem.CreateDirectory(path)

    End Sub
    ''' <summary>
    ''' Serializes an AcomException as an xml document in the Exceptions folder.
    ''' </summary>
    ''' <param name="ex">AmcomException class</param>
    ''' <remarks></remarks>
    Public Sub ExceptionLog(ByVal ex As AmcomException)
        If ex Is Nothing Then Throw New ArgumentNullException("ex")
        TraceLog(TraceLevel.Error, "EXCEPTION " & ex.Message)
        WriteException(ex)
    End Sub
    ''' <summary>
    ''' Displays an end-user dialog indicating a serious problem, with exception data.
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <remarks></remarks>
    Public Sub ExceptionShow(ByVal ex As AmcomException)
        If ex Is Nothing Then Throw New ArgumentNullException("ex")
        Dim f As New ExceptionPublisherForm(ex)
        f.ShowDialog()
    End Sub
    ''' <summary>
    ''' returns the instance of this Application
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared Function GetInstance() As Application
        mMutex.WaitOne()
        Try
            If mInstance Is Nothing Then
                mInstance = New Application
            End If
        Finally
            mMutex.ReleaseMutex()
        End Try
        Return mInstance
    End Function
    ''' <summary>
    ''' Returns a consistent root path no matter how application is running.  Requires Registry setting.
    ''' </summary>
    ''' <param name="appName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRootPath(ByVal appName As String) As String
        Using softKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software")
            Using SDCSolutionsKey As RegistryKey = softKey.OpenSubKey("Amcom Software")
                If SDCSolutionsKey Is Nothing Then
                    Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                End If
                Using appKey As RegistryKey = SDCSolutionsKey.OpenSubKey(appName)
                    If appKey Is Nothing Then
                        Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                    End If
                    Dim rootPath As String = CStr(appKey.GetValue("RootPath"))
                    If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
                    If rootPath.Length = 0 Then
                        Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                    End If
                    Return rootPath
                End Using
            End Using
        End Using
    End Function
    Private Sub GuiUnhandledExceptionHandler(ByVal Sender As Object, ByVal args As ThreadExceptionEventArgs)
        HandleUnhandledException(args.Exception)
    End Sub
    ''' <summary>
    ''' Encrypts a string using Rijndael and base64
    ''' </summary>
    ''' <param name="password"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HidePassword(ByVal password As String) As String
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As Encryption.Data
        encryptedData = sym.Encrypt(New Encryption.Data(password), key)
        Dim s As String = encryptedData.ToBase64
        Return s
    End Function
    ''' <summary>
    ''' Decryts a string from Rijndael / base64
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnPassword(ByVal data As String) As String
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As New Encryption.Data
        encryptedData.Base64 = data
        Dim decryptedData As Encryption.Data
        decryptedData = sym.Decrypt(encryptedData, key)
        Return decryptedData.ToString
    End Function
    ''' <summary>
    ''' Displays an error message box indicating an error in a Windows service.
    ''' </summary>
    ''' <param name="title"></param>
    ''' <param name="message"></param>
    ''' <param name="timeOutSeconds"></param>
    ''' <remarks></remarks>
    Public Sub ShowServiceMessageError(ByVal title As String, ByVal message As String, ByVal timeOutSeconds As Integer)
        ShowServiceMessage(title, message, MessageBoxIcon.Exclamation, timeOutSeconds)
    End Sub
    ''' <summary>
    ''' Displays a message box indicating an error in a Windows service.
    ''' </summary>
    ''' <param name="title"></param>
    ''' <param name="message"></param>
    ''' <param name="timeOutSeconds"></param>
    ''' <remarks></remarks>
    Public Sub ShowServiceMessage(ByVal title As String, ByVal message As String, ByVal icon As MessageBoxIcon, ByVal timeOutSeconds As Integer)
        Dim t As New Thread(AddressOf MessageFromService)
        Dim s() As String = {title, message, CStr(icon), CStr(timeOutSeconds)}
        t.Start(s)
    End Sub
    Private Sub MessageFromService(ByVal o As Object)
        Dim waitTime As Integer = 0
        Dim s() As String = o
        Dim timeOutSeconds As Integer = CInt(s(3))
        Dim t As New Thread(AddressOf ServiceMessagedisplayer)
        t.Start(o)
        Do While waitTime < timeOutSeconds
            Thread.Sleep(1000)
            waitTime += 1
        Loop
        t.Abort()
    End Sub
    Private Sub ServiceMessagedisplayer(ByVal o As Object)
        Dim s() As String = o
        Dim icon As MessageBoxIcon = CType(s(2), MessageBoxIcon)
        MessageBox.Show(s(1), s(0), MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
    End Sub
    Private Sub HandleUnhandledException(ByVal ex As Exception)
        _loggedExceptionCountToday += 1
        If mVisualStudioHosted AndAlso (mShowExceptionDialogInVStudio = False) Then Exit Sub
        ExceptionLog(New AmcomException("Unhandled exception", ex))
        Select Case mApplicationType
            Case BaseServices.ApplicationType.Windows, BaseServices.ApplicationType.Console
                Dim f As New ExceptionPublisherForm(ex)
                f.ShowDialog()
                Environment.Exit(1)
            Case BaseServices.ApplicationType.WindowsService
                Dim el As New EventLog()
                Try
                    'Register the App as an Event Source
                    If Not EventLog.SourceExists(My.Application.Info.AssemblyName) Then
                        EventLog.CreateEventSource(My.Application.Info.AssemblyName, "Application")
                    End If
                    el.Source = My.Application.Info.AssemblyName

                    el.WriteEntry("An unexpected exception has occurred in the service.  The details are below:" & vbCrLf & vbCrLf & ex.ToString, System.Diagnostics.EventLogEntryType.Error)
                Catch elEx As Exception
                    'nothing we can do!
                End Try
                Environment.Exit(1)
            Case BaseServices.ApplicationType.Console
            Case BaseServices.ApplicationType.NoConfigConsole
                Console.WriteLine("-- An unhandled exception has occurred in your application --")
                Console.WriteLine("The reported exception is: " & ex.Message)
                Console.WriteLine("The stack trace is shown below.")
                Console.WriteLine(ex.StackTrace)
                Console.WriteLine()
                Console.WriteLine("-The application will terminate-")
                Console.WriteLine("Press any key to continue")
                Console.ReadKey()
            Case Else
        End Select
    End Sub
    ''' <summary>
    ''' Validates a passed string to see if the value is a valid IP Address or DNS name.
    ''' </summary>
    ''' <param name="address"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsIPorDNS(ByVal address As String) As Boolean
        Dim r As New Text.RegularExpressions.Regex("\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
        Dim r2 As New Text.RegularExpressions.Regex("(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)")
        If Not r.IsMatch(address) AndAlso Not r2.IsMatch(address) Then
            Return False
        End If
        Return True
    End Function
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize()
        Initialize(BaseServices.ApplicationType.NoConfigConsole, Nothing)
    End Sub
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase)
        Initialize(appType, settings, "", ApplicationInstanceType.multiInstance)
    End Sub
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase, ByVal instanceType As ApplicationInstanceType)
        Initialize(appType, settings, "", instanceType)
    End Sub
    ''' <summary>
    ''' Initializes application, setting Unhandled Exception handler, default paths, and reads configuration,  A registry entry, located at HKEY_CURRENT_MACHINE\Amcom Software\APP_NAME, must be a string called "RootPath" with the path to the Install.  
    ''' </summary>
    ''' <param name="appType">Type of Application</param>
    ''' <param name="settings">Use My.Settings</param>
    ''' <param name="initializeFromApplicationDomain">Optional, domain name of application settings</param>
    ''' <param name="instanceType">single or multi-instance application</param>
    ''' <remarks>An exception will be thrown immediately if the Rootpath is not available.</remarks>
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase, ByVal initializeFromApplicationDomain As String, ByVal instanceType As ApplicationInstanceType)
        If mIsInitialized Then Exit Sub 'prevent the application from being reinitialized by applications compiled as dll's
        If instanceType = ApplicationInstanceType.singleInstanceOnly Then
            'By including Global\ in the prefix, will work across RDP sessions.
            _AppMutex = New Mutex(False, "Global\" & AppDomain.CurrentDomain.FriendlyName)

            If _AppMutex.WaitOne(0, False) = False Then
                _AppMutex.Close()
                _AppMutex = Nothing
                Throw New AppAlreadyRunningException(AppDomain.CurrentDomain.FriendlyName & " is already running on system " & My.Computer.Name)
            End If
            'if you get to this point it's frist instance
            'continue with app

        End If

        _appName = My.Application.Info.ProductName
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionHandler

        'Sanity Check -- Can we write file system?
        Dim checkPath As String = ""
        Try
            checkPath = IO.Path.Combine(My.Application.Info.DirectoryPath, "sanity.tmp")
            My.Computer.FileSystem.WriteAllText(checkPath, "test file system permissions", False)
        Catch ex As Exception
            Select Case appType
                Case BaseServices.ApplicationType.NoConfigConsole
                    Console.WriteLine("--Fatal Error Initializing Application --")
                    Console.WriteLine(My.Application.Info.AssemblyName & "  does not have rights to write the file system.")
                    Console.WriteLine("Error reported during write attempt: " & ex.Message)
                    Console.WriteLine("Check:  User Access Control (UAC), user at install time`")
                    Console.WriteLine("--Application will terminate--")
                    Console.WriteLine("Press any key to continue")
                    Console.ReadKey()
                    Environment.Exit(-1)
                Case BaseServices.ApplicationType.Windows, BaseServices.ApplicationType.Console
                    Throw New Exception(My.Application.Info.AssemblyName & "  does not have rights to write the file system. Check installation/Installation User.")
                    System.Windows.Forms.Application.Exit()
                Case BaseServices.ApplicationType.WindowsService
                    Dim el As New EventLog()
                    Try
                        'Register the App as an Event Source
                        If Not EventLog.SourceExists(My.Application.Info.AssemblyName) Then
                            EventLog.CreateEventSource(My.Application.Info.AssemblyName, "Application")
                        End If
                        el.Source = My.Application.Info.AssemblyName

                        el.WriteEntry(My.Application.Info.AssemblyName & "  does not have rights to write the file system.  Error reported during write attempt was: " & ex.Message & ".  Check User Access Control (UAC); Check user at installation time.", System.Diagnostics.EventLogEntryType.Error)
                    Catch elEx As Exception
                        'nothing we can do!
                    End Try
                    'This is exteme, but forces service to stop no matter what.
                    Process.GetCurrentProcess.Kill()
            End Select
        End Try
        If My.Computer.FileSystem.FileExists(checkPath) Then My.Computer.FileSystem.DeleteFile(checkPath)
        If initializeFromApplicationDomain.Length > 0 Then
            Select Case appType
                Case BaseServices.ApplicationType.Console, BaseServices.ApplicationType.Windows, BaseServices.ApplicationType.WindowsService
                    _alternameSettingsDomain = initializeFromApplicationDomain
                Case Else
                    Throw New ArgumentException("initializeFromApplicationsDomain can only be used on Console, Windows or Windows Service applications.")
            End Select
        End If
        mApplicationType = appType
        mSettings = settings
        Me.SendUDPLogPackets = False
        Select Case mApplicationType
            Case ApplicationType.Console
                InitializeForStandardDeployment()
            Case ApplicationType.Windows
                AddHandler Windows.Forms.Application.ThreadException, AddressOf GuiUnhandledExceptionHandler
                If ApplicationDeployment.IsNetworkDeployed Then
                    InitializeForNetworkDeployment()
                Else
                    InitializeForStandardDeployment()
                End If
            Case ApplicationType.WindowsService
                InitializeForStandardDeployment()
            Case BaseServices.ApplicationType.NoConfigConsole, BaseServices.ApplicationType.NoConfigComponent
                InitializeForNoConfigDeployment()

            Case Else
                Throw New AmcomException("Unexpected application type found.")
        End Select
        mIsInitialized = True
        'Dim authorizeFirewall As Boolean = False
        'Try
        '    authorizeFirewall = App.Config.GetBoolean(ConfigSection.configuration, "openFirewall", False)
        'Catch
        'End Try
        'If authorizeFirewall Then
        '    If Me.FullApplicationPath.Length > 0 Then
        '        If AuthorizeApplication(My.Application.Info.ProductName, Me.FullApplicationPath, NET_FW_SCOPE_.NET_FW_SCOPE_ALL, NET_FW_IP_VERSION_.NET_FW_IP_VERSION_ANY) Then
        '            App.TraceLog(TraceLevel.Info, "Application was authorized to open firewall")
        '        Else
        '            App.TraceLog(TraceLevel.Info, "Application FAILED authorization to open firewall")

        '        End If
        '    End If
        'End If
    End Sub

    Private Sub InitializeForNetworkDeployment()
        mName = ApplicationDeployment.CurrentDeployment.UpdatedApplicationFullName
        'mName = GetLeftItem(mName, "#")
        'mName = GetRightItem(mName, "/")
        mName = mName.Replace(".application", "")
        mRootPath = ApplicationDeployment.CurrentDeployment.DataDirectory
        If ApplicationDeployment.CurrentDeployment.IsFirstRun Then
            RaiseEvent FirstRun(Me, New EventArgs)
            CreateDirectoryStructure()
        End If
        _fullApplicationPath = IO.Path.Combine(ApplicationDeployment.CurrentDeployment.DataDirectory, mName & ".exe")
        mConfig = ConfigFile.GetInstance(ApplicationDeployment.CurrentDeployment.DataDirectory & "\App.xml")
        mConfig.LoadSettings()
        _settingsLoaded = True
        mVersion = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
    End Sub
    Private Sub InitializeForNoConfigDeployment()
        mConfig = Nothing
        mName = My.Application.Info.AssemblyName
        mRootPath = My.Application.Info.DirectoryPath
        CheckDirectoryStructure()
        _settingsLoaded = False
    End Sub
    Private Sub InitializeForStandardDeployment()
        If _alternameSettingsDomain.Length > 0 Then
            mName = _alternameSettingsDomain
        Else
            mName = AppDomain.CurrentDomain.FriendlyName
        End If
        mName = mName.Replace(".exe", "")
        If mName.IndexOf(".vshost") > 0 Then
            mVisualStudioHosted = True
            mName = mName.Replace(".vshost", "")
            Try
                mRootPath = GetRootPath(mName)
            Catch ex As AmcomException
                'did not find root path entry... Set root to app directory.
                mRootPath = My.Application.Info.DirectoryPath
                _usingAppPathAsBase = True
            End Try
            _fullApplicationPath = IO.Path.Combine(mRootPath, mName & ".exe")
            Dim configPath As String = System.Windows.Forms.Application.StartupPath.Replace("bin\Debug", "")
            configPath = configPath.Replace("bin\Release", "")

            Dim cf As String = IO.Path.Combine(configPath, "App.xml")
            If Not IO.File.Exists(cf) Then
                If _usingAppPathAsBase Then
                    Throw New AmcomException(String.Format("Development Environment detected.  For proper {1} Initialization, an App.xml file is required. No App.xml found at path '{0}'.  Please create an App.xml file and place in path at '{0}'", IO.Path.GetDirectoryName(cf), mApplicationType.ToString))

                Else
                    Throw New AmcomException(String.Format("No configuration file found at '{0}'", IO.Path.GetDirectoryName(cf)))
                End If
            End If
            mConfig = ConfigFile.GetInstance(cf)
        Else
            Try
                mRootPath = GetRootPath(mName)
            Catch ex As AmcomException
                'did not find root path entry... Set root to app directory.
                mRootPath = My.Application.Info.DirectoryPath
                _usingAppPathAsBase = True
            End Try
            mVisualStudioHosted = False
            Dim cf As String = IO.Path.Combine(IO.Path.Combine(mRootPath, "Config"), "App.xml")
            If Not IO.File.Exists(cf) Then
                If _usingAppPathAsBase Then
                    Throw New AmcomException(String.Format("No RootPath entry found in registry, so the default application path is used.  For proper {1} Initialization, an App.xml file is required. No App.xml found at path '{0}'.  Please create an App.xml file and place in path at '{0}'", IO.Path.GetDirectoryName(cf), mApplicationType.ToString))
                Else
                    Throw New AmcomException(String.Format("For proper {1} Initialization, an App.xml file is required. No App.xml found at path '{0}'.  Please create an App.xml file and place in path at '{0}'", IO.Path.GetDirectoryName(cf), mApplicationType.ToString))
                End If
            End If
            mConfig = ConfigFile.GetInstance(cf)
        End If
        CheckDirectoryStructure()
        mConfig.LoadSettings()
        _settingsLoaded = True
        mVersion = App.Config.GetString(ConfigSection.configuration, "version", "")
        _syslogEnabled = App.Config.SysLogEnabled
        _syslogServer = App.Config.SyslogServer
    End Sub

    Private Sub InitializeForWebDeployment()
        _fullApplicationPath = ""

        mRootPath = ""
        mConfig = Nothing
        mVersion = ""
        _syslogEnabled = App.Config.SysLogEnabled
        _syslogServer = App.Config.SyslogServer
    End Sub
    ''' <summary>
    ''' Depreciated. 
    ''' </summary>
    ''' <remarks>Do Not use</remarks>
    Private Sub StartUDPLogger()
        'mUDPLogger.Run()
    End Sub
    ''' <summary>
    ''' Utility function to wrap a string at a specific length
    ''' </summary>
    ''' <param name="text">content</param>
    ''' <param name="margin">maximum width in characters</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function WordWrap(ByVal text As String, ByVal margin As Integer) As String
        'Efficient word wrap for strings.
        Dim start As Integer = 0
        Dim [end] As Integer = 0
        Dim lines As New Text.StringBuilder
        text = System.Text.RegularExpressions.Regex.Replace(text, "\s", " ").Trim()

        While ((start + margin) < text.Length)
            [end] = start + margin
            While (text.Substring([end], 1) <> " " AndAlso [end] > start)
                [end] -= 1
            End While
            If ([end] = start) Then
                [end] = start + margin
            End If
            lines.Append(text.Substring(start, [end] - start))
            lines.Append(vbCrLf)
            start = [end] + 1
        End While

        If (start < text.Length) Then
            lines.Append(text.Substring(start))
            lines.Append(vbCrLf)
        End If
        Return lines.ToString.Replace(vbCrLf & vbCrLf, vbCrLf)
    End Function

    Private Function GetThreadName() As String
        Dim threadName As String = Threading.Thread.CurrentThread.Name
        Dim sb As New Text.StringBuilder
        sb.Append("[")
        If threadName Is Nothing Then
            sb.Append(Thread.CurrentThread.ManagedThreadId)
        Else
            If threadName.Length = 0 Then
                sb.Append(Thread.CurrentThread.ManagedThreadId)
            Else
                sb.Append(threadName)
            End If
        End If
        sb.Append("]")
        Return sb.ToString
    End Function
    Private Function FormatForConsoleLog(ByVal tl As TraceLevel, ByVal message As String) As String ' ByVal stackframe As Diagnostics.StackFrame) As String
        Dim timeStamp As String = Now.ToString("yyyy-MM-dd_HH:mm:ss.fff")
        Return timeStamp & " " & GetThreadName() & " " & message
    End Function
    Private Function FormatForSysLog(ByVal tl As TraceLevel, ByVal message As String) As String ', ByVal stackframe As Diagnostics.StackFrame) As String
        Return GetThreadName() & " " & message
    End Function
    Private Function GetLogName(ByVal alternateLogFolder As AlternateLogFolder) As String
        Dim baseName As String = ""
        Dim baseExt As String = ""
        Dim namePath As String = ""
        Dim xmlFileExtension = ".Log.xml"
        Dim flatFileExtension = ".Log.txt"
        Dim legacyExtension = ".txt"

        Dim count As Integer = 0
        Dim newbase As String = ""

        Dim altFolder As String
        Dim altFile As String
        Dim altFolderPath As String
        Dim altCurrentLogFilename As String

        If _maxLogSize = 0 Then
            _maxLogSize = mConfig.GetInteger(ConfigSection.configuration, "diagnostics/logfile/maxLogSizeMb", 5)
            _maxLogSize = _maxLogSize * 1000000
        End If

        'Check file extension against format
        If Not String.IsNullOrEmpty(_currentLogFilename) Then
            Select Case mConfig.LogFileTraceFormat
                Case LogFileTraceFormats.legacy
                    If _currentLogFilename.Length > 4 Then
                        If _currentLogFilename.Substring(_currentLogFilename.Length - 4) <> legacyExtension Then
                            _currentLogFilename = String.Empty
                        End If
                    Else
                        _currentLogFilename = String.Empty
                    End If
                Case LogFileTraceFormats.flatfile
                    If _currentLogFilename.Length > 8 Then
                        If _currentLogFilename.Substring(_currentLogFilename.Length - 8) <> flatFileExtension Then
                            _currentLogFilename = String.Empty
                        End If
                    Else
                        _currentLogFilename = String.Empty
                    End If
                Case LogFileTraceFormats.xml
                    If _currentLogFilename.Length > 8 Then
                        If _currentLogFilename.Substring(_currentLogFilename.Length - 8) <> xmlFileExtension Then
                            _currentLogFilename = String.Empty
                        End If
                    Else
                        _currentLogFilename = String.Empty
                    End If
            End Select
        End If

        If alternateLogFolder Is Nothing Then
            If Not String.IsNullOrEmpty(_currentLogFilename) Then
                'Check and see if size is an issue
                If My.Computer.FileSystem.FileExists(_currentLogFilename) Then
                    If My.Computer.FileSystem.GetFileInfo(_currentLogFilename).Length <= (_maxLogSize) Then
                        If DateDiff(DateInterval.Day, _currentLogDate, Now()) = 0 Then
                            Return _currentLogFilename
                        Else
                            _loggedExceptionCountToday = 0
                        End If
                    End If
                End If
            End If

            Select Case mConfig.LogFileTraceFormat
                Case LogFileTraceFormats.legacy
                    baseName = "Log" & Format(Now, "MMddyyyy")
                    _currentLogDate = Now()
                    baseExt = legacyExtension
                Case LogFileTraceFormats.flatfile
                    baseName = App.Name & "-" & Format(Now, "yyyy-MM-dd")
                    _currentLogDate = Now()
                    baseExt = flatFileExtension
                Case LogFileTraceFormats.xml
                    baseName = App.Name & "-" & Format(Now, "yyyy-MM-dd")
                    _currentLogDate = Now()
                    baseExt = xmlFileExtension
            End Select

            namePath = App.LogPath

            Do

                If count = 0 Then
                    newbase = baseName
                Else
                    newbase = baseName & "-" & count
                End If
                Dim temp As String = Path.Combine(namePath, newbase & baseExt)
                If My.Computer.FileSystem.FileExists(temp) Then
                    If My.Computer.FileSystem.GetFileInfo(temp).Length > (_maxLogSize) Then
                        count += 1
                    Else
                        _currentLogFilename = temp
                        Return temp
                    End If
                Else
                    _currentLogFilename = temp
                    Return temp
                End If

            Loop
        Else
            'an alternate folder and/or filename is being used

            Select Case mConfig.LogFileTraceFormat
                Case LogFileTraceFormats.legacy
                    baseName = "Log" & Format(Now, "MMddyyyy")
                    baseExt = legacyExtension
                Case LogFileTraceFormats.flatfile
                    baseName = App.Name & "-" & Format(Now, "yyyy-MM-dd")
                    baseExt = flatFileExtension
                Case LogFileTraceFormats.xml
                    baseName = App.Name & "-" & Format(Now, "yyyy-MM-dd")
                    baseExt = xmlFileExtension
            End Select

            altFolder = IIf(String.IsNullOrEmpty(alternateLogFolder.FolderName), String.Empty, alternateLogFolder.FolderName)
            altFile = IIf(String.IsNullOrEmpty(alternateLogFolder.FileNamePrefix), String.Empty, IIf(String.IsNullOrEmpty(altFolder), alternateLogFolder.FileNamePrefix, "\" & alternateLogFolder.FileNamePrefix))
            altFile = altFile & "_" & baseName
            altFolderPath = IIf(String.IsNullOrEmpty(altFolder), App.LogPath, Path.Combine(App.LogPath, altFolder))

            If Not My.Computer.FileSystem.DirectoryExists(altFolderPath) Then
                My.Computer.FileSystem.CreateDirectory(altFolderPath)
            End If

            Do
                If count = 0 Then
                    newbase = altFile
                Else
                    newbase = altFile & "-" & count
                End If
                If DateDiff(DateInterval.Day, _currentLogDate, Now()) <> 0 Then
                    _loggedExceptionCountToday = 0
                    _currentLogDate = Now()
                End If

                altCurrentLogFilename = Path.Combine(altFolderPath, newbase & baseExt)
                If My.Computer.FileSystem.FileExists(altCurrentLogFilename) Then
                    If My.Computer.FileSystem.GetFileInfo(altCurrentLogFilename).Length > (_maxLogSize) Then
                        count += 1
                    Else
                        Return altCurrentLogFilename
                    End If
                Else
                    Return altCurrentLogFilename
                End If
            Loop
        End If

        _currentLogFilename = String.Empty
        Return String.Empty
    End Function
    Private Function FormatValues(ByVal values() As Object) As Object()
        Dim l As New List(Of Object)
        For Each o As Object In values
            If TypeOf o Is Byte() Then
                l.Add(HexDump(o))
            Else
                l.Add(o)
            End If
        Next
        Return l.ToArray
    End Function
    ''' <summary>
    ''' Sends a message to the console and/or log file(s)
    ''' </summary>
    ''' <param name="level">Microsoft-standard tracelog level.</param>
    ''' <param name="message">log message content</param>
    ''' <param name="values">Values formatted in message string.  Uses String.Format syntax </param>
    ''' <remarks></remarks>
    <MethodImpl(MethodImplOptions.Synchronized)> _
    Public Sub TraceLog(ByVal level As TraceLevel, ByVal message As String, ByVal ParamArray values() As Object)
        TraceLog(level, message, Nothing, values)
    End Sub

    ''' <summary>
    ''' Sends a message to the console and/or log file(s)
    ''' </summary>
    ''' <param name="level">Microsoft-standard tracelog level.</param>
    ''' <param name="message"></param>
    ''' <param name="alternateLogFolder">an AlternateLogFolder object which identifies folder,stored under ..\Logs.  Using for specific module logging</param>
    ''' <param name="values">Values formatted in message string.  Uses String.Format syntax </param>
    ''' <remarks></remarks>
    <MethodImpl(MethodImplOptions.Synchronized)> _
    Public Sub TraceLog(ByVal level As TraceLevel, ByVal message As String, ByVal alternateLogFolder As AlternateLogFolder, ByVal ParamArray values() As Object)
        Dim logtext As String = String.Empty
        'If (mApplicationType = BaseServices.ApplicationType.NoConfigConsole) OrElse (mApplicationType = BaseServices.ApplicationType.NoConfigComponent) Then
        '    Dim msgbuffer As String = String.Empty
        '    If values.GetUpperBound(0) > -1 Then
        '        message = String.Format(message, FormatValues(values))
        '    End If
        '    logtext = FormatForConsoleLog(level, message) '), stackframe)
        '    msgbuffer = logtext
        '    If mApplicationType = BaseServices.ApplicationType.NoConfigConsole Then
        '        Console.WriteLine(msgbuffer)
        '        Using writer As New StreamWriter(Path.Combine(My.Application.Info.DirectoryPath, My.Application.Info.AssemblyName & ".log.txt"), True)
        '            writer.WriteLine(FormatForConsoleLog(level, message)) ', stackframe))
        '        End Using
        '    Else
        '        RaiseEvent TraceMessage(Me, New TraceMessageEventArgs(level, FormatForConsoleLog(level, message)))
        '    End If
        'End If
        Try
            If mConfig Is Nothing Then Exit Sub 'We have not been initialized yet!

            CheckForOldLogs(alternateLogFolder)  'check for logs to be deleted.

            Dim name As String = App.Name
            If mApplicationType = ApplicationType.Console Then
                If level <= mConfig.ConsoleTraceLevel Then
                    Dim msgbuffer As String = String.Empty
                    If values.GetUpperBound(0) > -1 Then
                        message = String.Format(message, FormatValues(values))
                    End If
                    logtext = FormatForConsoleLog(level, message) '), stackframe)
                    If _maxWordWrapLength = -1 Then
                        _maxWordWrapLength = App.Config.GetInteger(ConfigSection.diagnostics, "console/maxWidth", 0)
                    End If
                    If _maxWordWrapLength > 0 Then
                        msgbuffer = WordWrap(logtext, _maxWordWrapLength)
                    Else
                        msgbuffer = logtext
                    End If
                    Console.WriteLine(msgbuffer)
                End If

            End If
            If App.Config.SysLogEnabled Then
                If App.Config.SyslogServer.Length > 0 Then
                    SyslogSender.Send(App.Config.SyslogServer, level, DateTime.Now, FormatForSysLog(level, message)) ', stackframe))
                End If
            End If

            Dim namePath As String = GetLogName(alternateLogFolder)
            Dim separateNamePath As String = Path.Combine(Path.GetDirectoryName(namePath), GetThreadName() + ".txt")
            Try
                Select Case mConfig.LogFileTraceFormat
                    Case LogFileTraceFormats.legacy, LogFileTraceFormats.flatfile
                        If values.GetUpperBound(0) > -1 Then
                            message = String.Format(message, FormatValues(values))
                        End If
                        If level <= mConfig.LogfileTraceLevel Then
                            SyncLock _traceLogLock  'Attempt to protect the log file across multiple threads
                                If Not My.Computer.FileSystem.FileExists(namePath) Then
                                    Dim sr As New FileStream(namePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read)
                                    sr.Seek(0, SeekOrigin.End)
                                    Using writer As New StreamWriter(sr)
                                        writer.WriteLine(FormatForConsoleLog(TraceLevel.Info, App.Name & " " & My.Application.Info.Version.ToString)) ', stackframe))
                                    End Using
                                End If
                                Dim sr2 As New FileStream(namePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read)
                                sr2.Seek(0, SeekOrigin.End)
                                Using writer As New StreamWriter(sr2)
                                    writer.WriteLine(FormatForConsoleLog(level, message)) ', stackframe))
                                End Using
                            End SyncLock
                        End If
                    Case LogFileTraceFormats.xml
                        If values.GetUpperBound(0) > -1 Then
                            message = String.Format(message, FormatValues(values))
                        End If
                        If Not My.Computer.FileSystem.FileExists(namePath) Then
                            CreateInitalXMLLog(namePath) ', modInfo)
                        End If
                        AppendXMLLog(namePath, level, message)
                End Select
            Catch ex As IOException
                'we cannot call Exception log because we are having problems writing to the trace log
                WriteException(New AmcomException("Unable to write a message to the trace log", ex))
            Catch ex As Exception
            End Try
            If mSendUDPLogging Then
            End If
            logtext = FormatForConsoleLog(level, message) ', stackframe)
            RaiseEvent TraceMessage(Me, New TraceMessageEventArgs(level, logtext))
        Catch ex As Exception
        End Try
    End Sub
    Public Function GetLogTail(ByVal numrows As Integer) As String
        Dim lineCnt As Integer = 1
        Dim lines As New List(Of String)()
        Dim maxLines As Integer = numrows

        If numrows > 200 OrElse numrows < 1 Then Return ""
        Dim logfile As String = GetLogName(Nothing)

        If String.IsNullOrEmpty(logfile) Then
            Return String.Empty
        End If

        Dim br As New BackwardReader(logfile)
        While Not br.SOF
            Dim line As String = br.ReadLine()
            lines.Add(line)
            If lineCnt = maxLines Then
                Exit While
            End If
            lineCnt += 1
        End While
        lines.Reverse()
        Return String.Join(vbCrLf, lines.ToArray())
    End Function
    Private Sub CreateInitalXMLLog(ByVal filename As String)
        Dim ms As New MemoryStream
        Dim doc As New XmlDocument
        Using writer As XmlTextWriter = New XmlTextWriter(ms, Text.Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            writer.WriteStartDocument()
            writer.WriteStartElement(App.Name.Replace(".", ""))
            writer.WriteEndElement()
            writer.Flush()
            ms.Position = 0
            doc.LoadXml(New StreamReader(ms).ReadToEnd)
        End Using
        My.Computer.FileSystem.WriteAllText(filename, doc.OuterXml, True)
        AppendXMLLog(filename, TraceLevel.Info, App.Name & " " & My.Application.Info.Version.ToString)
    End Sub
    Private Sub AppendXMLLog(ByVal fileName As String, ByVal traceLevel As TraceLevel, ByVal message As String)
        Dim doc As New XmlDocument
        doc.Load(fileName)
        Dim component As String = ""

        Dim root As XmlNode = doc.SelectSingleNode(App.Name.Replace(".", ""))
        Dim le As XmlElement = doc.CreateElement("LogEntry")
        Dim ts As XmlElement = doc.CreateElement("timestamp")


        ts.InnerText = Now.ToString("yyyy-MM-dd_HH:mm:ss.fff")
        Dim mt As XmlElement = doc.CreateElement("messageType")
        mt.InnerText = traceLevel.ToString
        Dim msg As XmlElement = doc.CreateElement("message")
        msg.InnerText = message
        root = root.AppendChild(le)
        root.AppendChild(ts)
        root.AppendChild(mt)
        root = root.AppendChild(msg)
        'If Not compInfo.FileName.Contains("BaseServices") Then
        '    Dim ce As XmlElement = doc.CreateElement("fromComponent")
        '    root.AppendChild(ce)
        '    Dim fn As XmlElement = doc.CreateElement("file")
        '    fn.InnerText = compInfo.FullPath
        '    ce.AppendChild(fn)
        '    Dim api As XmlElement = doc.CreateElement("FileVersion")
        '    api.InnerText = compInfo.FileVersion
        '    ce.AppendChild(api)
        '    If compInfo.APIVersion.ToString.Length > 0 Then
        '        Dim dllApi As XmlElement = doc.CreateElement("APIVersion")
        '        dllApi.InnerText = compInfo.APIVersion.ToString
        '        ce.AppendChild(dllApi)
        '    End If
        'End If
        doc.Save(fileName)
    End Sub
    ''' <summary>
    ''' Returns a string indicating the uptime of the application in the format days:hours:minutes:seconds
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpTime() As String
        Dim secs = Math.Abs(DateDiff(DateInterval.Second, Now, mApplicationStartTime))
        Dim days As Integer = secs \ 86400
        secs = secs Mod 86400
        Dim hours As Integer = secs \ 3600
        secs = secs Mod 3600
        Dim mins As Integer = secs \ 60
        secs = secs Mod 60
        Dim sb As New Text.StringBuilder
        If days > 0 Then
            sb.Append(Format(days, "000"))
            sb.Append(" d ")
        End If
        sb.Append(Format(hours, "00"))
        sb.Append(":")
        sb.Append(Format(mins, "00"))
        sb.Append(":")
        sb.Append(Format(secs, "00"))
        Return sb.ToString
    End Function

    Private Sub UnhandledExceptionHandler(ByVal Sender As Object, ByVal args As UnhandledExceptionEventArgs)
        Dim ex As Exception = CType(args.ExceptionObject, Exception)
        HandleUnhandledException(ex)
    End Sub

    Public Sub WriteException(ByVal ex As AmcomException)
        If mApplicationType = BaseServices.ApplicationType.NoConfigConsole OrElse mApplicationType = BaseServices.ApplicationType.NoConfigComponent Then Exit Sub 'not initialized yet!
        Dim exPath As String = String.Empty
        If mConfig Is Nothing Then
            'an exception has occurred at initialization.
            exPath = Path.Combine(My.Application.Info.DirectoryPath, "Exceptions")
        Else
            exPath = Me.ExceptionPath
        End If
        If Not Directory.Exists(exPath) Then
            Directory.CreateDirectory(exPath)
        End If
        exPath = Path.Combine(exPath, ex.Id & ".xml")
        Using stream As FileStream = New FileStream(exPath, FileMode.Create)
            Dim context As StreamingContext = New StreamingContext(StreamingContextStates.File)
            Dim formatter As SoapFormatter = New SoapFormatter(Nothing, context)
            formatter.Serialize(stream, ex)
        End Using
    End Sub

    Private Sub WriteRegistryKeyForApplication()
        Using SDCSolutionsKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\" & mCompany, True)
            Using appKey As RegistryKey = SDCSolutionsKey.OpenSubKey(mName)
                If appKey Is Nothing Then
                    SDCSolutionsKey.CreateSubKey(mName)
                End If
            End Using
        End Using
    End Sub

    Private Sub WriteRegistryKeyForCompany()
        Using softKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software", True)
            Using SDCSolutionsKey As RegistryKey = softKey.OpenSubKey(mCompany)
                If SDCSolutionsKey Is Nothing Then
                    softKey.CreateSubKey(mCompany)
                End If
            End Using
        End Using
    End Sub

    Private Sub WriteRegistryKeyForRootPath(ByVal rootPath As String)
        WriteRegistryKeyForCompany()
        WriteRegistryKeyForApplication()
        Using appKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\" & mCompany & "\" & mName, True)
            appKey.SetValue("RootPath", rootPath)
        End Using
    End Sub
    Private Function HexDump(ByVal bytes As Byte()) As String
        Return Amcom.SDC.CodeUtilities.Util.DumpFormat.HexDump(bytes)
    End Function

    Private Function HexChar(ByVal value As Integer) As Char
        value = value And &HF
        If value >= 0 AndAlso value <= 9 Then
            Return ChrW(value)
        Else
            Return ChrW(65 + (value - 10))
        End If
    End Function

#End Region

#Region "Event Handlers"
    Private Sub OnSecurityLogWrite(ByVal sender As Object, ByVal e As EntryWrittenEventArgs)
        If e.Entry.InstanceId = 4616 Then
            Dim temp As String() = e.Entry.Message.Split(vbCr)
            Dim previousTime As DateTime = DateTime.MinValue
            Dim newTime As DateTime = DateTime.MinValue

            For Each line As String In temp

                line = line.Replace(vbCrLf, String.Empty)

                If line.Contains("Previous Time:") Then
                    Dim prev As String() = line.Split(":")
                    previousTime = DateTime.Parse(String.Join(":", prev, 1, prev.Count - 1))
                ElseIf line.Contains("New Time:") Then
                    Dim nt As String() = line.Split(":")
                    newTime = DateTime.Parse(String.Join(":", nt, 1, nt.Count - 1))
                End If
            Next

            If previousTime <> DateTime.MinValue AndAlso newTime <> DateTime.MinValue Then
                Dim ts As TimeSpan = newTime - previousTime
                If ts.TotalSeconds <> 0 Then
                    TraceLog(TraceLevel.Verbose, "OnSecurityLogWrite: System time change event occured, old time - {0}, new time - {1}, offset - {2}", previousTime.ToString(), newTime.ToString(), ts.TotalSeconds)
                    For Each d As SystemTimeAdjusted In _staHandlers
                        d(Me, New SystemTimeAdjustedArgs(ts.TotalSeconds, previousTime, newTime))
                    Next
                End If
            End If
        End If

    End Sub
#End Region

    Public Sub New()
        _componentInfoCache = New Dictionary(Of String, ComponentInfo)
        mShowExceptionDialogInVStudio = False
        mApplicationStartTime = Now
        _staHandlers = New List(Of SystemTimeAdjusted)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        If _securityLog IsNot Nothing Then
            RemoveHandler _securityLog.EntryWritten, AddressOf OnSecurityLogWrite
        End If
    End Sub

    Public ReadOnly Property LastExceptionDetails() As String Implements IStatistics.LastExceptionDetails
        Get
            Dim fi As FileInfo = Nothing
            Dim lastDate As Date = DateTime.MinValue
            Dim files As New System.IO.DirectoryInfo(Me.ExceptionPath)
            For Each f As FileInfo In files.GetFiles
                If f.LastWriteTime > lastDate Then
                    fi = f
                    lastDate = f.LastWriteTime
                End If
            Next
            If fi Is Nothing Then
                Return String.Empty
            End If
            Dim info As New Text.StringBuilder
            info.Append(String.Format("This exception occured on {0}, which was {1} ago.", lastDate, lastDate.Since))
            info.Append(vbCrLf)
            info.Append(My.Computer.FileSystem.ReadAllText(Path.Combine(Me.ExceptionPath, fi.Name)))
            Return info.ToString
        End Get
    End Property

    Public ReadOnly Property LoggedUnhandledExceptionCount() As Integer Implements IStatistics.LoggedUnhandledExceptionCount
        Get
            Return New System.IO.DirectoryInfo(Me.ExceptionPath).GetFiles.Count
        End Get
    End Property
    Public ReadOnly Property LoggedUnhandledExceptionCountToday() As Integer Implements IStatistics.LoggedUnhandledExceptionCountToday
        Get
            Return _loggedExceptionCountToday
        End Get
    End Property

    Public ReadOnly Property AppUpTime() As String Implements IStatistics.AppUppTime
        Get
            Return Me.UpTime
        End Get
    End Property
End Class

Public Class SystemTimeAdjustedArgs
    Inherits EventArgs

    Private _TimeBiasSeconds As Double
    Private _NewTime As DateTime
    Private _OldTime As DateTime


    Public Sub New(ByVal timeBiasSeconds As Double, ByVal oldTime As DateTime, ByVal newTime As DateTime)
        _TimeBiasSeconds = timeBiasSeconds
        _OldTime = oldTime
        _NewTime = newTime
    End Sub

    Public ReadOnly Property TimeBiasSeconds() As Double
        Get
            Return _TimeBiasSeconds
        End Get
    End Property

    Public ReadOnly Property OldTime() As DateTime
        Get
            Return _OldTime
        End Get
    End Property

    Public ReadOnly Property NewTime() As DateTime
        Get
            Return _NewTime
        End Get
    End Property
End Class

Public Class ApplicationMethods

    'used to implement the singleton pattern so that only one instance 
    'of this object is ever alive in the application
    Private Shared mInstance As ApplicationMethods
    Private Shared mMutex As New Mutex

    Private mApplicationType As ApplicationType           'the type of application
    Private ReadOnly mCompany As String = "Amcom Software"  'the name of our company
    Private mConfig As ConfigFile                         'the applications configuration file
    Private mIsInitialized As Boolean                     'has the application been initialized
    Private mName As String                               'name of te application
    Private mRootPath As String
    Private mSettings As ApplicationSettingsBase      'dll's can use
    Private mUserId As Integer
    Private mUserIsAdministrator As Boolean
    Private mUserName As String
    Private mVersion As String 'version
    Private mVisualStudioHosted As Boolean            'is the application being debugged by visual studio
    Private mUDPLogger As UDPLogger
    Private mSendUDPLogging As Boolean
    Private mShowExceptionDialogInVStudio As Boolean
    Public Event FirstRun As EventHandler(Of EventArgs)
    Public Event TraceMessage As EventHandler(Of TraceMessageEventArgs)
    Private ReadOnly mKey As String = "474AC74F-EE99-4bbf-ADF3-7BE96F9749E6"
    Private ReadOnly mKey2 As String = "5DF8E51C-A699-4a5d-990C-1BB5C9DAF712"
    Private mApplicationStartTime As Date = Nothing  'used to measure the running time of any application.  See Uptime()
    Private mLastLogCheck As Object = Nothing 'checking for old logs
    Private _OldLogDayMax As Integer = 0
    Private _settingsLoaded As Boolean = False
    Private _componentInfoCache As Dictionary(Of String, ComponentInfo)
    Private _alternameSettingsDomain As String = ""
    Private _syslogServer As String = ""
    Private _syslogEnabled As Boolean = False
    Private _fullApplicationPath As String = ""
    'FireWall COM interface Constants
    Private Const CLSID_FIREWALL_MANAGER As String = "{304CE942-6E39-40D8-943A-B913C40C9CD4}"
    ' ProgID for the AuthorizedApplication object
    Private Const PROGID_AUTHORIZED_APPLICATION As String = "HNetCfg.FwAuthorizedApplication"
#Region "Properties"

    Public ReadOnly Property ComponentInfoCache() As IDictionary(Of String, ComponentInfo)
        Get
            Return _componentInfoCache
        End Get
    End Property
    Public ReadOnly Property ApplicationType() As ApplicationType
        Get
            Return mApplicationType
        End Get
    End Property

    Public ReadOnly Property Config() As ConfigFile
        Get
            Return mConfig
        End Get
    End Property

    Public ReadOnly Property ConnectionString(ByVal applicationDatabase As String) As String
        Get
            Select Case ApplicationType
                Case ApplicationType.Console
                    Return WinConnectionInfo.ConnectionString(applicationDatabase)
                Case ApplicationType.Windows
                    Return WinConnectionInfo.ConnectionString(applicationDatabase)
                Case ApplicationType.WindowsService
                    Return WinConnectionInfo.ConnectionString(applicationDatabase)
            End Select
            Return ""
        End Get
    End Property
    Public ReadOnly Property ConnectionString() As String
        Get
            Select Case ApplicationType
                Case ApplicationType.Console
                    Return WinConnectionInfo.ConnectionString
                Case ApplicationType.Windows
                    Return WinConnectionInfo.ConnectionString
                Case ApplicationType.WindowsService
                    Return WinConnectionInfo.ConnectionString
            End Select
            Return ""
        End Get
    End Property

    Public ReadOnly Property ComputerAddress() As String
        Get
            Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Environment.MachineName)
            Dim IpAddr As IPAddress() = ipEntry.AddressList
            Return IpAddr(0).ToString()
        End Get
    End Property

    Public ReadOnly Property ComputerName() As String
        Get
            Return Dns.GetHostName
        End Get
    End Property
    Public ReadOnly Property PluginPath() As String
        Get
            Return Path.Combine(mRootPath, "Plugins")
        End Get
    End Property
    Public ReadOnly Property ConfigurationPath() As String
        Get
            Return Path.Combine(mRootPath, "Config")
        End Get
    End Property
    Public ReadOnly Property ExceptionPath() As String
        Get
            Return Path.Combine(mRootPath, "Exceptions")
        End Get
    End Property


    Public ReadOnly Property LogPath() As String
        Get
            Return Path.Combine(mRootPath, "Logs")
        End Get
    End Property

    Public ReadOnly Property FilePath() As String
        Get
            Return Path.Combine(mRootPath, "Files")
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get
            Return mName
        End Get
    End Property

    Public ReadOnly Property FullApplicationPath() As String
        Get
            Return _fullApplicationPath
        End Get
    End Property
    Public ReadOnly Property RootPath() As String
        Get
            Return mRootPath
        End Get
    End Property
    ''' <summary>
    ''' Obsolete.  Syslog enabled through App.xml, sent by tracelog
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Do not use</remarks>
    Public Property SendUDPLogPackets() As Boolean
        Get
            Return False
        End Get
        Set(ByVal value As Boolean)
            'mSendUDPLogging = value
            'If value = False Then
            '    If mUDPLogger IsNot Nothing Then
            '        mUDPLogger.Stop()
            '    End If
            'Else
            '    mUDPLogger = New UDPLogger
            '    Dim t As New Thread(AddressOf StartUDPLogger)
            '    t.IsBackground = True
            '    t.Start()
            'End If
        End Set
    End Property
    Public ReadOnly Property Settings() As ApplicationSettingsBase
        Get
            Return mSettings
        End Get
    End Property
    Public Property ShowExceptionDialogInVisualStudio() As Boolean
        Get
            Return mShowExceptionDialogInVStudio
        End Get
        Set(ByVal value As Boolean)
            mShowExceptionDialogInVStudio = value
        End Set
    End Property
    Public Property UserId() As Integer
        Get
            Return mUserId
        End Get
        Set(ByVal value As Integer)
            mUserId = value
        End Set
    End Property

    Public Property UserIsAdministrator() As Boolean
        Get
            Return mUserIsAdministrator
        End Get
        Set(ByVal value As Boolean)
            mUserIsAdministrator = value
        End Set
    End Property

    Public Property UserName() As String
        Get
            Return mUserName
        End Get
        Set(ByVal value As String)
            mUserName = value
        End Set
    End Property

    Public ReadOnly Property Version() As String
        Get
            Return mVersion
        End Get
    End Property

    Public ReadOnly Property VisualStudioHosted() As Boolean
        Get
            Return mVisualStudioHosted
        End Get
    End Property

    'Public ReadOnly Property FirewallEnabled() As Boolean
    '    Get
    '        Dim manager As INetFwMgr = GetFirewallManager()
    '        Return manager.LocalPolicy.CurrentProfile.FirewallEnabled
    '    End Get
    'End Property
#End Region

#Region "Methods"




    'Private Function GetFirewallManager() As NetFwTypeLib.INetFwMgr
    '    App.TraceLog(TraceLevel.Error, "FireWall Management disabled in this version")
    '    'Dim objectType As Type = Type.GetTypeFromCLSID(New Guid(CLSID_FIREWALL_MANAGER))
    '    'Return (TryCast(Activator.CreateInstance(objectType), NetFwTypeLib.INetFwMgr))
    'End Function



    'Public Function AuthorizeApplication(ByVal title As String, ByVal applicationPath As String, ByVal scope As NET_FW_SCOPE_, ByVal ipVersion As NET_FW_IP_VERSION_) As Boolean
    '    ' Create the type from prog id
    '    'Dim type__1 As Type = Type.GetTypeFromProgID(PROGID_AUTHORIZED_APPLICATION)
    '    'Dim auth As INetFwAuthorizedApplication = TryCast(Activator.CreateInstance(type__1), INetFwAuthorizedApplication)
    '    'Try
    '    '    auth.Name = title
    '    '    auth.ProcessImageFileName = applicationPath
    '    '    auth.Scope = scope
    '    '    auth.IpVersion = ipVersion
    '    '    auth.Enabled = True
    '    'Catch ex As Exception
    '    '    'fail for any number of reasons... platform issues?
    '    '    Return False
    '    'End Try

    '    'Dim manager As INetFwMgr = GetFirewallManager()
    '    'Try
    '    '    manager.LocalPolicy.CurrentProfile.AuthorizedApplications.Add(auth)
    '    'Catch ex As Exception
    '    '    Return False
    '    'End Try
    '    'Return True

    'End Function



    Private Sub CheckForOldLogs()
        Dim delList As New List(Of String)
        Dim process As Boolean = False
        If Not _settingsLoaded Then Exit Sub
        If _OldLogDayMax = 0 Then
            Dim _OldLogDayMax As Integer = Me.Config.GetInteger(ConfigSection.diagnostics, "deleteLogsAfterDays", 15)
        End If
        If _OldLogDayMax = -1 Then Exit Sub
        If mLastLogCheck Is Nothing Then
            mLastLogCheck = Now
            process = True
        ElseIf Math.Abs(DateDiff(DateInterval.Minute, mLastLogCheck, Now)) > 10 Then
            mLastLogCheck = Now
            process = True
        End If
        If Not process Then Exit Sub

        For Each file As String In My.Computer.FileSystem.GetFiles(Me.LogPath, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
            If DateDiff(DateInterval.Day, My.Computer.FileSystem.GetFileInfo(file).LastWriteTime, Now) > _OldLogDayMax Then
                delList.Add(file)
            End If
        Next
        If delList.Count = 0 Then Exit Sub
        For Each File As String In delList
            Try
                My.Computer.FileSystem.DeleteFile(File, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
                Me.TraceLog(TraceLevel.Verbose, "Log File " & File & " deleted as it is over " & _OldLogDayMax & " days old.")
            Catch ex As Exception
                Me.TraceLog(TraceLevel.Error, "failed to delete log file: " & File & " Error: " & ex.Message)
            End Try
        Next
    End Sub
    Public Sub Close()
        'do any cleanup here!
    End Sub
    Public Sub CreateDirectoryStructure()
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Exceptions"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Logs"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Files"))
        My.Computer.FileSystem.CreateDirectory(Path.Combine(mRootPath, "Plugins"))

    End Sub
    Public Function FileShortDateStamp() As String
        Dim d As Date = Now
        Dim u As TzTimeZone = TzTimeZone.CurrentTimeZone
        Dim t As Date = Now
        Dim sb As New Text.StringBuilder
        sb.Append(Format(d, "ddd"))
        sb.Append(" ")
        sb.Append(Format(d, "MMM"))
        sb.Append(" ")
        sb.Append(Format(d, "dd"))
        sb.Append(" ")
        sb.Append(Format(d, "HH:MM:ss"))
        sb.Append(" ")
        sb.Append(u.GetAbbreviation)
        sb.Append(" ")
        sb.Append(Format(d, "yyyy"))
        Return sb.ToString
    End Function
    Public Sub CheckDirectoryStructure()
        Dim Path As String = IO.Path.Combine(mRootPath, "Exceptions")
        If Not My.Computer.FileSystem.DirectoryExists(Path) Then My.Computer.FileSystem.CreateDirectory(Path)
        Path = IO.Path.Combine(mRootPath, "Logs")
        If Not My.Computer.FileSystem.DirectoryExists(Path) Then My.Computer.FileSystem.CreateDirectory(Path)
        Path = IO.Path.Combine(mRootPath, "Files")
        If Not My.Computer.FileSystem.DirectoryExists(Path) Then My.Computer.FileSystem.CreateDirectory(Path)
        Path = IO.Path.Combine(mRootPath, "PlugIns")
        If Not My.Computer.FileSystem.DirectoryExists(Path) Then My.Computer.FileSystem.CreateDirectory(Path)

    End Sub



    Public Sub ExceptionLog(ByVal ex As AmcomException)
        If ex Is Nothing Then Throw New ArgumentNullException("ex")
        TraceLog(TraceLevel.Error, "EXCEPTION " & ex.Message)
        WriteException(ex)
    End Sub

    Public Sub ExceptionShow(ByVal ex As AmcomException)
        If ex Is Nothing Then Throw New ArgumentNullException("ex")
        Dim f As New ExceptionPublisherForm(ex)
        f.ShowDialog()
    End Sub

    Friend Shared Function GetInstance() As ApplicationMethods
        mMutex.WaitOne()
        Try
            If mInstance Is Nothing Then
                mInstance = New ApplicationMethods
            End If
        Finally
            mMutex.ReleaseMutex()
        End Try
        Return mInstance
    End Function

    Public Function GetRootPath(ByVal appName As String) As String
        Using softKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software")
            Using SDCSolutionsKey As RegistryKey = softKey.OpenSubKey("Amcom Software")
                If SDCSolutionsKey Is Nothing Then
                    Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                End If
                Using appKey As RegistryKey = SDCSolutionsKey.OpenSubKey(appName)
                    If appKey Is Nothing Then
                        Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                    End If
                    Dim rootPath As String = CStr(appKey.GetValue("RootPath"))
                    If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
                    If rootPath.Length = 0 Then
                        Throw New AmcomException("Missing Registry Entry: A RootPath string field at HKEY_LOCAL_MACHINE\SOFTWARE\Amcom Software\" & mName & " is mssing. (Intstallation Error?)")
                    End If
                    Return rootPath
                End Using
            End Using
        End Using
    End Function

    Private Sub GuiUnhandledExceptionHandler(ByVal Sender As Object, ByVal args As ThreadExceptionEventArgs)
        HandleUnhandledException(args.Exception)
    End Sub
    Public Function HidePassword(ByVal password As String) As String
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As Encryption.Data
        encryptedData = sym.Encrypt(New Encryption.Data(password), key)
        Dim s As String = encryptedData.ToBase64
        Return s
    End Function
    Public Function ReturnPassword(ByVal data As String) As String
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As New Encryption.Data
        encryptedData.Base64 = data
        Dim decryptedData As Encryption.Data
        decryptedData = sym.Decrypt(encryptedData, key)
        Return decryptedData.ToString
    End Function
    Public Sub ShowServiceMessageError(ByVal title As String, ByVal message As String, ByVal timeOutSeconds As Integer)
        ShowServiceMessage(title, message, MessageBoxIcon.Exclamation, timeOutSeconds)
    End Sub
    Public Sub ShowServiceMessage(ByVal title As String, ByVal message As String, ByVal icon As MessageBoxIcon, ByVal timeOutSeconds As Integer)
        Dim t As New Thread(AddressOf MessageFromService)
        Dim s() As String = {title, message, CStr(icon), CStr(timeOutSeconds)}
        t.Start(s)

    End Sub
    Private Sub MessageFromService(ByVal o As Object)
        Dim waitTime As Integer = 0
        Dim s() As String = o
        Dim timeOutSeconds As Integer = CInt(s(3))
        Dim t As New Thread(AddressOf ServiceMessagedisplayer)
        t.Start(o)
        Do While waitTime < timeOutSeconds
            Thread.Sleep(1000)
            waitTime += 1
        Loop
        t.Abort()
    End Sub
    Private Sub ServiceMessagedisplayer(ByVal o As Object)
        Dim s() As String = o
        Dim icon As MessageBoxIcon = CType(s(2), MessageBoxIcon)
        MessageBox.Show(s(1), s(0), MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification)
    End Sub
    Private Sub HandleUnhandledException(ByVal ex As Exception)
        If mVisualStudioHosted AndAlso (mShowExceptionDialogInVStudio = False) Then Exit Sub
        ExceptionLog(New AmcomException("Unhandled exception", ex))
        If mApplicationType = ApplicationType.Windows Then
            Dim f As New ExceptionPublisherForm(ex)
            f.ShowDialog()
            Environment.Exit(1)
        Else
            ShowServiceMessage("Service Exception", "An Unhandled Exception has occurred in service " & App.Name & vbCrLf & vbCrLf & ". The service will terminate." & vbCrLf & vbCrLf & "The reported error was: " & vbCrLf & vbCrLf & ex.Message, MessageBoxIcon.Exclamation, 5)
            Environment.Exit(1)
        End If
    End Sub
    Public Function IsIPorDNS(ByVal address As String) As Boolean
        Dim r As New Text.RegularExpressions.Regex("\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
        Dim r2 As New Text.RegularExpressions.Regex("(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)")
        If Not r.IsMatch(address) AndAlso Not r2.IsMatch(address) Then
            Return False
        End If
        Return True
    End Function
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase)
        Initialize(appType, settings, "", ApplicationInstanceType.multiInstance)
    End Sub
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase, ByVal instanceType As ApplicationInstanceType)
        Initialize(appType, settings, "", instanceType)
    End Sub
    ''' <summary>
    ''' Called when starting a project.  Manditory to support app framework.  A registry entry, located at HKEY_CURRENT_MACHINE\SDC Soutions\APP_NAME, must be a string called "RootPath" with the path to the Install.  
    ''' </summary>
    ''' <param name="appType">Type of CopyOfApplication</param>
    ''' <param name="settings">Use My.Settings</param>
    ''' <param name="initializeFromApplicationDomain">Optional, domain name of application settings</param>
    ''' <remarks></remarks>
    <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.ControlAppDomain)> _
    Public Sub Initialize(ByVal appType As ApplicationType, ByVal settings As ApplicationSettingsBase, ByVal initializeFromApplicationDomain As String, ByVal instanceType As ApplicationInstanceType)
        If mIsInitialized Then Exit Sub 'prevent the application from being reinitialized by applications compiled as dll's
        If instanceType = ApplicationInstanceType.singleInstanceOnly Then
            'By including Global\ in the prefix, will work across RDP sessions.
            _AppMutex = New Mutex(False, "Global\" & AppDomain.CurrentDomain.FriendlyName)

            If _AppMutex.WaitOne(0, False) = False Then
                _AppMutex.Close()
                _AppMutex = Nothing
                Throw New AppAlreadyRunningException(AppDomain.CurrentDomain.FriendlyName & " is already running on system " & My.Computer.Name)
            End If
            'if you get to this point it's frist instance
            'continue with app

        End If
        If initializeFromApplicationDomain.Length > 0 Then
            Select Case appType
                Case BaseServices.ApplicationType.Console, BaseServices.ApplicationType.Windows, BaseServices.ApplicationType.WindowsService
                    _alternameSettingsDomain = initializeFromApplicationDomain
                Case Else
                    Throw New ArgumentException("initializeFromApplicationsDomain can only be used on Console, Windows or Windows Service applications.")
            End Select
        End If
        mApplicationType = appType
        mSettings = settings
        Me.SendUDPLogPackets = False
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionHandler
        'TODO: Subclass the CopyOfApplication object to return the correct application obect for the application type
        Select Case mApplicationType
            Case ApplicationType.Console
                InitializeForStandardDeployment()
            Case ApplicationType.Windows
                AddHandler Windows.Forms.Application.ThreadException, AddressOf GuiUnhandledExceptionHandler
                If ApplicationDeployment.IsNetworkDeployed Then
                    InitializeForNetworkDeployment()
                Else
                    InitializeForStandardDeployment()
                End If
            Case ApplicationType.WindowsService
                InitializeForStandardDeployment()
            Case Else
                Throw New AmcomException("Unexpected application type found.")
        End Select
        mIsInitialized = True
        Dim authorizeFirewall As Boolean = App.Config.GetBoolean(ConfigSection.configuration, "openFirewall", False)
        'If authorizeFirewall Then
        '    If Me.FullApplicationPath.Length > 0 Then
        '        If AuthorizeApplication(My.CopyOfApplication.Info.ProductName, Me.FullApplicationPath, NET_FW_SCOPE_.NET_FW_SCOPE_ALL, NET_FW_IP_VERSION_.NET_FW_IP_VERSION_ANY) Then
        '            App.TraceLog(TraceLevel.Info, "CopyOfApplication was authorized to open firewall")
        '        Else
        '            App.TraceLog(TraceLevel.Info, "CopyOfApplication FAILED authorization to open firewall")

        '        End If
        '    End If
        'End If
    End Sub

    Private Sub InitializeForNetworkDeployment()
        mName = ApplicationDeployment.CurrentDeployment.UpdatedApplicationFullName
        'mName = GetLeftItem(mName, "#")
        'mName = GetRightItem(mName, "/")
        mName = mName.Replace(".application", "")
        mRootPath = ApplicationDeployment.CurrentDeployment.DataDirectory
        If ApplicationDeployment.CurrentDeployment.IsFirstRun Then
            RaiseEvent FirstRun(Me, New EventArgs)
            CreateDirectoryStructure()
        End If
        _fullApplicationPath = IO.Path.Combine(ApplicationDeployment.CurrentDeployment.DataDirectory, mName & ".exe")
        mConfig = ConfigFile.GetInstance(ApplicationDeployment.CurrentDeployment.DataDirectory & "\App.xml")
        mConfig.LoadSettings()
        _settingsLoaded = True
        mVersion = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
    End Sub

    Private Sub InitializeForStandardDeployment()
        If _alternameSettingsDomain.Length > 0 Then
            mName = _alternameSettingsDomain
        Else
            mName = AppDomain.CurrentDomain.FriendlyName
        End If
        mName = mName.Replace(".exe", "")
        If mName.IndexOf(".vshost") > 0 Then
            mVisualStudioHosted = True
            mName = mName.Replace(".vshost", "")
            mRootPath = GetRootPath(mName)
            _fullApplicationPath = IO.Path.Combine(mRootPath, mName & ".exe")
            Dim configPath As String = System.Windows.Forms.Application.StartupPath.Replace("bin\Debug", "")
            configPath = configPath.Replace("bin\Release", "")

            mConfig = ConfigFile.GetInstance(configPath & "App.xml")
        Else
            mRootPath = GetRootPath(mName)
            mVisualStudioHosted = False
            mConfig = ConfigFile.GetInstance(mRootPath & "Config\App.xml")
        End If
        CheckDirectoryStructure()
        mConfig.LoadSettings()
        _settingsLoaded = True
        mVersion = App.Config.GetString(ConfigSection.configuration, "version", "")
        _syslogEnabled = App.Config.SysLogEnabled
        _syslogServer = App.Config.SyslogServer
    End Sub

    Private Sub InitializeForWebDeployment()
        _fullApplicationPath = ""

        mRootPath = ""
        mConfig = Nothing
        mVersion = ""
        _syslogEnabled = App.Config.SysLogEnabled
        _syslogServer = App.Config.SyslogServer
    End Sub
    ''' <summary>
    ''' Depreciated.  Not used
    ''' </summary>
    ''' <remarks>Do Not use</remarks>
    Private Sub StartUDPLogger()
        'mUDPLogger.Run()
    End Sub
    Private Function WordWrap(ByVal text As String, ByVal margin As Integer) As String
        'Efficient word wrap for strings.
        Dim start As Integer = 0
        Dim [end] As Integer = 0
        Dim lines As New Text.StringBuilder
        text = System.Text.RegularExpressions.Regex.Replace(text, "\s", " ").Trim()

        While ((start + margin) < text.Length)
            [end] = start + margin
            While (text.Substring([end], 1) <> " " AndAlso [end] > start)
                [end] -= 1
            End While
            If ([end] = start) Then
                [end] = start + margin
            End If
            lines.Append(text.Substring(start, [end] - start))
            lines.Append(vbCrLf)
            start = [end] + 1
        End While

        If (start < text.Length) Then
            lines.Append(text.Substring(start))
            lines.Append(vbCrLf)
        End If
        Return lines.ToString.Replace(vbCrLf & vbCrLf, vbCrLf)
    End Function

    Private Function GetThreadName() As String
        Dim threadName As String = Threading.Thread.CurrentThread.Name
        Dim sb As New Text.StringBuilder
        sb.Append("[")
        If threadName Is Nothing Then
            sb.Append(Thread.CurrentThread.ManagedThreadId)
        Else
            If threadName.Length = 0 Then
                sb.Append(Thread.CurrentThread.ManagedThreadId)
            Else
                sb.Append(threadName)
            End If
        End If
        sb.Append("]")
        Return sb.ToString
    End Function
    'Public Function GetCurrentModuleAndVersion(ByVal stackframe As Diagnostics.StackFrame) As ComponentInfo
    '    If _componentInfoCache Is Nothing Then _componentInfoCache = New Dictionary(Of String, ComponentInfo)
    '    Dim callingmethod As System.Reflection.MethodBase = stackframe.GetMethod
    '    Dim thisComponent As String = callingmethod.Module.FullyQualifiedName

    '    Dim componentInfo As ComponentInfo = Nothing
    '    If _componentInfoCache.ContainsKey(thisComponent) Then
    '        componentInfo = _componentInfoCache(thisComponent)
    '    End If
    '    If componentInfo Is Nothing Then
    '        componentInfo = New ComponentInfo(thisComponent)
    '        _componentInfoCache.Add(thisComponent, componentInfo)
    '    End If
    '    Return componentInfo
    'End Function
    Private Function FormatForConsoleLog(ByVal tl As TraceLevel, ByVal message As String) As String ' ByVal stackframe As Diagnostics.StackFrame) As String
        Dim timeStamp As String = Now.ToString("yyyy-MM-dd_HH:mm:ss.fff")

        Return timeStamp & " " & GetThreadName() & " " & message
        'Select Case mConfig.LogFileTraceFormat
        '    Case LogFileTraceFormats.legacy, LogFileTraceFormats.xml
        '        Return timeStamp & " " & GetThreadName() & " " & message
        '    Case LogFileTraceFormats.flatfile
        '        Return timeStamp & " " & GetThreadName() & " " & message
        '        'Return timeStamp & " " & GetCurrentModuleAndVersion(stackframe).ToString & " " & GetThreadName() & " " & message
        'End Select
        'Return timeStamp & " " & GetThreadName() & " " & message
        'Return timeStamp & " " & GetCurrentModuleAndVersion(stackframe).ToString & " " & GetThreadName() & " " & message
    End Function
    Private Function FormatForSysLog(ByVal tl As TraceLevel, ByVal message As String) As String ', ByVal stackframe As Diagnostics.StackFrame) As String
        Return GetThreadName() & " " & message
        'Return GetCurrentModuleAndVersion(stackframe).ToString & " " & GetThreadName() & " " & message
    End Function

    ''' <summary>
    ''' Sends a message to the console and/or log file(s)
    ''' </summary>
    ''' <param name="level">Microsoft-standard tracelog level.</param>
    ''' <param name="message"></param>
    ''' <param name="values">Values formatted in message string.  Uses String.Format syntax </param>
    ''' <remarks></remarks>
    <MethodImpl(MethodImplOptions.Synchronized)> _
    Public Sub TraceLog(ByVal level As TraceLevel, ByVal message As String, ByVal ParamArray values() As Object)
        TraceLog(level, message, Nothing, values)
    End Sub

    ''' <summary>
    ''' Sends a message to the console and/or log file(s)
    ''' </summary>
    ''' <param name="level">Microsoft-standard tracelog level.</param>
    ''' <param name="message"></param>
    ''' <param name="alternateLogFolder">an AlternateLogFolder object which identifies folder,stored under ..\Logs.  Using for specific module logging</param>
    ''' <param name="values">Values formatted in message string.  Uses String.Format syntax </param>
    ''' <remarks></remarks>
    <MethodImpl(MethodImplOptions.Synchronized)> _
    Public Sub TraceLog(ByVal level As TraceLevel, ByVal message As String, ByVal alternateLogFolder As AlternateLogFolder, ByVal ParamArray values() As Object)
        Dim logtext As String = String.Empty
        Try
            If mConfig Is Nothing Then Exit Sub 'We have not been initialized yet!

            'Dim stackframe As New StackFrame(1)
            'Dim callingmethod As System.Reflection.MethodBase = stackframe.GetMethod
            'If callingmethod.Name = "TraceLog" Then
            '    stackframe = New StackFrame(2)
            '    callingmethod = stackframe.GetMethod
            'End If
            'Dim trace As New StackTrace
            Dim name As String = App.Name
            'Dim name As String = callingmethod.Module.FullyQualifiedName
            ''CheckForOldLogs()  'delete old logs if necessary
            'Dim modInfo As ComponentInfo = GetCurrentModuleAndVersion(stackframe)
            If mApplicationType = ApplicationType.Console Then
                If level <= mConfig.ConsoleTraceLevel Then
                    Dim msgbuffer As String = String.Empty
                    If values.GetUpperBound(0) > -1 Then
                        message = String.Format(message, values)
                    End If
                    logtext = FormatForConsoleLog(level, message) '), stackframe)
                    msgbuffer = WordWrap(logtext, mConfig.MaxConsoleWidth)
                    Console.Write(msgbuffer)
                End If

            End If
            If App.Config.SysLogEnabled Then
                If App.Config.SyslogServer.Length > 0 Then
                    SyslogSender.Send(App.Config.SyslogServer, level, DateTime.Now, FormatForSysLog(level, message)) ', stackframe))
                End If
            End If


            Dim namePath As String
            Try
                Select Case mConfig.LogFileTraceFormat
                    Case LogFileTraceFormats.legacy
                        name = "Log" & Format(Now, "MMddyyyy") & ".txt"
                    Case LogFileTraceFormats.flatfile
                        name = App.Name & "-" & Format(Now, "yyyy-MM-dd") & ".Log.txt"
                    Case LogFileTraceFormats.xml
                        name = App.Name & "-" & Format(Now, "yyyy-MM-dd") & ".Log.xml"
                End Select
                If Not alternateLogFolder Is Nothing Then
                    namePath = Path.Combine(App.LogPath, alternateLogFolder.FolderName)
                    If Not My.Computer.FileSystem.DirectoryExists(namePath) Then
                        My.Computer.FileSystem.CreateDirectory(namePath)
                    End If
                Else
                    namePath = App.LogPath
                End If
                If values.GetUpperBound(0) > -1 Then
                    message = String.Format(message, values)
                End If
                namePath = Path.Combine(namePath, name)
                Select Case mConfig.LogFileTraceFormat
                    Case LogFileTraceFormats.legacy, LogFileTraceFormats.flatfile
                        If Not My.Computer.FileSystem.FileExists(namePath) Then
                            Using writer As New StreamWriter(namePath, True)
                                writer.WriteLine(FormatForConsoleLog(TraceLevel.Info, App.Name & " " & My.Application.Info.Version.ToString)) ', stackframe))
                            End Using
                        End If
                        Using writer As New StreamWriter(namePath, True)
                            writer.WriteLine(FormatForConsoleLog(level, message)) ', stackframe))
                        End Using
                    Case LogFileTraceFormats.xml
                        If Not My.Computer.FileSystem.FileExists(namePath) Then
                            CreateInitalXMLLog(namePath) ', modInfo)
                        End If
                        AppendXMLLog(namePath, level, message)
                End Select
            Catch ex As IOException
                'we cannot call Exception log because we are having problems writing to the trace log
                WriteException(New AmcomException("Unable to write a message to the trace log", ex))
            Catch ex As Exception
            End Try
            If mSendUDPLogging Then
            End If
            logtext = FormatForConsoleLog(level, message) ', stackframe)
            RaiseEvent TraceMessage(Me, New TraceMessageEventArgs(level, logtext))
        Catch ex As Exception
        End Try
    End Sub
    Private Sub CreateInitalXMLLog(ByVal filename As String)
        Dim ms As New MemoryStream
        Dim doc As New XmlDocument
        Using writer As XmlTextWriter = New XmlTextWriter(ms, Text.Encoding.UTF8)
            writer.Formatting = Formatting.Indented
            writer.WriteStartDocument()
            writer.WriteStartElement(App.Name.Replace(".", ""))
            writer.WriteEndElement()
            writer.Flush()
            ms.Position = 0
            Using sr As New StreamReader(ms)
                doc.LoadXml(sr.ReadToEnd)
            End Using
        End Using
        My.Computer.FileSystem.WriteAllText(filename, doc.OuterXml, True)
        AppendXMLLog(filename, TraceLevel.Info, App.Name & " " & My.Application.Info.Version.ToString)
    End Sub
    Private Sub AppendXMLLog(ByVal fileName As String, ByVal traceLevel As TraceLevel, ByVal message As String)
        Dim doc As New XmlDocument
        doc.Load(fileName)
        Dim component As String = ""

        Dim root As XmlNode = doc.SelectSingleNode(App.Name.Replace(".", ""))
        Dim le As XmlElement = doc.CreateElement("LogEntry")
        Dim ts As XmlElement = doc.CreateElement("timestamp")


        ts.InnerText = Now.ToString("yyyy-MM-dd_HH:mm:ss.fff")
        Dim mt As XmlElement = doc.CreateElement("messageType")
        mt.InnerText = traceLevel.ToString
        Dim msg As XmlElement = doc.CreateElement("message")
        msg.InnerText = message
        root = root.AppendChild(le)
        root.AppendChild(ts)
        root.AppendChild(mt)
        root = root.AppendChild(msg)
        'If Not compInfo.FileName.Contains("BaseServices") Then
        '    Dim ce As XmlElement = doc.CreateElement("fromComponent")
        '    root.AppendChild(ce)
        '    Dim fn As XmlElement = doc.CreateElement("file")
        '    fn.InnerText = compInfo.FullPath
        '    ce.AppendChild(fn)
        '    Dim api As XmlElement = doc.CreateElement("FileVersion")
        '    api.InnerText = compInfo.FileVersion
        '    ce.AppendChild(api)
        '    If compInfo.APIVersion.ToString.Length > 0 Then
        '        Dim dllApi As XmlElement = doc.CreateElement("APIVersion")
        '        dllApi.InnerText = compInfo.APIVersion.ToString
        '        ce.AppendChild(dllApi)
        '    End If
        'End If
        doc.Save(fileName)
    End Sub
    Public Function UpTime() As String
        Dim secs = Math.Abs(DateDiff(DateInterval.Second, Now, mApplicationStartTime))
        Dim days As Integer = secs \ 86400
        secs = secs Mod 86400
        Dim hours As Integer = secs \ 3600
        secs = secs Mod 3600
        Dim mins As Integer = secs \ 60
        secs = secs Mod 60
        Dim sb As New Text.StringBuilder
        If days > 0 Then
            sb.Append(Format(days, "000"))
            sb.Append(" d ")
        End If
        sb.Append(Format(hours, "00"))
        sb.Append(":")
        sb.Append(Format(mins, "00"))
        sb.Append(":")
        sb.Append(Format(secs, "00"))
        Return sb.ToString
    End Function

    Private Sub UnhandledExceptionHandler(ByVal Sender As Object, ByVal args As UnhandledExceptionEventArgs)
        Dim ex As Exception = CType(args.ExceptionObject, Exception)
        HandleUnhandledException(ex)
    End Sub

    Public Sub WriteException(ByVal ex As AmcomException)

        If mConfig Is Nothing Then Exit Sub 'not initialized yet!
        If Not Directory.Exists(Me.ExceptionPath) Then
            Directory.CreateDirectory(Me.ExceptionPath)
        End If
        Dim exPath As String = Path.Combine(Me.ExceptionPath, ex.Id & ".xml")
        Using stream As FileStream = New FileStream(exPath, FileMode.Create)
            Dim context As StreamingContext = New StreamingContext(StreamingContextStates.File)
            Dim formatter As SoapFormatter = New SoapFormatter(Nothing, context)
            formatter.Serialize(stream, ex)
        End Using
    End Sub

    Private Sub WriteRegistryKeyForApplication()
        Using SDCSolutionsKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\" & mCompany, True)
            Using appKey As RegistryKey = SDCSolutionsKey.OpenSubKey(mName)
                If appKey Is Nothing Then
                    SDCSolutionsKey.CreateSubKey(mName)
                End If
            End Using
        End Using
    End Sub

    Private Sub WriteRegistryKeyForCompany()
        Using softKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software", True)
            Using SDCSolutionsKey As RegistryKey = softKey.OpenSubKey(mCompany)
                If SDCSolutionsKey Is Nothing Then
                    softKey.CreateSubKey(mCompany)
                End If
            End Using
        End Using
    End Sub

    Private Sub WriteRegistryKeyForRootPath(ByVal rootPath As String)
        WriteRegistryKeyForCompany()
        WriteRegistryKeyForApplication()
        Using appKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\" & mCompany & "\" & mName, True)
            appKey.SetValue("RootPath", rootPath)
        End Using
    End Sub


#End Region

    Public Sub New()
        _componentInfoCache = New Dictionary(Of String, ComponentInfo)
        mShowExceptionDialogInVStudio = False
        mApplicationStartTime = Now
    End Sub
End Class
