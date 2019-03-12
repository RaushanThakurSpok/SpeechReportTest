Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Xml
Imports System.Xml.XPath
Imports Microsoft.Win32
Imports System.Security.Permissions
Imports System.Text

Public Enum ConfigSection
    control
    configuration
    database
    modules
    diagnostics
End Enum
Public Enum LogFileTraceFormats
    legacy
    flatfile
    xml
End Enum

Public Class ConfigFile

    'used to implement the singleton pattern so that only one instance 
    'of this object is ever alive in the application
    Private Shared mInstance As ConfigFile
    Private Shared mMutex As New Mutex

    Private mLastWriteTime As DateTime       'track changes to the configuration file
    Private mPath As String                  'path to the configuration file
    Private mSettings As XmlDocument         'the xml containging the configuration information
    Private mSettingsLock As Object          'lock for all the settings in the document
    Private mConsoleTraceLevel As TraceLevel        'the level of tracing for the app, taken out of Settings document for performance reasons
    Private mLogfileTraceLevel As TraceLevel        'the level of tracing for the app, taken out of Settings document for performance reasons
    Private mTraceLevel As TraceLevel        'the level of tracing for the app, taken out of Settings document for performance reasons
    Private mLogFileTraceFormat As LogFileTraceFormats  'format of log file
    Private mThrowConfigException As Boolean 'If key not found, throw an exception
    Private mSyslogServer As String
    Private msyslogEnabled As Boolean = False
    Private mMaxConsoleWidth As Integer
    Private mKey As String = "A7750BAF716D4fa2B3BFAF03245EE065"
#Region "Constructors"

    <PermissionSet(SecurityAction.Demand, Name:="FullTrust")> _
    Public Sub New(ByVal path As String)
        If File.Exists(path) = False Then
            Throw New AmcomException("The application configuration file does not exist  path: " & path)
        End If
        mPath = path
        mLastWriteTime = New DateTime
        mSettings = New XmlDocument
        mSettingsLock = New Object
        mTraceLevel = Diagnostics.TraceLevel.Info
        mLogfileTraceLevel = Diagnostics.TraceLevel.Info
        mConsoleTraceLevel = Diagnostics.TraceLevel.Info
        mLogFileTraceFormat = LogFileTraceFormats.legacy
        mSyslogServer = ""
        mMaxConsoleWidth = 60
    End Sub

#End Region

#Region "Properties"
    Public ReadOnly Property MaxConsoleWidth() As Integer
        Get
            Return mMaxConsoleWidth
        End Get
    End Property
    Public ReadOnly Property SyslogServer() As String
        Get
            Return mSyslogServer
        End Get
    End Property
    Public ReadOnly Property SysLogEnabled() As Boolean
        Get
            Return msyslogEnabled
        End Get
    End Property
    Public ReadOnly Property TraceLevel() As TraceLevel
        Get
            Return mTraceLevel
        End Get
    End Property
    Public ReadOnly Property ConsoleTraceLevel() As TraceLevel
        Get
            Return mConsoleTraceLevel
        End Get
    End Property
    Public ReadOnly Property LogfileTraceLevel() As TraceLevel
        Get
            Return mLogfileTraceLevel
        End Get
    End Property
    Public ReadOnly Property LogFileTraceFormat() As LogFileTraceFormats
        Get
            Return mLogFileTraceFormat
        End Get
    End Property

    Public Property ThrowConfigurationExceptions() As Boolean
        Get
            Return mThrowConfigException
        End Get
        Set(ByVal value As Boolean)
            mThrowConfigException = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub AddNode(ByVal section As ConfigSection, ByVal xpath As String, ByVal nodeString As String)
        AddNode(section, xpath, nodeString, False)
    End Sub

    Public Sub AddNode(ByVal section As ConfigSection, ByVal xpath As String, ByVal nodeString As String, ByVal forceEncrypt As String)
        SyncLock mSettingsLock
            Dim doc As New XmlDocument
            doc.LoadXml(nodeString)
            Dim p As String = "//configuration/" & section.ToString.ToLower
            If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
            Dim fullpath As String = p
            If xpath.Trim.Length > 0 Then
                fullpath = p & "/" & xpath
            End If
            Dim node As XmlNode = mSettings.SelectSingleNode(fullpath)
            Dim childNode As XmlNode = mSettings.ImportNode(doc.DocumentElement, True)
            Dim newNode As XmlNode = node.AppendChild(childNode)
            If forceEncrypt Then
                Dim attr As XmlAttribute = mSettings.CreateAttribute("encrypt")
                attr.Value = "true"
                newNode.Attributes.Append(attr)
            End If
            mSettings.Save(mPath)
        End SyncLock
    End Sub

    Public Sub CheckForChanges()
        If File.GetLastWriteTime(mPath) <> mLastWriteTime Then
            LoadSettings()
            mLastWriteTime = File.GetLastWriteTime(mPath)
        End If
    End Sub
    Public Sub CheckForConfigurationError(ByVal section As String, ByVal xpath As String)
        If mThrowConfigException Then
            Throw New InvalidConfigurationParameterException(String.Format("Setting {0} not found in section {1} of App.xml.  Check configuration.", xpath, section.ToString))
        End If
    End Sub
    Public Function GetBoolean(ByVal section As ConfigSection, ByVal xpath As String) As Boolean
        Return GetBoolean(section, xpath, False, False)
    End Function
    Public Function GetBoolean(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Boolean) As Boolean
        Return GetBoolean(section, xpath, [default], False)
    End Function
    Public Function GetBoolean(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Boolean, ByVal forceEncrypt As Boolean) As Boolean
        Return GetBoolean(section, xpath, [default], forceEncrypt, False)
    End Function

    Public Function GetBoolean(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Boolean, ByVal forceEncrypt As Boolean, ByVal createIfNotFound As Boolean) As Boolean
        SyncLock mSettingsLock
            Try
                Dim p As String = "//configuration/" & section.ToString.ToLower
                If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
                Dim fullpath As String = p & "/" & xpath
                Dim value As String = mSettings.SelectSingleNode(fullpath).InnerText
                Dim fileEncrypt As Boolean = False
                Dim attr As XmlAttribute = mSettings.SelectSingleNode(fullpath).Attributes.GetNamedItem("encrypt")
                If attr IsNot Nothing Then
                    fileEncrypt = (attr.Value.ToLower = "true") OrElse (attr.Value.ToLower = "yes")
                End If
                Dim isEnc As Boolean = False
                If value.Length >= 4 Then
                    isEnc = (value.Substring(0, 4) = "ENC(")
                End If
                If (forceEncrypt Or fileEncrypt) And Not isEnc Then
                    If Not String.IsNullOrEmpty(value) Then
                        WriteString(section, xpath, value, True)
                    End If
                End If
                Return CBool(DecryptConfigString(value))
            Catch ex As XmlException
                CheckForConfigurationError(section, xpath)
                Return [default]
            Catch ex As XPathException
                CheckForConfigurationError(section, xpath)
                Return [default]
            Catch ex As NullReferenceException
                'CheckForConfigurationError(section, xpath)
                If createIfNotFound Then
                    AddSetting(section, xpath, [default], forceEncrypt)
                End If
                Return [default]
            Catch ex As InvalidCastException
                If createIfNotFound Then
                    AddSetting(section, xpath, [default], forceEncrypt)
                End If
                Return [default]
            End Try
        End SyncLock
    End Function

    Public Shared Function GetInstance(ByVal path As String) As ConfigFile
        mMutex.WaitOne()
        Try
            If mInstance Is Nothing Then
                mInstance = New ConfigFile(path)
            End If
        Finally
            mMutex.ReleaseMutex()
        End Try
        Return mInstance
    End Function

    Public Function GetInteger(ByVal section As ConfigSection, ByVal xpath As String) As Integer
        Return GetInteger(section, xpath, 0, False)
    End Function
    Public Function GetInteger(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Integer) As Integer
        Return GetInteger(section, xpath, [default], False, False)
    End Function

    Public Function GetInteger(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Integer, ByVal createIfNotFound As Boolean) As Integer
        Return GetInteger(section, xpath, [default], createIfNotFound, False)
    End Function

    Public Function GetInteger(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Integer, ByVal createIfNotFound As Boolean, ByVal forceEncrypt As Boolean) As Integer
        SyncLock mSettingsLock
            Try
                Dim p As String = "//configuration/" & section.ToString.ToLower
                If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
				Dim fullpath As String = p & "/" & xpath
				Dim checkNode = mSettings.SelectSingleNode(fullpath)
				If (checkNode Is Nothing) Then
					If createIfNotFound Then
						AddSetting(section, xpath, [default], forceEncrypt)
					End If
					Return [default]
				End If
				Dim value As String = checkNode.InnerText
				Dim fileEncrypt As Boolean = False
				Dim attr As XmlAttribute = mSettings.SelectSingleNode(fullpath).Attributes.GetNamedItem("encrypt")
				If attr IsNot Nothing Then
					fileEncrypt = (attr.Value.ToLower = "true") OrElse (attr.Value.ToLower = "yes")
				End If
				If forceEncrypt Or fileEncrypt Then
					If Not String.IsNullOrEmpty(value) Then
						WriteString(section, xpath, value, True)
					End If
				End If
				Return CInt(DecryptConfigString(value))
			Catch ex As XmlException
				CheckForConfigurationError(section, xpath)
				Return [default]
			Catch ex As XPathException
				CheckForConfigurationError(section, xpath)
				Return [default]
			Catch ex As NullReferenceException
				'CheckForConfigurationError(section, xpath)
				If createIfNotFound Then
					AddSetting(section, xpath, [default], forceEncrypt)
				End If
				Return [default]
			Catch ex As InvalidCastException
				'CheckForConfigurationError(section, xpath)
				If createIfNotFound Then
					AddSetting(section, xpath, [default], forceEncrypt)
				End If
				Return [default]
			End Try
        End SyncLock
    End Function

    Public Function GetNode(ByVal section As ConfigSection, ByVal xpath As String) As XmlNode
        SyncLock mSettingsLock
            Try
                Dim p As String = "//configuration/" & section.ToString.ToLower
                If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
				Dim fullpath As String = p & "/" & xpath
				Dim selectedNode = mSettings.SelectSingleNode(fullpath)

				If (selectedNode Is Nothing) Then
					Return Nothing
				End If

				Return DecryptNode(selectedNode)
            Catch ex As XmlException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As XPathException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As NullReferenceException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As InvalidCastException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            End Try
        End SyncLock
    End Function

    Public Function GetNodeList(ByVal section As ConfigSection, ByVal xpath As String) As XmlNodeList
        SyncLock mSettingsLock
            Try
                Dim p As String = "//configuration/" & section.ToString.ToLower
                If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
                Dim fullpath As String = p & "/" & xpath
                Return mSettings.SelectNodes(fullpath)
            Catch ex As XmlException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As XPathException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As NullReferenceException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            Catch ex As InvalidCastException
                CheckForConfigurationError(section, xpath)
                Return Nothing
            End Try
        End SyncLock
    End Function
    Private Function DecryptNode(ByVal node As XmlNode) As XmlNode
        If Not String.IsNullOrEmpty(node.InnerText) Then
            If node.InnerText.Length >= 4 Then
                If node.InnerText.Substring(0, 4) = "ENC(" Then
                    node.InnerText = DecryptConfigString(node.InnerText)
                End If
            End If
        End If
        If node.HasChildNodes Then
            For Each cn As XmlNode In node.ChildNodes
                DecryptNode(cn)
            Next
        End If
        Return node
    End Function
    Public Function EncryptConfigString(ByVal value As String) As String
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As Encryption.Data
        encryptedData = sym.Encrypt(New Encryption.Data(value), key)
        Dim s As String = encryptedData.ToBase64
        Return "ENC(" & s & ")"
    End Function
    Public Function DecryptConfigString(ByVal value As String) As String
        If value.Length < 4 Then Return value
        If value.Substring(0, 4) <> "ENC(" Then
            'it's not encrypted!
            Return value
        End If
        value = value.Substring(4)
        value = value.Substring(0, value.Length - 1)
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data(mKey)
        Dim encryptedData As New Encryption.Data
        encryptedData.Base64 = value
        Dim decryptedData As Encryption.Data
        decryptedData = sym.Decrypt(encryptedData, key)
        Return decryptedData.ToString
    End Function

    Public Function GetString(ByVal section As ConfigSection, ByVal xpath As String) As String
        Return GetString(section, xpath, String.Empty, False, False)
    End Function
    Public Function GetString(ByVal section As ConfigSection, ByVal xpath As String, ByVal forceEncypt As Boolean) As String
        Return GetString(section, xpath, String.Empty, False, forceEncypt)
    End Function

    Public Function GetString(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As String) As String
        Return GetString(section, xpath, [default], False, False)
    End Function
    Public Function GetString(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As String, ByVal createIfNotFound As Boolean) As String
        Return GetString(section, xpath, [default], createIfNotFound, False)
    End Function

    Public Function GetString(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As String, ByVal createIfNotFound As Boolean, ByVal forceEncrypt As Boolean) As String
        SyncLock mSettingsLock
            Try
                Dim p As String = "//configuration/" & section.ToString.ToLower
                If section = ConfigSection.configuration Then p = "//" & section.ToString.ToLower
                Dim fullpath As String = p & "/" & xpath
                Dim value As String = mSettings.SelectSingleNode(fullpath).InnerText

                Dim fileEncrypt As Boolean = False
                Dim attr As XmlAttribute = mSettings.SelectSingleNode(fullpath).Attributes.GetNamedItem("encrypt")
                If attr IsNot Nothing Then
                    fileEncrypt = (attr.Value.ToLower = "true") OrElse (attr.Value.ToLower = "yes")
                End If
                If value.Length = 0 Then
                    Return [default]
                Else
                    Dim isEnc As Boolean = False
                    If value.Length >= 4 Then
                        isEnc = (value.Substring(0, 4) = "ENC(")
                    End If
                    If forceEncrypt AndAlso Not isEnc Then
                        WriteString(section, xpath, value, True)
                    End If
                    If isEnc Then
                        If Not fileEncrypt Then
                            WriteString(section, xpath, value, True)
                        End If
                        Return DecryptConfigString(value)
                    Else
                        If (forceEncrypt OrElse fileEncrypt) Then
                            WriteString(section, xpath, value, True)
                        End If
                        Return value
                    End If
                End If
                If forceEncrypt OrElse fileEncrypt Then
                    WriteString(section, xpath, value, True)
                End If
                Return value
            Catch ex As XmlException
                CheckForConfigurationError(section, xpath)
                Return [default]
            Catch ex As XPathException
                CheckForConfigurationError(section, xpath)
                Return [default]
            Catch ex As NullReferenceException
                If createIfNotFound Then
                    AddSetting(section, xpath, [default], forceEncrypt)
                End If
                Return [default]
            End Try
        End SyncLock
    End Function
    Private Sub AddSetting(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Object)
        AddSetting(section, xpath, [default], False)
    End Sub
    Private Sub AddSetting(ByVal section As ConfigSection, ByVal xpath As String, ByVal [default] As Object, ByVal forceEncrypt As Boolean)
        'If a settiing doesn't exist in the App.xml, add it in
        Dim nodes() = xpath.Split("/")
        Dim path As New StringBuilder
        Dim fullPath As New StringBuilder
        Try
            If section = ConfigSection.configuration Then
                path.Append("//" & section.ToString.ToLower)
            Else
                path.Append("//configuration/" & section.ToString.ToLower)
            End If

            For Each item As String In nodes
                If item.Trim.Length > 0 Then
                    If nodes.Length = (Array.IndexOf(nodes, item) + 1) Then
                        If mSettings.SelectSingleNode(path.ToString & fullPath.ToString & "/" & item) Is Nothing Then
                            'Node and Value don't exist
                            If forceEncrypt Then
                                AddNode(section, fullPath.ToString, String.Format("<{0}>{1}</{0}>", item, [default]), True)
                            Else
                                AddNode(section, fullPath.ToString, String.Format("<{0}>{1}</{0}>", item, [default]))
                            End If
                        Else
                            'Value doesn't exist
                            WriteString(section, fullPath.ToString & "/" & item, [default], forceEncrypt)
                        End If
                    Else
                        If mSettings.SelectSingleNode(path.ToString & fullPath.ToString & "/" & item) Is Nothing Then
                            'Node doesn't exist
                            If forceEncrypt Then
                                AddNode(section, fullPath.ToString, String.Format("<{0}></{0}>", item), True)
                            Else
                                AddNode(section, fullPath.ToString, String.Format("<{0}></{0}>", item))
                            End If
                        End If
                    End If
                    fullPath.Append("/" & item)
                End If
            Next
        Catch ex As Exception
            Throw New InvalidConfigurationParameterException(String.Format("Unable to add new setting to App.xml ({0}): {1}", path.ToString & xpath, ex.Message))
        End Try
    End Sub
    Private Function GetTraceLevel() As Integer
        'as of 7/23/09, GettraceLevel is obsolete -- use GetConsoleTraceLevel or GetLogTraceLevel
        'here for backward compatibility
        Try
            Dim fullpath As String = "//configuration/diagnostics/tracelevel"
            Return CInt(DecryptConfigString(mSettings.SelectSingleNode(fullpath).InnerText))
        Catch ex As XmlException
            Return TraceLevel.Warning
        Catch ex As InvalidCastException
            Return TraceLevel.Warning
        End Try
    End Function
    Private Function GetConsoleTraceLevel() As Integer
        Try
            Dim fullpath As String = "//configuration/diagnostics/console/tracelevel"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                'revert to default
                fullpath = "//configuration/diagnostics/tracelevel"
            End If
            Return CInt(mSettings.SelectSingleNode(fullpath).InnerText)
        Catch ex As XmlException
            Return TraceLevel.Warning
        Catch ex As InvalidCastException
            Return TraceLevel.Warning
        End Try
    End Function
    Private Function GetMaxConsoleWidth() As Integer
        Try
            Dim fullpath As String = "//configuration/diagnostics/console/maxWidth"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                'revert to default
                Return 120
            End If
            Return CInt(mSettings.SelectSingleNode(fullpath).InnerText)
        Catch ex As XmlException
            Return 60
        Catch ex As InvalidCastException
            Return 60
        End Try

    End Function
    Private Function GetLogfileTraceLevel() As Integer
        Try
            Dim fullpath As String = "//configuration/diagnostics/logfile/tracelevel"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                'revert to default
                fullpath = "//configuration/diagnostics/tracelevel"
            End If
            Return CInt(mSettings.SelectSingleNode(fullpath).InnerText)
        Catch ex As XmlException
            Return TraceLevel.Warning
        Catch ex As InvalidCastException
            Return TraceLevel.Warning
        End Try
    End Function
    Private Function GetSysLogServer() As String
        Try
            Dim fullpath As String = "//configuration/diagnostics/syslog/ipOrDnsName"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                Return ""
            End If
            fullpath = "//configuration/diagnostics/syslog/ipOrDnsName"
            Return mSettings.SelectSingleNode(fullpath).InnerText
        Catch ex As XmlException
            Return ""
        Catch ex As InvalidCastException
            Return ""
        End Try

    End Function
    Private Function GetSysLogEnabled() As Boolean
        Try
            Dim fullpath As String = "//configuration/diagnostics/syslog/enabled"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                Return False
            End If
            fullpath = "//configuration/diagnostics/syslog/enabled"
            Return CBool(mSettings.SelectSingleNode(fullpath).InnerText)
        Catch ex As XmlException
            Return False
        Catch ex As InvalidCastException
            Return False
        End Try

    End Function
    Private Function GetLogfileTraceFormat() As LogFileTraceFormats
        Try
            Dim fullpath As String = "//configuration/diagnostics/logfile/traceFormat"
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                'revert to default
                Return LogFileTraceFormats.legacy
            End If
            Return [Enum].Parse(GetType(LogFileTraceFormats), mSettings.SelectSingleNode(fullpath).InnerText)
        Catch ex As XmlException
            Return LogFileTraceFormats.legacy
        Catch ex As InvalidCastException
            Return LogFileTraceFormats.legacy
        End Try
    End Function

    <MethodImpl(MethodImplOptions.Synchronized)> _
    Public Sub LoadSettings()
        App.TraceLog(Diagnostics.TraceLevel.Info, IIf(App.ApplicationType = ApplicationType.WindowsService, "============New Service Instance Starting============", "============New Application Instance Starting============"))
        App.TraceLog(Diagnostics.TraceLevel.Info, "ConfigFile loading settings from " & mPath)
        Dim doc As New XmlDocument
        Try
            doc.Load(mPath)
            SyncLock mSettingsLock
                mSettings = doc
            End SyncLock
            mTraceLevel = CType(GetTraceLevel(), TraceLevel)  'cache the trace level for performance reasons
            mLogFileTraceFormat = GetLogfileTraceFormat()     'cache format
            mConsoleTraceLevel = CType(GetConsoleTraceLevel(), TraceLevel)
            mLogfileTraceLevel = CType(GetLogfileTraceLevel(), TraceLevel)
            mMaxConsoleWidth = GetMaxConsoleWidth()
            mSyslogServer = GetSysLogServer()
            msyslogEnabled = GetSysLogEnabled()
        Catch ex As XmlException
            Throw New AmcomException("The configuration file is not a valid xml document", ex)
        End Try

        App.TraceLog(Diagnostics.TraceLevel.Info, "ConfigFile finished loading settings")
    End Sub

    Public Sub RemoveNode(ByVal section As ConfigSection, ByVal xpath As String)
        SyncLock mSettingsLock
            Dim fullpath As String = "//" & section.ToString.ToLower & "/" & xpath
            Dim node As XmlNode = mSettings.SelectSingleNode(fullpath)
            node.ParentNode.RemoveChild(node)
            mSettings.Save(mPath)
        End SyncLock
    End Sub
    Public Sub WriteString(ByVal section As ConfigSection, ByVal xpath As String, ByVal value As String)
        WriteString(section, xpath, value, False)
    End Sub
    Public Sub WriteString(ByVal section As ConfigSection, ByVal xpath As String, ByVal value As String, ByVal forceEncrypt As Boolean)
        SyncLock mSettingsLock
            Dim fullpath As String = ""
            If section = ConfigSection.configuration Then
                fullpath = "//" & section.ToString.ToLower & "/" & xpath
            Else
                fullpath = "//configuration/" & section.ToString.ToLower & "/" & xpath
            End If
            If mSettings.SelectSingleNode(fullpath) Is Nothing Then
                AddSetting(section, xpath, String.Empty, forceEncrypt)
            End If
            Dim attr As XmlAttribute = mSettings.SelectSingleNode(fullpath).Attributes.GetNamedItem("encrypt")
            Dim fileEncrypt As Boolean = False
            If attr IsNot Nothing Then
                fileEncrypt = (attr.Value.ToLower = "true") OrElse (attr.Value.ToLower = "yes")
            End If
            If forceEncrypt Or fileEncrypt Then
                If Not value.StartsWith("ENC") Then
                    mSettings.SelectSingleNode(fullpath).InnerText = EncryptConfigString(value)
                End If
                If attr Is Nothing Then
                    attr = mSettings.CreateAttribute("encrypt")
                    attr.Value = "true"
                    mSettings.SelectSingleNode(fullpath).Attributes.Append(attr)
                Else
                    attr.Value = "true"
                End If
            Else
                mSettings.SelectSingleNode(fullpath).InnerText = value
            End If

            mSettings.Save(mPath)

        End SyncLock
    End Sub
    Public Sub WriteString(ByVal section As ConfigSection, ByVal xpath As String, ByVal value As String, ByVal settingsFile As String)
        WriteString(section, xpath, value, settingsFile, False)
    End Sub
    Public Sub WriteString(ByVal section As ConfigSection, ByVal xpath As String, ByVal value As String, ByVal settingsFile As String, ByVal forceEncrypt As Boolean)
        Dim doc As New XmlDocument
        doc.Load(settingsFile)
        Dim fullpath As String = ""
        If section = ConfigSection.configuration Then
            fullpath = "//" & section.ToString.ToLower & "/" & xpath
        Else
            fullpath = "//configuration/" & section.ToString.ToLower & "/" & xpath
        End If
        If forceEncrypt Then
            doc.SelectSingleNode(fullpath).InnerText = EncryptConfigString(value)
            Dim attr As XmlAttribute = mSettings.SelectSingleNode(fullpath).Attributes.GetNamedItem("encrypt")
            If attr Is Nothing Then
                attr = mSettings.CreateAttribute("encrypt")
                attr.Value = "true"
                mSettings.SelectSingleNode(fullpath).Attributes.Append(attr)
            Else
                attr.Value = "true"
            End If
        Else
            doc.SelectSingleNode(fullpath).InnerText = value
        End If
        doc.Save(settingsFile)
    End Sub


#End Region

End Class
