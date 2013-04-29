Imports System.Data.SqlClient
Imports System.xml
Public Class WinConnectionInfo
    Public Enum RecordSetShareOptionType
        disableMultiRecordsetsOnConnection
        enableMultiRecordsetsOnConnection
    End Enum

    Private Shared mConnectionBuilder As SqlConnectionStringBuilder = CreateConnectionBuilder()
    Private Sub New()
    End Sub
    Public Shared ReadOnly Property ConnectionString(ByVal applicationDatabase As String) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase)
            Return cb.ConnectionString
        End Get
    End Property
    Public Shared ReadOnly Property ConnectionString() As String
        Get
            Return mConnectionBuilder.ConnectionString
        End Get
    End Property
    Public Shared ReadOnly Property ConnectionString(ByVal allowMultiRecordsets As RecordSetShareOptionType) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(allowMultiRecordsets)
            Return cb.ConnectionString
        End Get
    End Property
    Public Shared ReadOnly Property ConnectionString(ByVal applicationDatabase As String, ByVal allowMultiRecordsets As RecordSetShareOptionType) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase, allowMultiRecordsets)
            Return cb.ConnectionString
        End Get
    End Property

    Public Shared ReadOnly Property DataSource(ByVal applicationDatabase As String) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase)
            Return cb.DataSource
        End Get
    End Property

    Public Shared ReadOnly Property DataSource(ByVal applicationDatabase As String, ByVal allowMultiRecordSets As RecordSetShareOptionType) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase, allowMultiRecordSets)
            Return cb.DataSource
        End Get
    End Property

    Public Shared ReadOnly Property DataSource() As String
        Get
            Return mConnectionBuilder.DataSource
        End Get
    End Property

    Public Shared ReadOnly Property Password(ByVal applicationDatabase As String) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase)
            Return cb.Password
        End Get
    End Property
    Public Shared ReadOnly Property Password() As String
        Get
            Return mConnectionBuilder.Password
        End Get
    End Property
    Public Shared ReadOnly Property UserId(ByVal applicationDatabase As String) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase)
            Return cb.UserID
        End Get
    End Property

    Public Shared ReadOnly Property UserId() As String
        Get
            Return mConnectionBuilder.UserID
        End Get
    End Property

    Public Shared ReadOnly Property InitialCatalog(ByVal applicationDatabase As String) As String
        Get
            Dim cb As SqlConnectionStringBuilder = CreateConnectionBuilder(applicationDatabase)
            Return cb.InitialCatalog
        End Get
    End Property
    Public Shared ReadOnly Property InitialCatalog() As String
        Get
            Return mConnectionBuilder.InitialCatalog
        End Get
    End Property

    Private Shared Function CreateConnectionBuilder() As SqlConnectionStringBuilder
        Dim builder As New SqlConnectionStringBuilder()
        Dim serverName As String = App.Config.GetString(ConfigSection.configuration, "serverName", "")
        builder.DataSource = App.Config.GetString(ConfigSection.database, "dataSource", "")
        builder.InitialCatalog = App.Config.GetString(ConfigSection.database, "initialCatalog", "")
        builder.IntegratedSecurity = True
        builder.Add("Workstation ID", My.Computer.Name)
        builder.Add("Persist Security Info", "false")
        Return builder
    End Function
    Private Shared Function CreateConnectionBuilder(ByVal allowMultiRecordSets As RecordSetShareOptionType) As SqlConnectionStringBuilder
        Dim builder As New SqlConnectionStringBuilder()
        Dim serverName As String = App.Config.GetString(ConfigSection.configuration, "serverName", "")
        builder.DataSource = App.Config.GetString(ConfigSection.database, "dataSource", "")
        builder.InitialCatalog = App.Config.GetString(ConfigSection.database, "initialCatalog", "")
        builder.IntegratedSecurity = True
        If allowMultiRecordSets = RecordSetShareOptionType.enableMultiRecordsetsOnConnection Then
            builder.Add("MultipleActiveResultSets", "True")
        End If
        builder.Add("Workstation ID", My.Computer.Name)
        builder.Add("Persist Security Info", "false")

        Return builder
    End Function

    Private Shared Function CreateConnectionBuilder(ByVal applicationDatabase As String, ByVal allowMultiRecordSets As RecordSetShareOptionType) As SqlConnectionStringBuilder
        Dim builder As New SqlConnectionStringBuilder()
        Dim serverName As String = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/serverName")
        builder.ApplicationName = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/applicationName", App.ApplicationName)
        builder.DataSource = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/serverName")
        builder.InitialCatalog = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/dbName")
        builder.IntegratedSecurity = App.Config.GetBoolean(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/integratedSecurity")
        builder.Add("User ID", App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/userName"))
        builder.Add("Password", App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/password"))
        If allowMultiRecordSets = RecordSetShareOptionType.enableMultiRecordsetsOnConnection Then
            builder.Add("MultipleActiveResultSets", "True")
        End If
        builder.Add("Workstation ID", My.Computer.Name)
        builder.Add("Persist Security Info", "false")
        Return builder
    End Function
    Private Shared Function CreateConnectionBuilder(ByVal applicationDatabase As String) As SqlConnectionStringBuilder
        Dim builder As New SqlConnectionStringBuilder()
        builder.ApplicationName = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/applicationName", App.ApplicationName)
        Dim serverName As String = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/serverName")
        builder.DataSource = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/serverName")
        builder.InitialCatalog = App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/dbName")
        builder.IntegratedSecurity = App.Config.GetBoolean(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/integratedSecurity")
        builder.Add("User ID", App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/userName"))
        builder.Add("Password", App.Config.GetString(ConfigSection.configuration, "databaseConnections/database[@application='" & applicationDatabase & "']/password"))
        builder.Add("Workstation ID", My.Computer.Name)
        builder.Add("Persist Security Info", "false")
        Return builder
    End Function

End Class
