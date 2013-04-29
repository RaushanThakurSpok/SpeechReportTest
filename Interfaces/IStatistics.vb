Public Interface IStatistics
    'provides Application Information data that is useful for the HealthStatusMonitor
    ReadOnly Property AppUppTime() As String
    ReadOnly Property LoggedUnhandledExceptionCount() As Integer
    ReadOnly Property LoggedUnhandledExceptionCountToday() As Integer
    ReadOnly Property LastExceptionDetails() As String

End Interface
