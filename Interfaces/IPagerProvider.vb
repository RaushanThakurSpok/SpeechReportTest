Imports System.Collections.ObjectModel
Namespace Notifications
    Public Interface IPagerProvider
        ''' <summary>
        ''' Set the configuration of the supplied paging provider.
        ''' </summary>
        ''' <param name="parameters">Keys and values required by the paging provider</param>
        ''' <remarks>Paramters are stored in App.xml</remarks>
        Sub SetConfiguration(ByVal parameters As Dictionary(Of String, String))
        ReadOnly Property Name() As String
        Sub SendPage(ByVal pagerNumber As String, ByVal message As String)
    End Interface
End Namespace