Imports System.Collections.ObjectModel
Namespace Notifications
    Public Interface IMailProvider
        ''' <summary>
        ''' Set the configuration of the supplied email provider.
        ''' </summary>
        ''' <param name="parameters">Keys and values required by the paging provider</param>
        ''' <remarks>Paramters are stored in App.xml</remarks>
        Sub SetConfiguration(ByVal parameters As Dictionary(Of String, String))
        ReadOnly Property Name() As String
        Sub SendMail(ByVal fromAddress As String, ByVal toAddress As String, ByVal subject As String, ByVal body As String, ByVal ParamArray attachments() As String)
    End Interface
End Namespace