Imports System.Collections.ObjectModel
Imports Amcom.SDC.BaseServices
Namespace TCP
    Public Class MessagingDaemonMessages

        Inherits Dictionary(Of String, MessagingDaemonMessage)

#Region "Constructor"

        Public Sub New()
            MyBase.new()
        End Sub

        Public Function DisplayCommands() As String
            Dim sbReturn As New Text.StringBuilder
            '
            For Each commandName As String In MyBase.Keys
                sbReturn.Append(vbCrLf)
                sbReturn.Append("Command: " & commandName & vbCrLf)
                '
                For Each paramName As String In MyBase.Item(commandName).QueryParameters.Keys
                    sbReturn.Append("ParamName: " & paramName & vbCrLf)
                    sbReturn.Append("ParamValue: " & MyBase.Item(commandName).QueryParameters.Item(paramName).Value & vbCrLf)
                Next
            Next
            '
            Return sbReturn.ToString
        End Function

        Public Sub AddCommand(ByVal Command As String, ByVal ParamsAndValues As Dictionary(Of String, String))
            Dim m As New MessagingDaemonMessage(Command)
            '
            If (ParamsAndValues.Count > 0) Then
                For Each paramName As String In ParamsAndValues.Keys
                    m.AddQueryParameter(paramName, ParamsAndValues.Item(paramName))
                Next
            End If
            '
            MyBase.Add(m.Command, m)
        End Sub
#End Region
    End Class
End Namespace
