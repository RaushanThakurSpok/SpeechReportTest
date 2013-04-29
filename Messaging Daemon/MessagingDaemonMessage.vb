Imports System.Xml
Imports System.IO
Imports System.Text
Namespace TCP
    Public Class MessagingDaemonMessage
        Implements IDisposable

        Private mQueryAddress As String
        Private mCommandName As String
        Private mQueryParameters As Parameters
        Private mResponseStatus As String
        Private mResponseValue As String
        Private mResponseData As Parameters
        Private mUniqueId As String
#Region "Constructor"
        Public Sub New(ByVal commandName As String, ByVal uniqueID As String)
            Initialize()
            mCommandName = commandName
            mUniqueId = uniqueID
            QueryParameters.Add("_uniqueid", mUniqueId)

        End Sub
        Public Sub New(ByVal commandName As String)
            Initialize()
            mCommandName = commandName
            mUniqueId = CreateGUID()
        End Sub
        Public Sub New()
            Initialize()
            mUniqueId = CreateGUID()
        End Sub
#End Region
#Region "Properties"
        Public Property Command() As String
            Get
                Return mCommandName
            End Get
            Set(ByVal value As String)
                mCommandName = value
            End Set
        End Property
        Public ReadOnly Property UniqueId() As String
            Get
                Return mUniqueId
            End Get
        End Property
        Public ReadOnly Property DisplayQuery() As String
            Get
                Dim sb As New StringBuilder
                sb.Append(mCommandName)
                If mQueryParameters.Keys.Count > 0 Then
                    Dim count As Integer = 0
                    sb.Append("(")
                    For Each p As Parameter In mQueryParameters.Values
                        sb.Append(p.Name)
                        If count > 0 Then sb.Append(", ")
                        sb.Append(":")
                        sb.Append(p.Value)
                        count = +1
                    Next
                    sb.Append("): ")
                End If
                Return sb.ToString
            End Get
        End Property
        Public ReadOnly Property DisplayResponse() As String
            Get
                Return DisplayResponse(" | ")
            End Get
        End Property
        Public ReadOnly Property DisplayResponse(ByVal sepChar As String) As String
            Get
                If mResponseStatus = "Packet" Then
                    Return mResponseValue
                End If
                If mResponseStatus <> "OK" Then
                    Return "Command Failed - " & mResponseValue
                End If
                If mCommandName.Length = 0 Then
                    Return mResponseValue
                End If
                If mResponseData.Count = 0 Then
                    Return mCommandName & ": " & mResponseValue
                End If
                Dim sb As New StringBuilder
                sb.Append(Me.DisplayQuery)
                Dim count As Integer = 0
                sb.Append(" ")
                For Each p As Parameter In mResponseData.Values
                    If count > 0 Then sb.Append(sepChar)
                    sb.Append(p.Name)
                    sb.Append(":")
                    sb.Append(p.Value)
                    count += 1
                Next
                Return sb.ToString
            End Get
        End Property
        Public ReadOnly Property QueryValues() As String()
            Get
                Dim i As Integer = 0
                Dim p(mQueryParameters.Count - 1) As String
                For Each param As Parameter In mQueryParameters.Values
                    p(i) = param.Value
                    i += 1
                Next
                Return p
            End Get
        End Property
        Public Property QueryAddress() As String
            Get
                Return mQueryAddress
            End Get
            Set(ByVal value As String)
                mQueryAddress = value
            End Set
        End Property
        Public ReadOnly Property QueryParameters() As Parameters
            Get
                Return mQueryParameters
            End Get
        End Property
        Public ReadOnly Property QueryDocument() As XmlDocument
            Get
                Dim ms As New MemoryStream
                Dim doc As New XmlDocument

                Using writer As XmlTextWriter = New XmlTextWriter(ms, Encoding.UTF8)
                    writer.Formatting = Formatting.Indented
                    writer.WriteStartDocument()
                    writer.WriteStartElement("sdcSolutions")
                    writer.WriteStartElement("messaging")
                    writer.WriteStartElement("command")
                    writer.WriteElementString("uniqueid", mUniqueId)
                    writer.WriteElementString("name", mCommandName)


                    writer.WriteStartElement("parameters")
                    For Each p As Parameter In mQueryParameters.Values
                        writer.WriteStartElement("parameter")
                        writer.WriteElementString("name", p.Name)
                        writer.WriteElementString("value", p.Value)
                        writer.WriteEndElement()
                    Next
                    writer.WriteEndElement() 'parameters
                    writer.WriteEndElement() 'command

                    writer.WriteStartElement("response")
                    writer.WriteElementString("status", mResponseStatus)
                    writer.WriteElementString("value", mResponseValue)

                    writer.WriteStartElement("parameters")
                    For Each p As Parameter In mQueryParameters.Values
                        writer.WriteElementString("name", p.Name)
                        writer.WriteElementString("value", p.Value)
                    Next

                    writer.WriteEndElement() 'parameters
                    writer.WriteEndElement() 'response

                    writer.WriteEndElement() 'messaging
                    writer.WriteEndElement() 'sdcSolutions
                    writer.Flush()
                    ms.Position = 0
                    Using sr As New StreamReader(ms)
                        doc.LoadXml(sr.ReadToEnd)
                    End Using
                End Using
                Return doc
            End Get
        End Property
        Public ReadOnly Property ResponseDocument() As XmlDocument
            Get
                Dim ms As New MemoryStream
                Dim doc As New XmlDocument

                Using writer As XmlTextWriter = New XmlTextWriter(ms, Encoding.UTF8)
                    writer.Formatting = Formatting.Indented
                    writer.WriteStartDocument()
                    writer.WriteStartElement("sdcSolutions")
                    writer.WriteStartElement("messaging")
                    writer.WriteStartElement("command")
                    writer.WriteElementString("uniqueid", mUniqueId)
                    writer.WriteElementString("name", mCommandName)

                    writer.WriteStartElement("parameters")
                    'gbh fix
                    For Each p As Parameter In mQueryParameters.Values
                        If p.ParameterList IsNot Nothing Then
                            writer.WriteStartElement("parameter")
                            writer.WriteElementString("name", p.Name)

                            writer.WriteStartElement("parameterlist")
                            For Each s As String In p.ParameterList.Keys
                                writer.WriteStartElement("parameter")
                                writer.WriteElementString("name", s)
                                writer.WriteElementString("value", p.ParameterList(s).Value)
                                writer.WriteEndElement()
                            Next
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("parameter")
                            writer.WriteElementString("name", p.Name)
                            writer.WriteElementString("value", p.Value)
                            writer.WriteEndElement()
                        End If
                    Next
                    writer.WriteEndElement() 'parameters
                    writer.WriteEndElement() 'command

                    writer.WriteStartElement("response")
                    writer.WriteElementString("status", mResponseStatus)
                    writer.WriteElementString("value", mResponseValue)

                    writer.WriteStartElement("parameters")
                    For Each p As Parameter In mResponseData.Values
                        If p.ParameterList IsNot Nothing Then
                            writer.WriteStartElement("parameter")
                            writer.WriteElementString("name", p.Name)

                            writer.WriteStartElement("parameterlist")
                            For Each s As String In p.ParameterList.Keys
                                writer.WriteStartElement("parameter")
                                writer.WriteElementString("name", s)
                                writer.WriteElementString("value", p.ParameterList(s).Value)
                                writer.WriteEndElement()
                            Next
                            writer.WriteEndElement()
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("parameter")
                            writer.WriteElementString("name", p.Name)
                            writer.WriteElementString("value", p.Value)
                            writer.WriteEndElement()
                        End If
                    Next

                    writer.WriteEndElement() 'parameters
                    writer.WriteEndElement() 'response

                    writer.WriteEndElement() 'messaging
                    writer.WriteEndElement() 'sdcSolutions
                    writer.Flush()
                    ms.Position = 0
                    Using sr As New StreamReader(ms)
                        doc.LoadXml(sr.ReadToEnd)
                    End Using
                End Using
                Return doc
            End Get
        End Property

        Public Property ResponseStatus() As String
            Get
                Return mResponseStatus
            End Get
            Set(ByVal value As String)
                mResponseStatus = value
            End Set
        End Property
        Public Property ResponseValue() As String
            Get
                Return mResponseValue
            End Get
            Set(ByVal value As String)
                mResponseValue = value
            End Set
        End Property
        Public ReadOnly Property ResponseData() As Parameters
            Get
                Return mResponseData
            End Get
        End Property

#End Region
#Region "Methods"
        Friend Sub AddQueryParameter(ByVal name As String, ByVal value As String)
            mQueryParameters.Add(name, New Parameter(name, value))
        End Sub
        Public Function Copy() As MessagingDaemonMessage
            Dim m As New MessagingDaemonMessage(mCommandName)
            For Each p As Parameter In mQueryParameters.Values
                m.AddQueryParameter(p.Name, p.Value)
            Next
            m.ResponseStatus = mResponseStatus
            m.ResponseValue = mResponseValue
            For Each p As Parameter In ResponseData.Values
                m.ResponseData.Add(p.Name, p)
            Next
            Return m
        End Function
        Public Sub Decode(ByVal doc As XmlDocument)
            mCommandName = doc.SelectSingleNode("//sdcSolutions/messaging/command/name").InnerText
            mUniqueId = doc.SelectSingleNode("//sdcSolutions/messaging/command/uniqueid").InnerText
            mQueryParameters.Clear()
            For Each p As XmlNode In doc.SelectNodes("//sdcSolutions/messaging/command/parameters/parameter")
                Dim par As New Parameter(p.SelectSingleNode("name").InnerText, p.SelectSingleNode("value").InnerText)
                mQueryParameters.Add(par.Name, par)
            Next
            mResponseStatus = doc.SelectSingleNode("//sdcSolutions/messaging/response/status").InnerText
            mResponseValue = doc.SelectSingleNode("//sdcSolutions/messaging/response/value").InnerText
            mResponseData.Clear()
            For Each p As XmlNode In doc.SelectNodes("//sdcSolutions/messaging/response/parameters/parameter")
                Dim list As XmlNodeList = p.SelectNodes("parameterlist/parameter")
                If list.Count > 0 Then
                    Dim ps As New Parameters
                    For Each node As XmlNode In list
                        ps.Add(node.SelectSingleNode("name").InnerText, node.SelectSingleNode("value").InnerText)
                    Next
                    mResponseData.Add(p.SelectSingleNode("name").InnerText, ps)
                Else
                    Dim par As New Parameter(p.SelectSingleNode("name").InnerText, p.SelectSingleNode("value").InnerText)
                    mResponseData.Add(par.Name, par)
                End If
            Next
        End Sub
        Friend Sub RemoveQueryParameter(ByVal name As String)
            mQueryParameters.Remove(name)
        End Sub
        Public Sub Initialize()
            mQueryAddress = ""
            mCommandName = ""
            mQueryParameters = New Parameters
            mResponseValue = ""
            mResponseData = New Parameters
        End Sub

#End Region

        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free other state (managed objects).
                    App.TraceLog(TraceLevel.Info, "MDM Dispose")
                    mQueryParameters.Dispose()
                    mResponseData.Dispose()
                    MyBase.Finalize()
                End If

                ' TODO: free your own state (unmanaged objects).
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Protected Overrides Sub Finalize()
			MyBase.Finalize()
        End Sub
    End Class
End Namespace
