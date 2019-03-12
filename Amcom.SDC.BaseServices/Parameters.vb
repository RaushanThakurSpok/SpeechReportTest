Imports System.Collections.ObjectModel
Public Class Parameters
    Inherits SortedDictionary(Of String, Parameter)
    Private mDirty As Boolean
  

    Public ReadOnly Property Clone() As Parameters
        Get
                Dim pl As New Parameters
            For Each p As Parameter In MyBase.Values
                Dim p2 As New Parameter(p.Name)
                p2.ObjectValue = p.ObjectValue
                pl.Add(p.Name, p2)
            Next
            Return pl
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get

            Return Me.Item("name").Value
        End Get
    End Property
    Public ReadOnly Property IsDirty() As Boolean
        Get
            Return mDirty
        End Get
    End Property

    Public ReadOnly Property DisplayList() As String
        Get
            Dim index As Integer = 0
            Dim sb As New Text.StringBuilder
            For Each k As String In MyBase.Keys
                If index > 0 Then sb.Append(", ")
                sb.Append(k)
                sb.Append(":")
                sb.Append(MyBase.Item(k).Value)
                sb.Append(" ")
            Next
            Return sb.ToString
        End Get
    End Property

    Public Sub New()

    End Sub
    Public Sub New(ByVal param As Parameter)
        MyBase.Add(param.Name, param)
    End Sub
    Public Sub New(ByVal ParamArray values() As Parameter)
        For Each p As Parameter In values
            MyBase.Add(p.Name, p)
        Next
    End Sub
    Public Overloads Sub Add(ByVal name As String, ByVal value As Object)
        Dim p As New Parameter(name, value)
        MyBase.Add(p.Name, p)
        ' AddHandler p.ValueChanged, AddressOf ValueChangedHandler
    End Sub

    Public Overloads Sub Add(ByVal name As String, ByVal value As Parameters)
        MyBase.Add(name, New Parameter(name, value))
    End Sub
    Public Overloads Sub Add(ByVal name As String, ByVal value As Parameter)
        MyBase.Add(Name, value)
        'AddHandler value.ValueChanged, AddressOf ValueChangedHandler
    End Sub
    Public Overloads Sub Add(ByVal name As String, ByVal value As String)
        Dim p As New Parameter(name, value)
        MyBase.Add(p.Name, p)
        ' AddHandler p.ValueChanged, AddressOf ValueChangedHandler
    End Sub
    Public Sub ValueChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
        mDirty = True
    End Sub
    Private Sub WriteParameterChildren(ByVal writer As System.Xml.XmlTextWriter, ByVal p As Parameter)
        If p.ParameterList Is Nothing Then
            writer.WriteAttributeString("islist", "false")
            writer.WriteElementString("value", p.Value)
        Else
            writer.WriteAttributeString("islist", "true")
            For Each subp As Parameter In p.ParameterList.Values
                writer.WriteStartElement("param")
                writer.WriteAttributeString("name", subp.Name)
                WriteParameterChildren(writer, subp)
                writer.WriteEndElement()
            Next
        End If
    End Sub
    Public ReadOnly Property ToXml() As String
        Get
            Dim ms As New IO.MemoryStream
            Dim doc As New System.Xml.XmlDocument
            Dim xml As String
            Using writer As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(ms, Text.Encoding.UTF8)
                writer.Formatting = System.Xml.Formatting.Indented
                writer.WriteStartDocument()
                writer.WriteStartElement("plist")
                For Each p As Parameter In MyBase.Values
                    writer.WriteStartElement("param")
                    writer.WriteAttributeString("name", p.Name)
                    WriteParameterChildren(writer, p)
                    writer.WriteEndElement()
                Next
                writer.WriteEndElement() 'plist
                writer.Flush()
                ms.Position = 0
                xml = New IO.StreamReader(ms).ReadToEnd
            End Using
            Return xml
        End Get
    End Property

    Public Shared Function FromXML(ByVal xml As String) As Parameters
        Dim params As New Parameters
        Dim currentItem As System.Xml.XmlNode = Nothing
        Dim currentParams As Parameters = Nothing
        Dim currentName As String = ""
        Dim doc As New System.Xml.XmlDocument
        Try
            doc.LoadXml(xml)
        Catch ex As Exception
            Return Nothing
        End Try
        For Each item As System.Xml.XmlNode In doc.SelectNodes("plist/param")
            params.ReadParameterChildren(item, params)
        Next
        Return params
    End Function

    Protected Sub ReadParameterChildren(ByVal node As System.Xml.XmlNode, ByVal params As Parameters)
        If node.Attributes.GetNamedItem("islist").Value = "true" Then
            Dim p As New Parameters
            params.Add(node.Attributes.GetNamedItem("name").Value, p)
            For Each childNode As System.Xml.XmlNode In node.SelectNodes("param")
                ReadParameterChildren(childNode, p)
            Next
        Else
            params.Add(node.Attributes.GetNamedItem("name").Value, node.SelectSingleNode("value").InnerText)
        End If
    End Sub
End Class
