Imports System.Xml
Imports System.Collections.ObjectModel
Namespace Notifications
    Friend Class Configuration
        Private _hasConfiguration As Boolean
        Private _hasPagingConfig As Boolean
        Private _hasEmailConfig As Boolean
        Private _pageList As ProviderConfiguration
        Private _emailList As ProviderConfiguration
        Private _pagingProviderConfigs As List(Of ProviderConfiguration)
        Private _emailProviderConfigs As List(Of ProviderConfiguration)

        Public Sub New()

            If App.ApplicationType = ApplicationType.NoConfigConsole OrElse App.ApplicationType = ApplicationType.NoConfigComponent Then
                _hasConfiguration = False
                Exit Sub
            End If
            Dim providers As XmlNode = App.Config.GetNode(ConfigSection.modules, "notificationProviders")
            If providers Is Nothing Then
                _hasConfiguration = False
                Exit Sub
            End If
            _pagingProviderConfigs = New List(Of ProviderConfiguration)
            _pageList = New ProviderConfiguration
            For Each pageServer As XmlNode In providers.SelectNodes("paging")
                _pageList.Add("priority", pageServer.Attributes.GetNamedItem("priority").Value)
                _pageList.Add("type", pageServer.Attributes.GetNamedItem("type").Value)
                For Each configParam As XmlNode In pageServer.SelectNodes("config")
                    For Each param As XmlNode In configParam.ChildNodes
                        _pageList.Add(param.Name, param.InnerText)
                    Next
                Next
                If _pagingProviderConfigs.Count = 0 Then
                    _pagingProviderConfigs.Add(_pageList)
                Else
                    For index As Integer = 0 To _pagingProviderConfigs.Count - 1
                        If _pagingProviderConfigs.Item(index)("priority") > _pageList("priority") Then
                            _pagingProviderConfigs.Insert(index, _pageList)
                            _pageList = Nothing
                            Exit For
                        End If
                    Next
                    If Not _pageList Is Nothing Then
                        _pagingProviderConfigs.Add(_pageList)
                    End If
                    _pageList = New ProviderConfiguration
                End If
            Next
            _hasPagingConfig = _pageList.Count > 0

            _emailProviderConfigs = New List(Of ProviderConfiguration)
            _emailList = New ProviderConfiguration
            For Each emailServer As XmlNode In providers.SelectNodes("email")
                _emailList.Add("priority", emailServer.Attributes.GetNamedItem("priority").Value)
                _emailList.Add("type", emailServer.Attributes.GetNamedItem("type").Value)
                For Each configParam As XmlNode In emailServer.SelectNodes("config")
                    For Each param As XmlNode In configParam.ChildNodes

                        If (param.Name.ToUpper = "USER" OrElse param.Name.ToUpper = "PASSWORD") AndAlso param.InnerText IsNot Nothing Then
                            Dim path As String = GetXPathToNode(param, 2)

                            App.Config.WriteString(ConfigSection.modules, path, param.InnerText, True)
                            _emailList.Add(param.Name, App.Config.GetString(ConfigSection.modules, path))
                        Else
                            _emailList.Add(param.Name, param.InnerText)
                        End If
                    Next
                Next
                If _emailProviderConfigs.Count = 0 Then
                    _emailProviderConfigs.Add(_emailList)
                Else
                    For index As Integer = 0 To _emailProviderConfigs.Count - 1
                        If _emailProviderConfigs.Item(index)("priority") > _emailList("priority") Then
                            _emailProviderConfigs.Insert(index, _emailList)
                            _emailList = Nothing
                            Exit For
                        End If
                    Next
                    If Not _emailList Is Nothing Then
                        _emailProviderConfigs.Add(_emailList)
                    End If
                    _emailList = New ProviderConfiguration
                End If
            Next
            _hasEmailConfig = _emailList.Count > 0

        End Sub
        Public ReadOnly Property HasConfiguration() As Boolean
            Get
                Return _hasConfiguration
            End Get
        End Property
        Public ReadOnly Property HasPagerProvider() As Boolean
            Get
                Return _hasPagingConfig
            End Get
        End Property
        Public ReadOnly Property HasEmailProvider() As Boolean
            Get
                Return _hasEmailConfig
            End Get
        End Property
        Public ReadOnly Property PagingProviders() As List(Of ProviderConfiguration)
            Get
                Return _pagingProviderConfigs
            End Get
        End Property
        Public ReadOnly Property EmailProviders() As List(Of ProviderConfiguration)
            Get
                Return _emailProviderConfigs
            End Get
        End Property
        Public Class ProviderConfiguration
            Inherits Dictionary(Of String, String)
        End Class
        Private Function GetXPathToNode(ByVal node As XmlNode, ByVal startNodeIndex As Integer)
            Dim temp As String = GetXPathToNode(node)
            Dim s() As String = temp.Split("/")
            Return "/" & String.Join("/", s, startNodeIndex + 1, s.Count - (startNodeIndex + 1))
        End Function
        Private Function GetXPathToNode(ByVal node As XmlNode) As String
            If (node.NodeType = XmlNodeType.Attribute) Then
                '' attributes have an OwnerElement, not a ParentNode; also they have
                '' to be matched by name, not found by position
                Return String.Format("{0}/@{1}", GetXPathToNode(DirectCast(node, XmlAttribute).OwnerElement), node.Name)
            End If
            'the only node with no parent is the root node, which has no path
            If node.ParentNode Is Nothing Then Return String.Empty

            Dim iIndex As Integer = 1
            Dim xnIndex As XmlNode = node
            While xnIndex.PreviousSibling IsNot Nothing
                iIndex += 1
                xnIndex = xnIndex.PreviousSibling
            End While
            'the path to a node is the path to its parent, plus "/node()[n]", where 
            'n is its position among its siblings.
            Return String.Format("{0}/node()[{1}]", GetXPathToNode(node.ParentNode), iIndex)
        End Function
    End Class
End Namespace
