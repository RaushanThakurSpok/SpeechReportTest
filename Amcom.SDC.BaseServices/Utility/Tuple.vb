Public Class Tuple(Of T, U)
    Public Property Item1() As T
        Get
            Return m_Item1
        End Get
        Private Set(ByVal value As T)
            m_Item1 = value
        End Set
    End Property
    Private m_Item1 As T
    Public Property Item2() As U
        Get
            Return m_Item2
        End Get
        Private Set(ByVal value As U)
            m_Item2 = value
        End Set
    End Property
    Private m_Item2 As U

    Public Sub New(ByVal item1__1 As T, ByVal item2__2 As U)
        Item1 = item1__1
        Item2 = item2__2
    End Sub
End Class