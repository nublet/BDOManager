Imports Newtonsoft.Json

Namespace Models.BDODAE

    Public Class Node

        Private _Connections As New List(Of String)
        Private _Distances As New List(Of NodeDistance)
        Private _Items As New List(Of NodeItem)

        Public Property Cost As Integer = -1
        Public Property Region As String = ""
        Public Property NodeName As String = ""
        Public Property NodeType As String = ""
        Public Property WorkLoadBase As Integer = -1
        Public Property WorkLoadCurrent As Integer = -1

        Public ReadOnly Property Connections As List(Of String)
            Get
                Return _Connections
            End Get
        End Property

        Public ReadOnly Property Distances As List(Of NodeDistance)
            Get
                Return _Distances
            End Get
        End Property

        Public ReadOnly Property Items As List(Of NodeItem)
            Get
                Return _Items
            End Get
        End Property

    End Class

End Namespace