Namespace Models

    Public Class Node

        Public Property CPCost As Integer = 0
        Public Property Name As String = ""
        Public Property Region As String = ""
        Public Property ResourceName As String = ""
        Public Property Resources As String = ""

        Public Sub New(cpCost As Integer, nodeName As String, nodeRegion As String, resources As String)
            resources = GetActualResources(resources).Trim()

            Me.CPCost = cpCost
            Me.Name = nodeName
            Me.Region = nodeRegion
            Me.Resources = resources
            Me.ResourceName = GetResourceNodeName(Me.Resources)

            If resources.ToLower().Contains("dried ") Then
                Me.Resources = resources.Replace("Dried ", "")
            End If
        End Sub

    End Class

End Namespace