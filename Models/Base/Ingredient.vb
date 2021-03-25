Namespace Models

    Public Class Ingredient

        Public Property Amount As Integer = 0
        Public Property CheapestIngredient As String = ""
        Public Property CheapestIngredientPrice As Long = 0
        Public Property Name As String = ""

        Public Sub New()

        End Sub

        Public Sub New(ByRef template As LifeBDO.Ingredient, ByRef validIds As List(Of Integer))
            Dim IngredientOptions As New List(Of String)

            For Each Item As LifeBDO.Item In template.item
                validIds.Add(Item.id)
                IngredientOptions.Add(GetActualResources(Item.name).Trim())

                Dim OPItem As Models.OmegaPepega.Item = RequestManager.GetEntity(Of Models.OmegaPepega.Item)(30, "https://omegapepega.com/eu/{0}/0".FormatWith(Item.id))
                If OPItem Is Nothing Then
                    Continue For
                End If

                If CheapestIngredientPrice <= 0 OrElse CheapestIngredientPrice > OPItem.pricePerOne Then
                    CheapestIngredient = GetActualResources(Item.name).Trim()
                    CheapestIngredientPrice = OPItem.pricePerOne
                End If
            Next

            If template.item.Count > 1 AndAlso CheapestIngredient.IsNotSet() Then
                CheapestIngredient = GetActualResources(template.item(0).name).Trim()
            End If

            Me.Amount = template.amount
            If IngredientOptions.Count = 1 Then
                Me.Name = IngredientOptions(0)
            ElseIf IngredientOptions.Count > 0 Then
                Me.Name = GetIngredientOptionName(IngredientOptions)
            Else
                Me.Name = "Unknown"
            End If
        End Sub

    End Class

End Namespace