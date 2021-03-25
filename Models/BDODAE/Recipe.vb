Imports Newtonsoft.Json

Namespace Models.BDODAE

    Public Class Recipe

        Private _Ingredients As New List(Of String)

        Public Property Category As String = ""
        Public Property Description As String = ""
        Public Property IsDropped As Boolean = False
        Public Property IsFarmed As Boolean = False
        Public Property IsFromNodes As Boolean = False
        Public Property IsGathered As Boolean = False
        Public Property Name As String = ""
        Public Property SubCategory As String = ""
        Public Property Vendor As String = ""
        Public Property VendorPrice As Long = 0
        Public Property Weight As Double = 0.0!

        Public ReadOnly Property Ingredients As List(Of String)
            Get
                Return _Ingredients
            End Get
        End Property

        Private Sub CheckCraftedIn()
            Dim TempDescription As String = Description.Trim().ToLower()

            Dim Temp As String = Description.Substring(TempDescription.IndexOf("crafted in ")).Trim()
            Temp = Temp.Substring(11)

            Category = "Production"
            SubCategory = Temp

            Description = Description.Substring(0, TempDescription.IndexOf("crafted in ")).Trim()
        End Sub

        Private Sub CheckFishing()
            Dim TempDescription As String = Description.Trim().ToLower()

            If TempDescription.IsEqualTo("white quality fish obtained by fishing") Then
                Category = "Gathering"
                SubCategory = "Fishing (White)"
            ElseIf TempDescription.IsEqualTo("green quality fish obtained by fishing") Then
                Category = "Gathering"
                SubCategory = "Fishing (Green)"
            ElseIf TempDescription.IsEqualTo("blue quality fish obtained by fishing") Then
                Category = "Gathering"
                SubCategory = "Fishing (Blue)"
            ElseIf TempDescription.IsEqualTo("yellow quality fish obtained by fishing") Then
                Category = "Gathering"
                SubCategory = "Fishing (Yellow)"
            End If

            Description = ""
        End Sub

        Private Sub CheckHarpooning()
            Dim TempDescription As String = Description.Trim().ToLower()

            If TempDescription.IsEqualTo("white quality fish obtained by harpooning") Then
                Category = "Gathering"
                SubCategory = "Harpooning (White)"
            ElseIf TempDescription.IsEqualTo("green quality fish obtained by harpooning") Then
                Category = "Gathering"
                SubCategory = "Harpooning (Green)"
            ElseIf TempDescription.IsEqualTo("blue quality fish obtained by harpooning") Then
                Category = "Gathering"
                SubCategory = "Harpooning (Blue)"
            ElseIf TempDescription.IsEqualTo("yellow quality fish obtained by harpooning") Then
                Category = "Gathering"
                SubCategory = "Harpooning (Yellow)"
            End If

            Description = ""
        End Sub

        Private Sub CheckProducedAt()
            Dim TempDescription As String = Description.Trim().ToLower()

            Dim Temp As String = Description.Substring(TempDescription.IndexOf("produced at ")).Trim()
            Temp = Temp.Substring(12)

            If Temp.ToLower().StartsWith("level ") Then
                Dim Level As String = Temp.Substring(6).Trim()
                Level = Level.Substring(0, Level.IndexOf(" "))

                SubCategory = "Level {0}".FormatWith(Level)
                Category = Temp.Substring(SubCategory.Length + 1).Trim()
            Else
                SubCategory = Temp
            End If

            Description = Description.Substring(0, TempDescription.IndexOf("produced at ")).Trim()
        End Sub

        Private Sub CheckProducedIn()
            Dim TempDescription As String = Description.Trim().ToLower()

            Dim Temp As String = Description.Substring(TempDescription.IndexOf("produced in ")).Trim()
            Temp = Temp.Substring(12)

            If Temp.ToLower().StartsWith("level ") Then
                Dim Level As String = Temp.Substring(6).Trim()
                Level = Level.Substring(0, Level.IndexOf(" "))

                SubCategory = "Level {0}".FormatWith(Level)
                Category = Temp.Substring(SubCategory.Length + 1).Trim()
            Else
                SubCategory = Temp
            End If

            Description = Description.Substring(0, TempDescription.IndexOf("produced in ")).Trim()
        End Sub

        Private Sub CheckVendor()
            Dim TempDescription As String = Description.Trim().ToLower()

            Dim TempPrice As String = ""
            If TempDescription.Contains(" for ") Then
                TempPrice = Description.Substring(TempDescription.IndexOf(" for ") + 5).Replace(","c, "").Replace("."c, "").Trim()
            End If

            If TempDescription.StartsWith("bought from ") Then
                Vendor = Description.Substring(TempDescription.IndexOf("bought from ") + 12).Trim()
                Description = ""
            ElseIf TempDescription.StartsWith("obtained from cooking vendors for ") Then
                Vendor = "Cooking"
                Description = ""
            ElseIf TempDescription.StartsWith("obtained from material vendors for ") Then
                Vendor = "Material"
                Description = ""
            ElseIf TempDescription.StartsWith("sold by ") Then
                Vendor = Description.Substring(TempDescription.IndexOf("sold by ") + 8).Trim()
            ElseIf TempDescription.Contains("sold by ") Then
                Vendor = Description.Substring(TempDescription.IndexOf("sold by ") + 8).Trim()
            Else
                AprBase.Extensions.System_String.ToLog("New Vendor Description found: {0}".FormatWith(Description), True)
            End If

            If Vendor.ToLower.Contains("amity ") Then
                Vendor = Vendor.Substring(0, Vendor.ToLower().IndexOf("amity ")).Trim()
            End If
            If Vendor.ToLower().Contains(" for ") Then
                Vendor = Vendor.Substring(0, Vendor.ToLower().IndexOf(" for ")).Trim()
            End If
            If Vendor.ToLower.Contains("vendors") Then
                Vendor = Replace(Vendor, "Vendors", "",,, CompareMethod.Text).Trim()
            End If

            If TempPrice.IsSet() Then
                VendorPrice = AprBase.Type.ToLongDB(TempPrice, 0)
            End If

            If Ingredients.Count <= 0 Then
                Category = "Vendor"
                SubCategory = Vendor
            End If
        End Sub

        Private Sub CheckDescription()
            Dim TempDescription As String = Description.Trim().ToLower()

            If TempDescription.Contains("dropped by mobs in ") Then
                Category = "Mob Drop"
                SubCategory = Description.Substring(TempDescription.IndexOf("dropped by mobs in ") + 18).Trim()
                If SubCategory.ToLower().Contains(" and ") Then
                    SubCategory = SubCategory.Substring(0, SubCategory.ToLower().IndexOf(" and ")).Trim()
                End If

                IsDropped = True
            ElseIf TempDescription.StartsWith("dropped from mobs in ") Then
                Category = "Mob Drop"
                SubCategory = Description.Substring(TempDescription.IndexOf("dropped from mobs in ") + 21).Trim()

                IsDropped = True
            End If

            If TempDescription.Contains("farmed") OrElse TempDescription.Contains("farming") Then
                IsFarmed = True
            End If

            If TempDescription.Contains("obtained from nodes") Then
                IsFromNodes = True
            End If

            If TempDescription.Contains("obtained from excavation node") Then
                IsFromNodes = True
            End If

            If TempDescription.Contains("obtained while fishing") Then
                Category = "Gathering"
                SubCategory = "Fishing"

                IsGathered = True
            End If

            If TempDescription.Contains("gathered") OrElse TempDescription.Contains("gathering") Then
                IsGathered = True
            End If

            If TempDescription.Contains("crafted in ") Then
                CheckCraftedIn()
            ElseIf TempDescription.Contains("obtained by fishing") Then
                CheckFishing()
            ElseIf TempDescription.Contains("obtained by harpooning") Then
                CheckHarpooning()
            ElseIf TempDescription.Contains("produced at ") Then
                CheckProducedAt()
            ElseIf TempDescription.Contains("produced in ") Then
                CheckProducedIn()
            End If

            If TempDescription.Contains("bought from ") OrElse TempDescription.Contains("sold by ") OrElse TempDescription.Contains(" vendors for ") Then
                CheckVendor()
            End If
        End Sub

        Public Sub CheckDetails()
            Try
                If Description.StartsWith("Furniture produced in ") Then
                    CheckDescription()
                ElseIf Description.IsEqualTo("Cooking vendor ingredient") Then
                    Description = ""
                    Vendor = "Cooking"

                    If Ingredients.Count <= 0 Then
                        Category = "Vendor"
                        SubCategory = Vendor
                    End If
                ElseIf Description.Contains(" or ") Then
                    Dim OriginalDescription As String = Description

                    For Each Current As String In Split(Description, " or ")
                        Description = Current.Trim()

                        If Description.StartsWith("or ") Then
                            Description = Description.Substring(3)
                        End If

                        CheckDescription()
                    Next

                    Description = OriginalDescription
                Else
                    CheckDescription()
                End If

                If IsDropped OrElse IsFarmed OrElse IsFromNodes OrElse IsGathered Then
                    Description = ""
                End If

                If (Vendor.IsSet() OrElse VendorPrice > 0) AndAlso Description.IsSet() Then
                    If Description.ToLower().Contains(" sold by ") Then
                        Description = Description.Substring(0, Description.ToLower().IndexOf("sold by ")).Trim()
                    ElseIf Description.ToLower().Contains("processed ") Then
                        Description = ""
                    ElseIf Description.ToLower().Contains("produced ") Then
                        Description = ""
                    ElseIf Description.ToLower().StartsWith("sold by ") Then
                        Description = ""
                    Else
                        Description = ""
                    End If
                End If

                If Description.StartsWith("+") Then
                    Description = " {0}".FormatWith(Description)
                End If

                Category = GetCategory(Category)
                SubCategory = GetSubCategory(SubCategory)

                If Description.Contains("Obtained by manufacture processing") Then
                    If Category.IsEqualTo("Processing") AndAlso SubCategory.IsEqualTo("Manufacture") Then
                        Description = Description.Replace("Obtained by manufacture processing", "").Trim()
                    Else
                        Description = Description
                    End If
                End If
            Catch ex As Exception
                ex.ToLog()
            End Try
        End Sub

#Region " Shared "

        Private Shared Function GetCategory(category As String) As String
            Dim Temp As String = category.ToLower()

            Return TextInfo.ToTitleCase(Temp)
        End Function

        Private Shared Function GetSubCategory(subCategory As String) As String
            Dim Temp As String = subCategory.ToLower()
            Temp = Temp.Replace("link ", "")
            Temp = Temp.Replace("list ", "")
            Temp = Temp.Replace(" button", "")
            Temp = Temp.Replace("sort_alchemy ", "")
            Temp = Temp.Replace("sort_cooking ", "")
            Temp = Temp.Replace("sort_material ", "")
            Temp = Temp.Replace("sort_processing ", "")
            Temp = Temp.Replace("sort_", "")
            Temp = Temp.Replace("_", " ").Trim()

            Return TextInfo.ToTitleCase(Temp)
        End Function

#End Region

    End Class

End Namespace