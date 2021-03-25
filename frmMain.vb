Imports System.Data.SqlTypes
Imports System.Diagnostics.SymbolStore
Imports System.Drawing.Drawing2D
Imports System.Net.Sockets
Imports System.Reflection
Imports System.Xml
Imports BDOManager.Models
Imports DevComponents.DotNetBar
Imports DevComponents.DotNetBar.Rendering
Imports Google.Apis.Sheets.v4.Data

Public Class frmMain

    Private _BackgroundWorker As ComponentModel.BackgroundWorker = Nothing
    Private _DocumentCompleted As Boolean = False

#Region " Form Events "

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        _Instance = Me
    End Sub

    Protected Overrides Sub Shutdown(ByRef e As FormClosingEventArgs)
        Me.ShutDown("")
        AprBase.CloseAndSave(Me)
    End Sub

    Protected Overrides Sub StartUp(ByRef stayHidden As Boolean)
        Me.StartUp("")

        Select Case AprBase.Settings.ComputerName
            Case "BJORN10"
                SetCacheLocations("C:\Projects\Cache\BDOData\")
            Case "NUBSERVER"
                SetCacheLocations("I:\Cache\BDOData\")
            Case "Study18"
                SetCacheLocations("D:\Projects\Cache\BDOData\")
            Case Else
                AprBase.Extensions.System_String.ToLog("frmMain->StartUp, unmatched Computer Name found: {0}".FormatWith(AprBase.Settings.ComputerName), True)
                Return
        End Select

        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3 Or System.Net.SecurityProtocolType.SystemDefault Or System.Net.SecurityProtocolType.Tls Or System.Net.SecurityProtocolType.Tls11 Or System.Net.SecurityProtocolType.Tls12

        _BackgroundWorker = New System.ComponentModel.BackgroundWorker()
        _BackgroundWorker.WorkerSupportsCancellation = True
        AddHandler _BackgroundWorker.DoWork, AddressOf BackgroundWorker_DoWork
        AddHandler _BackgroundWorker.RunWorkerCompleted, AddressOf BackgroundWorker_RunWorkerCompleted

        '_BackgroundWorker.RunWorkerAsync()
    End Sub

#End Region

#Region " JSON Updates "

    Private Sub AddLifeBDOItemIDs(ByRef validIds As List(Of Integer))
        Try
            AddStatus("   Getting LifeBDO Items...")
            Dim ItemList As Models.LifeBDO.ItemList = RequestManager.GetEntity(Of Models.LifeBDO.ItemList)(30, "https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/src/main/resources/items.json")
            For Each Current As Models.LifeBDO.Item In ItemList.items
                validIds.Add(Current.id)
            Next
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub AddBDOMarketplaceItems(ByRef validIds As List(Of Integer))
        Try
            AddStatus("   Getting BDO Marketplace Items...")
            For Each Current As Models.BDOMarketplace.Item In RequestManager.GetEntity(Of List(Of Models.BDOMarketplace.Item))(30, "https://raw.githubusercontent.com/kookehs/bdo-marketplace/master/items.json")
                validIds.Add(Current.id)
            Next
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub AddAlchemyRecipes(ByRef recipes As List(Of Models.Recipe), ByRef validIds As List(Of Integer))
        Try
            AddStatus("   Getting Alchemy Recipes...")
            Dim Recipes_Alchemy As Models.LifeBDO.RecipeList = RequestManager.GetEntity(Of Models.LifeBDO.RecipeList)(30, "https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/src/main/resources/recipesAlchemy.json")
            For Each Current As Models.LifeBDO.Recipe In Recipes_Alchemy.recipes
                recipes.Add(New Models.Recipe(Current, validIds))
            Next
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub AddCookingRecipes(ByRef recipes As List(Of Models.Recipe), ByRef validIds As List(Of Integer))
        Try
            AddStatus("   Getting Cooking Recipes...")
            Dim Recipes_Cooking As Models.LifeBDO.RecipeList = RequestManager.GetEntity(Of Models.LifeBDO.RecipeList)(30, "https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/src/main/resources/recipesCooking.json")
            For Each Current As Models.LifeBDO.Recipe In Recipes_Cooking.recipes
                recipes.Add(New Models.Recipe(Current, validIds))
            Next
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub AddProcessingRecipes(ByRef recipes As List(Of Models.Recipe), ByRef validIds As List(Of Integer))
        Try
            AddStatus("   Getting Processing Recipes...")
            Dim Recipes_Processing As Models.LifeBDO.RecipeList = RequestManager.GetEntity(Of Models.LifeBDO.RecipeList)(30, "https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/src/main/resources/recipesProcessing.json")
            For Each Current As Models.LifeBDO.Recipe In Recipes_Processing.recipes
                recipes.Add(New Models.Recipe(Current, validIds))
            Next
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub UpdateLifeSkills(ByRef validIds As List(Of Integer))
        Try
            Dim Recipes As New List(Of Models.Recipe)

            AddAlchemyRecipes(Recipes, validIds)

            AddCookingRecipes(Recipes, validIds)

            AddProcessingRecipes(Recipes, validIds)

            AddStatus("   Updating CraftPrices Sheet...")
            Dim RowDetails As New List(Of List(Of Object))
            For Each Current As Models.Recipe In Recipes.OrderBy(Function(o) o.Product)
                Dim RowData As New List(Of Object)

                RowData.Add(Current.Product)
                If Current.LifeSkill.IsEqualTo("Alchemy") OrElse Current.LifeSkill.IsEqualTo("Cooking") Then
                    RowData.Add(Current.LifeSkill)
                    RowData.Add(Current.LifeSkillLevel)
                    RowData.Add(Current.EXP)
                Else
                    RowData.Add("Processing")
                    RowData.Add(Current.LifeSkill)
                    RowData.Add(Current.EXP)
                End If
                RowData.Add("")

                For Each Ingredient As Models.Ingredient In Current.Ingredients.OrderBy(Function(o) o.Name)
                    RowData.Add(Ingredient.Amount)
                    RowData.Add(Ingredient.Name)
                    If Ingredient.Name.IsEqualTo(Ingredient.CheapestIngredient) Then
                        RowData.Add("")
                    Else
                        RowData.Add(Ingredient.CheapestIngredient)
                    End If
                Next
                Do While RowData.Count < 15
                    RowData.Add("")
                Loop

                RowDetails.Add(RowData)
            Next

            External.GoogleSheets.UpdateRange("CraftPrices!B4", External.GoogleSheets.SpreadSheetID, True, RowDetails)
            RowDetails.Clear()
            RowDetails.Add(New List(Of Object) From {AprBase.CurrentDate.GetSQLString("dd/MM/yy HH:mm")})
            External.GoogleSheets.UpdateRange("CraftPrices!B1", External.GoogleSheets.SpreadSheetID, True, RowDetails)
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub UpdateNodes(ByRef validIds As List(Of Integer))
        Try
            Dim Nodes As New List(Of Models.Node)

            AddStatus("   Getting Nodes...")
            Dim NodeList As Models.LifeBDO.NodeList = RequestManager.GetEntity(Of Models.LifeBDO.NodeList)(30, "https://raw.githubusercontent.com/Flockenberger/LifeBDO/master/LifeBDO/src/main/resources/nodes.json")
            For Each Current As Models.LifeBDO.Node In NodeList.nodes
                Nodes.Add(New Models.Node(Current.cpCost, Current.node, Current.region, ""))

                For Each NodeResource As LifeBDO.NodeResource In Current.subNodes
                    Dim ResourceNames As New List(Of String)

                    For Each Item As LifeBDO.Item In NodeResource.items
                        If Item Is Nothing Then
                            Continue For
                        End If
                        If Item.id <= 0 Then
                            Continue For
                        End If

                        ResourceNames.Add(Item.name)
                        validIds.Add(Item.id)
                    Next

                    Dim Resources As String = String.Join(", ", ResourceNames.Distinct().OrderBy(Function(o) o))

                    Nodes.Add(New Models.Node(NodeResource.cpCost, Current.node, Current.region, Resources))
                Next
            Next
            AddStatus("      Updating Sheet...")
            Dim RowDetails As New List(Of List(Of Object))
            For Each Current As Models.Node In Nodes.OrderBy(Function(o) o.ResourceName).ThenBy(Function(o) o.Region).ThenBy(Function(o) o.Name)
                Dim RowData As New List(Of Object)

                RowData.Add(GetName(Current.Region))
                RowData.Add(GetName(Current.Name))
                RowData.Add(Current.CPCost)
                RowData.Add("")
                RowData.Add(Current.ResourceName)
                RowData.Add(Current.Resources)
                RowData.Add("")

                RowDetails.Add(RowData)
            Next
            External.GoogleSheets.UpdateRange("Nodes!B4", External.GoogleSheets.SpreadSheetID, True, RowDetails)
            RowDetails.Clear()
            RowDetails.Add(New List(Of Object) From {AprBase.CurrentDate.GetSQLString("dd/MM/yy HH:mm")})
            External.GoogleSheets.UpdateRange("Nodes!C1", External.GoogleSheets.SpreadSheetID, True, RowDetails)
        Catch ex As Exception

        End Try
    End Sub

#End Region

    Private Sub AddStatus(message As String)
        Try
            lrMain.AddMessage("{0}: {1}".FormatWith(AprBase.CurrentDate.GetSQLString("HH:mm:ss"), message))
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Me.Opacity = 0.5

        Try
            Dim Nodes As New List(Of BDODAE.Node)

            _DocumentCompleted = False

            Dim URL As String = "https://www.bdodae.com/nodes/"

            wbMain.Navigate(URL)

            Do Until _DocumentCompleted
                System.Threading.Thread.Sleep(250)
                Application.DoEvents()
            Loop

            If wbMain.Document.Body Is Nothing Then
                Return
            End If

            Dim CurrentElement As HtmlElement = wbMain.Document.GetElementById("nodes_content")
            If CurrentElement Is Nothing Then
                Return
            End If

            For Each CurrentDiv As HtmlElement In CurrentElement.GetElementsByTagName("div")
                If CurrentDiv.Id.IsNotSet() Then
                    Continue For
                End If
                If CurrentDiv.Id.IsEqualTo("save_msg") Then
                    Continue For
                End If
                If CurrentDiv.Id.IsEqualTo("nw_msg") Then
                    Continue For
                End If

                If CurrentDiv.OuterHtml.Contains("class=""node list search sort_areas_on sort_nodes_on") Then
                    Dim NodeURL As String = "https://www.bdodae.com/nodes/index.php?node={0}".FormatWith(CurrentDiv.Id)

                    _DocumentCompleted = False

                    wbMain.Navigate(NodeURL)

                    Do Until _DocumentCompleted
                        System.Threading.Thread.Sleep(250)
                        Application.DoEvents()
                    Loop

                    If wbMain.Document.Body Is Nothing Then
                        Continue For
                    End If

                    Dim NodeElement As HtmlElement = wbMain.Document.GetElementById("node")
                    If NodeElement Is Nothing Then
                        Continue For
                    End If

                    Dim Node As New BDODAE.Node()

                    Node.NodeName = CurrentDiv.Id.Trim()
                    Node.NodeName = Node.NodeName.Replace("_"c, " "c)
                    Node.NodeName = TextInfo.ToTitleCase(Node.NodeName)

                    For Each Current As HtmlElement In NodeElement.Children
                        If Current.TagName.IsEqualTo("H3") Then
                            For Each Element As HtmlElement In Current.GetElementsByTagName("span")
                                Dim Value As String = Element.InnerText
                                Value = Value.Replace("CP", "").Trim()
                                Node.Cost = AprBase.Type.ToIntegerDB(Value)
                            Next
                        ElseIf Current.TagName.IsEqualTo("DIV") Then
                            If Current.OuterHtml.Contains("class=n_area") Then
                                Node.Region = Current.InnerText.Trim()
                            ElseIf Current.OuterHtml.Contains("class=n_type") Then
                                Node.NodeType = Current.InnerText.Trim()
                            ElseIf Current.OuterHtml.Contains("class=n_connected") Then
                                For Each Element As HtmlElement In Current.GetElementsByTagName("a")
                                    Node.Connections.Add(Element.InnerText.Trim())
                                Next
                            End If
                        End If
                    Next

                    Nodes.Add(Node)

                    Dim SubNodeElement As HtmlElement = wbMain.Document.GetElementById("subnodes")
                    If SubNodeElement Is Nothing Then
                        Continue For
                    End If

                    For Each sElement As HtmlElement In SubNodeElement.Children
                        Dim SubNode As New BDODAE.Node()

                        For Each sChildElement As HtmlElement In sElement.Children
                            If sChildElement.TagName.IsEqualTo("DIV") Then
                                If sChildElement.OuterHtml.Contains("class=s_title") Then
                                    For Each TitleElement As HtmlElement In sChildElement.Children
                                        If TitleElement.OuterHtml.Contains("class=s_subtype") Then
                                            Dim Temp As String = TitleElement.InnerHtml
                                            SubNode.NodeType = Temp.Substring(0, Temp.IndexOf("<")).Trim()
                                            Temp = Temp.Substring(Temp.IndexOf(">") + 1)
                                            Temp = Temp.Substring(0, Temp.IndexOf("<"))
                                            Temp = Temp.Replace("CP", "").Trim()
                                            SubNode.Cost = AprBase.Type.ToIntegerDB(Temp)
                                        ElseIf TitleElement.OuterHtml.Contains("class=s_workload") Then
                                            If TitleElement.InnerText.StartsWith("Base Workload:") Then
                                                Dim Temp As String = TitleElement.InnerText.Substring(TitleElement.InnerText.IndexOf(":"c) + 1).Trim()
                                                SubNode.WorkLoadBase = AprBase.Type.ToIntegerDB(Temp)
                                            ElseIf TitleElement.InnerText.StartsWith("Current Workload:") Then
                                                Dim Temp As String = TitleElement.InnerText.Substring(TitleElement.InnerText.IndexOf(":"c) + 1).Trim()
                                                SubNode.WorkLoadCurrent = AprBase.Type.ToIntegerDB(Temp)
                                            End If
                                        End If
                                    Next
                                End If
                            ElseIf sChildElement.TagName.IsEqualTo("TABLE") Then
                                If sChildElement.OuterHtml.Contains("class=s_items") Then
                                    For Each TableElement As HtmlElement In sChildElement.Children
                                        For Each RowElement As HtmlElement In TableElement.Children
                                            If RowElement.Children(0).InnerText.IsEqualTo("Avg") Then
                                                Continue For
                                            End If

                                            Dim NodeItem As New BDODAE.NodeItem()

                                            NodeItem.YieldAverage = AprBase.Type.ToDoubleDB(RowElement.Children(0).InnerText.Trim())
                                            NodeItem.ItemName = RowElement.Children(1).InnerText.Trim()
                                            NodeItem.YieldPerHour = AprBase.Type.ToDoubleDB(RowElement.Children(3).InnerText.Trim())

                                            SubNode.Items.Add(NodeItem)
                                        Next
                                    Next
                                ElseIf sChildElement.OuterHtml.Contains("class=s_distance") Then
                                    For Each TableElement As HtmlElement In sChildElement.Children
                                        For Each RowElement As HtmlElement In TableElement.Children
                                            If RowElement.Children(0).InnerText.IsEqualTo("Range") Then
                                                Continue For
                                            End If

                                            Dim NodeDistance As New BDODAE.NodeDistance()

                                            NodeDistance.Range = AprBase.Type.ToIntegerDB(RowElement.Children(0).InnerText.Trim())
                                            NodeDistance.City = RowElement.Children(2).InnerText.Trim()

                                            SubNode.Distances.Add(NodeDistance)
                                        Next

                                    Next
                                End If
                            End If
                        Next

                        SubNode.NodeName = "{0} - {1}".FormatWith(Node.NodeName, SubNode.NodeType)
                        SubNode.Region = Node.Region

                        Nodes.Add(SubNode)
                    Next
                End If
            Next

            wbMain.Navigate("")
            System.Threading.Thread.Sleep(5000)
            Application.DoEvents()

            Dim Cities As New List(Of String)
            Cities.AddRange(Nodes.SelectMany(Function(o) o.Distances.Select(Function(p) p.City)).Distinct().OrderBy(Function(o) o))

            Dim MaxItems As Integer = Nodes.Select(Function(o) o.Items.Count()).Max()

            Dim RowDetails As New List(Of List(Of Object))
            For Each Current As BDODAE.Node In Nodes.OrderBy(Function(o) o.Region).ThenBy(Function(o) o.NodeName)
                Dim RowData As New List(Of Object)

                RowData.Add(Current.Region)
                RowData.Add(Current.NodeName)
                RowData.Add(Current.Cost)

                Select Case Current.NodeType
                    Case "Excavation", "Farming", "Fish Drying Yard", "Gathering", "Lumbering", "Mining"
                        RowData.Add(1)
                    Case "Connection Node", "Town Node", "Worker Node"
                        RowData.Add(0)
                    Case Else
                        Current.NodeType = Current.NodeType
                End Select
                RowData.Add("")

                If Current.WorkLoadBase > 0 Then
                    RowData.Add(Current.WorkLoadBase)
                Else
                    RowData.Add("")
                End If
                If Current.WorkLoadCurrent > 0 Then
                    RowData.Add(Current.WorkLoadCurrent)
                Else
                    RowData.Add("")
                End If
                RowData.Add("")

                Dim Count As Integer = 0
                For Each NodeItem As BDODAE.NodeItem In Current.Items
                    RowData.Add(NodeItem.YieldAverage)
                    RowData.Add("")
                    RowData.Add(NodeItem.ItemName)

                    Count += 1
                Next
                Do While Count < MaxItems
                    RowData.Add("")
                    RowData.Add("")
                    RowData.Add("")
                    Count += 1
                Loop
                RowData.Add("")

                If Current.Connections.Count > 0 Then
                    RowData.Add(String.Join(", ", Current.Connections.Distinct().OrderBy(Function(o) o)))
                Else
                    RowData.Add("")
                End If
                RowData.Add("")

                For Each City As String In Cities
                    Dim NodeDistance As BDODAE.NodeDistance = Current.Distances.Where(Function(o) o.City.IsEqualTo(City)).FirstOrDefault()

                    If NodeDistance Is Nothing Then
                        RowData.Add("")
                    Else
                        RowData.Add(NodeDistance.Range)
                    End If
                Next
                RowData.Add("")

                RowDetails.Add(RowData)
            Next
            External.GoogleSheets.UpdateRange("Sheet38!C4", External.GoogleSheets.SpreadSheetID, True, RowDetails)

            RowDetails.Clear()
            Dim CityData As New List(Of Object)
            For Each City As String In Cities
                CityData.Add(City)
            Next
            RowDetails.Add(CityData)
            External.GoogleSheets.UpdateRange("Sheet38!AC3", External.GoogleSheets.SpreadSheetID, True, RowDetails)

            RowDetails.Clear()
            RowDetails.Add(New List(Of Object) From {AprBase.CurrentDate.GetSQLString("dd/MM/yy HH:mm")})
            External.GoogleSheets.UpdateRange("Sheet38!B1", External.GoogleSheets.SpreadSheetID, True, RowDetails)



            'Dim Recipes As New List(Of BDODAE.Recipe)

            'For Each Category As String In {"alchemy", "cooking", "processing", "production", "material"}
            '    _DocumentCompleted = False

            '    Dim URL As String = "https://www.bdodae.com/items/index.php?cat={0}".FormatWith(Category)

            '    wbMain.Navigate(URL)

            '    Do Until _DocumentCompleted
            '        System.Threading.Thread.Sleep(250)
            '        Application.DoEvents()
            '    Loop

            '    If wbMain.Document.Body Is Nothing Then
            '        Continue For
            '    End If

            '    Dim LinksElement As HtmlElement = wbMain.Document.GetElementById("links")
            '    If LinksElement IsNot Nothing Then
            '        For Each LinkDiv As HtmlElement In LinksElement.GetElementsByTagName("div")
            '            Dim Recipe As New BDODAE.Recipe()

            '            Recipe.Category = Category

            '            Recipe.SubCategory = LinkDiv.OuterHtml
            '            Try
            '                Recipe.SubCategory = Recipe.SubCategory.Substring(14, Recipe.SubCategory.IndexOf(">") - 15).Trim()
            '            Catch exInner As Exception
            '                Dim Test As String = ""
            '                Test = exInner.Message
            '            End Try

            '            If Recipe.SubCategory.ToLower.Contains("imperial") Then
            '                Continue For
            '            End If

            '            For Each ItemDiv As HtmlElement In LinkDiv.GetElementsByTagName("div")
            '                If ItemDiv.OuterHtml.Contains("class=link_title") Then
            '                    Recipe.Name = ItemDiv.InnerText
            '                ElseIf ItemDiv.OuterHtml.Contains("class=link_info") Then
            '                    Recipe.Description = ItemDiv.InnerText
            '                    If Recipe.Description.Contains(vbCrLf) Then
            '                        Recipe.Description = Recipe.Description.Substring(0, Recipe.Description.IndexOf(vbCrLf)).Trim()
            '                    End If

            '                    For Each ExtraDiv As HtmlElement In ItemDiv.GetElementsByTagName("div")
            '                        Dim IngredientString As String = ExtraDiv.InnerHtml
            '                        Dim IngredientStringLower As String = IngredientString.ToLower()

            '                        Do While IngredientStringLower.Contains("<a")
            '                            Dim IngredientValue As String = IngredientString.Substring(0, IngredientStringLower.IndexOf("<a")).Trim()
            '                            Dim IngredientName As String = IngredientString.Substring(IngredientStringLower.IndexOf(">") + 1)
            '                            IngredientName = IngredientName.Substring(0, IngredientName.IndexOf("<")).Trim()

            '                            Recipe.Ingredients.Add("{0} * {1}".FormatWith(IngredientName, IngredientValue))

            '                            IngredientString = IngredientString.Substring(IngredientStringLower.IndexOf("</a>") + 4)
            '                            IngredientStringLower = IngredientStringLower.Substring(IngredientStringLower.IndexOf("</a>") + 4)
            '                        Loop
            '                    Next
            '                End If
            '            Next

            '            If Recipe.Name.IsSet() Then
            '                If Recipe.Description.Contains("Obtained from quests and event rewards") Then
            '                    Continue For
            '                End If

            '                Recipe.CheckDetails()

            '                Recipes.Add(Recipe)
            '            End If
            '        Next
            '    End If

            '    wbMain.Navigate("")
            '    System.Threading.Thread.Sleep(5000)
            '    Application.DoEvents()
            'Next

            'For Each Current As BDODAE.Recipe In Recipes
            '    wbMain.Navigate("")
            '    System.Threading.Thread.Sleep(500)
            '    Application.DoEvents()

            '    _DocumentCompleted = False

            '    Dim URL As String = "https://www.bdodae.com/items/index.php?item={0}".FormatWith(Current.Name.ToLower().Replace("'"c, "").Replace(" "c, "_"c))

            '    wbMain.Navigate(URL)

            '    Do Until _DocumentCompleted
            '        System.Threading.Thread.Sleep(250)
            '        Application.DoEvents()
            '    Loop

            '    If wbMain.Document.Body Is Nothing Then
            '        Continue For
            '    End If

            '    Dim CurrentElement As HtmlElement = wbMain.Document.GetElementById("item_values")
            '    If CurrentElement IsNot Nothing Then
            '        For Each CurrentDiv As HtmlElement In CurrentElement.GetElementsByTagName("div")
            '            Dim ValueType As String = ""
            '            For Each ItemValueSpan As HtmlElement In CurrentDiv.GetElementsByTagName("span")
            '                ValueType = ItemValueSpan.InnerText
            '            Next
            '            Dim ValueAmount As String = Replace(CurrentDiv.InnerText, ValueType, "",,, CompareMethod.Text)
            '            Select Case ValueType.ToLower()
            '                Case "vendor sell:"
            '                Case "your value:"

            '                Case "trade value:"
            '                    Current.VendorPrice = AprBase.Type.ToLongDB(ValueAmount, 0)
            '                Case "vendor buy:"
            '                    Current.VendorPrice = AprBase.Type.ToLongDB(ValueAmount, 0)
            '                Case Else
            '                    AprBase.Extensions.System_String.ToLog("New ValueType found: {0}".FormatWith(ValueType.ToLower()), True)
            '            End Select
            '        Next
            '    End If

            '    CurrentElement = wbMain.Document.GetElementById("item_weight")
            '    If CurrentElement IsNot Nothing Then
            '        Dim Weight As String = CurrentElement.InnerText

            '        If Weight.Contains(":") Then
            '            Weight = Weight.Substring(Weight.IndexOf(":") + 3).Trim()
            '            Weight = Replace(Weight, " LT", "",,, CompareMethod.Text).Trim()
            '            Current.Weight = AprBase.Type.ToDoubleDB(Weight, 0.0!)
            '        ElseIf Weight.Contains(" - ") Then
            '            Weight = Weight.Substring(Weight.IndexOf(" - ") + 3).Trim()
            '            Weight = Replace(Weight, " LT", "",,, CompareMethod.Text).Trim()
            '            Current.Weight = AprBase.Type.ToDoubleDB(Weight, 0.0!)
            '        Else
            '            Weight = Weight
            '        End If
            '    End If

            '    CurrentElement = wbMain.Document.GetElementById("calc_main")
            '    If CurrentElement IsNot Nothing Then
            '        For Each CurrentDiv As HtmlElement In CurrentElement.GetElementsByTagName("div")
            '            If CurrentDiv.OuterHtml.Contains("<DIV class=calc_item_section>") Then
            '                For Each CalcItemDiv As HtmlElement In CurrentDiv.GetElementsByTagName("div")
            '                    If CalcItemDiv.OuterHtml.Contains("<DIV class=calc_item_craft>") Then
            '                        Dim Process As String = CalcItemDiv.InnerHtml
            '                        Process = Process.Substring(0, Process.IndexOf(" - ")).Trim()
            '                        If Current.Category.IsNotEqualTo(Process) Then
            '                            Process = Process

            '                            Current.Category = Process
            '                        End If

            '                        Dim Weight As String = CalcItemDiv.InnerText
            '                        If Weight.Contains(":") Then
            '                            Weight = Weight.Substring(Weight.IndexOf(":") + 3).Trim()
            '                            Weight = Replace(Weight, " LT", "",,, CompareMethod.Text).Trim()
            '                            Current.Weight = AprBase.Type.ToDoubleDB(Weight, 0.0!)
            '                        ElseIf Weight.Contains(" - ") Then
            '                            Weight = Weight.Substring(Weight.IndexOf(" - ") + 3).Trim()
            '                            Weight = Replace(Weight, " LT", "",,, CompareMethod.Text).Trim()
            '                            Current.Weight = AprBase.Type.ToDoubleDB(Weight, 0.0!)
            '                        Else
            '                            Weight = Weight
            '                        End If
            '                    End If
            '                Next
            '            End If
            '        Next
            '    End If

            '    CurrentElement = wbMain.Document.GetElementById("item_recipe")
            '    If CurrentElement IsNot Nothing Then
            '        If Current.Ingredients.Count <= 0 Then
            '            For Each CurrentDiv As HtmlElement In CurrentElement.Children
            '                If CurrentDiv.Id.IsEqualTo("recipe_options") Then
            '                    Continue For
            '                End If
            '                For Each ItemDiv As HtmlElement In CurrentDiv.GetElementsByTagName("div")
            '                    If ItemDiv.OuterHtml.Contains("<DIV class=""ing") AndAlso ItemDiv.OuterHtml.Contains("_item ing_item"">") Then
            '                        For Each ItemLink As HtmlElement In CurrentDiv.GetElementsByTagName("a")
            '                            Dim Ingredient As String = ItemLink.InnerText
            '                            If Ingredient.Contains(" (") Then
            '                                Dim IngredientName As String = Ingredient.Substring(0, Ingredient.IndexOf(" (")).Trim()
            '                                Dim IngredientValue As String = Ingredient.Substring(Ingredient.IndexOf(" (") + 2).Trim()
            '                                IngredientValue = IngredientValue.Substring(0, IngredientValue.IndexOf(")")).Trim()

            '                                Current.Ingredients.Add("{0} * {1}".FormatWith(IngredientName, IngredientValue))
            '                            Else
            '                                Ingredient = Ingredient
            '                            End If
            '                        Next
            '                    ElseIf ItemDiv.OuterHtml.Contains("<DIV class=ing") AndAlso ItemDiv.OuterHtml.Contains("_item ing_item>") Then
            '                        For Each ItemLink As HtmlElement In CurrentDiv.GetElementsByTagName("a")
            '                            Dim Ingredient As String = ItemLink.InnerText
            '                            If Ingredient.Contains(" (") Then
            '                                Dim IngredientName As String = Ingredient.Substring(0, Ingredient.IndexOf(" (")).Trim()
            '                                Dim IngredientValue As String = Ingredient.Substring(Ingredient.IndexOf(" (") + 2).Trim()
            '                                IngredientValue = IngredientValue.Substring(0, IngredientValue.IndexOf(")")).Trim()

            '                                Current.Ingredients.Add("{0} * {1}".FormatWith(IngredientName, IngredientValue))
            '                            Else
            '                                Ingredient = Ingredient
            '                            End If
            '                        Next
            '                    End If
            '                Next
            '            Next
            '        End If
            '    End If
            'Next

            'Dim ItemGroups As New List(Of BDODAE.Recipe)

            'Dim RowDetails As New List(Of List(Of Object))
            'For Each Current As BDODAE.Recipe In Recipes.OrderBy(Function(o) o.Name)
            '    If Current.Category.IsEqualTo("Material") Then
            '        If Current.SubCategory.IsEqualTo("Group") Then
            '            ItemGroups.Add(Current)
            '            Continue For
            '        End If
            '    End If

            '    Dim RowData As New List(Of Object)

            '    RowData.Add(Current.Name.Trim())
            '    RowData.Add(Current.Weight)
            '    RowData.Add("")
            '    RowData.Add(Current.Vendor.Trim())
            '    RowData.Add(Current.VendorPrice)
            '    RowData.Add("")
            '    RowData.Add(Current.Category.Trim())
            '    RowData.Add(Current.SubCategory.Trim())
            '    RowData.Add(Current.Description.Trim())
            '    RowData.Add("")
            '    RowData.Add(Current.IsDropped.GetString("Yes", ""))
            '    RowData.Add(Current.IsFarmed.GetString("Yes", ""))
            '    RowData.Add(Current.IsFromNodes.GetString("Yes", ""))
            '    RowData.Add(Current.IsGathered.GetString("Yes", ""))
            '    RowData.Add("")

            '    Dim Count As Integer = 0
            '    For Each Ingredient As String In Current.Ingredients.OrderBy(Function(o) o)
            '        Dim Ingredients As String() = Ingredient.Split("*"c)
            '        RowData.Add(Ingredients(1).Trim())
            '        RowData.Add(Ingredients(0).Trim())

            '        Count += 1
            '    Next

            '    Do While Count < 7
            '        RowData.Add("")
            '        RowData.Add("")
            '        Count += 1
            '    Loop
            '    RowData.Add("")

            '    RowDetails.Add(RowData)
            'Next
            'External.GoogleSheets.UpdateRange("BDODAE!B4", External.GoogleSheets.SpreadSheetID, True, RowDetails)

            ''RowDetails.Clear()
            ''For Each Current As BDODAE.Recipe In ItemGroups.OrderBy(Function(o) o.Name)
            ''    Dim RowData As New List(Of Object)

            ''    RowData.Add(Current.Name)
            ''    RowData.Add(Current.Description)

            ''    For Each Ingredient As String In Current.Ingredients.OrderBy(Function(o) o)
            ''        Dim Ingredients As String() = Ingredient.Split("*"c)
            ''        RowData.Add(Ingredients(1))
            ''        RowData.Add(Ingredients(0))
            ''    Next

            ''    RowDetails.Add(RowData)
            ''Next
            ''External.GoogleSheets.UpdateRange("BDODAE!AJ4", External.GoogleSheets.SpreadSheetID, True, RowDetails)

            'RowDetails.Clear()
            'RowDetails.Add(New List(Of Object) From {AprBase.CurrentDate.GetSQLString("dd/MM/yy HH:mm")})
            'External.GoogleSheets.UpdateRange("BDODAE!B1", External.GoogleSheets.SpreadSheetID, True, RowDetails)
        Catch ex As Exception
            ex.ToLog()
        Finally
            Me.Opacity = 1
        End Try

        '_BackgroundWorker.RunWorkerAsync()
    End Sub

    Private Sub wbMain_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles wbMain.DocumentCompleted
        _DocumentCompleted = True
    End Sub

    Private Sub BackgroundWorker_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs)
        Try
            Dim Test As Models.Ingredient = RequestManager.GetEntity(Of Models.Ingredient)(30, "https://www.bdodae.com/items/index.php?cat=alchemy")
            Return

            lrMain.ClearResults()
            AprBase.ProtocolErrors.Clear()
            RequestManager.JSONErrors.Clear()

            AddStatus("Updating BDO Details...")

            AddStatus("Clearing Cache...")
            DeleteJSONFolder("{0}BDOData\".FormatWith(JSONCacheLocation))

            Dim ItemIDs As New List(Of Integer)
            Dim Items As New List(Of Models.OmegaPepega.Item)

            ItemIDs.AddRange({4924, 6522, 6523, 9734, 11527})
            ItemIDs.AddRange({206, 207, 208, 212, 215, 504, 505, 506, 507, 508, 509, 512, 513, 514, 515, 516, 517, 518, 519, 520, 521, 522, 523, 524, 526, 528, 529, 531, 532, 566, 567, 568, 569, 570, 571, 575, 576, 578, 579, 586, 587, 588, 589, 591, 592, 593, 594, 595, 596, 597, 598, 601, 602, 603, 604, 605, 606, 607, 608, 609, 610, 611, 612, 613, 614, 615, 616, 617, 618, 619, 620, 621, 622, 623, 624, 625, 626, 627, 628, 629, 630, 631, 632, 633, 634, 635, 636, 637, 638, 639, 640, 641, 642, 643, 644, 645, 646, 647, 648, 649, 650, 651, 652, 653, 654, 655, 656, 657, 658, 659, 660, 661, 662, 663, 664, 665, 666, 667, 668, 669, 670, 671, 672, 673, 674, 675, 676, 677, 678, 679, 680, 681, 682, 683, 684, 685, 686, 687, 688, 689, 690, 691, 692, 693, 694, 695, 696, 697, 698, 699, 700, 701, 702, 703, 704, 705, 706, 707, 708, 709, 710, 711, 712, 713, 714, 715, 716, 717, 718, 719, 720, 721, 722, 723, 724, 725, 726, 727, 728, 729, 730, 732, 733, 734, 735, 736, 737, 738, 739, 740, 741, 744, 745, 748, 749, 750, 751, 752, 753, 754, 755, 756, 761, 762, 763, 764, 765, 766, 767, 768, 769, 770, 771, 773, 774, 775, 776, 777, 778, 779, 780, 781, 782, 783, 784, 785, 791, 792, 793, 794, 795, 796, 797, 2002, 2004, 2006, 2008, 2010, 2012, 2014, 2016, 2018, 2020, 2102, 2104, 2106, 2108, 2110, 2202, 2204, 2206, 2208, 2210, 2212, 2214, 2216, 2218, 2220, 2302, 2304, 2306, 2309, 2311, 2313, 2315, 2316, 2318, 2320, 2381, 2382, 2383, 2384, 2387, 2388, 2389, 2390, 2391, 2392, 2393, 2394, 2395, 2396, 2397, 2398, 2399, 2501, 2502, 2503, 2504, 2505, 2506, 2507, 2508, 2509, 2510, 2511, 2512, 2513, 2514, 2515, 2516, 2517, 2518, 2519, 2520, 2521, 2522, 2523, 2524, 2525, 2526, 2527, 2528, 2529, 2530, 2531, 2532, 2533, 2534, 2535, 2536, 2702, 2704, 2706, 2708, 2710, 2814, 2846, 3002, 3003, 3006, 3007, 3013, 3017, 3018, 3019, 3020, 3021, 3023, 3025, 3026, 3037, 3038, 3039, 3040, 3041, 3042, 3051, 3053, 3054, 3061, 3062, 3063, 3070, 3071, 3073, 3074, 3075, 3076, 3077, 3078, 3079, 3080, 3081, 3082, 3083, 3084, 3085, 3086, 3401, 3402, 3405, 3406, 3407, 3408, 3409, 3410, 3411, 3412, 3413, 3416, 3417, 3418, 3419, 3501, 3502, 3505, 3506, 3507, 3508, 3509, 3510, 3511, 3512, 3513, 3516, 3517, 3518, 3519, 3520, 3522, 3523, 3524, 3525, 3526, 3527, 3528, 3529, 3530, 3531, 3532, 3533, 3534, 3535, 3536, 3537, 3702, 3703, 3718, 3719, 3731, 3732, 3741, 3742, 4001, 4002, 4003, 4004, 4005, 4006, 4007, 4008, 4009, 4010, 4011, 4051, 4052, 4053, 4054, 4055, 4056, 4057, 4058, 4059, 4060, 4061, 4062, 4063, 4064, 4065, 4066, 4067, 4068, 4070, 4076, 4077, 4078, 4079, 4080, 4081, 4082, 4083, 4084, 4085, 4086, 4087, 4201, 4202, 4203, 4204, 4206, 4251, 4252, 4253, 4254, 4255, 4256, 4257, 4258, 4259, 4260, 4262, 4263, 4264, 4265, 4266, 4267, 4269, 4401, 4402, 4403, 4404, 4405, 4406, 4407, 4408, 4409, 4410, 4411, 4412, 4413, 4451, 4452, 4453, 4454, 4455, 4456, 4457, 4458, 4459, 4460, 4461, 4462, 4463, 4464, 4465, 4466, 4467, 4468, 4469, 4470, 4476, 4477, 4478, 4479, 4481, 4483, 4485, 4490, 4491, 4492, 4601, 4602, 4603, 4604, 4605, 4606, 4607, 4608, 4609, 4610, 4611, 4612, 4613, 4614, 4615, 4616, 4651, 4652, 4653, 4654, 4655, 4656, 4657, 4658, 4659, 4660, 4661, 4662, 4663, 4664, 4665, 4666, 4667, 4668, 4669, 4670, 4671, 4672, 4673, 4674, 4675, 4676, 4677, 4678, 4680, 4681, 4682, 4683, 4684, 4685, 4686, 4687, 4688, 4689, 4691, 4694, 4695, 4696, 4697, 4698, 4699, 4700, 4701, 4702, 4703, 4801, 4802, 4803, 4804, 4805, 4901, 4902, 4903, 4906, 4907, 4908, 4909, 4910, 4911, 4912, 4913, 4917, 4987, 4997, 4998, 4999, 5000, 5001, 5002, 5003, 5004, 5005, 5006, 5007, 5008, 5009, 5010, 5011, 5012, 5013, 5014, 5015, 5016, 5017, 5018, 5019, 5020, 5201, 5202, 5203, 5204, 5205, 5206, 5207, 5208, 5209, 5210, 5211, 5212, 5213, 5214, 5215, 5216, 5217, 5301, 5302, 5401, 5402, 5403, 5404, 5405, 5406, 5407, 5408, 5409, 5410, 5411, 5412, 5413, 5414, 5415, 5416, 5417, 5418, 5419, 5420, 5421, 5422, 5423, 5424, 5425, 5426, 5427, 5428, 5429, 5430, 5431, 5432, 5433, 5434, 5435, 5436, 5437, 5438, 5439, 5440, 5441, 5442, 5443, 5444, 5445, 5446, 5447, 5448, 5450, 5451, 5452, 5453, 5454, 5455, 5456, 5457, 5458, 5459, 5460, 5461, 5462, 5464, 5465, 5466, 5467, 5468, 5470, 5471, 5472, 5473, 5474, 5475, 5476, 5477, 5478, 5479, 5480, 5481, 5482, 5484, 5485, 5486, 5487, 5488, 5489, 5490, 5491, 5492, 5493, 5494, 5495, 5496, 5497, 5498, 5499, 5500, 5501, 5502, 5503, 5504, 5505, 5506, 5507, 5508, 5509, 5515, 5516, 5517, 5518, 5519, 5520, 5521, 5522, 5523, 5524, 5525, 5526, 5527, 5528, 5529, 5530, 5531, 5532, 5600, 5601, 5602, 5603, 5604, 5801, 5802, 5803, 5804, 5851, 5852, 5853, 5854, 5855, 5856, 5857, 5858, 5859, 5860, 5861, 5862, 5863, 5864, 5865, 5870, 5951, 5952, 5953, 5954, 5955, 5956, 5957, 5958, 5959, 5960, 5961, 5962, 5963, 6001, 6002, 6003, 6004, 6005, 6006, 6008, 6009, 6010, 6012, 6013, 6014, 6015, 6016, 6017, 6018, 6019, 6020, 6021, 6022, 6023, 6024, 6025, 6026, 6027, 6028, 6029, 6030, 6031, 6151, 6152, 6153, 6154, 6155, 6156, 6157, 6158, 6159, 6160, 6161, 6162, 6163, 6164, 6165, 6166, 6167, 6168, 6169, 6170, 6171, 6172, 6173, 6174, 6183, 6185, 6201, 6202, 6203, 6204, 6205, 6206, 6208, 6209, 6210, 6212, 6213, 6214, 6215, 6216, 6217, 6218, 6219, 6220, 6221, 6222, 6223, 6224, 6225, 6226, 6227, 6228, 6351, 6352, 6353, 6354, 6355, 6394, 6395, 6501, 6504, 6505, 6506, 6515, 6516, 6520, 6521, 6524, 6525, 6526, 6527, 6528, 6533, 6535, 6601, 6602, 6603, 6604, 6605, 6606, 6651, 6652, 6653, 6656, 6657, 6701, 6702, 6703, 6704, 6705, 6706, 6707, 6708, 6709, 6710, 6711, 6712, 6714, 6715, 6716, 6717, 6718, 6719, 6721, 6722, 6723, 6724, 6801, 6802, 6803, 6804, 6805, 6806, 6807, 6808, 6809, 6810, 6811, 6812, 6813, 6814, 6815, 6816, 6817, 6818, 6819, 6820, 6821, 6822, 6823, 6824, 6825, 6826, 6827, 6828, 6829, 6830, 6901, 6902, 6903, 6904, 6905, 6907, 6908, 6909, 6910, 6911, 6912, 6913, 6914, 6915, 6916, 6917, 6918, 6919, 6921, 6922, 6923, 6924, 6925, 6926, 6927, 6928, 6929, 6930, 6931, 6932, 6933, 6934, 6935, 6936, 6937, 6938, 6939, 6940, 6941, 6942, 6943, 6944, 6945, 6946, 6947, 6948, 6949, 6950, 6952, 6953, 6954, 6955, 6956, 6958, 6959, 6960, 6961, 6962, 6963, 6964, 6965, 6966, 6967, 6968, 6969, 6970, 6972, 6973, 6974, 6975, 6976, 6978, 6979, 6980, 6981, 6982, 6983, 6984, 6985, 6986, 6987, 6988, 6989, 6990, 6992, 6993, 6994, 6995, 6996, 6997, 6998, 6999, 7000, 7001, 7002, 7003, 7004, 7005, 7006, 7007, 7008, 7009, 7010, 7011, 7012, 7013, 7014, 7015, 7016, 7017, 7018, 7019, 7020, 7021, 7022, 7023, 7024, 7025, 7026, 7028, 7029, 7030, 7101, 7102, 7103, 7104, 7105, 7106, 7107, 7201, 7202, 7203, 7204, 7205, 7206, 7207, 7301, 7302, 7303, 7304, 7305, 7306, 7307, 7308, 7309, 7311, 7312, 7313, 7314, 7315, 7316, 7317, 7318, 7319, 7320, 7321, 7322, 7323, 7324, 7325, 7327, 7328, 7329, 7330, 7331, 7333, 7334, 7335, 7336, 7337, 7339, 7340, 7341, 7342, 7343, 7345, 7346, 7347, 7348, 7701, 7702, 7703, 7704, 7901, 7902, 7903, 7904, 7905, 7906, 7908, 7909, 7910, 7911, 7912, 7913, 7914, 7915, 7916, 7917, 7918, 7919, 7920, 7921, 7922, 7923, 7924, 7925, 7953, 7954, 7955, 7956, 7957, 7958, 8501, 8502, 8503, 8504, 8505, 8506, 8507, 8508, 8509, 8510, 8511, 8512, 8513, 8514, 8515, 8516, 8517, 8518, 8519, 8520, 8521, 8522, 8523, 8524, 8525, 8526, 8527, 8528, 8529, 8530, 8531, 8532, 8533, 8534, 8535, 8536, 8537, 8538, 8539, 8540, 8541, 8542, 8543, 8544, 8545, 8546, 8547, 8548, 8549, 8550, 8551, 8552, 8553, 8554, 8555, 8556, 8557, 8558, 8559, 8560, 8561, 8562, 8563, 8564, 8565, 8566, 8567, 8568, 8569, 8570, 8571, 8572, 8573, 8574, 8575, 8576, 8577, 8578, 8579, 8581, 8582, 8583, 8584, 8585, 8586, 8587, 8588, 8589, 8590, 8591, 8592, 8593, 8594, 8595, 8596, 8597, 8598, 8599, 8600, 8601, 8602, 8603, 8604, 8605, 8629, 8630, 8631, 8632, 8633, 8634, 8635, 8636, 8637, 8638, 8639, 8640, 8641, 8642, 8643, 8644, 8645, 8646, 8647, 8648, 8649, 8650, 8651, 8652, 8653, 8654, 8655, 8656, 8657, 8658, 9003, 9004, 9006, 9057, 9061, 9062, 9063, 9064, 9065, 9066, 9201, 9202, 9203, 9204, 9205, 9206, 9207, 9208, 9209, 9210, 9211, 9213, 9214, 9215, 9216, 9217, 9218, 9219, 9220, 9241, 9255, 9256, 9257, 9258, 9259, 9260, 9261, 9262, 9263, 9264, 9265, 9266, 9267, 9268, 9269, 9270, 9271, 9273, 9274, 9275, 9276, 9277, 9278, 9279, 9280, 9281, 9282, 9283, 9284, 9285, 9286, 9287, 9288, 9289, 9290, 9291, 9292, 9293, 9294, 9295, 9296, 9297, 9298, 9299, 9300, 9301, 9302, 9303, 9304, 9305, 9306, 9307, 9308, 9309, 9310, 9311, 9312, 9313, 9314, 9315, 9316, 9317, 9318, 9319, 9321, 9401, 9402, 9403, 9404, 9405, 9406, 9407, 9408, 9409, 9410, 9411, 9412, 9413, 9414, 9415, 9416, 9417, 9418, 9419, 9420, 9421, 9422, 9423, 9424, 9425, 9426, 9427, 9428, 9429, 9430, 9431, 9432, 9433, 9434, 9435, 9436, 9437, 9438, 9439, 9440, 9441, 9442, 9443, 9444, 9445, 9446, 9447, 9448, 9449, 9450, 9451, 9452, 9453, 9454, 9455, 9456, 9457, 9458, 9459, 9460, 9461, 9462, 9463, 9464, 9469, 9470, 9471, 9472, 9473, 9474, 9475, 9476, 9477, 9478, 9479, 9480, 9483, 9484, 9486, 9487, 9488, 9489, 9490, 9491, 9492, 9493, 9494, 9495, 9496, 9601, 9602, 9603, 9604, 9605, 9606, 9607, 9608, 9609, 9610, 9631, 9632, 9634, 9635, 9636, 9637, 9691, 9692, 9693, 9694, 9701, 9702, 9726, 9727, 9728, 9729, 9731, 9732, 9733, 9735, 9736, 9737, 9741, 9746, 9747, 9748, 9768, 9771, 9773, 10003, 10005, 10006, 10007, 10009, 10010, 10012, 10013, 10014, 10056, 10057, 10071, 10072, 10086, 10102, 10103, 10105, 10124, 10125, 10138, 10140, 10203, 10205, 10206, 10207, 10209, 10210, 10212, 10213, 10214, 10256, 10257, 10271, 10272, 10286, 10303, 10304, 10305, 10324, 10325, 10338, 10340, 10403, 10405, 10406, 10407, 10409, 10410, 10412, 10413, 10414, 10456, 10457, 10471, 10472, 10486, 10503, 10504, 10505, 10524, 10525, 10538, 10540, 10603, 10605, 10606, 10607, 10609, 10610, 10612, 10613, 10614, 10656, 10657, 10671, 10672, 10686, 10703, 10704, 10705, 10724, 10725, 10738, 10740, 10809, 10810, 10811, 10812, 10813, 10814, 10815, 10816, 10817, 10818, 10819, 10820, 10821, 10822, 10823, 10824, 10889, 10890, 10891, 10892, 10933, 10934, 10935, 10936, 10937, 10938, 10939, 10940, 11001, 11002, 11003, 11004, 11005, 11006, 11007, 11008, 11009, 11010, 11011, 11012, 11013, 11014, 11015, 11016, 11017, 11070, 11071, 11072, 11073, 11101, 11102, 11103, 11203, 11205, 11206, 11207, 11209, 11210, 11212, 11213, 11214, 11220, 11221, 11222, 11223, 11287, 11303, 11304, 11305, 11324, 11325, 11338, 11340, 11353, 11355, 11356, 11357, 11359, 11360, 11362, 11363, 11364, 11370, 11371, 11372, 11373, 11436, 11601, 11602, 11603, 11604, 11605, 11606, 11607, 11609, 11610, 11611, 11613, 11625, 11628, 11629, 11630, 11631, 11701, 11702, 11703, 11704, 11705, 11706, 11707, 11708, 11709, 11710, 11711, 11712, 11713, 11714, 11715, 11716, 11717, 11718, 11719, 11720, 11721, 11722, 11723, 11724, 11725, 11726, 11727, 11728, 11729, 11730, 11801, 11802, 11803, 11804, 11805, 11806, 11807, 11808, 11810, 11811, 11815, 11816, 11817, 11827, 11828, 11834, 11853, 11901, 11902, 11903, 11904, 11905, 11906, 11907, 11908, 11909, 11910, 11911, 11912, 11913, 11914, 11915, 11926, 11927, 11928, 12001, 12002, 12003, 12004, 12005, 12006, 12007, 12008, 12010, 12011, 12012, 12017, 12018, 12031, 12032, 12042, 12045, 12059, 12060, 12061, 12101, 12102, 12103, 12104, 12105, 12111, 12112, 12113, 12114, 12115, 12126, 12128, 12201, 12202, 12203, 12204, 12205, 12206, 12208, 12210, 12211, 12212, 12220, 12229, 12230, 12236, 12237, 12251, 12301, 12302, 12303, 12304, 12305, 12308, 12309, 12310, 12313, 12314, 12315, 12318, 12319, 12320, 12323, 12324, 12325, 12326, 12327, 12328, 12329, 12330, 12501, 12502, 12503, 12504, 12505, 12506, 12507, 12508, 12509, 12510, 12511, 12512, 12525, 12526, 12530, 12533, 12536, 12562, 12563, 12564, 12565, 12567, 12574, 12575, 12576, 12577, 12578, 12579, 12580, 12581, 12590, 12591, 12592, 12593, 12594, 12595, 12596, 12597, 12598, 12599, 12600, 12601, 12602, 12619, 12620, 12621, 12622, 12623, 12629, 12630, 12631, 12632, 12633, 12634, 12635, 12636, 12646, 12647, 12648, 12649, 12650, 12651, 12652, 12653, 12654, 12664, 12667, 12668, 12669, 12671, 12672, 12673, 12674, 12675, 12676, 12677, 12697, 12698, 12699, 12700, 12701, 12702, 12703, 12704, 12705, 12715, 12717, 12718, 12719, 12721, 12722, 12723, 12724, 12725, 12726, 12727, 12831, 13003, 13004, 13005, 13024, 13025, 13038, 13040, 13103, 13104, 13105, 13124, 13125, 13138, 13140, 13203, 13205, 13206, 13207, 13209, 13210, 13212, 13213, 13214, 13256, 13257, 13271, 13272, 13286, 13303, 13305, 13306, 13307, 13309, 13310, 13312, 13313, 13314, 13356, 13357, 13371, 13372, 13386, 13403, 13405, 13406, 13407, 13409, 13410, 13412, 13413, 13414, 13420, 13421, 13422, 13423, 13487, 13503, 13504, 13505, 13524, 13525, 13538, 13540, 13703, 13705, 13706, 13707, 13709, 13710, 13712, 13713, 13714, 13756, 13757, 13771, 13772, 13786, 13803, 13804, 13805, 13824, 13825, 13838, 13840, 13901, 13902, 13903, 14019, 14020, 14021, 14022, 14023, 14025, 14026, 14027, 14028, 14029, 14101, 14102, 14103, 14104, 14201, 14202, 14203, 14204, 14205, 14206, 14207, 14208, 14211, 14212, 14213, 14214, 14215, 14216, 14217, 14218, 14219, 14220, 14221, 14222, 14223, 14284, 14285, 14286, 14287, 14288, 14289, 14290, 14291, 14292, 14293, 14294, 14295, 14296, 14297, 14298, 14299, 14300, 14339, 14340, 14341, 14345, 14346, 14403, 14405, 14406, 14407, 14409, 14410, 14412, 14413, 14414, 14456, 14457, 14471, 14472, 14486, 14503, 14504, 14505, 14524, 14525, 14538, 14540, 14603, 14604, 14605, 14624, 14625, 14638, 14640, 14701, 14702, 14703, 14711, 14712, 14713, 14721, 14722, 14723, 14731, 14732, 14733, 14741, 14742, 14743, 14751, 14752, 14753, 14761, 14762, 14763, 14771, 14772, 14773, 14781, 14782, 14783, 14791, 14792, 14793, 14801, 14802, 14803, 14811, 14812, 14813, 14816, 14817, 14818, 14821, 14822, 14823, 14829, 14830, 14831, 14838, 14839, 14840, 14841, 14842, 14843, 14844, 14845, 14846, 14847, 14848, 14849, 14850, 14851, 14852, 14853, 14854, 14855, 14856, 14869, 14870, 14871, 14872, 14873, 14874, 14875, 14876, 14877, 14878, 14879, 14880, 15001, 15002, 15003, 15004, 15005, 15006, 15007, 15008, 15009, 15010, 15011, 15012, 15013, 15014, 15015, 15016, 15017, 15018, 15019, 15020, 15021, 15022, 15023, 15024, 15025, 15026, 15028, 15029, 15030, 15032, 15034, 15035, 15036, 15037, 15038, 15039, 15040, 15041, 15042, 15043, 15044, 15045, 15101, 15102, 15103, 15104, 15105, 15106, 15107, 15108, 15109, 15110, 15111, 15112, 15113, 15114, 15115, 15116, 15117, 15118, 15119, 15120, 15121, 15122, 15123, 15124, 15125, 15126, 15127, 15128, 15129, 15130, 15131, 15132, 15133, 15134, 15135, 15136, 15137, 15138, 15139, 15146, 15147, 15148, 15149, 15150, 15151, 15152, 15153, 15154, 15201, 15206, 15207, 15211, 15212, 15213, 15214, 15216, 15217, 15218, 15219, 15221, 15222, 15224, 15601, 15602, 15603, 15604, 15605, 15606, 15610, 15613, 15614, 15616, 15624, 15626, 15627, 15628, 15629, 15630, 15631, 15632, 15633, 15634, 15635, 15636, 15637, 15638, 15639, 15640, 15649, 15650, 15651, 15654, 15662, 15663, 15664, 15665, 15666, 15667, 15669, 15670, 15672, 15674, 15801, 15802, 15803, 15804, 15805, 15806, 15807, 15808, 15809, 15810, 15811, 15812, 15813, 15814, 15815, 15816, 15817, 15818, 16001, 16002, 16004, 16005, 16017, 16102, 16103, 16104, 16106, 16107, 16108, 16110, 16111, 16112, 16114, 16115, 16116, 16118, 16119, 16120, 16122, 16123, 16124, 16126, 16127, 16128, 16147, 16150, 16151, 16152, 16153, 16154, 16155, 16157, 16158, 16159, 16160, 16161, 16162, 16163, 16164, 16165, 16167, 16168, 16242, 16380, 16479, 16481, 16482, 16486, 16487, 16540, 16805, 16806, 16807, 16808, 16809, 16810, 16811, 16812, 16813, 16814, 16815, 16816, 16817, 16818, 16819, 16820, 16821, 16822, 16823, 16824, 16825, 16826, 16827, 16828, 16829, 16830, 16831, 16832, 16833, 16834, 16847, 16901, 16902, 16903, 17081, 17128, 17200, 17272, 17312, 17313, 17315, 17316, 17354, 17602, 17603, 17604, 17605, 17606, 17607, 17608, 17611, 17646, 17726, 17727, 17904, 17905, 17908, 17909, 17910, 17939, 17940, 17941, 17942, 17944, 17945, 17946, 17951, 17953, 17959, 17961, 17962, 17963, 17964, 17965, 17966, 17967, 17968, 17970, 17971, 17972, 17973, 17974, 17976, 17977, 18001, 18002, 18003, 18004, 18005, 18006, 18007, 18008, 18011, 18012, 18013, 18014, 18015, 18016, 18017, 18067, 18068, 18081, 18097, 18098, 18100, 18101, 18102, 18103, 18104, 18105, 18106, 18107, 18108, 18111, 18112, 18113, 18114, 18115, 18116, 18117, 18201, 18202, 18203, 18204, 18205, 18206, 18207, 18208, 18211, 18212, 18213, 18214, 18215, 18216, 18217, 18301, 18302, 18303, 18304, 18305, 18306, 18307, 18308, 18311, 18312, 18313, 18314, 18315, 18316, 18317, 18374, 18375, 18376, 18377, 18378, 18410, 18421, 18425, 18426, 18427, 18428, 18429, 18435, 18436, 18439, 18480, 18482, 18483, 18484, 18485, 18927, 18946, 19002, 19003, 19004, 19005, 19006, 19007, 19008, 19009, 19015, 19016, 19017, 19018, 19021, 19022, 19023, 19024, 19025, 19026, 19027, 19028, 19031, 19032, 19033, 19038, 19039, 19040, 19042, 19043, 19044, 19045, 19048, 19049, 19050, 19082, 19083, 19084, 19085, 19086, 19088, 19089, 19090, 19092, 19093, 19094, 19095, 19096, 19098, 19099, 19100, 19101, 19102, 19103, 19104, 19105, 19106, 19107, 19108, 19109, 19110, 19111, 19112, 19117, 19118, 19119, 19120, 19125, 19126, 19127, 19128, 19131, 19132, 19133, 19134, 19135, 19136, 19137, 19138, 19139, 19140, 19143, 19144, 19145, 19146, 19151, 19152, 19153, 19154, 19181, 19182, 19183, 19184, 19186, 19188, 19189, 19190, 19191, 19192, 19193, 19194, 19196, 19198, 19199, 19200, 19201, 19202, 19203, 19204, 19205, 19206, 19207, 19208, 19209, 19218, 19219, 19220, 19221, 19222, 19223, 19224, 19225, 19229, 19230, 19231, 19233, 19234, 19235, 19236, 19237, 19238, 19239, 19243, 19244, 19245, 19246, 19251, 19252, 19253, 19254, 19281, 19282, 19283, 19285, 19286, 19288, 19289, 19290, 19291, 19292, 19293, 19295, 19296, 19298, 19299, 19300, 19301, 19302, 19303, 19304, 19305, 19306, 19307, 19308, 19309, 19310, 19311, 19312, 19321, 19322, 19323, 19324, 19325, 19326, 19327, 19328, 19331, 19332, 19333, 19334, 19335, 19336, 19337, 19338, 19339, 19340, 19342, 19343, 19344, 19345, 19348, 19349, 19350, 19381, 19382, 19383, 19385, 19386, 19388, 19389, 19390, 19391, 19392, 19393, 19395, 19396, 19398, 19399, 19400, 19401, 19402, 19403, 19404, 19405, 19406, 19407, 19408, 19409, 19410, 19412, 19417, 19418, 19419, 19420, 19421, 19422, 19423, 19424, 19426, 19428, 19429, 19430, 19431, 19432, 19433, 19434, 19435, 19440, 19441, 19442, 19443, 19444, 19445, 19446, 19481, 19482, 19483, 19485, 19486, 19487, 19488, 19489, 19490, 19491, 19492, 19493, 19495, 19496, 19497, 19498, 19499, 19500, 19501, 19502, 19504, 19509, 19510, 19511, 19512, 19518, 19519, 19520, 19521, 19522, 19523, 19524, 19526, 19527, 19528, 19529, 19530, 19531, 19532, 19535, 19536, 19537, 19540, 19581, 19583, 19585, 19586, 19587, 19588, 19591, 19593, 19595, 19596, 19597, 19923, 20101, 20102, 20103, 20104, 20111, 20112, 20113, 20114, 20115, 20116, 20117, 20118, 20123, 20124, 20125, 20130, 20131, 20132, 20133, 20134, 20135, 20136, 20137, 20138, 20139, 20140, 20181, 20183, 20185, 20186, 20187, 20188, 20191, 20193, 20195, 20196, 20197, 20198, 20199, 20200, 20301, 20302, 20303, 20304, 20310, 20315, 20316, 20317, 20318, 20319, 20320, 20321, 20322, 20323, 20324, 20325, 20326, 20327, 20328, 20329, 20330, 20331, 20332, 20333, 20334, 20335, 20381, 20383, 20384, 20385, 20386, 20387, 20391, 20393, 20394, 20395, 20396, 20397, 21003, 21008, 21012, 21018, 21019, 21021, 21022, 21023, 21024, 21025, 21026, 21031, 21032, 21033, 21034, 21035, 21036, 21038, 21040, 21041, 21042, 21045, 21046, 21047, 21050, 21052, 21053, 21054, 21055, 21056, 21057, 21058, 21060, 21061, 21062, 21064, 21065, 21066, 21069, 21071, 21072, 21073, 21074, 21075, 21076, 21077, 21078, 21080, 21081, 21082, 21084, 21085, 21086, 21089, 21091, 21092, 21093, 21094, 21095, 21096, 21098, 21100, 21101, 21102, 21103, 21104, 21108, 21109, 21110, 21111, 21112, 21113, 21114, 21115, 21116, 21118, 21119, 21121, 21122, 21126, 21127, 21128, 21129, 21130, 21131, 21132, 21134, 21136, 21137, 21138, 21140, 21144, 21145, 21146, 21147, 21150, 21151, 21152, 21153, 21154, 21155, 21156, 21157, 21159, 21164, 21165, 21166, 21167, 21168, 21169, 21170, 21171, 21172, 21173, 21175, 21176, 21178, 21185, 21186, 21187, 21188, 21189, 21191, 21193, 21194, 21195, 21196, 21198, 21199, 21205, 21206, 21207, 21208, 21209, 21210, 21212, 21220, 21221, 21223, 21228, 21229, 21230, 21236, 21237, 21238, 21239, 21240, 21241, 21242, 21243, 21244, 21245, 21246, 21247, 21250, 21251, 21253, 21254, 21256, 21257, 21260, 21261, 21262, 21263, 21264, 21274, 21275, 21279, 21280, 21282, 21283, 21284, 21285, 21308, 21309, 21313, 21314, 21315, 21316, 21317, 21318, 21319, 21320, 21321, 21322, 21323, 21324, 21338, 21339, 21341, 21342, 21344, 21345, 21346, 21349, 21350, 21352, 21353, 21355, 21356, 21358, 21359, 21361, 21362, 21364, 21365, 21366, 21370, 21371, 21373, 21374, 21376, 21377, 21379, 21380, 21382, 21383, 21385, 21386, 21388, 21389, 21391, 21392, 21394, 21395, 21397, 21398, 21399, 21400, 21604, 21605, 21606, 21607, 21608, 21609, 21622, 21623, 21624, 21625, 21626, 21627, 21628, 21629, 21630, 21631, 21632, 21636, 21637, 21643, 21644, 21645, 21646, 21647, 21648, 21649, 21650, 21651, 21652, 21653, 21654, 21656, 21658, 21659, 21660, 21661, 21662, 21663, 21664, 21665, 21666, 21667, 21668, 21669, 21670, 21671, 21684, 21685, 21688, 21689, 21691, 21692, 21694, 21695, 21696, 21698, 21699, 21701, 21702, 21703, 21705, 21706, 21707, 21719, 21720, 21722, 21723, 21724, 21725, 21730, 21732, 21733, 21735, 21736, 21737, 21743, 21744, 21745, 21746, 21747, 21748, 21749, 21750, 21751, 21752, 21753, 21754, 21760, 21761, 21762, 21763, 21764, 21765, 21766, 21767, 21770, 21771, 21772, 21777, 21778, 21779, 21780, 21781, 21782, 21783, 21784, 21785, 21786, 21787, 21788, 21789, 21792, 21793, 21794, 21795, 21796, 21797, 21798, 21799, 21800, 21801, 21802, 21803, 21804, 21805, 21806, 21807, 21808, 21809, 21810, 21811, 21812, 21813, 21814, 21815, 21841, 21842, 21843, 21844, 21845, 21846, 21847, 21848, 21849, 21850, 21901, 21902, 21904, 21905, 21907, 21908, 21909, 21910, 21911, 21912, 21913, 21914, 21915, 21916, 21917, 21918, 21919, 21920, 21921, 21922, 21923, 21924, 21925, 21965, 21966, 21967, 21970, 21973, 21974, 21975, 22001, 22002, 22003, 22004, 22008, 22009, 22010, 22011, 22012, 22013, 22022, 22023, 22025, 22027, 22028, 22029, 22031, 22032, 22033, 22034, 22035, 22036, 22037, 22038, 22043, 22044, 22045, 22046, 22047, 22048, 22049, 22050, 22051, 22052, 22053, 22054, 22055, 22056, 22061, 22062, 22063, 22064, 22065, 22066, 22067, 22068, 22069, 22071, 22072, 22074, 22075, 22079, 22081, 22082, 22084, 22085, 22086, 22087, 22089, 22090, 22094, 22096, 22097, 22098, 22099, 22100, 22101, 22102, 22103, 22104, 22105, 22106, 22107, 22108, 22113, 22114, 22115, 22116, 22117, 22118, 22130, 22131, 22132, 22135, 22136, 22137, 22138, 22139, 22140, 22141, 22142, 22143, 22144, 22145, 22148, 22149, 22150, 22151, 22152, 22153, 22154, 22155, 22156, 22157, 22158, 22159, 22160, 22161, 22162, 22163, 22164, 22165, 22166, 22167, 22168, 22169, 22170, 22171, 22172, 22174, 22175, 22176, 22181, 22182, 22185, 22186, 22187, 22189, 22190, 22191, 22196, 22197, 22198, 22199, 22200, 22201, 22202, 22203, 22204, 22205, 22206, 22207, 22208, 22213, 22214, 22215, 22216, 22217, 22218, 22225, 22231, 22232, 22233, 22236, 22237, 22238, 22239, 22240, 22241, 22242, 22243, 22244, 22245, 22246, 22247, 22248, 22253, 22254, 22255, 22256, 22257, 22258, 22259, 22260, 22261, 22262, 22263, 22264, 22265, 22266, 22269, 22270, 22271, 22272, 22274, 22275, 22278, 22281, 22282, 22283, 22286, 22287, 22289, 22290, 22293, 22296, 22297, 22298, 22300, 22301, 22303, 22304, 22308, 22309, 22310, 22322, 22323, 22325, 22327, 22329, 22330, 22331, 22332, 22333, 22334, 22335, 22338, 22339, 22340, 22341, 22344, 22345, 22346, 22347, 22348, 22349, 22350, 22351, 22352, 22353, 22354, 22355, 22356, 22357, 22358, 22362, 22363, 22364, 22365, 22366, 22367, 22368, 22369, 22370, 22371, 22372, 22374, 22375, 22379, 22381, 22382, 22383, 22386, 22387, 22389, 22390, 22394, 22396, 22397, 22398, 22399, 22400, 22401, 22402, 22403, 22404, 22405, 22406, 22407, 22408, 22413, 22414, 22415, 22416, 22417, 22418, 22429, 22430, 22431, 22434, 22435, 22436, 22437, 22438, 22439, 22440, 22441, 22442, 22443, 22444, 22445, 22446, 22447, 22453, 22454, 22455, 22456, 22457, 22458, 22459, 22460, 22461, 22462, 22463, 22464, 22465, 22466, 22469, 22470, 22471, 22472, 22474, 22475, 22481, 22482, 22483, 22486, 22487, 22489, 22490, 22496, 22497, 22498, 22500, 22501, 22502, 22503, 22504, 22508, 22509, 22510, 22511, 22512, 22513, 22522, 22523, 22525, 22527, 22529, 22530, 22531, 22532, 22533, 22534, 22535, 22536, 22542, 22543, 22544, 22546, 22547, 22548, 22549, 22550, 22551, 22552, 22553, 22554, 22555, 22556, 22557, 22558, 22559, 22560, 22563, 22564, 22565, 22566, 22567, 22568, 22569, 22570, 22571, 22572, 22574, 22575, 22579, 22581, 22582, 22586, 22587, 22589, 22590, 22594, 22596, 22597, 22599, 22600, 22601, 22602, 22603, 22604, 22605, 22606, 22607, 22608, 22613, 22614, 22615, 22616, 22617, 22618, 22629, 22630, 22631, 22633, 22634, 22638, 22639, 22640, 22641, 22642, 22643, 22644, 22645, 22646, 22650, 22652, 22653, 22654, 22655, 22658, 22659, 22660, 22661, 22662, 22663, 22664, 22665, 22666, 22667, 22668, 22671, 22672, 22674, 22675, 22680, 22681, 22682, 22684, 22685, 22686, 22687, 22689, 22690, 22695, 22696, 22697, 22699, 22700, 22701, 22702, 22703, 22704, 22705, 22706, 22707, 22708, 22713, 22714, 22715, 22716, 22717, 22718, 22728, 22729, 22730, 22732, 22733, 22734, 22737, 22738, 22739, 22740, 22741, 22742, 22743, 22744, 22745, 22746, 22747, 22748, 22751, 22752, 22753, 22755, 22756, 22757, 22758, 22759, 22760, 22761, 22762, 22763, 22764, 22765, 22766, 22767, 22768, 22771, 22772, 22774, 22775, 22779, 22781, 22782, 22783, 22784, 22786, 22787, 22789, 22790, 22794, 22796, 22797, 22798, 22799, 22801, 22802, 22803, 22804, 22816, 22817, 22818, 22819, 22821, 22822, 22823, 22824, 22825, 22826, 22827, 22828, 22834, 22835, 22841, 22842, 22844, 22846, 22847, 22848, 22851, 22852, 22853, 22854, 22857, 22859, 22860, 22871, 22874, 22875, 22876, 22877, 22881, 22883, 22884, 22886, 22889, 22890, 22891, 22892, 22896, 22898, 22899, 22901, 22902, 22903, 22904, 22918, 22919, 22920, 22921, 22922, 22924, 22925, 22926, 22927, 22928, 22929, 22930, 22931, 22932, 22933, 22934, 22935, 22949, 22950, 22971, 22974, 22975, 22976, 22977, 22978, 22983, 22985, 22986, 22989, 22990, 22991, 22992, 22993, 22998, 23000, 23401, 23404, 23405, 23406, 23407, 23408, 23412, 23413, 23414, 23415, 23416, 23417, 23420, 23421, 23423, 23424, 23425, 23426, 23427, 23428, 23429, 23430, 23431, 23432, 23433, 23437, 23438, 23439, 23440, 23441, 23442, 23443, 23444, 23445, 23446, 23447, 23448, 23449, 23470, 23471, 23472, 23473, 23474, 23475, 23478, 23479, 23480, 23490, 23491, 23492, 23493, 23494, 23495, 23501, 23502, 23503, 23504, 23514, 23515, 23516, 23518, 23522, 23523, 23524, 23525, 23526, 23527, 23529, 23530, 23531, 23534, 23535, 23536, 23537, 23541, 23542, 23543, 23544, 23545, 23546, 23547, 23548, 23549, 23550, 23551, 23552, 23556, 23558, 23560, 23562, 23564, 23565, 23566, 23568, 23569, 23570, 23571, 23573, 23575, 23577, 23579, 23580, 23581, 23583, 23584, 23585, 23724, 23725, 23726, 23727, 23728, 23729, 23730, 23731, 23732, 23733, 23734, 23735, 23736, 23737, 23738, 23739, 23740, 23741, 23742, 23743, 23744, 23745, 23746, 23747, 23748, 23749, 23750, 23751, 23752, 23753, 23754, 23755, 23756, 23757, 23758, 23759, 23760, 23761, 23762, 23763, 23764, 23765, 23766, 23767, 23768, 23769, 23770, 23771, 23772, 23773, 23774, 23775, 23776, 23777, 23778, 23779, 23780, 23781, 23782, 23783, 23784, 23785, 23786, 23787, 23788, 23789, 23790, 23791, 23792, 23793, 23794, 23795, 23796, 23797, 23798, 23799, 23800, 23801, 23802, 23803, 23804, 23805, 23806, 24001, 24106, 24107, 24108, 24109, 24132, 24133, 24134, 24144, 24145, 24146, 24147, 24148, 24149, 24151, 24153, 24155, 24160, 24189, 24193, 24203, 24204, 24205, 24206, 24207, 24208, 24209, 24210, 24211, 24212, 24214, 24217, 24218, 24219, 24220, 24221, 24222, 24223, 24224, 24225, 24226, 24227, 24228, 24229, 24230, 24231, 24232, 24233, 24234, 24235, 24236, 24237, 24242, 24243, 24610, 24619, 24620, 24621, 24622, 24626, 24634, 24801, 24823, 24829, 24830, 24831, 24832, 24833, 24834, 24835, 24836, 24837, 24839, 24840, 24841, 24847, 24848, 24849, 24850, 24851, 24852, 24853, 24854, 24855, 24856, 24857, 24858, 24859, 24860, 24861, 24862, 24863, 24864, 24865, 24866, 24867, 24868, 24869, 24870, 24871, 24872, 24873, 24874, 24875, 24876, 24877, 25002, 25620, 25621, 25622, 25623, 25624, 25625, 26027, 26028, 26029, 26030, 26031, 26032, 26033, 26034, 26035, 26036, 26037, 26038, 26039, 26040, 26041, 26042, 26043, 26044, 26045, 26046, 26047, 26048, 26049, 26050, 26051, 26052, 26053, 26054, 26075, 26076, 26077, 26078, 26079, 26080, 26081, 26082, 26083, 26097, 26098, 26099, 26100, 26101, 26102, 26103, 26104, 26105, 26106, 26107, 26108, 26109, 26110, 26111, 26112, 26113, 26117, 26118, 26119, 26120, 26121, 26122, 26123, 28503, 28504, 28505, 28506, 28507, 28508, 28509, 28510, 28511, 28512, 28514, 28515, 28516, 28517, 28518, 28519, 28520, 28521, 28522, 28523, 28626, 28627, 28628, 28629, 28630, 28631, 28632, 28633, 28634, 28637, 28638, 28639, 28640, 28641, 28642, 28643, 28644, 28649, 28650, 28651, 28652, 28653, 28654, 28655, 28656, 28657, 28658, 28659, 28660, 28661, 28662, 28664, 28665, 28666, 28667, 28668, 28669, 28670, 28671, 28674, 28751, 28752, 28753, 28754, 28755, 28757, 28758, 28759, 28760, 28761, 28762, 28765, 28766, 28767, 28771, 28772, 28773, 28774, 28775, 28776, 28777, 28778, 28779, 28780, 28781, 28782, 28783, 28785, 28786, 28787, 28790, 28879, 28880, 28881, 28882, 28883, 28884, 28885, 28886, 28887, 28888, 28889, 28891, 28892, 28893, 28894, 28895, 28896, 28897, 28898, 28899, 28900, 28901, 28906, 29001, 29002, 29003, 29004, 29005, 29006, 29007, 29008, 29009, 29011, 29012, 29013, 29014, 29015, 29016, 29017, 29018, 29019, 29022, 29023, 29024, 29025, 29026, 29027, 29028, 29033, 29034, 29035, 29036, 29037, 29038, 29039, 29040, 29041, 29042, 29043, 29044, 29045, 29046, 29048, 29049, 29050, 29051, 29053, 29126, 29127, 29128, 29129, 29130, 29135, 29136, 29137, 29138, 29139, 29140, 29141, 29142, 29143, 29144, 29145, 29146, 29147, 29152, 29153, 29154, 29155, 29156, 29158, 29165, 29251, 29252, 29254, 29257, 29258, 29259, 29260, 29261, 29263, 29264, 29265, 29266, 29268, 29269, 29270, 29274, 29275, 29276, 29277, 29278, 29279, 29288, 29289, 29290, 29291, 29292, 29293, 29294, 29295, 29296, 29297, 29298, 29300, 29302, 29303, 29304, 29305, 29306, 29376, 29377, 29378, 29379, 29381, 29382, 29383, 29387, 29388, 29389, 29390, 29391, 29392, 29393, 29394, 29395, 29396, 29397, 29399, 29400, 29401, 29402, 29403, 29404, 29405, 29406, 29407, 29408, 29409, 29410, 29412, 29414, 29424, 29501, 29502, 29503, 29504, 29507, 29508, 29509, 29510, 29511, 29512, 29513, 29514, 29515, 29516, 29517, 29518, 29519, 29520, 29521, 29522, 29580, 29581, 29582, 29583, 29584, 29600, 29601, 29602, 29603, 29604, 29605, 29606, 29607, 29608, 29609, 29610, 29611, 29612, 29613, 29615, 29616, 29617, 29618, 29626, 29627, 29628, 29629, 29630, 29631, 29632, 29633, 29634, 29635, 29636, 29638, 29639, 29640, 29641, 29642, 29643, 29644, 29647, 29649, 29650, 29651, 29652, 29653, 29654, 29655, 29656, 29658, 29701, 29702, 29703, 29704, 29705, 29706, 29707, 29708, 29709, 29751, 29752, 29753, 29756, 29757, 29758, 29759, 29760, 29761, 29762, 29763, 29764, 29765, 29766, 29767, 29768, 29769, 29770, 29771, 29772, 29773, 29774, 29775, 29776, 29777, 29825, 29826, 29850, 29851, 29852, 29853, 29854, 29855, 29856, 29857, 29862, 29863, 29864, 29865, 29866, 29867, 29868, 29869, 29870, 29871, 29873, 29876, 29877, 29878, 29880, 29881, 29885, 29886, 29887, 29888, 29889, 29890, 29891, 29892, 29893, 29894, 29895, 29896, 29897, 29899, 29900, 29901, 29902, 29903, 29904, 29914, 29915, 29916, 40218, 40220, 40228, 40258, 40383, 40543, 40545, 40547, 40548, 43786, 44084, 44165, 44195, 44280, 44281, 44286, 44332, 44335, 44336, 44362, 44364, 44948, 44950, 44979, 45201, 45202, 45203, 45204, 45205, 45206, 45207, 45208, 45209, 45210, 45211, 45212, 45213, 45214, 45215, 45216, 45217, 45218, 45219, 45220, 45221, 45222, 45223, 45224, 45227, 45232, 45233, 45234, 45235, 45236, 45237, 45238, 45239, 45240, 45241, 45242, 45243, 45244, 45245, 45246, 45247, 45248, 45249, 45250, 45251, 45252, 45254, 45255, 45256, 45265, 45266, 45267, 45268, 45269, 45270, 45271, 45272, 45273, 45274, 45275, 45276, 45277, 45278, 45279, 45280, 45281, 45282, 45283, 45284, 45285, 45286, 45287, 45288, 45289, 45292, 45297, 45298, 45299, 45300, 45301, 45302, 45332, 45335, 45336, 45337, 45341, 45342, 46006, 46018, 46030, 46042, 46054, 46066, 46078, 46090, 46102, 46114, 46122, 46128, 46140, 46152, 46164, 46176, 46188, 46200, 46212, 46224, 46236, 46244, 46250, 46262, 46274, 46286, 46298, 46310, 46322, 46334, 46346, 46358, 46366, 46372, 46384, 46396, 46408, 46420, 46432, 46444, 46456, 46468, 46480, 46488, 46494, 46506, 46518, 46530, 46542, 46554, 46566, 46578, 46590, 46602, 46610, 46616, 46628, 46640, 46652, 46664, 46676, 46688, 46700, 46712, 46724, 46732, 46738, 46750, 46762, 46774, 46786, 46798, 46810, 46822, 46834, 46846, 46854, 49001, 49003, 49004, 49006, 49009, 49011, 49201, 49202, 49203, 49204, 49205, 49206, 49210, 49211, 49301, 49302, 49303, 49304, 49306, 49307, 49310, 49311, 49401, 49402, 49403, 49404, 49405, 49406, 49410, 49411, 49713, 49714, 49715, 49716, 49722, 49723, 49724, 49725, 49726, 49727, 49728, 49729, 49730, 49731, 49732, 49733, 49804, 49805, 49808, 50006, 50018, 50801, 50802, 50803, 50804, 50806, 50807, 50808, 50809, 50810, 51001, 51002, 51003, 51004, 51005, 51007, 51008, 51013, 52501, 52502, 52503, 52504, 52505, 52506, 52509, 52512, 52515, 52518, 52519, 52520, 52601, 52602, 52603, 52604, 52605, 52606, 52612, 52615, 52618, 52619, 52620, 52701, 52702, 52703, 52704, 52705, 52706, 52709, 52712, 52715, 52718, 52719, 52720, 52801, 52802, 52803, 52804, 52805, 52806, 52809, 52812, 52815, 52818, 52819, 52820, 52901, 52902, 52903, 52904, 52905, 52906, 52912, 52915, 52918, 52919, 52920, 53401, 53402, 53403, 53404, 53405, 53406, 53407, 53408, 53409, 53411, 53501, 53502, 53503, 53504, 53505, 53506, 53507, 53508, 53509, 53511, 53601, 53602, 53603, 53604, 53605, 53606, 53607, 53609, 53701, 53702, 53703, 53704, 53705, 53706, 53707, 53709, 53710, 53801, 53802, 54001, 54003, 54004, 54005, 54012, 54016, 54017, 54018, 54020, 54021, 54022, 54027, 54030, 54032, 54202, 54205, 54206, 54207, 54208, 54209, 54301, 54302, 54303, 54304, 54305, 54306, 54327, 54328, 54329, 54330, 54339, 54340, 54341, 54342, 54343, 54344, 54352, 54353, 54354, 54355, 54356, 54357, 54358, 54359, 54360, 54361, 54362, 54363, 54364, 54365, 54366, 54373, 54374, 54375, 54376, 54377, 54378, 54385, 54386, 54387, 54388, 54389, 54390, 54391, 54398, 54399, 54400, 54401, 54402, 54403, 54404, 54405, 54414, 54415, 54416, 54417, 54418, 54419, 54420, 54421, 54422, 54423, 54424, 54425, 54426, 54427, 54428, 54429, 54430, 54431, 54432, 54433, 54434, 54435, 54436, 54437, 54439, 54449, 54450, 54451, 56001, 56003, 56007, 56009, 56014, 56015, 56016, 56017, 56133, 57001, 57002, 57003, 57004, 57005, 57006, 57036, 57037, 57038, 57039, 57040, 57041, 57042, 57043, 57049, 57050, 57051, 57052, 57053, 57059, 57060, 57061, 57062, 57063, 57064, 57065, 57066, 57079, 57080, 57081, 57082, 57083, 57801, 57802, 57803, 57804, 57805, 57806, 57807, 57809, 57810, 57811, 57812, 57813, 57814, 57815, 57817, 57818, 57819, 57820, 57821, 57822, 57823, 57825, 57826, 57827, 57828, 57829, 57830, 57831, 57833, 57834, 57835, 57836, 57837, 57838, 57839, 57841, 57842, 57843, 57844, 57845, 57846, 57847, 57849, 57850, 57851, 57852, 57853, 57854, 57855, 57857, 57858, 57859, 57860, 57861, 57862, 57863, 57865, 57866, 57867, 57868, 57869, 57870, 57871, 57873, 57874, 57875, 57876, 57877, 57878, 57879, 57881, 57882, 57883, 57884, 57885, 57886, 57887, 57889, 57890, 57891, 57892, 57893, 57894, 57895, 57897, 57898, 57899, 57900, 57901, 57902, 57903, 57905, 57906, 57907, 57908, 57909, 57910, 57911, 57912, 57913, 57914, 57915, 57916, 57917, 57918, 57919, 57920, 57921, 57922, 57923, 57924, 57925, 57926, 59003, 59005, 59103, 59105, 59203, 59205, 59303, 59305, 59306, 59307, 290006, 320018, 331001, 331002, 331006, 331007, 331050, 331058, 336001, 336002, 336003, 336004, 336005, 336011, 336012, 336013, 336014, 336015, 340001, 340002, 340003, 340004, 340006, 340007, 340008, 340009, 340010, 340012, 340014, 340015, 340016, 340018, 340019, 340020, 340021, 340022, 340024, 340025, 340026, 340027, 340030, 340031, 340032, 340033, 340034, 340036, 340037, 340038, 340039, 340040, 340043, 340045, 340047, 340048, 340049, 340050, 340051, 340052, 340054, 340055, 340057, 340059, 340060, 340061, 340062, 340063, 340064, 340067, 340069, 340072, 340073, 340074, 340075, 340076, 340079, 340081, 340083, 340084, 340085, 340086, 340087, 340088, 340091, 340093, 340095, 340096, 340097, 340098, 340099, 340100, 340103, 340107, 340108, 340109, 340110, 340111, 340112, 340115, 340117, 340119, 340120, 340121, 340122, 340123, 340124, 340127, 340129, 340131, 340132, 340133, 340134, 340135, 340136, 340139, 340141, 340143, 340144, 340352, 340353, 340354, 340355, 340356, 340357, 340358, 340359, 340360, 340361, 340362, 340363, 340364, 340365, 340366, 340367, 340368, 340369, 340370, 340372, 340373, 340374, 340375, 340376, 340377, 340378, 340379, 340380, 340381, 340382, 340383, 340384, 340386, 340387, 340388, 340389, 340390, 340391, 340392, 340393, 340394, 340395, 340396, 340397, 340510, 340511, 340512, 340513, 340514, 340515, 340516, 340517, 340518, 340519, 340520, 340521, 340522, 340523, 340524, 340525, 340526, 340527, 340528, 340529, 340530, 340531, 340532, 340533, 340534, 340535, 340536, 340537, 340538, 340539, 340611, 340612, 340613, 340614, 340615, 340616, 340617, 340618, 340619, 340620, 340621, 340622, 340623, 340624, 340625, 340626, 340627, 340628, 340629, 340630, 340631, 340632, 340633, 340634, 340635, 340636, 340637, 340638, 340639, 340640, 340644, 340645, 340647, 400001, 400002, 400003, 400010, 400011, 400012, 400021, 400023, 400024, 400025, 400026, 410011, 410012, 410013, 410015, 410016, 410017, 410018, 420007, 420008, 420009, 420010, 420011, 420012, 420013, 420014, 420027, 420028, 420029, 420031, 420032, 420033, 420034, 420042, 420043, 420044, 420045, 430001, 430002, 430003, 430004, 430005, 430006, 430007, 430008, 430015, 430016, 430017, 430028, 430029, 430030, 430031, 440011, 440012, 440013, 440015, 440016, 440017, 440018, 440026, 440027, 440028, 440029, 440030, 450001, 450002, 450003, 450004, 450005, 450006, 450013, 450014, 450015, 450024, 450025, 450026, 450027, 450028, 450029, 460015, 460016, 460017, 460019, 460020, 460021, 460022, 470001, 470002, 470003, 470004, 470005, 470006, 470007, 470008, 470029, 470030, 470031, 470033, 470034, 470035, 470036, 470044, 470045, 470047, 480001, 480002, 480003, 480004, 480005, 480006, 480013, 480014, 480015, 480024, 480025, 480026, 480027, 480028, 480029, 490011, 490012, 490013, 490015, 490017, 490018, 500001, 500002, 500003, 500004, 500005, 500016, 500017, 500018, 500020, 500022, 500023, 510001, 510002, 510011, 510012, 510013, 510021, 510022, 510023, 510032, 510034, 510035, 510036, 510037, 520001, 520002, 520003, 520004, 520005, 520006, 520007, 520008, 520010, 520012, 520013, 520024, 520025, 520026, 520028, 520029, 520030, 520031, 520036, 520037, 520038, 520039, 520043, 520044, 520045, 530001, 530002, 530003, 530004, 530005, 530006, 530007, 530009, 530010, 530011, 530012, 530013, 530014, 530021, 530022, 530024, 530025, 530026, 530027, 530028, 530029, 530030, 530031, 530032, 530033, 530034, 530035, 530036, 530038, 530039, 530040, 530041, 530042, 530043, 530044, 530045, 530046, 530047, 530048, 530049, 530050, 530051, 530054, 530055, 530056, 530057, 530058, 530065, 530066, 530067, 530076, 530077, 530078, 530079, 530080, 530081, 530082, 540001, 540002, 540003, 540004, 540005, 540006, 540007, 540008, 540010, 540011, 540012, 540013, 540014, 540015, 540016, 540017, 540018, 540019, 540021, 540022, 540023, 540024, 540025, 540026, 540027, 540028, 540029, 540030, 540031, 540032, 540033, 540034, 540035, 540036, 540037, 540038, 540039, 540040, 540041, 540042, 540043, 540054, 540055, 540056, 540058, 540059, 540060, 540061, 540066, 540067, 540068, 540069, 540070, 540071, 540072, 540073, 540074, 540075, 540076, 550501, 550502, 550503, 551001, 551002, 551003, 551004, 551005, 551006, 551007, 551039, 551040, 551041, 551042, 551043, 551044, 551045, 551046, 551053, 551054, 551055, 551056, 551057, 551058, 551059, 551060, 551061, 551063, 551064, 551065, 551066, 551067, 551068, 551069, 551071, 551072, 551080, 551081, 551082, 551083, 551084, 551085, 551086, 551087, 551088, 551089, 551090, 551091, 551098, 551100, 551101, 551103, 551106, 551107, 551108, 551109, 551110, 551111, 551112, 551113, 551114, 551115, 551116, 551117, 551118, 551119, 551120, 551121, 551132, 551133, 551134, 551136, 551137, 551138, 551139, 551144, 551145, 551146, 551147, 551148, 551149, 551150, 551154, 551155, 551156, 551157, 551158, 551159, 551160, 551161, 551162, 551166, 551167, 551168, 551169, 551170, 560001, 560002, 560003, 560004, 560005, 560006, 560007, 560043, 560064, 560065, 560066, 560067, 560068, 560069, 560070, 560071, 560072, 560074, 560075, 560077, 560078, 560079, 560083, 560085, 560086, 560087, 560088, 560089, 560090, 560091, 560092, 560093, 560094, 560095, 560096, 560119, 560120, 580001, 580002, 580003, 580004, 580005, 580006, 580007, 580014, 580015, 580016, 580017, 580018, 580019, 580020, 580021, 580022, 580023, 580024, 580025, 580026, 590271, 590272, 590273, 590277, 590284, 590285, 590286, 590287, 590288, 590289, 590290, 590291, 590292, 590293, 590294, 590295, 590296, 590297, 590298, 590299, 590300, 590301, 590302, 590303, 590304, 590305, 590306, 590307, 590308, 590309, 590310, 590311, 590312, 590313, 590316, 590317, 590318, 590319, 590320, 590322, 590347, 590348, 590349, 590350, 590351, 590352, 590353, 590355, 590356, 590357, 590358, 590359, 590360, 590361, 590362, 590363, 590364, 590365, 590366, 590367, 590368, 590369, 590370, 590371, 590372, 590373, 590374, 590375, 590376, 590377, 590378, 590379, 590380, 590381, 590382, 590383, 590384, 590385, 590386, 590387, 590388, 590389, 590390, 590391, 590392, 590393, 590394, 590395, 590396, 590398, 590399, 590402, 590403, 590404, 590405, 590406, 590407, 590408, 590409, 590410, 590411, 590412, 590413, 590414, 590415, 590416, 590421, 590422, 590423, 590424, 590425, 590426, 590427, 590429, 590430, 590431, 590432, 590433, 590434, 590436, 590437, 590438, 590439, 590440, 590441, 590442, 590443, 590461, 590462, 590463, 590464, 590465, 590466, 590467, 590468, 590469, 590470, 590471, 590472, 590473, 590474, 590475, 590476, 590477, 590485, 590486, 590487, 590488, 590489, 600004, 600005, 600006, 600010, 600023, 600025, 600070, 600071, 600072, 600073, 600074, 600081, 600113, 600114, 600115, 600116, 600117, 701011, 701012, 701013, 701014, 702501, 702502, 702503, 702504, 702505, 702506, 702507, 702508, 702509, 702510, 702511, 702512, 702513, 702514, 702515, 702516, 702517, 702518, 702519, 702520, 702521, 702522, 702523, 702524, 702525, 702526, 702527, 702528, 702529, 702530, 702531, 702532, 702533, 702534, 705015, 705022, 705032, 705037, 705047, 705052, 705074, 705075, 705076, 705081, 705086, 705107, 705108, 705109, 705114, 705119, 705148, 705153, 705154, 705155, 705156, 705501, 705502, 705503, 705504, 705505, 705506, 705507, 705508, 705509, 705510, 705511, 705512, 705513, 705514, 705515, 705516, 705517, 705518, 715001, 715003, 715005, 715007, 715009, 715011, 715013, 715016, 715017, 715019, 715021, 718503, 718505, 718506, 718507, 718509, 718510, 718512, 718513, 718514, 718520, 718521, 718522, 718523, 718586, 718616, 719003, 719004, 719005, 719024, 719025, 719038, 719040, 719525, 719544, 719545, 720501, 720502, 720503, 720504, 720505, 720506, 720508, 720509, 720510, 720511, 720512, 720513, 720514, 720515, 720516, 720517, 720518, 720519, 720520, 720521, 720522, 720523, 720524, 720525, 720526, 720527, 720528, 720529, 720530, 721003, 754008, 754010, 754015, 754017, 754034, 754035, 754056, 754057, 754059, 754060, 754063, 754066, 754072, 754076})

            ' AddLifeBDOItemIDs(ItemIDs)

            ' AddBDOMarketplaceItems(ItemIDs)

            ' UpdateNodes(ItemIDs)

            ' UpdateLifeSkills(ItemIDs)

            ItemIDs = ItemIDs.Distinct().ToList()

            AddStatus("   Updating Market Prices...")
            Dim Index As Integer = 1
            For Each Current As Integer In ItemIDs.OrderBy(Function(o) o)
                AddStatus("      Getting Item: {0} ({1}/{2})...".FormatWith(Current, Index, ItemIDs.Count))
                Index += 1

                Dim URL As String = "https://omegapepega.com/eu/{0}/0".FormatWith(Current)
                If Not AprBase.UrlIsValid(URL) Then
                    Continue For
                End If

                Dim Item As Models.OmegaPepega.Item = RequestManager.GetEntity(Of Models.OmegaPepega.Item)(30, URL)
                If Item Is Nothing Then
                    Continue For
                End If
                If Item.pricePerOne <= 0 Then
                    Continue For
                End If

                Item.name = GetActualResources(Item.name).Trim()
                Items.Add(Item)
            Next
            AddStatus("      Updating MarketPrices Sheet...")
            Dim RowDetails As New List(Of List(Of Object))
            For Each Current As Models.OmegaPepega.Item In Items.OrderBy(Function(o) o.name)
                Dim RowData As New List(Of Object)

                RowData.Add(Current.mainKey)
                RowData.Add(Current.name)
                RowData.Add("")
                RowData.Add(Current.count)
                RowData.Add(Current.totalTradeCount)
                RowData.Add("")
                RowData.Add(Current.pricePerOne)

                RowDetails.Add(RowData)
            Next
            External.GoogleSheets.UpdateRange("MarketPrices_Old!B4", External.GoogleSheets.SpreadSheetID, True, RowDetails)
            RowDetails.Clear()
            RowDetails.Add(New List(Of Object) From {AprBase.CurrentDate.GetSQLString("dd/MM/yy HH:mm")})
            External.GoogleSheets.UpdateRange("MarketPrices_Old!C1", External.GoogleSheets.SpreadSheetID, True, RowDetails)

            AddStatus("BDO Details Updated.")
        Catch ex As Exception
            ex.ToLog()
        End Try
    End Sub

    Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs)
        Try
            Dim SBErrors As New System.Text.StringBuilder()
            Dim ErrorString As String = ""

            SBErrors.Clear()
            For Each kvp As KeyValuePair(Of String, List(Of String)) In RequestManager.JSONErrors
                If kvp.Key.IsNotSet() Then
                    Continue For
                End If
                If kvp.Value.Count <= 0 Then
                    Continue For
                End If

                SBErrors.AppendLine(kvp.Key)
                For Each Current As String In kvp.Value.Distinct().OrderBy(Function(o) o)
                    If Current.IsNotSet() Then
                        Continue For
                    End If

                    SBErrors.Append("   ")
                    SBErrors.AppendLine(Current)
                Next

                If SBErrors.Length > 0 Then
                    SBErrors.AppendLine("")
                End If
            Next
            If SBErrors.Length > 0 Then
                SBErrors.AppendLine("")
            End If
            ErrorString = SBErrors.ToString()
            If ErrorString.IsSet() Then
                AprBase.IORoutines.WriteToFile(True, True, GetErrorLogLocation(), False, ErrorString)
            End If

            SBErrors.Clear()

            SBErrors.Clear()
            For Each kvp As KeyValuePair(Of String, List(Of String)) In AprBase.ProtocolErrors
                If kvp.Key.IsNotSet() Then
                    Continue For
                End If
                If kvp.Value.Count <= 0 Then
                    Continue For
                End If

                SBErrors.AppendLine(kvp.Key)
                For Each Current As String In kvp.Value.Distinct().OrderBy(Function(o) o)
                    If Current.IsNotSet() Then
                        Continue For
                    End If

                    SBErrors.Append("   ")
                    SBErrors.AppendLine(Current)
                Next

                If SBErrors.Length > 0 Then
                    SBErrors.AppendLine("")
                End If
            Next
            If SBErrors.Length > 0 Then
                SBErrors.AppendLine("")
            End If
            ErrorString = SBErrors.ToString()
            If ErrorString.IsSet() Then
                AprBase.IORoutines.WriteToFile(True, True, GetErrorLogLocation(), False, ErrorString)
            End If
        Catch ex As Exception
            ex.ToLog()
        End Try

        Me.Opacity = 1

        Me.Close()
    End Sub

#Region " Shared "

    Private Shared _Instance As frmMain = Nothing

    Friend Shared ReadOnly Property Instance() As frmMain
        Get
            Return _Instance
        End Get
    End Property

#End Region

End Class