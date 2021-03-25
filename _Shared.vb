Public Module _Shared

    Private _TextInfo As Globalization.TextInfo = New Globalization.CultureInfo("en-GB", False).TextInfo

    Private _IngredientOptionNames As New Dictionary(Of String, String)
    Private _JSONCacheLocation As String = ""
    Private _ResourceNodeNames As New Dictionary(Of String, String)

    Public ReadOnly Property JSONCacheLocation As String
        Get
            If _JSONCacheLocation.IsNotSet() Then
                _JSONCacheLocation = "{0}Cache\BDOData\".FormatWith(AprBase.GetBaseDirectory(AprBase.BaseDirectory.AppDataCompanyProduct))
            End If
            Return _JSONCacheLocation
        End Get
    End Property

    Public ReadOnly Property TextInfo As Globalization.TextInfo
        Get
            Return _TextInfo
        End Get
    End Property

    Private Function GetIconSubFolder(iconName As String) As String
        iconName = iconName.ToLower()

        Dim Folder1 As String = ""
        Dim Folder2 As String = ""

        If iconName.Contains("_") Then
            Folder1 = iconName.Substring(0, iconName.IndexOf("_"))
            iconName = iconName.Substring(iconName.IndexOf("_") + 1)

            If iconName.Contains("_") Then
                Folder2 = iconName.Substring(0, iconName.IndexOf("_"))
                iconName = iconName.Substring(iconName.IndexOf("_") + 1)
            End If
        End If

        If Folder1.IsSet() Then
            Folder1 = "{0}{1}".FormatWith(Char.ToUpper(Folder1(0)), Folder1.Substring(1))

            If Folder2.IsSet() Then
                Folder2 = "{0}{1}".FormatWith(Char.ToUpper(Folder2(0)), Folder2.Substring(1))

                Return "{0}\{1}\".FormatWith(Folder1, Folder2)
            Else
                Return "{0}\".FormatWith(Folder1)
            End If
        End If

        Return ""
    End Function

    Public Function GetActualResources(resources As String) As String
        If resources.IsNotSet() Then
            Return ""
        End If

        Select Case resources
            Case "Magical Barley Seed"
                Return "Barley"
            Case "Magical Corn Seed"
                Return "Corn"
            Case "Magical Freekeh Seed"
                Return "Freekeh"
            Case "Magical Grape Seed"
                Return "Grape"
            Case "Magical Nutmeg Seed"
                Return "Nutmeg"
            Case "Magical Olive Seed"
                Return "Olive"
            Case "Magical Paprika Seed"
                Return "Paprika"
            Case "Magical Seed Potato"
                Return "Potato"
            Case "High-Quality Strawberry Seed, Strawberry Seed"
                Return "Strawberry"
            Case "High-Quality Sweet Potato, Magical Seed Sweet Potato, Special Sweet Potato"
                Return "Sweet Potato"
            Case "Magical Teff Seed"
                Return "Teff"
            Case "Magical Wheat Seed"
                Return "Wheat"

            Case "Brown Coal"
                Return "Coal"
            Case "Brown Coal, Powder of Crevice"
                Return "Coal, Powder of Crevice"
            Case "Brown Coal, Powder of Crevice, Rough Ruby"
                Return "Coal, Powder of Crevice, Rough Ruby"
            Case "Chicken Meat, Four-legged Dining Table"
                Return "Chicken Meat, Eggs"
            Case "Goblin Fighter Mask"
                Return "Fig"
            Case "Roast Mamott"
                Return "Roast Marmot"
            Case "[Spooky Halloween]  Eerie Pumpkin-colored Curtain"
                Return "Pumpkin"
        End Select

        Return resources
    End Function

    Public Function GetIngredientOptionName(ByRef ingredientOptions As List(Of String)) As String
        Dim Key As String = String.Join(", ", ingredientOptions.Distinct().OrderBy(Function(o) o)).ToLower()

        If Not _IngredientOptionNames.ContainsKey(Key) Then
            Dim Name As String = String.Join(", ", ingredientOptions.Distinct().OrderBy(Function(o) o))

            Select Case Key
                Case "apple, banana, cherry, grape, pear, pineapple, strawberry"
                    Name = "Fruit"
                Case "barley, corn, potato, sweet potato, wheat"
                    Name = "Starch"
                Case "barley dough, corn dough, potato dough, sweet potato dough, wheat dough"
                    Name = "Dough"
                Case "barley flour, corn flour, potato flour, sweet potato flour, wheat flour"
                    Name = "Flour"
                Case "bat blood, kuku bird blood, lizard blood, worm blood"
                    Name = "Blood - Legendary Beast"
                Case "bear blood, ogre blood, troll blood"
                    Name = "Blood - Tyrant"
                Case "bear meat, beef, deer meat, fox meat, lamb meat, pork, raccoon meat, rhino meat, weasel meat, wolf meat"
                    Name = "Meat - Red"
                Case "cabbage, olive, paprika, pumpkin, tomato"
                    Name = "Vegetable"
                Case "carp, crab, cuttlefish, jellyfish, octopus, shellfish, squid, starfish, terrapin"
                    Name = "Fish - Salt"
                Case "cheetah dragon blood, flamingo blood, rhino blood, wolf blood"
                    Name = "Blood - Clown"
                Case "cheetah dragon meat, lizard meat, waragon meat, worm meat"
                    Name = "Meat - Lizard"
                Case "deer blood, ox blood, pig blood, sheep blood, waragon blood"
                    Name = "Blood - Sinner"
                Case "dried amberjack, dried angler, dried arowana, dried atka mackerel, dried barbel steed, dried bass, dried beltfish, dried bigeye, dried bitterling, dried blackfin sweeper, dried bleeker, dried blue tang, dried bluefish, dried bluegill, dried bubble eye, dried butterflyfish, dried carp, dried catfish, dried cero, dried checkerboard wrasse, dried cherry salmon, dried clownfish, dried crab, dried crawfish, dried croaker, dried crucian carp, dried cuttlefish, dried cuvier, dried dace, dried dollarfish, dried dolphinfish, dried filefish, dried flatfish, dried flounder, dried flying fish, dried fourfinger threadfin, dried freshwater eel, dried goby minnow, dried golden-thread, dried grayling, dried greenling, dried grouper, dried grunt, dried gunnel, dried gurnard, dried herring, dried horn fish, dried jellyfish, dried john dory, dried kuhlia marginata, dried leather carp, dried lenok, dried mackerel, dried mackerel pike, dried mandarin fish, dried maomao, dried moray, dried mudfish, dried mudskipper, dried mullet, dried nibbler, dried notch jaw, dried octopus, dried pacu, dried perch, dried pintado, dried piranha, dried pomfret, dried porcupine fish, dried porgy, dried ray, dried rock hind, dried rockfish, dried rosefish, dried rosy bitterling, dried round herring, dried roundtail paradisefish, dried salmon, dried sandeel, dried sandfish, dried sardine, dried saurel, dried scorpion fish, dried sea bass, dried sea eel, dried seahorse, dried shellfish, dried siganid, dried silver stripe round herring, dried skipjack, dried smelt, dried smokey chromis, dried snakehead, dried soho bitterling, dried squid, dried starfish, dried striped catfish, dried striped shiner, dried sturgeon, dried sunfish, dried surfperch, dried sweetfish, dried swellfish, dried swiri, dried swordfish, dried tapertail anchovy, dried terrapin, dried tilefish, dried tongue sole, dried trout, dried tuna, dried whiting, dried yellow fin sculpin, dried yellow-head catfish"
                    Name = "Fish - Fresh"
                Case "fox blood, raccoon blood, weasel blood"
                    Name = "Blood - Wise Man"
                Case "rose, sunflower, tulip"
                    Name = "Flower"
                Case Else
                    AprBase.Extensions.System_String.ToLog("frmMain->GetIngredientOptionName, unmatched Key found: {0}{1}{2}".FormatWith(Key, Environment.NewLine, Name), True)
            End Select

            _IngredientOptionNames.Add(Key, Name)
        End If

        Return _IngredientOptionNames(Key)
    End Function

    Public Function GetName(itemDescription As String) As String
        If itemDescription Is Nothing Then
            Return ""
        End If
        If itemDescription.IsNotSet() Then
            Return ""
        End If
        itemDescription = itemDescription.Replace("_"c, " "c)

        Dim Result As String = ""

        For Each Current As String In itemDescription.Split(" "c)
            If Result.IsSet() Then
                Result = "{0} {1}{2}".FormatWith(Result, Current.Substring(0, 1).ToUpper(), Current.Substring(1).ToLower())
            Else
                Result = "{0}{1}".FormatWith(Current.Substring(0, 1).ToUpper(), Current.Substring(1).ToLower())
            End If
        Next

        Return Result
    End Function

    Public Function GetResourceNodeName(resources As String) As String
        If resources.IsNotSet() Then
            Return ""
        End If

        Dim Key As String = resources.ToLower()



        If Not _ResourceNodeNames.ContainsKey(Key) Then
            Dim Name As String = ""

            If Key.Contains("dried ") Then
                Name = "Fish"
            Else
                Select Case Key
                    Case "aloe"
                        Name = "Aloe"
                    Case "bag of muddy water, purified water"
                        Name = "Bag of Muddy Water"
                    Case "bracken"
                        Name = "Bracken"
                    Case "chicken meat, eggs"
                        Name = "Chicken"
                    Case "cinnamon"
                        Name = "Cinnamon"
                    Case "cotton, cotton yarn"
                        Name = "Cotton"
                    Case "date palm"
                        Name = "Date"
                    Case "fig"
                        Name = "Fig"
                    Case "flax, flax thread"
                        Name = "Flax"
                    Case "fleece, knitting yarn"
                        Name = "Fleece"
                    Case "freekeh"
                        Name = "Freekeh"
                    Case "cooking honey"
                        Name = "Honey"
                    Case "nutmeg"
                        Name = "Nutmeg"
                    Case "pistachio"
                        Name = "Pistachio"
                    Case "silk thread, silkworm cocoon"
                        Name = "Silk Thread"
                    Case "star anise mushroom"
                        Name = "Star Anise"
                    Case "teff"
                        Name = "Teff"

                    Case "grape"
                        Name = "Fruit - Grape"
                    Case "strawberry"
                        Name = "Fruit - Strawberry"

                    Case "dry mane grass"
                        Name = "Herb - Dry Mane Grass"
                    Case "silk honey grass"
                        Name = "Herb - Silk Honey Grass"
                    Case "silver azalea"
                        Name = "Herb - Silver Azalea"
                    Case "sunrise herb"
                        Name = "Herb - Sunrise Herb"

                    Case "arrow mushroom", "arrow mushroom, high-quality arrow mushroom, special arrow mushroom"
                        Name = "Mushroom - Arrow"
                    Case "blue umbrella mushroom, pink trumpet mushroom"
                        Name = "Mushroom - Blue"
                    Case "cloud mushroom"
                        Name = "Mushroom - Cloud"
                    Case "dwarf mushroom"
                        Name = "Mushroom - Dwarf"
                    Case "emperor mushroom"
                        Name = "Mushroom - Emperor"
                    Case "fortune teller mushroom"
                        Name = "Mushroom - Fortune Teller"
                    Case "ghost mushroom"
                        Name = "Mushroom - Ghost"
                    Case "green pendulous mushroom, volcanic umbrella mushroom"
                        Name = "Mushroom - Green"
                    Case "sky mushroom"
                        Name = "Mushroom - Sky"
                    Case "tiger mushroom"
                        Name = "Mushroom - Tiger"
                    Case "white flower mushroom, white umbrella mushroom"
                        Name = "Mushroom - White"

                    Case "coal, powder of crevice", "coal, powder of crevice, rough ruby"
                        Name = "Ore - Coal"
                    Case "copper ore", "copper ore, melted copper shard", "copper ore, powder of flame, rough opal", "copper ore, powder of flame", "copper ore, rough translucent crystal"
                        Name = "Ore - Copper"
                    Case "iron ore, powder of darkness, rough black crystal", "iron ore, rough lapis lazuli", "iron ore, rough mud crystal", "iron ore, powder of darkness"
                        Name = "Ore - Iron"
                    Case "lead ore, powder of crevice", "lead ore, powder of time"
                        Name = "Ore - Lead"
                    Case "noc ore, rough lapis lazuli", "noc ore, powder of crevice", "noc ore, powder of earth", "mythril, noc ore", "noc ore, powder of time"
                        Name = "Ore - Noc"
                    Case "powder of crevice, rough diamond, sulfur"
                        Name = "Ore - Sulfur"
                    Case "rough green crystal, tin ore", "rough red crystal, tin ore", "powder of earth, tin ore"
                        Name = "Ore - Tin"
                    Case "powder of flame, rough violet crystal, titanium ore", "powder of flame, rough blue crystal, titanium ore"
                        Name = "Ore - Titanium"
                    Case "powder of darkness, rough opal, vanadium ore", "powder of crevice, rough blue crystal, vanadium ore"
                        Name = "Ore - Vanadium"
                    Case "platinum ore, powder of time, zinc ore", "powder of time, rough red crystal, zinc ore", "mythril, zinc ore"
                        Name = "Ore - Zinc"

                    Case "barley"
                        Name = "Starch - Barley"
                    Case "corn"
                        Name = "Starch - Corn"
                    Case "potato"
                        Name = "Starch - Potato"
                    Case "sweet potato"
                        Name = "Starch - Sweet Potato"
                    Case "wheat"
                        Name = "Starch - Wheat"

                    Case "acacia sap, acacia timber, bloody tree knot"
                        Name = "Timber - Acacia"
                    Case "ash timber, spirit's leaf", "ash sap, ash timber", "ash timber"
                        Name = "Timber - Ash"
                    Case "birch sap, birch timber", "birch plank, birch sap, birch timber", "birch timber, red tree lump"
                        Name = "Timber - Birch"
                    Case "cactus rind, cactus sap, cactus thorn"
                        Name = "Cactus"
                    Case "cedar sap, cedar timber", "cedar timber, spirit's leaf", "cedar timber, monk's branch"
                        Name = "Timber - Cedar"
                    Case "elder tree plank, elder tree sap, elder tree timber"
                        Name = "Timber - Elder Tree"
                    Case "fir plank, fir sap, fir timber", "bloody tree knot, fir timber", "fir sap, fir timber"
                        Name = "Timber - Fir"
                    Case "loopy tree plank, loopy tree sap, loopy tree timber"
                        Name = "Timber - Loopy Tree"
                    Case "maple sap, maple timber", "maple timber, monk's branch", "maple sap, maple timber, old tree bark", "maple timber, red tree lump"
                        Name = "Timber - Maple"
                    Case "moss tree plank, moss tree sap, moss tree timber"
                        Name = "Timber - Moss Tree"
                    Case "pine plank, pine sap, pine timber", "pine sap, pine timber", "monk's branch, pine timber"
                        Name = "Timber - Pine"
                    Case "coconut, palm plank, palm timber", "palm plank, palm sap, palm timber"
                        Name = "Timber - Palm"
                    Case "thuja plank, thuja sap, thuja timber"
                        Name = "Timber - Thuja"
                    Case "bloody tree knot, white cedar sap, white cedar timber", "old tree bark, white cedar sap, white cedar timber"
                        Name = "Timber - White Cedar"

                    Case "trace of ascension, trace of the earth"
                        Name = "Trace - Ascension"
                    Case "trace of battle, trace of forest"
                        Name = "Trace - Battle"
                    Case "trace of chaos, trace of the earth"
                        Name = "Trace - Chaos"
                    Case "trace of death, trace of violence"
                        Name = "Trace - Death"
                    Case "trace of despair, trace of violence", "trace of despair, trace of forest"
                        Name = "Trace - Despair"
                    Case "trace of memory"
                        Name = "Trace - Memory"
                    Case "trace of origin", "trace of hunting, trace of origin"
                        Name = "Trace - Origin"
                    Case "trace of hunting, trace of savagery"
                        Name = "Trace - Savagery"
                    Case "trace of forest, vedelona"
                        Name = "Trace - Vedelona"

                    Case "olive"
                        Name = "Vegetable - Olive"
                    Case "paprika"
                        Name = "Vegetable - Paprika"
                    Case "pumpkin"
                        Name = "Vegetable - Pumpkin"

                    Case Else
                        AprBase.Extensions.System_String.ToLog("_Shared->GetResourceNodeName, unmatched Key found: {0}".FormatWith(Key), True)
                End Select
            End If

            _ResourceNodeNames.Add(Key, Name)
        End If

        Return _ResourceNodeNames(Key)
    End Function

    Public Function GetWoWIcon(iconName As String, size As Integer) As Drawing.Image
        If iconName.IsNotSet() Then
            Return New Drawing.Bitmap(23, 23)
        End If

        iconName = IO.Path.GetFileNameWithoutExtension(iconName)

        Dim FolderName As String = "{0}Icons\{1}".FormatWith(JSONCacheLocation, GetIconSubFolder(iconName))

        If Not IO.Directory.Exists(FolderName) Then
            IO.Directory.CreateDirectory(FolderName)
        End If

        Dim FileName As String = "{0}{1}.JPG".FormatWith(FolderName, iconName)
        Dim FileNamePNG As String = "{0}{1}.PNG".FormatWith(FolderName, iconName)

        If FileName.FileExists() Then
            Return Drawing.Image.FromFile(FileName).GetThumbnail(size)
        End If

        If FileNamePNG.FileExists() Then
            Return Drawing.Image.FromFile(FileNamePNG).GetThumbnail(size)
        End If

        Dim URL As String = "https://render-us.worldofwarcraft.com/icons/56/{0}.jpg".FormatWith(iconName)
        If Not AprBase.UrlIsValid(URL) Then
            URL = "http://eu.media.blizzard.com/wow/icons/56/{0}.jpg".FormatWith(iconName)
            If Not AprBase.UrlIsValid(URL) Then
                URL = "http://wow.zamimg.com/images/wow/icons/large/{0}.jpg".FormatWith(iconName)
                If Not AprBase.UrlIsValid(URL) Then
                    Return Nothing
                End If
            End If
        End If

        Using wc As New System.Net.WebClient()
            wc.DownloadFile(URL, FileName)
        End Using

        Return Drawing.Image.FromFile(FileName).GetThumbnail(size)
    End Function

    Public Sub DeleteJSONFolder(folderName As String)
        Try
            If IO.Directory.Exists(folderName) Then
                IO.Directory.Delete(folderName, True)
            End If
        Catch ex As Exception
            ex.ToLog(True)
        End Try
    End Sub

    Public Sub SetCacheLocations(jsonCache As String)
        _JSONCacheLocation = jsonCache

        If _JSONCacheLocation.IsSet() AndAlso Not _JSONCacheLocation.EndsWith("\") Then
            _JSONCacheLocation.Append("\")
        End If
    End Sub

End Module