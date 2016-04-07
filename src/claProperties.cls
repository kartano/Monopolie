VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claProperties
' Date      : Unknown
' Author    : Various
' Purpose   : Properties manager class
' 14/07/2004    SM:  Added use of new AppendToLog functionality
' 21/08/2004    SM:  Added GroupToString, made StringToGroup public
'---------------------------------------------------------------------------------------

'' SM:  This class used to be called "claProperty".
'' As this actually maintains ALL properties, I've
'' moved it into a claProperties class - less confusing.

'' Major overhaul done for V0.9.8b

Option Explicit

Private mProperties As Collection

Public Sub BuyProperty(ByVal player As Integer, _
                       ByVal Location As Integer)
    With property(Location)
        .Owner = player
        frmMain.Images(Location).ToolTipText = .Name & ": " & PlayerManager.player(player).Name
        AppendToLog "claProperties.BuyProperty"
        AppendToLog vbTab & "Player " & player & " (" & .Name & ") buying " & Location
    End With
End Sub

' Returns the number of railroads owned by the player
Public Function RailroadsOwnedByPlayer(thePlayer As Integer) As Integer
    Dim iCount As Integer
    Dim prop As claIProperty
    
    For Each prop In Properties
        If prop.Group = grpStation And prop.Owner = thePlayer Then
            iCount = iCount + 1
        End If
    Next prop
    RailroadsOwnedByPlayer = iCount
End Function

Private Function CalcUtilityRent(OwnerNum As Integer, Location As Integer) As Integer
    Dim iCount        As Integer
    Dim iRent         As Integer
    Dim DiceRollValue As Integer
    Dim utility As claUtility
    Dim ownersProperty As Collection
    Dim property As claIProperty
    
    Set utility = property(Location)
    Set ownersProperty = PropertiesForPlayer(PlayerManager.player(OwnerNum))
    iCount = 0
    For Each property In Properties
        If TypeOf property Is claUtility Then
            iCount = iCount + 1
        End If
    Next property
    
    DiceRollValue = DiceRoll1 + DiceRoll2
    
    If iCount = 1 Then
        iRent = utility.Multiplier1Utility * DiceRollValue
    Else
        iRent = utility.Multiplier2Utility * DiceRollValue
    End If
    CalcUtilityRent = iRent
End Function

Private Sub Class_Initialize()
    InitToolTips
    InitPlayerCoords
End Sub

Public Function GetLocName(ByVal Location As Integer) As String
    If (Location < 0) Or (Location > 39) Then
        GetLocName = "LOCATION OUT OF RANGE = " & CStr(Location)
    Else 'NOT (LOCATION...
        GetLocName = Properties(CStr(Location)).Name
    End If
End Function

Public Function GetOwner(ByVal iPropertyNumber As Integer) As Long
    GetOwner = property(iPropertyNumber).Owner
End Function

'' SM:  TO DO - Use a collection of all properties in a
'' given group.  Drop the use of the array.
Public Function GetPropertyNumbers(iLocation As Integer) As Integer()
    Dim xTemp(3) As Integer

    If iLocation = 1 Or iLocation = 3 Then                    'Purples
        xTemp(0) = 2
        xTemp(1) = 1
        xTemp(2) = 3
    ElseIf iLocation > 5 And iLocation < 10 Then    'Aquas'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 6
        xTemp(2) = 8
        xTemp(3) = 9
    ElseIf iLocation > 10 And iLocation < 15 Then 'Pinks'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 11
        xTemp(2) = 13
        xTemp(3) = 14
    ElseIf iLocation > 15 And iLocation < 20 Then 'Oranges'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 16
        xTemp(2) = 18
        xTemp(3) = 19
    ElseIf iLocation > 20 And iLocation < 25 Then 'Reds'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 21
        xTemp(2) = 23
        xTemp(3) = 24
    ElseIf iLocation > 25 And iLocation < 30 Then 'Yellows'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 26
        xTemp(2) = 27
        xTemp(3) = 29
    ElseIf iLocation > 30 And iLocation < 35 Then 'Greens'NOT ILOCATION...
        xTemp(0) = 3
        xTemp(1) = 31
        xTemp(2) = 32
        xTemp(3) = 34
    ElseIf iLocation = 37 Or iLocation = 39 Then                  'Blues'NOT ILOCATION...
        xTemp(0) = 2
        xTemp(1) = 37
        xTemp(2) = 39
    End If
    GetPropertyNumbers = xTemp
End Function

Private Sub InitToolTips()
    Dim property As claIProperty

    For Each property In Properties
        frmMain.Images(property.BoardLocation).ToolTipText = property.Name
    Next property
End Sub

Public Function isBuildableProperty(ByVal iPropertyNumber As Integer) As Boolean
    ' Checks whether the passed index is a property on which a player
    ' can buy houses.  Utilities and RailRoads are not "buildable"
    ' @author James
    ' @since version 0.9.5
    
    '' SM:  Changed to use new Properties collection
    If isProperty(iPropertyNumber) Then
        isBuildableProperty = PropertyManager.property(iPropertyNumber).CanImprove
    Else
        isBuildableProperty = False
    End If
End Function

Public Function isProperty(iPropertyNumber As Integer) As Boolean
    ' Checks if the passed value is a property.  Non properties are
    ' the four corners, chance and community chest, income tax
    ' and luxury tax.  AND's and <>'s are used for performace.
    ' @since version 0.9.5
    ' @author James

    '' SM:  Adjusted code to use new collections
    isProperty = IIf(PropertyManager.property(iPropertyNumber) Is Nothing, False, True)

''    Select Case iPropertyNumber
''    Case 0, 10, 20, 30, 2, 4, 7, 17, 22, 33, 36, 38
''        isProperty = False
''    Case Else
''        isProperty = True
''    End Select
End Function

Public Function PlayerNumberHotels(iPlayerNumber As Integer) As Integer
    Dim props As Collection
    Dim iHotelCount As Integer
    Dim prop As claIProperty
    
    Set props = Me.PropertiesForPlayer(PlayerManager.player(iPlayerNumber))
    For Each prop In props
        If prop.Houses = 5 Then iHotelCount = iHotelCount + 1
    Next prop
    PlayerNumberHotels = iHotelCount
End Function

Public Function PlayerNumberHouses(iPlayerNumber As Integer) As Integer
    Dim props As Collection
    Dim iHouseCount As Integer
    Dim prop As claIProperty
    
    Set props = Me.PropertiesForPlayer(PlayerManager.player(iPlayerNumber))
    For Each prop In props
        If prop.Houses < 5 Then iHouseCount = iHouseCount + 1
    Next prop
    PlayerNumberHouses = iHouseCount
End Function

' SM:  A convenience method. Returns all properties in
' a given property group, irrespective of owner or status.
Public Function PropertiesInGroup(theGroup As PropertyGroup) As Collection
    Dim theColl As Collection
    Dim prop As claIProperty
    
    '' SM:  TO DO:  Cache these - properties are obviously
    '' never going to switch between groups, so we should
    '' cache the collections are return them as needed.
    Set theColl = New Collection
    For Each prop In Properties
        If prop.Group = theGroup Then
            theColl.Add prop, CStr(prop.BoardLocation)
        End If
    Next prop
    Set PropertiesInGroup = theColl
End Function

Public Function PlayerOwnsMonopoly_ByPropGroup(player As Integer, propgroup As PropertyGroup, Optional IncludeMortgaged As Boolean = True) As Boolean
    Dim props As Collection
    Dim propertiesOwned As Integer
    Dim prop As claIProperty
    
    Set props = PropertiesInGroup(propgroup)
    
    For Each prop In props
        If prop.Owner = player Then
            If IncludeMortgaged = True Or (Not IncludeMortgaged And prop.MortgageStatus = staUnmortgaged) Then propertiesOwned = propertiesOwned + 1
        End If
    Next prop
    PlayerOwnsMonopoly_ByPropGroup = (propertiesOwned = props.Count)
End Function

Public Function PlayerOwnsMonopoly_ByPropNumber(player As Integer, propertyNum As Integer, Optional IncludeMortgaged As Boolean = True) As Boolean
  ' Checks to see if the player owns the entire monopoly for the given single property.
  ' Returns an integer.  It returns one of the constants true or false
  ' Used for houses and determining rent when the player owns the entire monopoly but
  ' has no houses.
    Dim groupToLookFor As PropertyGroup
    
    groupToLookFor = Me.property(propertyNum).Group
    PlayerOwnsMonopoly_ByPropNumber = PlayerOwnsMonopoly_ByPropGroup(player, groupToLookFor, IncludeMortgaged)
End Function

' SM:  sProperty is a Hex encoded string representing
' all of the property data being restored.
Public Sub RestoreProperty(ByVal sProperty As String)
    Dim iCounter    As Integer
    Dim sTempString As String
    Dim AryReturn   As Variant
    Dim iOffset     As Integer
    Dim Location As Integer
    
    sTempString = Convert2Ascii(sProperty)
    AryReturn = Split(sTempString, ".")
    iOffset = 1
    For iCounter = 1 To Properties.Count
        Location = Val(AryReturn(0))
        With PropertyManager.property(Location)
            .Owner = CInt(AryReturn(1))
            .MortgageStatus = CInt(AryReturn(2))
            .Houses = CInt(AryReturn(3))
        End With
    Next iCounter
End Sub

Public Function SaveProperty() As String
    Dim iCounter    As Integer
    Dim sTempString As String
    Dim prop As claIProperty
    
    For iCounter = 1 To Properties.Count
        Set prop = Properties.item(iCounter)
        sTempString = sTempString & "." & prop.BoardLocation & "." & prop.Owner & "." & prop.MortgageStatus & "." & prop.Houses
    Next iCounter
    sTempString = Convert2Hex(sTempString)
    SaveProperty = sTempString
End Function

' SM:  Will load properties collection on demand
Private Property Get Properties() As Collection
    If mProperties Is Nothing Then
        Set mProperties = New Collection
        LoadProperties
    End If
    Set Properties = mProperties
End Property

' SM:  ONLY gets called from Get Properties.  DO NOT
' call this anywhere else!
Private Sub LoadProperties()
    Dim newProperty As claIProperty
    Dim i As Long
    Dim setting As String
    Dim tokens As Collection
    Dim newUtility As claUtility
    Dim newStation As claStation
    Dim errMsg As String
    
    AppendToLog "claProperties.LoadProperties"
    
    For i = 1 To 22
        setting = INISetting("RealEstate", "LocName" & i, ConfigINI)
        Set tokens = Tokenize(setting, ",")
        If tokens.Count <> 13 Then
            If DebugMode Then
                errMsg = "WARNING!  Invalid number of arguments found for real estate item #" & i & " - please check the " & ConfigINI & " file"
                AppendToLog vbTab & errMsg
                MsgBox errMsg
                Stop
            End If
        Else
            Set newProperty = New claRealEstate
            With newProperty
                .Name = tokens(1)
                .BoardLocation = CInt(tokens(2))
                .PurchasePrice = CLng(tokens(3))
                .MortgagePrice = CLng(tokens(4))
                .BaseRent = CLng(tokens(5))
                .ImprovedRent(imp1House) = CLng(tokens(6))
                .ImprovedRent(imp2House) = CLng(tokens(7))
                .ImprovedRent(imp3House) = CLng(tokens(8))
                .ImprovedRent(imp4House) = CLng(tokens(9))
                .ImprovedRent(imp1Hotel) = CLng(tokens(10))
                .costPerHouse = CLng(tokens(11))
                .CostPerHotel = CLng(tokens(12))
                .Group = StringToGroup(CStr(tokens(13)))
            End With
            AppendToLog vbTab & "Add new property: " & newProperty.Name
            AppendToLog vbTab & vbTab & "Config INI string: " & setting
            mProperties.Add newProperty, CStr(newProperty.BoardLocation)
        End If
    Next i
    
    ''Electric Company,12,150,75,4,10
    For i = 1 To 2
        setting = INISetting("Utilities", "Utility" & i, ConfigINI)
        Set tokens = Tokenize(setting, ",")
        If tokens.Count <> 6 Then
            If DebugMode Then
                errMsg = "WARNING!  Invalid number of arguments found for Utility item #" & i & " - please check the " & ConfigINI & " file"
                AppendToLog errMsg
                MsgBox errMsg
                Stop
            End If
        Else
            Set newProperty = New claUtility
            '' SM:  Doing this because of how VB6 handles
            '' interfaces.
            Set newUtility = newProperty
            With newProperty
                newProperty.Name = tokens(1)
                newProperty.BoardLocation = CInt(tokens(2))
                newProperty.PurchasePrice = CLng(tokens(3))
                newProperty.MortgagePrice = CLng(tokens(4))
            End With
            newUtility.Multiplier1Utility = CInt(tokens(5))
            newUtility.Multiplier2Utility = CInt(tokens(6))
            AppendToLog vbTab & "Add new Utility: " & newProperty.Name
            AppendToLog vbTab & vbTab & "Config INI string: " & setting
            mProperties.Add newProperty, CStr(newProperty.BoardLocation)
        End If
    Next i
    
    For i = 1 To 4
        setting = INISetting("Stations", "Station" & i, ConfigINI)
        Set tokens = Tokenize(setting, ",")
        If tokens.Count <> 8 Then
            If DebugMode Then
                errMsg = "WARNING!  Invalid number of arguments found for station item #" & i & " - please check the " & ConfigINI & " file"
                AppendToLog errMsg
                MsgBox errMsg
                Stop
            End If
        Else
            Set newProperty = New claStation
            '' SM:  Doing this because of how VB6
            '' handles interfaces
            Set newStation = newProperty
            With newProperty
                .Name = tokens(1)
                .BoardLocation = CInt(tokens(2))
                .PurchasePrice = CLng(tokens(3))
                .MortgagePrice = CLng(tokens(4))
            End With
            newStation.RentIf1 = CLng(tokens(5))
            newStation.RentIf2 = CLng(tokens(6))
            newStation.RentIf3 = CLng(tokens(7))
            newStation.RentIf4 = CLng(tokens(8))
            AppendToLog vbTab & "Add new Station: " & newProperty.Name
            AppendToLog vbTab & vbTab & "Config INI string: " & setting
            mProperties.Add newProperty, CStr(newProperty.BoardLocation)
        End If
    Next i
    
    If mProperties.Count <> 28 Then
        errMsg = "WARNING!  Wrong number of properties were loaded!  Expected 28, loaded " & mProperties.Count
        AppendToLog errMsg
        MsgBox errMsg
        End
    End If
    
    Set tokens = Nothing
End Sub

' SM:  Name is confusing!!
' This is the default method for this class
Public Property Get property(Location As Integer) As claIProperty
Attribute property.VB_UserMemId = 0
    On Error GoTo err_property
    Set property = Properties(CStr(Location))
    Exit Property
err_property:
    Set property = Nothing
End Property

' For messages, debugging, etc.
Public Function GroupToString(theGroup As PropertyGroup) As String
    Select Case theGroup
    Case PropertyGroup.grpAqua
        GroupToString = "Aqua"
    Case PropertyGroup.grpBlue
        GroupToString = "Blue"
    Case PropertyGroup.grpGreen
        GroupToString = "Green"
    Case PropertyGroup.grpOrange
        GroupToString = "Orange"
    Case PropertyGroup.grpPink
        GroupToString = "Pink"
    Case PropertyGroup.grpPurple
        GroupToString = "Purple"
    Case PropertyGroup.grpRed
        GroupToString = "Red"
    Case PropertyGroup.grpYellow
        GroupToString = "Yellow"
    Case PropertyGroup.grpStation
        GroupToString = "Station"
    Case Else
        GroupToString = "Utility"
    End Select
End Function

' COnverts a string from an INI file to a group type
Public Function StringToGroup(theGroup As String) As PropertyGroup
    Dim errMsg As String
    
    Select Case UCase$(theGroup)
    Case "PURPLE"
        StringToGroup = grpPurple
    Case "AQUA"
        StringToGroup = grpAqua
    Case "PINK"
        StringToGroup = grpPink
    Case "ORANGE"
        StringToGroup = grpOrange
    Case "RED"
        StringToGroup = grpRed
    Case "YELLOW"
        StringToGroup = grpYellow
    Case "GREEN"
        StringToGroup = grpGreen
    Case "BLUE"
        StringToGroup = grpBlue
    Case Else
        '' Probably a programming error?
        '' If we get here, temporarily stop the code.
        '' Can use immediate window to find problem.
        errMsg = "Invalid Group type read from INI file"
        AppendToLog errMsg
        MsgBox errMsg
        Stop
    End Select
End Function

' Returns a collection of all property objects owned by
' the player.
' Basically just a convenience method.
Public Function PropertiesForPlayer(thePlayer As claIPlayer) As Collection
    Dim props As Collection
    Dim prop As claIProperty
    
    Set props = New Collection
    For Each prop In Properties
        If prop.Owner = thePlayer.Number Then
            props.Add prop, CStr(prop.BoardLocation)
        End If
    Next prop
    Set PropertiesForPlayer = props
End Function

Public Function UnownedProperties() As Collection
    Dim theColl As Collection
    Dim prop As claIProperty
    
    Set theColl = New Collection
    For Each prop In Properties
        If prop.Owner = 0 Then theColl.Add prop, CStr(prop.BoardLocation)
    Next prop
    Set UnownedProperties = theColl
End Function

' This allows us to use "for each" on this class
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = Properties.[_NewEnum]
End Function

Public Property Get Count() As Integer
    Count = Properties.Count
End Property
