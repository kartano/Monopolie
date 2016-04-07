Attribute VB_Name = "modCPU"
'---------------------------------------------------------------------------------------
' Module    : modCPU
' Date      : 11/13/2003
' Author    : Simon M. Mitchell
' Purpose   : Generic CPU decision making
' 21/08/2004    SM:  Added CheckAndDevelopProperty
' 05/30/2004    SM:  CPU players now mortgage unimproved properties if short of cash
'---------------------------------------------------------------------------------------

'' SM:  This code could be moved across to the claCPUPlayer
'' class eventually, although with multiple CPU players
'' I'm guessing having it seperately in a module will save
'' some memory?

' Checks CPUs desire to offer a trade
Public Sub CheckForTrade(cpuPlayer As claIPlayer)
    Dim player As claIPlayer
    Dim playerInJail As Boolean
    Dim trade As claTrade
    Dim TradeDone As Boolean
    
    ' DO I have a Get out of jail Free card and I'm not in jail?
    ' If so, is any player in jail?
    ' If so, are they on their second roll?
    ' If so, offer to sell it to them.
    ' The harder I am, the more cash I ask for.
    If (cpuPlayer.OOJailChance Or cpuPlayer.OOJailComm) And cpuPlayer.InJail = False Then
        TradeDone = False
        AppendToLog "modCPU.CheckForTrade"
        AppendToLog vbTab & "Checking for offer of GOOJF cards"
        For Each player In PlayerManager
            If player.number <> cpuPlayer.number Then
                If player.InJail And player.JailCount = 2 Then
                    AppendToLog vbTab & "Player " & player.Name & " in jail, building trade offer"
                    trade.Load cpuPlayer, player
                    Set trade = New claTrade
                    Select Case cpuPlayer.Difficulty
                    Case PlayerDifficulty.Easy
                        trade.Cash = CLng(player.Money * 0.1)
                    Case PlayerDifficulty.Medium
                        trade.Cash = CLng(player.Money * 0.2)
                    Case PlayerDifficulty.Hard
                        trade.Cash = CLng(player.Money * 0.3)
                    End Select
                    trade.ReceivingPlayerGetsCash = True
                    If cpuPlayer.OOJailChance Then
                        trade.ChanceJailCardOffered = True
                    Else
                        trade.CommChestJailCardOffered = True
                    End If
                    trade.DumpToLog
                    If player.AcceptTrade(trade) Then
                        TradeDone = True
                        trade.Execute
                        Exit For
                    End If
                End If
            End If
        Next player
    End If
    
    '' TO DO:  Check my cards - hunt for situations where:
    '' - I have almost all cards in a set, and someone else
    ''   has the card I need.
    '' - I have a card someone else wants and I can get a motza from
    ''   them to do it!
    
    '' SM:  In future, could also build a list of all
    '' trades I've offered and use those to make
    '' intelligent decisions based on any future/new trades.
End Sub

' Gets CPU player to decide how to handle a cash shortage
' NOTE:  If we owe the bank, theOtherPlayer will be NOTHING
Public Sub NotEnoughCash(cpuPlayer As claIPlayer, theOtherPlayer As claIPlayer)
    Dim properiesIOwn As Collection
    Dim prop As claIProperty
    Dim keepLooking As Boolean
    
    AppendToLog "modCPU.NotEnoughCash"
    AppendToLog vbTab & "Player with problem: " & cpuPlayer.Name
    
    Set propertiesIOwn = PropertyManager.PropertiesForPlayer(cpuPlayer)
    
    keepLooking = True
    ' STEP ONE:  Hunt for unmortgaged, unimproved props.
    AppendToLog vbTab & vbTab & "Looking for unmortgaged, unimproved props"
    For Each prop In propertiesIOwn
        If prop.MortgageStatus = staUnmortgaged Then
            If prop.CanImprove = False Then
                AppendToLog vbTab & vbTab & "Mortgaging: " & prop.Name
                '' TO DO:  SM:  Find a way to animate frmProperty to display a "mortgaged"
                '' message.  Just makes the interface consistent.
                SelfClosingMsgbox "CPU Player " & cpuPlayer.Name & " Mortgages " & prop.Name & " for " & prop.MortgagePrice
                prop.MortgageStatus = staMortgaged
                cpuPlayer.Money = cpuPlayer.Money + prop.MortgagePrice
                keepLooking = False
                Exit For
            End If
        End If
    Next prop
    
    If keepLooking Then
        AppendToLog vbTab & vbTab & "Looking for houses/hotels to remove"
        For Each prop In propertiesIOwn
            If prop.MortgageStatus = staUnmortgaged Then
                If prop.Houses > 0 Then
                    AppendToLog vbTab & vbTab & "Removing houses from " & prop.Name
                    '' TO DO:  Remove houses evenly
                    '' off this property group.
                    '' One by one.
                    keepLooking = False
                End If
            End If
        Next prop
    End If
    
    If keepLooking Then
        AppendToLog vbTab & vbTab & "Looking to offer a trade with someone"
        '' TO DO:  Offer a trade to another player?
    End If
    
    If keepLooking Then
        AppendToLog vbTab & vbTab & "FATALITY!  CPU module unable to reach a decision!"
        '' SM:  If we get here, something went wrong with
        '' the decision making process.
        MsgBox "CPU player has funds, but was unable to reach a decision when short of cash"
        Stop
    End If
End Sub

Public Function AcceptTradeOffer(theTrade As claTrade) As Boolean
    Dim CPUName As String
    
    AppendToLog "modCPU.AcceptTradeOffer"
    AppendToLog "Dump of trade on offer"
    theTrade.DumpToLog
    
    CPUName = theTrade.PlayerReceivingOffer.Name
    ' Sanity checks first.
    ' Human players love to pull these stunts
    With theTrade
        ' Player wants some of my properties for free
        If .Cash = 0 And .PropertiesOffered.Count = 0 And .PropertiesWanted.Count > 0 Then
            AppendToLog vbTab & "REJECTING TRADE OFFER - trade was ridiculous (wanted properties for nothing)"
            mMsgbox.SelfClosingMsgbox CPUName & " says: I'd be crazy to accept that kind of offer! No thanks!", vbOKOnly + vbInformation, "CPU Rejects the offer"
            AcceptTradeOffer = False
            Exit Function
        ' Player is offering the CPU properties for free
        ElseIf .Cash = 0 And .PropertiesWanted.Count > 0 And .PropertiesOffered.Count = 0 Then
            AppendToLog vbTab & "ACCEPTING TRADE OFFER - trade was ridiculous (offered properties for nothing)"
            mMsgbox.SelfClosingMsgbox CPUName & " says: Pffff!  Okay, I'll accept that!", vbOKOnly + vbInformation, "CPU Accepts offer"
            AcceptTradeOffer = True
            Exit Function
        ' Player if offering the CPU cash for nothing
        ElseIf .Cash >= 0 And .PropertiesOffered.Count = 0 And .PropertiesWanted.Count = 0 And .ReceivingPlayerGetsCash = True Then
            AppendToLog vbTab & "ACCEPTING TRADE OFFER - trade was ridiculous (offered cash for nothing)"
            mMsgbox.SelfClosingMsgbox CPUName & " says: Pffff!  Okay, I'll accept that!", vbOKOnly + vbInformation, "CPU Accepts offer"
            AcceptTradeOffer = True
            Exit Function
        ' Player is asking me for cash in exchange for nothing
        ElseIf .Cash >= 0 And .PropertiesOffered.Count = 0 And .PropertiesWanted.Count = 0 And .ReceivingPlayerGetsCash = False Then
            AppendToLog vbTab & "REJECTING TRADE OFFER - trade was ridiculous (wanted cash for nothing)"
            mMsgbox.SelfClosingMsgbox CPUName & " says: I'd be crazy to accept that kind of offer! No thanks!", vbOKOnly + vbInformation, "CPU Rejects the offer"
            AcceptTradeOffer = False
            Exit Function
        End If
    End With
    
    '' TO DO:  Logic that determines whether to accept
    '' an offer or not.
    '' For now - accept all offers if I'm easy.
    '' Otherwise, reject.
    If theTrade.PlayerMakingOffer.Difficulty = Easy Then
        AppendToLog vbTab & "ACCEPTING TRADE OFFER"
        mMsgbox.SelfClosingMsgbox CPUName & " says: I accept your trade offer!", vbOKOnly + vbInformation, "CPU accepts trade offer"
        '' TO DO:  Play happy, satisfied, chirpy sound!
        AcceptTradeOffer = True
    Else
        AppendToLog vbTab & "REJECTING TRADE OFFER"
        mMsgbox.SelfClosingMsgbox CPUName & " says: I reject your trade offer!", vbOKOnly + vbInformation, "CPU rejects trade offer"
        '' TO DO:  Play bad sound!
        AcceptTradeOffer = False
    End If
End Function

Public Sub MakeJailDecisions(thePlayer As claIPlayer)
    Dim UnownedProperties As Integer
    Dim totalProperties As Integer
    Dim unownedPerc As Integer
    Dim minPerc As Integer
    
    AppendToLog "modCPU.MakeJailDecisions"
    AppendToLog vbTab & "Player with problem: " & thePlayer.Name
    
    UnownedProperties = PropertyManager.UnownedProperties.Count
    totalProperties = PropertyManager.Count
    unownedPerc = Int((UnownedProperties / totalProperties) * 100)
    
    AppendToLog vbTab & "UnownedProperties = " & UnownedProperties
    AppendToLog vbTab & "totalProperties = " & totalProperties
    AppendToLog vbTab & "unownedPerc = " & unownedPerc
    
    ' How many unown properties left?
    ' If many properties still unowned, try to leave
    ' immediately.
    Select Case thePlayer.Difficulty
    Case PlayerDifficulty.Easy
        minPerc = 50
    Case PlayerDifficulty.Medium
        minPerc = 40
    Case Else
        minPerc = 30
    End Select
    
    AppendToLog vbTab & "minPerc = " & minPerc
    
    If unownedPerc >= minPerc Then
        AppendToLog vbTab & "Attempting to get out of jail IMMEDIATELY"
        '' Attempt to get out of jail immediately
    Else
        AppendToLog vbTab & "Waiting to get out of jail"
        '' - Take another roll attempt if I have some to
        ''   spare.
        '' - If this is my third roll attempt, use a
        ''   GOOJF if I have one.
        '' - Finally, pay $50.
    End If
End Sub

Public Sub PurchaseDecision(thePlayer As claIPlayer, theLocation As Integer, theProperty As claIProperty)
    '' TO DO:  Decide what to do.
    '' FOr now, the poota buys the property if it has the
    '' cash on hand to.
    If thePlayer.Money >= theProperty.PurchasePrice Then
        AppendToLog "modCPU.PurchaseDecision"
        AppendToLog vbTab & thePlayer.Name & " buying property"
        theProperty.DumpToLog
        thePlayer.SimpleTransaction Nothing, -theProperty.PurchasePrice
        PropertyManager.BuyProperty thePlayer.number, theLocation
        frmProperty.Run theLocation, True
    Else
        AppendToLog "modCPU.PurchaseDecision"
        AppendToLog vbTab & player.Name & " declines to buy property"
        theProperty.DumpToLog
        SelfClosingMsgbox thePlayer.Name & " declines to purchase this property"
    End If
End Sub

Public Sub HandleIncomeTax(thePlayer As claIPlayer)
    AppendToLog "modCPU.HandleIncomeTax"
    AppendToLog vbTab & "Player with problem: " & thePlayer.Name
    AppendToLog tbTab & "Total player value: " & thePlayer.TotalValue
    
    If (0.1 * thePlayer.TotalValue) < 200 Then
        AppendToLog vbTab & "Paying 10%"
        SelfClosingMsgbox thePlayerName & " opts to pay 10%"
        thePlayer.SimpleTransaction Nothing, (0.1 * thePlayer.TotalValue)
    Else
        AppendToLog vbTab & "Paying $200"
        SelfClosingMsgbox thePlayerName & " opts to pay $200"
        thePlayer.SimpleTransaction Nothing, 200
    End If
End Sub

' Checks to see if the CPU player owns a set
' If so, attempts to build on it.
Public Sub CheckAndDevelopProperty(thePlayer As claIPlayer)
    AppendToLog "modCPU.CheckAndDevelopProperty"
    AppendToLog vbTab & "Player: " & thePlayer.Name
    
    ' Sanity check
    If PropertyManager.PropertiesForPlayer(thePlayer).Count <= 1 Then
        AppendToLog vbTab & "Exiting - not enough properties"
        Exit Sub
    End If
    
    '' SM:  The order in which the tests for each
    '' group are called affect which set the poota
    '' will try to develop.  In future, this order could
    '' be optimised - for example, to upgrade properties
    '' over which other players will soon move
    DevelopSingleGroup thePlayer, grpAqua
    DevelopSingleGroup thePlayer, grpBlue
    DevelopSingleGroup thePlayer, grpGreen
    DevelopSingleGroup thePlayer, grpOrange
    DevelopSingleGroup thePlayer, grpPink
    DevelopSingleGroup thePlayer, grpPurple
    DevelopSingleGroup thePlayer, grpRed
    DevelopSingleGroup thePlayer, grpYellow
End Sub

Private Sub DevelopSingleGroup(player As claIPlayer, theGroup As PropertyGroup)
    Dim props As Collection
    Dim prop As claIProperty
    Dim CostPerHouse As Long
    Dim propToDevelop As claIProperty
    Dim lowestHouseCount As Integer
    Dim finished As Boolean
    Dim housesPurchased As Integer
    Dim found As Boolean
    
    AppendToLog "modCPU.DevelopSingleGroup"
    AppendToLog vbTab & "Player: " & player.Name
    AppendToLog vbTab & "Group to check: " & PropertyManager.GroupToString(theGroup)
    If Not CanPromote(player, theGroup) Then
        AppendToLog vbTab & vbTab & "Exiting - player can't promote this group"
        Exit Sub
    End If
    AppendToLog vbTab & "Group to promote: " & PropertyManager.GroupToString(theGroup)
    
    Set props = PropertyManager.PropertiesInGroup(theGroup)
    
    ' Sanity check - afford at least one house
    CostPerHouse = props.item(1).CostPerHouse
    AppendToLog vbTab & "Cost per house: " & CostPerHouse
        
    If CostPerHouse < player.Money Then
        AppendToLog vbTab & "CPU player trying to promote"
        ' Locate property to promote
        finished = False
        While Not finished
            AppendToLog vbTab & vbTab & "Looking for undeveloped property in set"
            ' Locate property with lowest development
            lowestHouseCount = -1
            found = False
            For Each prop In props
                If prop.Houses < 5 And lowestHouseCount <= prop.Houses Then
                    lowestHouseCount = prop.Houses
                    Set propToDevelop = prop
                    found = True
                End If
            Next prop
            ' Keep developing until set is fully promoted or
            ' we've run out of cash
            If propToDevelop Is Nothing Or found = False Then
                AppendToLog vbTab & vbTab & vbTab & "None found"
                finished = True
            Else
                AppendToLog vbTab & vbTab & vbTab & "Found:"
                AppendToLog vbTab & vbTab & vbTab & "============================"
                propToDevelop.DumpToLog
                AppendToLog vbTab & vbTab & vbTab & "============================"
                propToDevelop.Houses = propToDevelop.Houses + 1
                player.SimpleTransaction Nothing, CostPerHouse
                housesPurchased = housesPurchased + 1
                ' We're done when we can't afford any houses
                finished = CostPerHouse > player.Money
            End If
        Wend
        Select Case housesPurchased
        Case 0
            ' The CPU player should only ever enter this
            ' block of code where there is at least SOMETHING
            ' to build on.
            AppendToLog vbTab & "BIZARRE Problem!  CPU player couldn't find a property to develop!"
        Case 1
            SelfClosingMsgbox "CPU player " & player.Name & " purchases 1 house on the " & PropertyManager.GroupToString(theGroup) & " set.", vbOKOnly + vbInformation, "CPU Player Develops"
        Case Else
            SelfClosingMsgbox "CPU player " & player.Name & " purchases " & housesPurchased & " houses on the " & PropertyManager.GroupToString(theGroup) & " set.", vbOKOnly + vbInformation, "CPU Player Develops"
        End Select
    Else
        AppendToLog vbTab & "Exiting - can't afford to buy this house"
    End If
End Sub

Private Function CanPromote(player As claIPlayer, theGroup As PropertyGroup) As Boolean
    Dim props As Collection
    Dim prop As claIProperty
    Dim undevelopedCount As Integer
    
    ' Sanity check:  If no monopoly, jump ship
    If Not PropertyManager.PlayerOwnsMonopoly_ByPropGroup(player.number, theGroup, False) Then
        CanPromote = False
    Else
        Set props = PropertyManager.PropertiesInGroup(theGroup)
        For Each prop In props
            If prop.Houses < 5 Then undevelopedCount = undevelopedCount + 1
        Next prop
        CanPromote = (undevelopedCount <= props.Count)
    End If
End Function
