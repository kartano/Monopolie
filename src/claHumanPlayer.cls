VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claHumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claHumanPlayer
' Date      : 11/13/2003
' Author    : Simon M. Mitchell
' Purpose   : Human Player class
' 07/02/2004    SM:  Added use of Keep Money Private
' 07/14/2004    SM:  Changed to use new log file
' 05/18/2005    SM:  Added use of SendPlayerToJail
'---------------------------------------------------------------------------------------

Option Explicit

Implements claIPlayer

'' Generic player attributes
Private sName                As String
Private byNumber             As Long
Private lMoney               As Long
Private iLocation            As Integer
Private oPiece               As Token
Private byDoublesCount       As Long
Private bIsBankrupt          As Boolean

'' TO DO:  Replace this with a status value
Private bInJail              As Boolean

Private byJailCount          As Long
Private bOOJailComm          As Boolean
Private bOOJailChance        As Boolean
Private enStatus             As PlayerStatus

Private iReturnRoll          As Integer
Private enSpeed              As PlayerSpeed

' Current square on the board
Private mLocation As Integer

Private mtBar  As cExplorerBar
Private mtItem As cExplorerBarItem

Private Function claIPlayer_AcceptTrade(theTrade As claTrade) As Boolean
    '' TO DO:  If the cash that must change hands exceeds the
    '' amount of cash this player has, the trade should
    '' be automatically rejected.
    
    claIPlayer_AcceptTrade = frmViewTrade.Run(theTrade)
End Function

Private Sub claIPlayer_AddFrame()
    Dim thePlayer As claIPlayer
    
    Set thePlayer = Me
    With mdiMonopoly.SideMenu
        Set mtBar = .Bars.Add(, thePlayer.Name, thePlayer.Name)
        If Globals.MoneyAmountsPrivate Then
            Set mtItem = mtBar.Items.Add(, "Cash" & thePlayer.Number, "Cash: <Private!>", , 1)
            Set mtItem = mtBar.Items.Add(, "TotalValue" & thePlayer.Number, "Total Value: <Private!>", , 1)
        Else
            Set mtItem = mtBar.Items.Add(, "Cash" & thePlayer.Number, "Cash: " & thePlayer.Money, , 1)
            Set mtItem = mtBar.Items.Add(, "TotalValue" & thePlayer.Number, "Total Value: " & thePlayer.TotalValue, , 1)
        End If
        Set mtItem = mtBar.Items.Add(, "TotalProps" & thePlayer.Number, "Total Properties: ", , 1)
        Set mtItem = mtBar.Items.Add(, "Property" & thePlayer.Number, "Property Cards", , 0)
        Set mtItem = mtBar.Items.Add(, "Trade" & thePlayer.Number, "Trade", , 0)
        Set mtItem = mtBar.Items.Add(, "PropsOwned" & thePlayer.Number, "Properties Owned", , 0)
        Set mtItem = mtBar.Items.Add(, "Resign" & thePlayer.Number, "Resign", , 0)
    End With 'MDIMONOPOLY.SIDEMENU
    frmMain.Token(claIPlayer_Token).Visible = True
End Sub

Private Sub claIPlayer_BankruptThisPlayer(theBankruptor As claIPlayer)
    AppendToLog "claHumanPlayer.claIPlayer_BankruptThisPlayer"
    modPlayerUtils.BankruptThePlayer Me, theBankruptor
End Sub

Private Sub claIPlayer_ChangeMoney(theAmount As Long)
    AppendToLog "claHumanPlayer.claIPlayer_ChangeMoney"
    AppendToLog vbTab & "Amount: " & theAmount & " Balance before:" & lMoney & " Balance after will be: " & lMoney + theAmount
    lMoney = lMoney + theAmount
    If Globals.MoneyAmountsPrivate Then
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(1).Text = "Cash: <Private!>"
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(2).Text = "Total Value: <Private!>"
    Else
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(1).Text = "Cash: " & claIPlayer_Money
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(2).Text = "Total Value: " & claIPlayer_TotalValue
    End If
End Sub

'' SM:  Taken from the old claPlayer.inJail method
Private Function claIPlayer_DoSomethingIfInJail() As Boolean
    Dim iMsgboxResult As Integer
    
    If bInJail Then
        If bOOJailChance Or bOOJailComm Then
            'Returns-3,4,5 from buttons
            iMsgboxResult = AlternateMsgbox("You are in Jail, would you like to:", vbAbortRetryIgnore + vbExclamation, "In Jail", "Use Card", "Pay $50", "Roll")
        Else
            'Returns-6,7 from buttons
            iMsgboxResult = AlternateMsgbox("You are in Jail, would you like to:", vbYesNo + vbExclamation, "In Jail", "Pay $50", "Roll")
        End If
        
        Select Case iMsgboxResult
        Case 3 'Use Card
            If bOOJailChance And bOOJailComm Then   'If has both then ask which one
                iMsgboxResult = AlternateMsgbox("Which Out of Jail Card would you like to use?", vbYesNo + vbExclamation, "Which Card?", "Community Chest", "Chance")
                If iMsgboxResult = 6 Then
                    bOOJailComm = False
                End If
                If iMsgboxResult = 7 Then
                    bOOJailChance = False
                End If
            Else 'only has one card so set both to false'NOT OOJAILCHANCE...
                bOOJailChance = False
                bOOJailComm = False
            End If
            bInJail = False
        Case 4 'Pay 50
            claIPlayer_SimpleTransaction Nothing, 50
            bInJail = False
        Case 5 'Roll
            iReturnRoll = frmMain.RollDice(byNumber)
            If byDoublesCount > 0 Then
                bInJail = False
            End If
        Case 6 'Pay 50
            claIPlayer_SimpleTransaction Nothing, -50
            ''SimpleTransaction byNumber, -50, True
            bInJail = False
        Case 7 'Roll
            iReturnRoll = frmMain.RollDice(byNumber)
            If byDoublesCount > 0 Then
                bInJail = False
            End If
        End Select
        
        byDoublesCount = 0
        
        If bInJail Then
            byJailCount = byJailCount + 1
        End If
    End If ' in jail
    claIPlayer_DoSomethingIfInJail = bInJail
End Function

Private Property Let claIPlayer_DoublesCount(RHS As Integer)
    byDoublesCount = RHS
End Property

Private Property Get claIPlayer_DoublesCount() As Integer
    claIPlayer_DoublesCount = byDoublesCount
End Property

Private Property Let claIPlayer_InJail(RHS As Boolean)
    bInJail = RHS
End Property

Private Property Get claIPlayer_InJail() As Boolean
    claIPlayer_InJail = bInJail
End Property

Private Property Let claIPlayer_IsBankrupt(RHS As Boolean)
    bIsBankrupt = RHS
End Property

Private Property Get claIPlayer_IsBankrupt() As Boolean
    claIPlayer_IsBankrupt = bIsBankrupt
End Property

Private Property Let claIPlayer_Location(RHS As Integer)
    mLocation = RHS
End Property

Private Property Get claIPlayer_Location() As Integer
    claIPlayer_Location = mLocation
End Property

Private Property Let claIPlayer_Money(RHS As Long)
    AppendToLog vbTab & "Assignining money for player " & sName & ": " & RHS
    lMoney = RHS
End Property

Private Property Get claIPlayer_Money() As Long
    claIPlayer_Money = lMoney
End Property

Private Property Let claIPlayer_Name(RHS As String)
    sName = RHS
End Property

Private Property Get claIPlayer_Name() As String
    claIPlayer_Name = sName
End Property

Private Property Let claIPlayer_Number(RHS As Integer)
    byNumber = RHS
End Property

Private Property Get claIPlayer_Number() As Integer
    claIPlayer_Number = byNumber
End Property

Private Property Let claIPlayer_OOJailChance(RHS As Boolean)
    bOOJailChance = RHS
End Property

Private Property Get claIPlayer_OOJailChance() As Boolean
    claIPlayer_OOJailChance = bOOJailChance
End Property

Private Property Let claIPlayer_OOJailComm(RHS As Boolean)
    bOOJailComm = RHS
End Property

Private Property Get claIPlayer_OOJailComm() As Boolean
    claIPlayer_OOJailComm = bOOJailComm
End Property

Private Property Let claIPlayer_OOJailRoll(RHS As Integer)
    iReturnRoll = RHS
End Property

Private Property Get claIPlayer_OOJailRoll() As Integer
    claIPlayer_OOJailRoll = iReturnRoll
End Property

Private Property Get claIPlayer_PlayerType() As PlayerType
    claIPlayer_PlayerType = Human
End Property

Private Sub claIPlayer_PropertyLand(Location As Integer, Optional multiplier As Variant = 1)
    Dim iRentTemp       As Long
    Dim prop As claIProperty
    
    mLocation = Location
    Set prop = PropertyManager.property(Location)
    If PropertyManager.isProperty(Location) Then
        If prop.Owner = 0 Then
            frmProperty.Run Location, False
        Else 'property is owned pay rent'NOT IOWNERTEMP...
            frmMain.imgOwnedProp.Picture = frmProperty.P(Location)
            If prop.Owner <> PlayerManager.currentPlayer.Number Then
                If Globals.CollectRentInJail Or (PlayerManager.player(prop.Owner).InJail = False And Globals.CollectRentInJail = False) Then
                    iRentTemp = prop.CalculateRent * multiplier
                    SelfClosingMsgbox sName & " is paying $" & iRentTemp & " rent to: " & PlayerManager.player(prop.Owner).Name, vbOKOnly + vbInformation, "Rent Due", (5 - PlayerManager.currentPlayer.Speed)
                    PlayerManager.currentPlayer.SimpleTransaction PlayerManager.player(prop.Owner), -iRentTemp
                Else
                    SelfClosingMsgbox "No Rent collected, since " & PlayerManager.player(prop.Owner).Name & " is in jail.", vbOKOnly + vbInformation, "Rent Due", (5 - PlayerManager.currentPlayer.Speed)
                End If
            End If
        End If
    ElseIf Location = 30 Then 'NOT PROPERTIES.ISPROPERTY(LOCATION)...'PROPERTIES.ISPROPERTY(LOCATION) = FALSE/0
        modPlayerUtils.SendPlayerToJail Me
    ElseIf Location = 2 Or Location = 17 Or Location = 33 Then 'Community Chest'NOT LOCATION...
        frmDrawCard.CommChestShow
    ElseIf Location = 7 Or Location = 22 Or Location = 36 Then 'Chance'NOT LOCATION...
        frmDrawCard.ChanceShow
    ElseIf Location = 4 Then 'IncomeTax'NOT LOCATION...
        frmIncomeTax.Run
    ElseIf Location = 38 Then 'Luxury Tax'NOT LOCATION...
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -75
    ElseIf Location = 20 Then 'Free Parking'NOT LOCATION...
        If Globals.FreeParkingFees Then
            SelfClosingMsgbox "Free Parking Bonus of $" & Globals.FreeParkingBonusCash & "!", vbOKOnly + vbInformation, "Free Parking Bonus!", 5 - PlayerManager.currentPlayer.Speed
            PlayerManager.currentPlayer.SimpleTransaction Nothing, Globals.FreeParkingBonusCash
            Globals.FreeParkingBonusCash = 0
        End If
    ElseIf Location = 0 Then 'Go'NOT LOCATION...
    
    ElseIf Location = 10 Then 'NOT LOCATION...
        SelfClosingMsgbox "You are just visiting", vbOKOnly + vbInformation, "Just Visiting...", 5 - PlayerManager.currentPlayer.Speed
    End If
End Sub

Private Function claIPlayer_SavePlayer() As String
    Dim sTempString As String

    sTempString = sTempString & "." & sName
    sTempString = sTempString & "." & byNumber
    sTempString = sTempString & "." & lMoney
    sTempString = sTempString & "." & iLocation
    sTempString = sTempString & "." & oPiece
    sTempString = sTempString & "." & byDoublesCount
    sTempString = sTempString & "." & byJailCount
    sTempString = sTempString & "." & bOOJailComm
    sTempString = sTempString & "." & bOOJailChance
    sTempString = sTempString & "." & CInt(Human)
    sTempString = sTempString & "." & CInt(Easy)
    sTempString = sTempString & "." & CInt(Slow)
    sTempString = sTempString & "." & CInt(enStatus) & "."
    sTempString = Convert2Hex(sTempString)
    AppendToLog "claHumanPlayer.claIPlayer_SavePlayer"
    AppendToLog vbTab & "Save string: " & sTempString
    claIPlayer_SavePlayer = sTempString
End Function

Private Property Get claIPlayer_SideBar() As cExplorerBar
    Set claIPlayer_SideBar = mtBar
End Property

Private Sub claIPlayer_SimpleTransaction(theOtherPlayer As claIPlayer, requiredAmount As Long)
    Dim newTrade As claTrade
    Dim myself As claIPlayer
    
    AppendToLog "claIPlayer_SimpleTransaction"
    AppendToLog vbTab & "This player: " & sName
    AppendToLog vbTab & "This player current balance: " & lMoney
    AppendToLog vbTab & "This player can afford it: " & CBool(((lMoney) - Abs(requiredAmount)) >= 0)
    If Not (theOtherPlayer Is Nothing) Then
        AppendToLog vbTab & "Other Player: " & theOtherPlayer.Name
        AppendToLog vbTab & "Other player current balance: " & theOtherPlayer.Money
    Else
        AppendToLog vbTab & "Other player is the Bank (or nobody)"
    End If
    AppendToLog vbTab & "Amount of transaction: " & requiredAmount
    '' SM:  Really silly.
    '' Basically due to how VB6 handles interfaces.
    Set myself = Me
    
    With myself
        If requiredAmount < 0 Then ' we're paying OUT
            Do Until ((lMoney) - Abs(requiredAmount)) >= 0
                Select Case AlternateMsgbox("You do not have enough money, would you like to:", vbAbortRetryIgnore + vbExclamation, "Not Enough Money", "Mortgage", "Trade", "Bankrupt")
                Case 3 'Mortgage
                    frmCards.Run .Number
                Case 4 'Trade
                    If Not (theOtherPlayer Is Nothing) Then
                        Set newTrade = New claTrade
                        newTrade.Load Me, theOtherPlayer
                        If newTrade.Edit Then
                            If theOtherPlayer.AcceptTrade(newTrade) Then
                                newTrade.Execute
                            End If
                        End If
                        Set newTrade = Nothing
                    End If
                Case 5 'Bankrupt
                    claIPlayer_Status = Bankrupt
                    myself.BankruptThisPlayer theOtherPlayer
                    Exit Sub
                End Select
            Loop
        End If
        ' SM:  Req'd amount can be negative
        claIPlayer_ChangeMoney requiredAmount
        ' Transfer/receive cash to or from the other player
        If Not (theOtherPlayer Is Nothing) Then
            theOtherPlayer.SimpleTransaction Nothing, -requiredAmount
        End If
    End With
    
    ShowTransaction requiredAmount
    modSound.PlaySound sndKaching
End Sub

Private Property Let claIPlayer_Status(RHS As PlayerStatus)
    enStatus = RHS
End Property

Private Property Get claIPlayer_Status() As PlayerStatus
    claIPlayer_Status = enStatus
End Property

Private Property Let claIPlayer_Token(RHS As Token)
    oPiece = RHS
End Property

Private Property Get claIPlayer_Token() As Token
    claIPlayer_Token = oPiece
End Property

Private Property Get claIPlayer_TotalValue() As Long
    claIPlayer_TotalValue = modPlayerUtils.CalculateTotalValue(Me)
End Property

Private Sub Class_Initialize()
    lMoney = 1500
    bInJail = False
    bOOJailComm = False
    bOOJailChance = False
    enStatus = Active
End Sub
Private Property Let claIPlayer_JailCount(RHS As Integer)
    byJailCount = RHS
End Property

Private Property Get claIPlayer_JailCount() As Integer
    claIPlayer_JailCount = byJailCount
End Property

Private Sub claIPlayer_AddDoubles()
    byDoublesCount = byDoublesCount + 1
End Sub

Private Property Let claIPlayer_Speed(RHS As PlayerSpeed)
     enSpeed = RHS
End Property

Private Property Get claIPlayer_Speed() As PlayerSpeed
    claIPlayer_Speed = enSpeed
End Property

Private Property Let claIPlayer_Difficulty(RHS As PlayerDifficulty)
    '' SM:  Drop - not really used by the humans.
    ''enDifficulty = RHS
End Property

Private Property Get claIPlayer_Difficulty() As PlayerDifficulty
    ''claIPlayer_Difficulty = enDifficulty
    ''SM:  REturns static value.  Not really used by humans
    claIPlayer_Difficulty = Easy
End Property

