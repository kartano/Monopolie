VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCPUPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claCPUPlayer
' Date      : 11/13/2003
' Author    : Simon M. Mitchell
' Purpose   : CPU Player class
' 07/02/2004    SM:  Added use of Keep Money Private
' 07/14/2004    SM:  Added use of new Log File
' 08/21/2004    SM:  Added CheckForPropertyPromotion to PreRoll logic
' 05/18/2005    SM:  Added use of SendPlayerToJail
'---------------------------------------------------------------------------------------

' Poota player class!

Option Explicit

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
Private sNetworkAddress      As String
Private enStatus             As PlayerStatus

Private iReturnRoll          As Integer
Private enSpeed              As PlayerSpeed
Private enDifficulty         As PlayerDifficulty

' Current square on the board
Private mLocation As Integer

Private mtBar  As cExplorerBar
Private mtItem As cExplorerBarItem

Implements claIPlayer

Public Sub PreRollLogic()
    modCPU.CheckForTrade Me
    modCPU.CheckAndDevelopProperty Me
    '' TO DO:  Any decisions a CPU player should make
    '' before its roll will go here
End Sub

Public Sub PostRollLogic()
    modCPU.CheckForTrade Me
    '' TO DO:  Any decisions a CPU player should make
    '' after its roll will go here
End Sub

' claIPlayer interface
Private Function claIPlayer_AcceptTrade(theTrade As claTrade) As Boolean
    Dim returnValue As Boolean
    
    returnValue = modCPU.AcceptTradeOffer(theTrade)
    
    If Not returnValue Then
        mMsgbox.SelfClosingMsgbox claIPlayer_Name & " rejects your trade offer.", vbOKOnly + vbInformation, "Trade Offer Rejected", 4
    Else
        mMsgbox.SelfClosingMsgbox claIPlayer_Name & " accepts your trade offer!", vbOKOnly + vbInformation, "Trade Offer Rejected", 4
    End If
    claIPlayer_AcceptTrade = returnValue
End Function

Private Sub claIPlayer_AddDoubles()
    byDoublesCount = byDoublesCount + 1
End Sub

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
        Set mtItem = mtBar.Items.Add(, "PropsOwned" & thePlayer.Number, "Properties Owned", , 0)
    End With 'MDIMONOPOLY.SIDEMENU
    frmMain.Token(claIPlayer_Token).Visible = True
End Sub

Private Sub claIPlayer_BankruptThisPlayer(theBankruptor As claIPlayer)
    AppendToLog "claCPUPlayer.claIPlayer_BankruptThisPlayer"
    If theBankruptor Is Nothing Then
        AppendToLog vbTab & "Killed by Bank"
    Else
        AppendToLog vbTab & "Killed by " & theBankruptor.Name & " (" & theBankruptor.Number & ")"
    End If
    '' TO DO:  Bankrupt this player
End Sub

Private Sub claIPlayer_ChangeMoney(theAmount As Long)
    AppendToLog "claCPUPlayer.claIPlayer_ChangeMoney"
    AppendToLog vbTab & theAmount
    
    lMoney = lMoney + theAmount
    If Globals.MoneyAmountsPrivate Then
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(1).Text = "Cash: <Private!>"
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(2).Text = "Total Value: <Private!>"
    Else
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(1).Text = "Cash: " & claIPlayer_Money
        mdiMonopoly.SideMenu.Bars(byNumber + 1).Items(2).Text = "Total Value: " & claIPlayer_TotalValue
    End If
End Sub

Private Property Let claIPlayer_Difficulty(RHS As PlayerDifficulty)
    enDifficulty = RHS
End Property

Private Property Get claIPlayer_Difficulty() As PlayerDifficulty
    claIPlayer_Difficulty = enDifficulty
End Property

Private Function claIPlayer_DoSomethingIfInJail() As Boolean
    If bInJail Then
        modCPU.MakeJailDecisions Me
    End If
    claIPlayer_DoSomethingIfInJail = bInJail
End Function

Private Property Let claIPlayer_JailCount(RHS As Integer)
    byJailCount = RHS
End Property

Private Property Get claIPlayer_JailCount() As Integer
    claIPlayer_JailCount = byJailCount
End Property

Private Property Let claIPlayer_Location(RHS As Integer)
    mLocation = RHS
End Property

Private Property Get claIPlayer_Location() As Integer
    claIPlayer_Location = mLocation
End Property

Private Property Let claIPlayer_Money(RHS As Long)
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
    claIPlayer_PlayerType = Computer
End Property

Private Sub claIPlayer_PropertyLand(Location As Integer, Optional multiplier As Variant = 1)
    Dim iRentTemp       As Integer
    Dim prop As claIProperty
    
    mLocation = Location
    If PropertyManager.isProperty(Location) Then
        Set prop = PropertyManager.property(Location)
        If prop.Owner = 0 Then
            modCPU.PurchaseDecision Me, Location, prop
        Else 'property is owned pay rent'NOT IOWNERTEMP...
            frmMain.imgOwnedProp.Picture = frmProperty.P(Location)
            If prop.Owner <> PlayerManager.currentPlayer.Number Then
                If Globals.CollectRentInJail Or (PlayerManager.player(prop.Owner).InJail = False And Globals.CollectRentInJail = False) Then
                    iRentTemp = prop.CalculateRent * multiplier
                    SelfClosingMsgbox sName & " is paying $" & iRentTemp & " rent to: " & PlayerManager.player(prop.Owner).Name, vbOKOnly + vbInformation, "Rent Due", (5 - PlayerManager.currentPlayer.Speed)
                    claIPlayer_SimpleTransaction PlayerManager.player(prop.Owner), -iRentTemp
                Else
                    SelfClosingMsgbox "No Rent collected, since " & PlayerManager.player(prop.Owner).Name & " is in jail.", vbOKOnly + vbInformation, "Rent Due", (5 - PlayerManager.currentPlayer.Speed)
                End If
            End If
        End If
    ElseIf Location = 30 Then 'NOT PROPERTIES.ISPROPERTY(LOCATION)...'PROPERTIES.ISPROPERTY(LOCATION) = FALSE/0
        modPlayerUtils.SendPlayerToJail Me
    ElseIf Location = 2 Or Location = 17 Or Location = 33 Then 'Community Chest'NOT LOCATION...
        '' SM:  At the moment, draw card handles the CPU logic.
        '' This works for now, but should be changed later.
        '' frmDrawCard shouldn't need to know what kind of player
        '' it is handling.
        frmDrawCard.CommChestShow
    ElseIf Location = 7 Or Location = 22 Or Location = 36 Then 'Chance'NOT LOCATION...
        '' SM:  Ditto for the above.
        
        frmDrawCard.ChanceShow
    ElseIf Location = 4 Then 'IncomeTax'NOT LOCATION...
        modCPU.HandleIncomeTax Me
    ElseIf Location = 38 Then 'Luxury Tax'NOT LOCATION...
        claIPlayer_SimpleTransaction Nothing, -75
    ElseIf Location = 20 Then 'Free Parking'NOT LOCATION...
        If Globals.FreeParkingFees Then
            SelfClosingMsgbox "Free Parking Bonus of $" & Globals.FreeParkingBonusCash & "!", vbOKOnly + vbInformation, "Free Parking Bonus!", 5 - PlayerManager.currentPlayer.Speed
            PlayerManager.currentPlayer.SimpleTransaction Nothing, Globals.FreeParkingBonusCash
            Globals.FreeParkingBonusCash = 0
        End If
    ElseIf Location = 0 Then 'Go'NOT LOCATION...
        ''
    ElseIf Location = 10 Then 'NOT LOCATION...
        SelfClosingMsgbox sName & " is just visiting!", vbOKOnly + vbInformation, "Just Visiting...", 5 - PlayerManager.currentPlayer.Speed
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
    sTempString = sTempString & "." & CInt(Computer)
    sTempString = sTempString & "." & CInt(enDifficulty)
    sTempString = sTempString & "." & CInt(enSpeed)
    sTempString = sTempString & "." & CInt(enStatus) & "."
    sTempString = Convert2Hex(sTempString)
    AppendToLog "claCPUPlayer.claIPlayer_SavePlayer"
    AppendToLog vbTab & "Save string: " & sTempString
    claIPlayer_SavePlayer = sTempString
End Function

Private Property Get claIPlayer_SideBar() As cExplorerBar
    Set claIPlayer_SideBar = mtBar
End Property

Private Property Let claIPlayer_Speed(RHS As PlayerSpeed)
    enSpeed = RHS
End Property

Private Property Get claIPlayer_Speed() As PlayerSpeed
    claIPlayer_Speed = enSpeed
End Property

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

Private Sub claIPlayer_SimpleTransaction(theOtherPlayer As claIPlayer, requiredAmount As Long)
    With PlayerManager
        ' SM:  Check for bankrupt CPU player
        If requiredAmount < 0 Then
            If Abs(requiredAmount) > CalculateTotalValue(Me) Then
                BankruptThePlayer Me, theOtherPlayer
                Exit Sub
            End If
        End If
        
        Do Until ((lMoney) - Abs(requiredAmount)) >= 0
            modCPU.NotEnoughCash Me, theOtherPlayer
        Loop
        
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

