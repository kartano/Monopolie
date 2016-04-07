VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrawCard 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2685
   ClientLeft      =   4575
   ClientTop       =   3240
   ClientWidth     =   4350
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSComctlLib.ImageList imgChance 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   275
      ImageHeight     =   140
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1C522
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":38A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":54F66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":71B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":8E69A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":AABBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":C70DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":E3600
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":FFB22
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":11C044
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":138566
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":154A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":170FAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4095
   End
   Begin MSComctlLib.ImageList imgCommChest 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   275
      ImageHeight     =   140
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":18D4CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1A99EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1C5F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1E2432
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1FE954
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":21AE76
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":237398
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":2538BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":26FDDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":28C2FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":2A8820
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":2C4D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":2E1264
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":2FD786
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":319CA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgCard 
      Height          =   2100
      Left            =   120
      Top             =   120
      Width           =   4125
   End
End
Attribute VB_Name = "frmDrawCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDrawCard
' Date      : 4/14/2003
' Author    :
' Purpose   : This Form replaces two forms from previous versions
'                   It handles the loading of Chance/Comm Chest cards
'                   and takes the appropriate action for the selected card
' 02/07/2004    SM:  Added use of free parking fees
' 18/05/2005    SM:  Added use of SendPlayerToJail
' 03/03/2006    SM:  Removed some cruft, fixed indents
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub AdvanceTo(LocationNumber As Integer)
    Dim iCurrLoc As Integer

    iCurrLoc = PlayerManager.currentPlayer.Location
    If iCurrLoc > LocationNumber Then
        MovePlayerToken iCurrLoc, ((40 - iCurrLoc) + LocationNumber), 1
    Else
        MovePlayerToken iCurrLoc, (LocationNumber - iCurrLoc), 1
    End If
    PlayerManager.currentPlayer.Location = LocationNumber
End Sub

Private Sub Chance(CardNumber As Integer)
    Dim iReturn    As Integer
    Dim iTemptotal As Long
    Dim iTempLoc   As Integer

    Select Case CardNumber
    Case 1 'Building Loan Matures (+100)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 100
    Case 2 'Advance Nearest Rail (2xRental)
        AdvanceToRailRoad PlayerManager.currentPlayer.Location, 2
    Case 3 'Go to Jail
        modPlayerUtils.SendPlayerToJail PlayerManager.currentPlayer
    Case 4 'Take Ride of Reading
        AdvanceTo 5
    Case 5 'Advance to BoardWalk
        AdvanceTo 39
    Case 6 'Nearest Util (10x Dice Roll)
        iTempLoc = PlayerManager.currentPlayer.Location
        If iTempLoc = 7 Or iTempLoc = 36 Then 'The only two Chance Locations that can send you to space 12
            MovePlayerToken iTempLoc, (12 - iTempLoc), 1, 10
        Else 'NOT ITEMPLOC...
            MovePlayerToken iTempLoc, (28 - iTempLoc), 1, 10
        End If
    Case 7 'Advance to Illinois
        AdvanceTo 24
    Case 8 'Pay Poor Tax (-15)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -15
        Globals.CheckAddFreeParkingCash 15
    Case 9 'OOJail
        PlayerManager.currentPlayer.OOJailChance = True
    Case 10 'Advance to Go
        AdvanceTo 40
    Case 11 'Bank Pays Dividend (+50)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 50
    Case 12 'Go Back 3 Spaces
        iReturn = MovePlayerToken(PlayerManager.currentPlayer.Location, 3, 2)
    Case 13 'Advance to St Charles
        AdvanceTo 11
    Case 14 'Repairs (-25 per house, -100 per Hotel)
        iTemptotal = 25 * PropertyManager.PlayerNumberHouses(PlayerManager.currentPlayer.Number)
        iTemptotal = iTemptotal + (100 * PropertyManager.PlayerNumberHotels(PlayerManager.currentPlayer.Number))
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -iTemptotal
        Globals.CheckAddFreeParkingCash iTemptotal
    End Select
End Sub

Public Sub ChanceShow()
    Dim SelectedCard As Integer

    Me.Caption = "Chance"
    SelectedCard = ChanceCards.GetCardNumber
    If ChanceCards.PlayerHasOOJailCard And SelectedCard = 9 Then
        SelectedCard = ChanceCards.GetCardNumber
    ElseIf SelectedCard = 15 Then 'There are only 14 cards for Chance so select a different one'NOT CHANCECARDS.PLAYERHASOOJAILCARD...
        SelectedCard = ChanceCards.GetCardNumber
    End If
    Me.imgCard.Picture = imgChance.ListImages(SelectedCard).Picture
    '' SM:  DANGEROUS.  All form events can get triggered here!
    DoEvents
  
    If PlayerManager.currentPlayer.PlayerType = Computer Then
        tmrClose.Interval = tmrClose.Interval / PlayerManager.currentPlayer.Speed
        tmrClose.Enabled = True
    End If
    
    Me.Show vbModal
    
    Chance (SelectedCard)
    '' SM:  TO DO - should we unload here??
End Sub

Private Sub cmdOK_Click()
    '' TO DO:  Should this just me "hide"?
    Unload Me
End Sub

Private Sub CommChest(CardNumber As Integer)
    Dim iTemptotal  As Long
    Dim player As claIPlayer

    Select Case CardNumber
    Case 1 'Doc Fee (-50)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -50
        Globals.CheckAddFreeParkingCash 50
    Case 2 'School Tax (-150)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -150
        Globals.CheckAddFreeParkingCash 150
    Case 3 'Xmas Matures (+100)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 100
    Case 4 'Pay Hospital (-100)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -100
        Globals.CheckAddFreeParkingCash 100
    Case 5 'Inherit (+100)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 100
    Case 6 'Life Insuance Matures (+100)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 100
    Case 7 'Sale of Stock (+45)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 45
    Case 8 'Bank Error (+200)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 200
    Case 9 'Grand Opening (+50 From Every Player)
        For Each player In PlayerManager
            If player.Number <> PlayerManager.currentPlayer.Number Then
                PlayerManager.currentPlayer.SimpleTransaction Nothing, 50
            End If
        Next player
    Case 10 'OOJail
        PlayerManager.currentPlayer.OOJailComm = True
    Case 11 'Assesed (-40 Per House, -115 Per Hotel)
        iTemptotal = 40 * PropertyManager.PlayerNumberHouses(PlayerManager.currentPlayer.Number)
        iTemptotal = iTemptotal + (115 * PropertyManager.PlayerNumberHotels(PlayerManager.currentPlayer.Number))
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -iTemptotal
        Globals.CheckAddFreeParkingCash iTemptotal
    Case 12 'Income Tax refund (+20)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 20
    Case 13 'Second Prize Beauty (+10)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 10
    Case 14 'Advance to Go (+200)
        AdvanceTo 40
    Case 15 'Recieve for Service (+25)
        PlayerManager.currentPlayer.SimpleTransaction Nothing, 25
    End Select
End Sub

Public Sub CommChestShow()
    Dim SelectedCard As Integer

    Me.Caption = "Community Chest"
    SelectedCard = CommChestCards.GetCardNumber
    'If selected card is OOJail Card but Player already has one then pick another card
    If CommChestCards.PlayerHasOOJailCard And SelectedCard = 10 Then
        SelectedCard = CommChestCards.GetCardNumber
    End If
    Me.imgCard.Picture = imgCommChest.ListImages(SelectedCard).Picture
    '' SM: DANGEOUS!  All events fire here!
    DoEvents
    
    If PlayerManager.currentPlayer.PlayerType = Computer Then
        tmrClose.Interval = tmrClose.Interval / PlayerManager.currentPlayer.Speed
        tmrClose.Enabled = True
    End If
    
    Me.Show vbModal
    
    CommChest (SelectedCard)
    
    '' SM: TO do - unload here?
End Sub

Private Sub tmrClose_Timer()
    tmrClose.Enabled = False
    '' SM:  Should this just hide?
    cmdOK_Click
End Sub
