VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monopolie"
   ClientHeight    =   5640
   ClientLeft      =   4275
   ClientTop       =   2850
   ClientWidth     =   5130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdRolldice 
      Caption         =   "&Roll Dice"
      Default         =   -1  'True
      Height          =   495
      Left            =   3900
      TabIndex        =   1
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4995
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4995
      Begin VB.Image imgMortgaged 
         Height          =   450
         Left            =   2280
         Picture         =   "frmMain.frx":030A
         Top             =   3840
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgOwnedProp 
         Height          =   2970
         Left            =   1000
         Top             =   720
         Width           =   2925
      End
      Begin VB.Image die2 
         Height          =   360
         Left            =   3720
         Picture         =   "frmMain.frx":0703
         Top             =   3720
         Width           =   405
      End
      Begin VB.Image die1 
         Height          =   360
         Left            =   3240
         Picture         =   "frmMain.frx":0F25
         Top             =   3720
         Width           =   405
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   1
         Left            =   3960
         Tag             =   "2"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   615
         Index           =   0
         Left            =   4380
         Top             =   4380
         Width           =   615
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   39
         Left            =   4560
         Tag             =   "50"
         Top             =   3960
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   38
         Left            =   4560
         Top             =   3540
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   37
         Left            =   4560
         Tag             =   "35"
         Top             =   3120
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   36
         Left            =   4560
         Top             =   2700
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   35
         Left            =   4560
         Top             =   2280
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   34
         Left            =   4560
         Tag             =   "28"
         Top             =   1860
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   33
         Left            =   4560
         Top             =   1440
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   32
         Left            =   4560
         Tag             =   "26"
         Top             =   1020
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   31
         Left            =   4560
         Tag             =   "26"
         Top             =   600
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   615
         Index           =   30
         Left            =   4380
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   29
         Left            =   3960
         Tag             =   "24"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   28
         Left            =   3540
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   27
         Left            =   3120
         Tag             =   "22"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   26
         Left            =   2700
         Tag             =   "22"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   25
         Left            =   2280
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   24
         Left            =   1860
         Tag             =   "20"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   23
         Left            =   1440
         Tag             =   "18"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   22
         Left            =   1020
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   21
         Left            =   600
         Tag             =   "18"
         Top             =   0
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   615
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   19
         Left            =   0
         Tag             =   "16"
         Top             =   600
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   18
         Left            =   0
         Tag             =   "14"
         Top             =   1020
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   17
         Left            =   0
         Top             =   1440
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   16
         Left            =   0
         Tag             =   "14"
         Top             =   1860
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   15
         Left            =   0
         Top             =   2280
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   14
         Left            =   0
         Tag             =   "12"
         Top             =   2700
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   13
         Left            =   0
         Tag             =   "10"
         Top             =   3120
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   12
         Left            =   0
         Top             =   3540
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   11
         Left            =   0
         Tag             =   "10"
         Top             =   3960
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   615
         Index           =   10
         Left            =   0
         Top             =   4380
         Width           =   615
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   9
         Left            =   600
         Tag             =   "8"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   8
         Left            =   1020
         Tag             =   "6"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   7
         Left            =   1440
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   6
         Left            =   1860
         Tag             =   "6"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   5
         Left            =   2280
         Tag             =   "25"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   4
         Left            =   2700
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   3
         Left            =   3120
         Tag             =   "4"
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Images 
         Height          =   435
         Index           =   2
         Left            =   3540
         Top             =   4560
         Width           =   435
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   0
         Left            =   4560
         Picture         =   "frmMain.frx":1747
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   1
         Left            =   4560
         Picture         =   "frmMain.frx":1B13
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   10
         Left            =   4560
         Picture         =   "frmMain.frx":1ED4
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   2
         Left            =   4560
         Picture         =   "frmMain.frx":2290
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   3
         Left            =   4560
         Picture         =   "frmMain.frx":2652
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   4
         Left            =   4560
         Picture         =   "frmMain.frx":29F4
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   5
         Left            =   4560
         Picture         =   "frmMain.frx":2DCD
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   6
         Left            =   4560
         Picture         =   "frmMain.frx":319B
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   7
         Left            =   4560
         Picture         =   "frmMain.frx":3569
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   8
         Left            =   4560
         Picture         =   "frmMain.frx":391D
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Token 
         Height          =   345
         Index           =   9
         Left            =   4560
         Picture         =   "frmMain.frx":3CD7
         Top             =   4560
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Image Board 
         Height          =   4995
         Left            =   0
         Picture         =   "frmMain.frx":40A5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4995
      End
   End
   Begin MSComctlLib.ImageList DiceImages 
      Left            =   2760
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   27
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":555AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55DE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56613
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56E45
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57677
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57EA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image blank 
      Height          =   255
      Left            =   2280
      Top             =   6120
      Width           =   195
   End
   Begin VB.Image imgFourhouse 
      Height          =   345
      Left            =   1800
      Picture         =   "frmMain.frx":586DB
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgThreehouse 
      Height          =   345
      Left            =   1320
      Picture         =   "frmMain.frx":58A67
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgTwohouse 
      Height          =   345
      Left            =   780
      Picture         =   "frmMain.frx":58DEB
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgOnehouse 
      Height          =   345
      Left            =   480
      Picture         =   "frmMain.frx":59162
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgHotel 
      Height          =   345
      Left            =   180
      Picture         =   "frmMain.frx":594CA
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' Date      : 10/5/2003
' Author    : Various
' Purpose   : Main MDI form
' 11/17/2003        SM:  Fixes to allow for new player classes
'                        General tidying of code
' 07/14/2004        SM:  LogFile functionality now configurable - doesn't use #const
' 03/03/2006        SM:  Removed some cruft, cleaned up indents
'---------------------------------------------------------------------------------------

Option Explicit

Private DoNotTerminate      As Boolean
Private bEndTurn            As Boolean
Private bTurnInProgress     As Boolean

Private Sub cmdRolldice_Click()
    RollTheDice
End Sub

Private Sub die1_Click()
  RollTheDice
End Sub

Private Sub die2_Click()
  RollTheDice
End Sub

Public Sub DisplayHouses()
    Dim prop As claIProperty
    
    '' SM:  Overhauled to use new PropertyManager
    For Each prop In PropertyManager
        If prop.MortgageStatus = staMortgaged Then
            Images(prop.BoardLocation).Picture = imgMortgaged.Picture
        Else
            Select Case prop.Houses
            Case 0
                Images(prop.BoardLocation).Picture = blank.Picture
            Case 1
                Images(prop.BoardLocation).Picture = imgOnehouse.Picture
            Case 2
                Images(prop.BoardLocation).Picture = imgTwohouse.Picture
            Case 3
                Images(prop.BoardLocation).Picture = imgThreehouse.Picture
            Case 4
                Images(prop.BoardLocation).Picture = imgFourhouse.Picture
            Case 5
                Images(prop.BoardLocation).Picture = imgHotel.Picture
            End Select
        End If
    Next prop
End Sub

Private Sub Form_Load()
    Dim Seed As Single
    Dim iCounter As Integer
    
    ' NOTE - IF YOU WANT TO HAVE THE RANDOM NUMBER GERNERATOR
    ' RETURN THE SAME SEQUENCE OF RANDOM NUMBERS EVERY TIME
    ' THEN USE A NEGATIVE VALUE WHEN CALLING Randomize
    '
    ' THIS WILL ALLOW YOU TO GET THE SAME GAME EVERY TIME
    ' SO THAT FINDING AND REPLICATING BUGS DURING TESTING
    ' IS EASIER.
    '
    ' Randomize -100
    Seed = Timer
    Randomize Seed

    AppendToLog "Randomize seed value is: " & CStr(Seed)
    
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    For iCounter = 0 To Token.Count - 1
        'Set tokens to the Top so they are on top of the house pieces
        Token(iCounter).ZOrder 0
    Next
End Sub

Private Sub Form_Resize()
    'trying to give the user the ability to resize gameboard
    'this is a work in progress...:)
    '    fraGame.Width = Me.Width - 210
    '    Board.Width = Me.Width - 600
    '    Board.Height = Me.Height - 1455
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AppendToLog "Ending Log at: " & CStr(Time) & "  " & CStr(Date)
    
    Set ChanceCards = Nothing
    Set CommChestCards = Nothing
    If (DoNotTerminate = True) Then
        ' ONLY STARTING A NEW GAME
        DoNotTerminate = False
        frmSetup.Show vbModal
    Else 'NOT (DONOTTERMINATE...
        ' ENDING THE PROGRAM
        Unload Me
    End If
    frmDebug.CloseMe
End Sub

Private Sub Images_Click(index As Integer)
    ' Runs when user clicks on one of the properties on the board.  Displays the frmProperty form,
    ' with the selected property.
    '   -James
    '
    'properties without actions:
    'We might want to add actions later.
    '0      Go
    '2, 17, 33    Community Chest
    '7, 22, 36    Chance
    '10     Jail / Just visiting
    '20     Free Parking
    '30     GoTo Jail
    '38     Luxury Tax
    'make sure the user clicked on a property.

    ' SM:  Ignore the corner suares - each has an index divisible eveninly by 10
    If (index Mod 10) <> 0 Then
        If PropertyManager.isProperty(index) Then
            frmProperty.run index, False
        End If
    End If
End Sub

Public Sub PlayersTurn()
    Dim iDiceMove As Integer
    Dim cpuPlayer As claCPUPlayer
    
    cmdRolldice.Enabled = False

    With PlayerManager
        '' SM:  Thought about making "PreRollLogic" part
        '' of the player interface, but for network and
        '' human players that wouldn't make sense.
        '' Anyone have any other ideas??
        If .currentPlayer.PlayerType = Computer Then
            '' SM:  DOing this because of the way VB6
            '' handles interfaces.
            Set cpuPlayer = .currentPlayer
            cpuPlayer.PreRollLogic
        End If
        
        If .currentPlayer.InJail = False Then
            If .currentPlayer.OOJailRoll = 0 Then
                iDiceMove = RollDice(.currentPlayer.Number)
            Else
                iDiceMove = .currentPlayer.OOJailRoll
                .currentPlayer.OOJailRoll = 0
            End If
            
            If bEndTurn = False Then
                MovePlayerToken .currentPlayer.Location, iDiceMove, 1
            End If
        Else
            .currentPlayer.DoSomethingIfInJail
            If .currentPlayer.InJail And .currentPlayer.JailCount >= 3 Then
                .currentPlayer.SimpleTransaction Nothing, 50
                If Not .currentPlayer.IsBankrupt Then
                    .currentPlayer.InJail = False
                End If
            End If
        End If
    
        DisplayHouses
        
        If DebugMode Then frmDebug.RefreshDebug
        
        If .currentPlayer.DoublesCount = 0 Then
            .PassTheDice
            bEndTurn = False
        End If
        
        If .currentPlayer.PlayerType = Computer Then
            '' SM:  DOing this because of the way VB6
            '' handles interfaces.
            Set cpuPlayer = .currentPlayer
            cpuPlayer.PostRollLogic
            
        End If
        
        '' SM:  Do NOT use "CurrentPlayer" from this point
        '' down!  Once .PassTheDice is called, the next
        '' player is primed and ready.
    End With
    cmdRolldice.Enabled = True
End Sub

Public Function RollDice(ByVal playerNumber As Long) As Long
    DiceRoll1 = Int(Rnd * 6) + 1
    DiceRoll2 = Int(Rnd * 6) + 1
    SetDicePictures DiceRoll1, DiceRoll2

    AppendToLog PlayerManager.currentPlayer.Name & " Rolled: " & DiceRoll1 & " & " & DiceRoll2
  
    RollDice = DiceRoll1 + DiceRoll2
    mdiMonopoly.StatusBar.Panels(2).Text = "Last Roll: " & RollDice
    'The following line was added so that we can go to a specific place to help with debugging
    If DebugMode Then
        RollDice = frmDebug.RollValue(RollDice)
        If frmDebug.chkDoubles.Value = 1 Then
            DiceRoll1 = 1
            DiceRoll2 = 1
        End If
    End If
    If (DiceRoll1 = DiceRoll2) Then
        PlayerManager.currentPlayer.AddDoubles
        mdiMonopoly.StatusBar.Panels(3).Text = "Doubles Count: " & PlayerManager.currentPlayer.DoublesCount
        If PlayerManager.currentPlayer.DoublesCount = 3 Then
            AppendToLog "Player " & CStr(playerNumber) & " Rolled DOUBLES three times in a row and is going to jail."
      
            SelfClosingMsgbox "Rolled Doubles Three times go directly to Jail", vbCritical + vbOKOnly, "Go to Jail.", 5 - PlayerManager.currentPlayer.Speed
      
            PlayerManager.currentPlayer.DoublesCount = 0
            bEndTurn = True
    
            '' SM:  Dangerous!  All control events
            '' on any active form can fire here!
            '' This includes window closes!
            DoEvents
    
            ' GIVE PLAYER TIME TO APPRECIATE THAT THEY HAVE ROLLED DOUBLES FOR THE THIRD TIME
            Sleep 2137
            MovePlayerTokenDirect 10, PlayerManager.currentPlayer.Number
            PlayerManager.currentPlayer.InJail = True
        End If
    Else
        ' No double was rolled
        PlayerManager.currentPlayer.DoublesCount = 0
        mdiMonopoly.StatusBar.Panels(3).Text = "Doubles Count: 0"
    End If
End Function

Public Sub SetDicePictures(ByVal Die_1_Val As Integer, ByVal Die_2_Val As Integer)
    With DiceImages
        die1.Picture = .ListImages(Die_1_Val).Picture
        die2.Picture = .ListImages(Die_2_Val).Picture
    End With
End Sub

Public Property Get TurnInProgress() As Boolean
    TurnInProgress = bTurnInProgress
End Property

Private Sub RollTheDice()
    Dim iCounter As Integer
    For iCounter = 1 To 39
        'set all property pics to nothing
        '' SM:  Why?  Shouldn't we leave houses/hotels displayed?
        frmMain.Images(iCounter).Picture = frmMain.Images(0).Picture
    Next iCounter
    
    '' SM:  TO DO: Force use of "Roll dice". Players
    '' should all be allowed to make trade offers and such
    '' during their turn.
    '' We should also check all CPU players to ask if they
    '' want to make a trade as well.
    bTurnInProgress = True
    Do
        If PlayerManager.currentPlayer.IsBankrupt Then
            PlayerManager.PassTheDice
        Else
            PlayersTurn
            PlayerManager.CheckForCPUTrades
        End If
        If PlayerManager.CheckForVictory Then Exit Do
    Loop Until PlayerManager.currentPlayer.PlayerType <> Computer And PlayerManager.currentPlayer.IsBankrupt = False
    bTurnInProgress = False
End Sub
