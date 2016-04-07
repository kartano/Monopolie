VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debug Form"
   ClientHeight    =   5280
   ClientLeft      =   4710
   ClientTop       =   3915
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin Monpolie.GroupBox fraPlayer 
      Height          =   2490
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   4392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Player Debug:"
      Begin VB.Timer Timer 
         Interval        =   100
         Left            =   2280
         Top             =   960
      End
      Begin VB.CommandButton cmdBuyOneProp 
         Caption         =   "Own &Specific Property"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2070
         Width           =   1815
      End
      Begin VB.CommandButton cmdBuyAll 
         Caption         =   "Own &All Property"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1710
         Width           =   1815
      End
      Begin VB.ComboBox cboPlayerSpeed 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1470
         Width           =   1575
      End
      Begin VB.ComboBox cboPlayerDiff 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1110
         Width           =   1575
      End
      Begin VB.ComboBox cboPlayerType 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   750
         Width           =   1575
      End
      Begin VB.TextBox txtMoney 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdCommit 
         Caption         =   "&Change"
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   2070
         Width           =   1695
      End
      Begin VB.ComboBox cboPlayer 
         Height          =   315
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   225
         Width           =   2055
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Out Jail Chance"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1110
         Width           =   1455
      End
      Begin VB.TextBox txtJailCount 
         Height          =   285
         Left            =   1050
         TabIndex        =   8
         Top             =   630
         Width           =   375
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "In Jail"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   855
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Out Jail Community Chest"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label lblAddress 
         Caption         =   "Network Address"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1470
         Width           =   1695
      End
      Begin VB.Label lblToken 
         Caption         =   "Token"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboCommChest 
      Height          =   315
      ItemData        =   "frmDebug.frx":0000
      Left            =   2640
      List            =   "frmDebug.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox cboChance 
      Height          =   315
      ItemData        =   "frmDebug.frx":0106
      Left            =   2640
      List            =   "frmDebug.frx":0137
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtDoublesCount 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtRoll 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "5"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CheckBox chkRoll 
      Caption         =   "Roll:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox chkDoubles 
      Caption         =   "Force Doubles"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtStatus 
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Debug Text Logging Text Goes Here"
      Top             =   3360
      Width           =   4575
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' Module    : frmDebug
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Version   : 1.0
' Copyright : Copyright © Brian G. Schmitt
' Purpose   : Used during Developmnet to "Test" Features
'             To Use After Compiled run Monopolie with
'             Command Line Switch "/Debug"
' Revisions : Developer           Date     Comments
'             ---------           -------- --------
'             SM                  07/02/2004 Fixes due to new property manager class
'             SM                  07/07/2004 Fixed "error 13" when entering non-numeric
'                                            values for "own specific property"
' ======================================================================================
Option Explicit

Private lCount            As Long
Private bChangeStatus     As Boolean
Private bChangeMoney      As Boolean

Private Sub cboChance_Click()
'***********************************************************************************
' Procedure : cboChance_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to set the next Chance card that will be used
'***********************************************************************************
  If cboChance.ListIndex + 1 = 15 Then
    ChanceCards.CardNumber = 0
   Else 'NOT CBOCHANCE.LISTINDEX...
    ChanceCards.CardNumber = cboChance.ListIndex + 1
  End If

End Sub

Private Sub cboCommChest_Change()
'***********************************************************************************
' Procedure : cboCommChest_Change
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to set the Next CommChest Card for gameplay
'***********************************************************************************
  If cboCommChest.ListIndex + 1 = 16 Then
    CommChestCards.CardNumber = 0
   Else 'NOT CBOCOMMCHEST.LISTINDEX...
    CommChestCards.CardNumber = cboCommChest.ListIndex + 1
  End If

End Sub

Private Sub cboPlayer_Click()
'***********************************************************************************
' Procedure : cboPlayer_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Sets the Debug screen to the selected players values
'***********************************************************************************
  RefreshDebug

End Sub

Private Sub chkDoubles_Click()
'***********************************************************************************
' Procedure : chkDoubles_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to Set the next roll to doubles,
'             Note: The Dice will NOT display this!
'***********************************************************************************

  If chkDoubles.Value = 1 Then
    txtDoublesCount.Enabled = True
   Else 'NOT CHKROLL.VALUE...'NOT CHKDOUBLES.VALUE...
    txtDoublesCount.Enabled = False
  End If

End Sub

Private Sub chkRoll_Click()
'***********************************************************************************
' Procedure : chkRoll_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : If checked will set the Value of the next roll
'             To the value in the Txtbox below it
'***********************************************************************************

  If chkRoll.Value = 1 Then
    txtRoll.Enabled = True
   Else 'NOT CHKROLL.VALUE...
    txtRoll.Enabled = False
  End If

End Sub

Public Sub CloseMe()
'***********************************************************************************
' Procedure : CloseMe
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Unloads the form
'***********************************************************************************

  Unload Me

End Sub

Private Sub cmdBuyAll_Click()
'***********************************************************************************
' Procedure : cmdBuyAll_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to Buy All the property for the Selected User
'***********************************************************************************

    Dim iTemp As Integer

    With PropertyManager
        For iTemp = 1 To 39
            If .isProperty(iTemp) Then
                If .property(iTemp).CanImprove Then
                    If .property(iTemp).owner = 0 Then
                        .BuyProperty (cboPlayer.ListIndex + 1), iTemp
                    End If
                End If
            End If
        Next iTemp
    End With
End Sub

Private Sub cmdBuyOneProp_Click()
'***********************************************************************************
' Procedure : cmdBuyOneProp_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Buys ONE specific property by number must be unowned
'***********************************************************************************
    Dim iInput As Integer
    Dim inputBuffer As String

    inputBuffer = InputBox("Which Property Number would you like to own?", "Own Specific Property", 1)
    If IsNumeric(inputBuffer) Then
        iInput = CInt(inputBuffer)
        If PropertyManager.isProperty(iInput) Then
            If PropertyManager.property(CInt(iInput)).owner Then
                PropertyManager.BuyProperty (cboPlayer.ListIndex + 1), iInput
            End If
        Else
            MsgBox "Property number " & iInput & " is not a property."
        End If
    Else
        SelfClosingMsgbox "Numeric only input, please!", vbOKOnly + vbInformation, "Property Number"
    End If
End Sub

Private Sub cmdCommit_Click()
'***********************************************************************************
' Procedure : cmdCommit_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Commits the changes for that player
'***********************************************************************************

  If bChangeStatus = False Then
    Timer.Enabled = False
    cmdCommit.Caption = "&Commit Changes"
    cboPlayer.Enabled = False
    bChangeStatus = True
    UpdateControls
   Else 'NOT BCHANGESTATUS...
    cmdCommit.Caption = "&Change"
    cboPlayer.Enabled = True
    Timer.Enabled = True
    bChangeStatus = False
    UpdateControls
    PlayerManager.currentPlayer.InJail = chkOption(0).Value
    PlayerManager.currentPlayer.JailCount = txtJailCount.Text
    PlayerManager.currentPlayer.OOJailComm = chkOption(1).Value
    PlayerManager.currentPlayer.OOJailChance = chkOption(2).Value
    If bChangeMoney Then
      PlayerManager.player(cboPlayer.ListIndex + 1).ChangeMoney txtMoney.Text
    End If
  End If
  cboPlayer_Click

End Sub

Private Sub Form_Load()
'***********************************************************************************
' Procedure : Form_Load
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Form Loading--Load Defaults
'***********************************************************************************

  Me.Left = frmMain.Left + frmMain.Width
  Me.Top = frmMain.Top
  LoadComboBoxes
  cboPlayer.ListIndex = 0
  cboPlayer_Click
  cboChance.ListIndex = 14
  cboCommChest.ListIndex = 15
  bChangeStatus = False
  UpdateControls
  DebugMode = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'***********************************************************************************
' Procedure : Form_Unload
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Unloads the Form, Sets the Global Var of DebugMode to false
'***********************************************************************************

  DebugMode = False

End Sub

Private Sub LoadComboBoxes()
'***********************************************************************************
' Procedure : LoadComboBoxes
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to populate the combo boxes at form load
'***********************************************************************************

  Dim lCboTemp As Long
  Dim player As claIPlayer

  With cboPlayer
    .Clear
    ''For lCboTemp = 1 To PlayerManager.Count
    For Each player In PlayerManager
      .AddItem player.Name
    Next player
  End With 'CBOPLAYER
  With cboPlayerType
    .Clear
    .AddItem "Human", 0
    .AddItem "Computer", 1
  End With 'CBOPLAYERTYPE
  With cboPlayerDiff
    .Clear
    .AddItem "Easy", 0
    .AddItem "Medium", 1
    .AddItem "Hard", 2
  End With 'CBOPLAYERDIFF
  With cboPlayerSpeed
    .Clear
    .AddItem "Slow", 0
    .AddItem "Normal", 1
    .AddItem "Fast", 2
    .AddItem "SuperFast", 3
  End With 'CBOPLAYERSPEED

End Sub

Public Sub RefreshDebug()
'***********************************************************************************
' Procedure : RefreshDebug
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Updates the controls with the selected player values
'***********************************************************************************

  lblToken.Caption = PlayerManager.player(cboPlayer.ListIndex + 1).Token
  txtMoney.Text = PlayerManager.player(cboPlayer.ListIndex + 1).Money
  If PlayerManager.player(cboPlayer.ListIndex + 1).JailCount > 0 Then
    chkOption(0).Value = 1
   Else 'NOT PlayerManager.GETPLAYER(CBOPLAYER.LISTINDEX...
    chkOption(0).Value = 0
  End If
  txtJailCount.Text = PlayerManager.player(cboPlayer.ListIndex + 1).JailCount
  If PlayerManager.player(cboPlayer.ListIndex + 1).OOJailComm Then
    chkOption(1).Value = 1
   Else 'NOT PlayerManager.GETPLAYER(CBOPLAYER.LISTINDEX...
    chkOption(1).Value = 0
  End If
  If PlayerManager.player(cboPlayer.ListIndex + 1).OOJailChance Then
    chkOption(2).Value = 1
   Else 'NOT PlayerManager.GETPLAYER(CBOPLAYER.LISTINDEX...
    chkOption(2).Value = 0
  End If
  txtDoublesCount.Text = PlayerManager.player(cboPlayer.ListIndex + 1).DoublesCount
  cboPlayerType.ListIndex = PlayerManager.player(cboPlayer.ListIndex + 1).PlayerType
  cboPlayerSpeed.ListIndex = PlayerManager.player(cboPlayer.ListIndex + 1).Speed
  '' SM:  Network play not added yet
  ''lblAddress.Caption = PlayerManager.player(cboPlayer.ListIndex + 1).NetworkAddress
    cboPlayerDiff.ListIndex = PlayerManager.player(cboPlayer.ListIndex + 1).Difficulty
  lCount = 0

End Sub

Public Function RollValue(OldRoll As Long) As Long
'***********************************************************************************
' Procedure : RollValue
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used as a loop during gameplay if in DebugMode=true
' Returns   : If chkRoll is selected return the debug screen roll value
'             If chkRoll is not selected return what was actually rolled by the player
'***********************************************************************************

  If chkRoll.Value = 1 Then
    RollValue = txtRoll.Text
   Else 'NOT CHKROLL.VALUE...
    RollValue = OldRoll
  End If

End Function

Private Sub Timer_Timer()
'***********************************************************************************
' Procedure : Timer_Timer
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to keep the Form docked next to the frmMain
'***********************************************************************************

  Timer.Enabled = False
  If lCount = 100 Then
    lCount = 0
    cboPlayer_Click
   Else 'NOT LCOUNT...
    lCount = lCount + 1
  End If
  Me.Left = mdiMonopoly.Left + mdiMonopoly.Width
  Me.Top = mdiMonopoly.Top
  Timer.Enabled = True

End Sub

Private Sub txtDoublesCount_LostFocus()
'***********************************************************************************
' Procedure : txtDoublesCount_LostFocus
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to set the doubles count of the player
'             Usefull in testing the Jail on 3 Doubles rule
'***********************************************************************************

  PlayerManager.player(cboPlayer.ListIndex + 1).DoublesCount = txtDoublesCount

End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
'***********************************************************************************
' Procedure : txtMoney_KeyPress
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to determine if the user actually changed the money value
'             If they do not change it, it doesn't update the players money
'***********************************************************************************

  bChangeMoney = True

End Sub

Private Sub UpdateControls()
'***********************************************************************************
' Procedure : UpdateControls
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to enable/disable the controls to edit player values
'***********************************************************************************

  chkOption(0).Enabled = bChangeStatus
  txtJailCount.Enabled = bChangeStatus
  chkOption(1).Enabled = bChangeStatus
  chkOption(2).Enabled = bChangeStatus
  cboPlayerType.Enabled = bChangeStatus
  cboPlayerDiff.Enabled = bChangeStatus
  cboPlayerSpeed.Enabled = bChangeStatus
  txtMoney.Enabled = bChangeStatus

End Sub
