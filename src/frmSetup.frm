VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Game Setup Wizard"
   ClientHeight    =   4800
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   6870
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin Monpolie.GroupBox fraSetup1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Step 1:"
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   3
         Top             =   900
         Width           =   5055
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmSetup.frx":000C
         Left            =   1260
         List            =   "frmSetup.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   5055
      End
      Begin VB.ComboBox cboDifficulty 
         Height          =   315
         ItemData        =   "frmSetup.frx":0021
         Left            =   1260
         List            =   "frmSetup.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1980
         Width           =   5055
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         ItemData        =   "frmSetup.frx":0032
         Left            =   1260
         List            =   "frmSetup.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   5070
      End
      Begin VB.ComboBox cboToken 
         Height          =   315
         ItemData        =   "frmSetup.frx":0066
         Left            =   1260
         List            =   "frmSetup.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3060
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddPlayer 
         Caption         =   "&Add Player"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hello! Welcome to the Monopolie new game setup wizard. You will need to enter your name, and choose some settings for the game."
         Height          =   555
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Na&me:"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "&Type:"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "&Difficulty:"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1980
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "&Speed"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblToken 
         Alignment       =   1  'Right Justify
         Caption         =   "&Token:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3060
         Width           =   855
      End
   End
   Begin Monpolie.GroupBox fraSetup2 
      Height          =   4095
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Step 2:"
      Begin VB.CommandButton cmdAdvancedOptions 
         Caption         =   "View Advanced Game Options"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Width           =   3375
      End
      Begin VB.CommandButton cmdDeletePlayer 
         Caption         =   "&Delete Player"
         Height          =   375
         Left            =   5280
         TabIndex        =   16
         Top             =   3600
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstPlayers 
         Height          =   2775
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Player Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Token"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "These are the selected players for the game."
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next>>"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   4320
      Width           =   1155
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmSetup
' Date      : 10/12/2003
' Author    : Brian G. Schmitt
' Purpose   : To setup the board for a new game
' 07/09/2004            SM:  Removed use of redundant "PlayerCount" value
' 08/20/2004            SM:  Player names can no longer be numeric
' 03/03/2006            SM:  Removed a lot of cruft, tidied indents
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Run()
    Me.Show vbModal
    Unload Me
End Sub

Private Sub cboType_Click()
    If UCase$(cboType.Text) = "COMPUTER" Then
        cboDifficulty.Enabled = True
        cboSpeed.Enabled = True
    Else
        cboDifficulty.Enabled = False
        cboDifficulty.ListIndex = 0
        cboSpeed.Enabled = False
        cboSpeed.ListIndex = 0
    End If
End Sub

Private Sub cmdAddPlayer_Click()
    Dim newPlayer As claIPlayer
    Dim player As claIPlayer
    
    ' Sanity check
    If Len(Trim$(Me.txtName)) = 0 Then
        SelfClosingMsgbox "Please enter a player name", vbOKOnly + vbInformation, "Add Player"
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    ' SM:  Prevent error 457's when adding player frames
    For Each player In PlayerManager
        If UCase$(Me.txtName) = UCase$(player.Name) Then
            SelfClosingMsgbox "Player name '" & Me.txtName & "' has already been used!", vbOKOnly + vbInformation, "Add Player"
            Me.txtName.SetFocus
            Exit Sub
        End If
    Next player
    
    ' SM:  Prevent error 13 type mismatch when adding player frames
    If IsNumeric(Me.txtName) Then
        SelfClosingMsgbox "Player names must contain numbers and letters", vbOKOnly + vbInformation, "Add Player"
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    If cboType.ListIndex = 0 Then
        Set newPlayer = New claHumanPlayer
    Else
        Set newPlayer = New claCPUPlayer
    End If
    
    With newPlayer
        .Name = Me.txtName
        .Token = cboToken.ItemData(cboToken.ListIndex)
        .Speed = cboSpeed.ListIndex
        .Difficulty = cboDifficulty.ListIndex
    End With
    
    PlayerManager.Add newPlayer
  
    If PlayerManager.Count >= 2 Then
        cmdNext.Enabled = True
    End If
  
    txtName.Text = vbNullString
    cboToken.RemoveItem (cboToken.ListIndex)
    cboToken.ListIndex = 0
    txtName.SetFocus
    If PlayerManager.Count = 8 Then
        cmdAddPlayer.Enabled = False
        NextPlayer
    End If
End Sub

Private Sub cmdAdvancedOptions_Click()
    frmOptions.Run
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDeletePlayer_Click()
    Dim iDelPlayer As Integer
    Dim iCounter   As Integer

    iDelPlayer = lstPlayers.SelectedItem.index
  
    cboToken.AddItem PlayerManager.player(iDelPlayer).Token
    PlayerManager.Delete iDelPlayer
    
    For iCounter = iDelPlayer To PlayerManager.Count
        PlayerManager.player(iCounter).Number = iCounter
    Next iCounter
    
    If PlayerManager.Count < 2 Then
        cmdFinish.Enabled = False
    End If
    If PlayerManager.Count < 8 Then
        cmdAddPlayer.Enabled = True
    End If
    LoadListBox
End Sub

Private Sub cmdFinish_Click()
    Dim player As claIPlayer
    Dim Number As Integer
    
    InitClasses
    mdiMonopoly.ResetForNewGame
    '' SM:  No longer needed - players already loaded
    '' in Playermanager
    Number = 1
    For Each player In PlayerManager
        player.Number = Number
        Number = Number + 1
        player.AddFrame
    Next player
    Unload Me
    SelfClosingMsgbox "The Game is ready to begin.", vbOKOnly + vbInformation, "Ready", 5 - PlayerManager.currentPlayer.Speed
    mdiMonopoly.bNewGame = True
    Globals.GameInProgress = True
End Sub

Private Sub cmdNext_Click()
    NextPlayer
End Sub

Private Sub cmdPrevious_Click()
    fraSetup2.Visible = False
    fraSetup1.Visible = True
    cmdNext.Enabled = True
    cmdNext.Default = True
    cmdNext.SetFocus
    cmdPrevious.Enabled = False
    cmdFinish.Enabled = False
End Sub

Private Sub Form_Load()
    With cboType
        .Clear
        .AddItem "Human", 0
        .AddItem "Computer", 1
        .ListIndex = 0
    End With
    With cboDifficulty
        .Clear
        .AddItem "Easy", 0
        .AddItem "Medium", 1
        .AddItem "Hard", 2
        .ListIndex = 0
    End With
    With cboSpeed
        .Clear
        .AddItem "Slow", 0
        .AddItem "Normal", 1
        .AddItem "Fast", 2
        .AddItem "SuperFast", 3
        .ListIndex = 1
    End With
    With cboToken
        .Clear
        .AddItem "Cannon", 0
        .ItemData(0) = 0
        .AddItem "Car", 1
        .ItemData(1) = 1
        .AddItem "Dog", 2
        .ItemData(2) = 2
        .AddItem "Hat", 3
        .ItemData(3) = 3
        .AddItem "Horse", 4
        .ItemData(4) = 4
        .AddItem "Iron", 5
        .ItemData(5) = 5
        .AddItem "Money Bag", 6
        .ItemData(6) = 6
        .AddItem "Ship", 7
        .ItemData(7) = 7
        .AddItem "Shoe", 8
        .ItemData(8) = 8
        .AddItem "Thimble", 9
        .ItemData(9) = 9
        .AddItem "WheelBarrow", 10
        .ItemData(10) = 10
        .ListIndex = 0
    End With
  
    With lstPlayers
        .ColumnHeaders(1).Width = .Width / 3
        .ColumnHeaders(2).Width = .Width / 3
        .ColumnHeaders(3).Width = .Width / 3 - 100
    End With
    
    Label1.Caption = "Hello! Welcome to the Monopolie new game setup wizard." & vbNewLine & "You will need to enter your name, and choose some settings for the game."
    cmdNext.Enabled = False
    cmdFinish.Enabled = False
    cmdAddPlayer.Default = True
End Sub

Private Sub LoadListBox()
    '' SM:  Adjusted to use new Player Manager
    
    Dim player As claIPlayer
    Dim TempListItem As ListItem
    Dim iCount As Integer
    
    lstPlayers.ListItems.Clear
    iCount = 1
    For Each player In PlayerManager
        Set TempListItem = lstPlayers.ListItems.Add()
        TempListItem.Text = player.Name
        lstPlayers.ListItems(iCount).ListSubItems.Add , , PlayerTypeToWords(player.PlayerType)
        lstPlayers.ListItems(iCount).ListSubItems.Add , , TokenToWords(player.Token)
        iCount = iCount + 1
    Next player
End Sub

Private Sub NextPlayer()
    LoadListBox
    fraSetup2.Visible = True
    fraSetup1.Visible = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = True
    cmdFinish.Enabled = True
    cmdFinish.Default = True
    cmdFinish.SetFocus
End Sub

