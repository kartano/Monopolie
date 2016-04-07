VERSION 5.00
Begin VB.Form frmCards 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   7620
   ClientLeft      =   1665
   ClientTop       =   2160
   ClientWidth     =   9690
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPlayers 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   7260
      Width           =   2175
   End
   Begin VB.CheckBox chkShowAllCards 
      Caption         =   "Show &All Cards"
      Height          =   255
      Left            =   7380
      TabIndex        =   12
      Top             =   7260
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Purple"
      Height          =   2475
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   3
         Left            =   420
         Picture         =   "frmCards.frx":0000
         Top             =   500
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   1
         Left            =   60
         Picture         =   "frmCards.frx":77CA
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Aqua"
      Height          =   2475
      Left            =   2460
      TabIndex        =   1
      Top             =   0
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   9
         Left            =   780
         Picture         =   "frmCards.frx":EF94
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   8
         Left            =   420
         Picture         =   "frmCards.frx":1675E
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   6
         Left            =   60
         Picture         =   "frmCards.frx":1DF28
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pink"
      Height          =   2475
      Left            =   4860
      TabIndex        =   2
      Top             =   0
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   14
         Left            =   780
         Picture         =   "frmCards.frx":256F2
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   13
         Left            =   420
         Picture         =   "frmCards.frx":2CEBC
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   11
         Left            =   60
         Picture         =   "frmCards.frx":34686
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Orange"
      Height          =   2475
      Left            =   7260
      TabIndex        =   3
      Top             =   0
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   19
         Left            =   780
         Picture         =   "frmCards.frx":3BE50
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   18
         Left            =   420
         Picture         =   "frmCards.frx":4361A
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   16
         Left            =   60
         Picture         =   "frmCards.frx":4ADE4
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Red"
      Height          =   2475
      Left            =   60
      TabIndex        =   4
      Top             =   2460
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   24
         Left            =   780
         Picture         =   "frmCards.frx":525AE
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   23
         Left            =   420
         Picture         =   "frmCards.frx":59D78
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   21
         Left            =   60
         Picture         =   "frmCards.frx":61542
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Yellow"
      Height          =   2475
      Left            =   2460
      TabIndex        =   5
      Top             =   2460
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   29
         Left            =   780
         Picture         =   "frmCards.frx":68D0C
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   27
         Left            =   420
         Picture         =   "frmCards.frx":704D6
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   26
         Left            =   60
         Picture         =   "frmCards.frx":77CA0
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Green"
      Height          =   2475
      Left            =   4860
      TabIndex        =   6
      Top             =   2460
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   34
         Left            =   780
         Picture         =   "frmCards.frx":7F46A
         Top             =   900
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   32
         Left            =   420
         Picture         =   "frmCards.frx":86C34
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   31
         Left            =   60
         Picture         =   "frmCards.frx":8E3FE
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Blue"
      Height          =   2475
      Left            =   7260
      TabIndex        =   7
      Top             =   2460
      Width           =   2355
      Begin VB.Image P 
         Height          =   1530
         Index           =   39
         Left            =   420
         Picture         =   "frmCards.frx":95BC8
         Top             =   550
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   37
         Left            =   60
         Picture         =   "frmCards.frx":9D392
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Railroads"
      Height          =   2295
      Left            =   60
      TabIndex        =   8
      Top             =   4920
      Width           =   3975
      Begin VB.Image P 
         Height          =   1530
         Index           =   35
         Left            =   2400
         Picture         =   "frmCards.frx":A4B5C
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   25
         Left            =   1980
         Picture         =   "frmCards.frx":AC326
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   15
         Left            =   420
         Picture         =   "frmCards.frx":B3AF0
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   5
         Left            =   60
         Picture         =   "frmCards.frx":BB2BA
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Utilities"
      Height          =   2295
      Left            =   4080
      TabIndex        =   9
      Top             =   4920
      Width           =   2175
      Begin VB.Image P 
         Height          =   1530
         Index           =   28
         Left            =   600
         Picture         =   "frmCards.frx":C2A84
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image P 
         Height          =   1530
         Index           =   12
         Left            =   60
         Picture         =   "frmCards.frx":CA24E
         Top             =   200
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Get Out of Jail Free"
      Height          =   2295
      Left            =   6300
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
      Begin VB.Image imgJailComm 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1080
         Left            =   360
         Picture         =   "frmCards.frx":D1A18
         Top             =   960
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Image imgJailChance 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1080
         Left            =   45
         Picture         =   "frmCards.frx":D8C1A
         Top             =   300
         Visible         =   0   'False
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmCards
' Date      : 1/2/2003
' Author    :
' Purpose   :This was added a a possible cahnge to the existing Cards Form
' It positions the cards in the correct order, as well as saves room. Since a user can
' click on a card to view the information, I figured it was okay to cascade the cards this way.
' 21/08/2004    SM:  Now works for unowned properties
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub Run(thePlayerToShow As Integer)
    SelectedUser thePlayerToShow
    Me.Show vbModal
    Unload Me
End Sub

Private Sub cboPlayers_Click()
'***********************************************************************************
' Procedure : cboPlayers_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to select a player when viewing the form during gameplay
'***********************************************************************************
    SelectedUser cboPlayers.ListIndex
End Sub

Private Sub chkShowAllCards_Click()
'***********************************************************************************
' Procedure : chkShowAllCards_Click
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to show all Property Cards so the Player
'***********************************************************************************
    Dim iCount As Integer

    If chkShowAllCards.Value = 1 Then
        For iCount = 0 To 39
            If PropertyManager.isProperty(iCount) Then
                P(iCount).Visible = True
            End If
        Next iCount
    Else 'NOT CHKSHOWALLCARDS.VALUE...
        SelectedUser cboPlayers.ListIndex
    End If
End Sub

Private Sub Form_Activate()
'***********************************************************************************
' Procedure : Form_Activate
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : If the form goes out of focus then comes back
'             this refreshes the display
'***********************************************************************************
  SelectedUser cboPlayers.ListIndex

End Sub

Private Sub Form_Load()
'***********************************************************************************
' Procedure : Form_Load
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Loads the ComboBoxes upon Form load
'***********************************************************************************
    Dim lCount As Integer

    cboPlayers.Clear
    cboPlayers.AddItem "Unowned Properties"
    For lCount = 1 To PlayerManager.Count
        cboPlayers.AddItem (PlayerManager.player(lCount).Name)
    Next lCount
End Sub

Private Sub P_click(index As Integer)

  '
  ' User clicked on a property.  Display the property options form for the selected property
  '   -James
  '

  frmProperty.Run index, False

End Sub

Private Sub SelectedUser(ByVal byPlayerNumber As Integer)
    ' Displays the players cards in a new window.  Moves through each property, if it's
    ' owned by the player then it makes it visible.
    Dim prop As claIProperty
    
    For Each prop In PropertyManager
        If prop.Owner = byPlayerNumber Or Me.chkShowAllCards Then
            P(prop.BoardLocation).Visible = True
        Else
            P(prop.BoardLocation).Visible = False
        End If
    Next prop
    
    If byPlayerNumber > 0 Then
        imgJailChance.Visible = PlayerManager.player(byPlayerNumber).OOJailChance
        imgJailComm.Visible = PlayerManager.player(byPlayerNumber).OOJailComm
    End If
    cboPlayers.ListIndex = byPlayerNumber
End Sub

