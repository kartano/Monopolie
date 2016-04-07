VERSION 5.00
Begin VB.Form frmProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Property Options"
   ClientHeight    =   3735
   ClientLeft      =   4200
   ClientTop       =   3705
   ClientWidth     =   4215
   Icon            =   "frmProperty.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   3300
   End
   Begin VB.CommandButton cmdTrade 
      Caption         =   "&Trade"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdMortgage 
      Caption         =   "&Mortgage"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdHouses 
      Caption         =   "&Houses"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      Default         =   -1  'True
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "&Buy"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblMortgage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MORTGAGED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1365
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label OwnerStatus 
      Alignment       =   2  'Center
      Caption         =   "Property is available."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   4020
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   39
      Left            =   1260
      Picture         =   "frmProperty.frx":000C
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   37
      Left            =   1260
      Picture         =   "frmProperty.frx":1C716
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   35
      Left            =   1260
      Picture         =   "frmProperty.frx":38E20
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   34
      Left            =   1260
      Picture         =   "frmProperty.frx":5552A
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   32
      Left            =   1260
      Picture         =   "frmProperty.frx":71C34
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   31
      Left            =   1260
      Picture         =   "frmProperty.frx":8E33E
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   29
      Left            =   1260
      Picture         =   "frmProperty.frx":AAA48
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   28
      Left            =   1260
      Picture         =   "frmProperty.frx":C7152
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   27
      Left            =   1260
      Picture         =   "frmProperty.frx":E385C
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   26
      Left            =   1260
      Picture         =   "frmProperty.frx":FFF66
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   25
      Left            =   1260
      Picture         =   "frmProperty.frx":11C670
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   24
      Left            =   1260
      Picture         =   "frmProperty.frx":138D7A
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   23
      Left            =   1260
      Picture         =   "frmProperty.frx":155484
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   21
      Left            =   1260
      Picture         =   "frmProperty.frx":171B8E
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   19
      Left            =   1260
      Picture         =   "frmProperty.frx":18E298
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   18
      Left            =   1260
      Picture         =   "frmProperty.frx":1AA9A2
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   16
      Left            =   1260
      Picture         =   "frmProperty.frx":1C70AC
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   15
      Left            =   1260
      Picture         =   "frmProperty.frx":1E37B6
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   14
      Left            =   1260
      Picture         =   "frmProperty.frx":1FFEC0
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   13
      Left            =   1260
      Picture         =   "frmProperty.frx":21C5CA
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   12
      Left            =   1260
      Picture         =   "frmProperty.frx":238CD4
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   11
      Left            =   1260
      Picture         =   "frmProperty.frx":2553DE
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   9
      Left            =   1260
      Picture         =   "frmProperty.frx":271AE8
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   8
      Left            =   1260
      Picture         =   "frmProperty.frx":28E1F2
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   6
      Left            =   1260
      Picture         =   "frmProperty.frx":2AA8FC
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   5
      Left            =   1260
      Picture         =   "frmProperty.frx":2C7006
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   3
      Left            =   1260
      Picture         =   "frmProperty.frx":2E3710
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Image P 
      Height          =   2970
      Index           =   1
      Left            =   1260
      Picture         =   "frmProperty.frx":2FFE1A
      Top             =   120
      Visible         =   0   'False
      Width           =   2925
   End
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmProperty
' Date      : 10/5/2003
' Author    : Unknown
' Purpose   : Used to Display an individual property
' 03/03/2006            SM:  Standardised code header
'---------------------------------------------------------------------------------------

Option Explicit

Private iCurrentProperty      As Integer
Private iPurchaseAmount       As Long

Private mReadOnly As Boolean

' SM:  ReadOnly basically just disables the buttons so that
' the human players can see the property as the CPU buys it.
Public Sub run(theLocation As Integer, ReadOnly As Boolean)
    mReadOnly = ReadOnly
    iCurrentProperty = theLocation
    P(iCurrentProperty).Visible = True
    UpdateButtons theLocation
  
    Me.Show vbModal
  
    '' SM:  Probably don't need to do any of this if we
    '' are unloading the form anyway?
    P(iCurrentProperty).Visible = False
    iCurrentProperty = 0
    iPurchaseAmount = 0
    
    Unload Me
End Sub

Private Sub cmdBuy_Click()
    Dim iTemp As Integer

    'These actions are executed reguardless of which property was purchased
    iTemp = PlayerManager.currentPlayer.Number
    PropertyManager.BuyProperty iTemp, iCurrentProperty
      
    '' SM:  Supply an "other player" value of nothing, as
    '' we are dealing with the bank.
    PlayerManager.currentPlayer.SimpleTransaction Nothing, -iPurchaseAmount
    UpdateButtons (iCurrentProperty)
End Sub

Private Sub cmdDone_Click()
    Timer.Enabled = False
    Me.Hide
End Sub

Private Sub cmdHouses_Click()
    ' Assigns the current property to the houses dialogue and opens it
    ' If the selelcted property is a railroad or utility, do nothing.
    ' @author James
    '
    ' Check to make sure the player owns the entire monopoly.
    ' For multiplayer support, change the hardcoded "1" below
    ' to a variable to be set for each player.

    Select Case PropertyManager.PlayerOwnsMonopoly_ByPropNumber(1, iCurrentProperty)
    Case True
        Unload Me
        frmHousing.run iCurrentProperty
    Case False     'player does not own the entire monopoly
        Beep
        MsgBox "You may only purchase houses or hotels when you own the entire monopoly.", vbOKOnly, "Error: Monopoly unowned" 'display error message
    End Select
End Sub

Private Sub cmdMortgage_Click()
    Dim bResponse As Boolean
    Dim iReply    As Integer
    
    With PropertyManager.property(iCurrentProperty)
        If .MortgageStatus = staMortgaged Then
            .MortgageStatus = staUnmortgaged
            PlayerManager.player(.Owner).SimpleTransaction Nothing, -.MortgagePrice * 1.1
            If bResponse Then
                MsgBox "Property has been Unmortgaged", vbOKOnly, "Unmortgaged"
            End If
        Else ' property is unmortgaged
            If .Houses > 0 Then
                iReply = MsgBox("You cannot mortgage, please sell houses then try again." & vbNewLine & "Do you want to go there now?", vbYesNo, "Cannot Mortgage")
                If iReply = vbYes Then
                    cmdHouses_Click
                End If
            Else
                .MortgageStatus = staMortgaged
                PlayerManager.player(.Owner).SimpleTransaction Nothing, .MortgagePrice
                If bResponse Then
                    MsgBox "Property has been Unmortgaged", vbOKOnly, "Unmortgaged"
                End If
            End If
        End If
    End With
    UpdateButtons (iCurrentProperty)
End Sub

Private Sub SetPropertyOwner(PropertyNo As Integer)
    Dim prop As claIProperty
    
    '' SM:  CHanged to use new PropertyManager class
    Set prop = PropertyManager.property(PropertyNo)
    If prop.Owner = 0 Then
        OwnerStatus.Caption = "Available for purchase."
    Else
        OwnerStatus.Caption = "Owned by " & PlayerManager.player(prop.Owner).Name & "."
    End If
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    Me.Hide
End Sub

Private Sub UpdateButtons(Location As Integer)
    Dim prop As claIProperty
    
    Set prop = PropertyManager.property(Location)
    ' The CPU is buying this property
    If mReadOnly Then
        Me.cmdBuy.Enabled = False
        Me.cmdHouses.Enabled = False
        Me.cmdTrade.Enabled = False
        Me.cmdMortgage.Enabled = False
    ' The property is unowned
    ElseIf prop.Owner = 0 Then
        iPurchaseAmount = prop.PurchasePrice
        If PlayerManager.currentPlayer.Location = iCurrentProperty Then
            If PlayerManager.currentPlayer.Money > iPurchaseAmount Then
                cmdBuy.Enabled = True
            End If
        End If
        cmdMortgage.Enabled = False
        cmdHouses.Enabled = False
        cmdTrade.Enabled = False
    ' Someone owns the property
    Else 'NOT IOWNEDTEMP(0)...
        If PlayerManager.currentPlayer.Number = prop.Owner Then
            cmdMortgage.Enabled = True
            cmdHouses.Enabled = True
            cmdTrade.Enabled = True
        Else 'NOT PlayerManager.CURRENTPLAYER.NUMBER...
            cmdMortgage.Enabled = False
            cmdHouses.Enabled = False
            cmdTrade.Enabled = False
        End If
        cmdBuy.Enabled = False
        If prop.MortgageStatus = staMortgaged Then
            cmdMortgage.Caption = "Unmortgage"
            lblMortgage.Visible = True
        Else 'NOT PROPERTIES.PROPERTYSTATUS(LOCATION)...
            cmdMortgage.Caption = "Mortgage"
            lblMortgage.Visible = False
        End If
    End If
    SetPropertyOwner Location
End Sub
