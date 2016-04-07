VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade"
   ClientHeight    =   6330
   ClientLeft      =   2925
   ClientTop       =   3165
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstPropertiesWanted 
      Height          =   2775
      Left            =   3120
      TabIndex        =   17
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4895
      View            =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imgThumbnails"
      SmallIcons      =   "imgThumbnails"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstPropertiesOffered 
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4895
      View            =   2
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imgThumbnails"
      SmallIcons      =   "imgThumbnails"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgThumbnails 
      Left            =   5280
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":0000
            Key             =   "Blue"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":0502
            Key             =   "Green"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":0A04
            Key             =   "Aqua"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":0F06
            Key             =   "Orange"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":1408
            Key             =   "Pink"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":190A
            Key             =   "Purple"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":1E0C
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":230E
            Key             =   "Utility"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":2810
            Key             =   "Red"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrade.frx":2D12
            Key             =   "Yellow"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSelectReceivingNone 
      Caption         =   "Select None"
      Height          =   315
      Left            =   3240
      TabIndex        =   15
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelectReceivingAll 
      Caption         =   "Select All"
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelectOfferNone 
      Caption         =   "Select None"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelectOfferAll 
      Caption         =   "Select All"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   2895
   End
   Begin VB.ComboBox cboPlayerReceivingCash 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtCashInvolved 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.ComboBox cboToPlayer 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox cboFromPlayer 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Which player receives the cash?"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "How much cash is involved:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Properties wanted:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Properties on Offer:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Which player is the trade offered to?"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Which player offers this trade?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTrade
' Date      : 11/21/2003
' Author    : Various
' Purpose   : Trade edit form
'---------------------------------------------------------------------------------------

Option Explicit

Private mTrade As claTrade

Private mReturnValue As Boolean

Private mFromProps As Collection
Private mToProps As Collection

' Returns TRUE if details saved, FALSE if not
Public Function Run(theTrade As claTrade) As Boolean
    Dim row As Integer
    Dim theColl As Collection
    
    Set mTrade = theTrade
    
    mReturnValue = False
    
    BuildPlayerList Me.cboFromPlayer
    BuildPlayerList Me.cboToPlayer
    
    If Not mTrade.PlayerMakingOffer Is Nothing Then
        Me.cboFromPlayer.ListIndex = mTrade.PlayerMakingOffer.Number
    End If
    If Not mTrade.PlayerReceivingOffer Is Nothing Then
        Me.cboToPlayer.ListIndex = mTrade.PlayerReceivingOffer.Number
    End If
    
    buildPropertyOffered
    buildPropertyWanted
    
    With theTrade
        Me.txtCashInvolved = .Cash
        BuildReceiveList
    End With
        
    Me.Show vbModal
    
    If mReturnValue Then
        ' SM:  DOn't add one to the index.  Index 0 in the combos
        ' is a "None" entry.
        With theTrade
            Set .PlayerMakingOffer = PlayerManager.player(Me.cboFromPlayer.ListIndex)
            Set .PlayerReceivingOffer = PlayerManager.player(Me.cboToPlayer.ListIndex)
            .Cash = CLng(Me.txtCashInvolved)
            If .Cash > 0 Then
                If Me.cboPlayerReceivingCash.ListIndex = 0 Then
                    .ReceivingPlayerGetsCash = False
                Else
                    .ReceivingPlayerGetsCash = True
                End If
            End If
        End With
        
        ' Processed properties offered
        Set theColl = PropertyManager.PropertiesForPlayer(mTrade.PlayerMakingOffer)
        With Me.lstPropertiesOffered
            For row = 1 To .ListItems.Count
                If .ListItems(row).Selected Then
                    mTrade.PropertiesOffered.Add theColl.item(row)
                End If
            Next row
        End With
        
        ' Process properties wanted
        Set theColl = PropertyManager.PropertiesForPlayer(mTrade.PlayerReceivingOffer)
        With Me.lstPropertiesWanted
            For row = 1 To .ListItems.Count
                If .ListItems(row).Selected Then
                    mTrade.PropertiesWanted.Add theColl.item(row)
                End If
            Next row
        End With
    End If
    
    Run = mReturnValue
    Unload Me
End Function

Private Sub cboFromPlayer_LostFocus()
    buildPropertyOffered
    BuildReceiveList
End Sub

Private Sub cboToPlayer_LostFocus()
    buildPropertyWanted
    BuildReceiveList
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If DataIsOkay Then
        mReturnValue = True
        Me.Hide
    End If
End Sub

Private Function DataIsOkay() As Boolean
    ' Sanity checks - must have a FROM and a TO
    If Me.cboFromPlayer.ListIndex <= 0 Then
        MsgBox "An offering player must be selected", vbOKOnly + vbInformation, "Monopolie"
        Me.cboFromPlayer.SetFocus
        DataIsOkay = False
        Exit Function
    End If
    If Me.cboToPlayer.ListIndex <= 0 Then
        MsgBox "Please select a player to make this trade offer to", vbOKOnly + vbInformation, "Monopolie"
        Me.cboToPlayer.SetFocus
        DataIsOkay = False
        Exit Function
    End If
    If Me.cboToPlayer.ListIndex = Me.cboFromPlayer.ListIndex Then
        MsgBox "The player making the offer cannot be the same as the player the trade is offered to", vbOKOnly + vbInformation, "Monopolie"
        Me.cboFromPlayer.SetFocus
        DataIsOkay = False
        Exit Function
    End If
    If Val(Me.txtCashInvolved) < 0 Then
        MsgBox "The amount of Cash must be a postitive value"
        DataIsOkay = False
        Me.txtCashInvolved.SetFocus
        Exit Function
    End If
    
    '''' TO DO:  If any property we're asking for is developed,
    '''' reject this trade
    
    DataIsOkay = True
End Function

Private Sub cmdSelectOfferAll_Click()
    modUtils.AutoSelectListViewEntries Me.lstPropertiesOffered, True
End Sub

Private Sub cmdSelectOfferNone_Click()
    modUtils.AutoSelectListViewEntries Me.lstPropertiesOffered, False
End Sub

Private Sub cmdSelectReceivingAll_Click()
    modUtils.AutoSelectListViewEntries Me.lstPropertiesWanted, True
End Sub

Private Sub cmdSelectReceivingNone_Click()
    modUtils.AutoSelectListViewEntries Me.lstPropertiesWanted, False
End Sub

Private Sub txtCashInvolved_Validate(Cancel As Boolean)
    Cancel = InvalidAmount(Me.txtCashInvolved)
End Sub

Private Sub BuildPlayerList(theControl As ComboBox)
    Dim player As claIPlayer
    
    theControl.Clear
    theControl.AddItem "(None)"
    For Each player In PlayerManager
        theControl.AddItem player.Name & " (" & modUtils.PlayerTypeToWords(player.PlayerType) & ")"
    Next player
End Sub

Private Sub buildPropertyOffered()
    Dim player As claIPlayer
    Dim prop As claIProperty

    Me.lstPropertiesOffered.ListItems.Clear
    If Me.cboFromPlayer.ListIndex <= 0 Then Exit Sub
    Set player = PlayerManager.player(Me.cboFromPlayer.ListIndex)
    For Each prop In PropertyManager.PropertiesForPlayer(player)
        Me.lstPropertiesOffered.ListItems.Add Text:=prop.Name, SmallIcon:=PropertyManager.GroupToString(prop.Group)
    Next prop
    modUtils.AutoSelectListViewEntries Me.lstPropertiesOffered, False
End Sub

Private Sub buildPropertyWanted()
    Dim player As claIPlayer
    Dim prop As claIProperty

    Me.lstPropertiesWanted.ListItems.Clear
    If Me.cboToPlayer.ListIndex <= 0 Then Exit Sub
    Set player = PlayerManager.player(Me.cboToPlayer.ListIndex)
    For Each prop In PropertyManager.PropertiesForPlayer(player)
        Me.lstPropertiesWanted.ListItems.Add Text:=prop.Name, SmallIcon:=PropertyManager.GroupToString(prop.Group)
    Next prop
    modUtils.AutoSelectListViewEntries Me.lstPropertiesWanted, False
End Sub

Private Sub BuildReceiveList()
    Me.cboPlayerReceivingCash.Clear
    Me.cboPlayerReceivingCash.AddItem Me.cboFromPlayer
    Me.cboPlayerReceivingCash.AddItem Me.cboToPlayer

    If mTrade.ReceivingPlayerGetsCash Then
        Me.cboPlayerReceivingCash.ListIndex = 0
    Else
        Me.cboPlayerReceivingCash.ListIndex = 1
    End If
End Sub

' SM:  This can probably be moved into individual properies, or the property manager
