VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewTrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trade Offered"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstOffered 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      View            =   2
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
   Begin VB.CommandButton cmdViewMyProperty 
      Caption         =   "View &my properties"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdViewTheirProperty 
      Caption         =   "&View <other players> properties"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "&Reject"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgThumbnails 
      Left            =   4560
      Top             =   3840
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
            Picture         =   "frmViewTrade.frx":0000
            Key             =   "Blue"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":0502
            Key             =   "Green"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":0A04
            Key             =   "Aqua"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":0F06
            Key             =   "Orange"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":1408
            Key             =   "Pink"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":190A
            Key             =   "Purple"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":1E0C
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":230E
            Key             =   "Utility"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":2810
            Key             =   "Red"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTrade.frx":2D12
            Key             =   "Yellow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstWanted 
      Height          =   2175
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      View            =   2
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
   Begin VB.Label Label1 
      Caption         =   "Do you accept this trade?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblPlayerWants 
      Caption         =   "In exchange for these properties:"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblPlayerOffers 
      Caption         =   "<Player Offers caption>"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblCashDetails 
      Caption         =   "<Details about cash appear here>"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label lblTitle 
      Caption         =   "<Title for this Trade appears here>"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmViewTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmViewTrade
' Date      : Unknown
' Author    : Simon Mitchell
' Purpose   : Review a trade
' 21/08/2004    SM:  Changed to use pretty image icons
'---------------------------------------------------------------------------------------

Option Explicit

Dim mTrade As claTrade
Dim mReturnValue As Boolean

' Returns TRUE if the player accepted this trade
Public Function Run(theTrade As claTrade) As Boolean
    mReturnValue = False
    Set mTrade = theTrade
    SetupControls
    Me.Show vbModal
    Run = mReturnValue
    Unload Me
End Function

Private Sub SetupControls()
    Dim cashDetails As String
    Dim prop As claIProperty
    
    With mTrade
        Me.cmdViewTheirProperty.Caption = "&View " & .PlayerMakingOffer.Name & " property"
        Me.lblTitle.Caption = .PlayerMakingOffer.Name & " wants to trade with " & .PlayerReceivingOffer.Name & "!"
        Me.lblPlayerOffers.Caption = .PlayerMakingOffer.Name & "offers this property:"
        If .Cash = 0 Then
            cashDetails = "No cash is involved in this trade."
        Else
            If .ReceivingPlayerGetsCash Then
                cashDetails = .PlayerMakingOffer.Name & " agrees to pay you " & FormatCash(.Cash) & "."
            Else
                cashDetails = "You would agree to pay " & .PlayerMakingOffer.Name & " the sum of " & FormatCash(.Cash) & "."
            End If
        End If
        If .CommChestJailCardOffered Then
            cashDetails = cashDetails & vbCrLf & .PlayerMakingOffer.Name & " will give you their Community Chest Get out of Jail Free card."
        ElseIf .CommChestJailCardWanted Then
            cashDetails = cashDetails & vbCrLf & "You would agree to give " & .PlayerMakingOffer.Name & " your Community Chest Get out of Jail Free card."
        End If
        If .ChanceJailCardOffered Then
            cashDetails = cashDetails & vbCrLf & .PlayerMakingOffer.Name & " will give you their Chance Get out of Jail Free card."
        ElseIf .ChanceJailCardWanted Then
            cashDetails = cashDetails & vbCrLf & "You would agree to give " & .PlayerMakingOffer.Name & " your Chance Get out of Jail Free card."
        End If
        Me.lblCashDetails.Caption = cashDetails
    End With
    
    Me.lstOffered.ListItems.Clear
    For Each prop In mTrade.PropertiesOffered
        lstOffered.ListItems.Add Text:=prop.Name, SmallIcon:=PropertyManager.GroupToString(prop.Group)
    Next prop
    
    Me.lstWanted.ListItems.Clear
    For Each prop In mTrade.PropertiesWanted
        lstWanted.ListItems.Add Text:=prop.Name, SmallIcon:=PropertyManager.GroupToString(prop.Group)
    Next prop
End Sub

Private Sub cmdAccept_Click()
    mReturnValue = True
    Me.Hide
End Sub

Private Sub cmdReject_Click()
    Me.Hide
End Sub

Private Sub cmdViewMyProperty_Click()
    frmCards.Run mTrade.PlayerReceivingOffer.Number
End Sub

Private Sub cmdViewTheirProperty_Click()
    frmCards.Run mTrade.PlayerMakingOffer.Number
End Sub

Private Sub lstOffered_DblClick()
    ''
    MsgBox "TO DO:  Display the selected property"
    ''
End Sub

Private Sub lstWanted_DblClick()
    ''
    MsgBox "TO DO:  Display the select property"
    ''
End Sub
