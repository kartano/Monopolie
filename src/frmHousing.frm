VERSION 5.00
Begin VB.Form frmHousing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Housing"
   ClientHeight    =   5055
   ClientLeft      =   1995
   ClientTop       =   3510
   ClientWidth     =   9645
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSellAll 
      Caption         =   "&Sell One House From Each"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdBuyAll 
      Caption         =   "&Buy One House For Each"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdSellHouse 
      Caption         =   "&Sell House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuyHouse 
      Caption         =   "&Buy House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSellHouse 
      Caption         =   "&Sell House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuyHouse 
      Caption         =   "&Buy House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSellHouse 
      Caption         =   "&Sell House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuyHouse 
      Caption         =   "&Buy House"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image imgProperty 
      Height          =   2970
      Index           =   1
      Left            =   120
      Picture         =   "frmHousing.frx":0000
      Top             =   120
      Width           =   2925
   End
   Begin VB.Image imgHousebw 
      Height          =   720
      Left            =   720
      Picture         =   "frmHousing.frx":1C70A
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgHouse 
      Height          =   720
      Left            =   0
      Picture         =   "frmHousing.frx":1D1B8
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image house3 
      Height          =   360
      Index           =   1
      Left            =   6720
      Picture         =   "frmHousing.frx":1DCAE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house3 
      Height          =   360
      Index           =   2
      Left            =   7200
      Picture         =   "frmHousing.frx":1E7A4
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house3 
      Height          =   360
      Index           =   3
      Left            =   7680
      Picture         =   "frmHousing.frx":1F29A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house3 
      Height          =   360
      Index           =   4
      Left            =   8160
      Picture         =   "frmHousing.frx":1FD90
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house3 
      Height          =   360
      Index           =   5
      Left            =   8640
      Picture         =   "frmHousing.frx":20886
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblCash 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblCashText 
      Caption         =   "Cash $"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblImportant 
      Alignment       =   2  'Center
      Caption         =   "Houses / Hotels sell for 1/2 Purchase Price"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   4200
      Width           =   5295
   End
   Begin VB.Image house2 
      Height          =   360
      Index           =   5
      Left            =   5400
      Picture         =   "frmHousing.frx":2137C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image house2 
      Height          =   360
      Index           =   4
      Left            =   4920
      Picture         =   "frmHousing.frx":21E72
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house2 
      Height          =   360
      Index           =   3
      Left            =   4440
      Picture         =   "frmHousing.frx":22968
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house2 
      Height          =   360
      Index           =   2
      Left            =   3960
      Picture         =   "frmHousing.frx":2345E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house2 
      Height          =   360
      Index           =   1
      Left            =   3480
      Picture         =   "frmHousing.frx":23F54
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house1 
      Height          =   360
      Index           =   5
      Left            =   2160
      Picture         =   "frmHousing.frx":24A4A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   735
   End
   Begin VB.Image house1 
      Height          =   360
      Index           =   4
      Left            =   1680
      Picture         =   "frmHousing.frx":25540
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house1 
      Height          =   360
      Index           =   3
      Left            =   1200
      Picture         =   "frmHousing.frx":26036
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house1 
      Height          =   360
      Index           =   2
      Left            =   720
      Picture         =   "frmHousing.frx":26B2C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image house1 
      Height          =   360
      Index           =   1
      Left            =   240
      Picture         =   "frmHousing.frx":27622
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgProperty 
      Height          =   2970
      Index           =   3
      Left            =   6600
      Picture         =   "frmHousing.frx":28118
      Top             =   120
      Width           =   2925
   End
   Begin VB.Image imgProperty 
      Height          =   2970
      Index           =   2
      Left            =   3360
      Picture         =   "frmHousing.frx":44822
      Top             =   120
      Width           =   2925
   End
End
Attribute VB_Name = "frmHousing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHousing
' Date      : 10/5/2003
' Author    : Unknown
' Purpose   : Allow the player to purchase/sell houses/hotels
' 02/07/2004    SM:  Fixed to use new PropertyManager
' 08/21/2004    SM:  BuyAll/SellAll disable correctly
'                    Now prevents purchases player can't afford
' 03/03/2006    SM:  Standardised code header, removed cruft & fixed indents
'---------------------------------------------------------------------------------------

Option Explicit

Public iLocation      As Integer ' The physical square on the board
Private Property1     As Integer
Private Property2     As Integer
Private Property3     As Integer
Private iPrice        As Long

Public Sub run(ByVal Location As Integer)
    iLocation = Location
    If iLocation = 1 Or iLocation = 3 Or iLocation = 37 Or iLocation = 39 Then
        SetProperty3Visible (False) ' for purples or blues
    Else
        SetProperty3Visible (True) 'added to so that all other properties three cards will show **Brian
    End If
    iPrice = PropertyManager.property(Location).CostPerHouse
    SetCash
    GetPropertyNumbers
    DisplayHouses
    RefreshButtons
    RefreshProperties
    Me.Show vbModal
    Unload Me
End Sub

Private Sub cmdBuyAll_Click()
    Dim TempPrice As Long

    TempPrice = IIf(imgProperty(3).Visible, iPrice * 3, iPrice * 2)
    
    '' SM:  May need to change this later.
    '' bResponse can be set to false if there are not
    '' enough houses etc.
    If CanAfford(TempPrice) Then
        PropertyManager.property(Property1).Houses = PropertyManager.property(Property1).Houses + 1
        PropertyManager.property(Property2).Houses = PropertyManager.property(Property2).Houses + 1
        If (imgProperty(3).Visible = True) Then       'three properties
            PropertyManager.property(Property3).Houses = PropertyManager.property(Property3).Houses + 1
        End If
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -TempPrice
        SetCash
        DisplayHouses
        RefreshButtons
    End If
End Sub

Private Sub cmdBuyHouse_Click(index As Integer)
    ' Note that for the Buy#_click() routines, the player1money would have
    ' to be replaced with a variable for the current player for network
    ' games
    ' @author James
  
    'User clicked the first Buy House button
    'First, make sure the player can buy the house, then do it!
    
    If CanAfford(iPrice) Then
        Select Case index
        Case 0
            PropertyManager.property(Property1).Houses = PropertyManager.property(Property1).Houses + 1
        Case 1
            PropertyManager.property(Property2).Houses = PropertyManager.property(Property2).Houses + 1
        Case 2
            PropertyManager.property(Property3).Houses = PropertyManager.property(Property3).Houses + 1
        End Select
        
        PlayerManager.currentPlayer.SimpleTransaction Nothing, -iPrice
        ''SimpleTransaction PlayerManager.currentPlayer.Number, -iPrice, True
        SetCash
        DisplayHouses
        RefreshButtons
    End If
End Sub

Private Sub cmdDone_Click()
    frmMain.DisplayHouses
    Me.Hide
End Sub

Private Sub cmdSellAll_Click()
    ' Sells one house from each property
    ' @author James
    Dim bResponse As Boolean
    Dim TempPrice As Long

    PropertyManager.property(Property1).Houses = PropertyManager.property(Property1).Houses - 1
    PropertyManager.property(Property2).Houses = PropertyManager.property(Property2).Houses - 1
    If (imgProperty(3).Visible = True) Then       'three properties
        PropertyManager.property(Property3).Houses = PropertyManager.property(Property2).Houses - 1
        TempPrice = (iPrice / 2) * 3
    Else
        TempPrice = (iPrice / 2) * 2
    End If
    ''
    bResponse = True
    ''
    If bResponse Then
        PlayerManager.currentPlayer.SimpleTransaction Nothing, TempPrice
        ''SimpleTransaction PlayerManager.currentPlayer.Number, TempPrice, True
        SetCash
        DisplayHouses
        RefreshButtons
    End If
End Sub

Private Sub cmdSellHouse_Click(index As Integer)
    Dim TempPrice As Long
    Dim bResponse As Boolean

    Select Case index
    Case 0
        PropertyManager.property(Property1).Houses = PropertyManager.property(Property1).Houses - 1
    Case 1
        PropertyManager.property(Property2).Houses = PropertyManager.property(Property2).Houses - 1
    Case 2
        PropertyManager.property(Property3).Houses = PropertyManager.property(Property3).Houses - 1
    End Select
    
    bResponse = True
    
    TempPrice = (iPrice / 2)
    If bResponse Then
        PlayerManager.currentPlayer.SimpleTransaction Nothing, TempPrice
        SetCash
        DisplayHouses
        RefreshButtons
    End If
End Sub

Private Sub DisplayHouses()
    Dim i      As Integer
    Dim iCount As Integer

    i = 1
    'Set all houses to b&w, then set the correct number to color
    For i = 1 To 5
        house1(i).Picture = imgHousebw.Picture
        house2(i).Picture = imgHousebw.Picture
        house3(i).Picture = imgHousebw.Picture
    Next i
    iCount = PropertyManager.property(Property1).Houses
    For i = 1 To iCount
        If iCount < 6 Then
            house1(i).Picture = imgHouse.Picture
        End If
    Next i
    iCount = PropertyManager.property(Property2).Houses
    For i = 1 To iCount
        If iCount < 6 Then
            house2(i).Picture = imgHouse.Picture
        End If
    Next i
    If (imgProperty(3).Visible = True) Then
        iCount = PropertyManager.property(Property3).Houses
        For i = 1 To iCount
            If iCount < 6 Then
                house3(i).Picture = imgHouse.Picture
            End If
        Next i
    End If
End Sub

'' SM:  This function looks like a complete duplicate of
'' the one in property manager.
'' Was there a reason this was done seperately?
Private Sub GetPropertyNumbers()
    If iLocation = 1 Or iLocation = 3 Then                    'Purples
        Property1 = 1
        Property2 = 3
    ElseIf iLocation > 5 And iLocation < 10 Then    'Aquas'NOT ILOCATION...
        Property1 = 6
        Property2 = 8
        Property3 = 9
    ElseIf iLocation > 10 And iLocation < 15 Then 'Pinks'NOT ILOCATION...
        Property1 = 11
        Property2 = 13
        Property3 = 14
    ElseIf iLocation > 15 And iLocation < 20 Then 'Oranges'NOT ILOCATION...
        Property1 = 16
        Property2 = 18
        Property3 = 19
    ElseIf iLocation > 20 And iLocation < 25 Then 'Reds'NOT ILOCATION...
        Property1 = 21
        Property2 = 23
        Property3 = 24
    ElseIf iLocation > 25 And iLocation < 30 Then 'Yellows'NOT ILOCATION...
        Property1 = 26
        Property2 = 27
        Property3 = 29
    ElseIf iLocation > 30 And iLocation < 35 Then 'Greens'NOT ILOCATION...
        Property1 = 31
        Property2 = 32
        Property3 = 34
    ElseIf iLocation = 37 Or iLocation = 39 Then                  'Blues'NOT ILOCATION...
        Property1 = 37
        Property2 = 39
    End If
End Sub

Private Sub RefreshButtons()
    Dim PropStatus1 As Integer
    Dim PropStatus2 As Integer
    Dim PropStatus3 As Integer

    ' Get current house count for each of the properties.
    ' There could be 2 or 3 in the set.
    PropStatus1 = PropertyManager.property(Property1).Houses
    PropStatus2 = PropertyManager.property(Property2).Houses
    If Property3 <> 0 Then PropStatus3 = PropertyManager.property(Property3).Houses
    
    cmdBuyHouse(0).Enabled = False
    cmdBuyHouse(1).Enabled = False
    cmdBuyHouse(2).Enabled = False
    
    cmdSellHouse(0).Enabled = False
    cmdSellHouse(1).Enabled = False
    cmdSellHouse(2).Enabled = False
    
    '''' SM:  This code only allows houses to built in a
    '''' particular pattern - where the user can purchase
    '''' an uneven number of houses, they should be able
    '''' to choose which property in the set gets the "odd"
    '''' house.
    If PropStatus1 = PropStatus2 Then
        If PropStatus2 = PropStatus3 Then
            cmdBuyHouse(0).Enabled = PropStatus1 < 5
            cmdBuyHouse(1).Enabled = PropStatus1 < 5
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = PropStatus1 < 5
            cmdSellHouse(0).Enabled = (0 > PropStatus1 And PropStatus1 < 6)
            cmdSellHouse(1).Enabled = (0 > PropStatus2 And PropStatus2 < 6)
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = (0 > PropStatus3 And PropStatus3 < 6)
        End If
    End If
    If PropStatus1 < PropStatus2 Then
        If PropStatus2 = PropStatus3 Then
            cmdBuyHouse(0).Enabled = True
            cmdBuyHouse(1).Enabled = False
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = False
            cmdSellHouse(0).Enabled = False
            cmdSellHouse(1).Enabled = True
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = True
        End If
    End If
    If PropStatus1 > PropStatus2 Then
        If PropStatus2 = PropStatus3 Then
            cmdBuyHouse(0).Enabled = False
            cmdBuyHouse(1).Enabled = True
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = True
            cmdSellHouse(0).Enabled = True
            cmdSellHouse(1).Enabled = False
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = False
        End If
    End If
    If PropStatus1 = PropStatus2 Then
        If PropStatus2 < PropStatus3 Then
            cmdBuyHouse(0).Enabled = True
            cmdBuyHouse(1).Enabled = True
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = False
            cmdSellHouse(0).Enabled = False
            cmdSellHouse(1).Enabled = False
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = True
        End If
    End If
    If PropStatus1 = PropStatus2 Then
        If PropStatus2 > PropStatus3 Then
            cmdBuyHouse(0).Enabled = False
            cmdBuyHouse(1).Enabled = False
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = True
            cmdSellHouse(0).Enabled = True
            cmdSellHouse(1).Enabled = True
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = False
        End If
    End If
    If PropStatus1 < PropStatus2 Then
        If PropStatus2 > PropStatus3 Then
            cmdBuyHouse(0).Enabled = True
            cmdBuyHouse(1).Enabled = False
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = True
            cmdSellHouse(0).Enabled = False
            cmdSellHouse(1).Enabled = True
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = False
        End If
    End If
    If PropStatus1 > PropStatus2 Then
        If PropStatus2 < PropStatus3 Then
            cmdBuyHouse(0).Enabled = False
            cmdBuyHouse(1).Enabled = True
            If Property3 <> 0 Then cmdBuyHouse(2).Enabled = False
            cmdSellHouse(0).Enabled = True
            cmdSellHouse(1).Enabled = False
            If Property3 <> 0 Then cmdSellHouse(2).Enabled = True
        End If
    End If
    'disable buttons if not enough money
    If iPrice > PlayerManager.currentPlayer.Money Then
        cmdBuyHouse(0).Enabled = False
        cmdBuyHouse(1).Enabled = False
        If Property3 <> 0 Then cmdBuyHouse(2).Enabled = False
        cmdSellHouse(0).Enabled = True
        cmdSellHouse(1).Enabled = True
        If Property3 <> 0 Then cmdSellHouse(2).Enabled = True
    End If
    cmdBuyHouse(0).Caption = "Buy House"
    cmdBuyHouse(1).Caption = "Buy House"
    If Property3 <> 0 Then cmdBuyHouse(2).Caption = "Buy House"

    ' Enable/Disable "Buy all" if all houses already have hotels
    If Property3 <> 0 Then
        Me.cmdBuyAll.Enabled = PropStatus1 <> 5 Or PropStatus2 <> 5
    Else
        Me.cmdBuyAll.Enabled = PropStatus1 <> 5 Or PropStatus2 <> 5 Or PropStatus3 <> 5
    End If

    ' ENable/Disable "Sell all" if no houses on any property
    If Property3 <> 0 Then
        Me.cmdSellAll.Enabled = PropStatus1 > 0 Or PropStatus2 > 0
    Else
        Me.cmdSellAll.Enabled = PropStatus1 > 0 Or PropStatus2 > 0 Or PropStatus3 > 0
    End If
End Sub

Private Sub RefreshProperties()
    Select Case Property1
    Case 1
        imgProperty(1).Picture = frmProperty.P(1).Picture
        imgProperty(2).Picture = frmProperty.P(3).Picture
    Case 6
        imgProperty(1).Picture = frmProperty.P(6).Picture
        imgProperty(2).Picture = frmProperty.P(8).Picture
        imgProperty(3).Picture = frmProperty.P(9).Picture
    Case 11
        imgProperty(1).Picture = frmProperty.P(11).Picture
        imgProperty(2).Picture = frmProperty.P(13).Picture
        imgProperty(3).Picture = frmProperty.P(14).Picture
    Case 16
        imgProperty(1).Picture = frmProperty.P(16).Picture
        imgProperty(2).Picture = frmProperty.P(18).Picture
        imgProperty(3).Picture = frmProperty.P(19).Picture
    Case 21
        imgProperty(1).Picture = frmProperty.P(21).Picture
        imgProperty(2).Picture = frmProperty.P(23).Picture
        imgProperty(3).Picture = frmProperty.P(24).Picture
    Case 26
        imgProperty(1).Picture = frmProperty.P(26).Picture
        imgProperty(2).Picture = frmProperty.P(27).Picture
        imgProperty(3).Picture = frmProperty.P(29).Picture
    Case 31
        imgProperty(1).Picture = frmProperty.P(31).Picture
        imgProperty(2).Picture = frmProperty.P(32).Picture
        imgProperty(3).Picture = frmProperty.P(34).Picture
    Case 37
        imgProperty(1).Picture = frmProperty.P(37).Picture
        imgProperty(2).Picture = frmProperty.P(39).Picture
    End Select
End Sub

Private Sub SetCash()
    lblCash.Caption = PlayerManager.currentPlayer.Money
End Sub

Private Sub SetProperty3Visible(ByVal Choice As Boolean)
    Dim i As Integer

    For i = 1 To 5
        house3(i).Visible = Choice
    Next i
    imgProperty(3).Visible = Choice
    cmdBuyHouse(2).Visible = Choice
    cmdSellHouse(2).Visible = Choice
End Sub

Private Function CanAfford(cost As Long) As Boolean
    Dim bResult As Boolean
    
    bResult = (PlayerManager.currentPlayer.Money >= cost)
    If Not bResult Then SelfClosingMsgbox "This would cost $" & cost & " but you only have $" & PlayerManager.currentPlayer.Money & " on hand!", vbOKOnly + vbInformation, "Not enough cash"
    CanAfford = bResult
End Function
