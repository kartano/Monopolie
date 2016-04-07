VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claRealEstate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claRealEstate
' Date      : 11/17/2003
' Author    : Simon M. Mitchell
' Purpose   : Real Estate class
' 07/14/2004    SM:  Changed to use new Log file
' 03/03/2006    SM:  Changed the DumpToLog information
'---------------------------------------------------------------------------------------

Option Explicit

Implements claIProperty

Private mName As String
Private mBaseRent As Long
Private mRent1House As Long
Private mRent2Houses As Long
Private mRent3Houses As Long
Private mRent4Houses As Long
Private mRentHotel As Long
Private mRentMonopoly As Long
Private mBoardLocation As Integer
Private mGroup As PropertyGroup
Private mHouses As Integer
Private mMortgagePrice As Long
Private mPurchasePrice As Long
Private mMortgageStatus As PropertyStatus
Private mOwner As Integer
Private mCostPerHotel As Long
Private mCostPerHouse As Long

Private Property Let claIProperty_BaseRent(RHS As Long)
    mBaseRent = RHS
End Property
Private Property Get claIProperty_BaseRent() As Long
    claIProperty_BaseRent = mBaseRent
End Property

Private Property Let claIProperty_BoardLocation(RHS As Integer)
    mBoardLocation = RHS
End Property

Private Property Get claIProperty_BoardLocation() As Integer
    claIProperty_BoardLocation = mBoardLocation
End Property

Private Function claIProperty_CalculateRent() As Long
    Dim returnValue As Long
    
    If mMortgageStatus = staMortgaged Then
        returnValue = 0
    ElseIf PropertyManager.PlayerOwnsMonopoly_ByPropNumber(mOwner, mBoardLocation) Then
        Select Case mHouses
        Case 0
            returnValue = mBaseRent * 2
        Case 1
            returnValue = mRent1House
        Case 2
            returnValue = mRent2Houses
        Case 3
            returnValue = mRent3Houses
        Case 4
            returnValue = mRent4Houses
        Case 5
            returnValue = mRentHotel
        End Select
    Else
        returnValue = mBaseRent
    End If
    claIProperty_CalculateRent = returnValue
End Function

Private Property Get claIProperty_CanImprove() As Boolean
    claIProperty_CanImprove = True
End Property

Private Property Let claIProperty_CostPerHotel(RHS As Long)
    mCostPerHotel = RHS
End Property

Private Property Get claIProperty_CostPerHotel() As Long
    claIProperty_CostPerHotel = mCostPerHotel
End Property

Private Property Let claIProperty_CostPerHouse(RHS As Long)
    mCostPerHouse = RHS
End Property

Private Property Get claIProperty_CostPerHouse() As Long
    claIProperty_CostPerHouse = mCostPerHouse
End Property

Private Property Get claIProperty_CurrentPropertyValue() As Long
    Dim returnValue As Long
    
    If mMortgageStatus = staMortgaged Then
        returnValue = 0
    Else
        If mHouses = 0 Then
            returnValue = mMortgagePrice
        Else
            returnValue = mMortgagePrice + ((mHouses * mCostPerHouse) / 2)
        End If
    End If
    claIProperty_CurrentPropertyValue = returnValue
End Property

Private Sub claIProperty_DumpToLog()
    Dim logstring As String
    
    logstring = mName & " - PROPERTY - " & mHouses & " houses - "
    If mMortgageStatus = staMortgaged Then
        logstring = logstring & " MORTGAGED"
    Else
        logstring = logstring & " unmortgaged"
    End If
    If mOwner = 0 Then
        logstring = logstring & " - Unowned"
    Else
        With PlayerManager.player(mOwner)
            logstring = logstring & " - " & .Name
        End With
    End If
    AppendToLog logstring
End Sub

Private Property Let claIProperty_Group(RHS As PropertyGroup)
    mGroup = RHS
End Property

Private Property Get claIProperty_Group() As PropertyGroup
    claIProperty_Group = mGroup
End Property

Private Property Let claIProperty_Houses(RHS As Integer)
    mHouses = RHS
End Property

Private Property Get claIProperty_Houses() As Integer
    claIProperty_Houses = mHouses
End Property

Private Property Let claIProperty_ImprovedRent(improvement As PropertyImprovement, RHS As Long)
    Select Case improvement
    Case PropertyImprovement.imp1House
        mRent1House = RHS
    Case PropertyImprovement.imp2House
        mRent2Houses = RHS
    Case PropertyImprovement.imp3House
        mRent3Houses = RHS
    Case PropertyImprovement.imp4House
        mRent4Houses = RHS
    Case PropertyImprovement.imp1Hotel
        mRentHotel = RHS
    End Select
End Property

Private Property Get claIProperty_ImprovedRent(improvement As PropertyImprovement) As Long
    Select Case improvement
    Case PropertyImprovement.imp1House
        claIProperty_ImprovedRent = mRent1House
    Case PropertyImprovement.imp2House
        claIProperty_ImprovedRent = mRent2Houses
    Case PropertyImprovement.imp3House
        claIProperty_ImprovedRent = mRent3Houses
    Case PropertyImprovement.imp4House
        claIProperty_ImprovedRent = mRent4Houses
    Case PropertyImprovement.imp1Hotel
        claIProperty_ImprovedRent = mRentHotel
    End Select
End Property

Private Property Get claIProperty_MonopolyRent() As Long
    claIProperty_MonopolyRent = mRentMonopoly
End Property

Private Property Let claIProperty_MonopolyRent(RHS As Long)
    mRentMonopoly = RHS
End Property

Private Property Let claIProperty_MortgagePrice(RHS As Long)
    mMortgagePrice = RHS
End Property

Private Property Get claIProperty_MortgagePrice() As Long
    claIProperty_MortgagePrice = mMortgagePrice
End Property

Private Property Let claIProperty_MortgageStatus(RHS As PropertyStatus)
    mMortgageStatus = RHS
End Property

Private Property Get claIProperty_MortgageStatus() As PropertyStatus
    claIProperty_MortgageStatus = mMortgageStatus
End Property

Private Property Let claIProperty_Name(RHS As String)
    mName = RHS
End Property

Private Property Get claIProperty_Name() As String
    claIProperty_Name = mName
End Property

Private Property Let claIProperty_Owner(RHS As Integer)
    mOwner = RHS
End Property

Private Property Get claIProperty_Owner() As Integer
    claIProperty_Owner = mOwner
End Property

Private Property Let claIProperty_PurchasePrice(RHS As Long)
    mPurchasePrice = RHS
End Property

Private Property Get claIProperty_PurchasePrice() As Long
    claIProperty_PurchasePrice = mPurchasePrice
End Property

Private Sub claIProperty_RevertToBank()
    mOwner = 0
    mMortgageStatus = staUnmortgaged
    mHouses = 0
    AppendToLog "Real Estate " & claIProperty_Name & " reverted to Bank"
End Sub

