VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claStation
' Date      : 11/17/2003
' Author    : Simon M. Mitchell
' Purpose   : Station class
' 07/14/2004    SM:  Changed to use new Log file
' 03/03/2006    SM:  Improved DumpToLog info
'---------------------------------------------------------------------------------------

Option Explicit

Implements claIProperty

Private mName As String
Private mOwner As Integer
Private mRentIf1 As Long
Private mRentIf2 As Long
Private mRentIf3 As Long
Private mRentIf4 As Long
Private mBoardLocation As Long
Private mMortgageStatus As PropertyStatus
Private mPurchasePrice As Long
Private mMortgagePrice As Long

Public Property Get RentIf1() As Long
    RentIf1 = mRentIf1
End Property
Public Property Let RentIf1(RHS As Long)
    mRentIf1 = RHS
End Property
Public Property Get RentIf2() As Long
    RentIf1 = mRentIf2
End Property
Public Property Let RentIf2(RHS As Long)
    mRentIf2 = RHS
End Property
Public Property Get RentIf3() As Long
    RentIf1 = mRentIf3
End Property
Public Property Let RentIf3(RHS As Long)
    mRentIf3 = RHS
End Property
Public Property Get RentIf4() As Long
    RentIf1 = mRentIf4
End Property
Public Property Let RentIf4(RHS As Long)
    mRentIf4 = RHS
End Property

Private Property Let claIProperty_BaseRent(RHS As Long)
'' SM:  Rent for stations is calculated based on the
'' number of stations owned
End Property

Private Property Get claIProperty_BaseRent() As Long
'' TO DO:  Calculate rent based on how many stations
'' owned by the player
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
    Else
        Select Case PropertyManager.RailroadsOwnedByPlayer(mOwner)
        Case 1
            returnValue = mRentIf1
        Case 2
            returnValue = mRentIf2
        Case 3
            returnValue = mRentIf3
        Case Else
            returnValue = mRentIf4
        End Select
    End If
    claIProperty_CalculateRent = returnValue
End Function

Private Property Get claIProperty_CanImprove() As Boolean
    claIProperty_CanImprove = False
End Property

Private Property Let claIProperty_CostPerHotel(RHS As Long)
'' SM:  Not used by stations
End Property

Private Property Get claIProperty_CostPerHotel() As Long
    claIProperty_CostPerHotel = 0
End Property

Private Property Let claIProperty_CostPerHouse(RHS As Long)
'' SM:  Not used by stations
End Property

Private Property Get claIProperty_CostPerHouse() As Long
    claIProperty_CostPerHouse = 0
End Property

Private Property Get claIProperty_CurrentPropertyValue() As Long
    Dim returnValue As Long
    
    If mMortgageStatus = staMortgaged Then
        returnValue = 0
    Else
        returnValue = mMortgagePrice
    End If
    claIProperty_CurrentPropertyValue = returnValue
End Property

Private Sub claIProperty_DumpToLog()
    Dim logstring As String
    
    logstring = mName & " - STATION - "
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
'' SM:  Not used for stations
End Property

Private Property Get claIProperty_Group() As PropertyGroup
    claIProperty_Group = grpStation
End Property

Private Property Let claIProperty_Houses(RHS As Integer)
'' SM:  Drop - not used by stations
End Property

Private Property Get claIProperty_Houses() As Integer
    claIProperty_Houses = 0
End Property

Private Property Let claIProperty_ImprovedRent(improvement As PropertyImprovement, RHS As Long)
'' SM:  Not used by stations
End Property

Private Property Get claIProperty_ImprovedRent(improvement As PropertyImprovement) As Long
    claIProperty_ImprovedRent = 0
End Property

Private Property Let claIProperty_MonopolyRent(RHS As Long)
'' SM:  Not used by stations
End Property

Private Property Get claIProperty_MonopolyRent() As Long
    claIProperty_MonopolyRent = 0
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
    AppendToLog "Station " & claIProperty_Name & " reverted to Bank"
End Sub
