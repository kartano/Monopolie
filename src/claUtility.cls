VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claUtility
' Date      : 11/17/2003
' Author    : Simon M. Mitchell
' Purpose   : Utility class
' 07/14/2004    SM:  Changed to use new Log file
' 03/03/2006    SM:  Improvements to stuff dumped to the log
'---------------------------------------------------------------------------------------

Option Explicit

Implements claIProperty

Private mName As String
Private mOwner As Integer
Private mBaseRent As Long
Private mBoardLocation As Integer
Private mMortgageStatus As PropertyStatus
Private mPurchasePrice As Long
Private mMortgagePrice As Long
Private mMultiplierIf1Utility As Integer
Private mMultiplierIf2Utility As Integer

Public Property Let Multiplier1Utility(RHS As Integer)
    mMultiplierIf1Utility = RHS
End Property
Public Property Get Multiplier1Utility() As Integer
    Multiplier1Utility = mMultiplierIf1Utility
End Property
Public Property Let Multiplier2Utility(RHS As Integer)
    mMultiplierIf2Utility = RHS
End Property
Public Property Get Multiplier2Utility() As Integer
    Multiplier2Utility = mMultiplierIf2Utility
End Property

Private Property Let claIProperty_BaseRent(RHS As Long)
    '' SM:  Not used by utilities.  Rental is calculated
    '' on the dice roll and the number of utilities owned
End Property

Private Property Get claIProperty_BaseRent() As Long
    '' TO DO:  Calculate the rent based on the owner
    '' of this utility.  If they own both, multiplier is
    '' 10.  If they own one, multiplier is 4.
    claIProperty_BaseRent = 0
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
        If PropertyManager.PlayerOwnsMonopoly_ByPropNumber(mOwner, mBoardLocation) Then
            returnValue = (DiceRoll1 + DiceRoll2) * mMultiplierIf1Utility
        Else
            returnValue = (DiceRoll1 + DiceRoll2) * mMultiplierIf2Utility
        End If
    End If
    claIProperty_CalculateRent = returnValue
End Function

Private Property Get claIProperty_CanImprove() As Boolean
    claIProperty_CanImprove = False
End Property

Private Property Let claIProperty_CostPerHotel(RHS As Long)
'' SM:  Drop.  Not used for utility classes
End Property

Private Property Get claIProperty_CostPerHotel() As Long
    claIProperty_CostPerHotel = 0
End Property

Private Property Let claIProperty_CostPerHouse(RHS As Long)
'' SM:  Drop.  Not used for utility classes
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
    
    logstring = mName & " - UTILITY - "
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
'' SM:  Drop - not used by utilities
End Property

Private Property Get claIProperty_Group() As PropertyGroup
    claIProperty_Group = grpUtility
End Property

Private Property Let claIProperty_Houses(RHS As Integer)
'' SM:  Drop - not used by utilities
End Property

Private Property Get claIProperty_Houses() As Integer
    claIProperty_Houses = 0
End Property

Private Property Let claIProperty_ImprovedRent(improvement As PropertyImprovement, RHS As Long)
'' SM:  Drop - not used by utilities
End Property

Private Property Get claIProperty_ImprovedRent(improvement As PropertyImprovement) As Long
    claIProperty_ImprovedRent = 0
End Property

Private Property Let claIProperty_MonopolyRent(RHS As Long)
'' SM:  Drop - not used by utilities
End Property

Private Property Get claIProperty_MonopolyRent() As Long
    '' TO DO:  Calculate the rent based on the owner
    '' of this utility.  If they own both, multiplier is
    '' 10.  If they own one, multiplier is 4.
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
    AppendToLog "Utility " & claIProperty_Name & " reverted to Bank"
End Sub
