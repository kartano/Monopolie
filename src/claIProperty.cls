VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claIProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claIProperty
' Date      : 11/17/2003
' Author    : Simon M. Mitchell
' Purpose   : Property Interface class
'---------------------------------------------------------------------------------------

Option Explicit

' SM:  Groups to which properties can belong
Public Enum PropertyGroup
    grpPurple = 1
    grpAqua = 2
    grpPink = 3
    grpOrange = 4
    grpRed = 5
    grpYellow = 6
    grpGreen = 7
    grpBlue = 8
    ' Special cases
    grpStation = 100
    grpUtility = 101
End Enum

' Mortgage Status
Public Enum PropertyStatus
    staUnmortgaged = 0
    staMortgaged = 1
End Enum

Public Enum PropertyImprovement
    imp1House = 1
    imp2House = 2
    imp3House = 3
    imp4House = 4
    imp1Hotel = 5
End Enum

'' SM:  This is an abtract class used as the interface
'' for all properties - stations, utilities and the
'' big fellas.

Public Property Get Name() As String
''
End Property
Public Property Let Name(RHS As String)
''
End Property

Public Property Get Group() As PropertyGroup
''
End Property
Public Property Let Group(RHS As PropertyGroup)
''
End Property

Public Property Get PurchasePrice() As Long
''
End Property
Public Property Let PurchasePrice(RHS As Long)
''
End Property

Public Property Get MortgagePrice() As Long
''
End Property
Public Property Let MortgagePrice(RHS As Long)
''
End Property

' SM:  Base level rent
Public Property Get BaseRent() As Long
''
End Property
Public Property Let BaseRent(RHS As Long)
''
End Property
' SM:  Rent if unimproved but monopoly owned
Public Property Get MonopolyRent() As Long
''
End Property
Public Property Let MonopolyRent(RHS As Long)
''
End Property
'' SM:  etc etc
Public Property Get ImprovedRent(improvement As PropertyImprovement) As Long
''
End Property
Public Property Let ImprovedRent(improvement As PropertyImprovement, RHS As Long)
''
End Property

' SM:  Houses.  0 = undeveloped, 5 = hotel
' Utility/station subclasses should always return 0 for this
Public Property Get Houses() As Integer
''
End Property
Public Property Let Houses(RHS As Integer)
''
End Property

Public Property Get MortgageStatus() As PropertyStatus
''
End Property
Public Property Let MortgageStatus(RHS As PropertyStatus)
''
End Property

' SM:  Property owner.  Will be 0 if bank
Public Property Get Owner() As Integer
''
End Property
Public Property Let Owner(RHS As Integer)
''
End Property

Public Property Get BoardLocation() As Integer
''
End Property
Public Property Let BoardLocation(RHS As Integer)
''
End Property

Public Property Get CostPerHouse() As Long
''
End Property
Public Property Let CostPerHouse(RHS As Long)
''
End Property

Public Property Get CostPerHotel() As Long
''
End Property
Public Property Let CostPerHotel(RHS As Long)
''
End Property

Public Sub RevertToBank()
''
End Sub

Public Property Get CanImprove() As Boolean
''
End Property

Public Function CalculateRent() As Long
''
End Function

Public Property Get CurrentPropertyValue() As Long
''
End Property

Public Sub DumpToLog()
''
End Sub
