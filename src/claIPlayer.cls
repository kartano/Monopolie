VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claIPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claIPlayer
' Date      : 11/13/2003
' Author    : Simon M. Mitchell
' Purpose   : Player Interface class
' 07/09/2004    SM:  Added IsBankrupt
'                    Added BankruptThisPlayer
'---------------------------------------------------------------------------------------

Option Explicit

' This is an abstract class denoting common player
' attributes for both Humans and Computers.

Public Property Get Name() As String
    ''
End Property
Public Property Let Name(RHS As String)
    ''
End Property

'' TO DO:  This could be removed using a new player
'' control class!
Public Property Get Number() As Integer
    ''
End Property
Public Property Let Number(RHS As Integer)
    ''
End Property
'' END TO DO

Public Function AcceptTrade(theTrade As claTrade) As Boolean
    ''
End Function
Public Property Get Status() As PlayerStatus
    ''
End Property
Public Property Let Status(RHS As PlayerStatus)
    ''
End Property
Public Property Get PlayerType() As PlayerType
    ''
End Property
Public Function SavePlayer() As String
    ''
End Function
Public Sub AddFrame()
    ''
End Sub
Public Sub ChangeMoney(theAmount As Long)
    ''
End Sub
Public Property Get Money() As Long
    ''
End Property
Public Property Let Money(RHS As Long)
''
End Property
Public Property Get TotalValue() As Long
    ''
End Property
Public Property Get OOJailChance() As Boolean
    ''
End Property
Public Property Let OOJailChance(RHS As Boolean)
    ''
End Property
Public Property Get OOJailComm() As Boolean
    ''
End Property
Public Property Let OOJailComm(RHS As Boolean)
    ''
End Property
Public Property Get OOJailRoll() As Integer
    ''
End Property
Public Property Let OOJailRoll(RHS As Integer)
    ''
End Property
Public Property Get Token() As Token
''
End Property
Public Property Let Token(RHS As Token)
''
End Property
Public Sub SimpleTransaction(theOtherPlayer As claIPlayer, requiredAmount As Long)
''
End Sub

'' SM:  Sorry!  Couldn't think of a better name for this!!
Public Function DoSomethingIfInJail() As Boolean
''
End Function

Public Property Get DoublesCount() As Integer
''
End Property
Public Property Let DoublesCount(RHS As Integer)
''
End Property

Public Property Get JailCount() As Integer
''
End Property
Public Property Let JailCount(RHS As Integer)
''
End Property

Public Property Get InJail() As Boolean
''
End Property
Public Property Let InJail(RHS As Boolean)
''
End Property

Public Property Get Location() As Integer
''
End Property

Public Property Let Location(RHS As Integer)
''
End Property

Public Sub AddDoubles()
''
End Sub
Public Property Get Speed() As PlayerSpeed
''
End Property
Public Property Let Speed(RHS As PlayerSpeed)
''
End Property

' Properties specific to CPU players
Public Property Get Difficulty() As PlayerDifficulty
''
End Property
Public Property Let Difficulty(RHS As PlayerDifficulty)
''
End Property

Public Sub PropertyLand(Location As Integer, Optional multiplier = 1)
''
End Sub

Public Property Get SideBar() As cExplorerBar
''
End Property

Public Sub BankruptThisPlayer(theBankruptor As claIPlayer)
''
End Sub

Public Property Get IsBankrupt() As Boolean
''
End Property
Public Property Let IsBankrupt(RHS As Boolean)
''
End Property
