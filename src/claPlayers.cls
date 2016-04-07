VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claPlayers
' Date      : 11/14/2003
' Author    : Simon M. Mitchell
' Purpose   : Player manager class
' 07/02/2004    SM:  Fix to NewENum to prevent error 91s on new collections
' 07/09/2004    SM:  Added CheckForVictory
'---------------------------------------------------------------------------------------

' This class is used to control the players collections.
' If will merge in functionality from the PlayerManager
' and from the existing claPlayer class.

Option Explicit

Private mPlayers As Collection

Private mCurrentPlayerNumber As Integer

Public Enum PlayerType
    Human = 0
    Computer = 1
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Human, Computer
#End If

Public Enum PlayerDifficulty
    Easy = 0
    Medium = 1
    Hard = 2
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Easy, Medium, Hard
#End If

Public Enum PlayerSpeed
    Slow = 0
    Normal = 1
    Fast = 2
    SuperFast = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Slow, Normal, Fast, SuperFast
#End If

'' TO DO:  Include a status for "Active but in jail"
Public Enum PlayerStatus
    Bankrupt = 0
    Active = 1
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Bankrupt, Active
#End If
Public Enum Token
    Cannon = 0
    Car = 1
    Dog = 2
    Hat = 3
    Horse = 4
    Iron = 5
    Moneybag = 6
    Ship = 7
    Shoe = 8
    Thimble = 9
    Wheelbarrow = 10
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Cannon, Car, Dog, Hat, Horse, Iron, Moneybag, Ship, Shoe, Thimble, Wheelbarrow
#End If

Private Sub Class_Initialize()
    mCurrentPlayerNumber = 1
End Sub

' Load collection on demand
Private Property Get Players() As Collection
    If mPlayers Is Nothing Then Set mPlayers = New Collection
    Set Players = mPlayers
End Property

Public Property Get Count() As Integer
    Count = Players.Count
End Property

Public Sub Add(newPlayer As claIPlayer)
    Players.Add newPlayer
End Sub

' This allows us to use "for each" on this class
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    If mPlayers Is Nothing Then Set mPlayers = New Collection
    Set NewEnum = mPlayers.[_NewEnum]
End Function

Public Property Get player(index As Integer) As claIPlayer
    Set player = Players(index)
End Property

Public Property Get currentPlayer() As claIPlayer
    Set currentPlayer = player(mCurrentPlayerNumber)
End Property

' Basically:  Changes what the "current player" object is
Public Sub PassTheDice()
    Do
        If mCurrentPlayerNumber = Me.Count Then
            mCurrentPlayerNumber = 1
        Else
            mCurrentPlayerNumber = mCurrentPlayerNumber + 1
        End If
    Loop Until Me.currentPlayer.Status <> Bankrupt
    mdiMonopoly.StatusBar.Panels(1) = "Current Player: " & currentPlayer.Name
End Sub

Public Sub Delete(theIndex As Variant)
    AppendToLog vbTab & "Deleting player: " & theIndex
    Players.Remove theIndex
End Sub

' SM:  Give all CPU players a chance to trade between
' rolls.
Public Sub CheckForCPUTrades()
    Dim player As claIPlayer
    
    For Each player In Players
        If TypeOf player Is claCPUPlayer Then
            modCPU.CheckForTrade player
        End If
    Next player
End Sub

Private Sub Class_Terminate()
    AppendToLog vbTab & "Killing claPlayers"
    Set mPlayers = Nothing
End Sub

Public Function CheckForVictory() As Boolean
    Dim player As claIPlayer
    Dim bankruptCount As Integer
    Dim lastNonBankruptPlayer As claIPlayer
    
    bankruptCount = 0
    For Each player In mPlayers
        If player.IsBankrupt = True Then
            bankruptCount = bankruptCount + 1
        Else
            Set lastNonBankruptPlayer = player
        End If
    Next player
    If bankruptCount = mPlayers.Count - 1 Then
        SelfClosingMsgbox lastNonBankruptPlayer.Name & " IS THE VICTOR!!!!!!!", vbOKOnly + vbExclamation, "VICTORY!"
        CheckForVictory = True
    Else
        CheckForVictory = False
    End If
End Function
