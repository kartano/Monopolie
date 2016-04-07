VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claCards
' Date      : 4/14/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to control the deck of cards for Comm Chest and Chance
' 03/03/2006    SM:  Removed a lot of cruft, fixed indents
'---------------------------------------------------------------------------------------

Option Explicit

Private Cards(14)               As Integer
Private CardCount               As Integer
Private NextCard                As Integer
Public PlayerHasOOJailCard      As Boolean

Public Property Let CardNumber(ByVal Value As Integer)
    NextCard = Value
End Property

Private Sub Class_Initialize()
    ShuffleCards
    CardCount = 0
End Sub

Public Function GetCardNumber() As Integer
    GetCardNumber = Cards(CardCount)
    
    If DebugMode And NextCard > 0 Then GetCardNumber = NextCard
    CardCount = CardCount + 1
    If CardCount = 15 Then CardCount = 0
End Function

Private Sub ShuffleCards()
    Dim bCardOK   As Boolean
    Dim iFindCard As Integer
    Dim iTemp     As Integer
    Dim iShuffled As Integer

    For iShuffled = 0 To 14
        bCardOK = False
        Do While Not bCardOK
            bCardOK = True
            iTemp = Int((15 * Rnd) + 1)
            For iFindCard = 0 To iShuffled - 1
                If Cards(iFindCard) = iTemp Then
                    bCardOK = False
                    Exit For
                End If
            Next iFindCard
        Loop
        Cards(iShuffled) = iTemp
    Next iShuffled
End Sub
