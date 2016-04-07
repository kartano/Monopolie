VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : claTrade
' Date      : 11/13/2003
' Author    : Simon M. Mitchell
' Purpose   : Trade offer class
' 07/14/2004    SM:  Changed to use new Log file
'---------------------------------------------------------------------------------------

Option Explicit

Private mPlayerMakingOffer As claIPlayer
Private mPlayerReceivingOffer As claIPlayer

Private mPropertiesOffered As Collection
Private mPropertiesWanted As Collection

Public Cash As Integer
Public ReceivingPlayerGetsCash As Boolean
Public ChanceJailCardOffered As Boolean
Public CommChestJailCardOffered As Boolean
Public ChanceJailCardWanted As Boolean
Public CommChestJailCardWanted As Boolean

' I'm using this to spot bugs!
Private mLoaded As Boolean

' Always call this first.
' The trade class is used for human and CPU players.
Public Sub Load(PlayerMakingOffer As claIPlayer, PlayerReceivingOffer As claIPlayer)
    AppendToLog "claTrade.Load"
    If PlayerMakingOffer Is Nothing Then
        AppendToLog vbTab & "WARNING:  PlayerMakingOffer is NOTHING!"
    Else
        AppendToLog vbTab & "Player making offer: " & PlayerMakingOffer.Name & " (" & PlayerMakingOffer.Number & ")"
    End If
    If PlayerReceivingOffer Is Nothing Then
        AppendToLog vbTab & "WARNING:  PlayerReceivingOffer is NOTHING!"
    Else
        AppendToLog vbTab & "Player receiving offer: " & PlayerReceivingOffer.Name & " (" & PlayerReceivingOffer.Number & ")"
    End If
    
    Set mPlayerMakingOffer = PlayerMakingOffer
    Set mPlayerReceivingOffer = PlayerReceivingOffer
    mLoaded = True
End Sub

Public Property Get PropertiesOffered() As Collection
    If mPropertiesOffered Is Nothing Then Set mPropertiesOffered = New Collection
    Set PropertiesOffered = mPropertiesOffered
End Property
Public Property Get PropertiesWanted() As Collection
    If mPropertiesWanted Is Nothing Then Set mPropertiesWanted = New Collection
    Set PropertiesWanted = mPropertiesWanted
End Property

' Returns TRUE if user clicked OK, FALSE if not
Public Function Edit() As Boolean
    If Not mLoaded Then ImproperClassUse
    Edit = frmTrade.Run(Me)
End Function

Public Function TradeAccepted() As Boolean
    If Not mLoaded Then ImproperClassUse
    TradeAccepted = mPlayerReceivingOffer.AcceptTrade(Me)
End Function

Public Sub Execute()
    Dim prop As claIProperty
    
    If Not mLoaded Then ImproperClassUse
    
    AppendToLog "claTrade.Execute"
    
    ' Transfer cash
    AppendToLog vbTab & "Transferring cash " & Me.Cash
    If Me.ReceivingPlayerGetsCash Then
        PlayerReceivingOffer.SimpleTransaction mPlayerMakingOffer, Me.Cash
    Else
        PlayerReceivingOffer.SimpleTransaction mPlayerMakingOffer, -Me.Cash
    End If
    
    ' Transfer all properties
    AppendToLog vbTab & "Transferring properties Offered - " & mPropertiesOffered.Count & " properties"
    For Each prop In mPropertiesOffered
        prop.Owner = PlayerReceivingOffer.Number
    Next prop
    
    AppendToLog vbTab & "Transferring properties Wanted - " & mPropertiesWanted.Count & " properties"
    For Each prop In mPropertiesWanted
        prop.Owner = PlayerMakingOffer.Number
    Next prop
    
    ' Transfer get out of jail cards
    AppendToLog vbTab & "Transferring get out of Jail Frees"
    If ChanceJailCardOffered Then
        PlayerReceivingOffer.OOJailChance = True
        PlayerMakingOffer.OOJailChance = False
    ElseIf ChanceJailCardWanted Then
        PlayerReceivingOffer.OOJailChance = False
        PlayerMakingOffer.OOJailChance = True
    End If
    
    If CommChestJailCardOffered Then
        PlayerReceivingOffer.OOJailComm = True
        PlayerMakingOffer.OOJailComm = False
    ElseIf CommChestJailCardWanted Then
        PlayerReceivingOffer.OOJailComm = False
        PlayerMakingOffer.OOJailComm = True
    End If
    
    '' SM:  In network versions, the acceptance of a
    '' trade offer would need to be communicated to any
    '' clients to keep those games in synch.
End Sub

Private Sub Class_Initialize()
    mLoaded = False
End Sub

Private Sub ImproperClassUse()
    Dim errMsg As String
    
    ''
    
    errMsg = "PROGRAMMING ERROR - You are attempting to use a Trade object before you've properly loaded it!"
    AppendToLog errMsg
    MsgBox errMsg
    '' If you are here, you should check the call stack!
    '' Always call claTrade.Load before attempting to
    '' use this class.
    Stop
    ''
End Sub

Public Property Get PlayerMakingOffer() As claIPlayer
    If Not mLoaded Then ImproperClassUse
    Set PlayerMakingOffer = mPlayerMakingOffer
End Property

Public Property Set PlayerMakingOffer(RHS As claIPlayer)
    Set mPlayerMakingOffer = RHS
End Property

Public Property Get PlayerReceivingOffer() As claIPlayer
    If Not mLoaded Then ImproperClassUse
    Set PlayerReceivingOffer = mPlayerReceivingOffer
End Property

Public Property Set PlayerReceivingOffer(RHS As claIPlayer)
    Set mPlayerReceivingOffer = RHS
End Property

' Dumps contents of trade to the log file if active
Public Sub DumpToLog()
    Dim prop As claIProperty
    On Error Resume Next
    
    AppendToLog vbTab & "Trade object"
    AppendToLog vbTab & "Player making offer: " & mPlayerMakingOffer.Name
    AppendToLog vbTab & "Player receiving offer: " & mPlayerReceivingOffer.Name
    AppendToLog vbTab & "Properties Offered:"
    For Each prop In mPropertiesOffered
        prop.DumpToLog
    Next prop
    AppendToLog vbTab & "Properties Wanted:"
    For Each prop In mPropertiesWanted
        prop.DumpToLog
    Next prop
    AppendToLog vbTab & "Cash: " & Cash
    AppendToLog vbTab & "Receiving player gets cash: " & ReceivingPlayerGetsCash
    AppendToLog vbTab & "Chance GOOJF offered: " & ChanceJailCardOffered
    AppendToLog vbTab & "Comm chest GOOJF offered: " & CommChestJailCardOffered
    AppendToLog vbTab & "Chance GOOJF wanted: " & ChanceJailCardWanted
    AppendToLog vbTab & "Comm chest GOOJF wanted: " & CommChestJailCardWanted
End Sub

