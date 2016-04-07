Attribute VB_Name = "modPlayerUtils"
'---------------------------------------------------------------------------------------
' Module    : modPlayerUtils
' Date      : 13/11/2003
' Author    : Simon M. Mitchell
' Purpose   : Generic player utilities
' 05/18/2005    SM:  Added SendPlayerToJail
'---------------------------------------------------------------------------------------

Option Explicit

' Calculate the total value of all properties for a player
Public Function CalculateTotalValue(thePlayer As claIPlayer) As Long
    Dim returnValue As Long
    Dim boardSquare As Integer
    Dim propertiesOwned As Collection
    Dim prop As claIProperty
    
    returnValue = thePlayer.Money
    Set propertiesOwned = PropertyManager.PropertiesForPlayer(thePlayer)
    For Each prop In propertiesOwned
        returnValue = returnValue + prop.CurrentPropertyValue
    Next prop
    CalculateTotalValue = returnValue
End Function

' Performs the bankrupcy procedure for the player.
' If the player is bankrupted by the bank, set theBankruptor to NOTHING
Public Sub BankruptThePlayer(theBankruptPlayer As claIPlayer, theBankruptor As claIPlayer)
    Dim prop As claIProperty
    Dim newOwner As Integer
    Dim i As Long
    
    AppendToLog "modPlayerUtils.BankruptThePlayer"
    AppendToLog vbTab & "Player dead: " & theBankruptPlayer.Name
    If theBankruptor Is Nothing Then
        AppendToLog vbTab & "Killed by the bank"
    Else
        AppendToLog vbTab & "Killed by " & theBankruptor.Name
    End If
        
    SelfClosingMsgbox theBankruptPlayer.Name & " IS DECLARED BANKRUPT!!!!", vbOKOnly + vbExclamation, "BANKRUPT!"
    
    If theBankruptor Is Nothing Then
        newOwner = 0
    Else
        AppendToLog vbTab & "Transferring cash to the other player"
        newOwner = theBankruptor.number
        theBankruptor.Money = theBankruptor.Money + theBankruptPlayer.Money
    End If
    theBankruptPlayer.Money = 0
    
    If newOwner = 0 Then
        AppendToLog vbTab & "Transferring properties to bank"
    Else
        AppendToLog vbTab & "Transferring properties to other player"
    End If
    
    For Each prop In PropertyManager.PropertiesForPlayer(theBankruptPlayer)
        prop.DumpToLog
        If newOwner = 0 Then
            prop.RevertToBank
        Else
            prop.Owner = newOwner
        End If
        prop.Houses = 0
    Next prop
    
    ' Refresh the panel for this player - remove controls
    With theBankruptPlayer
        .IsBankrupt = True
        .SideBar.Items(1).Text = "Cash: *BANKRUPT*"
        .SideBar.Items(2).Text = "Total Value: *BANKRUPT*"
        For i = .SideBar.Items.Count To 3 Step -1
            .SideBar.Items.Remove i
        Next i
        .SideBar.State = eBarCollapsed
        .SideBar.Title = .SideBar.Title & " *R.I.P*"
    End With
End Sub

Public Sub SendPlayerToJail(thePlayer As claIPlayer)
    MovePlayerTokenDirect 10, thePlayer.number
    With PlayerManager.player(thePlayer.number)
        .InJail = True
        .JailCount = 0
        .Location = 10
        .DoublesCount = 0
    End With
End Sub
