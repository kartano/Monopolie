Attribute VB_Name = "PlayerCoords"
'---------------------------------------------------------------------------------------
' Module    : PlayerCoords
' Date      : Unknown
' Author    : Various
' Purpose   : Player coordinates and utilities
' 07/02/2004    SM:  Added double bonus for passing go
'---------------------------------------------------------------------------------------

Option Explicit

' THESE CONSTANTS ARE USED TO DETERMINE MOVEMENT DIRECTION
Private Const cForward          As Integer = 1
Private Const cBackward         As Integer = 2

' THESE ARE THE LOCATIONS PLAYER TOKENS ARE DRAWN AT ON THE BOARD
' TokenCoords(Player#, Location#, 0) = Left coordinate
' TokenCoords(Player#, Location#, 1) = Top coordinate

'' SM NOTE:  The arrays are base 0 - if the array is defined
'' to 40 places, this means 41 squares.  Was there a reason why we
'' did this??
Private TokenCoords(40, 1)      As Integer

' SM:  Array indexes for TokenCoords
Private Const cLeftCoord        As Integer = 0
Private Const cTopCoord         As Integer = 1

' USE TO HANDLE MOVEMENT, SIMPLIFIES HANDLING OF CASES
' WHERE THE LOCATIONS MAKE THE TRANSITION FROM 39 TO 0 AND OVER
' WHEN A PLAYER PASSES GO
' MAKING IT TWICE AS ALRGE AS IT NEEDS TO BE FOR CASES OF "ADVANCE TO xxx..."
' THAT CAN BE CAUSED BY CHANCE CARDS
'' SM:  Might be able to remove this.  See notes below.
Private MoveIndices(80)         As Integer

Public Sub AdvanceToRailRoad(ByVal CurrLocation As Integer, _
                             Optional MultiplierX As Integer = 1)

    Dim Destination    As Integer
    Dim iFinalLocation As Integer

    Select Case CurrLocation
    Case 0 To 4                     'Reading R.R.
        Destination = 5
    Case 5 To 14                    'Pennslyvania R.R.
        Destination = 15
    Case 15 To 24                   'B&O R.R.
        Destination = 25
    Case 25 To 34                   'Short Line R.R.
        Destination = 35
    Case 35 To 40                   'Reading R.R. plus pass GO
        Destination = 5
    End Select
    iFinalLocation = MovePlayerToken(CurrLocation, (Destination - CurrLocation), 1, MultiplierX)
End Sub

Public Function GetStreet(ByVal property) As Integer
  ' Returns the street the passed property is on or -1 for error.
  ' @author James
  '
    If property > 0 And property < 10 Then
        GetStreet = 1
    ElseIf property > 10 And property < 20 Then 'NOT PROPERTY...
        GetStreet = 2
    ElseIf property > 20 And property < 30 Then 'NOT PROPERTY...
        GetStreet = 3
    ElseIf property > 30 And property < 40 Then 'NOT PROPERTY...
        GetStreet = 4
    Else
        GetStreet = -1
        '' SM:  If we ever get here, we really need to know what
        '' went wrong as this would indicate a very nasty
        '' programming error somewhere.
        MsgBox "Invalid property value specified for PlayCoords.GetStreet!"
        Stop
    End If
End Function

Private Sub InitMoveIndices()
    Dim ndx As Integer

    For ndx = LBound(MoveIndices) To UBound(MoveIndices)
        If (ndx < 40) Then
        ' LOCATIONS 0-39
            MoveIndices(ndx) = ndx
        Else 'NOT (NDX...
            '' SM:  Do we need these?  Can we just take
            '' the ndx modulus with 40?  This would save
            '' having to have a large array!
            ' "WRAP-AROUND" LOCATIONS USED WHEN PASSING GO
            MoveIndices(ndx) = ndx - 40
        End If
    Next ndx
End Sub

Public Sub InitPlayerCoords()
    Dim i As Long
    Dim tokens As Collection
    Dim setting As String
    
    '' SM:  AGain, this is 41 squares in total.  Why are we doing this??
    For i = 0 To 40
        setting = modINIUtils.INISetting("PlayerCoords", "Coord" & i, ConfigINI)
        Set tokens = modUtils.Tokenize(setting, ",")
        SetPlayerCoords i, CInt(tokens(1)), CInt(tokens(2))
    Next i
    Set tokens = Nothing
End Sub

Public Function MovePlayerToken(ByVal CurrLocation As Integer, _
                                ByVal Steps As Integer, _
                                Optional Direction As Integer = cForward, _
                                Optional MultiplierX As Integer = 1) As Integer
  
    Dim LeftCoord As Integer
    Dim TopCoord  As Integer
    Dim MovePos   As Integer
    Dim ndx       As Integer
    Dim StartLoc  As Integer
    Dim FinishLoc As Integer
    Dim StepVal   As Integer

    StepVal = IIf(Direction = cForward, 1, -1)
    StartLoc = CurrLocation + StepVal
    FinishLoc = CurrLocation + (StepVal * Steps)
    
    For ndx = StartLoc To FinishLoc Step StepVal
        ' ADDING ndx TO Sleep TIME SO IT VARIES A BIT
        ' AND HOPEFULLY REDUCES ANY FLICKER
        Sleep 203 + ndx
    
        '' SM:  This might be dangerous as all button/menu
        '' events can fire while the token is moving!
    
        DoEvents
    
        Sleep 173 + ndx
    
        '' SM:  we can probably just take
        '' the modulus of ndx and 40 as the move pos rather
        '' than have a completely seperate indices array
        '' for this bit.
        ''MovePos = MoveIndices(ndx)
        MovePos = ndx Mod 40
        LeftCoord = TokenCoords(MovePos, cLeftCoord)
        TopCoord = TokenCoords(MovePos, cTopCoord)
        With PlayerManager.currentPlayer
            frmMain.Token(.Token).Left = LeftCoord
            frmMain.Token(.Token).Top = TopCoord
            frmMain.Refresh
            If (MovePos = 0) Then
                ' Player ends turn by LANDING on Go
                If ndx = FinishLoc Then
                    If Globals.GoPaysDouble Then
                        .SimpleTransaction Nothing, Globals.GoBonus * 2
                    Else
                        .SimpleTransaction Nothing, Globals.GoBonus
                    End If
                Else ' Play is passing go
                    .SimpleTransaction Nothing, Globals.GoBonus
                End If
                'mdiMonopoly.UpdateGameText "Passed Go", GameTransaction
            End If
        End With
        modSound.PlaySound sndTick
    Next ndx
  
    FinishLoc = FinishLoc Mod 40
    ''If FinishLoc >= 40 Then
    ''  FinishLoc = FinishLoc - 40
    ''End If
  
    MovePlayerToken = FinishLoc
  
    Sleep 1000   ' PAUSE PRIOR TO ANY OTHER ACTION
    
    '' SM:  Moved the property landing logic into the
    '' player classes.  Thought this would make handling
    '' CPU decisions a lot easier.
    PlayerManager.currentPlayer.PropertyLand FinishLoc, MultiplierX
    ''PropertyLand FinishLoc, MultiplierX

End Function

Public Sub MovePlayerTokenDirect(ByVal Location As Integer, _
                                 ByVal PlayerNum As Integer)
    Dim TopCoord  As Integer
    Dim LeftCoord As Integer

    ' THIS IS ALMOST EXACTLY THE SAME AS THE OLD MovePlayerToken() ROUTINE
    ' THE ONLY PLACE THAT IT WILL BE CALLED IS FROM GoDirectlyToJail() IN frmMain
    LeftCoord = TokenCoords(Location, cLeftCoord)
    TopCoord = TokenCoords(Location, cTopCoord)
    frmMain.Token(PlayerManager.player(PlayerNum).Token).Left = LeftCoord
    frmMain.Token(PlayerManager.player(PlayerNum).Token).Top = TopCoord
    PlayerManager.player(PlayerNum).Location = Location
End Sub

Private Sub SetPlayerCoords(ByVal Location As Integer, _
                            ByVal Left As Integer, _
                            ByVal Top As Integer)

    ' This routines stores player token screen coordinates for a given location into
    ' the global array TokenCoords (declared in Globals.bas)
    TokenCoords(Location, cLeftCoord) = Left
    TokenCoords(Location, cTopCoord) = Top
End Sub
