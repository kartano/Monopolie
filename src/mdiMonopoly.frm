VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.0#0"; "vbalExpBar6.ocx"
Begin VB.MDIForm mdiMonopoly 
   BackColor       =   &H8000000C&
   Caption         =   "Monopoly"
   ClientHeight    =   5085
   ClientLeft      =   1755
   ClientTop       =   2910
   ClientWidth     =   7755
   Icon            =   "mdiMonopoly.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl SideMenu 
      Align           =   4  'Align Right
      Height          =   4710
      Left            =   4920
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   8308
      UseExplorerTransitionStyle=   0   'False
      BackColorEnd    =   0
      BackColorStart  =   16777215
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4710
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8043
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuCustomGameOptions 
         Caption         =   "Custom Game Options ..."
      End
      Begin VB.Menu mnuLogFileOptions 
         Caption         =   "Log File Options ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpRules 
         Caption         =   "Rules"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelpSubmitBug 
         Caption         =   "Submit Bug"
      End
   End
End
Attribute VB_Name = "mdiMonopoly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : mdiMonopoly
' Date      : 10/8/2003
' Author    : Brian G. Schmitt
' Purpose   : Main Form to Start/End/Save/Open Games
' 07/30/2004            SM:  Added Log File settings to Options menu
' 21/08/2004            SM:  Improved explorer bar code, added "Resign"
' 03/03/2006            SM:  Removed a lot of cruft, standardised code header
'---------------------------------------------------------------------------------------

Option Explicit

Public bNewGame        As Boolean
Private FileDialog     As cFileDialog

'' SM:  Used to prevent the main MDI window being closed
'' when the user clicks on the "X".  This can cause serious
'' errors when the game is actually in progress.
Private mCanClose As Boolean

Private Sub MDIForm_Load()
    mCanClose = False
    Set FileDialog = New cFileDialog
    If Me.Height < 7700 Then Me.Height = 7700
    If Me.Width < 8300 Then Me.Width = 8300
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not mCanClose
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Not mCanClose Then
        Cancel = True
    Else
        Set PropertyManager = Nothing
        Set ChanceCards = Nothing
        Set CommChestCards = Nothing
    End If
End Sub

Private Sub mnuCustomGameOptions_Click()
    frmOptions.run
End Sub

Private Sub mnuFileExit_Click()
    mCanClose = True
    Unload Me
    End
End Sub

Private Sub mnuFileNew_Click()
    frmSetup.run
    If bNewGame Then
        frmMain.Show
        If DebugMode Then frmDebug.Show
        mdiMonopoly.StatusBar.Panels(1) = "Current Player: " & PlayerManager.currentPlayer.Name
    End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim iCount         As Integer
    Dim sNumPlayers    As String
    Dim sProperty      As String
    Dim sCurrentPlayer As String
    Dim sPlayer        As String
    Dim AryReturn      As Variant
    Dim FileNumber     As Integer

    On Error GoTo ErrorOpen
  
    FileNumber = FreeFile
    With FileDialog
        .hwnd = Me.hwnd
        .DialogTitle = "Monooplie Save File"
        .Filter = "Monopoly Save File: *.mon"
        .DefaultExt = "mon"
        .InitDir = App.Path
        .CancelError = True
        .ShowOpen
        Open .Filename For Input Shared As FileNumber
        InitClasses
        frmMain.Show
        ResetForNewGame
        Input #FileNumber, sNumPlayers
        For iCount = 1 To CInt(sNumPlayers)
            Input #FileNumber, sPlayer
            sPlayer = Convert2Ascii(sPlayer)
            AryReturn = Split(sPlayer, ".", -1, vbTextCompare)
            '' SM:  TO DO:  Write a "load" method for a player
            ''PlayerManager.AddPlayer CInt(AryReturn(2)), CStr(AryReturn(1)), CInt(AryReturn(10)), CInt(AryReturn(11)), CInt(AryReturn(12)), CInt(AryReturn(5))
            With PlayerManager.player(CInt(AryReturn(2)))
                .Money = CInt(AryReturn(3))
                .Location = AryReturn(4)
                MovePlayerTokenDirect (AryReturn(4)), CInt(AryReturn(2))
                .DoublesCount = AryReturn(6)
                .JailCount = AryReturn(7)
                .OOJailComm = AryReturn(8)
                .OOJailChance = AryReturn(9)
                .Status = CInt(AryReturn(13))
            End With
        Next iCount
        Input #FileNumber, sProperty, sCurrentPlayer
        PropertyManager.RestoreProperty sProperty
        Do
            PlayerManager.PassTheDice
        Loop Until PlayerManager.currentPlayer.Number = sCurrentPlayer
        Close #FileNumber
    End With

ErrorOpen:
    '' SM: TO DO - descriptive error message and append to log if necessary.
    '' Any catastrophies during an open should be reported!
    'User pressed the cancel button
End Sub

Private Sub mnuFileSave_Click()
    Dim iCount As Integer
    Dim FileNumber As Integer
    Dim player As claIPlayer

    On Error GoTo ErrorSave
  
    FileNumber = FreeFile
  
    With FileDialog
        .hwnd = Me.hwnd
        .DialogTitle = "Monooplie Save File"
        .Filter = "Monopoly Save File: *.mon"
        .DefaultExt = "mon"
        .InitDir = App.Path
        .CancelError = True
        .ShowSave
        Open .Filename For Output Shared As FileNumber
        Write #FileNumber, PlayerManager.Count
        For Each player In PlayerManager
            Write #FileNumber, player.SavePlayer
        Next player
        Write #FileNumber, PropertyManager.SaveProperty
        Write #FileNumber, PlayerManager.currentPlayer.Number
        Close #FileNumber
    End With

ErrorSave:
    '' SM:  Add a descriptive error message, as per the file/open
    'User pressed cancel button
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.run
End Sub

Private Sub mnuHelpRules_Click()
    frmRules.run
End Sub

Private Sub mnuHelpSubmitBug_Click()
    '' SM:  Make sure thius still works.
    ShellExecute Me.hwnd, vbNullString, "http://sourceforge.net/tracker/?func=add&group_id=56153&atid=479539", vbNullString, "C:\", 1
End Sub

Public Sub ResetForNewGame()
    Dim tBar     As cExplorerBar
    Dim iCounter As Integer

    SideMenu.Bars.Clear
    Set tBar = SideMenu.Bars.Add(, "Game", "Game")
    tBar.IsSpecial = True
    With tBar.Items
        .Add , "Last Transaction: ", "Last Transaction: ", , 1
        .Add , "Free Parking: ", "Free Parking: ", , 1
        .Add ItemType:=2 'used to bump the index for unowned props
        .Add ItemType:=2 'used to bump the index for unowned props
        .Add ItemType:=2 'used to bump the index for unowned props
        .Add , "Unowned Properties0", "Unowned Properties", , 0
    End With
    
    For iCounter = 0 To frmMain.Token.Count - 1
        frmMain.Token(iCounter).Visible = False
    Next iCounter
End Sub

Private Sub mnuLogFileOptions_Click()
    frmLogSettings.run
End Sub

Private Sub SideMenu_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    Dim iCounter        As Integer
    Dim iSelectedPlayer As Integer
    Dim newTrade As claTrade
    
    If frmMain.TurnInProgress Then Exit Sub
    
    '' SM:  This will not work if we allow for more than 9
    '' players.  Pretty unlikely, but be aware of this.
    iSelectedPlayer = Right$(itm.key, 1)
    
    Select Case UCase$(itm.Text)
    Case "PROPERTY CARDS"
        frmCards.run iSelectedPlayer
    Case "TRADE"
        Set newTrade = New claTrade
        newTrade.Load PlayerManager.currentPlayer, Nothing
        If newTrade.Edit Then
            If newTrade.TradeAccepted Then
                newTrade.Execute
            End If
        End If
        Set newTrade = Nothing
    Case "PROPERTIES OWNED"
        If iSelectedPlayer Then
            For iCounter = 1 To 39
                If PropertyManager.isProperty(iCounter) Then
                    If PropertyManager.GetOwner(iCounter) = iSelectedPlayer Then
                        frmMain.Images(iCounter).Picture = frmMain.Token(PlayerManager.player(iSelectedPlayer).Token).Picture
                    Else
                        frmMain.Images(iCounter).Picture = frmMain.Images(0).Picture
                    End If
                End If
            Next iCounter
        End If
    Case "RESIGN"
        If MsgBox("Are you really sure that you want to resign??", vbYesNo + vbQuestion, "Resign") = vbYes Then
            PlayerManager.player(iSelectedPlayer).BankruptThisPlayer Nothing
            ' If the current player just quit, pass the dice
            If iSelectedPlayer = PlayerManager.currentPlayer.Number Then
                PlayerManager.PassTheDice
            End If
        End If
    Case "UNOWNED PROPERTIES"
        frmCards.run 0
    End Select
End Sub

Public Sub UpdateFreeParkingText(theText As String)
    SideMenu.Redraw = False
    mdiMonopoly.SideMenu.Bars(1).Items(2).Text = "Free Parking: " & theText
    SideMenu.Redraw = True
End Sub

Public Sub UpdateTransactionText(theText As String)
    SideMenu.Redraw = False
    mdiMonopoly.SideMenu.Bars(1).Items(1).Text = "Last Transaction: " & theText
    SideMenu.Redraw = True
End Sub
