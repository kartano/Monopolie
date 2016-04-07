Attribute VB_Name = "MainLoop"
'---------------------------------------------------------------------------------------
' Module    : MainLoop
' Date      : Unknown
' Author    : Unknown
' Purpose   : Main line - program runs from here
'---------------------------------------------------------------------------------------

Option Explicit

Public PropertyManager    As claProperties
Public ChanceCards        As claCards
Public CommChestCards     As claCards

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Main()
    Globals.DefaultSettings
    InitCommonControls
    mdiMonopoly.Show
    If Command = "/Debug" Then DebugMode = True
End Sub

