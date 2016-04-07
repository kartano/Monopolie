Attribute VB_Name = "Globals"
'---------------------------------------------------------------------------------------
' Module    : Globals
' Date      : Unknown
' Author    : Unknown
' Purpose   : General global utilities and data??
' 07/02/2004    SM:  Added game options
' 07/07/2004    SM:  Added "Rent In Jail" option
' 07/14/2004    SM:  Added new Log file options
' 07/18/2004    SM:  CreateLogFile clears mLogFilename if log create fails
'---------------------------------------------------------------------------------------

Option Explicit

Private mDebugMode As Boolean
Public DiceRoll1 As Integer
Public DiceRoll2 As Integer
Public iHouseCount As Integer

' Game in progress
Private mGameInProgress As Boolean

' Game options
Private mFreeParkingFees As Boolean
Private mGoPaysDouble As Boolean
Private mSoundEffects As Boolean
Private mUnlimitedHousing As Boolean
Private mMoneyAmountsPrivate As Boolean
Private mCollectRentInJail As Boolean
Private mGoBonus As Long
Private mFreeParkingBonusCash As Long

' Log files
Private mLogFilename As String
Private mLogFile As claBasicLogFile

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'' SM:  The new Players manager class
'' Part of the CPU player upgrade
Private mPlayers As claPlayers

Public Property Get PlayerManager() As claPlayers
    If mPlayers Is Nothing Then Set mPlayers = New claPlayers
    Set PlayerManager = mPlayers
End Property

Public Property Get DebugMode() As Boolean
    DebugMode = mDebugMode
End Property

Public Property Let DebugMode(RHS As Boolean)
    mDebugMode = RHS
End Property

Public Sub InitClasses()
    Set PropertyManager = New claProperties
    Set ChanceCards = New claCards
    Set CommChestCards = New claCards
End Sub

Public Sub ShowTransaction(ByVal Amount As Integer)
    Dim sResult As String

    sResult = FormatCash(CLng(Amount))
    mdiMonopoly.UpdateTransactionText sResult
End Sub

Public Property Get GameInProgress() As Boolean
    GameInProgress = mGameInProgress
End Property
Public Property Let GameInProgress(RHS As Boolean)
    mGameInProgress = RHS
End Property
Public Property Get FreeParkingFees() As Boolean
    FreeParkingFees = mFreeParkingFees
End Property
Public Property Let FreeParkingFees(RHS As Boolean)
    mFreeParkingFees = RHS
End Property
Public Property Get GoPaysDouble() As Boolean
    GoPaysDouble = mGoPaysDouble
End Property
Public Property Let GoPaysDouble(RHS As Boolean)
    mGoPaysDouble = RHS
End Property
Public Property Get SoundEffects() As Boolean
    SoundEffects = mSoundEffects
End Property
Public Property Let SoundEffects(RHS As Boolean)
    mSoundEffects = RHS
End Property
Public Property Get UnlimitedHousing() As Boolean
    UnlimitedHousing = mUnlimitedHousing
End Property
Public Property Let UnlimitedHousing(RHS As Boolean)
    mUnlimitedHousing = RHS
End Property
Public Property Get MoneyAmountsPrivate() As Boolean
    MoneyAmountsPrivate = mMoneyAmountsPrivate
End Property
Public Property Let MoneyAmountsPrivate(RHS As Boolean)
    mMoneyAmountsPrivate = RHS
End Property

Public Property Get GoBonus() As Long
    GoBonus = mGoBonus
End Property
Public Property Let GoBonus(RHS As Long)
    mGoBonus = RHS
End Property

Public Property Get FreeParkingBonusCash() As Long
    FreeParkingBonusCash = mFreeParkingBonusCash
End Property
Public Property Let FreeParkingBonusCash(RHS As Long)
    mFreeParkingBonusCash = RHS
End Property

Public Property Get CollectRentInJail() As Boolean
    CollectRentInJail = mCollectRentInJail
End Property
Public Property Let CollectRentInJail(RHS As Boolean)
    mCollectRentInJail = RHS
End Property

Public Sub CheckAddFreeParkingCash(Amount As Long)
    If FreeParkingFees Then
        FreeParkingBonusCash = FreeParkingBonusCash + Amount
    End If
End Sub

' Set default globals
Public Sub DefaultSettings()
    GameInProgress = False
    SoundEffects = True
    CollectRentInJail = True
    GoBonus = 200
End Sub

' NOTE ABOUT LOG FILES
' A lot of this seems strange - I'm wrapping the log functionality
' in the Globals (as opposed to just exposing the mLogFile class) so
' that we can allow for the fact that log files can be turned on/off and
' filenames changed ad-hoc during program execution.
' This way we get better control.

' Create a fresh log file
' Silently close any existing log file
Public Sub CreateLogFile(Filename As String)
    If Not mLogFile Is Nothing Then
        mLogFile.CloseLog
        mLogFilename = Filename
    End If
    Set mLogFile = New claBasicLogFile
    If Not mLogFile.CreateNewLog(Filename) Then
        Set mLogFile = Nothing
        mLogFilename = vbNullString
    End If
End Sub

Public Property Get LogFilename() As String
    LogFilename = mLogFilename
End Property

' Close and kill the active log file
' Silently ignore if no log is active
Public Sub CloseLogFile()
    If Not mLogFile Is Nothing Then
        mLogFile.CloseLog
        Set mLogFile = Nothing
        mLogFilename = vbNullString
    End If
End Sub

' Append text to the log file - silently ignore if no log is active
Public Sub AppendToLog(theText As String)
    If Not mLogFile Is Nothing Then mLogFile.Append theText
End Sub

Public Property Get LogFileActive() As Boolean
    LogFileActive = Not (mLogFile Is Nothing)
End Property
