Attribute VB_Name = "modSound"
'---------------------------------------------------------------------------------------
' Module    : modSound
' Date      : 13/11/2003
' Author    : Simon M. Mitchell
' Purpose   : Sound module!
' 02/07/2004    Now uses sound configuration game option
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                                                             ByVal uFlags As Long) As Long
Private Const SND_ASYNC As Long = &H1
Private Const SND_LOOP As Long = &H8
Private Const SND_NODEFAULT As Long = &H2

' SM:  Need more sounds?
' 1)  Add an enumerator type to the SoundType
' 2)  Adjust the "select case" block in PlaySound
' 3)  Done!

Public Enum SoundType
    sndKaching = 1
    sndTick = 2
End Enum

Public Sub PlaySound(theSound As SoundType)
    Dim source As String
    
    If Globals.SoundEffects Then
        Select Case theSound
        Case SoundType.sndKaching
            source = "\kaching.wav"
        Case SoundType.sndTick
            source = "\tick.wav"
        Case Else
            source = vbNullString
        End Select
        
        If Len(source) > 0 Then sndPlaySound App.Path & source, SND_ASYNC Or SND_NODEFAULT
    End If
End Sub

