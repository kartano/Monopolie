VERSION 5.00
Begin VB.Form frmLogSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log File Settings"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.CheckBox chkKeepLogFile 
      Caption         =   "Keep a log file"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmLogSettings
' Date      : 16/07/2004
' Author    : Simon M. Mitchell
' Purpose   : Log file settings dialog
'---------------------------------------------------------------------------------------

Option Explicit

Private mOkay As Boolean

Public Sub Run()
    mOkay = False
    
    Me.chkKeepLogFile = IIf(Globals.LogFileActive, 1, 0)
    Me.txtFilename = Globals.LogFilename
    
    Me.Show vbModal
    
    If mOkay Then
        If Me.chkKeepLogFile = 1 And Globals.LogFileActive = False Then
            Globals.CreateLogFile Me.txtFilename
        ElseIf Me.chkKeepLogFile = 0 And Globals.LogFileActive = True Then
            Globals.CloseLogFile
        ElseIf Me.chkKeepLogFile = 1 And Globals.LogFileActive Then
            ' Different filename?
            If UCase$(Me.txtFilename) <> UCase$(Globals.LogFilename) Then
                Globals.CloseLogFile
                Globals.CreateLogFile Me.txtFilename
            End If
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    Dim mFileDialog As cFileDialog
    
    On Error Resume Next
    
    Err.Clear
    Set mFileDialog = New cFileDialog
    With mFileDialog
        .hwnd = Me.hwnd
        .DialogTitle = "Select Log File"
        .InitDir = Me.txtFilename
        .DefaultExt = "txt"
        .CancelError = True
        .ShowSave
        If Err.Number = 0 Then Me.txtFilename = .Filename
    End With
    Set mFileDialog = Nothing
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If DataIsOkay Then
        mOkay = True
        Me.Hide
    End If
End Sub

Private Function DataIsOkay() As Boolean
    If Me.chkKeepLogFile = 1 And Len(Me.txtFilename) = 0 Then
        SelfClosingMsgbox "Please enter a Log Filename", vbOKOnly + vbInformation, "Log Settings"
        Me.txtFilename.SetFocus
        DataIsOkay = False
        Exit Function
    End If
    
    DataIsOkay = True
End Function
