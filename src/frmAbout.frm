VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About MyApp"
   ClientHeight    =   3150
   ClientLeft      =   2910
   ClientTop       =   5565
   ClientWidth     =   4890
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2174.186
   ScaleMode       =   0  'User
   ScaleWidth      =   4591.963
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   60
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1080
      TabIndex        =   5
      Top             =   2700
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   3120
      TabIndex        =   6
      Top             =   2700
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4499.936
      Y1              =   1563.343
      Y2              =   1563.343
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   630
      TabIndex        =   3
      Top             =   945
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   630
      TabIndex        =   1
      Top             =   60
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4499.936
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   630
      TabIndex        =   2
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Home Page"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   660
      TabIndex        =   4
      Top             =   2400
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmAbout
' Date      : Unknown
' Author    : Unknown (assumed template from VB6)
' Purpose   : Supply general info about this app
' 07/30/2004    SM:  Tidied up the code indentation and silly MS default code
'                    Moved non-related code to the modUtils module
' 03/03/2006    SM:  Added a RUN method
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub run()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = "Monopolie is a clone of the favorite board game."
    Me.Show
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    StartSysInfo
End Sub

