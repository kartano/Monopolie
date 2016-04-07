VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3900
   ClientLeft      =   4680
   ClientTop       =   4035
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtGoBonus 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin Monpolie.GroupBox fraOptions 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Unofficial Rules"
      Begin VB.CheckBox chkRentInJail 
         Caption         =   "Players collect rent while in Jail"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox chkMoneyPrivate 
         Caption         =   "Keep my Money Amount Private"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox chkHousing 
         Caption         =   "Unlimited Housing"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1995
      End
      Begin VB.CheckBox chkSoundEffects 
         Caption         =   "Sound Effects"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox chkLandGo 
         Caption         =   "Double For Landing on Go"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkFreeParking 
         Caption         =   "Free Parking Collects Fees"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Money for passing GO:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmOptions
' Date      : Unknown
' Author    : Unknown
' Purpose   : Game options dialog
'---------------------------------------------------------------------------------------

Option Explicit

Private mChoice As Boolean

'' SM:  Call using the RUN method
'' Should not be able to change these options once
'' the game has started!
Public Sub Run()
    mChoice = False
    Me.chkFreeParking.Value = IIf(Globals.FreeParkingFees, 1, 0)
    Me.chkLandGo.Value = IIf(Globals.GoPaysDouble, 1, 0)
    Me.chkSoundEffects = IIf(Globals.SoundEffects, 1, 0)
    Me.chkHousing = IIf(Globals.UnlimitedHousing, 1, 0)
    Me.chkMoneyPrivate = IIf(Globals.MoneyAmountsPrivate, 1, 0)
    Me.chkRentInJail = IIf(Globals.CollectRentInJail, 1, 0)
    Me.txtGoBonus = Globals.GoBonus
    
    ' These options should not be changed after the game starts
    If Globals.GameInProgress Then
        Me.chkFreeParking.Enabled = False
        Me.chkLandGo.Enabled = False
        Me.chkHousing.Enabled = False
        Me.chkMoneyPrivate.Enabled = False
        Me.chkRentInJail.Enabled = False
        Me.txtGoBonus.Locked = True
    End If
    
    Me.Show vbModal
    
    If mChoice Then
        Globals.SoundEffects = CBool(Me.chkSoundEffects)
        If Not Globals.GameInProgress Then
            Globals.FreeParkingFees = CBool(Me.chkFreeParking)
            Globals.GoPaysDouble = CBool(Me.chkLandGo)
            Globals.UnlimitedHousing = CBool(Me.chkHousing)
            Globals.MoneyAmountsPrivate = CBool(Me.chkMoneyPrivate)
            Globals.CollectRentInJail = CBool(Me.chkRentInJail)
            Globals.GoBonus = CLng(Me.txtGoBonus)
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mChoice = True
    Me.Hide
End Sub

Private Sub txtGoBonus_Validate(Cancel As Boolean)
    Cancel = InvalidAmount(Me.txtGoBonus)
End Sub
