VERSION 5.00
Begin VB.Form frmIncomeTax 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Income Tax"
   ClientHeight    =   2505
   ClientLeft      =   4020
   ClientTop       =   4110
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd10 
      Caption         =   "&10%"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmd200 
      Caption         =   "$&200"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmIncomeTax.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4830
   End
End
Attribute VB_Name = "frmIncomeTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmIncomeTax
' Date      : Unknown
' Author    : Unknown
' Purpose   : Income tax information dialog
' 03/03/2006            SM:  Removed a lot of cruft and fixed some indexes
'---------------------------------------------------------------------------------------

Option Explicit

Private iTotalWorth     As Long

Public Sub run()
    Me.Show vbModal
    Unload Me
End Sub

Private Sub cmd10_Click()
    PlayerManager.currentPlayer.SimpleTransaction Nothing, -iTotalWorth
    Globals.CheckAddFreeParkingCash iTotalWorth
    Me.Hide
End Sub

Private Sub cmd200_Click()
    PlayerManager.currentPlayer.SimpleTransaction Nothing, -200
    Globals.CheckAddFreeParkingCash 200
    Me.Hide
End Sub

Private Sub Form_Load()
    iTotalWorth = PlayerManager.currentPlayer.TotalValue / 10
    mdiMonopoly.SideMenu.Bars(PlayerManager.currentPlayer.Number + 1).Items(2).Text = "-----"
End Sub
