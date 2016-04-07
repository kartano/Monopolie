VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRules 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Monopolie Rules"
   ClientHeight    =   4020
   ClientLeft      =   6405
   ClientTop       =   2610
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtbRules 
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6906
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmRules.frx":0000
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmRules
' Date      : 10/5/2003
' Author    : Brian G. Schmitt
' Purpose   : Used to display the Rules which is an external file "Rules.rtf"
' 03/03/2006            SM:  Standardised code header
'---------------------------------------------------------------------------------------

Option Explicit

Public Sub run()
    Me.Show
End Sub

Private Sub Form_Load()
    rtbRules.LoadFile (App.Path & "\doc\rules.rtf")
End Sub

Private Sub Form_Resize()
    rtbRules.Top = 60
    rtbRules.Left = 60
    rtbRules.Width = Me.Width - 240
    rtbRules.Height = Me.Height - 540
End Sub
