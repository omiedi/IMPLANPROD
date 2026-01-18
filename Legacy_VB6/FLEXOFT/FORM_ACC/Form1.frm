VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aprueba y Graba"
      Top             =   1200
      Width           =   650
   End
   Begin VB.PictureBox Picture16 
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2280
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
      Begin VB.Image Image3 
         Height          =   2025
         Left            =   0
         Picture         =   "Form1.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 3600
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If

End Sub

