VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Height          =   675
      Left            =   1800
      Picture         =   "prueba04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aprueba - Graba la Imputación"
      Top             =   600
      Width           =   650
   End
   Begin VB.PictureBox Picture16 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
      Begin VB.Image Image3 
         Height          =   2500
         Left            =   0
         Picture         =   "prueba04.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1065
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Me.Picture16.Visible = True
'      Me.BackColor = &HFCD7B6
'      Me.Frame1.BackColor = &HFCD7B6
'      Me.Frame2.BackColor = &HFCD7B6
'      Me.Frame3.BackColor = &HFCD7B6
'      Me.Frame4.BackColor = &HFCD7B6
'      Me.Frame5.BackColor = &HFCD7B6
'      MARCOBARRA.Top = 5500
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
'    End If
'&H8000000F

End Sub
