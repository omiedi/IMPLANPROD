VERSION 5.00
Begin VB.Form INIMASTER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicia Proceso de Maestro de Items de Stock"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3315
      Left            =   0
      Picture         =   "INIMASTER.frx":0000
      ScaleHeight     =   3315
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "    Maestro de Artículos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   3360
         TabIndex        =   1
         Top             =   210
         Width           =   1785
      End
   End
End
Attribute VB_Name = "INIMASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   VERANTER = "AMASTO01"
   '
   Call VERJOBID(OKEY%)
   Load AMASTOCK
   If OKEY% < 1 Then Call FINAL("")
   '
   Screen.MousePointer = 11
   DoEvents
   AMASTOCK.Show
   Screen.MousePointer = 1
   Hide
End Sub
