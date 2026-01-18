VERSION 5.00
Begin VB.Form INICLIEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes y Deudores"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3105
      Left            =   0
      Picture         =   "INICLIEN.frx":0000
      ScaleHeight     =   3105
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A-B-M Clientes "
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
         Left            =   3465
         TabIndex        =   1
         Top             =   210
         Width           =   1575
      End
   End
End
Attribute VB_Name = "INICLIEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   'Show
   VERANTER$ = "AMACLI01"
   Screen.MousePointer = 11
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   AMACLIEN.Show
   Screen.MousePointer = 1
   Hide
   '
End Sub

