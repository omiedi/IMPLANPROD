VERSION 5.00
Begin VB.Form INIPROVE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes y Deudores"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   2895
      Left            =   0
      Picture         =   "INIPROVE.frx":0000
      ScaleHeight     =   2895
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A-B-M Proveedores  "
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
         Left            =   3150
         TabIndex        =   1
         Top             =   210
         Width           =   1890
      End
   End
End
Attribute VB_Name = "INIPROVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   Show
   Screen.MousePointer = 11
   Call VERJOBID
   AMAPROVE.Show
   Screen.MousePointer = 1
   Hide
   '
End Sub

