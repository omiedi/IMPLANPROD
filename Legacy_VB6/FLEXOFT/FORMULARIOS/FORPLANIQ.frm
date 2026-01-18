VERSION 5.00
Begin VB.Form FORPLANIQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Producción"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2730
      Picture         =   "FORPLANIQ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   750
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      Height          =   855
      Left            =   105
      OleObjectBlob   =   "FORPLANIQ.frx":014A
      TabIndex        =   0
      Top             =   105
      Width           =   2475
   End
End
Attribute VB_Name = "FORPLANIQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Form_Load()
   Refresh
   DoEvents
End Sub
