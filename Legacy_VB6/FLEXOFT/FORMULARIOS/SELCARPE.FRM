VERSION 5.00
Begin VB.Form SELCARPE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carpeta de Exportación"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   645
      Left            =   2835
      Picture         =   "SELCARPE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1995
      Width           =   645
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   210
      TabIndex        =   1
      Top             =   525
      Width           =   2430
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   2430
   End
End
Attribute VB_Name = "SELCARPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Dir1_Change()
   On Error Resume Next
   RUTASAL$ = Dir1.Path
   WLITAB04.Label11.Caption = RUTASAL$
   On Error GoTo 0
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   On Error GoTo 0
End Sub
