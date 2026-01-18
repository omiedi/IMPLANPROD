VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   105
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3885
      Width           =   4425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   435
      Left            =   420
      TabIndex        =   6
      Top             =   4515
      Width           =   1170
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pegado a si Mismo"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   3570
      Width           =   2640
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Objeto Thing Temporal"
      Height          =   435
      Left            =   210
      TabIndex        =   4
      Top             =   2835
      Width           =   4110
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Liberar el Objeto Thing"
      Height          =   435
      Left            =   210
      TabIndex        =   3
      Top             =   1890
      Width           =   4110
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invertir el Nombre del Objeto Thing"
      Height          =   435
      Left            =   210
      TabIndex        =   2
      Top             =   1365
      Width           =   4110
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mostrar el Objeto Thing"
      Height          =   435
      Left            =   210
      TabIndex        =   1
      Top             =   840
      Width           =   4110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear nuevo Objeto Thing"
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   4110
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   540
      Left            =   2520
      TabIndex        =   7
      Top             =   4410
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FMGR1 As FMGR

Private Sub Form_Load()
   Set FMGR1 = New FMGR
End Sub
    
Private Sub Command6_Click()
   Label1 = FMGR1.EXISTE%(Text1)
   Label1.Refresh
End Sub
