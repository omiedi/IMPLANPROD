VERSION 5.00
Begin VB.Form FrmErrorControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Componentes de Windows"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Accder a Instalador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   120
      TabIndex        =   6
      Top             =   3885
      Width           =   2325
   End
   Begin VB.ListBox LISTADEOBJETOS 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      IntegralHeight  =   0   'False
      ItemData        =   "FrmErrorControl.frx":0000
      Left            =   210
      List            =   "FrmErrorControl.frx":0007
      TabIndex        =   2
      Top             =   840
      Width           =   4425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Instalar Ahora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   120
      TabIndex        =   1
      Top             =   4365
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar Sin Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2520
      TabIndex        =   0
      Top             =   3885
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmErrorControl.frx":001B
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
      Left            =   210
      TabIndex        =   5
      Top             =   0
      Width           =   4425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmErrorControl.frx":00CB
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
      Left            =   210
      TabIndex        =   4
      Top             =   3150
      Width           =   4530
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Top             =   3990
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "FrmErrorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label3 = "BAJAR"
   Hide
End Sub
Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command3_Click()
   Label3 = "FXSETUP"
   Hide

End Sub
