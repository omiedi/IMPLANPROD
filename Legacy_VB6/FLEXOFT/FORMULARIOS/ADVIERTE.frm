VERSION 5.00
Begin VB.Form ADVIERTE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalación de FLEXOFT"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   240
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      Caption         =   $"ADVIERTE.frx":0000
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenidos al programa de instalación de FLEXOFT"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "ADVIERTE.frx":00E0
      Top             =   480
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2055
      Left            =   240
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "ADVIERTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label3 = "1"
   Hide
End Sub

Private Sub Command2_Click()
   Label3 = "0"
   Unload Me
End Sub
