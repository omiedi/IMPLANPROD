VERSION 5.00
Begin VB.Form SELUDESTI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalación de FLEXOFT"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   240
      X2              =   5760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      Caption         =   $"SELUDESTI.frx":0000
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "SELUDESTI.frx":012D
      Top             =   480
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2295
      Left            =   240
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "SELUDESTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label3 = UCase$(Left$(Drive1.Drive, 1))
   Hide
End Sub

Private Sub Command2_Click()
   Label3 = ""
   Unload Me
End Sub
