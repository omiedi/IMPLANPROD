VERSION 5.00
Begin VB.Form FINSECPPT 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistemas FLEXOFT - Demostraciones"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   4305
      Width           =   5790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver a Ejecutar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   735
      TabIndex        =   6
      Top             =   4620
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Menu Principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3570
      TabIndex        =   5
      Top             =   4620
      Width           =   1800
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   210
      ScaleHeight     =   795
      ScaleWidth      =   5730
      TabIndex        =   4
      Top             =   210
      Width           =   5790
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones Informáticas Integrales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2100
         TabIndex        =   15
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   2205
      Width           =   5790
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000014&
      Height          =   435
      Index           =   1
      Left            =   225
      Top             =   7085
      Width           =   5790
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      Height          =   1590
      Index           =   1
      Left            =   225
      Top             =   5310
      Width           =   5790
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "FLEXOFT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   14
      Top             =   7140
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "es Marca Registrada propiedad del Ing. Ernesto Pablo Bauer"
      Height          =   270
      Left            =   1470
      TabIndex        =   13
      Top             =   7200
      Width           =   4485
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      Height          =   435
      Index           =   0
      Left            =   210
      Top             =   7070
      Width           =   5790
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comuniquese con Nosotros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   12
      Top             =   5385
      Width           =   3270
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      Height          =   1590
      Index           =   0
      Left            =   210
      Top             =   5295
      Width           =   5790
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"FINSECPPT.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   840
      TabIndex        =   11
      Top             =   5775
      Width           =   5160
   End
   Begin VB.Label Label9 
      Height          =   435
      Left            =   210
      TabIndex        =   10
      Top             =   4725
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"FINSECPPT.frx":00A7
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
      TabIndex        =   9
      Top             =   2415
      Visible         =   0   'False
      Width           =   5790
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   8
      Top             =   1785
      Width           =   5790
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   210
      Y2              =   1060
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   6000
      Y1              =   1060
      Y2              =   1060
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gracias por Evaluar la Secuencia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   1365
      Width           =   5790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"FINSECPPT.frx":015F
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
      TabIndex        =   1
      Top             =   3360
      Width           =   5790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   $"FINSECPPT.frx":0232
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   2415
      Width           =   5790
   End
   Begin VB.Image Image1 
      Height          =   8000
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6250
   End
End
Attribute VB_Name = "FINSECPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SETRAN

Private Sub Command1_Click()
    Label9.Caption = "1"
    Hide
End Sub

Private Sub Command2_Click()
    Label9.Caption = "0"
    Hide
End Sub

Private Sub Form_Activate()
   Label9.Caption = ""
   SETRAN = 0
   Command2.SetFocus
End Sub

Private Sub Form_Load()
   If MINIDEM% <> 1 Then
       FINSECUEN.Image1.Picture = MENUDEMO.Image1(1).Picture
       Picture1.Picture = PROPIPPT.Picture1.Picture
     Else
       Picture1.Picture = PROPINTEM.Picture1.Picture
   End If
End Sub

Private Sub Timer1_Timer()
   SETRAN = SETRAN + 1
   If SETRAN > 30 Then
     Label9.Caption = "1"
     On Error Resume Next
     Hide
     On Error GoTo 0
   End If
End Sub
