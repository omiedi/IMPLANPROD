VERSION 5.00
Begin VB.Form FINAGETEL 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistemas FLEXOFT - Agenda Telefónica"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Comuníquese con Nosotros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   11
      Top             =   4200
      Width           =   5790
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FINAGETEL.frx":0000
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
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   1
      Left            =   210
      TabIndex        =   9
      Top             =   3255
      Width           =   5790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reanudar Ejecución"
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
      TabIndex        =   8
      Top             =   3465
      Width           =   2010
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Terminar"
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
      Left            =   3465
      TabIndex        =   7
      Top             =   3465
      Width           =   2010
   End
   Begin VB.PictureBox Picture1 
      Height          =   1065
      Left            =   210
      Picture         =   "FINAGETEL.frx":00A9
      ScaleHeight     =   1005
      ScaleWidth      =   5730
      TabIndex        =   6
      Top             =   105
      Width           =   5790
      Begin VB.Label Label6 
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
         Left            =   2205
         TabIndex        =   13
         Top             =   735
         Width           =   3690
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   1680
      Width           =   5790
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   540
      Left            =   210
      TabIndex        =   2
      Top             =   5880
      Width           =   5790
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   105
         TabIndex        =   10
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "es Marca Registrada propiedad del Ing. Ernesto Pablo Bauer"
         Height          =   270
         Left            =   1260
         TabIndex        =   3
         Top             =   210
         Width           =   4480
      End
   End
   Begin VB.Label Label7 
      Caption         =   " "
      Height          =   435
      Left            =   105
      TabIndex        =   14
      Top             =   3465
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Line Line2 
      X1              =   5995
      X2              =   5995
      Y1              =   105
      Y2              =   1165
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   5995
      Y1              =   1165
      Y2              =   1165
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gracias por Usar la Agenda Telefónica FLEXOFT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   5790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"FINAGETEL.frx":62CB
      Height          =   645
      Left            =   105
      TabIndex        =   1
      Top             =   2415
      Width           =   6000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "El presente programa aplicativo FLEXOFT (r) se encuentra amparado por la Ley 11723 de Propiedad Intelectual."
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   1890
      Width           =   5790
   End
End
Attribute VB_Name = "FINAGETEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label7 = "REANUDA"
    Hide
End Sub

Private Sub Command2_Click()
    Unload Me
    'Hide
    'Call CIERRARCH("*.*")
    'Call FINAL("")
End Sub

Private Sub Form_Activate()
   Command1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Call CIERRARCH("*.*")
   'Call GENBACKUP
   'Call FINAL("")
End Sub
