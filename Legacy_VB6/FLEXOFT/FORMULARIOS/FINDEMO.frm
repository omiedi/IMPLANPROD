VERSION 5.00
Begin VB.Form FINDEMO 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistemas FLEXOFT - Demostraciones"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   240
      Picture         =   "FINDEMO.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   5730
      TabIndex        =   12
      Top             =   240
      Width           =   5790
      Begin VB.Label Label7 
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
         TabIndex        =   13
         Top             =   525
         Width           =   3480
      End
   End
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
      TabIndex        =   10
      Top             =   4410
      Width           =   5790
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FINDEMO.frx":11302
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
         TabIndex        =   11
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   3360
      Width           =   5790
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reanudar Demo"
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
      TabIndex        =   7
      Top             =   3675
      Width           =   1800
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
      Left            =   3570
      TabIndex        =   6
      Top             =   3675
      Width           =   1800
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   1785
      Width           =   5790
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   540
      Left            =   210
      TabIndex        =   2
      Top             =   6090
      Width           =   5790
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "FLEXOFT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   9
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
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6030
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line2 
      X1              =   6030
      X2              =   6030
      Y1              =   240
      Y2              =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "GRACIAS POR EVALUAR FLEXOFT !"
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
      TabIndex        =   4
      Top             =   1365
      Width           =   5790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   $"FINDEMO.frx":113AB
      Height          =   645
      Left            =   210
      TabIndex        =   1
      Top             =   2520
      Width           =   5790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "El presente programa de demostración FLEXOFT (r) se encuentra amparado por la Ley 11723 de Propiedad Intelectual."
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   1995
      Width           =   5790
   End
End
Attribute VB_Name = "FINDEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
    If YAPRELICEN% < 1 Then
      PROPINTEL.Show 1
      Exit Sub
    End If
    OPCIDEMO.Show 1
End Sub

Private Sub Command2_Click()
    Close
    End
End Sub

Private Sub Form_Activate()
   Command1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close
    End
    'Hide
End Sub
