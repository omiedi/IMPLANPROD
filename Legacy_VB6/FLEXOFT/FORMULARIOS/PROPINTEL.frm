VERSION 5.00
Begin VB.Form PROPINTEL 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistemas FLEXOFT - Demostraciones"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   240
      Picture         =   "PROPINTEL.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   5730
      TabIndex        =   11
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
         TabIndex        =   12
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.Frame Frame3 
      Height          =   25
      Left            =   210
      TabIndex        =   9
      Top             =   1785
      Width           =   5790
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   540
      Left            =   210
      TabIndex        =   6
      Top             =   4935
      Width           =   5790
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   10
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "es Marca Registrada propiedad del Ing. Ernesto Pablo Bauer"
         Height          =   270
         Left            =   1260
         TabIndex        =   7
         Top             =   225
         Width           =   4470
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1590
      Left            =   210
      TabIndex        =   2
      Top             =   3255
      Width           =   5790
      Begin VB.CommandButton Command2 
         Caption         =   "No, Cancelar"
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
         Left            =   3400
         TabIndex        =   5
         Top             =   945
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar DEMO"
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
         Left            =   525
         TabIndex        =   4
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "Si está de acuerdo con la condición precedente, haga click sobre 'Iniciar DEMO' para comenzar el ciclo de demostración."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   315
         TabIndex        =   3
         Top             =   315
         Width           =   5160
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
      BackColor       =   &H00FFFFC0&
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
      TabIndex        =   8
      Top             =   1365
      Width           =   5790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   $"PROPINTEL.frx":11302
      Height          =   645
      Left            =   210
      TabIndex        =   1
      Top             =   2520
      Width           =   5790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "El presente programa de demostración FLEXOFT (r) se encuentra amparado por la Ley 11723 de Propiedad Intelectual."
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   1995
      Width           =   5790
   End
End
Attribute VB_Name = "PROPINTEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    YAPRELICEN% = 1
    On Error Resume Next
    LODEMO& = FileLen("C:\DEMOFLEX.DRV")
    If LODEMO& > 0 Then GoTo 10
    '
    N1% = FreeFile
    Open "C:\DEMOFLEX.DRV" For Output As #N1%
    Print #N1%, Format$(Now, "dd-mm-yyyy hh:mm:ss")
    Close #N1%
    On Error GoTo 0
    Hide
    Call SECUDEMO("ANACOS")
    Exit Sub
    '
10  On Error GoTo 0
    Hide
    OPCIDEMO.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DEMOFLE1.Command1.Visible = True
End Sub
