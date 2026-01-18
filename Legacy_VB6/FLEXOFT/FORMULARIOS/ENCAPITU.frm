VERSION 5.00
Begin VB.Form ENCAPITU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demostración FLEXOFT"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   210
      ScaleHeight     =   795
      ScaleWidth      =   5730
      TabIndex        =   2
      Top             =   3570
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
         Left            =   1995
         TabIndex        =   6
         Top             =   525
         Width           =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " 4541-1515"
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
         Left            =   4060
         TabIndex        =   4
         Top             =   260
         Width           =   1590
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel: 4524-1284"
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
         Left            =   4060
         TabIndex        =   3
         Top             =   80
         Width           =   1590
      End
   End
   Begin VB.Label Label5 
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   6000
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   3150
      Y2              =   4000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2220
      Left            =   315
      TabIndex        =   1
      Top             =   1155
      Width           =   5685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capítulo 1:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   5370
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "ENCAPITU.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8670
   End
End
Attribute VB_Name = "ENCAPITU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If MINIDEM% = 1 Then
       Picture1.Picture = PROPINTEM.Picture1.Picture
     Else
       Picture1.Picture = PROPINTEL.Picture1.Picture
   End If
End Sub

Private Sub Timer1_Timer()
   If Timer > Val(Label5.Caption) Then
     If ENCAPITU.Visible = True Then
       On Error Resume Next
       Hide
       On Error GoTo 0
     End If
   End If
End Sub
