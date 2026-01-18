VERSION 5.00
Begin VB.Form WINMAIL04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Mail"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   14
      Top             =   840
      Width           =   4950
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   945
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1890
      Width           =   4950
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   945
      TabIndex        =   5
      Top             =   1365
      Width           =   4950
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   315
      Width           =   4950
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   6090
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "WINMAIL04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "WINMAIL04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Aprueba y Envía el Mail "
         Top             =   3465
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   4305
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "WINMAIL04.frx":0454
         Top             =   4620
         Width           =   1515
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adjunto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   9
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   945
      TabIndex        =   8
      Top             =   4620
      Width           =   4950
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   7
      Top             =   1890
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asunto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   1425
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dir-El.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   3
      Top             =   945
      Width           =   645
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   1
      Top             =   420
      Width           =   645
   End
End
Attribute VB_Name = "WINMAIL04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command14_Click()
   '
   '20/04/07 (OT-07-0158)
   If Text2 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Asunto.")
      Text2.SetFocus
      Exit Sub
   End If
   '
   If Text5 = "" Then
      Call MENSERR(24, "Imposible Enviar el Mensaje.\Debe Ingresar un Texto.")
      Text5.SetFocus
      Exit Sub
   End If
   '20/04/07 (FIN)
   '
   MAI_DESTI$ = TRIM$(Text1)
   MAI_DIREL$ = Text3                 ' "ebauer@flexoft.com.ar"
   MAI_ASUNT$ = TRIM$(Text2)
   MAI_TEXTO$ = Text5
   MAI_ADJUN$(0) = Label5
   '
   Call GENERAMAIL
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command22_Click()
   Hide '18/04/07 (OT-07-0158)
End Sub

