VERSION 5.00
Begin VB.Form COPRISIM 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   875
      Left            =   4440
      Picture         =   "COPRISIM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
      Top             =   300
      Width           =   540
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Impresión"
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
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   4155
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Impresión Fantasma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2640
         TabIndex        =   2
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Imprime Solicitud de Materiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   360
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1980
      End
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "COPRISIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

Private Sub Form_Load()
    '
    Check7.Value = 1
    Check2.Value = 0
    If CONTROL$(IDENTER$, "IMPSOLMA") = "SI" Then Check2.Value = 1
    '
End Sub
