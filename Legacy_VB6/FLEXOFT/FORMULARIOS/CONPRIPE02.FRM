VERSION 5.00
Begin VB.Form CONPRIPE02 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5235
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
      Left            =   4590
      Picture         =   "CONPRIPE02.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Width           =   4395
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
         Left            =   2950
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Imprime O/Despacho"
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
         Left            =   1400
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   1560
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Imprime O/Venta"
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
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1140
      End
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "CONPRIPE02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

