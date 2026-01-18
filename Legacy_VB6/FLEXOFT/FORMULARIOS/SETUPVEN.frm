VERSION 5.00
Begin VB.Form SETUPVEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Pedidos Electronicos"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Dirección Electronica Mails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   4215
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   6
         Top             =   840
         Width           =   3795
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   3795
      End
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4620
      Picture         =   "SETUPVEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guarda los Datos Ingresados"
      Top             =   945
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4620
      Picture         =   "SETUPVEN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4215
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   3795
      End
   End
   Begin VB.Line Line1 
      X1              =   4410
      X2              =   5565
      Y1              =   2615
      Y2              =   2615
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4020
      Picture         =   "SETUPVEN.frx":0454
      Top             =   2100
      Width           =   2010
   End
End
Attribute VB_Name = "SETUPVEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Cancel = 0
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command29_Click()
   '
   Command29.Enabled = False
   Call GRACONTROL("VENREMOT", "VENDEDOR", Text1)
   Call GRACONTROL("VENREMOT", "DESTIMAI", Text2)
   Call GRACONTROL("VENREMOT", "DIRELMAI", Text3)
   Command29.Enabled = True
   '
End Sub

Private Sub Form_Load()
   '
   Text1 = TRIM$(CONTROL$("VENREMOT", "VENDEDOR"))
   Text2 = TRIM$(CONTROL$("VENREMOT", "DESTIMAI"))
   Text3 = TRIM$(CONTROL$("VENREMOT", "DIRELMAI"))
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Hide
End Sub

