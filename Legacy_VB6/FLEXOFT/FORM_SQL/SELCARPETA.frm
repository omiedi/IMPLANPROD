VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SELCARPETA 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Carpeta"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10800
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   12600
      ScaleHeight     =   8235
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   -75
      Width           =   1245
      Begin VB.CommandButton command1 
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
         Height          =   700
         Left            =   180
         Picture         =   "SELCARPETA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar"
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
         Left            =   180
         Picture         =   "SELCARPETA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Trasla."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "SELCARPETA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   7920
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "SELCARPETA.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   7920
         Width           =   800
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   -405
         Picture         =   "SELCARPETA.frx":0DC8
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1850
         Width           =   2010
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      Begin MSComDlg.CommonDialog cmd 
         Left            =   1200
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text6 
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
         Height          =   1155
         Left            =   100
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "SELCARPETA.frx":43B2
         ToolTipText     =   "Doble Click Para Borrar Rua de  Carpeta"
         Top             =   480
         Width           =   12000
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Selección de Carpeta"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4515
      End
   End
End
Attribute VB_Name = "SELCARPETA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONTROL1$
Public CONTROL2$

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub


Private Sub Command2_Click()
    '
    'CONTROL1 Y 2 VARIABLES QUE SE ASIGNAN CUANDO SE LLAMA AL FORM SELCARPETA PARA QUE QUEDE A  MODO GENERAL
    Call GRACONTROL(CONTROL1$, CONTROL2$, Text6)
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   SELCARP07.Show 1
   CARPETAX$ = SELCARP07.Label1
   If CARPETAX$ <> "" Then
      Text6 = CARPETAX$
   End If
   Command3.Enabled = True
   Screen.MousePointer = 1
End Sub
'

Private Sub Form_Load()
    '
'    EJEMPLO CUANDO SE LLAMA DE OTRO FORMULARIO
'    SELCARPETA.CONTROL1 = "CARPIMAG"
'    SELCARPETA.CONTROL2 = "UBICIMAG"
'    SELCARPETA.Show 1

    Text6 = CONTROL(CONTROL1$, CONTROL2$)
    '
End Sub

Private Sub Text6_DblClick()
  Text6 = ""
End Sub
