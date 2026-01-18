VERSION 5.00
Begin VB.Form CFGNOPE09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Notas de Pedidos"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7230
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
   ScaleHeight     =   1410
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nota de Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   6060
      Begin VB.CommandButton Command7 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   5025
         Picture         =   "CFGNOPE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   195
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   4755
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4770
         TabIndex        =   4
         Top             =   300
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3405
         TabIndex        =   3
         Text            =   " "
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1020
         TabIndex        =   7
         Top             =   360
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   6735
      Left            =   6270
      ScaleHeight     =   6675
      ScaleWidth      =   1050
      TabIndex        =   0
      Top             =   -60
      Width           =   1110
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   150
         Picture         =   "CFGNOPE09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CFGNOPE09.frx":0614
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   960
         Width           =   1515
      End
   End
End
Attribute VB_Name = "CFGNOPE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()

    '
    Call GRACONTROL("", "FORNOPE", Text6.TEXT)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text6.TEXT = CONTROL$("", "FORNOPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
