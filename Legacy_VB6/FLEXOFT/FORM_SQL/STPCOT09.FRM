VERSION 5.00
Begin VB.Form STPCOT09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Impresión de  Cotizaciones"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5730
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
   ScaleHeight     =   1785
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   4560
      Begin VB.CommandButton Command6 
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
         Height          =   810
         Index           =   0
         Left            =   3690
         Picture         =   "STPCOT09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   420
         Width           =   735
      End
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
         Height          =   345
         Left            =   2085
         TabIndex        =   7
         Top             =   450
         Width           =   1320
      End
      Begin VB.CommandButton Command3 
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
         Height          =   330
         Left            =   3390
         TabIndex        =   4
         Top             =   450
         Width           =   175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   3405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Cotizacion:"
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
         Left            =   -420
         TabIndex        =   5
         Top             =   550
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404000&
      Height          =   3645
      Left            =   4800
      ScaleHeight     =   3585
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "STPCOT09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   110
         Picture         =   "STPCOT09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
   End
End
Attribute VB_Name = "STPCOT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command2_Click()
    Call GRACONTROL("", "FORCOTI", Text1.TEXT) 'FALATA AGREGAR EL CONTROL
    '
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
    Call CommSel("INDIFRM", Command3, Text1)
End Sub

Private Sub Command6_Click(Index As Integer)
    '
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    '
End Sub

Private Sub Form_Load()
    Text1.TEXT = CONTROL$("", "FORCOTI") 'FALTA AGREGAR EL CONTROL
End Sub

Private Sub Text1_Change()
    Label3 = TRIM$(ECOARCH$("INDIFRM", Text1))
End Sub

Private Sub Text1_DblClick()
    Call Command3_Click
End Sub
