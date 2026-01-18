VERSION 5.00
Begin VB.Form OPALMOD09 
   BackColor       =   &H00F0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formularios Almacen de Modelos"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6510
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
   ScaleHeight     =   5115
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   6180
      Left            =   5550
      ScaleHeight     =   6120
      ScaleWidth      =   1395
      TabIndex        =   18
      Top             =   -750
      Width           =   1455
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
         Height          =   660
         Left            =   120
         Picture         =   "OPALMOD09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1050
         Width           =   700
      End
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
         Height          =   660
         Left            =   120
         Picture         =   "OPALMOD09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1710
         Width           =   700
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -435
         Picture         =   "OPALMOD09.frx":0454
         Top             =   7860
         Width           =   2010
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -280
         Picture         =   "OPALMOD09.frx":3A3E
         Top             =   5280
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0E0E0&
      Caption         =   "Recepción de Modelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   12
      Top             =   3465
      Width           =   5160
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
         Height          =   360
         Left            =   2415
         TabIndex        =   16
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.CommandButton Command8 
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
         Height          =   360
         Left            =   3780
         TabIndex        =   15
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPALMOD09.frx":5960
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   17
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F0E0E0&
      Caption         =   "Remito a Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   6
      Top             =   1785
      Width           =   5160
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPALMOD09.frx":5C6A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
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
         Height          =   360
         Left            =   3780
         TabIndex        =   8
         Top             =   420
         Width           =   175
      End
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
         Height          =   360
         Left            =   2415
         TabIndex        =   7
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   11
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F0E0E0&
      Caption         =   "Ficha de Inspeccion de Modelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   5160
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
         Height          =   360
         Left            =   2415
         TabIndex        =   5
         Text            =   " "
         Top             =   420
         Width           =   1380
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
         Height          =   360
         Left            =   3780
         TabIndex        =   4
         Top             =   420
         Width           =   175
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
         Height          =   360
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
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
         Height          =   880
         Left            =   4095
         Picture         =   "OPALMOD09.frx":5F74
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   1
         Top             =   525
         Width           =   2325
      End
   End
End
Attribute VB_Name = "OPALMOD09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FICHINMO", Text1.TEXT)
    Call GRACONTROL("", "REMIMODE", Text3.TEXT)
    Call GRACONTROL("", "RECEMODE", Text6.TEXT)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub


Private Sub Command7_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FICHINMO")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
    Text3.TEXT = CONTROL$("", "REMIMODE")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    '
    Text6.TEXT = CONTROL$("", "RECEMODE")
    Text5.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text6_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
