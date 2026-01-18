VERSION 5.00
Begin VB.Form CFGPREPA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Preparación"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6525
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
   ScaleHeight     =   3225
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Transferencia entre Depósitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   5160
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
         Left            =   2100
         TabIndex        =   14
         Top             =   900
         Width           =   175
      End
      Begin VB.CommandButton Command5 
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
         Left            =   2100
         TabIndex        =   10
         Top             =   450
         Width           =   175
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
         Left            =   1710
         TabIndex        =   15
         Text            =   " "
         Top             =   900
         Width           =   420
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
         Left            =   2355
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   900
         Width           =   2610
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
         Left            =   1710
         TabIndex        =   11
         Text            =   " "
         Top             =   450
         Width           =   420
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
         Height          =   360
         Left            =   2355
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   450
         Width           =   2610
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. Preparación:"
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
         Left            =   -650
         TabIndex        =   16
         Top             =   990
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. Expedición:"
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
         Left            =   -650
         TabIndex        =   12
         Top             =   540
         Width           =   2325
      End
   End
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
      Height          =   855
      Left            =   5520
      Picture         =   "CFGPREPA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   230
      Width           =   855
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
      Height          =   855
      Left            =   5520
      Picture         =   "CFGPREPA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1170
      Width           =   855
   End
   Begin VB.Frame ETIQENV 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Etiqueta de Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   180
      TabIndex        =   2
      Top             =   90
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
         Picture         =   "CFGPREPA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
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
         Height          =   360
         Left            =   210
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   900
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
         TabIndex        =   4
         Top             =   390
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
         Height          =   360
         Left            =   2415
         TabIndex        =   3
         Text            =   " "
         Top             =   390
         Width           =   1380
      End
      Begin VB.Label Label4 
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
         TabIndex        =   7
         Top             =   540
         Width           =   2325
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4950
      Picture         =   "CFGPREPA.frx":075E
      Top             =   2550
      Width           =   2010
   End
End
Attribute VB_Name = "CFGPREPA"
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
    Call GRACONTROL("", "ETIQENV", Text6.TEXT)
    Call GRACONTROL("", "DEPEXPE", Text2.TEXT)
    Call GRACONTROL("", "DEPPREPA", Text4.TEXT)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text3.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text2.TEXT = VDEVU$
        Text1.TEXT = VALACT2$("TADEPOSI")
    End If
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
    Text6.TEXT = CONTROL$("", "ETIQENV")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text2.TEXT = CONTROL$("", "DEPEXPE")
    DEP% = CInt(Text2)
    Text1.TEXT = ECOARCH$("TADEPOSI", Chr$(DEP%))
    Text4.TEXT = CONTROL$("", "DEPPREPA")
    DEP% = CInt(Text4)
    Text3.TEXT = ECOARCH$("TADEPOSI", Chr$(DEP%))
    '
End Sub

Private Sub Text2_Change()
'    Call SELECHO("ECODEP")
'    VDEVU$ = VALACT1$("ECODEP")
'    If VDEVU$ <> "" Then
'        Text2.TEXT = VDEVU$
'        Text1.TEXT = VALACT2$("ECODEP")
'    End If
End Sub

Private Sub Text4_Change()
'    Call SELECHO("ECODEP")
'    VDEVU$ = VALACT1$("ECODEP")
'    If VDEVU$ <> "" Then
'        Text4.TEXT = VDEVU$
'        Text3.TEXT = VALACT2$("ECODEP")
'    End If
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub

Private Sub Text6_DblClick()
'    Call SELECHO("INDIFRM")
'    VDEVU$ = VALACT1$("INDIFRM")
'    If VDEVU$ <> "" Then
'        Text6.TEXT = VDEVU$
'        Text7.TEXT = VALACT2$("INDIFRM")
'    End If
End Sub

