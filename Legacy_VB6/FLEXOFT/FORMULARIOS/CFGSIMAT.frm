VERSION 5.00
Begin VB.Form CFGSIMAT 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Comprobante de Solicitud de Material"
   ClientHeight    =   2985
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
   ScaleHeight     =   2985
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Picture         =   "CFGSIMAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
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
      Picture         =   "CFGSIMAT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Comprobante de Solicitud de Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
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
         Height          =   855
         Left            =   210
         TabIndex        =   8
         Top             =   1575
         Width           =   4800
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime Impresión al Generar el Comprobante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   315
            TabIndex        =   9
            Top             =   315
            Width           =   4320
         End
      End
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
         Picture         =   "CFGSIMAT.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   510
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
         Top             =   1050
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
         Top             =   525
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
         Top             =   525
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
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4980
      Picture         =   "CFGSIMAT.frx":075E
      Top             =   2310
      Width           =   2010
   End
End
Attribute VB_Name = "CFGSIMAT"
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
    Call GRACONTROL("", "FORSIMT", Text6.TEXT)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("FORSIMT", "SUPRINT", SUPRINT$)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
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

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
    Text6.TEXT = CONTROL$("", "FORSIMT")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
    If CONTROL$("FORSIMT", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    '
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub


Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
