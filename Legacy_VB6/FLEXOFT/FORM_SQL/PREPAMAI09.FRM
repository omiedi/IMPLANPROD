VERSION 5.00
Begin VB.Form PREPAMAI09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prepara Envio por Mail"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   5985
      TabIndex        =   9
      ToolTipText     =   "Lista Completa de Contactos"
      Top             =   210
      Width           =   150
   End
   Begin VB.CommandButton COMMAND16 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   2730
      TabIndex        =   8
      ToolTipText     =   "Lista Completa de Contactos"
      Top             =   210
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4250
      TabIndex        =   6
      Top             =   1050
      Width           =   1905
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1050
      Width           =   2835
   End
   Begin VB.Label Label3 
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   2835
      Width           =   4740
   End
   Begin VB.Label Label2 
      Height          =   330
      Left            =   105
      TabIndex        =   10
      Top             =   2205
      Width           =   3795
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label LABEL52 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   555
      Width           =   4950
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prospecto:"
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
      Left            =   105
      TabIndex        =   4
      Top             =   270
      Width           =   960
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   210
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Oportunidad:"
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
      Left            =   2940
      TabIndex        =   2
      Top             =   270
      Width           =   1380
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4410
      TabIndex        =   1
      Top             =   210
      Width           =   1575
   End
End
Attribute VB_Name = "PREPAMAI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If Label26 = "" Then
       Call MENSERR(24, "Obligatorio Identificar Prospecto")
       GoTo 99
    End If
    '
    If Combo1.TEXT = "" Then
      If Combo1.ListCount > 0 Then
        Call MENSERR(24, "Obligatorio Identificar Destino")
        GoTo 99
      End If
    End If
    '
    FISECOM09.Label26 = Label26
    Label2 = FISECOM09.Text5
    Label3 = ""
    For KKI% = 1 To FISECOM09.Text6.Count
      If TRIM$(FISECOM09.Text6(KKI% - 1)) <> "" Then
        If FISECOM09.Text6(KKI% - 1) = Combo1.TEXT Then
          Label3 = TRIM$(FISECOM09.Text8(KKI% - 1).TEXT)
          GoTo 80
        End If
      End If
    Next KKI%
    '
80  Label1 = "OK"
    Command1.Enabled = True
    Hide
    Exit Sub
    '
99  Command1.Enabled = True
    '
End Sub

Private Sub COMMAND16_Click()
   ABUSX$ = "INDIPROS" ' : If USNBR% < 100 Then ABUSX$ = "BSEGALE"
21 Call SELECHO(ABUSX$ & "/Indice Prospect")
   VDEVU$ = VALACT1$(ABUSX$)
   If VDEVU$ <> "" Then
      Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command35_Click()

End Sub

Private Sub Command2_Click()
    Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
    CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
    If CODIOPO$ <> "" Then
       Label10 = CODIOPO$
    End If
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
End Sub

Private Sub Label26_CHANGE()
    '
    Label10 = ""
    LABEL52 = VALOBADA("BSEGCOM", "Raso_Cli", "CODI_CLI = '" & Label26 & "'")
    CONDI$ = "CodOport <>'' AND CODI_CLI = '" & Label26 & "'"
    Call CARGALISEL("!INDOPORT", "OPORTUNI", "CodOport/A12;RefeOport/A48", CONDI$, "NOMESS")
    '
    Call FISECOM09.CARLICONTAC(Combo1, Label26)
    '
End Sub
