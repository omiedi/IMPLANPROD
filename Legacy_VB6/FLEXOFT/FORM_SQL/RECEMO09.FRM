VERSION 5.00
Begin VB.Form RECEMO09 
   BackColor       =   &H00D0D0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibir Modelo de Proveedor"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   3075
      Left            =   8340
      ScaleHeight     =   3015
      ScaleWidth      =   1245
      TabIndex        =   14
      Top             =   -30
      Width           =   1305
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
         Height          =   660
         Left            =   120
         Picture         =   "RECEMO09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Aprueba recepción del modelo"
         Top             =   1650
         Width           =   700
      End
      Begin VB.CommandButton BOTEXIT 
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
         Picture         =   "RECEMO09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   1005
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "RECEMO09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cierra y Sale de la Aplicación"
         Top             =   210
         Width           =   700
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -270
         Picture         =   "RECEMO09.frx":075E
         Top             =   2370
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -435
         Picture         =   "RECEMO09.frx":2680
         Top             =   7860
         Width           =   2010
      End
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      MaxLength       =   5
      TabIndex        =   12
      Text            =   " "
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      MaxLength       =   5
      TabIndex        =   7
      Text            =   " "
      Top             =   870
      Width           =   750
   End
   Begin VB.CommandButton Command2 
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
      Height          =   285
      Left            =   1365
      TabIndex        =   6
      Top             =   870
      Width           =   195
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   1830
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   " "
      Top             =   210
      Width           =   4485
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      Left            =   1830
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   " "
      Top             =   540
      Width           =   4485
   End
   Begin VB.TextBox Text4 
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
      Height          =   285
      Left            =   1830
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   750
   End
   Begin VB.TextBox Text5 
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
      Height          =   285
      Left            =   2670
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   " "
      Top             =   870
      Width           =   3645
   End
   Begin VB.CommandButton Command9 
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
      Height          =   285
      Left            =   8085
      TabIndex        =   1
      Top             =   840
      Width           =   195
   End
   Begin VB.TextBox Text9 
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
      Height          =   285
      Left            =   7140
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   " "
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1830
      TabIndex        =   13
      Top             =   1920
      Width           =   6420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Recibimos de Ustedes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   210
      TabIndex        =   11
      Top             =   1560
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   210
      TabIndex        =   10
      Top             =   525
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6345
      TabIndex        =   9
      Top             =   900
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   6120
      TabIndex        =   8
      Top             =   270
      Width           =   1170
   End
End
Attribute VB_Name = "RECEMO09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   NPRO% = Val(Text1)
   If REGICLI&((-1) * NPRO%) < 1 Then
     Call MENSERR(24, "Seleccione Proveedor")
     Call Command2_Click
     GoTo 99
   End If
   '
   If Val(Text7) < 1 Then
     Call MENSERR(24, "Ingrese Número de Remito")
     GoTo 99
   End If
   '
   ' imprimir remito
   Call PRIRECEPMOD
   '
   CODMODE$ = AJUSTI$(ALTMOD09.Text3, 16)
   TTXX$ = Space$(76)
   TIREGI$ = "Recepción de Proveedor - Remito " + TRIM$(Text7)
   Call REPLA(TTXX$, TIREGI$, 11, 56)
   Call REPLA(TTXX$, Text9, 1, 8)
   Call REPLA(TTXX$, USERNAME$, 69, 24)
   ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
   TTXX$ = Space$(76)
   ' Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, TTXX$)
   '
   TTXX$ = Space$(76)
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, "            Origen: " + Text2)
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, "                     " + Text3)
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, "                     " + TRIM$(TRIM$(Text4) + " " + TRIM$(Text5)))
   Call ALTMOD09.AGREGA_ANOTAMODELO(CODMODE$, String$(88, "="))
   ' Call CIERRARCH("ANOTAMOD")
   '
   Command1.Enabled = True
   Hide
   '
99 Command1.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   '
   Call SELECHO("PROVED1")
   NCLIE% = Val(TRIM$(VALACT1$("PROVED1")))
   If NCLIE% > 0 Then
     Text1 = TRIM$(Str$(NCLIE%))
   End If
   '
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Form_Load()
    '
    Call BLANFORMU
    '
End Sub

Private Sub Text1_Change()
    '
    NCLIE% = Val(Text1)
    Call BLANFORMU
    If REGICLI&((-1) * NCLIE%) > 0 Then
      Call CARDACLI
    End If
    '
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub

Sub CARDACLI()
   '
   Screen.MousePointer = 11
   '
   NC% = (-1) * Val(Text1.TEXT)
   If NC% < 0 Then
       Label8 = CODICLI$(NC%)
       Text2.TEXT = RASOCLI$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
   '
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7 = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   '
End Sub

Private Sub Text11_GotFocus()
   On Error Resume Next 'x err '5' llamada a procedimiento no valida.
   Text1.SetFocus
   Screen.MousePointer = 1
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   Text1.SetFocus
End Sub

Private Sub PRIRECEPMOD()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "RECEMODE"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Recepción de Modelo.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    CODTABU1$(1) = "DES-MAT"
    CACOLTAB1% = 1
    '
    CAITAB1% = 0
    Call FRATEXTO(Label4, 60)
    For KKI% = 1 To CARETEX%
      CAITAB1% = CAITAB1% + 1
      Call REPLA(TTXX$, RETEX$(KKI%), 1, 60)
      TETABU1$(1, CAITAB1%) = TTXX$
    Next KKI%
    '
    CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = TRIM$(Text7)
    NC% = Val(Text1)
    If ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
      Screen.MousePointer = 0
      Exit Sub
    End If
    NC% = (-1) * NC%
    CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = RASOCLI$(NC%)
    CODPARAM$(3) = "CUEN-CLI": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "DOMI-CLI": DOCPARAM$(4) = DOMICLI$(NC%)
    CODPARAM$(5) = "CPOS-CLI": DOCPARAM$(5) = CPOSCLI$(NC%)
    CODPARAM$(6) = "LOCA-CLI": DOCPARAM$(6) = LOCACLI$(NC%)
    CODPARAM$(7) = "PIVA-CLI": DOCPARAM$(7) = POSIVAL$(NC%)
    CODPARAM$(8) = "CUIT-CLI": DOCPARAM$(8) = CUITCLI$(NC%)
    CODPARAM$(9) = "FECH-EMI": DOCPARAM$(9) = TRIM$(Text9)
    CAPARDOC% = 9
    '
    FORIPRE$ = TRIM$(CONTROL$("", "RECEMODE"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("RECEMODE")
    End If
    '
    Screen.MousePointer = 0
    '
End Sub

