VERSION 5.00
Begin VB.Form REMIMODE 
   BackColor       =   &H00FFF0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitir Modelo a Proveedor"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7455
      MaxLength       =   5
      TabIndex        =   16
      Text            =   " "
      Top             =   210
      Width           =   855
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
      Height          =   650
      Left            =   8660
      Picture         =   "REMIMODE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cierra y Sale de la Aplicación"
      Top             =   210
      Width           =   650
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
      Height          =   650
      Left            =   8660
      Picture         =   "REMIMODE.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancela la Pre-Selección"
      Top             =   3045
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
      Left            =   8660
      Picture         =   "REMIMODE.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Aprueba la Selección de Items a Comprar"
      Top             =   3675
      Width           =   650
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   630
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   7680
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remitimos a Ustedes:"
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
      TabIndex        =   15
      Top             =   1420
      Width           =   2115
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7980
      Picture         =   "REMIMODE.frx":075E
      Top             =   4410
      Width           =   2010
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
      Caption         =   "Número:"
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
      Left            =   6195
      TabIndex        =   8
      Top             =   270
      Width           =   1170
   End
End
Attribute VB_Name = "REMIMODE"
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
   Call PRIREMITO
   '
   CODMODE$ = AJUSTI$(ALTAMODE.Text3, 16)
   TTXX$ = Space$(76)
   TIREGI$ = "Despacho a Proveedor - Remito " + TRIM$(Text7)
   Call REPLA(TTXX$, TIREGI$, 11, 56)
   Call REPLA(TTXX$, Text9, 1, 8)
   Call REPLA(TTXX$, USERNAME$, 69, 24)
   Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
   TTXX$ = Space$(76)
   Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
   '
   TTXX$ = Space$(76)
   Call REGAPP("ANOTAMOD", CODMODE$ + "            Destino: " + Text2)
   Call REGAPP("ANOTAMOD", CODMODE$ + "                     " + Text3)
   Call REGAPP("ANOTAMOD", CODMODE$ + "                     " + TRIM$(TRIM$(Text4) + " " + TRIM$(Text5)))
   Call REGAPP("ANOTAMOD", CODMODE$ + String$(88, "="))
   Call CIERRARCH("ANOTAMOD")
   '
   Command1.Enabled = True
   Hide
   '
99 Command1.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   '
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
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

Private Sub PRIREMITO()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "REMIMODE"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Remito de Entrega de Modelo.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
    DATRANSPO$ = DTP$
    '
    Screen.MousePointer = 11
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CACOLTAB1% = 3
    '
    TETABU1$(1, 1) = TRIM$(ALTAMODE.Text3) ' Codigo
    TETABU1$(3, 1) = "1" ' Cantidad
    '
    CAITAB1% = 0
    Call FRATEXTO(Text6.TEXT, 50)
    For KKI% = 1 To CARETEX%
      CAITAB1% = CAITAB1% + 1
      Call REPLA(TTXX$, RETEX$(KKI%), 1, 50)
      TETABU1$(2, CAITAB1%) = TTXX$
      If CAITAB1% > 1 Then
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(3, CAITAB1%) = ""
      End If
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
    CODPARAM$(10) = "TRANSPOR": DOCPARAM$(10) = TRIM$(Mid$(DTP$, 1, 32))
    CODPARAM$(11) = "DOMI-TRA": DOCPARAM$(11) = TRIM$(Mid$(DTP$, 33, 32))
    CODPARAM$(12) = "CUIT-TRA": DOCPARAM$(12) = TRIM$(Mid$(DTP$, 65, 16))
    CODPARAM$(13) = "PATE-TRA": DOCPARAM$(13) = TRIM$(Mid$(DTP$, 81, 16))
    CODPARAM$(14) = "COND-TRA": DOCPARAM$(14) = TRIM$(Mid$(DTP$, 97, 32))
    CODPARAM$(15) = "DOCUCOND": DOCPARAM$(15) = TRIM$(Mid$(DTP$, 129, 24))
    CAPARDOC% = 15
    '
    FORIPRE$ = TRIM$(CONTROL$("", "REMIMODE"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("REMIMODE")
    End If
    '
    Screen.MousePointer = 0
    '
End Sub
