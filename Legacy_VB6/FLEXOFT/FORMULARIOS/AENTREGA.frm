VERSION 5.00
Begin VB.Form AENTREGA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizaciones de Entrega según Programa de Fabricación"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   5100
      TabIndex        =   8
      Top             =   210
      Width           =   5160
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedores Seleccionados"
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
         Left            =   210
         TabIndex        =   9
         Top             =   70
         Width           =   4425
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5565
      Picture         =   "AENTREGA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Invierte Selección de Proveedores"
      Top             =   2100
      Width           =   750
   End
   Begin VB.ListBox List2 
      Height          =   1740
      Left            =   210
      TabIndex        =   6
      Top             =   4620
      Visible         =   0   'False
      Width           =   6210
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5565
      Picture         =   "AENTREGA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Emite Autorizaciones de Entregas para los Proveedores Marcados"
      Top             =   2940
      Width           =   750
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   210
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   630
      Width           =   5160
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   5565
      ScaleHeight     =   75
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   1890
      Width           =   750
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
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
      Height          =   750
      Left            =   5565
      Picture         =   "AENTREGA.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5565
      Picture         =   "AENTREGA.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Line Line2 
      X1              =   210
      X2              =   5355
      Y1              =   4515
      Y2              =   4515
   End
   Begin VB.Line Line1 
      X1              =   5370
      X2              =   5370
      Y1              =   210
      Y2              =   4515
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4980
      Picture         =   "AENTREGA.frx":08A8
      Top             =   4020
      Width           =   2010
   End
End
Attribute VB_Name = "AENTREGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim CI%(32767)
 Dim PROVE%(32767)
 Dim UEC%(32767)
 Dim RUEC&(32767)
 Dim PUNX$(32767)
 '
 Dim YACAPRO%(32767)
 '
 Dim A$(10), SP%(10)
 Dim INIPER%(12)
 Dim SUMA#(16)
 Dim ENCA$, ENCB$
 '
 Dim MX%
 Dim FORMA% '  FORMA% = COMALTER%("Elija Formato de Impresion\\Por Proveedor\Por Material\Abandonar")
 Dim MODO%  '  MODO% = COMALTER%("Presentacion de Valores\\Cantidades\Importes\Abandonar")
 '

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command2_Click()
   '
   UREPLA& = ULTREG&("PLACOMP")
   '
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       PROVEI% = Val(Mid$(List1.List(U& - 1), 41))
       For KKI% = 1 To MX%
         PUNX$(KKI%) = ""
       Next KKI%
       RFF$ = RECFILT$("PLACOMP", 8, MKI$(PROVEI%))
       If Len(RFF$) > 0 Then
         '
         For KL& = 1 To Len(RFF$) Step 4
           RE& = CVS(Mid$(RFF$, KL&, 4))
           If RE& > 0 Then
             If RE& <= UREPLA& Then
               X$ = REGLEIDO$("PLACOMP", RE&)
               CII% = CVI(Mid$(X$, 1, 2))
               If CII% > 0 Then
                 If CII% <= MX% Then
                   PUNX$(CII%) = PUNX$(CII%) + MKS$(RE&)
                 End If
               End If
             End If
           End If
         Next KL&
         '
         Call GENEAUTEN
         If List2.ListCount > 0 Then
           Call PRINAUTEN(PROVEI%)
         End If
         '
       End If
     End If
   Next U&
   '
End Sub

Sub GENEAUTEN()
   '
   List2.Clear
   For CII% = 1 To MX%
     If PUNX$(CII%) <> "" Then
       TTXX$ = Space$(256)
       TTYY$ = "Codigo: " + TRIM$(CODEXT$(CII%))
       Call REPLA(TTXX$, TTYY$, 1, 36)
       Call REPLA(TTXX$, DESCRIT0$(CII%), 37, 36)
       '
       UEX = 0: BOLRE$ = "": REMI$ = "": CAN = 0
       If RUEC&(CII%) > 0 Then
         X$ = REGLEIDO$("MOVISTO", RUEC&(CII%))
         UEX = CVI(Left$(X$, 2))
         BOLRE$ = TRIM$(Mid$(X$, 23, 10))
         REMI$ = TRIM$(Mid$(X$, 33, 10))
         CAN = CVD(Mid$(X$, 88, 8))
       End If
       '
       R1$ = Space$(48): R2$ = Space$(48)
       If CAN > 0 Then
         UNIX$ = TRIM$(UNIMED$(CII%))
         If Left$(UNIX$, 1) = "U" Then UNIX$ = "Pzas."
         Call REPLA(R1$, " U.Entre:" + Str$(CAN) + " " + UNIX$ + " el " + FECHATEX$(UEX), 1, 40)
         Call REPLA(R2$, "          RT." + REMI$ + " - " + BOLRE$, 1, 40)
       End If
       Call REPLA(TTXX$, R1$, 73, 36)
       Call REPLA(TTXX$, R2$, 109, 36)
       '
       R3$ = Space$(112)
       For J% = 1 To Len(PUNX$(CII%)) Step 4
         RECNO& = CVS(Mid$(PUNX$(CII%), J%, 4))
         Y$ = REGLEIDO$("PLACOMP", RECNO&)
         FENTRE% = CVI(Mid$(Y$, 3, 2))
         If FENTRE% < INIPER%(1) Then FENTRE% = INIPER%(1)
         CANTU = CVS(Mid$(Y$, 21, 4))
         '
         For K% = 1 To 8
           If FENTRE% >= INIPER%(K%) Then
             If FENTRE% < INIPER%(K% + 1) Then
                CANTU = Int(CANTU + Val(Mid$(R3$, 10 * K% - 9, 10)) + 0.5)
                CANX$ = CSTRING$(MKS$(CANTU), 3, 10, 1, 2)
                Call REPLA(R3$, CANX$, 10 * K% - 9, 10)
             End If
           End If
         Next K%
       Next J%
       '
       Call REPLA(TTXX$, R3$, 145, 112)
       List2.AddItem TTXX$
       '
     End If
   Next CII%
   '
End Sub
'
Sub PRINAUTEN(PROVEI%)
   '
   NC1% = (-1) * PROVEI%
   '
   FORIPRE$ = "AENT-FAI"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
     Call MENSERR(24, "No Existe Formulario de Impresión '" + FORIPRE$ + "'.")
     Call FINAL("")
     Exit Sub
   End If
   '
   Call SETULTREG("!OBSERVOF", 0)
   Call CIERRARCH("!OBSERVOF")
   '
   FEX = HOY(HOS$)
   HOII% = FEX
   FECHDOC$ = HOS$
   NUMEDOC$ = ""
   TIPODOC$ = "Autorización de Entregas"
   DESTIDOC% = NC1%
   '
   CODPARAM$(1) = "FECH-LAN"
   CODPARAM$(2) = "REF-MAT1"
   CODPARAM$(3) = "REF-MAT2"
   CAPARDOC% = 3
   '
   DOCPARAM$(1) = "J.L.Suarez (B.A.), " + FETEXTO$(HOII%)
   DOCPARAM$(2) = ENCA$
   DOCPARAM$(3) = ENCB$
   '
   CODTABU1$(1) = "N-HPROF"
   CODTABU1$(2) = "N-HINSP"
   CODTABU1$(3) = "N-ORIGIN"
   CODTABU1$(4) = "N-OCOMPR"
   CODTABU1$(5) = "N-AUTENT"
   CODTABU1$(6) = "N-REMITO"
   CODTABU1$(7) = "N-BOLREC"
   CODTABU1$(8) = "POR-IVA"
   '
   CODTABU1$(9) = "MEDIDA-1"
   CODTABU1$(10) = "DES-MAT"
   CODTABU1$(11) = "MEDIDA-2"
   CODTABU1$(12) = "MEDIDA-3"
   CACOLTAB1% = 12
   '
   CAITAB1% = 0
   For U& = 1 To List2.ListCount
     TZ$ = List2.List(U& - 1)
     CAITAB1% = CAITAB1% + 1
     TETABU1$(9, CAITAB1%) = Left$(TZ$, 36)
     TETABU1$(10, CAITAB1%) = Mid$(TZ$, 37, 36)
     TETABU1$(11, CAITAB1%) = Mid$(TZ$, 73, 36)
     TETABU1$(12, CAITAB1%) = Mid$(TZ$, 109, 36)
     '
     Z$ = Mid$(TZ$, 145)
     TETABU1$(1, CAITAB1%) = Mid$(Z$, 1, 10)
     TETABU1$(2, CAITAB1%) = Mid$(Z$, 11, 10)
     TETABU1$(3, CAITAB1%) = Mid$(Z$, 21, 10)
     TETABU1$(4, CAITAB1%) = Mid$(Z$, 31, 10)
     TETABU1$(5, CAITAB1%) = Mid$(Z$, 41, 10)
     TETABU1$(6, CAITAB1%) = Mid$(Z$, 51, 10)
     TETABU1$(7, CAITAB1%) = Mid$(Z$, 61, 10)
     TETABU1$(8, CAITAB1%) = Mid$(Z$, 71, 10)
     '
   Next U&
   '
   Call PRINTDOC("@" + FORIPRE$)
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command3_Click()
   TTP = List1.TopIndex
   TTQ = List1.ListIndex
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
         List1.Selected(U& - 1) = False
       Else
         List1.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   List1.ListIndex = TTQ
   List1.TopIndex = TTP
   On Error GoTo 0
   OPENFORMS = DoEvents
End Sub

'
Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Show
    Refresh
    '
    Screen.MousePointer = 11
    Call ENCASEMA
    MX% = NUMAS%
    FORMA% = 1  ' POR PROVEEDOR
    MODO% = 1   ' SOLO CANTIDADES
    For I% = 1 To NUMAS%
      CI%(I%) = I%
    Next I%
    Screen.MousePointer = 1
    '
    Call CARPROVHA
    Call FULTIENT
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub CARPROVHA()
   '
   FIN& = ULTREG&("PLACOMP")
   Screen.MousePointer = 11
   For I& = 1 To FIN&
     Call TRACE(21, I&, FIN&)
     X$ = REGLEIDO$("PLACOMP", I&)
     CII% = CVI(Left$(X$, 2))
     PROVEI% = CVI(Mid$(X$, 33, 2))
     If CII% > 0 Then
       If CII% <= MX% Then
         If PROVE%(CII%) = 0 Then
           PROVE%(CII%) = PROVEI%
         End If
       End If
     End If
     '
     If PROVEI% > 0 Then
       If YACAPRO%(PROVEI%) < 1 Then
         TTXX$ = Space$(48)
         Call REPLA(TTXX$, RASOCLI$((-1) * PROVEI%), 1, 40)
         Call REPLA(TTXX$, Str$(PROVEI%), 41, 8)
         List1.AddItem TTXX$
         YACAPRO%(PROVEI%) = 1
         OPENFORMS = DoEvents
       End If
     End If
   Next I&
   Screen.MousePointer = 1
   '
End Sub
'
Sub FULTIENT()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("MOVISTO")
   For I& = 1 To FIN&
     Call TRACE(21, I&, FIN&)
     X$ = REGLEIDO$("MOVISTO", I&)
     CII% = CVI(Mid$(X$, 3, 2))
     FENTRE% = CVI(Left$(X$, 2))
     If CII% > 0 Then
       If CII% <= MX% Then
         If FENTRE% >= UEC%(CII%) Then
           If CVD(Mid$(X$, 88, 8)) > 0.005 Then
             If CVI(Mid$(X$, 86, 2)) = PROVE%(CII%) Then
               UEC%(CII%) = FENTRE%
               RUEC&(CII%) = I&
             End If
           End If
         End If
       End If
     End If
   Next I&
   Screen.MousePointer = 1
   '
End Sub
'
Sub ENCASEMA()
   '
   HO% = HOY(HOS$)
   FF& = HO%: SEMANACT% = SEMANU%(FF&, A1, B1)
   CACOL% = 9
   '
   FFX& = HO%: SEMAHO% = SEMANU%(FFX&, A2, B2)
   DIBAS = A2: SMNA% = SEMAHO%
   J% = 1
   INIPER%(1) = DIBAS
   ENCA$ = AJUSTD$("S." + TRIM$(Str$(SMNA%)), 10)
   ENCB$ = "  " + Left$(FECHATEX$(DIBAS), 5)
   While J% < CACOL%
98   DIBAS = DIBAS + 1
     FF = DIBAS: DIA = Val(FECHATEX$(FF)): If DIA < 1 Then GoTo 98
     FF = DIBAS: MES = Val(Mid$(FECHATEX$(FF), 4, 2)): If MES < 1 Then GoTo 98
                                                       If MES > 12 Then GoTo 98
99   DTXX$ = "31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31"
     For MESX% = 1 To MES
       DIAMES% = Val(Mid$(DTXX$, 4 * MESX% - 3, 2))
     Next MESX%
     If DIAMES% = 28 Then
       If (Val(Mid$(FECHATEX$(FF), 7, 2)) Mod 4) = 0 Then
         DIAMES% = 29
       End If
     End If
     If DIA > DIAMES% Then GoTo 98
     '
     FF = DIBAS: SMN% = SEMANU%(FF, A, B): If SMN% = SMNA% Then GoTo 98
     FF = DIBAS: If DIASEM%(FF) <> 1 Then GoTo 98
     '
     SMNA% = SMN%
     J% = J% + 1: INIPER%(J%) = DIBAS
     If J% <= 8 Then
       ENCA$ = ENCA$ + AJUSTD$("S." + TRIM$(Str$(SMNA%)), 8)
       ENCB$ = ENCB$ + "   " + Left$(FECHATEX$(DIBAS), 5)
     End If
     DIBAS = DIBAS + 6
   Wend
   '
End Sub
'
