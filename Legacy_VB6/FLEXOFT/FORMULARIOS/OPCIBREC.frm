VERSION 5.00
Begin VB.Form OPCIBREC 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Materiales - Opciones"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3990
      Picture         =   "OPCIBREC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   2520
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Consultas / Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   210
      TabIndex        =   15
      Top             =   2415
      Width           =   3585
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   23
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   -20
            Picture         =   "OPCIBREC.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Recepciones por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label LABEL5 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepciones por Rango de Fechas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   25
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0C0&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   19
         Top             =   1785
         Width           =   3165
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
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
            Height          =   225
            Left            =   1680
            TabIndex        =   21
            Top             =   300
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Pantalla"
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
            Left            =   315
            TabIndex        =   20
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   16
         Top             =   420
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "OPCIBREC.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Recepciones Pendientes de Aprobación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepciones Pen- dientes de Aprob."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   18
            Top             =   60
            Width           =   1590
         End
      End
   End
   Begin VB.CommandButton Command11 
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
      Height          =   720
      Left            =   3990
      Picture         =   "OPCIBREC.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Informes de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   5
         Top             =   1260
         Width           =   2325
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "OPCIBREC.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Anulación de Informes de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de I.R.M.P.'s"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   7
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   840
         Width           =   2325
         Begin VB.CommandButton Command22 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "OPCIBREC.frx":11A2
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir Informe de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver y Re-Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   10
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture15 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   11
         Top             =   450
         Width           =   2325
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "OPCIBREC.frx":14AC
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Modificar Informes de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificar I.R.M.P."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   13
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3990
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4200
      Width           =   770
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   720
      Left            =   3990
      Picture         =   "OPCIBREC.frx":15F6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3990
      Picture         =   "OPCIBREC.frx":1740
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   1785
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3405
      Picture         =   "OPCIBREC.frx":1A4A
      Top             =   4530
      Width           =   2010
   End
End
Attribute VB_Name = "OPCIBREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LETRANO%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call HELPONLINE("SIMPTEMP")
   Command11.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call OPNOIN("")
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
10 Call SELECHO("IREPRIM")
   NBOLRE& = Val(VALACT1$("IREPRIM"))
   If NBOLRE& > 0 Then
     TTXX$ = VALACT2$("IREPRIM")
     TTYY$ = "Fecha: " + TRIM$(Left$(TTXX$, 8)) + " - Remito: " + TRIM$(Mid$(TTXX$, 11, 16)) + " - Proveedor: " + TRIM$(Mid$(TTXX$, 28, 29))
     If COMALTER%("Acepta Anular IRMP " + TRIM$(Str$(NBOLRE&)) + "\" + TTYY$ + "\\Abandonar\Acepta Anular") = 2 Then
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Comentario - Motivo de Anulacion")
       If VTB% < 0 Then
         Call COMUNI("Anulación Cancelada")
         GoTo 90
       End If
       '
       Call BLOQARCH("IREPRIM")
       Call BLOQARCH("IREPRIPE")
       Call BLOQARCH("MOVISTO")
       Call BLOQARCH("REMAPRI")
       Call BLOQARCH("REPRIPEN")
       Call BLOQARCH("ANUIREC")
       '
       For KKU& = 1 To ULTREG("!OBSERVOF")
         XZ$ = REGLEIDO$("!OBSERVOF", KKU&)
         Call REGAPP("ANUIREC", MKS$(NBOLRE&) + XZ$)
       Next KKU&
       Call CIERRARCH("ANUIREC")
       '
       ICODI$ = ""
       XX$ = FILTERGETRECORD$("IREPRIM", 1, MKS$(NBOLRE&))
       If CVS(Left$(XX$, 4)) = NBOLRE& Then
         XX1$ = Mid$(XX$, 5, 56)
         If Left$(XX1$, 6) <> "(anul)" Then
           Call REPLA(XX$, "(anul) " + XX1$, 5, 56)
           Call FILTERPUTRECORD("IREPRIM", 1, MKS$(NBOLRE&), XX$)
         End If
       End If
       Call CIERRARCH("IREPRIM")
       '
       FIN& = ULTREG("IREPRIPE")
       JJ& = 0
       Screen.MousePointer = 11
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("IREPRIPE", U&)
         If CVS(Left$(XX$, 4)) <> NBOLRE& Then
           JJ& = JJ& + 1
           Call GRAREG("IREPRIPE", XX$, JJ&)
         End If
       Next U&
       Call SETULTREG("IREPRIPE", JJ&)
       Call CIERRARCH("IREPRIPE")
       '
       RFF$ = RECFILT$("REMAPRI", 1, MKS$(NBOLRE&))
       For U& = 1 To Len(RFF$) Step 4
         REBO& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("REMAPRI", REBO&)
         If Asc(Mid$(XX$, 98, 1)) <> 9 Then ' no esta anulado
           Call REPLA(XX$, Chr$(9), 98, 1)
           Call GRAREG("REMAPRI", XX$, REBO&)
           CIXI% = CVI(Mid$(XX$, 83, 2))
           DOPRIBUS$ = "BR." + TRIM$(Str$(NBOLRE&))
           If InStr(EMPREAC$, "SABO") > 0 Then DOPRIBUS$ = TRIM$(Mid$(XX$, 69, 12))
           '
           RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
           For KU& = 1 To Len(RGG$) Step 4
             REMO& = CVS(Mid$(RGG$, KU&, 4))
             MS$ = REGLEIDO$("MOVISTO", REMO&)
             If TRIM$(Mid$(MS$, 23, 10)) = DOPRIBUS$ Then
               TXT1$ = "ANULACION " + TRIM$(Mid$(MS$, 43, 30))
               Call REPLA(MS$, TXT1$, 43, 30)
               TXT1$ = Mid$(MS$, 96, 8) + Mid$(MS$, 88, 8)
               Call REPLA(MS$, TXT1$, 88, 16)
               Call REGAPP("MOVISTO", MS$)
             End If
           Next KU&
           ICODI$ = ICODI$ + MKI$(CIXI%)
         End If
       Next U&
       Call CIERRARCH("MOVISTO")
       Call CIERRARCH("REMAPRI")
       '
       FIN& = ULTREG("REPRIPEN")
       JJ& = 0
       Screen.MousePointer = 11
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("REPRIPEN", U&)
         If CVS(Left$(XX$, 4)) <> NBOLRE& Then
           JJ& = JJ& + 1
           Call GRAREG("REPRIPEN", XX$, JJ&)
         End If
       Next U&
       Call SETULTREG("REPRIPEN", JJ&)
       Call CIERRARCH("REPRIPEN")
       '
       For KU& = 1 To Len(ICODI$) Step 2
         CIXI% = CVI(Mid$(ICODI$, KU&, 2))
         Call VESARRAS(CIXI%)
       Next KU&
       Call CIERRARCH("MOVISTO")
       Call CIERRARCH("*.*")
       '
       Screen.MousePointer = 1
       Call COMUNI("Anulacion Completada")
       GoTo 10
       '
     End If
   End If
   '
90 Call CIERRARCH("*.*")
   '
   'Call BRECANUL
   Command14.Enabled = True
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    Call CONECRUN("FLSETUP/OPIRMP00", "Configuración de Funcionamiento")
    Call CIERRARCH("*.*")
    Call FINAL("")
    Command18.Enabled = True
    '
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If Option1.Value = True Then
        Call FINAL("*COBOLPE")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIBOLPE")
    End If
    Command2.Enabled = True
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    '
    ASELE$ = "IREPRIM"
    '
3   Call SELECHO(ASELE$ + "/Indice General de Informes de Recepción")
    NPXX$ = TRIM$(VALACT1$(ASELE$))
    If NPXX$ <> "" Then
      NBREC& = Val(NPXX$)
      RFF$ = RECFILT$("ANUIREC", 1, MKS$(NBREC&))
      If Len(RFF$) > 0 Then
        JJ& = 0
        For U& = 1 To Len(RFF$) Step 4
          REANU& = CVS(Mid$(RFF$, U&, 4))
          YY$ = Mid$(REGLEIDO$("ANUIREC", REANU&), 5)
          JJ& = JJ& + 1
          Call GRAREG("!OBSERVOF", YY$, JJ&)
        Next U&
        Call SETULTREG("!OBSERVOF", JJ&)
        Call CIERRARCH("!OBSERVOF")
        VTB% = VENTABU%("OBSERVOF/TITUPAN")
      End If
      '
      Call SHOWBREC(NPXX$)
      GoTo 3
    End If
    '
99  Command22.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    If Option1.Value = True Then
       Call MENSERR(24, "Opcion no Disponible por Pantalla.\  \Emita Listado Impreso y Exporte a Archivo de Texto.")
       GoTo 99
    End If
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("REMAPRI", "BOLSELEC")
    J& = 0
    FIN& = ULTREG&("REMAPRI")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("REMAPRI", U&)
       CIIX% = CVI(Mid$(X$, 83, 2))
       FERE% = CVI(Mid$(X$, 81, 2))
       '
       If FERE% >= FF1% Then
         If FERE% <= FF2% Then
           If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
             CODE$ = CODEXT$(CIIX%)
             If CODE$ >= cod1$ Then
               If CODE$ <= cod2$ Then
                 J& = J& + 1
                 Call GRAREG("BOLSELEC", X$, J&)
               End If
             End If
           End If
         End If
       End If
       '
    Next U&
    '
    Call SETULTREG("BOLSELEC", J&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango: " + RANGO$)
    Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
    '
99 Screen.MousePointer = 1
   '
   Command3.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call ACCDIR("REMAPRI")
   Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Sub COCOPEN()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = Asc(Mid$(XXX$, 122, 1))
         If MONEMI% <= 1 Then MONEX$ = "    "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         PREUNID# = CVD(Mid$(XXX$, 79, 8)) * TICAMBIO
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
         Call REPLA(XXY$, MKD$(PREUNID#), 133, 8)
         Call REPLA(XXY$, MKD$(PREUNID# * CAPEN), 141, 8)
         JJ& = JJ& + 1
         Call GRAREG("OPENDETA", XXY$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   Screen.MousePointer = 1
   '
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   OPX% = COMALTER%("Elija su Opción de Control:\\Cumplimiento\Calidad\Cancelar")
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  GoTo 10
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("REMAPRI")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("REMAPRI", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
            CAPEN(K&) = CAPEN(K&) - CAREC
            GoTo 19
          End If
        End If
      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
      Z0$ = REGBLAN$("COCOBREC")
      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
      FENTREPRO = CVI(Mid$(X$, 95, 2))
      '
      If Len(PUBOL$(K&)) < 4 Then
           DIATRA% = 0
           If FENTREPRO < HOII Then
             DIATRA% = DIENTRE%(FENTREPRO, HOII)
           End If
           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
           NCOCO& = NCOCO& + 1
           Call GRAREG("COCOBREC", Z0$, NCOCO&)
         Else
           For II& = 1 To Len(PUBOL$(K&)) Step 4
             Z$ = Z0$
             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
             Y$ = REGLEIDO$("REMAPRI", REBREC&)
             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
             FERECI = CVI(Mid$(Y$, 81, 2))
             DIATRA% = 0
             If FENTREPRO < FERECI Then
               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
             End If
             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
             NCOCO& = NCOCO& + 1
             Call GRAREG("COCOBREC", Z$, NCOCO&)
           Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("REMAPRI")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("REMAPRI", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                CI% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(CI%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                JJ& = JJ& + 1
                Call GRAREG("COCAREC", Z$, JJ&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
   LETRANO% = Val(CONTROL$("BOLRECEP", "LETRANO"))
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub BRECANUL()
    '
    If COMALTER%("Realmente Desea\ANULAR Informes de Recepción ?\\No, Abandonar\Si, Anular I/R's") <> 2 Then
      Exit Sub
    End If
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "Proceso de Anulación de Informes de Recepción"
10  TITXX$ = "Informes de Recepción para Anulación"
    YY$ = REGSEL$("IREPRIM/" + TITXX$)
    If Len(YY$) <= 4 Then GoTo 999
    '
    NPX& = CVS(Left$(YY$, 4))
    If NPX& < 1 Then GoTo 999
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("REMAPRI")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("REMAPRI", J&)
      If CVS(Left$(XX$, 4)) = NPX& Then
        XX$ = String$(256, 0)
        Call GRAREG("REMAPRI", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("REMAPRI")
    '
    FIN& = ULTREG&("REPRIPEN")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("REPRIPEN", J&)
      If CVS(Left$(XX$, 4)) = NPX& Then
        XX$ = String$(256, 0)
        Call GRAREG("REPRIPEN", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("REPRIPEN")
    '
    FIN& = ULTREG&("IREPRIM")
    JJ& = 0
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("IREPRIM", J&)
      If CVS(Left$(XX$, 4)) <> NPX& Then
        JJ& = JJ& + 1
        Call GRAREG("IREPRIM", XX$, JJ&)
      End If
    Next J&
    Call SETULTREG("IREPRIM", JJ&)
    Call CIERRARCH("IREPRIM")
    '
    FIN& = ULTREG&("IREPRIPE")
    JJ& = 0
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("IREPRIPE", J&)
      If CVS(Left$(XX$, 4)) <> NPX& Then
        JJ& = JJ& + 1
        Call GRAREG("IREPRIPE", XX$, JJ&)
      End If
    Next J&
    Call SETULTREG("IREPRIPE", JJ&)
    Call CIERRARCH("IREPRIPE")
    '
    FIN& = ULTREG&("OBSREPRI")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OBSREPRI", J&)
      If CVS(Left$(XX$, 4)) = NPX& Then
        XX$ = String$(256, 0)
        Call GRAREG("OBSREPRI", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("OBSREPRI")
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FRESHECHO("IREPRIM")
    Screen.MousePointer = 1
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
End Sub

Sub SHOWBREC(NPXX$)
    '
    TIDOCUM$ = "Informe de Recepción"
    NUDOCU$ = TRIM$(NPXX$)
    'While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    'NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        XX1$ = REPLACAR$(XX$, " ", "")
        If InStr(XX1$, TIDOCUM$ + NUDOCU$) > 0 Or InStr(XX1$, "IR-" + NUDOCU$ + "-") > 0 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
End Sub
