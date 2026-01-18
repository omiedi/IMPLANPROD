VERSION 5.00
Begin VB.Form FORCLATEC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificacion por Atributos"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6075
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5020
      Picture         =   "FORCLATEC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   300
      Width           =   855
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
      Height          =   855
      Left            =   5020
      Picture         =   "FORCLATEC.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Clasificación del Maestro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0FFD0&
         Caption         =   "Criterios y Tablas"
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
         Left            =   210
         TabIndex        =   7
         Top             =   1365
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Tablas de Clasificación"
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
            Index           =   6
            Left            =   420
            TabIndex        =   10
            Top             =   735
            Value           =   -1  'True
            Width           =   2745
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Tabla de Criterios Técnicos"
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
            Index           =   7
            Left            =   420
            TabIndex        =   9
            Top             =   420
            Width           =   2850
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
            Height          =   705
            Left            =   3255
            Picture         =   "FORCLATEC.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   4
         Top             =   525
         Width           =   2010
         Begin VB.CommandButton Command7 
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
            Picture         =   "FORCLATEC.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Actualización Características Técnicas de Item de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignación de Atributos"
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
            TabIndex        =   6
            Top             =   105
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   1
         Top             =   525
         Width           =   2010
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
            Picture         =   "FORCLATEC.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Búsqueda de Artículo por Características Técnicas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Búsqueda por Características"
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
            TabIndex        =   3
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4450
      Picture         =   "FORCLATEC.frx":1DCE
      Top             =   2460
      Width           =   2010
   End
End
Attribute VB_Name = "FORCLATEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CIXI% = COBUSCATRI%
   If CIXI% > 0 Then
     Clipboard.SetText (TRIM$(CODEXT$(CIXI%)))
     Call FINAL("")
   End If
   Command14.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
10 Call SELECHO("MASTER")
   CODD1$ = TRIM$(VALACT1$("MASTER"))
   If CODD1$ <> "" Then
     CIXI% = CODINT%(CODD1$)
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         Call CATRIBUS(CIXI%)
         GoTo 10
       End If
     End If
   End If
   Command7.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If Option1(6).Value = True Then
        Call CARTABUS
      ElseIf Option1(7).Value = True Then
        Call CARCRIBUS
      Else
        Call OPNOIN("")
    End If
    Command6.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     '
10   Caption = Caption + "  -  " + EPAC$
     APPX$ = App.EXEName
     If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     FMDX = FECHMODU%
     BB1 = HOY(HOS$)
     If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
       Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
     End If
     '
   End If
   '
   CODD1$ = TRIM$(APLINVO$)
   If CODD1$ <> "" Then
     CIXI% = CODINT%(CODD1$)
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         Call CATRIBUS(CIXI%)
         Call FINAL("")
         Exit Sub
       End If
     End If
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FINAL("")
End Sub

Private Sub Option1_DblClick(Index As Integer)
    '
    If Index = 6 Then
        Call CARTABUS
     ElseIf Index = 7 Then
        Call CARCRIBUS
    End If
    '
End Sub

Sub CARTABUS()
   '
10 X$ = REGSEL$("TACRIBUS")
   If X$ = "" Then Exit Sub
   '
   CLASI% = Asc(Left$(X$, 1))
   CLASX$ = TRIM$(Mid$(X$, 2, 31))
   If CLASI% < 1 Or CLASI% > 255 Or CLASX$ = "" Then Exit Sub
   '
   CLASX$ = REPLACAR$(CLASX$, "/", "-")
   ALISX$ = TRIM$(Str$(CLASI%))
   While Len(ALISX$) < 3: ALISX$ = "0" + ALISX$: Wend
   ACLASI$ = "CLATE" + ALISX$
   Call COPYSTRU("CLASISTO", ACLASI$)
   '
   Call TRALOCAL(ACLASI$, "CLASISTO")
   VTB% = VENTABU%("CLASIBUS/NC/TITUPAN=Clasificación Técnica: " + CLASX$)
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!CLASISTO")
         X$ = REGLEIDO$("!CLASISTO", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG(ACLASI$, X$, U&)
      Next U&
      Call SETULTREG(ACLASI$, U&)
      Call CIERRARCH(ACLASI$)
   End If
   Call CIERRARCH("*.*")
   GoTo 10
   '
End Sub

Sub CARCRIBUS()
   '
   Call TRALOCAL("TACRIBUS", "")
   VTB% = VENTABU%("TACRIBUS/NC/TITUPAN=Criterios Clasificación Técnica")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACRIBUS")
        If U& < 256 Then
          X$ = REGLEIDO$("!TACRIBUS", U&)
          Call REPLA(X$, Chr$(U&), 1, 1)
          Call GRAREG("TACRIBUS", X$, U&)
        End If
      Next U&
   End If
   If ULTREG&("TACRIBUS") > 255 Then
      Call SETULTREG("TACRIBUS", 255)
   End If
   Call CIERRARCH("TACRIBUS")
   '
End Sub

Sub CATRIBUS(CIXI%)
   '
   If CIXI% < 1 Or CIXI% > NUMAS% Then Exit Sub
   '
   CODESCRI$ = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
   M$ = FILTERGETRECORD$("TATRIBUS", 1, MKI$(CIXI%))
   CLASI$ = Left$(Mid$(M$, 3, 254) + String$(254, 0), 254)
   Call TRALOCAL("TACRIBUS", "")
   FINX& = ULTREG&("TACRIBUS")
   If FINX& > 254 Then FINX& = 254
   '
   For U& = 1 To FINX&
     VALOX% = Asc(Mid$(CLASI$, U&, 1))
     XX$ = REGLEIDO$("!TACRIBUS", U&)
     Call REPLA(XX$, Chr$(VALOX%), 49, 1)
     Call GRAREG("!TACRIBUS", XX$, U&)
   Next U&
   '
   CLASIA$ = CLASI$
   VTB% = VENTABU%("CLABUTEC/TITUPAN=" + CODESCRI$)
   If VTB% >= 0 Then
     CLASI$ = String$(254, 0)
     For U& = 1 To FINX&
       XX$ = REGLEIDO$("!TACRIBUS", U&)
       VALOX% = Asc(Mid$(XX, 49, 1))
       Mid$(CLASI$, U&, 1) = Chr$(VALOX%)
     Next U&
   End If
   If CLASI$ = CLASIA$ Then GoTo 90
   '
   M$ = MKI$(CIXI%) + CLASI$
   Call FILTERPUTRECORD("TATRIBUS", 1, MKI$(CIXI%), M$)
   Call CIERRARCH("TATRIBUS")
   '
90 End Sub
'
Function COBUSCATRI%()
   '
   Dim CODISE%(2048)
   '
   CBA% = 0
   CLASI$ = String$(254, 0)
   Call TRALOCAL("TACRIBUS", "")
   FINX& = ULTREG&("TACRIBUS")
   If FINX& > 254 Then FINX& = 254
   '
10 For U& = 1 To FINX&
     XX$ = REGLEIDO$("!TACRIBUS", U&)
     Call REPLA(XX$, Chr$(0), 49, 1)
     Call GRAREG("!TACRIBUS", XX$, U&)
   Next U&
   '
12 VTB% = VENTABU%("CLABUTEC/TITUPAN=Seleccione Argumentos de Búsqueda")
   If VTB% >= 0 Then
     CLASI$ = String$(254, 0)
     For U& = 1 To FINX&
       XX$ = REGLEIDO$("!TACRIBUS", U&)
       VALOX% = Asc(Mid$(XX, 49, 1))
       Mid$(CLASI$, U&, 1) = Chr$(VALOX%)
     Next U&
     '
     Screen.MousePointer = 11
     JJ& = 0
     FIN& = ULTREG&("TATRIBUS")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("TATRIBUS", U&)
       XY$ = Mid$(XX$, 3, FINX&)
       For JI% = 1 To Len(XY$)
         VALOX% = Asc(Mid$(CLASI$, JI%, 1))
         If VALOX% > 0 Then
           If Asc(Mid$(XY$, JI%, 1)) <> VALOX% Then
             GoTo 19
           End If
         End If
       Next JI%
       '
       If JJ& < 2048 Then
         JJ& = JJ& + 1
         CODISE%(JJ&) = CVI(Left$(XX$, 2))
       End If
19   Next U&
     Screen.MousePointer = 1
     '
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 12
       ElseIf JJ& = 1 Then
         CBA% = CODISE%(1)
25       CICDX% = CBA%
         REMASX& = CBA% + 1
         RMCDX$ = REGLEIDO$("MASTER", REMASX&) + String$(52, 0)
         Call REPLA(RMCDX$, MKS$(STOCKACT(CICDX%)), 129, 4)
         Call REPLA(RMCDX$, MKS$(STOMIN(CICDX%)), 133, 4)
         Call REPLA(RMCDX$, MKS$(LOTESTAN(CICDX%)), 137, 4)
         Call REPLA(RMCDX$, MKS$(LOREPOMA(CICDX%)), 141, 4)
         Call REPLA(RMCDX$, TALLE$(CICDX%), 145, 16)
         Call REPLA(RMCDX$, COLOR$(CICDX%), 161, 16)
         Call REPLA(RMCDX$, MKS$(PESUNI(CICDX%)), 177, 4)
         If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
            Call REPLA(RMCDX$, MKS$(99999999999#), 111, 4)
         End If
         OBXXX$ = OBJPLA$("MUESTRADATCOD", RMCDX$)
         If OBXXX$ = "" Then GoTo 12 ' pulso cancelar
         '
         COBUSCATRI% = CBA%
         Exit Function
         '
       ElseIf JJ& > 1 Then
         For U1& = 1 To JJ&
           TTXX$ = AJUSTI$(CODEXT$(CODISE%(U1&)), 16)
           TTXX$ = TTXX$ + DESCRIT$(CODISE%(U1&))
           Call GRAREG("!ECOMAST", TTXX$, U1&)
         Next U1&
         Call SETULTREG("!ECOMAST", JJ&)
         Call BAJALDISCO
         Call FRESHECHO("!ECOMAST")
         XX$ = REGSEL$("!ECOMAST")
         If Len(XX$) > 4 Then
           CBA% = CODINT%(Left$(XX$, 16))
           GoTo 25
         End If
         GoTo 12
         '
     End If
     '
   End If
   '
End Function
'
Sub CLASIRANGO()
   '
   CLR$ = OBJPLA$("CLASIRANGO", CLR$)
   Call CIERRARCH("*.*")
   If CLR$ = "" Then Exit Sub
   CLASI% = Asc(Left$(CLR$, 1))
   If CLASI% < 1 Or CLASI% > ULTREG&("TACRICLA") Then Exit Sub
   '
   ALIX$ = TRIM$(Str$(CLASI%))
   While Len(ALIX$) < 3: ALIX$ = "0" + ALIX$: Wend
   ACLASI$ = "CLASI" + ALIX$
   LU$ = LUDAT$
   Call CIERRARCH(ACLASI$)
   Call CIERRARCH("CLASISTO")
   If COPYFILE%(LU$ + ACLASI$ + ".DAT", LU$ + "CLASISTO.DAT") <> 1 Then
      Exit Sub
   End If
   '
   CLR$ = OBJPLA$("CLASIRANGO-2", CLR$)
   If CLR$ = "" Then Exit Sub
   VACLA% = Asc(Mid$(CLR$, 6, 1))
   If VACLA% < 1 Or VACLA% > ULTREG&(ACLASI$) Then Exit Sub
   '
   Screen.MousePointer = 11
   cod1$ = CODEXT$(CVI(Mid$(CLR$, 2, 2))): If TRIM$(cod1$) = "" Then Exit Sub
   cod2$ = CODEXT$(CVI(Mid$(CLR$, 4, 2))): If TRIM$(cod2$) = "" Then Exit Sub
   '
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               If REGMA& > 1 Then
                  MX$ = REGLEIDO$("MASTER", REGMA&)
                  If Mid$(MX$, 77, 16) >= cod1$ Then
                    If Mid$(MX$, 77, 16) <= cod2$ Then
                      Call REPLA(MX$, Chr$(VACLA%), 68 + CLASI%, 1)
                      Call GRAREG("MASTER", MX$, REGMA&)
                    End If
                  End If
               End If
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub

Sub CLASITAB()
   '
   Screen.MousePointer = 11
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               M$ = REGLEIDO$("MASTER", REGMA&)
               CLASI$ = Mid$(M$, 69, 8)
               J& = J& + 1
               Y$ = MKI$(CI%) + CLASI$
               Call GRAREG("!CLASIMAS", Y$, J&)
            End If
         End If
      End If
   Next I&
   '
   Call SETULTREG("!CLASIMAS", J&)
   Call CIERRARCH("!CLASIMAS")
   '
   ARX$ = ""
   For U% = 1 To ULTREG&("TACRICLA")
      U1$ = TRIM$(Str$(U% + 2))
      U2$ = TRIM$(Str$(U%)): While Len(U2$) < 3: U2$ = "0" + U2$: Wend
      ARX$ = ARX$ + "/ARCHECO(" + U1$ + ")=CLASI" + U2$
   Next U%
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("CLASIMAS/NC" + ARX$)
   If VTB% < 0 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("!CLASIMAS")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      Y$ = REGLEIDO$("!CLASIMAS", I&)
      CI% = CVI(Left$(Y$, 2))
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            CLASI$ = Mid$(Y$, 3, 8)
            REGMA& = CI% + 1
            If REGMA& > 1 Then
               MX$ = REGLEIDO$("MASTER", REGMA&)
               Call REPLA(MX$, CLASI$, 69, 8)
               Call GRAREG("MASTER", MX$, REGMA&)
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("!CLASIMAS")
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub
'


