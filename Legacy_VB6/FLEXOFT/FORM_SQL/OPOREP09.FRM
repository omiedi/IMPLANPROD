VERSION 5.00
Begin VB.Form OPOREP09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Reparación"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   4740
      Left            =   4335
      ScaleHeight     =   4680
      ScaleWidth      =   1155
      TabIndex        =   17
      Top             =   -45
      Width           =   1215
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   15
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   20
         Top             =   3120
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   105
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
         Height          =   660
         Left            =   105
         Picture         =   "OPOREP09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   150
         Width           =   700
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remito"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         Picture         =   "OPOREP09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Conexión con Remito"
         Top             =   1755
         Width           =   700
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -285
         Picture         =   "OPOREP09.frx":0454
         Top             =   3465
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver y Re-Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   10
      Top             =   2625
      Width           =   3900
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1425
         TabIndex        =   13
         Top             =   420
         Width           =   1485
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
            Left            =   0
            Picture         =   "OPOREP09.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Visualizar Presupuestos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Presu- puestos"
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
            TabIndex        =   14
            Top             =   105
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2100
         ScaleHeight     =   585
         ScaleWidth      =   1425
         TabIndex        =   11
         Top             =   420
         Width           =   1485
         Begin VB.CommandButton Command4 
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
            Picture         =   "OPOREP09.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Visualizar Ordenes de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Orden Repar."
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
            TabIndex        =   12
            Top             =   105
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Opciones de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3900
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   1
         Top             =   1470
         Width           =   2640
         Begin VB.CommandButton Command5 
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
            Picture         =   "OPOREP09.frx":298A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación de Orden en Curso en Curso"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Orden de Reparación"
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
            TabIndex        =   3
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   4
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command15 
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
            Picture         =   "OPOREP09.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cerrar Orden en Curso"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Cerrar Orden de Reparación en Curso"
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
            Top             =   50
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   7
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command9 
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
            Picture         =   "OPOREP09.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Generación a Partir de un Presupuesto Previo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Reparación"
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
            Left            =   945
            TabIndex        =   9
            Top             =   45
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "OPOREP09"
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

Private Sub Command15_Click()

    Call MENURE09.Command15_Click
    '
'   Command15.Enabled = False
'3  Screen.MousePointer = 11
'   Call COPYSTRU("ORDEREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     XX$ = REGLEIDO$("ORDEREPA", U&)
'     If CVI(Mid$(XX$, 65, 2)) < 3 Then
'       JJ& = JJ& + 1
'       Call GRAREG("PREPAEMI", XX$, JJ&)
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
'   Screen.MousePointer = 1
'   '
'   If JJ& < 1 Then
'     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Cerrarse.")
'     GoTo 99
'   End If
'   '
'   Call SELECHO("PREPAEMI")
'   NOREP& = Val(VALACT1$("PREPAEMI"))
'   If NOREP& < 1 Then GoTo 99
'   '
'   NOREX$ = TRIM$(Str$(NOREP&))
'   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
'   NOREX$ = "OR-" + NOREX$
'   '
'   VACON% = 1
'   Call CARCOMPO("", NOREX$, 1, VACON%, 2)
'   If VACON% < 0 Then GoTo 99
'   '
'   Screen.MousePointer = 11
'   NUPRESU& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     X$ = REGLEIDO$("ORDEREPA", U&)
'     If CVS(Left$(X$, 4)) = NOREP& Then
'       NUPRESU& = CVS(Mid$(X$, 125, 4))
'       REFERE$ = Mid$(X$, 86, 9)
'       REFOP$ = "(cerr) " + Mid$(X$, 5, 60)
'       CODREPA$ = TRIM$(Mid$(X$, 497, 16))
'       Call REPLA(X$, REFOP$, 5, 60)
'       Call REPLA(X$, MKI$(3), 65, 2) ' MARCA CERRADA
'       Call GRAREG("ORDEREPA", X$, U&)
'       GoTo 20
'     End If
'   Next U&
'   '
'20 NOREX$ = TRIM$(Str$(NOREP&))
'   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
'   NOREX$ = "OR-" + NOREX$
'   '
'   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
'   RCC$ = ""
'   URE& = ULTREG&("RESERVA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("RESERVA", RGG&)
'         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
'           Call GRAREG("RESERVA", String$(128, 0), RGG&)
'         End If
'       End If
'     End If
'   Next KK&
'   '
'   FECIE% = HOY(HOS$)
'   For U& = 1 To ULTREG&("!MATEROF")
'     ZZ$ = REGLEIDO$("!MATEROF", U&)
'     '
'     CI2% = CVI(Mid$(ZZ$, 1, 2))
'     CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
'     '
'     USOUNI# = CAN
'     If CI2% > 0 Then
'       If CI2% <= NUMAS% Then
'         If Abs(CAN) >= 0.05 Then
'           LOTE$ = ""
'           CMOV$ = "RP"
'           '
'           DOPRI$ = NOREX$
'           DOSEC$ = DOPRI$
'           DEPX% = 1
'           If TRIM$(REFERE$) <> "" Then DOSEC$ = REFERE$
'           REFE$ = "Reparación " + CODREPA$
'           If Abs(CAN) >= 0.05 Then
'              DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
''             Call Movisto(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
'              Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DepOrig%, CANTI)
'           End If
'         End If
'       End If
'     End If
'   Next U&
'   '
'   Call CIERRARCH("MOVISTO")
'   Call CIERRARCH("PRESREPA")
'   Call CIERRARCH("ORDEREPA")
'   Call CIERRARCH("RESERVA")
'   Screen.MousePointer = 1
'   '
'   Call COMUNI("Cierre " + NOREX$ + " Completo")
'   '
'99 Command15.Enabled = True

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVI(Mid$(XX$, 123, 2)) < 1 Then
     '  If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
         JJ& = JJ& + 1
         Call GRAREG("PREPAEMI", XX$, JJ&)
     '  End If
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
     GoTo 99
   End If
   '
   Call FRESHECHO("PREPAEMI")
   Call SELECHO("PREPAEMI/Indice de Reparaciones Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then GoTo 99
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   GoTo 99
   '
10 FOREMI09.REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", UI&)
   SEREF$ = "S/No." + TRIM$(Mid$(XX$, 71, 15)) + " - Ref: " + TRIM$(Mid$(XX$, 86, 9))
   'If CVI(Mid$(XX$, 121, 2)) < 1 Then
   '   Call COMUNI("Falta Generar y Emitir Presupuesto.")
   '   GoTo 99
   'End If
   '
   'If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
   '   Call COMUNI("Orden de Reparación ya Emitida.")
   '   GoTo 3
   'End If
   '
   'ORDREBOM.Label18 = TRIM$(Str$(NPX&))
   'ORDREBOM.Label2 = TRIM$(Str$(NUORDREP&))
   '
   NCLIE% = CVI(Mid$(XX$, 67, 2))
   FOREMI09.Text1 = TRIM$(Str$(NCLIE%))
   FOREMI09.Text1.TEXT = TRIM$(Str$(NCLIE%))
   FOREMI09.Text2.TEXT = RASOCLI$(NCLIE%)
   FOREMI09.Text3.TEXT = DOMICLI$(NCLIE%)
   FOREMI09.Text4.TEXT = CPOSCLI$(NCLIE%)
   FOREMI09.Text5.TEXT = LOCACLI$(NCLIE%)
   Call CARTRANSPOR
   FOREMI09.Text22 = TRIM$(Mid$(XX$, 129, 368)) + Chr$(13) + Chr$(10) + SEREF$
   FOREMI09.Show 1
   '
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Call PRESRE09.Command15_Click
'    Command3.Enabled = False
'    '
'3   Call SELECHO("PRESREPA/Indice Presupuestos de Reparación")
'    NPX& = Val(VALACT1$("PRESREPA"))
'    If NPX& < 1 Then GoTo 99
'    '
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    For UI& = ULTREG&("PRESREPA") To 1 Step -1
'      XX$ = REGLEIDO$("PRESREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    GoTo 99
'    '
'4   NUDOCU$ = TRIM$(Str$(NPX&))
'    Screen.MousePointer = 11
'    TIDOCUM$ = "Presupuesto de Reparación " + NUDOCU$
'    TIDOCU2$ = "Operaciones de Reparación PR-" + NUDOCU$
'    Screen.MousePointer = 11
'    '
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 5
'      End If
'    Next UI&
'    '
'    Screen.MousePointer = 1
'    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'    On Error Resume Next
'    GoTo 99
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 15
'      End If
'    Next UI&
'    '
'15  REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
'    GoTo 18
'    '
'17  Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'18  While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 19
'      L& = L& - 1
'    Wend
'    '
'19  For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'       TPSP$ = Mid$(TPSP$, PPXX% + 1)
'       '
'       If Len(COPRIAT$) > 0 Then
'         COPRI% = Asc(Left$(COPRIAT$, 1))
'         ATRIP$ = Mid$(COPRIAT$, 2)
'         Call SHOWCOMMAND(COPRI%, ATRIP$)
'       End If
'       GoTo 50
'       '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Visible = False
'      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 3
'    '
'99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    '
    Call ORREPA09.Command15_Click
    '
'    Command4.Enabled = False
'    '
'3   Call SELECHO("ORDEREPA/Indice General Ordenes de Reparación")
'    NPX& = Val(VALACT1$("ORDEREPA"))
'    If NPX& < 1 Then GoTo 99
'    '
'    For UI& = ULTREG&("ORDEREPA") To 1 Step -1
'      XX$ = REGLEIDO$("ORDEREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    Exit Sub
'    '
'4   NUDOCU$ = TRIM$(Str$(NPX&))
'    TIDOCUM$ = "Orden de Reparación " + NUDOCU$
'    TIDOCU2$ = "Procedimiento de Reparación OR-" + NUDOCU$
'    Screen.MousePointer = 11
'    '
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 5
'      End If
'    Next UI&
'    '
'    Screen.MousePointer = 1
'    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'    On Error Resume Next
'    GoTo 99
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCU2$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 15
'      End If
'    Next UI&
'    '
'15  REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
'    GoTo 18
'    '
'17  Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'18  While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 19
'      L& = L& - 1
'    Wend
'    '
'19  For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'    FIN& = Len(TPSP$)
'50  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 50
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
''    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 3
'    '
'99  Command4.Enabled = True
    '
End Sub

Private Sub Command5_Click()
    '
    Call MENURE09.Command5_Click
    '
'   Command5.Enabled = False
'3  Screen.MousePointer = 11
'   Call COPYSTRU("ORDEREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     XX$ = REGLEIDO$("ORDEREPA", U&)
'     If CVI(Mid$(XX$, 65, 2)) < 3 Then
'       JJ& = JJ& + 1
'       Call GRAREG("PREPAEMI", XX$, JJ&)
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
'   Screen.MousePointer = 1
'   '
'   If JJ& < 1 Then
'     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Anularse.")
'     GoTo 99
'   End If
'   '
'   Call SELECHO("PREPAEMI")
'   NOREP& = Val(VALACT1$("PREPAEMI"))
'   If NOREP& < 1 Then GoTo 99
'   '
'   If COMALTER%("Confirma Anulación de O/R " + TRIM$(Str$(NOREP&)) + "\" + TRIM$(VALACT2$("PREPAEMI")) + "\\Cancelar\Anular O/R") <> 2 Then
'      GoTo 99
'   End If
'   '
'   Screen.MousePointer = 11
'   NUPRESU& = 0
'   For U& = 1 To ULTREG&("ORDEREPA")
'     X$ = REGLEIDO$("ORDEREPA", U&)
'     If CVS(Left$(X$, 4)) = NOREP& Then
'       REFOP$ = "(anul) " + Mid$(X$, 5, 60)
'       NUPRESU& = CVS(Mid$(X$, 125, 4))
'       Call REPLA(X$, REFOP$, 5, 60)
'       Call REPLA(X$, MKI$(9), 65, 2) ' MARCA ANULADA
'       Call GRAREG("ORDEREPA", X$, U&)
'       GoTo 20
'     End If
'   Next U&
'   '
'20 If NUPRESU& > 0 Then
'     For U& = 1 To ULTREG&("PRESREPA")
'       X$ = REGLEIDO$("PRESREPA", U&)
'       If CVS(Left$(X$, 4)) = NUPRESU& Then
'         Call REPLA(X$, MKS$(0), 125, 4) ' QUITA NUMERO DE O/R
'         Call GRAREG("PRESREPA", X$, U&)
'         GoTo 30
'       End If
'     Next U&
'   End If
'   '
'30 NOREX$ = TRIM$(Str$(NOREP&))
'   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
'   NOREX$ = "OR-" + NOREX$
'   '
'   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
'   RCC$ = ""
'   URE& = ULTREG&("RESERVA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("RESERVA", RGG&)
'         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
'           Call GRAREG("RESERVA", String$(128, 0), RGG&)
'         End If
'       End If
'     End If
'   Next KK&
'   '
'   Call CIERRARCH("PRESREPA")
'   Call CIERRARCH("ORDEREPA")
'   Call CIERRARCH("RESERVA")
'   Screen.MousePointer = 1
'   '
'   Call COMUNI("Anulación Completa")
'   '
'99 Command5.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Call MENURE09.Command9_Click
    '
'   Command9.Enabled = False
'   '
'3  Screen.MousePointer = 11
'   Call COPYSTRU("PRESREPA", "PREPAEMI")
'   JJ& = 0
'   For U& = 1 To ULTREG&("PRESREPA")
'     XX$ = REGLEIDO$("PRESREPA", U&)
'     If CVI(Mid$(XX$, 121, 2)) > 0 Then
'       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
'         JJ& = JJ& + 1
'         Call GRAREG("PREPAEMI", XX$, JJ&)
'       End If
'     End If
'   Next U&
'   Call SETULTREG("PREPAEMI", JJ&)
'   Call CIERRARCH("PREPAEMI")
'   Screen.MousePointer = 1
'   '
'   If JJ& < 1 Then
'     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
'     GoTo 99
'   End If
'   '
'   Call FRESHECHO("PREPAEMI")
'   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
'   NPX& = Val(VALACT1$("PREPAEMI"))
'   If NPX& < 1 Then GoTo 99
'   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
'   Next UI&
'   Call COMUNI("Error de Indices - Consulte.")
'   Exit Sub
'   '
'10 REPRESUP = TRIM$(Str$(UI&))
'   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
'   If CVI(Mid$(XX$, 121, 2)) < 1 Then
'      Call COMUNI("Falta Generar y Emitir Presupuesto.")
'      GoTo 99
'   End If
'   '
'   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
'      Call COMUNI("Orden de Reparación ya Emitida.")
'      GoTo 3
'   End If
'   '
'   ORDREBOM.Label18 = TRIM$(Str$(NPX&))
'   ORDREBOM.Label2 = TRIM$(Str$(NUORDREP&))
'   '
'   NPP& = 1
'   FIN& = ULTREG&("PEDENCA")
'   If FIN& > 0 Then
'      For U& = FIN& To FIN& - 16 Step -1
'         If U& < 1 Then GoTo 99
'         XX$ = REGLEIDO$("PEDENCA", U&)
'         NPPX& = CVS(Mid$(XX$, 1, 4))
'         If NPPX& >= NPP& Then
'            NPP& = 1 + NPPX&
'         End If
'      Next U&
'   End If
'   ORDREBOM.Label25 = TRIM$(Str$(NPP&))
'   '
'   ORDREBOM.Show 1
'   GoTo 3
'   '
'99 Command9.Enabled = True
   '
End Sub

Function NUORDREP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("ORDEREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("ORDEREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUORDREP& = NPP&
   '
End Function
'

Sub CARTRANSPOR()
    '
    NCLIE% = Val(FOREMI09.Text1.TEXT)
    LENTREG$ = Space$(96)
    RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
    If Len(RFEX$) >= 4 Then
       RELENT& = CVS(Left$(RFEX$, 4))
       If RELENT& > 0 Then
         If RELENT& <= ULTREG&("LENTRECL") Then
           LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
         End If
       End If
    End If
    FOREMI09.Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

