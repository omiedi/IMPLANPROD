VERSION 5.00
Begin VB.Form MENUOREP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Reparación"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5265
   Icon            =   "MENUOREP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
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
      Height          =   650
      Left            =   4250
      Picture         =   "MENUOREP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   1630
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3400
      Left            =   4240
      TabIndex        =   24
      Top             =   2400
      Width           =   870
      Begin VB.CommandButton Command12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   120
         Picture         =   "MENUOREP.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Materiales Suministrados por el Cliente"
         Top             =   2700
         Width           =   600
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   105
         Picture         =   "MENUOREP.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Visualizar Comprobante de Recepción"
         Top             =   900
         Width           =   600
      End
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
         Height          =   560
         Left            =   105
         Picture         =   "MENUOREP.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Visualizar Ordenes de Reparación"
         Top             =   2100
         Width           =   600
      End
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
         Height          =   560
         Left            =   105
         Picture         =   "MENUOREP.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Visualizar Presupuestos"
         Top             =   1500
         Width           =   600
      End
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
         Height          =   560
         Left            =   105
         Picture         =   "MENUOREP.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Abrir Anotador de Seguimiento de Reparaciones"
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cotización"
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
      TabIndex        =   20
      Top             =   1575
      Width           =   3900
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   21
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command6 
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
            Picture         =   "MENUOREP.frx":17B6
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Generar Presupuesto de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Generación de Presu- puesto de Reparación"
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
            TabIndex        =   22
            Top             =   105
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pedido de Cotización"
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
      TabIndex        =   16
      Top             =   210
      Width           =   3900
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   17
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command8 
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
            Picture         =   "MENUOREP.frx":1AC0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Recepcion Equipos para Cotizar Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepción de Equipos a Reparar"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Expedición / Despacho"
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
      TabIndex        =   13
      Top             =   5355
      Width           =   3900
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   14
         Top             =   420
         Width           =   2640
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
            Picture         =   "MENUOREP.frx":1DCA
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Conexión con Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Remito de Expedición"
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
            TabIndex        =   15
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reparaciones"
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
      TabIndex        =   3
      Top             =   2940
      Width           =   3900
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   4
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
            Picture         =   "MENUOREP.frx":20D4
            Style           =   1  'Graphical
            TabIndex        =   5
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
            TabIndex        =   6
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
         TabIndex        =   7
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
            Picture         =   "MENUOREP.frx":23DE
            Style           =   1  'Graphical
            TabIndex        =   8
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
            TabIndex        =   9
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
         TabIndex        =   10
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
            Picture         =   "MENUOREP.frx":26E8
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   45
            Width           =   1485
         End
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
      Height          =   650
      Left            =   4250
      Picture         =   "MENUOREP.frx":29F2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4250
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   5880
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   2
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command11 
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
      Left            =   4250
      Picture         =   "MENUOREP.frx":2B3C
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Ayuda FLEXOFT en línea"
      Top             =   970
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3705
      Picture         =   "MENUOREP.frx":2E46
      Top             =   6090
      Width           =   2010
   End
End
Attribute VB_Name = "MENUOREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    '
3   Call SELECHO("PRESREPA/Indice Recepciones de Reparación")
    NPX& = Val(VALACT1$("PRESREPA"))
    If NPX& < 1 Then GoTo 99
    '
    NUDOCU$ = TRIM$(Str$(NPX&))
    For UI& = ULTREG&("PRESREPA") To 1 Step -1
      XX$ = REGLEIDO$("PRESREPA", UI&)
      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
    Next UI&
    Call COMUNI("Error de Indices - Consulte.")
    GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    TIDOCUM$ = "Recepción de Reparación " + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
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
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
50  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
       TPSP$ = Mid$(TPSP$, PPXX% + 1)
       '
       If Len(COPRIAT$) > 0 Then
         COPRI% = Asc(Left$(COPRIAT$, 1))
         ATRIP$ = Mid$(COPRIAT$, 2)
         Call SHOWCOMMAND(COPRI%, ATRIP$)
       End If
       GoTo 50
       '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Visible = False
      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command10.Enabled = True
    '
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call CONECRUN("MASUC09", "Material Suministrado por el Cliente")
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   CONFREPA.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command15_Click()
   '
   Command15.Enabled = False
3  Screen.MousePointer = 11
   Call COPYSTRU("ORDEREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("ORDEREPA")
     XX$ = REGLEIDO$("ORDEREPA", U&)
     If CVI(Mid$(XX$, 65, 2)) < 3 Then
       JJ& = JJ& + 1
       Call GRAREG("PREPAEMI", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Cerrarse.")
     GoTo 99
   End If
   '
   Call SELECHO("PREPAEMI")
   NOREP& = Val(VALACT1$("PREPAEMI"))
   If NOREP& < 1 Then GoTo 99
   '
   NOREX$ = TRIM$(Str$(NOREP&))
   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = "OR-" + NOREX$
   '
   Call SELECHO("TADEPOSI")
   DP2% = XVALO(VALACT1$("TADEPOSI"))
   If DP2% < 1 Then GoTo 99
   DEPOCON$ = VALACT2$("TADEPOSI")
   '
   VACON% = 1
   Call CARCOMPO("", NOREX$, 1, VACON%, 2)
   If VACON% < 0 Then GoTo 99
   '
   If COMALTER%("Confirma Cierre de " + NOREX$ + "\Consumiendo Componentes de Depósito '" + DEPOCON$ + "'\\Si, Cerrar\No, Cancelar") <> 1 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   NUPRESU& = 0
   For U& = 1 To ULTREG&("ORDEREPA")
     X$ = REGLEIDO$("ORDEREPA", U&)
     If CVS(Left$(X$, 4)) = NOREP& Then
       NUPRESU& = CVS(Mid$(X$, 125, 4))
       REFERE$ = Mid$(X$, 86, 9)
       REFOP$ = "(cerr) " + Mid$(X$, 5, 60)
       CODREPA$ = TRIM$(Mid$(X$, 497, 16))
       Call REPLA(X$, REFOP$, 5, 60)
       Call REPLA(X$, MKI$(3), 65, 2) ' MARCA CERRADA
       Call GRAREG("ORDEREPA", X$, U&)
       GoTo 20
     End If
   Next U&
   '
20 NOREX$ = TRIM$(Str$(NOREP&))
   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = "OR-" + NOREX$
   '
   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
   RCC$ = ""
   URE& = ULTREG&("RESERVA")
   For KK& = 1 To Len(RFF$) Step 4
     RGG& = CVS(Mid$(RFF$, KK&, 4))
     If RGG& > 0 Then
       If RGG& <= URE& Then
         XX$ = REGLEIDO$("RESERVA", RGG&)
         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
           Call GRAREG("RESERVA", String$(128, 0), RGG&)
         End If
       End If
     End If
   Next KK&
   '
   FECIE% = HOY(HOS$)
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(ZZ$, 1, 2))
     CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
     '
     USOUNI# = CAN
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           DOPRI$ = NOREX$
           DOSEC$ = DOPRI$
           DEPX% = 1
           If TRIM$(REFERE$) <> "" Then DOSEC$ = REFERE$
           REFE$ = "Reparación " + CODREPA$
           If Abs(CAN) >= 0.05 Then
              DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
              Call MOVISTO(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
           End If
         End If
       End If
     End If
   Next U&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("PRESREPA")
   Call CIERRARCH("ORDEREPA")
   Call CIERRARCH("RESERVA")
   Screen.MousePointer = 1
   '
   Call COMUNI("Cierre " + NOREX$ + " Completo")
   '
99 Command15.Enabled = True

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'Call CONECRUN("REMITO02", "Remito de Expedición")
   '
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
10 FOREMIBOM.REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", UI&)
   SEREF$ = "S/No." + TRIM$(Mid$(XX$, 71, 15)) + " - Ref:" + TRIM$(Mid$(XX$, 86, 9))
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
   FOREMIBOM.Text1 = TRIM$(Str$(NCLIE%))
   FOREMIBOM.Text1.TEXT = TRIM$(Str$(NCLIE%))
   FOREMIBOM.Text2.TEXT = RASOCLI$(NCLIE%)
   FOREMIBOM.Text3.TEXT = DOMICLI$(NCLIE%)
   FOREMIBOM.Text4.TEXT = CPOSCLI$(NCLIE%)
   FOREMIBOM.Text5.TEXT = LOCACLI$(NCLIE%)
   Call CARTRANSPOR
   FOREMIBOM.Text22 = TRIM$(Mid$(XX$, 129, 368)) + Chr$(13) + Chr$(10) + SEREF$
   FOREMIBOM.Show 1
   '
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
3   Call SELECHO("PRESREPA/Indice Presupuestos de Reparación")
    NPX& = Val(VALACT1$("PRESREPA"))
    If NPX& < 1 Then GoTo 99
    '
    NUDOCU$ = TRIM$(Str$(NPX&))
    For UI& = ULTREG&("PRESREPA") To 1 Step -1
      XX$ = REGLEIDO$("PRESREPA", UI&)
      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
    Next UI&
    Call COMUNI("Error de Indices - Consulte.")
    GoTo 99
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    TIDOCUM$ = "Presupuesto de Reparación " + NUDOCU$
    TIDOCU2$ = "Operaciones de Reparación PR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
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
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU2$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 15
      End If
    Next UI&
    '
15  REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
    GoTo 18
    '
17  Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
18  While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 19
      L& = L& - 1
    Wend
    '
19  For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
50  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
       COPRIAT$ = Left$(TPSP$, PPXX% - 1)
       TPSP$ = Mid$(TPSP$, PPXX% + 1)
       '
       If Len(COPRIAT$) > 0 Then
         COPRI% = Asc(Left$(COPRIAT$, 1))
         ATRIP$ = Mid$(COPRIAT$, 2)
         Call SHOWCOMMAND(COPRI%, ATRIP$)
       End If
       GoTo 50
       '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Visible = False
      If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    '
3   Call SELECHO("ORDEREPA/Indice General Ordenes de Reparación")
    NPX& = Val(VALACT1$("ORDEREPA"))
    If NPX& < 1 Then GoTo 99
    '
    For UI& = ULTREG&("ORDEREPA") To 1 Step -1
      XX$ = REGLEIDO$("ORDEREPA", UI&)
      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
    Next UI&
    Call COMUNI("Error de Indices - Consulte.")
    Exit Sub
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCUM$ = "Orden de Reparación " + NUDOCU$
    TIDOCU2$ = "Procedimiento de Reparación OR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
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
    GoTo 99
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
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
10  For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU2$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 15
      End If
    Next UI&
    '
15  REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
16  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 17
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 17
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 16
    If DCC& > DOCUNU& Then LS& = L&: GoTo 16
    GoTo 18
    '
17  Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
18  While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 19
      L& = L& - 1
    Wend
    '
19  For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
    FIN& = Len(TPSP$)
50  PPXX% = InStr(TPSP$, Chr$(255))
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
      GoTo 50
      '
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
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command4.Enabled = True
    '
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
3  Screen.MousePointer = 11
   Call COPYSTRU("ORDEREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("ORDEREPA")
     XX$ = REGLEIDO$("ORDEREPA", U&)
     If CVI(Mid$(XX$, 65, 2)) < 3 Then
       JJ& = JJ& + 1
       Call GRAREG("PREPAEMI", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Reparación en Curso\que Puedan Anularse.")
     GoTo 99
   End If
   '
   Call SELECHO("PREPAEMI")
   NOREP& = Val(VALACT1$("PREPAEMI"))
   If NOREP& < 1 Then GoTo 99
   '
   If COMALTER%("Confirma Anulación de O/R " + TRIM$(Str$(NOREP&)) + "\" + TRIM$(VALACT2$("PREPAEMI")) + "\\Cancelar\Anular O/R") <> 2 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   NUPRESU& = 0
   For U& = 1 To ULTREG&("ORDEREPA")
     X$ = REGLEIDO$("ORDEREPA", U&)
     If CVS(Left$(X$, 4)) = NOREP& Then
       REFOP$ = "(anul) " + Mid$(X$, 5, 60)
       NUPRESU& = CVS(Mid$(X$, 125, 4))
       Call REPLA(X$, REFOP$, 5, 60)
       Call REPLA(X$, MKI$(9), 65, 2) ' MARCA ANULADA
       Call GRAREG("ORDEREPA", X$, U&)
       GoTo 20
     End If
   Next U&
   '
20 If NUPRESU& > 0 Then
     For U& = 1 To ULTREG&("PRESREPA")
       X$ = REGLEIDO$("PRESREPA", U&)
       If CVS(Left$(X$, 4)) = NUPRESU& Then
         Call REPLA(X$, MKS$(0), 125, 4) ' QUITA NUMERO DE O/R
         Call GRAREG("PRESREPA", X$, U&)
         GoTo 30
       End If
     Next U&
   End If
   '
30 NOREX$ = TRIM$(Str$(NOREP&))
   While Len(NOREX$) < 6: NOREX$ = "0" + NOREX$: Wend
   NOREX$ = "OR-" + NOREX$
   '
   RFF$ = RECFILT$("RESERVA", 3, NOREX$)
   RCC$ = ""
   URE& = ULTREG&("RESERVA")
   For KK& = 1 To Len(RFF$) Step 4
     RGG& = CVS(Mid$(RFF$, KK&, 4))
     If RGG& > 0 Then
       If RGG& <= URE& Then
         XX$ = REGLEIDO$("RESERVA", RGG&)
         If TRIM$(Mid$(XX$, 3, 12)) = NOREX$ Then
           Call GRAREG("RESERVA", String$(128, 0), RGG&)
         End If
       End If
     End If
   Next KK&
   '
   Call CIERRARCH("PRESREPA")
   Call CIERRARCH("ORDEREPA")
   Call CIERRARCH("RESERVA")
   Screen.MousePointer = 1
   '
   Call COMUNI("Anulación Completa")
   '
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   '
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
       JJ& = JJ& + 1
       Call GRAREG("PREPAEMI", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Pendientes\para Emitir Presupuestos de Reparación.")
     GoTo 99
   End If
   '
2  VDEF$ = MKS$(0) + Space$(16)
3  OBX$ = OBJPLA$("ENCAREPA", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   NPX& = CVS(Left$(OBX$, 4))
   If NPX& < 1 Or ECOARCH$("PREPAEMI", MKS$(NPX&)) = "" Then GoTo 3
   NUPRESUP& = NPX&
   '
   COD$ = TRIM$(Mid$(OBX$, 5, 16))
   If COD$ = "" Or ECOARCH$("INDIREP", COD$) = "" Then GoTo 3
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   GoTo 99
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      Call COMUNI("Imposible Generar / Modificar Presupuesto.\Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   '
10 NC% = CVI(Mid$(XX$, 67, 2))
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       PREREBOM.Label18 = TRIM$(Str$(NPX&))
       PREREBOM.Label11 = TRIM$(Str$(NC%))
       PREREBOM.Label10 = COD$
       '
       Screen.MousePointer = 11
       RNC$ = RECFILT$("DESCREP", 1, COD$)
       URN& = ULTREG&("DESCREP")
       '
       For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
           If RENOTA& <= URN& Then
             TTYY$ = RTrim$(Mid$(REGLEIDO$("DESCREP", RENOTA&), 17))
             If TTYZ$ <> "" Then
               TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
             End If
             TTYZ$ = TTYZ$ + RTrim$(TTYY$)
           End If
         End If
       Next U&
       PREREBOM.Label8 = TTYZ$
       '
       COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
       If COEFI < 1 Then COEFI = 1
       CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
       PREREBOM.Label20 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
       '
       PREREBOM.List1.Clear
       SUREPUES = 0
       Call SETULTREG("!MATEROF", 0)
       '
       CODD$ = COD$
       Screen.MousePointer = 11
       RFF$ = RECFILT$("REPUESRE", 1, CODD$)
       For IK& = 1 To Len(RFF$) Step 4
          RERE& = CVS(Mid$(RFF$, IK&, 4))
          If RERE& > 0 Then
            If RERE& <= ULTREG&("REPUESRE") Then
              XX$ = REGLEIDO$("REPUESRE", RERE&)
              CIY% = CVI(Mid$(XX$, 17, 2))
              CODXX$ = CODEXT$(CIY%)
              DEXX$ = DESCRIT$(CIY%)
              COSUN = COEFI * COSUNI(CIY%)
              COSUX$ = CSTRING$(MKS$(COSUN), 4, 10, 2, 2)
              USO = CVS(Mid$(XX$, 19, 4))
              USOX$ = CSTRING$(MKS$(USO), 4, 10, 1, 2)
              COSTO = COSUN * USO
              COSTX$ = CSTRING$(MKS$(COSTO), 4, 12, 2, 2)
              SUREPUES = SUREPUES + COSTO
              '
              XXX$ = Space$(64)
              Call REPLA(XXX$, CODXX$, 1, 16)
              Call REPLA(XXX$, DEXX$, 17, 38)
              Call REPLA(XXX$, USOX$, 54, 10)
              Call REPLA(XXX$, COSUX$, 64, 10)
              Call REPLA(XXX$, COSTX$, 74, 12)
              PREREBOM.List1.AddItem XXX$
            End If
          End If
       Next IK&
       PREREBOM.SUMAREP = CSTRING$(MKS$(SUREPUES), 4, 12, 2, 2)
       '
       SUTIEREP = 0
       RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
       For I1& = 1 To Len(RFFX$) Step 4
         I& = CVS(Mid$(RFFX$, I1&, 4))
         X$ = REGLEIDO$("SECOPREP", I&)
         '
         CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
         SUTIEREP = SUTIEREP + CAOPS * CVS(Mid$(X$, 73, 4))
       Next I1&
       PREREBOM.Label19 = CSTRING$(MKS$(SUTIEREP), 4, 8, 1, 2)
       '
       Screen.MousePointer = 1
       '
       PREREBOM.Label18 = TRIM$(Str$(NUPRESUP&))
       PREREBOM.Show 1
       GoTo 2
     End If
   End If
   '
99 Command6.Enabled = True

End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ANOTAREP.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RECEBORE.Show 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVI(Mid$(XX$, 121, 2)) > 0 Then
       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
         JJ& = JJ& + 1
         Call GRAREG("PREPAEMI", XX$, JJ&)
       End If
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
   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then GoTo 99
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
   If CVI(Mid$(XX$, 121, 2)) < 1 Then
      Call COMUNI("Falta Generar y Emitir Presupuesto.")
      GoTo 99
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      Call COMUNI("Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   '
   ORDREBOM.Label18 = TRIM$(Str$(NPX&))
   ORDREBOM.Label2 = TRIM$(Str$(NUORDREP&))
   '
   NPP& = 1
   FIN& = ULTREG&("PEDENCA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 16 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PEDENCA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   ORDREBOM.Label25 = TRIM$(Str$(NPP&))
   '
   ORDREBOM.Show 1
   GoTo 3
   '
99 Command9.Enabled = True
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

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    'If HOY(HOS$) > FECHANUM("15/02/03") Then
    '  Call CONECRUN("REPABO02", "")
    '  Call FINAL("")
    'End If
    '
    CODGENER$ = TRIM$(CONTROL$("REPARA", "CODGENER"))
    If CODINT%(CODGENER$) < 1 Then
       Call MENSERR(24, "Falta Código Genérico Reparaciones.")
       CONFREPA.Show 1
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub CARTRANSPOR()
    '
    NCLIE% = Val(FOREMIBOM.Text1.TEXT)
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
    FOREMIBOM.Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

