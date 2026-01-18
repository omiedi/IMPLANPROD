VERSION 5.00
Begin VB.Form COBAVEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Básico de Ventas"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3990
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resúmenes Estadísticos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Index           =   1
      Left            =   210
      TabIndex        =   7
      Top             =   1575
      Width           =   3585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   945
         Width           =   2430
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
            Picture         =   "COBAVEN.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Ranking de Ventas por Cliente por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Ventas por Cliente"
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
            TabIndex        =   9
            Top             =   105
            Width           =   1785
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   400
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   10
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "COBAVEN.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Ranking de Ventas por Articulo por Rango de Fechas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking de Ventas por Producto"
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
            TabIndex        =   11
            Top             =   75
            Width           =   1740
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3990
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   5
      Top             =   2625
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3990
      Picture         =   "COBAVEN.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   855
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
      Height          =   855
      Left            =   3990
      Picture         =   "COBAVEN.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gestión Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   3585
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   420
         Width           =   2430
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
            Picture         =   "COBAVEN.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Acceso al Control General de Gestión Comercial"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a Control Gestión Comercial"
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
            Left            =   650
            TabIndex        =   2
            Top             =   85
            Width           =   1635
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   3780
      X2              =   5145
      Y1              =   3370
      Y2              =   3370
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3440
      Picture         =   "COBAVEN.frx":0EAA
      Top             =   2880
      Width           =   2010
   End
End
Attribute VB_Name = "COBAVEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub RANKCLI()
    '
    Dim TOTVENTA#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    VENDSEL% = 0
    RSX$ = REGSEL$("VENDEDOR")
    If Len(RSX$) >= 4 Then
      VENDSEL% = Asc(Left$(RSX$, 1))
      PERIO$ = PERIO$ + " - Vendedor: " + TRIM$(Mid$(RSX$, 2, 30))
    End If
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    FIN& = ULTREG&("FACTUR")
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      FF% = CVI(Mid$(X$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          NC% = CVI(Left$(X$, 2))
          If NC% > 0 Then
            If NC% <= 32767 Then
              If VENDSEL% = 0 Or VENDCLI%(NC%) = VENDSEL% Then
                    TI% = Asc(Mid$(X$, 13, 1))
                    If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
                    If TI% > 0 Then
                       If TI% < 7 Then
                          NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100
                          If TI% > 3 Then
                             NETO# = (-1) * Abs(NETO#)
                          End If
                          TOTVENTA#(NC%) = TOTVENTA#(NC%) + NETO#
                          TOTAGEN# = TOTAGEN# + NETO#
                       End If
                    End If
              End If
            End If
          End If
        End If
      End If
    Next U&
    '
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          COBAVEN.List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = COBAVEN.List1.ListCount
    '
    For U& = 1 To FIN&
        COBAVEN.List1.ListIndex = U& - 1
        NC% = Val(Mid$(COBAVEN.List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(COBAVEN.List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(COBAVEN.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    COBAVEN.List1.Clear
    '
    'If COBAVEN.Option1(4).Value = True Then
    '     Call FINAL("*RANKCLI")
    '   Else
         Call FINAL("*LRANCLI")
    'End If
    GoTo 20
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub

Sub RANKPIE()
    '
    Dim TOTVENTA#(32767)
    Dim TOTCANT#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '\\\OT-05-0594-WAR-22/08/05///
    OBX$ = OBJPLA$("DESHASFECHMAT", VDEF$)
    '\\\OT-05-0594-WAR-FIN///
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    TIPOSEL% = Asc(Mid$(OBX$, 5, 1))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("MOVISTO")
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                 If CI% <= 32767 Then
                  If TIPOSEL% >= 0 Then '\\\OT-05-0594-WAR-22/08/05///
                    If TIMATE%(CI%) = TIPOSEL% Or TIPOSEL% = 0 Then
                      COMO$ = UCase$(Mid$(X$, 33, 2)) ' Doc. secundario
                      If TRIM$(COMO$) = "" Then
                         COMO$ = UCase$(Mid$(X$, 23, 2)) ' Doc. primario
                      End If
                      If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                         NETO# = (CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))) * CVD(Mid$(X$, 73, 8))
                         CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                         TOTVENTA#(CI%) = TOTVENTA#(CI%) + NETO#
                         TOTCANT#(CI%) = TOTCANT#(CI%) + CANTI
                         TOTAGEN# = TOTAGEN# + NETO#
                      End If
                    End If
                  End If
                End If
              End If
           End If
        End If
    Next U&
    '
    For NCI% = 1 To NUMAS%
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
             RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
             PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
             RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
             RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
             RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
            COBAVEN.List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0
    FIN& = COBAVEN.List1.ListCount
    '
    For U& = 1 To FIN&
        COBAVEN.List1.ListIndex = U& - 1
        CI% = Val(Mid$(COBAVEN.List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(COBAVEN.List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(COBAVEN.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(CI%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTCANT#(CI%)), 21, 8)
        Call GRAREG("RANPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANPIE", J&)
    Call CIERRARCH("RANPIE")
    Call HEADERS("RANPIE", "")
    Call HEADERS("RANPIE", "Periodo: " + PERIO$)
    '\\\OT-05-0594-WAR-22/08/05///
    TIMA1$ = ECOARCH$("TATIMAT", MKI$(TIPOSEL%))
    If TRIM(TIMA1$) = "" Then
      TIMA1$ = "Sin Clasificación"
    End If
    Call HEADERS("RANPIE", "Tipo de Material: " + TIMA1$)
     '\\\OT-05-0594-WAR-FIN///
    Screen.MousePointer = 1
    COBAVEN.List1.Clear
    '
    Call FINAL("*LRNKPIE")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'\\\OT-05-0674-WAR-13/10/05///
Sub RANKRES() 'RANKING DE VTAS. RESUMEN POR TIPO DE MATERIAL
    '
    Dim TIMAT%(255) '255 MAXIMO CODIGO DE TIPO DE MATERIAL
    Dim TOTVENTA#(255)
    Dim TOTCANT#(255)
'    CR% = Val(ULTREG&("TATIMAT"))
'    ReDim TIMAT%(CR%)
'    ReDim TOTVENTA#(CR%)
'    ReDim TOTCANT#(CR%)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("MOVISTO")
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                 If CI% <= 32767 Then
                  TM% = TIMATE%(CI%)
                  COMO$ = UCase$(Mid$(X$, 33, 2)) ' Doc. secundario
                  If TRIM$(COMO$) = "" Then
                    COMO$ = UCase$(Mid$(X$, 23, 2)) ' Doc. primario
                  End If
                  If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                    NETO# = (CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))) * CVD(Mid$(X$, 73, 8))
                    CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                    TOTVENTA#(TM%) = TOTVENTA#(TM%) + NETO#
                    TOTCANT#(TM%) = TOTCANT#(TM%) + CANTI
                    TOTAGEN# = TOTAGEN# + NETO#
                  End If
                End If
              End If
           End If
        End If
    Next U&
    '
    For TTM% = 0 To 255
       If Abs(TOTVENTA#(TTM%)) > 0.005 Then
         RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(TTM%)), 4, 12, 2, 2)
         PORPAR = 100 * TOTVENTA#(TTM%) / TOTAGEN#
         RELIB$ = RELIB$ + CSTRING$(MKI$(TTM%), 1, 6, 0, 2)
         RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
         RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(TTM%)), 4, 12, 2, 2)
         COBAVEN.List1.AddItem RELIB$
       End If
    Next TTM%
    '
    J& = 0
    FIN& = COBAVEN.List1.ListCount
    '
    For U& = 1 To FIN&
        COBAVEN.List1.ListIndex = U& - 1
        TTM% = Val(Mid$(COBAVEN.List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(COBAVEN.List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(COBAVEN.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKTIMA")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(TTM%), 3, 1)
          DESTM$ = ECOARCH("TATIMAT", MKI$(TTM%))
          If TTM% = 0 Then DESTM$ = "Sin Clasificación"
        Call REPLA(Z$, DESTM$, 4, 20)
        Call REPLA(Z$, MKD$(TONEVEN#), 28, 8)
        Call REPLA(Z$, MKS$(PORPAR), 36, 4)
        Call REPLA(Z$, MKS$(PORACU), 40, 4)
        Call REPLA(Z$, MKS$(TOTCANT#(TTM%)), 24, 4)
        Call GRAREG("RANKTIMA", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKTIMA", J&)
    Call CIERRARCH("RANKTIMA")
    Call HEADERS("RANKTIMA", "")
    Call HEADERS("RANKTIMA", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    COBAVEN.List1.Clear
    '
    Call FINAL("*LRNKMAT")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'\\\OT-05-0674-WAR-FIN///
Private Sub Command1_Click()
'
Command1.Enabled = False
'\\\OT-05-0674-WAR-13/10/05///
RTA% = ALTERNA%("Detalle\Resumen")
If RTA% = 1 Then
  Call RANKPIE
ElseIf RTA% = 2 Then
  Call RANKRES 'NUEVO
End If
'\\\OT-05-0674-WAR-FIN///
Command1.Enabled = True
'
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command4_Click()
'
Command4.Enabled = False
COGEVE00.Show 1
Command4.Enabled = True
'
End Sub

Private Sub Command6_Click()
'
Command6.Enabled = False
Call RANKCLI
Command6.Enabled = True
'
End Sub
