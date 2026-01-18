VERSION 5.00
Begin VB.Form ESTADIPRO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadistica de Entregas"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
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
   ScaleHeight     =   5535
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "List"
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
      Left            =   9870
      Picture         =   "ESTADIPRO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Listado Resumen de Consumo Mensual por Cliente"
      Top             =   3570
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   4425
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   15
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3045
         TabIndex        =   14
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Número de Cliente:"
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
         Left            =   945
         TabIndex        =   13
         Top             =   525
         Width           =   1905
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   12
         Top             =   840
         Width           =   3975
      End
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
      Left            =   9870
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   2310
      Width           =   855
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
         TabIndex        =   10
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   4830
      TabIndex        =   4
      Top             =   105
      Width           =   4845
      Begin VB.CommandButton Command2 
         Caption         =   "Calc"
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
         Left            =   3885
         Picture         =   "ESTADIPRO.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cálculo:"
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
         TabIndex        =   8
         Top             =   945
         Width           =   750
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   7
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango:"
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
         TabIndex        =   6
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   5
         Top             =   420
         Width           =   2745
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
      Height          =   855
      Left            =   9870
      Picture         =   "ESTADIPRO.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cierra y Abandona la Aplicación"
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
      Left            =   9870
      Picture         =   "ESTADIPRO.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
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
      Left            =   9870
      Picture         =   "ESTADIPRO.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprime Ficha de Distribución de Entregas"
      Top             =   4515
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "List"
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
      Left            =   9870
      Picture         =   "ESTADIPRO.frx":0F12
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Listado Detalle Estadístico de Entregas Valorizadas"
      Top             =   2625
      Width           =   855
   End
End
Attribute VB_Name = "ESTADIPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
    
    Dim PUNMO$(32767), TICLIE%(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
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
    Label1.Caption = PERIO$
    Label1.Refresh
    '
    FIN& = ULTREG&("ABONOS")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      TTXX$ = REGLEIDO$("ABONOS", U&)
      NCLI% = CVI(Left$(TTXX$, 2))
      IABO# = CVD(Mid$(TTXX$, 5, 8))
      If NCLI% > 0 Then
        If IABO# >= 0.5 Then
          TICLIE%(NCLI%) = 1
        End If
      End If
    Next U&
    '
    'GRUCLIEN% = COMALTER%("Grupo de Clientes a Listar:\\Abonados\No Abonados\Todos")
    'If GRUCLIEN% < 1 Or GRUCLIEN% > 3 Then GoTo 99
    'If GRUCLIEN% = 2 Then GRUCLIEN% = 0
    GRUCLIEN% = 3  ' todos
    '
    Screen.MousePointer = 11
    '
    Call HEADERS("DETCLIEN", "")
    Call HEADERS("DETCLIEN", "Periodo: " + PERIO$)
    '
    FIN& = ULTREG&("MOVISTO")
    J& = 0
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           CLI% = CVI(Mid$(X$, 84, 2))
           If CLI% > 0 Then
             If CLI% <= 32767 Then
               If GRUCLIEN% = 3 Or TICLIE%(CLI%) = GRUCLIEN% Then
                 TIMO1$ = UCase$(Mid$(X$, 23, 2))
                 TIMO2$ = UCase$(Mid$(X$, 33, 2))
                 If TIMO1$ = "RT" Or TIMO2$ = "RT" Then
                    PUNMO$(CLI%) = PUNMO$(CLI%) + MKS$(U&)
                 End If
               End If
             End If
           End If
         End If
       End If
    Next U&
    '
    UREMO& = ULTREG&("MOVISTO")
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Y$ = REGLEIDO$("DEUDOR1", U&)
       CLI% = CVI(Mid$(Y$, 1, 2))
       If CLI% > 0 Then
          If CLI% <= 32767 Then
             For KKI% = 1 To Len(PUNMO$(CLI%)) Step 4
                REG& = CVS(Mid$(PUNMO$(CLI%), KKI%, 4))
                If REG& > 0 Then
                   If REG& <= UREMO& Then
                      X$ = REGLEIDO$("MOVISTO", REG&)
                      CLI% = CVI(Mid$(X$, 84, 2))
                      FF% = CVI(Left$(X$, 2))
                      CII% = CVI(Mid$(X$, 3, 2))
                      REMIFAC$ = Mid$(X$, 23, 10)
                      CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      VUNI = CVD(Mid$(X$, 73, 8))
                      PRELI = PRELISTA#("LIPRE001", CII%)
                      VALOUN = VUNI: If VUNI < 0.00005 Then VALOUN = PRELI
                      VATOT# = CANTU# * VALOUN
                      '
                      Z$ = REGBLAN$("DETCLIEN")
                      Call REPLA(Z$, MKI$(CLI%), 1, 2)
                      Call REPLA(Z$, MKI$(FF%), 3, 2)
                      Call REPLA(Z$, REMIFAC$, 5, 10)
                      Call REPLA(Z$, MKI$(CII%), 15, 2)
                      Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                      Call REPLA(Z$, MKS$(VUNI), 25, 4)
                      Call REPLA(Z$, MKS$(PRELI), 29, 4)
                      Call REPLA(Z$, MKD$(VATOT#), 33, 8)
                      '
                      J& = J& + 1
                      Call GRAREG("DETCLIEN", Z$, J&)
                   End If
                End If
             Next KKI%
             PUNMO$(CLI%) = ""
          End If
       End If
    Next U&
    '
    FIN& = 32767
    For U& = 1 To 32767
       Call TRACE(20, U&, FIN&)
       If PUNMO$(U&) <> "" Then
          For KKI% = 1 To Len(PUNMO$(U&)) Step 4
             REG& = CVS(Mid$(PUNMO$(U&), KKI%, 4))
             If REG& > 0 Then
                If REG& <= UREMO& Then
                   X$ = REGLEIDO$("MOVISTO", REG&)
                   CLI% = CVI(Mid$(X$, 84, 2))
                   FF% = CVI(Left$(X$, 2))
                   CII% = CVI(Mid$(X$, 3, 2))
                   REMIFAC$ = Mid$(X$, 23, 10)
                   CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                   VUNI = CVD(Mid$(X$, 73, 8))
                   PRELI = PRELISTA#("LIPRE001", CII%)
                   VALOUN = VUNI: If VUNI < 0.00005 Then VALOUN = PRELI
                   VATOT# = CANTU# * VALOUN
                   '
                   Z$ = REGBLAN$("DETCLIEN")
                   Call REPLA(Z$, MKI$(CLI%), 1, 2)
                   Call REPLA(Z$, MKI$(FF%), 3, 2)
                   Call REPLA(Z$, REMIFAC$, 5, 10)
                   Call REPLA(Z$, MKI$(CII%), 15, 2)
                   Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                   Call REPLA(Z$, MKS$(VUNI), 25, 4)
                   Call REPLA(Z$, MKS$(PRELI), 29, 4)
                   Call REPLA(Z$, MKD$(VATOT#), 33, 8)
                   '
                   J& = J& + 1
                   Call GRAREG("DETCLIEN", Z$, J&)
                End If
             End If
          Next KKI%
       End If
    Next U&
    '
    Call SETULTREG("DETCLIEN", J&)
    Call CIERRARCH("DETCLIEN")
    Call LIBARCH("DETCLIEN")
    '
    HOII% = HOY(HOS$)
    Call GRACONTROL("ESTADIP", "PERICALC", PERIO$)
    Label1.Caption = Control$("ESTADIP", "PERICALC")
    Call GRACONTROL("ESTADIP", "FECHACAL", HOS$ + " - " + Time$ + " hs")
    Label3.Caption = Control$("ESTADIP", "FECHACAL")
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command3_Click()
   PRINTFORM
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call FINAL("*DETACLI")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Call SELECHO("DEUDOR1")
   NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% > 0 Then
      Label8.Caption = TRIM$(Str$(NCXI%))
      Label6.Caption = RASOCLI$(NCXI%)
      ESTADIPRO.Refresh
      Call ESTADIP
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   NC% = Val(Label8.Caption)
   If NC% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
       Screen.MousePointer = 11
       ENCA1$ = "": ENCA2$ = ""
       Dim DETCAN$(16384), TOCAN(16384), TOIMPO#(16384)
       FIN& = ULTREG&("DETCLIEN")
       '
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         Z$ = REGLEIDO$("DETCLIEN", U&)
         CLI% = CVI(Left$(Z$, 2))
         If CLI% = NC% Then
           FEX = CVI(Mid$(Z$, 3, 2))
           REMI& = Val(Mid$(Z$, 8, 6))
           CI% = CVI(Mid$(Z$, 15, 2))
           CAN# = CVD(Mid$(Z$, 17, 8))
           VALO# = CVD(Mid$(Z$, 33, 8))
           '
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               REMIS$ = CSTRING$(MKS$(REMI&), 3, 6, 0, 2)
               For JJ% = 1 To Len(ENCA1$) Step 6
                 If Mid$(ENCA1$, JJ%, 6) = REMIS$ Then
                   COLUI% = JJ%
                   GoTo 15
                 End If
               Next JJ%
               ENCA1$ = ENCA1$ + REMIS$
               ENCA2$ = ENCA2$ + " " + Left$(FECHATEX$(FEX), 5)
               COLUI% = Len(ENCA1$) - 5
               '
15             While Len(DETCAN$(CI%)) < Len(ENCA1$)
                 DETCAN$(CI%) = DETCAN$(CI%) + Space$(6)
               Wend
               '
               TOCAN(CI%) = TOCAN(CI%) + CAN#
               TOIMPO#(CI%) = TOIMPO#(CI%) + VALO#
               CAN# = CAN# + Val(Mid$(DETCAN$(CI%), COLUI%, 6))
               CANS$ = CSTRING$(MKD$(CAN#), 3, 6, 0, 2)
               TTXX$ = DETCAN$(CI%)
               Call REPLA(TTXX$, CANS$, COLUI%, 6)
               DETCAN$(CI%) = TTXX$
               
               '
             End If
           End If
         End If
       Next U&
       '
       If Len(ENCA1$) > 72 Then
         Call MENSERR(24, "Proceso Imposible - Demasiados Movimientos.\   \Elija Período más Corto.")
         Exit Sub
       End If
       '
       ENCA1$ = Space$(48) + Left$(ENCA1$ + Space$(72), 72) + "     Total    Precio"
       ENCA2$ = "Codigo          Descripcion                     " + Left$(ENCA2$ + Space$(72), 72) + "   Consumo      Unit    Subtotal"
       '
       JK& = 0: RELI$ = REGBLAN$("COMENSU")
       For U& = 1 To NUINV%
         X$ = REGLEIDO$("INVERT", U& + 1)
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If Len(DETCAN$(CI%)) > 0 Then
               Z$ = RELI$
               Call REPLA(Z$, MKI$(CI%), 1, 2)
               Call REPLA(Z$, DETCAN$(CI%), 3, 72)
               If TRIM$(Mid$(Z$, 3, 1)) = "" Then
                 Call REPLA(Z$, ".", 3, 1)
               End If
               Call REPLA(Z$, MKS$(TOCAN(CI%)), 75, 4)
               Call REPLA(Z$, MKD$(TOIMPO#(CI%)), 83, 8)
               If TOCAN(CI%) > 0.5 Then
                 PREUNI = TOIMPO#(CI%) / TOCAN(CI%)
                 Call REPLA(Z$, MKS$(PREUNI), 79, 4)
               End If
               JK& = JK& + 1
               Call GRAREG("COMENSU", Z$, JK&)
             End If
           End If
         End If
       Next U&
       '
       Call SETULTREG("COMENSU", JK&)
       Call CIERRARCH("COMENSU")
       PERIOCLI$ = TRIM$(Label6.Caption) + " (" + TRIM$(Label1.Caption) + ")"
       Call HEADERS("COMENSU", "")
       Call HEADERS("COMENSU", "Corresponde a: " + PERIOCLI$)
       Call ENCACOL("COMENSU", "")
       Call ENCACOL("COMENSU", ENCA1$)
       Call ENCACOL("COMENSU", ENCA2$)
       '
       Screen.MousePointer = 1
       Call FINAL("*COMENSU")
       '
     End If
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Form_Load()
    Label1.Caption = Control$("ESTADIP", "PERICALC")
    Label3.Caption = Control$("ESTADIP", "FECHACAL")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Sub ESTADIP()
   '
   Screen.MousePointer = 11
   Dim TOVEN(32767), DEGRU$(32767)
   '
   ESTADIPRO.DrawWidth = 2
   ESTADIPRO.ScaleMode = 1
   Line (0, 1680)-(10965, 5880), ESTADIPRO.BackColor, BF
   TOTAVEN = 0
   NC% = Val(Label8.Caption)
   FIN& = ULTREG&("DETCLIEN")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("DETCLIEN", U&)
      CLI% = CVI(Left$(Z$, 2))
      If CLI% = NC% Then
         CI% = CVI(Mid$(Z$, 15, 2))
         GRUCO% = GRUCOVE%(CI%)
         TOVEN(GRUCO%) = TOVEN(GRUCO) + CVD(Mid$(Z$, 33, 8))
         TOTAVEN = TOTAVEN + CVD(Mid$(Z$, 33, 8))
      End If
   Next U&
   '
   If TOTAVEN < 1 Then
      Screen.MousePointer = 1
      Call MENSERR(24, "Sin Movimiento")
      Exit Sub
   End If
   '
   GRUCOMA% = 0
   TOVEN(0) = TOTAVEN
   DEGRU$(0) = "Varios"
   For U& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO$("GRUCOVEN", U&)
      GRUCO% = CVI(Left$(X$, 2))
      DEGRU$(GRUCO%) = Mid$(X$, 3, 30)
      If GRUCO% > GRUCOMA% Then GRUCOMA% = GRUCO%
      TOVEN(0) = TOVEN(0) - TOVEN(GRUCO%)
      'List1.AddItem TTT$
   Next U&
   '
   OFV = 1680: OFH = 210
   INI = 0: VERTI = 0
   ESTADIPRO.DrawWidth = 2
   ESTADIPRO.ScaleMode = 1
   For UI% = 0 To GRUCOMA%
     ROJO = ROJO + 192
     If ROJO > 255 Then VERDE = VERDE + 192: ROJO = ROJO - 256
     If VERDE > 255 Then AZUL = AZUL + 192: VERDE = VERDE - 256
     If AZUL > 256 Then AZUL = AZUL - 256
     If TOVEN(UI%) >= 0.005 Then
       ESTADIPRO.ForeColor = RGB(ROJO, VERDE, AZUL)
       '
       FINI = INI + 6.28 * TOVEN(UI%) / TOTAVEN
       For RADIO = 1 To 1500 Step 5
         ESTADIPRO.Circle (7300, 3500), RADIO, , INI, (FINI + INI) / 2
         ESTADIPRO.Circle (7300, 3500), RADIO, , (FINI + INI) / 2, FINI
       Next RADIO
       INI = FINI
       '
       VERTI = VERTI + 200
       X1 = 0 + OFH: X2 = 300 + OFH
       Y1 = VERTI + OFV: Y2 = VERTI + 120 + OFV
       Line (X1, Y1)-(X2, Y2), RGB(ROJO, VERDE, AZUL), BF
       TTT$ = CSTRING$(MKI$(UI%), 1, 3, 0, 2) + "  " + AJUSTI$(DEGRU$(UI%), 21) + CSTRING$(MKS$(TOVEN(UI%)), 4, 12, 2, 2)
       CurrentX = 350 + OFH
       CurrentY = VERTI - 50 + OFV
       ForeColor = QBColor(0)
       Print TTT$
     End If
   Next UI%
   '
   VERTI = VERTI + 200
   TTT$ = Space$(26) + "------------"
   CurrentX = 350 + OFH
   CurrentY = VERTI - 50 + OFV
   ForeColor = QBColor(0)
   Print TTT$
   '
   VERTI = VERTI + 200
   CurrentX = 350 + OFH
   CurrentY = VERTI - 50 + OFV
   TTT$ = Space$(5) + AJUSTI$("TOTAL GENERAL", 21) + CSTRING$(MKS$(TOTAVEN), 4, 12, 2, 2)
   Print TTT$
   '
   RFF$ = RECFILT$("ABONOS", 1, MKI$(NC%))
   If Len(RFF$) >= 4 Then
     REABO& = CVS(RFF$)
     If REABO& > 0 Then
       If REABO& <= ULTREG&("ABONOS") Then
         IABO# = CVD(Mid$(REGLEIDO$("ABONOS", REABO&), 5, 8))
         If IABO# >= 0.005 Then
           VERTI = VERTI + 400
           CurrentX = 350 + OFH
           CurrentY = VERTI - 50 + OFV
           TTT$ = Space$(5) + AJUSTI$("ACUERDO", 21) + CSTRING$(MKD$(IABO#), 4, 12, 2, 2)
           Print TTT$
         End If
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
