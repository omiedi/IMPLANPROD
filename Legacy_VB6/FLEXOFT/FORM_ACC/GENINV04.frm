VERSION 5.00
Begin VB.Form GENINV04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Inventarios a Fecha"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Valorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   210
      TabIndex        =   22
      Top             =   1700
      Width           =   4950
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1300
         TabIndex        =   24
         Text            =   " 1"
         Top             =   280
         Width           =   540
      End
      Begin VB.CommandButton Command4 
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
         Height          =   360
         Left            =   1800
         TabIndex        =   23
         Top             =   280
         Width           =   175
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2050
         TabIndex        =   26
         Top             =   285
         Width           =   2700
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
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
         Left            =   -650
         TabIndex        =   25
         Top             =   390
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000040C0&
      Height          =   4635
      Left            =   5490
      ScaleHeight     =   4575
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   -105
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   3800
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
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
         Left            =   105
         Picture         =   "GENINV04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit"
         Top             =   315
         Width           =   640
      End
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
         Height          =   650
         Left            =   105
         Picture         =   "GENINV04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Start"
         Top             =   1995
         Width           =   640
      End
      Begin VB.CommandButton Command29 
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
         Left            =   105
         Picture         =   "GENINV04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Help"
         Top             =   965
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "GENINV04.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4000
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   360
      Left            =   4980
      TabIndex        =   1
      Top             =   105
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   6
      Top             =   2700
      Width           =   4950
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   13
         Top             =   1365
         Width           =   225
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   11
         Top             =   1050
         Width           =   225
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   9
         Top             =   735
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   7
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación Inventario Analitico y Global"
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
         Left            =   1050
         TabIndex        =   14
         Top             =   1365
         Width           =   3585
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicialización de Archivos de Inventarios"
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
         Left            =   1050
         TabIndex        =   12
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Recálculo de Existencias"
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
         Left            =   1050
         TabIndex        =   10
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicialización Archivos de Transferencia"
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
         Left            =   1050
         TabIndex        =   8
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3930
      TabIndex        =   0
      Text            =   " "
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selección de Costos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   4950
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precios de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3675
         TabIndex        =   15
         Top             =   315
         Width           =   1170
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Costos Calculados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2100
         TabIndex        =   4
         Top             =   315
         Width           =   1485
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Costos Informados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   525
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inventario:"
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
      Left            =   1785
      TabIndex        =   5
      Top             =   210
      Width           =   1905
   End
End
Attribute VB_Name = "GENINV04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Command2.Enabled = True
    Hide
    Call FRESHALL
End Sub
'

Private Sub Command2_Click()
    Command2.Enabled = False
    Call GENINV2
    Command2.Enabled = True
End Sub
'

Private Sub Command29_Click()
   Command29.Enabled = False
   Call HELPONLINE("INVENTAR")
   Command29.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHANUM(FECHINVE$) > 0 Then
        Text1.TEXT = FECHINVE$
        Command2.SetFocus
      Else
        Text1.TEXT = HOS$
    End If
End Sub
'

Private Sub Command4_Click()
    '
    '\\\OT-07-0018-RG-23/01/07///
    Call SELECHO("TAMONED")
    MON$ = VALACT1$("TAMONED")
    If MON$ <> "" Then
       Text2.TEXT = MON$
    End If
    '\\\OT-07-0018-RG-FIN///
    '
End Sub


Private Sub Form_Load()
    HO% = HOY(HOS$)
    Text1.TEXT = HOS$
    '
    '\\\OT-07-0018-RG-23/01/07///
    MON% = XVALO(Text2)
    Label2.Caption = ECOARCH$("TAMONED", Chr(MON%))
    '\\\OT-07-0018-RG-FIN///
    '
End Sub '

Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHANUM(FECHINVE$) > 0 Then
        Text1.TEXT = FECHINVE$
        Command2.SetFocus
      Else
        Text1.TEXT = HOS$
    End If
End Sub

Sub GENINV2()
    '
    Dim VESAL$()
    ReDim VESAL$(NUMAS%)
    '
    HOII% = HOY(HOS$)
    FECHINVE$ = Text1
    If FECHANUM(FECHINVE$) < 1 Then
        Text1 = HOS$
        Text1.SetFocus
        Exit Sub
    End If
    '
    '\\\OT-07-0018-RG-23/01/07///
    If Text2 = "" Then
       Call MENSERR(24, "Falta Ingresar el Tipo de Moneda")
       Text2.SetFocus
       Exit Sub
    Else
       MUN% = XVALO(Text2)
       MONDESC$ = ECOARCH$("TAMONED", Chr(MUN%))
       If MUN% <= 0 Or MONDESC$ = "" Then
          Call MENSERR(24, "Tipo de Moneda No Valido")
          Text2 = "1"
          Text2.SetFocus
          Exit Sub
       End If
    End If
    '\\\OT-07-0018-RG-FIN///
    '
    Call GRACONTROL("INVENTAR", "FECHAINV", FECHINVE$)
    Text1 = FECHINVE$
    Text1.Refresh
    Call GRACONTROL("INVENTAR", "FEGENINV", "Generado el " + HOS$ + " - " + Left$(Time$, 5) + " hs.")
    FF% = FECHANUM(FECHINVE$)
    Call GRACONTROL("INVENTAR", "MONEDINV", Text2)
    
    '
    Screen.MousePointer = 11
    '
    Dim COSU%(62), SUCU$(62), DISTKK%(62)
    NUSU% = 1
    For I& = 1 To ULTREG&("SUCURSAL")
      X$ = REGLEIDO$("SUCURSAL", I&)
      If I& <= 62 Then
        If Asc(Left$(X$, 1)) > 32 Then
          NUSU% = I&
          COSU%(I&) = Asc(Mid$(X$, 64, 1))
          SUCU$(I&) = Left$(X$, 20)
          DISTKK%(I&) = 1
          If UCase$(Mid$(X$, 62, 2) = "NO") Then DISTKK%(I&) = 0
        End If
      End If
    Next I&
    Call CIERRARCH("SUCURSAL")
    '
100 GENINV04.Caption = "Generacion de Inventarios al " + FECHINVE$
    '
    Call COPYSTRU("SALDEPOS", "SALDEPO")
    '
    VALORI% = 1
    TICO$ = "Costos: Informados"
    If GENINV04.Option3 = True Then
        VALORI% = 2: TICOS$ = "Costos: Calculados"
      ElseIf GENINV04.Option4 = True Then
        VALORI% = 3: TICOS$ = "Costos: S/Precios de Venta"
        Dim APRELI$(128), MONELI%(128)
        CALIS% = ULTREG&("LISTAS"): If CALIS% > 64 Then CALIS% = 64
        For KKI% = 1 To CALIS%
          ULK& = KKI%
          XX$ = REGLEIDO$("LISTAS", ULK&)
          MONELI%(KKI%) = Asc(Mid$(XX$, 50, 1))
          APRELI$(KKI%) = TRIM$(Str$(KKI%))
          While Len(APRELI$(KKI%)) < 3: APRELI$(KKI%) = "0" + APRELI$(KKI%): Wend
          APRELI$(KKI%) = "LIPRE" + APRELI$(KKI%)
        Next KKI%
    End If
    '
    
' ********************* INICIALIZANDO ARCHIVO DE TRANSFERENCIA
    '
    GENINV04.Check1.Value = 2
    FIN& = NUMAS% + 1
    '
    'Call BLANARCH("SALDEPO")
    'N2% = FILEOPEN%("SALDEPO", 0, 256)
    'Dim SSQ As String * 256
    '
    FIN& = NUMAS% + 1
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        SSQ$ = String$(255, 0) + Chr$(1)
        Call GRAREG("SALDEPO", SSQ$, I&)
    Next I&
    Call SETULTREG("SALDEPO", FIN&)
    GENINV04.Check1.Value = 1
    '
    ' ******************************* RE-CALCULANDO STOCK AL  FFS$
    '
    GENINV04.Check2.Value = 2
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    FEC1# = CDbl(CVDAT(FF%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE int(cdbl(FechMov)) <= " & FEC1# & " "
    
    Set GENIVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With GENIVITMP
      On Error Resume Next
      .MoveLast
      TOT& = .RecordCount
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Movimientos Anteriores la la fecha: \ " & CVDAT(FF%) & "")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      CONTI& = 0
      Do While .EOF = False
       '
       Call TRACE(20, CONTI&, TOT&)
       '
       CI% = !Cod_Inte
       DEP% = !DepoMovi
       LOT$ = "": If IsNull(!IdenLote) = False Then LOT$ = !IdenLote
105    SUCX% = 1
       For U% = 1 To NUSU%
          If COSU%(U%) = DEP% Then SUCX% = U%: GoTo 110
       Next U%
110    CANTI = !CantIngre - !CantSalid
       CANTI = CANTI * DISTKK%(SUCX%)
       RESAL& = CI% + 1
       If CI% > 0 Then
         If CI% <= NUMAS% Then
            While Len(VESAL$(CI%)) < 4 * SUCX%
              VESAL$(CI%) = VESAL$(CI%) + Chr$(0)
            Wend
            SSQ$ = VESAL$(CI%)
            If Len(VESAL$(CI%)) > 0 Then
              'CANUE = CANTI + (Int(10 * CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4)) + 0.5)) / 10
              CANUE = CANTI + CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4))  ' modificado 27/05/06
              Mid$(SSQ$, 4 * SUCX% - 3, 4) = MKS$(CANUE)
              VESAL$(CI%) = SSQ$
            End If
          End If
       End If
       CONTI& = CONTI& + 1
       .MoveNext
      Loop
    End With
    Screen.MousePointer = 1
    '
    ' guardar en saldepo completando el vesal$ con chr$(0)'s
    For CIXO% = 1 To NUMAS%
      RESAL1& = CIXO% + 1
      If VESAL$(CIXO%) <> "" Then
        CIXS& = CIXO%
        SSQ1$ = REGLEIDO("SALDEPO", CIXS&)
        'Call REPLA(SSQ1$, VESAL$(CIXO%), 1, Len(VESAL$(CIXO%)))
        Call GRAREG("SALDEPO", VESAL$(CIXO%) + String$(256, 0), RESAL1&)
      End If
    Next CIXO%
    '
    Check2.Value = 1
    Call CIERRARCH("SALDEPO")
    '
' ********************** INICIALIZANDO ARCHIVOS DE INVENTARIOS

    GENINV04.Check3.Value = 2
    '
    Call HEADERS("INVEDEQ", "")
    Call HEADERS("INVEDEQ", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEDEQ", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEDEQ", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVEGEN", "")
    Call HEADERS("INVEGEN", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEGEN", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEGEN", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVANAL", "")
    Call HEADERS("INVANAL", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVANAL", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVANAL", "Lote / Partida: " + LOTAUT$)

120 Call BLANARCH("INVEDEP")
130 Call BLANARCH("INVEGEN")
140 Call BLANARCH("INVANAL")
   '
    Check3.Value = 1
    '
' ******************** GENERANDO INVENTARIO ANALITICO Y GLOBAL
    '
    GENINV04.Check4.Value = 2
    '
    FIN& = NUINV%
    '
    J& = 0: TTI% = 0
    URE& = ULTREG&("SALDEPO")
    '
    For I& = 2 To URE&
      Call TRACE(20, I&, FIN&)
      '
      CI% = I& - 1
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          '
          REG& = CI% + 1
          SSQ$ = REGLEIDO$("SALDEPO", REG&)
          '
          SALTOTU = 0
          For U% = 1 To NUSU%
            SALTOTU = SALTOTU + Abs(CVS(Mid$(SSQ$, 4 * U% - 3, 4)))
          Next U%
          '
          If Abs(SALTOTU) >= 0.05 Then
            TOTAL = 0
            COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%))
            '\\\OT-07-0018-RG-23/01/07///
            On Error Resume Next
            COU = COU / TICAMONE(MUN%)
            On Error GoTo 0
            '\\\OT-07-0018-RG-FIN///
            If VALORI% = 2 Then
                COU = COSCALCU(CI%, MODOCO%)
              ElseIf VALORI% = 3 Then
                COU = 0
                For KKI% = 1 To CALIS%
                  COU = PRELISTA#(APRELI$(KKI%), CI%)
                  If COU >= 0.00005 Then
                    COU = COU * TICAMONE(MONECOSTO%(MONELI%(KKI%)))
                    Exit For
                  End If
                Next KKI%
            End If
            '
            TMT% = TIMATE%(CI%)
            '
            For U% = 1 To NUSU%
              SALDE = CVS(Mid$(SSQ$, 4 * U% - 3, 4))
              If Abs(SALDE) >= 0.05 Then
                '
                TOTAL = TOTAL + SALDE
                '
                Z$ = REGBLAN$("INVANAL")
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                Call REPLA(Z$, Chr$(COSU%(U%)), 3, 1)
                Call REPLA(Z$, Chr$(TMT%), 4, 1)
                Call REPLA(Z$, MKS$(SALDE), 5, 4)
                Call REPLA(Z$, MKS$(0), 9, 4)
                Call REPLA(Z$, MKS$(0), 13, 4)
                Call REPLA(Z$, MKD$(0), 17, 8)
                Call REGAPP("INVANAL", Z$)
                REGA& = ULTREG&("INVANAL")
                '
                X$ = Left$(Z$, 8) + String$(12, 0)
                Call REPLA(X$, Chr$(TMT%), 4, 1)
                Call REPLA(X$, MKS$(COU), 9, 4)
                Call REPLA(X$, MKD$(CDbl(SALDE * COU)), 13, 8)
                Call REGAPP("INVEDEP", X$)
                '
              End If
            Next U%
            '
            If Abs(TOTAL) >= 0.05 Then
              Call REPLA(Z$, MKS$(TOTAL), 9, 4)
              Call REPLA(Z$, MKS$(COU), 13, 4)
              Call REPLA(Z$, MKD$(COU * TOTAL), 17, 8)
              Call GRAREG("INVANAL", Z$, REGA&)
              '
              Y$ = REGBLAN$("INVEGEN")
              Call REPLA(Y$, MKI$(CI%), 1, 2)
              Call REPLA(Y$, Chr$(0), 3, 1)
              Call REPLA(Y$, Chr$(TMT%), 4, 1)
              Call REPLA(Y$, MKS$(TOTAL), 5, 4)
              Call REPLA(Y$, MKS$(COU), 9, 4)
              Call REPLA(Y$, MKD$(CDbl(TOTAL * COU)), 13, 8)
              Call REGAPP("INVEGEN", Y$)
              '
            End If
          End If
        End If
      End If
199 Next I&
    '
    Call FILESORT("INVANAL", "", 1, 1, "ASC")
    Call FILESORT("INVEGEN", "", 1, 1, "ASC")
    Call FILESORT("INVEDEP", "", 1, 1, "ASC")
    Check4.Value = 1
    '
999 Call CIERRARCH("*.*")
    FiNVEN04.Text1 = FECHINVE$
    FiNVEN04.Text2 = TRIM$(CONTROL$("INVENTAR", "FEGENINV"))
    Screen.MousePointer = 1
    Hide
    Unload Me
    '
End Sub
'


Private Sub Text2_Change()
  '
  '\\\OT-07-0018-RG-23/01/07///
  If Text2 = "" Then Label2.Caption = ""
  On Error Resume Next
  If Text2 <> "" Then
     MUN% = XVALO(Text2)
     Text2.TEXT = MUN%
     Label2.Caption = Mid$(ECOARCH$("TAMONED", Chr(MUN%)), 1, 14)
  End If
  '\\\OT-07-0018-RG-FIN///
  '
End Sub

Private Sub Text2_DblClick()
   '
   '\\\OT-07-0018-RG-23/01/07///
   Call Command4_Click
   '\\\OT-07-0018-RG-FIN///
   '
End Sub


