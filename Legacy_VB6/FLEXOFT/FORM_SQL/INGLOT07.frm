VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INGLOT07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones Alternativas"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   21
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3675
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   3
         Left            =   960
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   4
         Top             =   1320
         Width           =   2430
         Begin VB.CommandButton AGenRep 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            Picture         =   "INGLOT07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cosultas y Listados   de Lotes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   5
            Top             =   100
            Width           =   1695
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   10
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   15
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   14
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   13
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   12
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   11
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   600
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   6
         Top             =   840
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
            Height          =   645
            Left            =   0
            Picture         =   "INGLOT07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Fracionamiento de Bobinas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fraccionamiento de Bobinas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   7
            Top             =   20
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   2430
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
            Height          =   645
            Left            =   0
            Picture         =   "INGLOT07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ingresos de Bobinas por propia Produccion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso de Bobinas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   9
            Top             =   120
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   2970
      Left            =   3960
      ScaleHeight     =   2940
      ScaleWidth      =   1230
      TabIndex        =   0
      Top             =   0
      Width           =   1260
      Begin VB.CommandButton Command10 
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
         Height          =   650
         Left            =   105
         Picture         =   "INGLOT07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   760
         Width           =   650
      End
      Begin VB.CommandButton Command11 
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
         Left            =   105
         Picture         =   "INGLOT07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INGLOT07.frx":0BB2
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INGLOT07.frx":0DE6
      TabIndex        =   20
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "INGLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command11_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call PARTEDIARIOPRO
    Command18.Enabled = True
End Sub

Sub PARTEDIARIOPRO()

    Call ABRESTOCKS
5   Call SELECHO("SELFECHA")
    FF1$ = VALACT1$("SELFECHA")
    If TRIM$(FF1$) = "" Then GoTo 99
    FF% = FECHANUM(FF1$)
    
10  Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    DEP% = XVALO(DEPO$)
    If DEP% = 0 Then GoTo 5
    
    Call SETULTREG("!PADIAPRO", 0)
    
    CONDI$ = "FECHALTA = '" & FETEXCOR1$(FF%) & "' AND DEPOUBIC = " & DEP%
    REG& = CantRegistros("LOTINGRE", CONDI$)
    If REG& > 0 Then
        SQLX$ = "SELECT IdenLote, COD_INTE, AnchoMili, CanAlta, OperProd, TurnoProd FROM LOTINGRE WITH(NOLOCK) "
        SQLX$ = SQLX$ + "WHERE " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
'        Dim PARDIATMP As ADODB.Recordset
'        Set PARDIATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Set PARDIATMP = READSET(SQLX$)
        TTXX$ = ""
        JJ& = 0
        With PARDIATMP
        Do While Not PARDIATMP.EOF
            '
            LOTTE$ = SAFETEXT(!IDENLOTE)
            CONDI$ = "IDENLOTE = '" & LOTTE$ & "' AND CODIMOVI = 'OF' AND DEPOMOVI = " & DEP%
            REGSS& = CantRegistros&("MOVISTO", CONDI$)
            If REGSS& > 0 Then
                Call REPLA(TTXX$, MKI$(XVALO(!COD_INTE)), 1, 2)
                Call REPLA(TTXX$, MKD$(XVALO(!AnchoMili)), 3, 8)
                Call REPLA(TTXX$, MKD$(XVALO(!CanAlta)), 11, 8)
                Call REPLA(TTXX$, SAFETEXT$(LOTTE$), 19, 16)
                '
                Call REPLA(TTXX$, SAFETEXT$(!OperProd), 35, 32)
                Call REPLA(TTXX$, SAFETEXT$(!TurnoProd), 67, 32)
                '
                JJ& = JJ& + 1
                Call GRAREG("!PADIAPRO", TTXX$, JJ&)
            End If
15      PARDIATMP.MoveNext
        '
        Loop
        End With
        PARDIATMP.Close
        Call SETULTREG("!PADIAPRO", JJ&)
    End If
    
20  VTB% = VENTABU%("PADIAPRO")
    If VTB% < 0 Then GoTo 10
    '
    'VERIFICA QUE NO SE MODIFIQUEN LOTES YA CONSUMIDOS
    CLM% = 0
    For U& = 1 To ULTREG&("!PADIAPRO")
        X$ = REGLEIDO$("!PADIAPRO", U&)
        KILLOO# = CVD(Mid$(X$, 11, 8))
        ANCHOO# = CVD(Mid$(X$, 3, 8))
        LOTTEE$ = TRIM$(Mid$(X$, 19, 16))
        OPERAPRO$ = TRIM$(Mid$(X$, 35, 32))
        TURNOPRO$ = TRIM$(Mid$(X$, 67, 32))

        CONDI$ = "IDENLOTE = '" & TRIM$(LOTTEE$) & "'"
        CNALTAA# = XVALO(VALOBADA("LOTINGRE", "CANALTA", CONDI$))
        ANCHOMIL# = XVALO(VALOBADA("LOTINGRE", "ANCHOMILI", CONDI$))
        TAOPERAPRO$ = VALOBADA("LOTINGRE", "OperProd", CONDI$)
        TATURNOPRO$ = VALOBADA("LOTINGRE", "TurnoProd", CONDI$)
        '
        CONSU# = ROUND(XVALO(VALOBADA("LOTINGRE", "CANCONSU", CONDI$, "NOMESS")), 1)
        If CONSU# > 0 Then
          If KILLOO# <> CNALTAA# Or ANCHOO# <> ANCHOMIL# Or OPERAPRO$ <> TRIM$(TAOPERAPRO$) Or TURNOPRO$ <> TRIM$(TATURNOPRO$) Then
            Call REPLA(X$, MKD$(CNALTAA#), 11, 8)
            Call REPLA(X$, MKD$(ANCHOMIL#), 3, 8)
            Call GRAREG("!PADIAPRO", X$, U&)
            CLM% = CLM% + 1
          End If
        End If
    Next U&
    If CLM% > 0 Then
        Call MENSERR(24, "No se Pueden Modificar Lotes ya Consumidos\Los Lotes se Ajustaron Según Valores Originales de Alta\Verifique y Vuelva a Grabar")
        GoTo 20
    End If
    '
    For U& = 1 To ULTREG&("!PADIAPRO")
        X$ = REGLEIDO$("!PADIAPRO", U&)
        CCII% = CVI(Mid$(X$, 1, 2))
        If CCII% = 0 Then
            Call MENSERR(24, "Codigo Invalido")
            GoTo 20
        End If
        If ESTRAZABLE%(CCII%) <> 1 Then
            Call MENSERR(24, "Solo Valido Articulos Trazables")
            GoTo 20
        End If
        '
        ANCHOO# = CVD(Mid$(X$, 3, 8))
        KILLOO# = CVD(Mid$(X$, 11, 8))
        LOTTEE$ = TRIM$(Mid$(X$, 19, 16))
        OPERAPRO$ = TRIM$(Mid$(X$, 35, 32))
        TURNOPRO$ = TRIM$(Mid$(X$, 67, 32))

        '
        If ANCHOO# <= 0.05 Then
            Call MENSERR(24, "No Valido Ancho(mm) Menores o Iguales a 0 (cero)")
            GoTo 20
        End If
        '
        If KILLOO# <= 0.05 Then
            Call MENSERR(24, "No Valido Kilos Menores o Iguales a 0 (cero)")
            GoTo 20
        End If
        '
        If OPERAPRO$ = "" Then
            Call MENSERR(24, "Falta Ingresar Operario")
            GoTo 20
        End If
        '
        If TURNOPRO$ = "" Then
            Call MENSERR(24, "Falta Ingresar Turno de Operario")
            GoTo 20
        End If
        '
        If LOTTEE$ = "" Then
            LOTTEE$ = NUMIDELOT$(FF%)
            CONDI$ = "IDENLOTE = '" & TRIM$(LOTTEE$) & "'"
            Call CREALOTIN(FF%, 0, 0, CCII%, LOTTEE$, "", "", KILLOO#, 0, ANCHOO#, DEP%)
            NRO& = ProNroComprobante("OF")
            DOPRILA$ = "OF-" & NRO&
            DoPriCor$ = DOPRILA$
            DOSELA$ = "PF-" & NRO&
            DOSECO$ = DOSELA$
            REFE$ = "Parte de Fabricacion"
            NUORIT% = U&
            Call MOVISTOK(FF%, CCII%, LOTTEE$, "OF", DOPRILA$, DoPriCor$, DOSELA$, DOSECO$, REFE$, NUORIT%, 0, 1, 0, DEP%, KILLOO#)
            '
         Else
            CONDI$ = "IDENLOTE = '" & TRIM$(LOTTEE$) & "'"
            Call ACSALOTE(CCII%, LOTTEE$, "NOMESS")
            '
            Call ACTUREGISTRO("LOTINGRE", "CANALTA", CONDI$, KILLOO#, REG&, "NOMESS")
            Call ACTUREGISTRO("LOTINGRE", "ANCHOMILI", CONDI$, ANCHOO#, REG&, "NOMESS")
            '
            CONDI2$ = "IDENLOTE = '" & LOTTEE$ & "' AND CODIMOVI = 'OF'"
            Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI2$, KILLOO#, REG&, "NOMESS")
        End If
        ' GUARDA FUERA DEL IF POR QUE DEBE GUARDAR EN LOS 2 CASOS
        Call ACTUREGISTRO("LOTINGRE", "OperProd", CONDI$, OPERAPRO$, REG&, "NOMESS")
        Call ACTUREGISTRO("LOTINGRE", "TurnoProd", CONDI$, TURNOPRO$, REG&, "NOMESS")

        Call ACSALOTE(CCII%, LOTTEE$, "NOMESS")
        '
89  Next U&
    Screen.MousePointer = 1
    Call COMUNI("Proceso Completado")
    
99
End Sub

Function NUMIDELOT$(FF%)
    '
    FEXX$ = FECHATEX$(FF%)
    '
    IDENLOT$ = Mid$(FEXX$, 7, 2) & Mid$(FEXX$, 4, 2)
    '
    VALOCAMP$ = "MAX (CAST(SUBSTRING(IDENLOTE, 6,4)AS INTEGER ))"
    CONDI$ = "IDENLOTE LIKE '" & IDENLOT$ & "%'"
    '
    IDENLO% = 1 + XVALO(VALOBADA("LOTINGRE", VALOCAMP$, CONDI$))
    '
    IDENLOTE$ = TRIM$(Str$(IDENLO%))
    While Len(IDENLOTE$) < 4: IDENLOTE$ = "0" + IDENLOTE$: Wend
    NUMIDELOT$ = IDENLOT$ & "-" & IDENLOTE$
    '
End Function

Private Sub Command3_Click()
    '
    FRABOBI09.Show 1
    '
End Sub


Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

