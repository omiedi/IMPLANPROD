VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RESCUP07 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Proveedores"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9300
   ControlBox      =   0   'False
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
   ScaleHeight     =   3240
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C7D9CC&
      Caption         =   "COMPOSICIÓN DE DEUDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   105
      TabIndex        =   21
      Top             =   2310
      Width           =   4005
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Reconstrucción Automática"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   630
         TabIndex        =   22
         Top             =   400
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   8400
      ScaleHeight     =   3450
      ScaleWidth      =   945
      TabIndex        =   18
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   31
         Top             =   2760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   32
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "RESCUP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "RESCUP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   240
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C7D9CC&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   4005
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
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
         Height          =   300
         Left            =   3580
         TabIndex        =   15
         Top             =   600
         Width           =   175
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2080
         TabIndex        =   17
         Top             =   375
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1380
         TabIndex        =   14
         Top             =   375
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C7D9CC&
      Caption         =   "RANGO DE CUENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   4005
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   7
         Top             =   315
         Width           =   855
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
         Height          =   300
         Left            =   3465
         TabIndex        =   6
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   315
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         TabIndex        =   4
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Height          =   300
         Left            =   3465
         TabIndex        =   3
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   315
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   1260
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C7D9CC&
      Caption         =   "SELECCION Y ORDENAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4200
      TabIndex        =   0
      Top             =   1540
      Width           =   4005
      Begin VB.Frame Frame2 
         BackColor       =   &H00C7D9CC&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   1000
         Width           =   3375
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Numérico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   150
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Alfabético"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   2000
            TabIndex        =   29
            Top             =   150
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C7D9CC&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Cuentas de Gastos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   3
            Left            =   2000
            TabIndex        =   27
            Top             =   120
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C7D9CC&
            Caption         =   "Cuentas Corrientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   945
         Y2              =   945
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RESCUP07.frx":0454
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "RESCUP07.frx":04CE
      Top             =   1200
   End
End
Attribute VB_Name = "RESCUP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command2_Click()
    CONTROLGRABA% = 0
    If Check1.Value = 1 Then CONTROLGRABA% = 1
    Hide
    Call FRESHALL
    Call GERESCUEPRO
    CONTROLGRABA% = 0
End Sub
'

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command8_Click()
'    Call SELECHO("PROVED1")
'    NCMAX% = XVALO(VALACT1$("PROVED1"))
    NCMAX% = ListaSeleccionProveedores()
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = rasocli$((-1) * NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
'    Call SELECHO("PROVED1")
'    NCMIN% = XVALO(VALACT1$("PROVED1"))
    NCMIN% = ListaSeleccionProveedores()
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = rasocli$((-1) * NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
      Option1(0) = False
      Option1(1) = True
    End If
End Sub

Private Sub Text10_Change()
    If Abs(XVALO(Text10)) < 32767 Then
       NCMIN% = XVALO(Text10.TEXT)
       Text9.TEXT = rasocli$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
      Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    If Abs(XVALO(Text8)) < 32767 Then
       NCMAX% = XVALO(Text8.TEXT)
       Text7.TEXT = rasocli$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text10_DblClick()
    Call Command9_Click
'    Call SELECHO("PROVED1")
'    NCMIN% = XVALO(VALACT1$("PROVED1"))
'    If NCMIN% > 0 Then
'      Text10.TEXT = TRIM$(Str$(NCMIN%))
'      Text9.TEXT = rasocli$((-1) * NCMIN%)
'    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
      Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call Command8_Click
'    Call SELECHO("PROVED1")
'    NCMAX% = XVALO(VALACT1$("PROVED1"))
'    If NCMAX% > 0 Then
'      Text8.TEXT = TRIM$(Str$(NCMAX%))
'      Text7.TEXT = rasocli$((-1) * NCMAX%)
'    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub GERESCUEPRO()
    '
    Screen.MousePointer = 11
    Des% = FECHANUM(Text6): Has% = FECHANUM(Text5)
    NCMINI = XVALO(Text10): NCMAXI = XVALO(Text8)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    '
    UJ& = 0
    Call BLANARCH("!RESCUPRO")
    '
    ConsultaSql$ = "SELECT * FROM PROVED12 INNER JOIN CAJABANC WITH(NOLOCK) ON NUME_CLI=NUPROVE WHERE STAT_CLI > 0"
    '
    If NCMINI > 0 Then CondicionSql$ = " AND PROVED12.NUME_CLI >= " & CStr(NCMINI)
    If NCMAXI > 0 Then CondicionSql$ = CondicionSql$ & " AND PROVED12.NUME_CLI <= " & CStr(NCMAXI)
    If Des% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA >= '" & FETEXCOR1$(Des%) & "'"
    If Has% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA <= '" & FETEXCOR1$(Has%) & "'"
    CondicionSql$ = CondicionSql$ & " AND (CAJABANC.TipoMovim = 'FCOM' OR CAJABANC.TipoMovim = 'OPAG' OR CAJABANC.TipoMovim = 'AJUP')"
    CondicionSql$ = CondicionSql$ & " AND CAJABANC.CodiEmpr = " & CodiEmp%
    If MODORDE% = 1 Then CondicionSql$ = CondicionSql$ & " ORDER BY PROVED12.NUME_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"
    If MODORDE% = 2 Then CondicionSql$ = CondicionSql$ & " ORDER BY PROVED12.RASO_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"

    ConsultaSql$ = ConsultaSql$ & CondicionSql$
    Dim RstCuentaCorriente As ADODB.Recordset
    Set RstCuentaCorriente = READSET(ConsultaSql$)
    '
    NumeroProveedorAnterior& = 0
    With RstCuentaCorriente
        Do While Not .EOF
          NumeroProveedor& = !nume_cli
            
          If NumeroProveedor& <> NumeroProveedorAnterior& Then
            If NumeroProveedorAnterior& <> 0 Then Call REGAPP("!RESCUPRO", String$(256, "="))

            REBLA$ = REGBLAN$("RESCUCLI")
            CLIE$ = TRIM$(SAFETEXT$(!raso_cli)) + " (" + TRIM$(Str$(NumeroProveedor&)) + ")"
            DOMIE$ = TRIM$(SAFETEXT$(!domi_cli))
            LOCAE$ = TRIM$(SAFETEXT$(!Codi_Cli))
            LOCAF$ = SAFETEXT$(!loca_Cli)
            If TRIM$(LOCAF$) <> "" Then
              LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
            End If
            TELEE$ = "Telefax: " + TRIM$(SAFETEXT$(!Tele_Cli))
            '
            Call REPLA(REBLA$, MKI$(NCLIEN), 1, 2)
            Call REPLA(REBLA$, "Sres.", 97, 6)
            Call REPLA(REBLA$, CLIE$, 103, 48)
            Call REPLA(REBLA$, DOMIE$, 151, 48)
            Call REPLA(REBLA$, LOCAE$, 199, 32)
            Call REPLA(REBLA$, TELEE$, 231, 24)
            '
            YATRANSPO% = 0: ITRANSPO# = 0
            NumeroProveedorAnterior& = NumeroProveedor&
            
            CondicionSql$ = "(CAJABANC.TipoMovim = 'FCOM' OR CAJABANC.TipoMovim = 'OPAG' OR CAJABANC.TipoMovim = 'AJUP') AND FECHCONTA<'" & FETEXCOR1$(Des%) & "' AND NUPROVE=" & CStr(NumeroProveedor&)
            SaldoDebe# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CondicionSql$))
            SaldoHabe# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CondicionSql$))
            SaldoAnterior# = SaldoHabe# - SaldoDebe#
          End If
            
          If YATRANSPO% = 0 Then
              REGJ$ = REBLA$
              FEX = Des%
              Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
              Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
              Call REPLA(REGJ$, MKD$(SaldoAnterior#), 63, 8)
              Call REGAPP("!RESCUPRO", REGJ$)
              Call REGAPP("!RESCUPRO", REBLA$)
              YATRANSPO% = 1
              SaldoAcumulado# = SaldoAnterior#
          End If
          
          COCOLA$ = SAFETEXT$(!CODICOM_LAR)
          TIX$ = Left$(!CODICOM_LAR, 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
          IMPORTECOMPROBANTE# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
          '
890       ImporteCbte# = ROUND(IMPORTECOMPROBANTE#)
          SaldoAcumulado# = SaldoAcumulado# + ImporteCbte#
          FechaEmision% = CVINT%(!FECHEMISI)
          NumeroComprobante& = !NumeComp
          Moneda% = ROUND(!Mone_Emis, 0)
          TIPOCAMBIO# = ROUND(!Tipo_Cam, 5)
          If TIPOCAMBIO# < 0.05 Then TIPOCAMBIO# = 1
          If TIPOCAMBIO# >= 0.01 Then
            ImporteDolar# = ROUND#(ImporteCbte# / TIPOCAMBIO#, 2)
          End If
          '
          Call REPLA(REGJ$, FECHATEX$(FechaEmision%), 3, 8)
          Call REPLA(REGJ$, TIPOS$, 11, 18)
          Call REPLA(REGJ$, COCOLA$, 29, 18)
          If ImporteCbte# > 0 Then
            Call REPLA(REGJ$, MKD$(0), 47, 8)
            Call REPLA(REGJ$, MKD$(ImporteCbte#), 55, 8)
          ElseIf ImporteCbte# < 0 Then
            Call REPLA(REGJ$, MKD$(Abs(ImporteCbte#)), 47, 8)
            Call REPLA(REGJ$, MKD$(0), 55, 8)
          Else
            Call REPLA(REGJ$, MKD$(0), 47, 8)
            Call REPLA(REGJ$, MKD$(0), 55, 8)
          End If
          Call REPLA(REGJ$, MKD$(SaldoAcumulado#), 63, 8)
          Call REGAPP("!RESCUPRO", REGJ$)
          '
         .MoveNext
        Loop
990 End With
    RstCuentaCorriente.Close
    Set RstCuentaCorriente = Nothing
    '
    If ULTREG&("!RESCUPRO") > 0 Then
        If NumeroProveedorAnterior& <> 0 Then Call REGAPP("!RESCUPRO", String$(256, "="))
        Call TRACENTRAL("RESCUPRO", "")
        Screen.MousePointer = 1
        Call CONECRUN("*RESCPRO", "Listado Resumen de Cuenta Deudores")
    Else
        Screen.MousePointer = 1
        Call COMUNI("No se Ha Encontrado Información para las Opciones Elegidas")
    End If
    
'    If Abs(XVALO(Text10)) >= 32767 Or Abs(XVALO(Text8)) >= 32767 Then
'       Call MENSERR(24, "Rango de Cuentas no Válido")
'       Screen.MousePointer = 1
'       Exit Sub
'    End If
'    '
'    Screen.MousePointer = 11
'    '
'    DES% = FECHANUM(Text6.TEXT)
'    HAS% = FECHANUM(Text5.TEXT)
'    NCMIN% = XVALO(Text10.TEXT)
'    NCMAX% = XVALO(Text8.TEXT)
'    '
'    MODORDE% = 1
'    If Option1(2) = True Then MODORDE% = 2
'    '
'    FIN& = ULTREG&("PROVED1")
'    UJ& = 0
'    List1.Clear
'    Call BLANARCH("RESCUPRO")
'    '
'    FIN& = ULTREG&("PROVED1")
'    For U& = 1 To FIN&
'      '
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("PROVED1", U&)
'      NC% = CVI(Left$(X$, 2))
'      '
'      If NC% >= NCMIN% Then
'        If NC% <= NCMAX% Then
'          '
'          TLISTA$ = Space$(36)
'          If MODORDE% = 1 Then
'               Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
'             ElseIf MODORDE% = 2 Then
'               Call REPLA(TLISTA$, rasocli$((-1) * NC%), 1, 30)
'          End If
'          Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 31, 6)
'          List1.AddItem TLISTA$
'          '
'        End If
'      End If
'      '
'    Next U&
'    '
'    FIN& = List1.ListCount
'    '
'    For U& = 1 To FIN&
'      '
'      Call TRACE(20, U&, FIN&)
'      NC% = XVALO(Mid$(List1.LIST(U& - 1), 31, 6))
'      RECUEP07.Text1 = TRIM$(Str$(NC%))
'      If CONTROLGRABA% = 1 Then
'        Call RECUEP07.Command4_Click
'      End If
'      'Call CARESCUEPRO
'      Call FRESHALL
'      '
'      REBLA$ = REGBLAN$("RESCUPRO")
'      CLIE$ = TRIM$(RECUEP07.Text2.TEXT) + " (" + TRIM$(Str$(NC%)) + ")"
'      DOMIE$ = TRIM$(RECUEP07.Text3.TEXT)
'      LOCAE$ = TRIM$(RECUEP07.Text5.TEXT)
'      If TRIM$(RECUEP07.Text4.TEXT) <> "" Then
'          LOCAE$ = TRIM$(RECUEP07.Text4.TEXT) + " - " + LOCAE$
'      End If
'      TELEE$ = "Telefax: " + TRIM$(RECUEP07.Text7.TEXT)
'      '
'      Call REPLA(REBLA$, MKI$(NC%), 1, 2)
'      Call REPLA(REBLA$, "Sres.", 97, 6)
'      Call REPLA(REBLA$, CLIE$, 103, 48)
'      Call REPLA(REBLA$, DOMIE$, 151, 48)
'      Call REPLA(REBLA$, LOCAE$, 199, 32)
'      Call REPLA(REBLA$, TELEE$, 231, 24)
'      '
'      YATRANSPO% = 0: ITRANSPO# = 0
'      For V& = 1 To RECUEP07.List1.ListCount
'        RECUEP07.List1.ListIndex = V& - 1
'        REGI$ = RECUEP07.List1.TEXT
'        FECHA$ = Left$(REGI$, 8)
'        FF% = FECHANUM(FECHA$)
'        If FF% < DES% Then
'            ITRANSPO# = XVALO(Mid$(REGI$, 64, 13))
'        End If
'        If FF% >= DES% Then
'          If FF% <= HAS% Then
'            If YATRANSPO% = 0 Then
'              REGJ$ = REBLA$
'              FEX = DES%
'              Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
'              Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
'              Call REPLA(REGJ$, MKD$(ITRANSPO#), 63, 8)
'              Call REGAPP("RESCUPRO", REGJ$)
'              Call REGAPP("RESCUPRO", REBLA$)
'              YATRANSPO% = 1
'            End If
'            '
'            REGJ$ = REBLA$
'            DOCU$ = Mid$(REGI$, 11, 11)
'            NUDO$ = Mid$(REGI$, 22, 18)
'            DEBE# = XVALO(Mid$(REGI$, 40, 11))
'            HABER# = XVALO(Mid$(REGI$, 53, 11))
'            SALDO# = XVALO(Mid$(REGI$, 64, 13))
'            '
'            Call REPLA(REGJ$, FECHA$, 3, 8)
'            Call REPLA(REGJ$, DOCU$, 11, 18)
'            Call REPLA(REGJ$, NUDO$, 29, 18)
'            Call REPLA(REGJ$, MKD$(DEBE#), 47, 8)
'            Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
'            Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
'            Call REGAPP("RESCUPRO", REGJ$)
'            '
'          End If
'        End If
'      Next V&
'      '
'      If YATRANSPO% = 1 Then
'          Call REGAPP("RESCUPRO", String$(256, "="))
'      End If
'      '
'    Next U&
'    Screen.MousePointer = 1
'    '
'    Call CONECRUN("*RESCPRO", "Listado Resumen de Cuenta Proveedores")
'    '
End Sub
'





