VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPVARFA07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complementos de Facturación"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   4740
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
   ScaleHeight     =   4050
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3150
      Top             =   1890
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   3840
      ScaleHeight     =   5970
      ScaleWidth      =   2190
      TabIndex        =   15
      Top             =   0
      Width           =   2220
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   34
         Top             =   3720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Correl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   28
         ToolTipText     =   "Recuperación de Comprobantes Electrónicos"
         Top             =   3240
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   27
         ToolTipText     =   "Test"
         Top             =   2700
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPVARFA07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Acceso al Setup de Ventas"
         Top             =   1520
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Token"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   24
         ToolTipText     =   "Genera Archivo TOKEN.TXT para Factura Electrónica."
         Top             =   2210
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPVARFA07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPVARFA07.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   820
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "LISTADOS DE CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      TabIndex        =   8
      Top             =   2100
      Width           =   3585
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   3
         Left            =   600
         ScaleHeight     =   615
         ScaleWidth      =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   2670
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
            Height          =   630
            Left            =   2090
            Picture         =   "OPVARFA07.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Generar  Etiquetas - Re-Imprimir"
            Top             =   0
            Width           =   560
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
            Height          =   630
            Left            =   0
            Picture         =   "OPVARFA07.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Control de Pre- cios de Venta"
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
            Left            =   720
            TabIndex        =   23
            Top             =   105
            Width           =   1395
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   1
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2640
         TabIndex        =   12
         Top             =   420
         Width           =   2670
         Begin VB.CommandButton Command10 
            Height          =   630
            Left            =   2090
            Picture         =   "OPVARFA07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Inicializa campo de simbolo de moneda"
            Top             =   0
            Width           =   560
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
            Height          =   630
            Left            =   0
            Picture         =   "OPVARFA07.frx":0FF4
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Control D.C"
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
            TabIndex        =   14
            Top             =   60
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CORREGIR Y ANULAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   0
         Left            =   600
         ScaleHeight     =   615
         ScaleWidth      =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   2670
         Begin VB.CommandButton Command11 
            Height          =   630
            Left            =   2100
            Picture         =   "OPVARFA07.frx":12FE
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Corregir Número de Documento"
            Top             =   0
            Width           =   560
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
            Height          =   630
            Left            =   0
            Picture         =   "OPVARFA07.frx":1448
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Corregir Fecha de Emisión"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Corregir Fecha / Nro. Documento"
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
            Left            =   650
            TabIndex        =   20
            Top             =   105
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   420
         ScaleHeight     =   615
         ScaleWidth      =   2640
         TabIndex        =   9
         Top             =   420
         Width           =   2670
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
            Height          =   630
            Left            =   2100
            Picture         =   "OPVARFA07.frx":1592
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Eliminación de Comprobantes Anulados sin CAE"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command45 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "OPVARFA07.frx":16DC
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Anulación de Comprobantes de Facturación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Comprobantes"
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
            Left            =   735
            TabIndex        =   10
            Top             =   60
            Width           =   1425
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
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
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
            TabIndex        =   7
            Top             =   1890
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
            TabIndex        =   6
            Top             =   1890
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
            TabIndex        =   5
            Top             =   1155
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
            TabIndex        =   4
            Top             =   1155
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
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
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
            TabIndex        =   2
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPVARFA07.frx":19E6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPVARFA07.frx":1C1A
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu CuContaPorCodigo 
         Caption         =   "Cuenta Contable por Código"
      End
      Begin VB.Menu mnuCtlCorr 
         Caption         =   "Control de Correlatividad "
      End
   End
End
Attribute VB_Name = "OPVARFA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim MATRIZ$() 'DECLARACION DE VECTOR DE RECUPREACION DE FACTURAS PARA MEDITEA
 Dim INDIC&
 Dim PVTA%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    '
    Call ABRECONEXION
    Call CREATABLA_CUECORRI
    '
    CONDI$ = "TIPOMOVIM = 'FVEN' AND (SIMMONE = '' OR SIMMONE IS NULL)"
    FIN& = CantRegistros("CAJABANC", CONDI$, "NOMESS")
    If FIN& > 0 Then
        '
        Call COMUNI("Deberan Realizarse Cambios Sobre la Base de Datos\Los Cuales Podran Demorar Algunos Minutos.\Favor de Esperar Mensaje: 'Proceso Terminado'")
        '
        Screen.MousePointer = 11
        SQLX$ = "SELECT DISTINCT(CODICOM_LAR), NUCLIEN, MONE_EMIS FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
        Set RSCABAN = READSET(SQLX$)
        i& = 0
        With RSCABAN
            Do While Not .EOF
                i& = i& + 1
                Call TRACE(24, i&, FIN&)
                NUFACTU$ = SAFETEXT$(!codicom_lar)
                MONE% = XVALO(!Mone_Emis)
                NCLI% = XVALO(!NUCLIEN)
                SIMBO$ = SIMBOLPES$(MONE%)
                CONDI2$ = "CODICOM_LAR='" & NUFACTU$ & "' AND NUCLIEN=" & NCLI%
                Call ACTUREGISTRO("CAJABANC", "SimMone", CONDI2$, SIMBO$, REG&, "NOMESS")
            .MoveNext
            Loop
        End With
        RSCABAN.Close
        Set RSCABAN = Nothing
        '
        Screen.MousePointer = 1
        Call COMUNI("Proceso Terminado")

    End If
    
    Command10.Enabled = True
    
End Sub

Sub Command11_Click()
    
    Command11.Enabled = False
    If DERACCE%("CORRFECH", "Corregir Número de Documento") < 2 Then
      Command11.Enabled = True
      Exit Sub
    End If
    '
    Call COMUNI("ATENCIÓN: No Podrá Modificarse el Número\de Comprobantes Correspondientes a Períodos Cerrados.\Para Estos Casos Deberá Re-Abrirse el Período Mensual.")
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
    FECHALU = FETEXCOR1$(CVINT(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then
       Command11.Enabled = True
       Exit Sub
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & "' "
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC "
    '
'   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CABATEMP As New ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    jj& = 0
    With CABATEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        On Error GoTo 0
        Do While Not .EOF
           NCL0 = XVALO(!NUCLIEN)
           IMPO# = XVALO(!ValAbsComp)
           FEMISI$ = FECHATEX$(CVINT(!FECHEMISI))
           NPX& = XVALO(Mid$(!codicom_lar, 9, 8))
           If Left$(!codicom_lar, 2) = "CC" Then
              NPX& = XVALO(Mid$(!codicom_lar, 4))
           End If
           Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + AJUSTD$(FEMISI$, 14)
           jj& = jj& + 1
           Call GRAREG("!SELCUERE", Y$, jj&)
        .MoveNext
        Loop
'      End If
    End With
    Call SETULTREG("!SELCUERE", jj&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Corregirse.")
       GoTo 10
    End If
    '
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 9, 8))
    FEMISI$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    '
15  NuevoNumFact$ = InputBox$("Nuevo Número de Documento", "Comprobante " + DOCULA$ + " del " + FEMISI$, NUOPANU&)
    If TRIM$(NuevoNumFact$) = "" Then       ' transaccion cancelada
       Command11.Enabled = True
       Exit Sub
    End If
    '
    'VALIDACION DEL NUMERO DE BOLETA INGRESADO
    '
    NUMECOMPR& = XVALO(NuevoNumFact$)
    If IsNumeric(NuevoNumFact$) = False Then
       Call COMUNI("Número de Factura no Procesable.")
       GoTo 15
    End If
    If NUMECOMPR& = NUOPANU& Then
       Call COMUNI("Número de Factura no Modificado.")
       GoTo 15
    End If
    '
    'FORMATEO DEL NUMER0 INGRESADO
    nufact$ = NuevoNumFact$
    While Len(nufact$) < 8: nufact$ = "0" + nufact$: Wend
    NUMEFACT$ = DOCULA$
    Mid(DOCULA$, 9, 8) = nufact$
    NUMCOMP$ = DOCULA$
    NUMCOMPCOR$ = Left$(DOCULA$, 3) + Mid$(DOCULA$, 11)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    CONDI$ = "CodiCom_Lar = '" & NUMCOMP$ & "'"
    If CantRegistros("CAJABANC", CONDI$) > 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "El Número Propuesto ya fue Utilizado.\Transacción Cancelada")
        GoTo 10
    End If
    '
    If COMALTER%("¿Realmente Quiere MODIFICAR el Número \del Documento: " & NUMEFACT$ & " ,\Por el Número de Documento '" + NUMCOMP$ + "'?\\No, Conservar " + NUMEFACT$ + "\Si, Cambiar al " + NUMCOMP$) <> 2 Then GoTo 10
    '
    Screen.MousePointer = 11
    '
    'GRABA EN SADEUPEN
    SQLX$ = "UPDATE SADEUPEN "
    SQLX$ = SQLX$ + "SET Codicom_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " Codicom_Cor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE Codicom_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NCLI
    FLEXCONN.Execute (SQLX$)
    '
    'GRABA EN MOVISTO
    SQLX$ = "UPDATE MOVISTO "
    SQLX$ = SQLX$ + "SET DosecLar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " DosecCor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE DosecLar = '" & NUMEFACT$ & "' ;"
    FLEXCONN.Execute (SQLX$)
    '
    'GRABA EN APLICOBRA
    SQLX$ = "UPDATE APLICOBRA "
    SQLX$ = SQLX$ + "SET CoDoCor_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " CoDoCor_Cor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE CoDoCor_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NCLI
    FLEXCONN.Execute (SQLX$)
    '
    'GRABA EN CAJABANC
    SQLX$ = "UPDATE CAJABANC "
    SQLX$ = SQLX$ + "SET Codicom_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " Codicom_Cor = '" & NUMCOMPCOR$ & "', "
    SQLX$ = SQLX$ + " NumeComp = " & NUMECOMPR& & " "
    SQLX$ = SQLX$ + " WHERE Codicom_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NCLI
    FLEXCONN.Execute (SQLX$)
    '
    'GRABAR EN TANUMSE
    Call CREATABLA_EQUIPOS
    CONDI$ = "NUMFACTURA = '" & NUMEFACT$ & "'"
    '
    SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + CONDI$
    Dim TANUTMP As New ADODB.Recordset
    Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With TANUTMP
     Do While Not .EOF
       NUSERRR$ = SAFETEXT$(!NUMEROSERIE)
       CUERPO$ = " Cambio Número de Factura " & NUMEFACT$ & vbCrLf
       CUERPO$ = CUERPO$ + " Por Número: " & NUMCOMP$
       Call ANOTANUSE(NUSERRR$, CUERPO$)
       .MoveNext
     Loop
    End With
    TANUTMP.Close
    Set TANUTMP = Nothing
    '
    Call ACTUREGISTRO("TANUMSE", "NUMFACTURA", CONDI$, NUMCOMP$, REGAF&, "NOMESS")
    '
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
       RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
       XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
       If TRIM$(Mid$(XXC$, 13, 18)) = TRIM$(NUMEFACT$) Then
         Call REPLA(XXC$, NUMCOMP$, 13, 18)
         Call GRAREG("CARBIMON", XXC$, RECABIN&)
       End If
       If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(NUMEFACT$) Then
         Call REPLA(XXC$, NUMCOMP$, 45, 18)
         Call GRAREG("CARBIMON", XXC$, RECABIN&)
       End If
    Next KUL&
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    Call COMUNI("Comprobante '" + TRIM$(NUMEFACT$) + "'\Cambiado por Número de Documento '" + NUMCOMP$ + "'.")
    GoTo 10
    '
    '//OT-07-0479//
End Sub

Private Sub Command12_Click()
  
     Command12.Enabled = False
     
     OPX% = COMALTER%("Opciones:\\Generar Etiquetas\Re-Imprimir Etiquetas")
     If OPX% < 1 Or OPX% > 2 Then GoTo 99
     
10   Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
     If PERIO$ = "" Then GoTo 99
     '
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Desde = FETEXCOR1$(DESDEX%)
     Hasta = FETEXCOR1$(HASTAX%)
     '
     Call ABRESTOCKS
     'LISTA DE SELECCION DE REMITOS
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT/Selección de Remitos")
     '
     NUREMITO$ = VALACT1$("SELREMIT")
     If NUREMITO$ = "" Then GoTo 10
     FE% = FECHANUM(Left$(VALACT2$("SELREMIT"), 8))

     If OPX% = 1 Then
       Call ETIQUETADEREMITO(NUREMITO$, FE%) 'Generar Etiquetas
     ElseIf OPX% = 2 Then
       Call REIMPRIME_ETIQUETAS(NUREMITO$)   '\Re-Imprimir Etiquetas
     End If
     
     GoTo 15
99  Command12.Enabled = True

End Sub
Sub ETIQUETADEREMITO(NUREMITO$, FE%)
    NUREMI$ = NUREMITO$
    FEMOVI = FETEXCOR1$(FE%)
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
        CMO$ = "RT"
    End If
    CANTI# = 0
    ACUM_CANTI# = 0
    'RECORRE EL REMITO Y ACUMULA LAS CANTIDAD DE BULTOS SEGUN CODIGO
    SQLX$ = "SELECT COD_INTE,SUM(CANTSALID)AS CANTI_SALIDA,FECHMOV,NUME_CLIE, NUPEDCLI FROM MOVISTO with(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE DOPRILAR = '" & NUREMI$ & "' "
    SQLX$ = SQLX$ + " AND FECHMOV = '" & FEMOVI & "' "
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'"
    SQLX$ = SQLX$ + " AND substring(DOPRILAR,1,2) = 'RT' "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,FECHMOV,NUME_CLIE,NUPEDCLI"
    
    Set Rst = READSET(SQLX$)
    '
    With Rst
        Do While Not .EOF
          CIXI% = XVALO(!cod_inte)
          CANTI# = XVALO(!CANTI_SALIDA)
          If VUELTA% < 1 Then
             NP& = XVALO(!NUPEDCLI)
             OCOM$ = TRIM$(VALOBADA("PEDENCA", "NROOCOM", "NROPED = " & NP&, "NOMESS"))
             NC% = XVALO(!NUME_CLIE)
             VUELTA% = 1
          End If
          '
          CANTICAJAS& = CANBULTOS(CIXI%, CANTI#, "VENTA") 'CANTIDAD DE BULTOS SEGUN CODIGO
          ACUM_CANTI# = ACUM_CANTI# + CANTICAJAS          'ACUMULA
          .MoveNext
        Loop
    End With
    '
    'GENERA LA CANTIDAD DE ETIQUETAS SEGUN BULTOS EN TOTAL DE TODOS LOS ARTICULOS
    Call PRIETIQUETA_REMITO(NUREMI$, NC%, ACUM_CANTI#, FE%, OCOM$)
    '
    Rst.Close
    Set Rst = Nothing
End Sub
Sub REIMPRIME_ETIQUETAS(NUREMITO$)
    'PRESENTA LISTA DE ETIQUETAS IMPRESAS DEL PDOC_TBL SEGUN REMITO LAS MISMAS TIENEN EL NUMERO DE ORDEN CONCATENADO AL N.REMITO
    
    NUREMI$ = NUREMITO$
    SQLX$ = "LEFT(NUMECOMPLE,18) = '" & NUREMI$ & "'"
    SQLX$ = SQLX$ + " AND TIPODOCU = 'ETIQUETAS DE REMITO'"
    Call CARGALISEL("INETIREM", "PDOC_TBL", "NUMECOMPLE/A24;REFERENCIA/A48", SQLX$, "DISTINCT")
10  Call SELECHO("INETIREM/SELECCIÓN DE ETIQUETAS")
    NUREMICOMPLE$ = TRIM$(VALACT1$("INETIREM"))
    If NUREMICOMPLE$ = "" Then Exit Sub
    '
    Call MUESTRADOC(DOCUNU&, "", "ETIQUETAS DE REMITO", "", NUREMICOMPLE$)
    GoTo 10
End Sub
Sub PRIETIQUETA_REMITO(NUREMI$, NC%, CANTI#, FE%, OCOM$)

    If NC% > 0 And CANTI > 0 And NUREMI$ <> "" Then
       
       IMPR$ = TRIM$(Control$(IDENTER$, "PRIETIRE"))
       If IMPR$ = "" Then IMPR$ = Control$("*", "PRIETIRE")
       Call SETPRINPORT(IMPR$)
       PRINTERPORT$ = ""
       
       RAZO$ = rasocli$(NC%)
       fech$ = FECHATEX$(FE%)
       FORIPRE$ = TRIM$(Control$("", "FORETIRE")) 'ETIQUETA CONFIGURADA EN SETUP DE STOCKS
       CANTICAJAS = CANTI#
       NOCONFIRM% = 1 ' PARA QUE NO PRESENTE MENSAJE DE CONFIRMACION
       'INICIALIZO VARIABLES PUBLICAS
       CACOPDOC% = 0
       CACOPI% = 0
       CAITAB1% = 0
       '
       CACOLTAB1% = 0
       CACOLTAB2% = 0
       DESTIDOCNUE = NC%
       '
       INICIOCAJA& = XVALO(InputBox("Iniciar en:", "Inicio de Impresion de Etiquetas", "1"))
       If INICIOCAJA& < 1 Then INICIOCAJA& = 1
       '
       If INICIOCAJA& > CANTICAJAS Then
            Call COMUNI("La Cantidad Ingresada Supera el Maximo de Cajas (" & TRIM$(Str$(CANTICAJAS)) & ")")
            GoTo 90
       End If
       
       For i& = INICIOCAJA& To CANTICAJAS
           'GENERA LA IMPRESION DE TODAS LAS ETIQUETAS CONCATENANDO UN NUMERO DE ORDEN 000N
           NUMEDOC$ = NUREMI$ & "-" & Right$("0000" + TRIM$(SAFETEXT$(i&)), 4)
           TIPODOC$ = "ETIQUETAS DE REMITO"
           CODPARAM$(1) = "NRO-REMI": DOCPARAM$(1) = SAFETEXT$(XVALO(Mid$(NUREMI$, 11)))
           CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = OCOM$
           CODPARAM$(3) = "CAN-CONJ": DOCPARAM$(3) = SAFETEXT$(i&)
           CODPARAM$(4) = "CAN-MPRI": DOCPARAM$(4) = SAFETEXT$(CANTICAJAS)
'           CODPARAM$(5) = "DOMI-CLI": DOCPARAM$(5) = DOMICLI$(NC%)
           CODPARAM$(5) = "FECH-EMI": DOCPARAM$(5) = fech$
'           CODPARAM$(7) = "RASO-CLI": DOCPARAM$(7) = RASOCLI$(NC%)
           CODPARAM$(6) = "ESP-MAT": DOCPARAM$(6) = LOCACLI$(NC%)
           CAPARDOC = 6
           '
           Call PRINTDOC("@" + FORIPRE$)
           '
       Next i&
90     NOCONFIRM% = 0 ' PARA QUE VUELVA A CONFIRMAR.
       '
       Call SETPRINPORT("RESTORE")
       '
    End If
End Sub
Sub Command2_Click()
    '
    Command2.Enabled = False
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          Call MENSERR(24, "Solo Disponible en Sistemas Bi- O Multi-Monetarios")
          GoTo 99
       End If
    End If
    '
10  FEX = HOY(HO$)
    Has% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASFECHCLI", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    CLISEL% = CVI(Mid$(OBX$, 5, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CARBIMON", U&)
      FEMI% = CVI(Mid$(X$, 1, 2))
      If FEMI% >= Des% Then
        If FEMI% <= Has% Then
          If CLISEL% < 1 Or CVI(Mid$(X$, 3, 2)) = CLISEL% Then
            '
            ' ENJUAGUE ASEMBLI-GRAMY
            MONEDAPLI% = CVI(Mid$(X$, 63, 2))
            If MONEDAPLI% > ULTREG&("TAMONED") Then
              MONEDAPLI% = MONEDAPLI% Mod 6
              Call REPLA(X$, MKI$(MONEDAPLI%), 63, 2)
              Call GRAREG("CARBIMON", X$, U&)
              If TRIM$(Mid$(X$, 45, 18)) = TRIM$(Mid$(X$, 13, 18)) Then
                Call REPLA(X$, MKI$(MONEDAPLI%), 31, 2)
                Call GRAREG("CARBIMON", X$, U&)
              End If
            End If
            ' FIN ENJUAGUE ASEMBLI-GRAMY
            '
            J& = J& + 1
            Call GRAREG("CARBIMOX", X$, J&)
          End If
        End If
      End If
    Next U&
    Call SETULTREG("CARBIMOX", J&)
    Call CIERRARCH("CARBIMOX")
    '
    Call FILESORT("CARBIMOX", "", 1, 6, "ASC")
    Call CONECRUN("*LISBIMON", "")
    '
99  Screen.MousePointer = 1
    Command2.Enabled = True
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    If DERACCE%("CORRFECH", "Corregir Fecha Emisión Facturas") < 2 Then
      Command3.Enabled = True
      Exit Sub
    End If
    '
    Call COMUNI("ATENCIÓN: No Podrá Modificarse la Fecha de Emisión\de Comprobantes Correspondientes a Períodos Cerrados.\   \Para Estos Casos Deberá Re-Abrirse el Período Mensual.")
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    'JANDY SQL ESTO FUE CAMBIADO SI NO HAY FECHA, LA LIMITE '01/01/1980'
    FECHALU = FETEXCOR1$(FECHALIM%)
    '
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then
       Command3.Enabled = True
       Exit Sub
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & "' "  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    jj& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCL0 = !NUCLIEN
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then
          '  IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
          '  If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            FEMISI$ = FECHATEX$(CVINT(!FECHEMISI))
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            If Left$(!codicom_lar, 2) = "CC" Then
              NPX& = XVALO(Mid$(!codicom_lar, 4))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + AJUSTD$(FEMISI$, 14)
            jj& = jj& + 1
            Call GRAREG("!SELCUERE", Y$, jj&)
          'End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
 
    Call SETULTREG("!SELCUERE", jj&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Corregirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    FEMISI$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    '
15  NUEFE$ = InputBox$("Nueva Fecha de Emisión", "Comprobante " + DOCULA$ + " del " + FEMISI$, FEMISI$)
    If NUEFE$ = "" Then GoTo 10
    If NUEFE$ = FEMISI$ Then
       Call COMUNI("Fecha no Modificada.")
       GoTo 10
    End If
    If FECHANUM(NUEFE$) < 1 Then
       Call MENSERR(24, "Formato de Fecha '" + NUEFE$ + "' no Válido")
       GoTo 15
    End If
    If FECHANUM(NUEFE$) < FECHALIM% Then
       Call MENSERR(24, "Fecha Ingresada no Válida.\Ese Período Mensual Está Cerrado.")
       GoTo 15
    End If
    '
    If COMALTER%("¿Realmente Quiere MODIFICAR la Fecha de Emisión\del Documento: " & DOCULA$ & " ,\del '" + FEMISI$ + "' al '" + NUEFE$ + "' ?\\No, Conservar " + FEMISI$ + "\Si, Cambiar al " + NUEFE$) <> 2 Then GoTo 10
    OKX% = (-1)
    '
    NUEVAFECHA = CVDAT(FECHANUM(NUEFE$))
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    
    'Jandy SQL Esta rutina fue cambiada
    ' CAJABANC
    strSQL = " Update CajaBanc set FechEmisi= '" & NUEVAFECHA & "'"
    strSQL = strSQL & ", FechConta= ' " & NUEVAFECHA & "'"
    strSQL = strSQL & " Where CodiCom_Lar = '" & DOCULA$ & "' "
    strSQL = strSQL & " And NuCLIEN = " & NCLI
    strSQL = strSQL & " And FechEmisi > '" & FECHALU & "' "
     ' SADEUPEN
    strSQL = strSQL & " Update SaDeuPen set FechEmisi= '" & NUEVAFECHA & "'"
    strSQL = strSQL & " Where CodiCom_Lar = '" & DOCULA$ & "' "
    strSQL = strSQL & " And NuCLIEN = " & NCLI
     ' MOVISTO
    strSQL = strSQL & " Update Movisto set FechMov= '" & NUEVAFECHA & "'"
    strSQL = strSQL & " Where DoSecLar = '" & DOCULA$ & "' "
    strSQL = strSQL & " And Nume_Clie = " & NCLI
     
     FLEXCONN.Execute (FILTSQL$(strSQL))
    If Err > 1 Then
     Call MENSERR(24, "Transacción Imposible. Consulte !")
     Exit Sub
    End If
'
'    With CajaBanc
'       SQLX$ = "(TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
'       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
'       SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
'       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
'       .FindFirst SQLX$
'       If .NoMatch = True Then
'          Call MENSERR(24, "Transacción Imposible. Consulte !")
'          GoTo 10
'       End If
'       REUTI% = 0
'       '
'       SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' "
'       SQLX$ = SQLX$ + "AND (NuClien = " & NCLI & " OR NuClien = 0) "
'       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
'19     .FindFirst SQLX$
'20     If .NoMatch = False Then
'          .Edit
'          !FechEmisi = NUEVAFECHA
'          !FechConta = NUEVAFECHA
'          .Update
'          .FindNext SQLX$
'          GoTo 20
'       End If
'    End With
'    '
'    With SaDeuPen
'22    .FindFirst " CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLI & " "
'23     If .NoMatch = False Then
'        .Edit
'        !FechEmisi = NUEVAFECHA
'        .Update
'        .FindNext SQLX$
'        GoTo 23
'      End If
'    End With
'    'ACA BUSCO EL REGISTRO EN MOVISTO Y PONGO EN BLANCO EL NºFACT
'    With Movisto
'       .FindFirst " DoSecLar = '" & DOCULA$ & "' AND Nume_Clie = " & NCLI & " "
'24     If .NoMatch = False Then
'        .Edit
'        !fechMov = NUEVAFECHA
'        .Update
'        .FindNext "DoSecLar = '" & DOCULA$ & "' AND Nume_Clie= " & NCLI & ""
'        GoTo 24
'      End If
'    End With
    
    
    '
'    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
'    For KUL& = Len(RFG$) - 3 To 1 Step -4
'      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'      If TRIM$(Mid$(XXC$, 13, 18)) = TRIM$(DOCULA$) Then
'        Call REPLA(XXC$, MKI$(CVINT(NUEVAFECHA)), 1, 2)
'        Call GRAREG("CARBIMON", XXC$, RECABIN&)
'      End If
'      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
'        Call REPLA(XXC$, MKI$(CVINT(NUEVAFECHA)), 33, 2)
'        Call GRAREG("CARBIMON", XXC$, RECABIN&)
'      End If
'    Next KUL&
    '
    Screen.MousePointer = 1
    Call COMUNI("Comprobante '" + TRIM$(DOCULA$) + "'\Fecha Emisión Cambiada a '" + NUEFE$ + "'.")
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    REBLA$ = REGBLAN$("LICONPRE")
    jj& = 0
    '
    DESDA = FECHATEX(Des%)
    HASDA = FECHATEX(Has%)
    
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDA & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDA & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With AVANCETMP
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
          FC% = CVINT(!FECHEMISI)
          FEMOVI$ = FETEXCOR1$(FC%)
          NUCO$ = SAFETEXT$(!codicom_lar)
          NCLIE% = XVALO(!NUCLIEN)
          IMPO# = XVALO(!IHABECOMP) - XVALO(!Idebecomp)
          '
          IREMIT# = 0: IPEDID# = 0: STRIPED$ = ""
          SQLX$ = "SELECT * FROM MOVISTO WHERE DOSECLAR = '" & NUCO$ & "' "
          SQLX$ = SQLX$ + " AND FECHMOV = '" & FEMOVI$ & "' "
          Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
          With MOVITMP
            If Not (.BOF And .EOF) Then
              Do While Not .EOF
                CIXI% = XVALO(!cod_inte)
                VUNI# = XVALO(!VALOUNIT)
                MONEXI% = XVALO(!MoneVal)
                TIPOCA = XVALO(!TIPOCAM)
                If MONEXI% > 1 Then
                   VUNI# = VUNI# * TIPOCA
                End If
                CANTU# = XVALO(!CANTSALID) - XVALO(!CANTINGRE)
                NUPEDI& = XVALO(!NUPEDCLI)
                NORPED% = XVALO(!NUMEORDPED)
                IREMIT# = IREMIT# + CANTU# * VUNI#
                '
                STRP$ = TRIM$(Str$(NUPEDI&)) + "-"
                  If InStr(STRIPED$, STRP$) < 1 Then
                  STRIPED$ = STRIPED$ + STRP$
                End If
                '
                CONDI$ = "NROPED = " & NUPEDI& & " AND NuOrItem = " & NORPED%
                PREPE# = XVALO(VALOBADA("PEDDETA", "PreUnid", CONDI$))
                If PREPE# < 0.00005 Then
                  CONDI$ = "NROPED = " & NUPEDI& & " AND CodiInt = " & CIXI%
                  PREPE# = XVALO(VALOBADA("PEDDETA", "PreUnid", CONDI$))
                End If
                IPEDID# = IPEDID# + CANTU# * PREPE#
              
              .MoveNext
              Loop
            End If
          End With
          '
          If Right$(STRIPED$, 1) = "-" Then STRIPED$ = Left$(STRIPED$, Len(STRIPED$) - 1)
          REGG$ = REBLA$
          Call REPLA(REGG$, MKI$(FC%), 1, 2)
          Call REPLA(REGG$, NUCO$, 3, 24)
          Call REPLA(REGG$, MKI$(NCLIE%), 27, 2)
          Call REPLA(REGG$, MKD$(IMPO#), 65, 8)
          Call REPLA(REGG$, MKD$(IREMIT#), 73, 8)
          Call REPLA(REGG$, STRIPED$, 81, 32)
          Call REPLA(REGG$, MKD$(IPEDID#), 113, 8)
            MAVERI$ = "OK"
            If Abs(IMPO# - IREMIT#) >= 0.005 Then
                MAVERI$ = "Error"
              ElseIf Abs(IMPO# - IPEDID#) >= 0.005 Then
                If STRIPED$ <> "0" Then
                  MAVERI$ = "Pr.Pedid"
                End If
            End If
          Call REPLA(REGG$, MAVERI$, 121, 8)
          jj& = jj& + 1
          Call GRAREG("LICONPRE", REGG$, jj&)
          '
        .MoveNext
        Loop
        '
        Call SETULTREG("LICONPRE", jj&)
        Call CIERRARCH("LICONPRE")
        Call CONECRUN("*LICONPRE", "Listado Control de Precios Facturacion")
        '
      End If
    End With
    '
99  Screen.MousePointer = 1
    Command4.Enabled = True
End Sub

Sub Command45_Click()
   '
   Command45.Enabled = False

   If DERACCE%("ANULAFAC", "Anulación de Comprobantes de Facturación") = 2 Then
     TTXX$ = "Pueden Anularse Comprobantes Emitidos\o Formularios Inutilizados"
     MODO% = COMALTER%(TTXX$ + "\\Comprobantes Emitidos\Formularios Arruinados\Cancelar")
     '
     If MODO% = 1 Then Call ANUFAC
     '
     If MODO% = 2 Then
         '
         NUMECORR% = 1
         QTRX$ = Control$("NUMEFACT", "NUMECORR")
         If QTRX$ = "INDEPEN" Then
              NUMECORR% = (-1)
           ElseIf QTRX$ = "UNIFICA" Then
              NUMECORR% = (-2)
         End If
         'Si la numeracion no es correlativa no se pueden
         'anular comprobantes.
         If NUMECORR% <= 0 Then
             Call COMUNI("Solo Disponible para Numeración Correlativa.")
             Command20.Enabled = True
             Exit Sub
           Else
             Call ANUFORM
             Call COMUNI("Re-Inicie Ahora la Aplicacion")
             Call FINAL("")
         End If
        '
     End If
   End If
   Command45.Enabled = True
   '
End Sub

Sub ANUFAC()
'ANULACION DE FACTURAS YA EMITIDAS
    '
    If DERACCE%("ANUFAEM", "Anulación de Facturas Emitidas") < 2 Then
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    'FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    'JANDY SQL ESTO FUE CAMBIADO SI NO HAY FECHA, LA LIMITE '01/01/1980'
    FECHALU = FETEXCOR1$(FECHALIM%)
    
    '
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & " '"  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FechEmisi ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'jandy sql
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    '
    jj& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLI1 = !NUPROVE
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            If Left$(!codicom_lar, 2) = "CC" Then
              NPX& = XVALO(Mid$(!codicom_lar, 4))
            End If
            jj& = jj& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(jj&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, jj&)
          'End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", jj&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
15  Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    IMPDOCU$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    If IMPDOCU$ = "" Then IMPDOCU$ = " 0.00"
    '
    If COMALTER%("¿Realmente Quiere ANULAR el Documento: " & DOCULA$ & " ,\con Importe: $" & IMPDOCU$ & "  de la Cuenta " & rasocli(NCLI) & " ?\\No, Cancelar\Si, Continuar") <> 2 Then GoTo 10
    '
    'EN EL CASO QUE SE QUIERA ELIMINAR EL REGISTRO VALIDA SI TIENE CAE
    CONDI$ = " (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    CONDI$ = CONDI$ + " AND CodiCom_Lar = '" & DOCULA$ & "' "
    CONDI$ = CONDI$ + " AND NuCLIEN = " & NCLI & " "
    CONDI$ = CONDI$ + " AND FechEmisi > '" & FECHALU & "' "
    SUCAE$ = TRIM$(VALOBADA("CAJABANC", "CAE", CONDI$, "NOMESS"))
    CbteOriginalExportaSimple$ = TRIM$(VALOBADA("CAJABANC", "codicom_lar_exsimple", CONDI$, "NOMESS"))
    If SUCAE$ <> "" And CbteOriginalExportaSimple$ = "" Then
       TXX1$ = "Imposible Anular el Comprobante: " & DOCULA$ & " con Cae: " & SUCAE$
       TXX1$ = TXX1$ + ".\Este Documento Electrónico Ya Se Encuentra Aprobado y Registrado en la AFIP."
       Call COMUNI(TXX1$)
       Screen.MousePointer = 1
       GoTo 15
    End If
    '
    
    If EXISTE_TABLA%("factuanti") > 0 Then
       CONDIFACTUANTI$ = "docu_orig = '" & DOCULA$ & "'  and (docu_apli <>  '' or docu_apli  <> null) AND NUCLIEN = " & NCLI
       If CantRegistros&("factuanti", CONDIFACTUANTI$, "NOMESS") > 0 Then
         Call MENSERR(24, "Imposible Anular!!!\Factura Elegida Es una Factura Anticipo\La Misma está Aplicada a Otras Facturas")
         GoTo 10
       End If
    End If

    OKX% = (-1)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    '
    SQLX$ = "Select * from CajaBanc where "
    SQLX$ = SQLX$ + CONDI$
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
    If (.EOF And .BOF) Then
        Call MENSERR(24, "Anulación Imposible. Consulte !")
        GoTo 10
    End If
    '
    REUTI% = 0
    If !OpciMovim = "CI" Then REUTI% = 1
    If REUTI% = 0 Then
       If DERACCE%("NULDEFAC/NOMESS", "Anulacion Reutilizable Facturas") = 2 Then
          RESERNUM% = COMALTER%("Desea Dejar Reservado\este Número de Comprobante ?\\Reservar Número\Re-Utilizarlo")
          If RESERNUM% < 1 Or RESERNUM% > 2 Then
             Call COMUNI("Transaccion Cancelada")
             GoTo 10
          End If
          REUTI% = RESERNUM% - 1
       End If
    End If
    '
    If REUTI% = 0 Then
         Call COMUNI("La Transacción RESERVARÁ el Número de Comprobante.")
       Else
         Call COMUNI("La Transacción PERMITIRÁ REUTILIZAR\el Número de Comprobante Anulado.")
    End If
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    'TOMO EL CLIREAL PARA USAR LUEGO EN EL MOVISTO
    CONDI$ = "CodiCom_Lar= '" & DOCULA$ & "' And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "'"
    CLIENRE = XVALO(VALOBADA("CAJABANC", "CLI_REAL", CONDI$, "NOMESS"))
    NC1 = NCLI
    If CLIENRE > 0 Then NC1 = CLIENRE
    '
    If REUTI% = 1 Then
         strSQL = "Delete CajaBanc where "
         strSQL = strSQL & " CodiCom_Lar= '" & DOCULA$ & "' "
         strSQL = strSQL & " And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "' "
         Call EJECUTACOMANDO(strSQL)
       Else
         strSQL = "Update CajaBanc set  ValAbsComp = 0,IDEBECOMP = 0,IHABECOMP = 0 "
         strSQL = strSQL & " where CodiCom_Lar= '" & DOCULA$ & "' "
         strSQL = strSQL & " And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "' "
         Call EJECUTACOMANDO(strSQL)
    End If
    '
    'jandy Delete en SaDeuPen
'    STRSQL = "Delete SaDeuPen where "
'    STRSQL = STRSQL & "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLI
'    Call EJECUTACOMANDO(STRSQL)

    'ACTUALIZADO CON BORRAREGISTRO
    CONDI$ = "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLI
    Call BORRAREGISTROS("SADEUPEN", CONDI$, REG&, "NOMESS")
    '
    'jandy Update en Movisto
'    STRSQL = "Update Movisto set DoSecLar= ' ', DoSecCor=' ' "
'    STRSQL = STRSQL & " where (DoSecLar = '" & DOCULA$ & "'"
'    STRSQL = STRSQL & " OR  DoPriLar = '" & DOCULA$ & "')"
'    STRSQL = STRSQL & " AND Nume_Clie = " & NCLI
'
'    Call EJECUTACOMANDO(STRSQL)
    'ACTUALIZADO CON ACTUREGISTRO
   
    CONDI$ = "(DoSecLar ='" & DOCULA$ & "' OR  DoPriLar = '" & DOCULA$ & "') AND  Nume_Clie = " & NC1
    'SE CAMBIA EL ORDEN PORQUE LA BUSQUEDA ES POR DOSECLAR, Y SI SE PONE EN BLANCO LUEGO NO ES POSIBLE
    'UBICAR EL REGISTRO PARA LIMPIAR EL CAMPO DOSECCOR
    Call ACTUREGISTRO("MOVISTO", "DOSECCOR", CONDI$, " ", REG&, "NOMESS")
    Call ACTUREGISTRO("MOVISTO", "DOSECLAR", CONDI$, " ", REG&, "NOMESS")
    'REALIZA LO MISMO SI HAY MOVIMIENTOS PARA LA MISMA CONDI EN MOVIREPA
'    If CantRegistros&("MOVIREPA", CONDI$, "NOMESS") > 0 Then ' OD-05/12/11 COMENTADO POR REDUNDANTE
     Call ACTUREGISTRO("MOVIREPA", "DOSECCOR", CONDI$, " ", REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "DOSECLAR", CONDI$, " ", REG&, "NOMESS")
'    End If
    
    
'    STRSQL = "Update Movisto set DoSecLar= ' ', DoSecCor=' ' "
'    STRSQL = STRSQL & ", CantIngre=0 , CantSalid=0 "
'    STRSQL = STRSQL & " where DoPriLar = '" & DOCULA$ & "' "
'    STRSQL = STRSQL & " AND Nume_Clie = " & NCLI
'    Call EJECUTACOMANDO(STRSQL)
    '
    'ACTUALIZADO CON ACTUREGISTRO
    CONDI$ = "DoPriLar = '" & DOCULA$ & "' AND Nume_Clie = " & NC1
    Call ACTUREGISTRO("MOVISTO", "DOSECLAR", CONDI$, " ", REG&, "NOMESS")
    Call ACTUREGISTRO("MOVISTO", "DOSECCOR", CONDI$, " ", REG&, "NOMESS")
    Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI$, 0, REG&, "NOMESS")
    Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, 0, REG&, "NOMESS")
    'REALIZA LO MISMO SI HAY MOVIMIENTOS PARA LA MISMA CONDI EN MOVIREPA
'    If CantRegistros&("MOVIREPA", CONDI$, "NOMESS") > 0 Then ' OD-05/12/11 COMENTADO POR REDUNDANTE
     Call ACTUREGISTRO("MOVIREPA", "DOSECLAR", CONDI$, " ", REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "DOSECCOR", CONDI$, " ", REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "CANTINGRE", CONDI$, 0, REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "CANTSALID", CONDI$, 0, REG&, "NOMESS")
'    End If
    '
'    'jandy Delete en APLICOBRA
'    STRSQL = "Delete APLICOBRA where "
'    STRSQL = STRSQL & " CoDocOr_Lar = '" & DOCULA$ & "' "
'    STRSQL = STRSQL & " OR CoDocAp_Lar = '" & DOCULA$ & "' "
'    STRSQL = STRSQL & " AND NuCLIEN = " & NCLI
'    Call EJECUTACOMANDO(STRSQL)
    'ACTUALIZADO CON BORRAREGISTRO
    CONDI$ = "(CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "') AND NuCLIEN = " & NCLI
    Call BORRAREGISTROS("APLICOBRA", CONDI$, REG&, "NOMESS")
    '
'    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
'    For KUL& = Len(RFG$) - 3 To 1 Step -4
'      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
'        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
'      End If
'    Next KUL&
    '
    'BLANQUEA EL NUMERO DE FACTURA EN LOS NUMEROS DE SERIE QUE TIENE A LA FACTURA ANULADA
    CONDI$ = "NUMFACTURA='" & DOCULA$ & "'"
    Call ACTUREGISTRO("TANUMSE", "NUMFACTURA", CONDI$, "", REGAF&, "NOMESS")
    '
    
    'BORRA NUMERO DE FACTURA DE LISTA DE EMPAQUE
    'SE VALIDA POR QUE ESTA TABLA SE GENERA FUERA DEL ABREBASE
    If EXISTELCAMPO%("", "ENCEMPAQ", "NUMFACTURA", 202, 18) > 0 Then
       CONDI$ = "NUMFACTURA = '" & DOCULA$ & "' AND CODI_CLI = " & NCLI
       Call ACTUREGISTRO("ENCEMPAQ", "NUMFACTURA", CONDI$, "", REGAF&, "NOMESS")
    End If
    '
    If EXISTE_TABLA%("factuanti") > 0 Then
       'arriba se valida si no tiene documentos aplicados
       CONDIFACTUANTI$ = "docu_orig = '" & DOCULA$ & "'  AND NUCLIEN = " & NCLI
       Call BORRAREGISTROS("factuanti", CONDIFACTUANTI$, REGAF&, "NOMESS")
    End If
    
    Screen.MousePointer = 1
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    RECUEC07.TEXT1.TEXT = ""
    RECUEC07.TEXT1.TEXT = TRIM$(Str$(NCLI))
    On Error Resume Next
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    RECUEC07.Show 1
    Call FRESHALL
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Sub ANUFORM()
    'ANULACION DE FACTURA DE FORMULARIO ARRUINADO
    ANUFORM07.Show 1
End Sub

Private Sub Command5_Click_vie()
    Command5.Enabled = False
    On Error Resume Next
    If DERACCE%("GETTOKEN", "Obtener TOKEN Diario del Web Service") >= 2 Then
       Kill LUDAT$ + "TOKEN.TXT"
       On Error GoTo 0
       '
       DINTER$ = LOADFILE$(LUDAT$ + "DINT_WSA.TXT")
       If DINTER$ = "" Then DINTER$ = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?WSDL"
10     DINTER$ = TRIM$(InputBox$("URL", "Direccion de Internet - Web Service Autenticacion", DINTER$))
       If DINTER$ = "" Then GoTo 99
       If Left$(DINTER$, 8) <> "https://" Then
          Call MENSERR(24, "Direccion URL no Válida.")
          GoTo 10
       End If
       Call SAVEFILE(LUDAT$ + "DINT_WSA.TXT", DINTER$)
       '
       ATRIBU$ = " -s " + Chr$(34) + "wsfe" + Chr$(34)
       ATRIBU$ = ATRIBU$ + " -c " + Chr$(34) + LUDAT$ + "MICERTIFICADO.PFX" + Chr$(34)
       ATRIBU$ = ATRIBU$ + " -w " + Chr$(34) + DINTER$ + Chr$(34)
       ATRIBU$ = ATRIBU$ + " -v ON"
       ' obtiene cuit de la empresa
       n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
       Dim X As String * 41
       Get #n1%, 6, X
       CUITT$ = REPLACAR$(TRIM$(DESENCRI$(X)), "-", "")
       Close #n1%
       '
       Call SAVEFILE(LUDAT$ + "CUIT.TXT", CUITT$)
       '
       ORIG$ = App.Path
       If Right$(ORIG$, 1) <> "\" Then ORIG$ = ORIG$ + "\"
       If EXISTE%(ORIG$ + "ClienteLoginCms_VB.EXE") < 1 Then
          ARCHOR$ = REPLACAR$(ORIG$ + "ClienteLoginCms_VB.EXE", "\", "/")
          Call MENSERR(24, "Falta Archivo '" + ARCHOR$ + "'")
          GoTo 99
       End If
       '
       Call FileCopy(ORIG$ + "ClienteLoginCms_VB.EXE", LUCOM$ + "ClienteLoginCms_VB.EXE")
       If EXISTE%(LUCOM$ + "ClienteLoginCms_vb.exe") = 1 Then
          On Error Resume Next
          XXX = Shell(LUCOM$ + "\ClienteLoginCms_VB.EXE" + ATRIBU$, 4)
          If Err Then
             On Error GoTo 0
             Call MENSERR(24, "No se ha Generado el 'TOKEN.TXT'")
             GoTo 99
          End If
          If TRIM$(Control$("FACTURA", "PUVENCAP")) <> "" Then ' TOKEN PARA BIENES DE CAPITAL
            Call COMUNI("Aguarde a que se Cierre\La Ventana de Generación de Token\Y Pulse Aceptar")
            Call REPLATEXT(ATRIBU$, "wsfe", "wsbfe")
            XXX = Shell(LUCOM$ + "\ClienteLoginCms_VB.EXE" + ATRIBU$, 4)
            If Err Then
               On Error GoTo 0
               Call MENSERR(24, "No se ha Generado el 'TOKEN_WSBFE.TXT'")
               GoTo 99
            End If
          End If
          On Error GoTo 0
       End If
       Call FINAL("")
    End If
    '
99  Command5.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    If DERACCE%("GETTOKEN", "Obtener TOKEN Diario del Web Service") >= 2 Then
       CARPEACTUAL$ = App.Path + "\"
       CARPEDESTIN$ = "C:\FLEXOFT\FACELEC\"
       '
 MsgBox "CARPEACTUAL " & CARPEACTUAL$
 MsgBox "CARPEDESTIN " & CARPEDESTIN$
       ' CREAR CARPETA DE TRABAJO
       On Error Resume Next
       MkDir "C:\FLEXOFT"
       MkDir "C:\FLEXOFT\FACELEC"
       On Error GoTo 0
       If EXISTE%("C:\FLEXOFT\FACELEC\.") < 1 Then
          Call MENSERR(24, "Imposible crear carpeta 'C:/FLEXOFT/FACELEC'")
          GoTo 99
       End If
       '
       'Eliminar archivos
       If EXISTE%(CARPEDESTIN$ + "WSFE_CLIENTE.EXE") > 0 Then
        MsgBox "2"
            FILEKILL (CARPEDESTIN$ + "WSFE_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSFEV1_Cliente.exe") > 0 Then
        MsgBox "3"
            FILEKILL (CARPEDESTIN$ + "WSFEV1_Cliente.exe")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSBFE_CLIENTE.EXE") > 0 Then
        MsgBox "4"
            FILEKILL (CARPEDESTIN$ + "WSBFE_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSBFEV1_CLIENTE.EXE") > 0 Then
       MsgBox "5"
            FILEKILL (CARPEDESTIN$ + "WSBFEV1_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSFEX_CLIENTE.EXE") > 0 Then
       MsgBox "6"
            FILEKILL (CARPEDESTIN$ + "WSFEX_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSFEXV1_Cliente.EXE") > 0 Then
       MsgBox "7"
            FILEKILL (CARPEDESTIN$ + "WSFEXV1_Cliente.EXE")
       End If

       '
       'Generar archivos LINK Y CUIT
       DINTER$ = LOADFILE$(LUDAT$ + "LINK.TXT")
       If DINTER$ = "" Then DINTER$ = "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL"
'10     DINTER$ = TRIM$(InputBox$("URL", "Direccion de Internet - Web Service Autenticacion", DINTER$))
'       If DINTER$ = "" Then GoTo 99
'       If Left$(DINTER$, 8) <> "https://" Then
'          Call MENSERR(24, "Direccion URL no Válida.")
'          GoTo 10
'       End If
       Call SAVEFILE(LUDAT$ + "LINK.TXT", DINTER$)
       '
       'Obtiene cuit de la empresa y lo valida
       n1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
       Dim X As String * 41
       Get #n1%, 6, X
       'VALIDACION
       CUITT$ = TRIM$(DESENCRI$(X))
       Valid% = VALICUIT%(CUITT$)
       If Valid% < 1 Then
          Call COMUNI("C.U.I.T. de Empresa No Válido")
          Call FINAL("")
       End If
       '
       CUITT$ = TRIM$(REPLACAR$(CUITT$, "-", ""))
       Close #n1%
       Call SAVEFILE(LUDAT$ + "CUIT.TXT", CUITT$)
       '
       'Copiar archivos a FACELEC
       Call FileCopy(LUDAT$ + "LINK.TXT", CARPEDESTIN$ + "LINK.TXT")
       Call FileCopy(LUDAT$ + "CUIT.TXT", CARPEDESTIN$ + "CUITEMP.TXT")
       '
       'Validar si existe archivo Copiarlo en FACELEC
       'ObtieneToken.exe
       If EXISTE%(CARPEACTUAL$ + "OBTIENETOKEN.EXE") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "OBTIENETOKEN.EXE")
          GoTo 99
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "OBTIENETOKEN.EXE", CARPEDESTIN$ + "OBTIENETOKEN.EXE")
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "ObtieneToken ya en Ejecución\" & Err.Description)
          GoTo 99
       End If
       On Error GoTo 0
       '
       'ClienteLoginCms_VB
       If EXISTE%(CARPEACTUAL$ + "ClienteLoginCms_VB.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "ClienteLoginCms_VB.exe")
          GoTo 99
       End If
       Call FileCopy(CARPEACTUAL$ + "ClienteLoginCms_VB.EXE", CARPEDESTIN$ + "ClienteLoginCms_VB.EXE")
       '
       'micertificado
       If EXISTE%(LUDAT$ + "micertificado.pfx") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(LUDAT$, "\", "/") + "micertificado.pfx")
          GoTo 99
       End If
       Call FileCopy(LUDAT$ + "micertificado.pfx", CARPEDESTIN$ + "micertificado.pfx")
       '
       ' *** COMUN
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSFE_Cliente.EXE"
       Kill CARPEDESTIN$ + "WSFEV1_Cliente.EXE"
       On Error GoTo 0
       If XVALO(Control$("FACTURA", "PUVENELE")) < 1 Then GoTo 70
       If EXISTE%(CARPEACTUAL$ + "WSFE_Cliente.EXE") < 1 And EXISTE%(CARPEACTUAL$ + "WSFEV1_CLIENTE.EXE") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "WSFEV1_Cliente.EXE")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSFE_Cliente.EXE", CARPEDESTIN$ + "WSFE_Cliente.EXE")
       Call FileCopy(CARPEACTUAL$ + "WSFEV1_Cliente.EXE", CARPEDESTIN$ + "WSFEV1_Cliente.EXE")
       On Error GoTo 0
       '
70     ' *** BIENES DE CAPITAL
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSBFE_Cliente.EXE"
       Kill CARPEDESTIN$ + "WSBFEV1_Cliente.EXE"
       On Error GoTo 0
       If TRIM$(Control$("FACTURA", "PUVENCAP")) = "" Then GoTo 75
       
       If EXISTE%(CARPEACTUAL$ + "WSBFE_Cliente.exe") < 1 And EXISTE%(CARPEACTUAL$ + "WSBFEV1_Cliente.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "WSBFE_Cliente.exe")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSBFE_Cliente.EXE", CARPEDESTIN$ + "WSBFE_Cliente.EXE")
       Call FileCopy(CARPEACTUAL$ + "WSBFEV1_Cliente.EXE", CARPEDESTIN$ + "WSBFEV1_Cliente.EXE")
       On Error GoTo 0
       
       '
75     ' *** EXPORTACION
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSFEX_Cliente.EXE"
       Kill CARPEDESTIN$ + "WSFEXV1_Cliente.EXE"
       On Error GoTo 0
       If TRIM$(Control$("FACTURA", "PUVEFAEX")) = "" Then GoTo 80
       If EXISTE%(CARPEACTUAL$ + "WSFEX_Cliente.exe") < 1 And EXISTE%(CARPEACTUAL$ + "WSFEXV1_Cliente.exe") < 1 Then
          Call COMUNI("Falta Ejecutable para el servicios WSFEX")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSFEX_Cliente.EXE", CARPEDESTIN$ + "WSFEX_Cliente.EXE")
       Call FileCopy(CARPEACTUAL$ + "WSFEXV1_Cliente.EXE", CARPEDESTIN$ + "WSFEXV1_Cliente.EXE")
       On Error GoTo 0
      
80     XXX = Shell(CARPEDESTIN$ + "OBTIENETOKEN.EXE", vbNormalFocus)
       
    End If
    '
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   
   If DERACCE%("STPFACTU", "Configuración de Funcionamiento Facturacion") > 1 Then
      OPFACTU07.Show 1
      Call FINAL("")
   End If
   '
   Command6.Enabled = True
   '
End Sub

Sub Command7_Click()
  TESTCAE09.MODOAUTO = "NO"
  TESTCAE09.Show 1
End Sub


Private Sub Command8_Click()
  
   Command8.Enabled = False
   '
   If DERACCE%("RECOMELE", "Recuperacion de Comprobantes Electronicos") < 2 Then Exit Sub
   '
   CARPETOK$ = "C:\FLEXOFT\FACELEC\"
   If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
   '
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   DESA = FETEXCOR1$(Des%)
   HASA = FETEXCOR1$(Has%)
   '
   'CREA TABLAS DE COMPROBANTES NO RECUPERADOS Y RECUPERADOS
   '
   Call CREACAMPO("", "FACNOENC", "NFACTURA", 10, 18)
   Call CREACAMPO("", "FACNOENC", "OBSERVA", 12, 0)
   '
   Call CREACAMPO("", "FARECAFIP", "FechEmis", 8, 0)
   Call CREACAMPO("", "FARECAFIP", "NFACTURA", 10, 18)
   Call CREACAMPO("", "FARECAFIP", "IMPORTE", 7, 0)
   Call CREACAMPO("", "FARECAFIP", "NuClien", 3, 2)
   '
   'MATRIZ DONDE SE ALMACENAN LOS COMPROBANTES A RECUPERAR
   '
   ReDim MATRIZ$(2, 2)
   
   'SELECCION DE PUNTO DE VENTA A REALIZAR PROCESO y VALIDACION DE EJECUTABLES DE PUNTO NET
   OPX% = COMALTER%("Seleccion de Punto de Venta:\\Factura Electrónica Comun\Factura Electrónica B.Capital\Factura Electrónica Exportación")
   Select Case OPX%
      Case 1
        PVTA% = XVALO(Control("FACTURA", "PUVENELE"))
        If EXISTE%(CARPETOK$ + "wsfe_cliente.exe") = 0 Then
           Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
           GoTo 99
        End If
        '
      Case 2
          PVTA% = XVALO(Control("FACTURA", "PUVENCAP"))
          If EXISTE%(CARPETOK$ + "wsbfe_cliente.exe") = 0 Then
             Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
             GoTo 99
          End If
      Case 3
          PVTA% = XVALO(Control("FACTURA", "PUVEFAEX"))
          If EXISTE%(CARPETOK$ + "wsfex_cliente.exe") = 0 Then
             Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
             GoTo 99
          End If
          '
      Case Else: GoTo 99
   End Select
   '
   Screen.MousePointer = 11
   '
   'BORRA EL CONTENIDO DE LA TABLA DE FACTURAS NO RECUPERADAS SEGUN PUNTO DE VENTA ELEGIDO
   '
   PVTAX$ = "0000" & PVTA%: PVTAX$ = Right(PVTAX$, 4)
   Condi1$ = "SUBSTRING(NFACTURA, 4, 4) = '" & PVTAX$ & "'"
   Call BORRAREGISTROS("FACNOENC", Condi1$, REGAF&, "NOMESS")
   '
   'CONDICION DE CONSULTA
   '
   CONDI$ = " tipomovim = 'fven' "
   CONDI$ = CONDI$ + " AND NUMECOMP > 0 "
   CONDI$ = CONDI$ + " AND  NUMETALO =  " & PVTA%
   CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESA & "'"
   CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HASA & "'"
   CONDI$ = CONDI$ + " ORDER BY OPCIMOVIM, VARIMOVIM, NUMECOMP asc"
   '
   FIN& = CantRegistros("CAJABANC", CONDI$)
   '
   'RECORRO CAJABANC VERIFICANDO CORRELATIVIDAD Y GENERO LA MATRIZ CON LOS
   'DOCUMENTOS A RECUPERAR
   '
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   '
   i& = 0: J& = 0: VUELTA% = 0: VARI$ = "": OPCI$ = "": U& = 0
   Dim CABATMP As New ADODB.Recordset
   Set CABATMP = FLEXCONN.Execute(SQLX$)
   With CABATMP
       Do While Not .EOF
          '
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          'COMPARA LOS VALORES PARA VER SI CAMBIO EJEMPLO DE 'FC' A 'NC' O LA LETRA
          If SAFETEXT$(!VARIMOVIM) <> VARI$ Or SAFETEXT$(!OpciMovim) <> OPCI$ Then
              i& = XVALO(!NumeComp)
              VARI$ = SAFETEXT$(!VARIMOVIM)
              OPCI$ = SAFETEXT$(!OpciMovim)
           End If
          '
10        If i& < XVALO(!NumeComp) Then 'SI SE ROMPE LA CARRELATIVIDAD
'           'CARGA EL VECTOR
            J& = J& + 1
            ReDim Preserve MATRIZ$(2, J&)
            'ARMO EL STRING DE LA FACTURA DEL HUECO
            NUMCOM$ = "00000000" & (i&): NUMCOM$ = Right(NUMCOM$, 8)
            'ASIGNO A LA MATRIZ$ LA LECTURA DEL CAE ANTERIOR AL HUECO Y FACTURA DEL HUECO
            MATRIZ$(1, J&) = OPCI$ + "-" + PVTAX$ + "-" + NUMCOM$ + "-" + VARI$
            MATRIZ$(2, J&) = ID$
            i& = i& + 1
            '
            If i& < XVALO(!NumeComp) Then ' SI HAY MAS DE UN HUECO CONSECUTIVO
              GoTo 10                     ' VUELVE Y COMPARA A VER SI EL SIGUIENTE
            End If                        ' ES EL QUE SIGUE
            '
         End If
         '
         i& = i& + 1
         NUFACTU$ = SAFETEXT$(!codicom_lar)
         ID$ = SAFETEXT$(!IDCAE)
         .MoveNext
        Loop
    End With
    '
    INDIC& = J&
    If INDIC& > 0 Then ' SI HAY HUECOS EN LAS FACTURAS HABILITO EL TIMER PARA INICIAR PROCESO
      Timer1.Enabled = True
    Else
      Screen.MousePointer = 1
      Call COMUNI("No Se Encontraron Facturas No Correlativas Para el Punto de Venta: " & PVTA%)
    End If
    '
99  Command8.Enabled = True
    '
End Sub

Sub Command9_Click()
   Command9.Enabled = False
   
   If DERACCE%("ELCOANU", "Eliminación de Comprobantes Anulados sin CAE") >= 2 Then
      Screen.MousePointer = 11
      CONDI$ = "Abs(IDEBECOMP) < 0.005 And Abs(IHABECOMP) < 0.005 And NUMECOMP = (-1)"
      Call BORRAREGISTROS("CAJABANC", CONDI$, REGAF&, "NOMESS")
      Screen.MousePointer = 1
      Call COMUNI("Proceso Terminado")
   End If
   '
   Command9.Enabled = True
End Sub

Private Sub CuContaPorCodigo_Click()

20  Call SELECHO("GRUCOVEN")
    GRUCO$ = VALACT1$("GRUCOVEN")

    If GRUCO$ = "" Then
        OPX% = COMALTER%("¿Continua con la Asignación de Cuentas Contables por Artículo?\\No, Finalizar\Si, Continuar")
        If OPX% = 1 Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    Campos$ = "Codigo/A20;Descripción/A48;C.Contable/A24"
    Call CARGA_ENCABEZADO(MUESTRADATOS5.GRID, Campos$, MUESTRADATOS5, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(MUESTRADATOS5.MSF_2, Campos$, MUESTRADATOS5, 1, 0, ANTOT, , 1)
    
    'PRESENTACION DE LOS COMPROBANTES QUE NO TIENEN CUIT
30  SQLX$ = " SELECT Codigo, Descripcion, CueContaI, Grucoven FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Status > 0"
    If GRUCO$ <> "" Then SQLX$ = SQLX$ + " AND Grucoven=" & GRUCO$
    SQLX$ = SQLX$ + " ORDER BY Codigo ASC"
    Set rs1 = READSET(SQLX$)
    J& = 0
    With rs1
      Do While Not .EOF
        Codigo$ = SAFETEXT$(!Codigo)
        DESCRI$ = SAFETEXT$(!Descripcion)
        CUECO% = XVALO(!CUECONTAI)
        '
        J& = J& + 1
        MUESTRADATOS5.GRID.Rows = J& + 1
        MUESTRADATOS5.GRID.TextMatrix(J&, 1) = Codigo$
        MUESTRADATOS5.GRID.TextMatrix(J&, 2) = DESCRI$
        MUESTRADATOS5.GRID.TextMatrix(J&, 3) = CODCUE(CUECO%)
        .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    If J& > 0 Then
       Screen.MousePointer = 1
       MUESTRADATOS5.GRID.Visible = True
       MUESTRADATOS5.GRID.ZOrder 0
       MUESTRADATOS5.Caption = "Asignación de Cuenta Contable por Código"
       MUESTRADATOS5.Show 1
       GoTo 20
    Else
       Call COMUNI("No se Han Encontrado Artículos Que Pertenezcan al Grupo Contable")
       Screen.MousePointer = 1
       GoTo 20
    End If

End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture2)
End Sub

Private Sub mnuCtlCorr_Click()
'   If CLACONTRA% > 0 Then
     CTRCORREL.Show 1
'   End If
End Sub


Private Sub Timer1_Timer()
      '
      Static i&, PROCESOACTIVO%
      Static FACTU$
      FIN1& = INDIC&
      CARPETOK$ = "C:\FLEXOFT\FACELEC\"
      '
      'VERIFICO SI EXISTE UN PROCESO ACTIVO DONDE SE ESTA BUSCANDO
      'LOS DATOS DEL COMPROBANTE FALTANTE O SI NO HAY MAS COMPROBANTES
      If PROCESOACTIVO% = 0 And i& < FIN1& Then
          'TRATO DE RECUPERAR EL PROXIMO COMPROBANTE
          i& = i& + 1
          FACTU$ = MATRIZ$(1, i&)
          IDXCAE$ = MATRIZ$(2, i&)
          'SE VALIDA QUE TENGA ID DE TRANSACCION, CASO CONTRARIO NO RECUPERA EL COMPROBANTE
          If TRIM$(IDXCAE$) <> "" Then
                PROCESOACTIVO% = 1
                EJECUTABLE$ = "wsfe_cliente.exe "
                If PVTA% = XVALO(Control("FACTURA", "PUVENCAP")) Then
                   EJECUTABLE$ = "wsbfe_cliente.exe "
                ElseIf PVTA% = XVALO(Control("FACTURA", "PUVEFAEX")) Then
                   EJECUTABLE$ = "wsfex_cliente.exe "
                End If
                OPX% = COMALTER%("Desea Recuperar Comprobante '" & FACTU$ & "'\\No Cancelar\Si, Recuperar")
                If OPX% = 1 Then
                   Timer1.Enabled = False
                   Screen.MousePointer = 1
                   Call COMUNI("Proceso de Recuperacion Cancelado")
                   Exit Sub
                End If
                XXX = Shell(CARPETOK$ & EJECUTABLE$ & " RECU " & FACTU$ & " " & IDXCAE$, 4)
          End If
      Else
          'VERIFICA SI SE PUDO RECUPERAR EL COMPROBANTE DESDE LA AFIP
          Timer1.Enabled = False  ' DETIENE LA BUSQUEDA HASTA VER QUE PASO CON EL ULTIMO COMPROBANTE
          KAE$ = LEEARCHDEVU$(FACTU$, MOTIVO$) ' LEO EL ARCHIVO DEVUELTO.
          '
          If XVALO(KAE$) > 0 Then ' SI LO DEVUELTO TIENE CAE PUEDO RECUPERAR EL COMPROBANTE
               'RECUPERE EL COMPROBANTE Y GRABO LA INFORMACION EN LAS TABLAS
               Call GrabFactRecuperada(FACTU$)
          Else
               'NO PUDE RECUPERAR EL COMPROBANTE Y LO GUARDO EN LA LISTA DE NO RECUPERADOS
               OBSERVACION$ = MOTIVO$
               If InStr(KAE$, "CANCELAR") > 0 Then OBSERVACION$ = KAE$
               Call GrabFactNoRecup(FACTU$, OBSERVACION$)
          End If
          '
          If i& >= FIN1& Then ' SI SE ACABARON LOS COMPROBANTES, FINALIZO EL PROCESO
            Timer1.Enabled = False
            i& = 0
            FIN1& = 0
            Screen.MousePointer = 1
            Call COMUNI("Proceso de Recuperacion Terminado")
            Exit Sub
          End If
          '
          PROCESOACTIVO% = 0 ' LIBERO LA POSIBILIDAD DE RECUPERAR EL PROXIMO COMPROBANTE
          Timer1.Enabled = True ' INICIO NUEVAMENTE EL PROCESO DE RECUPERACION
          '
      End If

End Sub

Function LEEARCHDEVU$(FACTU$, MOTIVO$)
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LEEARCHDEVU$ = ""
    Cae$ = ""
    MOTIVO$ = ""
    '
    'VERIFICO SI EXISTE EL ARCHIVO CON LOS DATOS DEL COMPROBANTE A RECUPERAR
    If EXISTE%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT") > 0 Then
        'VERIFICO SI EL CONTENIDO DEL ARCHIVO TIENE EL CAE. EN CASO DE QUE SI,
        'SE ASUME QUE EL COMPROBANTE SE HA PODIDO RECUPERAR
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cae>")) + 5
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             Cae$ = Mid$(TXT$, 1, POS2% - 1)
             Exit Do
          End If
        Loop
        Close NARCHIVO%
        
        LEEARCHDEVU$ = Cae$

        'SI EL ARCHIVO NO CONTIENE EL CAE, ENTONCES NO SE HA PODIDO RECUPERAR EL COMPROBANTE
        'SE BUSCA SI EXISTE ALGUN MOTIVO POR EL CUAL NO SE PUDO RECUPERAR EL COMPROBANTE
        If XVALO(Cae$) < 1 Then
          MOTIVO1$ = ""
          MOTIVO2$ = ""
          VUELTA% = 0 ' LECTURA DE MOTIVO DEL ARCHIVO CAE.TXT
          NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
          Do While Not EOF(NARCHIVO%) ' RECORRE
              Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
              If InStr(XX$, "<motivo>") > 0 Or InStr(XX$, "<Obs>") Then
                 POS1% = InStr(XX$, ">")
                 POS2% = InStr(XX$, "</")
                 Hasta% = POS2% - (POS1% + 1)
                 If VUELTA% < 1 Then ' SI ES LA PRIMERA VEZ ASIGNA MOTIVO1
                  MOTIVO1$ = Mid$(XX$, POS1% + 1, Hasta%)
                  VUELTA% = VUELTA% + 1
                 ElseIf VUELTA% > 0 Then ' LA SEGUNDA VEZ QUE PASA CARGA MOTIVO2
                  MOTIVO2$ = "-" & Mid$(XX$, POS1% + 1, Hasta%)
                 End If
              End If
          Loop
          '
          Close NARCHIVO%
          MOTIVO$ = MOTIVO1$ & MOTIVO2$ ' ASIGNA LOS DOS MOTIVOS
          LEEARCHDEVU$ = MOTIVO$
          '
       End If
    End If
    '
End Function

Sub GrabFactRecuperada(FACTU$)
 
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    '
    Dim CUEIVX$(2)
    CONT% = 0
    TOTAGEN# = 0
    NENETO# = 0
    IMPOIVA# = 0
    PERIB7# = 0
    TICAMBIOX# = 0

    NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    '
    'PUNTO DE VENTA COMUN
    If PVTA% = XVALO(Control("FACTURA", "PUVENELE")) Then '
         Do While Not EOF(NARCHIVO%) ' RECORRE
            Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            POS1% = InStr(XX$, ">")
            POS2% = InStr(XX$, "</")
            Hasta% = POS2% - (POS1% + 1)
            If Hasta% < 0 Then Hasta% = 1
            
            TX$ = Mid$(XX$, POS1% + 1, Hasta%)

               Select Case CONT%
                  Case 4
                  IDECAE$ = TRIM$(TX$)
                 Case 16
                    NC% = CLIENTCUIT&(TX$)
                 Case 21
                   TOTAGEN# = ROUND#(XVALO(TX$), 2)
                 Case 23
                   NENETO# = ROUND#(XVALO(TX$), 2)
                 Case 24
                   IMPOIVA# = ROUND#(XVALO(TX$), 2)
                 Case 22
                   PERIB7# = ROUND#(XVALO(TX$), 2)
                 Case 28
                   Cae$ = TRIM$(TX$)
                 Case 29
                   FFI% = INVERFECHA%(TX$)
               End Select
            Loop
    End If
    '
    'PUNTO DE VENTA DE BIENEES DE CAPITAL
    If PVTA% = XVALO(Control("FACTURA", "PUVENCAP")) Then
      Do While Not EOF(NARCHIVO%) ' RECORRE
            Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            POS1% = InStr(XX$, ">")
            POS2% = InStr(XX$, "</")
            Hasta% = POS2% - (POS1% + 1)
            If Hasta% < 0 Then Hasta% = 1
            TX$ = Mid$(XX$, POS1% + 1, Hasta%)
               Select Case CONT%
                 Case 4
                  IDECAE$ = TRIM$(TX$)
                 Case 7 'CUIT
                    NC% = CLIENTCUIT&(TX$)
                 Case 11
                   TOTAGEN# = ROUND#(TX$, 2)
                 Case 13
                   NENETO# = ROUND#(TX$, 2)
                 Case 14
                   IMPOIVA# = ROUND#(TX$, 2)
                 Case 18
                   PERIB7# = ROUND#(TX$, 2)
                 Case 23
                   FFI% = INVERFECHA%(TX$)
                 Case 26
                   Cae$ = TRIM$(TX$)
                 Case 31
               End Select
            Loop
    End If
    '
   
    'PUNTO DE VENTA DE EXPORTACION
    If PVTA% = XVALO(Control("FACTURA", "PUVEFAEX")) Then
      IMPOIVA# = 0: PERIB7# = 0
      Do While Not EOF(NARCHIVO%) ' RECORRE
            Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            POS1% = InStr(XX$, ">")
            POS2% = InStr(XX$, "</")
            Hasta% = POS2% - (POS1% + 1)
            If Hasta% < 0 Then Hasta% = 1 ' por si no encuentra las posiciones
            TX$ = Mid$(XX$, POS1% + 1, Hasta%)
               Select Case CONT%
                 Case 4
                  IDECAE$ = TRIM$(TX$)
                 Case 5
                   FFI% = INVERFECHA%(TX$)
                 Case 15
                    NC% = CLIENTCUIT&(TX$)
                 Case 19
                   TOTAGEN# = ROUND#(XVALO(TX$), 2)
                   NENETO# = ROUND#(XVALO(TX$), 2)
                 Case 37
                   Cae$ = TRIM$(TX$)
                 Case 16
                   DENOMIMONE$ = TRIM$(TX$)
                 Case 17
                   TICAMBIOX# = XVALO(TX$)
               End Select
            Loop
    End If
    
    Close NARCHIVO%
    
    'GRABACION EN EL CAJABANC
    '************************
    '
    MsgBox IDECAE$
    MsgBox NC%
    MsgBox TOTAGEN#
    MsgBox NENETO#
    MsgBox IMPOIVA#
    MsgBox PERIB7#
    MsgBox FFI%
    MsgBox Cae$
    '
    'OBTENCION DE LOS DATOS
    NFCY$ = FACTU$
    NFCX$ = Left$(FACTU$, 3) + Mid$(FACTU$, 11, 6) + Right(FACTU$, 2)
    OPMOVIX$ = UCase(Left$(FACTU$, 2))
    LETRA$ = Right(NFCY$, 1)
    PTOVTASTR$ = Mid$(NFCY$, 4, 4)
    SIGX% = 1
    If OPMOVIX$ = "NC" Then SIGX% = (-1)  ' SI ES NOTA DE CREDITO
    '
    'RE-CALCULO DE IMPORTES SEGUN MONEDA DE EMISION DE LA FACTURA. CASO EXPORTACION
    If TICAMBIOX# < 0.00005 Then TICAMBIOX# = 1
    TOTAGEN# = TOTAGEN# * TICAMBIOX#
    NENETO# = NENETO# * TICAMBIOX#
    TOTADOLAR# = TOTAGEN# / TICAMONE(2)
    '
    'COTIMONEM# = TICAMBIOX#
    MONEMI% = TIPOMONEDA%(DENOMIMONE$)
    
    'GRABA REGISTRO POR EL TOTAL DE LA FACTURA
    CoComCor$ = NFCX$
    CoComLar$ = NFCY$
    NumeClie = Int(GRUDIVCLI(NC%))
    NumeProv% = 0
    FechEmi = CVDAT(FFI%)
    Referen$ = rasocli$(NC%)
    FeContab = CVDAT(FFI%)
    FeAcredi = CVDAT(FFI%)
    CuentaCo% = CUECOCLI%(NC%)
    TiMovi$ = "FVEN"
    OpMovi$ = OPMOVIX$
    VaMovi$ = LETRA$
    DOCUMEN$ = "Factura:": If OPMOVIX$ = "NC" Then DOCUMEN$ = "N.Crédito:": If OPMOVIX$ = "ND" Then DOCUMEN$ = "N.Débito:"
    Descrip$ = DOCUMEN$ + " " + LETRA$ + " (Mat.Codif.)" + " " + PTOVTASTR$ + "-" + TRIM$(Str(XVALO(Mid$(NFCY$, 9, 8))))
    NumeTal = XVALO(PTOVTASTR$)
    NumeCom = XVALO(Mid$(NFCY$, 9, 8))
    StatInit = 0
    ImpoDeb = TOTAGEN# * SIGX%
    ValoDola = TOTADOLAR# * SIGX% ' VER COMO SE VA A GRABAR SI ES A VALOR DOLAR ACTUALIZADO
    Observa$ = rasocli$(NC%)
    NumeroCae$ = Cae$
    IdentiCae$ = IDECAE$
    ObservaFactu$ = "Factura Recuperada Automaticamente Desde AFIP"
    Call GRACABAN
    '
    'GRABA REGISTRO POR EL IMPORTE NETO DE LA FACTURA
    TOTANETX# = NENETO#
    If TOTANETX# >= 0.005 Then
      PPVI% = PIVACLI%(NC%)
      If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
      CUENTAX$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
      QQ1% = CUECOINT%(CUENTAX$)
      CuentaCo% = QQ1%
      TiMovi$ = "VENTA"
      OpMovi$ = OPMOVIX$
      VaMovi$ = LETRA$
      Descrip$ = "Mercaderia Vendida " + NFCX$
      StatInit = 0
      ImpoDeb = TOTANETX# * SIGX% * (-1)
        NENETODOL# = (NENETO# / TICAMONE(2))
      ValoDola = NENETODOL# * SIGX% * (-1)
      Call GRACABAN
    End If
    '
    'GRABA REGISTRO POR EL IMPORTE DE IVA DE LA FACTURA
    CUEIVX$(0) = Control$("SVD451S", "CUEIVA")
    CUEIVX$(1) = Control$("SVD451S", "CUEIVB")
    If TRIM$(CUEIVX$(1)) = "" Or CUECOINT%(CUEIVX$(1)) < 1 Then
     CUEIVX$(1) = CUEIVX$(0)
    End If
    TOTANETX# = IMPOIVA#
    If TOTANETX# >= 0.005 Then
       CuentaCo% = CUECOINT%(CUEIVX$(0)) 'SE IMPUTA A LA CUENTA CONTABLE DE IVA PRINCIPAL
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = FORMATNUM$((IMPOIVA# / NENETO#) * 100, "F5.2") 'SE ESTA CALCULANDO UN PORCENTAJE PROMEDIO DEL IVA PROMEDIO
            If Abs(XVALO(VaMovi$) - 10.5) > 0.001 Then CuentaCo% = CUECOINT%(CUEIVX$(1))
       Descrip$ = "Iva Debito Fiscal " + NFCX$
       StatInit = 0
       ImpoDeb = TOTANETX# * SIGX% * (-1)
       ValoDola = NENETODOL# * SIGX% * (-1)
       BasImpo# = NENETO#
       Call GRACABAN
     End If
   BasImpo# = 0
   '
   'GRABA REGISTRO POR EL IMPORTE DE PERCEPCION DE IIBB DE LA FACTURA
   TOTANETX# = PERIB7#
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = Control$("PERIBRU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-B"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB.Bs.As. " + NFCX$
     StatInit = 0
     ImpoDeb = TOTANETX# * SIGX% * (-1)
     ValoDola = NENETODOL# * SIGX% * (-1)
     BasImpo# = TOTANETX#
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'GRABA LA DEUDA EN SAUDEPEN
   '**************************
   '
   COPAGI% = CPAGCLI%(NC)
   VARIDOC% = 1
   '
   If LETRA$ = "B" Then VARIDOC% = 2
   If LETRA$ = "E" Then VARIDOC% = 3
   Tipo% = 2
   If OPMOVIX$ = "ND" Then Tipo% = 3
   If OPMOVIX$ = "NC" Then Tipo% = 6
   '
   Call FORFASQL.GRADEUFA(Tipo%, VARIDOC%, NumeCom, NC%, FFI%, COPAGI%, TOTAGEN#)
   '
   'GRABAR COMPROBANTE EN LISTA DE FACTURAS RECUPERADAS
   '***************************************************
   '
10 RNDX$ = RANDSTRING$(-10)
   If CantRegistros("FARECAFIP", "NFACTURA = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
   Call CreaRegistro("FARECAFIP", "NFACTURA", RNDX$, "NOMESS")
   CONDI$ = "NFACTURA = '" & RNDX$ & "'"
   Call ACTUREGISTRO("FARECAFIP", "FECHEMIS", CONDI$, FETEXCOR1$(FFI%), REGAF&, "NOMESS")
   Call ACTUREGISTRO("FARECAFIP", "IMPORTE", CONDI$, TOTAGEN#, REGAF&, "NOMESS")
   Call ACTUREGISTRO("FARECAFIP", "NuClien", CONDI$, NC%, REGAF&, "NOMESS")
   Call ACTUREGISTRO("FARECAFIP", "NFACTURA", CONDI$, NFCY$, REGAF&, "NOMESS")
   '
99 End Sub

Function CLIENTCUIT&(TX$)
   'OBTIENE NUMERO DE CUENTA DE CLIENTE A PARTIR DE DOCUMENTO O CUIT
   NC1& = 0
   CUICLI$ = TRIM$(TX$)
   If CUICLI$ <> "" Then
      CUICLI$ = FormatearCuitValido$(CUICLI$)
      'BUSCA POR COINCIDENCIA EXACTA POR SI SE TRATA DE UN NUMERO DE DOCUMENTO
      NC1& = XVALO(VALOBADA("DEUDOR12", "NUME_CLI", "CUIT_CLI = '" & CUICLI$ & "'"))
      If NC1& < 1 Then
        'BUSCA POR CUIT UTILIZANDO VALICUIT PARA FORMATEARLO CON LOS GUIONES
        A% = VALICUIT%(CUICLI$)
        NC1& = XVALO(VALOBADA("DEUDOR12", "NUME_CLI", "CUIT_CLI = '" & CUICLI$ & "'"))
      End If
   End If
   '
   CLIENTCUIT& = NC1&
End Function

Function INVERFECHA%(TX$)
   'CONVIERTE FECHA EN FORMATO aaaammdd EN ddmmaa Y DEVUELVE SU REPRESENTACION EN NUMERO
   FE% = 0
   TX1$ = TRIM$(TX$)
   If TX1$ <> "" Then
      DIA1 = Right(TX1$, 2)
      MES1 = Mid$(TX1$, 5, 2)
      ANIO1 = Mid$(TX1$, 3, 2)
      FCH$ = DIA1 + "/" + MES1 + "/" + ANIO1
      FE% = FECHANUM(FCH$)
   End If
   INVERFECHA% = FE%
   '
End Function
'
Sub GrabFactNoRecup(FACTU$, INFORME$)
   'GRABAR EN ARCHIVO DE FACTURAS NO RECUPERADAS
10 RNDX$ = RANDSTRING$(-10)
   '
   CONDI$ = "NFACTURA = '" & RNDX$ & "'"
   If CantRegistros("FACNOENC", "NFACTURA = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
   Call CreaRegistro("FACNOENC", "NFACTURA", RNDX$, "NOMESS")
   Call ACTUREGISTRO("FACNOENC", "OBSERVA", CONDI$, INFORME$, REGAF&, "NOMESS")
   Call ACTUREGISTRO("FACNOENC", "NFACTURA", CONDI$, FACTU$, REGAF&, "NOMESS")
   '
End Sub

Function TIPOMONEDA%(CODIMONE$)
    'DEVUELVE EL NUMERO DE MONEDA SEGUN FLEXOFT A PARTIR DE LA CODIGO DE
    'MONEDA QUE TIENE LA AFIP, EL CUAL ES PASADO COMO PARAMETRO
    MONE$ = CODIMONE$
    MONEMIX% = 1
    '
    If UCase$(MONE$) = "PES" Then
       MONEMIX% = 1
       GoTo 99
    ElseIf UCase$(MONE$) = "DOL" Then
       MONEMIX% = 2
       GoTo 99
    End If
    '
    DESCRIMONE$ = ""
    If MONE$ = "060" Then DESCRIMONE$ = "EURO"
    If MONE$ = "012" Then DESCRIMONE$ = "REAL"
    For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         REG% = Asc(Mid$(X$, 1, 1))
         If Left$(UCase$(Mid$(X$, 2, 21)), 3) = MONE$ Then
             MONEMIX% = REG%
             Exit For
         End If
    Next U&
    '
    Call CIERRARCH("TAMONED")

99  TIPOMONEDA% = MONEMIX%

End Function

