VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CFGSALSTO 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Stock"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cálculo / Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   3510
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   175
         ScaleHeight     =   585
         ScaleWidth      =   3120
         TabIndex        =   20
         Top             =   420
         Width           =   3180
         Begin VB.CommandButton Command21 
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
            Left            =   2490
            Picture         =   "CFGSALSTO.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Reporte Valorizado"
            Top             =   0
            Width           =   630
         End
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
            Height          =   600
            Left            =   1840
            Picture         =   "CFGSALSTO.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   0
            Width           =   650
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
            Height          =   600
            Left            =   0
            Picture         =   "CFGSALSTO.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Analisis de Stocks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   720
            TabIndex        =   21
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   600
         TabIndex        =   19
         Top             =   2250
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1935
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   3360
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   4530
      Left            =   3720
      ScaleHeight     =   4470
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "CFGSALSTO.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   3480
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   120
         Picture         =   "CFGSALSTO.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "CFGSALSTO.frx":0BB2
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3840
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   3540
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1650
         Left            =   120
         TabIndex        =   11
         Top             =   1050
         Width           =   3285
         Begin VB.CheckBox Check2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Incluye Consignaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox Text1 
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
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton Command2 
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
            Left            =   2880
            TabIndex        =   12
            Top             =   720
            Width           =   175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito Expedición:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   375
            TabIndex        =   16
            Top             =   800
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   15
            Top             =   1155
            Width           =   2850
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   8
         Top             =   350
         Width           =   2535
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
            Left            =   1930
            Picture         =   "CFGSALSTO.frx":2AD4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Tabla de Sectores"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command25 
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
            Picture         =   "CFGSALSTO.frx":2C1E
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Tabla de Depósitos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Armadores  /  Sectores"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   10
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CFGSALSTO.frx":2D68
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CFGSALSTO.frx":2F9C
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "CFGSALSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STOCKXX(32768, 1)
Dim ACUMULA_X_SECTOR#(10)

Sub CARGAGRILLA_STOCK_ARM()
        '
    'LEO LOS DATOS PARA EL ENCABEZADO
    INFO_GRILLA07.MSF.Rows = 1
    INFO_GRILLA07.MSF.Cols = 37  '/*/**
    '
    'CARGO ENCABEZADO
    '****************
    'LEO LA TABLA DE LOS ARMADORES
    '-----------------------------
    Dim DEPOTALLER$(12)
    ULTRE& = ULTREG&("DEPOARMA")
    If ULTRE& < 1 Then
       Call COMUNI("Falta Cargar Depósitos Taller")
       Exit Sub
    End If
    For i& = 1 To ULTRE&
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPOX% = Asc(Mid$(X$, 1, 1))
       DEPOTALLER$(i&) = ECOARCH$("TADEPOSI", Chr$(DEPOX%))
    Next i&
    '
    'ASIGNO TAMAÑO A LAS COLUMNAS
    For J% = 0 To 36
         NUCOLU% = J%: INFO_GRILLA07.TEPRUEBA = "123456789012": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width
    Next J%
    i& = 0
    '
    INFO_GRILLA07.MSF.Row = 0
    INFO_GRILLA07.MSF.Rows = 2
    For J% = 3 To 37 Step 3
         i& = i& + 1
         NUCOLU% = J%: INFO_GRILLA07.TEPRUEBA = "12345678901234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 1: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = DEPOTALLER$(i&)
    Next J%
    '
    Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HFFFFC0, 1)
    J% = 0
    INFO_GRILLA07.MSF.Rows = 3
    INFO_GRILLA07.MSF.Row = 2
    For i& = 0 To 36
       If i& = 0 Then NUCOLU% = 1:: GoTo 50
       
       INFO_GRILLA07.MSF.Rows = INFO_GRILLA07.MSF.Rows + 1
       J% = J% + 1
       Select Case J%
         Case 1: TITU$ = "STOCK"
         Case 2: TITU$ = "ENIVIADO"
         Case 3: TITU$ = "RECIBIDO"
       End Select
       If J% = 3 Then J% = 0
       
       NUCOLU% = i& + 1: INFO_GRILLA07.MSF.TextMatrix(1, NUCOLU% - 1) = TITU$
50  Next i&
'-----------------------------------------------------------------------------------------
'CARGA DETALLE LA GRILLA
'***************
    
    SQLX$ = "SELECT * FROM STOCK_ARMADOR WITH(NOLOCK) "
    SQLX$ = SQLX$ + " ORDER BY NUORDEN"
    INFO_GRILLA07.MSF.Cols = 37
    REG& = INFO_GRILLA07.MSF.Row

    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          REG& = REG& + 1
          INFO_GRILLA07.MSF.Row = INFO_GRILLA07.MSF.Row + 1
          Call TRACE(24, REG&, FIN&)
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 0) = FECHATEX$(CVINT(!FECHA))
          INFO_GRILLA07.MSF.ColAlignment(0) = 1
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 1) = FORMATNUM$(!Stock_1, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(1) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 2) = FORMATNUM$(!Enviado_1, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(2) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 3) = FORMATNUM$(!Recibido_1, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(3) = 6

          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 4) = FORMATNUM$(!Stock_2, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(4) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 5) = FORMATNUM$(!Enviado_2, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(5) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 6) = FORMATNUM$(!Recibido_2, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(6) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 7) = FORMATNUM$(!Stock_3, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(7) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 8) = FORMATNUM$(!Enviado_3, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(8) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 9) = FORMATNUM$(!Recibido_1, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(9) = 6
         '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 10) = FORMATNUM$(!Stock_4, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(10) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 11) = FORMATNUM$(!Enviado_4, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(11) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 12) = FORMATNUM$(!Recibido_4, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(12) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 13) = FORMATNUM$(!Stock_5, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(13) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 14) = FORMATNUM$(!Enviado_5, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(14) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 15) = FORMATNUM$(!Recibido_5, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(15) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 16) = FORMATNUM$(!Stock_6, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(16) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 17) = FORMATNUM$(!Enviado_6, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(17) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 18) = FORMATNUM$(!Recibido_6, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(18) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 19) = FORMATNUM$(!Stock_7, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(19) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 20) = FORMATNUM$(!Enviado_7, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(20) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 21) = FORMATNUM$(!Recibido_7, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(21) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 22) = FORMATNUM$(!Stock_8, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(22) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 23) = FORMATNUM$(!Enviado_8, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(23) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 24) = FORMATNUM$(!Recibido_8, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(24) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 25) = FORMATNUM$(!Stock_9, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(25) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 26) = FORMATNUM$(!Enviado_9, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(26) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 27) = FORMATNUM$(!Recibido_9, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(27) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 28) = FORMATNUM$(!Stock_10, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(28) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 29) = FORMATNUM$(!Enviado_10, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(29) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 30) = FORMATNUM$(!Recibido_10, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(30) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 31) = FORMATNUM$(!Stock_11, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(31) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 32) = FORMATNUM$(!Enviado_11, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(32) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 33) = FORMATNUM$(!Recibido_11, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(33) = 6
          '
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 34) = FORMATNUM$(!Stock_12, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(34) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 35) = FORMATNUM$(!Enviado_12, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(35) = 6
          INFO_GRILLA07.MSF.TextMatrix(REG& - 1, 36) = FORMATNUM$(!Recibido_12, "F14.2")
          INFO_GRILLA07.MSF.ColAlignment(36) = 6
          .MoveNext
       Loop
    End With
    INFOTMP.Close
    Set INFOTMP = Nothing
    '
    INFO_GRILLA07.Label1 = "Cálculo Realizado El Dìa: " & TRIM$(Control("CFGSALSTO", "RECAVALO"))
    INFO_GRILLA07.MSF.BackColor = &HFFFFFF
    INFO_GRILLA07.Show 1
    
    
End Sub




Function STOCK_ARM_RECIBIDO(DEPO%, J%)

     Condi$ = "DEPOMOVI = " & DEPO%
     Condi$ = Condi$ + " AND FECHMOV = '" & FETEXCOR1$(J%) & "'"
     STOCK_ARM_RECIBIDO = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", Condi$, "NOMESS"))
     
End Function

Function STOCK_ARM_ENVIADO(DEPO%, J%)

     Condi$ = "DEPOMOVI = " & DEPO%
     Condi$ = Condi$ + " AND FECHMOV = '" & FETEXCOR1$(J%) & "'"
     STOCK_ARM_ENVIADO = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", Condi$, "NOMESS"))
     
End Function

'

Sub STOCK_INFO_SECCION(DIAX%)

    'CRITERIO -///////////********************************
    'RECORRO LA TABLA INFOFAB
    'TOMO EL IMPORTE BUSCANDO EL PRECIO DE LA LISTA VIGENTE
    'Y ACUMULO EL IMPORTE DE TODAS LOS ARTICULOS AGRUPADOS POR SECCION
    'CON EL NUMERO DE LA MAQUINA OBTENGO LA SECCION Y SI MACHEA CON LAS CARGADAS EN LA TABLA DE SECCIONES DE ALEJPA
    'EL VECTOR ACUMULA_X_SECTOR() ESTA DECLARADO A NIVEL DE FRM
    
    DESDEX = FETEXCOR1$(DIAX%)
    HASTAX = FETEXCOR1$(DIAX%, "HASTA")
    '
    Condi$ = "FEHORLIB >= '" & DESDEX & "' AND FEHORLIB <= '" & HASTAX & "'"
    Condi$ = Condi$ + " AND COD_INTE > 0  AND RETRABA <= 0 "
    FIN& = CantRegistros&("INFOFAB", Condi$, "NOMESS")
    '
    For i% = 1 To 10
      ACUMULA_X_SECTOR#(i%) = 0
    Next i%
    '
    SQLX$ = "SELECT SUM(CANTREAL) AS INFORMADA, COD_INTE, MAQUINA FROM  INFOFAB WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,MAQUINA"
    '
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          U& = U& + 1
          Call TRACE(24, U&, FIN&)
          CIXI% = XVALO(!cod_inte)
          AAAAAAA = CODEXT$(CIXI%)
          MAQUI$ = TRIM$(SAFETEXT(!MAQUINA))
          MAQ = SECPERTE(MAQUI$)
          '
          CANTISECT = XVALO(!INFORMADA)
          IMPORTESEC = CANTISECT * PRELISTACT#(CIXI%, TIMONE%)
          If TIMONE% < 1 Then TIMONE% = 1
          IMPORTESEC = IMPORTESEC * TICAMONE(TIMONE%)
          Select Case NUMESECT(MAQ)
             Case Is = 1
                 ACUMULA_X_SECTOR#(1) = ACUMULA_X_SECTOR#(1) + IMPORTESEC
             Case Is = 2
                 ACUMULA_X_SECTOR#(2) = ACUMULA_X_SECTOR#(2) + IMPORTESEC
             Case Is = 3
                 ACUMULA_X_SECTOR#(3) = ACUMULA_X_SECTOR#(3) + IMPORTESEC
             Case Is = 4
                 ACUMULA_X_SECTOR#(4) = ACUMULA_X_SECTOR#(4) + IMPORTESEC
             Case Is = 5
                 ACUMULA_X_SECTOR#(5) = ACUMULA_X_SECTOR#(5) + IMPORTESEC
             Case Is = 6
                 ACUMULA_X_SECTOR#(6) = ACUMULA_X_SECTOR#(6) + IMPORTESEC
             Case Is = 7
                 ACUMULA_X_SECTOR#(7) = ACUMULA_X_SECTOR#(7) + IMPORTESEC
             Case Is = 8
                 ACUMULA_X_SECTOR#(8) = ACUMULA_X_SECTOR#(8) + IMPORTESEC
             Case Is = 9
                 ACUMULA_X_SECTOR#(9) = ACUMULA_X_SECTOR#(9) + IMPORTESEC
             Case Is = 10
                 ACUMULA_X_SECTOR#(10) = ACUMULA_X_SECTOR#(10) + IMPORTESEC
          End Select
          '
          .MoveNext
       Loop
    End With
    '
End Sub

Private Sub AGenRep_Click()
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Unload Me
   Command1.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    '
    Call BLANARCH("!STOKSECC")
    '
    Call TRALOCAL("STOKSECC", "")

30  VTB% = VENTABU%("STOKSECC")
    If VTB% < 0 Then GoTo 99
    '
    If ULTREG("!STOKSECC") > 10 Then
       Call COMUNI("Super Cantidad Máxima de Registros (10)")
       GoTo 30
    End If
    '
    Call TRACENTRAL("STOKSECC", "")
    Call CIERRARCH("STOKSECC")
    '
99  Command13.Enabled = True

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text1.TEXT = VDEV$
    End If
    Command2.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call FINAL("?ANASTOK")
   Command23.Enabled = True

End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    '
     OPX1% = COMALTER%("Opciones de Reporte\\Listado Valorizado de Stocks (Pesos) " & "\Listado de Stocks Talleres(Cantidades)")
    '
    If OPX1% = 1 Then
        'MENSAJE DE SELECCION CALCULAR O PRESENTAR REPORTE
        '-------------------------------------------------
        FECHCALCU1$ = TRIM$(Control("CFGSALSTO", "RECAVALO"))
        If FECHCALCU1$ <> "" Then FECHCALCU$ = " (Último Cálculo: " + FECHCALCU1$ + ")"
        OPX% = COMALTER%("Opciones de Reporte\\Ver Reporte" + FECHCALCU$ & "\Recalcular")
        '-********************************************************************************
        
        'VALIDACION DE OPCION
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        '-********************************************************************************
        '
        'EJECUCION DE OPCION Y GRABACION DE CONROL ULTIMO CALCULO
        '-------------------------------------------------
        If OPX% = 1 Then
           Call CARGAGRILLA
        ElseIf OPX% = 2 Then
           Call RECALCUVALO
        End If
        '-********************************************************************************
        '
    ElseIf OPX1% = 2 Then
        'MENSAJE DE SELECCION CALCULAR O PRESENTAR REPORTE
'        Call COMUNI("Aun No Disponible")
'        GoTo 99
        '-------------------------------------------------
        FECHCALCU1$ = TRIM$(Control("CFGSALSTO", "RESTKTAL"))
        If FECHCALCU1$ <> "" Then FECHCAL$ = " (Último Cálculo: " + FECHCALCU1$ + ")"
        OPX2% = COMALTER%("Opciones de Reporte\\Ver Reporte" + FECHCAL$ & "\Recalcular")
        '*********************************************************************************
        'VALIDACION DE OPCION
        If OPX2% < 1 Or OPX2% > 2 Then GoTo 99
        '-********************************************************************************
        '
        'EJECUCION DE OPCION Y GRABACION DE CONROL ULTIMO CALCULO
        '-------------------------------------------------
        If OPX2% = 1 Then
           Call CARGAGRILLA_STOCK_ARM
        ElseIf OPX2% = 2 Then
           Call STOCKS_TALLERES
        End If
        
    End If
99  Command21.Enabled = True

End Sub
Sub RECALCUVALO(Optional MODO$)
    '
    'MUESTRA SELECCION DE FECHAS
    '--------------------------
    If MODO$ = "AUTO" Then
      Has% = HOY(HO$)
      RANGOX% = XVALO(Control("SITUASTO", "RANGPROC"))
      If RANGOX% > 0 Then
          DES% = Has% - (RANGOX% - 1)
      Else
          DES% = Has% - 10 ' POR DEFECTO
      End If
      GoTo 20
    End If
    
10  Call DESHASFECHA(DES%, Has%, PERIO$)
    If PERIO$ = "" Then
        Exit Sub
    End If
20  DESDE = DES%: HASTA = Has%
    Call SETULTREG("STVALFEC", 0)
    '*******************************************************************
    '
    'VALIDO CANTIDAD DE DIAS
    '--------------------------
    CANDIA% = 1 + DIENTRE(DESDE, HASTA)
    If CANDIA% > 31 Then
       Call COMUNI("Rango Seleccionado\Supera el Máximo de Días Permitido (31)")
       GoTo 10
       Exit Sub
    End If
    '*******************************************************************
    '
    Screen.MousePointer = 11
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        h% = i%
        FEC$ = FECHATEX$(h%)
        FEC$ = AJUSTD(FEC$, 12)
        Call REPLA(TX1$, FEC$, PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Rango de Fechas", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 1)

    'TOTAL VALORIZADO  REMITIDO
    '--------------------------
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        FCREM# = FACTURADO_X_REM(i%)
        Call REPLA(TX1$, FORMATNUM$(FCREM#, "F12.2"), PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Facturado X Rem. Total", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 2)
    '*******************************************************************
    '
    'TOTAL ENVIADO A ARMADORES
    '-------------------------
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        TTENV_AR# = TTENVIADO_ARMADOR(i%)
        Call REPLA(TX1$, FORMATNUM$(TTENV_AR#, "F12.2"), PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Total Enviado Armadores", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 3)
    '*******************************************************************
    '
    'TOTAL ENTREGADO A ARMADORES
    '--------------------------
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        TTENTR_ARM# = TTEENTREGADO_DEL_ARMADOR(i%)
        Call REPLA(TX1$, FORMATNUM$(TTENTR_ARM#, "F12.2"), PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Total Ent. Armadores", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 4)
    '*******************************************************************
    '
    'TOTAL DEPOSITO Armadores
    '--------------------------
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        STARM# = STOCK_TOT_ARMA#(i%, Has%)
        Call REPLA(TX1$, FORMATNUM$(STARM#, "F12.2"), PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Total Dep. Armadores", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 5)
     '*******************************************************************
    '
    'TOTAL DEPOSITO EXPEDICION
    '--------------------------
    PINIX% = 24: TX1$ = ""
    For i% = DES% To Has%
        STEXPE# = STOCK_EXPEDICION(i%, Has%)
        Call REPLA(TX1$, FORMATNUM$(STEXPE#, "F12.2"), PINIX%, 12)
        PINIX% = PINIX% + 12
    Next i%
    Call REPLA(TX1$, "Total Dep. Expedición", 1, 23)
    Call GRAREG("STVALFEC", TX1$, 6)
   
    '*******************************************************************
    'STOCK INFORMADOS POR SECCIONES
    '--------------------------
    Dim TITULOX$(10)
    For J& = 1 To 10
       TITULOX$(J&) = ""
    Next J&
    
    PINIX% = 24: TX1$ = "": J& = 0
    'CARGO LOS TITULOS DE LAS SECCIONES
    For J& = 1 To 10
       If J& <= ULTREG("STOKSECC") Then
          X$ = REGLEIDO$("STOKSECC", J&)
          COD = CVS(Mid$(X$, 1, 4))
          If COD > 0 Then TITULOX$(J&) = TRIM$(ECOARCH("SECPROD", COD))
       End If
    Next J&

    ' VECTOR 'ACUMULA_X_SECTOR#' DECLARADO A NIVEL DE FORMULARIO
    For i% = DES% To Has%
       Call STOCK_INFO_SECCION(i%)
       Call REPLA(TX1$, FORMATNUM$(ACUMULA_X_SECTOR#(1), "F12.2"), PINIX%, 12)
       Call REPLA(TX2$, FORMATNUM$(ACUMULA_X_SECTOR#(2), "F12.2"), PINIX%, 12)
       Call REPLA(TX3$, FORMATNUM$(ACUMULA_X_SECTOR#(3), "F12.2"), PINIX%, 12)
       Call REPLA(TX4$, FORMATNUM$(ACUMULA_X_SECTOR#(4), "F12.2"), PINIX%, 12)
       Call REPLA(TX5$, FORMATNUM$(ACUMULA_X_SECTOR#(5), "F12.2"), PINIX%, 12)
       Call REPLA(TX6$, FORMATNUM$(ACUMULA_X_SECTOR#(6), "F12.2"), PINIX%, 12)
       Call REPLA(TX7$, FORMATNUM$(ACUMULA_X_SECTOR#(7), "F12.2"), PINIX%, 12)
       Call REPLA(TX8$, FORMATNUM$(ACUMULA_X_SECTOR#(8), "F12.2"), PINIX%, 12)
       Call REPLA(TX9$, FORMATNUM$(ACUMULA_X_SECTOR#(9), "F12.2"), PINIX%, 12)
       Call REPLA(TX10$, FORMATNUM$(ACUMULA_X_SECTOR#(10), "F12.2"), PINIX%, 12)
       PINIX% = PINIX% + 12
    Next i%
    '
    'LE AGREGO EL TITULO
    Call REPLA(TX1$, TITULOX$(1), 1, 23)
    Call REPLA(TX2$, TITULOX$(2), 1, 23)
    Call REPLA(TX3$, TITULOX$(3), 1, 23)
    Call REPLA(TX4$, TITULOX$(4), 1, 23)
    Call REPLA(TX5$, TITULOX$(5), 1, 23)
    Call REPLA(TX6$, TITULOX$(6), 1, 23)
    Call REPLA(TX7$, TITULOX$(7), 1, 23)
    Call REPLA(TX8$, TITULOX$(8), 1, 23)
    Call REPLA(TX9$, TITULOX$(9), 1, 23)
    Call REPLA(TX10$, TITULOX$(10), 1, 23)
    '
    'GRABO EN ARCHVO
    Call GRAREG("STVALFEC", TX1$, 7)
    Call GRAREG("STVALFEC", TX2$, 8)
    Call GRAREG("STVALFEC", TX3$, 9)
    Call GRAREG("STVALFEC", TX4$, 10)
    Call GRAREG("STVALFEC", TX5$, 11)
    Call GRAREG("STVALFEC", TX6$, 12)
    Call GRAREG("STVALFEC", TX7$, 13)
    Call GRAREG("STVALFEC", TX8$, 14)
    Call GRAREG("STVALFEC", TX9$, 15)
    Call GRAREG("STVALFEC", TX10$, 16)
    '
    Call SETULTREG("STVALFEC", 16)
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    If MODO$ <> "AUTO" Then Call COMUNI("Proceso de Cálculo Finalizado, Puede Emitir el Reporte")
    '
    
    FEC$ = FECHATEX$(HOY(HO$))
    Call GRACONTROL("CFGSALSTO", "RECAVALO", FEC$)
    '
End Sub
Sub STOCKS_TALLERES(Optional MODO$)

    'MUESTRA SELECCION DE FECHAS
    '--------------------------
    If MODO$ = "AUTO" Then
      Has% = HOY(HO$)
      RANGOX% = XVALO(Control("SITUASTO", "RANGPROC"))
      If RANGOX% > 0 Then
          DES% = Has% - (RANGOX% - 1)
      Else
          DES% = Has% - 10 ' POR DEFECTO
      End If
      GoTo 20
    End If


10  Call DESHASFECHA(DES%, Has%, PERIO$)
    If PERIO$ = "" Then
        Exit Sub
    End If
20  DESDE = DES%: HASTA = Has%
   '
   'VALIDO CANTIDAD DE DIAS
   '--------------------------
   CANDIA% = 1 + DIENTRE(DESDE, HASTA)
   If CANDIA% > 31 Then
       Call COMUNI("Rango Seleccionado\Supera el Máximo de Días Permitido (31)")
       GoTo 10
       Exit Sub
    End If

   'CARGO LOS DEPOSITOS DE ARMADORES EN UN VECTOR
   '----------------------------------------------
    Dim DEPOX%(64)
    ULTRE& = ULTREG&("DEPOARMA")
    If ULTRE& < 1 Then
       Call COMUNI("Falta Cargar Depósitos Taller")
       Exit Sub
    End If

    For i& = 1 To ULTRE&
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPOX%(i&) = Asc(Mid$(X$, 1, 1))
    Next i&
   '----------------------------------------------
   '
   'BORRO LOS DATOS DE LA TABLA
   '----------------------------------------------
    Call BORRAREGISTROS("STOCK_ARMADOR", "NUORDEN > 0", REGAF&, "NOMESS")
   '----------------------------------------------
   '
   'AGREGO LA INFORMACION A LA TABLA
   '----------------------------------------------
   
   SQLX$ = "SELECT * FROM STOCK_ARMADOR WHERE NUORDEN < 1 "
   Dim STKMP As ADODB.Recordset
   Set STKMP = New ADODB.Recordset
25 On Error Resume Next
   STKMP.Open FILTSQL(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With STKMP
   Screen.MousePointer = 11
   For J% = DES% To Has%
              k% = k% + 1
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              '
              !FECHA = FECHATEX$(J%)
              !NUORDEN = k%
              '
              !Stock_1 = STOCK_ARM(DEPOX%(1), J%)
              !Enviado_1 = STOCK_ARM_ENVIADO(DEPOX%(1), J%)
              !Recibido_1 = STOCK_ARM_RECIBIDO(DEPOX%(1), J%)
              '
              If ULTRE& < 2 Then GoTo 80 ' NEXT
              !Stock_2 = STOCK_ARM(DEPOX%(2), J%)
              !Enviado_2 = STOCK_ARM_ENVIADO(DEPOX%(2), J%)
              !Recibido_2 = STOCK_ARM_RECIBIDO(DEPOX%(2), J%)
              '
              If ULTRE& < 3 Then GoTo 80 ' NEXT
              !Stock_3 = STOCK_ARM(DEPOX%(3), J%)
              !Enviado_3 = STOCK_ARM_ENVIADO(DEPOX%(3), J%)
              !Recibido_3 = STOCK_ARM_RECIBIDO(DEPOX%(3), J%)
              '
              If ULTRE& < 4 Then GoTo 80 ' NEXT
              !Stock_4 = STOCK_ARM(DEPOX%(4), J%)
              !Enviado_4 = STOCK_ARM_ENVIADO(DEPOX%(4), J%)
              !Recibido_4 = STOCK_ARM_RECIBIDO(DEPOX%(4), J%)
              '
              If ULTRE& < 5 Then GoTo 80 ' NEXT
              !Stock_5 = STOCK_ARM(DEPOX%(5), J%)
              !Enviado_5 = STOCK_ARM_ENVIADO(DEPOX%(5), J%)
              !Recibido_5 = STOCK_ARM_RECIBIDO(DEPOX%(5), J%)
              '
              If ULTRE& < 6 Then GoTo 80 ' NEXT
              !Stock_6 = STOCK_ARM(DEPOX%(6), J%)
              !Enviado_6 = STOCK_ARM_ENVIADO(DEPOX%(6), J%)
              !Recibido_6 = STOCK_ARM_RECIBIDO(DEPOX%(6), J%)
              '
              If ULTRE& < 7 Then GoTo 80 ' NEXT
              !Stock_7 = STOCK_ARM(DEPOX%(7), J%)
              !Enviado_7 = STOCK_ARM_ENVIADO(DEPOX%(7), J%)
              !Recibido_7 = STOCK_ARM_RECIBIDO(DEPOX%(7), J%)
              '
              If ULTRE& < 8 Then GoTo 80 ' NEXT
              !Stock_8 = STOCK_ARM(DEPOX%(8), J%)
              !Enviado_8 = STOCK_ARM_ENVIADO(DEPOX%(8), J%)
              !Recibido_8 = STOCK_ARM_RECIBIDO(DEPOX%(8), J%)
              '
              If ULTRE& < 9 Then GoTo 80 ' NEXT
              !Stock_9 = STOCK_ARM(DEPOX%(9), J%)
              !Enviado_9 = STOCK_ARM_ENVIADO(DEPOX%(9), J%)
              !Recibido_9 = STOCK_ARM_RECIBIDO(DEPOX%(9), J%)
              '
              If ULTRE& < 10 Then GoTo 80 ' NEXT
              !Stock_10 = STOCK_ARM(DEPOX%(10), J%)
              !Enviado_10 = STOCK_ARM_ENVIADO(DEPOX%(10), J%)
              !Recibido_10 = STOCK_ARM_RECIBIDO(DEPOX%(10), J%)

              If ULTRE& < 11 Then GoTo 80 ' NEXT
              !Stock_11 = STOCK_ARM(DEPOX%(11), J%)
              !Enviado_11 = STOCK_ARM_ENVIADO(DEPOX%(11), J%)
              !Recibido_11 = STOCK_ARM_RECIBIDO(DEPOX%(11), J%)
              
              If ULTRE& < 12 Then GoTo 80 ' NEXT
              !Stock_12 = STOCK_ARM(DEPOX%(12), J%)
              !Enviado_12 = STOCK_ARM_ENVIADO(DEPOX%(12), J%)
              !Recibido_12 = STOCK_ARM_RECIBIDO(DEPOX%(12), J%)
80            .Update

   Next J%
   Set STKMP = Nothing
   .Close
   End With
   Screen.MousePointer = 1
   If MODO$ <> "AUTO" Then Call COMUNI("Proceso de Cálculo Finalizado, Puede Emitir el Reporte")
   
   FEC$ = FECHATEX$(HOY(HO$))
   Call GRACONTROL("CFGSALSTO", "RESTKTAL", FEC$)

End Sub

Function STOCK_ARM(DEPO%, DIAX%)
    '
    'CRITERIO
    '--------
    'GENERO CONSULTA QUE DEVUELVE DE LOS DEPOSITOS DE ARMADORES
    'DEVUELVE EL STOCKS CONSIDERADO A FECHA.
    FECHAX = FETEXCOR1$(DIAX%)
    '
    Condi$ = "FECHMOV <= '" & FECHAX & "'" & " And DEPOMOVI = " & DEPO%
    '
    STCKACMLDO = 0
    SQLX$ = "SELECT SUM(CANTSALID) AS CASALIDA,SUM(CANTINGRE)AS CAINGRESADA FROM  MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + " GROUP BY FECHMOV , DEPOMOVI"
    '
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
      Do While Not .EOF
        i& = i& + 1
        Call TRACE(24, i&, FIN&)
        CANTIX = XVALO(!CAINGRESADA) - XVALO(!CASALIDA)
        STCKACMLDO = STCKACMLDO + (CANTIX)
        .MoveNext
      Loop
    End With
    '
    INFOTMP.Close
    Set INFOTMP = Nothing
    STOCK_ARM = STCKACMLDO
    '
End Function
Function FACTURADO_X_REM(DIAX%)
    'SE CONSULTA LOS MOVIMIENTOS CON RT TOMANDO PARA CODIGO EL TOTAL DE LA CANTIDAD REMITIDA
    'SE OBTIENE EL VALOR DE LA LISTA DE PRECIOS VIGENTE Y SE MULTIPICA POR LA CANTIDAD DE CADA ARTICULO
    'SE ACUMLA LA MISMA  Y DEVUELVE EL RESULTADO PARA EL DIA PASADO COMO PARAMETRO
    
    FECHAX = FETEXCOR1$(DIAX%)
    '
    Condi$ = "CANTSALID > 0 AND CODIMOVI = 'RT' AND FECHMOV = '" & FECHAX & "' AND COD_INTE > 0"
    '
    FACTURADO_X_REM = 0
    ACUMULADO = 0
    '
    SQLX$ = "SELECT SUM(CANTSALID) AS CANTI, VALOUNIT, TIPOCAM, COD_INTE  FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, VALOUNIT, TIPOCAM"
    
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          i& = i& + 1
          Call TRACE(24, i&, FIN&)
          CANTIX = XVALO(!CANTI)
          CIXI% = XVALO(!cod_inte)
          If CIXI% > 0 Then
             TICAM# = XVALO(!TIPOCAM): If TICAM# < 1 Then TICAM# = 1
             IMPORUNI# = XVALO(!VALOUNIT) * TICAM#
             ACUMULADO = ACUMULADO + (CANTIX * IMPORUNI#)
             .MoveNext
          End If
       Loop
    End With

    FACTURADO_X_REM = ACUMULADO
    '
End Function
Function TTENVIADO_ARMADOR(DIAX%)
    
    'CRITERIO -///////////********************************
    'TOMA LA CANTIDAD INGRESADA DE LOS DEPOSITOS DE ARMADORES
    'ACUMULADO POR CODIGO, VALOR Y TIPO DE CAMBIO Y LUEGO VALORIZA
    'LA AGRUPACION SE REALIZA POR ESTOS 3 PARA PODER CALCULAR LA CANTIDAD VALORIZADA POR EL VALOR
    'YA QUE EXISTE LA POSIBILIDAD QUE DURANTE EL MISMO DIA CAMBIE EL VALOR DE LA MONEDA O EL PRECIO DE LISTA
    'Y SE GUARDEN MAS DE UNA TRANFERENCIA PARA UN MISMO ITEM CON DIFERENTES VALORES.

    FECHAX = FETEXCOR1$(DIAX%)
    '
    Condi$ = "CANTINGRE > 0 AND CODIMOVI = 'TR' AND FECHMOV = '" & FECHAX & "'"
    Condi$ = Condi$ + " AND COD_INTE > 0 AND TIPOTRAN <= 0"
    FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
    '
    STOCK_ACUMULADO = 0
    
    For i& = 1 To ULTREG&("DEPOARMA")
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPO% = Asc(Mid$(X$, 1, 1))
       If i& = 1 Then
          sqlx1$ = " AND (DEPOMOVI = " & DEPO%
       Else
          sqlx1$ = sqlx1$ + " OR DEPOMOVI = " & DEPO%
       End If
    Next i&
    If i& > 1 Then sqlx1$ = sqlx1$ + ")"

    '
    SQLX$ = "SELECT SUM(CANTINGRE) AS CANTI ,COD_INTE, VALOUNIT ,TIPOCAM FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + sqlx1$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, VALOUNIT, TIPOCAM"

    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          i& = i& + 1
          Call TRACE(24, i&, FIN&)
          CANTIX = XVALO(!CANTI)
          CIXI% = XVALO(cod_inte)
          TICAM# = XVALO(!TIPOCAM): If TICAM# < 1 Then TICAM# = 1
          IMPORUNI# = XVALO(!VALOUNIT) * TICAM#
          ACUMULADO = ACUMULADO + (CANTIX * IMPORUNI#)
          .MoveNext

       Loop
    End With
    TTENVIADO_ARMADOR = ACUMULADO
End Function
Function TTEENTREGADO_DEL_ARMADOR(DIAX%)
    'CRITERIO -///////////********************************
    'TOMA LA CANTIDAD SALIDA DE LOS DEPOSITOS DE ARMADORES
    'ACUMULADO POR CODIGO, VALOR Y TIPO DE CAMBIO Y LUEGO VALORIZA
    'LA AGRUPACION SE REALIZA POR ESTOS 3 PARA PODER CALCULAR LA CANTIDAD VALORIZADA POR EL VALOR
    'YA QUE EXISTE LA POSIBILIDAD QUE DURANTE EL MISMO DIA CAMBIE EL VALOR DE LA MONEDA O EL PRECIO DE LISTA
    'Y SE GUARDEN MAS DE UNA TRANFERENCIA PARA UN MISMO ITEM CON DIFERENTES VALORES.
    
    FECHAX = FETEXCOR1$(DIAX%)
    '
    Condi$ = "CANTSALID > 0 AND CODIMOVI = 'TR' AND FECHMOV = '" & FECHAX & "'"
    Condi$ = Condi$ + " AND COD_INTE > 0 AND TIPOTRAN <= 0"
    FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
    '
    STOCK_ACUMULADO = 0
    
    For i& = 1 To ULTREG&("DEPOARMA")
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPO% = Asc(Mid$(X$, 1, 1))
       If i& = 1 Then
          sqlx1$ = " AND (DEPOMOVI = " & DEPO%
       Else
          sqlx1$ = sqlx1$ + " OR DEPOMOVI = " & DEPO%
       End If
    Next i&
    '
    If i& > 1 Then sqlx1$ = sqlx1$ + ")"
    '
    SQLX$ = "SELECT SUM(CANTSALID) AS CANTI,COD_INTE, VALOUNIT, TIPOCAM  FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + sqlx1$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, VALOUNIT, TIPOCAM"
    
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          i& = i& + 1
          Call TRACE(24, i&, FIN&)
          CANTIX = XVALO(!CANTI)
          CIXI% = XVALO(!cod_inte)
          If CIXI% > 0 Then
             TICAM# = XVALO(!TIPOCAM): If TICAM# < 1 Then TICAM# = 1
             IMPORUNI# = XVALO(!VALOUNIT) * TICAM#
             ACUMULADO = ACUMULADO + (CANTIX * IMPORUNI#)
             .MoveNext
          End If
       Loop
    End With
    TTEENTREGADO_DEL_ARMADOR = ACUMULADO
End Function
Function STOCK_TOT_ARMA#(DIAX%, HASTAX%)
    'CRITERIO -///////////********************************
    'RECORRO EL STOCK EXISTENTE AGRUPADO POR ARTICULO
    'TOMO EL IMPORTE BUSCANDO EL COSTO MAS PROXIMO A LA FECHA DEL MOVIMIENTO
    'Y ACUMULO EL IMPORTE DE TODAS LOS ARTICULOS.
    '
    'PRIMERA VEZ QUE INGRESA VALORIZA POR LISTA VIGENTE TODO EL STOCKS
    'CONSIDERANDO QUE LA FECHA SEA <= A LA FECHA DESDE
    'DESPUES DE LA PRIMERA VEZ CALCULA EL STOCK CORRESPONDIENTE AL DIA PARAMETRIZADO
    'Y LO ACUMALA AL ANTERIOR , ES PARA QUE NO CALCULE TODO EL STOCKS CADA VEZ QUE INGRESA
    'POR QUE SE MULTIPLICA DEMORA INNECESARIA.
    
    Static ENTRADAS%, STCKACMLDO_TDS_DIAS
    FECHAX = FETEXCOR1$(DIAX%)
        '
    For i& = 1 To ULTREG&("DEPOARMA")
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPO% = Asc(Mid$(X$, 1, 1))
       If i& = 1 Then
          sqlx1$ = " AND (DEPOMOVI = " & DEPO%
       Else
          sqlx1$ = sqlx1$ + " OR DEPOMOVI = " & DEPO%
       End If
    Next i&
    '
    If i& > 1 Then sqlx1$ = sqlx1$ + ")"
    
    'PRIMER ENTRADA
    If ENTRADAS% < 1 Then
        ENTRADAS% = 1
        Condi$ = " FECHMOV <= '" & FECHAX & "'" & sqlx1$
        FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
        '
        STCKACMLDO_TDS_DIAS = 0
        '
        SQLX$ = "SELECT COD_INTE ,SUM(CANTSALID) AS CASALIDA,SUM(CANTINGRE) AS CAINGRESADA FROM  MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & Condi$
        SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
        
        Dim INFOTMP As ADODB.Recordset
        Set INFOTMP = READSET(SQLX$)
        With INFOTMP
           Do While Not .EOF
              i& = i& + 1
              Call TRACE(24, i&, FIN&)
              CANTIX = XVALO(!CAINGRESADA) - XVALO(!CASALIDA)
              CIXI% = XVALO(!cod_inte)
              '
              'SE MODIFICA PARA QUE COINCIDA CON EL REPORTE DE STOCKS VALORIZADO POR LISTA DE PRECIOS
              IMPORUNI# = PRELISTACT#(CIXI%, TIMONE%, 1)
              If TIMONE% < 1 Then TIMONE% = 1
              IMPORUNI# = IMPORUNI# * TICAMONE(TIMONE%)
'              'OBTENGO EL PRECIO DEL MOVISTO CONSIDERANDO EL MOMVIMIENTO MAS CERCANO A LA FECHA CON PRECIO
'              CONDI$ = "VALOUNIT > 0 AND FECHMOV <= '" & FECHAX & "' AND COD_INTE = " & CIXI% & "  AND DEPOMOVI = " & DEPO% & " ORDER BY FECHMOV DESC"
'              IMPORUNI# = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
'              TICAM# = XVALO(VALOBADA("MOVISTO", "TIPOCAM", CONDI$, "NOMESS"))
'              If TICAM# < 1 Then TICAM# = 1
'              IMPORUNI# = IMPORUNI# * TICAM#
              '
              'CALCULO PARA ESTE CODIGO EL IMPORTE POR LA CANTIDAD
              STCKACMLDO_TDS_DIAS = STCKACMLDO_TDS_DIAS + (CANTIX * IMPORUNI#)
              '
              .MoveNext
           Loop
        End With
        INFOTMP.Close
        Set INFOTMP = Nothing

        '
    Else ' ENTRADAS SIGUIENTES
        Condi$ = "FECHMOV = '" & FECHAX & "'"
        FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
        '
        STOCK_ACUMULADO = 0
        '
        SQLX$ = "SELECT COD_INTE ,SUM(CANTSALID) AS CASALIDA,SUM(CANTINGRE)AS CAINGRESADA FROM  MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & Condi$
        SQLX$ = SQLX$ + sqlx1$
        SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
        
'        Dim INFOTMP1 As ADODB.Recordset
        Set INFOTMP1 = READSET(SQLX$)
        With INFOTMP1
           Do While Not .EOF
              i& = i& + 1
              Call TRACE(24, i&, FIN&)
              CANTIX = XVALO(!CAINGRESADA) - XVALO(!CASALIDA)
              CIXI% = XVALO(!cod_inte)
               'SE MODIFICA PARA QUE COINCIDA CON EL REPORTE DE STOCKS VALORIZADO POR LISTA DE PRECIOS
              IMPORUNI# = PRELISTACT#(CIXI%, TIMONE%, 1)
              If TIMONE% < 1 Then TIMONE% = 1
              IMPORUNI# = IMPORUNI# * TICAMONE(TIMONE%)
             '
'              'OBTENGO EL PRECIO DEL MOVISTO CONSIDERANDO EL MOMVIMIENTO MAS CERCANO A LA FECHA CON PRECIO
'              CONDI$ = "VALOUNIT > 0 AND FECHMOV <= '" & FECHAX & "' AND COD_INTE = " & CIXI% & "  AND DEPOMOVI = " & DEPO% & " ORDER BY FECHMOV DESC"
'              IMPORUNI# = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
'              TICAM# = XVALO(VALOBADA("MOVISTO", "TIPOCAM", CONDI$, "NOMESS"))
'              If TICAM# < 1 Then TICAM# = 1
'              IMPORUNI# = IMPORUNI# * TICAM#
              '
              'CALCULO PARA ESTE CODIGO EL IMPORTE POR LA CANTIDAD
              STCKACMLDO_TDS_DIAS = STCKACMLDO_TDS_DIAS + (CANTIX * IMPORUNI#)
              '
              .MoveNext
           Loop
        End With
        INFOTMP1.Close
        Set INFOTMP1 = Nothing

    End If
    '
    STOCK_TOT_ARMA# = STCKACMLDO_TDS_DIAS
    '
    'SI ES LA ULTIMA ENTRADA LIMPIA LAS VARIABLES ESTATICAS
    If DIAX% = HASTAX% Then STCKACMLDO_TDS_DIAS = 0: ENTRADAS% = 0

End Function
Function STOCK_EXPEDICION(DIAX%, HASTAX%)
    'CRITERIO -///////////********************************
    'RECORRO EL STOCK EXISTENTE AGRUPADO POR ARTICULO
    'TOMO EL IMPORTE BUSCANDO EL COSTO MAS PROXIMO A LA FECHA DEL MOVIMIENTO
    'Y ACUMULO EL IMPORTE DE TODAS LOS ARTICULOS.
    '
    'PRIMERA VEZ QUE INGRESA VALORIZA POR LISTA VIGENTE TODO EL STOCKS
    'CONSIDERANDO QUE LA FECHA SEA <= A LA FECHA DESDE
    'DESPUES DE LA PRIMERA VEZ CALCULA EL STOCK CORRESPONDIENTE AL DIA PARAMETRIZADO
    'Y LO ACUMALA AL ANTERIOR , ES PARA QUE NO CALCULE TODO EL STOCKS CADA VEZ QUE INGRESA
    'POR QUE SE MULTIPLICA DEMORA INNECESARIA.
    
    Static ENTRADAS%, STCKACMLDO_TDS_DIAS
    FECHAX = FETEXCOR1$(DIAX%)
        '
    DEPO% = XVALO(Control("ANALSTOK", "DEPOEXPE"))
    'PRIMER ENTRADA
    If ENTRADAS% < 1 Then
        ENTRADAS% = 1
        Condi$ = "FECHMOV <= '" & FECHAX & "'  AND DEPOMOVI = " & DEPO%
        FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
        '
        STCKACMLDO_TDS_DIAS = 0
        '
        SQLX$ = "SELECT COD_INTE ,SUM(CANTSALID) AS CASALIDA,SUM(CANTINGRE)AS CAINGRESADA FROM  MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & Condi$
        SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
        
        Dim INFOTMP As ADODB.Recordset
        Set INFOTMP = READSET(SQLX$)
        With INFOTMP
           Do While Not .EOF
              i& = i& + 1
              Call TRACE(24, i&, FIN&)
              CANTIX = XVALO(!CAINGRESADA) - XVALO(!CASALIDA)
              CIXI% = XVALO(!cod_inte)
              '
              'SE MODIFICA PARA QUE COINCIDA CON EL REPORTE DE STOCKS VALORIZADO POR LISTA DE PRECIOS
              IMPORUNI# = PRELISTACT#(CIXI%, TIMONE%, 1)
              If TIMONE% < 1 Then TIMONE% = 1
              IMPORUNI# = IMPORUNI# * TICAMONE(TIMONE%)
              '
              'CALCULO PARA ESTE CODIGO EL IMPORTE POR LA CANTIDAD
              STCKACMLDO_TDS_DIAS = STCKACMLDO_TDS_DIAS + (CANTIX * IMPORUNI#)
              '
              .MoveNext
           Loop
        End With
        INFOTMP.Close
        Set INFOTMP = Nothing

        '
    Else ' ENTRADAS SIGUIENTES
        Condi$ = "FECHMOV = '" & FECHAX & "'  AND DEPOMOVI = " & DEPO%
        FIN& = CantRegistros&("MOVISTO", Condi$, "NOMESS")
        '
        STOCK_ACUMULADO = 0
        '
        SQLX$ = "SELECT COD_INTE ,SUM(CANTSALID) AS CASALIDA,SUM(CANTINGRE)AS CAINGRESADA FROM  MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE " & Condi$
        SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
        
'        Dim INFOTMP1 As ADODB.Recordset
        Set INFOTMP1 = READSET(SQLX$)
        With INFOTMP1
           Do While Not .EOF
              i& = i& + 1
              Call TRACE(24, i&, FIN&)
              CANTIX = XVALO(!CAINGRESADA) - XVALO(!CASALIDA)
              CIXI% = XVALO(!cod_inte)
               'SE MODIFICA PARA QUE COINCIDA CON EL REPORTE DE STOCKS VALORIZADO POR LISTA DE PRECIOS
              IMPORUNI# = PRELISTACT#(CIXI%, TIMONE%, 1)
              If TIMONE% < 1 Then TIMONE% = 1
              IMPORUNI# = IMPORUNI# * TICAMONE(TIMONE%)
              '
              'CALCULO PARA ESTE CODIGO EL IMPORTE POR LA CANTIDAD
              STCKACMLDO_TDS_DIAS = STCKACMLDO_TDS_DIAS + (CANTIX * IMPORUNI#)
              '
              .MoveNext
           Loop
        End With
        INFOTMP1.Close
        Set INFOTMP1 = Nothing
    End If
    '
    STOCK_EXPEDICION = STCKACMLDO_TDS_DIAS
    '
    'SI ES LA ULTIMA ENTRADA LIMPIA LAS VARIABLES ESTATICAS
    If DIAX% = HASTAX% Then STCKACMLDO_TDS_DIAS = 0: ENTRADAS% = 0

End Function

Private Sub Command25_Click()
    Command25.Enabled = False
    '
    
    Call BLANARCH("!DEPOARMA")
    '
    Call TRALOCAL("DEPOARMA", "")

30  VTB% = VENTABU%("DEPOARMA")
    If VTB% < 0 Then GoTo 99
    '
    Call TRACENTRAL("DEPOARMA", "")
    Call CIERRARCH("DEPOARMA")
    '
99  Command25.Enabled = True
   
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    DEP% = XVALO(Text1)
    DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEP%))
    If DEP% > 0 And TRIM$(DEPOSI$) = "" Then
       Call COMUNI("Depósito Inexistente o no Válido")
       GoTo 99
    End If
    '
    OPCIONX$ = "NO"
    If Check2.Value = 1 Then OPCIONX$ = "SI"
    Call GRACONTROL("ANALSTOK", "INCLUCON", OPCIONX$)
    '
    Call GRACONTROL("ANALSTOK", "DEPOEXPE", TRIM$(Text1.TEXT))
99  Command4.Enabled = True
    '
    Call Command1_Click ' CIERRA LA VENTANA
    
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    '
    DEP% = XVALO(Text1)
    DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEP%))
    If TRIM$(DEPOSI$) = "" Then
       Call COMUNI("Falta Configurar Depósito de Consignación")
       GoTo 99
    End If
    '
    INCLUYECONSIGNACION% = 0
    If Control$("ANALSTOK", "INCLUCON") = "SI" Then
        INCLUYECONSIGNACION% = 1
        If ULTREG("STOKSECC") < 1 Then
           Call COMUNI("Falta Configurar Depósito de Terceros")
        End If
    End If
    '
    Screen.MousePointer = 11
    Call CREACAMPO("", "ANASTOK", "Cod_Inte", 3, 0, 1)
    Call CREACAMPO("", "ANASTOK", "Cod_Exte", 10, 24)
    Call CREACAMPO("", "ANASTOK", "Idsubor", 10, 12)
    Call CREACAMPO("", "ANASTOK", "NroOperacion", 3, 0)
    Call CREACAMPO("", "ANASTOK", "Sector1", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector2", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector3", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector4", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector5", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector6", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector7", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector8", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector9", 7, 0)
    Call CREACAMPO("", "ANASTOK", "Sector10", 7, 0)
    Call CREACAMPO("", "ANASTOK", "StkTerceros", 7, 0)
    Call CREACAMPO("", "ANASTOK", "StkExpedicion", 7, 0)
    '
    DEPO% = XVALO(Control("ANALSTOK", "DEPOEXPE"))

    Call CARGA_VECTOR_STOCKS(DEPO%)
    Call BORRAREGISTROS("ANASTOK", "(IDSUBOR <> '' Or IDSUBOR Is Null Or IDSUBOR = '')", REGAF&, "NOMESS")
    'CONDI$ = "ORDEFABR.STATUORF < 3"
    'FIN& = CantRegistros("OPERFAB WITH(NOLOCK) INNER JOIN OPERFAB ON INFOFAB.IDSUBOR = ORDEFABR.IDSUBOR", CONDI$, "NOMESS")
    
'    SQLX$ = "SELECT OPERFAB.IDSUBOR, OPERFAB.COD_INTE, OPERFAB.NUMEOPER, OPERFAB.CANTPROG, OPERFAB.CANTREAL"
'    SQLX$ = SQLX$ + " FROM OPERFAB WITH(NOLOCK) INNER JOIN ORDEFABR "
'    SQLX$ = SQLX$ + " ON OPERFAB.IDSUBOR = ORDEFABR.IDSUBOR "
'    SQLX$ = SQLX$ + " WHERE ORDEFABR.STATUORF < 3 "
''    SQLX$ = SQLX$ + " GROUP BY INFOFAB.IDSUBOR, INFOFAB.NUMEOPER,INFOFAB.COD_INTE "
'    SQLX$ = SQLX$ + " ORDER BY OPERFAB.COD_INTE, OPERFAB.IDSUBOR,  OPERFAB.NUMEOPER "
    SQLX$ = "SELECT MASTER.CODINT, OPERFAB.IDSUBOR, OPERFAB.NUMEOPER, OPERFAB.CANTPROG, OPERFAB.CANTREAL"
    SQLX$ = SQLX$ + " FROM MASTER WITH(NOLOCK) LEFT JOIN OPERFAB "
    SQLX$ = SQLX$ + " ON MASTER.CODINT = OPERFAB.COD_INTE "
    SQLX$ = SQLX$ + " WHERE MASTER.STATUS > 0 "
    SQLX$ = SQLX$ + " ORDER BY MASTER.CODINT, OPERFAB.IDSUBOR, OPERFAB.NUMEOPER "
    '
    FIN& = XVALO(VALOBADA("MASTER", "COUNT(*)", "MASTER.STATUS > 0"))
    OF_ANTERIOR$ = ""
    i& = 1
    '
    Dim INFOTMP As ADODB.Recordset
    Set INFOTMP = READSET(SQLX$)
    With INFOTMP
       Do While Not .EOF
          Call TRACE(24, i&, FIN&)
          OF$ = TRIM$(SAFETEXT$(!IDSUBOR))
          SALDO = 0: NUMOPER% = 0
          If OF$ <> "" Then
            Stat% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IDSUBOR='" & OF$ & "'"))
            If Stat% > 2 Then GoTo 90
            NUMOPER% = XVALO(!NUMEOPER)
            CANTIPROGRAMADA = XVALO(!CANTPROG)
            CANTREALIZADA = XVALO(!CANTREAL)
            SALDO = CANTIPROGRAMADA - CANTREALIZADA
          End If
          CIXI% = XVALO(!CODINT)
          If SALDO < 0.05 Then SALDO = 0
          If CIXI% = CIXI_ANTERIOR% Then
             Call GRABA_ANASTOK(CIXI%, NUMOPER%, OF$, SALDO, 0)
          Else
             Call GRABA_ANASTOK(CIXI%, NUMOPER%, OF$, SALDO, 1)
          End If
          '
          CIXI_ANTERIOR% = CIXI%
          i& = i& + 1
90        .MoveNext
       Loop
    End With
    '
    INFOTMP.Close
    Set INFOTMP = Nothing
    Screen.MousePointer = 1
    Call COMUNI("Proceso Terminado")
99  Command7.Enabled = True

End Sub
Sub CARGA_VECTOR_STOCKS(DEPO%)
   '
   'LIMPIA LA MATRIZ
   For i% = 1 To NUMAS%
      STOCKXX(i%, 0) = 0
      STOCKXX(i%, 1) = 0
   Next i%
   i% = 0
   
   'RECORRE EL MASTER LE AGREGA A LA MATRIZ (DECLARADA A NIVEL DE FRM) STOCK EN DEPO% Y TERCEROS
   
   For i% = 1 To NUMAS%
      I1& = i%
      FIN& = NUMAS%
      Call TRACE(20, I1&, FIN&)
      CIXI% = i%
      If STATITEM%(CIXI%) > 0 Then
         STOCKXX(CIXI%, 0) = STOCKTERCEROS(CIXI%)
         STOCKXX(CIXI%, 1) = STODEPOS(CIXI%, DEPO%)
      End If
   Next i%
   '
End Sub
Sub GRABA_ANASTOK(CIXI%, NUMOPER%, OF$, SALDO, GRABA_STOCK%)
        'NOTA *****  MATRICES DECLARADAS A NIVEL DE FRM
        'STOCKXX(CIXI%, 0) MATRIZ CON EL STOCK EN TERCEROS SEGUN CIXI%
        'STOCKXX(CIXI%, 1) MATRIZ CON EL STOCK EN EXPEDICION SEGUN CIXI%
        
       Condi$ = "CODICONJ = " & CIXI & " AND NUMEOPER = " & NUMOPER%
       MAQUI$ = TRIM$(VALOBADA("SECOPER", "CODMAQ0", Condi$, "NOMESS"))
       MAQ = SECPERTE(MAQUI$)
        
       SQLX$ = "SELECT * FROM ANASTOK WHERE Cod_Inte < 1 "
       Dim ANATMP As ADODB.Recordset
       Set ANATMP = New ADODB.Recordset
25     On Error Resume Next
       ANATMP.Open FILTSQL(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
          On Error GoTo 0
          Call CapturaErrorOpen
          GoTo 25
       End If
       On Error GoTo 0
       '
       With ANATMP
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !IDSUBOR = OF$
          !cod_inte = CIXI%
          !COD_EXTE = CODEXT$(CIXI%)
          !NroOperacion = NUMOPER%
          !Sector1 = 0: !Sector2 = 0: !Sector3 = 0: !Sector4 = 0: !Sector5 = 0
          !Sector6 = 0: !Sector7 = 0: !Sector8 = 0: !Sector9 = 0: !Sector10 = 0
          Select Case NUMESECT(MAQ)
             Case Is = 1: !Sector1 = SALDO
             Case Is = 2: !Sector2 = SALDO
             Case Is = 3: !Sector3 = SALDO
             Case Is = 4: !Sector4 = SALDO
             Case Is = 5: !Sector5 = SALDO
             Case Is = 6: !Sector6 = SALDO
             Case Is = 7: !Sector7 = SALDO
             Case Is = 8: !Sector8 = SALDO
             Case Is = 9: !Sector9 = SALDO
             Case Is = 10: !Sector10 = SALDO
          End Select
          '
          If GRABA_STOCK% > 0 Then
             !StkTerceros = STOCKXX(CIXI%, 0)
             !StkExpedicion = STOCKXX(CIXI%, 1)
          Else
             !StkTerceros = 0
             !StkExpedicion = 0
          End If
          '
          .Update
       End With
       On Error Resume Next
       ANATMP.Close
       Set ANATMP = Nothing
       On Error GoTo 0
       
''10      RNDX$ = RANDSTRING$(-10)
''        CONDI$ = "Idsubor = '" & RNDX$ & "'"
''        '
''        If CantRegistros("ANASTOK", CONDI$, "NOMESS") > 0 Then GoTo 10
''        Call CreaRegistro("ANASTOK", "Idsubor", RNDX$, "NOMESS")
''
''        Call ACTUREGISTRO("ANASTOK", "Cod_Inte", CONDI$, CIXI%, RAFE&, "NOMESS")
''        Call ACTUREGISTRO("ANASTOK", "Cod_Exte", CONDI$, CODEXT$(CIXI%), REGAF&, "NOMESS")
''        Call ACTUREGISTRO("ANASTOK", "NroOperacion", CONDI$, NUMOPER%, RAFE&, "NOMESS")
''        Call ACTUREGISTRO("ANASTOK", "Sector" & NUMESECT(MAQ), CONDI$, SALDO, RAFE&, "NOMESS")
''        If GRABA_STOCK% > 0 Then
''           '***** MATRICES DECLARADAS A NIVEL DE FRM
''           Call ACTUREGISTRO("ANASTOK", "StkTerceros", CONDI$, STOCKXX(CIXI%, 0), RAFE&, "NOMESS")
''           Call ACTUREGISTRO("ANASTOK", "StkExpedicion", CONDI$, STOCKXX(CIXI%, 1), RAFE&, "NOMESS")
''        End If
''        '
'        Call ACTUREGISTRO("ANASTOK", "Idsubor", CONDI$, OF$, RAFE&, "NOMESS")

End Sub

Function NUMESECT(NUMESEC)
    NUMESECTOR = 0
    'DEVUELVE LA POSICION DEL SECTOR EN LA TABLA
    '
    For i& = 1 To ULTREG&("!STOKSECC")
       X$ = REGLEIDO$("STOKSECC", i&)
       CODISEC = CVS(Mid$(X$, 1, 4))
       If CODISEC = NUMESEC Then
          NUMESECTOR = i&
          Exit For
       End If
    Next i&
    NUMESECT = NUMESECTOR

End Function

Function STOCKTERCEROS(CIXI%)

    STOCK_ACUMULADO = 0
    For i& = 1 To ULTREG&("DEPOARMA")
       X$ = REGLEIDO$("DEPOARMA", i&)
       DEPO% = Asc(Mid$(X$, 1, 1))
       STOCK_ACUMULADO = STOCK_ACUMULADO + STODEPOS(CIXI%, DEPO%)
    Next i&
    '
    STOCKTERCEROS = STOCK_ACUMULADO
    '
End Function

Private Sub Form_Load()
    Check2.Value = 0
    If Control$("ANALSTOK", "INCLUCON") = "SI" Then Check2.Value = 1
    '
    Call APLICACIONSKINS(Me, Picture1)
    
    Text1 = Control("ANALSTOK", "DEPOEXPE")
    'CREO LOS CAMPOS PAR LA TABLA QUE VA A CONTENER LA INFO DEL REPORTE
    'STOCKS DE TALLERES CANTIDADES
    
    Call CREACAMPO("", "STOCK_ARMADOR", "NUORDEN", 3, 0, 1)
    Call CREACAMPO("", "STOCK_ARMADOR", "Fecha", 8, 0, 1)

    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_1", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_1", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_1", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_2", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_2", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_2", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_3", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_3", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_3", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_4", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_4", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_4", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_5", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_5", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_5", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_6", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_6", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_6", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_7", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_7", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_7", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_8", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_8", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_8", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_9", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_9", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_9", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_10", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_10", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_10", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_11", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_11", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_11", 7, 0)
    '
    Call CREACAMPO("", "STOCK_ARMADOR", "Stock_12", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Enviado_12", 7, 0)
    Call CREACAMPO("", "STOCK_ARMADOR", "Recibido_12", 7, 0)
  
End Sub

Private Sub Text1_Change()
    Label1.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text1))))
End Sub
Sub CARGAGRILLA()
    '
    'LEO LOS DATOS PARA EL ENCABEZADO
    Dim CAMPO1(32)
    PINI% = 24
'    INFO_GRILLA07.MSF.Rows =
    INFO_GRILLA07.MSF.Rows = 1
    INFO_GRILLA07.MSF.Cols = 33

    '
    X$ = REGLEIDO$("STVALFEC", 1)
    CAMPO1(1) = Mid$(X$, 1, 23)
    NUCOLU% = 1: INFO_GRILLA07.TEPRUEBA = "12345678901234567890123456": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 1: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Depósito"
    '
    PINI% = 24:
    For i% = 2 To 32
       CAMPO1(i%) = Mid$(X$, PINI%, 12)
       PINI% = PINI% + 12
       NUCOLU% = i%: INFO_GRILLA07.TEPRUEBA = "123456789012": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = TRIM$(CAMPO1(i%))
    Next i%
    '
'-----------------------------------------------------------------------------------------
'CARGA DETALLE LA GRILLA
'***************

    
    INFO_GRILLA07.MSF.Rows = 3
    For h& = 2 To ULTREG&("STVALFEC")
        PINI% = 24
        X$ = REGLEIDO$("STVALFEC", h&)
        VALORX$ = Mid$(X$, 1, 23)
        INFO_GRILLA07.MSF.Rows = h& + 1
        INFO_GRILLA07.MSF.TextMatrix(h& - 1, 0) = VALORX$
        For COL& = 2 To 32
            VALORX$ = Mid$(X$, PINI%, 12)
            INFO_GRILLA07.MSF.TextMatrix(h& - 1, COL& - 1) = VALORX$
            CAMPO1(h&) = Mid$(X$, PINI%, 12)
            PINI% = PINI% + 12
        Next COL&
        
    Next h&
    '
    INFO_GRILLA07.Label1 = "Cálculo Realizado El Dìa: " & TRIM$(Control("CFGSALSTO", "RECAVALO"))
    INFO_GRILLA07.Show 1
    
    
End Sub

Sub COLOREAR_GRILLASQL(MSF As MSFlexGrid, COLOR As Variant, Optional REG&)
   With MSF
        .Row = REG&
        For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .COL = iCol%
          .CellBackColor = COLOR
        Next iCol%
   End With

End Sub

