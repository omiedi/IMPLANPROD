VERSION 5.00
Begin VB.Form INIFLEX07 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Industrial"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INIFLEX07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   540
      Left            =   420
      TabIndex        =   35
      Top             =   6930
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   420
   End
   Begin VB.CommandButton Command17 
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
      Left            =   2520
      Picture         =   "INIFLEX07.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Depósito Caotico SAPORITI"
      Top             =   6195
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command23 
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
      Left            =   3675
      Picture         =   "INIFLEX07.frx":1E74
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Expedición - Remito de Ventas - Clientes "
      Top             =   5145
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command28 
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
      Left            =   3360
      Picture         =   "INIFLEX07.frx":217E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Recepción de Proveedores Externos"
      Top             =   5985
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   29
      Top             =   210
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   420
      Top             =   0
   End
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
      Left            =   8190
      Picture         =   "INIFLEX07.frx":2488
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Solicitud Interna de Materiales"
      Top             =   6050
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   26
      Top             =   210
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
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
      Height          =   600
      Left            =   5880
      Picture         =   "INIFLEX07.frx":3182
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Informes de Producción"
      Top             =   7200
      Visible         =   0   'False
      Width           =   600
   End
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
      Left            =   9450
      Picture         =   "INIFLEX07.frx":35C4
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Programación por Pedidos de Clentes"
      Top             =   4410
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
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
      Height          =   600
      Left            =   6615
      Picture         =   "INIFLEX07.frx":38CE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5040
      Width           =   600
   End
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
      Left            =   4305
      Picture         =   "INIFLEX07.frx":3BD8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ordenes de Trabajo y Fabricación"
      Top             =   6195
      Width           =   600
   End
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
      Height          =   600
      Left            =   1785
      Picture         =   "INIFLEX07.frx":3EE2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
      Width           =   600
   End
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
      Left            =   7980
      Picture         =   "INIFLEX07.frx":41EC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Planificación de Necesidades de Materiales - Técnica MRP-II"
      Top             =   3360
      Width           =   600
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
      Height          =   600
      Left            =   630
      Picture         =   "INIFLEX07.frx":44F6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   8295
      Picture         =   "INIFLEX07.frx":4800
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   7
      Top             =   315
      Width           =   1905
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   165
         TabIndex        =   16
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   14
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   13
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   10900
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":BF92
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cambio de Clave Personal de Acceso"
         Top             =   6650
         Width           =   700
      End
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":C29C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
         Top             =   7350
         Width           =   700
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":C5A6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   1620
         Width           =   700
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C.R.P."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   710
         Left            =   140
         Picture         =   "INIFLEX07.frx":C6F0
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Programación de Carga de Máquinas"
         Top             =   3510
         Width           =   700
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":C9FA
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5940
         Width           =   700
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":CE3C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5230
         Width           =   700
      End
      Begin VB.CommandButton Command9 
         Caption         =   "P-Cap"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   710
         Left            =   140
         Picture         =   "INIFLEX07.frx":D146
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Planeamiento de Capacidad de Fábrica"
         Top             =   4215
         Width           =   700
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Costos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":DD88
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2805
         Width           =   700
      End
      Begin VB.CommandButton Command2 
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":E092
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   920
         Width           =   700
      End
      Begin VB.CommandButton Command1 
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
         Height          =   700
         Left            =   140
         Picture         =   "INIFLEX07.frx":E39C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   700
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   6720
      Picture         =   "INIFLEX07.frx":E4E6
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8085
      Picture         =   "INIFLEX07.frx":15FE0
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   3570
      Width           =   2115
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4350
      Picture         =   "INIFLEX07.frx":1DC72
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   6405
      Width           =   2010
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1890
      Picture         =   "INIFLEX07.frx":25A3C
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   735
      Picture         =   "INIFLEX07.frx":2D66A
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   3570
      Width           =   2010
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   2940
      Picture         =   "INIFLEX07.frx":35298
      ScaleHeight     =   1485
      ScaleWidth      =   5055
      TabIndex        =   5
      Top             =   1575
      Width           =   5055
   End
   Begin VB.Line Line7 
      X1              =   5250
      X2              =   9135
      Y1              =   2310
      Y2              =   4200
   End
   Begin VB.Line Line6 
      X1              =   5355
      X2              =   7770
      Y1              =   2310
      Y2              =   5565
   End
   Begin VB.Line Line5 
      X1              =   5355
      X2              =   5355
      Y1              =   2310
      Y2              =   6300
   End
   Begin VB.Line Line4 
      X1              =   3045
      X2              =   5355
      Y1              =   5565
      Y2              =   2310
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   5355
      Y1              =   4095
      Y2              =   2310
   End
End
Attribute VB_Name = "INIFLEX07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CONVIDOCDIGI()
    '
    'INICIALIZO VARIABLE
    CONT% = 0
    Static INGRESO%
    Static ULTIMODOCU&
    '
    If INGRESO% < 1 Then
        '1.-CREO LOS CAMPOS
        Call ABRECONEXION
        Call CREACAMPO("", "PDOC_TBL", "NUMEDOC", 4, 0, 1)
        Call CREACAMPO("", "PDOC_TBL", "Referencia", 10, 128)
        Call CREACAMPO("", "PDOC_TBL", "FechayHora", 8, 0)
        Call CREACAMPO("", "PDOC_TBL", "Usuario", 10, 64)
        Call CREACAMPO("", "PDOC_TBL", "Terminal", 10, 64)
        Call CREACAMPO("", "PDOC_TBL", "Impresora", 10, 64)
        Call CREACAMPO("", "PDOC_TBL", "Imagen", 11, 0)
        Call CREACAMPO("", "PDOC_TBL", "TIPODOCU", 10, 32, 1)
        Call CREACAMPO("", "PDOC_TBL", "NUMECOMPLE", 10, 32, 1)
        Call CREACAMPO("", "PDOC_TBL", "NUMEBUSQUE", 10, 16, 1)
        INGRESO% = 1
    End If
    '
    '2.-Recorro el  el pdocclst
    FIN& = ULTREG&("PDOCLST")
    If ULTIMODOCU& < 1 Then ULTIMODOCU& = FIN&
    For i& = ULTIMODOCU& To 1 Step -1
       '
       If CONT% > 5 Then GoTo 99
       '
       CUENT& = i&
       '
       TX$ = REGLEIDO$("PDOCLST", i&)
       
       'LEO LOS VALORES DE PDOCLST
       NUMDOC1& = CVS(Left$(TX$, 4))
       '
       '3.-VALIDO SI ESTA EN LA TABLA SQL
       CONDI$ = "NUMEDOC =" & NUMDOC1&
       EXISREG& = CantRegistros("PDOC_TBL", CONDI$, "NOMESS")
       If EXISREG& > 0 Then GoTo 50 '(NEXT)
       '
       'LEO EL RESTO DE LOS DATOS
       REFER$ = TRIM$(Mid$(TX$, 5, 60))
       FE% = CVI(Mid(TX$, 65, 2))
       HORX% = CVI(Mid(TX$, 67, 2))
       FEYHO = FETEXCOR1$(FE%) & " " & HORATEX$(HORX%)
       USUA$ = TRIM$(Mid(TX$, 69, 30))
       TERMI$ = TRIM$(Mid(TX$, 99, 30))
       IMPRE$ = TRIM$(Mid(TX$, 129, 128))
       IMAG2$ = IMAGENPDOC$(NUMDOC1&) 'IMAGEN DE PDOCSPL
       
       '3.-VALIDO SI ESTA EN LA TABLA SQL NUEVAMENTE
       EXISREG& = CantRegistros("PDOC_TBL", CONDI$, "NOMESS")
       If EXISREG& > 0 Then GoTo 50 '(NEXT)
       '
       ' 4.GRABO EN SQL SI EL REGISTRO NO EXISTE
       'CREO EL REGISTRO
       Call CreaRegistro("PDOC_TBL", "NUMEDOC", NUMDOC1&, "NOMESS/NOAUDIT")
       'GRABO
       Call ACTUREGISTRO("PDOC_TBL", "REFERENCIA", CONDI$, REFER$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "FECHAYHORA", CONDI$, FEYHO, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "USUARIO", CONDI$, USUA$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "TERMINAL", CONDI$, TERMI$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "IMPRESORA", CONDI$, IMPRE$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "IMAGEN", CONDI$, IMAG2$, REGAF&, "NOMESS/NOAUDIT/RETRY")
       '
       'GRABO EN LOS NUEVOS CAMPOS
       Call NUEVOS_DATOS(REFER$, NUMDOC1&)
       '
       CONT% = CONT% + 1
50  Next i&

99  ULTIMODOCU& = CUENT&
    
End Sub

Sub NUEVOS_DATOS(REFE$, numeroDocumento&)
        
         POS1% = 0
         POS2% = 0
         REFE$ = UCase$(REFE$)
         If Left(REFE$, Len("AJUSTE DE STOCKS")) = "AJUSTE DE STOCKS" Then
            'Ajuste de Stocks 10000
            TIPOX$ = "AJUSTE DE STOCKS"
            
         ElseIf Left$(REFE$, Len("AJUSTE SALDO CLIENTE")) = "AJUSTE SALDO CLIENTE" Then
            'Ajuste Saldo Cliente AS-101019-03 - RONALD FREDY IRIARTE - 19/10/10
            TIPOX$ = "AJUSTE SALDO CLIENTE"
         
         ElseIf Left$(REFE$, Len("ANALISIS DE COSTO")) = "ANALISIS DE COSTO" Then
            'ANALISIS DE COSTO 0014 - 17/01/11
            TIPOX$ = "ANALISIS DE COSTO"
         
         ElseIf Left$(REFE$, Len("BAJA DE NUMERO DE SERIE")) = "BAJA DE NUMERO DE SERIE" Then
            'BAJA DE NUMERO DE SERIE 889040742170
            TIPOX$ = "BAJA DE NUMERO DE SERIE"
            
         ElseIf Left$(REFE$, Len("BOLETA DE RECEPCIÓN")) = "BOLETA DE RECEPCIÓN" Then
            'BOLETA DE RECEPCIÓN  1 - PROVEEDOR DE PRUEBA - 02/10/08
            TIPOX$ = "BOLETA DE RECEPCIÓN"
           
         ElseIf Left$(REFE$, Len("COMPROBANTE DE DIARIO")) = "COMPROBANTE DE DIARIO" Then
           ' COMPROBANTE DE DIARIO 90107-001 - 07/01/09
            TIPOX$ = "COMPROBANTE DE DIARIO"
            
         ElseIf Left$(REFE$, Len("DEPOSITO BANCARIO")) = "DEPOSITO BANCARIO" Then
           'DEPOSITO BANCARIO DB-101126-00 - 26/11/10
            TIPOX$ = "DEPOSITO BANCARIO"
         
          ElseIf Left$(REFE$, Len("DESPACHO DE MATERIALES")) = "DESPACHO DE MATERIALES" Then
            'Despacho de Materiales DM.30765 - SIMTECH LTDA - 01/10/10
            TIPOX$ = "DESPACHO DE MATERIALES"
         
         ElseIf Left$(REFE$, Len("ETIQUETA DE EXPEDICIÓN")) = "ETIQUETA DE EXPEDICIÓN" Then
            'ETIQUETA DE EXPEDICIÓN 07441 - 04/11/10
            TIPOX$ = "ETIQUETA DE EXPEDICIÓN"
            
         ElseIf Left$(REFE$, Len("FACTURA NRO.")) = "FACTURA NRO." Then
            'FACTURA NRO.0001-00001250 - CLIENTE DE PRUEBA - 02/10/08
            TIPOX$ = "FACTURA NRO."
         ElseIf Left$(REFE$, Len("NOTA DE CREDITO")) = "NOTA DE CREDITO" Then
            'NOTA DE CREDITO 0001-00000156 - 3 WAY SOLUTION SA - 15/04/11
            TIPOX$ = "NOTA DE CREDITO"
            
         ElseIf Left$(REFE$, Len("NOTA DE PEDIDO")) = "NOTA DE PEDIDO" Then
            'NOTA DE PEDIDO 002136-1 - COELAR S.R.L. - 24 DE ENERO DE 2011
            TIPOX$ = "NOTA DE PEDIDO"
            
         ElseIf Left$(REFE$, Len("NOTA DE DEBITO")) = "NOTA DE DEBITO" Then
            'Nota de Debito 0001-00033662 - NUSA DUA SA - 15/09/10
            TIPOX$ = "NOTA DE DEBITO"
 
         ElseIf Left$(REFE$, Len("FUGA A TIERRA")) = "FUGA A TIERRA" Then
            'FUGA A TIERRA 2429123456-0
            TIPOX$ = "FUGA A TIERRA"
         ElseIf Left$(REFE$, Len("ORDEN DE FABRICACIÓN")) = "ORDEN DE FABRICACIÓN" Then
            'ORDEN DE FABRICACIÓN OF-022349-1
            TIPOX$ = "ORDEN DE FABRICACIÓN"
         
         ElseIf Left$(REFE$, Len("ORDEN DE PAGO NRO")) = "ORDEN DE PAGO NRO" Then
           'ORDEN DE PAGO NRO. 2 - WAGNER JAVIER - 03/10/08
            TIPOX$ = "ORDEN DE PAGO NRO."
            
         ElseIf Left$(REFE$, Len("ORDEN DE DESPACHO")) = "ORDEN DE DESPACHO" Then
            'ORDEN DE DESPACHO 002726 - PVR SA - 18/01/11
            TIPOX$ = "ORDEN DE DESPACHO"
            
         ElseIf Left$(REFE$, Len("ORDEN DE SERVICIO")) = "ORDEN DE SERVICIO" Then
            'ORDEN DE SERVICIO 460 - WAGNER JAVIER FERNANDO - 20/09/10
            TIPOX$ = "ORDEN DE SERVICIO"
            
         ElseIf Left$(REFE$, Len("PEDIDO INTERNO")) = "PEDIDO INTERNO" Then
            'PEDIDO INTERNO 000001 - CLIENTE DE PRUEBA - 02/10/08
            TIPOX$ = "PEDIDO INTERNO"
            
         ElseIf Left$(REFE$, Len("PEDIDO DE CLIENTE")) = "PEDIDO DE CLIENTE" Then
            'PEDIDO DE CLIENTE 002727 - PVR SA - 18/01/11
            TIPOX$ = "PEDIDO DE CLIENTE"
         
         ElseIf Left$(REFE$, Len("PLANILLA DE REPARACIÓN")) = "PLANILLA DE REPARACIÓN" Then
           'Planilla de Reparación 1 - ADAMOLI VALERIA-CLEMENTE INMACULA
            TIPOX$ = "PLANILLA DE REPARACIÓN"
   
         ElseIf Left$(REFE$, Len("PRESUP.REPARACIÓN EXTERNO")) = "PRESUP.REPARACIÓN EXTERNO" Then
            'Presup.Reparación Externo 19 - CENTURION CRISTALDO MARIA ESTHER - 17/03/11
            TIPOX$ = "PRESUP.REPARACIÓN EXTERNO"
 
         ElseIf Left$(REFE$, Len("PRESUPUESTO DE REPARACIÓN")) = "PRESUPUESTO DE REPARACIÓN" Then
            'Presupuesto de Reparación 27 - 4 DE FEBRERO S.R.L. - 04/04/11
            TIPOX$ = "PRESUPUESTO DE REPARACIÓN"
            
         ElseIf Left$(REFE$, Len("PRESUPUESTO")) = "PRESUPUESTO" Then
            'Presupuesto 74 - OBRA SOCIAL DEL PERSONAL DE TV - 14/09/09
            TIPOX$ = "PRESUPUESTO"
            
         ElseIf Left$(REFE$, Len("PUESTA A TIERRA")) = "PUESTA A TIERRA" Then
            'PUESTA A TIERRA 2429123456-0
            TIPOX$ = "PUESTA A TIERRA"
            
         ElseIf Left$(REFE$, Len("RECIBO NRO.")) = "RECIBO NRO." Then
           'Recibo Nro. 29038 - KAIDAR MARIEL FABIANA - 19/10/10
            TIPOX$ = "RECIBO NRO."
          
         ElseIf Left$(REFE$, Len("RECHAZO EN LÍNEA")) = "RECHAZO EN LÍNEA" Then
            'RECHAZO EN LÍNEA 002077 - 02/11/10
            TIPOX$ = "RECHAZO EN LÍNEA"
          
         ElseIf Left$(REFE$, Len("REGISTRO DE CONTROL FINAL")) = "REGISTRO DE CONTROL FINAL" Then
            'REGISTRO DE CONTROL FINAL 586312100014-0
            TIPOX$ = "REGISTRO DE CONTROL FINAL"
            
         ElseIf Left$(REFE$, Len("REGISTRO DE MEDICION")) = "REGISTRO DE MEDICION" Then
            'REGISTRO DE MEDICION 2429123456-0
            TIPOX$ = "REGISTRO DE MEDICION"
          
         ElseIf Left$(REFE$, Len("REMITO DE DEVOLUCIÓN")) = "REMITO DE DEVOLUCIÓN" Then
            'Remito de Devolución RD.1281 - ACTION TRAVEL - 13/10/10
            TIPOX$ = "REMITO DE DEVOLUCIÓN"
          
         ElseIf Left$(REFE$, Len("REMITO DE TRASLADO")) = "REMITO DE TRASLADO" Then
            'Remito de Traslado RC-751 - ROBERTO PERSIANI - 06/09/10
            TIPOX$ = "REMITO DE TRASLADO"
            
         ElseIf Left$(REFE$, Len("REMITO NRO.")) = "REMITO NRO." Then
            'Remito Nro.0001-00040261 - JUAN JOSE LOVERA E HIJOS - 14/09/10
            TIPOX$ = "REMITO NRO."
            
         ElseIf Left$(REFE$, Len("TRANSFERENCIA DE STOCK")) = "TRANSFERENCIA DE STOCK" Then
            'Transferencia de Stock 486
            TIPOX$ = "TRANSFERENCIA DE STOCK"
          
         ElseIf Left$(REFE$, Len("VALE DE CONSUMO DE MATERIALES")) = "VALE DE CONSUMO DE MATERIALES" Then
            'Vale de Consumo de Materiales VM.1144 - 20/09/10
            TIPOX$ = "VALE DE CONSUMO DE MATERIALES"
         
         ElseIf Left$(REFE$, Len("VALE DE MATERIALES")) = "VALE DE MATERIALES" Then
            'Vale de Materiales 023801 - 10/01/11
            TIPOX$ = "VALE DE MATERIALES"
          
         End If
         '
         POS1% = Len(TIPOX$)
         If POS1% > 0 Then
            POS1% = POS1% + 1
            TX$ = LTrim(Mid$(REFE$, POS1%))
            POS2% = InStr(TX$, " ")
            NUCOMP$ = TRIM$(Mid$(TX$, 1, POS2%))
            If POS2% < 1 Then NUCOMP$ = TRIM$(Mid$(REFE, POS1%))
            '
            NUBUSQ$ = NUBUSQUE$(NUCOMP$, TIPOX$)
            '
            'PARA EL CASO DE FACTURA, REMITO, RECIBO QUE TERMINAN CON 'NRO.'
            If Right$(TIPOX$, 4) = "NRO." Then TIPOX$ = TRIM$(Mid$(TIPOX$, 1, Len(TIPOX$) - 4))
            '
            CONDI$ = "NUMEDOC =" & numeroDocumento&
            Call ACTUREGISTRO("PDOC_TBL", "TIPODOCU", CONDI$, TIPOX$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("PDOC_TBL", "NUMECOMPLE", CONDI$, NUCOMP$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("PDOC_TBL", "NUMEBUSQUE", CONDI$, NUBUSQ$, REGAF&, "NOMESS")
 
         End If
End Sub

Function IMAGENPDOC$(DOCUNU&)

    Dim TD0 As String * 512
    TPSP$ = ""
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
      Close #NX0%
      NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    End If
    '
    LOFI = LOF(NX0%)
    URE& = LOFI / 512
    If LOFI < 0 Then URE& = 4194304
    '
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
'For KKII% = 1 To 24: Get #NX0%, KKII%, TD0$: TTXX$ = TTXX$ + Str$(CVS(Left$(TD0$, 4))): Next KKII%: MsgBox TTXX$
7   Close #NX0%
    IMAGENPDOC$ = ""
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      Get #NX0%, L&, TD0$
      DCC& = CVS(Left$(TD0$, 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      Get #NX0%, KKL&, TD0$
      TBUF$ = TD0$
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  Close #NX0%
    IMAGENPDOC$ = TPSP$
    '
99  End Function



Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  If PROCACTI% < 1 Then
    Call GENBACKUP
    Call CANCELUCOM
    Call FINAL("")
  End If
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If EXISTE%("IFABRI07.EXE") > 0 Then
     Call CONECRUN("IFABRI07", "Informes de Fabricación")
   Else
     Call COMUNI("Falta módulo IFABRI07")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG07")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO07.EXE") > 0 Then
     Call CONECRUN("CONSTO07", "Control de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO07")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("PROMRP07.EXE") > 0 Then
     Call CONECRUN("PROMRP07", "Programación de Necesidades MRP")
   Else
     Call COMUNI("Falta módulo PROMRP07")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA07", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
'   If EXISTE%("ORFABR04.EXE") > 0 Then
'     Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
'   Else
'     Call COMUNI("Falta módulo ORFABR04")
'   End If
  '\\\OT-05-0415-WAR-22/06/05///
   If EXISTE%("ORPROD09.EXE") > 0 Then
      Call CONECRUN("ORPROD09.EXE", "Sistema de Ordenes de Fabricación y Envios de Prendas")
     ElseIf EXISTE%("ORCONF07.EXE") > 0 Then
       Call CONECRUN("ORCONF07", "Sistema de Ordenes de Fabricación y Envios de Prendas")
     ElseIf EXISTE%("ORFABR07.EXE") > 0 Then
       Call CONECRUN("ORFABR07", "Sistema de Ordenes de Fabricación")
     ElseIf EXISTE%("MINORF07.EXE") > 0 Then
       Call CONECRUN("MINORF07", "Sistema Basico de Fabricación")
     Else
       Call COMUNI("Falta módulo ORFABR07")
   End If
   '\\\OT-05-0415-WAR-FIN///
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
     Call CANCELUCOM
99   ' conecta con FLOGON.EXE
     On Error Resume Next
     RETOR$ = TRIM$(UCase$(App.EXEName))
     Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
     If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
     End If
     Call FINAL("")
   End If
   '
   Command16.Enabled = True
End Sub


Private Sub Command17_Click()
   Command17.Visible = False
   ACONEC$ = "TRANSA07"
   Call CONECRUN(ACONEC$, "Deposito Caotico SAPORITI")
   Command17.Visible = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("ORCOMP07.EXE") > 0 Then
     Call CONECRUN("ORCOMP07", "Sistema de Ordenes de Compra")
   Else
     Call COMUNI("Falta módulo ORCOMP07")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   'Call CONECRUN("FSETUP04/CLAVUSER", "Cambio de Clave Personal de Acceso")
   FSETUPNEW.Command5_Click
99 Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command20_Click()
   Timer3.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO07"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command28_Click()
   Command28.Visible = False
   ACONEC$ = "BRECEP07"
   Call CONECRUN(ACONEC$, "Recepcion de Materiales")
   Command28.Visible = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   '
   If EXISTE%("FLGEST07.EXE") > 0 Then
     OPX% = COMALTER%("Opciones de Conexión a Accesos Directos:\\FLEXOFT Gestión Administrativa\Otros Accesos Directos")
     If OPX% = 1 Then
         Call CONECRUN("FLGEST07", "Sistema de Gestión Administrativa")
       ElseIf OPX% = 2 Then
         GoTo 90
     End If
     GoTo 99
   End If
   '
90 Call ACCDIR("FLEXO00")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("FSETUP07/", "CONFIGURACION GENERAL") < 2 Then GoTo 99
   FSETUPNEW.Show 1
   Call FINAL("")

'   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
'   Call FINAL("")
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
If EXISTE%("PREACH07.EXE") > 0 Then
     Call CONECRUN("PREACH07", "Presupuesto ACHERNAR")
     GoTo 99
End If
   If EXISTE%("ANACOS07.EXE") > 0 Then
     Call CONECRUN("ANACOS07", "Análisis de Costos y Presupuestación de Producto")
   Else
     Call COMUNI("Falta módulo ANACOS07")
   End If
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If DERACCE%("GENREP", "Generador de Reportes") >= 2 Then
     Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If EXISTE%("CARMAQ07.EXE") > 0 Then
     Call CONECRUN("CARMAQ07", "Cálculo de Capacidad y Secuencia de Máquina")
   Else
     Call COMUNI("Falta módulo CARMAQ07")
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
'   OPX% = 2
'   If EXISTE%("PRESUP07.EXE") > 0 Then
'     OPX% = ALTERNA%("Generación de Presupuestos\Ordenes de Venta")
'   End If
'   If OPX% = 1 Then
'        Call CONECRUN("PRESUP07", "Presupuestos a Clientes")
'      '16/03/07(OT-01-0114)
'      'PREGUNTO SI LA OPCION ELEGIDA POR EL USUARIO ES 'PEDIDOS DE CLIENTES'
'      'SI ELIGE CANCELAR O HACE CLIC EN LA X, ME MUESTRA NUEVAMENTE EL MENU
'      ElseIf OPX% = 2 Then
'      '16/03/07(FIN)
'        Call CONECRUN("PEDIDO07", "Pedidos de Clientes")
'   End If
   If EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 And EXISTE%("COTCLI09.EXE") > 0 Then
      OPX% = ALTERNA%("Emisión de Presupuestos\Pedidos de Clientes\Solicitud de Cotización")
   ElseIf EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 Then
      OPX% = ALTERNA%("Emisión de Presupuestos\Pedidos de Clientes")
   End If
   If OPX% = 1 Then GoTo 10
   If OPX% = 2 Then GoTo 20
   If OPX% = 3 Then GoTo 30
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   '
   If EXISTE%("PRESUP07.EXE") > 0 Then
10    Call CONECRUN("PRESUP07", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
   If EXISTE%("PEDIDO07.EXE") > 0 Then
20    Call CONECRUN("PEDIDO07", "Gestión de Pedidos de Clientes")
      GoTo 99
   End If
   '
   If EXISTE%("COTCLI09.EXE") > 0 Then
30    Call CONECRUN("COTCLI09", "Solicitud de Cotización")
      GoTo 99
   End If
   '
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("CALCAP07.EXE") > 0 Then
     Call CONECRUN("CALCAP07", "Cálculo de Capacidad")
   Else
     Call COMUNI("Falta módulo CALCAP07")
   End If
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   CCOLO% = 192
   BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture1.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture2.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture3.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture4.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture5.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture6.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture8.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10    Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call PONELOGO
   '
   If EXISTE%("IFABRI07.EXE") > 0 Then
      Command10.Visible = True
   End If
   '
   If EXISTE%("SOLIMA02.EXE") > 0 Then
      Command14.Visible = True
   End If
   '
   If EXISTE%("ANACOS07.EXE") > 0 Then
      Command5.Enabled = True
   End If
   '
   If EXISTE%("CARMAQ07.EXE") > 0 Then
      Command7.Enabled = True
   End If
   '
   If EXISTE%("CALCAP07.EXE") > 0 Then
      Command9.Enabled = True
   End If
   '
   If EXISTE%("BRECEP07.EXE") > 0 Then
      Command28.Visible = True
   End If
   '
   If EXISTE%("REMITO07.EXE") > 0 Then
      Command23.Visible = True
   End If
   '
   If EXISTE%("TRANSA07.EXE") > 0 Then
      Command17.Visible = True
   End If
   '
'   If EXISTE%("PEDIDO07.EXE") > 0 Then
'      Command8.Visible = True
'      If EXISTE%("PRESUP07.EXE") > 0 Then
'        Command8.ToolTipText = "Acceso a Presupuestos y Ordenes de Venta"
'      End If
'   End If
   If EXISTE%("PEDIDO07.EXE") > 0 Then
      Command8.Visible = True
      Command8.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
      'ped_clientes.Enabled = True
      'ped_clientes1.Enabled = True
   End If
   If EXISTE%("PRESUP07.EXE") > 0 Then
      Command8.ToolTipText = "Sub-Sistema de Pedidos de Clientes"
      Command8.Visible = True
      'pre_clientes.Enabled = True
   End If
   If EXISTE%("COTCLI09.EXE") > 0 Then
      Command8.ToolTipText = "Solicitud de Cotizaciones"
      Command8.Visible = True
   End If
  If EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 And EXISTE%("COTCLI09.EXE") > 0 Then
      Command8.ToolTipText = "Sub-Sistemas de Presupuestos, Pedidos de Clientes y Solicitud de Cotizaciones"
   ElseIf EXISTE%("PEDIDO07.EXE") > 0 And EXISTE%("PRESUP07.EXE") > 0 Then
      Command8.ToolTipText = "Sub-Sistemas de Presupuestos y Pedidos de Clientes"
   End If
   '
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image1.Picture = LoadPicture("")
        On Error Resume Next
        Image1.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        ' Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
        If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") < 1 Then ' PARA CASO GENTILI
           Picture9.Visible = True
           IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
           '
           FORMALTER.Image3.Picture = LoadPicture("")
           On Error Resume Next
           FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
           On Error GoTo 0
           GoTo 12
        End If
        '
      End If
   End If
   '
   Picture9.Visible = False
   If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") > 0 Then
     Picture9.Visible = True
     IMAGEFILE$ = LUDAT$ + "LOGOEMPR.BMP"
     '
     FORMALTER.Image3.Picture = LoadPicture("")
     FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
     '
12   HIMAGEN = FORMALTER.Image3.Picture.Width
     VIMAGEN = FORMALTER.Image3.Picture.Height
     '
     HCUADRO = Picture9.Width
     VCUADRO = Picture9.Height
     '
     REDUH = HIMAGEN / HCUADRO
     REDUV = VIMAGEN / VCUADRO
     REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
     If REDUC < 1 Then REDUC = 1
     '
     With Picture9
       .Height = VCUADRO * REDUV / REDUC
       .Width = HCUADRO * REDUH / REDUC
       .Left = .Left + (HCUADRO - Picture9.Width) / 2
         If .Left < 300 Then Left = 300
       .Top = .Top - (VCUADRO - Picture9.Height) / 2
         If .Top < 400 Then .Top = 400
     End With
     '
     Image1.Width = Picture9.Width
     Image1.Height = Picture9.Height
     Image1.Stretch = True
     On Error Resume Next
     Image1.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
     On Error GoTo 0
     '
     FORMALTER.Image3.Picture = LoadPicture("")
   End If
   '
End Sub

'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
End Sub

Private Sub Timer2_Timer()
Exit Sub
   If EXISTE%("C:\FLEXOFT\PRUEBA.PDF") > 0 Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ <> "" Then
       RUTOR$ = ACROL$ + " "
       TEXACT$ = "" '  "Adobe"
       RUTAR$ = "C:\FLEXOFT\"
       ANAME$ = RUTOR$ + RUTAR$ + "PRUEBA.PDF"
       RETURNVALUE = Shell(ANAME$, 1)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [PRUEBA.pdf]" ' Adobe Reader"
       'If Err Then
         'Resume 20
20       'On Error GoTo 0
         'TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         'GoTo 18
       'End If
       On Error GoTo 0
       '
     End If
   End If
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      x$ = TRIM$(UCase$(File1.LIST(U& - 1)))
      If x$ <> "" Then
        If x$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + x$) > 0 Then
            On Error GoTo 10
            Kill LUCOM$ + x$
            On Error GoTo 0
          End If
        End If
      End If
   Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(x$, Len(x$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function

Private Sub Timer3_Timer()
   '
   Static CONTA%
   CONTA% = CONTA% + 1
   If CONTA% = 5 Then
      CONTA% = 0
      Call CONVIDOCDIGI
   End If
'   Static CONTAXX%
'   Static MAXI%
'   CONTA% = CONTA% + 1
'   If CONTAXX% < 1 Then
'10   MAXI% = XVALO(InputBox$("Ingrese MAXIMO "))
'     If MAXI% < 1 Then GoTo 10
'   End If
'   '
'   CONTAXX% = CONTAXX% + 1
'   If CONTA% = 5 Then
'      CONTA% = 0
'      If (CONTAXX% * 5) >= MAXI% Then Timer3.Enabled = False
'      Call CONVIDOCDIGI
'   End If
End Sub

