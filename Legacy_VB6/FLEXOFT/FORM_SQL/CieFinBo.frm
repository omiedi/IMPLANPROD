VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CieFinBo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cierre Final (Consumo de Bolsas para Envaplast)"
   ClientHeight    =   7305
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   18105
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7305
   ScaleWidth      =   18105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   18135
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   4
         Left            =   14880
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Estado(0-Ab.,3Cerr.)"
         Top             =   200
         Width           =   2000
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   14880
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Doble clic para abrir Orden Cerrada"
         Top             =   440
         Width           =   2000
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   14595
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   3
            Left            =   12000
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Kg. Acumulados de Cierres"
            Top             =   180
            Width           =   2460
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   12000
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   2460
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   420
            Width           =   2235
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   4330
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   180
            Width           =   4330
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "Cant.Proyec."
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   9060
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   180
            Width           =   2235
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1260
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   6660
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   180
            Width           =   1260
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   2
            Left            =   9060
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Cant.Aprob."
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "REALIZAR CONSUMO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15480
         TabIndex        =   4
         Top             =   6840
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   5595
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   17955
         _ExtentX        =   31671
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   5786388
         ForeColorFixed  =   16777215
         BackColorSel    =   10468997
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material Transferido a la O.Fab. Activa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H009FBE85&
         Caption         =   "En linea Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2900
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CieFinBo.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CieFinBo.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Transacciones"
   End
   Begin VB.Menu mnuCtalis 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
   End
   Begin VB.Menu mnuacces 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "CieFinBo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HAYFALTANTESCTP%
   Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ActualizaListaSeleccion As Boolean




Private Sub Command2_Click()
  Command2.Enabled = False
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  NUORIT% = 0
  CICONJ% = CODINT%(Text11)
  'VALIDO
  For i& = 1 To MSF1.Rows - 1
    CODICOMPO% = CODINT%(MSF1.TextMatrix(i&, 1))
    LOTE$ = MSF1.TextMatrix(i&, 3)
    METROS# = XVALO(MSF1.TextMatrix(i&, 4))
    PESOKG# = XVALO(MSF1.TextMatrix(i&, 5))
    If CODICOMPO% < 1 Then
       Call COMUNI("Falta Còdigo de Componente en Fila: " & i&)
       GoTo 99
    End If
    If ESTRAZABLE(CODICOMPO%) = 1 And TRIM$(LOTE$) = "" Then
       Call COMUNI("Falta Lote en Fila: " & i&)
       GoTo 99
    End If
    If METROS# < 1 Then
       Call COMUNI("Falta Cantidad en  Componente en Fila: " & i&)
       GoTo 99
    End If
  Next i&
  NORFA$ = Text10
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  For i& = 1 To MSF1.Rows - 1
    CODICOMPO% = CODINT%(MSF1.TextMatrix(i&, 1))
    LOTE$ = MSF1.TextMatrix(i&, 3)
    METROS# = XVALO(MSF1.TextMatrix(i&, 4))
    CANTI = METROS#
    If Unimed(CODICOMPO%) = "U." Then 'CASO QUILMES
         CANTI = 1 'UNA BOBINA POR LOTE
    End If

    PESOKG# = XVALO(MSF1.TextMatrix(i&, 7))
    NUORIT% = NUORIT% + 1
    If InStr(EPAC$, "QUILMES") > 0 Then
        TIP% = TIMATE%(CODICOMPO%)
        If TIP% = 41 Then 'QUILMES CASO DE BOBINA
           METROS# = 1
        End If
        'PARA QUE EN QUILMES CONSUMA LA CANTIDAD DE METROS TOTALES DE LA BOBINA
        METROS# = XVALO(MetrosSegunLote#(LOTE$))
        CODFAMI$ = TRIM$(CODFAMIL$(CODICOMPO%))
        CODFAMICONJ$ = TRIM$(CODFAMIL$(CICONJ%))
        'SI EL COMPONENTE ES TELA IMPRESA (ES BOLSA) SI EL ITEM A CERRAR ES T.IMPRESA ES EL COMPONENTE ES BOBINA BASE.
        'LE AGREGUE QUE  SI EL CONJUNTO ES UNA BOLSA Y EL COMPONENTE PAPEL TAMBIEN CONSUMA DEL 3 AGREGADO TAMBIEN TL
        If Left$(CODFAMI$, 2) = "TL" Or Left$(CODFAMI$, 2) = "TI" Or Left$(CODFAMICONJ$, 2) = "TI" Or Left$(CODFAMI$, 2) = "TT" Or (Left$(CODFAMICONJ$, 2) = "BO" And Left$(CODFAMI$, 2) = "PA") Then
          DEPOCORTE% = XVALO(Control("CONTEPLA", "DEBOAIMP"))
        End If
    End If
    
    Call GENERACONSUMO(Text10, CODINT(Text11), CODICOMPO%, CANTI, DEPOCORTE%, NUORIT%, LOTE$, PESOKG#, METROS#)
  Next i&
  Screen.MousePointer = 1
  Dim ctp As New Conteplast
  Call COMUNI("Consumo Generado")
  Call ctp.CargaMatTranfXOrden(CieFinBo.MSF1, TRIM$(NORFA$), DEPOCORTE%, 0, 0, 1)
99 Command2.Enabled = True
  
End Sub

Private Sub Form_Load()
  
  Call CARGA_ENCABEZADO(MSF1, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
  
End Sub


Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    If REG& < 1 Then Exit Sub
    Unload VENTABNEW

    Campos$ = "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf/F16.2"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "CONSUMO"
    VENTANA.CampEditables = "7"
    VENTANA.Inicializar = 1
    '
    Call EDITARFILA(MSF1, VENTABNEW.MSF, REG&)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call PEGARFILAEDITADA(VENTABNEW.MSF, MSF1, REG&)
       LOTE$ = MSF1.TextMatrix(REG&, 3)
       PESOKG# = XVALO(MSF1.TextMatrix(REG&, 7))
       METROSSEGUNPESO = XVALO(MetrosSegunPesoSegúnLote#(LOTE$, PESOKG#))
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(METROSSEGUNPESO, "F10.1")
    End If
    
    Unload VENTABNEW
End Sub

Private Sub Text10_Change()
    NORFA$ = TRIM$(Text10)
    MSF1.Rows = 1

    If NORFA$ = "" Then
      Exit Sub
    End If
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    Dim ctp As New Conteplast
    DEPOCORTE% = XVALO(Control("", "DEPODEST"))
    CI1% = CODIFAB%(NORFA$)
'    If TRIM$(UCase$(CODFAMIL$(CI1%))) = "CC" Then
'      Me.LBLDEPOCONS = XVALO(CONTROL("CONTEPLA", "DEPOCINT"))
'      Me.LBLDEPOENT = XVALO(CONTROL("", "DEPODEST"))  'DEPOCORTE%
'    End If
    TXTCODINT = CI1%
    Text11 = CODEXT$(CI1%)
    Text4 = DESCRIT0$(CI1%)
    CANTI = CantiProyectadaOF(NORFA$)
    Text3 = CANTI
    Text1 = CantiAprobadaOF(NORFA$)
    CONDI$ = "CODIMOVI = 'OF'  AND CANTINGRE > 0 AND DOPRICOR = '" & NORFA$ & "'"
    KGAcumulados# = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI$, "NOMESS"))
    Text5 = FORMATNUM$(KGAcumulados#, "f10.0")
    'AGREGADO ESTE PARRAFO EL 30/06/24
    CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
    If InStr(EPAC$, "QUILMES") > 0 Then
       If Left$(CODFAMI$, 2) = "BO" Then
          DEPOCORTE% = XVALO(Control("CONTEPLA", "DEBOAIMP"))
       End If
    End If

    Call ctp.CargaMatTranfXOrdenConsumoFinalBolsasQuilmes(CieFinBo.MSF1, TRIM$(NORFA$), DEPOCORTE%, 0, 0, 1)
    Text6 = XVALO(STATORFA%(TRIM$(NORFA$)))
    Set ctp = Nothing
End Sub

Private Sub Text6_DblClick()
    
    NORFA$ = TRIM$(Text10)
    If STATORFA%(NORFA$) = 3 Then
       CANTIDADINFORMADA$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad Proyectada", XVALO(Text3)))
       If XVALO(CANTIDADINFORMADA$) < 1 Then Exit Sub
       
       CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
       Call ACTUREGISTRO("ORDEFABR", "canproy", CONDI$, XVALO(CANTIDADINFORMADA$), REGAF&, "NOMESS")
       Call ACTUREGISTRO("ORDEFABR", "STATUORF", CONDI$, 0, REGAF&, "NOMESS")
    End If
End Sub
