VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InfDiaSec 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "REPORTES SECTORES IMPRESIÓN - ACCESORIOS - SEMITERMINADOS - LINEA"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   17805
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   17805
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "InfDiaSec.frx":0000
      Left            =   4320
      List            =   "InfDiaSec.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   34
      ToolTipText     =   "Estados de Orden de Fabricación  (0: Todas - 1: Abiertas - 3: Cerradas)"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H009FBE85&
      Caption         =   "Sector de Talleres (Ubicaciones 30 a50)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   20
      TabIndex        =   31
      Top             =   5640
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid MSF6 
         Height          =   1395
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2461
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H009FBE85&
      Caption         =   "Sector Linea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   0
      TabIndex        =   25
      Top             =   7320
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1400
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Clic Boton Derecho Para Dar X Consumido La Orden de Fabricación Seleccionada"
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2461
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009FBE85&
      Caption         =   "Sector Accesorios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1715
      Left            =   20
      TabIndex        =   7
      Top             =   600
      Width           =   17790
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   1400
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2461
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
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Listar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H009FBE85&
      Caption         =   "Sector Semiterminados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1700
      Left            =   20
      TabIndex        =   12
      Top             =   3960
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   1400
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2461
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      Caption         =   "Sector Impresión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1715
      Index           =   1
      Left            =   20
      TabIndex        =   10
      Top             =   2260
      Width           =   17805
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1400
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2461
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
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   9120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   20
      TabIndex        =   2
      Top             =   0
      Width           =   17790
      Begin VB.CommandButton Command10 
         Caption         =   "Imp. Talleres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14280
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   28
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imp.Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   27
         Top             =   240
         Width           =   1600
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir Linea"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   15960
         TabIndex        =   24
         Top             =   240
         Width           =   1600
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Imp. S.Terminados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12480
         TabIndex        =   23
         Top             =   240
         Width           =   1700
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imp. Impresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10800
         TabIndex        =   22
         Top             =   240
         Width           =   1600
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   765
         TabIndex        =   19
         Text            =   " "
         Top             =   240
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
         Height          =   285
         Left            =   1815
         TabIndex        =   18
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   2565
         TabIndex        =   15
         Text            =   " "
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
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
         Height          =   285
         Left            =   3615
         TabIndex        =   14
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   864
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Ingresar Código de Producto Terminado"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O.Fab.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   5160
         TabIndex        =   30
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5880
         TabIndex        =   29
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   16
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   6810
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -510
         TabIndex        =   3
         Top             =   1035
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "InfDiaSec.frx":000E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "InfDiaSec.frx":0242
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7920
      TabIndex        =   17
      Top             =   465
      Width           =   615
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   225
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnu 
         Caption         =   "Cerrar OF y Cancelar Reserevas de Of Activa"
      End
   End
   Begin VB.Menu mnuCtaListado 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuInfSect 
         Caption         =   "Informe de Sectores"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver "
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuOrigScrap 
         Caption         =   "Informar Origen de Scrap"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTabAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuTablaConfScrap 
         Caption         =   "Configuración Tabla de Scrap"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuIngresoMatCorte 
         Caption         =   "Ingreso de Bobinas a Corte"
      End
      Begin VB.Menu mnuTransfLinea 
         Caption         =   "Transferencia a Linea"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinformarRetiCorte 
         Caption         =   "Informar Retiro de Cortes"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelprinot 
         Caption         =   "Datos Extras Orden de Fabricación Activas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUbic 
         Caption         =   "Ubicaciones"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "InfDiaSec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'Sub Carga_Operaciones(MSF As MSFlexGrid, CIFAB%)
''    SQLX$ = " SELECT * FROM SECOPER where codIconj = " & CIFAB%
''    Set rs = READSET(SQLX$)
'    campos$ = "Cód.Operaciones/F10;" '1
'    campos$ = campos$ + "Descripción/A64;"      '2
'
'    Call CARGA_ENCABEZADO(MSF6, campos$, Me)
'    SQLX$ = "SELECT * FROM CODIGORANGO"
'    SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
'    Dim obj As New RECORXET
'    CAMPOSTABLA$ = "NUMEOPER ,DESCOPER"
'    TABLA$ = "SECOPER"
'    CONDI$ = "codIconj = " & CIFAB%
'    '
'    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
'    Call Carga_MSF_Recordset(RS1, campos$, Me.MSF6, 1)
'    On Error Resume Next
'    Set obj = Nothing
'    RS1.Close
'    Set RS1 = Nothing
'    On Error GoTo 0
'
'End Sub

Sub CARGAGRILLA3(NORFA$, CantAFabricar&)


   Dim AcumulaCantidadTransferida(32767)
   MaxIndex = 0: MaxIndex2 = 0
   
   Dim ctp As New Conteplast
   DEPOCORTE% = XVALO(Control("", "DEPODEST"))
   CIFAB% = CODIFAB%(NORFA$)
   Dim CantAsignada(), CantDisponible()
   SQLX$ = "SELECT CodiElem, CodXElem, UsoBruto, UMED_FORM, COMPONENTE, MODUXBB, CORTE  "
   SQLX$ = SQLX$ + " FROM  formulas WITH(NOLOCK)"
   SQLX$ = SQLX$ + " where codIconj = " & CIFAB%
   SQLX$ = SQLX$ + " order by CodiElem"
   Set rs = READSET(SQLX$)

   i& = 0
   With rs
    Do While Not .EOF
      i& = i& + 1
      CI1% = XVALO(!codielem)
      If MaxIndex2 < CI1% Then
         ReDim Preserve CantAsignada(CI1%)
         ReDim Preserve CantDisponible(CI1%)
         MaxIndex2 = CI1%
      End If
      
      CODI$ = SAFETEXT$(!CodXElem)
      CIRANGO& = CODINT%(CODI$)

      UsoUnit = XVALO(!usobruto)
      CantCportesXUnidad = XVALO(!moduxbb)
      NecesidadDeCortes = CantCportesXUnidad * CantAFabricar&
      DESCROPER$ = SAFETEXT$(!COMPONENTE)
      NUOPER% = ctp.NumeOperXNombre%(CIFAB%, DESCROPER$)
      '
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NUOPER%
      CortesInformados = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
  
      NECESIDAD = XVALO(UsoUnit * CantAFabricar&)
      FALTA = CantTransfDisponible - NECESIDAD
      If FALTA > 0 Then
         FALTA = ""
      Else
         FALTA = Abs(FALTA)
      End If
      CortesFaltan = NecesidadDeCortes - CortesInformados
      If CortesFaltan < 1 Then
        CortesFaltan = ""
      End If
      
      MSF3.Rows = i& + 1
      MSF3.TextMatrix(i&, 1) = CODI$
      MSF3.TextMatrix(i&, 2) = DESCRIT$(CI1%)
      MSF3.TextMatrix(i&, 3) = SAFETEXT$(!UMED_FORM)
      MSF3.TextMatrix(i&, 4) = FORMATNUM$(NECESIDAD, "f12.2")
      
      MSF3.TextMatrix(i&, 7) = FORMATNUM$(NecesidadDeCortes, "f12.2")
      MSF3.TextMatrix(i&, 8) = FORMATNUM$(CortesInformados, "f12.2")
      MSF3.TextMatrix(i&, 9) = FORMATNUM$(CortesFaltan, "f12.2")
      
      MSF3.TextMatrix(i&, 10) = NUOPER%
      MSF3.TextMatrix(i&, 11) = DESCROPER$
      CantAsignada(CI1%) = CantAsignada(CI1%) + NECESIDAD 'GUARDA EN UN VECTOR LA CANTIDAD ASIGNADA
      CI1_ANT% = CI1%
      .MoveNext
    Loop
   End With
   '
   'CARGO EN UN VECTOR LA CANTIDADES TRANSFERIDA A CADA CODIGO RANGO
   
   For i& = 1 To MSF2.Rows - 1
      RangoX$ = MSF2.TextMatrix(i&, 9)
      CODRANGO& = CODINT%(RangoX$)
      CANTI = XVALO(MSF2.TextMatrix(i&, 4))
      AcumulaCantidadTransferida(CODRANGO&) = AcumulaCantidadTransferida(CODRANGO&) + CANTI
   Next i&

   'ASIGNO LAS CANTIDADES A LA NECESIDAD

   For i = 1 To MSF3.Rows - 1
      NECESIDAX = XVALO(MSF3.TextMatrix(i&, 4))
      RangoX$ = MSF3.TextMatrix(i&, 1)
      CODRANGO& = CODINT%(RangoX$)
      If AcumulaCantidadTransferida(CODRANGO&) <= NECESIDAX Then
         AsignoCantidad = AcumulaCantidadTransferida(CODRANGO&)
         FALTA = NECESIDAX - AcumulaCantidadTransferida(CODRANGO&)
      Else
         AsignoCantidad = NECESIDAX
         FALTA = 0
      End If
      MSF3.TextMatrix(i&, 5) = FORMATNUM$(AsignoCantidad, "f12.2")
      MSF3.TextMatrix(i&, 6) = FORMATNUM$(FALTA, "f12.2")
      AcumulaCantidadTransferida(CODRANGO&) = AcumulaCantidadTransferida(CODRANGO&) - AsignoCantidad
   Next i&
   
   
   
   'PRESENTO EN LA GRILLA INFERIOR EL EXCEDENTE
   For i& = 1 To MSF2.Rows - 1
      CODI$ = MSF2.TextMatrix(i&, 9) 'CODIGO rango
      CODRANGO& = CODINT%(CODI$)
      CANTI = XVALO(AcumulaCantidadTransferida(CODRANGO&))
      'SI LA CANTIDAD DISPONIBLE ES MAYOR A LA CANTIDAD ASIGNADA LA MUESTRA
      If CANTI > 0 And CODRANGO& <> CODRANGO_ANT& Then
         J& = J& + 1
         MSF5.Rows = J& + 1
         MSF5.TextMatrix(J&, 1) = CODI$
         MSF5.TextMatrix(J&, 2) = MSF2.TextMatrix(J&, 2)
         MSF5.TextMatrix(J&, 3) = FORMATNUM$(CANTI, "F12.2")
      End If
      AcumulaCantidadTransferida(CODRANGO&) = 0 'lo pongo en 0 para que no repita
    Next i&
End Sub

Sub CargarCompleto()
   MSF5.Rows = 1
   Dim VECTORFECHAIN(), VECTORINGRE(), VECTOREGRE(), VECTORFECHAEG()
   ReDim Preserve VECTORFECHAIN(1), VECTORINGRE(1), VECTOREGRE(1), VECTORFECHAEG(1)
   '
   J& = 0
   'RECORRO LA GRILLA DE LOS INGRESOS
   For i& = 1 To MSF2.Rows - 1
     NORFA$ = TRIM$(MSF2.TextMatrix(i&, 1))
     If NORFA_ANT$ <> NORFA$ Then
       J& = J& + 1
       MSF5.Rows = J& + 1
       MSF5.TextMatrix(J&, 1) = MSF2.TextMatrix(i&, 1)
       MSF5.TextMatrix(J&, 2) = MSF2.TextMatrix(i&, 2)
       MSF5.TextMatrix(J&, 3) = MSF2.TextMatrix(i&, 3)
       MSF5.TextMatrix(J&, 4) = MSF2.TextMatrix(i&, 4)
       MSF5.TextMatrix(J&, 5) = MSF2.TextMatrix(i&, 5)
       MSF5.TextMatrix(J&, 6) = MSF2.TextMatrix(i&, 6)
     End If
     OFNUM& = XVALO(Mid$(NORFA$, 4, 6))
     If UBound(VECTORINGRE) < OFNUM& Then
        ReDim Preserve VECTORINGRE(OFNUM&)
     End If
    
     VECTORINGRE(OFNUM&) = XVALO(VECTORINGRE(OFNUM&)) + XVALO(MSF2.TextMatrix(i&, 10)) 'ACUMULA LOS KG.
     
     NORFA_ANT$ = TRIM$(MSF2.TextMatrix(i&, 1))
     
   Next i&
   '***********
   'RECORRO LA GRILLA DE LOS EGRESOS
   'ACA GUARDO LA ULTIMA FECHA DEL INFORME DE CADA OF Y ACUMULO LOS KG.
   For i& = 1 To MSF3.Rows - 1
     NORFA$ = TRIM$(MSF3.TextMatrix(i&, 1))
     OFNUM& = XVALO(Mid$(NORFA$, 4, 6))

     If UBound(VECTOREGRE) < OFNUM& Then
        ReDim Preserve VECTOREGRE(OFNUM&), VECTORFECHAEG(OFNUM&)
     End If
     VECTORFECHAEG(OFNUM&) = MSF3.TextMatrix(i&, 6)
     
     VECTOREGRE(OFNUM&) = VECTOREGRE(OFNUM&) + XVALO(MSF3.TextMatrix(i&, 11)) 'ACUMULA LOS KG.
     NORFA_ANT$ = TRIM$(MSF3.TextMatrix(i&, 1))
   Next i&
   '
   'AHORA RECORRO EL MSF5 Y SEGUN LA OF APLICO LOS VECTORES
   For i& = 1 To MSF5.Rows - 1
     NORFA$ = TRIM$(MSF5.TextMatrix(i&, 1))
     OFNUM& = XVALO(Mid$(NORFA$, 4, 6))
     MSF5.TextMatrix(i&, 7) = FORMATNUM$(VECTORINGRE(OFNUM&), "F12.2")
     MSF5.TextMatrix(i&, 8) = VECTORFECHAEG(OFNUM&)
     MSF5.TextMatrix(i&, 9) = FORMATNUM$(VECTOREGRE(OFNUM&), "F12.2")
     Diferencia = XVALO(VECTORINGRE(OFNUM&)) - XVALO(VECTOREGRE(OFNUM&))
     MSF5.TextMatrix(i&, 10) = FORMATNUM$(Diferencia, "F12.2")
   Next i&
End Sub

Private Sub Command14_Click()
   Label4 = OrdenFabricacionAbiertas$(1) 'abiertas y cerradas
End Sub

Sub SALDOSENLINEA(Desde%, Hasta%, STAT%, MSF As msFlexGrid, NORFA$)
    If Desde% > 0 Then
      Des$ = FETEXCOR1$(Desde%)
    End If
    If Hasta% > 0 Then
      Has$ = FETEXCOR1$(Hasta%)
    End If
    'TRAE TODO LO QUE ESTA EN LINEA Y QUE TODAVIA TIENEN SALDO (SE CALCULA LA CANTIDAD REALI MENOS EL CONSUMO)
    SQLX$ = "SELECT o.idsubor,FeHorInf, NUMEOPER,DescOper,FeHorInf,COORDUBIC,IDCOORDUBIC,O.Cod_Exte, CORTE,INFOFAB_ID,IDREGPADRE,OperFucionada "
    SQLX$ = SQLX$ + ", O.NumeOrFa,I.FeHorLib,O.StatuOrf,O.FechGen,O.Nume_Cli,I.CodiElem,i.peso,CANTREAL"
    SQLX$ = SQLX$ + " ,I.IdenLote,(I.cantreal - I.CONSUMO) AS SALDO,CONSUMO "
    SQLX$ = SQLX$ + "   FROM infofab I  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON O.IDSUBOR = I.IDSUBOR"
    SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2 WHERE I.INFOFAB_ID = I2.IDREGPADRE) "
    SQLX$ = SQLX$ + " AND idcoordubic = 100 "
    SQLX$ = SQLX$ + " AND (I.cantreal - I.CONSUMO) > 0"
    
    'ESTE AGREGADO ES PARA QUE SOLO TRAIGA LOS QUE YA ESTAN EN FOTODETACIERRE3 Y QUE NO TRAIGA LOS CIERREAS ANTERIORES
    '**** SE LO QUITE POR QUE NO ESTABA TRAYENDO LOS QUE TODAVIA NO TENIAN NINGUN CIERRE 30/09/21 ******
'    SQLX$ = SQLX$ + " and  EXISTS(SELECT 1 FROM FOTODETACIERRE3 F INNER JOIN MOVISTO M ON M.MOVISTO_ID = F.IDMOVISTO AND I.IDSUBOR  = M.IdenLote) "
    '
    If Desde% > 0 Then SQLX$ = SQLX$ + " AND I.FeHorLib >= '" & Des$ & "'"
    If Hasta% > 0 Then SQLX$ = SQLX$ + " AND I.FeHorLib <= '" & Has$ & "'"
    If STAT% > 0 Then SQLX$ = SQLX$ + " AND O.StatuOrf =  " & STAT%
    If TRIM$(NORFA$) <> "" Then SQLX$ = SQLX$ + " AND I.IDSUBOR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CANTREAL > CONSUMO ORDER BY I.IDSUBOR,DESCOPER"
    
    MSF.Rows = 1: REG& = 0
    Set rs = READSET(SQLX$)
      With rs
        Do While Not .EOF
         If XVALO(!CANTREAL) > 0 Then
           IDCOR% = XVALO(!IDCOORDUBIC)
           CORDENADAUBICACION = SAFETEXT$(!COORDUBIC)
           If IDCOR% = 0 Then 'VALIDO PARA VER SI TIENE  POSCIION EN RACK SIGUE SI ES = 0 Y NO TIENE POSICION EN RACK SE VA
             If XVALO(CORDENADAUBICACION) < 1 Then GoTo 80
           End If
           REG& = REG& + 1
           MSF.Rows = REG& + 1
           NC = XVALO(!nume_cli)
           CIELEM% = XVALO(!codielem)
    
           MSF.TextMatrix(REG&, 1) = SAFETEXT$(!idsubor)
           MSF.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FECHGEN))
           MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Cod_Exte)
           MSF.TextMatrix(REG&, 4) = NC
           MSF.TextMatrix(REG&, 5) = rasocli$(NC)
           MSF.TextMatrix(REG&, 6) = FECHATEX$(CVINT(!FeHorLib))
           MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
           MSF.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(!CANTREAL), "F12.0")
           MSF.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(!CONSUMO), "F12.0")
           MSF.TextMatrix(REG&, 10) = FORMATNUM$(XVALO(!SALDO), "F12.0")
           
          End If
80        .MoveNext
        Loop
     End With
     rs.Close
     Set rs = Nothing


End Sub

Private Sub Command10_Click()
    Campos$ = "O.Fab./A12;Fecha/D8;Producto/A14;N°/a0;R.Social/A16;F.Ing./D8;Prod.Intermedio/a15;Cant./F9;Kgs./F6 "
    Campos = Campos$ + ";Ubicación/A30"

    Call CARGA_TABLA_IMPRE(Me.MSF6, Campos$, "REPODIAR", Frame6.Caption, "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then
       FEDESDE$ = "-"
    Else
       FEDESDE$ = TRIM$(FECHATEX$(Desde%))
    End If
    
    If Hasta% < 1 Then
       FEHASTA$ = "-"
    Else
       FEHASTA$ = TRIM$(FECHATEX$(Hasta%))
    End If
   
    FILTX$ = FEDESDE$ & ";" & FEDESDE$ & ";" & Frame6.Caption

    
    Call STDFORM("RDAC-CTP", FILTX$)
End Sub

Private Sub Command2_Click()
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Prod.Intermedio/a24;Cant./F9;Kgs./F9.2"

    Call CARGA_TABLA_IMPRE(Me.MSF3, Campos$, "REPODIAR", Frame1(1).Caption, "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then
       FEDESDE$ = "-"
    Else
       FEDESDE$ = TRIM$(FECHATEX$(Desde%))
    End If
    
    If Hasta% < 1 Then
       FEHASTA$ = "-"
    Else
       FEHASTA$ = TRIM$(FECHATEX$(Hasta%))
    End If
   
    FILTX$ = FEDESDE$ & ";" & FEDESDE$ & ";" & Frame1(1).Caption
    Call STDFORM("RDAC-CTP", FILTX$)

End Sub





Private Sub Command3_Click()
    Command3.Enabled = False
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Prod.Intermedio/a24;Cant./F9;Kgs./F9.2"

    Call CARGA_TABLA_IMPRE(Me.MSF2, Campos$, "REPODIAR", Frame3.Caption, "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then
       FEDESDE$ = "-"
    Else
       FEDESDE$ = TRIM$(FECHATEX$(Desde%))
    End If
    
    If Hasta% < 1 Then
       FEHASTA$ = "-"
    Else
       FEHASTA$ = TRIM$(FECHATEX$(Hasta%))
    End If
   
    FILTX$ = FEDESDE$ & ";" & FEDESDE$ & ";" & Frame3.Caption
    Call STDFORM("RDAC-CTP", FILTX$)
    Command3.Enabled = True
End Sub


Private Sub Command8_Click()
    Campos$ = "O.Fab./A12;Fecha/D8;Producto/A14;N°/a0;R.Social/A16;F.Ing./D8;Prod.Intermedio/a15;Cant./F9;Kgs./F6 "
    Campos = Campos$ + ";Ubicación/A30"

    Call CARGA_TABLA_IMPRE(Me.MSF4, Campos$, "REPODIAR", Frame5.Caption, "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then
       FEDESDE$ = "-"
    Else
       FEDESDE$ = TRIM$(FECHATEX$(Desde%))
    End If
    
    If Hasta% < 1 Then
       FEHASTA$ = "-"
    Else
       FEHASTA$ = TRIM$(FECHATEX$(Hasta%))
    End If
   
    
    FILTX$ = FEDESDE$ & ";" & FEDESDE$ & ";" & Frame5.Caption

    Call STDFORM("RDAC-CTP", FILTX$)

End Sub




Private Sub Command4_Click()

    Campos$ = "O.Fabric./A12;Referencia/A48;Q.Proy./F9.1;Q.Aprob/F9.1;ID_ORFA/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Ordenes de Fabricación Pendientes de Cierre"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_OF_ABIERTAS_2(1, 1)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Width = 12500
    FRMZELECHO.Show 1
    Label4 = RESP_ZELE_VECT(1)

End Sub

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text4.TEXT = VDEV$
  End If
End Sub

Private Sub Command6_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text1.TEXT = VDEV$
  End If
End Sub

 Sub Command7_Click()
    MODO% = 0
    Screen.MousePointer = 11
    MSF2.Rows = 1: MSF3.Rows = 1: MSF5.Rows = 1
    
    'MODOORDEN% 1 ABIERAS, 3 CERRADAS,  0 TODAS
    Desde% = FECHANUM(TRIM$(Text1))
    Hasta% = FECHANUM(TRIM$(Text4))
    Call CargarReporteDiario(Desde%, Hasta%, XVALO(Combo1.TEXT), 2, 2, MSF2, Label4)
    Call CargarReporteDiario(Desde%, Hasta%, XVALO(Combo1.TEXT), 1, 1, MSF3, Label4)
    Call CargarReporteDiario(Desde%, Hasta%, XVALO(Combo1.TEXT), 3, 40, MSF4, Label4)
    Call CargarReporteDiario(Desde%, Hasta%, XVALO(Combo1.TEXT), 30, 50, MSF6, Label4) '30 AL 50 REFIERES A UBICACIONES DE RACK
    'Call CargarReporteDiario(Desde%, Hasta%, XVALO(Combo1.TEXT), 100, 100, MSF5, Label4)
    Call SALDOSENLINEA(Desde%, Hasta%, XVALO(Combo1.TEXT), MSF5, Label4)
    Screen.MousePointer = 1
    
End Sub


Private Sub Command9_Click()
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Prod.Intermedio/a24;Cant./F9;Kgs./F9.2"


    Call CARGA_TABLA_IMPRE(Me.MSF5, Campos$, "REPODIAR", Frame4.Caption, "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then
       FEDESDE$ = "-"
    Else
       FEDESDE$ = TRIM$(FECHATEX$(Desde%))
    End If
    
    If Hasta% < 1 Then
       FEHASTA$ = "-"
    Else
       FEHASTA$ = TRIM$(FECHATEX$(Hasta%))
    End If
   
    
    FILTX$ = FEDESDE$ & ";" & FEDESDE$ & ";" & Frame4.Caption
    
    Call STDFORM("RDAC-CTP", FILTX$)

End Sub

Private Sub Form_Load()
    '
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Prod.Intermedio/a24;Cant./F9;Kgs./F9"
    Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
    Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
    Call CARGA_ENCABEZADO(MSF5, Campos$, Me)
    
    Campos = Campos$ + ";Ubicación/A32"
    Call CARGA_ENCABEZADO(MSF4, Campos$, Me)
    Call CARGA_ENCABEZADO(MSF6, Campos$, Me)
    Combo1.Clear
    Combo1.AddItem "0"
    Combo1.AddItem "1"
    Combo1.AddItem "3"
    
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Prod.Intermedio/a24;Cant./F9;Consumo/F9;Saldo/F9"
    Call CARGA_ENCABEZADO(MSF5, Campos$, Me)

'********************************************************************************
End Sub


Private Sub Frame1_DblClick(Index As Integer)
   If Frame1(1).Height < 2200 Then
      Frame1(1).Height = 7575
      Frame1(1).Top = 600
      MSF3.Height = 7170
      Frame1(1).ZOrder 0
   Else
      Frame1(1).Height = 2100
      Frame1(1).Top = 2760
      MSF3.Height = 2850
   End If
End Sub

Private Sub Frame3_DblClick()
   If Frame3.Height < 3000 Then
      Frame3.Height = 7575
      MSF2.Height = 7170
      Frame3.ZOrder 0
   Else
      Frame3.Height = 2175
      Frame3.Top = 600
      MSF2.Height = 2250
   End If
End Sub

Private Sub Frame4_DblClick()

   If Frame4.Height < 2200 Then
      Frame4.Height = 7575
      Frame4.Top = 600
      MSF5.Height = 7170
      Frame4.ZOrder 0
   Else
      Frame4.Height = 2100
      Frame4.Top = 6840
      MSF5.Height = 1770
   End If

End Sub

Private Sub Frame5_DblClick()
   If Frame5.Height < 2200 Then
      Frame5.Height = 7575
      Frame5.Top = 600
      MSF4.Height = 7170
      Frame5.ZOrder 0
   Else
      Frame5.Height = 2100
      Frame5.Top = 4800
      MSF4.Height = 1770
   End If
End Sub

Private Sub Label1_DblClick()
   Label1 = ""
End Sub


Private Sub Label4_Change()
'   MSF1.Rows = 1
   MSF2.Rows = 1
   MSF3.Rows = 1
   MSF5.Rows = 1
   If NORFA$ = "" Then Exit Sub
'   Label2 = RESP_ZELE_VECT(2)
'   CI1% = CODIFAB%(NORFA$)
'   TXTCODINT = CI1%
'   Text10 = CODEXT$(CI1%)
'   Label3 = DESCRIT0$(CI1%)
'   CIFAB% = CODINT%(Text10)
'   NROPED& = NroPedidoSegunOf&(NORFA$)
'   NC = cliente_SEGUN_PEDIDO(NROPED&)
'   Text6 = NC
'   Text12 = NROPED&
'   CANTI = CantiProyectadaOF(NORFA$)
'   Text14 = CantiProyectadaOF(NORFA$)
'   CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CIFAB%
'   Text13 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
'   'Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$)
'   Dim CTP As New Conteplast
'   DEPOCORTE% = XVALO(Control("", "DEPODEST"))
'   Call CTP.CargaMatTranfXOrdenParaCierre(MSF2, TRIM$(NORFA$), DEPOCORTE%, 0, 1)
'   MSF2.COL = 9 'RANGO
'   MSF2.Sort = 3
'   Set CTP = Nothing
'
''   Call cargarTransfDeOf_DeSalidasDeposito(NORFA$, XVALO(Me.lbldepogeneral), MSF2)
'   Call CARGA_ESTRUCTURA_FICHA(CIFAB%, CANTI, MSF1)
'   Call CARGA_INFORMECONSUMO
'   Call Carga_Operaciones(MSF6, CIFAB%)
End Sub
Sub CargarIngresos(NORFA1$, NORFA2$, Desde%, Hasta%, MODO%)
'Sub CargaMatTranfXOrdenParaCierre(MSF1 As MSFlexGrid, NORFA$, DEPOCORTE%, Optional FILTROOPERACION%, Optional AGREGARCODIGORANGO%)
   'SE CONSIDERA TAMBIEN LO TRANSFERIDO AL DEPOSITO DE ESLINGA
   

  MSF2.Rows = 1
  DEPOESLINGA% = XVALO(Control("CONTEPLA", "DEPOESLI"))
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  AGREGARCODIGORANGO% = 1
  'ES CargaMatTranfXOrden MODIFICADO PARA PRESENTAR LA PANTALLA DE EVALUACION PARA CIERRE DE FABRICACION
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION,DOSECLAR"
  SQLX$ = SQLX$ + "  FROM MOVISTO WITH(NOLOCK) "
  SQLX$ = SQLX$ + "  WHERE DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + "  AND CODIMOVI = 'TR' and LEFT(DOSECLAR,3) = 'OF-'"
  If FILTROOPERACION% > 0 Then SQLX$ = SQLX$ + " AND ID_OPERACION = " & FILTROOPERACION%
  
  'FILTROS DE PANTALLA
  If XVALO(Mid$(NORFA1$, 4)) > 0 And XVALO(Mid$(NORFA2$, 4)) > 0 Then
     SQLX$ = SQLX$ + "  AND CAST(substring(DOSECLAR,4,6) AS INT) >= " & XVALO(Mid$(NORFA1$, 4))
     SQLX$ = SQLX$ + "  AND CAST(substring(DOSECLAR,4,6) AS INT) <= " & XVALO(Mid$(NORFA2$, 4))
  End If
  '
  If Desde% > 0 And Hasta% > 0 Then
     SQLX$ = SQLX$ + "  AND FECHMOV >= '" & FETEXCOR1$(Desde%) & "'"
     SQLX$ = SQLX$ + "  AND FECHMOV <= '" & FETEXCOR1$(Hasta%) & "'"
  End If
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION,DOSECLAR"
  SQLX$ = SQLX$ + " having  SUM(CANTINGRE-CANTSALID) >0 "
  SQLX$ = SQLX$ + " ORDER BY DOSECLAR,ID_OPERACION,COD_INTE "
  MSF3.Rows = 1: REG& = 0
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
    'MODOORDEN% 1 ABIERAS, 3 CERRADAS,  0 TODAS
      If XVALO(!Resultado) > 0 Then
       NORFA$ = SAFETEXT$(!DOSECLAR)
       St% = STATORFA%(NORFA$)
       If MODO% = 1 Then
       '  CASO SOLO ABIERTAS
          If MODO% = 1 And St% <= 1 Then
             GoTo 20
          Else
             GoTo 80
          End If
       End If
       If MODO% = 3 Then
          '  CASO SOLO CERRADAS
            If MODO% = 3 And St% = 3 Then
              GoTo 20
            Else
              GoTo 80
            End If
       End If
       
20     REG& = REG& + 1
       MSF2.Rows = REG& + 1
       OFNUM = Mid$(NORFA$, 4, 6)
       FEGEN% = FECHGENORFA%(NORFA$)
       CIFAB% = CODIFAB%(NORFA$)
       NC = NroClienteSegunOf&(NORFA$)
       CI1% = XVALO(!cod_inte)
       PRIMERINGRESO% = FechaPrimerTransfMaterialesSegunOf%(NORFA$, CI1%)
       CIRANG% = CODIGORANGO&(CI1%)
       CANTMETROSBOBINA = XVALO(!Resultado) 'TOTAL DE MTS
       'esoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       
       PesoPorMetro# = PESMETRO(CI1%)  'TRAIGO EL DE LA FORMULA POR QUE NO PUEDO CALCULARLO EN UNA CONSULTA AGRUPADA
       SaldoBobinaEnKilos# = CANTMETROSBOBINA * PesoPorMetro#
       MSF2.TextMatrix(REG&, 1) = NORFA$
       MSF2.TextMatrix(REG&, 2) = FECHATEX$(FEGEN%)
       MSF2.TextMatrix(REG&, 3) = CODEXT$(CIFAB%)
       MSF2.TextMatrix(REG&, 4) = NC
       MSF2.TextMatrix(REG&, 5) = rasocli$(NC)
       MSF2.TextMatrix(REG&, 6) = FECHATEX$(PRIMERINGRESO%)
       MSF2.TextMatrix(REG&, 7) = CODEXT$(CIRANG%)
       MSF2.TextMatrix(REG&, 8) = CODEXT$(CI1%)
       MSF2.TextMatrix(REG&, 9) = TRIM$(FORMATNUM$(CANTMETROSBOBINA, "F12.0"))
       MSF2.TextMatrix(REG&, 10) = TRIM$(FORMATNUM$(SaldoBobinaEnKilos#, "F12.0"))
       'MSF2.TextMatrix(REG&, 11) = OFNUM
      End If
80     .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
End Sub

 Sub CargarReporteDiario(Desde%, Hasta%, MODO%, UbicDesde%, UbicHasta%, MSF As msFlexGrid, Optional NORFA$)
 
    'ES CargaMatTranfXOrden MODIFICADO PARA PRESENTAR LA PANTALLA DE EVALUACION PARA CIERRE DE FABRICACION
    SQLX$ = "SELECT o.idsubor,FeHorInf, NUMEOPER,DescOper,CANTREAL,FeHorInf,COORDUBIC,IDCOORDUBIC,O.Cod_Exte"
    SQLX$ = SQLX$ + ", CORTE,INFOFAB_ID,IDREGPADRE,OperFucionada,O.NumeOrFa,I.FeHorLib,O.StatuOrf,O.FechGen"
    SQLX$ = SQLX$ + ",O.Nume_Cli,I.CodiElem,i.peso"
    SQLX$ = SQLX$ + " FROM INFOFAB I with(NOLOCK)   INNER JOIN ORDEFABR O ON O.IDSUBOR = I.IDSUBOR"
    'SQLX$ = SQLX$ + " Where (IDREGPADRE < 1 or IDREGPADRE is null)" ' AND IDCOORDUBIC < 100 "
    SQLX$ = SQLX$ + " WHERE NOT EXISTS"
    SQLX$ = SQLX$ + " (SELECT * FROM INFOFAB AS I2"
    SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE and I.IdRegPadre <> NULL)"
    'QUE POR EJMPLOS LOS QUE BUSCO VAN A TENER '30;4;1' ENTONCES LA TERECER POSICION DEBERIA SER = ;
    'Y QUE EL LEFT 2 DEBE SER UN VALOR >= 30 Y <= 50
    
    If UbicDesde% = 3 Then ' CASO DE SEMITERMINADOS
        'ESTE SIGUIENTE PARRAFO ES PARA DEJAR FUERA LAS POSICIONES DE RACK DE TALLERES DEL 30 A 50
'        SQLX$ = SQLX$ + " AND NOT  EXISTS (SELECT * FROM INFOFAB AS I3  "
'        SQLX$ = SQLX$ + " WHERE substring(i.COORDUBIC,3,1) = ';'"
'        SQLX$ = SQLX$ + " AND convert(int,(left(i.COORDUBIC,2))) >= 30"
'        SQLX$ = SQLX$ + " AND  convert(int,(left(i.COORDUBIC,2))) <= 50)"
        
        SQLX$ = SQLX$ + " AND ("
        SQLX$ = SQLX$ + " IDCOORDUBIC >= " & UbicDesde%
        SQLX$ = SQLX$ + " AND IDCOORDUBIC <= " & UbicHasta% '& ")" antes cerraba el parentesis aca
        SQLX$ = SQLX$ + " or IDCOORDUBIC = 0" & ")"                 'ahora lo cierra aqui
'        SQLX$ = SQLX$ + ")"
    ElseIf UbicDesde% = 30 Then 'CASO TALLERES UBICACIONES RESERVADAS 30 A 50
    'QUE POR EJMPLOS LOS QUE BUSCO VAN A TENER '30;4;1' ENTONCES LA TERECER POSICION DEBERIA SER = ;
    'Y QUE EL LEFT 2 DEBE SER UN VALOR >= 30 Y <= 50
        SQLX$ = SQLX$ + " AND IDCOORDUBIC < 1  "
        SQLX$ = SQLX$ + " AND substring(COORDUBIC,3,1) = ';' "
        SQLX$ = SQLX$ + " AND convert(int,(left(COORDUBIC,2))) >= " & UbicDesde%
        SQLX$ = SQLX$ + " AND  convert(int,(left(COORDUBIC,2))) <= " & UbicHasta%
        
    Else
        SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & UbicDesde%
        SQLX$ = SQLX$ + " AND IDCOORDUBIC <= " & UbicHasta%
    End If



  If Desde% > 0 And Hasta% > 0 Then
     SQLX$ = SQLX$ + "  AND FeHorLib >= '" & FETEXCOR1$(Desde%) & "'"
     SQLX$ = SQLX$ + "  AND FeHorLib <= '" & FETEXCOR1$(Hasta%) & "'"
  End If
  If MODO% > 0 Then
    'MODOORDEN% 1 ABIERAS, 3 CERRADAS,  0 TODAS
    SQLX$ = SQLX$ + "  AND StatuOrf = " & MODO%
  End If
  '
  If TRIM$(NORFA$) <> "" Then
     SQLX$ = SQLX$ + " AND I.IDSUBOR = '" & TRIM$(NORFA$) & "'"
  End If
  
  SQLX$ = SQLX$ + " ORDER BY o.idsubor,I.FeHorInf"
  
  MSF.Rows = 1: REG& = 0
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
     If XVALO(!CANTREAL) > 0 Then
       IDCOR% = XVALO(!IDCOORDUBIC)
       CORDENADAUBICACION = SAFETEXT$(!COORDUBIC)
       If IDCOR% = 0 Then 'VALIDO PARA VER SI TIENE  POSCIION EN RACK SIGUE SI ES = 0 Y NO TIENE POSICION EN RACK SE VA
         If XVALO(CORDENADAUBICACION) < 1 Then GoTo 80
       End If
       REG& = REG& + 1
       MSF.Rows = REG& + 1
       NC = XVALO(!nume_cli)
       CIELEM% = XVALO(!codielem)
       PesoPorMetro# = PESMETRO(CIELEM%) 'TRAIGO EL DE LA FORMULA POR QUE NO PUEDO CALCULARLO EN UNA CONSULTA AGRUPADA

       MSF.TextMatrix(REG&, 1) = SAFETEXT$(!idsubor)
       MSF.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FECHGEN))
       MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Cod_Exte)
       MSF.TextMatrix(REG&, 4) = NC
       MSF.TextMatrix(REG&, 5) = rasocli$(NC)
       MSF.TextMatrix(REG&, 6) = FECHATEX$(CVINT(!FeHorLib))
'       MSF.TextMatrix(REG&, 7) = XVALO(!NUMEOPER)
       MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
       MSF.TextMatrix(REG&, 8) = XVALO(!CANTREAL)
'       TOTMET# = XVALO(!CANTREAL) * XVALO(!CORTE)
'       MSF.TextMatrix(REG&, 10) = FORMATNUM$(TOTMET#, "F12.2")
       MSF.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(!Peso), "F12.2")
       If MSF.Name = "MSF4" Or MSF.Name = "MSF6" Then
            CORDENADAUBICACION = SAFETEXT$(!COORDUBIC)
            COORDENADA = ""
            If XVALO(CORDENADAUBICACION) > 0 Then
              COORDENADA = CORDENADAUBICACION
              CONDI$ = "CONVERT(int,campo1) = " & XVALO(CORDENADAUBICACION)
              CONDI$ = CONDI$ + " AND LISTA = 'CONFUBIX'"
              CORDENADAUBICACION = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
            End If
            MSF.TextMatrix(REG&, 10) = CORDENADAUBICACION & " " & COORDENADA
       End If
      End If
80    .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
End Sub





Private Sub Label4_DblClick()
   Label4 = ""
End Sub

Private Sub mnuinformarRetiCorte_Click()
INFOCORTE.Show 1
End Sub

Private Sub mnuIngresoMatCorte_Click()
   INGREBOB.Show 1

End Sub

Private Sub mnuTabAux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)
End Sub

Private Sub mnuTransfLinea_Click()
  TRAFALINEA.Show 1
End Sub

Private Sub mnuUbic_Click()
   INFUBIPT.NOUBICAR = 1
   INFUBIPT.Show 1
End Sub

Private Sub Text14_Change()
   TXTCANTICALCULO = Text14
End Sub

Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub



Private Sub MSF5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      REG& = MSF5.MouseRow
      If MSF5.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF5, REG&)
            NORFA$ = TRIM$(MSF5.TextMatrix(REG&, 1))
            If DERACCE%("DARXCUMP", "DAR POR CUMPLIDO CONSUMOS DE ORDEN DE FABRICACION") < 2 Then Exit Sub
            OPX% = COMALTER("Realmente Desea Dar por Cunsumido Todo el Material de Esta Orden: " & NORFA$ & "\\Cancelar\Sí, Consumir")
            If OPX% = 2 Then
              
              STAT% = STATORFA%(NORFA$)
              If STAT% <> 3 Or STAT% <> 9 Then
                 Call COMUNI("Atención ! La Orden de Fabricación Seleccionada No Se Encuentra Cerrada"): Exit Sub
              End If
              Call DarXCumplidoConsumoDeOF(NORFA$)
              Call Command7_Click
            End If
        End If
      End If
    End If

     NORFA$ = TRIM$(MSF5.TextMatrix(i&, 1))
     OFNUM& = XVALO(Mid$(NORFA$, 4, 6))


End Sub
Sub DarXCumplidoConsumoDeOF(NORFA$)
  'IGUALA LA CANATIDAD REAL RESERVADA LA ASIGNA AL CAMPO CONSUMO
  SQLX$ = "UPDATE INFOFAB SET CONSUMO = CANTREAL WHERE  IDSUBOR = '" & NORFA$ & "'"
  FLEXCONN.Execute (SQLX$)
End Sub
