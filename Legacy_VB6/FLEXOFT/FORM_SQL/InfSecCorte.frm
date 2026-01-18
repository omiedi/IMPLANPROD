VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InfSecCorte 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Reporte Sector Corte"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   17805
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   17805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H009FBE85&
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   50
      TabIndex        =   10
      Top             =   1320
      Width           =   17790
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2250
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Materiales Transferidos a Sector de Corte"
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   3969
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
      Height          =   855
      Left            =   7680
      TabIndex        =   31
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H009FBE85&
      Caption         =   "Completo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   45
      TabIndex        =   15
      Top             =   7200
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1545
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Muestra la Diferencia entre lo ingresado en kg. y los Cortes Primarios Informados."
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2725
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   16680
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   1
      Left            =   45
      TabIndex        =   13
      Top             =   3960
      Width           =   17805
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2850
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Cortes informados Primarios (Es decir no Considera Sub-Procesos)"
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   5027
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
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   50
      TabIndex        =   3
      Top             =   0
      Width           =   17790
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir Completo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12120
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Imprimir Egresos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10680
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir Ingresos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9240
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   1000
         Left            =   5640
         TabIndex        =   27
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton Todas 
            Appearance      =   0  'Flat
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   650
            Width           =   1500
         End
         Begin VB.OptionButton Pendientes 
            Appearance      =   0  'Flat
            Caption         =   "Pendientes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   415
            Width           =   1500
         End
         Begin VB.OptionButton Cerradas 
            Appearance      =   0  'Flat
            Caption         =   "Cerradas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   200
            Width           =   1500
         End
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
         Left            =   4125
         TabIndex        =   25
         Text            =   " "
         Top             =   360
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
         Left            =   5175
         TabIndex        =   24
         Top             =   360
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
         Left            =   4125
         TabIndex        =   21
         Text            =   " "
         Top             =   720
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
         Left            =   5175
         TabIndex        =   20
         Top             =   720
         Width           =   175
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
         Left            =   2760
         TabIndex        =   17
         Top             =   720
         Width           =   175
      End
      Begin VB.CommandButton Command3 
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
         Left            =   2760
         TabIndex        =   2
         Top             =   360
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   16680
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
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
         Left            =   3480
         TabIndex        =   26
         Top             =   465
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Left            =   3480
         TabIndex        =   22
         Top             =   825
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
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
         Left            =   -480
         TabIndex        =   19
         Top             =   795
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Left            =   870
         TabIndex        =   18
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   720
         Width           =   1890
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
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   6810
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
         Left            =   870
         TabIndex        =   6
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   360
         Width           =   1890
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
         TabIndex        =   5
         Top             =   1035
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
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
         Left            =   -480
         TabIndex        =   4
         Top             =   435
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "InfSecCorte.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "InfSecCorte.frx":0234
      TabIndex        =   9
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
      TabIndex        =   23
      Top             =   465
      Width           =   615
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   12
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
      Begin VB.Menu mnuContConsumo 
         Caption         =   "Control de Consumos x OF"
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
      End
      Begin VB.Menu mnuUbic 
         Caption         =   "Ubicaciones"
      End
   End
End
Attribute VB_Name = "InfSecCorte"
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
      CantCportesXUnidad = XVALO(!MODUXBB)
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
   Dim VECTORINGRE(), VECTOREGRE(), VECTORFECHAEG()
   ReDim Preserve VECTORINGRE(1), VECTOREGRE(1), VECTORFECHAEG(1)
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
   Label9.Caption = "": Label10.Caption = ""
   For i& = 1 To MSF5.Rows - 1
     NORFA$ = TRIM$(MSF5.TextMatrix(i&, 1))
     OFNUM& = XVALO(Mid$(NORFA$, 4, 6))
     If UBound(VECTORINGRE) < OFNUM& Then ReDim Preserve VECTORINGRE(OFNUM&)
     If UBound(VECTORFECHAEG) < OFNUM& Then ReDim Preserve VECTORFECHAEG(OFNUM&)
     If UBound(VECTOREGRE) < OFNUM& Then ReDim Preserve VECTOREGRE(OFNUM&)
        
     MSF5.TextMatrix(i&, 7) = FORMATNUM$(VECTORINGRE(OFNUM&), "F12.2")
     MSF5.TextMatrix(i&, 8) = VECTORFECHAEG(OFNUM&)
     MSF5.TextMatrix(i&, 9) = FORMATNUM$(VECTOREGRE(OFNUM&), "F12.2")
     diferencia = XVALO(VECTORINGRE(OFNUM&)) - XVALO(VECTOREGRE(OFNUM&))
     MSF5.TextMatrix(i&, 10) = FORMATNUM$(diferencia, "F12.2")
     Label9.Caption = FORMATNUM$(XVALO(Label9.Caption) + XVALO(MSF5.TextMatrix(i&, 7)), "F12.1")
     Label10.Caption = FORMATNUM$(XVALO(Label10.Caption) + XVALO(MSF5.TextMatrix(i&, 9)), "F12.1")
   Next i&
   
End Sub

Private Sub Command14_Click()
   Label4 = OrdenFabricacionAbiertas$(1) 'abiertas y cerradas
End Sub

Private Sub Command2_Click()
    Campos$ = "O.Fab./A11;F.O.Fab/D8;Producto/A14;Cte.N°/F6;R.Social/A24;F.Ing./D8;Rgo./a4;Cód.Mat./A16;Mts./F6;Kgs./F6;Lote/A12"

    Call CARGA_TABLA_IMPRE(Me.MSF2, Campos$, "INGREMAT", "INFORME DE INGRESOS", "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    If Desde% < 1 Then Desde% = 100
    If Hasta% < 1 Then Hasta% = HOY(HO$) + 30

    NORFA1$ = TRIM$(Label4): If NORFA1$ = "" Then NORFA1$ = "-"
    NORFA2$ = TRIM$(Label1): If NORFA2$ = "" Then NORFA2$ = "-"
    If NORFA2$ = "" Then NORFA2$ = "-"
    FILTX$ = NORFA1$ & ";" & NORFA2$ & ";" & TRIM$(FECHATEX$(Desde%)) & ";" & TRIM$(FECHATEX$(Hasta%)) & ";" & TRIM$(Label9.Caption)
    Call STDFORM("INGR-CTP", FILTX$)

End Sub

Private Sub Command3_Click()
    '
    MODO% = 0
    If Cerradas.Value = True Then MODO% = 3
    If Pendientes.Value = True Then MODO% = 1
    Label4 = OrdenFabricacionAbiertas$(MODO%) 'abiertas y cerradas
    '
    
    Exit Sub
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
    End If
    '
    CIRANG = CODIGORANGO&(CIXI%)
    CONDI$ = "COD_INTE = " & CIXI%
    CONDI$ = CONDI$ + " AND STATUORF < 3"
    REG& = CantRegistros&("ORDEFABR", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Ordenes de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!ECORDEF1/Ordenes de Fabricación")
    If VALACT1$("!ECORDEF1") <> "" Then
     Label4 = VALACT1$("!ECORDEF1")
     Label2 = VALACT2$("!ECORDEF1")
    End If
    '
End Sub




Private Sub Command8_Click()
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F8;R.Social/A24;F.Eg./D8;N.OPER/A0;Prod Interm./a16;Cant/f8;Mts./F9.2;Kgs./F9.2;Sector/A12"

    Call CARGA_TABLA_IMPRE(Me.MSF3, Campos$, "EGREMAT", "INFORME DE EGRESOS", "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    NORFA1$ = TRIM$(Label4): If NORFA1$ = "" Then NORFA1$ = "-"
    NORFA2$ = TRIM$(Label1): If NORFA2$ = "" Then NORFA2$ = "-"
    If Desde% < 1 Then Desde% = 100
    If Hasta% < 1 Then Hasta% = HOY(HO$) + 30

    FILTX$ = NORFA1$ & ";" & NORFA2$ & ";" & TRIM$(FECHATEX$(Desde%)) & ";" & TRIM$(FECHATEX$(Hasta%)) & ";" & TRIM$(Label10.Caption)
    Call STDFORM("EGRE-CTP", FILTX$)

End Sub




Private Sub Command4_Click()
    MODO% = 0
    If Cerradas.Value = True Then MODO% = 3
    If Pendientes.Value = True Then MODO% = 1
    Label1 = OrdenFabricacionAbiertas$(MODO%) 'abiertas y cerradas

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

Private Sub Command7_Click()
    MODO% = 0
    Screen.MousePointer = 11
    MSF2.Rows = 1: MSF3.Rows = 1: MSF5.Rows = 1
    
    If Cerradas.Value = True Then MODO% = 3
    If Pendientes.Value = True Then MODO% = 1
    'MODOORDEN% 1 ABIERAS, 3 CERRADAS,  0 TODAS
    
    Call CargarIngresos(Label4, Label1, FECHANUM(Text1), FECHANUM(Text4), MODO%)
    Call CargarEgresos(Label4, Label1, FECHANUM(Text1), FECHANUM(Text4), MODO%)
    Call CargarCompleto
    Screen.MousePointer = 1
    
End Sub

Private Sub Command9_Click()
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F8;R.Social/A24;F.Ing./D8"                              '  1
    Campos$ = Campos$ + ";Kg/F10.2;F.Egreso/d8;Kg/F10.2"                    '  2
    Campos$ = Campos$ + ";Diferencia./f12.2"

    Call CARGA_TABLA_IMPRE(Me.MSF5, Campos$, "INEGEMAT", "INFORME DE INGRESOS Y EGRESOS", "TABLA_IMPRE", , 1)
    '*****************************************************'
    Desde% = FECHANUM(Text1.TEXT)
    Hasta% = FECHANUM(Text4.TEXT)
    NORFA1$ = TRIM$(Label4): If NORFA1$ = "" Then NORFA1$ = "-"
    NORFA2$ = TRIM$(Label1): If NORFA2$ = "" Then NORFA2$ = "-"
    
    If NORFA2$ = "" Then NORFA2$ = "-"
    FILTX$ = NORFA1$ & ";" & NORFA1$ & ";" & TRIM$(FECHATEX$(Desde%)) & ";" & TRIM$(FECHATEX$(Hasta%))
    Call STDFORM("INEG-CTP", FILTX$)

End Sub

Private Sub Form_Load()
    '
    Todas.Value = True
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F6;R.Social/A24;F.Ing./D8;Cód.Rgo./a10;Cód.Mat./A16;Mts./F9;Kgs./F9;Lote/A12"
    Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
    
    
    Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F8;R.Social/A24;F.Eg./D8;N.OPER/A0;Prod Interm./a16;Cant/f8;Mts./F9;Kgs./F9;Sector/A28;Rango/A12"
'  O.Fab./A12                   '  1
'  F.O.Fab/D8                   '  2
'  Producto/A16                 '  3
'  Cte.N°/F8                    '  4
'  R.Social/A24                 '  5
'  F.Eg./D8                     '  6
'  N.OPER/A0                    '  7
'  Prod Interm./a16             '  8
'  Cant/f8                      '  9
'  Mts./F9                      '  10
'  Kgs./F9                      '  11
'  Sector/A28                   '  12
    Call CARGA_ENCABEZADO(MSF3, Campos$, Me)

'********************************************************************************
  Campos$ = "O.Fab./A12;F.O.Fab/D8;Producto/A16;Cte.N°/F8;R.Social/A24;F.Ing./D8"                              '  1
  Campos$ = Campos$ + ";Kg/F10.2;F.Egreso/d8;Kg/F10.2"                    '  2
  Campos$ = Campos$ + ";Diferencia./f12.2"
  '
'  O.Fab./A12                   '  1
'  F.O.Fab/D8                   '  2
'  Producto/A16                 '  3
'  Cte.N°/F8                    '  4
'  R.Social/A24                 '  5
'  F.Ing./D8                    '  6
'  Kg/F10.2                     '  7
'  F.Egreso/d8                  '  8
'  Kg/F10.2                     '  9
'  Diferencia./f12.2            '  10
 
  
  Call CARGA_ENCABEZADO(MSF5, Campos$, Me)
'********************************************************************************
End Sub


Private Sub Frame1_DblClick(Index As Integer)
   If Frame1(1).Height < 3400 Then
      Frame1(1).Height = 7575
      Frame1(1).Top = 1320
      MSF3.Height = 7170
      Frame1(1).ZOrder 0
   Else
      Frame1(1).Height = 3255
      Frame1(1).Top = 3960
      MSF3.Height = 2850
   End If
End Sub

Private Sub Frame3_DblClick()
   If Frame3.Height < 3000 Then
      Frame3.Height = 7575
      MSF2.Height = 7170
      Frame3.ZOrder 0
   Else
      Frame3.Height = 2655
      MSF2.Height = 2250
   End If
End Sub

Private Sub Frame5_DblClick()
   If Frame5.Height < 2000 Then
      Frame5.Height = 7575
      Frame5.Top = 1320
      MSF5.Height = 7170
      Frame5.ZOrder 0
   Else
      Frame5.Height = 1815
      Frame5.Top = 7200
      MSF5.Height = 2850
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
   
  Dim ctp As New Conteplast
  MSF2.Rows = 1
  DEPOESLINGA% = XVALO(Control("CONTEPLA", "DEPOESLI"))
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  AGREGARCODIGORANGO% = 1
  'ES CargaMatTranfXOrden MODIFICADO PARA PRESENTAR LA PANTALLA DE EVALUACION PARA CIERRE DE FABRICACION
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION,DOSECLAR,FECHMOV,IDENLOTE"
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
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION,DOSECLAR,FECHMOV,IDENLOTE"
  SQLX$ = SQLX$ + " having  SUM(CANTINGRE-CANTSALID) > 0 "
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
       'EXCEPTUAR BOLSAS FAMILIA = BO
       CI1% = XVALO(!cod_inte)
       If CODFAMIL$(CI1%) = "BO" Then GoTo 80
       
20     REG& = REG& + 1
       MSF2.Rows = REG& + 1
       OFNUM = Mid$(NORFA$, 4, 6)
       FEGEN% = FECHGENORFA%(NORFA$)
       CIFAB% = CODIFAB%(NORFA$)
       NC = NroClienteSegunOf&(NORFA$)
       CI1% = XVALO(!cod_inte)
       'PRIMERINGRESO% = FechaPrimerTransfMaterialesSegunOf%(NORFA$, CI1%)
       FEMOV% = CVINT(!FechMov)
       COEFPESOPORMETRO· = CoeficientePesoPorMetroSegunLote#(SAFETEXT$(!idenlote))
       CIRANG% = CODIGORANGO&(CI1%)
       CANTMETROSBOBINA = XVALO(!Resultado) 'TOTAL DE MTS
       'esoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       
'       PesoPorMetro# = PESMETRO(CI1%)  'TRAIGO EL DE LA FORMULA POR QUE NO PUEDO CALCULARLO EN UNA CONSULTA AGRUPADA
PesoPorMetro# = COEFPESOPORMETRO·
       SaldoBobinaEnKilos# = CANTMETROSBOBINA * PesoPorMetro#
       MSF2.TextMatrix(REG&, 1) = NORFA$
       MSF2.TextMatrix(REG&, 2) = FECHATEX$(FEGEN%)
       MSF2.TextMatrix(REG&, 3) = CODEXT$(CIFAB%)
       MSF2.TextMatrix(REG&, 4) = NC
       MSF2.TextMatrix(REG&, 5) = rasocli$(NC)
       'MSF2.TextMatrix(REG&, 6) = FECHATEX$(PRIMERINGRESO%)
       MSF2.TextMatrix(REG&, 6) = FECHATEX$(FEMOV%)
       MSF2.TextMatrix(REG&, 7) = CODEXT$(CIRANG%)
       MSF2.TextMatrix(REG&, 8) = CODEXT$(CI1%)
       MSF2.TextMatrix(REG&, 9) = TRIM$(FORMATNUM$(CANTMETROSBOBINA, "F12.0"))
       MSF2.TextMatrix(REG&, 10) = TRIM$(FORMATNUM$(SaldoBobinaEnKilos#, "F12.0"))
       MSF2.TextMatrix(REG&, 11) = TRIM$(!idenlote)
       'MSF2.TextMatrix(REG&, 11) = OFNUM
      End If
80     .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
End Sub

 Sub CargarEgresos(NORFA1$, NORFA2$, Desde%, Hasta%, MODO%)
    SQLX$ = "SELECT o.idsubor,FeHorInf, NUMEOPER,DescOper,CANTREAL,FeHorInf,COORDUBIC,IDCOORDUBIC,O.Cod_Exte"
    SQLX$ = SQLX$ + ", CORTE,INFOFAB_ID,IDREGPADRE,OperFucionada,O.NumeOrFa,I.FeHorLib,O.StatuOrf,O.FechGen"
    SQLX$ = SQLX$ + ",O.Nume_Cli,I.CodiElem,i.peso"
    SQLX$ = SQLX$ + " FROM INFOFAB I with(NOLOCK)   INNER JOIN ORDEFABR O ON O.IDSUBOR = I.IDSUBOR"
    SQLX$ = SQLX$ + " Where (IDREGPADRE < 1 or IDREGPADRE is null)" ' AND IDCOORDUBIC < 100 "
    SQLX$ = SQLX$ + " and (OperFucionada < 1 or OperFucionada is null)"
  
  'FILTROS DE PANTALLA
  If XVALO(Mid$(NORFA1$, 4)) > 0 And XVALO(Mid$(NORFA2$, 4)) > 0 Then
     SQLX$ = SQLX$ + "  AND NUMEORFA >= " & XVALO(Mid$(NORFA1$, 4))
     SQLX$ = SQLX$ + "  AND NUMEORFA <= " & XVALO(Mid$(NORFA2$, 4))
  End If
  '
  If Desde% > 0 And Hasta% > 0 Then
     SQLX$ = SQLX$ + "  AND FeHorLib >= '" & FETEXCOR1$(Desde%) & "'"
     SQLX$ = SQLX$ + "  AND FeHorLib <= '" & FETEXCOR1$(Hasta%) & "'"
  End If
  If MODO% > 0 Then
    'MODOORDEN% 1 ABIERAS, 3 CERRADAS,  0 TODAS
    SQLX$ = SQLX$ + "  AND StatuOrf = " & MODO%
  End If
  SQLX$ = SQLX$ + " ORDER BY o.idsubor,I.FeHorInf"
  
  MSF3.Rows = 1: REG& = 0
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
     If XVALO(!CANTREAL) > 0 Then
       REG& = REG& + 1
       MSF3.Rows = REG& + 1
       NC = XVALO(!nume_cli)
       CIELEM% = XVALO(!codielem)
       PesoPorMetro# = PESMETRO(CIELEM%) 'TRAIGO EL DE LA FORMULA POR QUE NO PUEDO CALCULARLO EN UNA CONSULTA AGRUPADA

       MSF3.TextMatrix(REG&, 1) = SAFETEXT$(!idsubor)
       MSF3.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FECHGEN))
       MSF3.TextMatrix(REG&, 3) = SAFETEXT$(!Cod_Exte)
       MSF3.TextMatrix(REG&, 4) = NC
       MSF3.TextMatrix(REG&, 5) = rasocli$(NC)
       MSF3.TextMatrix(REG&, 6) = FECHATEX$(CVINT(!FeHorLib))
       MSF3.TextMatrix(REG&, 7) = XVALO(!NUMEOPER)
       MSF3.TextMatrix(REG&, 8) = SAFETEXT$(!descoper)
       MSF3.TextMatrix(REG&, 9) = XVALO(!CANTREAL)
       TOTMET# = XVALO(!CANTREAL) * XVALO(!corte)
       MSF3.TextMatrix(REG&, 10) = FORMATNUM$(TOTMET#, "F12.2")
       MSF3.TextMatrix(REG&, 11) = FORMATNUM$(XVALO(!Peso), "F12.2")
       CORDENADAUBICACION = SAFETEXT$(!COORDUBIC)
       If XVALO(CORDENADAUBICACION) > 0 Then
         CONDI$ = "CONVERT(int,campo1) = " & XVALO(CORDENADAUBICACION)
         CONDI$ = CONDI$ + " AND LISTA = 'CONFUBIX'"
         CORDENADAUBICACION = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
       End If
       MSF3.TextMatrix(REG&, 12) = CORDENADAUBICACION
       MSF3.TextMatrix(REG&, 13) = CODEXT$(CODIGORANGO&(CIELEM%))

       
      End If
      .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
End Sub


Sub CARGA_INFORMECONSUMO()
    'PRIMERO CARGO LA GRILLA
    Dim ACUMLACANTIDAD(32767)

    '
    'CARGO LA GRILLLA 3 EN BASE A LA QUE CONTIENE LA ESTRUCTORA SEGUN LA FICHA
    'Y GUARDO EN UN VECTOR EL ACUMULADO TRANSFERIDO DE CADA CODIGO
    CIFAB% = CODINT%(Text10)

    Call CARGAGRILLA3(Label4, XVALO(Text14))
'    MSF3.COL = 10
'    MSF3.Sort = 1

    Exit Sub
End Sub


Private Sub Label4_DblClick()
   Label4 = ""
End Sub

Private Sub mnuContConsumo_Click()
   ContConsumo.Show 1
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



Private Sub MSF2_Click()
   REG& = MSF2.MouseRow
   J& = MSF2.MouseCol
   
   If REG& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(InfSecCorte.MSF2, J&, MODODEORDEN$)
     Exit Sub
   End If
   
End Sub


Private Sub MSF3_Click()
   REG& = MSF3.MouseRow
   J& = MSF3.MouseCol
   
   If REG& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(InfSecCorte.MSF3, J&, MODODEORDEN$)
     Exit Sub
   End If
End Sub


Private Sub MSF5_Click()
   REG& = MSF5.MouseRow
   J& = MSF5.MouseCol
   
   If REG& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(InfSecCorte.MSF5, J&, MODODEORDEN$)
     Exit Sub
   End If
End Sub


