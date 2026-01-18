VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form LOTEXSQR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   13500
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10080
      ScaleHeight     =   75
      ScaleWidth      =   1260
      TabIndex        =   37
      Top             =   480
      Width           =   1320
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12960
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   430
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Busqueda x N° Serie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11160
      TabIndex        =   34
      Top             =   2880
      Width           =   2205
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
      Left            =   3120
      TabIndex        =   30
      ToolTipText     =   "Ordenes de Fabricación (Todas)"
      Top             =   120
      Width           =   175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "XXX"
      Top             =   2880
      Width           =   405
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "XXX"
      Top             =   2880
      Width           =   405
   End
   Begin VB.TextBox TXTCI 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12000
      TabIndex        =   23
      Top             =   120
      Width           =   1335
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
      Height          =   300
      Left            =   2760
      TabIndex        =   19
      ToolTipText     =   "Ordenes de Fabricación (Cerradas)"
      Top             =   120
      Width           =   175
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   80
      TabIndex        =   4
      Top             =   8040
      Width           =   13335
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "Agregar Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10800
         TabIndex        =   36
         ToolTipText     =   "Permite Agregar Manualmente un Accesorio "
         Top             =   160
         Width           =   2460
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Caption         =   "Grabar Carga de Lotes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   18
         Top             =   160
         Width           =   2205
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   17
         Top             =   160
         Width           =   3015
      End
      Begin VB.TextBox TXTREVISION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   160
         Width           =   1400
      End
      Begin VB.TextBox TXTLOTE 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   160
         Width           =   1400
      End
      Begin VB.TextBox TXTCODIGO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   160
         Width           =   1400
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Lectura:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   -50
         TabIndex        =   5
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   13800
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5280
      OleObjectBlob   =   "LOTEXSQR.frx":0000
      Top             =   4680
   End
   Begin MSFlexGridLib.MSFlexGrid MSFRESERVA 
      Height          =   2925
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Boton derecho Opciones: quitar Accesorio/Informar Lote a Componentes de la Formula"
      Top             =   5160
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSERIESVINC 
      Height          =   2085
      Left            =   6720
      TabIndex        =   13
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3678
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFSegunOf 
      Height          =   2325
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "Clic Para Seleccionar N.Serie/Boton derecho para borrar solo datos CPU"
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFESTADOVINCULADAS 
      Height          =   1485
      Left            =   120
      TabIndex        =   24
      ToolTipText     =   "Doble Click Para Seleccionar O.Fab. Como  Activa"
      Top             =   3240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   2619
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   9600
      TabIndex        =   33
      Top             =   165
      Width           =   405
   End
   Begin VB.Label LBLREVISION 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   10080
      TabIndex        =   32
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   6
      Left            =   3840
      TabIndex        =   31
      Top             =   4785
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "O.Fabricación Abiertas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   28
      Top             =   2925
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "O.Fabricación Cerradas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   27
      Top             =   2925
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vinculaciones del N.Serie Activo (Todas)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   4365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   4
      Left            =   2280
      TabIndex        =   22
      Top             =   4785
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Series de Orden de Fabricación Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   21
      Top             =   765
      Width           =   3765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "N° Serie  (Vinculación Inmediata)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   6720
      TabIndex        =   14
      Top             =   435
      Width           =   3285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reservas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   4785
      Width           =   2205
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   4065
   End
   Begin VB.Label Label44 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. O/F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu mnuArchivox2 
      Caption         =   "Archivo"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Begin VB.Menu mnuEtiqAut 
            Caption         =   "Etiqueta Autoadhesiva"
         End
         Begin VB.Menu mnuEtiqFirma 
            Caption         =   "Etiqueta con Firma"
         End
         Begin VB.Menu mnuRegistroArmado 
            Caption         =   "Registro de Armado"
         End
         Begin VB.Menu mnuPrintEmbalajeC15 
            Caption         =   "Etiqueta de Embalaje C15"
         End
         Begin VB.Menu mnuPrintEtiqEquipoC15 
            Caption         =   "Etiqueta de Equipo C15"
         End
      End
      Begin VB.Menu mnuGurdaDatosCPU 
         Caption         =   "Guardar Datos de CPU"
      End
      Begin VB.Menu mnuCierrreOFActiva 
         Caption         =   "Cierre de Orden de Fabricación Activa"
      End
      Begin VB.Menu mnuGLA 
         Caption         =   "Generación de Lotes Automáticos"
      End
      Begin VB.Menu mnuDesvioTransitorio 
         Caption         =   "Desvío Transitorio"
      End
      Begin VB.Menu modDesvioTransitorio 
         Caption         =   "Modificar Desvio Transitorio"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnularOF 
         Caption         =   "Anular Cierre de O.Fabricación"
      End
   End
   Begin VB.Menu mnuEngancharNSerie 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuEngancharNSerieAOF 
         Caption         =   "Enganchar N.Serie a Orden de Fabricación "
      End
      Begin VB.Menu mnuInicalizaDVT 
         Caption         =   "Inicializar Desvio Transitorio"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgregarItemaOF 
         Caption         =   "Agregar Item a Estrucutura de una Orden de Fabricación"
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuConslis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuDesvTrans 
         Caption         =   "Consulta de Desvios Transitorios"
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Visible         =   0   'False
      Begin VB.Menu mnuMenu3 
         Caption         =   "Revertir Selección"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuVincularOrdAnt 
         Caption         =   "Vincular Ordenes Anteriores"
      End
      Begin VB.Menu mnuControles 
         Caption         =   "Controles "
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImpoRegArmado 
         Caption         =   "Importación de Regisrtro de Armado desde Excel"
      End
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuHNS 
         Caption         =   "Historial de N.Serie"
      End
      Begin VB.Menu mnurecuuperalotes 
         Caption         =   "Proceso de Recuperacion de formato de  lotes"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuconfigimp 
         Caption         =   "Configurar Impresoras de Etiquetas"
      End
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuGennserie 
         Caption         =   "Generador de Números de Serie"
      End
   End
End
Attribute VB_Name = "LOTEXSQR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim MSFGLOBALFRM As msFlexGrid

Dim LecturaDatosOriginales As Byte
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub AGREGAITEM_A_RESERVA(NORFA$, CICOMP%, USOII)
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
     newItem = ProxNumOrdItemResrerva&(NORFA$)
     SQLX$ = " SELECT * FROM Reserva where IdSubOr = '" & NORFA$ & "' "
     rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     'AGREGA A LA RESERVA COMO CANTIDAD EL USO POR EL SALDO QUE TIENE LA OF
     'SI NO TIENE CIERRRES POR EL TOTAL SI TIENE PARCIALES POR LA DIFERENCIA
     KANAPRO = CantiAprobadaOF(NORFA$)
     CANTITOTAL = USOII * (CantiProyectadaOF(NORFA$) - KANAPRO)
        '
     If Abs(CANTITOTAL) >= 0.05 Then
          With rs
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !COD_INTE = CICOMP%
            !COD_EXTE = CODEXT$(CICOMP%)
            !Descrip = DESCRIT0$(CICOMP%)
            !IDSUBOR = NORFA$
            !FechRes = Now
            !CANTRES = CANTITOTAL
            !CANTCONS = KANAPRO
            !Saldo_Entre = CANTITOTAL
            !idenlote = " "
            !DEPOSITO = 0
            !USOUNIT = USOII
            !NuOrItem = newItem
            !NUME_SOLINT = 0
            .Update
          End With
     End If
     rs.Close
     Set rs = Nothing
End Sub

Sub AgregarEditarAccesorios(CODIGOEXT$, NORFA$, NSERIE$)
       'ATENCION VIENE DE LA RUTINA :  IMPRIMIR_ETIQUETAS_F054_EYM
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "CODIGO/A16;DESCRIPCIÓN/A32;LOTE-SN/A16;ID/A0;CI/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        
        VENTANA.CampEditables = "1;3"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Impresión F054"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 7500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        
        SQLX$ = "SELECT COD_INTE,COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION "
        SQLX$ = SQLX$ + " ,IDSUBOR,NUMEROSERIE,IDSUBOR,OF_ARRASTRA,TRAZABLE"
        SQLX$ = SQLX$ + " ,FECHALTA,CANTI,ORD,MARBORRA "
        
        SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE (IDSUBOR = '" & NORFA$ & "' OR OF_ARRASTRA = '" & NORFA$ & "') "
        SQLX$ = SQLX$ + " AND NUMEROSERIE = '" & NSERIE$ & "' AND COD_EXTE <> '" & CODIGOEXT$ & "'"
        
        SQLX$ = SQLX$ + " ORDER BY TRAZABLE,ORD"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            CI1% = XVALO(!COD_INTE)
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            '
            If ES_CODMODELO%(CI1%) > 0 Then
               VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!NumeroSerie)
            Else
               VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!idenlote)
            End If
            '
            VENTABNEW.MSF.TextMatrix(k&, 4) = XVALO(!TRAZAXSERIE_ID)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        '
        VENTABNEW.Show 1
        '
        If VENTABNEW.APROBADO% > 0 Then

            On Error GoTo ERROR_GUARDAR
            FLEXCONN.BeginTrans
            Dim rs1 As ADODB.Recordset
            Set rs1 = New ADODB.Recordset
            J& = 0
            SQLX$ = "SELECT COD_INTE,COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION "
            SQLX$ = SQLX$ + " ,IDSUBOR,NUMEROSERIE,IDSUBOR,OF_ARRASTRA,TRAZABLE"
            SQLX$ = SQLX$ + " ,FECHALTA,CANTI,ORD,MARBORRA "
            
            SQLX$ = SQLX$ + " FROM TRAZAXSERIE "
            SQLX$ = SQLX$ + " WHERE (IDSUBOR = '" & NORFA$ & "' OR OF_ARRASTRA = '" & NORFA$ & "') "
            rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

            For REG& = 1 To VENTABNEW.MSF.Rows - 1
              ID = XVALO(VENTABNEW.MSF.TextMatrix(REG&, 4))
              If ID < 1 Then 'SI SON DATOS NUEVOS LOS GUARDA NO ASI CON LA YA EXISTENTE VALIDO CON EL ID
                CODI$ = VENTABNEW.MSF.TextMatrix(REG&, 1)
                LOTE$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 3))
                DESXRIP$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 2))
                With rs1
50                  rs1.AddNew   ' YA CONVERTIDO
                    J& = J& + 1
                    On Error Resume Next
                    rs1!CodiEmpr = CodiEmp%
                    On Error GoTo 0
                    rs1!IDSUBOR = NORFA$
                    rs1!OF_ARRASTRA = ""
                    rs1!idenlote = LOTE$
                    rs1!NumeroSerie = NSERIE$
                    rs1!TRAZABLE = 1
                    rs1!COD_INTE = CODINT(CODI$)
                    rs1!Descripcion = DESXRIP$
                    rs1!COD_EXTE = CODI$
                    rs1!CANTI = 1
                    rs1!FECHALTA = CVDAT(HOY(HO$))
                    rs1!ORD = J&
                    rs1!MARBORRA = 0
                    rs1.Update
                 End With
              End If
            Next REG&
            rs1.Close: Set rs1 = Nothing
        Else
            Exit Sub
        End If
        
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
    End If

End Sub

Sub ComponerLotesenTrazaxserie()
  SQLX$ = "SELECT COD_INTE,IDENLOTE FROM TRAZAXSERIE  WHERE (LEFT(IDENLOTE,2) = 'OF' OR LEFT(IDENLOTE,2) = 'PR' OR LEFT(IDENLOTE,2) = 'BR' ) "
  SQLX$ = SQLX$ + " AND SUBSTRING(IDENLOTE,3,1)  <> '-'  order by idenlote"
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
      Do While Not .EOF
         LEIDO$ = SAFETEXT$(!idenlote)
         CI1% = XVALO(!COD_INTE)
         If Left(LEIDO$, 2) = "OF" Then
            If IsNumeric(Mid(LEIDO$, 3)) Then
               !idenlote = "OF-" & FORMATOCON0$(Mid(LEIDO$, 3), 6) & "-1"
            End If
         End If
         If Left(LEIDO$, 2) = "PR" Or Left(SAFETEXT$(!idenlote), 2) = "BR" Then
         
             LEIDORECE$ = Mid(LEIDO$, 3)
             'SI TIENE GUION SE LO QUITO
             If InStrRev(LEIDORECE$, "-") = 1 Then LEIDORECE$ = Left(LEIDORECE$, Len(LEIDORECE$) - 1)
             NUMERECE = XVALO(LEIDORECE$)
             SQLX$ = "SELECT * FROM bolrecep WITH(NOLOCK) WHERE NUMEBORE = " & NUMERECE
             SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
             Set RSS = READSET(SQLX$)
             With RSS
               If Not .EOF Then
                  LOTEBOLRECEP$ = SAFETEXT$(!idenlote)
                  rs!idenlote = LOTEBOLRECEP$
                  rs.Update
               End If
             End With
             RSS.Close: Set RSS = Nothing
         End If
      
        .MoveNext
      Loop
  End With
End Sub

Sub GuardarDatosCPUaCodigosDB()
      'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13)
'        If NORFA$ = "" Then
'           Call COMUNI("No Hay una Orden de Fabricación Activa")
'           Exit Sub
'        End If
        '
        
'*** OJO ESTA RUTINA ES IGUAL PERO CON MAS REGISTROS --- RUTINA ORIGINAL GuardarDatosCPU ----

        NSERIE$ = TRIM$(Label2(4))
'        If NSERIE$ = "" Then
'           Call COMUNI("Falta Seleccionar Un Número de Serie")
'           Exit Sub
'        End If
        '
        CI1% = XVALO(TXTCI.TEXT)
        CODIGOEXT$ = TRIM$(CODEXT$(CI1%))
        
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        'VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        
        VENTANA.CampEditables = "3;4"
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Datos CPU (Versión Software - Número de serie CPU) Para N° Serie: " & NSERIE$
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 4500
        VENTABNEW.Width = 11500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        '
        'VECTOR CARGO LOS DOS  PRIMEROS REGISTROS EN LAS 2 COLUMNAS CON EL DATO FIJO
        Dim CODI$(6), DESX$(6)
        CODI$(1) = "Responsable": DESX$(1) = "Preparado Por: "
        CODI$(2) = "Cpu Nº:": DESX$(2) = "Número de serie CPU"
        
        CODI$(3) = "Software CPU": DESX$(3) = "Versión Software CPU"
        CODI$(4) = "Software DESC": DESX$(4) = "Versión Software DESC"
        CODI$(5) = "Software ECG": DESX$(5) = "Versión Software ECG"
        CODI$(6) = "Software AUDIO": DESX$(6) = "Versión Software AUDIO"

        '
        VENTABNEW.MSF.Rows = 8
        VENTABNEW.MSF.TextMatrix(1, 1) = CODI$(1)
        VENTABNEW.MSF.TextMatrix(1, 2) = DESX$(1)
        VENTABNEW.MSF.TextMatrix(2, 1) = CODI$(2)
        VENTABNEW.MSF.TextMatrix(2, 2) = DESX$(2)
        VENTABNEW.MSF.TextMatrix(3, 1) = CODI$(3)
        VENTABNEW.MSF.TextMatrix(3, 2) = DESX$(3)
        
        VENTABNEW.MSF.TextMatrix(4, 1) = CODI$(4)
        VENTABNEW.MSF.TextMatrix(4, 2) = DESX$(4)
        VENTABNEW.MSF.TextMatrix(5, 1) = CODI$(5)
        VENTABNEW.MSF.TextMatrix(5, 2) = DESX$(5)
        VENTABNEW.MSF.TextMatrix(6, 1) = CODI$(6)
        VENTABNEW.MSF.TextMatrix(6, 2) = DESX$(6)
        
        'FIN DE CARGA POR DEFECTO
        
        'CARGO PRIMERO LOS DATOS DEL CPU
        SQLX$ = "SELECT CODIGO,DESCRIPCION,IDENTIFICACION,DESV_TRANS,DATOSCPU_ID FROM DATOSCPU WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CI1%
        SQLX$ = SQLX$ + " ORDER BY DATOSCPU_ID "
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            If k& <= 6 Then
              VENTABNEW.MSF.TextMatrix(k&, 1) = CODI$(k&)
              VENTABNEW.MSF.TextMatrix(k&, 2) = DESX$(k&)
            Else
              VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!Codigo)
              VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            End If
            VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!IDENTIFICACION)
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = SAFETEXT$(!DATOSCPU_ID)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% > 0 Then
            On Error GoTo ERROR_GUARDAR
            FLEXCONN.BeginTrans
            Dim rs1 As ADODB.Recordset
            Set rs1 = New ADODB.Recordset
            SQLX$ = "SELECT * FROM DATOSCPU WHERE NUMEROSERIE=''" '
            rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            J& = 0
            
            For k& = 1 To VENTABNEW.MSF.Rows - 1
              ID = XVALO(VENTABNEW.MSF.TextMatrix(k&, 5))
              If ID > 0 Then
                SQLX$ = "UPDATE DATOSCPU "
                SQLX$ = SQLX$ + " SET Codigo = '" & VENTABNEW.MSF.TextMatrix(k&, 1) & "'"
                SQLX$ = SQLX$ + " , DESCRIPCION = '" & VENTABNEW.MSF.TextMatrix(k&, 2) & "'"
                SQLX$ = SQLX$ + " , IDENTIFICACION = '" & VENTABNEW.MSF.TextMatrix(k&, 3) & "'"
                SQLX$ = SQLX$ + " , DESV_TRANS = '" & VENTABNEW.MSF.TextMatrix(k&, 4) & "'"
                SQLX$ = SQLX$ + "  WHERE DATOSCPU_ID = " & ID
              Else
                If TRIM$(VENTABNEW.MSF.TextMatrix(k&, 1)) <> "" Then
                    SQLX$ = "INSERT INTO DATOSCPU ("
                    SQLX$ = SQLX$ + "  Codigo,DESCRIPCION,IDENTIFICACION,DESV_TRANS,IDSUBOR,NUMEROSERIE,COD_INTE)"
                    SQLX$ = SQLX$ + " VALUES('" & VENTABNEW.MSF.TextMatrix(k&, 1) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 2) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 3) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 4) & "'"
                    SQLX$ = SQLX$ + " ,'" & Label13 & "'"
                    SQLX$ = SQLX$ + " ,'" & Label2(4) & "'"
                    SQLX$ = SQLX$ + " ," & CI1%
                    SQLX$ = SQLX$ + ")"
                End If
              End If
              FLEXCONN.Execute (SQLX$)
           Next k&
           FLEXCONN.CommitTrans
           Exit Sub
        End If
        
        
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
    End If

End Sub

Sub GuardarLoteAlComponenteEnTablaLOTEEYM(FECHA%, CODICONJ%, CIELELM%, LOTE$, NSERIE$, NORFA$, LoteDelComponente$, CANTASALI)
        '
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans

              CONDI$ = "NUMEROSERIE= '" & NSERIE$ & "'"
              CONDI$ = CONDI$ + " AND CODICONJ = " & CODICONJ%
              CONDI$ = CONDI$ + " AND CODIELEM = " & CIELELM%
              CONDI$ = CONDI$ + " AND IDSUBOR = '" & NORFA$ & "'"
              CONDI$ = CONDI$ + " AND CodiEmpr = " & CodiEmp%

              ID = XVALO(VALOBADA("LOTEEYM", "LOTEEYM_ID", CONDI$, "NOMESS"))
              If ID > 0 Then
                SQLX$ = "UPDATE LOTEEYM "
                SQLX$ = SQLX$ + " SET IDENLOTE = '" & LOTE$ & "'"
                SQLX$ = SQLX$ + " , FEREMOVI = '" & FETEXCOR1$(FECHA%) & "'"
              Else
                If CODICONJ% > 0 And CIELELM% > 0 And LOTE$ <> "" And NSERIE$ <> "" Then
                    SQLX$ = "INSERT INTO LOTEEYM ("
                    SQLX$ = SQLX$ + "  CODICONJ,CODIELEM,IDSUBOR,NUMEROSERIE,FEREMOVI,IDENLOTE,CodiEmpr)"
                    SQLX$ = SQLX$ + " VALUES(" & CODICONJ%
                    SQLX$ = SQLX$ + " ," & CIELELM%
                    SQLX$ = SQLX$ + " ,'" & NORFA$ & "'"
                    SQLX$ = SQLX$ + " ,'" & NSERIE$ & "'"
                    SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(FECHA%) & "'"
                    SQLX$ = SQLX$ + " ,'" & LOTE$ & "'"
                    SQLX$ = SQLX$ + " ," & CodiEmp%
                    SQLX$ = SQLX$ + ")"
                 End If
              End If
              FLEXCONN.Execute (SQLX$)
              'GRABA EN EL MOVISTO EL LOTE AL COMPONENTE CONSUMIDO EN LA OF DEL AC0008
              'P Q. EL AC0008 SE CERRO CON ANTERIORIDAD Y YA CONSUMIO SUS COMPONENTES
              'GUARDAMOS EL N.SERIE EN IDENMOVI PARA TENER EL CONTROL DE LOS LOTES YA APLICADOS
              'ESTA COMIBINACION DEBE SER UNICA
              DOSEC$ = "P" & Mid$(LoteDelComponente$, 2) 'PF-005453-1
              CONDI$ = "CODIMOVI = 'CF' AND DOPRILAR = '" & LoteDelComponente$ & "' AND COD_INTE=" & CIELELM% & "  AND DOSECLAR = '" & DOSEC$ & "'"
              CONDI$ = CONDI$ + " AND CANTSALID = " & CANTASALI
              CONDI$ = CONDI$ + " and CodiEmpr = " & CodiEmp%
              CONDI2$ = CONDI$ + " AND IDENMOVI = '" & NSERIE$ & "'" 'LO PONGO APARTE POR QUE LUEGO USO SOLO CONDI$
              CANTREG& = CantRegistros&("MOVISTO", CONDI2$, "NOMESS")
              If CANTREG& < 1 Then
                 'SI LA COMBINACION NO EXISTE REALIZO EL UPDATE ENTRARIA LA PRIMERA VEZ
                 SQLX$ = "UPDATE TOP(1) MOVISTO SET IDENMOVI = '" & NSERIE$ & "'"
                 SQLX$ = SQLX$ + " , IDENLOTE = '" & LOTE$ & "'"
                 SQLX$ = SQLX$ + " WHERE " & CONDI$
'                 FLEXCONN.Execute (SQLX$)
                 FLEXCONN.Execute SQLX$, affectedRecords
                 GoTo 98
              Else 'ACA ENTRARIA SI SE REPITE EL PASO CON EL MISMO LOTE O CON OTRO EN ESTE CASO LO ACTUAIZARIA
                 CONDI$ = "CODIMOVI = 'CF' AND DOPRILAR = '" & LoteDelComponente$ & "' AND COD_INTE=" & CIELELM% & "  AND DOSECLAR = '" & DOSEC$ & "'"
                 CONDI$ = CONDI$ + " AND CANTSALID = " & CANTASALI
                 CONDI$ = CONDI$ + " and CodiEmpr = " & CodiEmp%
                 CONDI2$ = CONDI$ + " AND IDENMOVI = '" & NSERIE$ & "' " 'LO PONGO APARTE POR QUE LUEGO USO SOLO CONDI$
                 SQLX$ = "UPDATE TOP(1) MOVISTO SET IDENLOTE = '" & LOTE$ & "'"
                 SQLX$ = SQLX$ + " WHERE " & CONDI2$
'                 FLEXCONN.Execute (SQLX$)
                 FLEXCONN.Execute SQLX$, affectedRecords
                 If affectedRecords > 0 Then GoTo 98 'SI NO ENCONTRO SIGUE
              End If
              'SI NO ENTRO EN LAS ANTERIORES ES POR QUE NO ENCONTRO Y APLICA ENTONCES N.SERIE Y LOTE Y CONTROLA LA CANTIDAD
              CONDI$ = "CODIMOVI = 'CF' AND DOPRILAR = '" & LoteDelComponente$ & "' AND COD_INTE=" & CIELELM%
              CONDI$ = CONDI$ + " AND DOSECLAR = '" & DOSEC$ & "' AND (IDENMOVI= '' OR IDENMOVI IS NULL)"
              CONDI$ = CONDI$ + " AND (IDENLOTE= '' OR IDENLOTE IS NULL)"
              CONDI$ = CONDI$ + " AND CodiEmpr = " & CodiEmp%
              CONDI$ = CONDI$ + " ORDER BY CANTSALID "
              SQLX$ = "SELECT IDENLOTE,IDENMOVI,CANTSALID,MOVISTO_ID FROM MOVISTO WITH(NOLOCK) WHERE " & CONDI$
              Set rs = READSET(SQLX$)
              SaldoAAplicar = CANTASALI
              With rs
              Do While Not .EOF
                CantiEnBD = XVALO(!CANTSALID)
                ID = XVALO(!MOVISTO_ID)
'                If SaldoAAplicar > 0 Then
                    If CantiEnBD <= SaldoAAplicar Then
                       SQLX$ = "UPDATE TOP(1) MOVISTO SET IDENMOVI = '" & NSERIE$ & "'"
                       SQLX$ = SQLX$ + " , IDENLOTE = '" & LOTE$ & "'"
                       SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & ID
                       FLEXCONN.Execute SQLX$, affectedRecords
                       SaldoAAplicar = SaldoAAplicar - CantiEnBD
                    Else
                       RetID& = 1
                       Call DuplicarRegistro("MOVISTO", "MOVISTO_ID = " & ID, RetID&)
                       If RetID& > 0 Then
                            'ACTUALIZO EL REGISTRO EXISTENTE
                            CantiEnRegistroExistente = CantiEnBD - SaldoAAplicar
                            SQLX$ = "UPDATE MOVISTO SET CANTSALID =  " & CantiEnRegistroExistente
                            SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & ID
                            FLEXCONN.Execute SQLX$, affectedRecords
                            'ACTUALIZO EL REGISTRO DUPLICADO
                            SQLX$ = "UPDATE MOVISTO SET IDENMOVI = '" & NSERIE$ & "'"
                            SQLX$ = SQLX$ + " , IDENLOTE = '" & LOTE$ & "'"
                            SQLX$ = SQLX$ + " , CANTSALID = " & SaldoAAplicar
                            SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & ID
                            FLEXCONN.Execute SQLX$, affectedRecords 'si entro en este else aplica todo
                            Exit Do
                       End If
                    End If
'                 End If
                 If SaldoAAplicar <= 0 Then Exit Do

                .MoveNext
              Loop
              End With
98
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       rs.Close: Set rs = Nothing
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Transacción Realizada")
    End If
             
              

99
End Sub

Sub LISTAOF_ABIERTAS_CERRADAS(Optional TODAS%)
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "IDSUBOR/A12;REFEREN/A48;C.Proy/F8.1;C.Apr./F8.1;Status/F7;Fecha/D8;"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
   Dim obj As New RECORXET
   Set rs1 = obj.RS_OF_ABYCERR(TODAS%)
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 11000
   FRMZELECHO.Show 1
'   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
'      Label13 = TRIM$(RESP_ZELE_VECT(1))
'   End If

End Sub

Sub ActualizaTanumseSiesProductoTerminado(NORFA$, NSERIE$)
    'REEMPLAZA LOS DATOS EN TANUMSE
    'CUANDO SE CREO SE HIZO CON OTRO CODIGO AHORA SI ESTE ES EL PRODUCTO QUE VA A LA VENTA
    'SE ACTUALIZA EL CODIGO, DESCRICION, Y LA OF Y EL DEPOSITO
    CI1% = CODIFAB%(NORFA$)
    SQLX$ = " SELECT CODIGO,DESCRIPCION FROM MASTER WITH(NOLOCK) WHERE CODINT = " & CI1%
    Set rs = READSET(SQLX$)
    If Not rs.EOF Then
       CODI$ = SAFETEXT$(rs!Codigo)
       DESX$ = SAFETEXT$(rs!Descripcion)
    End If
    rs.Close
    Set rs = Nothing
    '
    DEPOCIERRE% = XVALO(Control("", "DEPOCIE"))
    If DEPOCIERRE% < 1 Then DEPOCIERRE% = 1
    'si coincide con la configuracion de los productos terminado
    If TIMATE%(CI1%) = XVALO(Control("", "DEFPROTR")) Then
        SQLX$ = "UPDATE TANUMSE "
        SQLX$ = SQLX$ + " SET COD_INTE = " & CI1%
        SQLX$ = SQLX$ + " , COD_EXTE = '" & TRIM$(CODI$) & "'"
        SQLX$ = SQLX$ + " , DESCRIP = '" & TRIM$(DESX$) & "'"
        SQLX$ = SQLX$ + " , IDSUBOR = '" & TRIM$(NORFA$) & "'"
        SQLX$ = SQLX$ + " , DEPONUMSE = " & DEPOCIERRE%
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & TRIM$(NSERIE$) & "'"
        FLEXCONN.Execute (SQLX$)
    End If
End Sub

Function FALTAN_INFORMAR_LOTES%(CI1%, NORFA$, NROSERIE$)
   'VALIDA SI HAY LOTES SIN INFORMAR EN LA TABLA TRAZAXSERIE
   
   CONDI$ = "Idsubor = '" & NORFA$ & "' AND NUMEROSERIE = '" & NROSERIE$ & "'  AND TRAZABLE = 1 AND (idenlote is null or idenlote = '')"
   FALTAN_INFORMAR_LOTES% = XVALO(CantRegistros&("TRAZAXSERIE", CONDI$, "NOMESS"))
End Function

Function GrabarLotes$()
     'grabo los lotes que estan en pantalla por si modifico algo y no grabo
     GrabarLotes$ = "OK"
     If MSFRESERVA.Rows <= 1 Then Exit Function
     If ValidacionDeLotes%(Me.MSFRESERVA) > 0 Then
        OPX% = COMALTER%("Puede Continuar Dando Por Informados Los Lotes Existentes o Nó\Sabiendo Que este Dato Quedará Registrado Para la Trazabilidad\Pero el Consumo se Realizará a Modo Fifo\Sobre Lotes Existentes.\\Cancelar\Aceptar Condición")
        If OPX% < 2 Then
          GrabarLotes$ = ""
          GoTo 99
        End If
     End If
     '
     i& = 0
     CONDI$ = " IdSubOr = '" & NORFA$ & "' AND TRAZABLE = 1 AND NUMEROSERIE = '" & NSERIE$ & "'"
     On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     Dim rs As ADODB.Recordset
     
     For i& = 1 To Me.MSFRESERVA.Rows - 1
        ID = XVALO(Me.MSFRESERVA.TextMatrix(i&, 6))
        SQLX$ = "SELECT IDENLOTE, DESV_TRANS FROM TRAZAXSERIE WHERE TRAZAXSERIE_ID = " & ID
        Set rs = New ADODB.Recordset
        rs.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Not rs.EOF Then
           rs!idenlote = FORMATOBR$(Me.MSFRESERVA.TextMatrix(i&, 5))
           DVT = DESVIOTRANS(XVALO(MSFRESERVA.TextMatrix(i&, 7)), TRIM$(MSFRESERVA.TextMatrix(i&, 5)), TRIM$(MSFRESERVA.TextMatrix(i&, 4)), REFERENCIA$)
'           If DVT = 0 Then
'              DVT = DESVIOTRANS(XVALO(MSFRESERVA.TextMatrix(i&, 7)), TRIM$(MSFRESERVA.TextMatrix(i&, 5)), "", REFERENCIA$)
'           End If
           rs!DESV_TRANS = DVT
           MSFRESERVA.TextMatrix(i&, 8) = DVT
           rs.Update
        End If
     Next i&
     GrabarLotes$ = "OK"
     
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       rs.Close: Set rs = Nothing
       GrabarLotes$ = ""
       On Error GoTo 0
       'SI HUBO ERROR REFRESCA LA PANTALLA
       NORFA$ = Label13.Caption
       Label13.Caption = ""
       Label13.Caption = NORFA$
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      rs.Close: Set rs = Nothing
    End If


99
    Call PINTATEXT(Text1)


End Function

Sub IMPRIMIR_ETIQUETAS_F054_EYM()
        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Label2(4))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        Call AgregarEditarAccesorios(Label44, NORFA$, NSERIE$) ' ANTES ESTABA TODO ESTE CODIGO (EL COMENTADO) AQUI SE MUDO A ESTA SUB
'         Unload VENTABNEW
'        VENTANA.Inicializar = 0
'        Campos$ = "CODIGO/A16;DESCRIPCIÓN/A32;LOTE-SN/A16;ID/A0;CI/A0"
'        VENTANA.Campox = Campos$
'        VENTANA.UtilizaFormula = 1
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
'
'        VENTANA.CampEditables = "1;3"
'        VENTANA.AgregaRegistro = 1
'        VENTANA.UtilizaFormula = 1
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.HabilitaInsertar = 1
'        VENTANA.NoMostrarBotonTilde = 0
'        VENTANA.TITULO = "Impresión F054"
'        VENTANA.Inicializar = 1
'        VENTABNEW.Height = 7500
'        VENTABNEW.MSF.Font = "Roboto Mono"
'        VENTABNEW.MSF.Font.Size = 8
'
'        SQLX$ = "SELECT COD_INTE,COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION "
'        SQLX$ = SQLX$ + " ,IDSUBOR,NUMEROSERIE,IDSUBOR,OF_ARRASTRA,TRAZABLE"
'        SQLX$ = SQLX$ + " ,FECHALTA,CANTI,ORD,MARBORRA "
'
'        SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) "
'        SQLX$ = SQLX$ + " WHERE (IDSUBOR = '" & NORFA$ & "' OR OF_ARRASTRA = '" & NORFA$ & "') "
'        SQLX$ = SQLX$ + " AND NUMEROSERIE = '" & NSERIE$ & "' AND COD_EXTE <> '" & Label44 & "'"
'
'        SQLX$ = SQLX$ + " ORDER BY TRAZABLE,ORD"
'        Set rs = READSET(SQLX$)
'        With rs
'          Do While Not .EOF
'            k& = k& + 1
'            VENTABNEW.MSF.Rows = k& + 1
'            CI1% = XVALO(!COD_INTE)
'            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
'            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
'            '
'            If ES_CODMODELO%(CI1%) > 0 Then
'               VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!NUMEROSERIE)
'            Else
'               VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!idenlote)
'            End If
'            '
'            VENTABNEW.MSF.TextMatrix(k&, 4) = XVALO(!TRAZAXSERIE_ID)
'            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!COD_INTE)
'            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
'            .MoveNext
'          Loop
'        End With
'        rs.Close
'        Set rs = Nothing
'        '
'        VENTABNEW.Show 1
'        '
'        If VENTABNEW.APROBADO% > 0 Then
'
'            On Error GoTo ERROR_GUARDAR
'            FLEXCONN.BeginTrans
'            Dim RS1 As ADODB.Recordset
'            Set RS1 = New ADODB.Recordset
'            J& = 0
'            SQLX$ = "SELECT COD_INTE,COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION "
'            SQLX$ = SQLX$ + " ,IDSUBOR,NUMEROSERIE,IDSUBOR,OF_ARRASTRA,TRAZABLE"
'            SQLX$ = SQLX$ + " ,FECHALTA,CANTI,ORD,MARBORRA "
'
'            SQLX$ = SQLX$ + " FROM TRAZAXSERIE "
'            SQLX$ = SQLX$ + " WHERE (IDSUBOR = '" & NORFA$ & "' OR OF_ARRASTRA = '" & NORFA$ & "') "
'            RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'
'            For REG& = 1 To VENTABNEW.MSF.Rows - 1
'              ID = XVALO(VENTABNEW.MSF.TextMatrix(REG&, 4))
'              If ID < 1 Then 'SI SON DATOS NUEVOS LOS GUARDA NO ASI CON LA YA EXISTENTE VALIDO CON EL ID
'                CODI$ = VENTABNEW.MSF.TextMatrix(REG&, 1)
'                LOTE$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 3))
'                DESXRIP$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 2))
'                With RS1
'50                  RS1.AddNew   ' YA CONVERTIDO
'                    J& = J& + 1
'                    On Error Resume Next
'                    RS1!CodiEmpr = CodiEmp%
'                    On Error GoTo 0
'                    RS1!IDSUBOR = NORFA$
'                    RS1!OF_ARRASTRA = ""
'                    RS1!idenlote = LOTE$
'                    RS1!NUMEROSERIE = NSERIE$
'                    RS1!TRAZABLE = 1
'                    RS1!COD_INTE = CODINT(CODI$)
'                    RS1!Descripcion = DESXRIP$
'                    RS1!COD_EXTE = CODI$
'                    RS1!CANTI = 1
'                    RS1!FechAlta = CVDAT(HOY(HO$))
'                    RS1!ORD = J&
'                    RS1!MARBORRA = 0
'                    RS1.Update
'                 End With
'              End If
'            Next REG&
'            RS1.Close: Set RS1 = Nothing
'        Else
'            Exit Sub
'        End If
'
'ERROR_GUARDAR:
'    If Err.Number <> 0 Then
'       FLEXCONN.RollbackTrans
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
'       On Error GoTo 0
'       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
'    End If
    '
    VENTABNEW.MSF.Rows = 16
    FECHALTA$ = Label2(6)
    VENTABNEW.MSF.TextMatrix(13, 1) = Label44
    VENTABNEW.MSF.TextMatrix(14, 1) = Label13
    VENTABNEW.MSF.TextMatrix(15, 1) = FECHALTA$
    
End Sub

Sub INGRESO_n_d_t(NUME&)
      If IsNumeric(NUME&) Then
        Call CreaRegistro("DES_TRANS", "DESV_TRANS", NUME&, "NOMESS")
      End If
End Sub

Sub LIMPIAR()
   Me.MSFRESERVA.Rows = 1
   Me.MSFSegunOf.Rows = 1
   Me.MSFSERIESVINC.Rows = 1
   Me.MSFESTADOVINCULADAS.Rows = 1
   Label2(4) = ""
   Label2(6) = ""
   Campos$ = "Alta/D8;O.F./A12; N.Serie/A12;ID/A0;Código/A16"
   Call CARGA_ENCABEZADO(Me.MSFSERIESVINC, Campos$, Me)
    
   Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;N.Serie/A16;Lote/A16;ID/A0;CI/A0"
   Call CARGA_ENCABEZADO(Me.MSFRESERVA, Campos$, Me)

   Campos$ = "Alta/D8;O.F./A20; N.Serie/A20;ID/A0"
   Call CARGA_ENCABEZADO(Me.MSFSegunOf, Campos$, Me)
   Text3 = ""
End Sub



Function LOTE_A_CONSUMIR_EYM$(CI1%, LOTE$, CANTI, NSERIE$)
    LOTEX1$ = REPLACAR(LOTE$, "PR-", "BR-")
    'DEVUELVE EL LOTE CORRESPONDIENTE PARA QUE CONSUMA
    'SI NO TIENE STOCK O NO LO ENCUENTRA DEVUELVE EL TEXTO LOTE_A_CONSUMIR_EYM$
    LOTE_A_CONSUMIR_EYM$ = LOTEX1$ ' POR DEFECTO CONSUME EL LOTE QUE VIENE DE PARAMETRO
    
    'VEO SI ES LOTE DEL PROVEEDOR (QUE PUEDEN CARGAR CUANDO REALIZAN LA RECEPCION
    CONDI$ = "LoteCol_Prov = '" & LOTEX1$ & "'"
    If CantRegistros&("BOLRECEP", CONDI$, "NOMESS") > 0 Then
       LOTEX$ = VALOBADA("BOLRECEP", "IDENLOTE", CONDI$, "NOMESS")
       If LOTEX$ <> "" Then LOTEX1$ = LOTEX$
    End If
    '
    'SI ES TRAZABLE -1 o ES TRAZABLE 2 DEVUELVE COMO LOTE LA OF SEGUN EL NUMERO DE SERIE.
    If ESTRAZABLE%(CI1%) = 2 Then 'TRAZABLE 2 o -1
       LOTE_A_CONSUMIR_EYM$ = OrdenFabricacionSegunNumSerie$(CI1%, NSERIE$)
       Exit Function
    End If
    '
    CIX& = CI1% ' EL FORMATO QUE LE PONE DE MODO AUTOMATICO POSEE EL N. DE CODINT
    LOTEAUTO$ = "LA-" & FORMATOCON0$(CIX&, 6) & "-1"
    '
    'VALIDO SI EL PASADO POR PARAMETRO NO POSEE LOTE CONSUME DEL LOTE AUTOMATICO
    If SaldoLote(CI1%, LOTEX1$) >= CANTI Then
       LOTE_A_CONSUMIR_EYM$ = LOTEX1$
    ElseIf SaldoLote(CI1%, LOTEAUTO$) >= CANTI Then
       LOTE_A_CONSUMIR_EYM$ = LOTEAUTO$
    Else
       LOTE_A_CONSUMIR_EYM$ = "SIN STOCK"
    End If
    
End Function

Private Sub asignans_Click()

'    OBX$ = OBJPLA$("CARNUMSE", VDEF$) ' OBJETO TANUMSE PARA INGRESAR LOS NUMEROS DE SERIE
'    VDEF$ = OBX$
'    If OBX$ = "" Then GoTo 5
'    NUSERIEE$ = TRIM$(Mid$(VDEF$, 1, 32))
'    CI% = CVI(Mid$(VDEF$, 33, 2))
'    fech% = CVI(Mid$(VDEF$, 35, 2))
'    FECC = FECHATEX(fech%)
'    TENSION$ = Mid$(VDEF$, 37, 6)
'    NuDepo% = CVI(Mid$(VDEF$, 43, 2))
'
'    CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIEE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
'    '
'    If CANTIRE& > 0 Then Call MENSERR(24, "Número de Serie Existente"): GoTo 10 ' VALIDACIONES
'    If CI% < 1 Then Call MENSERR(24, "Código no Válido"): GoTo 10
'    '
'    If InStr(EMPREAC$, "MEDITEA") > 0 Then
'      If Len(NUSERIEE$) > 4 Then
'        Ulti4Cod$ = Right(CODEXT$(CI%), 4) ' Valida que los ultimos cuatro numeros del codigo coincidan con los 4 primeros del numero de serie
'        Prim4NuSe1$ = Left$(NUSERIEE$, 4)
'       '
'       If Ulti4Cod$ <> Prim4NuSe1$ Then
'         Call COMUNI("Número de Serie Difiere de Código")
'         GoTo 10
'       End If
'      Else
'        Call MENSERR(24, "Número de Serie Cantidad de Caracteres menor a 5"): GoTo 10
'      End If
'    End If
'    '
'    DEPOSITO$ = ECOARCH$("TADEPOSI", Chr$(NuDepo%))
'    If DEPOSITO$ = "" Then Call MENSERR(24, "Depósito no Válido"): GoTo 10
'    '
'    MESYANIO$ = Mid$(NUSERIEE$, 5, 4)
'    NORDEN% = XVALO(Mid$(NUSERIEE$, 9, 4))
'    '
'    FECHAA$ = FECHATEX$(fech%)
'    CMOV$ = TRIM$(CONTROL$("", "CODAJUST")) ' SE OBTIENE PROXIMO NUMERO DE AJUSTE
'    If CMOV$ = "" Then CMOV$ = "AJ"
'    NVM$ = TRIM$(ProNroComprobante("AJ"))
'    While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
'    NRO& = XVALO(NVM$)
'    If CMOV$ = "" Then CMOV$ = "AJ"
'    DOSEC$ = "AJ." + Mid$(FECHAA$, 7, 2) + Mid$(FECHAA$, 4, 2) + Left$(FECHAA$, 2)
'    DOPRI$ = CMOV$ + "." + NVM$
'    REFE$ = "AJUSTE POR CARGA DIRECTA"
'    DEPX% = NuDepo%
'
'    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE=0" ' SI PASO LAS VALIDACIONES AGREGA EL NUEVO NUMERO DE SERIE
'    Dim TANUTMP As ADODB.Recordset
'    Set TANUTMP = New ADODB.Recordset
'    TANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'      With TANUTMP
'        .AddNew   ' YA CONVERTIDO
'        On Error Resume Next
'        !CodiEmpr = CodiEmp%
'        On Error GoTo 0
'        !NumeroSerie = NUSERIEE$
'        !MESANIO = MESYANIO$
'        !NUMEORD = NORDEN%
'        !COD_INTE = CI%
'        !COD_EXTE = CODEXT(CI%)
'        Descripcion$ = DESCRIT(CI%)
'        !Descrip = Descripcion$
'        !FechAlta = FECC
'        !TEN_ENSA = XVALO(TENSION$)
'        !DepoNumSe = NuDepo%
'        .Update
'        '
'        If RTA1% = 1 Then Call MOVISTOK(fech%, CI%, NUSERIEE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORDEN%, 0, 1, NC%, DEPX%, 1)
'        '
'      End With
'    TANUTMP.Close
'    Set TANUTMP = Nothing
'    '
'    'Call ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
'
'    End If
End Sub

'Sub RECUPERARRESERVAS_NSER(NORFA$, NSERIE$)
'   'RECORRO EL  TRAZAXSERIE DONDE TAMBIEN ESTAN LAS RESERVAS (POR QUE EN RESERVAS SE PUEDEN HABER BORRADO)
'   'SELECT * FROM TRAZAXSERIE WHERE IDSUBOR = 'OF-002210-1' AND NUMEROSERIE = '185-1474' AND TRAZABLE >= 0 AND TRAZABLE <= 1
'    Dim ACUCANT(32767): CODIXT$ (32767): DESCRX (32767)
'    SQLX$ = "SELECT * FROM TRAZAXSERIE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' "
'    SQLX$ = SQLX$ + " AND NUMEROSERIE = '" & NSERIE$ & "' AND TRAZABLE >= 0 AND TRAZABLE <= 1"
'    SQLX$ = SQLX$ + " ORDER BY COD_INTE"
'    Set rs = READSET(SQLX$)
'    With rs
'      Do While Not .EOF
'         CI1% = XVALO(!COD_INTE)
'         CANTI = XVALO(!CANTI)
'         ACUCANT(CI1%) = ROUND(ACUCANTCONSU(CI1%) + CANTI, 0)  'CANTIDAD ACUMULADA A DESCONTAR DE UN MISMO ITEM.
'         CODIXT$(CI1%) = SAFETEXT$(!Cod_Exte)
'         DESCRX(CI1%) = SAFETEXT$(!Descripcion)
'         .MoveNext
'      Loop
'    End With
'    rs.Close
'    Set rs = Nothing
'    '
'    FE% = HOY$(HO$)
'    J& = 0
'    For i% = 1 To 32767
'      If ACUCANT(i%) > 0 Then
'       SQLX$ = " SELECT * FROM Reserva WHERE"
'       SQLX$ = SQLX$ + "  IDSUBOR = '" & NORFA$ & "' AND COD_INTE =  " & i%
'       Dim RESERVVVA As ADODB.Recordset
'       Set RESERVVVA = New ADODB.Recordset
'       RESERVVVA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'       With RESERVVVA
'          If Not .EOF Then
'            !CANTCONS = XVALO(!CANTCONS) - ACUCANT(CI1%)
'            .Update
'          Else
'            .AddNew
'            !FEREMOVI = CVDAT(FE%)
'            !CODIEMPR = CODIEMPR%
'            !COD_INTE = i%
'            !Cod_Exte = CODIXT$(i%)
'            !Descrip = DESCRX(i%)
'            !IdSubOr = NORFA$
'            J& = J& + 1
'            !NuOrItem = J&
'            !FechRes = CVDAT(FE%)
'            !CantRes = ACUCANT(CI1%)
'            !CANTCONS = ACUCANT(CI1%)
'            !Saldo_Entre = ACUCANT(CI1%)
'            !Deposito = 0
'            !USOUNIT = ACUCANT(CI1%)
'            .Update
'          End If
'       End With
'       RESERVVVA.Close
'       Set RESERVVVA = Nothing
'    Next i%
'
'
'End Sub

Function ReemplazaXGuionCaracterDiferenteEYM(TEXTO$)
    'ESTA FUNCION SE CREO PARA EYM POR QUE LA LECTORA REEMPLAZABA AL GUION POR UN APOSTROFE
    'EL REPLACAR NO FUNCIONABA POR QUE NO ERA LA COMILLA SIMPLE COMUN
    RESTOTEXTO$ = TEXTO$
    ReemplazaXGuionCaracterDiferenteEYM = ""
    For i& = 1 To Len(TEXTO$)
       LECTURA$ = Left(RESTOTEXTO$, 1)
       AXC% = Asc(LECTURA$)
       'SOLO DEJA LAS LETRAS (MAYUSCULAS O MINUSCULAS), GUION, ESPACIOS,PUNTO Y NUMEROS
       
       If (AXC% >= 48 And AXC% <= 57) Or (AXC% >= 65 And AXC% <= 90) Or (AXC% >= 97 And AXC% <= 122) Or AXC% = 45 Or AXC% = 32 Or AXC% = 46 Then
         ReemplazaXGuionCaracterDiferenteEYM = ReemplazaXGuionCaracterDiferenteEYM & LECTURA$
       Else
         ReemplazaXGuionCaracterDiferenteEYM = ReemplazaXGuionCaracterDiferenteEYM & "-"
       End If
       RESTOTEXTO$ = Mid$(RESTOTEXTO$, 2)
    Next i&
End Function


Function SERIES_TODAS$()
        SERIES_TODAS$ = ""
        CONDI$ = " COD_INTE > 0"
        CONDI$ = CONDI$ + "  ORDER BY NUMEROSERIE"
        Campos$ = "N.SERIE/A20;FECHA ALTA"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        Dim obj As New RECORXET
        Set rs1 = obj.RS_NSERIE(CONDI$)
        Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2)
        Set obj = Nothing
        rs1.Close
        RESP_ZELE_VECT(1) = ""
        FRMZELECHO.Caption = "N° de Series Exisstentes"


        FRMZELECHO.Width = 10000
        FRMZELECHO.Show 1
        If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
          SERIES_TODAS$ = TRIM$(RESP_ZELE_VECT(1))
        End If
End Function

Private Sub Command2_Click()
   Command2.Enabled = False
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "IDSUBOR/A12;REFEREN/A48;C.Proy/F8.1"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set rs1 = obj.RS_OF_ABIERTAS(1)
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 9500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
      Label13 = TRIM$(RESP_ZELE_VECT(1))
   End If
    
   
99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "O.Fabric./A12;Referencia/A48;C.Proy./F8.1;C.Aprob./F8.1"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set rs1 = obj.RS_OF_ABIERTAS_2()
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 9500
   FRMZELECHO.Height = 7500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
      Label13 = TRIM$(RESP_ZELE_VECT(1))
   End If
   
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
     Call LIMPIAR
     Label44 = ""
     Label3 = ""
     Label13 = ""
     LBLREVISION = ""
     Me.TXTCI = ""
     Label2(4) = "": Label2(6) = ""
     NSERIE$ = SERIES_TODAS$
     If NSERIE$ = "" Then GoTo 99
     
     SQLX$ = "SELECT FechAlta, IDSUBOR,COD_EXTE,DESCRIPCION,STATUORF,COD_INTE,TRAZAXSERIE_ID,NUMEROSERIE"
     SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where(TRAZABLE = 2 Or TRAZABLE = -1) And NUMEROSERIE = '" & NSERIE$ & "' "
     SQLX$ = SQLX$ + " ORDER BY TRAZABLE DESC, STATUORF ,TRAZAXSERIE_ID"
      Set rs = READSET(SQLX$)
      With rs
         Do While Not .EOF
            i& = i& + 1
            Me.MSFESTADOVINCULADAS.Rows = i& + 1
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 1) = SAFETEXT$(!FECHALTA)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 3) = SAFETEXT$(!COD_EXTE)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 4) = SAFETEXT$(!Descripcion)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 5) = SAFETEXT$(!NumeroSerie)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 6) = SAFETEXT$(!TRAZAXSERIE_ID)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 7) = XVALO(!StatuOrf)
            If XVALO(!StatuOrf) <> 3 Then
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFESTADOVINCULADAS, &HFF0000, i&)
            Else
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFESTADOVINCULADAS, vbRed, i&)
            End If
    
            .MoveNext
         Loop
       End With
       rs.Close: Set rs = Nothing

99 Command4.Enabled = True
'
End Sub


Private Sub Command5_Click()
        Command5.Enabled = False
        NORFA$ = TRIM$(Label13)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Label2(4))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        
        Codigo$ = TRIM$(Label44)
        If CODINT%(Codigo$) < 1 Then
           Call COMUNI("Código No Válido")
           Exit Sub
        End If
        
        Call AgregarEditarAccesorios(Codigo$, NORFA$, NSERIE$)
        
        
        Command5.Enabled = True
   
End Sub

Private Sub Command8_Click()
     Command8.Enabled = False
     NORFA$ = Label13
     
     If NORFA$ = "" Then GoTo 99
     NSERIE$ = TRIM$(Me.MSFRESERVA.TextMatrix(1, 4)) 'LEO EL N.SERIE EN EL PRIMER REGISTRO(TODOS VAN A SER IGUALES
     If NSERIE$ = "" Then
        Call COMUNI("No hay N°.Series Seleccionados")
        GoTo 99
     End If
     '
     X = GrabarLotes$
99   Command8.Enabled = True
End Sub

Function ValidacionDeLotes%(MSF As msFlexGrid)
     ValidacionDeLotes% = 0
     PRESENTAR_ZELECHO% = 0: J& = 0
     Campos$ = "CI/A0;Código/A16;LOTE/A18"
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
     FRMZELECHO.Caption = "Lotes No Reconocidos en el Sistema"
     
     'VALIDA QUE LOS LOTES AL MENOS FIGUREN EL EL LOTINGRE
     For i& = 1 To Me.MSFRESERVA.Rows - 1
        CODI$ = Me.MSFRESERVA.TextMatrix(i&, 1)
        LOTE$ = TRIM$(Me.MSFRESERVA.TextMatrix(i&, 5))
        CI1% = XVALO(Me.MSFRESERVA.TextMatrix(i&, 7))
        CONDI$ = "IdenLote = '" & LOTE$ & "'"
        CONDI$ = CONDI$ + " AND Cod_Inte = " & CI1%
        CONDI$ = CONDI$ + " AND CANTINGRE > 0 "
        If CantRegistros("MOVISTO", CONDI$, "NOMESS") < 1 Then
             PRESENTAR_ZELECHO% = 1
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = CI1%
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = CODI$
             FRMZELECHO.MSF2.TextMatrix(J&, 3) = LOTE$
        End If
     Next i&
     '
     If PRESENTAR_ZELECHO% = 1 Then
        FRMZELECHO.Show 1
        ValidacionDeLotes% = 1
     End If
     '
End Function
Private Sub Form_Activate()
MENU1% = 0
MENU2% = 0
MENU3% = 0
SeleccionaFilaDobleClick% = 0
End Sub


Private Sub Form_Load()
     Call CREATABLA_STOCKS

     Campos$ = "Alta/D8;O.F./A12; N.Serie/A12;ID/A0;Código/A16"
     Call CARGA_ENCABEZADO(Me.MSFSERIESVINC, Campos$, Me)
     
     Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;N.Serie/A16;Lote/A16;ID/A0;CI/A0"
     Call CARGA_ENCABEZADO(Me.MSFRESERVA, Campos$, Me)

     Campos$ = "Alta/D8;O.F./A20; N.Serie/A20;ID/A0"
     Call CARGA_ENCABEZADO(Me.MSFSegunOf, Campos$, Me)
     
     Campos$ = "Alta/D8;O.F./A16; Código/A16;Descripción/A32;N.Serie/A20;ID/A0;Status/F7"
     Call CARGA_ENCABEZADO(Me.MSFESTADOVINCULADAS, Campos$, Me)
     Call GENERAR_ESTRUCTURAS_SQL
End Sub
Private Sub Form_Resize()
'   On Error Resume Next
'   Me.MSF_2.Width = Me.Width - 300
'   Me.GRID.Width = Me.Width - 300
'   Me.MSF_2.Height = Me.Height - 1600
'   Me.GRID.Height = Me.Height - 1600
'   Frame1.Top = Me.Height - 1300
'   Frame2.Top = Me.Height - 1300
'   On Error GoTo 0
End Sub


Private Sub GRID_Click()
'    Static MODO  As Boolean
'    i& = GRID.MouseRow
'    J& = GRID.MouseCol
'
'    If (GRID.MouseRow = 0) Then
'
'        ' Ordena en forma ascendente
'        If MODO Then
'            GRID.COL = GRID.MouseCol
'            GRID.Sort = 2
'            MODO = False
'        ' Ordena en forma descendente
'        Else
'            GRID.COL = GRID.MouseCol
'            GRID.Sort = 1
'            MODO = True
'        End If
'    ElseIf NOPINTAR% > 0 Then Exit Sub
'    End If
'    If J& = 0 Then
'       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
'    End If
End Sub

Private Sub GRID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If GRID.Rows <= 1 Then Exit Sub
  With GRID
    ' verifica que el GRIDlexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
        ' le asigna el dato con TextMatrix
        .ToolTipText = VectorToolTipText(.MouseCol)
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub MC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MC.Visible = False
End Sub

Private Sub Label13_Change()
   Call LIMPIAR
   NORFA$ = Label13
   CI1% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   TXTCI = CI1%
   If CI1% > 0 Then
     Label44 = CODEXT(CI1%)
     Label3 = DESCRIT$(CI1%)
     '
     ESEXCEPCION_SERIE% = ES_EXCEPCION_GENERAR_SERIE%(Label44)
     If ESEXCEPCION_SERIE% > 0 Then
         Campos$ = "Alta/D8;O.F./A12; Identificador/A14;ID/A0;Código/A16"
         Call CARGA_ENCABEZADO(Me.MSFSERIESVINC, Campos$, Me)

         Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;Identificador/A14;Lote/A16;ID/A0;CI/A0"
         Call CARGA_ENCABEZADO(Me.MSFRESERVA, Campos$, Me)

         Campos$ = "Alta/D8;O.F./A20; Identificador/A14;ID/A0"
         Call CARGA_ENCABEZADO(Me.MSFSegunOf, Campos$, Me)
     End If
     'CARGA DE VINCULADOS
'    '============================
     i& = 0
     CI1% = CODINT%(Label44)
     SQLX$ = "SELECT FECHALTA,IDSUBOR,NUMEROSERIE,TRAZAXSERIE_ID,STATUORF,COD_EXTE FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where OF_ARRASTRA = '" & NORFA$ & "'"
     Set rs = READSET(SQLX$)
     With rs
        Do While Not .EOF
           i& = i& + 1
           Me.MSFSERIESVINC.Rows = i& + 1
           Me.MSFSERIESVINC.TextMatrix(i&, 1) = SAFETEXT$(!FECHALTA)
           Me.MSFSERIESVINC.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
           Me.MSFSERIESVINC.TextMatrix(i&, 3) = SAFETEXT$(!NumeroSerie)
           Me.MSFSERIESVINC.TextMatrix(i&, 4) = XVALO(!TRAZAXSERIE_ID)
           Me.MSFSERIESVINC.TextMatrix(i&, 5) = SAFETEXT$(!COD_EXTE)
'            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 5) = XVALO(!StatuOrf)
            If XVALO(!StatuOrf) <> 3 Then
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFSERIESVINC, &HFF0000, i&)
            Else
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFSERIESVINC, vbRed, i&)
            End If
           
           .MoveNext
        Loop
     End With
     rs.Close: Set rs = Nothing
     
     'CARGA NUMEROS DE SERIE GENERADOS CON ESTA ORDEN DE FABRICACION
'    '============================
     i& = 0
     CI1% = CODINT%(Label44)
     SQLX$ = "SELECT FECHALTA,IDSUBOR,NUMEROSERIE,TRAZAXSERIE_ID,StatuOrf FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where IDSUBOR = '" & NORFA$ & "' AND (TRAZABLE = 2  OR TRAZABLE = -1)"
     SQLX$ = SQLX$ + " ORDER BY IDSUBOR,NUMEROSERIE"
     Set rs = READSET(SQLX$)
     With rs
        Do While Not .EOF
           i& = i& + 1
           Me.MSFSegunOf.Rows = i& + 1
           Me.MSFSegunOf.TextMatrix(i&, 1) = SAFETEXT$(!FECHALTA)
           Me.MSFSegunOf.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
           Me.MSFSegunOf.TextMatrix(i&, 3) = SAFETEXT$(!NumeroSerie)
           Me.MSFSegunOf.TextMatrix(i&, 4) = XVALO(!TRAZAXSERIE_ID)
           If XVALO(!StatuOrf) <> 3 Then
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFSegunOf, &HFF0000, i&)
            Else
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFSegunOf, vbRed, i&)
           End If

           .MoveNext
        Loop
     End With
     rs.Close: Set rs = Nothing
   End If
   '

End Sub

Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub mnuMenu2_Click()
  MENU2% = 2
  Hide
End Sub

Private Sub mnuBorrar_Click()
    If CLACONTRA% <> 1 Then Exit Sub
    On Error GoTo ERROR_GUARDAR

    FLEXCONN.BeginTrans

    Dim rs As ADODB.Recordset
    SQLX$ = "SELECT *  FROM TRAZAXSERIE "
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE LIKE '184-9%' ORDER BY NUMEROSERIE"
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
      Do While Not .EOF
         NS$ = TRIM$(SAFETEXT$(!NumeroSerie))
         ORX% = XVALO(Right(NS$, 3))
         ORX% = ORX% + 452
         !NumeroSerie = Left$(NS$, 5) & TRIM$(SAFETEXT$(ORX%))
         .Update
         .MoveNext
      Loop
    End With
    rs.Close: Set rs = Nothing
    
    Dim rs1 As ADODB.Recordset
    SQLX$ = "SELECT *  FROM TANUMSE "
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE LIKE '184-9%' ORDER BY NUMEROSERIE"
    Set rs1 = New ADODB.Recordset
    rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs1
      Do While Not .EOF
         NS$ = TRIM$(SAFETEXT$(!NumeroSerie))
         ANOX = SAFETEXT$(!MESANIO)
         ORX% = XVALO(Right(NS$, 3))
         ORX% = ORX% + 452
         !NumeroSerie = Left$(NS$, 5) & TRIM$(SAFETEXT$(ORX%))
         !MESANIO = Left(ANOX, 2)
         !NUMEORD = XVALO(!NUMEORD) + 452
         .Update
         .MoveNext
      Loop
    End With
    rs1.Close: Set rs1 = Nothing
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("PROCESO TERMINADO")
    End If

End Sub



Private Sub MNUBORRARTABLA_Click()

End Sub

Private Sub mnuAgregarItemaOF_Click()
10  NORFA$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NORFA$) = "" Then
         Exit Sub
     End If
     '
     NORFA$ = AJUSTI$(NORFA$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 10
     End If
     '
40   OBX$ = OBJPLA$("CODICANT", VDEF$)
     VDEF$ = OBX$
     If OBX$ = "" Then
        GoTo 10
     End If
     '
     CICOMP% = CVI(Left$(OBX$, 2))
     CANTI = CVS(Mid$(OBX$, 3, 4))
     If CICOMP% < 1 Then
        Call COMUNI("Falta Seleccionar Código")
        GoTo 40
     End If
     
     If CANTI < 0.0005 Then
        Call COMUNI("Falta Ingresar Cantidad o No Válida")
        GoTo 40
     End If
     '
     TR% = ESTRAZABLE%(CICOMP%, 1)
     If TR% = 2 Or TR% = -1 Then
        Call COMUNI("Imposible Agregar Códigos con Trazabilidad de N° de Serie como Componente")
        GoTo 40
     End If
     '
     On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     
     Call AGREGAITEM_A_RESERVA(NORFA$, CICOMP%, CANTI)
     
     If TR% = 1 Then
       
       SQLX$ = "select numeroserie  from trazaxserie WITH(NOLOCK) "
       SQLX$ = SQLX$ + " where idsubor = '" & NORFA$ & "' AND TRAZABLE = 1 AND ( STATUORF < 3  or STATUORF  is null ) "
       SQLX$ = SQLX$ + " group by numeroserie"
       Set rs = READSET(SQLX$)
       With rs
           Do While Not .EOF
             NSERIE$ = SAFETEXT$(!NumeroSerie$)
             ValProxOrdTraza% = 1 + XVALO(VALOBADA("TRAZAXSERIE", "MAX(ORD)", "idsubor = '" & NORFA$ & "'   AND NUMEROSERIE = '" & NSERIE$ & "' AND TRAZABLE = 1", "NOMESS"))
             
             IDTRAZA& = XVALO(VALOBADA("TRAZAXSERIE", "TRAZAXSERIE_ID", "idsubor = '" & NORFA$ & "'   AND NUMEROSERIE = '" & NSERIE$ & "' AND TRAZABLE = 1  AND  (STATUORF < 3 or STATUORF  is null )", "NOMESS"))
             SQLX$ = " INSERT INTO TRAZAXSERIE  (CodiEmpr, IDSUBOR, FechAlta, Cod_Inte, COD_EXTE, Descripcion, NumeroSerie,TRAZABLE,IDENLOTE,CANTI,ORD)"
             
             SQLX$ = SQLX$ + " SELECT CodiEmpr, IDSUBOR, FechAlta "
             SQLX$ = SQLX$ + "," & CICOMP%
             SQLX$ = SQLX$ + ", '" & CODEXT$(CICOMP%) & "'"
             SQLX$ = SQLX$ + ", '" & DESCRIT(CICOMP%) & "'"
             SQLX$ = SQLX$ + ", NumeroSerie "
             SQLX$ = SQLX$ + "," & TR%
             SQLX$ = SQLX$ + ",''"
             SQLX$ = SQLX$ + "," & CANTI
             SQLX$ = SQLX$ + "," & ValProxOrdTraza%
             SQLX$ = SQLX$ + "  FROM TRAZAXSERIE "
             SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE =  '" & NSERIE$ & "' AND TRAZAXSERIE_ID = " & IDTRAZA& 'condiciono el idtraza para que solo sea un registro
             FLEXCONN.Execute (SQLX$)
             .MoveNext
           Loop
       End With
     End If
     rs.Close
     Set rs = Nothing
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Se Ha Ingresado El Item")
    End If
    
End Sub
Private Sub mnuAnularOF_Click()
   'VALIDACIONES
   If DERACCE%("ANULAORF", "ANULACIÓN DE ORDEN DE FABRICACIÓN") < 2 Then Exit Sub
   '
   NORFA$ = TRIM$(Label13)
   If NORFA$ = "" Then
      Call COMUNI("No Hay una Orden de Fabricación Activa")
      Exit Sub
   End If
   '
   NSERIE$ = TRIM$(Label2(4))
   If NSERIE$ = "" Then
      Call COMUNI("Falta Seleccionar Un Número de Serie")
      Exit Sub
   End If
   CI1% = XVALO(TXTCI)
   
   If STATORFA_TRAZAXSERIE%(NORFA$, NSERIE$) <> 3 Then
      Call COMUNI("La Orden de Fabricación No Está Cerrada")
      Exit Sub
   End If
   '
   ABUSX$ = "CONSUMO " & TRIM$(Label44)
   CONDIMOV$ = "CODIMOVI= 'CF' AND DOPRILAR = '" & NORFA$ & "'"
   CONDIMOV$ = CONDIMOV$ + " AND REFERCOR =  '" & ABUSX$ & "'"
   CONDIMOV$ = CONDIMOV$ + " AND CODIMOVI = 'CF' AND PLC_Has = '" & NSERIE$ & "'"
   If CantRegistros("MOVISTO", CONDIMOV$, "NOMESS") < 1 Then
      OPX% = COMALTER%("Atención!!\Confirma Anular Esta Orden de Fabricación\\Cancelar\Continuar y Anular O.F.")
      If OPX% < 2 Then Exit Sub
   End If
   'FIN VALIDACIONES
   '
   'TRANSACCIONES
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   Call BORRAREGISTROS("MOVISTO", CONDIMOV$, REGAF&, "NOMESS")
   'BORRAR LA MARCA DE CERRADA EN TRAZAXSERIE
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
   Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", CONDI$, 0, REGAF&, "NOMESS")
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
   Call ACTUREGISTRO("TRAZAXSERIE", "StatuOrf", CONDI$, 0, REGAF&, "NOMESS")
   Call ACTUREGISTRO("TRAZAXSERIE", "CONSUMIDO", CONDI$, 0, REGAF&)
   'VUELVO A 0 EL DEPOSITO EN TANUMSE
   Call GENERES_ANUCIEOF(CI1%, 1, NORFA$)   ' SIEMPRE ANULA DE A UNA
   'FIN TRANSACCIONES
   '
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
    End If
   '
   'REFRESCA LA INFO
    NORFA$ = Label13
    Label13 = ""
    Label13 = NORFA$

99
End Sub

Private Sub mnuCierrreOFActiva_Click()
   NORFA$ = TRIM$(Label13)
   If NORFA$ = "" Then
      Call COMUNI("No Hay una Orden de Fabricación Activa")
      Exit Sub
   End If
   '
   NSERIE$ = TRIM$(Label2(4))
   If NSERIE$ = "" Then
      Call COMUNI("Falta Seleccionar Un Número de Serie")
      Exit Sub
   End If
'
   CI1% = XVALO(TXTCI)
'   lo hace mas abajo recorriendo la grilla
'   If FALTAN_INFORMAR_LOTES%(CI1%, NORFA$, NSERIE$) > 0 Then
'      Call COMUNI("Imposible Continuar !!! Faltan Asignar Lotes ")
'      Exit Sub
'   End If
   
   If GrabarLotes$ <> "OK" Then Exit Sub
   
   If STATORFA_TRAZAXSERIE%(NORFA$, NSERIE$) = 3 Then
      Call COMUNI("La Orden de Fabricación Ya ha Sido Cerrada")
      Exit Sub
   End If
   '
   '====================================================================
   'SI EN EL PRIMER NIVEL DE LA OF ACTIVA ESTAN LOS LOTES
   For i& = 1 To MSFRESERVA.Rows - 1
       LOTE$ = TRIM$(Me.MSFRESERVA.TextMatrix(i&, 5))
       If LOTE$ = "" Then
          Call COMUNI("Imposible Continuar Falta Asignar Lotes a Componentes Críticos")
          Exit Sub
       End If
   Next i&

   'IF VALIDAR_OF_VINCULADAS_CERRADAS%(NORFA$,
   If VALIDAR_OF_VINCULADAS_CERRADAS%(CI1%, NSERIE$, RETORNO$) > 0 Then
      Call COMUNI("Imposible Cerrar la Orden de Fabricación Seleccionada\Faltan Cerrar Otras Pertenecientes a La Estructura\Qué Aún Estan Abiertas\(" & RETORNO$ & ")")
      Exit Sub
   End If
   '
   TABLA$ = TRIM$(REFE_TABLA_CONTROL_EYM$(CI1%)) ' TRAE EL NOMBRE DE LA TABLA
   
   If TABLA$ <> "" Then 'SI DEVUELVE LA TABLA ES POR QUE ESTA EN LA LISTA DE LOS QUE LLEVAN CONTROL
    If RESULTADO_CONTROL_EYM$(TABLA$, "RESULTADO", NORFA$, NSERIE$) <> "C" Then
         OPX% = COMALTER%("Orden de Fabricación Sin Control Aprobado\\Cancelar\Continuar y Cerrar Orden Activa")
         If OPX% < 2 Then Exit Sub
    End If
   End If
   '
'   TIPOCIERRE% = COMALTER%("Cerrar la Orden de Fabricación Para el N° de Serie Activo\\Con Consumo de Componentes y Alta de Producto\Solo Marcar Como Cerrada sin Mov. de Stock")
'
'   If TIPOCIERRE% <= 1 Then ' SOLO VALIDA STOCKS SI ES UN CIERRE CON CONSUMO
    VALSTOCK% = VALIDA_STOCK_COMPONENTES%(NORFA$, NSERIE$)
    If VALSTOCK% <= 0 Then
      If VALSTOCK% < 0 Then
       Call COMUNI("Imposible Continuar Se han Detectado Componentes Con Faltantes de Stock")
      End If
      Exit Sub
    End If
'   End If
   '

   Call BLOQUEAR_CONTROLES(LOTEXSQR, 2, 1, 1, 1)
   'Call Command8_Click 'GRABA LA INFORMACION QUE ESTA EN PANTALLA, POR SI REALIZO ALGUN CAMBIO
   On Error GoTo ERROR_GUARDAR

   FLEXCONN.BeginTrans
   '
'   If TIPOCIERRE% > 1 Then 'CIERRE SIN MOVIMIENTO DE STOCKS - SOLO MARCA EL ESTADO DE LA ORDEN EN TRAZAXSERIE
'      Call ACTUREGISTRO("TRAZAXSERIE", "StatuOrf", "IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'", 3, REGAF&, "NOMESS")
'      Call ActualizaTanumseSiesProductoTerminado(NORFA$, NSERIE$)
'      GoTo 80
'   End If

   'Call GENERAR_RESERVAS_DE_TRAZABLE_EYM(NORFA$, NSERIE$)
   
   Call Cierre_Orden_Fabriacion_EYM(NORFA$, NSERIE$)
   
   Screen.MousePointer = 1
   
80 CantidadProyectada = CantiProyectadaOF(NORFA$)
   CantidadAprobada = CantiAprobadaOF(NORFA$)
   If CantidadAprobada >= CantidadProyectada Then
      Call ACTUREGISTRO("ORDEFABR", "StatuOrf", "IDSUBOR = '" & NORFA$ & "'", 3, REGAF&, "NOMESS")
   Else
       MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
       If MDXX% = 2 Then
          Call ACTUREGISTRO("ORDEFABR", "StatuOrf", "IDSUBOR = '" & NORFA$ & "'", 3, REGAF&, "NOMESS")
       End If
   End If
   '
'   Call CIERRORFA_EYM(NORFA$, 1, NSERIE$)
90
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
    End If

   'RECORDAR QUE EL CONSUMO VA A SER LOS TRAZABLES Y LOS NO TRAZABLES.
   'EN EL MOVISTO SE DEBERA PONER COMO IDENLOTE EL N.SERIE CUANDO CORRRESPONDA
   'VER CUANDO SE CAMBIA EN EL TANUMSE LA OF Y EL CODIGO.
   '
   'REFRESCA LA INFO
    NORFA$ = Label13
    Label13 = ""
    Label13 = NORFA$

99  Call BLOQUEAR_CONTROLES(LOTEXSQR, 1, 1, 1, 1)

End Sub


Private Sub mnuconfigimp_Click()
   FSTF054.Show 1
End Sub

Private Sub mnuControles_Click()
    CONDI$ = " LISTA ='CONFICONTROL'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      Call PRESENTA_CONTROLES
      GoTo 99
    End If
99
End Sub

Private Sub mnuDesvioTransitorio_Click()
    If TRIM$(Label13) = "" Then
       Call COMUNI("No se ha Seleccionado Ningúna Orden de Facturación")
       Exit Sub
    End If
    '
    NSERIE$ = TRIM$(Label2(4))
    If NSERIE$ = "" Then
       Call COMUNI("Falta Seleccionar Un Número de Serie")
       Exit Sub
    End If

    If MSFRESERVA.Rows - 1 = 0 Then Exit Sub
    
    Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;N.Serie/A16;Lote/A16;ID/A0;CI/A0;N°D.Trans./f11"

10  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Width = 12000
    
    Call CopiarMsfAOtroMsf(Me.MSFRESERVA, FRMZELECHO.MSF2, 1, MSFRESERVA.Width)
    FRMZELECHO.Show 1
    If TRIM(RESP_ZELE_VECT(4)) = "" Then Exit Sub
    NSERIE$ = TRIM(RESP_ZELE_VECT(4))
    PARTEX$ = TRIM(RESP_ZELE_VECT(5))
    ID_TRASERIE& = XVALO(RESP_ZELE_VECT(6))
    CI1% = XVALO(RESP_ZELE_VECT(7))
    DESVTRAN = XVALO(RESP_ZELE_VECT(8))
    
    OBSERVAX$ = ""
    If DESVTRAN > 0 Then
       OBSERVAX$ = VALOBADA("DES_TRANS", "REF_DESVIO", "DESV_TRANS = " & DESVTRAN, "NOMESS")
    End If
    
30  REFERENCIA$ = TRIM$(InputBox$("Observaciones", "Ingrese Observaciones", OBSERVAX$))
    If REFERENCIA$ = "" Then GoTo 10
    Call PRINTETI.Alta_Actu_Desvio_Transitorio(DESVTRAN, PARTEX$, CI1%, REFERENCIA$, NSERIE$, ID_TRASERIE&)
    
   'REFRESCA LA INFO
    GoTo 10
End Sub

Private Sub mnuDesvTrans_Click()
   Screen.MousePointer = 11
20 Campos$ = "N° Desvío/F8;Referencia/A32;Código/A16;Descripción/A24;Lote/A16;N° Serie/A16"
   
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   J& = 0
   SQLX$ = " select DESV_TRANS, REF_DESVIO,IDENLOTE,NUMEROSERIE,COD_EXTE,COD_INTE  "
   SQLX$ = SQLX$ + " FROM des_trans WITH(NOLOCK) order by cod_inte"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
        J& = J& + 1
        '
        NRO& = XVALO(!DESV_TRANS)
        REFDES$ = SAFETEXT$(!REF_DESVIO)
        CI1% = XVALO(!COD_INTE)
        CODI$ = SAFETEXT$(!COD_EXTE)
        DESX$ = DESCRIT0$(CI1%)
        LOTE$ = SAFETEXT$(!idenlote)
        NS$ = SAFETEXT$(!NumeroSerie)
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = NRO&
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = REFDES$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = CODI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = DESX$
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = LOTE$
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = NS$
       .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   '
   Screen.MousePointer = SCMA
    
    TX$ = "Desvios Transitorios" & PERIO$
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    
'    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
    NRO_DESV_TRANS& = 0
    GRILLAGRAL.Show 1
    
    NRO_DESV_TRANS& = XVALO(RESP_ZELE_VECT(1))
    If NRO_DESV_TRANS& > 0 Then
     'PARA QUE SE VEA DIFERENTE Y NO SEA CONFUSO
     GRILLAGRAL.Frame1.BackColor = &H9FBE85
     GRILLAGRAL.Frame2.BackColor = &H9FBE85
     GRILLAGRAL.Option1.BackColor = &H9FBE85
     GRILLAGRAL.Option2.BackColor = &H9FBE85
     GRILLAGRAL.Option3.BackColor = &H9FBE85
     GRILLAGRAL.BackColor = &H9FBE85
     Call ReporteDondeSeUtilizoDesvioTransitorio(NRO_DESV_TRANS&)
     
     For i& = 0 To UBound(RESP_ZELE_VECT)
       RESP_ZELE_VECT(i&) = ""
     Next i&
     GRILLAGRAL.Frame1.BackColor = &H8000000F
     GRILLAGRAL.Frame2.BackColor = &H8000000F
     GoTo 20
   End If
    
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DESVTRAN", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

End Sub

Sub ReporteDondeSeUtilizoDesvioTransitorio(NroDesvTransitorio&)
10  J& = 0
    
    Campos$ = "CI./a0;" '1
    Campos$ = Campos$ + "Código/a16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "Lote./A12;"      '4
    Campos$ = Campos$ + "O.F./A12;"         '5
    Campos$ = Campos$ + "N°Serie/A10;"       '6
    Campos$ = Campos$ + "Referencia./A60;"   '7
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    SQLX$ = "select T.cod_inte,T.Cod_Exte,T.idenlote,T.NUMEROSERIE,D.REF_DESVIO,T.DESV_TRANS,T.IDSUBOR"
    SQLX$ = SQLX$ + " from TRAZAXSERIE T WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN DES_TRANS  D ON D.DESV_TRANS = T.DESV_TRANS"
    SQLX$ = SQLX$ + "  WHERE T.DESV_TRANS = " & NroDesvTransitorio&
    SQLX$ = SQLX$ + "  ORDER BY T.NumeroSerie,T.IDENLOTE"
    Set rs = READSET(SQLX$)
    With rs
     Do While Not .EOF
        CI1% = XVALO(!COD_INTE)
        Codigo$ = SAFETEXT$(!COD_EXTE)
        DESXRI$ = DESCRIT0$(CI1%)
        LOTE$ = SAFETEXT$(!idenlote)
        NORFA$ = SAFETEXT$(!IDSUBOR)
        NUSERIE$ = SAFETEXT$(!NumeroSerie)
        REFERENCIA$ = SAFETEXT$(!REF_DESVIO)
        
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = LOTE$
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = NORFA$
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = NUSERIE$
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = REFERENCIA$
       .MoveNext
     Loop
    End With
    rs.Close
    Set rs = Nothing
    GRILLAGRAL.Width = 12500
    If J& = 0 Then Call COMUNI("No Hay Registros Con la Condición Seleccionada"): Exit Sub
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    
    GRILLAGRAL.Caption = "Detalle de Aplicaciones de Desvios Transitorios: " & REFERENCIA$
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOSCRP", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    
'   NORFA$ = OrdenFabricacionAbiertas$(1)
'   If TRIM$(NORFA$) <> "" Then GoTo 10
End Sub


Private Sub mnuEngancharNSerieAOF_Click()
   ENGANNUMSERIE.Show 1
   NORFA$ = Label13
   Label13 = ""
   Label13 = NORFA$
End Sub

Private Sub MSFASIGNADOS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     
     If DERACCE("MODPAPRO", "MODIFICAR PARTE DE PRODUCCIÓN") < 2 Then Exit Sub
     Set MSFGLOBALFRM = MSFASIGNADOS
     GLOBALFRM_INDIX% = Index
     REG& = MSFASIGNADOS.MouseRow
     col% = MSFASIGNADOS.MouseCol
  End If
  
End Sub
 Sub INVICTRL()
   mnucopiar.Visible = Not (mnucopiar.Visible)
   mnuQuitar.Visible = Not (mnuQuitar.Visible)
End Sub
Sub MENU_CELDA(MSF As msFlexGrid, REG&, col%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(col%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, col%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu mnuArchivo
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, col%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

Private Sub mnuEtiqAut_Click()
        Call IMPRIMIR_ETIQUETAS_F054_EYM
        PDES$ = Control$(IDENTER$, "PRTQEYM1")
        If PDES$ = "" Then PDES$ = Control$("*", "PRTQEYM1")
        RUTAEXCEL$ = Control("TRAZABLE", "EXCEAUTO")
        If VENTABNEW.APROBADO% > 0 Then
           Call Param_Print_Excel(RUTAEXCEL$, "Hoja3", VENTABNEW.MSF, PDES$, "0;1;3", 6, 13, "1")
        End If
        Unload VENTABNEW

End Sub

Private Sub mnuEtiqFirma_Click()
        Call IMPRIMIR_ETIQUETAS_F054_EYM
        PDES$ = Control$(IDENTER$, "PRTQEYM2")
        If PDES$ = "" Then PDES$ = Control$("*", "PRTQEYM2")
        RUTAEXCEL$ = Control("TRAZABLE", "EXCEAUTO")
        
        If VENTABNEW.APROBADO% > 0 Then
           Call Param_Print_Excel(RUTAEXCEL$, "Hoja1", VENTABNEW.MSF, PDES$, "0;1;3", 6, 13, "1")
           Unload VENTABNEW
        End If
End Sub


Private Sub mnuGennserie_Click()
    GENUMSE07.Show 1
End Sub

Private Sub mnuGLA_Click()
  If CLACONTRA% <> 1 Then Exit Sub
  '
  'ESTE PROCESO LE APLICA UN LOTE AUTOMATICO A TODOS LOS ITEMS TRAZABLES EN EL MOVISTO
  'EL FORMATO DEL LOTE ES: "LA-" & FORMATOCON0$(XVALO(!COD_INTE), 6) & "-1" =    LA-000389-1
  'ACTUALIZA EL LOTE EXISTENTE (QUE DEBERIA QUEDAR EN 0
  'ACTUALIZA EL LOTE AUTOMATICO Y TAMBIEN HACE UN CREALOTIN
  Dim rs As ADODB.Recordset
  SQLX$ = "SELECT MA.CODIGO, MO.COD_INTE,MO.IDENLOTE,MO.DOCUORIG  FROM MOVISTO MO INNER JOIN MASTER MA"
  SQLX$ = SQLX$ + " ON MA.CODINT = MO.COD_INTE "
  SQLX$ = SQLX$ + " WHERE MA.TRAZABLE = 1"
  SQLX$ = SQLX$ + " ORDER BY MO.COD_INTE"
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
    HOII% = HOY(HOS$)
    VUELTA% = 0: J& = 0
    Do While Not .EOF
       If VUELTA% < 1 Then
          FIN& = XVALO(.RecordCount)
          VUELTA% = 1
       End If
       J& = J& + 1
       Call TRACE(20, J&, FIN&)
       CI1% = SAFETEXT$(!COD_INTE)
       LOTE$ = SAFETEXT$(!idenlote$)
       LALOTE$ = "LA-" & FORMATOCON0$(XVALO(!COD_INTE), 6) & "-1"
       !idenlote = LALOTE$
       !DOCUORIG = LOTE$
       .Update
       .MoveNext
       If XVALO(!COD_INTE) <> CI1% And J& <> 1 Then
          Call ACSALOTE(CI1%, LOTE$, "NOMESS")
          Call CREALOTIN(HOII%, 0, 0, CI1%, LALOTE$, "", "", 1)
          Call ACSALOTE(CI1%, LALOTE$)
       End If
    Loop
  End With
  rs.Close
  Set rs = Nothing
  
End Sub

Private Sub mnuGurdaDatosCPU_Click()
        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        
        NSERIE$ = TRIM$(Label2(4))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        
        CI1% = XVALO(TXTCI.TEXT)
        CODIGOEXT$ = TRIM$(CODEXT$(CI1%))
        
        If Left(CODIGOEXT$, 2) = "DB" Then
            Call GuardarDatosCPUaCodigosDB
            Exit Sub
        End If
        
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        'VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        
        VENTANA.CampEditables = "3;4"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Datos CPU (Versión Software - Número de serie CPU) Para N° Serie: " & NSERIE$
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 4500
        VENTABNEW.Width = 11500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        '
        'VECTOR CARGO LOS DOS  PRIMEROS REGISTROS EN LAS 2 COLUMNAS CON EL DATO FIJO
        Dim CODI$(3), DESX$(3)
        CODI$(1) = "Responsable": DESX$(1) = "Preparado Por: "
        CODI$(2) = "Software CPU": DESX$(2) = "Versión Software CPU"
        CODI$(3) = "Cpu Nº:": DESX$(3) = "Número de serie CPU"

        '
        VENTABNEW.MSF.Rows = 8
        VENTABNEW.MSF.TextMatrix(1, 1) = CODI$(1)
        VENTABNEW.MSF.TextMatrix(1, 2) = DESX$(1)
        VENTABNEW.MSF.TextMatrix(2, 1) = CODI$(2)
        VENTABNEW.MSF.TextMatrix(2, 2) = DESX$(2)
        VENTABNEW.MSF.TextMatrix(3, 1) = CODI$(3)
        VENTABNEW.MSF.TextMatrix(3, 2) = DESX$(3)
        
        'FIN DE CARGA POR DEFECTO
        
        'CARGO PRIMERO LOS DATOS DEL CPU
        SQLX$ = "SELECT CODIGO,DESCRIPCION,IDENTIFICACION,DESV_TRANS,DATOSCPU_ID FROM DATOSCPU WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CI1%
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            If k& <= 3 Then
              VENTABNEW.MSF.TextMatrix(k&, 1) = CODI$(k&)
              VENTABNEW.MSF.TextMatrix(k&, 2) = DESX$(k&)
            Else
              VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!Codigo)
              VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            End If
            VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!IDENTIFICACION)
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = SAFETEXT$(!DATOSCPU_ID)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% > 0 Then
            On Error GoTo ERROR_GUARDAR
            FLEXCONN.BeginTrans
            Dim rs1 As ADODB.Recordset
            Set rs1 = New ADODB.Recordset
            SQLX$ = "SELECT * FROM DATOSCPU WHERE NUMEROSERIE=''" '
            rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            J& = 0
            
            For k& = 1 To VENTABNEW.MSF.Rows - 1
              ID = XVALO(VENTABNEW.MSF.TextMatrix(k&, 5))
              If ID > 0 Then
                SQLX$ = "UPDATE DATOSCPU "
                SQLX$ = SQLX$ + " SET Codigo = '" & VENTABNEW.MSF.TextMatrix(k&, 1) & "'"
                SQLX$ = SQLX$ + " , DESCRIPCION = '" & VENTABNEW.MSF.TextMatrix(k&, 2) & "'"
                SQLX$ = SQLX$ + " , IDENTIFICACION = '" & VENTABNEW.MSF.TextMatrix(k&, 3) & "'"
                SQLX$ = SQLX$ + " , DESV_TRANS = '" & VENTABNEW.MSF.TextMatrix(k&, 4) & "'"
                SQLX$ = SQLX$ + "  WHERE DATOSCPU_ID = " & ID
              Else
                If TRIM$(VENTABNEW.MSF.TextMatrix(k&, 1)) <> "" Then
                    SQLX$ = "INSERT INTO DATOSCPU ("
                    SQLX$ = SQLX$ + "  Codigo,DESCRIPCION,IDENTIFICACION,DESV_TRANS,IDSUBOR,NUMEROSERIE,COD_INTE)"
                    SQLX$ = SQLX$ + " VALUES('" & VENTABNEW.MSF.TextMatrix(k&, 1) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 2) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 3) & "'"
                    SQLX$ = SQLX$ + " ,'" & VENTABNEW.MSF.TextMatrix(k&, 4) & "'"
                    SQLX$ = SQLX$ + " ,'" & Label13 & "'"
                    SQLX$ = SQLX$ + " ,'" & Label2(4) & "'"
                    SQLX$ = SQLX$ + " ," & CI1%
                    SQLX$ = SQLX$ + ")"
                End If
              End If
              FLEXCONN.Execute (SQLX$)
           Next k&
           FLEXCONN.CommitTrans
           Exit Sub
        End If
        
        
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
    End If

End Sub

Private Sub mnuHNS_Click()
    CONUSE07.Show 1
End Sub

Private Sub mnuImpoRegArmado_Click()
   ImpoRegAr.Show 1
End Sub

Private Sub mnuInicalizaDVT_Click()
    'DESVIO TRANSITORIO
30  REFERENCIA$ = TRIM$(InputBox$("Inicializar", "Ingrese Último Número de Desvío Transitorio Utilizado", NUME$))
    If XVALO(REFERENCIA$) < 1 Then Exit Sub
    Call INGRESO_n_d_t(XVALO(REFERENCIA$))

End Sub

Private Sub mnuPrintEmbalajeC15_Click()

    NSERIE$ = RegCFinC15Eym.SERIES_CONTROLFINAL_C15
    If NSERIE$ = "" Then Exit Sub
    
    PDES$ = Control$(IDENTER$, "PRIREGCF") 'la configuracion ya esta. ESTA YA ESTABA
    If PDES$ = "" Then PDES$ = Control$("*", "PRIREGCF")

    PIMP_HOJA_NORMAL$ = Control$(IDENTER$, "PRTQEYM2") 'HOJA NORMAL CONFIGURADO EN FSTF054
    If PIMP_HOJA_NORMAL$ = "" Then PIMP_HOJA_NORMAL$ = Control$("*", "PRTQEYM2")
    
    PIMP_ETIQ_GDE$ = Control$(IDENTER$, "PORTSTIK")   'ETIQUETA GRANDE EN FSTF054
    If PIMP_ETIQ_GDE$ = "" Then PIMP_ETIQ_GDE$ = Control$("*", "PORTSTIK")
    
    RUTAEXCEL$ = Control$("ETIQUC15", "ETEMBC15")
    If EXISTE%(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Archivo Excel de Etiqueta de Embalaje\Configuración dentro del Formulario de Control Final C15")
       'G
       Exit Sub
    End If
    '
    
     OPX% = COMALTER%("Seleccione Impresora Configurada\\Cancelar\Imp. de Reg.Datos\Imp. Hoja Normal\Imp. Etiq. Gde.")
     If OPX% < 2 Then Exit Sub
     Select Case OPX%
      Case 3
        PDES$ = PIMP_HOJA_NORMAL$
      Case 4
        PDES$ = PIMP_ETIQ_GDE$
      Case Else
        Exit Sub
     End Select
     
    'CARGO LOS DATOS QUE YA LOS TENGO EN LA GRILLA DE SELECCION
    CODIGOX$ = TRIM$(RESP_ZELE_VECT(2))
    If Right(CODIGOX$, 1) <> "+" Then CODIGOX$ = CODIGOX$ & "+"
    Descripcion$ = TRIM$(RESP_ZELE_VECT(3))
    NORFA$ = TRIM$(RESP_ZELE_VECT(4))
    FEALTA$ = TRIM$(RESP_ZELE_VECT(5))
    '
    Unload VENTABNEW
    VENTABNEW.MSF.Rows = 6
    VENTANA.Inicializar = 0
    Campos$ = "DATOS/A48"
    VENTANA.Campox = Campos$
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
      
    VENTANA.CampEditables = ""
    VENTANA.AgregaRegistro = 1
    VENTANA.Inicializar = 1
    VENTABNEW.Height = 7500
    VENTABNEW.MSF.Font = "Roboto Mono"
    VENTABNEW.MSF.Font.Size = 8
    'CARGO MANULAMENTE LOS DATOS EN LA MISMA COLUMNA (MISMO ORDEN EN QUE SE VAN A IMPRIMIR EN EL EXCEL
    VENTABNEW.MSF.TextMatrix(1, 1) = CODIGOX$
    VENTABNEW.MSF.TextMatrix(2, 1) = CODIGOX$
    
    VENTABNEW.MSF.TextMatrix(3, 1) = NSERIE$
    VENTABNEW.MSF.TextMatrix(4, 1) = NSERIE$
    
    If FEALTA$ <> "" Then
      FE% = FECHANUM(FEALTA$)
      ANO = Right(FETEXCOR1$(FE%), 4)
      VENTABNEW.MSF.TextMatrix(5, 1) = ANO
    End If
    
      ''''''''''''''''''''''''''''''''''
     IMAG$ = "QR_ETIEMBALAJE.JPG"
     IDENTIFICADOR$ = "QR_ETIEMBALAJE"

     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     
     TEXTOQR$ = CODIGOX$ & "|" & NSERIE$
     On Error GoTo 0
     
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", 0, 0))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)

    
    ''''''''''''''''''''''''''''''''''
    'VENTABNEW.Show 1
    'Etiq embalaje es la  hoja
    Call Param_Print_Excel(RUTAEXCEL$, "Etiq embalaje", VENTABNEW.MSF, PDES$, "1", 9, 13, "", 1, 0, IMAG$)      'imprime en excel la 3 columna de mi msf en la columna 1 a partir de la fila 54 en el excel  (este imprime con los 3 cambios )

'      Param_Print_Excel

End Sub

Private Sub mnuPrintEtiqEquipoC15_Click()
    NSERIE$ = RegCFinC15Eym.SERIES_CONTROLFINAL_C15
    If NSERIE$ = "" Then Exit Sub
    PDES$ = Control$(IDENTER$, "PRIREGCF") 'la configuracion ya esta.
    If PDES$ = "" Then PDES$ = Control$("*", "PRIREGCF")

    RUTAEXCEL$ = Control$("ETIQUC15", "ETEQUC15")
    If EXISTE%(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Archivo Excel de Etiqueta de Equipo\Configuración dentro del Formulario de Control Final C15")
       'G
       Exit Sub
    End If

      
    'CARGO LOS DATOS QUE YA LOS TENGO EN LA GRILLA DE SELECCION
    CODIGOX$ = TRIM$(RESP_ZELE_VECT(2))
    Descripcion$ = TRIM$(RESP_ZELE_VECT(3))
    NORFA$ = TRIM$(RESP_ZELE_VECT(4))
    FEALTA$ = TRIM$(RESP_ZELE_VECT(5))
    '
    Unload VENTABNEW
    VENTABNEW.MSF.Rows = 4
    VENTANA.Inicializar = 0
    Campos$ = "DATOS/A48"
    VENTANA.Campox = Campos$
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
      
    VENTANA.CampEditables = ""
    VENTANA.AgregaRegistro = 1
    VENTANA.Inicializar = 1
    VENTABNEW.Height = 7500
    VENTABNEW.MSF.Font = "Roboto Mono"
    VENTABNEW.MSF.Font.Size = 8
    'CARGO MANULAMENTE LOS DATOS EN LA MISMA COLUMNA (MISMO ORDEN EN QUE SE VAN A IMPRIMIR EN EL EXCEL
    VENTABNEW.MSF.TextMatrix(1, 1) = CODIGOX$
    VENTABNEW.MSF.TextMatrix(2, 1) = NSERIE$
    If FEALTA$ <> "" Then
      FE% = FECHANUM(FEALTA$)
      ANO = Right(FETEXCOR1$(FE%), 4)
      VENTABNEW.MSF.TextMatrix(3, 1) = ANO
    End If
    'VENTABNEW.Show 1
    'Etiq embalaje es la  hoja
    Call Param_Print_Excel(RUTAEXCEL$, "Etiq equipo", VENTABNEW.MSF, PDES$, "1", 5, 14, "", 1, 0)        'imprime en excel la 3 columna de mi msf en la columna 1 a partir de la fila 54 en el excel  (este imprime con los 3 cambios )


End Sub


Private Sub mnuquitar_Click()
       FILX& = GLOBALFRM_FILA&
       Call cmdEliminaItem(MSFGLOBALFRM, GLOBALFRM_FILA&)
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub mnurecuuperalotes_Click()
   If CLACONTRA% < 1 Then Exit Sub
   Screen.MousePointer = 11
   Call ComponerLotesenTrazaxserie
   Call COMUNI("Proceso Terminado")
   Screen.MousePointer = 1
End Sub

Private Sub mnuRegistroArmado_Click()

        PDES$ = Control$(IDENTER$, "PRTQEYM3")
        If PDES$ = "" Then PDES$ = Control$("*", "PRTQEYM3")
        RUTAEXCEL$ = Control("TRAZABLE", "EXCEARMA")
        If Left(TRIM(UCase$(Label44)), 2) = "DB" Then
           Call IMPRIMIR_REG_ARMADO_C15_EYM
        Else
           Call IMPRIMIR_REG_ARMADO_EYM
        End If
'        If VENTABNEW.APROBADO% > 0 Then
'           Call Param_Print_Excel(RUTAEXCEL$, "Hoja1", VENTABNEW.MSF, PDES$, "1;2;3;4", 16, 9)
'           Unload VENTABNEW
'         End If
End Sub
Sub IMPRIMIR_REG_ARMADO_EYM()
        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Label2(4))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        'VALIDA SI YA SEW GRABARON LOS DATOS DEL CPU
        CIFAB% = XVALO(Me.TXTCI)
        
        CONDI$ = "NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CIFAB%
        If CantRegistros&("DATOSCPU", CONDI$, "NOMESS") < 3 Then
           OPX% = COMALTER%("No se Han Guardado Datos Del CPU o Datos Incompletos\\Cancelar\Continuar e Imprimir")
           If OPX% < 2 Then Exit Sub
        End If
        '
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        'VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        
        VENTANA.CampEditables = "1;2;3;4"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Impresión F054"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 7500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        FALTACPU% = 0
        'CARGO PRIMERO LOS DATOS DEL CPU
        SQLX$ = "SELECT CODIGO,DESCRIPCION,IDENTIFICACION,DESV_TRANS,DATOSCPU_ID,COD_INTE FROM DATOSCPU WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & XVALO(Me.TXTCI)
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!Codigo)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!IDENTIFICACION)
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!DATOSCPU_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!IDENTIFICACION)) = "" Then
               FALTACPU% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If
            
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        If k& < 3 Then k& = 2
        FALTALOTE% = 0
        CONDI$ = "COD_EXTE = '" & Label44 & "' AND  IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
        OF_NIVEL_SUPERIOR$ = VALOBADA("TRAZAXSERIE", "OF_ARRASTRA", CONDI$, "NOMESS")
        'TRAE TODOS LOS COMPONENTES CRITICOS CON EXCEPCION DE LOS DEL PRIMER NIVEL DE LA OF SELECCIONADA
        'SELECT * FROM TRAZAXSERIE WITH(NOLOCK)  WHERE NUMEROSERIE = '185-1473'  AND IDSUBOR <> 'OF-002221-1'
        SQLX$ = "SELECT COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION,DESV_TRANS FROM TRAZAXSERIE WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
        SQLX$ = SQLX$ + " AND  COD_EXTE <> '" & Label44 & "' AND  IDSUBOR <> '" & OF_NIVEL_SUPERIOR$ & "'"
        SQLX$ = SQLX$ + " ORDER BY  FECHALTA DESC"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            LOTEX$ = SAFETEXT$(!idenlote)
            LOTEX$ = REPLACAR(LOTEX$, "BR", "PR")
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = LOTEX$
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!TRAZAXSERIE_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!idenlote)) = "" Then
               FALTALOTE% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        '
30      VENTABNEW.Show 1 ' NO LO MUESTRO (ABAJO ESTA EL CODIGO SI SE MUESTRA PARA GRABAR MODIFICACIONES.
        If VENTABNEW.APROBADO < 1 Then Exit Sub
        
        If VENTABNEW.APROBADO > 0 And (FALTALOTE% + FALTACPU%) > 0 Then
             Call COMUNI("Imposible Continuar Faltan Datos de Lotes o del CPU ")
             GoTo 30
        End If
        Dim ObjetoExcel As Excel.Application
        Dim Libro As Excel.Workbook
        Dim Hoja As Excel.Worksheet
        
        PDES$ = Control$(IDENTER$, "PRTQEYM3")
        If PDES$ = "" Then PDES$ = Control$("*", "PRTQEYM3")
        RUTAEXCEL$ = Control("TRAZABLE", "EXCEARMA")
    
        Set ObjetoExcel = CreateObject("Excel.Application")
        Set Libro = ObjetoExcel.Workbooks.Open(RUTAEXCEL$)
        With ObjetoExcel
            '
            'LIMPIO PRIMERO LOS CAMOS EN EL EXCEL
            For R& = 16 To 37
                Libro.Worksheets("Hoja1").Cells(R&, 6) = "N/A"
                Libro.Worksheets("Hoja1").Cells(R&, 8) = " "
             Next R&
            PREPARADOPORX$ = VENTABNEW.MSF.TextMatrix(1, 3)
            
            'ESCRIBO LOS DATOS EN OTROS CASILLEROS QUE ESTAN REFERENCIADOS
            For i& = 2 To VENTABNEW.MSF.Rows - 1
                CODI$ = UCase$(TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)))
                IDENTIF$ = VENTABNEW.MSF.TextMatrix(i&, 3)
                DV$ = CEROIGUALVACIO(VENTABNEW.MSF.TextMatrix(i&, 4))
                For R& = 16 To 37
                   CODIEXCEL$ = UCase$(TRIM$(Libro.Worksheets("Hoja1").Cells(R&, 2)))
                   If CODIEXCEL$ = CODI$ Then
                      Libro.Worksheets("Hoja1").Cells(R&, 6) = IDENTIF$
                      Libro.Worksheets("Hoja1").Cells(R&, 8) = CEROIGUALVACIO(DV$)
                      Exit For
                   End If
                Next R&
            Next i&
    '            For J& = 1 To 4
    '               DATO$ = TRIM$(VENTABNEW.MSF.TextMatrix(i&, J&))
    '               If DATO$ = "" Then DATO$ = "-----"
    '               Select Case J&
    '                 Case 1: Libro.Worksheets("Hoja1").Cells(16 + i& - 1, 2 + J& - 1) = DATO$ 'imprime en columna 2
    '                 Case 2: Libro.Worksheets("Hoja1").Cells(16 + i& - 1, 2 + J& - 1) = DATO$ 'imprime en columna 3
    '                 Case 3: Libro.Worksheets("Hoja1").Cells(16 + i& - 1, 4 + J& - 1) = DATO$ 'imprime en columna 6
    '                 Case 4: Libro.Worksheets("Hoja1").Cells(16 + i& - 1, 5 + J& - 1) = DATO$ 'imprime en columna 8
    '
    '               End Select
    '            Next J&
    '         Next i&
             '
             'ESTOS DATOS SON DE ENCABEZADO LOS ESCRIBO DIRECTAMENTE EN LA CELDA CORRESPONDIENTE
             For i& = 1 To 7
              Select Case i&
                Case 1: Dato$ = Label44: Libro.Worksheets("Hoja1").Cells(4, 6) = Dato$
                Case 2: Dato$ = LBLREVISION: Libro.Worksheets("Hoja1").Cells(5, 6) = Dato$
                Case 3: Dato$ = Label2(4): Libro.Worksheets("Hoja1").Cells(6, 6) = Dato$
                Case 4:
                'valido por que aca va la fecha
                Dato$ = Label2(6)
'                If IsDate(DATO$) And ((Mid$(DATO$, 3, 1) = "-" And Mid$(DATO$, 6, 1) = "-") Or (Mid$(DATO$, 3, 1) = "/" And Mid$(DATO$, 6, 1) = "/")) Then
'                    Libro.Worksheets("Hoja1").Cells(7, 6).NumberFormat = "d/m/yyyy;@"
'                End If
                Libro.Worksheets("Hoja1").Cells(7, 6) = Dato$
                
                Case 5: Dato$ = Label13: Libro.Worksheets("Hoja1").Cells(8, 6) = Dato$
                Case 6: Dato$ = DESVIOTRANSITORIO$: Libro.Worksheets("Hoja1").Cells(9, 6) = Dato$
                Case 7: Dato$ = PREPARADOPORX$: Libro.Worksheets("Hoja1").Cells(10, 6) = Dato$
              End Select
             Next i&
             '
             'MUESTRO E IMPRIMO
             ObjetoExcel.ActiveWorkbook.Save
             Set Hoja = Libro.Sheets("Hoja1")
             Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=PDES$
             '
        End With
        
    
        'CIERRA HOJA, LIBRO Y EXCEL
99      On Error Resume Next
        Set Hoja = Nothing
        Libro.Close: Set Libro = Nothing
        ObjetoExcel.Close: Set ObjetoExcel = Nothing
        On Error GoTo 0
    
    '        '
    '
    '        If VENTABNEW.APROBADO% > 0 Then
    '            On Error GoTo ERROR_GUARDAR
    '            FLEXCONN.BeginTrans
    '            Dim RS1 As ADODB.Recordset
    '            Set RS1 = New ADODB.Recordset
    '            SQLX$ = "SELECT * FROM TRAZAXSERIE WHERE COD_INTE=0" '
    '            RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '            J& = 0
    '
    '            For REG& = 1 To VENTABNEW.MSF.Rows - 1
    '              ID = XVALO(VENTABNEW.MSF.TextMatrix(REG&, 5))
    '              If ID < 1 Then 'SI SON DATOS NUEVOS LOS GUARDA NO ASI CON LA YA EXISTENTE VALIDO CON EL ID
    '                CODI$ = VENTABNEW.MSF.TextMatrix(REG&, 1)
    '                LOTE$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 2))
    '                DESXRIP$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, 2))
    '                With RS1
    '50                  RS1.AddNew   ' YA CONVERTIDO
    '                    J& = J& + 1
    '                    On Error Resume Next
    '                    RS1!CodiEmpr = CodiEmp%
    '                    On Error GoTo 0
    '                    RS1!IDSUBOR = NORFA$
    '                    RS1!OF_ARRASTRA = ""
    '                    RS1!IDENLOTE = LOTE$
    '                    RS1!NUMEROSERIE = NSERIE$
    '                    RS1!TRAZABLE = 1
    '                    RS1!COD_INTE = CODINT(CODI$)
    '                    RS1!Descripcion = DESXRIP$
    '                    RS1!COD_EXTE = CODI$
    '                    RS1!CANTI = 1
    '                    RS1!FECHALTA = CVDAT(HOY(HO$))
    '                    RS1!ORD = J&
    '                    RS1!MARBORRA = 0
    '                    RS1.Update
    '                 End With
    '              End If
    '            Next REG&
    '            RS1.Close: Set RS1 = Nothing
    '        Else
    '            Exit Sub
    '        End If
    '
    'ERROR_GUARDAR:
    '    If Err.Number <> 0 Then
    '       FLEXCONN.RollbackTrans
    '       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
    '       On Error GoTo 0
    '       Exit Sub
    '    Else
    '      FLEXCONN.CommitTrans
    '      On Error GoTo 0
    '    End If
    '    '
End Sub

Sub IMPRIMIR_REG_ARMADO_C15_EYM()
        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13)
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Label2(4))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        'VALIDA SI YA SEW GRABARON LOS DATOS DEL CPU
        CIFAB% = XVALO(Me.TXTCI)
        
        CONDI$ = "NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CIFAB%
        If CantRegistros&("DATOSCPU", CONDI$, "NOMESS") < 3 Then
           OPX% = COMALTER%("No se Han Guardado Datos Del CPU o Datos Incompletos\\Cancelar\Continuar e Imprimir")
           If OPX% < 2 Then Exit Sub
        End If
        '
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        'VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        
        VENTANA.CampEditables = "1;2;3;4"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Impresión F054"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 7500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        FALTACPU% = 0
        'CARGO PRIMERO LOS DATOS DEL CPU
        SQLX$ = "SELECT CODIGO,DESCRIPCION,IDENTIFICACION,DESV_TRANS,DATOSCPU_ID,COD_INTE FROM DATOSCPU WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & XVALO(Me.TXTCI)
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!Codigo)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!IDENTIFICACION)
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!DATOSCPU_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!IDENTIFICACION)) = "" Then
               FALTACPU% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If
            
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        If k& < 3 Then k& = 2
        FALTALOTE% = 0
        CONDI$ = "COD_EXTE = '" & Label44 & "' AND  IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
        OF_NIVEL_SUPERIOR$ = VALOBADA("TRAZAXSERIE", "OF_ARRASTRA", CONDI$, "NOMESS")
        'TRAE TODOS LOS COMPONENTES CRITICOS CON EXCEPCION DE LOS DEL PRIMER NIVEL DE LA OF SELECCIONADA
        'SELECT * FROM TRAZAXSERIE WITH(NOLOCK)  WHERE NUMEROSERIE = '185-1473'  AND IDSUBOR <> 'OF-002221-1'
        SQLX$ = "SELECT COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION,DESV_TRANS FROM TRAZAXSERIE WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
        SQLX$ = SQLX$ + " AND  COD_EXTE <> '" & Label44 & "' AND  IDSUBOR <> '" & OF_NIVEL_SUPERIOR$ & "'"
        SQLX$ = SQLX$ + " ORDER BY  FECHALTA DESC"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            LOTEX$ = SAFETEXT$(!idenlote)
            LOTEX$ = REPLACAR(LOTEX$, "BR", "PR")
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = LOTEX$
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!TRAZAXSERIE_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!idenlote)) = "" Then
               FALTALOTE% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        '
        'AHORA BUSCO EL NSERIE Y LA OF DEL EMBALAJE ASOCIADO EL EJ. DB0001+ (COMO SE MUESTRA EN LA GRILLA DEL MEDIO)
        SQLX$ = "SELECT FechAlta, IDSUBOR,COD_EXTE,DESCRIPCION,STATUORF,COD_INTE,TRAZAXSERIE_ID,NUMEROSERIE"
        SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK)  Where TRAZABLE = -1 "
        SQLX$ = SQLX$ + "And NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR <> '" & NORFA$ & "'"
        SQLX$ = SQLX$ + "ORDER BY TRAZABLE DESC, STATUORF ,TRAZAXSERIE_ID"
        Set rs = READSET(SQLX$)
        With rs
          NorfaEmbalaje$ = SAFETEXT$(!IDSUBOR)
          NSERIEEMBALAJE$ = SAFETEXT$(!NumeroSerie)
        End With
        rs.Close
        Set rs = Nothing
        If TRIM$(NorfaEmbalaje$) = "" Or TRIM$(NSERIEEMBALAJE$) = "" Then
           Call COMUNI("No se pudo encontrar  el Código del Embalaje") 'si entra aca en el frm deberia estar vacio la grilla del medio
           Exit Sub
        End If
        
        'AHORA LOS ACCESORIOS (CONTENIDO DE EMBALAJE)
        SQLX$ = "SELECT NumeroSerie,TRAZAXSERIE_ID,IDENLOTE,COD_EXTE,CANTI,DESCRIPCION,COD_INTE,DESV_TRANS"
        SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE TRAZABLE = 1 AND NUMEROSERIE = '" & NSERIEEMBALAJE$ & "' AND IDSUBOR = '" & NorfaEmbalaje$ & "'"

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            LOTEX$ = SAFETEXT$(!idenlote)
            LOTEX$ = REPLACAR(LOTEX$, "BR", "PR")
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = LOTEX$
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!TRAZAXSERIE_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!idenlote)) = "" Then
               FALTALOTE% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        On Error Resume Next
        'AGREGO EL CODIGO EMBALAJE EL QUE ESTA EN LA GRILLA DEL MEDIO CON EL SIMBOLO +
        If Me.MSFESTADOVINCULADAS.Rows > 1 Then
            NorfaEmbalaje$ = Me.MSFESTADOVINCULADAS.TextMatrix(1, 2)
            CodExtEmbalaje$ = Me.MSFESTADOVINCULADAS.TextMatrix(1, 3)
            DescCodEmbalaje$ = Me.MSFESTADOVINCULADAS.TextMatrix(1, 4)
        End If
        On Error GoTo 0
        
        
        k& = k& + 1
        VENTABNEW.MSF.Rows = k& + 1
        VENTABNEW.MSF.TextMatrix(k&, 1) = CodExtEmbalaje$
        VENTABNEW.MSF.TextMatrix(k&, 2) = DescCodEmbalaje$
        VENTABNEW.MSF.TextMatrix(k&, 3) = NorfaEmbalaje$
        NroRegistroCodEmbalaje& = k&




30      VENTABNEW.Show 1 ' NO LO MUESTRO (ABAJO ESTA EL CODIGO SI SE MUESTRA PARA GRABAR MODIFICACIONES.
        If VENTABNEW.APROBADO < 1 Then Exit Sub
        
        If VENTABNEW.APROBADO > 0 And (FALTALOTE% + FALTACPU%) > 0 Then
             Call COMUNI("Imposible Continuar Faltan Datos de Lotes o del CPU ")
             GoTo 30
        End If
        Dim ObjetoExcel As Excel.Application
        Dim Libro As Excel.Workbook
        Dim Hoja As Excel.Worksheet
        
        PDES$ = Control$(IDENTER$, "PRTQEYM3")
        If PDES$ = "" Then PDES$ = Control$("*", "PRTQEYM3")
        RUTAEXCEL$ = Control("TRAZABLE", "EXARMA15")
        
    
        Set ObjetoExcel = CreateObject("Excel.Application")
        Set Libro = ObjetoExcel.Workbooks.Open(RUTAEXCEL$)
        With ObjetoExcel
            '
            'LIMPIO PRIMERO LOS CAMOS EN EL EXCEL
            For R& = 15 To 45
               If R& <> 28 And R& <> 29 And R& <> 30 Then ' en estos 3 registros no deberia imprmir
                Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 5) = "N/A"
                Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 7) = " "
               End If
             Next R&
            PREPARADOPORX$ = VENTABNEW.MSF.TextMatrix(1, 3)
            
            'ESCRIBO LOS DATOS EN OTROS CASILLEROS QUE ESTAN REFERENCIADOS
            For i& = 1 To VENTABNEW.MSF.Rows - 1
                CODI$ = UCase$(TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)))
                IDENTIF$ = VENTABNEW.MSF.TextMatrix(i&, 3)
                DV$ = CEROIGUALVACIO(VENTABNEW.MSF.TextMatrix(i&, 4))
                For R& = 15 To 45
                   CODIEXCEL$ = UCase$(TRIM$(Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 2)))
                   If CODIEXCEL$ = "AC0040" Then
                     A = 1
                   End If
                   If CODIEXCEL$ = CODI$ And TRIM$(Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 5)) = "N/A" Then
                      Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 5) = IDENTIF$
                      Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 7) = CEROIGUALVACIO(DV$)
                      Exit For
                   End If
                   SV$ = TRIM$(Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(R&, 5))
                Next R&
            Next i&
             '
             'ESTOS DATOS SON DE ENCABEZADO LOS ESCRIBO DIRECTAMENTE EN LA CELDA CORRESPONDIENTE
             For i& = 1 To 7
              Select Case i&
                Case 1: Dato$ = Label44: Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(3, 10) = Dato$
                Case 2: Dato$ = LBLREVISION: Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(4, 10) = Dato$
                Case 3: Dato$ = Label2(4): Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(5, 10) = Dato$
                Case 4:
                'valido por que aca va la fecha
                Dato$ = Label2(6)
'                If IsDate(DATO$) And ((Mid$(DATO$, 3, 1) = "-" And Mid$(DATO$, 6, 1) = "-") Or (Mid$(DATO$, 3, 1) = "/" And Mid$(DATO$, 6, 1) = "/")) Then
'                    Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(7, 6).NumberFormat = "d/m/yyyy;@"
'                End If
                Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(6, 10) = Dato$
                
                Case 5: Dato$ = Label13: Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(7, 10) = Dato$
                Case 6: Dato$ = DESVIOTRANSITORIO$: Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(8, 10) = Dato$
                Case 7: Dato$ = PREPARADOPORX$: Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(9, 10) = Dato$
              End Select
             Next i&
             '
             Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(30, 2) = CodExtEmbalaje$
             Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(30, 3) = DescCodEmbalaje$
             Libro.Worksheets("Registro Armado C-15 Rev.B").Cells(30, 5) = NorfaEmbalaje$
             
             'MUESTRO E IMPRIMO
             ObjetoExcel.ActiveWorkbook.Save
             Set Hoja = Libro.Sheets("Registro Armado C-15 Rev.B")
             Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=PDES$
             '
        End With
        
    
        'CIERRA HOJA, LIBRO Y EXCEL
99      On Error Resume Next
        Set Hoja = Nothing
        Libro.Close: Set Libro = Nothing
        ObjetoExcel.Close: Set ObjetoExcel = Nothing
        On Error GoTo 0

End Sub


Private Sub mnuTablasAuxiliares_Click()
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

Private Sub mnuVincularOrdAnt_Click()
    ENGASERIE_ANT.Show 1
End Sub

Private Sub modDesvioTransitorio_Click()
    DESVTRANS.Show 1
    
End Sub

Private Sub MSFESTADOVINCULADAS_DblClick()
     REG& = MSFESTADOVINCULADAS.MouseRow
     col% = MSFESTADOVINCULADAS.MouseCol
     If MSFESTADOVINCULADAS.Rows <= 1 Then Exit Sub
     
     OF$ = MSFESTADOVINCULADAS.TextMatrix(REG&, 2)
     NSERIE_SELEC$ = MSFESTADOVINCULADAS.TextMatrix(REG&, 5)
     
     Label13 = OF$ 'ACTUALIZA LA LISTA SUPERIOR
     
     Call T_Espera(1 / 4)
     For i& = 1 To MSFSegunOf.Rows - 1
        NS$ = MSFSegunOf.TextMatrix(i&, 3)
        If NS$ = NSERIE_SELEC$ Then
           MSFSegunOf.row = i&
           MSFSegunOf.TopRow = i& 'SE POSICIONA EN LA FILA PARA QUE QUEDE VISIBLE
           Call SELECCIONARFILA(MSFSegunOf, i&)
           Exit For
        End If
      Next i&
     
     Call PINTATEXT(Text1)
End Sub

Private Sub MSFRESERVA_Click()
    Call PINTATEXT(Text1)
End Sub

Private Sub MSFRESERVA_DblClick()
    i& = MSFRESERVA.MouseRow
    J& = MSFRESERVA.MouseCol
    If i& = 0 Or i& > MSFRESERVA.Rows Then Exit Sub
    '
    On Error Resume Next
    MSFRESERVA.row = i&
    MSFRESERVA.col = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    
    
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;N.Serie/A16;Lote/A16;ID/A0;CI/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.CampEditables = "5"
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 0
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Editar Lote"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 6500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        Call CopiarMsfAOtroMsf(MSFRESERVA, VENTABNEW.MSF, 1)
        '****************************************************************
        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
        
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% > 0 Then
           Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSFRESERVA, 1)
    '       MSF.COL = 1
    '       MSF.Sort = 3
        End If
        Set VENTANA = Nothing
        Unload VENTABNEW
        
        Text1.SetFocus
End Sub

Private Sub MSFRESERVA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
     REG& = MSFRESERVA.MouseRow
     col% = MSFRESERVA.MouseCol
     If MSFRESERVA.Rows <= 1 Then Exit Sub
     
     ID = TRIM$(MSFRESERVA.TextMatrix(REG&, 6))
     CODICONJ% = XVALO(MSFRESERVA.TextMatrix(REG&, 7))
     Codigo$ = TRIM$(MSFRESERVA.TextMatrix(REG&, 1))
     LoteDelComponente$ = TRIM$(MSFRESERVA.TextMatrix(REG&, 5)) 'es el lote asignado al codigo seleccionado en la grilla inferior
     
     
     OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Confirmar Borrar el Accesorio ( Código: " & Codigo$ & ")\Agregar Lote En Estructura del Código: " & Codigo$ & "")
     
     filaSeleccionada = XVALO(Text3)
     If filaSeleccionada > 0 Then
        If OPX% = 2 Then
           CONDI$ = "TRAZAXSERIE_ID = " & ID
           Call BORRAREGISTROS("TRAZAXSERIE", CONDI$, REGAF&, "NOMESS")
           If REGAF& > 0 Then
              Call COMUNI("Transacción Realizada")
              filaSeleccionada = XVALO(Text3)
              If filaSeleccionada > 0 Then
                MSFSegunOf.row = XVALO(Text3)
                Call MSFSegunOf_Click
              End If
           End If
        ElseIf OPX% = 3 Then
            If LoteDelComponente$ <> "" Then
              Call AgregarLoteAEstructura(CODICONJ%, LoteDelComponente$)
            Else
              Call COMUNI("Debe Asignar el Lote Al Código Seleccionado ( Código: " & Codigo$ & ")")
            End If
        End If
     
     End If
   End If
   
     
End Sub

Sub AgregarLoteAEstructura(CODICONJ%, LoteDelComponente$)
'ESTO SE AGREGO PARA PODER AGREGAR UN LOTE A UN COMPONENTE QUE  PERTENEZCA AL ITEM SELECCIONADO
'ESTA ACCION SOLO REGISTRA EN LA TABLA LOTEEYM EL LOTE A UN COMPONENTE EN LA TABLA SE GUARDA
'EL LOTE ASOCIADO AL COMPONENTE HIJO, EL COMPONENTE PADRE, LA OF, EL  NSERIE,Y LA FECHA
'SOLICITADO POR INES DONDE HABIA ITEMS QUE NO ERAN TRAZABLES Y SE LE HA EXIGIDO IGUAL QUE SE REGISTRE SU LOTE.

   If CODICONJ% < 1 Then Exit Sub
   NORFA$ = TRIM$(Label13)
   NSERIE$ = TRIM$(Label2(4))
   'MUESTRA LA FORMULA DEL ITEM SELECCIONADO
    TituloLista$ = "Componentes Trazobles del Código: " & CODEXT$(CODICONJ%)
    Campos$ = "CI/A0; Código/A20; Descripción/A64/; Uso Unit./F12.1"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    CONDI$ = "CODICONJ = " & CODICONJ%
    CONDI$ = CONDI$ + " ORDER BY CODIELEM"
    Set rs1 = obj.RS_GENERAL("FORMULAS INNER JOIN MASTER M WITH(NOLOCK) ON M.CODINT = CODIELEM AND M.TRAZABLE=1", " CODIELEM, CODXELEM, m.Descripcion, UsoBruto", CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    '
    '
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    'PRESENTA OBJETO PARA PODER ASIGNARLE UN LOTE AL COMPONENTE QUE SE SELECCIONO
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
       CIELEM% = XVALO(RESP_ZELE_VECT(1))
       USOUNIT = XVALO(RESP_ZELE_VECT(4))
    Else
       GoTo 99
    End If
    CONDI$ = "NUMEROSERIE= '" & NSERIE$ & "'"
    CONDI$ = CONDI$ + " AND CODICONJ = " & CODICONJ%
    CONDI$ = CONDI$ + " AND CODIELEM = " & CIELEM%
    CONDI$ = CONDI$ + " AND IDSUBOR = '" & NORFA$ & "'"

   SELCODLOT07.Label28 = VALOBADA("LOTEEYM", "IDENLOTE", CONDI$, "NOMESS")
   LOTE$ = LOTSELEC$(CIELEM%)
   If LOTE$ = "" Then GoTo 99
   VDEF$ = String$(64, 0)
   DEP% = DEPOLOT%(CIELEM%, LOTE$)
   Call REPLA(VDEF$, MKI$(CIELEM%), 1, 2)
   Call REPLA(VDEF$, LOTE$, 3, 12)
   Call REPLA(VDEF$, MKS$(USOUNIT), 15, 4)
   Call REPLA(VDEF$, MKI$(0), 20, 1)
   '

20 OBX$ = OBJPLA$("COLODECA", VDEF$)

   If OBX$ = "" Then GoTo 99
   
   If CIELEM% < 1 And TRIM(LOTE$) = "" Then
       GoTo 99
    End If
   '
   '
   DEPOCONSUM% = Asc(Mid$(OBX$, 20, 1))
   CANTI = CVS(Mid$(OBX$, 15, 4))
   '

   '
   If XVALO(CANTI) < 0.05 Then
      Call COMUNI("Falta Ingresar Cantidad")
      VDEF$ = OBX$
      GoTo 20
   End If
   '
   If DEPOCONSUM% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%)) = "" Then
     Call MENSERR(24, "Depósito de Consumo No Válido")
     VDEF$ = OBX$
     GoTo 20
   End If
   '
   If TRIM$(LOTE$) <> "" Then
    If DEPOCONSUM% <> DEP% Then
       Call COMUNI("Depósito Seleccionado Difiere de Ubicación de Lote")
       Call REPLA(OBX$, MKI$(DEP%), 20, 1)
       VDEF$ = OBX$
       GoTo 20
    End If
   End If
   '
   '
   If TRIM$(LOTE$) <> "" Then
     CONDI$ = "Cod_Inte =" & Ci% & " AND IdenLote ='" & LOTE$ & "'"
     CANTX = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
     If CANTX < XVALO(CANTI) Then
       Call COMUNI("Cantidad Solicitada Supera Saldo de Lote Seleccionado")
       GoTo 20
     End If
   End If
   '
   Call GuardarLoteAlComponenteEnTablaLOTEEYM(HOY(HO$), CODICONJ%, CIELEM%, LOTE$, NSERIE$, NORFA$, LoteDelComponente$, CANTI)
   'GUARDA EN UNA TABLA (LOTEEYM) DONDE REGISTRA EL N.OF, EL N.SERIE,EL CI PADRE, EL CI COMPONENTE, Y EL LOTE ASIGNADO.
   '
   Screen.MousePointer = 1
99
End Sub

 Sub MSFSegunOf_Click()
     REG& = MSFSegunOf.MouseRow
     COLUM% = MSFSegunOf.MouseCol
     Text3 = REG&
     If REG& = 0 Or REG& > MSFSegunOf.Rows - 1 Then Exit Sub
     
     NSERIE$ = TRIM$(MSFSegunOf.TextMatrix(REG&, 3))
     '
     'CARGA DE RESERVAS
     '============================
     NORFA$ = Label13
     If NORFA$ = "" Then Exit Sub
     i& = 0
     CONDI$ = " IdSubOr = '" & NORFA$ & "' AND TRAZABLE = 1 AND NUMEROSERIE = '" & NSERIE$ & "'"
     SQLX$ = "SELECT NumeroSerie,TRAZAXSERIE_ID,IDENLOTE,COD_EXTE,CANTI,DESCRIPCION,COD_INTE,DESV_TRANS "
     SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) WHERE " & CONDI$
     Campos$ = "Código/A12; Descrip/A20; Cant.Res./F8.1;N.Serie/A16;Lote/A16;ID/A0;CI/A0;N°D.Trans./f11"
     Call CARGA_ENCABEZADO(MSFRESERVA, Campos$, Me)
     Set rs = READSET(SQLX$)
     With rs
        Do While Not .EOF
           i& = i& + 1
           Me.MSFRESERVA.Rows = i& + 1
           Me.MSFRESERVA.TextMatrix(i&, 1) = SAFETEXT$(!COD_EXTE)
           Me.MSFRESERVA.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
           Me.MSFRESERVA.TextMatrix(i&, 3) = FORMATNUM$(XVALO(!CANTI), "F8.1")
           Me.MSFRESERVA.TextMatrix(i&, 4) = SAFETEXT$(!NumeroSerie)
           LOTE$ = SAFETEXT$(!idenlote)
           Me.MSFRESERVA.TextMatrix(i&, 5) = LOTE$
           Me.MSFRESERVA.TextMatrix(i&, 6) = XVALO(!TRAZAXSERIE_ID)
           Me.MSFRESERVA.TextMatrix(i&, 7) = SAFETEXT$(!COD_INTE)
           DVT = DESVIOTRANS(XVALO(SAFETEXT$(!COD_INTE)), LOTE$, SAFETEXT$(!NumeroSerie), REFERENCIA$)
           If DVT = 0 Then
              DVT = DESVIOTRANS(XVALO(SAFETEXT$(!COD_INTE)), LOTE$, "", REFERENCIA$)
           End If
           MSFRESERVA.TextMatrix(i&, 8) = DVT

           .MoveNext
        Loop
     End With
     rs.Close: Set rs = Nothing
     Label2(0).Caption = "Reservas Para N° Serie: "
     Label2(4) = NSERIE$
     Label2(6) = MSFSegunOf.TextMatrix(REG&, 1)
     '
     Call PINTATEXT(Text1)
     '
        'EN ESTE SE VEN TODAS LAS OF VINCULADAS Y SI LAS MISMAS ESTAN CERRADAS
     i& = 0
     SQLX$ = "SELECT FechAlta, IDSUBOR,COD_EXTE,DESCRIPCION,STATUORF,COD_INTE,TRAZAXSERIE_ID,NUMEROSERIE"
     SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where(TRAZABLE = 2 Or TRAZABLE = -1) And NUMEROSERIE = '" & NSERIE$ & "' AND IDSUBOR <> '" & NORFA$ & "' "
     SQLX$ = SQLX$ + " ORDER BY TRAZABLE DESC, STATUORF ,TRAZAXSERIE_ID"
      Set rs = READSET(SQLX$)
      With rs
         Do While Not .EOF
            i& = i& + 1
            Me.MSFESTADOVINCULADAS.Rows = i& + 1
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 1) = SAFETEXT$(!FECHALTA)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 2) = SAFETEXT$(!IDSUBOR)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 3) = SAFETEXT$(!COD_EXTE)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 4) = SAFETEXT$(!Descripcion)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 5) = SAFETEXT$(!NumeroSerie)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 6) = SAFETEXT$(!TRAZAXSERIE_ID)
            Me.MSFESTADOVINCULADAS.TextMatrix(i&, 7) = XVALO(!StatuOrf)
'            If XVALO(!StatuOrf) <> 3 Then
'               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFESTADOVINCULADAS, &HFF0000, i&)
'            Else
'               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFESTADOVINCULADAS, vbRed, i&)
'            End If
    
            .MoveNext
         Loop
       End With
       rs.Close: Set rs = Nothing
       Call PINTATEXT(Text1)
       
End Sub

Sub AGREGARLECTURA(CI1%, LOTE$)

   With MSFRESERVA
     SEASIGNOLOTE% = 0
     For REG& = 1 To .Rows - 1
        CI2% = .TextMatrix(REG&, 7)
        If CI1% = CI2% Then
          LOTEENGRILLA$ = TRIM$(.TextMatrix(REG&, 5))
          If LOTEENGRILLA$ = "" Then
             .TextMatrix(REG&, 5) = LOTE$
             SEASIGNOLOTE% = 1
             Exit For
          Else
             REGISTRO_COINCIDENCIA% = REG&
          End If
        End If
      Next REG&
   End With
  
  
'  If SEASIGNOLOTE% < 1 Then
'     Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSFRESERVA, vbYellow, REGISTRO_COINCIDENCIA%)
'     Call COMUNI("No se Asignó Lote\Ya Posee Asignados")
'     Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSFRESERVA, vbWhite, REGISTRO_COINCIDENCIA%)
'   End If
 
End Sub

Private Sub MSFSegunOf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BORRAR LOS DATOS DE CPU PARA LOS CASOS QUE SE GUARDARON CON EL FORMATO VIEJO. QUE ERAN 3 REGISTROS Y AHORA SON 6 PARA LOS DB-nnnnnn
 '
    If Button = 2 Then
     REG& = MSFSegunOf.MouseRow
     col% = MSFSegunOf.MouseCol
     If MSFSegunOf.Rows <= 1 Then Exit Sub
     
     OF$ = TRIM$(MSFSegunOf.TextMatrix(REG&, 2))
     NS$ = TRIM$(MSFSegunOf.TextMatrix(REG&, 3))
     CI1% = XVALO(TXTCI)
     OPX% = COMALTER%("Desea Borrar Los Datos de CPU Guardados (" & OF$ & "  /  " & NS$ & ")\\Cancelar\Confirmar Borrar")
     If OPX% > 1 Then
        If CI1% > 0 And OF$ <> "" And NS$ <> "" Then
            CONDI$ = "NUMEROSERIE = '" & NS$ & "'"
            CONDI$ = CONDI$ + " AND IDSUBOR = '" & OF$ & "'"
            CONDI$ = CONDI$ + " AND COD_INTE = " & CI1%
            Call BORRAREGISTROS("DATOSCPU", CONDI$, REGAF&, "NOMESS")
            Call COMUNI("Transacción Realizada Se ha Borrado " & REGAF& & " Registros (" & OF$ & "  /  " & NS$ & ")")
        End If
     End If
    
    End If
End Sub

Private Sub MSFSERIESVINC_Click()
   Call PINTATEXT(Text1)
End Sub

Private Sub MSFSERIESVINC_DblClick()
     REG& = MSFSERIESVINC.MouseRow
     col% = MSFSERIESVINC.MouseCol
     If MSFSERIESVINC.Rows <= 1 Then Exit Sub
     
     OF$ = MSFSERIESVINC.TextMatrix(REG&, 2)
     NS$ = MSFSERIESVINC.TextMatrix(REG&, 3)
     Label13 = OF$
     '
'     Call T_Espera(1 / 4)
'     For i& = 1 To MSFSegunOf.Rows - 1
'       If MSFSegunOf.TextMatrix(i&, 3) = NS$ Then
'          MSFSegunOf.Row = i&
'          Call SELECCIONARFILA(MSFSegunOf, i&)
'          Call MSFSegunOf_Click
'          Exit For
'       End If
'     Next i&
     
End Sub


'
'
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        sLecturaCodigoBarras$ = ReemplazaXGuionCaracterDiferenteEYM(TRIM$(Me.Text1))
        'sLecturaCodigoBarras$ = REPLACAR(sLecturaCodigoBarras$, "'", "-") 'REEMPLAZO POR QUE LA LECTORA REEMPLAZA EL '-' POR "'"
        
        Call TEXTOLOTES(sLecturaCodigoBarras$, " ")
'            On Error Resume Next
        For i& = 1 To UBound(CADENATEX$) 'RECORRO EL VECTOR QUE TIENE CUALES SON LAS CELDAS EDITABLES
            Select Case i&
            Case 1
             CI1% = CODINT%(CADENATEX$(i&)) 'SI ES CODIGO INTERNO
             If CI1% < 1 Then
                Me.TXTCODIGO.TEXT = CODEXT$(XVALO(CADENATEX$(i&))) 'SI ES CODIGO INTERNO
                CI1% = XVALO(CADENATEX$(i&))
             Else
                Me.TXTCODIGO.TEXT = CODEXT$(CI1%)
             End If
             
            Case 2
             LEIDO$ = TRIM$(CADENATEX$(i&))
             If LEIDO$ <> "Rev." Then LEIDO$ = "Rev." & LEIDO$ 'SI NO DICE 'Rev.' se lo agrega
             Me.TXTREVISION = LEIDO$
            Case 3
             LEIDO$ = CADENATEX$(i&)
             Me.TXTLOTE = FORMATOBR$(LEIDO$, CI1%) ' SI LE FAL
            End Select
        Next i&
'            On Error GoTo 0
        Call AGREGARLECTURA(CI1%, TXTLOTE)
        Text1 = ""
    End If
    '
    If KeyAscii >= 32 And KeyAscii <= 126 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       'agrego el caracter por que si no solo toma el foco
'       CARACTERX$ = Chr$(KeyAscii)
'       Text1.TEXT = Text1.TEXT & CARACTERX$
'       Text1.SelStart = Len(Text1)
    ElseIf KeyAscii = 8 Then
       On Error Resume Next
       Text1 = Left$(Text1, Len(Text1) - 1)
       On Error GoTo 0
    End If
   
'         Call PINTATEXT(Text1)
         
End Sub
Sub CIERRORFA_EYM(NORFA$, CANTI, NSERIE$)
    CI1% = CODIFAB%(NORFA$)
    FECHACIERRE% = HOY(HO$)
    
    Call ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FECHACIERRE%, 1, "Fabricación Propia")
    '
    CAN# = CANTI
    Call CARCOMPO(CODEXT$(CI1%), NORFA$, CAN#, 1, 1, 5)

    Call CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FECHACIERRE%, XVALO(Text1), NSERIE$)
    '
    'ACTUALIZA LA RESERVA
    FIN& = ULTREG&("!MATEROF")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CICOMPO% = CVI(Left$(ZZ$, 2))
      CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
      If TRIM$(UCase$(Unimed$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
      '
      FECHA% = FECHACIERRE%
      Call ACTURESERVA(CICOMPO%, NORFA$, CANCONSUM)
    Next U&
    '
    Call ANOTAENORFA_FACTURADO(CI1%, NORFA$, CANTI, FECHACIERRE%, "Cierre de O/Fabricación (Proc.Másivo)")
    Call ACTUALIZAORFA(CI1%, NORFA$, CANTI, FECHACIERRE%)
    
End Sub

Sub Cierre_Orden_Fabriacion_EYM(NORFA$, NSERIE$)
       CIFAB% = CODIFAB%(NORFA$)
       FECHACIERRE% = HOY(HO$)
       CODICONJ$ = CODEXT$(CIFAB%)
       ESEXCEPCION_SERIE% = ES_EXCEPCION_GENERAR_SERIE%(CODICONJ$)
       
       CONDI$ = "NUMEROSERIE = '" & NSERIE$ & "' AND (" 'VA ARMANDO DENTRO DEL FOR EL RESTO DE LA CONSULTA
       DOPRI$ = NORFA$
       DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
       REFE$ = "CONSUMO " + CODICONJ$
       CMOV$ = "CF"

       ENTRO% = 0
       Dim MSF As msFlexGrid
       Set MSF = VENTABNEW.MSF

       For i& = 1 To MSF.Rows - 1
          CI1% = XVALO(MSF.TextMatrix(i&, 1))
          If CI1% > 0 Then
            COD$ = MSF.TextMatrix(i&, 2)
            Descrip$ = MSF.TextMatrix(i&, 3)
            CANTI = XVALO(MSF.TextMatrix(i&, 4))
            LOTE$ = ""
            If ESTRAZABLE%(CI1%) >= 1 Then
               LOTE$ = MSF.TextMatrix(i&, 5)
               LOTE$ = LOTE_A_CONSUMIR_EYM(CI1%, LOTE$, CANTI, NSERIE$) 'FUNCION
            End If
            If LOTE$ <> "" Then
              If ENTRO% < 1 Then
                ENTRO% = 1
                CONDI$ = CONDI$ & "COD_INTE = " & CI1%
              Else
                CONDI$ = CONDI$ & " OR COD_INTE = " & CI1%
              End If
            End If
            NUORIT% = NUORIT% + 1
            Call MOVISTOK(FECHACIERRE%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, 1, CANTI * (-1))
          End If
          Call ACTURESERVA(CI1%, NORFA$, CANTI)
       Next i&
       CONDI$ = CONDI$ & ")"
       '
       
       Call ACTUREGISTRO("TRAZAXSERIE", "CONSUMIDO", CONDI$, "1", REGAF&)
       Call ACTUREGISTRO("TRAZAXSERIE", "StatuOrf", "IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'", "3", REGAF&)
       '
       NS$ = NSERIE$: If ESEXCEPCION_SERIE% > 0 Then NS$ = "" 'EN CASO DE SER EXCEPCION DE N.SERIE INGRESA COMO PRODUCTO TERMINADO EL LOTE OF
       Call ENRAPRODUCTO_TERMINADO(CIFAB%, NORFA$, 1, FECHACIERRE%, 1, "Fabricación Propia", NS$)
       '
       Call ANOTAENORFA_FACTURADO(CIFAB%, NORFA$, 1, FECHACIERRE%, "Cierre de O/Fabricación (" & NS$ & ")")
       Call ACTUALIZAORFA(CIFAB%, NORFA$, 1, FECHACIERRE%)
       Call ActualizaTanumseSiesProductoTerminado(NORFA$, NSERIE$)
       '
    
       '
End Sub
Function VALIDA_STOCK_COMPONENTES%(NORFA$, NSERIE$)
    '0 SI CEERRO LA VENTANTA DE COMPONENTES SIN APROBAR
    '1 SI HAY STOCK PARA TODOS LOS COMPONENTES
    '-1 SI FALTA STOCK PARA UNO O MAS COMPONENTES.
    
    'AQUI SE VALIDA EL STOCK DE LOS COMPONENTES, SI FALTA STOCK SE MUESTRA UN SELECHO INFORMANDO EL DETALLE
    VALIDA_STOCK_COMPONENTES% = 0
    Unload VENTABNEW
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = ""
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Lista de Componentes a Consumir (Todos Trazables o Nó)"
    Campos$ = "CI/A0;Código/A16;Descripción/A28;Cant./F8.1;Lote/A16;Tr./f3;N°Serie/A12"
    
    VENTANA.Campox = Campos$
    VENTANA.ColEcoCampo(2) = "CODIGO;DESCRIPCION;MASTER"
    VENTANA.Inicializar = 1
    '****************************************************************
    Call COMPLETAR_CONSUMO_OF_SERIE(VENTABNEW.MSF, NORFA$, NSERIE$, RET_COUNT_REG&)
    If RET_COUNT_REG& < 1 Then
       Call COMUNI("No se Encontraron Reservas Para esta Orden de Fabricación")
       Exit Function
    End If
30  VENTABNEW.Show 1
    
    Dim MSF As msFlexGrid
    Set MSF = VENTABNEW.MSF
    '
    Campos$ = "CI/A0;Código/A16;Descripción/A28;Cant./F8.1;Lote/A16;Tr./f3;Referencia/A24"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Items Sin Sock"
    FRMZELECHO.Width = 12000
    '
    J& = 0
    If VENTABNEW.APROBADO% > 0 Then
       VALIDA_STOCK_COMPONENTES% = 1
       For i& = 1 To MSF.Rows - 1
          CI1% = XVALO(MSF.TextMatrix(i&, 1))
          If CI1% > 0 Then
            COD$ = MSF.TextMatrix(i&, 2)
            Descrip$ = MSF.TextMatrix(i&, 3)
            CANTI = XVALO(MSF.TextMatrix(i&, 4))
            LOTE$ = ""
            If ESTRAZABLE%(CI1%) = 1 Then
               LOTE$ = MSF.TextMatrix(i&, 5)
               LOTE$ = LOTE_A_CONSUMIR_EYM$(CI1%, LOTE$, CANTI, NSERIE$) 'FUNCION
               If LOTE$ = "SIN STOCK" Then GoSub LLENAR_ZELECHO
            Else
               If STOCKACT(CI1%) < CANTI Then
                  GoSub LLENAR_ZELECHO
               End If
            End If
          End If
       Next i&
    Else
       VALIDA_STOCK_COMPONENTES% = 0
    End If
    If J& > 0 Then
      FRMZELECHO.Show 1
    End If
    Exit Function
    
LLENAR_ZELECHO:
                  VALIDA_STOCK_COMPONENTES% = -1
                  J& = J& + 1
                  FRMZELECHO.MSF2.Rows = J& + 1
                  FRMZELECHO.MSF2.TextMatrix(J&, 1) = CI1%
                  FRMZELECHO.MSF2.TextMatrix(J&, 2) = COD$
                  FRMZELECHO.MSF2.TextMatrix(J&, 3) = Descrip$
                  FRMZELECHO.MSF2.TextMatrix(J&, 4) = CANTI
                  FRMZELECHO.MSF2.TextMatrix(J&, 5) = MSF.TextMatrix(i&, 5)
                  FRMZELECHO.MSF2.TextMatrix(J&, 6) = MSF.TextMatrix(i&, 6)
                  FRMZELECHO.MSF2.TextMatrix(J&, 7) = "SIN STOCK"
Return
End Function

Sub COMPLETAR_CONSUMO_OF_SERIE(MSF As msFlexGrid, NORFA$, NSERIE$, COUNT_REG&)
    CANPRO& = CantiProyectadaOF(NORFA$)
    REG& = 0

    CAMPOS0$ = "CI/A0;Código/A16;Descripción/A28;Cant./F8.1;Lote/A16;Tr./f3;N°Serie/A12"
    Campos$ = "COD_INTE,COD_EXTE,DESCRIPCION,CANTI,IDENLOTE,TRAZABLE,NUMEROSERIE"
    Call CARGA_ENCABEZADO(VENTABNEW.MSF, CAMPOS0$, VENTABNEW)
    SQLX$ = "select " & Campos$ & " from trazaxserie WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE OF_ARRASTRA = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
    SQLX$ = SQLX$ + " AND (CONSUMIDO < 1  OR CONSUMIDO IS NULL) AND (TRAZABLE = -1 OR TRAZABLE = 2)"
    Set rs = READSET(SQLX$)
    With rs
'        XIX1% = XVALO(!COD_INTE)

        Do While Not .EOF

            REG& = REG& + 1
            MSF.Rows = REG& + 1
            MSF.TextMatrix(REG&, 1) = XVALO(!COD_INTE)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!COD_EXTE)
            MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descripcion)
            MSF.TextMatrix(REG&, 4) = "1.0"
            MSF.TextMatrix(REG&, 5) = SAFETEXT$(!NumeroSerie)
            MSF.TextMatrix(REG&, 6) = XVALO(!TRAZABLE)
            MSF.TextMatrix(REG&, 7) = SAFETEXT$(!NumeroSerie)
            .MoveNext
        Loop
    End With
    rs.Close: Set rs = Nothing
    
    SQLX$ = "select " & Campos$ & " from trazaxserie WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "' AND TRAZABLE = 1 AND (CONSUMIDO < 1  OR CONSUMIDO IS NULL)"
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
            
            REG& = REG& + 1
            MSF.Rows = REG& + 1
            MSF.TextMatrix(REG&, 1) = XVALO(!COD_INTE)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!COD_EXTE)
            MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descripcion)
            MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTI), "f8.1")
            MSF.TextMatrix(REG&, 5) = SAFETEXT$(!idenlote)
            MSF.TextMatrix(REG&, 6) = XVALO(!TRAZABLE)
            MSF.TextMatrix(REG&, 7) = SAFETEXT$(!NumeroSerie)
    
            .MoveNext
        Loop
    End With
    rs.Close: Set rs = Nothing
    
    SQLX$ = "select R.COD_INTE,R.COD_EXTE,R.Descrip,R.CANTRES from RESERVA R WITH(NOLOCK)"
    SQLX$ = SQLX$ + "  INNER JOIN MASTER M ON M.CODINT = R.COD_INTE WHERE R.IDSUBOR = '" & NORFA$ & "' AND (M.TRAZABLE = 0 OR M.TRAZABLE IS NULL)"
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
            REG& = REG& + 1
            MSF.Rows = REG& + 1
            SERIEX$ = ""
'            If XIX1% = XVALO(!COD_INTE) Then SERIEX$ = NSERIE$
            MSF.TextMatrix(REG&, 1) = XVALO(!COD_INTE)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!COD_EXTE)
            MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
            MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTRES) / CANPRO&, "f8.1")
            MSF.TextMatrix(REG&, 5) = 0
            MSF.TextMatrix(REG&, 6) = ""
            MSF.TextMatrix(REG&, 7) = SERIEX$
            .MoveNext
        Loop
    End With
    rs.Close: Set rs = Nothing
    COUNT_REG& = REG&

End Sub

Private Sub TXTCI_Change()
  LBLREVISION.Caption = REVICOD$(XVALO(Me.TXTCI))
End Sub
Sub PRESENTA_CONTROLES()
    Campos$ = "CONTROL/A32;FORMULARIO/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE CONTROLES"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_CONTROLES_EYM()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1
    LECTURA1X$ = TRIM$(RESP_ZELE_VECT(1))
    If LECTURA1X$ = "" Then Exit Sub
    FORMU = RESP_ZELE_VECT(2)
    Select Case TRIM$(FORMU)
      Case "REGIDATO"
        RegDatoEym.Show 1
      Case "CONTFINEYM"
        RegCFinEym.Command3.Enabled = False
        RegCFinEym.Show 1
      Case "RegCFinC15Eym"
        RegCFinC15Eym.Show 1
      Case Else
         Call COMUNI("No se Encontró El Control Seleccionado")
         Exit Sub
    End Select
    
      

End Sub



