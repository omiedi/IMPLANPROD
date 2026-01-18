VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ImpoRegAr 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formulario Registro Control Final"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14985
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
   ScaleHeight     =   9945
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   4080
      TabIndex        =   11
      Top             =   2040
      Width           =   6735
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Registro Seleccionado"
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
         Index           =   11
         Left            =   2400
         TabIndex        =   35
         Top             =   2280
         Width           =   2085
      End
      Begin VB.Label LBLREGISTROSELECCIONADO 
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
         Left            =   4560
         TabIndex        =   29
         Top             =   2280
         Width           =   1545
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
         Index           =   7
         Left            =   4560
         TabIndex        =   28
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FIRMA"
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
         Index           =   8
         Left            =   3360
         TabIndex        =   27
         Top             =   1365
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
         Index           =   6
         Left            =   4560
         TabIndex        =   26
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARO"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   1005
         Width           =   1485
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
         Index           =   5
         Left            =   4560
         TabIndex        =   24
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N° DE D.T."
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
         Index           =   6
         Left            =   3360
         TabIndex        =   23
         Top             =   645
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
         Index           =   4
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OF N°"
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
         Index           =   5
         Left            =   3360
         TabIndex        =   21
         Top             =   285
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
         Index           =   3
         Left            =   1440
         TabIndex        =   20
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
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
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1365
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
         Index           =   2
         Left            =   1440
         TabIndex        =   18
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE N°"
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
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1005
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
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REVISION"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   645
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
         Index           =   0
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
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
         Left            =   240
         TabIndex        =   12
         Top             =   285
         Width           =   1125
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Bindings        =   "ImpoRegAr..frx":0000
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8705
      _Version        =   393216
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
   Begin MSComDlg.CommonDialog cmd 
      Left            =   13800
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9960
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   1320
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.CommandButton Command5 
         Caption         =   "DETENER"
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
         Left            =   7440
         TabIndex        =   40
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar (4)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10320
         TabIndex        =   37
         ToolTipText     =   $"ImpoRegAr..frx":0014
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TXTCOL 
         Appearance      =   0  'Flat
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
         Left            =   13680
         TabIndex        =   34
         Text            =   " "
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox TXTREG 
         Appearance      =   0  'Flat
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
         Left            =   13680
         TabIndex        =   33
         Text            =   " "
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mostrar Datos en Pantalla (3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9120
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cargar Grilla de Registros (2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7380
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CommandButton cmdconn 
         Caption         =   "Conectar Con Excel"
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
         Left            =   7800
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Registro Identificador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   10
         Left            =   12360
         TabIndex        =   32
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Columna Identificador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   9
         Left            =   12360
         TabIndex        =   31
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Ruta Archivo Excel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Selección de Hoja de B.Datos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Bindings        =   "ImpoRegAr..frx":00A9
      Height          =   2775
      Left            =   10800
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4895
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Bindings        =   "ImpoRegAr..frx":00BD
      Height          =   2775
      Left            =   12960
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4895
      _Version        =   393216
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
   Begin MSFlexGridLib.MSFlexGrid MSF5 
      Bindings        =   "ImpoRegAr..frx":00D1
      Height          =   4575
      Left            =   12120
      TabIndex        =   38
      Top             =   5040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8070
      _Version        =   393216
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N.Existentes"
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
      Index           =   12
      Left            =   12120
      TabIndex        =   39
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   4080
      X2              =   10800
      Y1              =   3720
      Y2              =   3720
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
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   45
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   2745
      Left            =   0
      Picture         =   "ImpoRegAr..frx":00E5
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4020
   End
End
Attribute VB_Name = "ImpoRegAr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public objexcel As Excel.Application
'Public Libro As Excel.Workbook
'Public Hoja As Excel.Worksheet

Sub ImortacionRegDatoEym(RUTA_ARCH$, NombreHoja$, MSF As MSFlexGrid, IMPRESORA$, Columnas_a_Leer$, A_Partir_Fila%, En_Columna%, Optional datofijo$, Optional SIVACIO%, Optional NOIMPRIMIR%, Optional ARCHIVOIMAGEN$, Optional TOPX, Optional LEFTX, Optional HEIGHTX, Optional WIDTHX, Optional grabar%)
    'Columnas_a_Leer$ = NUMEROS DE COLUMNAS QUE VA A LEER DEL MSF SEPARADOS PRO PUNTO Y COMA "1;3;5"
    'A_Partir_Fila% = A PARTIR DE QUE FILA DEL EXCEL EMPIEZA A ESCRIBIR LOS DATOS
    'En_Columna%= EN QUE COLUMNA DEL EXCEL VA A ESCRIBIR LOS DATOS
    'SI Columnas_a_Leer$ ES MAS DE UNA LAS ESCRIBE EN EL EXCEL EN LA COLUMNA DE AL LADO
    
    'EJEMPLO SI Columnas_a_Leer$ = "1;4" Y En_Columna% = 5
    'VA A ESCRIBIR LA INFORMACION EN EL EXCEL _
    EN LA COLUMNA 5 Y 6
    'SI SE PASA UN VALOR COMO DATO FIJO Y DENTRO DE LAS COLUMNAS A LEER INCLUYE EL 0 CUANDO LEA ESTA ASIGNA EL VALORFIJO
    '
    'ESTA RUTINA SE ADAPTA PARA QUE ESCRIBA EN EL EXCEL DE REGISTRO DE DATOS ENVIANDO DE EYM
    'EN UNA POSICIION CONFIGURADA EL NRO DE ORDEN DE REGISTRO (EN EL EXCEL CON ESTE CAMBIO ACTUALIZA LOS DATOS)
    'LUEGO SE LEEN TODOS LOS DATOS Y SE IMPORTAN A SQL EN TRES TABLAS
    'EncaRegArma,CpuRegArma, DetaRegArma CREADAS PARA ESTE FIN RELACIONADAS POR EL NRO DE SERIE
    ImpoRegAr.msf2.Rows = 1
    ImpoRegAr.MSF4.Rows = 1
    '
    ImpoRegAr.Command4.Enabled = True 'HABILITO EL BOTON GRABAR AL COMIENZO SI HAY CODIGOS INEXISTENTES SE DEAHABILITA
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    Call TEXTOLOTES(Columnas_a_Leer$, ";")
    
    Set ObjetoExcel = CreateObject("Excel.Application")
    With ObjetoExcel
        ' ARCHIVO EXCEL / ABRE EL LIBRO
        Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)

        If TRIM$(ARCHIVOIMAGEN$) <> "" Then
          'IMPRIME UNA IMAGEN EN LA HOJA
          'objHoja.Shapes.AddPicture <Name file>, False, True, 445, 200, 250, 150
          Libro.Worksheets(NombreHoja$).Shapes.AddPicture "C:\FLEXOFT\QR_ETIQPTERM.JPG", False, True, LEFTX, TOPX, WIDTHX, HEIGHTX
        End If
         k& = 0
         For i& = 1 To MSF.Rows - 1
         Openforms = DoEvents
         If Command5.Caption = "CONTINUAR" Then GoTo 98 'boton de detencion
         
            For j& = 1 To UBound(CADENATEX$)
               Col_a_Leer% = CADENATEX$(j&)
               DATO$ = TRIM$(MSF.TextMatrix(i&, Col_a_Leer%))
               If DATO$ = "" And SIVACIO% < 1 Then
                  DATO$ = "-----"
               End If
               If Left$(DATO$, 3) = "BR-" Then DATO$ = REPLACAR(DATO$, "BR-", "PR-")
               
               If Col_a_Leer% = 0 Then
                  If datofijo$ <> "" And STRINGFILA_STRING$(MSF, i&) <> "" Then
                     DATO$ = datofijo$
                  Else
                      DATO$ = "--"
                  End If
               End If
               Call SELECCIONARFILA(ImpoRegAr.MSF3, i&)
               Call UbicarFocoTop(ImpoRegAr.MSF3, i&)
               Libro.Worksheets(NombreHoja).Cells(6, 31) = DATO$
               Libro.Worksheets(NombreHoja).Cells(7, 31) = "1" 'LE AGREGO UNO PARA QUE CAMBIE DE CELDA
               'MsgBox Libro.Worksheets(NombreHoja).Cells(6, 20)
               GoSub ASIGNARVALORENCABEZADO
               GoSub ASIGNARVALORCPU
               GoSub ASIGNARDETALLE
               
               
            Next j&
            k& = k& + 1
         Next i&
         
98       ObjetoExcel.ActiveWorkbook.Save
'         Set Hoja = Libro.Sheets(NombreHoja$)
         If NOIMPRIMIR% > 0 Then GoTo 99
         Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
         'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada

    End With

    'CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0
    Exit Sub
    
ASIGNARVALORENCABEZADO:
    REG_ENCA_DESDE = 4 '
    REG_ENCA_HASTA = 8
    COL_ENCA = 20
    ImpoRegAr.LBLREGISTROSELECCIONADO = DATO$
    REG& = REG_ENCA_DESDE
    IndiceObj = 0
    Do While REG& <= REG_ENCA_HASTA
        If REG& = 0 Then
           If Abs(CODINT%(DATO$)) < 1 Then 'ABS POR SI ESTA DE BAJA LO DEJO PASAR IGUAL
             REGMSF5 = MSF5.Rows 'SI TIENE DATOS EN ESTA GRILLA NO VA A PERMITIR GRABAR
             MSF5.Rows = MSF5.Rows + 1
             MSF5.TextMatrix(REGMSF5, 1) = DATO$
             MSF5.TextMatrix(REGMSF5, 2) = "CODIGO INEXISTENTE"
             ImpoRegAr.Command4.Enabled = False
           End If
        End If
        DATO$ = SAFETEXT$(Libro.Worksheets(NombreHoja).Cells(REG&, COL_ENCA))
        ImpoRegAr.Label13(IndiceObj) = DATO$: ImpoRegAr.Label13(IndiceObj).Refresh
        IndiceObj = IndiceObj + 1
        REG& = REG& + 1
     Loop
     Return
     
     
ASIGNARVALORCPU:
    j& = 0
    
    REG_ENCA_DESDE = 5 '
    REG_ENCA_HASTA = 12
    COL_ENCA = 22
    REG& = REG_ENCA_DESDE
    Do While REG& <= REG_ENCA_HASTA
        DATO$ = Libro.Worksheets(NombreHoja).Cells(REG&, COL_ENCA)
        j& = j& + 1
        ImpoRegAr.msf2.Rows = j& + 1
        ImpoRegAr.msf2.TextMatrix(j&, 1) = DATO$
        REG& = REG& + 1
     Loop
     Return
     
ASIGNARDETALLE:
    j& = 0
    
    REG_ENCA_DESDE = 16 '
    REG_ENCA_HASTA = 41
    COL_COD = 2
    COL_DESCRIP = 6
    COL_IDEN = 20
    COL_NDET = 22
    
    REG& = REG_ENCA_DESDE
    Do While REG& <= REG_ENCA_HASTA
        Codigo$ = SAFETEXT$(Libro.Worksheets(NombreHoja).Cells(REG&, COL_COD))
        If Left$(Codigo$, 13) = "Formulario Nº" Then
           GoTo 80
        End If
        Descripcion$ = SAFETEXT$(Libro.Worksheets(NombreHoja).Cells(REG&, COL_DESCRIP))
        IDENTIFICACION$ = SAFETEXT$(Libro.Worksheets(NombreHoja).Cells(REG&, COL_IDEN))
        NDET$ = SAFETEXT$(Libro.Worksheets(NombreHoja).Cells(REG&, COL_NDET))
        j& = j& + 1
        ImpoRegAr.MSF4.Rows = j& + 1
        ImpoRegAr.MSF4.TextMatrix(j&, 1) = Codigo$
        ImpoRegAr.MSF4.TextMatrix(j&, 2) = Descripcion$
        ImpoRegAr.MSF4.TextMatrix(j&, 3) = IDENTIFICACION$
        ImpoRegAr.MSF4.TextMatrix(j&, 4) = NDET$
80      REG& = REG& + 1
     Loop
     If grabar% > 0 Then
        Call GRABARIMPORTACIONABD
     End If
     Return

End Sub





Sub cerrar_excel()

  On Error Resume Next
   ObjetoExcel.Application.Quit ' Cierra Excel
   Libro.Close (False)
   Set ObjetoExcel = Nothing
   Set Libro = Nothing
   Set Hoja = Nothing
   
   '
   Libro.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set Libro = Nothing     ') libera memoria
   Set Hoja = Nothing    ') utilizada por
   '
   On Error GoTo 0
End Sub
Private Sub cmdconn_Click()
      If TRIM$(Text1) = "" Then
          MsgBox "Falta Seleccionar Ruta Archivo Excel"
          Call PINTATEXT(Text1)
          Exit Sub
      End If

'      Call HAB_O_DESCOMMAND(Me, 1)
'      Dim obj As New LeerTxtExcel
'
'      Set DataGrid1.DataSource = obj.Leer_Excel(Text1, Combo1.TEXT, Combo3.TEXT)
      
'      Set obj = Nothing
'      On Error Resume Next
'      rs.Close
    
      Call cerrar_excel
'      DataGrid1.Row = 0
      Call CARGAR_MSF_DE_EXCEL_SOLOUNACOLUMNA(Text1, MSF3, Combo1.TEXT)
      Call HAB_O_DESCOMMAND(Me, 0)
End Sub
  Sub CARGAR_MSF_DE_EXCEL_SOLOUNACOLUMNA(Ruta$, MSF As MSFlexGrid, NOMHOJA$, Optional ColumnaALeerEnExcel%)
     'ESTA RUTINA LEE SOLO UNA COLUMNA DEL EXCEL Y LA ESCRIBE EN UN FLEX GRID
     On Error GoTo 10
     MSF.Rows = 1
     COL% = ColumnaALeerEnExcel%
'     On Error GoTo Errores
j& = 0
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    Set ObjetoExcel = CreateObject("Excel.Application")
    Set Libro = ObjetoExcel.Workbooks.Open(Ruta$)

     AA = Combo1.TEXT
     CANTCOL% = numero_de_columnas
     ObjetoExcel.Sheets(NOMHOJA$).Select
     With ObjetoExcel.Sheets(NOMHOJA$)
       UltFila = .Range("A65536").End(xlUp).Row
50     Do While i < UltFila
'Libro.Worksheets(NombreHoja).Cells(6, 20)
          i = i + 1
          If XVALO(.Cells(i, ColumnaALeerEnExcel%)) > 0 Then
             j& = j& + 1
             MSF.Rows = j& + 1
             AAA = .Cells(i, ColumnaALeerEnExcel%)
             MSF.TextMatrix(j&, ColumnaALeerEnExcel%) = .Cells(i, ColumnaALeerEnExcel%)   'Leemos los datos apartir de la celda 1 y metemos la informacion a un combo
           End If
       Loop
     End With
     
10  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0
    End Sub
Private Sub Combo1_Click()

   '
End Sub

Private Sub Command2_Click()
'    For i& = 1 To 6
'       MSF.Rows = i& + 1
'       MSF.TextMatrix(i&, 1) = i&
'    Next i&
    Call CARGAR_MSF_DE_EXCEL_SOLOUNACOLUMNA(Text1, MSF3, Combo1.TEXT, 1)

    Call COMUNI("TAREA FINALIZADA")
End Sub
Sub Param_Print_Excel(RUTA_ARCH$, NombreHoja$, MSF As MSFlexGrid, IMPRESORA$, Columnas_a_Leer$, A_Partir_Fila%, En_Columna%, Optional datofijo$, Optional SIVACIO%, Optional NOIMPRIMIR%, Optional ARCHIVOIMAGEN$, Optional TOPX, Optional LEFTX, Optional HEIGHTX, Optional WIDTHX)
    'Columnas_a_Leer$ = NUMEROS DE COLUMNAS QUE VA A LEER DEL MSF SEPARADOS PRO PUNTO Y COMA "1;3;5"
    'A_Partir_Fila% = A PARTIR DE QUE FILA DEL EXCEL EMPIEZA A ESCRIBIR LOS DATOS
    'En_Columna%= EN QUE COLUMNA DEL EXCEL VA A ESCRIBIR LOS DATOS
    'SI Columnas_a_Leer$ ES MAS DE UNA LAS ESCRIBE EN EL EXCEL EN LA COLUMNA DE AL LADO
    
    'EJEMPLO SI Columnas_a_Leer$ = "1;4" Y En_Columna% = 5
    'VA A ESCRIBIR LA INFORMACION EN EL EXCEL _
    EN LA COLUMNA 5 Y 6
    'SI SE PASA UN VALOR COMO DATO FIJO Y DENTRO DE LAS COLUMNAS A LEER INCLUYE EL 0 CUANDO LEA ESTA ASIGNA EL VALORFIJO
    
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    Call TEXTOLOTES(Columnas_a_Leer$, ";")
    
    Set ObjetoExcel = CreateObject("Excel.Application")
    With ObjetoExcel
        ' ARCHIVO EXCEL / ABRE EL LIBRO
        Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)

        If TRIM$(ARCHIVOIMAGEN$) <> "" Then
          'IMPRIME UNA IMAGEN EN LA HOJA
          'objHoja.Shapes.AddPicture <Name file>, False, True, 445, 200, 250, 150
          Libro.Worksheets(NombreHoja$).Shapes.AddPicture "C:\FLEXOFT\QR_ETIQPTERM.JPG", False, True, LEFTX, TOPX, WIDTHX, HEIGHTX
        End If
         k& = 0
         For i& = 1 To MSF.Rows - 1
            For j& = 1 To UBound(CADENATEX$)
               Col_a_Leer% = CADENATEX$(j&)
               DATO$ = TRIM$(MSF.TextMatrix(i&, Col_a_Leer%))
               If DATO$ = "" And SIVACIO% < 1 Then
                  DATO$ = "-----"
               End If
               If Left$(DATO$, 3) = "BR-" Then DATO$ = REPLACAR(DATO$, "BR-", "PR-")
               
               If Col_a_Leer% = 0 Then
                  If datofijo$ <> "" And STRINGFILA_STRING$(MSF, i&) <> "" Then
                     DATO$ = datofijo$
                  Else
                      DATO$ = "--"
                  End If
               End If

               Libro.Worksheets(NombreHoja).Cells(A_Partir_Fila% + k&, En_Columna% + j& - 1) = DATO$
            Next j&
            k& = k& + 1
         Next i&
         
         ObjetoExcel.ActiveWorkbook.Save
         Set Hoja = Libro.Sheets(NombreHoja$)
         If NOIMPRIMIR% > 0 Then GoTo 99
         Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
         'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada

    End With

    'CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0





End Sub

 Sub Command3_Click()
    OPX% = COMALTER%("CAMBIO LA HOJA?\\Sí,Continuar\Cancelar")
    If OPX% > 1 Then Exit Sub

DoEvents
    Call ImortacionRegDatoEym(Text1, TRIM$(Combo1.TEXT), MSF3, "", 1, 31, 6, "", 0, 1)
    Call COMUNI("TAREA FINALIZADA")
    
End Sub

 Sub GRABARIMPORTACIONABD()
 
  MSF5.Rows = 1
  NSERIE$ = Label13(2)
  REGERROR% = 0
  NROREGISTROEYM$ = LBLREGISTROSELECCIONADO
  'ESCRIBO EN UNA GRILLA LOS NSERIE EXISTENTES PREVIAMENTE

  If NumeroRegistroSegunSerieEym&(NSERIE$) > 0 Then
     MSF5.Rows = MSF5.Rows + 1
     REGERROR% = REGERROR% + 1
     MSF5.TextMatrix(REGERROR%, 1) = NSERIE$
     MSF5.TextMatrix(REGERROR%, 2) = "EncaRegArma"
     Exit Sub
  End If
  If ExisteNSerieEnTanumse%(NSERIE$) > 0 Then
     MSF5.Rows = MSF5.Rows + 1
     REGERROR% = REGERROR% + 1
     MSF5.TextMatrix(REGERROR%, 1) = NSERIE$
     MSF5.TextMatrix(REGERROR%, 2) = "Tanumse"
     Exit Sub
  End If
  '
  NORFA$ = Left$(Label13(4), 2)
  If Left(NORFA$, 2) = "OF" Then
    NORFA1$ = FORMATOCON0$(Mid$(Label13(4), 3), 6)
    NORFA$ = NORFA$ & "-" & NORFA1$ & "-1"
  Else
    NORFA$ = Label13(4)
  End If
'  On Error GoTo ERROR_GUARDAR
'  FLEXCONN.BeginTrans
  '
  SQLX$ = "SELECT * FROM EncaRegArma"
  SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NSERIE$ & "'"
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
    .AddNew   ' YA CONVERTIDO
    On Error Resume Next
    !CodiEmpr = CodiEmp%
    On Error GoTo 0
    !NUMEROSERIE = NSERIE$
    !NroRegistro = NROREGISTROEYM$
    !MODELO = Label13(0)
    !Revision = Label13(1)
    !FECHA = CVDAT(CVINT(Label13(3)))
    !IDSUBOR = NORFA$
    !DESV_TRANS = Label13(5)
    !PREPARO = Label13(6)
    !MARBORRA = 0
    .Update
  End With
  rs.Close
  Set rs = Nothing
  '
  'GRABO EN TABLA CPUREGARMA EL DATO DEL CPU (ORDEN ES EL DE PRESENTACION
  SQLX$ = "SELECT * FROM CpuRegArma"
  SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NSERIE$ & "'"
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
    For i& = 1 To msf2.Rows - 1
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
        On Error GoTo 0
       !ORDEN = i&
       !NUMEROSERIE = NSERIE$
       !CPU = msf2.TextMatrix(i&, 1)
       !MARBORRA = 0
       .Update
     Next i&
  End With
  rs.Close
  Set rs = Nothing
  
  'GRABO EN TABLA DetaRegArma EL DETALLE
  SQLX$ = "SELECT * FROM DetaRegArma"
  SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NSERIE$ & "'"
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
    For i& = 1 To MSF4.Rows - 1
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
        On Error GoTo 0
       !ORDEN = i&
       !NUMEROSERIE = NSERIE$
       !Codigo = Left$(MSF4.TextMatrix(i&, 1), 24)
       !Descripcion = MSF4.TextMatrix(i&, 2)
       !IDENTIFICACION = MSF4.TextMatrix(i&, 3)
       !NROdeDT = MSF4.TextMatrix(i&, 4)
       !COD_INTE = CODINT%(MSF4.TextMatrix(i&, 1))
       !MARBORRA = 0
       .Update
    Next i&
  End With
  rs.Close
  Set rs = Nothing
  '
  MESYANIO$ = Left(NSERIE$, 2)
  NORDEN% = XVALO(Right(NSERIE$, 3))
  
  '100-1001 (ejemplo de n. serie) 'calculo aprocimadamente la fecha de fabricacion
  FECOMIENZOAÑO = FECHANUM("01/01" & MESYANIO$)
  Semana$ = Mid$(NSERIE$, 3, 3): Semana$ = REPLACAR(Semana$, "-", "")
  CantidadediasConsiderandoLaSemana% = XVALO(Semana$) * 7
  fech% = DIASPOST(FECOMIENZOAÑO, CantidadediasConsiderandoLaSemana%)
  FETE$ = FECHATEX$(fech%)
  CI1% = CODINT%(Label13(0))
  'GRABO EN TANTUMSE LA MARCA QUE ES DATO IMPORTADO
  SQLX$ = "SELECT * FROM TANUMSE"
  SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NSERIE$ & "'"
  Set rs = New ADODB.Recordset
  rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With rs
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !NserieImpor = 1
 
        !NUMEROSERIE = NSERIE$
        !MESANIO = MESYANIO$
        !NUMEORD = NORDEN%
        !COD_INTE = CI1%
        !COD_EXTE = Label13(0)
        Descripcion$ = Left$(DESCRIT$(CI1%), 64)    'modificado por que devolvia un error de datos
        !Descrip = Descripcion$
        !FECHALTA = CVDAT(fech%)
        !TEN_ENSA = "220"
        !DEPONUMSE = 0
         FEGARANTIA% = DIASPOST(fech%, 730)
        !FEVENGAR = FETEXCOR1$(FEGARANTIA%)
        !IDSUBOR = NORFA$
        !NumRemito = "RT-X-9999-99999999"
        .Update
    End With
    rs.Close
    Set rs = Nothing
    '
    Call GENUMSE07.ACTLISELNUMSER   'Actualiza lista de selecciona de numeros de serie
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Alta N/Serie X Importación Excel"
    Call ANOTANUSE(NSERIE$, CUERPO$)

  
'ERROR_GUARDAR:
'    If Err.Number <> 0 Then
'       FLEXCONN.RollbackTrans
'       MsgBox Err.Description
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
'       On Error GoTo 0
'    Else
'      FLEXCONN.CommitTrans
'      On Error GoTo 0
'    End If
'


End Sub

Private Sub Command4_Click()
    Call ImortacionRegDatoEym(Text1, TRIM$(Combo1.TEXT), MSF3, "", 1, 31, 6, "", 0, 1, "", 0, 0, 0, 0, 1)

End Sub

Private Sub Command5_Click()
  If Command5.Caption = "DETENER" Then
     Command5.Caption = "CONTINUAR"
  Else
     Command5.Caption = "DETENER"
  End If
End Sub

Private Sub Form_Load()
Call GENUMSE07.ACTLISELNUMSER
'   Call CREAR_TABLA_IMPOREGDATOEYM
   Me.TXTREG = 6
   Me.TXTCOL = 31
   Campos$ = "REFERENCIA/A10"
   Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
   Campos$ = "CPU/A10"
   Call CARGA_ENCABEZADO(msf2, Campos$, Me)
   Campos$ = "CODIGO/A16;DESCRIPCION/A48;IDENTIFICACION/A32;N° DE D.T.:/F/8"
   Call CARGA_ENCABEZADO(MSF4, Campos$, Me)
   Campos$ = "N.SERIE/A16;TABLA/A24)"
   Call CARGA_ENCABEZADO(MSF5, Campos$, Me)
   Openforms = DoEvents
End Sub

Private Sub LBLREGISTROSELECCIONADO_Change()
Exit Sub
'ACA LEER EL EXCEL DE ACUERDO AL REGISTRO Y POSCIIONAR LOS VALORES EN PANTALLA
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    REGID = LBLREGISTROSELECCIONADO
    REG_ENCA_DESDE = 4 '
    REG_ENCA_HASTA = 8
    COL_ENCA = 20
    '
    REG& = REG_ENCA_DESDE
    IndiceObj = 0
    NombreHoja = Combo1
    Set ObjetoExcel = CreateObject("Excel.Application")
    Set Libro = ObjetoExcel.Workbooks.Open(Text1)
    With Libro.Worksheets(NombreHoja)
        
        Do While REG& <= REG_ENCA_HASTA
               DATO$ = .Cells(REG&, COL_ENCA)
               Me.Label13(IndiceObj) = DATO$
               IndiceObj = IndiceObj + 1
               REG& = REG& + 1
         Loop
    End With

    'CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0

End Sub

Private Sub MSF3_Click()
     REG& = MSF3.MouseRow
     COLUM% = MSF3.MouseCol
     If REG& = 0 Or REG& > MSF3.Rows - 1 Then Exit Sub
     LBLREGISTROSELECCIONADO = MSF3.TextMatrix(REG&, COLUM%)
End Sub

Private Sub Text1_Change()
  Combo1.Clear
  '
  
  If Text1.TEXT <> "" Then
'  MSF1.Clear
  msf2.Rows = 1
  MSF3.Rows = 1
  MSF4.Rows = 1
  Combo1.Clear
  'Combo1.Text = ""
  '
  If Text1.TEXT <> "" Then
     'Dim objexcel As Excel.Application
'     Set objexcel = New Excel.Application
'     Set Libro = objexcel.Workbooks.Open(Text1)
     Dim ObjetoExcel As Excel.Application
     Dim Libro As Excel.Workbook
     Dim Hoja As Excel.Worksheet
     Set ObjetoExcel = New Excel.Application
     Set Libro = ObjetoExcel.Workbooks.Open(Text1)
     '
     CANOJ% = Libro.Sheets.Count
     For KKI% = 1 To CANOJ%
        Combo1.AddItem Libro.Sheets(KKI%).Name
        If KKI% = 1 Then
           Combo1.TEXT = Libro.Sheets(KKI%).Name
        End If
     Next KKI%
'     Call Combo1_Click
    On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0
  End If

Exit Sub

Errores:
If Err.Number <> 0 Then
  ObjetoExcel.Quit
  Set ObjetoExcel = Nothing
  MsgBox Err.Description, vbCritical, CStr(Err.Number)
  Err.Clear
End If
End If
End Sub

Private Sub Combo1_Change()

   Data1.RecordSource = Combo1.TEXT
   Data1.Refresh
   Data1.Recordset.MoveFirst

End Sub

Sub CargoEncabezadoadeGrilla(MSF As MSFlexGrid, CMB As ComboBox, MASNUM%)
    CMB.Clear
    NUMERO$ = ""
    For c = 1 To MSF.Cols - 1
       If MASNUM% > 0 Then NUMERO$ = c & " - "
       CMB.AddItem NUMERO$ & MSF.TextMatrix(0, c)
    Next

End Sub


Private Sub Combo3_Click()
   LBLORDENA.Caption = Left(TRIM$(Combo3.TEXT), 2)
   lblcampo.Caption = TRIM(Mid(TRIM(Combo3.TEXT), 4))
End Sub
Function PRESENTAR_DIALOGO_GUARDAR$(CommonDialog1 As CommonDialog, TIPOARCHIVO$)
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Escriba/Seleccione el Archivo Excel"
        '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = TIPOARCHIVO$
        CommonDialog1.FilterIndex = 1
        On Error Resume Next
        CommonDialog1.ShowSave
        If Err > 0 Then GoTo 99
        VUEL% = 0
        DoEvents
        PRESENTAR_DIALOGO_GUARDAR$ = UCase$(CommonDialog1.FileName)
99
End Function

Private Sub Command1_Click()
     Text1 = PRESENTAR_DIALOGO_GUARDAR$(cmd, "Archivo Excel XLS|*.xls|XLSX|*.xlsx|")
     On Error Resume Next
      rs.Close
     On Error GoTo 0
     Call COMUNI("PASO 1 REALIZADO")
'    Call cerrar_excel

End Sub
