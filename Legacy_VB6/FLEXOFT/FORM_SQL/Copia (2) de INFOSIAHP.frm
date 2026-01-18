VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form INFOSIAHP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE SITUACIÓN DE STOCKS DE ITEM ACTIVO"
   ClientHeight    =   8760
   ClientLeft      =   -3150
   ClientTop       =   540
   ClientWidth     =   15150
   ControlBox      =   0   'False
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   15150
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   15135
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   1080
         Width           =   1470
      End
      Begin VB.TextBox TXTCOSTO 
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
         Height          =   380
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "Costo de Maestro de Articulos"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.TextBox Text40 
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
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Vía de Entr."
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox TXTVIAENTREGA 
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
         Height          =   380
         Left            =   6870
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1320
         Width           =   1300
      End
      Begin VB.TextBox Text6 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Proveedor:"
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox TXTNUMEPROVE 
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
         Height          =   380
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   2880
         TabIndex        =   36
         ToolTipText     =   "Lista de Proveedores Relacionados con el Item Activo"
         Top             =   1320
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text6 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Código Cliente:"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtcodigo 
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
         Height          =   380
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox Text6 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text7 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   360
         Width           =   3345
      End
      Begin VB.TextBox TXTDESCRIPCION 
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
         Height          =   380
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   630
         Width           =   3345
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ing. sin Stocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   700
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Cant. Solicitada:"
         Top             =   360
         Width           =   1560
      End
      Begin VB.TextBox TXTSOLICITADA 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   380
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   1560
      End
      Begin VB.TextBox TXTSERVTECNICO 
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
         Height          =   1380
         Left            =   11040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         ToolTipText     =   "Especificación Del Maestro de Artículos"
         Top             =   360
         Width           =   3900
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Texto Técnico:"
         Top             =   360
         Width           =   1680
      End
      Begin VB.TextBox TXTPESO 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   380
         Left            =   6970
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   6970
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Peso en Kg.:"
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir sin Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1250
         Width           =   1695
      End
      Begin VB.TextBox TXTCODICLIENTE 
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
         Height          =   380
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Height          =   380
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6135
         Sorted          =   -1  'True
         TabIndex        =   42
         Text            =   "Combo6"
         Top             =   1315
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4200
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   15135
      Begin VB.CommandButton Command1 
         Caption         =   "Ingreso con Lotes"
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
         Left            =   9240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   3075
      End
      Begin VB.TextBox TXTSELECCIONADA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Cant.Selecc.:"
         Top             =   240
         Width           =   1440
      End
      Begin VB.TextBox Text6 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "En transito (Otros):"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TXTTRANSITO_OTROS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox Text36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Últimas 3 Cotizaciones"
         Top             =   600
         Width           =   7950
      End
      Begin VB.TextBox Text37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Últimas 3 Ventas"
         Top             =   2400
         Width           =   7950
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Estadisticas de Ventas"
         Top             =   2400
         Width           =   5750
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Estadísticas de Cotizaciones"
         Top             =   600
         Width           =   5750
      End
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   1395
         Left            =   9240
         TabIndex        =   14
         Top             =   840
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   2461
         _Version        =   393216
         BackColorFixed  =   8421504
         SelectionMode   =   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1395
         Left            =   9240
         TabIndex        =   15
         Top             =   2640
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   2461
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         AllowUserResizing=   3
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   1395
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   2461
         _Version        =   393216
         BackColorFixed  =   16501405
         SelectionMode   =   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   1395
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   2461
         _Version        =   393216
         BackColorFixed  =   16501405
         SelectionMode   =   1
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
      Caption         =   "Selección "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   15135
      Begin VB.TextBox TXTMOVIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Height          =   315
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2355
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   4154
         _Version        =   393216
         BackColorFixed  =   16501405
         SelectionMode   =   1
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "INFOSIAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODO1$
Public FILAX&
'1.CODIGO
'2.CODIGO CLIENTE
'3.DESCRIPCION
'4.CANTIDAD
'5.NUMERO PROVEEDOR ' OCULTO
'6.R.S.PROVEEDOR
'7.COSTO
'8.PR.VENTA
'9.IMPORTE
'10.VIA ENTREGA
'11.LOTES OCULTO
'12.CANTIDAD DE LOTES OCULTO
'13.NUMERO DE PROVEEDOR OCULTO
'14.KILOS OCULTO
'15.DEMORA OCULTO
Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If TRIM$(txtcodigo.TEXT) = "" Then
       Call COMUNI("Código Inexistente o No Válido")
       GoTo 90
    End If
    '
    CANTSELEC = XVALO(TXTSELECCIONADA)
    NECESIDAD = XVALO(TXTSOLICITADA)
    If CANTSELEC > NECESIDAD Then
       Call COMUNI("Cantidad Seleccionada Supera a la Cantidad a Cotizar")
       GoTo 99
    End If
    '
    If NECESIDAD > CANTSELEC Then
       Call COMUNI("Cantidad Seleccionada No Alcanza Para la Cantidad a Cotizar\Modifique Cantidad")
       GoTo 99
    End If
    '
    'VALIDO QUE NO SE HAYA SELECCIONADO MAS CANTIDAD DE LOS LOTES DISPONIBLES
    If FILAX& = 0 Then
      Call CONTROL_CANT_LOTE_MSF1(OK%)
      If OK% < 1 Then GoTo 99
    End If
    '
    If XVALO(TXTSELECCIONADA) < 1 Then
      Call COMUNI("Falta Seleccionar Cantidad de Lote")
      GoTo 99
    End If
    '
    EDICION% = 0
    '
    For I& = 1 To MSF1.Rows - 1
     CANTI = XVALO(MSF1.TextMatrix(I&, 11))
     If CANTI > 0 Then
        FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
        FRMPRESUAHP.Text15 = FORMATNUM$(XVALO(MSF1.TextMatrix(I&, 7)) + XVALO(MSF1.TextMatrix(I&, 8)), "F12.2")
        '
        FRMPRESUAHP.Text38 = "LOCAL" ' MSF1.TextMatrix(I&, 2)
        FRMPRESUAHP.TXTPREUNIT = 250 'FORMATNUM$(XVALO(MSF1.TextMatrix(I&, 5)), "F10.2")
        FRMPRESUAHP.TXTCANTIDAD = TRIM$(FORMATNUM$(CANTI, "F9.1"))
        FRMPRESUAHP.TXTPESOP = TRIM$(FORMATNUM$(XVALO(TXTPESO), "F9.2"))
        Call FRMPRESUAHP.cargar_datos ' cargo los datos para crear el nuevo registro en la pantalla de presupuesto
          NPROV = XVALO(MSF1.TextMatrix(I&, 14))
          'EL NUMERO DEL PROVEEDOR SE ESTA GUARDANDO EN DOS CELDAS DIFERENTES !!!HAY QUE CONTROLARLO!!!
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 5) = NPROV
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 13) = NPROV
        
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 6) = TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 11) = MSF1.TextMatrix(I&, 13)
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 14) = FRMPRESUAHP.TXTPESOP
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 15) = "Inmediata" 'DEBERIA SER SIEMPRE INMEDIATA POR QUE AQUI ES CON STOCKS.

     End If
    Next I&
    '
    Call FRMPRESUAHP.COLOREARSTOCKS(REG&)

    MODO1$ = "ACTUALIZAR"
90  Hide
    
99  Command1.Enabled = True

End Sub
Sub ACTUCANTI()
   ACUMULACANTI = 0
   'If TXTMOVIL.Visible = True Then Exit Sub
   For I& = 1 To MSF1.Rows - 1
        CANTI = XVALO(MSF1.TextMatrix(I&, 11))
        ACUMULACANTI = ACUMULACANTI + CANTI
   Next I&
   TXTSELECCIONADA = FORMATNUM$(ACUMULACANTI, "F10.1")
End Sub
Sub CONTROL_CANT_LOTE_MSF1(OK%)
   OK% = 1
   CI1% = CODINT(txtcodigo)
   For I& = 1 To MSF1.Rows - 1
        LOTE$ = MSF1.TextMatrix(I&, 13)
        ST = XVALO(MSF1.TextMatrix(I&, 12))
        CANTI = XVALO(MSF1.TextMatrix(I&, 11))
        '
        If CANTI > ST Then
             Call COMUNI("Cantidad Seleccionada para Lote: " & LOTE$ & " Supera el Stock del Mismo.")
             OK% = 0
             Exit For
        End If
        '
        'VALIDA SI HAY CANTIDAD YA SELECCIONADA EN EL PRESUPUESTO + LA SELECCIONADA SI ALCANZA EL STOCKS
        CANTIGRI = CANTILOTE_EN_GRILLA(CI1%, LOTE$) ' CANTIDAD YA INGRESADA EN GRILLA DE PRESUPUESTO DEL MISMO LOTE
        If CANTIGRI > 0 Then
           If (CANTI + CANTIGRI) > ST Then
              TX$ = "La Cantidad (" & CANTI & ") Seleccionada para Lote: " & LOTE$
              TX$ = TX$ + "\Sumada a La Cantidad (" & CANTIGRI & ") Ya ingresada en Detalle de Presupuesto Activo"
              TX$ = TX$ + "\Supera La Cantidad (" & ST & ") del Stock Disponible."
              Call COMUNI(TX$)
              OK% = 0
              Exit For
           End If
        End If
   Next I&
End Sub

Function CANTILOTE_EN_GRILLA(CI1%, LOTEPARAM$)
    ACUCANTI = 0
    With FRMPRESUAHP
        LOTEPARAM$ = TRIM$(LOTEPARAM$)
        For I& = 1 To .MSF1.Rows - 1
           STRINGLOTES$ = .MSF1.TextMatrix(I&, 11)
           If InStr(STRINGLOTES$, LOTEPARAM$) > 0 Then
                  CANTI = XVALO(.MSF1.TextMatrix(I&, 11))
                  ACUCANTI = ACUCANTI + CANTI
           End If
        Next I&
    End With
    '
    CANTILOTE_EN_GRILLA = ACUCANTI
    '
End Function

Private Sub Command2_Click()
   FRMPRESUAHP.EDICION = 0
   MODO1$ = "CANCEL"
   Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
     CANTI = XVALO(TXTSOLICITADA)
     If CANTI > 0 Then
        FRMPRESUAHP.txtcodigo = txtcodigo 'LE ASIGNO EL CODIGO VERDADERO POR SI VIENE DE UN CODIGO DE RELACION (SI NO TENGO PROBLEMAS EN LA VALIDACION)
        FRMPRESUAHP.Text15 = FORMATNUM$(XVALO(MSF1.TextMatrix(I&, 7)) + XVALO(MSF1.TextMatrix(I&, 8)), "F12.2")
        '
        FRMPRESUAHP.Text38 = "LOCAL" ' MSF1.TextMatrix(I&, 2)
        FRMPRESUAHP.TXTPREUNIT = 250 'FORMATNUM$(XVALO(MSF1.TextMatrix(I&, 5)), "F10.2")
        FRMPRESUAHP.TXTCANTIDAD = TRIM$(FORMATNUM$(XVALO(TXTSOLICITADA), "F9.1"))
        FRMPRESUAHP.TXTPESOP = TRIM$(FORMATNUM$(XVALO(TXTPESO), "F9.2"))
        NPROV = XVALO(TXTNUMEPROVE)
        FRMPRESUAHP.TXTNUMPROV = NPROV 'TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor
        Call FRMPRESUAHP.cargar_datos_sin_stock ' cargo los datos para crear el nuevo registro en la pantalla de presupuesto
          'EL NUMERO DEL PROVEEDOR SE ESTA GUARDANDO EN DOS CELDAS DIFERENTES !!!HAY QUE CONTROLARLO!!!
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 5) = NPROV
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 13) = NPROV
        
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 6) = TEL2CLI$(-1 * NPROV) 'agrego en el nuevo registro el nombre del proveedor
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 11) = ""
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 14) = TXTPESO
        FRMPRESUAHP.MSF1.TextMatrix(FRMPRESUAHP.MSF1.Rows - 1, 15) = "Inmediata" 'DEBERIA SER SIEMPRE INMEDIATA POR QUE AQUI ES CON STOCKS.
        Hide
     End If
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()
  Command5.Enabled = False
         Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         TXTNUMEPROVE = NPRO%
   Command5.Enabled = True
End Sub

Private Sub MSF1_Click()
          COLUM = MSF1.MouseCol
          If COLUM <> 11 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF1.COL = 11
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            .TEXT = MSF1.TEXT
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
          Call ACTUCANTI
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub

Private Sub TXTDESCRIPCION_Change()

    Call CARGA_ENCABEZADO(INFOSIAHP.MSF3, "Presup./f8.0;Cliente/A24;M./A3;P.Unit./F10.2;Fecha/D8", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF2, "Factura./A20;Cliente/A24;P.Unit./F10.2", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF4, "Referencia/A32;Cant.Cotiz./F15", INFOSIAHP)
    Call CARGA_ENCABEZADO(INFOSIAHP.MSF5, "Referencia/A32;Cant.Ventas/F15", INFOSIAHP)
    '
    NC = XVALO(FRMPRESUAHP.Text1)
    INFOSIAHP.MSF2.Rows = 1: INFOSIAHP.MSF3.Rows = 1: INFOSIAHP.MSF4.Rows = 1: INFOSIAHP.MSF5.Rows = 1
    
    If TRIM$(TXTDESCRIPCION.TEXT) <> "" Then
       CI1% = CODINT(txtcodigo.TEXT)
       If CI1% > 0 Then
          Call ULTIMOS3PRESUPUESTOS(INFOSIAHP.MSF3, CI1%, 0)
          Call ULTIMOS3VENTAS(INFOSIAHP.MSF2, CI1%, 0)
          Call ESTADIS_COTIZA(INFOSIAHP.MSF4, CI1%)
          Call ESTADIS_VENTAS(INFOSIAHP.MSF5, CI1%)
          TXTSERVTECNICO = ESPECIFIT$(CI1%)
       End If
    End If

End Sub
Sub ESTADIS_VENTAS(MSF As MSFlexGrid, CI1%)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT90%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_90 = INGRE - SALE
    '
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT180%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_180 = INGRE - SALE
    '
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FECHMOV >= '" & FETEXCOR1$(DIASANT365%) & "' AND CODIMOVI = 'RT' AND COD_INTE = " & CI1%
    SALE = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
    INGRE = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
    ULT_365 = INGRE - SALE
    '
    MSF.Rows = 4: MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    '
    
End Sub

Sub ESTADIS_COTIZA(MSF As MSFlexGrid, CI1%)
    DIASANT90% = DIANTE(HOY(HO$), 90)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT90%) & "' AND CodiInt = " & CI1%
    ULT_90 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    DIASANT180% = DIANTE(HOY(HO$), 180)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT180%) & "'AND CodiInt = " & CI1%
    ULT_180 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    DIASANT365% = DIANTE(HOY(HO$), 365)
    CONDI$ = "FechEmis >= '" & FETEXCOR1$(DIASANT365%) & "'AND CodiInt = " & CI1%
    ULT_365 = CantRegistros("PRESUVEN", CONDI$, "NOMESS")
    '
    MSF.Rows = 4
    MSF.Cols = 3
    MSF.TextMatrix(1, 1) = "Últimos 3 Meses": MSF.TextMatrix(1, 2) = ULT_90
    MSF.TextMatrix(2, 1) = "Últimos 6 Meses": MSF.TextMatrix(2, 2) = ULT_180
    MSF.TextMatrix(3, 1) = "Últimos 12 Meses": MSF.TextMatrix(3, 2) = ULT_365
    '
End Sub
Sub ULTIMOS3PRESUPUESTOS(MSF As MSFlexGrid, CI1%, NC)
     MSF.Rows = 1: MSF.Cols = 6
     SQLX$ = "SELECT TOP 3 NroPres, RasoClie, PreUnit,NroClie,MonedEmi,COTIZAMONEDA, FechEmis"
     SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE CODIINT =  " & CI1%
     If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND NroClie =  " & NC
     SQLX$ = SQLX$ + " ORDER BY NroPres desc"
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        j& = j& + 1
        MSF.Rows = j& + 1
        MSF.TextMatrix(j&, 1) = XVALO(!NroPres)
        MSF.TextMatrix(j&, 2) = RASOCLI$(XVALO(!NroClie))
        MON% = XVALO(!MonedEmi)
        PRECIOUNITARIO# = XVALO(!PreUnit) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZAMONEDA) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TIPOCAMBIODEPRESUP# ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               MON% = 2
           End If
        End If
        MSF.TextMatrix(j&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(j&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        FE% = CVINT(!FechEmis)
        
        MSF.TextMatrix(j&, 5) = FECHATEX$(FE%)
        .MoveNext
       Loop
     End With

End Sub
Sub ULTIMOS3VENTAS(MSF As MSFlexGrid, CI1%, NC)

     MSF.Rows = 1: MSF.Cols = 5
     SQLX$ = "SELECT TOP 3 DOSECLAR, VALOUNIT,Nume_Clie"
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE COD_INTE =  " & CI1%
     If NC > 0 Then SQLX$ = SQLX$ + " AND  Nume_Clie =  " & NC
     SQLX$ = SQLX$ + " AND CODIMOVI = 'RT'"
     SQLX$ = SQLX$ + " ORDER BY DOSECLAR DESC "
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        j& = j& + 1
        MSF.Rows = j& + 1
        MSF.TextMatrix(j&, 1) = SAFETEXT$(!DOSECLAR)
        MSF.TextMatrix(j&, 2) = RASOCLI$(XVALO(!Nume_Clie))
        MSF.TextMatrix(j&, 3) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
        '
        'MON% = XVALO(!Mone_Emis)
        PRECIOUNITARIO# = XVALO(!VALOUNIT) 'POR DEFECTO TRAE EL IMPORTE EN EL QUE FUE GUARDADO.
        If MON% = 1 Then 'SI LA MONEDA EN QUE ESTA GUARDADA AL PRESUPUESTO ES PESOS
           TIPOCAMBIODEPRESUP# = XVALO(!COTIZA_DOLAR) 'TOMA EL CAMBIO EN QUE FUE GUARDADO
           If TIPOCAMBIODEPRESUP# > 0.0005 Then       'SI  TIENE GUARDADO IMPORTE DE COTIZACION PRESENTA EN ESE (DOLARES)
               PRECIOUNITARIO# = XVALO(!PreUnit) / TIPOCAMBIODEPRESUP# ' CALCULA
           Else
               PRECIOUNITARIO# = XVALO(!PreUnit) / TICAMONE(2)    'CASO CONTRARIO TOMA EL VALOR ACTUAL DE DOLARES
               MON% = 2
           End If
        End If
        
        MSF.TextMatrix(j&, 3) = SIMBOLPES$(MON%)
        MSF.TextMatrix(j&, 4) = TRIM$(FORMATNUM$(PRECIOUNITARIO#, "F12.2"))
        '
        .MoveNext
       Loop
     End With
     
End Sub


Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If XVALO(TXTMOVIL) > XVALO(MSF1.TextMatrix(MSF1.Row, 12)) Then
       Call COMUNI("Cantidad Ingresada Supera la Disponible, Verifique!!!!")
       TXTMOVIL = ""
     Else
      MSF1.TEXT = TXTMOVIL
      TXTMOVIL.Visible = False
      Call ACTUCANTI
     End If
   End If
End Sub

Private Sub TXTNUMEPROVE_Change()
    NPROV = XVALO(TXTNUMEPROVE)
    If NPROV > 0 Then
      Text1 = ECOARCH$("PROVED1", MKI$(NPROV))
    Else
      Text1 = ""
    End If
    
End Sub
Private Sub Combo6_Click()
    TXTVIAENTREGA.TEXT = Combo6.TEXT
End Sub

'Private Sub TXTMOVIL_LostFocus()
'   Call TXTMOVIL_KeyPress(13)
'End Sub
Private Sub TXTSOLICITADA_Change()

End Sub

Private Sub TXTSOLICITADA_DblClick()
    If TXTSOLICITADA.Locked = False Then
       TXTSOLICITADA.BackColor = &HE0E0E0
       TXTSOLICITADA.Locked = True
    Else
       TXTSOLICITADA.Locked = False
       TXTSOLICITADA.BackColor = &H80000005
       Call PINTATEXT(TXTSOLICITADA)
    End If

End Sub
Sub CARGARCOMBOVIAS_TODOS()
     Combo6.Clear
     FIN& = ULTREG&("KILOVIAS")
     '
     For I& = 1 To FIN&
        X$ = REGLEIDO$("KILOVIAS", I&)
        VIAX$ = TRIM$(CVI(Mid$(X$, 1, 2)))
        REFVIA$ = ECOARCH$("VIASENTR", MKI$(XVALO(VIAX$)))
        Call REPLA(TX$, VIAX$, 1, 2)
        Call REPLA(TX$, "-", 3, 1)
        Call REPLA(TX$, REFVIA$, 4, 8)
        Combo6.AddItem TX$
     Next I&
     '

End Sub

Function ViaEntrega(LOTE$)
   CONDI$ = "IdenLote = '" & LOTE$ & "'"
   NOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
   If NOC > 0 Then
      CONDI$ = "Nume_Ocom = " & NOC&
      ViaEntrega = XVALO(VALOBADA("OCOMENCA", "Via_Entrega", CONDI$, "NOMESS"))
   End If
End Function

