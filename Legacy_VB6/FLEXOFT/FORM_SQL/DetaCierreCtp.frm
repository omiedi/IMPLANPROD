VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DetaCierreCtp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Situación de Componentes para el Cierre de Orden de Fabricación"
   ClientHeight    =   8970
   ClientLeft      =   75
   ClientTop       =   585
   ClientWidth     =   18105
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8970
   ScaleWidth      =   18105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   18135
      Begin VB.CommandButton Command2 
         Caption         =   ">>>>>>Continuar >>>>>>"
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
         TabIndex        =   34
         Top             =   8520
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ver Detalle de Mat.en Linea"
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
         Left            =   14760
         TabIndex        =   33
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TXTCODINT 
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
         Height          =   350
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   8760
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Cód.Interno"
         Top             =   8520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H009FBE85&
         Height          =   855
         Left            =   120
         TabIndex        =   4
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
            Index           =   7
            Left            =   13320
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Cerrar X"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTACERRAR 
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
            Left            =   13320
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            Left            =   12180
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Cant.Aprob."
            Top             =   180
            Width           =   1140
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
            Left            =   9900
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   180
            Width           =   1140
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
            Left            =   9900
            Locked          =   -1  'True
            TabIndex        =   16
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
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   180
            Width           =   2235
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
            Left            =   12180
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            Left            =   11040
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Cant.Proyec."
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
            Left            =   11040
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   180
            Width           =   3800
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
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   3800
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
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   420
            Width           =   2235
         End
         Begin VB.TextBox Text14 
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
            Index           =   8
            Left            =   3465
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "N°.O."
            Top             =   180
            Width           =   500
         End
         Begin VB.TextBox Text14 
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
            Index           =   6
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "F.Técnica"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTFTECNICA 
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
            Height          =   350
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Index           =   5
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Cotización"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTNROCOTIZA 
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
            Height          =   350
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox Text14 
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
            Index           =   3
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Pedido"
            Top             =   180
            Width           =   1140
         End
         Begin VB.TextBox TXTNROPED 
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
            Height          =   350
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   420
            Width           =   1140
         End
         Begin VB.TextBox TXTNORPED 
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
            Height          =   350
            Left            =   3465
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   420
            Width           =   500
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4365
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   17955
         _ExtentX        =   31671
         _ExtentY        =   7699
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2670
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   4710
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2325
         Left            =   120
         TabIndex        =   27
         Top             =   6360
         Width           =   17955
         _ExtentX        =   31671
         _ExtentY        =   4101
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
      Begin VB.Label Label3 
         Caption         =   $"DetaCierreCtp.frx":0000
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2280
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   9495
      End
      Begin VB.Label Label1 
         BackColor       =   &H009FBE85&
         Caption         =   "Necesidad de Consumo"
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
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   6000
         Width           =   2415
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
         TabIndex        =   31
         Top             =   2900
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H009FBE85&
         Caption         =   "Material En linea "
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
         TabIndex        =   30
         Top             =   960
         Width           =   1695
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
      OleObjectBlob   =   "DetaCierreCtp.frx":00D7
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DetaCierreCtp.frx":030B
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
      Begin VB.Menu mnuINGTELBOL 
         Caption         =   "Ingreso de Telas a Dep. Corte P/Bolsas"
      End
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
      Begin VB.Menu mnuIngresoMatCorte 
         Caption         =   "Ingreso de Bobinas a Corte"
      End
      Begin VB.Menu mnuTransfLinea 
         Caption         =   "Transferencia a Linea"
      End
      Begin VB.Menu s 
         Caption         =   "- "
      End
      Begin VB.Menu mnuinformarRetiCorte 
         Caption         =   "Informar Retiro de Cortes"
      End
      Begin VB.Menu mnuRepDiario 
         Caption         =   "Reporte Diario"
      End
   End
End
Attribute VB_Name = "DetaCierreCtp"
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



Function APLICARCONSUMO%(NORFA$, DescripOper$, CANTI_ACERRAR)
'****  tener en cuenta la operacion como filtro *****
'ESTA FUNCION DEVUELVE 1 SI PUDO APLICAR EL CONSUMO A LA OPERACION PARAMETRIZADA
'RECORRE LA TABLA INFOFAB REVISANDO LOS INFORMES DE CORTES QUE ESTAN EN LA LINEA
'Y REGISTRA EN EL CAMPO CONSUMO LA CANTIDAD, EN LA MEDIDA QUE CONSUMO TENGA UN SALDO (cantreal - CONSUMO)
'VALIDO PARA BOLSONES
'
        APLICARCONSUMO% = 0
        CIFAB% = CODIFAB%(NORFA$)
        CantDeCierre = CANTI_ACERRAR
        Dim CTP As New Conteplast
        'ACA TENGO LOS SALDOS DE CADA INFORME DE FABRICACION
        SQLX$ = "SELECT   I.DescOper"
        SQLX$ = SQLX$ + "  ,I.IdenLote,(I.cantreal - I.CONSUMO) AS SALDO"
        SQLX$ = SQLX$ + " FROM infofab I  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2"
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
        SQLX$ = SQLX$ + " AND idcoordubic = 100"
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " AND I.DescOper = '" & DescripOper$ & "'"
        SQLX$ = SQLX$ + " order by saldo "
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With rs
          Do While Not .EOF
            SaldoActual = XVALO(!SALDO)
            If SALDO >= CantDeCierre Then
               Aconsumir = CantDeCierre
            Else
               Aconsumir = SaldoActual
            End If
            !CONSUMO = XVALO(!CONSUMO) + Aconsumir
            .Update
            ConsumoAplicado = ConsumoAplicado + Aconsumir
            If ConsumoAplicado >= CantDeCierre Then
               APLICARCONSUMO% = 1
               Exit Do 'si ya cumplio con la asignacion del consumo sale
            End If
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing



End Function

Sub CARGARMSF1(NORFA$)
        REG& = 0
        Call CREATABLA_OFABRIC
        Call PoneEnCeroNull("INFOFAB", "CONSUMO", REGAF&, "NOMESS")

        Campos$ = "Descripción Rango/A48;N.Op./f5;Desc.Oper/A24;Cortes Disponibles/f24;Fact.Corte/a0;CORTE/a0;Cod.Rango/A12"
        Call CARGA_ENCABEZADO(MSF1, Campos$, Me)
'        SQLX$ = " SELECT I.numeoper , I.DescOper ,SUM(cantreal - Consumo) AS CantCortes, FACTORCORTE,CORTE ,c.CODRANGO "
        SQLX$ = " SELECT I.numeoper , I.DescOper,I.FACTORCORTE,SUM(cantreal - Consumo) AS CantCortes,I.CORTE,c.CODRANGO ,C.Descripcion"
        SQLX$ = SQLX$ + " FROM infofab I WITH(NOLOCK) "
        SQLX$ = SQLX$ + " inner join CODIGORANGO  c on  i.CodiElem = c.CIMATPRIMA "
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2 "
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE and I.INFOFAB_ID = I2.OperFucionada) "
        SQLX$ = SQLX$ + " AND idcoordubic = 100 " 'LINEA
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        
        SQLX$ = SQLX$ + " and LEFT(I.DescOper,8) <> 'CINTILLA'"
        SQLX$ = SQLX$ + " and LEFT(I.DescOper,4) <> 'HILO'"
        
'        SQLX$ = SQLX$ + " GROUP BY  i.Cod_Exte, I.numeoper , I.DescOper ,i.CodiElem,I.CodXElem"
'        SQLX$ = SQLX$ + " ,I.FACTORCORTE,I.CORTE,c.CODRANGO "
        SQLX$ = SQLX$ + " GROUP BY   I.numeoper , I.DescOper,I.FACTORCORTE,I.CORTE,c.CODRANGO,c.Descripcion "

        SQLX$ = SQLX$ + " order by numeoper  "
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            REG& = REG& + 1
            MSF1.Rows = REG& + 1
            MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Descripcion)
            MSF1.TextMatrix(REG&, 2) = XVALO(!NUMEOPER)
'            DESOP$ = SAFETEXT$(!descoper)
'            POSGUION% = InStr(DESOP$, "-") ' PARA ACUMULAR LATERALES Y LATERALES-SEMITERMINADOS
'            MSF1.TextMatrix(REG&, 3) = Left$(DESOP$, POSGUION% - 1)
            MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!descoper)
            MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CantCortes), "F9")
            MSF1.TextMatrix(REG&, 5) = XVALO(!FACTORCORTE)
            MSF1.TextMatrix(REG&, 6) = XVALO(!corte)
            MSF1.TextMatrix(REG&, 7) = SAFETEXT$(!CODRANGO)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
        


                    
    
End Sub

Sub NuevaCARGARMSF1(NORFA$)
        REG& = 0
        Call CREATABLA_OFABRIC
        Call PoneEnCeroNull("INFOFAB", "CONSUMO", REGAF&, "NOMESS")

        Campos$ = "Descripción Rango/A48;N.Op./f5;Desc.Oper/A24;Cortes Disponibles/f24;Fact.Corte/a0;CORTE/a0;Cod.Rango/A12"
        Call CARGA_ENCABEZADO(MSF1, Campos$, Me)
        SQLX$ = " SELECT CodiElem,CodXElem,Descrip, NumeOper,DescOper,FACTORCORTE,CORTE,IDENLOTE"
        SQLX$ = SQLX$ + " , SUM(CASE WHEN CantReal > 0 THEN CantReal  ELSE -CantSalid END) AS saldoCantiCortes"
        SQLX$ = SQLX$ + " From InfoFab WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE    IDSUBOR = '" & NORFA$ & "' AND IDCOORDUBIC = 100" 'linea
        SQLX$ = SQLX$ + " and LEFT(DescOper,8) <> 'CINTILLA' and LEFT(DescOper,4) <> 'HILO'"
        SQLX$ = SQLX$ + " GROUP BY CodiElem,CodXElem,Descrip,NumeOper,DescOper,FACTORCORTE,CORTE,IDENLOTE"
        SQLX$ = SQLX$ + " order by  NUMEOPER,IDENLOTE"
        
        
'        SQLX$ = " SELECT SUM(CASE WHEN cantreal > 0 THEN PESO  ELSE -PESO END) AS kg,"
'        SQLX$ = SQLX$ + " SUM(CASE WHEN cantreal > 0 THEN cantreal  ELSE -cantsalid END) AS saldoCantiCortes,"
'        SQLX$ = SQLX$ + " COUNT(*) AS CANTREG ,descrip,cod_inte,idenlote, NumeOper"
'        SQLX$ = SQLX$ + " ,descoper,codxelem ,Cod_Exte,COORDUBIC,IDCOORDUBIC,FACTORCORTE,CORTE"
'        SQLX$ = SQLX$ + " ,NUMOPERARIO,NUMOPERARIO,NameOperario,c.CODRANGO ,C.Descripcion"
'        SQLX$ = SQLX$ + " From infofab WITH(NOLOCK)   inner join CODIGORANGO  c on  CodiElem = c.CIMATPRIMA"
'
'        SQLX$ = SQLX$ + " WHERE    IDSUBOR = '" & NORFA$ & "' AND IDCOORDUBIC = 100" 'linea
'        SQLX$ = SQLX$ + " and LEFT(DescOper,8) <> 'CINTILLA' and LEFT(DescOper,4) <> 'HILO'"
'        SQLX$ = SQLX$ + " group by Cod_Exte,idenlote,NumeOper,COORDUBIC,IDCOORDUBIC,idenlote,descrip,cod_inte,idenlote, descoper"
'        SQLX$ = SQLX$ + " , codxelem,NUMOPERARIO,NUMOPERARIO,NameOperario,c.CODRANGO ,C.Descripcion,FACTORCORTE,CORTE"
'        SQLX$ = SQLX$ + " order by COORDUBIC, NUMEOPER,Cod_inte,IDENLOTE"

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            REG& = REG& + 1
            MSF1.Rows = REG& + 1
            CODRANGO$ = CODIGORANGOEXTE$(XVALO(!codielem))
            
            MSF1.TextMatrix(REG&, 1) = DESCRIPCIONCODIGORANGO(CODRANGO$)  'SAFETEXT$(!Descripcion)
            MSF1.TextMatrix(REG&, 2) = XVALO(!NUMEOPER)
'            DESOP$ = SAFETEXT$(!descoper)
'            POSGUION% = InStr(DESOP$, "-") ' PARA ACUMULAR LATERALES Y LATERALES-SEMITERMINADOS
'            MSF1.TextMatrix(REG&, 3) = Left$(DESOP$, POSGUION% - 1)
            MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!descoper)
            MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!saldoCantiCortes), "F9")
            MSF1.TextMatrix(REG&, 5) = XVALO(!FACTORCORTE)
            MSF1.TextMatrix(REG&, 6) = XVALO(!corte)
            MSF1.TextMatrix(REG&, 7) = CODRANGO$ 'SAFETEXT$(!CODRANGO)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
        


                    
    
End Sub


Sub CARGARMSF3BOLSAS(NORFA$, CantiaCerrar&)
    CIFAB% = CODIFAB%(NORFA$)
    Campos$ = "Código Rango/A12;Descripción/A64;Neces.Mts/f12;Disponible/F12.2;Falta/f12.2"
        Call CARGA_ENCABEZADO(MSF3, Campos$, Me)

    SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CodIConj=" & CIFAB%
    Set RstFormula = READSET(SQLX$)
    J& = 0
    With RstFormula
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                CodigoIntComponentes% = XVALO(!codielem)
                CodigoExtComponentes$ = SAFETEXT$(!CodXElem)
                CantidadUnitariaComponente = XVALO(!usobruto)
                NECESIDAD = CantiaCerrar& * CantidadUnitariaComponente
                DESCRIX$ = DESCRIT0$(CodigoIntComponentes%)
                J& = J& + 1
                MSF3.Rows = J& + 1
                MSF3.TextMatrix(J&, 1) = CodigoExtComponentes$
                MSF3.TextMatrix(J&, 2) = DESCRIX$
                MSF3.TextMatrix(J&, 3) = FORMATNUM$(NECESIDAD, "f12.2")
                MSF3.TextMatrix(J&, 4) = ""
                MSF3.TextMatrix(J&, 5) = ""
                .MoveNext
            Loop
        End If
    End With
    RstFormula.Close
    Set RstFormula = Nothing
End Sub



Sub LIMPIAR()

    Text11 = ""
    Text4 = ""
    Text10 = ""
    MSF.Rows = 1
    '
    Text6 = ""

End Sub








Private Sub Command1_Click()
  If Command1.Caption = "Ver Detalle de Mat.en Linea" Then
     Command1.Caption = "Ver Mat.en Linea Acumulado"
     MSF1.Visible = False
  Else
     Command1.Caption = "Ver Detalle de Mat.en Linea"
     MSF1.Visible = True
  End If
End Sub

Private Sub Command2_Click()
   HAYFALTANTESCTP% = 0
   For i& = 1 To MSF3.Rows - 1
      If XVALO(MSF3.TextMatrix(i&, 11)) > 0 Then
         HAYFALTANTESCTP% = 1
         Exit For
      End If
   Next i&
   Me.Hide
End Sub

Private Sub Form_Load()
  MSF1.Height = 4605
  MSF1.Left = 120
  MSF1.Top = 1400
  MSF1.Width = 17955
  
  MSF2.Height = 4605
  MSF2.Left = 120
  MSF2.Top = 1400
  MSF2.Width = 17955
  Me.HAYFALTANTESCTP% = 0
End Sub
Sub instructivo_Cintilla()

'24/09/2020 ' usa funcion 'RecuperarLoteCintilla' se pone en funcionamiento 25/09/20

''1.-Se deben transferir las telas para fabricacion de cintillas mediante formlario INGCINTILLA
' 2.-Se generan orden de fabricacion de modo normal y simple por una cantidad, siendo estas unidades
' 3 los codigos de  cintillas  van a poseer codigo rango de comoponente
' 4 Se cierra la orden de fabricacion y consume en modo fifo las telas transferidas al deposito de cintilla
' 5 Las cintillas se dan de alta en el deposito de corte.
' 6 Cuando se cierran las of de los bolsones se consumen del deposito de corte
' 7 La cantidad a consumir CUANDO SE CIERRA LA DEL BOLSON no va a ser  la de la formula en mts,
'   si no la del moduxxb.
' 8 El alta de la cintilla terminado se debe dar en el deposito de corte para que consumo de alli.

'**************************************************************************************
'CARGA DE FICHA TECNICA
'Cotizacion: En Coti1 se debera cargar la cintilla producto terminado (que podra estar asociada a un Codigo Rango)
'Este codigo Rango debera coincidir la unidad de medida 'U.'
'Se debera Poner un costo a la Cintilla P.Term. El mismo debera ser en mts. como en  las bobinas
'En la coti1 asi calcula de acuerdo al corte el costo
'En el armado de la Formula (Aprobacion de Coti) se considera que si es familia cintilla
'y es unidad de medida U. que el uso sea el valor del campo Moduxbb

'**************************************************************************************
'******* AVISAR ******* A RAUL QUE LAS CINTILLAS P.TERMI. NO DEBEN ESTAR DENTRO DE CODIGO RANGO.
'DADO QEU EL CODIGO RANGO TIENE SENTIDO CUANDO SE FABRICA LA CINTILLA, NO CUANDO ESTA SE USA DE COMPONENTE
'EN LOS BOLSONES.
'**************************************************************************************

'Configuraciones
'las Cintillas Prod.Term. Deberan tener un codigo para cada medida
'No Deberan Pertenecer a un codigo Rango, se deberan consmir en el cierre de modo fifo del deposito de corte
'Ver en el cierre el consumo
'Nota:
'  LA PANTALLA DE TRANSFERENCIA DE TELAS PARA CINTILLAS SE ACTUALIZA MOSTRANDO EL SALDO.


End Sub
Sub instructivo_Liner()
  'EJEMPLO OF 62
End Sub

Sub instructivo_Cortes()
'
'1 Debe estar dado de alta como familia de cortes (CT), CONTENER EL COMPONENTE (RANGO)
'2 Generar pedido
'3.-Genera of de pedido
'4.-Transfiere tela para corte sin operacion (ingreso de bobinas a corte)
'5.-NO genera informe de corte.
'6.-No hace falta que envie a linea
'7.-Cerrar la orden (En pantalla de Situacion de Cierre esta Capturado que si es Corte muestra lo transferido)
End Sub

Private Sub mnuIngresoMatCorte_Click()
INGREBOB.TXTOF = Text10
INGREBOB.Show 1
End Sub

Private Sub mnuTransfLinea_Click()
  TRAFALINEA.TXTOT = Text10
  TRAFALINEA.Show 1
End Sub
Private Sub mnuinformarRetiCorte_Click()
INFOCORTE.Label4 = Text10
INFOCORTE.Show 1
End Sub
Private Sub mnuRepDiario_Click()
   InfDiaSec.Label4 = Text10
   InfDiaSec.Show 1
End Sub

Private Sub Text10_Change()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NORFA$ = TRIM$(Text10)
   CI1% = CODIFAB%(NORFA$)
   TXTCODINT = CI1%
   Text3 = CantiProyectadaOF(NORFA$)
   Text1 = CantiAprobadaOF(NORFA$)
   TXTNROPED = NroPedidoSegunOf&(NORFA$)
   ORDENPEDIDO% = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
   Me.TXTNORPED = ORDENPEDIDO%
   TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPED), ORDENPEDIDO%)
   Codigo$ = CODEXT(CI1%)
   
   CF$ = TRIM$(UCase$(CODFAMIL$(CI1%))) 'CODIGO DE FAMILIA
   'SI ES CORTES O BOLSAS
   If Left$(Codigo$, 3) = "CTP" Or CF$ = "CT" Or CF$ = "SA" Then   'BOLSAS O CORTES  SOLO MUESTRA LO TRANSFERIDO PARA LA ORDEN (NO CORTES COMO BOLSONES
     Command1.Visible = False
     
     Call CARGARMSF1BOLSAS(NORFA$)
     Call CARGARMSF3BOLSAS(NORFA$, XVALO(TXTACERRAR))
     Call AsignarDisponibilidad_a_Necesidad
   Else
    'BOLSONES BIG BAG
   CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
   If (XVALO(Mid$(NORFA$, 4, 6)) >= 2341 Or XVALO(Mid$(NORFA$, 4, 6)) = 2325) And InStr(CODFAMI$, "BB") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then  'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
        Call NuevaCARGARMSF1(NORFA$)
        Call NuevaCargarMsf2(NORFA$)
   Else
        Call CARGARMSF1(NORFA$)
        Call CargarMsf2(NORFA$)
   End If
    
    Call CargarMsf3(NORFA$, XVALO(TXTACERRAR), XVALO(TXTCODINT), XVALO(TXTNROCOTIZA))
    For i& = 1 To MSF1.Rows - 1
        Campos$ = "Código/A18;N.Op./f5;Desc.Oper/A24;Cant.Cortes/f9;Fact.Corte/F12;CORTE/F10.2;Cod.Rango/A12"
        DescripOper$ = MSF1.TextMatrix(i&, 3)
        CantidadCortesDisponibles& = XVALO(MSF1.TextMatrix(i&, 4))
        Call AsignarOperacionANecesidad(DescripOper$, CantidadCortesDisponibles&)
    Next i&
   End If
End Sub
Sub AsignarDisponibilidad_a_Necesidad()
    Campos$ = "Código Rango/A12;Descripción/A24;Neces.Mts/f12;Disponible/F12.2;Falta/f12.2"
    '  Código/A20                   '  1
'  Descripción/A40              '  2
'  Lote/A16                     '  3
'  Transf.Mts./f12.1            '  4
'  CI/a0                        '  5
'  Op.N°/F6                     '  6
'  Kg.Transf/F11.2              '  7
'  Consumo Mts./F12.2           '  8
'  Saldo Mts./F12.2             '  9
'  Rango./a12                   '  10

   MaxIndex = 0
   Dim CantSaldo()
   ReDim Preserve CantSaldo(1)

   'EL SALDO A VECTOR POR CODIGO RANGO
   For i& = 1 To MSF1.Rows - 1
      SALDO = XVALO(MSF1.TextMatrix(i&, 9))
      Rango& = CODINT%(MSF1.TextMatrix(i&, 10))
      If MaxIndex < Rango& Then
         ReDim Preserve CantSaldo(Rango&)
         MaxIndex = Rango&
      End If
      CantSaldo(Rango&) = CantSaldo(Rango&) + SALDO
   Next i&
   '
    
   For i& = 1 To MSF3.Rows - 1
      NECESIDADMTS = XVALO(MSF3.TextMatrix(i&, 3))
      Rango& = CODINT%(MSF3.TextMatrix(i&, 1))
      If Rango& > UBound(CantSaldo) Then ReDim Preserve CantSaldo(Rango&)
      MSF3.TextMatrix(i&, 4) = CantSaldo(Rango&)
      
     If CantSaldo(Rango&) >= NECESIDADMTS Then
        FALTAN = 0
      Else
        FALTAN = NECESIDADMTS - CantSaldo(Rango&)
      End If
      MSF3.TextMatrix(i&, 5) = FALTAN
      If NECESIDADMTS > CantSaldo(Rango&) Then
         Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbRed, i&, 5)
      End If
      CantSaldo(Rango&) = CantSaldo(Rango&) - NECESIDADMTS: If CantSaldo(Rango&) < 0 Then CantSaldo(Rango&) = 0
   Next i&
   
End Sub
Sub AsignarOperacionANecesidad(DescripOperMSF1$, CantidadCortesDisponibles&)
        Campos$ = "Código/A1 '  1"
        Campos$ = Campos$ + ";Descripción/a24"             '  2"
        Campos$ = Campos$ + ";Componente/A24"               '  3"
        Campos$ = Campos$ + ";Cielem/A0"                    '  4"
        Campos$ = Campos$ + ";Moduxb/F12"                   '  5"
        Campos$ = Campos$ + ";CORTE/F10.2"                  '  6"
        Campos$ = Campos$ + ";Nec.M."                       '  7"
        Campos$ = Campos$ + ";Cod.Rango/A12"                '  8"
        Campos$ = Campos$ + ";Nec.Cortes/f10.1"             '  9
        Campos$ = Campos$ + ";Disponibilidad"               '  10
        Campos$ = Campos$ + ";Falta"                        '  11
    Dim CTP As New Conteplast
    For i& = 1 To MSF3.Rows - 1
       DescriOperMsf3$ = MSF3.TextMatrix(i&, 3)
       SectSinParentesisMSF3$ = CTP.SectorSinParentesis$(DescriOperMsf3$)
       CantCaracteres% = Len(SectSinParentesisMSF3$)
       If TRIM$(SectSinParentesisMSF3$) = TRIM$(Left(DescripOperMSF1$, CantCaracteres%)) Then
       'acumula por si esta mas de una vez en la grilla msf1
          MSF3.TextMatrix(i&, 10) = CantidadCortesDisponibles& + XVALO(MSF3.TextMatrix(i&, 10))
          FALTAN& = XVALO(MSF3.TextMatrix(i&, 9)) - XVALO(MSF3.TextMatrix(i&, 10))
          If FALTAN& < 1 Then FALTAN& = 0
          If FALTAN& > 0 Then
             Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbRed, i&, 11)
          Else
             Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbBlack, i&, 11)
          End If
          MSF3.TextMatrix(i&, 11) = FALTAN&
       End If
    Next i&
    
    
    Set CTP = Nothing
    
End Sub
Sub CargarMsf2(NORFA$)
        Campos$ = "Código/A18;N.Op./f5;Desc.Oper/A24;Cant.Cortes Originales/f22;Fact.Corte/F12;CORTE/F10.2;Lote/a16;Peso/F10.2;Cod.Rango/A12;Saldo/F10.2"
        Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
        Call PoneEnCeroNull("INFOFAB", "CONSUMO", REGAF&, "NOMESS")

        SQLX$ = "SELECT  i.CodXElem, I.numeoper , I.DescOper ,I.cantreal AS CantCortes,FACTORCORTE"
        SQLX$ = SQLX$ + ",I.CORTE,I.IdenLote,I.Peso,c.CODRANGO,SUM(I.cantreal - I.CONSUMO) AS SALDO"
        SQLX$ = SQLX$ + " FROM infofab I  WITH(NOLOCK)"
        SQLX$ = SQLX$ + "inner join CODIGORANGO  c on  I.CodiElem = c.CIMATPRIMA"
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2"
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE and I.INFOFAB_ID = I2.OperFucionada)"
        SQLX$ = SQLX$ + "AND idcoordubic = 100"
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " group by i.CodXElem, I.numeoper , I.DescOper ,I.cantreal ,i.FACTORCORTE"
        SQLX$ = SQLX$ + ",I.CORTE,I.IdenLote,I.Peso,c.CODRANGO"
        SQLX$ = SQLX$ + " order by I.DescOper,saldo,idenlote"
        REG& = 0

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            REG& = REG& + 1
            MSF2.Rows = REG& + 1
            MSF2.TextMatrix(REG&, 1) = SAFETEXT$(!CodXElem)
            MSF2.TextMatrix(REG&, 2) = XVALO(!NUMEOPER)
            MSF2.TextMatrix(REG&, 3) = SAFETEXT$(!descoper)
            MSF2.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CantCortes), "F9")
            MSF2.TextMatrix(REG&, 5) = XVALO(!FACTORCORTE)
            MSF2.TextMatrix(REG&, 6) = XVALO(!corte)
            MSF2.TextMatrix(REG&, 7) = SAFETEXT$(!idenlote)
            MSF2.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(!Peso), "F9.1")
            MSF2.TextMatrix(REG&, 9) = SAFETEXT$(!CODRANGO)
            MSF2.TextMatrix(REG&, 10) = SAFETEXT$(!SALDO)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
End Sub

Sub NuevaCargarMsf2(NORFA$)
        Campos$ = "Código/A18;N.Op./f5;Desc.Oper/A24;Cant.Cortes Originales/f22;Fact.Corte/F12;CORTE/F10.2;Lote/a16;Peso/F10.2;Cod.Rango/A12;Saldo/F10.2"
        Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
        Call PoneEnCeroNull("INFOFAB", "CONSUMO", REGAF&, "NOMESS")

        SQLX$ = "SELECT  i.CodXElem, I.numeoper , I.DescOper ,I.cantreal AS CantCortes,FACTORCORTE"
        SQLX$ = SQLX$ + ",I.CORTE,I.IdenLote,I.Peso,c.CODRANGO,"
        SQLX$ = SQLX$ + " SUM(CASE WHEN cantreal > 0 THEN cantreal  ELSE -cantsalid END) AS SALDO"
        
        SQLX$ = SQLX$ + " FROM infofab I  WITH(NOLOCK)"
        SQLX$ = SQLX$ + "inner join CODIGORANGO  c on  I.CodiElem = c.CIMATPRIMA"
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2"
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE and I.INFOFAB_ID = I2.OperFucionada)"
        SQLX$ = SQLX$ + " AND idcoordubic = 100"
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " group by i.CodXElem, I.numeoper , I.DescOper ,I.cantreal ,i.FACTORCORTE"
        SQLX$ = SQLX$ + ",I.CORTE,I.IdenLote,I.Peso,c.CODRANGO"
        SQLX$ = SQLX$ + " order by I.DescOper,saldo,idenlote"
        REG& = 0

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            REG& = REG& + 1
            MSF2.Rows = REG& + 1
            MSF2.TextMatrix(REG&, 1) = SAFETEXT$(!CodXElem)
            MSF2.TextMatrix(REG&, 2) = XVALO(!NUMEOPER)
            MSF2.TextMatrix(REG&, 3) = SAFETEXT$(!descoper)
            MSF2.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CantCortes), "F9")
            MSF2.TextMatrix(REG&, 5) = XVALO(!FACTORCORTE)
            MSF2.TextMatrix(REG&, 6) = XVALO(!corte)
            MSF2.TextMatrix(REG&, 7) = SAFETEXT$(!idenlote)
            MSF2.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(!Peso), "F9.1")
            MSF2.TextMatrix(REG&, 9) = SAFETEXT$(!CODRANGO)
            MSF2.TextMatrix(REG&, 10) = SAFETEXT$(!SALDO)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
End Sub

Sub CargarMsf3(NORFA$, CANTI_ACERRAR, CIFAB%, NROCOTI&)
        Campos$ = "Código/A16 '  1"
        Campos$ = Campos$ + ";Descripción/a48"             '  2"
        Campos$ = Campos$ + ";Componente/A24"               '  3"
        Campos$ = Campos$ + ";Cielem/A0"                    '  4"
        Campos$ = Campos$ + ";Moduxb/A0"                   '  5"
        Campos$ = Campos$ + ";CORTE/A0"                  '  6"
        Campos$ = Campos$ + ";Nec.M./A0"                       '  7"
        Campos$ = Campos$ + ";Cod.Rango/A12"                '  8"
        Campos$ = Campos$ + ";Nec.Cortes/f10.1"            '  9
        Campos$ = Campos$ + ";Disponibilidad/f14"              '  10
        Campos$ = Campos$ + ";Falta/f10"                       '  11
        
        Call CARGA_ENCABEZADO(MSF3, Campos$, Me)

        If CANTI_ACERRAR < 1 Then CANTI_ACERRAR = 1
        SQLX$ = "SELECT CC.CODRANGO, C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE, SUM(C.MODUXBB * C.CORTE * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_EN_M"
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_DE_CORTES"
        
        SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " inner join CODIGORANGO  CC on C.CODINTELEM = CC.CIMATPRIMA"
        SQLX$ = SQLX$ + " Where CODINTCONJ = " & CIFAB%
        SQLX$ = SQLX$ + "  And CODINTELEM > 0 "
        SQLX$ = SQLX$ + " And NROCOTIZACION = " & NROCOTI&
        
        SQLX$ = SQLX$ + " and LEFT(COMPONENTE,8) <> 'CINTILLA'"
        SQLX$ = SQLX$ + " and LEFT(COMPONENTE,4) <> 'HILO'"

        SQLX$ = SQLX$ + " GROUP BY  CC.CODRANGO,CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        REG& = 0

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            REG& = REG& + 1
            MSF3.Rows = REG& + 1
            MSF3.TextMatrix(REG&, 1) = SAFETEXT$(!Codigo)
            MSF3.TextMatrix(REG&, 2) = SAFETEXT$(!Descripcion)
            MSF3.TextMatrix(REG&, 3) = SAFETEXT$(!COMPONENTE)
            MSF3.TextMatrix(REG&, 4) = XVALO(!CODINTELEM)
            MSF3.TextMatrix(REG&, 5) = XVALO(!moduxbb)
            MSF3.TextMatrix(REG&, 6) = XVALO(!corte)
            MSF3.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!NECESIDAD_EN_M), "F9.1")
            MSF3.TextMatrix(REG&, 8) = SAFETEXT$(!CODRANGO)
            MSF3.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(!NECESIDAD_DE_CORTES), "F9.0")
            MSF3.TextMatrix(REG&, 10) = ""
            Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbRed, REG&, 11)
            MSF3.TextMatrix(REG&, 11) = FORMATNUM$(XVALO(!NECESIDAD_DE_CORTES), "F9.0") 'POR DEFECO FALTA = NECESIDAD
                                                                        ', LUEGO SI HAY DISPONIBLIDAD SE SOBRE ESCRIBE
            
            
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
End Sub

Private Sub Text5_LostFocus()
  Text5 = FORMATNUM$(XVALO(Text5), "F10.0")
End Sub

Private Sub txtCodigo_Change()
End Sub

Private Sub TXTCODINT_Change()
   CI1% = XVALO(TXTCODINT)
   Dim CTP As New Conteplast
   NROFT& = CTP.NroFTecnicaXCodigo(CI1%)
   TXTFTECNICA = NROFT&
   Text11 = CODEXT(CI1%)
   DESCRIT0$ (CI1%)
   CONDI$ = "CODICONJ = " & CI1%
   Text4 = DESCRIT0$(CI1%)
   TXTCODIGO = VALOBADA("FORMULAS", "CODXELEM", CONDI$, "NOMESS")
  
End Sub

Sub PRINTETIQUETACTP_DESDEOF(CI1%, NORFA$, CODIGOEDITABLE$)
    If CI1% < 1 Then
       Call COMUNI("No Hay un Código Activo")
       Exit Sub
    End If
    '
    NORFA$ = TRIM$(NORFA$)
    If NORFA$ = "" Then
       Call COMUNI("Falta Orden de Fabriación")
       Exit Sub
    End If
    '
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")
   
    '
    RUTAEXCEL$ = Control("COTICONT", "EXCEETIQ")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    '
    IMAG$ = "QR_ETIQPTERM.JPG"

    On Error Resume Next
    Kill "C:\FLEXOFT\" & IMAG$
    On Error GoTo 0
'1 razon social
'2 orden de compra
'3 codigo p.t.
'4 medidas
'5 descripcion de p.t.
'6 orden de fabricacion
    Dim CTP As New Conteplast
    NC = cliente_SEGUN_PEDIDO(XVALO(Me.TXTNROPED))
    NP& = XVALO(TXTNROPED)
    CI1% = XVALO(Me.TXTCODINT)
    RASO$ = rasocli$(NC)
    OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    CODI$ = CODIGOEDITABLE$
    DESKRI$ = TRIM$(Text4)
    NF& = XVALO(TXTFTECNICA)
    ANCHOX = CTP.ANCHOCTP(CI1%)
    LARGOX = CTP.LARGOCTP(CI1%)
    ALTOX = CTP.ALTOCTP(CI1%)
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    '
    With VENTABNEW.MSF
       .Rows = 10: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = OC$
       .TextMatrix(3, 1) = CODI$
       .TextMatrix(4, 1) = CTP.CodEspTecCliente$(CI1%)
       .TextMatrix(5, 1) = DESKRI$
       .TextMatrix(6, 1) = NORFA$
       .TextMatrix(7, 1) = ""
       .TextMatrix(8, 1) = ""
       .TextMatrix(9, 1) = ""
    End With
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "ETIQOF", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0)                  'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
End Sub
Sub CARGARMSF2BOLSAS(NORFA$)
'  Código/A20                   '  1
'  Descripción/A40              '  2
'  Lote/A16                     '  3
'  Transf.Mts./f12.1            '  4
'  CI/a0                        '  5
'  Op.N°/F6                     '  6
'  Kg.Transf/F11.2              '  7
'  Consumo Mts./F12.2           '  8
'  Saldo Mts./F12.2             '  9
    Dim CTP As New Conteplast
    Call CARGA_ENCABEZADO(MSF2, "Código/A20;Descripción/A40;Lote/A16;Transf.Mts./f12.1;CI/a0;Op.N°/F6;Kg.Transf/F11.2;Consumo Mts./F12.2;Saldo Mts./F12.2", Me)
    DEPOCORTE% = XVALO(Control("", "DEPODEST"))
    Call CTP.CargaMatTranfXOrden(MSF2, TRIM$(NORFA$), DEPOCORTE%)
    Set CTP = Nothing
    For i& = 1 To MSF2.Rows - 1
       LOTE$ = MSF2.TextMatrix(i&, 3)
       MTSTRANSFERIDOS = XVALO(MSF2.TextMatrix(i&, 4))
        CONSUMIDO = XVALO(CTP.MTSCONSUMIDOS(NORFA$, LOTE$, DEPOSITO%))
       MSF2.TextMatrix(i&, 8) = FORMATNUM$(CONSUMIDO, "F12.2")
        SALDO = MTSTRANSFERIDOS - CONSUMIDO
       MSF2.TextMatrix(i&, 9) = FORMATNUM$(SALDO, "F12.2")
    Next i&
    Set CTP = Nothing
End Sub
Sub CARGARMSF1BOLSAS(NORFA$)
'  Código/A20                   '  1
'  Descripción/A40              '  2
'  Lote/A16                     '  3
'  Transf.Mts./f12.1            '  4
'  CI/a0                        '  5
'  Op.N°/F6                     '  6
'  Kg.Transf/F11.2              '  7
'  Consumo Mts./F12.2           '  8
'  Saldo Mts./F12.2             '  9
'  Rango./a12                   '  10

    Dim CTP As New Conteplast
    Call CARGA_ENCABEZADO(MSF1, "Código/A20;Descripción/A40;Lote/A16;Transf.Mts./f12.1;CI/a0;Op.N°/F6;Kg.Transf/F11.2;Consumo Mts./F12.2;Saldo Mts./F12.2;Rango/A12", Me)
    DEPOCORTE% = XVALO(Control("", "DEPODEST"))
    Call CTP.CargaMatTranfXOrden(MSF1, TRIM$(NORFA$), DEPOCORTE%)
    Set CTP = Nothing
    For i& = 1 To MSF1.Rows - 1
       LOTE$ = MSF1.TextMatrix(i&, 3)
       MTSTRANSFERIDOS = XVALO(MSF1.TextMatrix(i&, 4))
        CONSUMIDO = XVALO(CTP.MTSCONSUMIDOS(NORFA$, LOTE$, DEPOSITO%))
       MSF1.TextMatrix(i&, 8) = FORMATNUM$(CONSUMIDO, "F12.2")
        SALDO = MTSTRANSFERIDOS - CONSUMIDO
       MSF1.TextMatrix(i&, 9) = FORMATNUM$(SALDO, "F12.2")
        CI1% = XVALO(MSF1.TextMatrix(i&, 5))
        Rango$ = CODIGORANGOEXTE$(CI1%)
       MSF1.TextMatrix(i&, 10) = Rango$
    Next i&
    Set CTP = Nothing
End Sub


