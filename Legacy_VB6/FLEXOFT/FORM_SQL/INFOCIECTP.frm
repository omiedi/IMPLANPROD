VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFOCIECTP 
   BackColor       =   &H009FBE85&
   Caption         =   "Situación de Orden de Fabricación y Cierre "
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11280
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H009FBE85&
      Caption         =   "Scrap Teórico (Solo Items Correspondientes a la Ficha Técnica)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   50
      TabIndex        =   31
      Top             =   7440
      Width           =   17775
      Begin MSFlexGridLib.MSFlexGrid MSF5 
         Height          =   1485
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   17565
         _ExtentX        =   30983
         _ExtentY        =   2619
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
            Name            =   "Roboto Mono"
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
      Caption         =   "Comparativo Ficha Técnica vs Transferido"
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
      Left            =   4365
      TabIndex        =   29
      ToolTipText     =   "Doble Clic Sobre Leyenda Para Ampliar Ancho de Grilla"
      Top             =   4080
      Width           =   13485
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2850
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   13200
         _ExtentX        =   23283
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
            Name            =   "Roboto Mono"
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
      Caption         =   "Operaciones Maestro"
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
      Index           =   0
      Left            =   50
      TabIndex        =   27
      Top             =   4080
      Width           =   4230
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2850
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
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
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF6 
         Height          =   2850
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
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
            Name            =   "Roboto Mono"
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
      Caption         =   "Transferencias Realizadas de Depósito General para la Orden De Fabricación Activa"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   17790
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2250
         Left            =   120
         TabIndex        =   13
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
            Name            =   "Roboto Mono"
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
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009FBE85&
      Caption         =   "Parte de O/F a Informar"
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
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar "
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
         Left            =   16320
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "!"
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
         Left            =   2960
         TabIndex        =   38
         ToolTipText     =   "Ordenes de Fabricación Abiertas y Cerradas"
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox TXTCANTICALCULO 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   " "
         ToolTipText     =   "Doble Click Para Cambiar Canidad de Cálculo"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox TXTPENDIENTE 
         Alignment       =   1  'Right Justify
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   " "
         Top             =   720
         Width           =   1155
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   22
         Text            =   " "
         Top             =   720
         Width           =   1920
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
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
         Left            =   14310
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   240
         Width           =   720
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
         Height          =   480
         Left            =   15525
         Picture         =   "INFOCIECTP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   660
         Width           =   600
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
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. de Cálculo:"
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
         Left            =   10680
         TabIndex        =   37
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   " Pendiente O.F.:"
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
         Index           =   2
         Left            =   7560
         TabIndex        =   35
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transf. Desde Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   26
         Top             =   700
         Width           =   1530
      End
      Begin VB.Label lbldepogeneral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4095
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4695
         TabIndex        =   24
         Top             =   720
         Width           =   2325
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Proy.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -45
         TabIndex        =   23
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.del Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   12975
         TabIndex        =   21
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label21 
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
         Left            =   12960
         TabIndex        =   19
         Top             =   240
         Width           =   3210
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11160
         TabIndex        =   18
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10920
         TabIndex        =   17
         Top             =   300
         Width           =   1125
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   435
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   6810
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFOCIECTP.frx":014A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFOCIECTP.frx":037E
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   14
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
      Begin VB.Menu mnuInCorte 
         Caption         =   "Informe Sector Corte "
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
Attribute VB_Name = "INFOCIECTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Sub Carga_Operaciones(MSF As msFlexGrid, CIFAB%)
'    SQLX$ = " SELECT * FROM SECOPER where codIconj = " & CIFAB%
'    Set rs = READSET(SQLX$)
    Campos$ = "Cód.Operaciones/F10;" '1
    Campos$ = Campos$ + "Descripción/A64;"      '2
    
    Call CARGA_ENCABEZADO(MSF6, Campos$, Me)
    SQLX$ = "SELECT * FROM CODIGORANGO"
    SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "NUMEOPER ,DESCOPER"
    TABLA$ = "SECOPER"
    CONDI$ = "codIconj = " & CIFAB%
    '
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, Me.MSF6, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0

End Sub

Sub CARGAGRILLA3(NORFA$, CantAFabricar&)
'  "CODIGO/A14                  '  1
'  Descripción/A24              '  2
'  LOTE/A0                    '  3
'  Transf.Mts./f10.2            '  4
'  CI/A0                        '  5
'  Op.N°/F6                     '  6
'  Ref.Oper./A16                '  7
'  Kg.Transf/A0              '  8
'  Rango/A5                     '  9
'  Descrip.Rango/A20            '  10

'MSF3
'  Código/A12                   '  1
'  Descripción/A28              '  2
'  U./A2                        '  3
'  Necesidad/F10.2              '  4
'  Cant-Trans./F10.2            '  5
'  Falta/F9.2                   '  6
'  Nec.Corte/F9.2               '  7
'  Informados/F9.2              '  8
'  Falta/F9.1                   '  9
'  N.O./F4                      '  10
'  Operación/A16                '  11
'  ULTFILA/F0"                  '  12

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

Private Sub Command14_Click()
   Label4 = OrdenFabricacionAbiertas$(1) 'abiertas y cerradas
End Sub

Private Sub Command2_Click()
'ESTE CODIGO ESTA DENTRO DE LA RUTINA ORFCIE, ESTA COMENTADO,
'       CODFAMI$ = TRIM$(CODFAMIL$(CI2%))
'       If InStr(CODFAMI$, "CC") > 0 Then
'         DEPOCINTILLA% = XVALO(CONTROL("CONTEPLA", "DEPOCINT"))
'         Call RecuperarLoteCintilla(0, NORFA$, DEPOCINTILLA%)
'       ElseIf InStr(CODFAMI$, "BO") > 0 Then
'         DEPOMATBOL% = XVALO(CONTROL("CONTEPLA", "DEMATBOL"))
'         Call RecuperarLoteCintilla(0, NORFA$, DEPOMATBOL%)
'       End If
'     End If

End Sub

Private Sub Command3_Click()
    '
    Label4 = OrdenFabricacionAbiertas$
    
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
    Call CIERRARCH("*.*")
    Unload Me
End Sub




Private Sub Form_Load()
    '
    Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
    Call CARGA_ENCABEZADO(MSF2, "CODIGO/A16;Descripción/A28;LOTE/A0;Transf.Mts./f10.2;CI/A0;Op.N°/F6;Ref.Oper./A20;Kg.Transf/A0;Rango/A12;Descrip.Rango/A40", Me)
    '  "CODIGO/A14                  '  1
    '  Descripción/A24              '  2
    '  LOTE/A0                      '  3
    '  Transf.Mts./f10.2            '  4
    '  CI/A0                        '  5
    '  Op.N°/F6                     '  6
    '  Ref.Oper./A16                '  7
    '  Kg.Transf/A0                 '  8
    '  Rango/A5                     '  9
    '  Descrip.Rango/A20            '  10
    Call CARGA_ENCABEZADO(Me.MSF3, "Código/A12;Descripción/A28;U./A2;Necesidad/F10.2;Cant-Trans./F10.2;Falta/F9.2;Nec.Corte/F9.2;Informados/F9.2;Falta/F9.1;Op.N°/F6;Ref.Oper./A20;ULTFILA/F0", Me)
    '  Operación/A16                '  11
'  ULTFILA/F0"                  '  12

    Call CARGA_ENCABEZADO(Me.MSF1, "Código/A12;Descripción/A28;U./A2;Cantidad/F12", Me)
    Text4 = FECHATEX$(HOY(HO$))
    lbldepogeneral = Control$("", "DEPOCEN")
    
'********************************************************************************
  Campos$ = "Código Rango/A20"                              '  1
  Campos$ = Campos$ + ";Descripción/A40"                    '  2
  Campos$ = Campos$ + ";Cant.Excedente.Mts./f24.2"          '  3
  Call CARGA_ENCABEZADO(MSF5, Campos$, Me)
'********************************************************************************
End Sub


Private Sub Frame1_DblClick(Index As Integer)
   If Frame1(1).Width < 13600 Then
      Frame1(1).Width = 17820
      Frame1(1).Left = 50
      MSF3.Width = Frame1(1).Width - 285
   Else
         Frame1(1).Width = 13485
         Frame1(1).Left = 4365
         MSF3.Width = 13200
   End If
End Sub

Private Sub Label4_Change()
'   MSF1.Rows = 1
   MSF2.Rows = 1
   MSF5.Rows = 1
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then MSF1.Rows = 1: Exit Sub
   Label2 = RESP_ZELE_VECT(2)
   CI1% = CODIFAB%(NORFA$)
   TXTCODINT = CI1%
   Text10 = CODEXT$(CI1%)
   Label3 = DESCRIT0$(CI1%)
   CIFAB% = CODINT%(Text10)
   NROPED& = NroPedidoSegunOf&(NORFA$)
   NC = cliente_SEGUN_PEDIDO(NROPED&)
   Text6 = NC
   Text12 = NROPED&
   CANTI = CantiProyectadaOF(NORFA$)
   Text14 = CantiProyectadaOF(NORFA$)
   CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CIFAB%
   Text13 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
   Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$)
   Dim ctp As New Conteplast
   DEPOCORTE% = XVALO(Control("", "DEPODEST"))
   Call ctp.CargaMatTranfXOrdenParaCierre(MSF2, TRIM$(NORFA$), DEPOCORTE%, 0, 1)
   MSF2.COL = 9 'RANGO
   MSF2.Sort = 3
   Set ctp = Nothing

'   Call cargarTransfDeOf_DeSalidasDeposito(NORFA$, XVALO(Me.lbldepogeneral), MSF2)
   Call CARGA_ESTRUCTURA_FICHA(CIFAB%, CANTI, MSF1)
   Call CARGA_INFORMECONSUMO
   Call Carga_Operaciones(MSF6, CIFAB%)
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

Private Sub lbldepogeneral_Change()
   DEPO% = XVALO(lbldepogeneral)
   Label1 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))

End Sub

Private Sub mnuInCorte_Click()
   InfSecCorte.Show 1
End Sub

Private Sub mnuinformarRetiCorte_Click()
    NORFA$ = TRIM$(Label4)
    If NORFA$ = "" Then Exit Sub

    INFOCORTE.Label4 = NORFA$

    INFOCORTE.Show 1
End Sub

Private Sub mnuIngresoMatCorte_Click()
    NORFA$ = TRIM$(Label4)
    If NORFA$ = "" Then Exit Sub

    INGREBOB.TXTOF = NORFA$
    INGREBOB.Show 1
End Sub

Private Sub mnuOrigScrap_Click()
   FRMSCRAP.Show 1
End Sub

Private Sub mnuSelprinot_Click()
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then Exit Sub
   CI1% = CODIFAB%(NORFA$)

   Call OFABRI07.PresentaSeleccionImpresion(CI1%, NORFA$)
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

Private Sub mnuTablaConfScrap_Click()
        Call mostrarTablaScrap
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

Sub cargarTransfDeOf_DeSalidasDeposito(NORFA$, DEPO%, MSF As msFlexGrid)


    REG& = 0
    SQLX$ = "Select sum(cantsalid - cantingre) as CANTI,idenlote,cod_exte, DepoMovi,CODIMOVI,DESCRIP,ID_OPERACION"
    SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) "
    SQLX$ = SQLX$ + " where doseclar = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
    SQLX$ = SQLX$ + " AND CODIMOVI = 'TR'"
    SQLX$ = SQLX$ + " group by idenlote,cod_exte ,DepoMovi, CODIMOVI,DESCRIP,ID_OPERACION"
    SQLX$ = SQLX$ + " order by idenlote"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         REG& = REG& + 1
         MSF.Rows = REG& + 1
         MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
         MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
         MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
         MSF.TextMatrix(REG&, 4) = XVALO(!CANTI)
         MSF.TextMatrix(REG&, 5) = XVALO(!CANTI)
         MSF.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
         MSF.TextMatrix(REG&, 7) = SAFETEXT$(!ID_OPERACION)
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub

Private Sub TXTCANTICALCULO_DblClick()
      CANTIX$ = TRIM$(InputBox$("Informar", "Ingrese Código", TXTCANTICALCULO))
      If CANTIX$ = "" Then
         TXTCANTICALCULO = Text14
      Else
         TXTCANTICALCULO = ROUND(XVALO(CANTIX$))
      
      End If
      CANTI& = XVALO(TXTCANTICALCULO)
      Call CARGAGRILLA3(Label4, CANTI&)
End Sub

