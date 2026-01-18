VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFUBIPT 
   BackColor       =   &H009FBE85&
   Caption         =   "Ubicaciones de Productos Terminados"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   15285
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   7920
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      Caption         =   "Informar Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin VB.TextBox TXTUBICSELEC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14520
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   " "
         ToolTipText     =   "ID (Tabla Movisto)"
         Top             =   960
         Width           =   585
      End
      Begin VB.TextBox TXTTOTALCIERRES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   " "
         Top             =   9240
         Width           =   1440
      End
      Begin VB.TextBox TXTPENDIENTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   " "
         Top             =   9240
         Width           =   1440
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10420
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   " "
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   " "
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   " "
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   " "
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   240
         Width           =   1440
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   720
      End
      Begin VB.TextBox lblidregistro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         ToolTipText     =   "ID (Tabla Movisto)"
         Top             =   960
         Width           =   945
      End
      Begin VB.TextBox txtMarcarOcupadas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1050
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "INFUBIPT.frx":0000
         Top             =   4680
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   240
         Width           =   945
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
         Height          =   670
         Left            =   14475
         Picture         =   "INFUBIPT.frx":0043
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir"
         Top             =   240
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
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTCANTICORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   " "
         Top             =   600
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "O.K."
         Top             =   960
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   5130
         Left            =   3960
         TabIndex        =   10
         ToolTipText     =   "Doble Clic para Ubicar Lote o Quitar Lote-"
         Top             =   1320
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   9049
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
         Height          =   2925
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Clic para Filtrar por Rack"
         Top             =   1320
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   5159
         _Version        =   393216
         Cols            =   4
         BackColor       =   8421504
         ForeColor       =   16777152
         BackColorFixed  =   16501405
         SelectionMode   =   1
         AllowUserResizing=   1
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
         Height          =   2250
         Left            =   3960
         TabIndex        =   28
         Top             =   6720
         Width           =   11145
         _ExtentX        =   19659
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cierres Realizados:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   39
         Top             =   9240
         Width           =   2565
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad.Pendiente de Cierre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10560
         TabIndex        =   37
         Top             =   9240
         Width           =   3045
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Original Pedido.:"
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
         Left            =   7800
         TabIndex        =   35
         Top             =   1020
         Width           =   2565
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Proy.O.F.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4680
         TabIndex        =   33
         Top             =   1020
         Width           =   1485
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7800
         TabIndex        =   31
         Top             =   600
         Width           =   4050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   4560
         TabIndex        =   30
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ord. N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5040
         TabIndex        =   24
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7800
         TabIndex        =   22
         Top             =   240
         Width           =   4050
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Id Registro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   12120
         TabIndex        =   19
         Top             =   1035
         Width           =   1335
      End
      Begin VB.Label LBLNUMEORDENINFO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   240
         TabIndex        =   16
         Top             =   4200
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro x Rack"
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Entrega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11880
         TabIndex        =   12
         Top             =   315
         Width           =   1485
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   8
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. de Cierre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   11760
         TabIndex        =   5
         Top             =   675
         Width           =   1815
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFUBIPT.frx":018D
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFUBIPT.frx":03C1
      TabIndex        =   2
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuCtaListado 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver "
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuOcupada 
         Caption         =   "Marcar/Desmarcar Ubicaciones ocupadas"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTabAux 
         Caption         =   "Tablas Auxiliares"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "INFUBIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NOUBICAR%




Sub BLANFORMU()


    MSF3.Rows = 1
    Text12 = ""
    Text1 = ""
    Text6 = ""
    Label21 = ""
    Text4 = ""
    TXTCANTICORTES = ""
    lblidregistro = ""
    TXTUBICSELEC = ""
End Sub

Sub CARGARGRILARACK()
   REG& = 0
   SQLX$ = "SELECT  CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'CONFUBPT' ORDER BY CAMPO1"
   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
       REG& = REG& + 1
       MSF2.Rows = REG& + 1
       MSF2.TextMatrix(REG&, 1) = XVALO(!CAMPO1)
       MSF2.TextMatrix(REG&, 2) = SAFETEXT$(!campo2)
       .MoveNext
      Loop
   End With
End Sub



Sub DarPorEntregado(idRegistro)
    If idRegistro < 1 Then Exit Sub
    Erase RESP_ZELE_VECT
    Campos$ = "Nro Remito/A18;Fecha/D8"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "Doprilar,Fechmov"
    TABLA$ = "MOVISTO "
    CONDI$ = "IDENLOTE= '" & Label4 & "' AND CODIMOVI = 'RT' GROUP BY  Doprilar,Fechmov ORDER BY FECHMOV "
    
  
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    If RESP_ZELE_VECT(1) <> "" Then '
       Remito$ = TRIM$(RESP_ZELE_VECT(1))
       Call ACTUREGISTRO("MOVISTO", "DocuOrig", "MOVISTO_ID  = " & idRegistro, Remito$, REGAF&, "NOMESS")
       Label4_Change
    End If
End Sub

Sub GRABARUBICA(DESCRACK$, RK, FL, KL, NORFA$, CANTI, NC, NPED&, ORDENPED, FECHENT%, Optional LblId1&)
'    NPED& = NroPedidoSegunOf&(NORFA$)
'    NC = XVALO(VALOBADA("PEDENCA", "NroClie", "NROPED = " & NPED&, "NOMESS"))
    rs$ = rasocli$(NC)
    CIFAB% = CODIFAB%(NORFA$)
    CONDI$ = "RACK = " & RK
    CONDI$ = CONDI$ + " AND FILA = " & FL
    CONDI$ = CONDI$ + " AND COLUMNA = " & KL
    EXISTEFILA% = CantRegistros&("UBILOTE_ENVAP", CONDI$, "NOMESS")
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    '
    If EXISTEFILA% > 0 Then
       SQLX$ = "UPDATE UBILOTE_ENVAP SET LOTE = '" & NORFA$ & "'"
       SQLX$ = SQLX$ + ", CANTIDAD = " & CANTI
       SQLX$ = SQLX$ + ", MARBORRA = 0"
       SQLX$ = SQLX$ + ", CODINT = " & CIFAB%
       SQLX$ = SQLX$ + ", NUCLIEN = " & NC
       SQLX$ = SQLX$ + ", FECHEMISI = '" & FETEXCOR1$(HOY(HO$)) & "'"
       SQLX$ = SQLX$ + ", IDSUBOR =   '" & NORFA$ & "'"
       SQLX$ = SQLX$ + ", RASO_CLI = '" & rs$ & "'"
       SQLX$ = SQLX$ + ", DescOper = ''"
       SQLX$ = SQLX$ + ", NumeOper = 0"
       SQLX$ = SQLX$ + ", NROPED = " & NPED&
       SQLX$ = SQLX$ + ", CodiEmpr = " & CodiEmp%
       SQLX$ = SQLX$ + ", DescriRack = '" & DESCRACK$ & "'"
       SQLX$ = SQLX$ + ", FESOLENT = '" & FECHATEX(FECHENT%) & "'"
       SQLX$ = SQLX$ + ", ID_MOVISTO = " & LblId1&
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
    Else
       SQLX$ = "INSERT INTO UBILOTE_ENVAP "
       SQLX$ = SQLX$ + " (LOTE,CANTIDAD,MARBORRA,CODINT,NUCLIEN,FECHEMISI"
       SQLX$ = SQLX$ + " ,IDSUBOR,RASO_CLI,DescOper,NumeOper,NROPED,CodiEmpr,RACK,FILA,COLUMNA,DescriRack,FESOLENT,ID_MOVISTO)"
       SQLX$ = SQLX$ + " VALUES('" & NORFA & "'"
       SQLX$ = SQLX$ + " ," & CANTI '+ AGREGACANT
       SQLX$ = SQLX$ + " ,0"
       SQLX$ = SQLX$ + " ," & CIFAB%
       SQLX$ = SQLX$ + " ," & NC
       SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'"
       SQLX$ = SQLX$ + " ,'" & NORFA$ & "'"
       SQLX$ = SQLX$ + " ,'" & rs$ & "'"
       SQLX$ = SQLX$ + " ,''"
       SQLX$ = SQLX$ + " , 0"
       SQLX$ = SQLX$ + " ," & NPED&
       SQLX$ = SQLX$ + " ," & CodiEmp%
       SQLX$ = SQLX$ + " ," & RK
       SQLX$ = SQLX$ + " ," & FL
       SQLX$ = SQLX$ + " ," & KL
       SQLX$ = SQLX$ + " ,'" & DESCRACK$ & "'"
       SQLX$ = SQLX$ + ",'" & FECHATEX(FECHENT%) & "'"
       SQLX$ = SQLX$ + " ," & LblId1&
       SQLX$ = SQLX$ + ")"
       FLEXCONN.Execute (SQLX$)
    End If
    '
'    'Guarda las Cordenadas de ubicacion in infofab
'    CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
'    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
'    CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
'    'Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
'    CONDI$ = CONDI$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
'    CONDI$ = CONDI$ + " AND  NUORINF = " & XVALO(LBLNUMEORDENINFO)
'    CONDI2$ = CONDI$ + " AND COORDUBIC = '" & CoordenadasUbicacion$ & "'"
'    'ESTO SE AGREGO DESPUES SI SE LEVANTA DE LA PANTALLA DE INFOCORTE VIENE CON EL ID PARA ASEGURAR LA TRANSACCION
'    If LblId1& > 0 Then
'       CONDI$ = "INFOFAB_ID =" & LblId1&
'    End If
'
'    If CantRegistros("INFOFAB", CONDI2$, "NOMESS") > 0 Then
'         Call COMUNI("Atención !!! Imposible Ingresar Misma Operación y Lote en en Misma Ubicación\Libere la Ubicación y Vuelva a Intentar o Seleccione una Nueva Ubicación")
'       Exit Sub
'    End If
'    Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, CoordenadasUbicacion$, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, 0, REGAF&, "NOMESS")


'    Call ACTUREGISTRO("INFOFAB", "IDENLOTE", Condi$, LOTE$, REGAF&, "NOMESS")
    
ERROR_GUARDAR:
    If Err <> 0 Then
       On Error Resume Next
       On Error GoTo 0
       AAA$ = Err.Description
       Call FRATEXTO(AAA$, 50)
       TEXT$ = ""
       For KU& = 1 To CARETEX%
         TEXT$ = TEXT$ + RETEX$(KU&) + "\"
       Next KU&

       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION\" & TEXT$)
       FLEXCONN.RollbackTrans
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
    End If

End Sub

Sub MARCAROCUPADAS(RK, FL, KL)
    CONDI$ = "RACK = " & RK
    CONDI$ = CONDI$ + " AND FILA = " & FL
    CONDI$ = CONDI$ + " AND COLUMNA = " & KL
    EXISTEFILA% = CantRegistros&("UBILOTE_ENVAP", CONDI$, "NOMESS")
    If XVALO(MSF1.TextMatrix(FL, 5)) < 1 Then
    
    
    End If

    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    '
    If EXISTEFILA% > 0 And AGREGACANT < 1 Then
       SQLX$ = "UPDATE UBILOTE_ENVAP SET LOTE = 'OCUPADO'"
       SQLX$ = SQLX$ + ", IDSUBOR =   'OCUPADO'"
       SQLX$ = SQLX$ + ", CodiEmpr = " & CodiEmp%
       SQLX$ = SQLX$ + ", DescriRack = '" & DESCRACK$ & "'"
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
    Else
       SQLX$ = "INSERT INTO UBILOTE_ENVAP "
       SQLX$ = SQLX$ + " (LOTE,CodiEmpr,IDSUBOR,DescriRack"
       SQLX$ = SQLX$ + " VALUES('" & LOTE$ & "'"
       SQLX$ = SQLX$ + " ," & CodiEmp%
       SQLX$ = SQLX$ + " ,'OCUPADO'"

       SQLX$ = SQLX$ + ")"
       FLEXCONN.Execute (SQLX$)
    End If
    
ERROR_GUARDAR:
    If Err <> 0 Then
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       FLEXCONN.RollbackTrans
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
    End If
    
    
End Sub

Private Sub Command3_Click()
    '
    Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
    If VALACT1$("ECORDEF") <> "" Then
     Label4 = VALACT1$("ECORDEF")
    End If
    '
End Sub

Private Sub Command4_Click(Index As Integer)
  
End Sub '

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
End Sub

Private Sub Command6_Click()
    '
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
        If TRIM$(Label4) = "" Then
            Call MENSERR(24, "Primero debe Seleccionar la Orden de Fabricacion")
            Exit Sub
        End If
    End If
    '
    CONDI$ = "CodIConj = " & CIXI%
    REG& = CantRegistros&("SECOPER", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Operaciones de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!OPERINF1/Operaciones de Ordenes de Fabricacion")
    If TRIM$(VALACT1$("!OPERINF1")) <> "" Then
      Text3.TEXT = TRIM$(VALACT1$("!OPERINF1"))
'      Label7.Caption = TRIM$(VALACT2$("!OPERINF1"))
    End If
    '
End Sub

Private Sub Command7_Click()
   '

   '
End Sub

Private Sub Command8_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub



Private Sub Command9_Click()
   
End Sub

Private Sub Form_Load()
    '
    Call CREATABLA_UBILOTE_ENVAP
    Campos$ = "Rack/F3;" '1
    Campos$ = Campos$ + "Descripción/A24;" '2
    Campos$ = Campos$ + "Piso/F4;" '3
    Campos$ = Campos$ + "Picking/F7;" '4
    Campos$ = Campos$ + "Cant./F5" '5
    Campos$ = Campos$ + ";O.F./A12;" '6
    Campos$ = Campos$ + "MOVISTO_ID/A10;" '7
    Campos$ = Campos$ + "NC/A0;" '8
    Campos$ = Campos$ + "Razon Social/A24;" '9
    Campos$ = Campos$ + "N°Pedido/f7;" '10
    Campos$ = Campos$ + "F.Entrega/D8;" '11
    Campos$ = Campos$ + "UBILOTE_ID/A0" '11
    
    Call CARGA_ENCABEZADO(MSF1, Campos$, Me)

    Campos$ = "CI/A0;" '1
    Campos$ = Campos$ + "Código/A16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "F.Emis./D8;" '4
    Campos$ = Campos$ + "Cantidad./F9.1" '5
    Campos$ = Campos$ + ";Movisto_id/A0;" '6
    Campos$ = Campos$ + "Ubicación/A12;" '7
    Campos$ = Campos$ + "Cant.Pq./f8;" '8
    Campos$ = Campos$ + "Remito/A18;" '9
    Campos$ = Campos$ + "Nada/A0;" '10
    
    Call CARGA_ENCABEZADO(MSF3, Campos$, Me)

    MSF1.FixedCols = 4
    MSF1.ScrollTrack = True
    MSF1.ScrollBars = flexScrollBarBoth
    MSF1.MergeCells = flexMergeFree
    MSF1.MergeCol(1) = True
    MSF1.MergeCol(2) = True
    MSF1.MergeCol(3) = True
    Call ARMARGRILLAUBICA(NRACK%)
    Call CARGA_ENCABEZADO(MSF2, "Rack/F4;Descripción/A28;Nada/A0", Me)
    Call CARGARGRILARACK
    
    
End Sub

Sub CargaMatTranfXOrdenSoloIngresos()
 '********************************************************************
 'VER ACA SI ESTA BIEN QUE MUESTRE SOLO CANTINGRE EN VEZ DE CANTINGRE - CANTSALID
  '********************************************************************

  MSF1.Rows = 1: REG& = 0
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  
  SQLX$ = "SELECT sum(CANTINGRE) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION "
  SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where doprilar = 'TR." & TRIM(Label4) & "'  AND DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
  SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!Resultado), "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
 
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   If Me.Height > 10500 Then
       Frame1.Width = Me.Width - 200
       Frame1.Height = Me.Height - 800
       Me.MSF1.Width = Frame1.Width - 4060
       Me.MSF1.Height = Me.Height - 3500 - MSF3.Height
       
       Me.MSF3.Width = Frame1.Width - 4060
       Me.MSF3.Height = Me.Height - 2900 - MSF1.Height
       Me.MSF3.Top = MSF1.Height + MSF1.Top + 200
    '   Frame1.Top = Me.Height - 1300
       Me.Label10.Top = Me.MSF3.Height + Me.MSF3.Top + 50
       Me.Label9.Top = Me.MSF3.Height + Me.MSF3.Top + 50
       Me.TXTPENDIENTE.Top = Me.MSF3.Height + Me.MSF3.Top + 50
       Me.TXTTOTALCIERRES.Top = Me.MSF3.Height + Me.MSF3.Top + 50
   Else
        Frame1.Height = 9645
        Frame1.Width = 15195
        
        MSF1.Height = 5130
        MSF1.Width = 11145
        
        MSF3.Height = 2250
        MSF3.Width = 11145
        MSF3.Top = 6720
        
       Me.Label10.Top = 9240
       Me.Label9.Top = 9240
       Me.TXTPENDIENTE.Top = 9240
       Me.TXTTOTALCIERRES.Top = 9240
        
   End If
   
   On Error GoTo 0
End Sub

Private Sub Label3_Change()
   Call ARMARGRILLAUBICA(XVALO(Label3))
End Sub

Private Sub Label3_DblClick()
  Label3 = ""
End Sub

Private Sub Label4_Change()
    J& = 0
    MSF3.Rows = 1
    NORFA$ = TRIM$(Label4)
    Call BLANFORMU
    If NORFA$ = "" Then Exit Sub
    Dim CTP As New Conteplast
    CI1% = CODIFAB%(NORFA$)
    Text2 = CODEXT$(CI1%)
    Label6 = DESCRIT0$(CI1%)
    NROPED& = NroPedidoSegunOf&(NORFA$)
    Text12 = NROPED&
    NC = cliente_SEGUN_PEDIDO(NROPED&)
    Text6 = NC
    Text1 = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
    Text3.TEXT = CStr(CantiProyectadaOF(NORFA$))
    CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CI1%
    Text5 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
    Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CI1%, NORFA$)
    TXTTOTALCIERRES = CantiAprobadaOF(NORFA$)
    Text4 = FECHATEX$(CTP.FechaEntregaSegunNroPedyOrdenPedido(NROPED&, XVALO(Text1)))
    PENDIENTE = XVALO(Text3) - XVALO(TXTTOTALCIERRES)
    If PENDIENTE < 0 Then PENDIENTE = 0
    Me.TXTPENDIENTE = PENDIENTE
    EPAC$ = UCase$(TRIM$(EMPREAC$))

   
    If NORFA$ = "" Then Exit Sub
    
    SQLX$ = "SELECT  COD_INTE,COD_EXTE,DESCRIP,FECHMOV, CANTINGRE,MOVISTO_ID,IdCalidad,DOCUORIG,IDENLOTE,CANTSALID "
    SQLX$ = SQLX$ + " FROM MOVISTO  WITH(NOLOCK)  "
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'OF' AND CANTINGRE > 0" 'PARA QUE NO TRAIGA LOS CIERRES ANULADOS
    SQLX$ = SQLX$ + " AND DOPRILAR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND (DOCUORIG = '' OR  DOCUORIG IS NULL)" 'PARA QUE NO TRAIGA LOS REMITIDOS
    SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID DESC"
    
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        LOTE$ = SAFETEXT$(!idenlote)
        CI1% = XVALO(!cod_inte)
        If ESTRAZABLE%(CI1%) = 1 Then
          If SaldoLote(CI1%, LOTE$) > 0.5 Then 'PARA UNIFICAR POR LOS QUE SE CARGAN EN EL REMITO SIN IDMOVISTO Y NO QUEDA
             GoTo 60 'AGREGA REGISTROS
          Else
             GoTo 80 'NEXT
          End If
        Else
60          J& = J& + 1                    'RESUELTO DE GRABAR EL NRO DE REMITO EN EL DOCUORIG
            MSF3.Rows = J& + 1
            MSF3.TextMatrix(J&, 1) = CI1%
            MSF3.TextMatrix(J&, 2) = SAFETEXT$(!Cod_Exte)
            MSF3.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
            MSF3.TextMatrix(J&, 4) = FECHATEX$(CVINT(!FechMov))
            MSF3.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!CANTINGRE), "F9.1")
            MSF3.TextMatrix(J&, 6) = XVALO(!MOVISTO_ID)
            MSF3.TextMatrix(J&, 7) = CTP.CordenadasUbicaciónSegunMovistoid(XVALO(!MOVISTO_ID))
            MSF3.TextMatrix(J&, 8) = XVALO(!IdCalidad)
            MSF3.TextMatrix(J&, 9) = SAFETEXT$(!DOCUORIG) 'no va a estar por que si tiene remito no se va a mostrar
            MSF3.TextMatrix(J&, 10) = "" 'este esta invisible
        End If
80      .MoveNext
      Loop
    End With
    
    rs.Close
    Set rs = Nothing
    Set CTP = Nothing
End Sub

Private Sub mnuOcupada_Click()
   If txtMarcarOcupadas.Visible = True Then
      txtMarcarOcupadas.Visible = False
   Else
      txtMarcarOcupadas.Visible = True
   End If
    
End Sub

Private Sub mnuTabAux_Click()

' If DERACCE%("TABAUXGE", "ACCESO A TABLAS AUXILIARES GENERALES") < 2 Then Exit Sub
' Call MOSTRARTABLASCONFIGURADAS(Me.Name)
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



Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
    CANTI = XVALO(MSF1.TextMatrix(REG&, 5))
    LblId_MOVISTO& = XVALO(lblidregistro)
    LblId_UBILOTE& = XVALO(MSF1.TextMatrix(REG&, 12))
    
    RK = XVALO(MSF1.TextMatrix(REG&, 1))
    DRK$ = MSF1.TextMatrix(REG&, 2)
    FL = XVALO(MSF1.TextMatrix(REG&, 3))
    KL = XVALO(MSF1.TextMatrix(REG&, 4))
    NORFA$ = TRIM$(Label4)
    If CANTI <= 0 And LblId_MOVISTO& < 1 Then Exit Sub
    
    If CANTI > 0 Then
      TXXK$ = "RACK: " & RK & " (" & DRK$ & ")"
      TXXK$ = TXXK$ + " -Piso: " & FL
      TXXK$ = TXXK$ + " -Picking: " & KL
      OPX% = COMALTER%("" & TXXK$ & "\\Ubicación Ocupada!!\Quitar Pallet de Producto Terminado de la Ubicación Seleccionada")
      If OPX% = 2 Then
        Call ELIMINAR_CONTENIDO_UBICACION(LblId_UBILOTE&)
        Call ARMARGRILLAUBICA(XVALO(Label3))
      End If
    Else
      If XVALO(TXTUBICSELEC) > 0 Then
        On Error Resume Next
        textox$ = MSF2.TextMatrix(REG&, 10)
        strArray = Split(textox$, ";")
        RK = XVALO(strArray(0))
        FL = XVALO(strArray(1))
        KL = XVALO(strArray(2))
        On Error GoTo 0

        UBIX$ = "Rack: " & RK & " Piso: " & FL & " Picking: " & KL
        Call COMUNI("Cierre Seleccionado Tiene Ubicación (Coordenadas: " & UBIX$ & ")\Debe Liberar La Ubicación Asingada y Luego Asignar una Nueva")
        Exit Sub
      End If
      
      Call GRABARUBICA(DRK$, RK, FL, KL, NORFA$, XVALO(TXTCANTICORTES), XVALO(Text6), XVALO(Text12), XVALO(Text1), FECHANUM(Text4), XVALO(lblidregistro))
    End If
    
    Call MostrarUbicacionesGuardadas 'refresca la info guardada en bd.
    Call Label4_Change

'
'    End If
'        Campos$ = "CI/A0;" '1
'    Campos$ = Campos$ + "Código/A16;" '2
'    Campos$ = Campos$ + "Descripción/A32;" '3
'    Campos$ = Campos$ + "F.Emis./D8;" '4
'    Campos$ = Campos$ + "Cantidad./F9.1" '5
'    Campos$ = Campos$ + ";Movisto_id/A0;" '6
'    Campos$ = Campos$ + "Nada/A0;" '7
'
'    If LblId_MOVISTO& > 0 Then
'      TXXK$ = "RACK: " & RK & " (" & DRK$ & ")"
'      TXXK$ = TXXK$ + " -Piso: " & FL
'      TXXK$ = TXXK$ + " -COLUMNA: " & KL
'      OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Quitar Pallet de Producto Terminado de la Ubicación Seleccionada")
'      If OPX% = 2 Then
'        Call ELIMINAR_CONTENIDO_UBICACION(LblId_UBILOTE&)
'        Exit Sub
'      End If
'    End If
'    '
'    TXTCANTICORTES = FORMATNUM$(CANTI, "f9.1")
'    Label4 = NORFA$
'
'
'
'
'
'''    OPCION DE INGRESAR LOTE A UBICACION SELECCIONADA.
'    If OPX% = 3 Then
''       If NOUBICAR% > 0 Then Call COMUNI("Imposible Ubicar Pallet!\Debe Ingresarlo Mediante la Selección de un informe Activo")
'       If CANTI > 0 Then
'          OPX% = COMALTER("Se Detecta un lote en Ubicación Seleccionada, (Lote: " & LOTE$ & ")\\Cancelar\Sumar Lote en Posición Seleccionada")
'          If OPX% < 2 Then Exit Sub
'          If OPX% = 2 Then
'              Call GRABARUBICA(DRK$, RK, FL, KL, NORFA$, CANTI, NC, NROPED&, ORDPEDIDO&, FECHENT%, XVALO(lblidregistro))
'          End If
'       Else
'              Call GRABARUBICA(DRK$, RK, FL, KL, NORFA$, CANTI, NC, NROPED&, ORDPEDIDO&, FECHENT%, XVALO(lblidregistro))
'       End If
'    End If
'
'60
'
'    Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(MSF1, 5) 'limpia los registros solo borra el contenido a paritr de la columna 5
'    Call MostrarUbicacionesGuardadas 'refresca la info guardada en bd.

End Sub

Private Sub Text19_Change()
    On Error Resume Next
    TXT$ = TRIM$(ECOARCH$("TAMOTIV", Chr$(Text19)))
    On Error GoTo 0
    If TXT$ <> "" Then
       Label55 = TXT$
    Else
       Label55 = ""
    End If
End Sub

Private Sub MSF3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
     EPAC$ = UCase$(TRIM$(EMPREAC$))

     REG& = MSF3.MouseRow
     COL% = MSF3.MouseCol
     If REG& = 0 Or REG& >= MSF3.Rows Then Exit Sub
     Call SELECCIONARFILA(MSF3, REG&)
     Remito$ = TRIM$(MSF3.TextMatrix(REG&, 9))
'     If Remito$ <> "" Then
'        Call COMUNI("Imposible dar Ubicaciòn o Anular Cierre\El Pallet Seleccionado ha Sido Remitido")
'        Exit Sub
'     End If
     '
     IDMOVISTO& = XVALO(MSF3.TextMatrix(REG&, 6))
     NROOF$ = TRIM$(Label4)

     OPX% = COMALTER%("Opción de Ubicaciones\\Canelar\Imprimir Etiqueta de Pallet De Cierre de O.F. Seleccionada\Anular Cierre Seleccionado\Dar Por Entregado ")
     CANTI = XVALO(MSF3.TextMatrix(REG&, 5))
     
     If OPX% < 2 Then Exit Sub
     If OPX% = 2 Then
        'VALIDAR SI TIENE UBICACION SI NO PUEDE IMPRIMIR LA ETIQUETA
        UBICACION$ = TRIM$(MSF3.TextMatrix(J&, 7))
        If UBICACION$ = "" Then
          OPX% = COMALTER%("Producto sin Ubicaciòn\\Cancelar\Imprimir de Todos Modos (Solo con Derechos)")
          If OPX% < 2 Then Exit Sub
          If DERACCE%("IMPRESION DE ETIQUETA DE PALLET DE CIERRE DE O.F.", "IMPETPAL") < 2 Then
             Exit Sub 'SI NO TIENE DERECHOS SALE
          End If
        End If
        FECHA = (MSF3.TextMatrix(REG&, 4))
        CANTPAQ = XVALO(MSF3.TextMatrix(REG&, 8))
        Codigo$ = TRIM$(Text2): CI1% = CODINT%(Codigo$)
        Dim CTP As New Conteplast
45      CANTX$ = InputBox("Cantidad de Etiquetas", "Imprimir Etiquetas", CANTX$)
        If XVALO(CANTX$) > 0 Then
           CANTIMPRE% = XVALO(CANTX$)
           If CANTIMPRE% > 10 Then
              Call COMUNI("Cantidad Máxima 10 Etiquetas")
              GoTo 45
           End If
           If CANTIMPRE% < 1 Then Exit Sub
        End If
        TextoQuilmes$ = ""
        If InStr(EPAC$, "QUILMES") > 0 Then 'MODIF.25_01_24 SI SON ESTAS FAMILIAS IMPRIME LA ETIQUETA IGUAL QUE RIVADAVIA
          CODFAMI$ = TRIM$(CODFAMIL$(CI1%))
          If InStr(CODFAMI$, "R1") > 0 Or InStr(CODFAMI$, "R2") > 0 Or InStr(CODFAMI$, "TP") > 0 Or InStr(CODFAMI$, "TT") > 0 Or InStr(CODFAMI$, "PL") > 0 Or InStr(CODFAMI$, "TL") > 0 Or InStr(CODFAMI$, "EE") > 0 Or InStr(CODFAMI$, "CT") > 0 Then
                CONDI$ = "MOVISTO_ID = " & IDMOVISTO&
                LoteNorfa$ = VALOBADA("MOVISTO", "IDENLOTE", CONDI$, "NOMESS")
                pesoReal = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
                metrosReal = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
                If metrosReal < 1 Then
                   metrosReal = XVALO(PESUNI(CI1%))
                End If
                Call CIERRE_CMB.PRINTETIQORFA_CAMBACERES(NROOF$, LoteNorfa$, pesoReal, metrosReal, 1)
                Call Label4_Change 'REFRESCO
                Exit Sub
          Else 'SI NO ESTA DENTRO DE ESAS FAMILIAS IMPRIME COMO ANTES LA DE HOJA GRANDE
                TextoQuilmes$ = InputBox("Ingresar Referencia", "Referencia", TextoQuilmes$)
                Call PRINTETIQUETACTPQUILMES(CI1%, NROOF$, CODIGOEDITABLE$, CANTI, FECHA, CANTPAQ, IDMOVISTO&, CANTIMPRE%, TRIM$(TextoQuilmes$))
                Call Label4_Change 'REFRESCO
                Exit Sub
          End If
        End If
        Call CTP.PRINTETIQUETACTP(CI1%, NROOF$, CODIGOEDITABLE$, CANTI, FECHA, CANTPAQ, IDMOVISTO&, CANTIMPRE%, TRIM$(TextoQuilmes$))
        Call Label4_Change 'REFRESCO
     ElseIf OPX% = 3 Then
        If DERACCE%("ANULAROF", "Anulación de Orden de Fabricaciòn") < 2 Then
           Exit Sub
        End If

        If XVALO(MSF3.TextMatrix(REG&, 7)) > 0 Then
           Call COMUNI("Imposible Anular Cierre Seleccionado\Quite la Ubicación del Mismo y Vuelva a Intentar")
           Exit Sub
        End If
        Call AnularCierre(NROOF$, IDMOVISTO&, CANTI)
     ElseIf OPX% = 4 Then
        Call DarPorEntregado(XVALO(Me.lblidregistro))
        Exit Sub
     End If
     
  End If
End Sub
Sub PRINTETIQUETACTPQUILMES_anterior(CI1%, NORFA$, CODIGOEDITABLE$, CANTI, FECHA, CANTPAQ, Optional IDMOVISTO&, Optional CANTIMPRE%, Optional TextoQuilmes$)
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
'    Call SETPRINPORT(PDES$)
    FORIPRE$ = "ETIQ-CMB"
    '
    Dim CTP As New Conteplast
    NP& = CTP.NroPedidoSegunOf2&(NORFA$)
    NC = cliente_SEGUN_PEDIDO(XVALO(NP&))
    RASO$ = rasocli$(NC): If TRIM$(RASO$) = "" Then RASO$ = "-"
    If TextoQuilmes$ <> "" Then 'a pedido de centurion se cambia la orden de compra por un texto de referencia
       OC$ = TextoQuilmes$      'si viene vacio imprime la o.compra
    Else
       OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    End If
    CODI$ = CODEXT$(CI1%): If TRIM$(CODI$) = "" Then CODI$ = "-"
    DESKRI$ = DESCRIT0$(CI1%): If TRIM$(DESKRI$) = "" Then DESKRI$ = "-"
    NF& = FICHATECNCIA_SEGUN_CODIGO&(CI1%)
    ANCHOX = CTP.ANCHOCTP(CI1%)
    LARGOX = CTP.LARGOCTP(CI1%)
    ALTOX = CTP.ALTOCTP(CI1%)
    CODCLIENTE$ = CTP.CodEspTecCliente$(CI1%): If TRIM$(CODCLIENTE$) = "" Then CODCLIENTE$ = "-"
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    If TRIM$(NORFA$) = "" Then NORFA$ = "-"
    If TRIM$(FECHA) = "" Then FECHA = "-"
    If XVALO(CANTI) < 1 Then CANTI = "1"
    '
'    IdCalidad
    CANTIPAQUETE$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad de Paquetes", XVALO(CANTPAQ)))
    If XVALO(CANTIPAQUETE$) < 1 Then CANTIPAQUETE$ = "1"
    
'    With VENTABNEW.MSF
'       .Rows = 11: .Cols = 2
'       .TextMatrix(1, 1) = RASO$
'       .TextMatrix(2, 1) = OC$
'       .TextMatrix(3, 1) = CODI$
'       .TextMatrix(4, 1) = CODCLIENTE$
'       .TextMatrix(5, 1) = DESKRI$
'       .TextMatrix(6, 1) = NORFA$
'       .TextMatrix(7, 1) = FECHA
'       .TextMatrix(8, 1) = CANTIPAQUETE$
'       .TextMatrix(9, 1) = CANTI
'    End With

     IMAG$ = "QR_ETIQPTERM.JPG"
     Identificador$ = "QR_ETIQPTERM"

     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     TEXTOQR$ = CODI$ & "|" & NORFA$ & "|" & CANTI & "|" & NC & "|" & CI1% & "|" & NP& & "|" & IDMOVISTO&

     On Error GoTo 0
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", 0, 0))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, Identificador$)

    For i% = 1 To CANTIMPRE%
      FILTX$ = TRIM$(RASO$) & ";" & OC$ & ";" & CODI$ & ";" & CODCLIENTE$ & ";" & DESKRI$ & ";" & NORFA$ & ";" & FECHA & ";" & CANTIPAQUETE$ & ";" & TRIM$(CStr(CANTI))
     '   Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0, CANCOPIAS&)
      Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0)
    Next i%


        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
'    If CANTIMPRE% < 1 Then CANTIMPRE% = 1
'    For i% = 1 To CANTIMPRE%
'      Call Param_Print_Excel(RUTAEXCEL$, "ETIQUETAS", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0, "C:\FLEXOFT\" & IMAG$, 550, 29, 130, 130)                     'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
'    Next i%
    
'    Unload VENTABNEW
    
    'ACTUALIZO LA CANTIDAD DE PAQUETES
    Call ACTUREGISTRO("MOVISTO", "IDCALIDAD", "MOVISTO_ID = " & IDMOVISTO&, CANTIPAQUETE$, REGAF&, "NOMESS")
    
End Sub

Sub PRINTETIQUETACTPQUILMES(CI1%, NORFA$, CODIGOEDITABLE$, CANTI, FECHA, CANTPAQ, Optional IDMOVISTO&, Optional CANTIMPRE%, Optional TextoQuilmes$)
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
'    Call SETPRINPORT(PDES$)
'    '
    
    RUTAEXCEL$ = Control("COTICONT", "EXCEETIQ")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    FORIPRE$ = TRIM$(Control("ETIQPRTE", "ETPROTER"))
    If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Etiqueta de Cierre de Orden de Fabricación\Configurar Formulario desde la Ficha Técnica > Congiguración General")
       Exit Sub
    End If

    Dim CTP As New Conteplast
    NP& = CTP.NroPedidoSegunOf2&(NORFA$)
    NC = cliente_SEGUN_PEDIDO(XVALO(NP&))
    RASO$ = rasocli$(NC): If TRIM$(RASO$) = "" Then RASO$ = "-"
    If TextoQuilmes$ <> "" Then 'a pedido de centurion se cambia la orden de compra por un texto de referencia
       OC$ = TextoQuilmes$      'si viene vacio imprime la o.compra
    Else
       OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    End If
    CODI$ = CODEXT$(CI1%): If TRIM$(CODI$) = "" Then CODI$ = "-"
    DESKRI$ = DESCRIT0$(CI1%): If TRIM$(DESKRI$) = "" Then DESKRI$ = "-"
    NF& = FICHATECNCIA_SEGUN_CODIGO&(CI1%)
    ANCHOX = CTP.ANCHOCTP(CI1%)
    LARGOX = CTP.LARGOCTP(CI1%)
    ALTOX = CTP.ALTOCTP(CI1%)
    CODCLIENTE$ = CTP.CodEspTecCliente$(CI1%): If TRIM$(CODCLIENTE$) = "" Then CODCLIENTE$ = "-"
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    
    If TRIM$(NORFA$) = "" Then NORFA$ = "-"
    If TRIM$(FECHA) = "" Then FECHA = "-"
    If XVALO(CANTI) < 1 Then CANTI = "1"
    '
'    IdCalidad
    CANTIPAQUETE$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad de Paquetes", XVALO(CANTPAQ)))
    If XVALO(CANTIPAQUETE$) < 1 Then CANTIPAQUETE$ = "1"
    
    With VENTABNEW.MSF
       .Rows = 11: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = OC$
       .TextMatrix(3, 1) = CODI$
       .TextMatrix(4, 1) = CODCLIENTE$
       .TextMatrix(5, 1) = DESKRI$
       .TextMatrix(6, 1) = NORFA$
       .TextMatrix(7, 1) = FECHA
       .TextMatrix(8, 1) = CANTIPAQUETE$
       .TextMatrix(9, 1) = CANTI
    End With

     IMAG$ = "QR_ETIQPTERM.JPG"
     Identificador$ = "QR_ETIQPTERM"

     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     TEXTOQR$ = CODI$ & "|" & NORFA$ & "|" & CANTI & "|" & NC & "|" & CI1% & "|" & NP& & "|" & IDMOVISTO&

     On Error GoTo 0
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", 0, 0))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, Identificador$)

'    For i% = 1 To CANTIMPRE%
'      FILTX$ = TRIM$(RASO$) & ";" & OC$ & ";" & CODI$ & ";" & CODCLIENTE$ & ";" & DESKRI$ & ";" & NORFA$ & ";" & FECHA & ";" & CANTIPAQUETE$ & ";" & TRIM$(CStr(CANTI))
'     '   Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0, CANCOPIAS&)
'      Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0)
'    Next i%
    For i% = 1 To CANTIMPRE%
      Call Param_Print_Excel(RUTAEXCEL$, "ETIQUETAS", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0, "C:\FLEXOFT\" & IMAG$, 550, 29, 130, 130)                     'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Next i%
    Unload VENTABNEW
    Call SETPRINPORT("RESTORE")
    PRINTERPORT$ = ""
    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
'    If CANTIMPRE% < 1 Then CANTIMPRE% = 1
'    For i% = 1 To CANTIMPRE%
'      Call Param_Print_Excel(RUTAEXCEL$, "ETIQUETAS", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0, "C:\FLEXOFT\" & IMAG$, 550, 29, 130, 130)                     'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
'    Next i%
    
'    Unload VENTABNEW
    
    'ACTUALIZO LA CANTIDAD DE PAQUETES
    Call ACTUREGISTRO("MOVISTO", "IDCALIDAD", "MOVISTO_ID = " & IDMOVISTO&, CANTIPAQUETE$, REGAF&, "NOMESS")
    
End Sub


Sub AnularCierre(NROOF$, IDMOVISTO&, CANTANUL)
        If DERACCE%("ANULAROF", "Anulación de Orden de Fabricaciòn") < 2 Then
           Exit Sub
        End If
        Dim CTP As New Conteplast
        If CTP.SaldoTallerSegunOF(NROOF$) > 0 Then
           Call COMUNI("Atenciòn Esta Orden de Fabriaciòn Tiene Materiales en Taller\Si Confima la Misma Deberà Resolver Dicha Situaciòn\Desde Las Pantallas De Entrega o Recepciòn a Taller")
        End If
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        SQLX$ = "UPDATE MOVISTO SET CANTINGRE = 0,CANTSALID =  0, ReferCor =  LEFT('Anul.' + ReferCor,64),PESOKG=0,METROS=0"
        SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IDMOVISTO&
        FLEXCONN.Execute (SQLX$)
   
        SQLX$ = "UPDATE OrdeFabr SET StatuOrf = 1,CanApro = CanApro - " & TRIM$(FORMATNUM$(CANTANUL, "f12.1")) & ""
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "'"
        FLEXCONN.Execute (SQLX$)
        CUERPO$ = "Anulación de Cierre x " & TRIM$(FORMATNUM$(CANTANUL, "f12.1"))
        Call AGREGA_ANOTAORFAB(NROOF$, CUERPO$)
        
        CI1% = CODIFAB%(NROOF$)
        Call GENERES_ANUCIEOF(CI1%, CANTANUL, NROOF$)
        
        
        COTI& = NROCOTIZA(XVALO(Text12), XVALO(Text1))
        Call CTP.RecuperarConsumo_x_AnulaciónDeCierre(NROOF$, CANTANUL, COTI&)
        
        
'       IDMovistoSeguOf& = XVALO(VALOBADA("MOVISTO", "MOVISTO_ID", "DOPRILAR = '" & NORFA$ & "'", "NOMESS"))
'       DOCUORIGINAL$ = VALOBADA("MOVISTO", "DOCUORIG", "IDMOVISTO = '" & IDMOVISTO& & "'", "NOMESS") 'SE GUARDA LA RECEPCION DE TODOS LOS ITEMS QUE INGRESARON PARA ESA RECEPCION
'       Call BORRAREGISTROS("MOVTALLER", "DOCUORIG = '" & DOCUORIGINAL$ & "'", REGAF&, "NOMESS")
'       If REGAF& > 0 Then
'         Call COMUNI("Se han Borrado las Recepciones Asociadas a Este Cierre")
'       End If
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
      Call Label4_Change
    End If
End Sub
Private Sub Text2_LostFocus()
  Text2 = HORATEX$(HORANUM(Text2.TEXT))
End Sub

Sub ACTUALIZARUBICACION(RK, DESCRACK$, FL, KL, Optional LblId1&)

    '

End Sub
Sub ELIMINAR_CONTENIDO_UBICACION(UBI_ID&)
    CONDI$ = "UBILOTE_ENVAP_ID =" & UBI_ID&
    Call BORRAREGISTROS("UBILOTE_ENVAP", CONDI$, REGAF&, "NOMESS")
End Sub
    
    
'    If NUID& < 1 Then Exit Sub
'
'    NOP% = XVALO(RESP_ZELE_VECT(3))
'    NORFA$ = TRIM$(RESP_ZELE_VECT(5))
'    LOTE$ = TRIM$(RESP_ZELE_VECT(6))
'    CANTI = XVALO(RESP_ZELE_VECT(7))
'    CODIEXLM$ = TRIM$(TXTCODELEM)
'
'
'    OPX% = COMALTER%("Opción de Ubicaciones\\Canelar\Quitar Solo Lote Seleccionado")
'    If OPX% < 2 Then Exit Sub
'    If OPX% = 2 Then
'        Call BORRAREGISTROS("UBILOTE_ENVAP", "UBILOTE_ENVAP_ID = " & NUID&, REGAF&, "NOMESS")
'        CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
'        Condi$ = "IDSUBOR = '" & Label4 & "'"
'        Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
'        Condi$ = Condi$ + " AND CANTREAL = " & CANTI
'        Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
'        Condi$ = Condi$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
'        Condi$ = Condi$ + " AND  COORDUBIC = '" & CoordenadasUbicacion$ & "'"
'        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
'nload Me
'      U
'    End If

'End Sub
Function SELECCIONLOTEENUBICACION$(RK, FL, KL)
'    'SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED FROM UBILOTE_ENVAP
'    Campos$ = "O.F/a10" '1
'    Campos$ = Campos$ + " ;Componente/A32" '2
'    Campos$ = Campos$ + " ;CiComp/A0" '3
'    Campos$ = Campos$ + " ;Lote/a12" '4
'    Campos$ = Campos$ + " ;N°OP/f5" '5
'    Campos$ = Campos$ + " ;Desc./A14" '6
'    Campos$ = Campos$ + " ;Cant./F7.1" '7
'    Campos$ = Campos$ + " ;Orden/A5" '8
'
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'    CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
'    CONDI$ = " COORDUBIC = '" & CoordenadasUbicacion$ & "'"
'    SQLX$ = "SELECT IDSUBOR,CODXELEM,CODIELEM,IDENLOTE,NUMEOPER,DESCOPER,CANTREAL,NUORINF FROM INFOFAB WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE " & CONDI$
'    Dim obj As New RECORXET
'    Set RS1 = READSET(SQLX$)
'    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
'    Set obj = Nothing
'    RS1.Close
'    Set RS1 = Nothing
'
'    FRMZELECHO.Caption = "Rack: " & RK & " Piso: " & FL & " Columna: " & KL
'    FRMZELECHO.MSF1.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
'    FRMZELECHO.MSF2.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
'    FRMZELECHO.Width = 13000
'    FRMZELECHO.Show 1
'
'    SELECCIONLOTEENUBICACION$ = RESP_ZELE_VECT(1)

    'SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED FROM UBILOTE_ENVAP
    Campos$ = "Cta./F7;R.Social/A24;OP/f3;Desc./A12;O.F/a10;Lote/a12;Cant./F7.1;Ped./F7;ID/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    CONDI$ = " RACK = " & RK
    CONDI$ = CONDI$ + " AND FILA = " & FL
    CONDI$ = CONDI$ + " AND COLUMNA = " & KL

    SQLX$ = "SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED,UBILOTE_ENVAP_ID FROM UBILOTE_ENVAP WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim obj As New RECORXET
    Set rs1 = READSET(SQLX$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing

    FRMZELECHO.Caption = "Rack: " & RK & " Piso: " & FL & " Picking: " & KL
    FRMZELECHO.MSF1.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
    FRMZELECHO.MSF2.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
    FRMZELECHO.Width = 13000
    FRMZELECHO.Show 1

    SELECCIONLOTEENUBICACION$ = XVALO(RESP_ZELE_VECT(9))
End Function
Private Sub MSF2_Click()
    REG& = MSF2.MouseRow
    COL& = MSF2.MouseCol
    If REG& = 0 Or REG& >= MSF2.Rows Then Exit Sub
    RACK% = XVALO(MSF2.TextMatrix(REG&, 1))
    Label3 = RACK%

End Sub

Private Sub MSF3_Click()
    REG& = MSF3.MouseRow
    COL& = MSF3.MouseCol
    If REG& = 0 Or REG& >= MSF3.Rows Then Exit Sub
    '
    
    Call CARGARCASILLEROS(MSF3, REG&)
End Sub
Sub CARGARCASILLEROS(MSF As msFlexGrid, REG&)

    Campos$ = "CI/A0;" '1
    Campos$ = Campos$ + "Código/A16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "F.Emis./D8;" '4
    Campos$ = Campos$ + "Cantidad./F9.1" '5
    Campos$ = Campos$ + ";Movisto_id/A0;" '6
    Campos$ = Campos$ + "CANTPAQ/A0;" '7
    '
    
    TXTCANTICORTES = XVALO(MSF.TextMatrix(REG&, 5))
    lblidregistro = XVALO(MSF.TextMatrix(REG&, 6))
    TXTUBICSELEC = MSF.TextMatrix(REG&, 7)
    
    
    
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM$(XVALO(Text7), "F8.1")
End Sub

Sub ARMARGRILLAUBICA(NRACK%)
    ' **** SE ARMA LA GRILLA CON TODAS LAS UBICACIONES EN BLANCO ****
    MSF1.Rows = 1
    J& = 0
    VUELTA% = 0
    SQLX$ = "SELECT * FROM DATASQL WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE LISTA = 'CONFUBPT' AND CONVERT(int, campo1) > 0 "
    If NRACK% > 0 Then SQLX$ = SQLX$ + " AND CONVERT(int, campo1)= " & NRACK%
    SQLX$ = SQLX$ + " ORDER BY CAMPO1,CAMPO3,CAMPO4"
    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
         VUELTA% = VUELTA% + 1
         RACK% = XVALO(!CAMPO1)
         FIL% = XVALO(!CAMPO3)
         COL% = XVALO(!CAMPO4)
         For F% = 1 To FIL%
            For i% = 1 To COL%
               J& = J& + 1
               MSF1.Rows = J& + 1
               Call NEGRITA_SOLO_UNA_CELDA(MSF1, J&, 2)
               MSF1.TextMatrix(J&, 1) = XVALO(!CAMPO1)
               MSF1.TextMatrix(J&, 2) = SAFETEXT$(!campo2)
               MSF1.TextMatrix(J&, 3) = F% 'FILA
               MSF1.TextMatrix(J&, 4) = i% 'COLUMNA
               If (VUELTA% Mod 2) = 0 Then
                 Call COLOREAR_GRILLA_SOLAFILA(MSF1, &H9FBE85, J&, , 1, 1)
               End If
             Next i%
         Next F%
         .MoveNext
       Loop
    End With
    rs.Close: Set rs = Nothing
    
    'FIN **** SE ARMA LA GRILLA CON TODAS LAS UBICACIONES EN BLANCO ****
    '
    Call MostrarUbicacionesGuardadas
'    'RECORRO LAS UBICACIONES ASIGNADAS Y LO IMPRIMO EN GRILLA
'    SQLX$ = "SELECT *  FROM UBILOTE_ENVAP WITH(NOLOCK)"
'    SQLX$ = SQLX$ + "ORDER BY RACK,FILA,COLUMNA,LOTE,IDSUBOR"
'    Set RS = READSET(SQLX$)
'    With RS
''    .CellForeColor = vbRed
'      Do While Not .EOF
'         RACK% = XVALO(!RACK)
'         FIL% = XVALO(!fila)
'         COL% = XVALO(!COLUMNA)
'         CAN = XVALO(!CANTIDAD)
'         LOTE$ = XVALO(!LOTE)
'         Call UBICAR_UBICACION_EN_GRILLA(Me.MSF1, RACK%, FIL%, COL%, CAN, LOTE$)
'        .MoveNext
'      Loop
'    End With
'    RS.Close: Set RS = Nothing
End Sub
Sub MostrarUbicacionesGuardadas()

'ORIGINAL DESDE TABLA UBILOTE_ENVAP
    SQLX$ = "SELECT *  FROM UBILOTE_ENVAP WITH(NOLOCK)"
    SQLX$ = SQLX$ + "ORDER BY RACK,FILA,COLUMNA,LOTE,IDSUBOR"
    Set rs = READSET(SQLX$)
    With rs
'
      Do While Not .EOF
         RACK% = XVALO(!RACK)
         FIL% = XVALO(!Fila)
         COL% = XVALO(!COLUMNA)
         CAN = XVALO(!Cantidad)
         LOTE$ = SAFETEXT$(!LOTE)
         NORFA$ = SAFETEXT$(!idsubor)
         NC = XVALO(!NUCLIEN)
         NPED& = XVALO(!NROPED)
         FEENT% = CVINT(!FeSolEnt)
         IDMOVISTO& = XVALO(!id_movisto)
         IDUBILOTE& = XVALO(!UBILOTE_ENVAP_ID)
         
         Call UBICAR_UBICACION_EN_GRILLA(Me.MSF1, RACK%, FIL%, COL%, CAN, NORFA$, NC, NPED&, FEENT%, IDMOVISTO&, IDUBILOTE&)
        .MoveNext
      Loop
    End With
    rs.Close: Set rs = Nothing
End Sub
Sub UBICAR_UBICACION_EN_GRILLA(MSF As msFlexGrid, RACK%, FIL%, COL%, CAN, NORFA$, NC, NPED&, FEENT%, Optional LblId1&, Optional IDUBILOTE&)
    Campos$ = "Rack/F3;" '1
    Campos$ = Campos$ + "Descripción/A24;" '2
    Campos$ = Campos$ + "Piso/F4;" '3
    Campos$ = Campos$ + "Picking/F7;" '4
    Campos$ = Campos$ + "Cant./F5" '5
    Campos$ = Campos$ + ";O.F./A12;" '6
    Campos$ = Campos$ + "MOVISTO_ID/A10;" '7
    Campos$ = Campos$ + "NC/A0;" '8
    Campos$ = Campos$ + "Razon Social/A24;" '9
    Campos$ = Campos$ + "N°Pedido/f7;" '10
    Campos$ = Campos$ + "F.Entrega/D8;" '11
    Campos$ = Campos$ + "UBILOTE_ID/A0" '11
    A = rasocli$(1) 'solo para refrescar
    For i% = 1 To MSF.Rows - 1
       RACKENGRILLA% = XVALO(MSF.TextMatrix(i%, 1))
       If RACKENGRILLA% = RACK% Then
          FILAENGRILLA% = XVALO(MSF.TextMatrix(i%, 3))
          If FILAENGRILLA% = FIL% Then
             COLUENGRILLA% = XVALO(MSF.TextMatrix(i%, 4))
             If COLUENGRILLA% = COL% Then
               PCOMA$ = ""
               LoteAsignar$ = LOTEENGRILLA$ & PCOMA$ & LOTE$
               MSF.TextMatrix(i%, 5) = CAN
               MSF.TextMatrix(i%, 6) = NORFA$
               MSF.TextMatrix(i%, 7) = LblId1&

               MSF.TextMatrix(i%, 8) = NC
               MSF.TextMatrix(i%, 9) = rasocli$(NC)
               MSF.TextMatrix(i%, 10) = NPED&
               MSF.TextMatrix(i%, 11) = FECHATEX$(FEENT%)
               MSF.TextMatrix(i%, 12) = XVALO(IDUBILOTE&)
              
               MSF.Row = i%
               MSF.COL = 5: MSF.CellForeColor = &H584B14
               MSF.COL = 6: MSF.CellForeColor = &H584B14
               MSF.COL = 7: MSF.CellForeColor = &H584B14
               MSF.COL = 8: MSF.CellForeColor = &H584B14
               MSF.COL = 9: MSF.CellForeColor = &H584B14
               MSF.COL = 10: MSF.CellForeColor = &H584B14
               MSF.COL = 11: MSF.CellForeColor = &H584B14
               MSF.COL = 12: MSF.CellForeColor = &H584B14

               Exit For
             End If
          End If
       End If
    Next i%
End Sub


Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub


