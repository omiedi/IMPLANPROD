VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ContConsumo 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18795
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   18795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   -120
      Width           =   18855
      Begin VB.TextBox Text10 
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
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Ingresar Código de Producto Terminado"
         Top             =   720
         Visible         =   0   'False
         Width           =   1890
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
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   360
         Width           =   720
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
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox Text13 
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
         Left            =   12825
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox Text14 
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
         Left            =   1230
         LinkTimeout     =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   1080
         Width           =   1080
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
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   1080
         Width           =   1185
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
         Index           =   1
         Left            =   9510
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   1080
         Width           =   720
      End
      Begin VB.TextBox TXTNROCOTIZA 
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
         Left            =   12825
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   " "
         Top             =   1080
         Width           =   720
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
         Left            =   3360
         TabIndex        =   9
         ToolTipText     =   "Cerradas"
         Top             =   360
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
         Left            =   3120
         TabIndex        =   5
         ToolTipText     =   "Abiertas"
         Top             =   360
         Width           =   175
      End
      Begin VB.Label Label2 
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
         Left            =   3510
         TabIndex        =   28
         Top             =   360
         Width           =   4770
      End
      Begin VB.Label Label6 
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
         Height          =   300
         Left            =   -240
         TabIndex        =   27
         Top             =   795
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   3390
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   4890
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
         Left            =   8310
         TabIndex        =   25
         Top             =   420
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
         Left            =   8430
         TabIndex        =   24
         Top             =   720
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
         Left            =   10320
         TabIndex        =   23
         Top             =   360
         Width           =   3210
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.del Pedido:"
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
         Index           =   0
         Left            =   11175
         TabIndex        =   22
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.O.F..:"
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
         Index           =   1
         Left            =   75
         TabIndex        =   21
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   " Pendiente O.F.:"
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
         Index           =   2
         Left            =   5430
         TabIndex        =   20
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   8550
         TabIndex        =   19
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización N°:"
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
         Index           =   1
         Left            =   10900
         TabIndex        =   18
         Top             =   1080
         Width           =   1845
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
         Left            =   -120
         TabIndex        =   7
         Top             =   435
         Width           =   1335
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
         Left            =   1230
         TabIndex        =   6
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   360
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      Caption         =   "Consumos Realizados en Metros y Kilos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Index           =   1
      Left            =   -120
      TabIndex        =   2
      Top             =   5640
      Width           =   18915
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   3650
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Cortes informados Primarios "
         Top             =   240
         Width           =   18700
         _ExtentX        =   32994
         _ExtentY        =   6429
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
      Caption         =   "Materiales Transferidos al Depòsito de Corte Segùn O.F. Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   18800
      Begin VB.TextBox Text9 
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
         Left            =   17445
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   3840
         Width           =   1260
      End
      Begin VB.TextBox Text8 
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
         Left            =   17445
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   3480
         Width           =   1260
      End
      Begin VB.TextBox TXTMETROSCONSUMIDOS 
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
         Left            =   15165
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   3840
         Width           =   1260
      End
      Begin VB.TextBox TXTKILOSCONSUMIDOS 
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
         Left            =   15165
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   3480
         Width           =   1260
      End
      Begin VB.TextBox TXTMETROSTRANSFERIDOS 
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
         Left            =   12045
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   3840
         Width           =   1260
      End
      Begin VB.TextBox TXTKILTRANSFERIDOS 
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
         Left            =   12045
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   3480
         Width           =   1260
      End
      Begin VB.TextBox TXTDifKgsEntreTransfeInf 
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
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox TXTKgsSegunTransf 
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
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox TXTDIFERENCIAENMETROS 
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
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   1800
         Width           =   800
      End
      Begin VB.TextBox TxtCortesEnKilos 
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
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   2040
         Width           =   800
      End
      Begin VB.TextBox TXTCTESINFTRANSFERIDOSMTS 
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
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   1080
         Width           =   800
      End
      Begin VB.TextBox Text3 
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
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox TXTLOTE 
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
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   " "
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox TXTCODELEM 
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
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   " "
         Top             =   600
         Width           =   1656
      End
      Begin VB.TextBox TXTCORTE 
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
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   " "
         ToolTipText     =   "Es la medida del corte en mts."
         Top             =   2400
         Width           =   780
      End
      Begin VB.TextBox TXTCANTICORTES 
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
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   " "
         ToolTipText     =   "La cantidad de cortes que tiene por unidad "
         Top             =   2400
         Width           =   780
      End
      Begin VB.TextBox TXTCORTESOPYLOTE 
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
         Left            =   14640
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   " "
         Top             =   1440
         Width           =   800
      End
      Begin VB.TextBox TXTCORTESOPERACION 
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
         Left            =   17445
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   " "
         Top             =   1200
         Width           =   1260
      End
      Begin VB.TextBox TXTNECESIDADCORTES 
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
         Left            =   17445
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   " "
         Top             =   240
         Width           =   1260
      End
      Begin VB.TextBox TXTRESTAN 
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
         Left            =   17445
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   " "
         Top             =   600
         Width           =   1260
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   3900
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Materiales Transferidos a Sector de Corte"
         Top             =   360
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   6879
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
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos Transferiddos"
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
         Index           =   21
         Left            =   10200
         TabIndex        =   72
         Top             =   3600
         Width           =   1725
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dif.Metros:"
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
         Index           =   20
         Left            =   16305
         TabIndex        =   71
         Top             =   3915
         Width           =   1215
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dif.Kilos:"
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
         Index           =   19
         Left            =   16440
         TabIndex        =   70
         Top             =   3555
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Metros Consumidos"
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
         Index           =   18
         Left            =   13320
         TabIndex        =   67
         Top             =   3915
         Width           =   1815
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos Consumidos"
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
         Index           =   17
         Left            =   13320
         TabIndex        =   66
         Top             =   3540
         Width           =   1725
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Metros Transferidos"
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
         Index           =   16
         Left            =   10365
         TabIndex        =   63
         Top             =   3915
         Width           =   1695
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Index           =   4
         Left            =   9600
         TabIndex        =   62
         Top             =   3555
         Width           =   765
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   10560
         X2              =   18600
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia Entre Transferido e informado en Kilos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   15
         Left            =   13920
         TabIndex        =   59
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos Segùn Transferido: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   14
         Left            =   14880
         TabIndex        =   56
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia Entre Transferido e informado en Metros:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   13
         Left            =   10725
         TabIndex        =   54
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados en Kilos del Lote Activo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   12
         Left            =   13800
         TabIndex        =   52
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados en Metros de operaciòn y Lote Activo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   11
         Left            =   10275
         TabIndex        =   51
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Cortes Teoricos Segun lo Transferido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   10
         Left            =   10650
         TabIndex        =   49
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   15200
         X2              =   18720
         Y1              =   1000
         Y2              =   1000
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
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
         Left            =   10200
         TabIndex        =   48
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label7 
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
         Left            =   12240
         TabIndex        =   47
         Top             =   240
         Width           =   2850
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Index           =   9
         Left            =   12960
         TabIndex        =   46
         Top             =   675
         Width           =   525
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código: "
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
         Index           =   8
         Left            =   10560
         TabIndex        =   45
         Top             =   675
         Width           =   765
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Index           =   1
         Left            =   10275
         TabIndex        =   44
         Top             =   2475
         Width           =   765
      End
      Begin VB.Label LABEL12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Cortes:"
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
         Index           =   0
         Left            =   11880
         TabIndex        =   43
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados de Operación Activa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   15600
         TabIndex        =   42
         Top             =   1140
         Width           =   1725
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Necesidad de Cortes Para la Operación Activa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   5
         Left            =   15390
         TabIndex        =   41
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Restan:"
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
         Index           =   7
         Left            =   15840
         TabIndex        =   40
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Label LABEL12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. Cortes:"
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
      Index           =   3
      Left            =   2325
      TabIndex        =   30
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label LABEL12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Corte:"
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
      Index           =   2
      Left            =   0
      TabIndex        =   29
      Top             =   75
      Width           =   765
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ContConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    MODO% = 3
    Label4 = OrdenFabricacionAbiertas$(MODO%) 'abiertas y cerradas

End Sub

Private Sub Command3_Click()
    MODO% = 0
'    If Cerradas.Value = True Then MODO% = 3
'    If Pendientes.Value = True Then MODO% = 1
    Label4 = OrdenFabricacionAbiertas$(MODO%) 'abiertas y cerradas

End Sub

Private Sub Form_Load()
  Call CARGA_ENCABEZADO(MSF2, "CODIGO/A16;Descripción/A32;LOTE/A11;Transf.Mts./f9.1;CI/A0;Op.N°/F4;Kg.Transf./F8.1", Me) ';Cant.Cortes/F8.1;Mts.Cortes/F8.1;Kgs.Cortes/F8.1;Dif.Mts/F8.1;Dif.Kgs/F8.1"
  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Cons.Mts./F10.1;Kgs./F8.1", Me)

End Sub

Private Sub Label4_Change()
   Dim CTP As New Conteplast
   DEPOCORTE% = XVALO(CONTROL("", "DEPODEST"))
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then Exit Sub
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
   Text14 = CANTI
   CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CIFAB%
   Text13 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
   Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$)
 
   TXTNROCOTIZA = NROCOTIZA(XVALO(NROPED&), XVALO(Text1))
    
    
   Call CTP.CargaMatTranfXOrdenNoConsideraConsumos(MSF2, TRIM$(NORFA$), DEPOCORTE%, 0, 0, 1, "COD_EXTE,ID_OPERACION, IDENLOTE")
   
   Call CTP.MostrarConsumoSegunOf(MSF3, NORFA$, Me)
    
    Set CTP = Nothing
End Sub


Private Sub MSF2_Click()
    REG& = MSF2.MouseRow
   j& = MSF2.MouseCol
   
   If REG& = 0 Then
     Static MODODEORDEN$
     If MODODEORDEN$ = "DESC" Then
        MODODEORDEN$ = ""
     Else
        MODOORD% = 0
        MODODEORDEN$ = "DESC"
     End If
     Call ORDENAR_MSF(INFOCORTE.MSF2, j&, MODODEORDEN$)
     Exit Sub
   End If
   
   If REG& >= MSF2.Rows Then Exit Sub
   
   Call SELECCIONARFILA(MSF2, REG&)
'   Me.LblKgImpresos = ""
   NOP% = XVALO(MSF2.TextMatrix(REG&, 6))
   Text3 = NOP%
   TXTLOTE = MSF2.TextMatrix(REG&, 3)
   
   TXTCODELEM = MSF2.TextMatrix(REG&, 1)
   CIELEM% = XVALO(MSF2.TextMatrix(REG&, 5))
   CODIGOFAB = CODINT%(Text10)
   
   Condi1$ = "CodIConj = " & CODIGOFAB & " and  NumeOper = " & NOP%
   Label7 = VALOBADA("SECOPER", "DescOper", Condi1$)
   
   CONDI$ = "CODICONJ = " & CODIGOFAB
   CONDI$ = CONDI$ + " AND (CODIELEM = " & CIELEM% & "  or CODIELEM = " & CODIGORANGO&(CIELEM%) & ")"
   DESOP$ = TRIM$(Label7)
   
   If DESOP$ <> "" Then
    CONDI$ = CONDI$ + " AND COMPONENTE = '" & DESOP$ & "'"
   End If
   'LO PASO A MTS 'EL CORTE EN LA TABLA FORMULAS ESTA EN CM POR QUE ASI LO CARGAN EN LA COTIZACION
   TXTCORTE = XVALO(VALOBADA("FORMULAS", "CORTE", CONDI$, "NOMESS")) / 100
   Me.TXTCANTICORTES = XVALO(VALOBADA("FORMULAS", "MODUXBB", CONDI$, "NOMESS"))
   
   METROS# = XVALO(MSF2.TextMatrix(REG&, 4))
   TXTCTESINFTRANSFERIDOSMTS = ""
   If XVALO(TXTCORTE) > 0 Then
     TXTCTESINFTRANSFERIDOSMTS = CInt(METROS# / XVALO(TXTCORTE))
   End If
   
   KILOSTRANSFERIDOS# = XVALO(MSF2.TextMatrix(REG&, 7))
   TXTKgsSegunTransf = FORMATNUM$(KILOSTRANSFERIDOS#, "F10.1")
   
   NORFA$ = TRIM$(Label4)
   NOPER% = XVALO(Text3)
   LOTE$ = TRIM$(TXTLOTE)
   'CANTIDAD DE CORTES DE LA OF X LOTE Y OPERACION
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER%
   PesoDeLosCortesSegunLote = XVALO(VALOBADA("INFOFAB", "sum(peso)", CONDI$, "NOMESS"))
   TxtCortesEnKilos = FORMATNUM$(PesoDeLosCortesSegunLote, "f10.1")
   IDREGSELEC& = XVALO(VALOBADA("INFOFAB", "INFOFAB_ID", CONDI$, "NOMESS"))
   CONDI2$ = "IdRegPadre = " & IDREGSELEC&
   TieneHerencia& = XVALO(VALOBADA("InfoFab", "IdRegPadre", CONDI2$, "NOMESS"))
   If TieneHerencia& > 0 Then '14/07/21
      'cuando se realiza subprocesos se pone cantreal en 0 y se aplica el valor al nuevo registro generado
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("InfoFab", "sum(CantOrigCte)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CantOrigCte)", CONDI$, "NOMESS"))
   ElseIf TRIM$(UCase$(CODFAMIL$(CIELEM%))) = "EE" Then ' ESLINGAS
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("InfoFab", "sum(CANTREAL)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
   Else
      Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
      CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
      Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
'21/07/21 se corrigio a esta linea queda igual que if tieneherencia pero lo dejo separado para ver si va todo bien
      
   End If
   
   'CANTIDAD DE CORTES DE LA OF X  OPERACION
   Me.TXTNECESIDADCORTES = FORMATNUM$(XVALO(TXTCANTICORTES) * XVALO(Text14), "F8")
   TXTRESTAN = FORMATNUM$(XVALO(TXTNECESIDADCORTES) - XVALO(TXTCORTESOPERACION), "F8")
End Sub


Private Sub MSF3_DblClick()
    REG& = MSF3.MouseRow
    COL& = MSF3.MouseCol
    If REG& = 0 Or REG& >= MSF3.Rows Then Exit Sub
    Call SELECCIONARFILA(MSF3, REG&)
    
    DEPOCORTE% = XVALO(CONTROL("", "DEPODEST"))


   IDLOT$ = MSF3.TextMatrix(REG&, 4)
   TXTCODIGOELEM = MSF3.TextMatrix(REG&, 2)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TXTCODIGOELEM
     CONSUMLOT07.Text4 = DEPOCORTE%
     CONSUMLOT07.Label33.Caption = "P" & Mid$(Label4.Caption, 2) 'PF-001043-1
     CONSUMLOT07.Label33.Visible = True
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub


Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub


Private Sub TxtCortesEnKilos_Change()
   Me.TXTDifKgsEntreTransfeInf = actualizaDiferenciaEntreTeoricoYCortesKILOS
End Sub

Private Sub TXTCORTESOPYLOTE_Change()
   TXTDIFERENCIAENMETROS = actualizaDiferenciaEntreTeoricoYCortesMts
End Sub

Private Sub TXTCTESINFTRANSFERIDOSMTS_Change()
   TXTDIFERENCIAENMETROS = actualizaDiferenciaEntreTeoricoYCortesMts
End Sub
Function actualizaDiferenciaEntreTeoricoYCortesMts()
   TXTDIFERENCIAENMETROS.ForeColor = Black
   actualizaDiferenciaEntreTeoricoYCortesMts = CInt(XVALO(TXTCTESINFTRANSFERIDOSMTS) - XVALO(TXTCORTESOPYLOTE))
   If actualizaDiferenciaEntreTeoricoYCortesMts < 0 Then
      TXTDIFERENCIAENMETROS.ForeColor = vbRed
   End If
End Function
Function actualizaDiferenciaEntreTeoricoYCortesKILOS()
   TXTDifKgsEntreTransfeInf.ForeColor = Black
   actualizaDiferenciaEntreTeoricoYCortesKILOS = CInt(XVALO(TXTKgsSegunTransf) - XVALO(TxtCortesEnKilos))
   If actualizaDiferenciaEntreTeoricoYCortesKILOS < 0 Then
      TXTDifKgsEntreTransfeInf.ForeColor = vbRed
   End If
End Function

Private Sub TXTKgsSegunTransf_Change()
   Me.TXTDifKgsEntreTransfeInf = actualizaDiferenciaEntreTeoricoYCortesKILOS
End Sub


Private Sub TXTKILOSCONSUMIDOS_Change()
    Text8 = FORMATNUM$(XVALO(TXTKILTRANSFERIDOS) - XVALO(TXTKILOSCONSUMIDOS), "f10.1")
End Sub

Private Sub TXTKILTRANSFERIDOS_Change()
    Text8 = FORMATNUM$(XVALO(TXTKILTRANSFERIDOS) - XVALO(TXTKILOSCONSUMIDOS), "f10.1")
End Sub


Private Sub TXTMETROSCONSUMIDOS_Change()
    Text9 = FORMATNUM$(XVALO(TXTMETROSTRANSFERIDOS) - XVALO(TXTMETROSCONSUMIDOS), "f10.1")
End Sub

Private Sub TXTMETROSTRANSFERIDOS_Change()
    Text9 = FORMATNUM$(XVALO(TXTMETROSTRANSFERIDOS) - XVALO(TXTMETROSCONSUMIDOS), "f10.1")

End Sub


