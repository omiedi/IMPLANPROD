VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RECUEP07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Acreedores (Gestion de Compras)"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   15810
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid mfgCuentaCorrienteAHP 
      Height          =   2300
      Left            =   90
      TabIndex        =   86
      Top             =   2400
      Visible         =   0   'False
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   5273178
      BackColorBkg    =   14737632
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid mfgCuentaCorriente 
      Height          =   2500
      Left            =   90
      TabIndex        =   85
      Top             =   2400
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   4419
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   5273178
      BackColorBkg    =   14737632
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid mfgComposicionDeuda 
      Height          =   2535
      Left            =   90
      TabIndex        =   84
      Top             =   5000
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   5273178
      BackColorBkg    =   14737632
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   10680
      Picture         =   "RECUEP12.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Actualizar"
      Top             =   8040
      Width           =   405
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   2750
      Index           =   8
      Left            =   15000
      TabIndex        =   59
      Top             =   5040
      Width           =   15
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   3720
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      ScaleHeight     =   345
      ScaleWidth      =   10935
      TabIndex        =   38
      Top             =   7520
      Width           =   10965
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9195
         TabIndex        =   57
         Top             =   45
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   50
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7680
         TabIndex        =   40
         Top             =   45
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   45
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2560
      Left            =   90
      ScaleHeight     =   2535
      ScaleWidth      =   9855
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   9890
      Begin VB.Frame Frame7 
         Caption         =   "Frame11"
         Height          =   2500
         Index           =   5
         Left            =   9600
         TabIndex        =   37
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame10"
         Height          =   2500
         Index           =   4
         Left            =   7680
         TabIndex        =   36
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame9"
         Height          =   2500
         Index           =   3
         Left            =   6195
         TabIndex        =   35
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame8"
         Height          =   2500
         Index           =   2
         Left            =   4830
         TabIndex        =   34
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2500
         Index           =   1
         Left            =   2480
         TabIndex        =   33
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame6"
         Height          =   2500
         Index           =   0
         Left            =   1080
         TabIndex        =   32
         Top             =   0
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   -10
         TabIndex        =   31
         Top             =   360
         Width           =   9885
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha     Documento  Número                   Debe        Haber        Saldo"
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
         Height          =   330
         Left            =   0
         TabIndex        =   49
         Top             =   60
         Width           =   10650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   9555
      Left            =   14880
      ScaleHeight     =   9525
      ScaleWidth      =   1050
      TabIndex        =   20
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   1075
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Búsqueda por Texto"
         Top             =   1680
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   81
         Top             =   7320
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   82
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "I-Cre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Imputación de Créditos y Anticipos"
         Top             =   4485
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Impresión de Cuenta Corriente"
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Note"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Notas de Proveedores"
         Top             =   3090
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "R-Apl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":1F6E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   5280
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":20B8
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   140
         Picture         =   "RECUEP12.frx":23C2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   300
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   90
      ScaleHeight     =   2535
      ScaleWidth      =   9855
      TabIndex        =   17
      Top             =   5000
      Width           =   9885
      Begin VB.Frame Frame7 
         Caption         =   "Frame6"
         Height          =   2500
         Index           =   10
         Left            =   1320
         TabIndex        =   48
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2500
         Index           =   11
         Left            =   3705
         TabIndex        =   47
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame8"
         Height          =   2500
         Index           =   12
         Left            =   4800
         TabIndex        =   46
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame9"
         Height          =   2500
         Index           =   13
         Left            =   6435
         TabIndex        =   45
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame10"
         Height          =   2500
         Index           =   14
         Left            =   7905
         TabIndex        =   44
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame11"
         Height          =   2500
         Index           =   15
         Left            =   9450
         TabIndex        =   43
         Top             =   0
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   -10
         TabIndex        =   42
         Top             =   360
         Width           =   9900
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento   Número              Emisión      Importe  Vencimiento   Saldo ($)"
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
         Height          =   345
         Left            =   -10
         TabIndex        =   50
         Top             =   60
         Width           =   10650
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Width           =   10650
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   8640
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   8040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "DATOS DE LA CUENTA ACTIVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   90
      TabIndex        =   2
      Top             =   0
      Width           =   10965
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Vista Ampliada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Left            =   1245
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C7D9CC&
         Caption         =   "Moneda - Cotización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   6500
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   4305
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   105
            TabIndex        =   56
            ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
            Top             =   200
            Width           =   2430
         End
         Begin VB.Label TICAMBIS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   285
            Left            =   3120
            TabIndex        =   55
            ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T-C"
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
            Left            =   2520
            TabIndex        =   54
            ToolTipText     =   "Doble Click para Configurar Cliente"
            Top             =   315
            Width           =   510
         End
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   8760
         TabIndex        =   18
         Text            =   " "
         Top             =   1275
         Width           =   1905
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   8760
         TabIndex        =   14
         Text            =   " "
         Top             =   1920
         Width           =   1905
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1245
         TabIndex        =   11
         Text            =   " "
         Top             =   1920
         Width           =   4530
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
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   1155
         Width           =   180
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1245
         TabIndex        =   8
         Text            =   " "
         Top             =   780
         Width           =   4530
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1245
         TabIndex        =   7
         Text            =   " "
         Top             =   1155
         Width           =   4530
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   1245
         TabIndex        =   6
         Text            =   " "
         Top             =   1545
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   2085
         TabIndex        =   5
         Text            =   " "
         Top             =   1545
         Width           =   3690
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   6555
         TabIndex        =   4
         Text            =   " "
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   6555
         TabIndex        =   3
         Text            =   " "
         Top             =   1275
         Width           =   1695
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   1545
         Width           =   970
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta"
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
         Index           =   2
         Left            =   8550
         TabIndex        =   19
         Top             =   1005
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Acreedor"
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
         Index           =   3
         Left            =   8520
         TabIndex        =   15
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Index           =   1
         Left            =   6450
         TabIndex        =   13
         Top             =   1665
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Left            =   130
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono"
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
         Index           =   0
         Left            =   6450
         TabIndex        =   9
         Top             =   1005
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   1800
      OleObjectBlob   =   "RECUEP12.frx":250C
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "RECUEP12.frx":2586
      Top             =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1095
      Left            =   75
      TabIndex        =   51
      Top             =   8040
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1931
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2570
      Left            =   90
      ScaleHeight     =   2535
      ScaleWidth      =   14640
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   14670
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   5
         Left            =   9465
         TabIndex        =   71
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame10"
         Height          =   2550
         Index           =   4
         Left            =   7770
         TabIndex        =   70
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame9"
         Height          =   2550
         Index           =   3
         Left            =   6195
         TabIndex        =   69
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2550
         Index           =   2
         Left            =   4830
         TabIndex        =   68
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame7"
         Height          =   2550
         Index           =   1
         Left            =   2400
         TabIndex        =   67
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame6"
         Height          =   2550
         Index           =   0
         Left            =   1080
         TabIndex        =   66
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   6
         Left            =   10680
         TabIndex        =   65
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   7
         Left            =   11880
         TabIndex        =   64
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   8
         Left            =   12360
         TabIndex        =   63
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   9
         Left            =   12720
         TabIndex        =   62
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame11"
         Height          =   2550
         Index           =   10
         Left            =   13080
         TabIndex        =   61
         Top             =   0
         Width           =   15
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -10
         TabIndex        =   75
         Top             =   360
         Width           =   14700
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha      Documento    Numero                       Debe           Haber          Saldo"
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
         Left            =   0
         TabIndex        =   72
         Top             =   80
         Width           =   14700
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   14700
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Diferencia"
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
      Left            =   10800
      TabIndex        =   80
      ToolTipText     =   "Doble Click para Configurar Cliente"
      Top             =   5220
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label LBLTOTDIFERENCIA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   13080
      TabIndex        =   79
      Top             =   5160
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label LBLTOTREAL_DEBE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roboto Mono Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   13080
      TabIndex        =   78
      Top             =   5760
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real - Debe"
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
      Left            =   10800
      TabIndex        =   77
      ToolTipText     =   "Doble Click para Configurar Cliente"
      Top             =   5820
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
   End
   Begin VB.Menu mnuConsultasylistado 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ComposicionDeuda 
         Caption         =   "Composición de Deuda a Fecha"
      End
      Begin VB.Menu ComposicionDeudaProveedor 
         Caption         =   "Composición de Deuda por Proveedor"
      End
      Begin VB.Menu ObligacionesPagarPorFecha 
         Caption         =   "Obligaciones a Pagar por Rango de Fecha"
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar texto en Cuenta Corriente"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuImprimirCtaCte 
         Caption         =   "Imprimir Cta. Corriente"
      End
      Begin VB.Menu mnuNotaProveedor 
         Caption         =   "Notas de Proveedor"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuReapli 
         Caption         =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
      End
      Begin VB.Menu mnuImputacionCredAnt 
         Caption         =   "Imputación de Créditos y Anticipos"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
   End
   Begin VB.Menu mnuADirectos 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "RECUEP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%
Dim MONRESU%
Public ARGU$



Sub ExportaAExcelVistaAmpliada(NombreDeArchivo$)

   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
   Dim FORMATONUMERO As String
   Dim IDIOMA As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   'CARGAENCABEZADO
   RENG1& = 1
   XLSHEET.Cells(RENG1&, 1) = "Fecha":
   XLSHEET.Cells(RENG1&, 2) = "Tipo Cbte."
   XLSHEET.Cells(RENG1&, 3) = "Documento"
   XLSHEET.Cells(RENG1&, 4) = "Debe"
   XLSHEET.Cells(RENG1&, 5) = "Haber"
   XLSHEET.Cells(RENG1&, 6) = "Saldo"
   XLSHEET.Cells(RENG1&, 7) = "ME"
   XLSHEET.Cells(RENG1&, 8) = "Cotiz."
   XLSHEET.Cells(RENG1&, 9) = "Val.Compra"
   XLSHEET.Cells(RENG1&, 10) = "G. Varios"
   XLSHEET.Cells(RENG1&, 11) = "V.Real"
   XLSHEET.Cells(RENG1&, 12) = "Diferencia"

   'RECORRO LAS FILAS Y VOY AGREGANDO LA INFO A EXCEL
   For U& = 1 To mfgCuentaCorrienteAHP.Rows - 1
         RENG1& = RENG1& + 1
         FECHAX$ = mfgCuentaCorrienteAHP.TextMatrix(U&, 1)
           XLSHEET.Cells(RENG1&, 1).NumberFormat = "dd/mm/yyyy"
           XLSHEET.Cells(RENG1&, 1) = CDate(FECHAX$)
         TIPOS$ = mfgCuentaCorrienteAHP.TextMatrix(U&, 2)
           XLSHEET.Cells(RENG1&, 2) = TIPOS$
         NROX$ = mfgCuentaCorrienteAHP.TextMatrix(U&, 3)
           XLSHEET.Cells(RENG1&, 3) = NROX$
         DEBE$ = mfgCuentaCorrienteAHP.TextMatrix(U&, 4)
           XLSHEET.Cells(RENG1&, 4).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 4) = XVALO(DEBE$)
         HABER = mfgCuentaCorrienteAHP.TextMatrix(U&, 5)
           XLSHEET.Cells(RENG1&, 5).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 5) = XVALO(HABER)
         SALDOX = mfgCuentaCorrienteAHP.TextMatrix(U&, 6)
           XLSHEET.Cells(RENG1&, 6).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 6) = XVALO(SALDOX)
         SIMBOLOX$ = mfgCuentaCorrienteAHP.TextMatrix(U&, 7)
           XLSHEET.Cells(RENG1&, 7) = SIMBOLOX$
         TICAMBIOX = mfgCuentaCorrienteAHP.TextMatrix(U&, 8)
           XLSHEET.Cells(RENG1&, 8).NumberFormat = "#,##0.0000"
           XLSHEET.Cells(RENG1&, 8) = XVALO(TICAMBIOX)
         COMPRA = mfgCuentaCorrienteAHP.TextMatrix(U&, 9)
           XLSHEET.Cells(RENG1&, 9).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 9) = XVALO(COMPRA)
         GASTOS = mfgCuentaCorrienteAHP.TextMatrix(U&, 10)
           XLSHEET.Cells(RENG1&, 10).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 10) = XVALO(GASTOS)
         VALREAL = mfgCuentaCorrienteAHP.TextMatrix(U&, 11)
           XLSHEET.Cells(RENG1&, 11).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 11) = XVALO(VALREAL)
         DIFER_COMPRA_HABER = mfgCuentaCorrienteAHP.TextMatrix(U&, 12)
           XLSHEET.Cells(RENG1&, 12).NumberFormat = "#,##0.00"
           XLSHEET.Cells(RENG1&, 12) = XVALO(DIFER_COMPRA_HABER)
   Next U&


'   'RECORRO LAS FILAS Y VOY AGREGANDO LA INFO A EXCEL
'   For U& = 0 To List3.ListCount - 1
'         X$ = List3.LIST(U&)
'         FECHAX$ = Mid$(X$, 1, 8)
'         RENG1& = RENG1& + 1
'           VALORX = FETEXCOR2$(Int(FECHANUM(FECHAX$)))
'           XLSHEET.Cells(RENG1&, 1) = VALORX
'         TIPOS$ = Mid$(X$, 11, 11)
'           XLSHEET.Cells(RENG1&, 2) = TIPOS$
'         NROX$ = Mid$(X$, 22, 18)
'           XLSHEET.Cells(RENG1&, 3) = NROX$
'           TEXTODEBE$ = TRIM$(Mid$(X$, 40, 24))
'           If Left$(TEXTODEBE$, 24) = ".... Moneda: Pesos ...." Then
'             XLSHEET.Cells(RENG1&, 4) = TEXTODEBE$
'           Else
'           DEBE = ROUND(XVALO(Mid$(X$, 40, 11)), 2)
'           XLSHEET.Cells(RENG1&, 4) = DEBE
'
'           End If
'         HABER = ROUND(XVALO(Mid$(X$, 53, 11)), 2)
'           XLSHEET.Cells(RENG1&, 5) = HABER
'         SALDOX = ROUND(XVALO(Mid$(X$, 64, 13)), 2)
'           XLSHEET.Cells(RENG1&, 6) = SALDOX
'         SIMBOLOX$ = Mid$(X$, 79, 2)
'           XLSHEET.Cells(RENG1&, 7) = SIMBOLOX$
'         TICAMBIOX = ROUND(XVALO(Mid$(X$, 81, 9)), 4)
'           XLSHEET.Cells(RENG1&, 8) = TICAMBIOX
'         COMPRA = ROUND(XVALO(Mid$(X$, 90, 11)), 2)
'           XLSHEET.Cells(RENG1&, 9) = COMPRA
'         GASTOS = ROUND(XVALO(Mid$(X$, 101, 11)), 2)
'           XLSHEET.Cells(RENG1&, 10) = GASTOS
'         VALREAL = ROUND(XVALO(Mid$(X$, 112, 13)), 2)
'           XLSHEET.Cells(RENG1&, 11) = VALREAL
'         DIFER_COMPRA_HABER = ROUND(XVALO(Mid$(X$, 125, 13)), 2)
'           XLSHEET.Cells(RENG1&, 12) = DIFER_COMPRA_HABER
'   Next U&
'
'   '
'   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '


End Sub

Sub pantallaAhp()
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       '
'       Text7.FontName = "Lucida Console"
'       Text7.FontSize = 9
'       Text7.FontBold = True
'       Text7.Height = 300
'       '
'       Text6.FontName = "Lucida Console"
'       Text6.FontSize = 9
'       Text6.FontBold = True
'       Text6.Height = 300
'       '
'       Frame6.Visible = True
'       '
       'List3.FontName = "Lucida Console"
'       List3.FontSize = 9
'       List3.FontBold = False
'       '
'       Label14.Caption = Label9.Caption + " Val.Compra   G.Varios   Valor Real   Diferencia"
'       Label14.FontName = "Lucida Console"
'       Label14.FontSize = 9
'       Label14.FontBold = False
'       Label14.Top = 90
       '
       'List2.FontName = "Lucida Console"
'       List2.FontSize = 9
'       List2.FontBold = False
       '
'       Frame8(0).Left = 1080
'       Frame8(1).Left = 2230
'       Frame8(2).Left = 4255
'       Frame8(3).Left = 5350
'       Frame8(4).Left = 6700
'       Frame8(5).Left = 8050
'       Frame8(6).Left = 8620
'       Frame8(7).Left = 9400
'       Frame8(8).Left = 10580
'       Frame8(9).Left = 11800
'       Frame8(10).Left = 13060
''
'       Frame8(0).ZOrder 0
'       Frame8(1).ZOrder 0
'       Frame8(2).ZOrder 0
'       Frame8(3).ZOrder 0
'       Frame8(4).ZOrder 0
'       Frame8(5).ZOrder 0
'       Frame8(6).ZOrder 0
'       Frame8(7).ZOrder 0
'       Frame8(8).ZOrder 0
'       Frame8(9).ZOrder 0
'       '
'       Frame8(10).ZOrder 0
       '
       'Label6.Left = 4200
       'Label5.Left = 6850
       'Label13.Left = 8350
       Label13.Visible = True
   End If
End Sub

Private Sub Check2_Click()

   If Check2.Value = 1 Then
     'Picture1.Visible = True
      mfgCuentaCorrienteAHP.Visible = True
      Me.Width = 15915
      'Picture1.Left = 13400
      'Line2.X2 = 9800
      Picture3.Left = 14880
      Frame2.Width = 14715
      msf2.Width = 14280
      Command43.Left = 14400
      Call CARESCUEPRO_ahp
   Else
     'Picture1.Visible = False
     mfgCuentaCorrienteAHP.Visible = False
     Frame2.Width = 10965
     'Line2.X2 = 9800
     msf2.Width = 10560
     Picture3.Left = 11265
     Picture3.ZOrder 0
     Picture4.Height = 2385
     Me.Width = 12270
     Command43.Left = 10680
   End If
   
   
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text10 = ""
    Label26 = ""
    '
    List1.Clear
    List2.Clear
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
    '
    'Call SELECHO("PROVED1")
    'NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    NC% = ListaSeleccionProveedores()
    If NC% > 0 Then
        Text1.TEXT = TRIM$(Str$(NC%))
    End If
    '
End Sub

Private Sub Command20_Click()
   Dim LIST As ListBox
   Set LIST = Me.List1
   If List3.Visible = True Then Set LIST = Me.List3
   
   Command20.Enabled = False ' BUSQUEDA POR NUMERO DE SERIE
   BAND% = 0 ' BANDERA PARA VERIFICAR SI ENCONTRO EL ARGUMENTO
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda por Texto en Cta. Corriente", ARGU$)))
   If ARGU$ = "" Then GoTo 99
   
   Screen.MousePointer = 11
   If LIST.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For i& = 1 To LIST.ListCount
'        If TRIM$(ARGU$) = TRIM$(List1.List(I& - 1)) Then
        If InStr((TRIM$(UCase$(Left$(LIST.LIST(i& - 1), 77)))), ARGU$) Then
          LIST.ListIndex = i& - 1
          LIST.SetFocus
          BAND% = 1
          Exit For
        Else
        End If
     Next i&
   End If
   If BAND% < 1 Then Call MENSERR(24, "Argumento No Encontrado")
   Screen.MousePointer = 1
99 Command20.Enabled = True

End Sub


Private Sub ComposicionDeuda_Click()
   Call FAPEPAG07.ComposicionDeudaAFecha
End Sub

Private Sub ComposicionDeudaProveedor_Click()
     Call FAPEPAG07.ComposicionDeudaPorProveedorFecha
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  '\\\OT-07-0046-RG-01/02/07///
   If KeyCode = vbKeyF3 Then
       If ARGU$ <> "" Then
          VUELTA% = 0
          INI& = List1.ListIndex + 2
    '
10        For U& = INI& To List1.ListCount
              If InStr((UCase$(Left$(List1.LIST(U& - 1), 77))), ARGU$) Then
                 List1.ListIndex = U& - 1
                 Exit Sub
              End If
          Next U&
    '
          If VUELTA% < 1 Then
             VUELTA% = 1
             INI& = 1
             GoTo 10
          End If
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If
   '\\\OT-07-0046-RG-FIN///
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
  '\\\OT-07-0046-RG-01/02/07///
   If KeyCode = 114 Then
       If ARGU$ <> "" Then
          VUELTA% = 0
          INI& = List3.ListIndex + 2
    '
10        For U& = INI& To List3.ListCount
'              If InStr(UCase$(List3.List(U& - 1)), ARGU$) > 0 Then
              If InStr(TRIM$(UCase$(Left$(List3.LIST(U& - 1), 77))), ARGU$) Then
                 List3.ListIndex = U& - 1
                 Exit Sub
              End If
          Next U&
    '
          If VUELTA% < 1 Then
             VUELTA% = 1
             INI& = 1
             GoTo 10
          End If
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If
   '\\\OT-07-0046-RG-FIN///
End Sub


Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    Screen.MousePointer = 11
    '
    If ULTREG&("PROVED1") > 0 Then
        X$ = REGLEIDO$("PROVED1", 1)      ' fuerza la
        Call GRAREG("PROVED1", X$, 1)     ' reindexacion
        Call CIERRARCH("PROVED1")         ' del archivo
    End If
    '
    Screen.MousePointer = 11
    '
    If Int(XVALO(Text1.TEXT)) > 0 Then
        If Int(XVALO(Text1.TEXT)) <= 32767 Then
            NCMIN% = Int(XVALO(Text1.TEXT))
            NCMAX% = NCMIN%
            GoTo 10
        End If
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("PROVED1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
10  Screen.MousePointer = 1
    RESCUP07.Text10.TEXT = TRIM$(Str$(NCMIN%))
    RESCUP07.Text9.TEXT = rasocli$((-1) * NCMIN%)
    RESCUP07.Text8.TEXT = TRIM$(Str$(NCMAX%))
    RESCUP07.Text7.TEXT = rasocli$((-1) * NCMAX%)
    '
    HOO% = HOY(HO$)
    RESCUP07.Text6.TEXT = "01" + Mid$(HO$, 3)
    RESCUP07.Text5.TEXT = HO$
    '
    If mfgCuentaCorrienteAHP.Visible = True Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call ExportaAExcelVistaAmpliada(AREXPO$)
        End If
    Else
      RESCUP07.Show 1
    End If
    '
99  Command3.Enabled = True
End Sub

Sub Command4_Click()
    '
    Command4.Enabled = False
    If MONRESU% > 1 Then
      Call MENSERR(24, "Proceso Imposible - Cambie Moneda a '" + TRIM$(DEMONECO$(1)) + "'")
      GoTo 99
    End If

    NC% = XVALO(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       GoTo 99
    End If
    '
    'If TICUEPRO%((-1) * NC%) <> 0 Then
    '   Call MENSERR(24, "Proceso Imposible - Cuenta de Gastos")
    '   GoTo 99
    'End If
    '
    If CONTROLGRABA% = 1 Then GoTo 80
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       GoTo 99
    End If
    '
80  Call RAPLIPRO(NC%)
    Text1 = ""
    Text1 = TRIM$(Str$(NC%))
    '
99  Command4.Enabled = True
End Sub

Private Sub Command43_Click()
   Screen.MousePointer = 11
   Call CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(XVALO(Text1), RECUEP07.msf2, MONRESU%)
   Screen.MousePointer = 1
End Sub

Private Sub Command5_Click()
    NCLI$ = TRIM$(Text1.TEXT)
    If NCLI$ <> "" Then
        Command5.Enabled = False
        Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Proveedor")
    End If
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    CLIPROACTIVO = (-1) * XVALO(Text1)
    Call FACOCOMP07.Command3_Click     ' Imputación de Créditos
    Text1 = ""
    Text1 = TRIM$(Str$(Abs(CLIPROACTIVO)))
    CLIPROACTIVO = 0
    Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    MONRESU% = 1
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    TICAMBIS = FORMATNUM$(TICAMONE(MONRESU%), "G9.4")
    '
    Me.Width = 11145 + 1165
    Frame2.Width = 10965
    msf2.Width = 10560
    Command43.Left = 10680
    Picture3.Left = 10100 + 1165
    Picture3.ZOrder 0
    Label6.Left = 6170
    Label5.Left = 9195
    '
    ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
    Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A22;Número/A18;Debe/F12.2;Haber/F12.2;Saldo/F14.2", Me)
    For i = 1 To 6
         mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
    Next i
    mfgCuentaCorriente.Font.Size = 9
    '
    ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA
    Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A20;Número/A18;Emisión/D8;Importe/F16.2;Vencimiento/D8;Saldo ($)/F16.2", Me)
    For i = 1 To 6
         mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
    Next i
    mfgComposicionDeuda.Font.Size = 9
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       '
       Frame6.Visible = True
       Label6.Left = 5160
       Label5.Left = 7680
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
       Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A14;Número/A18;Debe/F10.2;Haber/F10.2;Saldo/F12.2;Mon/A3;Cotiz/F7.2", Me)
       For i = 1 To 8
            mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
       Next i
       mfgCuentaCorriente.Font.Size = 9
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA EN MODALIDAD BIMONETARIA
       mfgComposicionDeuda.Cols = 1
       mfgComposicionDeuda.Rows = 1
       Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A14;Número/A18;Emisión/D8;Importe/F12.2;Vencimiento/D8;Saldo ($)/F12.2; Saldo (U$)/F12.2", Me)
       For i = 1 To 7
            mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
       Next i
       mfgComposicionDeuda.Font.Size = 9
       
       Label9.Caption = "Fecha     Documento   Número                  Debe        Haber        Saldo   ME  Cotiz."
       Label9.Font.Size = 9
       
       Label13.Visible = True
       'If InStr(EPAC$, "AHP") > 0 Then Call pantallaAhp
       If InStr(EPAC$, "AHP") > 0 Then
            ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
            Call CARGA_ENCABEZADO(mfgCuentaCorrienteAHP, "Fecha/D8;Documento/A10;Número/A18;Debe/F9.2;Haber/F9.2;Saldo/F10.2;Mon/A3;Cotiz/F7.2;Val.Compra/F9.2;G.Varios/F9.2;Valor Real/F9.2;Diferencia/F9.2", Me)
            For i = 1 To 12
                 mfgCuentaCorrienteAHP.COL = i: mfgCuentaCorrienteAHP.CellFontBold = False: mfgCuentaCorrienteAHP.CellFontSize = 8
            Next i
            mfgCuentaCorrienteAHP.Font.Size = 8
       End If
   End If
 
   RECUEP07.Height = 8650
   If InStr(EPAC$, "AHP") > 0 Then
     RECUEP07.Height = 9420 'MUESTRA LA LISTA DE FACTURAS CON LOS VALORES SEGUN ORDEN DE COMPRA.
     RECUEP07.Picture3.Height = 9435
     Call CARGA_ENCABEZADO(msf2, "Fecha/D8;Factura/A18;Gasto Aplic./F12.22", SELFACTUMAT)
     For i = 1 To 3
        msf2.COL = i: msf2.CellFontBold = False: msf2.CellFontSize = 8
     Next i
     msf2.Font.Size = 8
   End If
   '
   TamanioFuente% = 10
   If MODOMONE$ = "MULTIMON" Then
       List1.Font.Size = 9
       List2.Font.Size = 9
       List3.Font.Size = 9
   End If
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
   'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
   'Call REMPLAGRID(Me, GRID1, Picture2, Label4, List2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub GRID2_DblClick()
   NGRID = GRID2.Row
   NLIST = List1.ListCount
   If (NLIST) >= NGRID - 1 Then
     List1.ListIndex = NGRID - 1
     List1_DblClick
   End If

End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDIACRE")
   CPROX$ = TRIM$(VALACT1$("INDIACRE"))
   If CPROX$ <> "" Then
     NCXI% = NUMEPRO%(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&((-1) * NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
       End If
     End If
   End If
End Sub

'Private Sub Label8_Change()
'    Label12 = FORMATNUM$(XVALO(Label13) + XVALO(Label8), "F12.2")
'End Sub

Private Sub List1_DblClick()
   '
   RELIS$ = List1.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 11))
   If UCase$(TIDOCUM$) <> "O.PAGO" And UCase$(TIDOCUM$) <> "O/P F.FIJO" Then
      Call MUEFACO
      Exit Sub
   End If
   '
   TIDOCUM1$ = "Orden de Pago"
   NPX& = XVALO(Mid$(RELIS$, 32, 8))
   NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
   If UCase$(TIDOCUM$) = "O/P F.FIJO" Then
        TIDOCUM1$ = "Liquidacion de Fondo Fijo"
        NPX& = XVALO(Mid$(RELIS$, 25, 8))
        NUDOCU$ = TIDOCUM1$ & " " & TRIM$(Str$(NPX&))
   End If
   
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM1$) > 4 Then
       PPXX% = InStr(XX$, "Nro.")
       If PPXX% > 0 Then
         XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
       End If
       If InStr(XX$, NUDOCU$) > 4 Then
         DOCUNU& = CVS(Left$(XX$, 4))
         GoTo 5
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
3  Call COMUNI("Imposible Mostrar Documento")
   Exit Sub
   '
5  Call MUESTRADOC(DOCUNU&)
   '
End Sub

Private Sub List3_DblClick()
   RELIS$ = List3.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 11))
   If UCase$(TIDOCUM$) <> "O.PAGO" Then
      Call MUEFACO_AHP
      Exit Sub
   End If
   '
   TIDOCUM1$ = "Orden de Pago"
   NPX& = XVALO(Mid$(RELIS$, 32, 8))
   NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM1$) > 4 Then
       PPXX% = InStr(XX$, "Nro.")
       If PPXX% > 0 Then
         XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
       End If
       If InStr(XX$, NUDOCU$) > 4 Then
         DOCUNU& = CVS(Left$(XX$, 4))
         GoTo 5
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
3  Call COMUNI("Imposible Mostrar Documento")
   Exit Sub
   '
5  Call MUESTRADOC(DOCUNU&)

End Sub


Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 2 Then
    
    RELIS$ = List3.TEXT
    TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 11))
    DOCU$ = TRIM$(Mid$(RELIS$, 22, 18))
    MONRES% = XVALO(Mid$(RELIS$, 140, 2))
   
    VALCOMPRA# = ValorCompraFacturaExpo#(XVALO(Text1), DOCU$, MONRES%) ' SOLO VALOR ORIGINAL
    VALCOMPRA_AJUSTADO# = ValorCompraFacturaExpo#(XVALO(Text1), DOCU$, MONRES%, 1) ' valor con ajuste
   '
5   AJUSTEX$ = XVALO(InputBox$("Valor Original Orden de Compra: " & TRIM$(FORMATNUM$(VALCOMPRA#, "F12.2")) & vbCrLf & vbCrLf & " - Ingresar Valor Ajustado", "Factura: " & DOCU$, TRIM$(FORMATNUM$(VALCOMPRA_AJUSTADO#, "F12.2"))))
    If XVALO(AJUSTEX) = 0 Then Exit Sub
    '
    VALORAJUSTADO# = XVALO(AJUSTEX)
    If VALORAJUSTADO# > VALCOMPRA# Then
       AJUSTEAREALIZAR# = VALORAJUSTADO# - VALCOMPRA#
    Else
        AJUSTEAREALIZAR# = VALORAJUSTADO# - VALCOMPRA#
    End If
    CONDI$ = "Nume_Cli = " & XVALO(Text1) & " AND TIPO_FACTURA = 'M'"
    CONDI$ = CONDI$ + " AND FACTURAPROV = '" & DOCU$ & "'"
    Call ACTUREGISTRO("FACTURAYGASTOS", "AJUSTE_VAL_COMPRA", CONDI$, FORMATNUM$(AJUSTEAREALIZAR#, "F12.2"), REGAF&, "NOMESS")
    Call CARESCUEPRO_ahp
   End If
End Sub


Private Sub mfgCuentaCorriente_DblClick()
   '
   FilaSeleccionada& = mfgCuentaCorriente.Row
   TIDOCUM$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2))
   If UCase$(TIDOCUM$) <> "ORDEN DE PAGO" And UCase$(TIDOCUM$) <> "O/P F.FIJO" Then
      Call MUEFACO
      Exit Sub
   End If
   '
   NUDOCU$ = XVALO(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 11, 8))
   FE% = FECHANUM(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
   If UCase$(TIDOCUM$) = "O/P F.FIJO" Then
        TIDOCUM$ = "Liquidacion de Fondo Fijo"
        NUDOCU$ = XVALO(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 4, 8))
   End If
   
   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU$, NUDOCU$)
   Call MUESTRADOC(DOCUNU&)
   '
End Sub


Private Sub mfgCuentaCorrienteAHP_DblClick()
   '
   FilaSeleccionada& = mfgCuentaCorriente.Row
   TIDOCUM$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2))
   If UCase$(TIDOCUM$) <> "O.PAGO" And UCase$(TIDOCUM$) <> "O/P F.FIJO" Then
      Call MUEFACO_AHP
      Exit Sub
   End If
   '
   TIDOCUM$ = "Orden de Pago"
   NUDOCU$ = XVALO(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3), 11, 8))
   FE% = FECHANUM(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
   If UCase$(TIDOCUM$) = "O/P F.FIJO" Then
        TIDOCUM$ = "Liquidacion de Fondo Fijo"
   End If
   
   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU$, NUDOCU$)
   Call MUESTRADOC(DOCUNU&)
   '
End Sub


Private Sub mfgCuentaCorrienteAHP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    '
    If Button = 2 Then
         FilaSeleccionada& = mfgCuentaCorrienteAHP.Row
         TIDOCUM$ = TRIM$(mfgCuentaCorrienteAHP.TextMatrix(FilaSeleccionada&, 2))
         DOCU$ = TRIM$(mfgCuentaCorrienteAHP.TextMatrix(FilaSeleccionada&, 3))
         MONRES% = TRIM$(mfgCuentaCorrienteAHP.TextMatrix(FilaSeleccionada&, 7))
        
         VALCOMPRA# = ValorCompraFacturaExpo#(XVALO(Text1), DOCU$, MONRES%) ' SOLO VALOR ORIGINAL
         VALCOMPRA_AJUSTADO# = ValorCompraFacturaExpo#(XVALO(Text1), DOCU$, MONRES%, 1) ' valor con ajuste
        '
5        AJUSTEX$ = XVALO(InputBox$("Valor Original Orden de Compra: " & TRIM$(FORMATNUM$(VALCOMPRA#, "F12.2")) & vbCrLf & vbCrLf & " - Ingresar Valor Ajustado", "Factura: " & DOCU$, TRIM$(FORMATNUM$(VALCOMPRA_AJUSTADO#, "F12.2"))))
         If XVALO(AJUSTEX) = 0 Then Exit Sub
         '
         VALORAJUSTADO# = XVALO(AJUSTEX)
         If VALORAJUSTADO# > VALCOMPRA# Then
            AJUSTEAREALIZAR# = VALORAJUSTADO# - VALCOMPRA#
         Else
             AJUSTEAREALIZAR# = VALORAJUSTADO# - VALCOMPRA#
         End If
         CONDI$ = "Nume_Cli = " & XVALO(Text1) & " AND TIPO_FACTURA = 'M'"
         CONDI$ = CONDI$ + " AND FACTURAPROV = '" & DOCU$ & "'"
         Call ACTUREGISTRO("FACTURAYGASTOS", "AJUSTE_VAL_COMPRA", CONDI$, FORMATNUM$(AJUSTEAREALIZAR#, "F12.2"), REGAF&, "NOMESS")
         Call CARESCUEPRO_ahp
    End If
    '
End Sub




Private Sub mnuBuscar_Click()
   Call Command20_Click
End Sub

Private Sub mnuImprimirCtaCte_Click()
   Call Command3_Click
End Sub

Private Sub mnuImputacionCredAnt_Click()
    Call Command6_Click
End Sub

Private Sub mnuNotaProveedor_Click()
  Call Command5_Click
End Sub

Private Sub mnuReapli_Click()
  Call Command4_Click
End Sub

Private Sub MONEMIS_Change()
   TICAMBIS = FORMATNUM$(TICAMONE(MONRESU%), "G9.4")
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If MONRESU% <> 1 And InStr(EPAC$, "AHP") > 0 Then 'VISTA AMPLIADA AHP
      Check2.Visible = True
   Else
      Check2.Visible = False
   End If
   If Check2.Value = 1 Then
     Check2.Value = 0
   End If
   
End Sub

Private Sub MONEMIS_DblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   If Abs(XVALO(Text1)) < 32767 Then
     NCX% = XVALO(Text1)
     If NCX% > 0 Then
       If ECOARCH$("PROVED1", MKI$(NCX%)) <> "" Then
         GoTo 20
       End If
     End If
   End If
   Exit Sub
   '
20 MONEMIAR% = MONRESU%
   MONRESU% = MONRESU% + 1
   If MONRESU% > CAMON% Then MONRESU% = 1
   MONEMIS = DMONE$(MONRESU%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONRESU% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   Call CARESCUEPRO
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "AHP") > 0 Then
        'Call CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(XVALO(Text1), RECUEP07.MSF2, MONRESU%)
        Call CARESCUEPRO_ahp
   End If
   
   '
End Sub

Private Sub msf2_Click()
    'AL REALIZAAR DOBLE CLICK SE POSICIONA CON LA IDEM EN LA LISTA DE ARRIBA.
    i& = msf2.MouseRow
    J& = msf2.MouseCol
    If i& = 0 Or i& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = i&
    msf2.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If

   BAND% = 0 ' BANDERA PARA VERIFICAR SI ENCONTRO EL ARGUMENTO
10 ARGU$ = UCase$(TRIM$(msf2.TextMatrix(REG&, 2)))
   If ARGU$ <> "" Then
    Screen.MousePointer = 11
    If List1.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For i& = 1 To List1.ListCount
        If TRIM$(ARGU$) = TRIM$(Mid$(List1.LIST(i& - 1), 22, 18)) Then
          List1.ListIndex = i& - 1
          BAND% = 1
          Exit For
        Else
        End If
     Next i&
    End If
   End If
   
'   If BAND% < 1 Then Call MENSERR(24, "Argumento No Encontrado")
   Screen.MousePointer = 1

End Sub

Private Sub MSF2_DblClick()
    i& = msf2.MouseRow
    J& = msf2.MouseCol
    If i& = 0 Or i& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = i&
    msf2.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    
    FACTU$ = UCase$(TRIM$(msf2.TextMatrix(REG&, 2)))
    J& = 0
    Campos$ = "Fecha/D8;Prov./F6;R.Social/A24;Factura/A20;Imp.Fact.Mon.Orig./F20.2;Imp.Gastos $/F20.2"
    
    'MUESTRA EL DETALLE DE GASTOS APLICADOS LA  FACTURA PARAMETRO.
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    FRMZELECHO.Caption = "Gastos en $ de Factura: " & FACTU$
    '
    FRMZELECHO.MSF1.Rows = 1

    SQLX$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE FACTURAPROV_ASOC = '" & FACTU$ & "'"
    SQLX$ = SQLX$ + " AND NUME_CLI_ASOC = " & XVALO(Text1)
    SQLX$ = SQLX$ + " AND TIPO_FACTURA = 'G' "
    SQLX$ = SQLX$ + " AND ABS(IMPO_GASTO) > 0 "
    
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
             '
             J& = J& + 1
             FRMZELECHO.msf2.Rows = J& + 1
             FRMZELECHO.msf2.TextMatrix(J&, 1) = FECHATEX$(CVINT(!FECHEMISI))
             FRMZELECHO.msf2.TextMatrix(J&, 2) = XVALO(!nume_cli)
             FRMZELECHO.msf2.TextMatrix(J&, 3) = rasocli$(-1 * XVALO(!nume_cli))
             FRMZELECHO.msf2.TextMatrix(J&, 4) = SAFETEXT$(!facturaprov)
              TICAM# = XVALO(!TIPOCAM)
              MONEX% = XVALO(!MONEVAL)
              GASTOX# = XVALO(!IMPO_GASTO)
              GASTOX1$ = FORMATNUM$(GASTOX#, "F12.2") 'GASTO MONEDA ORIGINAL
              GASTOXPESOS$ = FORMATNUM$(GASTOX# * TICAM#, "F12.2") 'GASTO CONVERTIDO A $
             FRMZELECHO.msf2.TextMatrix(J&, 5) = GASTOX1$
             FRMZELECHO.msf2.TextMatrix(J&, 6) = GASTOXPESOS$
            .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Show 1
End Sub

Private Sub ObligacionesPagarPorFecha_Click()
    FAPEPAG07.ObligacionesPagarPorFecha_Click
End Sub

Private Sub Text1_Change()
   If XVALO(Text1) >= 1 Then
     If REGICLI&((-1) * XVALO(Text1)) > 0 Then
       Label26 = CODICLI$((-1) * XVALO(Text1))
'       Call CARESCUEPRO
        Timer2.Enabled = True
       Exit Sub
     End If
   End If
   '
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text10 = ""
   Label26 = ""
   '
   List1.Clear
   List2.Clear
   mfgCuentaCorriente.Rows = 1
   mfgCuentaCorrienteAHP.Rows = 1
   mfgComposicionDeuda.Rows = 1
   '
   '
   'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
   'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
   '
End Sub
Private Sub Timer2_Timer()
    Static TextA$, FUEAMOSTRAR%
    '
    'If Text1 = TEXTA$ Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '
    FUEAMOSTRAR% = 1
    Call CARESCUEPRO
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    If InStr(EPAC$, "AHP") > 0 Then
      RECUEP07.Height = 9825 'MUESTRA LA LISTA DE FACTURAS CON LOS VALORES SEGUN ORDEN DE COMPRA.
      RECUEP07.Picture3.Height = 9435
      'Call CARGA_ENCABEZADO(MSF2, "Fecha/D8;Factura/A26;Importe/F20.2;Gastos Aplicado/F20.2;LIBRE/A0", SELFACTUMAT)
      'Call CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(XVALO(Text1), RECUEP07.MSF2, MONEMI%)
      Call CARESCUEPRO_ahp
    End If
    Timer2.Enabled = False
    '
    FUEAMOSTRAR% = 0
    '
End Sub
Sub CARGA_FACTURAS_PROVEEDOR_VALCOMPRA(NPROV, MSF As MSFlexGrid, MONEX%)
    '
    If NPROV < 1 Then Exit Sub
    MSF.Rows = 1
    '
    msf2.Rows = 1
    Label8 = FORFAC07.Label43 ' EL TOTAL DE LA FACTURA
    Diferencia# = 0
    '
    CONDI$ = "Nume_Cli = " & NPROV & " AND TIPO_FACTURA = 'M' "
    FIN& = CantRegistros("FACTURAYGASTOS", CONDI$, "NOMESS")
    Screen.MousePointer = 11
    If MONEX% < 1 Then MONEX% = 1
    SQLX1$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    SQLX1$ = SQLX1$ + " WHERE " & CONDI$ & " order by fechemisi"
   ' SQLX1$ = SQLX1$ + " AND VAL_COMPRA > 0 AND VAL_COMPRA <> IMPORTE_FACTURA" ' NO MUESTRO  LAS QUE ES EL MISMO VALOR P Q NO TIENE SENTIDO
    Set RS1 = READSET(SQLX1$)
    With RS1
       Do While Not .EOF
         'ESCRIBO SIEMPRE LAS 3 CAMPOS
         J& = J& + 1
         Call TRACE(20, J&, FIN&)
         MSF.Rows = J& + 1
         
         MSF.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHEMISI))
           facturaprov$ = TRIM$(SAFETEXT$(!facturaprov))
         MSF.TextMatrix(J&, 2) = facturaprov$
           CONDI$ = "TIPO_FACTURA = 'G' AND FACTURAPROV_ASOC = '" & facturaprov$ & "'"
         APLICADO# = GastosAplicados(facturaprov$, NPROV, MONEX%) ' TRAE LOS GASTOS APLICADOS CONVERTIDOS A DOLARES.
'         IMPOFAC# = XVALO(!IMPORTE_FACTURA)
'         VALOR_COMPRA# = XVALO(!VAL_COMPRA)
'         MSF.TextMatrix(j&, 3) = FORMATNUM$(VALOR_COMPRA#, "F12.2") 'EL IMPORTE DE LA COMPRA SIEMPRE SE GUARDA EN $

         MSF.TextMatrix(J&, 3) = TRIM$(FORMATNUM$(APLICADO#, "F12.2"))
'         MSF.TextMatrix(j&, 5) = TRIM$(FORMATNUM$(IMPOFAC#, "F12.2"))
'         DIFERENCIA# = DIFERENCIA# + (VALOR_COMPRA# - IMPOFAC#)

         .MoveNext
       Loop
    End With
    Screen.MousePointer = 1
    RS1.Close
    Set RS1 = Nothing
'    Label8 = FORMATNUM$(DIFERENCIA#, "F12.2")
'    Call Label8_Change
End Sub
Private Sub Text1_DblClick()
    '
    'Call SELECHO("PROVED1")
    'NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    NC% = ListaSeleccionProveedores()
    Text1.TEXT = TRIM$(Str$(NC%))
    '
End Sub

Private Sub Text10_Click()
   Text1.SetFocus
End Sub

Private Sub Text2_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text3_Click()
   Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text5_Click()
   Text1.SetFocus
End Sub

Private Sub Text6_Click()
   Text1.SetFocus
End Sub

Private Sub Text7_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text8_Click()
   Text1.SetFocus
End Sub

Private Sub Text9_CLICK()
   Text1.SetFocus
End Sub

Sub MUEFACO()
   '
   FilaSeleccionada& = mfgCuentaCorriente.Row
   TIPODOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2)) 'FC-A
   NUDOCOR$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3)) 'NUMERO DE DOC COMP
   NP1& = XVALO(Text1)
   FECHDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
   'FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   FECDOC1 = FETEXCOR1$(FECHANUM(TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1)))) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   'jandy sql
   Dim SHOFACOTMP As ADODB.Recordset
   Set SHOFACOTMP = New ADODB.Recordset
   SHOFACOTMP.CursorLocation = adUseClient
   SQLZ$ = " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLZ$ = SQLZ$ + " AND NuProve =  " & NP1& & " "
   SQLX$ = " SELECT * FROM CAJABANC " + SQLZ$
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
25 On Error Resume Next
   SHOFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = TIPODOCU$ + " " + !CODICOM_LAR + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '\\\OT-08-988-29/10/08-FL
     SHOWFACO.Label19 = ""
     '\\\OT-08-988-FIN-FL
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(XVALO(!CUECONTAI))
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DESCRIMOV
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.List1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !Idebecomp) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !Idebecomp)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !OBSERGEN
          '\\\OT-08-988-29/10/08-FL
          'SE CAPTURA UN POSIBLE ERROR PRODUCTO QUE EL CAMPO IDENALTA CONTENGA INFORMACION INVALIDA
          On Error Resume Next
          USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
          On Error GoTo 0
          SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
          '\\\OT-08-988-FIN-FL
        End If
        '
     .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub

Sub MUEFACO_AHP()
   '
   FilaSeleccionada& = mfgCuentaCorriente.Row
   TIPODOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2)) 'FC-A
   NUDOCOR$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3)) 'NUMERO DE DOC COMP
   NP1& = XVALO(Text1)
   FECHDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
   'FECDOC1 = CDate(TRIM$(Mid$(List1.TEXT, 1, 9))) 'FECHA DEL DOC.
   FECDOC1 = FETEXCOR1$(FECHANUM(TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1)))) 'FECHA DEL DOC.
   
'   TIPODOCU$ = TRIM$(Mid$(List3.TEXT, 10, 17)) 'FC-A
'   NUDOCOR$ = TRIM$(Mid$(List3.TEXT, 22, 18)) 'NUMERO DE DOC COMP
'   NP1& = XVALO(Text1)
'   FECHDOCU$ = TRIM$(Left$(List3.TEXT, 9))
'   'FECDOC1 = CDate(TRIM$(Mid$(list3.TEXT, 1, 9))) 'FECHA DEL DOC.
'   FECDOC1 = FETEXCOR1$(FECHANUM(TRIM$(Mid$(List3.TEXT, 1, 9)))) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   'jandy sql
   Dim SHOFACOTMP As ADODB.Recordset
   Set SHOFACOTMP = New ADODB.Recordset
   SHOFACOTMP.CursorLocation = adUseClient
   SQLZ$ = " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLZ$ = SQLZ$ + " AND NuProve =  " & NP1& & " "
   SQLX$ = " SELECT * FROM CAJABANC " + SQLZ$
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
25 On Error Resume Next
   SHOFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = TIPODOCU$ + " " + !CODICOM_LAR + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '\\\OT-08-988-29/10/08-FL
     SHOWFACO.Label19 = ""
     '\\\OT-08-988-FIN-FL
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CUECONTAI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DESCRIMOV
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.List1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or IsEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !Idebecomp) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !Idebecomp)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !OBSERGEN
          '\\\OT-08-988-29/10/08-FL
          'SE CAPTURA UN POSIBLE ERROR PRODUCTO QUE EL CAMPO IDENALTA CONTENGA INFORMACION INVALIDA
          On Error Resume Next
          USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
          On Error GoTo 0
          SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
          '\\\OT-08-988-FIN-FL
        End If
        '
     .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub

Sub CARESCUEPRO()
    '
    Static ENPROCESO%
    '
    Text1.Enabled = False
5   If ENPROCESO% > 0 Then Exit Sub
    ENPROCESO% = 1
    '
10  Screen.MousePointer = 11
    NC% = XVALO(Text1)
    NCX7% = XVALO(Text1)
    NICONFLIM% = NIVCONFI%
    Command4.Enabled = True
    '
    Call APERBASDAT("CABAN")
    '
    If NC% > 0 Then
        NC1% = (-1) * NC%
        Label26 = CODICLI$(NC1%)
        Text2 = rasocli$(NC1%)
        Text3 = DOMICLI$(NC1%)
        Text4 = CPOSCLI$(NC1%)
        Text5 = LOCACLI$(NC1%)
        Text7 = TEL1CLI$(NC1%)
        Text10 = TRIM$(FORMATNUM$(SALDPRO#(NC%), "G14.2"))
        TCUE% = TICUEPRO%(NC1%)
        If TCUE% = 1 Then
            Text6.TEXT = "GASTOS"
            Picture4.Height = 5340
            Frame7(0).Height = 5340
            Frame7(1).Height = 5340
            Frame7(2).Height = 5340
            Frame7(3).Height = 5340
            Frame7(4).Height = 5340
            Frame7(5).Height = 5340
            List1.Height = 4750
            mfgCuentaCorrienteAHP.Height = 5500
            mfgCuentaCorriente.Height = 5500
          Else
            Text6.TEXT = "CORRIENTE"
            Picture4.Height = 2430
            Frame7(0).Height = 2415
            Frame7(1).Height = 2415
            Frame7(2).Height = 2415
            Frame7(3).Height = 2415
            Frame7(4).Height = 2415
            Frame7(5).Height = 2415
            List1.Height = 2085
            mfgCuentaCorrienteAHP.Height = 2300
            mfgCuentaCorriente.Height = 2450
        End If
        DoEvents
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    'Call REMPLAGRID(Me, GRID1, Picture2, Label9, List2)
    'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    SALDO# = 0: SARRAS# = 0
    FilaCuentaCorriente& = 0
    MaximoDeFilasGrid& = 1000
    mfgComposicionDeuda.Rows = 1: Label6 = "": Label5 = "": Label13 = ""
    mfgCuentaCorriente.Rows = 1: mfgCuentaCorriente.Rows = MaximoDeFilasGrid&
    '
    RFF$ = RECFILT$("CUECOPRO", 1, MKI$(NC%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOPRO", RELE&)
      TIPOS$ = Mid$(XX$, 5, 16)
      IMPO# = CVD(Mid$(XX$, 25, 8))
      SALDO# = SALDO# - IMPO#
      FEX = CVI(Mid$(XX$, 3, 2))
      NROL& = CVS(Mid$(XX$, 5, 4))
      NROX$ = Mid$(XX$, 33, 24)
      '
      Z$ = Space$(96)
      Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
      Call REPLA(Z$, TIPOS$, 11, 11)
      Call REPLA(Z$, NROX$, 22, 18)
      IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
      SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
      If IMPO# > 0 Then
          Call REPLA(Z$, IMPTEX$, 40, 11)
        ElseIf IMPO# < 0 Then
          Call REPLA(Z$, IMPTEX$, 53, 11)
      End If
      Call REPLA(Z$, SALTEX$, 64, 13)
      
      FilaCuentaCorriente& = FilaCuentaCorriente& + 1
      mfgCuentaCorriente.Rows = mfgCuentaCorriente.Rows + 1
      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
      If IMPO# > 0 Then
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = TRIM$(IMPTEX$)
      ElseIf IMPO# < 0 Then
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = TRIM$(IMPTEX$)
      End If
      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = TRIM$(SALTEX$)
      'List1.AddItem z$
    Next KU&
    DoEvents
    
    AjustaCentavos% = 0: ImporteAjustaCentavos# = 0
    If InStr(1, EMPREAC$, "AHP") > 0 Then
        DEB# = ROUND#(XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", "NUPROVE=" & Text1 & " AND (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP')")), 2)
        HAB# = ROUND#(XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", "NUPROVE=" & Text1 & " AND (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP')")), 2)
        SALDO1# = DEB# - HAB#
        
        DESACRE# = ROUND#(XVALO(VALOBADA("SACREPEN", "SUM(SaldDebe)", "NUPROVE=" & Text1)), 2)
        HASACRE# = ROUND#(XVALO(VALOBADA("SACREPEN", "SUM(SaldAcre)", "NUPROVE=" & Text1)), 2)
        
        SALDO2# = DESACRE# - HASACRE#
        
        ImporteAjustaCentavos# = SALDO2# - SALDO1#
        
    End If
    
    ' OBTIENE LA CANTIDAD DE REGISTROS DE LA CONSULTA DE CUENTA CORRIENTE
    ' PARA AGREGAR DICHA CANTIDAD DE FILAS A LA GRILLA mfgCuentaCorriente
    CondicionSql$ = " NuProve = " & NC% & " AND (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP')"
    CantidadRegistros& = CantRegistros("CAJABANC", CondicionSql$)
    mfgCuentaCorriente.Rows = CantidadRegistros& + 1

    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CondicionSql$
    SQLX$ = SQLX$ + " ORDER BY FECHCONTA, FEHOREAL, CODICOM_LAR ASC "
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    If NCX7% <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          DocumentoProveedor$ = SAFETEXT$(!CODICOM_LAR)
          TIX$ = Left$(DocumentoProveedor$, 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "RC" Then TIPOS$ = "Recibo"
                              If TIX$ = "NV" Then TIPOS$ = "Nota de Venta"
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = XVALO(ROUND#(!Idebecomp, 2)) - XVALO(ROUND#(!IHABECOMP, 2))
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             Command4.Enabled = False            ' NO PERMITE REAPLICAR
             GoTo 899
          End If
          '
          ' AGREGADO PARA AGRUPAR O/P'S MULTIPLES A EGRESOS Y GASTOS - 27/05/06
          If mfgCuentaCorriente.Rows > 0 Then
            If TIX$ = "OP" Or TIX$ = "FF" Then
              If DocumentoProveedor$ = NROX$ Then
                If CVINT%(!FECHEMISI) = FEX Then
                  SALDO# = SALDO# - IMPO#
                  'IMPO# = IMPO# + XVALO(Mid$(z$, 40, 11)) - XVALO(Mid$(z$, 53, 11))
                  IMPO# = IMPO# + XVALO(mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4)) - XVALO(mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5))
                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
                  Call REPLA(Z$, Space$(24), 40, 24)
                  If IMPO# > 0 Then
                      'Call REPLA(z$, IMPTEX$, 40, 11)
                    ElseIf IMPO# < 0 Then
                      'Call REPLA(z$, IMPTEX$, 53, 11)
                  End If
                  'Call REPLA(z$, SALTEX$, 64, 13)
                  '
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
                  If IMPO# > 0 Then
                      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
                  ElseIf IMPO# < 0 Then
                      mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
                  End If
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
                  '
                  GoTo 899
                End If
              End If
            End If
          End If
'          If List1.ListCount > 0 Then
'            If TIX$ = "OP" Or TIX$ = "FF" Then
'              If SAFETEXT$(!CODICOM_LAR) = NROX$ Then
'                If CVINT%(!FECHEMISI) = FEX Then
'                  SALDO# = SALDO# - IMPO#
'                  IMPO# = IMPO# + XVALO(Mid$(z$, 40, 11)) - XVALO(Mid$(z$, 53, 11))
'                  List1.RemoveItem List1.ListCount - 1
'                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
'                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
'                  Call REPLA(z$, Space$(24), 40, 24)
'                  If IMPO# > 0 Then
'                      Call REPLA(z$, IMPTEX$, 40, 11)
'                    ElseIf IMPO# < 0 Then
'                      Call REPLA(z$, IMPTEX$, 53, 11)
'                  End If
'                  Call REPLA(z$, SALTEX$, 64, 13)
'                  List1.AddItem z$
'                  GoTo 899
'                End If
'              End If
'            End If
'          End If
          ' FIN AGREGADO 27/05/06
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!FECHEMISI)
              NROL& = !NumeComp
              NROX$ = SAFETEXT$(!CODICOM_LAR) ' LO AGREGUE ACA TAMBIEN POR QUE SI NO CUANDO ERA DOLARES SE QUEDABA CON EL ULTIMO QUE LEIA Y NO ACTUALIZABA
              'Call REPLA(z$, FECHATEX$(FEX), 1, 8)
              'Call REPLA(z$, TIPOS$, 11, 11)
              'Call REPLA(z$, NROX$, 22, 18)
              '  TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
              '  While Len(TTX1$) < 23: TTX1$ = "." + TTX1$ + ".": Wend
              'Call REPLA(z$, " " + Left$(TTX1$, 23) + " ", 40, 25)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          
          '
890       SALDO# = SALDO# - IMPO#
          If AjustaCentavos% = 0 Then
            SALDO# = SALDO# - ImporteAjustaCentavos#
            AjustaCentavos% = 1
          End If
          FEX = CVINT%(!FECHEMISI)
          NROX$ = DocumentoProveedor$
          '
          'z$ = Space$(96)
          'Call REPLA(z$, FECHATEX$(FEX), 1, 8)
          'Call REPLA(z$, TIPOS$, 11, 11)
          'Call REPLA(z$, NROX$, 22, 18)
          IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
          SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
          If IMPO# > 0 Then
              'Call REPLA(z$, IMPTEX$, 40, 11)
            ElseIf IMPO# < 0 Then
              'Call REPLA(z$, IMPTEX$, 53, 11)
          End If
          'Call REPLA(z$, SALTEX$, 64, 13)
          '
          If MODOMONEI% = 2 Then
              MONEM% = ROUND(!Mone_Emis, 0) Mod 6   ' enjuague gramy
              TICAM = ROUND(!Tipo_Cam, 5)
              If TICAM < 0.05 Then TICAM = 1
              '
              'Call REPLA(z$, SIMBOLPES$(MONEM%), 79, 2)
              'If MONEM% > 1 Then
                 'Call REPLA(z$, FORMATNUM$(TICAM, "F9.4"), 81, 9)
              'End If
              'Call REPLA(z$, FORMATNUM$(MONEM%, "I2"), 103, 2)
              'Call REPLA(z$, COCOLA$, 105, 24)
              '
          End If
          
      'List1.AddItem z$
895       FilaCuentaCorriente& = FilaCuentaCorriente& + 1
          'mfgCuentaCorriente.Rows = mfgCuentaCorriente.Rows + 1
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
          If IMPO# > 0 Then
              mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          ElseIf IMPO# < 0 Then
              mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          End If
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
          If MODOMONEI% = 2 Then
            mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 7) = TRIM$(SIMBOLPES$(MONEM%))
            mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 8) = TRIM$(FORMATNUM$(TICAM, "G9.4"))
          End If
          If MONRESU% > 1 Then
              MONEM% = ROUND(!Mone_Emis, 0)
              If MONEM% <> MONRESU% Then
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = "..........."
                    LeyendaMoneda$ = TRIM$(DEMONECO$(MONEM%))
                    If Len(LeyendaMoneda$) < 11 Then Blancos = Int((10 - Len(LeyendaMoneda$)) / 2)
                    LeyendaMoneda$ = LeyendaMoneda$ & Space(Blancos)
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = LeyendaMoneda$
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = "............"
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 7) = ""
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 8) = ""
              End If
          End If
          '
899      .MoveNext
        Loop
      End If
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    List1.Refresh
    mfgCuentaCorriente.Rows = FilaCuentaCorriente& + 1
    Openforms = DoEvents
    '
    'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    '
    FilaCompDeuda& = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    mfgComposicionDeuda.Rows = 1
    CondicionSql$ = " NuProve = " & NC% & " AND (abs(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005)"
    CantidadRegistros& = CantRegistros("SACREPEN", CondicionSql$)
    mfgComposicionDeuda.Rows = CantidadRegistros& + 1
    
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    If TCUE% <> 0 Then GoTo 999
    '
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM SACREPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CondicionSql$
    SQLX$ = SQLX$ + " ORDER BY FeVencim ASC;"
26  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             GoTo 993
          End If
          TIX$ = UCase$(Left$(SAFETEXT$(!CODICOM_LAR), 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "RC" Then TIPOS$ = "Recibo"
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = (-1) * XVALO(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
          SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = SAFETEXT$(!CODICOM_LAR)
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
          '
          SALTEXDOL$ = ""
          COCOLAR$ = !CODICOM_LAR
          '
          SALDOLA# = 0 ' CAMBIO 13/07/15
          'If MODOMONEI% = 2 THEN 'ANTES
          If MODOMONEI% = 2 And XVALO(!Mone_Emis) = 2 Then ' CAMBIO 13/07/15
             SALDOLA# = ROUND(XVALO(!SaldAcre_DOLA) - XVALO(!SaldDebe_DOLA))
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          
20        Z$ = Space$(90)
          'Call REPLA(z$, TIPOS$, 1, 11)
          'Call REPLA(z$, NROX$, 12, 18)
          'Call REPLA(z$, FECHATEX$(FEMI), 32, 8)
          
          'Call REPLA(z$, IMPTEX$, 40, 12)
          'Call REPLA(z$, FECHATEX$(FEVE), 56, 8)
          'Call REPLA(z$, SALTEX$, 65, 12)
          'Call REPLA(z$, SALTEXDOL$, 77, 13)
          SUDEUDOL# = SUDEUDOL# + SALDOLA#
          '
'          List2.AddItem z$
          '
          FilaCompDeuda& = FilaCompDeuda& + 1
          'mfgComposicionDeuda.Rows = mfgComposicionDeuda.Rows + 1
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 1) = TIPOS$
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 2) = TRIM$(NROX$)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 3) = FECHATEX$(FEMI)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 5) = FECHATEX$(FEVE)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
          If MODOMONE$ = "MULTIMON" Then mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 7) = FORMATNUM$(XVALO(SALTEXDOL$), "G14.2")
          '
993     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
'    If MONRESU% <= 1 Then
    If MODOMONEI% <= 1 Then
        If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
           If Command4.Enabled = True Then
             If CONTROLGRABA% = 0 Then
               ATENTI% = 1
              '
              ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
              ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR PAGAR APARECIA
              ' DESACTUALIZADO.
              '
              'Call RAPLIPRO(NC%)
              'GoTo 100
    '           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar y Reconstruir\Continuar sin Reaplicar"
    '           OPX1% = COMALTER%(TTXX$)
    '           If OPX1% = 2 Then 'CONTINUAR Y RECONSTRUIR
    '               Call RAPLIPRO(NC%)
    '               GoTo 100
    '           ElseIf OPX1% = 3 Then ' CONTINUAR SIN REAPLICAR
    '               GoTo 995
    '           Else
    '               ATENTI% = 0 '  CANCELA
    '               Call CIERRARCH("*.*")
    '               Call FINAL("")
    '           End If
             End If
           End If
        End If
    End If
    '
'    If MONRESU% <= 1 Then
'       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
'         If CONTROLGRABA% = 0 Then
'           ATENTI% = 1
'           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'           If COMALTER%(TTXX$) = 2 Then
'               Call RAPLIPRO(NC%)
'               GoTo 100
'             Else
'               ATENTI% = 0
'               Call CIERRARCH("*.*")
'               Call FINAL("")
'           End If
'         End If
'       End If
'    End If
    '

995 List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label5.Caption = TRIM$(FORMATNUM$(SUMADEU#, "G14.2"))
    Label5.Refresh
    Label6.Caption = TRIM$(FORMATNUM$(SUMADEU#, "G14.2"))
    Label6.Refresh
    Label13.Caption = TRIM$(FORMATNUM$(SUDEUDOL#, "G14.2"))
    Label13.Refresh
    
    mfgCuentaCorriente.Rows = FilaCuentaCorriente& + 1
    POSICIONAR_ULTIMAFILA mfgCuentaCorriente
    mfgComposicionDeuda.Rows = FilaCompDeuda& + 1
    POSICIONAR_ULTIMAFILA mfgComposicionDeuda
   
    Call POSICIONAR_ULTIMAFILA(mfgCuentaCorriente)
    Call POSICIONAR_ULTIMAFILA(mfgComposicionDeuda)
    'Call REMPLAGRID(Me, GRID1, Picture2, Label4, List2)
    '
999 Screen.MousePointer = 1
    ENPROCESO% = 0
    Text1.Enabled = True
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    '
End Sub
'
Sub CARESCUEPRO_ahp()
    '
    'If Picture1.Visible = False Then Exit Sub
    If mfgCuentaCorrienteAHP.Visible = False Then Exit Sub
    Static ENPROCESO%
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    Text1.Enabled = False
5   If ENPROCESO% > 0 Then Exit Sub
    ENPROCESO% = 1
    '
10  Screen.MousePointer = 11
    NC% = XVALO(Text1)
    NCX7% = XVALO(Text1)
    NICONFLIM% = NIVCONFI%
    Command4.Enabled = True
    '
    If NC% > 0 Then
        NC1% = (-1) * NC%
        Label26 = CODICLI$(NC1%)
        Text2 = rasocli$(NC1%)
        Text3 = DOMICLI$(NC1%)
        Text4 = CPOSCLI$(NC1%)
        Text5 = LOCACLI$(NC1%)
        Text7 = TEL1CLI$(NC1%)
        Text10 = TRIM$(FORMATNUM$(SALDPRO#(NC1%), "G14.2"))
        TCUE% = TICUEPRO%(NC1%)
        DoEvents
        Refresh
    End If
    '
100 List3.Clear
    List3.Refresh
    List2.Clear
    List2.Refresh
    SALDO# = 0: SARRAS# = 0
    FilaCuentaCorriente& = 0
    MaximoDeFilasGrid& = 10000
    mfgCuentaCorrienteAHP.Rows = 1: mfgCuentaCorrienteAHP.Rows = MaximoDeFilasGrid&
    '
    RFF$ = RECFILT$("CUECOPRO", 1, MKI$(NC%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOPRO", RELE&)
      TIPOS$ = Mid$(XX$, 5, 16)
      IMPO# = CVD(Mid$(XX$, 25, 8))
      SALDO# = SALDO# - IMPO#
      FEX = CVI(Mid$(XX$, 3, 2))
      NROL& = CVS(Mid$(XX$, 5, 4))
      NROX$ = Mid$(XX$, 33, 24)
      '
      Z$ = Space$(96)
      Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
      Call REPLA(Z$, TIPOS$, 11, 11)
      Call REPLA(Z$, NROX$, 22, 18)
      IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
      SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
      If IMPO# > 0 Then
          Call REPLA(Z$, IMPTEX$, 40, 11)
        ElseIf IMPO# < 0 Then
          Call REPLA(Z$, IMPTEX$, 53, 11)
      End If
      Call REPLA(Z$, SALTEX$, 64, 13)
      
      FilaCuentaCorriente& = FilaCuentaCorriente& + 1
      mfgCuentaCorrienteAHP.Rows = mfgCuentaCorrienteAHP.Rows + 1
      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
      If IMPO# > 0 Then
          mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
      ElseIf IMPO# < 0 Then
          mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
      End If
      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
      
      'List3.AddItem z$
    Next KU&
    DoEvents
    '
    Call APERBASDAT("CABAN")
    'jandy sql
    SALDOAHPGEN# = 0: ULTIMORECORRIDO# = 0: ACUMULAHABERAHP# = 0: TOTGASFLETE1AHP# = 0: ACUMULA_DIFER# = 0: ACUMULA_DEBE# = 0: ACUMULA_COMPYGASTO# = 0
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    'SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, VAL(MID(CodiCom_Lar,11)) ASC;"
    SQLX$ = SQLX$ + " ORDER BY FECHCONTA, FEHOREAL ASC"
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    If NCX7% <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
'    If InStr(1, EMPREAC$, "AHP") > 0 Then
'        DEB# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "NUPROVE=" & Text1))
'        HAB# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", "NUPROVE=" & Text1))
'        SALDO1# = Abs(DEB# - HAB#)
'
'        DESACRE# = XVALO(VALOBADA("SACREPEN", "SaldDebe", "NUPROVE=" & Text1))
'        HASACRE# = XVALO(VALOBADA("SACREPEN", "SaldAcre", "NUPROVE=" & Text1))
'
'        SALDO2# = Abs(DESACRE# - HASACRE#)
'
'        A = SALDO1# - SALDO2#
'
'    End If
    
    SALDOAHPGEN# = 0
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = Left$(SAFETEXT$(!CODICOM_LAR), 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "RC" Then TIPOS$ = "Recibo"
                              If TIX$ = "NV" Then TIPOS$ = "Nota de Venta"
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             Command4.Enabled = False            ' NO PERMITE REAPLICAR
             GoTo 899
          End If
          '
          ' AGREGADO PARA AGRUPAR O/P'S MULTIPLES A EGRESOS Y GASTOS - 27/05/06
          If mfgCuentaCorrienteAHP.Rows > 0 Then
            If TIX$ = "OP" Or TIX$ = "FF" Then
              If SAFETEXT$(!CODICOM_LAR) = NROX$ Then
                If CVINT%(!FECHEMISI) = FEX Then
                  SALDO# = SALDO# - IMPO#
                  IMPO# = IMPO# + XVALO(Mid$(Z$, 40, 11)) - XVALO(Mid$(Z$, 53, 11))
                  Call cmdEliminaItem(mfgCuentaCorrienteAHP, mfgCuentaCorrienteAHP.Rows - 1)
                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
                  Call REPLA(Z$, Space$(24), 40, 24)
                  If IMPO# > 0 Then
                      Call REPLA(Z$, IMPTEX$, 40, 11)
                    ElseIf IMPO# < 0 Then
                      Call REPLA(Z$, IMPTEX$, 53, 11)
                  End If
                  Call REPLA(Z$, SALTEX$, 64, 13)
                  
                  FilaCuentaCorriente& = FilaCuentaCorriente& + 1
                  mfgCuentaCorrienteAHP.Rows = mfgCuentaCorrienteAHP.Rows + 1
                  mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
                  mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
                  mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
                  If IMPO# > 0 Then
                      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
                  ElseIf IMPO# < 0 Then
                      mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
                  End If
                  mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
                  
                  GoTo 899
                End If
              End If
            End If
          End If
          
'          If List3.ListCount > 0 Then
'            If TIX$ = "OP" Or TIX$ = "FF" Then
'              If SAFETEXT$(!CODICOM_LAR) = NROX$ Then
'                If CVINT%(!FECHEMISI) = FEX Then
'                  SALDO# = SALDO# - IMPO#
'                  IMPO# = IMPO# + XVALO(Mid$(z$, 40, 11)) - XVALO(Mid$(z$, 53, 11))
'                  List3.RemoveItem List3.ListCount - 1
'                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
'                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
'                  Call REPLA(z$, Space$(24), 40, 24)
'                  If IMPO# > 0 Then
'                      Call REPLA(z$, IMPTEX$, 40, 11)
'                    ElseIf IMPO# < 0 Then
'                      Call REPLA(z$, IMPTEX$, 53, 11)
'                  End If
'                  Call REPLA(z$, SALTEX$, 64, 13)
'                  List3.AddItem z$
'                  GoTo 899
'                End If
'              End If
'            End If
'          End If
          
          ' FIN AGREGADO 27/05/06
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!FECHEMISI)
              NROL& = XVALO(!NumeComp)
              NROX$ = SAFETEXT$(!CODICOM_LAR)
              Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
              Call REPLA(Z$, TIPOS$, 11, 11)
              Call REPLA(Z$, NROX$, 22, 18)
                TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
                While Len(TTX1$) < 23: TTX1$ = "." + TTX1$ + ".": Wend
              Call REPLA(Z$, " " + Left$(TTX1$, 23) + " ", 40, 25)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          
          '
890       SALDO# = SALDO# - IMPO#
          FEX = CVINT%(!FECHEMISI)
          NROX$ = SAFETEXT$(!CODICOM_LAR)
          '
          Z$ = Space$(148)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(Z$, TIPOS$, 11, 11)
          Call REPLA(Z$, NROX$, 22, 18)
          IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
          SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
          If IMPO# > 0 Then
              Call REPLA(Z$, IMPTEX$, 40, 11)
              ACUMULA_DEBE# = ACUMULA_DEBE# + IMPO# ' ACUMULO LA COLUMNA DEBE
            ElseIf IMPO# < 0 Then
              Call REPLA(Z$, IMPTEX$, 53, 11)
          End If
          Call REPLA(Z$, SALTEX$, 64, 13)
          '
          If MODOMONEI% = 2 Then
              MONEM% = ROUND(!Mone_Emis, 0) Mod 6   ' enjuague gramy
              TICAM = ROUND(!Tipo_Cam, 5)
              If TICAM < 0.05 Then TICAM = 1
              '
              Call REPLA(Z$, SIMBOLPES$(MONEM%), 79, 2)
              If MONEM% > 1 Then
                 Call REPLA(Z$, FORMATNUM$(TICAM, "F9.4"), 81, 9)
              End If
              
              Call REPLA(Z$, FORMATNUM$(MONEM%, "I2"), 140, 2)
              'Call REPLA(z$, COCOLA$, 105, 24)
              '
          End If
           VALCOMPRA# = 0: GASFLETE1# = 0
           VALCOMPRA# = ValorCompraFacturaExpo#(NC%, NROX$, MONRESU%, 1)
           If VALCOMPRA# <> 0 Then
              GASFLETE1# = GASFLETE#(NC%, NROX$, MONRESU%)
              IMPO11# = (VALCOMPRA# + GASFLETE1#)
              
           Else
              VALCOMPRA# = 0
           End If
             IMPO11# = (-1 * IMPO11#) 'SI NO ES FACTURA AGREGA EL MISMO IMPORTE NEGATIVO SI ES UN PAGO O NC.
            Call REPLA(Z$, FORMATNUM$(VALCOMPRA#, "F11.2"), 90, 11)
            Call REPLA(Z$, FORMATNUM$(GASFLETE1#, "F11.2"), 101, 11)
            Call REPLA(Z$, FORMATNUM$(VALCOMPRA# + GASFLETE1#, "F13.2"), 112, 13)
            ACUMULA_COMPYGASTO# = ACUMULA_COMPYGASTO# + VALCOMPRA# + GASFLETE1#
            'ESTO ES P Q SE VA A CALCULAR A PARTIR DEL 01/07/15 PERO DEBERA SUMAR EL ULTIMO REGISTRO ANTERIOR A ESTA FECHA QUE SE GUARDA EN ULTIMORECORRIDO#

             If FEX >= FECHANUM("01/07/15") And InStr(EPAC$, "AHP") > 0 Then
              SALDOAHPGEN# = SALDOAHPGEN# + VALCOMPRA#
              TOTGASFLETE1AHP# = TOTGASFLETE1AHP# + GASFLETE1#
              If IMPO# < 0 Then
                ACUMULAHABERAHP# = ACUMULAHABERAHP# + Abs(IMPO#)
              End If
              'COLUMNA AGREGADA DE DIFERENCIA COMPRA - HABER
              DIFER_COMPRA_HABER# = 0
              If VALCOMPRA# > 0 Then
                 DIFER_COMPRA_HABER# = (VALCOMPRA# + GASFLETE1#) - Abs(IMPO#)
              End If
              Call REPLA(Z$, FORMATNUM$(DIFER_COMPRA_HABER#, "F13.2"), 125, 13)

            Else
              ULTIMORECORRIDO# = VALCOMPRA# + GASFLETE1# 'SIEMPRE VA A QUEDAR EL ULTIMO ANTES DE LA FECHA 01/07/15
            End If
            ACUMULA_DIFER# = ACUMULA_DIFER# + DIFER_COMPRA_HABER#

895         FilaCuentaCorriente& = FilaCuentaCorriente& + 1
            'mfgCuentaCorrienteAHP.Rows = mfgCuentaCorrienteAHP.Rows + 1
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 3) = TRIM$(NROX$)
            If IMPO# > 0 Then
                mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
            ElseIf IMPO# < 0 Then
                mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
            End If
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
            
            If MODOMONEI% = 2 Then
                mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 7) = MONEM%
                If MONEM% > 1 Then mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 8) = TRIM$(FORMATNUM$(TICAM, "G9.4"))
            End If
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 9) = TRIM$(FORMATNUM$(VALCOMPRA#, "G12.2"))
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 10) = TRIM$(FORMATNUM$(GASFLETE1#, "G12.2"))
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 11) = TRIM$(FORMATNUM$(VALCOMPRA# + GASFLETE1#, "G13.2"))
            mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 12) = TRIM$(FORMATNUM$(DIFER_COMPRA_HABER#, "G13.2"))
            
            If MONRESU% > 1 Then
                MONEM% = ROUND(!Mone_Emis, 0)
                If MONEM% <> MONRESU% Then
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 4) = "..........."
                        LeyendaMoneda$ = TRIM$(DEMONECO$(MONEM%))
                        If Len(LeyendaMoneda$) < 10 Then Blancos = Int((9 - Len(LeyendaMoneda$)) / 2)
                        LeyendaMoneda$ = LeyendaMoneda$ & Space(Blancos)
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 5) = LeyendaMoneda$
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 6) = "............"
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 7) = ""
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 8) = ""
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 9) = ""
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 10) = ""
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 11) = ""
                    mfgCuentaCorrienteAHP.TextMatrix(FilaCuentaCorriente&, 12) = ""
                End If
            End If
'       List3.AddItem z$

'             Debug.Print z$
          '
899      .MoveNext
        Loop
      End If
990 End With

    'TOTALIZADORES
'    Me.LBLTOTAHP.Caption = FORMATNUM$(SALDOAHPGEN# + ULTIMORECORRIDO#, "F15.2")
'    Me.LBLTOTHABER = FORMATNUM$(ACUMULAHABERAHP#, "F15.2")
'    Me.LBLTOTMISELANIAS = FORMATNUM$(TOTGASFLETE1AHP#, "F15.2")
'    Me.LBLRESULTADO = FORMATNUM$(ACUMULAHABERAHP# - (XVALO(Me.LBLTOTAHP.Caption) + TOTGASFLETE1AHP#), "F15.2")

    Me.LBLTOTDIFERENCIA.Caption = TRIM$(FORMATNUM$(ACUMULA_DIFER#, "G15.2"))
    Me.LBLTOTREAL_DEBE = TRIM$(FORMATNUM$(ACUMULA_COMPYGASTO# - ACUMULA_DEBE#, "G15.2"))
    
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    List3.Refresh
    Openforms = DoEvents
    '
    'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List3)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    FilaCompDeuda& = 0
    mfgComposicionDeuda.Rows = 1: mfgComposicionDeuda.Rows = MaximoDeFilasGrid&
    '
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    '
    If TCUE% <> 0 Then GoTo 999
    '
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM SACREPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (abs(SaldAcre) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SaldDebe) >= 0.005) "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
26  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NICON% = XVALO(!Niv_Clasi)
          If NICON% > NICONFLIM% Then
             GoTo 993
          End If
          TIX$ = UCase$(Left$(SAFETEXT$(!CODICOM_LAR), 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
                              If TIX$ = "RC" Then TIPOS$ = "Recibo"
                              If TIX$ = "ND" Then TIPOS$ = "N.Debito"
                              If TIX$ = "CP" Then TIPOS$ = "Créd.Prop."
                              If TIX$ = "CO" Then TIPOS$ = "Comprob."
                              If TIX$ = "OP" Then TIPOS$ = "O.Pago": SG% = 1
                              If TIX$ = "FF" Then TIPOS$ = "O/P F.Fijo": SG% = 1
                              If TIX$ = "NC" Then TIPOS$ = "N.Crédito": SG% = 1
                              If TIX$ = "DP" Then TIPOS$ = "Deb.Propio": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                              '
          IMPO# = (-1) * XVALO(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = XVALO(!SaldAcre) - XVALO(!SaldDebe)
          SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = SAFETEXT$(!CODICOM_LAR)
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
          '
          SALTEXDOL$ = ""
          COCOLAR$ = !CODICOM_LAR
          '
          SALDOLA# = 0 ' CAMBIO 13/07/15
          'If MODOMONEI% = 2 THEN 'ANTES
          If MODOMONEI% = 2 And XVALO(!Mone_Emis) = 2 Then ' CAMBIO 13/07/15
             SALDOLA# = ROUND(XVALO(!SaldAcre_DOLA) - XVALO(!SaldDebe_DOLA))
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          
20        Z$ = Space$(90)
          Call REPLA(Z$, TIPOS$, 1, 11)
          Call REPLA(Z$, NROX$, 12, 18)
          Call REPLA(Z$, FECHATEX$(FEMI), 32, 8)
          
          Call REPLA(Z$, IMPTEX$, 40, 12)
          Call REPLA(Z$, FECHATEX$(FEVE), 56, 8)
          Call REPLA(Z$, SALTEX$, 65, 12)
          Call REPLA(Z$, SALTEXDOL$, 77, 13)
          SUDEUDOL# = SUDEUDOL# + SALDOLA#
          '
'          List2.AddItem z$
          '
          FilaCompDeuda& = FilaCompDeuda& + 1
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 1) = TIPOS$
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 2) = TRIM$(NROX$)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 3) = FECHATEX$(FEMI)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 5) = FECHATEX$(FEVE)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 7) = FORMATNUM$(XVALO(SALTEXDOL$), "G14.2")
          '
993     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
'    If MONRESU% <= 1 Then
    If MODOMONEI% <= 1 Then
        If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
           If Command4.Enabled = True Then
             If CONTROLGRABA% = 0 Then
               ATENTI% = 1
              '
              ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
              ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR PAGAR APARECIA
              ' DESACTUALIZADO.
              '
              Call RAPLIPRO(NC%)
              GoTo 100
    '           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar y Reconstruir\Continuar sin Reaplicar"
    '           OPX1% = COMALTER%(TTXX$)
    '           If OPX1% = 2 Then 'CONTINUAR Y RECONSTRUIR
    '               Call RAPLIPRO(NC%)
    '               GoTo 100
    '           ElseIf OPX1% = 3 Then ' CONTINUAR SIN REAPLICAR
    '               GoTo 995
    '           Else
    '               ATENTI% = 0 '  CANCELA
    '               Call CIERRARCH("*.*")
    '               Call FINAL("")
    '           End If
             End If
           End If
        End If
    End If
    '
'    If MONRESU% <= 1 Then
'       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
'         If CONTROLGRABA% = 0 Then
'           ATENTI% = 1
'           TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'           If COMALTER%(TTXX$) = 2 Then
'               Call RAPLIPRO(NC%)
'               GoTo 100
'             Else
'               ATENTI% = 0
'               Call CIERRARCH("*.*")
'               Call FINAL("")
'           End If
'         End If
'       End If
'    End If
    '

995 List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    mfgComposicionDeuda.Refresh
    Label5.Caption = TRIM$(FORMATNUM$(SUMADEU#, "G14.2"))
    Label5.Refresh
    Label6.Caption = TRIM$(FORMATNUM$(SUMADEU#, "G14.2"))
    Label6.Refresh
    Label13.Caption = TRIM$(FORMATNUM$(SUDEUDOL#, "G14.2"))
    Label13.Refresh
    'Call REMPLAGRID(Me, GRID1, Picture2, Label4, List2)
    
    mfgCuentaCorrienteAHP.Rows = FilaCuentaCorriente + 1
    mfgComposicionDeuda.Rows = FilaCompDeuda& + 1
    Call POSICIONAR_ULTIMAFILA(mfgCuentaCorriente)
    Call POSICIONAR_ULTIMAFILA(mfgComposicionDeuda)
    '
999 Screen.MousePointer = 1
    ENPROCESO% = 0
    Text1.Enabled = True
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    '
End Sub
Private Sub Time2_Timer()

End Sub


