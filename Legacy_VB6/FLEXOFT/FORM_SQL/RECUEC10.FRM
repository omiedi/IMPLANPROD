VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RECUEC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Deudores"
   ClientHeight    =   8175
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid mfgCuentaCorriente 
      Height          =   2655
      Left            =   120
      TabIndex        =   79
      Top             =   2400
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   6442536
      BackColorBkg    =   14737632
      ScrollBars      =   2
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
      Left            =   120
      TabIndex        =   81
      Top             =   5130
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   6442536
      BackColorBkg    =   14737632
      ScrollBars      =   2
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
   Begin VB.CommandButton Command12 
      Caption         =   "Paso 1"
      Height          =   255
      Left            =   1800
      TabIndex        =   70
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Paso2"
      Height          =   255
      Left            =   2760
      TabIndex        =   69
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2D3C0&
      Caption         =   "Vista Clásica"
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
      Height          =   195
      Left            =   100
      TabIndex        =   66
      Top             =   2120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   10935
      TabIndex        =   58
      Top             =   7660
      Width           =   10960
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
         Height          =   345
         Left            =   9120
         TabIndex        =   62
         Top             =   60
         Width           =   1545
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
         Height          =   345
         Left            =   6090
         TabIndex        =   61
         Top             =   60
         Width           =   1515
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
         Left            =   30
         TabIndex        =   60
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Height          =   345
         Left            =   9240
         TabIndex        =   59
         Top             =   60
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   1080
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GRID1 
      Height          =   765
      Left            =   1560
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   11160
      ScaleHeight     =   8205
      ScaleWidth      =   1785
      TabIndex        =   47
      ToolTipText     =   "Re-Aplicación de Comprobantes"
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   75
         Top             =   7440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   76
            Top             =   0
            Width           =   12000
         End
      End
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
         Left            =   105
         Picture         =   "RECUEC10.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Búsqueda por Texto"
         Top             =   1700
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Acceso A Control de Factura Anticipo"
         Top             =   5150
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   7080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Impresión de Cuenta Corriente"
         Top             =   2380
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":1B1A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Notas de Clientes"
         Top             =   3080
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":1E24
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
         Top             =   4450
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":1F6E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1000
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":2278
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Finaliza la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "D-C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":23C2
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
         Top             =   3770
         Visible         =   0   'False
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "RECUEC10.frx":26CC
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Re-Aplicacion Manual de Créditos"
         Top             =   5840
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   6720
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   12
      Left            =   12000
      TabIndex        =   44
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   10
      Left            =   9500
      TabIndex        =   43
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   8
      Left            =   16000
      TabIndex        =   42
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   7
      Left            =   2310
      TabIndex        =   41
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   9
      Left            =   3240
      TabIndex        =   40
      Top             =   5040
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   11
      Left            =   4305
      TabIndex        =   39
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   13
      Left            =   6300
      TabIndex        =   38
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2750
      Index           =   15
      Left            =   7455
      TabIndex        =   37
      Top             =   5050
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   0
      Left            =   1290
      TabIndex        =   36
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   1
      Left            =   3500
      TabIndex        =   35
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   2
      Left            =   4340
      TabIndex        =   34
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   3
      Left            =   5910
      TabIndex        =   33
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   4
      Left            =   7690
      TabIndex        =   32
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   5
      Left            =   9500
      TabIndex        =   31
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Height          =   2450
      Index           =   6
      Left            =   12000
      TabIndex        =   30
      Top             =   2415
      Width           =   15
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CRÉDITO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8230
      TabIndex        =   19
      Top             =   360
      Width           =   2820
      Begin VB.TextBox Text12 
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
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   1100
         TabIndex        =   26
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1245
         Width           =   1590
      End
      Begin VB.TextBox Text6 
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
         Left            =   1100
         TabIndex        =   24
         Tag             =   "TBOX"
         Text            =   " "
         ToolTipText     =   "Doble Click para Listar Detalle"
         Top             =   915
         Width           =   1590
      End
      Begin VB.TextBox Text11 
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
         Left            =   1100
         TabIndex        =   22
         Tag             =   "TBOX"
         Top             =   255
         Width           =   1590
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
         Left            =   1100
         TabIndex        =   20
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
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
         Left            =   75
         TabIndex        =   27
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
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
         Left            =   75
         TabIndex        =   25
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lim. Crédito"
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
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Top             =   345
         Width           =   1005
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Cte"
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
         Left            =   75
         TabIndex        =   21
         Top             =   690
         Width           =   885
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   9870
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   5460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2430
      Left            =   120
      ScaleHeight     =   2400
      ScaleWidth      =   9690
      TabIndex        =   14
      Top             =   2400
      Width           =   9725
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Height          =   5580
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   10515
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            IntegralHeight  =   0   'False
            Left            =   -15
            TabIndex        =   16
            Top             =   3220
            Width           =   9720
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            IntegralHeight  =   0   'False
            Left            =   0
            TabIndex        =   67
            Top             =   320
            Width           =   9720
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"RECUEC10.frx":2816
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
            TabIndex        =   71
            Top             =   70
            Width           =   10410
         End
         Begin VB.Line Line7 
            X1              =   1155
            X2              =   1155
            Y1              =   1470
            Y2              =   3255
         End
         Begin VB.Label Label10 
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
            Height          =   340
            Left            =   -10
            TabIndex        =   72
            Top             =   -10
            Width           =   10410
         End
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   6405
         Y1              =   2400
         Y2              =   2400
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4F9FD&
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   9690
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   9725
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFE0&
         BorderStyle     =   0  'None
         Height          =   5630
         Left            =   0
         TabIndex        =   77
         Top             =   320
         Width           =   10515
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            IntegralHeight  =   0   'False
            Left            =   -10
            TabIndex        =   78
            Top             =   0
            Width           =   9725
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   9660
            Y1              =   2100
            Y2              =   2100
         End
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"RECUEC10.frx":28B8
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
         TabIndex        =   73
         Top             =   60
         Width           =   15000
      End
      Begin VB.Label Label11 
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
         Height          =   345
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Visible         =   0   'False
         Width           =   15000
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECUEC10.frx":2957
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECUEC10.frx":2B8B
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
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
      Height          =   1995
      Left            =   105
      TabIndex        =   2
      Top             =   50
      Width           =   7815
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
         Left            =   1365
         TabIndex        =   11
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1560
         Width           =   4290
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
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   800
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
         Left            =   1050
         TabIndex        =   1
         Top             =   800
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
         Left            =   1365
         TabIndex        =   8
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   420
         Width           =   4290
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
         Left            =   1365
         TabIndex        =   7
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   800
         Width           =   4290
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
         Left            =   1365
         TabIndex        =   6
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1180
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
         Left            =   2205
         TabIndex        =   5
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1180
         Width           =   3450
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
         Left            =   5880
         TabIndex        =   4
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1560
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
         Height          =   300
         Left            =   5880
         TabIndex        =   3
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   800
         Width           =   1695
      End
      Begin VB.Label Label26 
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
         Height          =   300
         Left            =   210
         TabIndex        =   28
         Top             =   1155
         Width           =   1020
      End
      Begin VB.Label label2 
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
         Height          =   195
         Index           =   1
         Left            =   5880
         TabIndex        =   13
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label label3 
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1650
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
         Height          =   330
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Refrescar"
         Top             =   550
         Width           =   960
      End
      Begin VB.Label label2 
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
         Height          =   285
         Index           =   0
         Left            =   5880
         TabIndex        =   9
         Top             =   555
         Width           =   960
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfgCuentaCorrienteAnticipos 
      Height          =   2655
      Left            =   120
      TabIndex        =   82
      Top             =   2400
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   14737632
      BackColorFixed  =   14737632
      BackColorSel    =   6442536
      BackColorBkg    =   14737632
      ScrollBars      =   2
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
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ver Facturas de Anticipo Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   4
      Left            =   7540
      TabIndex        =   65
      ToolTipText     =   "Doble Click Para Ver Detalle"
      Top             =   2100
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   8280
      TabIndex        =   29
      ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
End
Attribute VB_Name = "RECUEC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%
Dim MONRESU%
'jandy sql
Dim ctrl As Object
Public ARGU$
Sub ClearControls()
    For Each ctrl In Me.Controls
        If ctrl.Tag = "TBOX" Then ctrl.TEXT = ""
    Next
    Label26 = "": Label6 = "": Label5 = ""
End Sub

Private Sub recuperaapli()
    Call ABRECONEXION
    'NO LO MODIFICO  MAS POR QUE NO SE USA ,
    
    FF% = FECHANUM("01/11/10")
    FECHA$ = FETEXCOR1$(FF%)
    For II% = 1 To 32767
        JJ& = II%
        Call TRACE(24, JJ&, 32767)
        NCLI$ = CStr(II%)
        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NuClien = " & NCLI$ & " AND IHABECOMP >0 AND TIPOMOVIM='FVEN' AND FECHCONTA>'" & FECHA$ & "'"
        Set RST = READSET(SQLX$)
        With RST
            Do While Not .EOF
                DOCRED$ = SAFETEXT$(!CodiCom_Lar)
                DOCREDCOR$ = SAFETEXT$(!CODICOM_COR)
                If CantRegistros("APLICOBRA", "CoDocAp_Lar='" & DOCRED$ & "'") > 0 Then GoTo 20
                '
                IMPO = XVALO(!IHABECOMP)
                'FF% = CVINT(!FECHCONTA)
                CONDI$ = "NUCLIEN=" & NCLI$ & " AND TIPOMOVIM='FVEN' AND OPCIMOVIM='FC' AND ABS((IDEBECOMP-" & TRIM$(FORMATNUM(IMPO, "F12.2")) & "))<0.005"
                DOCDEB$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$))
                DOCDEBCOR$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_COR", CONDI$))
                'Call COMUNI(CONDI$)
                'Call COMUNI(DOCDEB$)
                If TRIM$(DOCDEB$) <> "" Then
                    'OPX% = COMALTER%("Nota Credito: '" & DOCRED$ & "'   Factura: '" & DOCDEB$ & "'\\No, Cancelar\Si, Continuar")
                    'If OPX% = 1 Then GoTo 20
                    '
                    Dim IMPUMANU As ADODB.Recordset
                    Set IMPUMANU = New ADODB.Recordset
25                  On Error Resume Next
                    IMPUMANU.Open FILTSQL$("APLICOBRA"), FLEXCONN, adOpenKeyset, adLockPessimistic
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                      On Error GoTo 0
                      Call CapturaErrorOpen
                      GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CI% = Int(NCLI$)
                    RAZON$ = rasocli(CI%)
                    NUCLI$ = CODICLI$(CI%)
                    '
                    With IMPUMANU
                    .AddNew   ' YA CONVERTIDO
                    On Error Resume Next
                    !CodiEmpr = CodiEmp%
                    !CodiEmpr = 0
                    !CodiSucu = 0
                    !nuclien = CI%
                    !CODICLIE = CODICLI$(CI%)
                    !RaSoClie = Mid(RAZON$, 1, 24)
                    !codocap_lar = DOCRED$
                    !CODOCAP_COR = DOCREDCOR$
                    !IMPOAPLI = IMPO
                    !refeapli = "Reaplicación de Comprobantes"
                    !codocor_lar = DOCDEB$
                    !CODOCOR_COR = DOCDEBCOR$
                    !ITOTORIG = 0
                    !IIVAORIG = 0
                    .Update
                    On Error GoTo 0
                    '
                    End With
                Else
                    Call COMUNI("Fallo Nota Credito: '" & DOCRED$ & "'   Factura: '" & DOCDEB$ & "'")
                End If
20              .MoveNext
            Loop
        End With
    Next II%
End Sub


Private Sub Check1_Click()
   If Check1.Value = 1 Then
      List3.Visible = False
      mfgCuentaCorrienteAnticipos.Visible = False
   Else
      List3.Visible = True
      List3.ZOrder 0
      mfgCuentaCorrienteAnticipos.ZOrder 0
      mfgCuentaCorrienteAnticipos.Visible = True
   End If
End Sub

Private Sub Command10_Click()
'ESTAS DOS INSTRUCCIONES ESTAN PROBADAS Y OK, SOLUCIONARON EL PROBLEMA
If CLACONTRA% > 0 Then

'    SQLX$ = "update movisto set descri_lar= 'Ant.FC-0001-00005234-A', VALTOTITIVA  = -59755.9"
'    SQLX$ = SQLX$ + " WHERE DOSECLAR = 'FC-0001-00005292-A'"
'    SQLX$ = SQLX$ + " AND NUME_CLIE = 722   AND COD_INTE = 0 and valtotit = -59755.9"
'    FLEXCONN.Execute (SQLX$)
'    '
'
'    'ACA BORRA UNA APLICACION DE UNA FACTURA Y UN AJUSTE POR QUE LA FACTURA APLICADA ORIGINAL SE HABIA ANULADO LUEGO DE HABERSE APLICADO
'    CONDI$ = "DOCU_APLI = 'AJ-7'  AND NUCLIEN = 722"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    CONDI$ = "DOCU_APLI = 'FC-0001-00005273-A'  AND NUCLIEN = 722"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    '
'    'FIN CLIENTE 722
'    '-70313
'     CONDI$ = " DOSECLAR= 'FC-0001-00005243-A'  AND NUME_CLIE= 19 AND COD_INTE = 0 and valtotit = -70313"
'     Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-70313), REGAF&, "NOMESS")
'     Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Ant.FC-0001-00005239-A", REGAF&, "NOMESS")
'
'     'ACA BORRA UNA APLICACION DE UNA FACTURA Y UN AJUSTE POR QUE LA FACTURA APLICADA ORIGINAL SE HABIA ANULADO LUEGO DE HABERSE APLICADO
'    CONDI$ = "DOCU_APLI = 'AJ-3'  AND NUCLIEN = 276"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    CONDI$ = "DOCU_ORIG = 'FC-0001-00005205-A'  AND NUCLIEN = 276"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    '
'
'
'    CONDI$ = " DOSECLAR ='FC-0001-00005213-A' And NUME_CLIE = 954 AND NUORITEM = 3 "
'    Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-641.13), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "DESCUENTO: 30.326 %", REGAF&, "NOMESS")
'
'    '
'    CONDI$ = " DOSECLAR ='FC-0001-00005213-A' And NUME_CLIE = 954 AND NUORITEM = 2 "
'    Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-1472.98), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Ant.FC-0001-00005184-A", REGAF&, "NOMESS")
'
'    CONDI$ = " DOSECLAR ='FC-0003-00000070-A' And NUME_CLIE = 973 AND NUORITEM = 39 "
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Ant.FC-0003-00000004-A (Iva 21%)", REGAF&, "NOMESS")
'
'    CONDI$ = " DOSECLAR ='FC-0003-00000076-A' And NUME_CLIE = 973 AND NUORITEM = 7 "
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Ant.FC-0003-00000004-A (Iva 21%)", REGAF&, "NOMESS")
'    FE% = FECHANUM("06/08/14")
'    '
'    CONDI$ = "DOSECLAR ='FC-0003-00000182-A' AND NUORITEM = 17 and NUME_CLIE = 63"
'    Call ACTUREGISTRO("MOVISTO", "valtotit", CONDI$, (-7601.25), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-7601.25), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALOUNIT", CONDI$, (-7601.25), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALOUNITIVA", CONDI$, (-7601.25), REGAF&, "NOMESS")
'
'    'CONDI$ = "DOSECLAR ='FC-0003-00000182-A' AND NUORITEM = 17"
'    'Call ACTUREGISTRO("MOVISTO", "NUME_CLIE", CONDI$, 63, REGAF&, "NOMESS")
'
'    CONDI$ = "DOSECLAR ='FC-0003-00000179-A' And NUME_CLIE = 441 AND NUORITEM = 2"
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Ant.FC-0003-00000081-A (Iva 21%)", REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, 1, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-12032.13), REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALOUNITIVA", CONDI$, (-12032.13), REGAF&, "NOMESS")
'
'    CONDI$ = "DOSECLAR ='FC-0003-00000179-A' And NUME_CLIE = 441 AND descri_lar like 'Desc.: 10 %'"
'    Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, 1, REGAF&, "NOMESS")
'    '
'
'    'FE% = FECHANUM("13/03/14")
'    '
'    'Call MOVISTOK(FE%, 0, "", "RT", "FC-0003-00000179-A", "FC-000179-A", "FC-0003-00000179-A", "FC-000179-A", "FORTES CRISTIAN MAXIMILIANO", 2, (-12032.13), 1, 441, 1, -1, 0, 0, 0, 0, , (-12032.13), 0, "")
'    CONDI$ = " DOSECLAR ='FC-0001-00005292-A' And NUME_CLIE = 722  and nuoritem = 10"
'    Call ACTUREGISTRO("MOVISTO", "descri_lar", CONDI$, "Desc.: 10 %", REGAF&, "NOMESS")
'    Call ACTUREGISTRO("MOVISTO", "VALTOTITIVA", CONDI$, (-6639.55), REGAF&, "NOMESS")
'    '
'    'DELETE FROM FACTUANTI WHERE DOCU_ORIG = 'FC-0001-00005088-A' 'LA FACTURA ORIGINAL DONDE SE APLICARON LOS ANTICIPOS SE ANULO
'    CONDI$ = "DOCU_ORIG = 'FC-0001-00005086-A' AND NUCLIEN = 946"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    CONDI$ = "DOCU_ORIG = 'FC-0001-00005087-A' AND NUCLIEN = 946"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
'    CONDI$ = "DOCU_ORIG = 'FC-0001-00005088-A' AND NUCLIEN = 946"
'    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
    '
    'ACA BORRA UNA APLICACION DE UNA FACTURA Y UN AJUSTE POR QUE LA FACTURA APLICADA ORIGINAL SE HABIA ANULADO LUEGO DE HABERSE APLICADO EN LA CUENTA 12
    CONDI$ = "DOCU_APLI = 'AJ-1'  AND NUCLIEN = 12"
    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")
    CONDI$ = "DOCU_APLI = 'AJ-2'  AND NUCLIEN = 12"
    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")

    CONDI$ = "DOCU_APLI = 'FC-0001-00005132-A'  AND NUCLIEN = 12"
    Call BORRAREGISTROS("FACTUANTI", CONDI$, REGAF&, "NOMESS")

End If
End Sub

Private Sub Command11_Click()

   Command11.Enabled = False
   If DERACCE%("Ver Aplicaciones de Facturas Anticipos", "VERAPLFA") < 2 Then GoTo 99
   
   CONTROLFANT.Show 1
   
99 Command11.Enabled = True

End Sub

Private Sub Command12_Click()
'todo esto es para corregir el cliente 441
    If CLACONTRA% > 0 Then
'        Call MOVISTOK(FE%, 0, "", "RT", "FC-0003-00000179-A", "FC-000179-A", "FC-0003-00000179-A", "FC-000179-A", "FORTES CRISTIAN MAXIMILIANO", 2, (-12032.13), 1, 441, 1, -1, 0, 0, 0, 0, , (-12032.13), 0, "")
'        Call MOVISTOK(FE%, 0, "", "RT", "FC-0003-00000179-A", "FC-000179-A", "FC-0003-00000179-A", "FC-000179-A", "FORTES CRISTIAN MAXIMILIANO", 3, (-1315.94), 1, 441, 1, -1, 0, 0, "Desc.: 10 %", 0, , (-1315.94), 0, "")
'
'        'todo esto es para corregir el cliente 63
'        Call MOVISTOK(FE%, 0, "", "RT", "FC-0003-00000182-A", "FC-000182-A", "FC-0003-00000182-A", "FC-000182-A", "SEMESURT S.A.", 16, (-68415), 1, 63, 1, -1, 0, 0, "Ant.FC-0003-00000157-A", 0, , (-68415), 0, "")
'        Call MOVISTOK(FE%, 0, "", "RT", "FC-0003-00000182-A", "FC-000182-A", "FC-0003-00000182-A", "FC-000182-A", "SEMESURT S.A.", 17, (-7602), 1, 63, 1, -1, 0, 0, "Desc.: 10 %", 0, , (-7602), 0, "")
    
        'todo esto es para corregir el cliente 12
        FE% = HOY(HO$)
        Call MOVISTOK(FE%, 0, "", "RT", "FC-0001-00005134-A", "FC-005134-A", "FC-0001-00005134-A", "FC-005134-A", "RES-SUR", 23, (-6001.28), 1, 12, 1, -1, 0, 0, "Desc.: 10 %", 0, , (-6001.28), 0, "")
        SQLX$ = "update movisto set descri_lar = 'Ant.FC-0001-00005085-A' where doseclar = 'Fc-0001-00005134-a' and nuoritem = 22"
        FLEXCONN.Execute (SQLX$)
    End If
End Sub



Sub RECADEUDA(NCLI, SALIN#, FECORTE$)

End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command20_Click()
   Dim List As ListBox
   Set List = Me.List1
   If List3.Visible = True Then Set List = Me.List3
   
   Command20.Enabled = False ' BUSQUEDA POR NUMERO DE SERIE
   BAND% = 0 ' BANDERA PARA VERIFICAR SI ENCONTRO EL ARGUMENTO
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda por Número de Serie", ARGU$)))
   If ARGU$ = "" Then Exit Sub
   
   Screen.MousePointer = 11
   If List.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For i& = 1 To List.ListCount
'        If TRIM$(ARGU$) = TRIM$(List1.List(I& - 1)) Then
        If InStr((TRIM$(UCase$(Left$(List.List(i& - 1), 77)))), ARGU$) Then
          List.ListIndex = i& - 1
          List.SetFocus
          BAND% = 1
          Exit For
        Else
        End If
     Next i&
   End If
   If BAND% < 1 Then Call MENSERR(24, "Argumento No Encontrado")
   Screen.MousePointer = 1
   Command20.Enabled = True

End Sub

'Sub MUESTRASTSCREEN(NC%, NUDOCUM$, FE%, OK%)
'    '
'    ' OJO  SE GENERA ESTA RUTINA PARA MOSTRAR LAS FACTURAS CON EL STDSCREEN
'    ' SE LLAMAN A LOS SQR COMO ESTAN EN MEDITEA SE VA A TENER QUE VER SI
'    'SE VAN A CONFIGURAR PARA QUE NO QUEDE PERSONALIZADO
'
'    'VALIDAR SI ES EXISTE
''    If EXISTE%("VENGES09.EXE") < 1 Then
''       Exit Sub
''    End If
'    '
'    FEX = FETEXCOR1$(FE%)
'    CONDI$ = "TIPOMOVIM = 'FVEN' AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NC% & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
'    FACTURAX$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
'    '
'    If TRIM$(FACTURAX$) <> "" Then
'        TIPOCOM$ = Mid$(FACTURAX$, 1, 2)
'        LETR$ = Right(FACTURAX$, 1)
'        '
'        CONDI$ = "DOSECLAR='" & FACTURAX$ & "' AND NUME_CLIE= " & NC% & " AND COD_INTE > 0"
'        CANTREG& = XVALO(CantRegistros("MOVISTO", CONDI$, "NOMESS"))
'        EMP$ = "1-MDT": If CANTREG& = 0 Then EMP$ = "T-MDT"
'        ARCSQR$ = TIPOCOM$ & LETR$ & EMP$
'         '
'        If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
'            Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
'            Exit Sub
'        End If
'       '
'       OK% = 1
'       FILTX$ = TRIM$(FACTURAX$) & ";" & TRIM$(Str$(NC%))
'       Call STDFORM(ARCSQR$, FILTX$, MODO$)
'    End If
'
'End Sub
'
Private Sub Command6_Click()
   Command6.Enabled = False
   Call OPVARFA07.Command2_Click
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   NC% = XVALO(Text1)
   If NC% < 1 Then Exit Sub
   If REGICLI&(NC%) < 1 Then Exit Sub
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     'If SALDOL#(KKI%) <= 0.005 Then
     '  If SALPES#(KKI%) > 0 Then
         XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
         FEMI = CVI(Mid$(XXC$, 1, 2))
         Z$ = Space$(60) + MKS$(0)
         Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
         Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
         Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
         Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
         Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
         Call REGAPP("SELDOCAP", Z$)
     '  End If
     'End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Call COMUNI("No Aparecen Créditos Pendientes\para la Cuenta Elegida.")
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   '
End Sub

Private Sub Command8_Click()

    Command8.Enabled = False
    CLIPROACTIVO = XVALO(Text1)
    Call RECICO07.IMPUCREVEN     ' Imputación de Créditos
    Text1 = ""
    If CLIPROACTIVO > 0 Then
        Text1 = TRIM$(Str$(Abs(CLIPROACTIVO)))
        CLIPROACTIVO = 0
    End If
    Command8.Enabled = True
    
'10  SELECHO ("DEUDOR1")
'    NCLI$ = TRIM$(VALACT1("DEUDOR12"))       'Muestra Clientes
'    If NCLI$ <> "" Then
'       Call COPYSTRU("SELDOCAP", "SELDOCA1")
'       Call CARGALISEL("SELDOCA1", "CAJABANC", "CodiCom_Lar/A24;FechEmisi/D8;IHABECOMP/F12.2", "NuClien = " & NCLI$ & " AND IHABECOMP >0 AND (TIPOMOVIM='FVEN' OR TIPOMOVIM='RCOB')")
'15     Call SELECHO("SELDOCA1")             'Muestra Docu de Credito
'       DOCRED$ = TRIM$(VALACT1("SELDOCA1"))
'       If DOCRED$ <> "" Then
'          Call COPYSTRU("SELDOCAP", "SELDOCA2")
'          Call CARGALISEL("SELDOCA2", "CAJABANC", "CodiCom_Lar/A24;FechEmisi/D8;IDEBECOMP/F12.2", "NuClien = " & NCLI$ & " AND IDEBECOMP >0 AND TIPOMOVIM='FVEN'")
'          Call SELECHO("SELDOCA2")          ' Muestra Docu de Debito
'          DOCDEB$ = TRIM$(VALACT1("SELDOCA2"))
'          If DOCDEB$ <> "" Then
'             XX1$ = Space$(53)
'             Call REPLA(XX1$, DOCRED$, 1, 24)
'             Call REPLA(XX1$, DOCDEB$, 25, 24)
'20           RESP$ = OBJPLA$("INPUMANU", XX1$)      'Usuario ingresa imputacion manual
'             INPUMANUAL# = CVS(Mid(RESP$, 49, 4))
'             If INPUMANUAL# > 0.0005 Then
'                'Se obtienen los datos para grabar registro en APLICOBRA
'                CONDI$ = "CodiCom_Lar='" & DOCRED$ & "' AND NuClien=" & NCLI$ & " AND IHABECOMP >0 AND (TIPOMOVIM='FVEN' OR TIPOMOVIM='RCOB')"
'                IMPOCRE# = VALOBADA("CAJABANC", "IHABECOMP", CONDI$)
'                DOCREDCOR$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$)
'                '
'                CONDI$ = "CodiCom_Lar='" & DOCDEB$ & "' AND NuClien=" & NCLI$ & " AND IDEBECOMP >0 AND TIPOMOVIM='FVEN'"
'                ImpoDeb# = VALOBADA("CAJABANC", "IDEBECOMP", CONDI$)
'                DOCDEBCOR$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$)
'                TIPOCAMBIO# = VALOBADA("CAJABANC", "TIPO_CAM", CONDI$)
'
'                'Se extrae el iva del documento de debito
'                CONDI$ = "CodiCom_Lar='" & DOCDEB$ & "' AND NuClien=" & NCLI$ & " AND TIPOMOVIM='IVA'"
'                IMPIVADEBE$ = VALOBADA("CAJABANC", "IHABECOMP", CONDI$)
'
'                '
'                If (INPUMANUAL# >= IMPOCRE#) Or (INPUMANUAL# >= ImpoDeb#) Then
'                   Call COMUNI("Importe Ingresado Supera al Importe de los Comprobantes.")
'                   GoTo 20
'                 End If
'                 '
'                 Call BORRAREGISTROS("APLICOBRA", "CODOCAP_LAR='" & DOCRED$ & "'", REG&)
'                 Call BORRAREGISTROS("APLICOBRA", "CODOCOR_LAR='" & DOCDEB$ & "'", REG&)
'                 '
'                 Dim IMPUMANU As ADODB.Recordset
'                 Set IMPUMANU = New ADODB.Recordset
'25               On Error Resume Next
'                 IMPUMANU.Open FILTSQL$("APLICOBRA"), FLEXCONN, adOpenKeyset, adLockPessimistic
'                 If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'                   On Error GoTo 0
'                   Call CapturaErrorOpen
'                   GoTo 25
'                 End If
'                 On Error GoTo 0
'                 '
'                 NCLIE = XVALO(NCLI$)
'                 RAZON$ = RASOCLI(NCLIE)
'                 NUCLI$ = CODICLI$(NCLIE)
'                 '
'                 With IMPUMANU
'                 .AddNew   ' YA CONVERTIDO
'                 On Error Resume Next
'                 !CodiEmpr = CodiEmp%
'                 On Error GoTo 0
'                 !CodiEmpr = 0
'                 !CodiSucu = 0
'                 !nuclien = NCLI$
'                 !CODICLIE = NUCLI$
'                 !RasoClie = RAZON$
'                 !codocap_lar = DOCRED$
'                 !CODOCAP_COR = DOCREDCOR$
'                 !IMPOAPLI = INPUMANUAL#
'                 !refeapli = "Reaplicación de Comprobantes"
'                 !codocor_lar = DOCDEB$
'                 !CODOCOR_COR = DOCDEBCOR$
'                 !ItotOrig = ImpoDeb#
'                 !IIVAORIG = XVALO(IMPIVADEBE$)
'                 .Update
'                 '
'                 End With
'                 Text1 = NCLI$
'                 Call Command4_Click
'             End If
'           Else
'             GoTo 15
'           End If
'        Else
'          GoTo 10
'       End If
'    End If
       
End Sub

Private Sub Command9_Click()
Call recuperaapli
End Sub

Private Sub GRID1_DblClick()
   NGRID = GRID1.Row
   NLIST = List1.ListCount
   If (NLIST) >= NGRID - 1 Then
     List1.ListIndex = NGRID - 1
     List1_DblClick
   End If
End Sub

Private Sub GRID1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      PVER = 0: PVER2 = 0: KKF% = GRID1.TopRow
      For KKII% = 0 To KKF% - 1
         PVER = PVER + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
      Next KKII%
      KKII% = KKF% - 1
      While PVER2 < Y - GRID1.RowHeight(KKII%)
         KKII% = KKII% + 1
         PVER = PVER + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
         PVER2 = PVER2 + GRID1.RowHeight(KKII%) + GRID1.GridLineWidth + 7
      Wend
      NGRID = KKII%
      GRID1.Row = NGRID
      NLIST = List1.ListCount
      If (NLIST) >= NGRID - 1 Then
         List1.ListIndex = NGRID - 1
         Call List1_MouseDown(Button, Shift, X, Y)
      End If
   End If
End Sub

Private Sub Label1_DblClick()
   Call Text1_Change
End Sub

Private Sub Label2_DblClick(Index As Integer)
    'ESTA RUTINA LA REALIZA CON LAS MISMAS CONSULTAS DEL FRM CON
    If Index <> 4 Then Exit Sub
    NC = XVALO(Text1)
    If NC < 1 Then Exit Sub
     J& = 0

    'Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
    CAMPOS$ = "Fact.Anticipo/A20;Imp.Fact.Ant./F12.2;Imp.Aplicado/F12.2;Saldo-Fact./F12.2;Saldo-Tot/F12.2;Emisión/D8;Docu.Aplicado/A20;N.Aj./F5.O"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    FRMZELECHO.Caption = "Facturas Anticipo de Cliente N°: " & NC
    'ESTILO DE COLUMNA MANTIENE SIN SEPARACION SI REPITE EL NUMERO DE CLIENTE.
    FRMZELECHO.MSF1.MergeCells = flexMergeRestrictColumns
    FRMZELECHO.msf2.MergeCells = flexMergeRestrictColumns
    FRMZELECHO.MSF1.MergeCol(1) = True
    FRMZELECHO.msf2.MergeCol(1) = True
    '
    FRMZELECHO.MSF1.Rows = 1

    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE)AS IHABE ,FECHEMISI, DOCU_APLI,DOCU_ORIG  from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien =" & NC
    SQLX$ = SQLX$ + " AND TIPOMOVIM = 'FVEN'"
    SQLX$ = SQLX$ + " AND IMPODEBE > 0.005"
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN,DOCU_ORIG,DOCU_APLI,FECHEMISI"
    Set RS1 = READSET(SQLX$)
    With RS1
        Do While Not .EOF
            NC = XVALO(!nuclien)
            RASO$ = rasocli$(NC)
            NUFACTU$ = TRIM$(SAFETEXT$(!DOCU_ORIG))
            '
            SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE)AS IHABE ,FECHEMISI, DOCU_APLI,DOCU_ORIG, AJUSTE_NUMERO  from factuanti with(nolock)"
            SQLX$ = SQLX$ + " where nuclien =" & NC
            SQLX$ = SQLX$ + " AND  DOCU_ORIG ='" & NUFACTU$ & "'"
            SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'FVEN' or TIPOMOVIM = 'AJUS')"
            SQLX$ = SQLX$ + " GROUP BY NUCLIEN,DOCU_ORIG,DOCU_APLI,FECHEMISI, AJUSTE_NUMERO"
            SQLX$ = SQLX$ + " ORDER BY FECHEMISI"
            Set rs = READSET(SQLX$)
            With rs
                Do While Not .EOF
                   NC = XVALO(!nuclien)
                   '
                   DOCUORIG$ = SAFETEXT$(!DOCU_ORIG)

'                   If DOCUORIG_ANT$ <> "" And DOCUORIG_ANT$ <> DOCUORIG$ Then
'                      j& = j& + 1
'                   End If
                   If DOCUORIG_ANT$ <> DOCUORIG$ Then
                     SALDOFACTURA# = 0
                     If COLOR1 = &H8000000F Then
                        COLOR1 = &HFFFFFF
                     Else
                        COLOR1 = &H8000000F
                     End If
                   End If
                   J& = J& + 1
                   FRMZELECHO.msf2.Rows = J& + 1
                   FRMZELECHO.msf2.TextMatrix(J&, 1) = DOCUORIG$
                   FRMZELECHO.msf2.TextMatrix(J&, 2) = FORMATNUM$((XVALO(!IDEBE)), "G12.2")
                   FRMZELECHO.msf2.TextMatrix(J&, 3) = FORMATNUM$((XVALO(!IHABE)), "G12.2")
                   SALDOFACTURA# = SALDOFACTURA# + XVALO(!IDEBE) - XVALO(!IHABE)
                   SALDO# = SALDO# + XVALO(!IDEBE) - XVALO(!IHABE)
                   FRMZELECHO.msf2.TextMatrix(J&, 4) = FORMATNUM$(SALDOFACTURA#, "G12.2")
                   FRMZELECHO.msf2.TextMatrix(J&, 5) = FORMATNUM$(SALDO#, "G12.2")
                   FRMZELECHO.msf2.TextMatrix(J&, 6) = FECHATEX$(CVINT(!FECHEMISI))
                   FRMZELECHO.msf2.TextMatrix(J&, 7) = SAFETEXT$(!DOCU_APLI)
                   FRMZELECHO.msf2.TextMatrix(J&, 8) = SAFETEXT$(!AJUSTE_NUMERO)
                   Call COLOREAR_GRILLA_SOLO_UNA_AFILA(FRMZELECHO.msf2, COLOR1, FRMZELECHO.msf2.Rows - 1)
                   DOCUORIG_ANT$ = DOCUORIG$
                  .MoveNext
                Loop
            End With
            rs.Close
            Set rs = Nothing
          .MoveNext
        Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Show 1
End Sub


Private Sub Label5_Change()
   Text10 = TRIM$(Label5)
End Sub


Private Sub List3_DblClick()
   OPX% = 0: TIPOCOM% = 0
   RELIS$ = List3.TEXT
   FE% = FECHANUM(Mid$(RELIS$, 1, 8))
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   NUDOCU1$ = Str$(XVALO(NUDOCU$))
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Or InStr(UCase$(TIDOCUM$), "FACT.ANTICIPO") > 0 Or InStr(UCase$(TIDOCUM$), "N.C.ANTICIPO") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU2$ = Mid$(NUDOCU$, 4, 13)
     End If
     TIPOCOM% = 1
   ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Or InStr(UCase$(TIDOCUM$), "COMPENSACIÓN DE CUENTAS") > 0 Then
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU1$ = ""
        NUDOCU2$ = NUDOCU$
     End If
     TIPOCOM% = 2
   ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
     NUDOCU1$ = NUDOCU$
     NUDOCU2$ = NUDOCU$
     TIPOCOM% = 3
   End If
   OPX% = COMALTER%("Opciones de Trabajo:\\Documento Digital\Re-Emitir Comprobante")
   If OPX% = 2 Then
      Call REPRIFAC(NUDOCU$)
      Exit Sub
   End If
   '
Screen.MousePointer = 11
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FERVI") > 0 Then
   If OPX% > 0 And TIPOCOM% = 1 Then
      If FE% < FECHANUM("14/01/13") Then
         NUDOQUI& = FACTUFERVI&(TIDOCUM$, TRIM$(NUDOCU1$), FE%)
         If NUDOQUI& > 0 Then Call MUESTRADOC(NUDOQUI&, "", TIDOCUM$, NUDOCU1$, NUDOCU2$): Exit Sub
      End If
   End If
   
End If

   'MsgBox TIDOCUM$
   'MsgBox NUDOCU1$
   'MsgBox NUDOCU2$
   If TRIM$(TIDOCUM$) = "Fact.Anticipo" Then TIDOCUM$ = "Factura"
   If TRIM$(TIDOCUM$) = "N.C.Anticipo" Then TIDOCUM$ = "Nota de Credito"
   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, FECHATEX$(FE%))
   If DOCUNU& < 1 Then
        If OPX% > 0 And TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            CONSIDERAR_MOVIVIE% = 1
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
            CONSIDERAR_MOVIVIE% = 0
        Else
            Call COMUNI("Documento No Encontrado")
        End If
   Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
   End If
   Screen.MousePointer = 1

End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  '\\\OT-07-0046-RG-01/02/07///
   If KeyCode = vbKeyF3 Then
       If ARGU$ <> "" Then
          VUELTA% = 0
          INI& = List1.ListIndex + 2
    '
10        For U& = INI& To List1.ListCount
              If InStr((UCase$(Left$(List1.List(U& - 1), 77))), ARGU$) Then
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
              If InStr(TRIM$(UCase$(Left$(List3.List(U& - 1), 77))), ARGU$) Then
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

Private Sub mfgCuentaCorriente_DblClick()
    '
    FilaSeleccionada& = mfgCuentaCorriente.Row
    FE% = FECHANUM(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 1))
    TIDOCUM$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 2))
    If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
    If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
    If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
    NUDOCU$ = " " + CStr(XVALO(Mid$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 7), 11, 8))) + " "
    NUDOCU1$ = Str$(XVALO(NUDOCU$))
    If InStr(UCase$(TIDOCUM$), "CUOTA") > 0 Or InStr(UCase$(TIDOCUM$), "MATRICULA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
        NUDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 7))
        CONDI$ = "DOSECLAR = '" & NUDOCU$ & "' AND Nume_Clie= " & Text1
        CI1 = XVALO(VALOBADA("MOVISTO", "COD_INTE", CONDI$, "NOMESS"))
        CODAULA& = XVALO(VALOBADA("MASTER", "ID_AULA", "CODINT = " & CI1, "NOMESS"))
        IDPROF% = XVALO(VALOBADA("MASTER", "ID_PROFESOR", "CODINT = " & CI1, "NOMESS"))
        DES$ = VALOBADA("MASTER", "H_DESDE", "CODINT = " & CI1, "NOMESS")
        HAS$ = VALOBADA("MASTER", "H_HASTA", "CODINT = " & CI1, "NOMESS")
        HORARIOX$ = DES$ & " a " & HAS$
        DIA$ = VALOBADA("MASTER", "DESC_DIA1", "CODINT = " & CI1, "NOMESS")
        DIA2$ = VALOBADA("MASTER", "DESC_DIA2", "CODINT = " & CI1, "NOMESS")
        If DIAS2$ <> "" Then DIAS$ = DIAS$ + " Y " & DIAS2$
        CANTMEX = XVALO(VALOBADA("MASTER", "CantMeses", "CODINT = " & CI1, "NOMESS"))
        AUL$ = AULA$(CODAULA&)
        DetalleFactura$ = DetalleFactura$ + VALOBADA("CAJABANC", "NOTAS", "CODICOM_LAR='" & NUDOCU$ & "' AND NUCLIEN=" & Text1)
        DetalleFactura$ = DetalleFactura$ + "\Se Dicta en el Aula: " & AUL$
        DetalleFactura$ = DetalleFactura$ + "\Los días: " & DIA$
        DetalleFactura$ = DetalleFactura$ + "\En Horario de: " & HORARIOX$
        DetalleFactura$ = DetalleFactura$ + "\Con el profesor: " & NameProfesor$(IDPROF%)
        DetalleFactura$ = DetalleFactura$ + "\Meses de Duración: " & CANTMEX
        Call COMUNI("Referencia del Curso: \ \" & DetalleFactura$)
        
        Exit Sub
    ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Then
      If TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3)) <> "" Then
         NUDOCU$ = TRIM$(mfgCuentaCorriente.TextMatrix(FilaSeleccionada&, 3))
         NUDOCU1$ = ""
         NUDOCU2$ = NUDOCU$
      End If
      TIPOCOM% = 2
    ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
      NUDOCU1$ = NUDOCU$
      NUDOCU2$ = NUDOCU$
      TIPOCOM% = 3
      'EN EL CASO DEL RECIBO SE BUSCA LA FECHA REAL DE GRABACION POR QUE EN LOS CASOS QUE SE CAMBIO
      'NO MACHEA CON EL PDOC_TBL P Q ALLI GUARDA LA FECHA REAL
      CONDI$ = "NUMECOMP= " & XVALO(NUDOCU$) & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
      FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", CONDI$, "NOMESS"))
    End If
    
    Screen.MousePointer = 11
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "FERVI") > 0 Then
        If OPX% > 0 And TIPOCOM% = 1 Then
           If FE% < FECHANUM("14/01/13") Then
              NUDOQUI& = FACTUFERVI&(TIDOCUM$, TRIM$(NUDOCU1$), FE%)
              If NUDOQUI& > 0 Then Call MUESTRADOC(NUDOQUI&, "", TIDOCUM$, NUDOCU1$, NUDOCU2$): Exit Sub
           End If
        End If
    End If
    VARCONTROL$ = FECHATEX$(FE%)
    If TIDOCUM$ = "Recibo" Then VARCONTROL$ = TRIM$(Text2)
    DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, VARCONTROL$)
    If DOCUNU& < 1 Then
        If TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            CONSIDERAR_MOVIVIE% = 1
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
            CONSIDERAR_MOVIVIE% = 0
        Else
            Call COMUNI("Documento No Encontrado")
        End If
    Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
    End If
    
99  Screen.MousePointer = 1

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
       If rasocli$(NCX%) <> "" Then
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
   MONEMIS.ForeColor = &H624E28
   If MONRESU% > 1 Then
      MONEMIS.ForeColor = &HC0&
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   Call CARESCUE
   '
End Sub

Private Sub Text10_GotFocus()
    Command1.SetFocus
End Sub
Private Sub Text11_GotFocus()
    Command1.SetFocus
End Sub

Private Sub Text12_GotFocus()
    Command1.SetFocus
End Sub



Private Sub Text2_gotfocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub


Private Sub Text3_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text4_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text5_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text6_GotFocus()
    Command1.SetFocus
End Sub


Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command2_Click()
    '
    NCLIEN = ListaSeleccionClientes()
    If NCLIEN > 0 Then
        Text1.TEXT = CStr(NCLIEN)
    End If

'    Call SELECHO("DEUDOR1")
'    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
'    If VSELE$ <> "" Then
'       NCLIEN = XVALO(VSELE$)
'       Text1.TEXT = TRIM$(Str$(NCLIEN))
'       'Call CARESCUE
'    End If
    '
End Sub

Sub Command3_Click()
    '
    Dim oReporte As cGenrep
    Dim RST As ADODB.Recordset
    
    NCMINI = XVALO(Text1): NCMAXI = XVALO(Text1)
    If NCMINI <= 0 Then Exit Sub
    
    Screen.MousePointer = 11
    Set oReporte = New cGenrep
    Set RST = oReporte.GeneraReporteConTablaTemporal("RESCCLI", "LISTADO RESUMEN DE CUENTAS DEUDORAS", "Fecha/D8; Documento/A24; RefCurso(T=Ref-Curso)/A24; Debe/F14.2; Haber/F14.2; Saldo/F14.2")
    '
    MODORDE% = 1
    'If Option1(2) = True Then MODORDE% = 2
    '
    UJ& = 0
    'List1.Clear
    Call BLANARCH("!RESCUCLI")
    '
    ConsultaSql$ = "SELECT * FROM DEUDOR12 INNER JOIN CAJABANC WITH(NOLOCK) ON NUME_CLI=NUCLIEN WHERE STAT_CLI > 0"
    '
    If NCMINI > 0 Then CondicionSql$ = " AND DEUDOR12.NUME_CLI >= " & CStr(NCMINI)
    If NCMAXI > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.NUME_CLI <= " & CStr(NCMAXI)
    If DES% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA >= '" & FETEXCOR1$(DES%) & "'"
    If HAS% > 0 Then CondicionSql$ = CondicionSql$ & " AND CAJABANC.FECHCONTA <= '" & FETEXCOR1$(HAS%) & "'"
    If GRUCLI% > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.GRCO_CLI = " & CStr(GRUCLI%)
    If VENDE% > 0 Then CondicionSql$ = CondicionSql$ & " AND DEUDOR12.VEND_CLI = " & CStr(VENDE%)
    CondicionSql$ = CondicionSql$ & " AND (CAJABANC.TipoMovim = 'FVEN' OR CAJABANC.TipoMovim = 'RCOB' OR CAJABANC.TipoMovim = 'AJUD')"
    If MODORDE% = 1 Then CondicionSql$ = CondicionSql$ & " ORDER BY DEUDOR12.NUME_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"
    If MODORDE% = 2 Then CondicionSql$ = CondicionSql$ & " ORDER BY DEUDOR12.RASO_CLI, CAJABANC.FECHEMISI, CAJABANC.FEHOREAL"

    ConsultaSql$ = ConsultaSql$ & CondicionSql$
    Dim RstCuentaCorriente As ADODB.Recordset
    Set RstCuentaCorriente = READSET(ConsultaSql$)
    '
    NumeroClienteAnterior& = 0
    With RstCuentaCorriente
        Do While Not .EOF
          NumeroCliente& = !nume_cli
            
          If NumeroCliente& <> NumeroClienteAnterior& Then
            If NumeroClienteAnterior& <> 0 Then Call REGAPP("!RESCUCLI", String$(256, "="))

            REBLA$ = REGBLAN$("RESCUCLI")
            CLIE$ = TRIM$(SAFETEXT$(!raso_cli)) + " (" + TRIM$(Str$(NumeroCliente&)) + ")"
            DOMIE$ = TRIM$(SAFETEXT$(!domi_cli))
            LOCAE$ = TRIM$(SAFETEXT$(!Codi_Cli))
            LOCAF$ = SAFETEXT$(!loca_Cli)
            If TRIM$(LOCAF$) <> "" Then
              LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
            End If
            TELEE$ = "Telefax: " + TRIM$(SAFETEXT$(!Tele_Cli))
            '
            Call REPLA(REBLA$, MKI$(NCLIEN), 1, 2)
            Call REPLA(REBLA$, "Sres.", 97, 6)
            Call REPLA(REBLA$, CLIE$, 103, 48)
            Call REPLA(REBLA$, DOMIE$, 151, 48)
            Call REPLA(REBLA$, LOCAE$, 199, 32)
            Call REPLA(REBLA$, TELEE$, 231, 24)
            '
            YATRANSPO% = 0: ITRANSPO# = 0
            NumeroClienteAnterior& = NumeroCliente&
            
            CondicionSql$ = "(CAJABANC.TipoMovim = 'FVEN' OR CAJABANC.TipoMovim = 'RCOB' OR CAJABANC.TipoMovim = 'AJUD') AND FECHCONTA<'" & FETEXCOR1$(DES%) & "' AND NUCLIEN=" & CStr(NumeroCliente&)
            SaldoDebe# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CondicionSql$))
            SaldoHabe# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CondicionSql$))
            SaldoAnterior# = SaldoDebe# - SaldoHabe#
          End If
            
          If YATRANSPO% = 0 Then
              REGJ$ = REBLA$
              FEX = DES%
              Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
              Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
              Call REPLA(REGJ$, MKD$(SaldoAnterior#), 63, 8)
              Call REGAPP("!RESCUCLI", REGJ$)
              Call REGAPP("!RESCUCLI", REBLA$)
              YATRANSPO% = 1
              SaldoAcumulado# = SaldoAnterior#
          End If
          
          COCOLA$ = SAFETEXT$(!CodiCom_Lar)
          TIX$ = Left$(!CodiCom_Lar, 2)
          SG% = (-1)
          Observa$ = TRIM$(SAFETEXT$(!OBSERGEN))
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              If TIX$ = "AS" Then
                                    If Observa$ = "RECARGO" Then
                                        TIPOS$ = "Ajuste (Recargo)"
                                    ElseIf Observa$ = "DESCUENTO" Then
                                        TIPOS$ = "Ajuste (Descuento)"
                                    End If
                              End If
                              If TIX$ = "FC" Then
                                    TIPOS$ = SAFETEXT$(VALOBADA("MOVISTO", "IDENMOVI", "DOSECLAR='" & COCOLA$ & "' AND Nume_Clie=" & TRIM$(Text1)))
                                    COCOLA$ = Observa$
                              End If
          ImporteComprobante# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
          '
890       ImporteCbte# = ROUND(ImporteComprobante#)
          SaldoAcumulado# = SaldoAcumulado# + ImporteCbte#
          FechaEmision% = CVINT%(!FECHEMISI)
          NumeroComprobante& = !NumeComp
          Moneda% = ROUND(!Mone_Emis, 0)
          TipoCambio# = ROUND(!Tipo_Cam, 5)
          If TipoCambio# < 0.05 Then TipoCambio# = 1
          If TipoCambio# >= 0.01 Then
            ImporteDolar# = ROUND#(ImporteCbte# / TipoCambio#, 2)
          End If
          '
           With RST
                    .AddNew
                    .Fields(0) = CVDAT(FechaEmision%)
                    .Fields(1) = Left(TIPOS$, 24)
                    .Fields(2) = Left(COCOLA$, 24)
                    If ImporteCbte# > 0 Then
                      .Fields(3) = ImporteCbte#
                      .Fields(4) = 0
                    ElseIf ImporteCbte# < 0 Then
                      .Fields(3) = 0
                      .Fields(4) = Abs(ImporteCbte#)
                    Else
                      .Fields(3) = 0
                      .Fields(4) = 0
                    End If
                    .Fields(5) = SaldoAcumulado#
                    .Update
            End With
          '
         .MoveNext
        Loop
990 End With
    RstCuentaCorriente.Close
    Set RstCuentaCorriente = Nothing
    '
    oReporte.MostrarReporte
    '
    Screen.MousePointer = 1
    '
End Sub

Sub Command4_Click()
    '
    If MONRESU% > 1 Then
      Call MENSERR(24, "Proceso Imposible - Cambie Moneda a '" + TRIM$(DEMONECO$(1)) + "'")
      Exit Sub
    End If
    '
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = RGB(0, 0, 255)
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    '
    NCX = XVALO(Text1.TEXT)
    If NCX < 1 Or rasocli$(NCX) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       Exit Sub
    End If
    '
    If CONTROLGRABA% = 1 Then GoTo 80
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       Exit Sub
    End If
    '
80  Call REAPLICRE(NCX)
    Text1.TEXT = TRIM$(Str$(NCX))
    List1.Clear
    List2.Clear
    List3.Clear
    List3.Refresh
    
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    Call CARESCUE
    DoEvents
    '
End Sub

Private Sub Command5_Click()
    NCX = XVALO(TRIM$(Text1.TEXT))
    If NCX > 0 And rasocli$(NCX) <> "" Then
        Command5.Enabled = False
        Call NotaCli(NCX, 2)
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Cliente")
    End If
End Sub

Private Sub Form_Activate()
    Refresh
    If XVALO(Text1.TEXT) > 0 Then
        NCLIE = XVALO(Text1.TEXT)
      ElseIf NCLIE > 0 Then
        Text1.TEXT = NC%
    End If
    'List1.Clear
    'List2.Clear
    'If NC% > 0 Then
    '  If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
    '    Call CARESCUE
    '  End If
    'End If
End Sub

Private Sub Form_Load()
    '
'    Dim oCont As CONTROL
'    For Each oCont In Me.Controls
'        If TypeOf oCont Is TextBox Then
'            Call AplicaFuenteAnchoFijo(oCont, , 9)
'        ElseIf TypeOf oCont Is ListBox Then
'            Call AplicaFuenteAnchoFijo(oCont, , 10)
'        End If
'    Next
'    Call AplicaFuenteAnchoFijo(Text11, , 9)
'    Call AplicaFuenteAnchoFijo(Text10, , 9)
'    Call AplicaFuenteAnchoFijo(Text6, , 9)
'    Call AplicaFuenteAnchoFijo(Text12, , 9)
'
'    Call AplicaFuenteAnchoFijo(List1, , 10)
'    Call AplicaFuenteAnchoFijo(List2, , 10)
'    Call AplicaFuenteAnchoFijo(List3, , 10)
'    Call AplicaFuenteAnchoFijo(Label6, , 10)
'    Call AplicaFuenteAnchoFijo(Label5, , 10)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    MONRESU% = 1
    MONEMIS = DEMONECO$(MONRESU%)
    MONEMIS.ForeColor = &H624E28
    If MONRESU% > 1 Then
       MONEMIS.ForeColor = &HC0&
    End If
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONRESU%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    '
    ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
    Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A18;Ref-Curso/A18;Debe/F14.2;Haber/F14.2;Saldo/F14.2;Numero/A18", Me)
    For i = 1 To 6
         mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
    Next i
    mfgCuentaCorriente.Font.Size = 9
    mfgCuentaCorriente.ColWidth(7) = 0
    Call CARGA_ENCABEZADO(mfgCuentaCorrienteAnticipos, "Fecha/D8;Documento/A18;Mes-Número/A18;Debe/F14.2;Haber/F14.2;Saldo/F14.2", Me)
    For i = 1 To 6
         mfgCuentaCorrienteAnticipos.COL = i: mfgCuentaCorrienteAnticipos.CellFontBold = False: mfgCuentaCorrienteAnticipos.CellFontSize = 9
    Next i
    mfgCuentaCorrienteAnticipos.Font.Size = 9
    '
    ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA
    Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A19;Ref-Curso/A18;Emisión/D9;Importe/F15.2;Vencimiento/D10;Saldo ($)/F15.2", Me)
    For i = 1 To 6
         mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
    Next i
    mfgComposicionDeuda.Font.Size = 9
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Or MODOMONE$ = "BIMONETA" Then
       MODOMONEI% = 2
       '
       Label2(2).Top = 400
       Label2(3).Top = 720
       label(0).Top = 1040
       label(1).Top = 1360
       '
       MONEMIS.Visible = True
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE CUENTA CORRIENTE
       Call CARGA_ENCABEZADO(mfgCuentaCorriente, "Fecha/D8;Documento/A14;Ref-Curso/A18;Debe/F10.2;Haber/F10.2;Saldo/F12.2;Mon/A3;Cotiz/F7.2", Me)
       For i = 1 To 8
            mfgCuentaCorriente.COL = i: mfgCuentaCorriente.CellFontBold = False: mfgCuentaCorriente.CellFontSize = 9
       Next i
       mfgCuentaCorriente.Font.Size = 9
       '
       ' *** DEFINE ENCABEZADO DEL GRID DE COMPOSICION DE DEUDA EN MODALIDAD BIMONETARIA
       mfgComposicionDeuda.Cols = 1
       mfgComposicionDeuda.Rows = 1
       Call CARGA_ENCABEZADO(mfgComposicionDeuda, "Documento/A14;Ref-Curso/A18;Emisión/D8;Importe/F12.2;Vencimiento/D8;Saldo ($)/F12.2; Saldo (U$)/F12.2", Me)
       For i = 1 To 7
            mfgComposicionDeuda.COL = i: mfgComposicionDeuda.CellFontBold = False: mfgComposicionDeuda.CellFontSize = 9
       Next i
       mfgComposicionDeuda.Font.Size = 9
              
       Label8.Caption = "     Fecha        Documento                          Nro                      Debe                       Haber                       Saldo  ME    Cotiz."
       Label9.Caption = " Documento                         Nro        Emisión                          Importe     Vencim.                      Saldo ($)           Saldo (U$)"
       List1.FontName = "Lucida Console"
       List1.FontSize = 9
       List1.FontBold = False
       '
       List2.FontName = "Lucida Console"
       List2.FontSize = 9
       List2.FontBold = False
       '
       Frame7(0).Left = 1140
       Frame7(1).Left = 3030
       Frame7(2).Left = 3785
       Frame7(3).Left = 5230
       Frame7(4).Left = 6740
       Frame7(5).Left = 8200
       Frame7(6).Left = 8620
       Frame7(7).Left = 2100
       Frame7(9).Left = 2750
       Frame7(11).Left = 3800
       Frame7(13).Left = 5500
       Frame7(15).Left = 6550
       Frame7(8).Left = 8220
       Frame7(10).Left = 9540
       Frame7(12).Left = 9540
       '
       Label6.Left = 5160
       Label5.Left = 7635
       Label4.Visible = True
   End If
   Call APLICACIONSKINS(Me, Picture3)
   Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI = NUMECLI(CPROX$)
     If NCXI > 0 Then
       If REGICLI&(NCXI) > 0 Then
         Label26 = CPROX$
         'NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NCXI))
       End If
     End If
   End If
End Sub

Private Sub List1_DblClick()
   '
   OPX% = 0: TIPOCOM% = 0
   RELIS$ = List1.TEXT
   FE% = FECHANUM(Mid$(RELIS$, 1, 8))
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If TIDOCUM$ = "Recibo Cobranza" Then TIDOCUM$ = "Recibo"
   If TIDOCUM$ = "Ajuste de Saldo" Then TIDOCUM$ = "Ajuste Saldo Cliente"
   If TIDOCUM$ = "Comp.Ctas." Then TIDOCUM$ = "Compensación de Cuentas"
   NUDOCU$ = " " + TRIM$(Mid$(RELIS$, 27, 8)) + " "
   NUDOCU1$ = Str$(XVALO(NUDOCU$))
   If InStr(UCase$(TIDOCUM$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
     NUDOCU$ = TRIM$(NUDOCU$)
     While Len(NUDOCU$) < 8: NUDOCU$ = "0" + NUDOCU$: Wend
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU2$ = Mid$(NUDOCU$, 4, 13)
     End If
     TIPOCOM% = 1
   ElseIf InStr(UCase$(TIDOCUM$), "AJUSTE SALDO") > 0 Or InStr(UCase$(TIDOCUM$), "COMPENSACIÓN DE CUENTAS") > 0 Then
     If TRIM$(Mid$(RELIS$, 257, 18)) <> "" Then
        NUDOCU$ = TRIM$(Mid$(RELIS$, 257, 18))
        NUDOCU1$ = ""
        NUDOCU2$ = NUDOCU$
     End If
     TIPOCOM% = 2
   ElseIf InStr(UCase$(TIDOCUM$), "RECIBO") > 0 Then
     NUDOCU1$ = NUDOCU$
     NUDOCU2$ = NUDOCU$
     TIPOCOM% = 3
     'EN EL CASO DEL RECIBO SE BUSCA LA FECHA REAL DE GRABACION POR QUE EN LOS CASOS QUE SE CAMBIO
     'NO MACHEA CON EL PDOC_TBL P Q ALLI GUARDA LA FECHA REAL
     CONDI$ = "NUMECOMP= " & XVALO(NUDOCU$) & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
     FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", CONDI$, "NOMESS"))
'   ElseIf InStr(UCase$(TIDOCUM$), "Nota de Debito") > 0 Then
'     if cantregistros
   End If
'   OPX% = COMALTER%("Opciones de Trabajo:\\Documento Digital\Re-Emitir Comprobante")
'   If OPX% = 2 Then
'      Call REPRIFAC(NUDOCU$)
'      Exit Sub
'   End If
   '
Screen.MousePointer = 11
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FERVI") > 0 Then
   If OPX% > 0 And TIPOCOM% = 1 Then
      If FE% < FECHANUM("14/01/13") Then
         NUDOQUI& = FACTUFERVI&(TIDOCUM$, TRIM$(NUDOCU1$), FE%)
         If NUDOQUI& > 0 Then Call MUESTRADOC(NUDOQUI&, "", TIDOCUM$, NUDOCU1$, NUDOCU2$): Exit Sub
      End If
   End If
   
End If

'MsgBox TIDOCUM$
'MsgBox NUDOCU1$
'MsgBox NUDOCU2$
'SQLX$ = "UPDATE  PDOC_TBL SET FECHAYHORA = '01/FEB/1980' WHERE TIPODOCU = 'FACTURA' AND NUMEBUSQUE = '1951' AND NUMECOMPLE = '0003-00001951'"
'FLEXCONN.Execute (SQLX$)

   DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, FECHATEX$(FE%))
'MsgBox DOCUNU&
   If DOCUNU& < 1 Then
        If TIPOCOM% = 1 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
            CONSIDERAR_MOVIVIE% = 1
            Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$)
            CONSIDERAR_MOVIVIE% = 0
        Else
            Call COMUNI("Documento No Encontrado")
        End If
   Else
        OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
        Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
   End If
   Screen.MousePointer = 1

'   RACLI$ = TRIM$(Left$(RASOCLI$(XVALO(Text1.TEXT)), 8))
'   Screen.MousePointer = 11
'   VUELTA% = 0
'10 For UI& = ULTREG&("PDOCLST") To 1 Step -1
'     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 66)
'     FEMIX% = CVI(Mid$(XX$, 65, 2))
'     If FEMIX% > 8224 Then
'        If FEMIX% < FE% - 32 Then
'          GoTo 12
'        End If
'     End If
'     If InStr(XX$, NUDOCU$) > 4 Then
'       If InStr(XX$, TIDOCUM$) > 4 Then
'         'SE OMITE LA COMPARACION POR RAZON SOCIAL, ASUMIENDO QUE EXISTE
'         'LA POSIBILIDAD DE QUE ESTA CAMBIE
'         If InStr(XX$, RACLI$) > 16 Or VUELTA% > 0 Then
'           DOCUNU& = CVS(Left$(XX$, 4))
'           GoTo 5
'         End If
'       End If
'     End If
'   Next UI&
'12 VUELTA% = VUELTA% + 1
'   If VUELTA% < 2 Then GoTo 10
'   Screen.MousePointer = 1
'   '
'   If OPX% > 0 Then ' SI OPX% ES MAYOR A 0 POR QUE ES FACTURA O NOTA DE ..
'      Call MUESTRASTSCREEN(XVALO(Text1.TEXT), NUDOCU$, FE%, OK%)
'   End If
'
'   If OK% < 1 Then
'   Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
'   End If
'   Exit Sub
'   '
'5  OPPXX$ = ""
'   If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
'   Call MUESTRADOC(DOCUNU&, OPPXX$)
'   '
End Sub

Sub REPRIFAC(NUDOCU$)
   Call COMUNI("No Disponible por el Momento")
End Sub
'
Private Sub Text1_Change()
   List1.Clear
   List2.Clear
   List3.Clear
   Call ClearControls
   Label2(4).Visible = False
   mfgCuentaCorrienteAnticipos.Rows = 1
   mfgCuentaCorrienteAnticipos.Visible = False
   mfgCuentaCorriente.Rows = 1
   mfgComposicionDeuda.Rows = 1
   '
   If Abs(XVALO(Text1)) <= 999999 Then
     NCX = XVALO(Text1)
     If NCX > 0 Then
       If rasocli$(NCX) <> "" Then
         Call CARESCUE
         Call VERIFICA_SALDOFACTUANTICIPO
         '
         Check1.Visible = False: List3.Visible = False: Check1.Value = 0
         If TIENE_FACTURAS_ANTICIPO%(NCX) > 0 Then
            Check1.Visible = True
            List3.Visible = True
            Call CARESCUE_ANTICIPO
         End If
         '
         LMCR# = LIMICRE(NCX)  'MUESTRA QUE LA CUENTA DEL CLIENTE ESTA SUSPENDIDA
         If LMCR# < 0 Then
            Text11 = "Suspendida"
         End If
       End If
    End If
   End If
   
End Sub
Sub VERIFICA_SALDOFACTUANTICIPO()

    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where  NUCLIEN = " & XVALO(Text1)
    SQLX$ = SQLX$ + " and (tipomovim = 'FVEN'"
    SQLX$ = SQLX$ + " OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN"
    SQLX$ = SQLX$ + " HAVING (SUM(impodebe))- (sum(impohabe)) > 0.1"
    '
    Set rs = READSET(SQLX$)
    If Not (rs.EOF And rs.BOF) Then
       Label2(4).Visible = True
    End If

End Sub
Private Sub Text1_DblClick()
    '
    NCLIEN = ListaSeleccionClientes()
    If NCLIEN > 0 Then
        Text1.TEXT = CStr(NCLIEN)
    End If
    
'    Call SELECHO("DEUDOR1")
'    VSELE$ = TRIM$(VALACT1$("DEUDOR1"))
'    If VSELE$ <> "" Then
'       NCX = XVALO(VSELE$)
'       Text1.TEXT = TRIM$(Str$(NCX))
'       'Call CARESCUE
'    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       'If XVALO(Text1.TEXT) > 0 Then
       '   If XVALO(Text1.TEXT) <= 32767 Then
       '      NC% = XVALO(Text1.TEXT)
       '      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
       '         Call CARESCUE
       '         Exit Sub
       '      End If
       '   End If
       'End If
       '
       'Beep
       'Text1.TEXT = ""
       '
    End If
    '
End Sub

Sub Text6_DBLCLICK()
Exit Sub
   Screen.MousePointer = 11
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
        NCX = XVALO(Text1.TEXT)
        'SELCHE04.LICHESEL.Style = 0
        SELCHE07.Caption = "Cheques Pendientes de Acreditaciòn: " + rasocli$(NCX)
        SELCHE07.List1.Clear
        SELCHE07.LICHESEL.Clear
        '
        FEXI% = DIANTE(DIANTE(HOY(HOS$)))
        'AYER = CDbl(CVDAT(FEXI%))
        AYER = FETEXCOR1$(FEXI%)
                       
15      Call APERBASDAT("CABAN")
        Dim CABATEMP As ADODB.Recordset
        Set CABATEMP = New ADODB.Recordset
        CABATEMP.CursorLocation = adUseClient
        
        SQLX$ = "SELECT * FROM CAJABANC "
        SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
        SQLX$ = SQLX$ + "AND NuClien = " & NCX & " "
        SQLX$ = SQLX$ + "AND Status >0 "
        SQLX$ = SQLX$ + "AND FechAcred >= '" & AYER & "'"
        SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
        '
25      On Error Resume Next
        CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        JJ& = 0: TRANSPOR# = 0
        FORMALTER.LINOSORT.Clear
        With CABATEMP
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do While Not .EOF
              If TRIM(!OPCIMOVIM) <> "CH" Then GoTo 19 ' no es un cheque
              If IsNull(!FECHENTRE) = False Then
                If CVINT(!FECHENTRE) > 0 Then
                  GoTo 19    ' salio de cartera
                End If
              End If
              '
              FEACRED% = CVINT(!fechacred)
              BANUFE$ = !DescriMov
              ICHEQ# = !ValAbsComp
              '
              If TRIM$(Left$(BANUFE$, 12)) = "" Then
                Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
              End If
              RECHE& = !NumeComp
              '
              Z$ = "A" + Space$(80)
              Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
              Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
              Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
              Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
              SELCHE07.List1.AddItem Z$
              '
19          .MoveNext
            Loop
          End If
          On Error GoTo 0
        End With
        '
        SUCHETO# = 0
        For i& = 1 To SELCHE07.List1.ListCount
           Z$ = Mid$(SELCHE07.List1.List(i& - 1), 9)
           SELCHE07.LICHESEL.AddItem Z$
           SUCHETO# = SUCHETO# + XVALO(Mid$(Z$, 41, 12))
        Next i&
        '
        SELCHE07.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
        SELCHE07.Label4.Refresh
        '
        Screen.MousePointer = 1
        SELCHE07.Show 1
        DoEvents
        '
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARESCUE()
Static VUELTA%
    '
    If ATENTI% = 1 Then
       Exit Sub
    End If
    '
10  Screen.MousePointer = 11
    NCX7 = XVALO(Text1.TEXT)
    If NCX7 > 0 Then
        List1.Clear
        List1.Refresh
        List2.Clear
        List2.Refresh
        List3.Clear
        List3.Refresh
        Text2.TEXT = rasocli$(NCX7)
        Text3.TEXT = DOMICLI$(NCX7)
        Text4.TEXT = CPOSCLI$(NCX7)
        Text5.TEXT = LOCACLI$(NCX7)
        Text7.TEXT = TEL1CLI$(NCX7)
        Label26 = CODICLI$(NCX7)
        Label5.Caption = ""
        Text11.TEXT = TRIM$(FORMATNUM$(LIMICRE(NCX7), "G14.2"))
        Refresh
        Text6.TEXT = TRIM$(FORMATNUM$(CHEPEACRE#(NCX7), "G14.2"))
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    List3.Clear
    List3.Refresh
    SALDO# = 0: SARRAS# = 0
    
    If NCX7 = 459 And InStr(1, EMPREAC$, "FERVI") > 0 Then
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001227'", REGAF&)
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001441'", REGAF&)
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001490'", REGAF&)
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001538'", REGAF&)
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001575'", REGAF&)
        Call BORRAREGISTROS("CAJABANC", "CODICOM_LAR='AS-X-0000-00001626'", REGAF&)
    End If
    '
'    RFF$ = RECFILT$("CUECOCLI", 1, MKI$(NCX7))
'    For KU& = 1 To Len(RFF$) Step 4
'      RELE& = CVS(Mid$(RFF$, KU&, 4))
'      XX$ = REGLEIDO$("CUECOCLI", RELE&)
'      TIPOS$ = Mid$(XX$, 5, 16)
'      IMPO# = ROUND(CVD(Mid$(XX$, 25, 8)))
'      SALDO# = SALDO# + IMPO#
'      FEX = CVI(Mid$(XX$, 3, 2))
'      NROL& = CVS(Mid$(XX$, 21, 4))
'      '
'      Z$ = Space$(96)
'      Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
'      Call REPLA(Z$, TIPOS$, 11, 16)
'      Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
'      IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
'      SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
'      If IMPO# >= 0.005 Then
'          Call REPLA(Z$, IMPTEX$, 35, 13)
'        ElseIf IMPO# <= (-0.005) Then
'          Call REPLA(Z$, IMPTEX$, 50, 13)
'        Else
'          Call REPLA(Z$, "  ...... comp.anulado ...... ", 35, 29)
'      End If
'      Call REPLA(Z$, SALTEX$, 63, 14)
'      List1.AddItem Z$
'    Next KU&
'    DoEvents
'    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    '
'    Call APERBASDAT("CABAN")
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
'    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
'    SQLX$ = SQLX$ + "ORDER BY FechEmisi, fehoreal ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim CuecoTemp As ADODB.Recordset
'    Set CuecoTemp = New ADODB.Recordset
'    CuecoTemp.CursorLocation = adUseClient
'25  On Error Resume Next
    'CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    
    ' LIMPIA GRILLA CUENTA CORRIENTE Y LABELS DE TOTALES
    ' OBTIENE LA CANTIDAD DE REGISTROS DE LA CONSULTA DE CUENTA CORRIENTE
    ' PARA AGREGAR DICHA CANTIDAD DE FILAS A LA GRILLA mfgCuentaCorriente
    mfgCuentaCorriente.Rows = 1: mfgComposicionDeuda.Rows = 1: Label6 = "": Label5 = "": Label4 = "":
    CondicionSql$ = " NuClien = " & NCX7 & " AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD')"
    CantidadRegistros& = CantRegistros("CAJABANC", CondicionSql$)
    mfgCuentaCorriente.Rows = CantidadRegistros& + 1
    
    SQLX$ = " EXECUTE CTACTE_CLIENTE  " & NCX7
    SQLX$ = SQLX$ + "," & CodiEmp%

    Set CuecoTemp = FLEXCONN.Execute(SQLX$)
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
    On Error GoTo 0
    '
    If NCX7 <> XVALO(Text1) Then GoTo 10  ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        On Error GoTo 0
'        .MoveFirst
        Do While Not .EOF
          COCOLA$ = !CodiCom_Lar
          TIX$ = Left$(!CodiCom_Lar, 2)
          SG% = (-1)
          Observa$ = TRIM$(SAFETEXT$(!OBSERGEN))
          TIPOS$ = "Factura":
                              If TIX$ = "AS" Then
                                    If Observa$ = "RECARGO" Then
                                        TIPOS$ = "Ajuste (Recargo)"
                                    ElseIf Observa$ = "DESCUENTO" Then
                                        TIPOS$ = "Ajuste (Descuento)"
                                    End If
                              End If
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
          IMPO# = !IDEBECOMP - !IHABECOMP
          '
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!FECHEMISI)
              NROL& = !NumeComp
              Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
              Call REPLA(Z$, TIPOS$, 11, 16)
              Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
                TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
                While Len(TTX1$) < 26: TTX1$ = "." + TTX1$ + ".": Wend
              Call REPLA(Z$, "  " + Left$(TTX1$, 26) + " ", 35, 29)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          '
890       IMPO1# = ROUND(IMPO#)    '(Int(CDbl(100 * IMPO# + 0.5))) / 100
          SALDO# = SALDO# + IMPO1#  '(Int(100 * IMPO# + 0.5)) / 100
          FEX = CVINT%(!FECHEMISI)
          NROL& = !NumeComp
          '
          Z$ = Space$(128)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(Z$, TIPOS$, 11, 16)
          'Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
          Call REPLA(Z$, FORMATNUM$(NROL&, "F8.0"), 27, 8)
          IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          If IMPO# >= 0.005 Then
              Call REPLA(Z$, IMPTEX$, 35, 13)
            ElseIf IMPO# <= (-0.005) Then
              Call REPLA(Z$, IMPTEX$, 50, 13)
            Else
              Call REPLA(Z$, "  ...... comp.anulado ...... ", 35, 29)
          End If
          Call REPLA(Z$, SALTEX$, 63, 14)
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
             Call REPLA(Z$, FORMATNUM$(MONEM%, "I2"), 103, 2)
             Call REPLA(Z$, COCOLA$, 105, 24)
             '
          End If
          'SE AGREGA EL NUMERO COMPLETO DEL AJUSTE PARA QUE SE PUEDA DIFERENCIAR SI
          'SE REALIZA UN AJUSTE EL MISMO DIA.
895       If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "NC" Or TIX$ = "AS" Or TIX$ = "CC" Then
              Z$ = Z$ + Space$(128) + COCOLA$ ' Mid$(COCOLA$, 4, 13)
          End If
          
          If TIX$ = "FC" Then
                TIPOS$ = SAFETEXT$(VALOBADA("MOVISTO", "IDENMOVI", "DOSECLAR='" & COCOLA$ & "' AND Nume_Clie=" & TRIM$(Text1)))
                COCOLA$ = Observa$
          End If
          '
          'List1.AddItem Z$
          FilaCuentaCorriente& = FilaCuentaCorriente& + 1
          'mfgCuentaCorriente.Rows = mfgCuentaCorriente.Rows + 1
          
          ' *** OMAR
           If FilaCuentaCorriente& >= mfgCuentaCorriente.Rows Then mfgCuentaCorriente.Rows = FilaCuentaCorriente& + 1

          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 3) = COCOLA$
          If IMPO# > 0 Then
              mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          ElseIf IMPO# < 0 Then
              mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          Else
              mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = "Anulado"
          End If
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
          mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 7) = !CodiCom_Lar
          '
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 4) = "..........."
                    LeyendaMoneda$ = TRIM$(DEMONECO$(MONEM%))
                    If Len(LeyendaMoneda$) < 11 Then Blancos = Int((10 - Len(LeyendaMoneda$)) / 2)
                    LeyendaMoneda$ = LeyendaMoneda$ & Space(Blancos)
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 5) = LeyendaMoneda$
                  mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 6) = "............"
            End If
          End If
          '
          If MODOMONEI% = 2 Then
             MONEM% = XVALO(!Mone_Emis)
             TICAM = ROUND(XVALO(!Tipo_Cam), 5)
             If TICAM < 0.05 Then TICAM = 1
             '
             mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 7) = SIMBOLPES$(MONEM%)
             If MONEM% > 1 Then mfgCuentaCorriente.TextMatrix(FilaCuentaCorriente&, 8) = TRIM$(FORMATNUM$(TICAM, "F9.5"))
          End If
          '
         .MoveNext
        Loop
'      End If
      On Error GoTo 0
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh
    '
    ' OBTIENE LA CANTIDAD DE REGISTROS DE LA CONSULTA DE CUENTA CORRIENTE
    ' PARA AGREGAR DICHA CANTIDAD DE FILAS A LA GRILLA mfgCuentaCorriente
    FilaCompDeuda& = 0
    CondicionSql$ = " NuClien = " & NCX7 & " AND (abs(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005"
    If MODOMONEI% = 2 Then
       CondicionSql$ = CondicionSql$ + " OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
    Else
       CondicionSql$ = CondicionSql$ + ") "       ' PARA CERRAR EL PARENTESIS
    End If
    CantidadRegistros& = CantRegistros("SADEUPEN", CondicionSql$)
    mfgComposicionDeuda.Rows = CantidadRegistros& + 1
    FilaCompDeuda& = 0
    '
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
    SQLX$ = SQLX$ + "AND (ABS(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005 "
    If MODOMONEI% = 2 Then
         SQLX$ = SQLX$ + "OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
       Else
         SQLX$ = SQLX$ + ") "       ' PARA CERRAR EL PARENTESIS
    End If
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    If MODOMONEI% < 1 Then MODOMONEI% = 1
'    SQLX$ = " EXECUTE SALDOCLIENTE  " & MODOMONEI%
'    SQLX$ = SQLX$ + " , " & NCX7
'    SQLX$ = SQLX$ + " , " & CodiEmp%
    
'    Dim FacPenTemp As ADODB.Recordset
'    Set FacPenTemp = New ADODB.Recordset
'    FacPenTemp.CursorLocation = adUseClient

    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = FLEXCONN.Execute(SQLX$)
    'SQLX$ = SQLX$ + "," & CodiEmp%

    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
26  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0

    If NCX7 <> XVALO(Text1) Then GoTo 10   ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    SUDEUDOL# = 0
    With FacPenTemp
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        On Error GoTo 0
'        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CodiCom_Lar, 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 ''\\\OT-06-0112-WAR-30/03/06///
                              '
          If TIX$ = "AS" Then
               Comprobante$ = SAFETEXT$(!CODICOM_COR)
               Observa$ = SAFETEXT$(VALOBADA("CAJABANC", "OBSERGEN", "CODICOM_LAR='" & Comprobante$ & "' AND NUCLIEN=" & TRIM$(Text1)))
               If Observa$ = "RECARGO" Then
                   TIPOS$ = "Ajuste (Recargo)"
               ElseIf Observa$ = "DESCUENTO" Then
                   TIPOS$ = "Ajuste (Descuento)"
               End If
          End If
                              
          IMPO# = (-1) * ROUND(!IBruComp) * SG%
          SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = ROUND(!SaldDebe - !SALDAcre)
          SUMADEU# = SUMADEU# + SALDO#    ' (Int(100 * SALDO# + 0.5)) / 100
          '
          NROL& = XVALO(Mid$(!CODICOM_COR, 4, 6))
          'NROX$ = CSTRING$(MKS$(NROL&), 3, 8, 0, 2)
          NROX$ = FORMATNUM$(NROL&, "F8.0")
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          SALTEXDOL$ = ""
          '
          COCOLAR$ = !CodiCom_Lar
          If TIX$ = "FC" Then
                TIPOS$ = SAFETEXT$(VALOBADA("MOVISTO", "IDENMOVI", "DOSECLAR='" & COCOLAR$ & "' AND Nume_Clie=" & TRIM$(Text1)))
                COCOLAR$ = SAFETEXT$(VALOBADA("CAJABANC", "OBSERGEN", "CODICOM_LAR='" & COCOLAR$ & "' AND NUCLIEN=" & TRIM$(Text1)))
          End If
          '
          If MODOMONEI% = 2 Then
             SALDOLA# = ROUND(!SaldDebe_DOLA - !SaldAcre_DOLA)
             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
             GoTo 20
          End If
          '
          ' rutina incluida para procesamiento bimonetario
          RECABIN& = 0
'          If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
'             MONEMX% = 0
'             On Error Resume Next   ' por si no está el campo
'             MONEMX% = !Mone_Emis
'             On Error GoTo 0
'             If MONEMX% < 1 Then
'                COCOLAR$ = TRIM$(!CODICOM_LAR)
'                RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7))
'                For KUL& = Len(RFG$) - 3 To 1 Step -4
'                   RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'                   XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'                   If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
'                      MONEMX% = CVI(Mid$(XXC$, 31, 2))
'                      GoTo 15
'                   End If
'                Next KUL&
'15              If MONEMX% < 1 Then MONEMX% = 1
'             End If
'             '
'             If MONEMX% > 1 Then
'               ' LO QUE SIGUE ES UN ENJUAGUE PARA GRAMY PARA RECUPERAR LA MONEDA DE EMISIÓN
'               If MONEMX% > ULTREG&("TAMONED") Then
'                 Call COMUNI("Moneda no Válida en Comprobante '" + COCOLAR$ + "'.\Debe Seleccionar la Moneda de Emisión.")
'                 Call SELECHO("TAMONED")
'                 If XVALO(VALACT1$("TAMONED")) > 0 Then
'                   MONEMX% = XVALO(VALACT1$("TAMONED"))
'                   If RECABIN& > 0 Then
'                     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'                     Call REPLA(XXC$, MKI$(MONEMX%), 31, 2)
'                     Call GRAREG("CARBIMON", XXC$, RECABIN&)
'                   End If
'                 End If
'               End If
'               '
'               RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7))
'               For KUL& = Len(RFG$) - 3 To 1 Step -4
'                 RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'                 XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'                 If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
'                   SALDOPES# = ROUND(CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8)))
'                   If Abs(SALDOPES# - SALDO#) <= 0.005 Then
'                     SALTEX$ = FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F14.2")
'                     SBPES$ = TRIM$(SIMBOLPES$(MONEMX%))
'                     If SBPES$ <> "" Then
'                       Call REPLA(SALTEX$, SBPES$, 1, Len(SPPES$))
'                     End If
'                     GoTo 20
'                   End If
'                 End If
'               Next KUL&
'             End If
'          End If
         
          ' fin rutina para procesamiento bimonetario
          '
20        Z$ = Space$(80)
          Call REPLA(Z$, TIPOS$, 1, 16)
          Call REPLA(Z$, NROX$, 17, 8)
          Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
          Call REPLA(Z$, IMPTEX$, 37, 14)
          Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
          Call REPLA(Z$, SALTEX$, 63, 14)
          Call REPLA(Z$, SALTEXDOL$, 77, 13)
          SUDEUDOL# = SUDEUDOL# + SALDOLA#
          '
          'List2.AddItem Z$
          '
          FilaCompDeuda& = FilaCompDeuda& + 1
          'mfgComposicionDeuda.Rows = mfgComposicionDeuda.Rows + 1
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 1) = TIPOS$
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 2) = COCOLAR$
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 3) = FECHATEX$(FEMI)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 5) = FECHATEX$(FEVE)
          mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 6) = FORMATNUM$(XVALO(SALTEX$), "G14.2")
          If MODOMONEI% = 2 Then mfgComposicionDeuda.TextMatrix(FilaCompDeuda&, 7) = FORMATNUM$(XVALO(SALTEXDOL$), "G14.22")
          '
        .MoveNext
        Loop
'      End If
'      On Error GoTo 0
    End With
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    '
    If MONRESU% <= 1 Then
       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
          ATENTI% = 1
          Call REAPLICRE(NCX7)
          If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then ATENTI% = 0: GoTo 899 'provisorio FFFERVI
          GoTo 100
          '
          ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
          ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR COBRAR APARECIA
          ' DESACTUALIZADO.
          '
'          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'          If COMALTER%(TTXX$) = 2 Then
'              Call REAPLICRE(NCX7)
'              GoTo 100
'            Else
'              ATENTI% = 0
'              Call CIERRARCH("*.*")
'              Call FINAL("")
'          End If
       End If
    End If
    '
899 List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label5.Caption = TRIM$(FORMATNUM$(SUMADEU#, "G14.2"))
    Label5.Refresh
    Label6.Caption = TRIM$(FORMATNUM$(SUMAFAC#, "G14.2"))
    Label6.Refresh
    Label4.Caption = TRIM$(FORMATNUM$(SUDEUDOL#, "G14.2"))
    Label4.Refresh
    '
    mfgCuentaCorriente.Rows = FilaCuentaCorriente& + 1
    mfgComposicionDeuda.Rows = FilaCompDeuda& + 1
    Call POSICIONAR_ULTIMAFILA(mfgCuentaCorriente)
    Call POSICIONAR_ULTIMAFILA(mfgComposicionDeuda)
    '
999 Screen.MousePointer = 1
    '
    If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then GoTo 900 'provisorio FFFERVI

    If MONRESU% <= 1 Then
       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
          ATENTI% = 1
          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
          If Visible = True Then
              TTXX$ = TTXX$ + " Elegida."
            Else
              TTXX$ = TTXX$ + ":\   \'" + TRIM$(rasocli$(NCX7)) + "'"
          End If
          TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
          '
          If COMALTER%(TTXX$) = 2 Then
              Call REAPLICRE(NCX7)
              GoTo 100
            Else
              ATENTI% = 0
              Call CIERRARCH("*.*")
              Call FINAL("")
          End If
       End If
    End If
    '
900 DISPONI = XVALO(Text11.TEXT) - XVALO(Text10.TEXT) - XVALO(Text6.TEXT)
    Text12.ForeColor = &H404080
    If DISPONI < 0 Then Text12.ForeColor = &HFF&
    Text12.TEXT = TRIM$(FORMATNUM$(DISPONI, "G14.2"))
    '
    Screen.MousePointer = 1
    '
End Sub

Sub CARESCUE_ANTICIPO()

Static VUELTA%
    '
    If ATENTI% = 1 Then
       Exit Sub
    End If
    '
10  Screen.MousePointer = 11
    NCX7 = XVALO(Text1.TEXT)
    '
100 ' *** OCULTA GRID PRINCIPAL DE CUENTA CORRIENTE Y ACTIVA EL DE LA MODALIDAD DE ANTICIPOS
    mfgCuentaCorrienteAnticipos.Rows = 1
    mfgCuentaCorrienteAnticipos.ZOrder (0)
    mfgCuentaCorrienteAnticipos.Visible = True
    '
    List3.Clear
    List3.Refresh
    YAPASO% = 0
    Call REMPLAGRID(Me, GRID1, Picture1, Label8, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    SALDO# = 0: SARRAS# = 0
    mfgCuentaCorrienteAnticipos.Rows = 1
    CondicionSql$ = " NuClien = " & NCX7 & " AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD')"
    CantidadRegistros& = CantRegistros("CAJABANC", CondicionSql$)
    mfgCuentaCorrienteAnticipos.Rows = CantidadRegistros& + 1
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, fehoreal ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    If NCX7 <> XVALO(Text1) Then Call CARESCUE: GoTo 10   ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          COCOLA$ = !CodiCom_Lar
          TIX$ = Left$(!CodiCom_Lar, 2)
          SG% = (-1)
          ESFACTUANTICIPO% = 0 'INICIALIZO NUEVAMENTE

          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              If EsFacturaAnticipo%(NCX7, COCOLA$) > 0 Then
                                ESFACTUANTICIPO% = 1
                                If TIPOS$ = "Factura" Then TIPOS$ = "Fact.Anticipo"
                                If TIPOS$ = "Nota de Credito" Then TIPOS$ = "N.C.Anticipo"
                              End If
                              'ACA VALIDAR SI ES UNA FACTURA ANTICIPO, DEBERA IR EN 0
                              '-----------------------------------------------------------------------------------
                              'SI ES UNA APLICACION DEBERIA TRAER EL VALOR DE LO REMITIDO VER CON VALOUNITIVA
                              '
          IMPO# = !IDEBECOMP - !IHABECOMP
          If ESFACTUANTICIPO% > 0 Then
              IMPO# = 0
          Else
             If EsAplicacionDeAnticipo%(NCX7, COCOLA$) > 0 Then
                IMPO# = ImporteRealFactura(NCX7, COCOLA$) ' es el importe real de la facturas 0,01
             End If
          End If
          '
          'SE INTENTA TRAER LOS AJUSTES DE LA TABLA FACTUANTI  Y SE CARGA EN UN VECTOR
          If YAPASO% < 1 Then
            Dim NUMEAJUSTE&(), FECHAAJUSTE%(), DOCUAJUSTE$(), IMPODEBEAJUSTE#(), IMPOHABEAJUSTE#()
            YAPASO% = 1
            IX& = 0: UltimoInd& = 0 'se usa mas abajo
            NUMAJUS& = 0
'            FE5% = FECHANUM("26/05/15"): FECHAINICIO_IMPLEMNTACION = FETEXCOR1$(FE5%) ' FECHA EN LA CUAL SE IMPLEMENTO EL CAMBIO DE LA NUEVA CUENA CORRIENTE.
             'AND FechEmisi <= '" & FECHAINICIO_IMPLEMNTACION & "' 'SE QUITA PARA QUE MUESTRE TODOS
            CONDI$ = "AJUSTE_NUMERO > " & NUMAJUS& & "  AND NUCLIEN = " & NCX7 & " AND TIPOMOVIM = 'AJUS' ORDER BY FECHEMISI"
            SQLX$ = "SELECT * FROM FACTUANTI WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$
            Set rs = READSET(SQLX$)
            With rs
                Do While Not .EOF
                    ReDim Preserve FECHAAJUSTE%(IX&), DOCUAJUSTE$(IX&), DOCUAJUSTE$(IX&), IMPODEBEAJUSTE#(IX&), IMPOHABEAJUSTE#(IX&), NUMEAJUSTE&(IX&)
                    FECHAAJUSTE%(IX&) = CVINT(!FECHEMISI)
                    'SI ES NOTA DE CREDITO MUESTRA EL TEXTO Ajuste x N.C. Y EL NUMERO DE NUMECOM DE LA N.C.
                    DOCUAPLICADO$ = SAFETEXT$(!DOCU_APLI)
                    NUMEROAJUSTE& = XVALO(!AJUSTE_NUMERO)
                    
                    If Left$(DOCUAPLICADO$, 3) = "NC-" Then
                       DOCUAPLICADO$ = "Ajuste x N.C."
                       NUMEROAJUSTE& = XVALO(Mid$(DOCUAPLICADO$, 9, 8))
                    End If
                    DOCUAJUSTE$(IX&) = DOCUAPLICADO$
                    NUMEAJUSTE&(IX&) = NUMEROAJUSTE&
                         
                    IMPODEBEAJUSTE#(IX&) = XVALO(!IMPODEBE)
                    IMPOHABEAJUSTE#(IX&) = XVALO(!IMPOHABE)
                    
                    QW = !FECHEMISI
                    IX& = IX& + 1
                   .MoveNext
                Loop
            End With
          End If
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
              Z$ = Space$(128)
              FEX = CVINT%(!FECHEMISI)
              NROL& = !NumeComp
              Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
              Call REPLA(Z$, TIPOS$, 11, 16)
              Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
                TTX1$ = " Moneda: " + TRIM$(DEMONECO$(MONEM%)) + " "
                While Len(TTX1$) < 26: TTX1$ = "." + TTX1$ + ".": Wend
              Call REPLA(Z$, "  " + Left$(TTX1$, 26) + " ", 35, 29)
              GoTo 895
            End If
            TICAM = ROUND(!Tipo_Cam, 5)
            If TICAM < 0.05 Then TICAM = 1
            IMPO# = IMPO# / TICAM
          End If
          '
890       IMPO1# = ROUND(IMPO#)    '(Int(CDbl(100 * IMPO# + 0.5))) / 100
          SALDO# = SALDO# + IMPO1#  '(Int(100 * IMPO# + 0.5)) / 100
          FEX = CVINT%(!FECHEMISI)
          NROL& = !NumeComp
          '
          Z$ = Space$(128)
          Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
          Call REPLA(Z$, TIPOS$, 11, 16)
          Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 27, 8)
          IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
          If IMPO# >= 0.005 Then
              Call REPLA(Z$, IMPTEX$, 35, 13)
            ElseIf IMPO# <= (-0.005) Then
              Call REPLA(Z$, IMPTEX$, 50, 13)
            Else
              If ESFACTUANTICIPO% < 1 Then
                Call REPLA(Z$, "  ...... comp.anulado ...... ", 35, 29)
              End If
          End If
          Call REPLA(Z$, SALTEX$, 63, 14)
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
             Call REPLA(Z$, FORMATNUM$(MONEM%, "I2"), 103, 2)
             Call REPLA(Z$, COCOLA$, 105, 24)
             '
          End If
          'SE AGREGA EL NUMERO COMPLETO DEL AJUSTE PARA QUE SE PUEDA DIFERENCIAR SI
          'SE REALIZA UN AJUSTE EL MISMO DIA.
895       If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "NC" Or TIX$ = "AS" Or TIX$ = "CC" Then
              Z$ = Z$ + Space$(128) + COCOLA$ ' Mid$(COCOLA$, 4, 13)
          End If

          FilaCuentaCorriente& = FilaCuentaCorriente& + 1
          'mfgCuentaCorrienteAnticipos.Rows = mfgCuentaCorrienteAnticipos.Rows + 1
          mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FEX)
          mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
          mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 3) = COCOLA$
          If IMPO# > 0.005 Then
              mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          ElseIf IMPO# < (-0.005) Then
              mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(XVALO(IMPTEX$), "G14.2")
          Else
            If ESFACTUANTICIPO% < 1 Then
              mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 4) = "Anulado"
            End If
          End If
          mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(XVALO(SALTEX$), "G15.2")
          '
          If MONRESU% > 1 Then
            MONEM% = ROUND(!Mone_Emis, 0)
            If MONEM% <> MONRESU% Then
                  mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 4) = "..........."
                    LeyendaMoneda$ = TRIM$(DEMONECO$(MONEM%))
                    If Len(LeyendaMoneda$) < 11 Then Blancos = Int((10 - Len(LeyendaMoneda$)) / 2)
                    LeyendaMoneda$ = LeyendaMoneda$ & Space(Blancos)
                  mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 5) = LeyendaMoneda$
                  mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 6) = "............"
            End If
          End If
          '
          If MODOMONEI% = 2 Then
             MONEM% = XVALO(!Mone_Emis)
             TICAM = ROUND(XVALO(!Tipo_Cam), 5)
             If TICAM < 0.05 Then TICAM = 1
             '
             mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 7) = SIMBOLPES$(MONEM%)
             If MONEM% > 1 Then mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 8) = TRIM$(FORMATNUM$(TICAM, "F9.5"))
          End If
          
          'List3.AddItem Z$
          '
          If mfgCuentaCorrienteAnticipos.Rows > 1 Then
            If IX& > 0 Then 'VARIABLE DEL VECTOR SI ES > 0 POR QUE HAY AJUSTES.
             FechEmisAnterior% = FECHANUM(mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente& - 1, 1))
             For i& = UltimoInd& To UBound(FECHAAJUSTE%)
              If FECHAAJUSTE%(i&) >= FechEmisAnterior% And FECHAAJUSTE%(i&) < FEX Then
                UltimoInd& = UltimoInd& + 1
                Z2$ = ""
                Call REPLA(Z2$, FECHATEX$(FECHAAJUSTE%(i&)), 1, 8)
                Call REPLA(Z2$, DOCUAJUSTE$(i&), 11, 16)
                Call REPLA(Z2$, AJUSTD$(TRIM$(SAFETEXT$(NUMEAJUSTE&(i&))), 8), 27, 8)
                Call REPLA(Z2$, FORMATNUM$(IMPOHABEAJUSTE#(i&), "F13.2"), 35, 13)
                Call REPLA(Z2$, FORMATNUM$(IMPODEBEAJUSTE#(i&), "F13.2"), 50, 13)
                  IMPOAJUST# = IMPOHABEAJUSTE#(i&) - IMPODEBEAJUSTE#(i&)
                  SALDO# = SALDO# + IMPOAJUST#
                Call REPLA(Z2$, AJUSTD$(TRIM$(FORMATNUM$(SALDO#, "F14.2")), 14), 63, 14)
                'List3.AddItem Z2$
                '
                FilaCuentaCorriente& = FilaCuentaCorriente& + 1
                mfgCuentaCorrienteAnticipos.Rows = mfgCuentaCorrienteAnticipos.Rows + 1
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FECHAAJUSTE%(i&))
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 3) = DOCUAJUSTE$(i&)
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(IMPOHABEAJUSTE#(i&), "G14.2")
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(IMPODEBEAJUSTE#(i&), "G14.2")
                mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(SALDO#, "G14.2")
              Else
                Exit For
              End If
             Next i&
            End If
          End If
          
          
'          If List3.ListCount > 0 Then
'            If IX& > 0 Then 'VARIABLE DEL VECTOR SI ES > 0 POR QUE HAY AJUSTES.
'             IND& = List3.ListCount - 1
'             TTXX$ = List3.List(IND& - 1)
'             FechEmisAnterior% = FECHANUM(Mid$(TTXX$, 1, 8))
'             For i& = UltimoInd& To UBound(FECHAAJUSTE%)
'              If FECHAAJUSTE%(i&) >= FechEmisAnterior% And FECHAAJUSTE%(i&) < FEX Then
'                UltimoInd& = UltimoInd& + 1
'                Z2$ = ""
'                Call REPLA(Z2$, FECHATEX$(FECHAAJUSTE%(i&)), 1, 8)
'                Call REPLA(Z2$, DOCUAJUSTE$(i&), 11, 16)
'                Call REPLA(Z2$, AJUSTD$(TRIM$(SAFETEXT$(NUMEAJUSTE&(i&))), 8), 27, 8)
'                Call REPLA(Z2$, FORMATNUM$(IMPOHABEAJUSTE#(i&), "F13,2"), 35, 13)
'                Call REPLA(Z2$, FORMATNUM$(IMPODEBEAJUSTE#(i&), "F13,2"), 50, 13)
'                  IMPOAJUST# = IMPOHABEAJUSTE#(i&) - IMPODEBEAJUSTE#(i&)
'                  SALDO# = SALDO# + IMPOAJUST#
'                Call REPLA(Z2$, AJUSTD$(TRIM$(FORMATNUM$(SALDO#, "F14,2")), 14), 63, 14)
''                Call REPLA(Z$, AJUSTD$(TRIM$(FORMATNUM$(SALDO# + IMPO1#, "F14,2")), 14), 63, 14)
'                List3.AddItem Z2$
'              Else
'                Exit For
'              End If
'             Next i&
'            End If
'          End If
          '
         .MoveNext
        Loop
      End If
      On Error GoTo 0
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    '
    'MUESTRO AJUSTES DEL FINAL
    If IX& > 0 Then 'VARIABLE DEL VECTOR SI ES > 0 POR QUE HAY AJUSTES.
        If UltimoInd& <= UBound(FECHAAJUSTE%) Then
                 For i& = UltimoInd& To UBound(FECHAAJUSTE%)
                    Z3$ = ""
                    Call REPLA(Z3$, FECHATEX$(FECHAAJUSTE%(i&)), 1, 8)
                    Call REPLA(Z3$, DOCUAJUSTE$(i&), 11, 16)
                    Call REPLA(Z3$, AJUSTD$(TRIM$(SAFETEXT$(NUMEAJUSTE&(i&))), 8), 27, 8)
                    Call REPLA(Z3$, FORMATNUM$(IMPOHABEAJUSTE#(i&), "F13.2"), 35, 13)
                    Call REPLA(Z3$, FORMATNUM$(IMPODEBEAJUSTE#(i&), "F13.2"), 50, 13)
                      IMPOAJUST# = IMPOHABEAJUSTE#(i&) - IMPODEBEAJUSTE#(i&)
                      SALDO# = SALDO# + IMPOAJUST#
                    Call REPLA(Z3$, AJUSTD$(TRIM$(FORMATNUM$(SALDO#, "F14.2")), 14), 63, 14)
                    'List3.AddItem Z3$
                    
                    FilaCuentaCorriente& = FilaCuentaCorriente& + 1
                    mfgCuentaCorrienteAnticipos.Rows = mfgCuentaCorrienteAnticipos.Rows + 1
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 1) = FECHATEX$(FECHAAJUSTE%(i&))
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 2) = TIPOS$
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 3) = DOCUAJUSTE$(i&)
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 4) = FORMATNUM$(IMPOHABEAJUSTE#(i&), "G14.2")
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 5) = FORMATNUM$(IMPODEBEAJUSTE#(i&), "G14.2")
                    mfgCuentaCorrienteAnticipos.TextMatrix(FilaCuentaCorriente&, 6) = FORMATNUM$(SALDO#, "G14.2")
                 Next i&
        End If
    End If
    List3.ListIndex = List3.ListCount - 1
    List3.Refresh
    
    mfgCuentaCorrienteAnticipos.Rows = FilaCuentaCorriente& + 1
    Call POSICIONAR_ULTIMAFILA(mfgCuentaCorrienteAnticipos)
    
    Screen.MousePointer = 1
    '(Iva al 21.00%)
    
'991 SARRAS# = SALDO#
'    SUMADEU# = 0
'    List1.ListIndex = List1.ListCount - 1
'    List1.Refresh
'    '
'    SQLX$ = "SELECT * FROM SADEUPEN "
'    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
'    SQLX$ = SQLX$ + "AND (ABS(SaldAcre) >= 0.005 OR ABS(SaldDebe) >= 0.005 "
'    If MODOMONEI% = 2 Then
'         SQLX$ = SQLX$ + "OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005) "
'       Else
'         SQLX$ = SQLX$ + ") "       ' PARA CERRAR EL PARENTESIS
'    End If
'    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
'    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim FacPenTemp As ADODB.Recordset
'    Set FacPenTemp = New ADODB.Recordset
'    FacPenTemp.CursorLocation = adUseClient
'26  On Error Resume Next
'    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 26
'    End If
'    On Error GoTo 0
'
'    '
'    If NCX7 <> XVALO(Text1) Then GoTo 10   ' PARA ATAJAR ERROR CUANDO SE  INGRESA NUMERO DE CUENTA A MANO
'    '
'    SUDEUDOL# = 0
'    With FacPenTemp
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        On Error GoTo 0
'        .MoveFirst
'        Do While Not .EOF
'          TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
'          SG% = (-1)
'          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
'                              If TIX$ = "AJ" Then TIPOS$ = "Ajuste de Saldo"
'                              If TIX$ = "RB" Then TIPOS$ = "Recibo Cobranza"
'                              If TIX$ = "VC" Then TIPOS$ = "Venta Contado"
'                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
'                              If TIX$ = "DC" Then TIPOS$ = "Descuento Caja"
'                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito"
'                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 ''\\\OT-06-0112-WAR-30/03/06///
'                              '
'          IMPO# = (-1) * ROUND(!IBruComp) * SG%
'          SUMAFAC# = SUMAFAC# + IMPO#
'          '
'          SALDO# = ROUND(!SaldDebe - !SALDAcre)
'          SUMADEU# = SUMADEU# + SALDO#    ' (Int(100 * SALDO# + 0.5)) / 100
'          '
'          NROL& = XVALO(Mid$(!CODICOM_COR, 4, 6))
'          NROX$ = CSTRING$(MKS$(NROL&), 3, 8, 0, 2)
'          FEMI = CVINT(!FECHEMISI)
'          FEVE = CVINT(!FEVENCIM)
'          '
'          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
'          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
'          SALTEXDOL$ = ""
'          '
'          COCOLAR$ = !CODICOM_LAR
'          '
'          If MODOMONEI% = 2 Then
'             SALDOLA# = ROUND(!SaldDebe_Dola - !SaldAcre_Dola)
'             SALTEXDOL$ = FORMATNUM$(SALDOLA#, "F13.2")
'             GoTo 20
'          End If
'          '
'          ' rutina incluida para procesamiento bimonetario
'          RECABIN& = 0
''          If Abs(Abs(IMPO#) - Abs(SALDO#)) < 0.005 Then ' NO ESTA CANCELADA
''             MONEMX% = 0
''             On Error Resume Next   ' por si no está el campo
''             MONEMX% = !Mone_Emis
''             On Error GoTo 0
''             If MONEMX% < 1 Then
''                COCOLAR$ = TRIM$(!CODICOM_LAR)
''                RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7))
''                For KUL& = Len(RFG$) - 3 To 1 Step -4
''                   RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
''                   XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
''                   If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
''                      MONEMX% = CVI(Mid$(XXC$, 31, 2))
''                      GoTo 15
''                   End If
''                Next KUL&
''15              If MONEMX% < 1 Then MONEMX% = 1
''             End If
''             '
''             If MONEMX% > 1 Then
''               ' LO QUE SIGUE ES UN ENJUAGUE PARA GRAMY PARA RECUPERAR LA MONEDA DE EMISIÓN
''               If MONEMX% > ULTREG&("TAMONED") Then
''                 Call COMUNI("Moneda no Válida en Comprobante '" + COCOLAR$ + "'.\Debe Seleccionar la Moneda de Emisión.")
''                 Call SELECHO("TAMONED")
''                 If XVALO(VALACT1$("TAMONED")) > 0 Then
''                   MONEMX% = XVALO(VALACT1$("TAMONED"))
''                   If RECABIN& > 0 Then
''                     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
''                     Call REPLA(XXC$, MKI$(MONEMX%), 31, 2)
''                     Call GRAREG("CARBIMON", XXC$, RECABIN&)
''                   End If
''                 End If
''               End If
''               '
''               RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCX7))
''               For KUL& = Len(RFG$) - 3 To 1 Step -4
''                 RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
''                 XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
''                 If TRIM$(Mid$(XXC$, 45, 18)) = COCOLAR$ Then
''                   SALDOPES# = ROUND(CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8)))
''                   If Abs(SALDOPES# - SALDO#) <= 0.005 Then
''                     SALTEX$ = FORMATNUM$(CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8)), "F14.2")
''                     SBPES$ = TRIM$(SIMBOLPES$(MONEMX%))
''                     If SBPES$ <> "" Then
''                       Call REPLA(SALTEX$, SBPES$, 1, Len(SPPES$))
''                     End If
''                     GoTo 20
''                   End If
''                 End If
''               Next KUL&
''             End If
''          End If
'
'          ' fin rutina para procesamiento bimonetario
'          '
'20        Z$ = Space$(80)
'          Call REPLA(Z$, TIPOS$, 1, 16)
'          Call REPLA(Z$, NROX$, 17, 8)
'          Call REPLA(Z$, FECHATEX$(FEMI), 27, 8)
'          Call REPLA(Z$, IMPTEX$, 37, 14)
'          Call REPLA(Z$, FECHATEX$(FEVE), 53, 8)
'          Call REPLA(Z$, SALTEX$, 63, 14)
'          Call REPLA(Z$, SALTEXDOL$, 77, 13)
'          SUDEUDOL# = SUDEUDOL# + SALDOLA#
'          '
'          List2.AddItem Z$
'          '
'        .MoveNext
'        Loop
'      End If
'      On Error GoTo 0
'    End With
'    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
'    '
'    If MONRESU% <= 1 Then
'       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
'          ATENTI% = 1
'          Call REAPLICRE(NCX7)
'          If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then ATENTI% = 0: GoTo 899 'provisorio FFFERVI
'          GoTo 100
'          '
'          ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
'          ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR COBRAR APARECIA
'          ' DESACTUALIZADO.
'          '
''          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
''          If COMALTER%(TTXX$) = 2 Then
''              Call REAPLICRE(NCX7)
''              GoTo 100
''            Else
''              ATENTI% = 0
''              Call CIERRARCH("*.*")
''              Call FINAL("")
''          End If
'       End If
'    End If
'    '
'899 List2.ListIndex = List2.ListCount - 1
'    List2.Refresh
'    Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
'    Label5.Refresh
'    Label6.Caption = CSTRING$(MKD$(SUMAFAC#), 4, 14, 2, 2)
'    Label6.Refresh
'    Label4.Caption = CSTRING$(MKD$(SUDEUDOL#), 4, 14, 2, 2)
'    Label4.Refresh
'    '
'999 Screen.MousePointer = 1
'    '
'    If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then GoTo 900 'provisorio FFFERVI
'
'    If MONRESU% <= 1 Then
'       If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
'          ATENTI% = 1
'          TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta"
'          If Visible = True Then
'              TTXX$ = TTXX$ + " Elegida."
'            Else
'              TTXX$ = TTXX$ + ":\   \'" + TRIM$(RASOCLI$(NCX7)) + "'"
'          End If
'          TTXX$ = TTXX$ + "\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Créditos y Cobranzas\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
'          '
'          If COMALTER%(TTXX$) = 2 Then
'              Call REAPLICRE(NCX7)
'              GoTo 100
'            Else
'              ATENTI% = 0
'              Call CIERRARCH("*.*")
'              Call FINAL("")
'          End If
'       End If
'    End If
'    '
'900 DISPONI = XVALO(Text11.TEXT) - XVALO(Text10.TEXT) - XVALO(Text6.TEXT)
'    Text12.ForeColor = &H80000008
'    If DISPONI < 0 Then Text12.ForeColor = &HFF&
'    Text12.TEXT = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
'    '
    '
End Sub

'
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   ' RESUMEN - OD-15/12/11
   ' SI EL SISTEMA NO ES MULTIMONEDA Y SE PRESIONO BOTON DERECHO
   ' PRESENTA LOS DOCUMENTOS APLICADOS A LA FACTURA
   '
   ' SI EL SISTEMA ES MULITIMONEDA Y SE PRESIONO BOTON DERECHO SE PRESENTAN OPCIONES
   ' PARA AJUSTE DE TIPO DE CAMBIO DE DOCUMENTOS Y MOSTRAR DOCUMENTOS APLICADOS
   '
   If Button = 2 Then
       If MODOMONEI% <> 2 Then
           OPX% = 1
         Else
           OPX% = COMALTER%("Opciones de Trabajo\\Documentos Aplicados\Ajuste de Cotización")
           If OPX% < 1 Or OPX% > 2 Then Exit Sub
       End If
       '
       If OPX% = 1 Then
            TTXX1$ = List1.TEXT
            Call MUESTRADOCUAPLI(TTXX1$)
          Else
            If MONRESU% > 1 Then
               Call COMUNI("Imposible Ajustar Tipo de Cambio\en Resumen en Moneda Extranjera.")
               Exit Sub
            End If
            If DERACCE%("AJUTICAM", "Ajuste de Tipo de Cambio Documentos Emitidos") < 2 Then
               Exit Sub
            End If
            '
            Call AJUCOTIZ
       End If
   End If
   '
End Sub

Sub AJUCOTIZ()
    '
    List1.ListIndex = List1.TopIndex + Y / 190   ' OJO QUE EL 190 DEPENDE DE LA FUENTE
    TTXX$ = List1.TEXT
    COCOLA$ = TRIM$(Mid$(TTXX$, 105, 24))
    If COCOLA$ = "" Then
       Call MENSERR(24, "Imposible Ajustar Parámetros.")
       Exit Sub
    End If
    MONEM% = XVALO(Mid$(TTXX$, 103, 2))
    IMPESO# = XVALO(TRIM$(Mid$(TTXX$, 35, 28)))
    TICAM = XVALO(Mid$(TTXX$, 83, 9))
    If TICAM < 0.01 Then TICAM = 1
    VDEF$ = AJUSTI$(COCOLA$, 24) + Chr$(MONEM%) + Chr$(0) + MKD$(IMPESO#) + MKS$(TICAM)
20    OBX$ = OBJPLA$("AJUCOTFAC", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    MONEM% = Asc(Mid$(OBX$, 25, 1))
    If MONEM% < 1 Then GoTo 20
    TICAM = CVS(Mid$(OBX$, 35, 4))
    If MONEM% = 1 Then TICAM = 1
    If TICAM < 0.01 Then GoTo 20
    '
    NCX7 = XVALO(Text1)
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7 & " "
    SQLX$ = SQLX$ + "AND CODICOM_LAR = '" & COCOLA$ & "' "
    'Set CUECO = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim CUECO As ADODB.Recordset
    Set CUECO = New ADODB.Recordset
25  On Error Resume Next
    CUECO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
    End If
    On Error GoTo 0
    '
    With CUECO
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IPESO# = ROUND(!IDEBECOMP - !IHABECOMP)
          '.Edit
          !Mone_Emis = MONEM%
          !Tipo_Cam = TICAM
          !Impo_Dola = ROUND#(IPESO# / TICAM)
          .Update
        .MoveNext
        Loop
      End If
    End With
    Set CUECO = Nothing
    '
    TOPI = List1.TopIndex
    LINDI = List1.ListIndex
    Call CARESCUE
    On Error Resume Next
    List1.TopIndex = TOPI
    List1.ListIndex = LINDI
    '
End Sub

Sub REAPLICRE(NCX)
    '
    'If InStr(UCase$(EMPREAC$), "FERVI") > 0 Then Exit Sub 'provisorio FFFERVI
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#()
    RAMAX& = 0
    ATENTI% = 1
    '
100 Screen.MousePointer = 11
    '
    NUDOCU% = 0
'    RFF$ = RECFILT$("CUECOCLI", 1, MKI$(NCX%))
'    For KU& = 1 To Len(RFF$) Step 4
'      RELE& = CVS(Mid$(RFF$, KU&, 4))
'      XX$ = REGLEIDO$("CUECOCLI", RELE&)
'      NUDOCU% = NUDOCU% + 1
'      If NUDOCU% > RAMAX& Then
'        RAMAX& = RAMAX& + 128
'        ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
'      End If
'      DOCULAR$(NUDOCU%) = TRIM$(Mid$(XX$, 33, 24))
'      DOCUCOR$(NUDOCU%) = TRIM$(Mid$(XX$, 57, 12))
'      IMPO#(NUDOCU%) = ROUND(CVD(Mid$(XX$, 25, 8)))
'      IORI#(NUDOCU%) = Abs(IMPO#(NUDOCU%))
'      FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 3, 2))
'      CODIEM%(NUDOCU%) = CVI(Mid$(XX$, 69, 2))
'      CODISU%(NUDOCU%) = CVI(Mid$(XX$, 71, 2))
'      STATDOC%(NUDOCU%) = CVI(Mid$(XX$, 73, 2))
'      SALDOLA#(NUDOCU%) = 0
'      '
'      NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
'      For KKU& = 1 To List2.ListCount
'        TTXY$ = List2.List(KKU& - 1)
'        If XVALO(Mid$(TTXY$, 17, 8)) = NRODOC& Then
'          If FECHANUM(Mid$(TTXY$, 27, 8)) = FFDOC%(NUDOCU%) Then
'            If Abs(Abs(XVALO(Mid$(TTXY$, 37, 14))) - IORI#(NUDOCU%)) < 0.005 Then
'              FVENC%(NUDOCU%) = FECHANUM(Mid$(TTXY$, 53, 8))
'              GoTo 219
'            End If
'          End If
'        End If
'      Next KKU&
'      FVENC%(NUDOCU%) = 0
'219 Next KU&
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    SALDO# = 0: SARRAS# = 0
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = TRIM$(!CodiCom_Lar)   ' ANTES ESTABA SIN TRIM - 24/03/07
          DOCUCOR$(NUDOCU%) = !CODICOM_COR
          IMPO#(NUDOCU%) = ROUND(!IDEBECOMP - !IHABECOMP)
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!FECHEMISI)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          SALDOLA#(NUDOCU%) = 0
          If IsNull(!Status) = False Then STATDOC%(NUDOCU%) = !Status
          '
          NRODOC& = XVALO(Mid$(DOCULAR$(NUDOCU%), 11))
          For KKU& = 1 To List2.ListCount
            TTXY$ = List2.List(KKU& - 1)
            If XVALO(Mid$(TTXY$, 17, 8)) = NRODOC& Then
              If FECHANUM(Mid$(TTXY$, 27, 8)) = FFDOC%(NUDOCU%) Then
                If Abs(Abs(XVALO(Mid$(TTXY$, 37, 14))) - IORI#(NUDOCU%)) < 0.005 Then
                  FVENC%(NUDOCU%) = FECHANUM(Mid$(TTXY$, 53, 8))
                  GoTo 119
                End If
              End If
            End If
          Next KKU&
          FVENC%(NUDOCU%) = 0
          '
119     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    ' AQUI SE BLANQUEABA SI PASABA POR CERO. vIGENTE HASTA EL 14/02/08
    ' SE SUPRIMIÓ PARA PRIMERO PRECESAR LAS APLICACIONES EXPLÍCITAS
    ' Y RECIÉN DESPUÉS LAS IMPLÍCITAS. RECLAMADO POR GUIDI CT-100330-FL
    '
'    ' blanquea si pasa por cero
'    FIN& = NUDOCU%
'    SUMAPEN# = 0
'    For KKI% = 1 To NUDOCU%
'      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
'      If Abs(SUMAPEN#) < 0.005 Then
'        For KKJ% = 1 To KKI%
'          IMPO#(KKJ%) = 0
'        Next KKJ%
'      End If
'    Next KKI%
    '
    SQLX$ = "SELECT * FROM APLICOBRA "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCX & " "
    'Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ApliTemp As ADODB.Recordset
    Set ApliTemp = New ADODB.Recordset
    ApliTemp.CursorLocation = adUseClient
26  On Error Resume Next
    ApliTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0
    '
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          IAPLI# = !IMPOAPLI
          RECI$ = TRIM$(!codocap_lar)
          FACU$ = TRIM$(!codocor_lar)
          '
          DOCUAP% = 0: DOCUOR% = 0
          For KKI% = 1 To NUDOCU%
            If DOCULAR$(KKI%) = RECI$ And DOCUAP% < 1 Then
                DOCUAP% = KKI%
            End If
            If DOCULAR$(KKI%) = FACU$ And DOCUOR% < 1 Then
                DOCUOR% = KKI%
            End If
          Next KKI%

AA1$ = DOCULAR$(DOCUAP%)
aa2 = DOCULAR$(DOCUOR%)
          If DOCUAP% > 0 And DOCUOR% > 0 Then
            SIGAP% = 1: If IAPLI# < 0 Then SIGAP% = (-1)
            If SIGAP% * IAPLI# > SIGAP% * IMPO#(DOCUOR%) Then IAPLI# = IMPO#(DOCUOR%)
            'If SIGAP% * IAPLI# > (-1) * SIGAP% * IMPO#(DOCUAP%) Then
            '   IAPLI# = (-1) * IMPO#(DOCUAP%)   ' QUEDA POR INVESTIGAR SI
            'End If                              ' ESTA CONDICION CORRECTA - EPB 11/10/08
            IMPO#(DOCUOR%) = IMPO#(DOCUOR%) - IAPLI#
            IMPO#(DOCUAP%) = IMPO#(DOCUAP%) + IAPLI#
          End If
          '
'          IRECIMA# = 0: IFACMAX# = 0
'          For KKI% = 1 To NUDOCU%
'             If DOCULAR$(KKI%) = RECI$ Then
'               If IMPO#(KKI%) < 0 Then
'                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
'               End If
'             End If
'             If Mid$(DOCULAR$(KKI%), 9, 8) = Mid$(FACU$, 9, 8) Then
'               If IMPO#(KKI%) > 0 Then
'                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
'               End If
'             End If
'          Next KKI%
'          '
'          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
'          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
'          '
' If IAPLI# < 1 Then
' AAA = BBB
' End If
'          IRECIMA# = IAPLI#
'          IFACMAX# = IAPLI#
'          '
'          If IAPLI# > 0 Then
'             For KKI% = 1 To NUDOCU%
'                If DOCULAR$(KKI%) = RECI$ Then
'                   If IMPO#(KKI%) < 0 Then
'                      IAP1# = IRECIMA#
'                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
'                      IRECIMA# = IRECIMA# - IAP1#
'                   End If
'                End If
'                If Mid$(DOCULAR$(KKI%), 9, 8) = Mid$(FACU$, 9, 8) Then
'                  If IMPO#(KKI%) > 0 Then
'                      IAP2# = IFACMAX#
'                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
'                      IFACMAX# = IFACMAX# - IAP2#
'                   End If
'                End If
'             Next KKI%
'          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    If InStr(1, EMPREAC$, "AHP") > 0 Or InStr(1, EMPREAC$, "CEYEQ") > 0 Then GoTo 24
    If InStr(1, EMPREAC$, "FERVI") > 0 Or InStr(1, EMPREAC$, "AMBEST") > 0 Then GoTo 241
    '
    SUMAPEN# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      If Abs(SUMAPEN#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    GoTo 242
    
241 SUMAPEN# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      If FFDOC%(KKI%) > FECHANUM("01/10/15") Then Exit For
      If Abs(SUMAPEN#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    
    '
242 For KKI% = 1 To NUDOCU%
      For KKJ% = KKI% + 1 To NUDOCU%
        If (Abs(IMPO#(KKI%) + IMPO#(KKJ%))) < 0.005 And FFDOC%(KKI%) < FECHANUM("01/10/15") And FFDOC%(KKJ%) < FECHANUM("01/10/15") Then
          IMPO#(KKI%) = 0
          IMPO#(KKJ%) = 0
        End If
      Next KKJ%
    Next KKI%
          
    KKM% = 0
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
      If STATDOC%(KKI%) = 8 Then
        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      End If
    Next KKI%
    '
    If Abs(SUMAPEN#) < 0.005 Then
      For KKI% = 1 To NUDOCU%
        If STATDOC%(KKI%) = 8 Then
          IMPO#(KKI%) = 0
        End If
      Next KKI%
      GoTo 19
    End If
    '
    For KKI% = NUDOCU% To 1 Step -1
      If STATDOC%(KKI%) = 8 Then
        If Sgn(IMPO#(KKI%)) <> Sgn(SUMAPEN#) Then
            'SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
            IMPO#(KKI%) = 0
          Else
            If Abs(SUMAPEN#) >= Abs(IMPO#(KKI%)) Then
                SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
                'IMPO#(KKI%) = 0
              Else
                IMPO#(KKI%) = SUMAPEN#
                SUMAPEN# = 0
            End If
        End If
      End If
    Next KKI%
    '
19  If MODOMONEI% = 2 Then       ' cargar saldo en dolares de carbimon.dat
       RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCX))
       FIN& = Len(RFF$)
       For KKU& = 1 To FIN& Step 4
          Call TRACE(20, KKU&, FIN&)
          U& = CVS(Mid$(RFF$, KKU&, 4))
          XX$ = REGLEIDO$("CARBIMON", U&)
          NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
          DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
          HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
          '
          For JJ% = 1 To NUDOCU%
            If Left$(DOCULAR$(JJ%), 16) = Left$(NUDOCX$, 16) Then
               SALDOLA#(JJ%) = SALDOLA#(JJ%) + DEBEDOL# - HABEDOL#
               GoTo 191
            End If
          Next JJ%
          '
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENC%(RAMAX&), STATDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = NUDOCX$
          DOCUCOR$(NUDOCU%) = Left$(NUDOCX$, 3) + Right$(NUDOCX$, 8)
          IMPO#(NUDOCU%) = 0
          IORI#(NUDOCU%) = CVD(Mid$(XX$, 65, 8))
          FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 33, 2))
          CODIEM%(NUDOCU%) = 0
          CODISU%(NUDOCU%) = 0
          STATDOC%(NUDOCU%) = 1
          SALDOLA#(NUDOCU%) = DEBEDOL# - HABEDOL#
          '
191    Next KKU&
    End If
    '
    For JJ% = 1 To NUDOCU%
      For JJ1% = JJ% + 1 To NUDOCU%
        If Abs(IMPO#(JJ%) + IMPO#(JJ1%)) < 0.005 And FFDOC%(JJ%) < FECHANUM("01/10/15") And FFDOC%(JJ1%) < FECHANUM("01/10/15") Then
          If Abs(SALDOLA#(JJ%) + SALDOLA#(JJ1%)) < 0.005 Then
            IMPO#(JJ%) = 0
            SALDOLA#(JJ%) = 0
            IMPO#(JJ1%) = 0
            SALDOLA#(JJ1%) = 0
          End If
        End If
      Next JJ1%
    Next JJ%
    '
    If CONTROLGRABA% = 1 Then GoTo 24
    Screen.MousePointer = 1
    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(rasocli$(NCX))
    'TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
    '
    ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
    ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR COBRAR APARECIA
    ' DESACTUALIZADO.
    '
'    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
'       GoTo 999
'    End If
    '
24  Screen.MousePointer = 11
    SQLX$ = "DELETE FROM SaDeuPen WHERE NuClien = " & NCX & " "
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    ' AQUI ABRIR UN RECORDSET PARA GRABAR EN SADEUPEN
    Dim SaDeuPen As ADODB.Recordset
    Set SaDeuPen = New ADODB.Recordset
    SaDeuPen.CursorType = adOpenDynamic
    SaDeuPen.LockType = adLockPessimistic
    SQLX$ = "Select *  from SaDeuPen WHERE NuClien = " & NCX & " "
27  On Error Resume Next
    SaDeuPen.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
    End If
    On Error GoTo 0
    '
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    LU$ = LUDAT$
    Call APERBASDAT("CABAN")
    NOCLOSE% = 1
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Or Abs(SALDOLA#(U&)) >= 0.005 Then
        IMORIGDOL# = 0: MONEMIOR% = 0
        If MODOMONEI% = 2 Then
          ' BUSCAR MONEDA E IMPORTE ORIGINAL DEL DOCUMENTO
          Call APERBASDAT("CABAN")
          SQLX$ = " SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & DOCULAR$(U&) & "' "
          Dim CAJJJABANC As ADODB.Recordset
          Set CAJJJABANC = New ADODB.Recordset
28        On Error Resume Next
          CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 28
          End If
          On Error GoTo 0
          With CAJJJABANC
            'CajaBanc.FindFirst "CodiCom_Lar = '" & DOCULAR$(U&) & "' "
            'If CajaBanc.NoMatch = False Then
            On Error Resume Next
            If Not (.EOF And .BOF) Then
                On Error GoTo 0
                IMORIGDOL# = !Impo_Dola
                MONEMIOR% = !Mone_Emis Mod 6
                TICAMO = !Tipo_Cam
                VALACOM1# = !ValAbsComp
                On Error GoTo 0
                If MONEMIOR% > 1 Then
                  If Abs(IMORIGDOL#) < 0.005 Then
                    If TICAMO > 0.1 Then
                      IMORIGDOL# = ROUND(VALACOM1# / TICAMO)
                      '.Edit
                      !Impo_Dola = IMORIGDOL#
                      .Update
                    End If
                  End If
                End If
            End If
            On Error GoTo 0
          End With
          CAJJJABANC.Close
          Set CAJJJABANC = Nothing
        End If
        '
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
'        CodiEmp% = CODIEM%(U&)
'        CodiSuc% = CODISU%(U&)
        NumeClie = NCX
        NumeProv% = 0
        FechEmi = CVDAT(FFDOC%(U&))
        IBruTot# = IORI#(U&)
        IDolTot# = IMORIGDOL#
        '
        If DOCULAR$(U&) <> DOCULARANT$ Then NuorVen% = 0
        NuorVen% = NuorVen% + 1
        '
        FVECIX% = FVENC%(U&)
        If FVECIX% < 1 Then
           COPAGI% = CPAGCLI%(NCX)
           FEDOCU% = FFDOC%(U&)
           Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
           aaaa = FEDOCU%
           BBBB = FEDOCU%
           PLAPA% = 0: If Len(PLA$) > 0 Then PLAPA% = Asc(Left$(PLA$, 1))
           While DIENTRE%(aaaa, BBBB) < PLAPA%
             BBBB = DIPOST(BBBB)
           Wend
           FVECIX% = BBBB
        End If
        '
        FeVenc = CVDAT(FVECIX%)
        ImVenc# = IMPO#(U&)
        ImVDol# = SALDOLA#(U&)
        DOCULARANT$ = DOCULAR$(U&)
        '
        GoSub GRADEURECU
      End If
    Next U&
    SaDeuPen.Close
    NOCLOSE% = 0
    '
999 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
    Exit Sub
    '
    '
GRADEURECU:
   SQLX$ = "SELECT * FROM SADEUPEN"
   Dim SADDDEUPEN As ADODB.Recordset
   Set SADDDEUPEN = New ADODB.Recordset
29 On Error Resume Next
   SADDDEUPEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 29
   End If
   On Error GoTo 0

   With SADDDEUPEN
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !nuclien = NumeClie
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie - NumeProv%), 16)
      !RaSoCliPro = Left$(rasocli$(NumeClie - NumeProv%), 48)
      !CODICOM_COR = Left$(CoComCor$, 12)
      !CodiCom_Lar = Left$(CoComLar$, 24)
      !FECHEMISI = FechEmi
        On Error Resume Next
      !Mone_Emis = MONEMIOR%
        On Error GoTo 0
        SIGX% = 1
        If Abs(ImVenc#) > 0 Then SIGX% = Sgn(ImVenc#)
      !IBruComp = IBruTot# * SIGX%
      !IIvaComp = IIvaTot#
      !Ibru_Dola = IDolTot# * SIGX%
      !OrVencim = NuorVen%
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
          !SaldDebe = ImVenc#
          !SALDAcre = 0
        Else
          !SALDAcre = Abs(ImVenc#)
          !SaldDebe = 0
      End If
      If ImVDol > 0 Then
          !SaldDebe_DOLA = ImVDol#
          !SaldAcre_DOLA = 0
        Else
          !SaldAcre_DOLA = Abs(ImVDol#)
          !SaldDebe_DOLA = 0
      End If
      !Status = 0
     .Update
   End With
   SADDDEUPEN.Close
   Set SADDDEUPEN = Nothing
Return
'
End Sub

Private Sub Text7_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text8_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Private Sub Text9_GotFocus()
'\\\OT-06-0554-RG-13/12/06///
    Text1.SetFocus
'\\\OT-06-0554-RG-FIN///
End Sub

Sub MUESTRADOCUAPLI(RELIS$)
    '
    'MUESTRA LOS DOCUMENTOS APLICADOS A LA FACTURA  O NOTA DE DEBITO SELECCIONADA
    Screen.MousePointer = 11
    'LEO LOS IMPORTES Y FECHAS
    IMPODEBE# = XVALO(Mid$(RELIS$, 35, 13))
    IMPOHABE# = XVALO(Mid$(RELIS$, 50, 13))
    fech% = FECHANUM(TRIM$(Mid$(RELIS$, 1, 8)))
    FEX = FETEXCOR1$(fech%)
    NCX = XVALO(Text1)
    TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))

    'CASO FACTURAS, N.DEBITO. AJUSTES EN POSITIVO
    If IMPODEBE# > 0.005 Then
      '
      NUDOCUM$ = TRIM$(Mid$(RELIS$, 257, 13))
      TIMOV$ = "FVEN"
      If TIDOCUM$ = "Ajuste de Saldo" Then TIMOV$ = "AJUD"
      CONDI$ = "TIPOMOVIM =  '" & TIMOV$ & "'  AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NCX & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
      NUDOCU$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
      '
      CONDI$ = "codocor_lar = '" & NUDOCU$ & "' AND NUCLIEN = " & NCX & ""
      '
      Call CARGALISEL("DOCUMAPL", "APLICOBRA", "codocap_lar/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      '
    'CASO DE NOTAS DE CREDITO , RECIBOS , AJUSTES NEGATIVOS
    ElseIf IMPOHABE# > 0.005 Then
      NUDOCUM$ = TRIM$(Mid$(RELIS$, 27, 8))
      
      TIMOV$ = "FVEN"
      If TIDOCUM$ = "Recibo Cobranza" Then TIMOV$ = "RCOB"
      If TIDOCUM$ = "Ajuste de Saldo" Then TIMOV$ = "AJUD"
      CONDI$ = "TIPOMOVIM =  '" & TIMOV$ & "'  AND fechemisi = '" & FEX & "' AND NUCLIEN = " & NCX & "  AND  codicom_lar like '%" & NUDOCUM$ & "%'"
      NUDOCU$ = VALOBADA("CAJABANC", "CODICOM_LAR", CONDI$, "NOMESS")
      '
      If TIDOCUM$ = "Nota de Credito" Then
         CONDI$ = "CODOCOR_LAR = '" & NUDOCU$ & "' AND NUCLIEN = " & NCX & ""
         Call CARGALISEL("DOCUMAPL", "APLICOBRA", " CODOCAP_LAR/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      Else
         CONDI$ = "CODOCAP_LAR = '" & NUDOCU$ & "' AND NUCLIEN = " & NCX & ""
         Call CARGALISEL("DOCUMAPL", "APLICOBRA", " CODOCOR_LAR/A24;IMPOAPLI/F14.2", CONDI$, "ORDER BY codocor_lar")
      End If
    End If
    
90  Screen.MousePointer = 1
    ENCAB$ = "Comprobante: " & NUDOCU$
    Call SELECHO("DOCUMAPL/" & ENCAB$)
    '
End Sub




