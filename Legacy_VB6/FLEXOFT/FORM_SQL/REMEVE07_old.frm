VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form REMEVE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Ventas"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   58
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   9480
      ScaleHeight     =   6705
      ScaleWidth      =   1170
      TabIndex        =   42
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   61
         Top             =   5760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   62
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3600
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Archivos de Transferencia Percepciones I.V.A. / Ingresos Brutos"
         Top             =   2325
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   53
         Top             =   4320
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cust."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Acceso a Maestro de Clientes"
         Top             =   1685
         Visible         =   0   'False
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
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   4725
         Visible         =   0   'False
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
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   955
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
         Height          =   640
         Left            =   105
         Picture         =   "REMEVE07.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2965
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E2D3C0&
      Caption         =   "DISPOSITIVO DE SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   5400
      Width           =   4635
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Pantalla"
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
         Left            =   840
         TabIndex        =   32
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Impresión"
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
         TabIndex        =   31
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PERÍODO MENSUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   4635
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3600
         ScaleHeight     =   510
         ScaleWidth      =   975
         TabIndex        =   60
         Top             =   180
         Width           =   1000
         Begin VB.Image Image1 
            Height          =   480
            Left            =   0
            Picture         =   "REMEVE07.frx":15F6
            Top             =   20
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   460
            Picture         =   "REMEVE07.frx":1900
            Top             =   20
            Width           =   480
         End
      End
      Begin VB.CommandButton Command12 
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
         Left            =   3360
         TabIndex        =   28
         Top             =   410
         Width           =   175
      End
      Begin VB.TextBox MESAN 
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
         Left            =   1575
         TabIndex        =   27
         Top             =   410
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año"
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
         Left            =   -735
         TabIndex        =   29
         Top             =   480
         Width           =   2220
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "GESTIÓN DE VENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   25
      Top             =   1100
      Width           =   4635
      Begin VB.TextBox Text2 
         Height          =   855
         Left            =   3240
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   3240
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1365
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   33
         Top             =   1470
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "REMEVE07.frx":1C0A
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Facturación en Cuenta Corriente y de Contado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   735
            TabIndex        =   35
            Top             =   0
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   36
         Top             =   945
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "REMEVE07.frx":1F14
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Débito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   38
            Top             =   150
            Width           =   1410
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   735
         ScaleHeight     =   615
         ScaleWidth      =   2295
         TabIndex        =   39
         Top             =   420
         Width           =   2325
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
            Height          =   630
            Left            =   0
            Picture         =   "REMEVE07.frx":221E
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas de Crédito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   41
            Top             =   150
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "LIBROS DE VENTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   4635
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2970
         TabIndex        =   19
         Top             =   525
         Width           =   3000
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   0
            Picture         =   "REMEVE07.frx":2528
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Emite Subdiario de Ventas por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command21 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   1840
            Picture         =   "REMEVE07.frx":2832
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Emite Subdiario de I.V.A.- Ventas por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command18 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   0
            Picture         =   "REMEVE07.frx":2B3C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Emite Subdiario de Compras por Impresora"
            Top             =   640
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   2420
            Picture         =   "REMEVE07.frx":2E46
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Resumen Mensual I.V.A. Ventas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión de Subdiarios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   24
            Top             =   105
            Width           =   1005
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OTROS REPORTES MENSUALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ESTADÍSTICAS DE VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   48
         Top             =   4620
         Width           =   4005
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   230
            ScaleHeight     =   615
            ScaleWidth      =   3510
            TabIndex        =   49
            Top             =   420
            Width           =   3540
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   2970
               Picture         =   "REMEVE07.frx":3150
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "Detalle de Facturación por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command30 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   1835
               Picture         =   "REMEVE07.frx":345A
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Ranking de Ventas por Cliente por Rango de Fechas"
               Top             =   0
               Width           =   555
            End
            Begin VB.CommandButton Command11 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   2390
               Picture         =   "REMEVE07.frx":389C
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Ranking de Ventas por Articulo por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command10 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "REMEVE07.frx":3BA6
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Resumen por Provincia y Grupo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísticos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   675
               TabIndex        =   51
               Top             =   90
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2D3C0&
         Caption         =   "COBRANZA Y AJUSTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2505
            TabIndex        =   4
            Top             =   840
            Width           =   2535
            Begin VB.CommandButton Command2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "REMEVE07.frx":3EB0
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Cobranza en Cuenta Corriente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   6
               Top             =   105
               Width           =   1850
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2505
            TabIndex        =   7
            Top             =   420
            Width           =   2535
            Begin VB.CommandButton Command7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "REMEVE07.frx":41BA
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   0
               Left            =   735
               TabIndex        =   9
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "PROCESOS CONTABLES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   2
         Top             =   3320
         Width           =   4005
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   3165
            TabIndex        =   13
            Top             =   420
            Width           =   3200
            Begin VB.CommandButton Command20 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   2040
               Picture         =   "REMEVE07.frx":44C4
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command19 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "REMEVE07.frx":47CE
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   2600
               Picture         =   "REMEVE07.frx":4AD8
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Re-Imputación de Comprobantes de Ventas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Ventas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   17
               Top             =   105
               Width           =   1365
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "RESUMEN OPERATIVO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   210
         TabIndex        =   1
         Top             =   2020
         Width           =   4005
         Begin VB.PictureBox V 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2085
            TabIndex        =   10
            Top             =   420
            Width           =   2115
            Begin VB.CommandButton Command6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "REMEVE07.frx":4C22
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Mensual por Cuenta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   12
               Top             =   105
               Width           =   1485
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "REMEVE07.frx":5064
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "REMEVE07.frx":5298
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuCerrarPeriodoActivo 
         Caption         =   "Cerrar Período Activo"
      End
      Begin VB.Menu mnuAbrirPeriodoActivo 
         Caption         =   "Abrir Período Activo"
      End
   End
   Begin VB.Menu mnuConsultasyListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuCuentaCorriente 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu separa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEstadoCuenta 
         Caption         =   "Estado de Cuenta"
         Begin VB.Menu mnuLimitesDeCredito 
            Caption         =   "Listado de Límites de Crédito y Deudas"
         End
         Begin VB.Menu mnuCuentasPorCobrar 
            Caption         =   "Cuentas por Cobrar"
         End
         Begin VB.Menu mnuSaldosyDeudas 
            Caption         =   "Saldos y Deudas de Clientes"
         End
         Begin VB.Menu mnuAvanceFacturacion 
            Caption         =   "Avance de Facturación"
         End
         Begin VB.Menu mnuFlujoDeCajaProyectado 
            Caption         =   "Flujo de Caja Proyectado"
         End
      End
      Begin VB.Menu mnuGestionVentas 
         Caption         =   "Gestión de Ventas"
         Begin VB.Menu mnuFacturacion 
            Caption         =   "Facturación en Cuenta Corriente"
         End
         Begin VB.Menu mnuNotasDebito 
            Caption         =   "Notas de Débito"
         End
         Begin VB.Menu mnuNotasCredito 
            Caption         =   "Notas de Crédito"
         End
         Begin VB.Menu mnuDetalleCobranza 
            Caption         =   "Detalle de Cobranza"
         End
         Begin VB.Menu mnuResumenMensualCuenta 
            Caption         =   "Resumen Mensual por Cuenta"
         End
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFactAist 
         Caption         =   "Facturación de Asistencias Ténicas"
      End
      Begin VB.Menu mnuLibrosDeVentas 
         Caption         =   "Libro de Ventas"
         Begin VB.Menu mnuSubdiarioVentas 
            Caption         =   "Subdiario de Ventas"
         End
         Begin VB.Menu mnuLibroIvaVentas 
            Caption         =   "Libro Iva Ventas"
         End
         Begin VB.Menu mnuResumenIvaVentas 
            Caption         =   "Resumen Iva Ventas"
         End
      End
      Begin VB.Menu mnuEstadisticas 
         Caption         =   "Estadísticas"
         Begin VB.Menu mnuResumenProvinciaGrupo 
            Caption         =   "Resumen por Provincia y Grupo"
         End
         Begin VB.Menu mnuRankingVentasCliente 
            Caption         =   "Ranking de Ventas por Cliente"
         End
         Begin VB.Menu mnuRankingVentasArticulo 
            Caption         =   "Ranking de Ventas por Artículio"
         End
         Begin VB.Menu mnuDetalleFacturacion 
            Caption         =   "Detalle de Facturación"
         End
         Begin VB.Menu linea20 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEstadisticasConfiguraciones 
            Caption         =   "Configuraciones"
            Begin VB.Menu mnuEstadisticasTablaProductos 
               Caption         =   "Tabla de Productos"
            End
            Begin VB.Menu mnuEstadisticasTablaVendedores 
               Caption         =   "Tabla de Vendedores"
            End
            Begin VB.Menu mnuCuentasEgreso 
               Caption         =   "Tabla de Cuentas de Egresos"
            End
            Begin VB.Menu mnuGrupoMediosCobranza 
               Caption         =   "Tabla de Grupos de Medios de Cobranza"
            End
         End
         Begin VB.Menu mnuFacturaNetavsCmv 
            Caption         =   "Facturación Neta vs CMV"
            Begin VB.Menu mnuFacturaNetavsCmvMes 
               Caption         =   "Mes por Mes"
            End
            Begin VB.Menu mnuFacturaNetavsCmvGrupo 
               Caption         =   "Por Grupo"
            End
         End
         Begin VB.Menu mnuComparativaPorAnio 
            Caption         =   "Comparativa por Año"
            Begin VB.Menu mnuComparativaFacturacionNeta 
               Caption         =   "Facturación Neta y Proyección"
            End
            Begin VB.Menu mnuComparativaExportacion 
               Caption         =   "Exportaciones"
            End
            Begin VB.Menu lineacomparativa 
               Caption         =   "-"
            End
            Begin VB.Menu mnuComparativaEquipos 
               Caption         =   "Equipos"
               Begin VB.Menu mnuComparativaEquiposCantidad 
                  Caption         =   "Cantidad y Proyeccion"
               End
            End
         End
         Begin VB.Menu mnuEquiposPorMes 
            Caption         =   "Equipos por Mes"
         End
         Begin VB.Menu mnuVentasPorProvincia 
            Caption         =   "Ventas por Provincia"
            Begin VB.Menu mnuVentasPorProvinciaCantidades 
               Caption         =   "Cantidades"
            End
            Begin VB.Menu mnuVentasPorProvinciaFacturacion 
               Caption         =   "Facturación Neta"
            End
         End
         Begin VB.Menu mnuVentasPorCanales 
            Caption         =   "Ventas por Canales"
            Begin VB.Menu mnuVentasTodosLosCanales 
               Caption         =   "Todos los Canales"
            End
            Begin VB.Menu mnuVentasSoloVendedores 
               Caption         =   "Solo Vendedores"
            End
            Begin VB.Menu mnuVentasVendodoresProductos 
               Caption         =   "Vendedores y Productos"
            End
            Begin VB.Menu mnuVentasProductosVendodores 
               Caption         =   "Productos y Vendedores"
            End
         End
         Begin VB.Menu mnuEgresos 
            Caption         =   "Resumen de Egresos"
         End
         Begin VB.Menu mnuCashFlow 
            Caption         =   "Cash Flow"
         End
         Begin VB.Menu mnuAnalisisFacturacionPedidos 
            Caption         =   "Análisis de Facturación de Pedidos"
         End
      End
      Begin VB.Menu mnuProcesosContables 
         Caption         =   "Procesos Contables"
         Begin VB.Menu mnuContabilizacionVentasPorCuenta 
            Caption         =   "Contabilizacion de Ventas por Cuenta"
         End
         Begin VB.Menu mnuAsientoVentas 
            Caption         =   "Asiento Mensual de Ventas"
         End
      End
      Begin VB.Menu mnuAnalizarCorrelatividadComprobantes 
         Caption         =   "Analizar Correlatividad de Comprobantes"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuAjustes 
         Caption         =   "Ajustes de Saldo"
      End
      Begin VB.Menu mnuReimputacion 
         Caption         =   "Re-imputación de Comprobantes"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu mnuImpresionLibroIva 
         Caption         =   "Impresión Libro de Iva"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaestroClientes 
         Caption         =   "Maestros de Clientes"
      End
      Begin VB.Menu mnuLiquidaComis 
         Caption         =   "Liquidación de Comisiones"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Impuestos 
      Caption         =   "Impuestos"
      Begin VB.Menu IIBB 
         Caption         =   "Ingresos Brutos"
         Begin VB.Menu IIBBpercep 
            Caption         =   "Percepciones (Propias)"
            Begin VB.Menu IIBBpercepBsAS 
               Caption         =   "Buenos Aires"
               Begin VB.Menu IIBBpercepExpo 
                  Caption         =   "Archivo de Transferencia"
               End
            End
         End
         Begin VB.Menu IIBBpercepCABA 
            Caption         =   "CABA"
            Begin VB.Menu IIBBPercepRetenCABAExpo 
               Caption         =   "Percepciones y Retenciones - Archivo de Transferencia"
            End
            Begin VB.Menu IIBBpercepCABANCExpo 
               Caption         =   "Notas de Crédito - Archivo de Transferencia"
            End
         End
         Begin VB.Menu IIBBreten 
            Caption         =   "Retenciones"
            Begin VB.Menu IIBBretenLis 
               Caption         =   "Listado de Retenciones"
            End
            Begin VB.Menu IIBBretenExpo 
               Caption         =   "Archivo de Transferencia"
            End
         End
      End
      Begin VB.Menu Iva 
         Caption         =   "I.V.A."
         Begin VB.Menu IvaPercep 
            Caption         =   "Percepciones (Propias)"
            Begin VB.Menu IvaPercepExpo 
               Caption         =   "Archivo de Transferencia"
            End
         End
         Begin VB.Menu IvaReten 
            Caption         =   "Retenciones"
            Begin VB.Menu IvaRetenLis 
               Caption         =   "Listado de Retenciones"
            End
            Begin VB.Menu IvaRetenExpo 
               Caption         =   "Archivo de Transferencia"
            End
         End
      End
      Begin VB.Menu RegInfoVentas 
         Caption         =   "Régimen Informativo de Ventas"
      End
      Begin VB.Menu mnuLibroIvaDigital 
         Caption         =   "Libro Iva Digital"
      End
   End
End
Attribute VB_Name = "REMEVE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDE%, HASTA%, CORRESP$
Dim DESDAT As String
Dim HASDAT As String
Function ACTUVENDEDOR%(DES1, HAS1)

    ACTVDDOR% = 1 ' 1 si tiene facturas con vendedor sin asignar caso contrario 0
    '
    Call COPYSTRU("SELCUERE", "SELFACNC") ' LISTADO DE SELECCION
    '
20  CONDI$ = "TIPOMOVIM = 'FVEN' "
    CONDI$ = CONDI$ + " and FECHEMISI >= '" & DES1 & "'"
    CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HAS1 & "'"
    CONDI$ = CONDI$ + " AND (VEND_CLI < 1 OR VEND_CLI IS NULL )"
    CONDI$ = CONDI$ + " ORDER BY FechEmisi "
    '
    Call CARGALISEL("SELFACNC", "CAJABANC", "NUMECOMP/I10;CODICOM_LAR/A20;FechEmisi/D8;Nuclien/I8", CONDI$, "DISTINCT")
    Call TRALOCAL("SELFACNC", "")
    '
    If ULTREG&("SELFACNC") < 1 Then ACTVDDOR% = 0: GoTo 99
    
25  Call SELECHO("SELFACNC")
    If VALACT2$("SELFACNC") = "" Then GoTo 99
    FilaSeleccionada$ = VALACT2$("SELFACNC")
    NFACT$ = Left$(FilaSeleccionada$, 18)
    fech% = FECHANUM(Mid$(FilaSeleccionada$, 22, 8))
    '
    'CONDI1$ = " CODICOM_LAR = '" & NFACT$ & "'"
    'FECH = VALOBADA("CAJABANC", "FechEmisi", CONDI1$, "NOMESS")
    'NC% = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI1$, "NOMESS"))
    NCLIE = XVALO(Mid$(FilaSeleccionada$, 31))
    RazonSoc$ = rasocli$(NCLIE)
    '
    'Call REPLA(VDEF$, MKI$(CVINT(FECH)), 1, 2) ' INFORMACION PARA EL OBJETO
    Call REPLA(VDEF$, MKI$(fech), 1, 2) ' INFORMACION PARA EL OBJETO
    Call REPLA(VDEF$, NFACT$, 3, 18)
    Call REPLA(VDEF$, MKS$(NCLIE), 21, 4)
    Call REPLA(VDEF$, RazonSoc$, 25, 30)
    '
30  OBX$ = OBJPLA$("ACTUVENDE", VDEF$) ' ABRE EL OBJETO
    '
    If OBX$ = "" Then GoTo 25
    '
    On Error Resume Next ' POR SI INGRESA UN VENDEDOR > 255
    VDDOR% = CVI(Mid$(OBX$, 55, 2)) ' VALIDA EL NUMERO DE VENDEDOR
    If VDDOR% < 1 Or ECOARCH$("VENDEDOR", Chr$(VDDOR%)) = "" Then
      On Error GoTo 0
      Call COMUNI("Vendedor No Válido")
      VDEF$ = OBX$
      GoTo 30
    End If
    On Error GoTo 0
    '
    'CONFIRMACION Y GRABACION
    OPC% = COMALTER%("Confirma El Ingreso de Vendedor:\\No\Si")
    '
    If OPC% = 2 Then ' GRABACION
        'FECHX = FETEXCOR1(CVINT(FECH))
        FECHX = FETEXCOR1(fech%)
        CONDI2$ = " CODICOM_LAR = '" & NFACT$ & "' AND FechEmisi = '" & FECHX & "' AND NUCLIEN = " & NCLIE & ""
        Call ACTUREGISTRO("CAJABANC", "VEND_CLI", CONDI2$, VDDOR%, REGAF&, "NOMESS")
        GoTo 20 ' VUELVE AL SELECHO ACTUALIZADO
    End If
    GoTo 25 ' MUESTRA EL SELECHO SIN ACTUALIZAR POR QUE NO SE MODIFICÓ

99  ACTUVENDEDOR% = ACTVDDOR%

End Function

Sub AnalizaCorrelativaDeComprobantes(MODO As Integer, Optional MostrarResultado As Boolean = True)

    Screen.MousePointer = 11
    
    PuntoVentaFacturaComun% = XVALO(CONTROL("FACTURA", "PUVENELE"))
    PuntoVentaFacturaCredito% = XVALO(CONTROL("FACTURA", "PUVECRED"))
    PuntoVentaBienesDeCapital% = XVALO(CONTROL("FACTURA", "PUVENCAP"))
    PuntoVentaExportacion% = XVALO(CONTROL("FACTURA", "PUVEFAEX"))
    
    ComunA$ = "NO APLICA"
    ComunB$ = "NO APLICA"
    NDComunA$ = "NO APLICA"
    NDComunB$ = "NO APLICA"
    NCComunA$ = "NO APLICA"
    NCComunB$ = "NO APLICA"
    Credito$ = "NO APLICA"
    NDCredito$ = "NO APLICA"
    NCCredito$ = "NO APLICA"
    BienesA$ = "NO APLICA"
    BienesB$ = "NO APLICA"
    NDBienesA$ = "NO APLICA"
    NDBienesB$ = "NO APLICA"
    NCBienesA$ = "NO APLICA"
    NCBienesB$ = "NO APLICA"
    Expo$ = "NO APLICA"
    NDExpo$ = "NO APLICA"
    NCExpo$ = "NO APLICA"

    FechaFin% = DIANTE(HOY(HO$), 45)
    Fechaini% = DIANTE(FechaFin%, 30)
    
    ' ACTUALIZA STORE PROCEDURE DE BUSQUEDA DE HUECOS
    Call BusquedaHuecos_StoreProcedure
    
    Call TRACE(24, 1, 100)
    If PuntoVentaFacturaComun% > 0 Then
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then ComunA$ = BusquedaHuecos("fc", "a", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 5, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then ComunB$ = BusquedaHuecos("fc", "b", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 10, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDComunA$ = BusquedaHuecos("nd", "a", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 15, 100)
    
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDComunB$ = BusquedaHuecos("nd", "b", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 20, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "a", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCComunA$ = BusquedaHuecos("nc", "a", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 25, 100)
    
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "b", PuntoVentaFacturaComun%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCComunB$ = BusquedaHuecos("nc", "b", PuntoVentaFacturaComun%, LIMITEINFERIOR&, LIMITESUPERIOR)
    End If

    Call TRACE(24, 30, 100)
    If PuntoVentaFacturaCredito% > 0 Then
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then Credito$ = BusquedaHuecos("fc", "a", PuntoVentaFacturaCredito%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 35, 100)

        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDCredito$ = BusquedaHuecos("nd", "a", PuntoVentaFacturaCredito%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 40, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "a", PuntoVentaFacturaCredito%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCCredito$ = BusquedaHuecos("nc", "a", PuntoVentaFacturaCredito%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 50, 100)
    End If

    Call TRACE(24, 60, 100)
    If PuntoVentaBienesDeCapital% > 0 Then
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then BienesA$ = BusquedaHuecos("fc", "a", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 62, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then BienesB$ = BusquedaHuecos("fc", "b", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 68, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDBienesA$ = BusquedaHuecos("nd", "a", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 70, 100)
    
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDBienesB$ = BusquedaHuecos("nd", "b", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 72, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "a", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCBienesA$ = BusquedaHuecos("nc", "a", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 75, 100)
    
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "b", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCBienesB$ = BusquedaHuecos("nc", "b", PuntoVentaBienesDeCapital%, LIMITEINFERIOR&, LIMITESUPERIOR)
    End If

    Call TRACE(24, 80, 100)
    If PuntoVentaExportacion% > 0 Then
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("fc", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("fc", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then Expo$ = BusquedaHuecos("fc", "e", PuntoVentaExportacion%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 85, 100)

        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nd", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nd", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NDExpo$ = BusquedaHuecos("nd", "e", PuntoVentaExportacion%, LIMITEINFERIOR&, LIMITESUPERIOR): Call TRACE(24, 90, 100)
        
        LIMITEINFERIOR& = BusquedaHuecos_LimiteInferior("nc", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        LIMITESUPERIOR& = BusquedaHuecos_LimiteSuperior("nc", "e", PuntoVentaBienesDeCapital%, Fechaini%, FechaFin%)
        If LIMITEINFERIOR& > 0 And LIMITEINFERIOR& < LIMITESUPERIOR& Then NCExpo$ = BusquedaHuecos("nc", "e", PuntoVentaExportacion%, LIMITEINFERIOR&, LIMITESUPERIOR)
    End If

    Call TRACE(24, 100, 100)
    Screen.MousePointer = 1

    Resultado$ = "" _
    & " Fc Comun A: " & ComunA$ & vbCrLf _
    & " Fc Comun B: " & ComunB$ & vbCrLf _
    & " Fc Credito: " & Credito$ & vbCrLf _
    & " Fc Bienes A: " & BienesA$ & vbCrLf _
    & " Fc Bienes B: " & BienesB$ & vbCrLf _
    & " Fc Expo: " & Expo$ & vbCrLf _
    & " ND Comun A: " & NDComunA$ & vbCrLf _
    & " ND Comun B: " & NDComunB$ & vbCrLf _
    & " ND Credito: " & NDCredito$ & vbCrLf _
    & " ND Bienes A: " & NDBienesA$ & vbCrLf _
    & " ND Bienes B: " & NDBienesB$ & vbCrLf _
    & " ND Expo: " & NDExpo$ & vbCrLf _
    & " NC Comun A: " & NCComunA$ & vbCrLf _
    & " NC Comun B: " & NCComunB$ & vbCrLf _
    & " NC Credito: " & NCCredito$ & vbCrLf _
    & " NC Bienes A: " & NCBienesA$ & vbCrLf _
    & " NC Bienes B: " & NCBienesB$ & vbCrLf _
    & " NC Expo: " & NCExpo$ & vbCrLf

    If MostrarResultado = True Then MsgBox Resultado$, , "Correlatividad de comprobantes"

End Sub

Function BusquedaHuecos(TipoComprobante As String, LETRA As String, Talonario As Integer, Inicio As Long, FIN As Long) As String
    
    CantidadDias% = 45
    
    Dim RstFC As ADODB.Recordset
    
    On Error GoTo ControlError
    Set RstFC = FLEXCONN.Execute("exec HUECOS_EN_FACTURACION '" & TipoComprobante & "','" & LETRA$ & "'," & CStr(Talonario) & "," & Inicio & "," & FIN)
    Resultado$ = "OK"
    If Not (RstFC.EOF And RstFC.EOF) Then
        Resultado$ = ""
        Do While Not RstFC.EOF
            Comprobante$ = SAFETEXT$(RstFC!COMPROBANTES)
            If TRIM$(Comprobante$) <> "" Then
                Resultado$ = Resultado$ & SAFETEXT$(RstFC!COMPROBANTES) & ", "
            End If
            RstFC.MoveNext
        Loop
    End If
    On Error GoTo 0
    
    On Error Resume Next
    RstFC.Close
    On Error GoTo 0

    BusquedaHuecos = Resultado$
    Exit Function

ControlError:
    MsgBox Err.Description
    On Error GoTo 0
    BusquedaHuecos = "Error"

End Function


Function BusquedaHuecos_LimiteInferior(Comprobante As String, LETRA As String, Talonario As Integer, FechaInicio As Integer, FechaFin As Integer) As Long

    LIMITEINFERIOR& = XVALO(VALOBADA("cajabanc", "max(numecomp)", "fechconta>='" & FETEXCOR1$(FechaInicio) & "' and fechconta<='" & FETEXCOR1$(FechaFin) & "' and tipomovim='fven' and opcimovim='" & Comprobante & "' and varimovim='" & LETRA & "' and numetalo=" & CStr(Talonario)))
    If LIMITEINFERIOR& < 1 Then
        LIMITEINFERIOR& = XVALO(VALOBADA("cajabanc", "max(numecomp)", "fechconta<='" & FETEXCOR1$(FechaFin) & "' and tipomovim='fven' and opcimovim='" & Comprobante & "' and varimovim='" & LETRA & "' and numetalo=" & CStr(Talonario)))
    End If
    
    BusquedaHuecos_LimiteInferior = LIMITEINFERIOR&
        
End Function

Function BusquedaHuecos_LimiteSuperior(Comprobante As String, LETRA As String, Talonario As Integer, FechaInicio As Integer, FechaFin As Integer) As Long

    LIMITESUPERIOR& = XVALO(VALOBADA("cajabanc", "max(numecomp)", "fechconta>='" & FETEXCOR1$(FechaFin) & "' and tipomovim='fven' and opcimovim='" & Comprobante & "' and varimovim='" & LETRA & "' and numetalo=" & CStr(Talonario)))

    BusquedaHuecos_LimiteSuperior = LIMITESUPERIOR&
    
End Function


Sub CargaAsientoConciTarjeta(CONDIX$, DESDE%, HASTA%, ARCHIX$, OK%, RET_TX$)

            DES1$ = FETEXCOR1$(DESDE%)
            HAS1$ = FETEXCOR1$(HASTA%)
            RET_TX$ = ""
            JX& = FRMZELECHO.MSF2.Rows - 1
            '
            Fallo% = 0

            SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
            SQLX$ = SQLX$ + "WHERE FECHEMISI >= '" & DES1$ & "'"
            SQLX$ = SQLX$ + "AND FECHEMISI <= '" & HAS1$ & "' "
            SQLX$ = SQLX$ + "AND TIPOMOVIM = 'GASTC'"
            SQLX$ = SQLX$ + " AND " & CONDIX$
            SQLX$ = SQLX$ + " ORDER BY FECHEMISI ASC "
            Set rs = READSET(SQLX$)
            
            With rs
              Do While Not .EOF
                   FalloNoExisteCuentaEnProvincia% = 0
                   FalloCuitAgente% = 0
                   FalloComprobanteAplicadoImporte% = 0
                   FalloComprobanteAplicadoConstancia% = 0
                   FalloComprobanteAplicadoFactura% = 0
                   FalloComprobanteAplicadoTipoComprobante% = 0
                   FalloComprobanteLetra% = 0
                   
                   ' *** SI LA VARIABLE YA CONTIENE RETENCIONES, ENTONCES AGREGA UN SALTO DE LINEA
                   If TRIM$(RET_TX$) <> "" Then RET_TX$ = RET_TX$ + vbCrLf
                
                   ' *** OBTIENE LA JURISDICCION SEGUN LA CUENTA DE RETENCIONES
                   CuentaContableMedioCobranza% = XVALO(!CUECONTAI)
                   RegistroProvincia$ = FILTERGETRECORD$("PROVINB", 5, MKI$(CuentaContableMedioCobranza%))
                   If TRIM$(RegistroProvincia$) = "" Then FalloNoExisteCuentaEnProvincia% = 1   ' VERIFICA SI EXISTE CTA CONTABLE DE RETENCION
                                                                                                ' EN LA TABLA DE PROVINCIAS
                   CodigoDeJurisdiccion$ = CStr(CVI(Mid$(RegistroProvincia$, 37, 2)))
                   nombreProvincia$ = TRIM$(Mid$(RegistroProvincia$, 2, 25))
                      
                   ' *** DATOS COMPLEMENTARIOS
                   FE% = CVINT(!FECHEMISI)
                   FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
                   SUCURSAL$ = "0001"
                   DOCUMENTO$ = SAFETEXT$(!codicom_lar)
                   IMPORTEAPLIC# = XVALO(!idebecomp)
                   DATOCOBRA$ = SAFETEXT$(!DESCRIMOV)
                   DATOCOBRA$ = Mid$(DATOCOBRA$, 27)
                   TICOMP$ = "": NUMCONSTANCIA$ = "": LET1$ = "": COMPROBANTEORIGINAL$ = ""
                   
                   ' OBTIENE CUIT DEL AGENTE CORRESPONDIENTE AL BANCO. LO SACA A PARTIR DE LA COINCIDENCIA DE LA CUENTA CONTABLE DEL MOVIMIENTO QUE
                   ' SE GRABA CON LA LIQUIDACION DE CUPONES Y LA CUENTA MADRE DE LA TABLA DE PROVEEDORES.
                   CContableContraPartidaCupon% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", "TIPOMOVIM='CONCITC' AND CODICOM_LAR='" & DOCUMENTO$ & "' AND IDEBECOMP > 0.005"))
                   CuitAgenteRetencion$ = SAFETEXT$(VALOBADA("PROVED12", "CUIT_CLI", "CUMA_CLI='" & CStr(CContableContraPartidaCupon%) & "'"))
                   
                   Valid% = VALICUIT%(CuitAgenteRetencion$)
                   If Valid% < 1 Then
                        FalloCuitAgente% = 1
                   Else
                        CUITX$ = FormatearCuitValido$(CuitAgenteRetencion$) ' LE AGREGA FORMATO CON LOS SEPARADORES - EN CASO QUE NO LOS TENGA
                   End If
                  
                   NC = XVALO(VALOBADA("PROVED12", "NUME_CLI", "CUMA_CLI='" & CStr(CContableContraPartidaCupon%) & "'"))
                   NC = (-1) * NC
                   
                   Erase CADENATEX$
    
                   Call TEXTOLOTES$(DATOCOBRA$, ";")

                   ' *** DATOS DEL COMPROBANTE QUE ORIGINA LA RETENCION Y LA CONSTANCIA
                   On Error Resume Next
                   For i& = 1 To UBound(CADENATEX$)
                     Select Case i&
                       Case 1
                         TICOMP$ = Left$(TRIM$(CADENATEX$(i&)), 1)
                       Case 2
                         NUMCONSTANCIA$ = Left$(TRIM$(CADENATEX$(i&)), 16)
                       Case 3
                         LET1$ = Left$(TRIM$(CADENATEX$(i&)), 1)
                       Case 4
                         COMPROBANTEORIGINAL$ = Left$(TRIM$(CADENATEX$(i&)), 20)
                     End Select
                   Next i&
                                     
                   If IMPORTEAPLIC# < 0.005 Then FalloComprobanteAplicadoImporte% = 1
                   If TRIM$(NUMCONSTANCIA$) = "" Then FalloComprobanteAplicadoConstancia% = 1
                   If TRIM$(DOCUMENTO$) = "" Then FalloComprobanteAplicadoFactura% = 1
                   If TICOMP$ = "" Then FalloComprobanteAplicadoTipoComprobante% = 1
                   If LET1$ = "" Then FalloComprobanteLetra% = 1
                    
                   ' *** CARGA DETALLE DE LOS REGISTROS QUE NO CUMPLEN CON LAS VALIDACIONES
                   If FalloNoExisteCuentaEnProvincia% = 1 Or FalloCuitAgente% = 1 Or FalloComprobanteAplicadoImporte% = 1 Or FalloComprobanteAplicadoConstancia% = 1 Or FalloComprobanteAplicadoFactura% = 1 Or FalloComprobanteAplicadoTipoComprobante% = 1 Or FalloComprobanteLetra% = 1 Then
                         JX& = JX& + 1
                         FRMZELECHO.MSF2.Rows = JX& + 1
                         FRMZELECHO.MSF2.TextMatrix(JX&, 1) = DOCUMENTO$
                         FRMZELECHO.MSF2.TextMatrix(JX&, 2) = (-1) * NC
                         FRMZELECHO.MSF2.TextMatrix(JX&, 3) = rasocli$(NC)
                         FRMZELECHO.MSF2.TextMatrix(JX&, 4) = XVALO(!NumeComp)
                         FRMZELECHO.MSF2.TextMatrix(JX&, 5) = FECHATEX$(CVINT(!FECHEMISI))
                         Separador1$ = "": TXTERROR$ = ""
                         If FalloNoExisteCuentaEnProvincia% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "No Existe Provincia Vinculada con Cuenta Contable de Retención"
                         If FalloCuitAgente% > 0 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Cuit Inválido del Agente de Retención."
                         If FalloComprobanteAplicadoImporte% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Importe Retenido Menor a Cero"
                         If FalloComprobanteAplicadoConstancia% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Numero de Constancia"
                         If FalloComprobanteAplicadoPuntoVenta% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Punto de Venta"
                         If FalloComprobanteAplicadoFactura% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Numero de Factura"
                         If FalloComprobanteAplicadoTipoComprobante% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Tipo de Comprobante"
                         If FalloFacturaSinCae% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Factura Aplicada Sin CAE": Separador1$ = " - "
                         
                         FRMZELECHO.MSF2.TextMatrix(JX&, 6) = TXTERROR$
                    End If
                   
                    ' *** ARMA REGISTRO DEL ARCHIVO DE RETENCIONES
72                  GENERO_ARCHIVO% = 1
                     If Len(CodigoDeJurisdiccion$) < 3 Then CodigoDeJurisdiccion$ = String$(3, "0") & CodigoDeJurisdiccion$: CodigoDeJurisdiccion$ = Right(CodigoDeJurisdiccion$, 3)
                    Call REPLA(TX$, CodigoDeJurisdiccion$, 1, 3)
                     If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
                    Call REPLA(TX$, CUITX$, 4, 13)
                     If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
                    Call REPLA(TX$, FEMIS$, 17, 10)
                    Call REPLA(TX$, SUCURSAL$, 27, 4)
                     If Len(NUMCONSTANCIA$) < 16 Then NUMCONSTANCIA$ = String$(16, "0") & NUMCONSTANCIA$: NUMCONSTANCIA$ = Right(NUMCONSTANCIA$, 16)
                    Call REPLA(TX$, NUMCONSTANCIA$, 31, 16)
                    Call REPLA(TX$, TICOMP$, 47, 1)
                    Call REPLA(TX$, LET1$, 48, 1)
                     If Len(COMPROBANTEORIGINAL$) < 20 Then COMPROBANTEORIGINAL$ = String(20, "0") & COMPROBANTEORIGINAL$: COMPROBANTEORIGINAL$ = Right$(COMPROBANTEORIGINAL$, 20)
                    Call REPLA(TX$, COMPROBANTEORIGINAL$, 49, 20)
                    IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F11.2"))
                    ' FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
                    If InStr(IMPORTEAPL$, ".") > 0 Then
                     IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
                    Else
                      IMPORTEAPL$ = IMPORTEAPL$ & ",00"
                    End If
                    If Len(IMPORTEAPL$) < 11 Then IMPORTEAPL$ = String$(11, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 11)
                  
                    Call REPLA(TX$, IMPORTEAPL$, 69, 11)
                        RET_TX$ = RET_TX$ + TX$ ' CARGO EN UNA VARIABLE EL TEXTO QUE VA DE RETORNO A LA LLAMADA A ESTA RUTINA.
                        
                    TTXX$ = REGBLAN$("LIREPEIB")
                    Call REPLA(TTXX$, MKI(FE%), 1, 2)
                    Call REPLA(TTXX$, DOCUMENTO$, 3, 18)
                    Call REPLA(TTXX$, rasocli$(NC), 21, 48)
                    Call REPLA(TTXX$, NUMCONSTANCIA$, 69, 24)
                    Call REPLA(TTXX$, MKD$(IMPORTEAPLIC#), 93, 8)
                    Call REPLA(TTXX$, Space(91), 101, 91)
                    If TXTERROR$ <> "" Then
                        TextoError$ = TXTERROR$
                        Call FRATEXTO(TextoError$, 60)
                        For KKII% = 1 To CARETEX%
                            Call REPLA(TTXX$, RETEX(KKII%), 101, 60)
                            Call REGAPP("LIREPEIB", TTXX$)
                        Next KKII%
                    Else
                        Call REGAPP("LIREPEIB", TTXX$)
                    End If
                    
                   .MoveNext
              Loop
              '
            End With
End Sub


Sub CargaConciETENCIONIVA(Des%, Has%, ARCHIX$, OK%, RET_TX$)
    RET_TX$ = ""
    DES1$ = FETEXCOR1$(DESDE%)
    HAS1$ = FETEXCOR1$(HASTA%)
    
    Campos$ = "Cod.Jurisdicción/A3;Cuit/a13;F.Asiento/A10;N.Conciliacion/A12;Const/A20;Debe/F12.2;Haber/F12.2;Prov/F7;Fallo;DH/A4"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = 12000
    FRMZELECHO.Caption = "Asientos No Conformes Para Emisión de Archivo de Transferencia"
    
    
    CUECONTA$ = CONTROL$("OPRECIBO", "CURETIVA")
    CUE% = CUECOINT%(CUECONTA$)
    If CUE% < 1 Then
      Call COMUNI("Falta Configurar Cuenta Contable de Retención de I.V.A.\Configurar la Misma Desde Setup de Recibos")
      OK% = -1
      GoTo 99
    End If
    
    CODREG& = XVALO(CONTROL$("OPRECIBO", "CREREIVA"))
    If CODREG& < 1 Then
      OK% = -1
      Call COMUNI("Falta Configurar Código de Retención de I.V.A.\Configurar la Misma Desde Setup de Recibos")
      GoTo 99
    End If
    '
'    CODREGXTC = XVALO(CONTROL("OPCHEQ07", "CRETIVAT"))
'    If CODREGXTC < 1 Then
'      OK% = -1
'      Call COMUNI("Falta Configurar Código de Retención de I.V.A. de Tarjetas de Crédito\Configurar la Misma Desde Setup de Gestión de Valores")
'      GoTo 99
'    End If

    JX& = 0
    'RECORRO EL DETAASIEN SEGUN CONDICION
    SQLX$ = "SELECT * FROM DETAASIEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE FECHASIEN >= '" & DES1$ & "'"
    SQLX$ = SQLX$ + "AND FECHASIEN <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND CUECOINT = " & CUE%
    SQLX$ = SQLX$ + " AND STATPASE >=0 "
    SQLX$ = SQLX$ + " ORDER BY NUIDASIEN, FECHASIEN ASC "
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF

        NUMEAS$ = nuidasien
        IMPODEBE1 = XVALO(!idebepase)
        IMPOHABE1 = XVALO(!ihabepase)
        FE% = CVINT(!FECHASIEN)
        CUE% = XVALO(!CUECOINT)
        ASIENTO$ = SAFETEXT$(!nuidasien)
        TTXX$ = SAFETEXT$(!DATRETPER)
        
        COMPROBAX$ = Mid$(TTXX$, 1, 1)
        CERTIFICADOX$ = Mid$(TTXX$, 2, 20)
        LETRAX$ = Mid$(TTXX$, 22, 1)
        'CARGO LA INFORMACION EN LA GRILLA
        FRMZELECHO.MSF2.Rows = JX& + 1
        FRMZELECHO.MSF2.TextMatrix(JX&, 1) = CODREG&
        FRMZELECHO.MSF2.TextMatrix(JX&, 2) = ""
        FRMZELECHO.MSF2.TextMatrix(JX&, 3) = FECHATEX$(FE%)
        FRMZELECHO.MSF2.TextMatrix(JX&, 4) = ASIENTO$
        FRMZELECHO.MSF2.TextMatrix(JX&, 5) = CERTIFICADOX$
         If IMPODEBE1 > 0 Then DH$ = "D"
         If IMPOHABE1 > 0 Then DH$ = "H"
        FRMZELECHO.MSF2.TextMatrix(JX&, 6) = FORMATNUM$(IMPODEBE1, "F12.2")
        FRMZELECHO.MSF2.TextMatrix(JX&, 7) = FORMATNUM$(IMPOHABE1, "F12.2")
        FRMZELECHO.MSF2.TextMatrix(JX&, 8) = Proveedor
        FRMZELECHO.MSF2.TextMatrix(JX&, 9) = "" 'SI EL IMPORTE ESTA DEL LADO DEL DEBE O DEL HABER
        FRMZELECHO.MSF2.TextMatrix(JX&, 10) = DH$ 'SI EL IMPORTE ESTA DEL LADO DEL DEBE O DEL HABER
        .MoveNext
      Loop
      End With
      '
      'RECORRER LA GRILLA Y TOMAR LOS OTROS VALORES
      'AHORA RECORRO LA GRILLA
      JX& = 0: Fallo% = 0
      With FRMZELECHO
       For JX& = 1 To .MSF2.Rows - 1
          ASIENTO$ = FRMZELECHO.MSF2.TextMatrix(JX&, 3)
          DH$ = FRMZELECHO.MSF2.TextMatrix(JX&, 10)
          If DH$ = "D" Then
             importex# = XVALO(FRMZELECHO.MSF2.TextMatrix(JX&, 9))
             CONDI$ = "NUIDASIEN = " & ASIENTO$ & " AND IHABEPASE > 0  "
          Else
             importex# = XVALO(FRMZELECHO.MSF2.TextMatrix(JX&, 9))
             CONDI$ = "NUIDASIEN = " & ASIENTO$ & " AND IDEBEPASE > 0  "
          End If
          'VALIDO QUE HAYA UNA SOLA CUENTA DE CONTRAPARTIDA
          If CantRegistros&("DETAASIEN", CONDI$, "NOMESS") = 1 Then
             CUE% = VALOBADA("DETAASIEN", "CUECOINT", CONDI$, "NOMESS")
             NPROV = PROV_SegCtaCble(CUE%)
             PROVINCIAX$ = PROVCLI$(-1 * NPROV)
             CUPRO$ = CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%) 'RET_REIB% ES DE RETORNO
             FRMZELECHO.MSF2.TextMatrix(JX&, 1) = CUPRO$
                CUITX$ = CuitCli(-1 * NPROV)
             FRMZELECHO.MSF2.TextMatrix(JX&, 2) = CUITX$
           Else
             Fallo% = 1
             FRMZELECHO.MSF2.TextMatrix(JX&, 11) = "MAS DE 1"
             FRMZELECHO.MSF2.TextMatrix(JX&, 12) = "FALLO"
           End If
       Next JX&
      End With
      'VALIDA SI HAY FALLOS MUESTRA LOS FALLOS
      If Fallo% > 0 Then
        OK% = -1
        With FRMZELECHO
          For JX1& = 1 To .MSF2.Rows - 1
            If TRIM$(.MSF2.TextMatrix(JX1&, 12)) <> "FALLO" Then
              Call cmdEliminaItem(.MSF2, JX1&)
              JX1& = JX1& - 1
              If JX1& >= .MSF2.Rows Then Exit For
            End If
          Next JX1&
          Call COMUNI("Imposible Procesar Archivo de Transferencia\Al Aceptar Este Mensaje\Se Visualizará el Detalle de Los Mismos")
          .Show 1
          GoTo 99
        End With
      Else 'SI NO HAY FALLOS AGREGA AL ARCHIVO
        OK% = 1
        With FRMZELECHO
          For JX& = 1 To .MSF2.Rows - 1
'      1.-Cod.Jurisdicción
'      2.-Cuit
'      3.-F.Asiento
'      4.-Asiento
'      5.-Const
'      6.-Debe
'      7.-Haber
'      8.-Prov
'      9.-Fallo
'      10.-DH
            CUPRO$ = .MSF2.TextMatrix(JX&, 1)
            CUITX$ = .MSF2.TextMatrix(JX&, 2)
            FE% = FECHANUM(FRMZELECHO.MSF2.TextMatrix(JX&, 3))
              FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
            NUMCONSTANCIA$ = .MSF2.TextMatrix(JX&, 5)
            IMPORTEAPLIC# = 0
            DH$ = FRMZELECHO.MSF2.TextMatrix(JX&, 10)
            If DH$ = "D" Then IMPORTEAPLIC# = XVALO(.MSF2.TextMatrix(JX&, 6))
            If DH$ = "H" Then IMPORTEAPLIC# = XVALO(.MSF2.TextMatrix(JX&, 7))

            If Len(CUPRO$) < 3 Then CUPRO$ = String$(3, "0") & CUPRO$: CUPRO$ = Right(CUPRO$, 3)
            Call REPLA(TX$, CUPRO$, 1, 3)
             If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
            Call REPLA(TX$, CUITX$, 4, 13)
             If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
            Call REPLA(TX$, FEMIS$, 17, 10)
              If Len(NUMCONSTANCIA$) < 16 Then NUMCONSTANCIA$ = String$(16, "0") & NUMCONSTANCIA$: NUMCONSTANCIA$ = Right(NUMCONSTANCIA$, 16)
            Call REPLA(TX$, NUMCONSTANCIA$, 27, 16)
            IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F15.2"))
            'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
            If InStr(IMPORTEAPL$, ".") > 0 Then
              IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
            Else
              IMPORTEAPL$ = IMPORTEAPL$ & ",00"
            End If
            If Len(IMPORTEAPL$) < 16 Then IMPORTEAPL$ = String$(16, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 16)
            Call REPLA(TX$, IMPORTEAPL$, 69, 16)
            RET_TX$ = RET_TX$ + TX$ + vbCrLf 'CARGO EN UNA VARIABLE EL TEXTO QUE VA DE RETORNO A LA LLAMADA A ESTA RUTINA.
          Next JX&
        End With
      End If
   'HASTA ACA
      
99
End Sub

'Sub CargaAsientosRETENCIONIIBB_TMP(CONDIX$, DES%, HAS%)
'
'    DES1$ = FETEXCOR1$(DESDE%)
'    HAS1$ = FETEXCOR1$(HASTA%)
'    CONTRAPARDEBE% = 0
'    CONTRAPARHABE% = 0
'    NUMEAS_ANT$ = ""
'    CAMPOS$ = "Cod.Jurisdicción/A3;Cuit/a13;F.Asiento/A10;Sucur/A4;Const/A15;T.C./A4;Let/A3;Cmbpte.Orig./A14;Imp.Ret./F12.2;Prov/F7;Prov/a4;Fallo/f2"
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
'    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
'    If ANTOT < 5500 Then ANTOT = 5500
'    FRMZELECHO.Width = 12000
'    FRMZELECHO.Caption = "Recibos Sin Imputar"
'    JX& = 0
'    SQLX$ = "SELECT * FROM DETAASIEN WITH(NOLOCK)"
'    SQLX$ = SQLX$ + "WHERE FECHASIEN >= '" & DES1$ & "'"
'    SQLX$ = SQLX$ + "AND FECHASIEN <= '" & HAS1$ & "' "
'    SQLX$ = SQLX$ + " AND " & CONDIX$
'    SQLX$ = SQLX$ + " AND STATPASE >=0 "
'    SQLX$ = SQLX$ + " ORDER BY NUIDASIEN, FECHASIEN ASC "
'    Set RS = READSET(SQLX$)
'    With RS
'      Do While Not .EOF
'        NUMEAS$ = NUIDASIEN
'        If NUMEAS_ANT$ <> NUMEAS$ And NUMEAS_ANT$ <> "" Then
'            FALLO% = 0
'            If CONTRAPARDEBE% = 1 Then
'               CONDI$ = "NUIDASIEN = " & NUMEAS_ANT$ & " AND IDEBEPASE > 0  "
'               CUE% = VALOBADA("DETAASIEN", "CUECOINT", CONDI$, "NOMESS")
'               CTAMADREPROV& = PROV_SegCtaCble(CUE%)
'               NPROV = PROV_SegCtaCble(CUE%)
'               PROVINCIAX$ = PROVCLI$(-1 * NPROV)
'               PROVINCIAX$ = PROVCLI$(-1 * NPROV)
'               CUPRO$ = CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%) 'RET_REIB% ES DE RETORNO
'            ElseIf CONTRAPARHABE% = 1 Then
'               CONDI$ = "NUIDASIEN = " & NUMEAS_ANT$ & " AND IHABEPASE > 0  "
'               CUE% = VALOBADA("DETAASIEN", "CUECOINT", CONDI$, "NOMESS")
'            Else
'                FALLO% = 1
'            End If
'            NPROV = PROV_SegCtaCble(CUE%): If NPROV < 1 Then FALLO% = 1
'            CUITX$ = CUITCLI(-1 * NPROV): If TRIM$(CUITX$) = "" Then FALLO% = 1
'            PROVINCIAX$ = PROVCLI$(-1 * NPROV): If TRIM$(ECOARCH$("PROVINC", PROVINCIAX$)) = "" Then FALLO% = 1
'            CUPRO$ = TRIM$(CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%)) 'RET_REIB% ES DE RETORNO)
'            If CUPRO$ = "" Then FALLO% = 1
'            If FALLO% > 0 Then
'                JX& = JX& + 1
'                'CAMPOS$ = "Cod.Jurisdicción/A3;Cuit/a13;F.Asiento/A10;Sucur/A4;Const/A15;T.C./A4;Let/A3;Cmbpte.Orig./A14;Imp.Ret./F12.2;Prov/F7;Prov/a4;Fallo/f2"
'
'                FRMZELECHO.msf2.Rows = JX& + 1
'                FRMZELECHO.msf2.TextMatrix(JX&, 1) = CUPRO$
'                FRMZELECHO.msf2.TextMatrix(JX&, 2) = CUITX$
'                FRMZELECHO.msf2.TextMatrix(JX&, 3) = FECHATEX$(FE%)
'                FRMZELECHO.msf2.TextMatrix(JX&, 4) = NUSUCU$
'                FRMZELECHO.msf2.TextMatrix(JX&, 5) = CERTIFICADOX$
'                FRMZELECHO.msf2.TextMatrix(JX&, 6) = COMPROBAX$
'                FRMZELECHO.msf2.TextMatrix(JX&, 7) = LETRAX$
'                FRMZELECHO.msf2.TextMatrix(JX&, 8) = ASIENTO$
'                 If IMPODEBE1 > 0 Then IMPORETE# = IMPODEBE1
'                 If IMPOHABE1 > 0 Then IMPORETE# = IMPOHABE1
'
'                FRMZELECHO.msf2.TextMatrix(JX&, 9) = FORMATNUM$(IMPORETE#, "F12.2")
'                FRMZELECHO.msf2.TextMatrix(JX&, 10) = NPROV
'                FRMZELECHO.msf2.TextMatrix(JX&, 11) = ECOARCH$("PROVINC", PROVINCIAX$)
'
'            End If
'
'        End If
'
'            IMPODEBE1 = XVALO(!IDEBEPASE)
'            IMPOHABE1 = XVALO(!IHABEPASE)
'            FE% = CVINT(!FECHASIEN)
'            NUSUCU$ = "0001"
'            CUE% = XVALO(!CUECOINT)
'            ASIENTO$ = SAFETEXT$(!NUIDASIEN)
'            TXTDATO$ = SAFETEXT$(!DATREPER)
'            TTXX$ = SAFETEXT$(!DATRETPER)
'
'            COMPROBAX$ = Mid$(TTXX$, 1, 1)
'            CERTIFICADOX$ = Mid$(TTXX$, 2, 20)
'            LETRAX$ = Mid$(TTXX$, 22, 1)
'
'            If IMPODEBE1 > 0 Then
'              CONTRAPARDEBE% = CONTRAPARDEBE% + 1
'
'            ElseIf IMPOHABE1 > 0 Then
'              CONTRAPARHABE% = CONTRAPARHABE% + 1
'            End If
'
'            NUMEAS_ANT$ = NUMEAS$
'            .MoveNext
'      Loop
'      End If
'
'                           If Len(CUPRO$) < 3 Then CUPRO$ = String$(3, "0") & CUPRO$: CUPRO$ = Right(CUPRO$, 3)
'                    Call REPLA(TX$, CUPRO$, 1, 3)
'                     If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
'                    Call REPLA(TX$, CUITX$, 4, 13)
'                     If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
'                    Call REPLA(TX$, FEMIS$, 17, 10)
'                    Call REPLA(TX$, SUCURSAL$, 27, 4)
'                      If Len(NUMCONSTANCIA$) < 16 Then NUMCONSTANCIA$ = String$(16, "0") & NUMCONSTANCIA$: NUMCONSTANCIA$ = Right(NUMCONSTANCIA$, 16)
'                    Call REPLA(TX$, NUMCONSTANCIA$, 31, 16)
'                    Call REPLA(TX$, TICOMP$, 47, 1)
'                    Call REPLA(TX$, LET1$, 48, 1)
'                      If Len(PVTA1$) < 4 Then PVTA1$ = String(4, "0") & PVTA1$: PVTA1$ = Right$(PVTA1$, 4)
'                      If Len(nufact$) < 16 Then nufact$ = String(16, "0") & nufact$: nufact$ = Right$(nufact$, 16)
'                      facturaaplicada$ = PVTA1$ & nufact$
'                    Call REPLA(TX$, facturaaplicada$, 49, 20)
'                    IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F11.2"))
'                    'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
'                    If InStr(IMPORTEAPL$, ".") > 0 Then
'                      IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
'                    Else
'                      IMPORTEAPL$ = IMPORTEAPL$ & ",00"
'                    End If
'                    If Len(IMPORTEAPL$) < 11 Then IMPORTEAPL$ = String$(11, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 11)
'
'                    Call REPLA(TX$, IMPORTEAPL$, 69, 11)
'                    Print #NARCHI2%, TX$
'
'
'End Sub
'
'
Function CodigoDocumento(PosicionIva%) As String

    Select Case PosicionIva%
        Case 0: CodiDoc$ = "96"
        Case Else: CodiDoc$ = 80
    End Select
    
    CodigoDocumento$ = CodiDoc$
    
End Function

Function CodigoMoneda(Moneda%) As String

    Select Case Moneda%
        Case 1: CodiMo$ = "PES"
        Case 2: CodiMo$ = "DOL"
        Case Else: CodiMo$ = "000"
    End Select
    
    LeyendaMoneda$ = DEMONECO$(Moneda%)
    If InStr(1, UCase(LeyendaMoneda$), "EURO") > 0 Then
        CodiMo$ = "060"
        GoTo 99
    End If
    
    If InStr(1, UCase(LeyendaMoneda$), "DOLA") > 0 Then
        CodiMo$ = "DOL"
        GoTo 99
    End If
    
99  CodigoMoneda = CodiMo$
    
End Function



Sub BusquedaHuecos_StoreProcedure()

    If EXISTE_SP%("HUECOS_EN_FACTURACION") > 0 Then
      Procedimiento$ = "alter procedure HUECOS_EN_FACTURACION " & vbCrLf
    Else
      Procedimiento$ = "create procedure HUECOS_EN_FACTURACION " & vbCrLf
      EJECUTA_DIRECTO% = 1
    End If

    Procedimiento$ = Procedimiento$ & " @comprobante nvarchar(16)," & vbCrLf _
    & " @letra nvarchar(1)," & vbCrLf _
    & " @talonario int," & vbCrLf _
    & " @inicio int," & vbCrLf _
    & " @fin int" & vbCrLf _
    & " as" & vbCrLf _
    & " begin" & vbCrLf _
    & " declare @resultado nvarchar(max)" & vbCrLf _
    & " declare @fechamin datetime" & vbCrLf _
    & " set @resultado = '' " & vbCrLf _
    & " set @fechamin = getdate()-45-30" & vbCrLf
 
    Procedimiento$ = Procedimiento$ & " set nocount on" & vbCrLf _
    & " if @fin > @inicio " & vbCrLf _
    & " begin" & vbCrLf _
    & "  if object_id('HUECOS_EN_FACTURA') is null" & vbCrLf _
    & "  begin" & vbCrLf _
    & "      create table HUECOS_EN_FACTURA" & vbCrLf _
    & "      ( numero int )" & vbCrLf _
    & "  end" & vbCrLf _
    & "  delete from HUECOS_EN_FACTURA" & vbCrLf _
    & " " & vbCrLf _
    & "  while @inicio < @fin + 1 " & vbCrLf _
    & "  begin" & vbCrLf _
    & "      insert into HUECOS_EN_FACTURA values (@inicio)" & vbCrLf _
    & "      set @inicio = @inicio + 1" & vbCrLf _
    & "  end" & vbCrLf _
    & " " _
    & "  select * from (select t1.numero+1 AS Comprobantes FROM HUECOS_EN_FACTURA as t1 LEFT OUTER JOIN (select numecomp from CAJABANC with(nolock) where opcimovim=@comprobante and varimovim=@letra and numetalo=@talonario and fechconta>@fechamin) as t2 ON (t1.numero+1=t2.NumeComp) WHERE t2.NumeComp IS NULL) t3 where t3.Comprobantes<@fin" & vbCrLf _
    & " " _
    & " end" & vbCrLf _
    & " else" & vbCrLf _
    & " begin" & vbCrLf _
    & "  select '' as Comprobantes" & vbCrLf _
    & " end" & vbCrLf _
    & " set nocount off" & vbCrLf _
    & " end"

    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (Procedimiento$)
    Else
      If CAMBIOSP("HUECOS_EN_FACTURACION", Procedimiento$) Then
        FLEXCONN.Execute (Procedimiento$)
      End If
    End If

End Sub

Sub DETAVEN()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     MESAN = MESASEL$
   End If
   Call FRESHALL
    
   If VALIMESAN% < 1 Then Exit Sub
   FF1% = FECHANUM(DESDAT)
   FF2% = FECHANUM(HASDAT)
   FECHDE = FETEXCOR1(FF1%)
   FECHHA = FETEXCOR1(FF2%)
   '
   Call APERBASDAT("STOCKS")
   '
   CONDI$ = " FECHMOV >= '" & FECHDE & "' AND FechMov <= '" & FECHHA & "' AND CodiMovi ='RT'"
   '
   XX& = CantRegistros&("MOVISTO", CONDI$)
   '
   If XX& < 1 Then
      Call MENSERR(24, "No Existen Registros Según Consulta SQL")
      Exit Sub
   End If
   '
   SQLX$ = "SELECT * FROM MOVISTO"
   SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + " ORDER BY Doprilar"
   '
   Dim DETATEMP As ADODB.Recordset
   Set DETATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   DETATEMP.MoveFirst
   JJ& = 0
   With DETATEMP
        Do While Not DETATEMP.EOF
            Call TRACE(24, UU&, XX&)
            UU& = UU& + 1
            ZZ$ = REBLA$
            NUCOMP$ = SAFETEXT$(TRIM$(!DOSECLAR))
            COMP$ = UCase$(Left$(NUCOMP$, 2))
            LETRA$ = UCase$(Right$(NUCOMP$, 1))
            NUMCOM$ = Mid$(NUCOMP$, 4, 13)
            If COMP$ = "FC" Then Tcomp$ = "FACTURA"
            If COMP$ = "ND" Then Tcomp$ = "NOTA DE DEBITO"
            If COMP$ = "NC" Then Tcomp$ = "NOTA DE CREDITO"
            Call REPLA(ZZ$, Tcomp$, 1, 16)
            Call REPLA(ZZ$, LETRA$, 17, 2)
            Call REPLA(ZZ$, NUMCOM$, 19, 16)
            FECHEMIS$ = VALOBADA("CAJABANC", "FechConta", "CodiCom_Lar = '" & NUCOMP$ & "'")
            Call REPLA(ZZ$, FECHEMIS$, 35, 12)
            NCLI = XVALO(!NUME_CLIE)
            RASO$ = rasocli(NCLI)
            Call REPLA(ZZ$, RASO$, 47, 48)
            DOMCLI$ = DOMICLI$(NCLI)
            Call REPLA(ZZ$, DOMCLI$, 95, 48)
            LOCALI$ = LOCACLI$(NCLI)
            Call REPLA(ZZ$, LOCALI$, 143, 32)
            CUITT$ = CuitCli$(NCLI)
            Call REPLA(ZZ$, CUITT$, 175, 13)
            Call REPLA(ZZ$, MKI$(XVALO(!COD_INTE)), 188, 2)
            Call REPLA(ZZ$, MKS$(XVALO(!Cantsalid)), 190, 4)
            Call REPLA(ZZ$, MKD$(XVALO(!VALOUNIT)), 194, 8)
            Call REPLA(ZZ$, SAFETEXT(!DOPRILAR), 202, 18)
            Call REPLA(ZZ$, SAFETEXT(!FeReMovi), 220, 12)
                NUMEROSERIE$ = VALOBADA("TANUMSE", "NUMEROSERIE", "NUMREMITO='" & !DOPRILAR & "' AND COD_INTE=" & CStr(!COD_INTE))
            Call REPLA(ZZ$, NUMEROSERIE$, 232, 24)
            
            '
            JJ& = JJ& + 1
            Call GRAREG("DETAVEN", ZZ$, JJ&)
            .MoveNext
       Loop
  
    Call SETULTREG("DETAVEN", JJ&)
    Call CIERRARCH("DETAVEN")
    Call FILESORT("DETAVEN", "", 4, 4, "ASC")

   End With
   
   Call FINAL("*DETAVEN")

End Sub

Sub DETAVEN_MDT()
   Call FRESHALL
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     MESAN = MESASEL$
   End If
   Call FRESHALL
    
   If VALIMESAN% < 1 Then Exit Sub
   FF1% = FECHANUM(DESDAT)
   FF2% = FECHANUM(HASDAT)
   FECHDE = FETEXCOR1(FF1%)
   FECHHA = FETEXCOR1(FF2%)
   '
   Call APERBASDAT("STOCKS")
   '
   CONDI$ = " FECHMOV >= '" & FECHDE & "' AND FechMov <= '" & FECHHA & "' AND CodiMovi ='RT'"
   '
   XX& = CantRegistros&("MOVISTO", CONDI$)
   '
   If XX& < 1 Then
      Call MENSERR(24, "No Existen Registros Según Consulta SQL")
      Exit Sub
   End If
   '
   Campos$ = "Tcomp/A16;L/A4;Documento/A18;Fecha/A8;Cta/F7;R.Social/A32;"
   Campos$ = Campos$ + "Domicilio/A32;Localidad/A24;Cuit/A13;Código/A18;Descripción/A32;Cantidad/F11.1;"
   Campos$ = Campos$ + "P.Unit./F12.2;Importe/F12.2;T.X.Fact./F12.2;Remito/A18;F.Rem./A8;N.Serie/A16;Cae/A20;Cod.Nomenclador/A24"
   Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
   Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
   Screen.MousePointer = 11
   Call VectArticulos.cargarvector
   Call VectCliente.cargarvector
   ACUIMPO# = 0
   SQLX$ = "SELECT * FROM MOVISTO"
   SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
   SQLX$ = SQLX$ + " ORDER BY Doprilar,fechmov"
   '
   Dim DETATEMP As ADODB.Recordset
   Set DETATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   DETATEMP.MoveFirst
   JJ& = 0
   With DETATEMP
        Do While Not DETATEMP.EOF
            Call TRACE(24, JJ&, XX&)
            JJ& = JJ& + 1
            GRILLAGRAL.GRID.Rows = JJ& + 1
            NUCOMP$ = SAFETEXT$(TRIM$(!DOSECLAR))
            NUCOMPLETO$ = SAFETEXT$(TRIM$(!DOSECLAR))
            NCLI = XVALO(!NUME_CLIE)

            COMP$ = UCase$(Left$(NUCOMP$, 2))
            LETRA$ = UCase$(Right$(NUCOMP$, 1))
            NUMCOM$ = Mid$(NUCOMP$, 4, 13)
            If COMP$ = "FC" Then Tcomp$ = "FACTURA"
            If COMP$ = "ND" Then Tcomp$ = "NOTA DE DEBITO"
            If COMP$ = "NC" Then Tcomp$ = "NOTA DE CREDITO"
            FECHEMIS$ = FECHATEX$(CVINT(VALOBADA("CAJABANC", "FechConta", "CodiCom_Lar = '" & NUCOMP$ & "'")))
            RASO$ = VectCliente.nombre(NCLI)
            DOMCLI$ = VectCliente.Domicilio(NCLI)
            LOCALI$ = VectCliente.localidad(NCLI)
            CUITT$ = TRIM$(VectCliente.CuitCli(NCLI))
            CI1% = (XVALO(!COD_INTE))
            CODI$ = VectArticulos.Codigo(CI1%)
            DESC$ = VectArticulos.Descripcion(CI1%)
            If COMP$ = "NC" Then
              CANTI = XVALO(!CANTINGRE)
              CANTI = CANTI * -1
            Else
              CANTI = XVALO(!Cantsalid)
            End If
            IMPO# = XVALO(!VALOUNIT)
            'IMPO# = XVALO(!VALOUNITIVA)
            REMI$ = SAFETEXT(!DOPRILAR)
            FEREMI$ = FECHATEX$(CVINT(!FeReMovi))
            'NUMEROSERIE$ = VALOBADA("TANUMSE", "NUMEROSERIE", "NUMREMITO='" & !DOPRILAR & "' AND COD_INTE=" & CStr(!COD_INTE))
            NUMEROSERIE$ = SAFETEXT(!IDENLOTE)
            
            Cae$ = TRIM$(VALOBADA("CAJABANC", "CAE", "CodiCom_Lar = '" & NUCOMP$ & "'"))
            CODINOMENCLADOR$ = VectArticulos.Codigo_Ncm(XVALO(CI1%))
            GRILLAGRAL.GRID.TextMatrix(JJ&, 1) = Tcomp$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 2) = LETRA$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 3) = NUMCOM$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 4) = FECHEMIS$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 5) = NCLI
            GRILLAGRAL.GRID.TextMatrix(JJ&, 6) = RASO$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 7) = DOMCLI$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 8) = LOCALI$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 9) = CUITT$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 10) = CODI$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 11) = DESC$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 12) = FORMATNUM$(CANTI, "F9.1")
            GRILLAGRAL.GRID.TextMatrix(JJ&, 13) = FORMATNUM$(IMPO#, "F12.2")
            
            GRILLAGRAL.GRID.TextMatrix(JJ&, 14) = FORMATNUM$(CANTI * IMPO#, "F12.2")
            ACUIMPO# = ACUIMPO# + XVALO(GRILLAGRAL.GRID.TextMatrix(JJ&, 14))
            If NUCOMPLETO_ant$ <> NUCOMPLETO$ Then
               If TRIM$(NUCOMPLETO_ant$) <> "" Then
                 Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, JJ& - 1, 15)
                 ACU_IMPO_FACT = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP - IHABECOMP)", "TIPOMOVIM = 'FVEN' AND CodiCom_Lar = '" & NUCOMPLETO_ant$ & "' AND NUCLIEN = " & NCLIEN_ant, "NOMESS"))
                 GRILLAGRAL.GRID.TextMatrix(JJ& - 1, 15) = FORMATNUM$(ACU_IMPO_FACT, "F12.2")
               End If
            End If
            GRILLAGRAL.GRID.TextMatrix(JJ&, 16) = REMI$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 17) = FEREMI$
            If Cae$ = "" Then Cae$ = "------"
            If CODINOMENCLADOR$ = "" Then CODINOMENCLADOR$ = "------"
            GRILLAGRAL.GRID.TextMatrix(JJ&, 18) = NUMEROSERIE$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 19) = Cae$
            GRILLAGRAL.GRID.TextMatrix(JJ&, 20) = CODINOMENCLADOR$
            '
            NCLIEN_ant = NCLI
            NUCOMPLETO_ant$ = NUCOMPLETO$
            .MoveNext
       Loop
       'estas 2 lineas para totalizar el ultimo registro
       Call NEGRITA_SOLO_UNA_CELDA(GRILLAGRAL.GRID, JJ&, 15)
       ACU_IMPO_FACT = XVALO(VALOBADA("CAJABANC", "IDEBECOMP - IHABECOMP", "CodiCom_Lar = '" & NUCOMPLETO$ & "'", "NOMESS"))
       ACU_IMPO_FACT = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP - IHABECOMP)", "TIPOMOVIM = 'FVEN' AND CodiCom_Lar = '" & NUCOMPLETO$ & "' AND NUCLIEN = " & NCLI, "NOMESS"))
       GRILLAGRAL.GRID.TextMatrix(JJ&, 15) = FORMATNUM$(ACU_IMPO_FACT, "F12.2")

       JJ& = JJ& + 1
       GRILLAGRAL.GRID.Rows = JJ& + 2
       GRILLAGRAL.GRID.TextMatrix(JJ&, 14) = FORMATNUM$(ACUIMPO#, "F12.2")
       Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H8000000F, JJ&)

   End With
   Screen.MousePointer = 1
   TX$ = "Detalle de Ventas " & MESASEL$
   GRILLAGRAL.Caption = TX$
   '*******************************************
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
'   Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)

    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DETAVENTA", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing

End Sub









Function GeneraEstadisticaDeVentas(CodigoDesde$, CodigoHasta$, DesdeFecha%, HastaFecha%, Optional IncluyeTextoLibre%) As Boolean

        IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
        CondicionSqlCodigos$ = ""
        If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 And CodigoDesde$ = "" Then
            ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
            Dim RstCodigos As ADODB.Recordset
            Set RstCodigos = READSET(ConsultaSql$)
    
            With RstCodigos
                Do While Not .EOF
                    CodigoProducto$ = SAFETEXT$(!CAMPO2)
                    If CondicionSqlCodigos$ = "" Then
                        CondicionSqlCodigos$ = "m.codigo='" & CodigoProducto$ & "'"
                    Else
                        CondicionSqlCodigos$ = CondicionSqlCodigos$ & " or m.codigo='" & CodigoProducto$ & "'"
                    End If
    
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            RstCodigos.Close
            Set RstCodigos = Nothing
            On Error GoTo 0
        Else
            If CodigoDesde$ = "" Then
                racod$ = OBJPLA$("DESHASCOD", VDEF$)
                If racod$ = "" Then GoTo 99
                CI1% = CVI(racod$)
                CI2% = CVI(Mid$(racod$, 3))
                If CI1% = 0 Or CI2% = 0 Then GoTo 99
                CodigoDesde$ = CODEXT$(CI1%)
                CodigoHasta$ = CODEXT$(CI2%)
            End If
            If UCase(CodigoDesde$) <> "ALL" Then
                If CodigoDesde$ > CodigoHasta$ Then GoTo 99
            End If
        End If

        ConsultaSql$ = "CREATE VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select h.fecoscal, h.cod_inte from histocostos h inner join master m on (h.cod_inte=m.codint) "
        If CondicionSqlCodigos$ = "" Then
            If UCase(CodigoDesde$) <> "ALL" Then
                ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') h"
            Else
                ConsultaSql$ = ConsultaSql$ & ") h"
            End If
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") h"
        End If
        ConsultaSql$ = ConsultaSql$ & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        Dim rst_costos As New ADODB.Recordset
        On Error Resume Next
        FLEXCONN.Execute "DROP VIEW COSTOS"
        On Error GoTo 0
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        On Error GoTo 0

        ' ( (master inner join (cajabanc inner join movisto) ) left join costos
        ' con master se filtra por rango de codigos o por tabla de productos
        ' con el master inner join solo se trabaja con aquellos codigos que tienen movimientos.
                
        FiltroMaster$ = "inner"
        If UCase(CodigoDesde$) = "ALL" Then
            FiltroMaster$ = "right"
        End If
        If IncluyeTextoLibre% = 0 Then
            NoIncluyeCodigoCero$ = " and not (cod_inte=0 and valounit=0)"
        End If
        
        ConsultaSql$ = "create view estadistica as " _
        & " (select case when vendedor is null then 0 else vendedor end as vendedor, cliente, codinte, codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio, costo from " _
        & " (select vendedor, cliente, codinte, codexte, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo from" _
        & " (select vn.vendedor, vn.cliente, codinte, m.codigo codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from master m " & FiltroMaster$ & " join" _
        & " (select c.vend_cli vendedor, c.nuclien cliente, m.cod_inte codinte, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar and c.nuclien=m.nume_clie)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt' and c.numecomp>0 " & NoIncluyeCodigoCero$ _
        & " group by c.vend_cli, c.nuclien, m.cod_inte, month(c.fechconta), year(c.fechconta)) vn" _
        & " on (m.codint=vn.codinte)"
        If CondicionSqlCodigos$ = "" Then
            If UCase(CodigoDesde$) <> "ALL" Then
                ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') ilp_vn"
            Else
                ConsultaSql$ = ConsultaSql$ & ") ilp_vn"
            End If
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") ilp_vn"
        End If
        ConsultaSql$ = ConsultaSql$ & " left join costos" _
        & " on (ilp_vn.codinte=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv)"
        
        Text2 = ConsultaSql$
        Dim rst_estadisticas As New ADODB.Recordset
        On Error Resume Next
        FLEXCONN.Execute "DROP VIEW ESTADISTICA"
        On Error GoTo 0
        On Error GoTo ControlError
        rst_estadisticas.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_estadisticas = Nothing
        On Error GoTo 0
        
        GeneraEstadisticaDeVentas = True
        Exit Function
        
99:
    GeneralEstadisticaDeVentas = False
    Exit Function
    
ControlError:
    GeneraEstadisticaDeVentas = False
    Exit Function
    
End Function

Function NumeroCbte(NumeroComprobante$) As String

    NroComprobante$ = REPLACAR$(NumeroComprobante$, "-", "")
    NroComprobante$ = Mid$(NroComprobante$, 7, 8)
    
    Do While Len(NroComprobante$) < 20
        NroComprobante$ = "0" & NroComprobante$
    Loop
    
    NumeroCbte = NroComprobante$
    
End Function

Function PuntoVenta(NumeroComprobante$) As String

    PVenta$ = Mid$(NumeroComprobante$, 4, 4)
    
    PuntoVenta$ = "0" & PVenta$
    
End Function

Private Sub RegimenInformativoVentas()
    
    On Error Resume Next
    Call TRALOCAL("LIVAVEN", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "REGINFO_CV_VENTAS_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX1% = FILEOPEN%(LUCOM$ & "REGINFO_CV_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    
    Dim VGD_FechaComprobante$(3)
    Dim VGD_TipoCbte$(3)
        VGD_TipoCbte$(1) = TipoComprobante("FC-0000-00000000-B")
        VGD_TipoCbte$(2) = TipoComprobante("ND-0000-00000000-B")
        VGD_TipoCbte$(3) = TipoComprobante("NC-0000-00000000-B")
    Dim VGD_PuntoVta$(3)
    Dim VGD_NumeroCbtDesde$(3)
    Dim VGD_NumeroCbtHasta$(3)
    VGD_CodigoDocComprador$ = "99"
    VGD_NumeroIdentificacionComprador$ = RIVFormatoNumero$(0, 20)
    VGD_RazonSocial$ = AJUSTI$("VENTA GLOBAL DIARIA", 30)
    Dim VGD_ImporteTotal$(3)
    Dim VGD_ImporteNoGravado$(3)
    VGD_PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
    Dim VGD_ImporteOperacionExenta$(3)
    Dim VGD_ImportePercepcionesNacionales$(3)
    Dim VGD_ImportePercepcionIB$(3)
    VGD_ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
    VGD_ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
    VGD_CodMoneda$ = CodigoMoneda(1)
    VGD_TipoCambio$ = RIVFormatoNumero$(1, 10, 6)
    Dim VGD_CantidadAlicuotas$(3)
    VGD_CodigoOperacion$ = " "
    VGD_OtrosTributos$ = RIVFormatoNumero$(0, 15)
    VGD_FechaVencimiento$ = "00000000"
    Dim VGD_RegistroRIV$(3)
    Dim INDICE%
    
    Dim RegistroRIV_Alicuotas$
    
    Dim FlagError%
    Dim ListaDeErrores$
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVAVEN")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    For i& = 1 To FIN&
        Call TRACE(24, i&, ULTREG&("!LIVAVEN"))
        
        REGISTRO$ = REGLEIDO$("!LIVAVEN", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
        End If
        
        If NumeroComprobante$ <> UltimoNumeroComprobante$ Then
            '
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)
                       
            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    RegistroRIV_Alicuotas$ = ""
                End If
            End If
            FlagError% = 0
            
25
            
            fechaCbte$ = FETEXCOR1$(CVI(Mid$(REGISTRO$, 3, 2)))
            NumeroCliente& = XVALO(VALOBADA("CAJABANC", "NuClien", "(CODICOM_LAR='" & NumeroComprobante$ & "' OR CODICOM_LAR_EXSIMPLE='" & NumeroComprobante$ & "') AND FechConta='" & fechaCbte$ & "'"))
            If NumeroCliente& = 0 Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "Imposible Procesar el Movimiento. No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PUNTOVTA$ = PuntoVenta$(NumeroComprobante$)
            
            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
            
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroCliente&))
            
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
'            If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
'                NumeroIdentificacionComprador$ = CUITCLI$(NumeroCliente&)
'            End If
            If NumeroIdentificacionComprador$ = "" Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If
            PosicionIva% = PIVACLI%(NumeroCliente&)
            If PosicionIva% <> 5 And PosicionIva% <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
                End If
            End If
            NumeroIdentificacionComprador$ = SOLONUM$(NumeroIdentificacionComprador$)
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")
            
            razonSocial$ = AJUSTI$(rasocli$(NumeroCliente&), 30)
            If PIVACLI%(NumeroCliente&) = 0 And (TRIM$(razonSocial$) = "" Or InStr(1, razonSocial$, "CONSUMIDOR") > 0) Then
                razonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)), 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
            End If
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
            
            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
            
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(0, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)), 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                
                MONE% = XVALO(VALOBADA("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            CodMoneda$ = CodigoMoneda(MONE%)
            If TRIM$(CodMoneda$) = "000" Then CodMoneda$ = "PES"

                TCAMBIO# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            TIPOCAMBIO$ = RIVFormatoNumero$(TCAMBIO#, 10, 6)
            If CodMoneda$ = "PES" And (XVALO(TIPOCAMBIO$) / 1000000) > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
            End If
            If XVALO(TIPOCAMBIO$) = 0 And CodMoneda$ = "PES" Then TIPOCAMBIO$ = RIVFormatoNumero(1, 10, 6)
                
                
            ' Debe existir por lo menos una alicuota aunque sea no gravado o exento. No permite 0 (cero)
                CantAlicuotas% = 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
            
            If PIVACLI%(NumeroCliente&) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroCliente&) = 5 Then
                CodigoOperacion$ = "X"
            ElseIf Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroCliente&) = 4 Then
                CodigoOperacion$ = "E"
            Else
                CodigoOperacion$ = " "
            End If
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            ' Si es una Exportacion o Consumidor Final, es la misma que la fecha del comprobante.
            ' Caso contrario la calcula
            FechaVencimientoCbte$ = FechaComprobante$
            If PIVACLI%(NumeroCliente&) <> 5 And PIVACLI%(NumeroCliente&) <> 7 And PIVACLI%(NumeroCliente&) <> 0 Then
                FechaVencimientoCbte$ = FechaComprobante$ ' calcular!!!
            ElseIf PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                FechaVencimientoCbte$ = "00000000"
            End If
            
            ' Arma el registro para el regimen informativo
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                            & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                            & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                            & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
              
            UltimoNumeroComprobante$ = NumeroComprobante$
            
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            If InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                
                RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            Else
                ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuClien='" & CStr(NumeroCliente&) & "'"))
                If (ConIva# < 0.01 Or (XVALO(ImportePercepcionIB$) / 100) > 0.005) And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                    ImporteNeto$ = ImporteTotal$
                    If (XVALO(ImportePercepcionIB$) / 100) > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                    
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteNoGravado$, 124, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                    
                ElseIf (XVALO(ImporteOperacionExenta$) / 100) > 0.005 Then
                    
                    ImporteNeto$ = ImporteTotal$
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 154, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                End If
            End If
        Else
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 124, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8))
                If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
                End If
            End If
           
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 154, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
                          
            IMPO# = XVALO(Mid$(RegistroRIV$, 184, 15)) / 100
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)))
                               
                CantAlicuotas% = XVALO(CantidadAlicuotas$)
                If Mid$(REGISTRO$, 129, 8) <> "NO GRAV." Then CantAlicuotas% = CantAlicuotas% + 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
               
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
               
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)))
            If Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC") And ImporteIvaActual# > 0.005 Then ' InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 Then
                ImporteNeto$ = RIVFormatoNumero(ImporteNeto0#, 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                ImporteIva$ = RIVFormatoNumero(ImporteIvaActual#, 15)
                   
                RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            End If
        End If
50  Next i&
    
    If NumeroComprobante$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    Close #NX1%
        
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 109, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 124, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 139, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 154, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 169, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 184, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 199, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 214, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 244, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Percep. No Cat: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then SIGNO% = -1
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 29, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 48, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Ventas"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    Else
        Call COMUNI("Regimen Informativo de Ventas: Generación Ok")
    End If
    
99  Exit Sub
    
    
InicializarCamposVGD:
    
    VGD_FechaComprobante$(INDICE%) = "00000000"
    VGD_PuntoVta$(INDICE%) = ""
    VGD_NumeroCbtDesde$(INDICE%) = "0"
    VGD_NumeroCbtHasta$(INDICE%) = "0"
    VGD_ImporteTotal$(INDICE%) = "0"
    VGD_ImporteNoGravado$(INDICE%) = "0"
    VGD_ImporteOperacionExenta$(INDICE%) = "0"
    VGD_ImportePercepcionesNacionales$(INDICE%) = "0"
    VGD_ImportePercepcionIB$(INDICE%) = "0"
    VGD_CantidadAlicuotas$(INDICE%) = "0"
    VGD_RegistroRIV$(INDICE%) = ""

    Return
    
End Sub

Function CodigoAlicuota(Porcentaje$) As String

    Dim Codigo$
    If (XVALO(Porcentaje$) = 10.5) Then
        Codigo$ = "0004"
    ElseIf (XVALO(Porcentaje$) = 21) Then
        Codigo$ = "0005"
    ElseIf (Porcentaje$ = "NO GRAV.") Then
        Codigo$ = "0000"
    End If
    
    CodigoAlicuota = Codigo$
    
End Function


Function RIVFormatoNumero(NUMERO#, Optional Longitud%, Optional Decimales%, Optional ConsiderarSigno%) As String
    
    If Longitud <= 0 Then Longitud = 15
    If Decimales% <= 0 Then Decimales% = 2
    valor$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
    valor$ = REPLACAR$(valor$, ".", "")
    If valor$ = "" Then valor$ = "0"
    
    If ConsiderarSigno% = 0 Then
        Do While Len(valor$) < Longitud%
            valor$ = "0" & valor$
        Loop
    Else
        If NUMERO# >= 0 Then
            Do While Len(valor$) < Longitud%
                valor$ = "0" & valor$
            Loop
        Else
            'VALOR$ = Mid$(VALOR$, 2)
            Do While Len(valor$) < (Longitud% - 1)
                valor$ = "0" & valor$
            Loop
            valor$ = "-" & valor$
        End If
    End If
    
    RIVFormatoNumero = valor$
    
'    If Longitud <= 0 Then Longitud = 15
'    If Decimales% <= 0 Then Decimales% = 2
'    valor$ = TRIM$(FORMATNUM$(Abs(NUMERO#), "F" & CStr(Longitud%) & "." & Decimales%))
'    valor$ = REPLACAR$(valor$, ".", "")
'    If valor$ = "" Then valor$ = "0"
'
'    Do While Len(valor$) < Longitud%
'        valor$ = "0" & valor$
'    Loop
'
'    RIVFormatoNumero = valor$
    
End Function

Function TipoComprobante(NumeroComprobante$) As String

    IdCbte$ = Mid$(NumeroComprobante$, 1, 2) & Mid$(NumeroComprobante$, Len(NumeroComprobante$), 1)
    
    Tipo$ = ""
    Select Case IdCbte$
        Case "FCA": Tipo$ = "001"
        Case "NDA": Tipo$ = "002"
        Case "NCA": Tipo$ = "003"
        Case "FCB": Tipo$ = "006"
        Case "NDB": Tipo$ = "007"
        Case "NCB": Tipo$ = "008"
        Case "FCE": Tipo$ = "019"
        Case "NDE": Tipo$ = "020"
        Case "NCE": Tipo$ = "021"
    End Select
    
    TipoComprobante$ = Tipo$
    
End Function

Sub VentasPorProvincia(Opcion As Byte)

    If DERACCE%("VENPROV", "Ventas por Provincia") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    AnioElegido% = ListaDeAnios
    If AnioElegido% = 0 Then GoTo 99
    AnioElegido% = AnioElegido% Mod 100
    d$ = "01/01/" & CStr(AnioElegido%)
    H$ = "31/12/" & CStr(AnioElegido%)

    ' ***
    DesdeFecha% = FECHANUM(d$)
    HastaFecha% = FECHANUM(H$)
    If Opcion = 1 Then
        If GeneraEstadisticaDeVentas("", "", DesdeFecha%, HastaFecha%, 1) = False Then GoTo 99
    Else
        If GeneraEstadisticaDeVentas("all", "all", DesdeFecha%, HastaFecha%, 1) = False Then GoTo 99
    End If

    ' *** OBTIENE CANTIDAD DE PROVINCIAS
    Dim rstTotalProvincias As ADODB.Recordset
    ConsultaSql$ = "select count(provincias) cantprov from" _
    & " (select count(*) provincias from" _
    & " (select prov_cli, sum(ingreso) ingreso, sum(salida) salida, mes, anio from estadistica e with(nolock) inner join deudor12 d with(nolock) on (e.cliente=d.nume_cli) group by prov_cli, mes, anio) a group by prov_cli) b"
    Set rstTotalProvincias = READSET(ConsultaSql$)
    CantidadProvincias& = rstTotalProvincias!cantprov
    
    On Error Resume Next
    rstTotalProvincias.Close
    Set rstTotalProvincias = Nothing
    On Error GoTo 0
    
    ' *** DIBUJA EL INFORME
    GRILLAGRAL.GRID.Clear
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, "Cod/A3;Provincia/A24;Enero/F12.2;Febrero/F12.2;Marzo/F12.2;Abril/F12.2;Mayo/F12.2;Junio/F12.2;Julio/F12.2;Agosto/F12.2;Septiembre/F12.2;Octubre/F12.2;Noviembre/F12.2;Diciembre/F12.2;Tot. Provincia/F14.2", GRILLAGRAL, 1)
    GRILLAGRAL.GRID.FixedCols = 3
    GRILLAGRAL.GRID.Rows = CantidadProvincias& + 3
    GRILLAGRAL.GRID.TextMatrix(CantidadProvincias& + 2, 2) = "TOTALES"
    
    ' *** COLOREA INFORME
    With GRILLAGRAL
        For fi% = 1 To CantidadProvincias&
            For co% = 3 To 15
                Select Case co%
                Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(162, 210, 196), fi%, co%, 1)
                Case 15
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(201, 184, 143), fi%, co%, 1)
                End Select
                
                Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(207, 226, 243), CantidadProvincias& + 2, co%, 1)
            Next co%
        Next fi%
    End With
    
    ' *** OBTIENE CANTIDADES VENDIDAS POR PROVINCIA Y VUELCA EN LA GRILLA
    Dim rstCantidadPorProvincia As ADODB.Recordset
    ConsultaSql$ = "select prov_cli provincia, sum(ingreso) ingreso, sum(salida) salida, sum(valunit_salida) haber, sum(valunit_ingreso) debe, mes, anio from estadistica e with(nolock) inner join deudor12 d with(nolock) on (e.cliente=d.nume_cli) group by prov_cli, mes, anio"
    Set rstCantidadPorProvincia = READSET(ConsultaSql$)

    With rstCantidadPorProvincia
        ProximaProvincia% = 1
        Do While Not .EOF
            CodigoProvincia$ = SAFETEXT$(!Provincia)
            MES% = !MES
            If Opcion = 1 Then
                ImporteMes# = XVALO(!SALIDA) - XVALO(!INGRESO)
            ElseIf Opcion = 2 Then
                ImporteMes# = XVALO(!HABER) - XVALO(!DEBE)
            End If
            
            With GRILLAGRAL
                fi% = 0
                For i% = 1 To CantidadProvincias&
                    If .GRID.TextMatrix(i%, 1) = CodigoProvincia$ Then
                        fi% = i%
                        Exit For
                    End If
                Next i%
                If fi% = 0 Then
                    fi% = ProximaProvincia%
                    .GRID.TextMatrix(fi%, 1) = CodigoProvincia$
                    .GRID.TextMatrix(fi%, 2) = ECOARCH$("PROVINC", CodigoProvincia$)
                    ProximaProvincia% = ProximaProvincia% + 1
                End If
                co% = 2 + MES%
                
                Acumulado# = XVALO(.GRID.TextMatrix(fi%, co%))
                Acumulado# = Acumulado# + ImporteMes#
                .GRID.TextMatrix(fi%, co%) = TRIM$(FORMATNUM$(Acumulado#, "F12.2"))
            End With
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstCantidadPorProvincia.Close
    Set rstCantidadPorProvincia = Nothing
    On Error GoTo 0
    
    ' *** TOTALES POR PROVINCIA
    With GRILLAGRAL
        For fi% = 1 To CantidadProvincias&
            Acumulado# = 0
            For co% = 3 To 14
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next co%
            .GRID.TextMatrix(fi%, 15) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next fi%
    End With
    
    ' *** TOTALES POR MES
    With GRILLAGRAL
        For co% = 3 To 14
            Acumulado# = 0
            For fi% = 1 To CantidadProvincias&
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next fi%
            .GRID.TextMatrix(CantidadProvincias& + 2, co%) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next co%
    End With
    
    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    If Opcion = 1 Then
        GRILLAGRAL.Caption = "Ventas por Provincia (Cantidades)"
    ElseIf Opcion = 2 Then
        GRILLAGRAL.Caption = "Ventas por Provincia (Pesos)"
    End If
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub





Sub VentasPorCanales(Opcion As Byte)

    If DERACCE%("VENCANAL", "Ventas por Canales") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    AnioElegido% = ListaDeAnios
    If AnioElegido% = 0 Then GoTo 99
    AnioElegido% = AnioElegido% Mod 100
    d$ = "01/01/" & CStr(AnioElegido%): DesdeFecha% = FECHANUM(d$)
    H$ = "31/12/" & CStr(AnioElegido%): HastaFecha% = FECHANUM(H$)
    
    ' *** GENERA LA INFORMACION
    If GeneraEstadisticaDeVentas("all", "all", DesdeFecha%, HastaFecha%) = False Then GoTo 99

    ' *** MODO
    If Opcion = 2 Then
        IDLIST& = ObtenerIdTablaAuxiliar("ESTADIVENDE")
        ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
        Set rstTablaVendedores = READSET(ConsultaSql$)
        With rstTablaVendedores
            Do While Not .EOF
                CodigoVendedor$ = SAFETEXT$(!CAMPO2)
                
                If CondicionSql$ = "" Then
                    CondicionSql$ = "vendedor=" & CodigoVendedor$
                Else
                    CondicionSql$ = CondicionSql$ & " or vendedor=" & CodigoVendedor$
                End If
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rstTablaVendedores.Close
        Set rstTablaVendedores = Nothing
        On Error GoTo 0
    End If
    
    ' *** OBTIENE CANTIDAD POR VENDEDOR
    Dim rstTotalProvincias As ADODB.Recordset
    ConsultaSql$ = "select count(*) cantvendedores from" _
    & " (select vendedor from estadistica with(nolock)"
    If Opcion = 2 Then
        ConsultaSql$ = ConsultaSql$ & " where " & CondicionSql$
    End If
    ConsultaSql$ = ConsultaSql$ & " group by vendedor) vende"
    Set rstTotalVendedores = READSET(ConsultaSql$)
    CantidadVendedores& = rstTotalVendedores!cantvendedores
    
    On Error Resume Next
    rstTotalVendedores.Close
    Set rstTotalVendedores = Nothing
    On Error GoTo 0
    
    ' *** DIBUJA EL INFORME
    GRILLAGRAL.GRID.Clear
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, "Cod/A3;Vendedor/A24;Enero/F12.2;Febrero/F12.2;Marzo/F12.2;Abril/F12.2;Mayo/F12.2;Junio/F12.2;Julio/F12.2;Agosto/F12.2;Septiembre/F12.2;Octubre/F12.2;Noviembre/F12.2;Diciembre/F12.2;Tot. Canal/F14.2;% Canal/F14.2;Promedio/F14.2", GRILLAGRAL, 1)
    GRILLAGRAL.GRID.FixedCols = 3
    GRILLAGRAL.GRID.Rows = CantidadVendedores& + 3
    GRILLAGRAL.GRID.TextMatrix(CantidadVendedores& + 2, 2) = "TOTALES"
    
    ' *** COLOREA INFORME
    With GRILLAGRAL
        For fi% = 1 To CantidadVendedores&
            For co% = 3 To 15
                Select Case co%
                Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(162, 210, 196), fi%, co%, 1)
                Case 15
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(201, 184, 143), fi%, co%, 1)
                End Select
                
                Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(207, 226, 243), CantidadVendedores& + 2, co%, 1)
            Next co%
        Next fi%
    End With
    
    ' *** OBTIENE CANTIDADES VENDIDAS POR VENDEDOR Y VUELCA EN LA GRILLA
    Dim rstCantidadPorVendedor As ADODB.Recordset
    ConsultaSql$ = "select vendedor, sum(ingreso) ingreso, sum(salida) salida, sum(valunit_salida) haber, sum(valunit_ingreso) debe, mes, anio from estadistica e with(nolock)"
    If Opcion = 2 Then
        ConsultaSql$ = ConsultaSql$ & " where " & CondicionSql$
    End If
    ConsultaSql$ = ConsultaSql$ & " group by vendedor, mes, anio"
    Set rstCantidadPorVendedor = READSET(ConsultaSql$)

    With rstCantidadPorVendedor
        ProximoVendedor% = 1
        Do While Not .EOF
            CodigoVendedor$ = SAFETEXT$(!vendedor)
            MES% = !MES
            ImporteMes# = XVALO(!HABER) - XVALO(!DEBE)
            
            With GRILLAGRAL
                fi% = 0
                For i% = 1 To CantidadVendedores&
                    If .GRID.TextMatrix(i%, 1) = CodigoVendedor$ Then
                        fi% = i%
                        Exit For
                    End If
                Next i%
                If fi% = 0 Then
                    fi% = ProximoVendedor%
                    .GRID.TextMatrix(fi%, 1) = CodigoVendedor$
                    .GRID.TextMatrix(fi%, 2) = ECOARCH$("VENDEDOR", CodigoVendedor$)
                    ProximoVendedor% = ProximoVendedor% + 1
                End If
                co% = 2 + MES%
                
                Acumulado# = XVALO(.GRID.TextMatrix(fi%, co%))
                Acumulado# = Acumulado# + ImporteMes#
                .GRID.TextMatrix(fi%, co%) = TRIM$(FORMATNUM$(Acumulado#, "F12.2"))
            End With
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstCantidadPorVendedor.Close
    Set rstCantidadPorVendedor = Nothing
    On Error GoTo 0
    
    ' *** TOTALES POR VENDEDOR
    With GRILLAGRAL
        For fi% = 1 To CantidadVendedores&
            Acumulado# = 0
            For co% = 3 To 14
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next co%
            .GRID.TextMatrix(fi%, 15) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next fi%
    End With
    
    ' *** TOTALES POR MES
    With GRILLAGRAL
        For co% = 3 To 15
            Acumulado# = 0
            For fi% = 1 To CantidadVendedores&
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next fi%
            .GRID.TextMatrix(CantidadVendedores& + 2, co%) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next co%
    End With
    
    ' *** % POR VENDEDOR
    With GRILLAGRAL
        TotalTodosVendedores# = XVALO(.GRID.TextMatrix(CantidadVendedores& + 2, 15))
        If TotalTodosVendedores# > 0.005 Then
            For fi% = 1 To CantidadVendedores&
                Acumulado# = XVALO(.GRID.TextMatrix(fi%, 15))
                Acumulado# = (Acumulado# / TotalTodosVendedores#) * 100
                .GRID.TextMatrix(fi%, 16) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
            Next fi%
        End If
    End With
    
    ' *** PROMEDIO
    With GRILLAGRAL
        For fi% = 1 To CantidadVendedores&
            TotalVendedor# = XVALO(.GRID.TextMatrix(fi%, 15))
            CantidadMesesConVenta# = 0
            For co% = 3 To 14
                ImporteMes# = XVALO(.GRID.TextMatrix(fi%, co%))
                If ImporteMes# > 0.005 Then CantidadMesesConVenta# = CantidadMesesConVenta# + 1
            Next co%
            Promedio# = 0
            If CantidadMesesConVenta# > 0 Then Promedio# = (TotalVendedor# / CantidadMesesConVenta#)
            .GRID.TextMatrix(fi%, 17) = TRIM$(FORMATNUM$(Promedio#, "F14.2"))
        Next fi%
    End With

    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    If Opcion = 1 Then
        GRILLAGRAL.Caption = "Ventas por Canales (Todos)"
    ElseIf Opcion = 2 Then
        GRILLAGRAL.Caption = "Ventas por Canales (Solo Vendedores)"
    End If
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub


End Sub

Sub VentasCanalesPorProducto()

    If DERACCE%("VENCANAL", "Ventas por Canales") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    Call DESHASFECHA(DesdeFecha%, HastaFecha%, PERIODO$)
    
    ' *** GENERA LA INFORMACION
    If GeneraEstadisticaDeVentas("all", "all", DesdeFecha%, HastaFecha%) = False Then GoTo 99
    
    ' *** DIBUJA EL INFORME
    GRILLAGRAL.GRID.Clear
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, "Cod/A3;Vendedor/A24;Codigo/A18;Descripción/A48;Cantidad/F12.1", GRILLAGRAL, 1)
    GRILLAGRAL.GRID.FixedCols = 3
    
    ' *** OBTIENE CANTIDADES VENDIDAS POR VENDEDOR Y PRODUCTO Y VUELCA EN LA GRILLA
    Dim rstCantidadPorVendedor As ADODB.Recordset
    ConsultaSql$ = "select vendedor, codinte, sum(ingreso) ingreso, sum(salida) salida from estadistica e with(nolock)" _
    & " where codinte>0" _
    & " group by vendedor, codinte" _
    & " order by vendedor, codinte"
    Set rstCantidadPorVendedor = READSET(ConsultaSql$)

    With rstCantidadPorVendedor
        fi% = 0
        Do While Not .EOF
            CodigoVendedor$ = TRIM$(SAFETEXT$(!vendedor))
            If CodigoVendedor$ = "" Then CodigoVendedor$ = "0"
            CodigoInternoProducto% = XVALO(!codinte)
            CantidadVendida# = XVALO(!SALIDA) - XVALO(!INGRESO)
            
            fi% = fi% + 1
            With GRILLAGRAL
                .GRID.Rows = fi% + 1
                .GRID.TextMatrix(fi%, 1) = CodigoVendedor$
                .GRID.TextMatrix(fi%, 2) = ECOARCH$("VENDEDOR", CodigoVendedor$)
                .GRID.TextMatrix(fi%, 3) = CODEXT$(CodigoInternoProducto%)
                .GRID.TextMatrix(fi%, 4) = DESCRIT0$(CodigoInternoProducto%)
                .GRID.TextMatrix(fi%, 5) = TRIM$(FORMATNUM$(CantidadVendida#, "F12.1"))
            End With
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstCantidadPorVendedor.Close
    Set rstCantidadPorVendedor = Nothing
    On Error GoTo 0

    ' ********************
    GRILLAGRAL.GRID.MergeCells = flexMergeRestrictColumns
    GRILLAGRAL.GRID.MergeCol(1) = True
    GRILLAGRAL.GRID.MergeCol(2) = True
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Ventas por Canales y Productos (" & PERIODO$ & ")"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub


End Sub

Sub VentasProductosPorCanales()

    If DERACCE%("VENCANAL", "Ventas por Canales") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    Call DESHASFECHA(DesdeFecha%, HastaFecha%, PERIODO$)
    
    ' *** GENERA LA INFORMACION
    If GeneraEstadisticaDeVentas("", "", DesdeFecha%, HastaFecha%) = False Then GoTo 99
    
    ' *** DIBUJA EL INFORME
    GRILLAGRAL.GRID.Clear
    Campos$ = "CI/I0;Codigo/A18;Descripción/A48;0/I0;S-Vendedor/F24.1"
    
    IDLIST& = ObtenerIdTablaAuxiliar("ESTADIVENDE")
    ConsultaSql$ = "select campo2, campo3 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
    Set rstTablaVendedores = READSET(ConsultaSql$)
    With rstTablaVendedores
        Do While Not .EOF
            CodigoVendedor$ = SAFETEXT$(!CAMPO2)
            DenominacionVendedor$ = SAFETEXT$(!CAMPO3)
            Campos$ = Campos$ & ";" & CodigoVendedor$ & "/I0;" & DenominacionVendedor$ & "/F24.1"
            
            If CondicionSql$ = "" Then
                CondicionSql$ = "vendedor=" & CodigoVendedor$
            Else
                CondicionSql$ = CondicionSql$ & " or vendedor=" & CodigoVendedor$
            End If
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rstTablaVendedores.Close
    Set rstTablaVendedores = Nothing
    On Error GoTo 0
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    

    ' *** OBTIENE CANTIDADES VENDIDAS POR VENDEDOR Y PRODUCTO Y VUELCA EN LA GRILLA
    Dim rstCantidadPorVendedor As ADODB.Recordset
    ConsultaSql$ = "select codinte, vendedor, sum(ingreso) ingreso, sum(salida) salida from estadistica e with(nolock)" _
    & " where codinte>0 " & "and (" & CondicionSql$ & ")" _
    & " group by codinte, vendedor" _
    & " order by codinte, vendedor"
    Set rstCantidadPorVendedor = READSET(ConsultaSql$)

    With rstCantidadPorVendedor
        fi% = 0
        Do While Not .EOF
            CodigoVendedor$ = CStr(XVALO(!vendedor))
            If CodigoVendedor$ = "" Then CodigoVendedor$ = "0"
            CodigoInternoProducto% = XVALO(!codinte)
            CantidadVendida# = XVALO(!SALIDA) - XVALO(!INGRESO)

            With GRILLAGRAL
                ColumnaVendedor% = 0
                For VENDE% = 4 To .GRID.Cols - 1 Step 2
                    If CodigoVendedor$ = .GRID.TextMatrix(0, VENDE%) Then
                        ColumnaVendedor% = VENDE% + 1
                        Exit For
                    End If
                Next VENDE%
                
                If .GRID.Rows > 1 Then
                    fi% = 0
                    For filaArticulo% = 2 To .GRID.Rows - 1
                        If CStr(CodigoInternoProducto%) = .GRID.TextMatrix(filaArticulo%, 1) Then
                            fi% = filaArticulo%
                            Exit For
                        End If
                    Next filaArticulo%
                    If fi% = 0 Then
                        fi% = .GRID.Rows
                        .GRID.Rows = fi% + 1
                        
                        .GRID.TextMatrix(fi%, 1) = CStr(CodigoInternoProducto%)
                        .GRID.TextMatrix(fi%, 2) = CODEXT$(CodigoInternoProducto%)
                        .GRID.TextMatrix(fi%, 3) = DESCRIT0$(CodigoInternoProducto%)
                    End If
                Else
                    fi% = 2
                    .GRID.Rows = fi% + 1
                    
                    .GRID.TextMatrix(fi%, 1) = CStr(CodigoInternoProducto%)
                    .GRID.TextMatrix(fi%, 2) = CODEXT$(CodigoInternoProducto%)
                    .GRID.TextMatrix(fi%, 3) = DESCRIT0$(CodigoInternoProducto%)
                End If
                               
                .GRID.TextMatrix(fi%, ColumnaVendedor%) = TRIM$(FORMATNUM$(CantidadVendida#, "F12.1"))
            End With

50          .MoveNext
        Loop
    End With

    On Error Resume Next
    rstCantidadPorVendedor.Close
    Set rstCantidadPorVendedor = Nothing
    On Error GoTo 0

    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Ventas de Productos por Vendedores (" & PERIODO$ & ")"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub


End Sub


Private Sub AGenRep_Click()
    'VISDATA
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'VISDATA BOTON DERECHO
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command10_Click()
    Command10.Enabled = False
    'Call FORFASQL.INICIALIZACLIREAL ' INICIALIZA EL CAMPO AGREGADO DEL CLIENTE REAL
    
    If InStr(EMPREAC$, "DOSIVA") > 0 Then
       TTXX$ = "Ventas por Provincia\Informe Estadístico de Ventas"
       OPX% = ALTERNA%(TTXX$)
       If OPX% < 1 Or OPX% > 2 Then GoTo 99
       If OPX% = 1 Then
          Call VENTAPRODOSI
          GoTo 99
       ElseIf OPX% = 2 Then
          Call CONECRUN("ESTAVE07", "Informe Estadístico de Ventas")
          GoTo 99
       End If
       GoTo 99
    End If
    
    'jandy sql POR AHORA SE DEJA COMO ESTA
    'IDEAL TABLA PCIA, JOINS CON CLIENTE,PCIA,CAJABANC SALDOS EN LA SELECT
    ' IGESVE07
    '
    TTXX$ = "Ventas por Provincia y Artículo\Ventas por Provincia y Cliente"
    OPMAX% = 2
    If EXISTE%("IGESVE07.EXE") > 0 Then
       TTXX$ = TTXX$ + "\Informe Estadístico de Gestión"
       OPMAX% = OPMAX% + 1
       OPCIG% = OPMAX%
    End If
    If EXISTE%("ESTAVE07.EXE") > 0 Then
       TTXX$ = TTXX$ + "\Informe Estadístico de Ventas"
       OPMAX% = OPMAX% + 1
       OPCIV% = OPMAX%
    End If
    If InStr(EMPREAC$, "MEDITEA") > 0 Then
       TTXX$ = TTXX$ + "\Ventas por Vendedor"
       OPMAX% = OPMAX% + 1
       OPVEN% = OPMAX%
    End If
    '
    If InStr(EMPREAC$, "MEDITEA") > 0 Then
       TTXX$ = TTXX$ + "\Detalle de Ventas"
       OPMAX% = OPMAX% + 1
       OPDEVE% = OPMAX%
    End If
    '
    If InStr(EMPREAC$, "FERVI") > 0 Then
       TTXX$ = TTXX$ + "\Informe Ventas vs Costos"
       OPMAX% = OPMAX% + 1
       OPESTAVE% = OPMAX%
    End If
    '
    OPX% = ALTERNA%(TTXX$)
    If OPX% < 1 Or OPX% > OPMAX% Then GoTo 99
    If OPX% = 2 Then
         'Call VENPROCLI **AHORA SE LLAMA AL CREADO EN EL GENREP07**
         Call FINAL("?VENPORPROVIN")
         GoTo 99
       ElseIf OPX% = OPCIG% Then
         Call CONECRUN("IGESVE07", "Informe Estadístico de Gestión")
         GoTo 99
       ElseIf OPX% = OPCIV% Then
         Call CONECRUN("ESTAVE07", "Informe Estadístico de Ventas")
         GoTo 99
       ElseIf OPX% = OPVEN% Then
         Call FACTURAPORVENDEDOR
         GoTo 99
       ElseIf OPX% = OPDEVE% Then
        If InStr(EMPREAC$, "MEDITEA") > 0 Then
         OPX% = COMALTER%("Listado\\Reporte Tradicional\Reporte con Cae, N°.Serie, Nom. Merco Sur")
         If OPX% < 2 Then
            Call DETAVEN
         Else
           Call DETAVEN_MDT
         End If
         GoTo 99
       End If
       ElseIf OPX% = OPESTAVE% Then
         Call CONECRUN("IGESVE09", "Informe Ventas vs Costos")
         GoTo 99
    End If
    '
    Dim CPROVI$(), CIXX%(), TOVEN#(), GRUCO%()
    '
    difacu# = 0
    MODOTANIT% = 0
    If EXISTE%("LIVENTAN.PRF") > 0 Then MODOTANIT% = 1
    '
'1- Ingresar rango de fechas desde - hasta
     Call DESHASFECHA(Des%, Has%, PERIO$)
     If Des% < 1 Then GoTo 99
       DESX = FETEXCOR1$(Des%)
       HASX = FETEXCOR1$(Has%)
     CAREGIS% = 0
     '
'1'- Copiar la estructura del archivo GRUCOVEN.rfs
    Call COPYSTRU("GRUCOVEN", "GRUCOVE1")
'1''-Copio los registros del GRUCOVEN.DAT al GRUCOVE1.DAT
    YY$ = REGBLAN$("GRUCOVE1")
    U& = 0
    For J& = 1 To ULTREG&("GRUCOVEN")
      X$ = REGLEIDO("GRUCOVEN", J&)
      U& = U& + 1
      Call GRAREG("GRUCOVE1", X$, U&)
    Next J&
'1'''-Agregar un ultimo registro a GRUCOVE1.DAT
    YY$ = REGBLAN$("GRUCOVE1")
    Call REPLA(YY$, MKI(99), 1, 2)
    Call REPLA(YY$, "No Clasificado", 3, 30)
    Call GRAREG("GRUCOVE1", YY$, U& + 1)
    Call SETULTREG("GRUCOVE1", U& + 1)
    '
    If MODOTANIT% = 1 Then
       DESDAT = DESX
       HASDAT = HASX
       Call GENIVAVEN
    End If
    '
'2- Seleccionar de CajaBanc las facturas del período - similar a libro IVA
'   con el cliente obtener la provincia
    'jandy sql
    Call APERBASDAT("CABAN")
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechEmisi >= '" & DESX & "' "
    SQLX$ = SQLX$ + "AND FechEmisi <= '" & HASX & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'VENTA' or TipoMovim = 'NDCLI') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
24  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
    End If
    On Error GoTo 0
    '
    
    SQLY0$ = "SELECT * FROM CAJABANC "
    SQLY0$ = SQLY0$ + "where FechEmisi >= '" & DESX & "' "
    SQLY0$ = SQLY0$ + "AND FechEmisi <= '" & HASX & "' "
    SQLY0$ = SQLY0$ + "AND TipoMovim = 'EMBAL' "
    'Set SELEMBA = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
    '
    Dim SELEMBA As ADODB.Recordset
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      U& = 0
      .MoveFirst
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        Screen.MousePointer = 11
        If TRIM$(UCase$(SAFETEXT$(!OpciMovim))) <> "ANUL" Then
          NUMEFA$ = SAFETEXT$(!codicom_lar)
'          NCLIE% = !NuClien
          'TOMA EL CLIENTE REAL DEL MOVIMIENTO PARA HACER COINCIDIR CON EL DEL MOVISTO
          NCLI = XVALO(!Cli_Real) '
          
          ITOTA# = XVALO(!ihabecomp) - XVALO(!idebecomp)
          SIG% = 1: If ITOTA# < 0 Then SIG% = (-1)
          '
          
         ' AQUI OBTENGO EL IMPORTE EMBALAJE
         
          'SELEMBA.FindFirst "CodiCom_Lar = '" & NUMEFA$ & "'"
          'If SELEMBA.NoMatch = False Then
          
          'jandy sql SELEMBA
          IEMBAL# = 0
          SQLY$ = SQLY0$ + " AND CodiCom_Lar = '" & NUMEFA$ & "'"
          Set SELEMBA = FLEXCONN.Execute(FILTSQL$(SQLY$))

          'Set SELEMBA = FLEXCONN.Execute(SQLY$)
          If Not (SELEMBA.EOF And SELEMBA.BOF) Then
            IEMBAL# = XVALO(SELEMBA!ihabecomp) - XVALO(SELEMBA!idebecomp)
          End If
          Set SELEMBA = Nothing
          '
          'jandy sql fin SELEMBA
          
          ' LO QUE SIQUE ES PARA OBTENER EL NETO GRAVADO DE DÉBITOS LIBRES - TANIT
          If Left$(UCase$(NUMEFA$), 2) = "ND" Then
            If MODOTANIT% = 1 Then
              RFVT$ = FILTERGETRECORD$("LIVENTAN", 5, NUMEFA$)
              If TRIM$(Mid$(RFVT$, 11, 18)) = TRIM$(NUMEFA$) Then
                ITOTA# = SIG% * Abs(CVD(Mid$(RFVT$, 137, 8)))
              End If
            End If
          End If
          '
          MOBASI$ = SAFETEXT$(!TIPOMOVIM)
          COPROVIN$ = PROVCLI$(NCLI)
          If TRIM$(COPROVIN$) = "" Or ECOARCH$("PROVINC", COPROVIN$) = "" Then
            Call MENSERR(24, "Falta Codigo de Provincia\para Cliente Numero " + TRIM$(Str$(NCLI)) + ".")
          End If
          SUMAFA# = 0
          '
' COEFICIENTE DE PRECIOS
          Call APERBASDAT("STOCKS")
          
          SQLY$ = "SELECT * FROM MOVISTO "
          SQLY$ = SQLY$ + "where DoSecLar = '" & NUMEFA$ & "' "
          SQLY$ = SQLY$ + "AND (Nume_Clie = " & NCLI & ""
'             ORGD% = ORIGRUDI%(NCLIE%)
'             If ORGD% > 0 Then
'               If ORGD% <> NCLIE% Then
'                  SQLY$ = SQLY$ + " OR Nume_Clie = " & ORGD% & ""
'               End If
'             End If
          SQLY$ = SQLY$ + ") "
          SQLY$ = SQLY$ + "Order by NuorItem ASC "
          '
          'Set MOVISTEMP = Stocks.OpenRecordset(SQLY$, dbOpenSnapshot)
          Dim MOVISTEMP As ADODB.Recordset
          Set MOVISTEMP = New ADODB.Recordset
26        On Error Resume Next
          MOVISTEMP.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 26
          End If
          On Error GoTo 0
          '
          On Error Resume Next
          MOVISTEMP.MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 35 'Stop ' PROXIMA FACTURA
          End If
          On Error GoTo 0
          SUMATVEN# = 0
          Do While Not MOVISTEMP.EOF
            CIXI% = XVALO(MOVISTEMP!COD_INTE)
            CANTU = (Abs(XVALO(MOVISTEMP!Cantsalid) - XVALO(MOVISTEMP!CANTINGRE))) * SIG%
            PREUN# = XVALO(MOVISTEMP!VALOUNIT)
            SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            MOVISTEMP.MoveNext
          Loop
          COEFIMA# = 1
          If Abs(SUMATVEN#) > 0.005 Then
            COEFIMA# = Abs(ITOTA# / SUMATVEN#)
          End If
          '
'3.- para cada factura, con codicom_cor seleccionar los movimientos de stock correspondientes y obtener CI%, Preuni, cantidad y con eso el importe.
25        Call APERBASDAT("STOCKS")
          SQLY$ = "SELECT * FROM MOVISTO "
          SQLY$ = SQLY$ + "where DoSecLar = '" & NUMEFA$ & "' "
          SQLY$ = SQLY$ + "AND (Nume_Clie = " & NCLI & ""
'             ORGD% = ORIGRUDI%(NCLIE%)
'             If ORGD% > 0 Then
'               If ORGD% <> NCLIE% Then
'                  SQLY$ = SQLY$ + " OR Nume_Clie = " & ORGD% & ""
'               End If
'             End If
          SQLY$ = SQLY$ + ") "
          SQLY$ = SQLY$ + "Order by NuorItem ASC "
          '
          'Set MOVISTEMP = Stocks.OpenRecordset(SQLY$, dbOpenSnapshot)
          Set MOVISTEMP = New ADODB.Recordset
28        On Error Resume Next
          MOVISTEMP.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 28
          End If
          On Error GoTo 0
          '
          On Error Resume Next
          MOVISTEMP.MoveFirst
          If Err > 0 Then
            On Error GoTo 0
            GoTo 35 'Stop ' PROXIMA FACTURA
          End If
          On Error GoTo 0
          TDC$ = UCase$(Left$(SAFETEXT$(MOVISTEMP!DOSECLAR), 2))
          Do While Not MOVISTEMP.EOF
            CIXI% = XVALO(MOVISTEMP!COD_INTE)
            CANTU = (Abs(XVALO(MOVISTEMP!Cantsalid) - XVALO(MOVISTEMP!CANTINGRE))) * SIG%
            PREUN# = XVALO(MOVISTEMP!VALOUNIT) * COEFIMA#
            GRUMA% = GRUCOVEI%(CIXI%)
            If GRUMA% < 1 Then GRUMA% = 99
            SUMAFA# = SUMAFA# + PREUN# * CANTU
            '
            ' aqui eliminar notas de debito y credito de texto libre
'            If TDC$ = "NC" Or TDC$ = "ND" Then
'              If CIXI% < 1 Then
'                GoTo 32
'              End If
'            End If
'            '
'4.- Buscar en el vector de detalle coincidencia de provincia y articulo. si machea sumas cantidad e importe
'    Si no machea agregar un elemento al vector
            For KU% = CAREGIS% To 1 Step -1
              If CPROVI$(KU%) = COPROVIN$ Then
                If CIXX%(KU%) = CIXI% Then
                  TOVEN#(KU%) = TOVEN#(KU%) + PREUN# * CANTU
                  GoTo 32
                End If
              End If
            Next KU%
            CAREGIS% = CAREGIS% + 1
            ReDim Preserve CPROVI$(CAREGIS%), CIXX%(CAREGIS%), TOVEN#(CAREGIS%), GRUCO%(CAREGIS%)
            CPROVI$(CAREGIS%) = COPROVIN$
            CIXX%(CAREGIS%) = CIXI%
            TOVEN#(CAREGIS%) = PREUN# * CANTU
            GRUCO%(CAREGIS%) = GRUMA%
            '
'5.- repetir para los restantes items de la factura
32        MOVISTEMP.MoveNext
          Loop
35        DIFIMP# = IEMBAL# + ITOTA# - SUMAFA#
' MsgBox NUMEFA$ & "   " & ITOTA# & "   " & DIFIMP#
          If Abs(DIFIMP#) >= 0.005 Then
            CIXI% = 0
            PREUN# = DIFIMP#
            CANTU = 1
            GRUMA% = 99
            For KU% = CAREGIS% To 1 Step -1
              If CPROVI$(KU%) = COPROVIN$ Then
                If CIXX%(KU%) = CIXI% Then
                  TOVEN#(KU%) = TOVEN#(KU%) + (Int(100 * PREUN# * CANTU + 0.5)) / 100
                  GoTo 40
                End If
              End If
            Next KU%
            CAREGIS% = CAREGIS% + 1
            ReDim Preserve CPROVI$(CAREGIS%), CIXX%(CAREGIS%), TOVEN#(CAREGIS%), GRUCO%(CAREGIS%)
            CPROVI$(CAREGIS%) = COPROVIN$
            CIXX%(CAREGIS%) = CIXI%
            TOVEN#(CAREGIS%) = (Int(100 * PREUN# * CANTU + 0.5)) / 100
            GRUCO%(CAREGIS%) = GRUMA%
          End If
          '
        End If
'6.- siguiente factura
40    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
'7.- ordenar creciente por provincia y por tipo de material
'8.- generar listador
    '
    U& = 0
    YYY$ = REGBLAN("VEPROGRU")
    For J& = 1 To CAREGIS%
      Call REPLA(YYY$, CPROVI$(J&), 1, 1)
      Call REPLA(YYY$, MKI$(GRUCO%(J&)), 3, 2)
      Call REPLA(YYY$, MKI(CIXX%(J&)), 5, 2)
      Call REPLA(YYY$, MKD(TOVEN#(J&)), 7, 8)
      Call REPLA(YYY$, MKD(TOVEN#(J&)), 15, 8)
      U& = U& + 1
      Call GRAREG("VEPROGRU", YYY$, U&)
    Next J&
    '
    Call SETULTREG("VEPROGRU", U&)
    Call FILESORT("VEPROGRU", "VEPROGRU", 1, 5, "ASC")
    '
    REBLA$ = REGBLAN$("VEPROGRU")
    SUMAGRU# = 0: U2& = 0: XANT$ = REBLA$
    FIN& = ULTREG&("VEPROGRU")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("VEPROGRU", U&)
      Y$ = REBLA$
      If U& < FIN& Then
        Y$ = REGLEIDO$("VEPROGRU", U& + 1)
      End If
      '
      SUMAGRU# = SUMAGRU# + CVD(Mid$(X$, 7, 8))
      CORTECON% = 0
      '
      Call REPLA(X$, String$(16, 0), 23, 16)
      If Left$(Y$, 1) <> Left$(X$, 1) Then     'Or CVI(Mid$(Y$, 3, 2)) <> CVI(Mid$(X$, 3, 2)) Then ' CAMBIO PROVINCIA O GRUPO
          Call REPLA(X$, MKD$(SUMAGRU#), 23, 8)
          SUMAGRU# = 0
          UGRU% = CVI(Mid$(X$, 3, 2))
          CORTECON% = 1
      End If
      U2& = U2& + 1
      Call GRAREG("!VEPROGRU", X$, U2&)
    Next U&
    Call SETULTREG("!VEPROGRU", U2&)
    Call CIERRARCH("!VEPROGRU")
    Call TRACENTRAL("VEPROGRU", "")
    '
    Screen.MousePointer = 1
    Call FINAL("*VEPROGRU")
    '
99  Command10.Enabled = True
End Sub


Private Sub VENTAPRODOSI()
    '
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    CORRESP$ = MESANO(MESASEL$, Des%, Has%)
    '
    DES22$ = FETEXCOR1$(Des%)
    HAS22$ = FETEXCOR1$(Has%)
    '
    OPXX% = COMALTER%("Clasificación Según Tipo de Cliente:\\Usuarios\Comerciales\Todos\Cancelar")
    If OPXX% < 1 Or OPXX% > 3 Then Exit Sub
    '
    SQLX$ = "SELECT * FROM CAJABANC where FechConta >= '" + DES22$ + "' AND FechConta <= '" + HAS22$ + "' AND TipoMovim = 'VENTA' Order by FechEmisi ASC, NumeComp ASC"
    '
    Call APERBASDAT("CABAN")
    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = New ADODB.Recordset
    AVANCETMP.CursorLocation = adUseClient
25  On Error Resume Next
    AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With AVANCETMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Screen.MousePointer = 1
          Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
          Exit Sub
        End If
        JJ& = 0
        Do While Not .EOF
            NCLIE% = !NUCLIEN
            Select Case OPXX%
               Case 1: If CATIBRU$(NCLIE%) = "3" Then GoTo 19
               Case 2: If CATIBRU$(NCLIE%) <> "3" Then GoTo 19
            End Select
            CODPROV$ = PROVCLI$(NCLIE%)
            FECHAEMI% = CVINT(!FECHEMISI)
            NUMEFAC$ = SAFETEXT(!CodiCom_Cor)
            '
            If InStr(NUMEFAC$, "AJ") > 0 Then GoTo 19
            '
            CuitClien$ = CuitCli$(NCLIE%)
            IMPONETO# = !ihabecomp - !idebecomp
            '
            SQLX1$ = "SELECT * FROM CAJABANC WHERE CODICOM_LAR = '" + !codicom_lar + "' AND TIPOMOVIM='IVA'"
            Set EXENTOTMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
            If Not (EXENTOTMP.EOF And EXENTOTMP.BOF) Then
               XX$ = REGBLAN("COMULTI")
               Call REPLA(XX$, CODPROV$, 1, 1)
               Call REPLA(XX$, MKI$(FECHAEMI%), 2, 2)
               Call REPLA(XX$, NUMEFAC$, 4, 18)
               Call REPLA(XX$, MKI$(NCLIE%), 22, 2)
               Call REPLA(XX$, CuitClien$, 24, 16)
               Call REPLA(XX$, MKD$(IMPONETO#), 40, 8)
               '
               JJ& = JJ& + 1
               Call GRAREG("COMULTI", XX$, JJ&)
            Else
               XX$ = REGBLAN("VEPROVIX")
               Call REPLA(XX$, CODPROV$, 1, 1)
               Call REPLA(XX$, MKI$(FECHAEMI%), 2, 2)
               Call REPLA(XX$, NUMEFAC$, 4, 10)
               Call REPLA(XX$, MKI$(NCLIE%), 14, 2)
               Call REPLA(XX$, MKD$(IMPONETO#), 16, 8)
               '
               II& = II& + 1
               Call GRAREG("VEPROVIX", XX$, II&)
            End If
19       .MoveNext
         Loop
    End With
    '
    ENCABE$ = "Clientes: NO COMERCIALES"
    Select Case OPXX%
       Case 2: ENCABE$ = "Clientes: COMERCIALES"
       Case Else: ENCABE$ = "Clientes: TODOS"
    End Select
    '
    If JJ& > 0 Then
      Call SETULTREG("COMULTI", JJ&)
      Call CIERRARCH("COMULTI")
      Call FILESORT("COMULTI", "", 1, 1, "ASC")
      Call HEADERS("COMULTI", "")
      Call HEADERS("COMULTI", "Corresponde a: " + TRIM$(MESASEL$) + " - Generado " + FECHATEX(HOY(HO$)))
      Call FINAL("*COMULTI")
    End If
    '
    If II& > 0 Then
      Call SETULTREG("VEPROVIX", II&)
      Call CIERRARCH("VEPROVIX")
      Call FILESORT("VEPROVIX", "", 1, 1, "ASC")
      Call HEADERS("VEPROVIX", "")
      Call HEADERS("VEPROVIX", "Corresponde a: " + TRIM$(MESASEL$) + " - Generado " + FECHATEX(HOY(HO$)))
      Call FINAL("*VEPROLX")
    End If
End Sub


Sub VENPROCLI()

'FALTA COMPLETAR

    Dim TOTVENTA#(32767)
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where FechEmisi >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechEmisi <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'VENTA' or TipoMovim = 'NDCLI') "
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    Screen.MousePointer = 11
    CLIMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         NCLIE% = !NUCLIEN
         If NCLIE% > 0 Then
           NUMEFA$ = !codicom_lar
           If NCLIE% > CLIMAX% Then CLIMAX% = NCLIE%
           '
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + (!ihabecomp - !idebecomp)
           ' agrega EL IMPORTE EMBALAJE
           IEMBAL# = 0
           'SELEMBA.FindFirst "CodiCom_Lar = '" & NUMEFA$ & "'"
           SQLY$ = "SELECT * FROM CAJABANC "
           SQLY$ = SQLY$ + "where FechEmisi >= '" & DESDAT & "' "
           SQLY$ = SQLY$ + "AND FechEmisi <= '" & HASDAT & "' "
           SQLY$ = SQLY$ + "AND TipoMovim = 'EMBAL' "
           SQLY$ = SQLY$ + " AND CodiCom_Lar = '" & NUMEFA$ & "'"
           'Set SELEMBA = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
           Dim SELEMBA As ADODB.Recordset
           Set SELEMBA = FLEXCONN.Execute(FILTSQL$(SQLY$))

           'If SELEMBA.NoMatch = False Then
           On Error Resume Next
           With SELEMBA
             .MoveFirst
             If Err < 1 Then
               On Error GoTo 0
               'IEMBAL# = SELEMBA!IHABECOMP - SELEMBA!IDEBECOMP
               IEMBAL# = !ihabecomp - !idebecomp
             End If
             On Error GoTo 0
           End With
           Set SELEMBA = Nothing
           TOTVENTA#(NCLIE%) = TOTVENTA#(NCLIE%) + IEMBAL#
           '
         End If
35    .MoveNext
      Loop
    End With
    Set CABATEMP = Nothing
    '
    JJ& = 0
    REBLA$ = REGBLAN$("VEPROVCL")
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          NCLIE% = NCI%
          COPROVIN$ = PROVCLI$(NCLIE%)
          XX$ = REBLA$
          Call REPLA(XX$, COPROVIN$, 1, 2)
          Call REPLA(XX$, MKI$(NCLIE%), 3, 2)
          Call REPLA(XX$, MKD$(TOTVENTA#(NCI%)), 5, 8)
          JJ& = JJ& + 1
          Call GRAREG("VEPROVCL", XX, JJ&)
       End If
    Next NCI%
    Call SETULTREG("VEPROVCL", JJ&)
    '
    Call FILESORT("VEPROVCL", "", 1, 3)
    Call HEADERS("VEPROVCL", "")
    Call HEADERS("VEPROVCL", "Corresponde a: " + PERIO$)
    Call FINAL("*VEPROVCL")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")

End Sub
'
Sub Command11_Click()
    Command11.Enabled = False
    
    'Call FORFASQL.INICIALIZACLIREAL ' INICIALIZA EL CAMPO AGREGADO DEL CLIENTE REAL

'\\\OT-5-0384-WAR-07/06/05///
    Dim TOTVENTA#(32767)
'\\\OT-06-0494-RG-20/11/06///
    Dim TOTCANT#(32767)
'\\\OT-06-0494-RG-FIN///
    '
10  Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    '\\\OT-06-0494-RG-20/11/06///
    opt% = COMALTER%("Opciones de Formato de Salida\\Ranking de Ventas\Ranking de Margen Bruto")
    If opt% = 0 Then GoTo 99
    '
    COSTCALC% = 0
    If opt% = 2 Then
       OPT2% = COMALTER%("Opciones \\Costo de Reposición\Costo Calculado")
       If OPT2% = 2 Then COSTCALC% = 1
    End If
    
    '\\\OT-06-0494-RG-FIN///
    
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
    Next U&
    '
    'DESDAT = FECHATEX(DES%)
    'HASDAT = FECHATEX(HAS%)
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    '
    Call APERBASDAT("CABAN")
    'jandy SQL (No seria mejor un join entre cajabanc y movisto?)
    
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(XVALO(!ihabecomp) - XVALO(!idebecomp), 2)
'       NCLI = XVALO(!NuClien) TOMA EL CLIENTE REAL DEL MOVIMIENTO PARA HACER COINCIDIR
        NCLI = XVALO(!Cli_Real) ' CON EL DEL MOVISTO
        If Abs(TOTFAC#) >= 0.005 Then
          NROFAC$ = SAFETEXT$(!codicom_lar)
          SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
          SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
          SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLI & " "
'             ORGD% = ORIGRUDI%(NCLI)
'             If ORGD% > 0 Then
'               If ORGD% <> NCLI Then
'                  SQLX$ = SQLX$ + " OR Nume_Clie = " & ORGD% & ""
'               End If
'             End If
          SQLX$ = SQLX$ + ") "
          SQLX$ = SQLX$ + "Order by NuorItem ASC "
          '
          'Set STOTEMP1 = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
          Dim STOTEMP1 As ADODB.Recordset
          Set STOTEMP1 = New ADODB.Recordset
          STOTEMP1.CursorLocation = adUseClient
26        On Error Resume Next
          STOTEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
          End If
          On Error GoTo 0
          '
          With STOTEMP1
            On Error Resume Next
            .MoveFirst
            If Err > 0 Then
               TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC#
               On Error GoTo 0
               GoTo 30
            End If
            On Error GoTo 0 '
            '
            'COEFICIENTE DE PRECIOS
            SUMATVEN# = 0
            Do While Not .EOF
              CIXI% = XVALO(!COD_INTE)
              CANTU = (XVALO(!Cantsalid) - XVALO(!CANTINGRE))
              PREUN# = XVALO(!VALOUNIT)
              If XVALO(!MoneVal) > 1 Then
                PREUN# = PREUN# * XVALO(!TIPOCAM)
              End If
              SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            .MoveNext
            Loop
            COEFIMA# = 1
            If Abs(SUMATVEN#) > 0.005 Then
              COEFIMA# = Abs(TOTFAC# / SUMATVEN#)
            End If
            '
            .MoveFirst
            Do While Not .EOF
              NART% = XVALO(!COD_INTE): If NART% < 0 Then NART% = 0
              If NART% > ARTMAX% Then ARTMAX% = NART%
              '\\\OT-06-0494-RG-20/11/06///
              CANTI# = (XVALO(!Cantsalid) - XVALO(!CANTINGRE))
              TOTCANT#(NART%) = TOTCANT#(NART%) + CANTI#
              '\\\OT-06-0494-RG-FIN///
              PREUN# = XVALO(!VALOUNIT)
              NETO# = (XVALO(!Cantsalid) - XVALO(!CANTINGRE)) * PREUN# * COEFIMA#
              If XVALO(!MoneVal) > 1 Then
                NETO# = NETO# * XVALO(!TIPOCAM)
              End If
              NETO# = ROUND(NETO#, 2)
              TOTVENTA#(NART%) = TOTVENTA#(NART%) + NETO#
              TOTFAC# = TOTFAC# - NETO#
            .MoveNext
            Loop
If Abs(TOTFAC#) >= 0.005 Then
  AA1 = NROFAC$
  aa2 = NCLI
  aa3 = aa2
End If
            TOTVENTA#(0) = TOTVENTA#(0) + TOTFAC#
          End With                    ' SOBRE MOVISTO
          STOTEMP1.Close
          Set STOTEMP1 = Nothing
        End If
        '
30    .MoveNext
      Loop
      '
    End With                  ' SOBLE CAJABANC
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 0 To ARTMAX%
      TOTAGEN# = TOTAGEN# + TOTVENTA#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 0 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Or (Abs(TOTVENTA#(NCI%)) < 0.005 And TOTCANT#(NCI%) > 0) Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        W$ = REGBLAN$("RANMABRU")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call REPLA(Z$, MKD$(TOTCANT#(NC%)), 21, 8) '10/05/07 (OT-07-0188)
        TOTBRUTO# = TONEVEN# * (1 + PORCEIVA(NC%) / 100)
        Call REPLA(Z$, MKD$(TOTBRUTO#), 29, 8)
        Call GRAREG("RANKPIE", Z$, J&)
        '\\\OT-06-0494-RG-20/11/06///
        Call REPLA(W$, MKI$(JI%), 1, 2)
        Call REPLA(W$, MKI$(NC%), 3, 2)
        Call REPLA(W$, MKD$(TOTCANT#(NC%)), 5, 8)
        Call REPLA(W$, MKD$(TONEVEN#), 13, 8)
        If COSTCALC% > 0 Then
          CU# = COSCALC#(NC%) ' COSTO CALCULADO
        Else
          CU# = COSPESOS(NC%) ' COSTO REPOSICION
        End If
        Call REPLA(W$, MKD$(CU#), 21, 8)
        ct# = CU# * TOTCANT#(NC%)
        Call REPLA(W$, MKD$(ct#), 29, 8)
        UB# = TONEVEN# - ct#
        Call REPLA(W$, MKD$(UB#), 37, 8)
        Call REPLA(W$, MKS$(PORPAR), 45, 4)
        Call REPLA(W$, MKS$(PORACU), 49, 4)
        '\\\OT-07-0014-RG-24/1/07///
        IMPIVA# = (TONEVEN# * PORCEIVA(NC%)) / 100
        Call REPLA(W$, MKD$(IMPIVA#), 53, 8)
        ITOVE# = TONEVEN# + IMPIVA#
        Call REPLA(W$, MKD$(ITOVE#), 61, 8)
        '\\\OT-07-0014-RG-FIN///
        Call GRAREG("RANMABRU", W$, J&)
        '\\\OT-06-0494-RG-FIN///
    Next U&
    '
    
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    '\\\OT-06-0494-RG-20/11/06///
    Call SETULTREG("RANMABRU", J&)
    Call CIERRARCH("RANMABRU")
    Call HEADERS("RANMABRU", "")
    Call HEADERS("RANMABRU", "Periodo: " + PERIO$)
    '\\\OT-06-0494-RG-FIN///
    Screen.MousePointer = 1
    List1.Clear
    '
    
    If opt% = 1 Then
      If Option1.Value = True Then
         Call FINAL("*RANKPIE")
      Else
         Call FINAL("*LRANPIE")
      End If
      '\\\OT-06-0494-RG-20/11/06///
      ElseIf opt% = 2 Then
          If Option1.Value = True Then
             Call FINAL("*RANMABRU")
          Else
             Call COMUNI("Solamente válido para impresión")
             Call FINAL("*RANMABRU")
          End If
      '\\\OT-06-0494-RG-FIN///
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command11.Enabled = True
    '
    '\\\OT-5-0384-WAR-07/06/05///
End Sub

Private Sub Command12_Click()
    '
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
    End If
    Call FRESHALL
    '
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENIVAVEN
    FIN& = ULTREG&("REIVAVEN")
    If FIN& < 1 Then GoTo 99
    '
    OPX% = COMALTER%("Filtrado de Salida:\\Totales\Local\Com.Exterior")
    If OPX% = 1 Then
        Call FINAL("*SOREIVEN")
      ElseIf OPX% = 2 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVAVEN", U&)
          If InStr(XXX$, "COM.EXT") < 1 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVAVEN", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVAVEN", JJ&)
        Call FINAL("*SOREIVEN")
      ElseIf OPX% = 3 Then
        JJ& = 0
        For U& = 1 To FIN&
          XXX$ = REGLEIDO$("REIVAVEN", U&)
          If InStr(XXX$, "COM.EXT") > 0 Then
            JJ& = JJ& + 1
            Call GRAREG("REIVAVEN", XXX$, JJ&)
          End If
        Next U&
        Call SETULTREG("REIVAVEN", JJ&)
        Call FINAL("*SOREIVEN")
    End If
    '
99  Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If DERACCE%("STPREMEV", "Setup Resumen Mensual de Ventas") = 2 Then
        '\\\OT-06-0558-RG-21/12/06///
        opt% = ALTERNA%("Imputación Contable\Libro de IVA")
        If opt% = 1 Then
            Call COGEVE07.Command18_Click
          ElseIf opt% = 2 Then
            FOLIVAVE07.Show 1
        End If
        '\\\OT-06-0558-RG-FIN///
   End If
99 Command14.Enabled = True
End Sub




 Sub Command15_Click()
    '
    Command15.Enabled = False
    
    'Call FORFASQL.INICIALIZACLIREAL ' INICIALIZA EL CAMPO AGREGADO DEL CLIENTE REAL
    '
    Dim TOTVENTA#(32767), TOTCANT#(32767)
    '
10  Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
15  OPX% = COMALTER%("Opciones de Trabajo:\\Por Cliente\Por Articulo\Por Fecha")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    '
    If OPX% = 2 Then 'Si la opcion elegida es Por Articulo, permite filtrar por codigo de articulo
        VDEF$ = MKI$(PRICOD%) + MKI$(ULTCOD%) + Chr$(0)
20      racod$ = OBJPLA$("DESHASCOD", VDEF$)
        If racod$ = "" Then
            GoTo 15
        End If
        CI1% = CVI(Left$(racod$, 2)): COD1$ = CODEXT$(CI1%)
        CI2% = CVI(Mid$(racod$, 3, 2)): COD2$ = CODEXT$(CI2%)
        If CI1% = 0 Or CI2% = 0 Then GoTo 20
    End If
    
    FiltroMovimientos% = COMALTER%("Tipos de Movimiento:\\Solo Facturas\Incluye Además Notas de Débito y Crédito")
    
25  Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
    Next U&
    '
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    REBLA$ = REGBLAN$("DETFACPE")
    '
    Call APERBASDAT("CABAN")
    '
    'VALIDAMOS QUE EXISTAN REGISTROS EN PERIODO ELEGIDO
    CONDI$ = " FechConta >= '" & DESDAT & "' AND FechConta <= '" & HASDAT & "' "
    CONDI$ = CONDI$ + "AND TipoMovim = 'VENTA' "
    '
    FIN& = CantRegistros&("CAJABANC", CONDI$)
    If FIN& < 1 Then
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
    End If
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE " & CONDI$
    '
    Dim STOTEMP1 As New ADODB.Recordset
    Dim CABATEMP  As New ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With CABATEMP
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(XVALO(!ihabecomp) - XVALO(!idebecomp), 2)
'       NCLIE = XVALO(!NuClien) TOMA EL CLIENTE REAL DEL MOVIMIENTO PARA HACER COINCIDIR
        NCLIE = XVALO(!Cli_Real) ' CON EL DEL MOVISTO        If Abs(TOTFAC#) >= 0.005 Then
        If Abs(TOTFAC#) >= 0.005 Then
          FF% = CVINT(!fechconta)
          NROFAC$ = SAFETEXT$(!codicom_lar)
          '
          If (FiltroMovimientos% = 0 Or FiltroMovimientos% = 1) And Left$(NROFAC$, 2) <> "FC" Then GoTo 30 ' solo facturas
          '
          SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
          SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
          SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLIE & " "
''             ORGD% = ORIGRUDI%(NCLIE) COMENTADO POR QUE TOMA EL CLI_REAL QUE COINCIDE
''             If ORGD% > 0 Then         CON EL DEL MOVISTO
''               If ORGD% <> NCLIE Then
''                  SQLX$ = SQLX$ + " OR Nume_Clie = " & ORGD% & ""
''               End If
''             End If
          SQLX$ = SQLX$ + ") "
          SQLX$ = SQLX$ + "Order by NuorItem ASC "
          '
          Set STOTEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
          '
          With STOTEMP1
            On Error Resume Next
'            .MoveFirst
            If .EOF = True Or .BOF = True Then
               On Error GoTo 0
               NPED& = 0
               NART% = 0
               CANTU = 0
               PRUNI# = 0
               NETO# = TOTFAC#
               GoSub AGREDET
               GoTo 30
            End If
            On Error GoTo 0 '
            '
            'COEFICIENTE DE PRECIOS
            SUMATVEN# = 0
            Do While Not .EOF
              CIXI% = !COD_INTE
              CANTU = (!Cantsalid - !CANTINGRE)
              PREUN# = !VALOUNIT
              If XVALO(!MoneVal) > 1 Then
                PREUN# = PREUN# * XVALO(!TIPOCAM)
              End If
              SUMATVEN# = SUMATVEN# + PREUN# * CANTU
            .MoveNext
            Loop
            COEFIMA# = 1
            If Abs(SUMATVEN#) > 0.005 Then
              COEFIMA# = Abs(TOTFAC# / SUMATVEN#)
            End If
            '
            .MoveFirst
            Do While Not .EOF
              NART% = !COD_INTE: If NART% < 0 Then NART% = 0
              If NART% > ARTMAX% Then ARTMAX% = NART%
              NPED& = !NUPEDCLI
              CANTU = (!Cantsalid - !CANTINGRE)
              PRUNI# = !VALOUNIT * COEFIMA#
                If !MoneVal > 1 Then PRUNI# = PRUNI# * !TIPOCAM
              NETO# = ROUND(CANTU * PRUNI#)
              TOTFAC# = TOTFAC# - NETO#
              GoSub AGREDET
            .MoveNext
            Loop
            If Abs(TOTFAC#) >= 0.005 Then
               NPED& = 0
               NART% = 0
               CANTU = 0
               PRUNI# = 0
               NETO# = TOTFAC#
               GoSub AGREDET
            End If
          End With                    ' SOBRE MOVISTO
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    '
    Call SETULTREG("DETFACPE", J&)
    Call CIERRARCH("DETFACPE")
    Call HEADERS("DETFACPE", "")
    Call HEADERS("DETFACPE", "Periodo: " + PERIO$)
    If OPX% = 1 Then
        Call FILESORT("DETFACPE", "", 1, 4, "ASC")  'ORDENADO POR CLIENTE Y Nº FACTURA
    ElseIf OPX% = 3 Then
        Call FILESORT("DETFACPE", "", 3, 4, "ASC")  'ORDENADO POR FECHA Y Nº FACTURA
    End If
    
    '
    Call SETULTREG("DETFACOD", i&)
    Call CIERRARCH("DETFACOD")
    Call HEADERS("DETFACOD", "")
    Call HEADERS("DETFACOD", "Periodo: " + PERIO$)
    Call HEADERS("DETFACOD", "Desde Codigo: " + CODEXT$(CI1%) + " - " + DESCRIT$(CI1%))
    Call HEADERS("DETFACOD", "Hasta Codigo: " + CODEXT$(CI2%) + " - " + DESCRIT$(CI2%))
    Call FILESORT("DETFACOD", "", 3, 3, "ASC")
    Call FILESORT("DETFACOD", "", 1, 1, "ASC")
    '
    Screen.MousePointer = 1
    '
    If OPX% = 1 Then
        Call FINAL("*DETFACPE")     'ORDENADO POR CLIENTE Y Nº FACTURA
    ElseIf OPX% = 2 Then
        Call FINAL("*DETFACOD")
    ElseIf OPX% = 3 Then
        Call FINAL("*DETFACPE")     'ORDENADO POR FECHA Y Nº FACTURA
    End If
    GoTo 99
    '
AGREDET:
    'DETALLE DE FACTURACION POR CLIENTE
    Z$ = REBLA$
    Call REPLA(Z$, MKI$(NCLIE), 1, 2)
    Call REPLA(Z$, MKI$(FF%), 3, 2)
    Call REPLA(Z$, NROFAC$, 5, 18)
    Call REPLA(Z$, NROREM$, 23, 18)
    Call REPLA(Z$, MKS$(NPED&), 41, 4)
    Call REPLA(Z$, MKI$(NART%), 45, 2) '10/05/07 (OT-07-0188)
    Call REPLA(Z$, MKS$(CANTU), 47, 4)
    Call REPLA(Z$, MKD$(PRUNI#), 51, 8)
    Call REPLA(Z$, MKD$(NETO#), 59, 8)
    J& = J& + 1
    Call GRAREG("DETFACPE", Z$, J&)
    '
    'DETALLE DE FACTURACION POR ARTICULO
    If CODEXT$(NART%) >= COD1$ Then
        If CODEXT$(NART%) <= COD2$ Then
            XZ$ = REBLA$
            Call REPLA(XZ$, MKI$(NART%), 1, 2)
            Call REPLA(XZ$, MKI$(NCLIE), 3, 2)
            Call REPLA(XZ$, MKS$(CANTU), 5, 4)
            Call REPLA(XZ$, MKD$(PRUNI#), 9, 8)
            Call REPLA(XZ$, MKD$(NETO#), 17, 8)
            i& = i& + 1
            Call GRAREG("DETFACOD", XZ$, i&)
        End If
    End If
    
Return
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    STOTEMP1.Close
    Set STOTEMP1 = Nothing
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command15.Enabled = True
End Sub



Private Sub Command16_Click()
    Command16.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Option2.Value = True
    Call GENSUDIVEN
    Call FINAL("*SUDIVEN")
    '
99  Command16.Enabled = True
End Sub







Sub Command19_Click()
    Command19.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    Call GENSUDIVEN
    If COPYFILE%(LUDAT$ + "SUDIVEN.DAT", LUDAT$ + "COSUDIV.DAT") > 0 Then
      Call FILESORT("COSUDIV", "", 1, 3, "ASC")
      If Option1.Value = True Then
          Call CONECRUN("*COSUDIV", "Consulta por Cuenta Contable")
          GoTo 99
        ElseIf Option2.Value = True Then
          Call FILESORT("COSUDIV", "", 6, 6, "ASC")
          Call CONECRUN("*LISUDIV", "Listado de Ventas por Cuenta Contable")
          GoTo 99
      End If
    End If
99  Command19.Enabled = True
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    If VALIMESAN% < 1 Then GoTo 99
    '
    MOX% = ALTERNA%("Secuencia por Recibo\Detalle por Cliente")
    If MOX% < 1 Or MOX% > 2 Then
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    'Jandy SQL
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where  FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "Order by FechConta, NUMECOMP, nuclien ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP  As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    
    J1& = 0: J2& = 0
    REBLA1$ = REGBLAN$("DETACCX")
    REBLA1$ = REGBLAN$("RECOBCLI")
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        ZZ$ = REBLA1$
        NCLI = !NUCLIEN
        FECHAD = Int(CDbl(!fechconta))
        INETO# = !ValAbsComp
        ITOTA# = !ValAbsComp
        '
        Call REPLA(ZZ$, Chr$(4), 1, 1)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
        Call REPLA(ZZ$, MKS$(NCLI), 9, 4)
        Call REPLA(ZZ$, MKD$(INETO#), 13, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 53, 8)
        J1& = J1& + 1
        Call GRAREG("DETACCX", ZZ$, J1&)
        '
        ZZ$ = REBLA2$
        Call REPLA(ZZ$, MKS$(!NUCLIEN), 1, 4)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 5, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
        Call REPLA(ZZ$, MKD$(INETO#), 27, 8)
        J2& = J2& + 1
        Call GRAREG("RECOBCLI", ZZ$, J2&)
        '
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    '
    Call SETULTREG("DETACCX", J1&)
    Call SETULTREG("RECOBCLI", J2&)
    '
    Screen.MousePointer = 1
    On MOX% GoTo 203, 204
    GoTo 99
    '
203 Screen.MousePointer = 11
    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
    Call HEADERS("DETACCX", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
    GoTo 99
    '
204 Call HEADERS("RECOBCLI", "")
    Call HEADERS("RECOBCLI", "Corresponde a: " + TRIM$(MESAN))
    Call HEADERS("RECOBCLI", "Rubro: Recibos de Cobranza")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*COCOBCL")
      ElseIf Option2.Value = True Then
         Call FINAL("*SOCOBCL")
    End If
    GoTo 99
    '
99 Command2.Enabled = True
   Screen.MousePointer = 1
End Sub
'

Sub Command20_Click()
   '
   Command20.Enabled = False
   '
   If DERACCE%("VERASIVE", "Visualizacion de Asiento de Ventas") < 2 Then
      GoTo 99
   End If
   '
   If VALIMESAN% < 1 Then GoTo 99
   '
   Call GENSUDIVEN
   '
   JJ& = 0
   TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HASTA%)
   FIN& = ULTREG&("ASIDIAR")
   '
   For IL& = 1 To FIN&
     Call TRACE(20, IL& - IL0&, FIN& - IL0&)
     X$ = REGLEIDO$("ASIDIAR", IL&)
     If Left$(X$, 4) = TTXXY$ Then
       CUENUI% = CVI(Mid$(X$, 5, 2))
       IDEB# = CVD(Mid$(X$, 39, 8))
       If Abs(IDEB#) >= 0.005 Then
         IHAB# = 0
         If IDEB# < 0 Then
           IHAB# = Abs(IDEB#)
           IDEB# = 0
         End If
         CUES$ = CODCUE$(CUENUI%)
         DENOCUEX$ = ECOARCH$("ECUECON", CUES$)
         '
         Z$ = REGBLAN$("ASIDIAFI")
         Call REPLA(Z$, CUES$, 1, 12)
         Call REPLA(Z$, DENOCUEX$, 13, 30)
         Call REPLA(Z$, MKD$(IDEB#), 43, 8)
         Call REPLA(Z$, MKD$(IHAB#), 51, 8)
         '
         JJ& = JJ& + 1
         Call GRAREG("ASIDIAFI", Z$, JJ&)
       End If
     End If
   Next IL&
   Call SETULTREG("ASIDIAFI", JJ&)
   '
   Call FILESORT("ASIDIAFI", "", 1, 1, "ASC")
   Call TRALOCAL("ASIDIAFI", "")
   '
   Screen.MousePointer = 1
   '\\\OT-05-0710-WAR-03/11/05///
   If ULTREG&("ASIDIAFI") > 0 Then
   '\\\OT-05-0710-WAR-FIN///
    VTB% = VENTABU%("ASIDIAFI/TITUPAN=Asiento Resumen Mensual de Ventas - " + CORRESP$)
   End If
   '
99 Command20.Enabled = True
   '
End Sub

'Sub Command21_Click()
'    Command21.Enabled = False
'    If VALIMESAN% < 1 Then GoTo 99
'    '
'    RESP1% = ALTERNA%("Selección de Comprobantes a Listar.\Listado Libro de IVA - Ventas.")
'    If RESP1% = 1 Then
'        Call EXCLUIVA
'      ElseIf RESP1% = 2 Then
'        Option2.Value = True
'        'INCLUIDO PARA PODER LISTAR EL LIBRO IVA POR FECHA CONTABLE POR LOS DOCUMENTOS RECUPERADOS
'        MODO% = 0
'        If InStr(EMPREAC$, "MEDITEA") > 0 Then
'           If COMALTER("Elija Formato de Salida:\\Fecha Emision\Fecha Contable") = 2 Then MODO% = 1
'        End If
'        Call GENIVAVEN(MODO%)
'        '\\\OT-08-0890-OD-30/09/08///
'        If InStr(EMPREAC$, "DOSIVA") > 0 Then
'           Call IVADOSI
'           GoTo 99
'        End If
'        '\\\OT-08-0890-OD-FIN///
'        If EXISTE%("LIVENTAN.PRF") > 0 Then
'             'INCLUIDO PARA PODER LISTAR EL LIBRO IVA POR FECHA CONTABLE POR LOS DOCUMENTOS RECUPERADOS
'             LIBROIVA$ = "*LIVENTAN"
'             If MODO% = 1 Then
'                Call FILESORT("LIVENTAN", "", 2, 2, "ASC")
'                LIBROIVA$ = "*LIVENMED"
'             End If
'             Call FINAL(LIBROIVA$)
'           Else
'             FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVAVEN"))
'             If FORIPRE$ <> "" Then
'               If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
'                 If COMALTER("Elija Formato de Salida:\\Planilla Plana\Formulario") = 2 Then
'                   Call FORMIVAVE
'                   GoTo 99
'                 End If
'               End If
'             End If
'             Call FINAL("*LIVENMEN")
'        End If
'    End If
'    '
'99  Command21.Enabled = True
'End Sub

Sub FORMIVAVE()
    '
    Opciones% = COMALTER%("Formato de salida\\Impresión\Exportar a Excel")
    If Opciones% < 1 Or Opciones% > 2 Then Exit Sub
    
    If Opciones% = 2 Then GoTo 50
    
    FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVAVEN"))
    FEX = HOY(HOS$)
    FECHDOC$ = FECHATEX$(FEX)
    NUMEDOC$ = MESAN
    TIPODOC$ = "Libro IVA Ventas"
    '
    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    If ITPORHOJA% < 1 Then
      ITPORHOJA% = 32767
    End If
    '
    CODPARAM$(1) = "IMP-NETO" ': DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
    CODPARAM$(2) = "IMP-DESC" ': DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
    CODPARAM$(3) = "ITO-IVA"  ': DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
    CODPARAM$(4) = "IMP-RNI"  ': DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
    CODPARAM$(5) = "IMPNEDES" ': DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
    CODPARAM$(6) = "REF-MAT1" ': DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
    CODPARAM$(7) = "REF-MAT2" ': DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
    CAPARDOC% = 7
    '
    CODTABU1$(1) = "FECH-REC"
    CODTABU1$(2) = "TALLE"
    CODTABU1$(3) = "NRO-LOTE"
    CODTABU1$(4) = "RASO-CLI"
    CODTABU1$(5) = "CUIT-CLI"
    CODTABU1$(6) = "PIVA-CLI"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "BRU-ITEM"
    CODTABU1$(9) = "IMP-IVA"
    CODTABU1$(10) = "IMP-IVA1"
    CODTABU1$(11) = "IMP-TOTA"
    CODTABU1$(12) = "PERC-IVA"
    CODTABU1$(13) = "PERC-BRU"
    CACOLTAB1% = 13
    '
    DESTIDOCNUE = 0
    CAITAB1% = 0
    '
    SUMA1# = 0: SUMA2# = 0: SUMA3# = 0
    SUMA4# = 0: SUMA5# = 0: SUMA6# = 0: SUMA7# = 0
    URE& = ULTREG&("LIVAVEN")
    For YX& = 1 To URE&
       '
       YY$ = REGLEIDO$("LIVAVEN", YX&)
       FEMI = CVI(Mid$(YY$, 3, 2))
       NDOC$ = Mid$(YY$, 11, 18)
       Tipo$ = Left(NDOC$, 2) + Mid$(NDOC$, 18, 1)
       NUME$ = Mid$(NDOC$, 4, 13)
       RACLI$ = Mid$(YY$, 33, 30)
       CUITI$ = TRIM$(Mid$(YY$, 63, 16))
       PIVAI$ = TRIM$(Mid$(YY$, 79, 12))
       TOFA0# = CVD(Mid$(YY$, 91, 8))
       '
       NEGRA0# = 0: NEXEN0# = 0: NIVA21# = 0: NIVA105# = 0: RETIVA# = 0: RETGAN# = 0
       '
25     NETO1# = CVD(Mid$(YY$, 137, 8)) + CVD(Mid$(YY$, 185, 8))
       iva1# = CVD(Mid$(YY$, 145, 8))
       TASA1 = CVD(Mid$(YY$, 153, 8))
       If Abs(iva1#) < 0.005 Then
            NEXEN0# = NEXEN0# + NETO1#
          Else
            NEGRA0# = NEGRA0# + NETO1#
            If TASA1 >= 10 And TASA1 <= 11 Then
                 NIVA105# = NIVA105# + iva1#
               Else
                 NIVA21# = NIVA21# + iva1#
            End If
       End If
       RETIVA# = RETIVA# + CVD(Mid$(YY$, 161, 8))
       RETGAN# = RETGAN# + CVD(Mid$(YY$, 169, 8))
       '
       If YX& < URE& Then
         YZ$ = REGLEIDO$("LIVAVEN", YX& + 1)
         If TRIM$(Mid$(YZ$, 11, 18)) = "" Then
           YX& = YX& + 1
           YY$ = YZ$
           GoTo 25
         End If
       End If
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = FECHATEX$(FEMI)
       TETABU1$(2, CAITAB1%) = Tipo$
       TETABU1$(3, CAITAB1%) = NUME$
       TETABU1$(4, CAITAB1%) = RACLI$
       TETABU1$(5, CAITAB1%) = CUITI$
       TETABU1$(6, CAITAB1%) = PIVAI$
       TETABU1$(7, CAITAB1%) = FORMATNUM$(NEGRA0#, "F15.2")
          SUMA1# = SUMA1# + ROUND#(NEGRA0#)
       TETABU1$(8, CAITAB1%) = FORMATNUM$(NEXEN0#, "F15.2")
          SUMA2# = SUMA2# + ROUND#(NEXEN0#)
       TETABU1$(9, CAITAB1%) = FORMATNUM$(NIVA21#, "F15.2")
          SUMA3# = SUMA3# + ROUND#(NIVA21#)
       TETABU1$(10, CAITAB1%) = FORMATNUM$(NIVA105#, "F15.2")
          SUMA4# = SUMA4# + ROUND#(NIVA105#)
       TETABU1$(11, CAITAB1%) = FORMATNUM$(TOFA0#, "F15.2")
          SUMA5# = SUMA5# + ROUND#(TOFA0#)
       TETABU1$(12, CAITAB1%) = FORMATNUM$(RETIVA#, "F15.2")
          SUMA6# = SUMA6# + ROUND#(RETIVA#)
       TETABU1$(13, CAITAB1%) = FORMATNUM$(RETGAN#, "F15.2")
          SUMA7# = SUMA7# + ROUND#(RETGAN#)
          '
       If CAITAB1% >= ITPORHOJA% Then
         If YX& < URE& Then
            DOCPARAM$(1) = TRIM$(FORMATNUM$(SUMA1#, "F15.2"))
            DOCPARAM$(2) = TRIM$(FORMATNUM$(SUMA2#, "F15.2"))
            DOCPARAM$(3) = TRIM$(FORMATNUM$(SUMA3#, "F15.2"))
            DOCPARAM$(4) = TRIM$(FORMATNUM$(SUMA4#, "F15.2"))
            DOCPARAM$(5) = TRIM$(FORMATNUM$(SUMA5#, "F15.2"))
            DOCPARAM$(6) = TRIM$(FORMATNUM$(SUMA6#, "F15.2"))
            DOCPARAM$(7) = TRIM$(FORMATNUM$(SUMA7#, "F15.2"))
            Call PRINTDOC("@" + FORIPRE$)
            CAITAB1% = 0
         End If
       End If
    Next YX&
    '
    DOCPARAM$(1) = TRIM$(FORMATNUM$(SUMA1#, "F15.2"))
    DOCPARAM$(2) = TRIM$(FORMATNUM$(SUMA2#, "F15.2"))
    DOCPARAM$(3) = TRIM$(FORMATNUM$(SUMA3#, "F15.2"))
    DOCPARAM$(4) = TRIM$(FORMATNUM$(SUMA4#, "F15.2"))
    DOCPARAM$(5) = TRIM$(FORMATNUM$(SUMA5#, "F15.2"))
    DOCPARAM$(6) = TRIM$(FORMATNUM$(SUMA6#, "F15.2"))
    DOCPARAM$(7) = TRIM$(FORMATNUM$(SUMA7#, "F15.2"))
    Call PRINTDOC("@" + FORIPRE$)
    
    Exit Sub

50  Dim iFila As Long, iCol As Integer, i As Integer
    
    On Error GoTo CreaExcel
    Dim OExcel As Excel.Application
    Dim oWBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet

    Set OExcel = New Excel.Application
    Set oWBook = OExcel.Workbooks.Add
    Set oSheet = oWBook.Worksheets.Add
    On Error GoTo 0

    CAITAB1% = 1
    oSheet.Cells(CAITAB1%, 1) = "FECHA"
    oSheet.Cells(CAITAB1%, 2) = "TIPO"
    oSheet.Cells(CAITAB1%, 3) = "NUMERO"
    oSheet.Cells(CAITAB1%, 4) = "RAZON SOCIAL"
    oSheet.Cells(CAITAB1%, 5) = "CUIT"
    oSheet.Cells(CAITAB1%, 6) = "COND. IVA"
    oSheet.Cells(CAITAB1%, 7) = "NETO GRAV. %"
    oSheet.Cells(CAITAB1%, 8) = "CONC. Y OPERAC. EXCL."
    oSheet.Cells(CAITAB1%, 9) = "IVA 21%"
    oSheet.Cells(CAITAB1%, 10) = "IVA 10.5%"
    oSheet.Cells(CAITAB1%, 11) = "IMPORTE TOTAL FACT."
    oSheet.Cells(CAITAB1%, 12) = "RETENC. IVA"
    oSheet.Cells(CAITAB1%, 13) = "RETENC. GCIAS."

    SUMA1# = 0: SUMA2# = 0: SUMA3# = 0
    SUMA4# = 0: SUMA5# = 0: SUMA6# = 0: SUMA7# = 0
    URE& = ULTREG&("LIVAVEN")
    CAITAB1% = 2
    For YX& = 1 To URE&
       '
       YY$ = REGLEIDO$("LIVAVEN", YX&)
       FEMI = CVI(Mid$(YY$, 3, 2))
       NDOC$ = Mid$(YY$, 11, 18)
       Tipo$ = Left(NDOC$, 2) + Mid$(NDOC$, 18, 1)
       NUME$ = Mid$(NDOC$, 4, 13)
       RACLI$ = Mid$(YY$, 33, 30)
       CUITI$ = TRIM$(Mid$(YY$, 63, 16))
       PIVAI$ = TRIM$(Mid$(YY$, 79, 12))
       TOFA0# = CVD(Mid$(YY$, 91, 8))
       '
       NEGRA0# = 0: NEXEN0# = 0: NIVA21# = 0: NIVA105# = 0: RETIVA# = 0: RETGAN# = 0
       '
60     NETO1# = CVD(Mid$(YY$, 137, 8)) + CVD(Mid$(YY$, 185, 8))
       iva1# = CVD(Mid$(YY$, 145, 8))
       TASA1 = CVD(Mid$(YY$, 153, 8))
       If Abs(iva1#) < 0.005 Then
            NEXEN0# = NEXEN0# + NETO1#
          Else
            NEGRA0# = NEGRA0# + NETO1#
            If TASA1 >= 10 And TASA1 <= 11 Then
                 NIVA105# = NIVA105# + iva1#
               Else
                 NIVA21# = NIVA21# + iva1#
            End If
       End If
       RETIVA# = RETIVA# + CVD(Mid$(YY$, 161, 8))
       RETGAN# = RETGAN# + CVD(Mid$(YY$, 169, 8))
       '
       If YX& < URE& Then
         YZ$ = REGLEIDO$("LIVAVEN", YX& + 1)
         If TRIM$(Mid$(YZ$, 11, 18)) = "" Then
           YX& = YX& + 1
           YY$ = YZ$
           GoTo 60
         End If
       End If
       '
       CAITAB1% = CAITAB1% + 1
          oSheet.Cells(CAITAB1%, 1).NumberFormat = "dd/mm/yyy":
          FECHA = FECHATEX$(FEMI)
       oSheet.Cells(CAITAB1%, 1) = CDate(FECHA)
       oSheet.Cells(CAITAB1%, 2) = Tipo$
       oSheet.Cells(CAITAB1%, 3) = NUME$
       oSheet.Cells(CAITAB1%, 4) = RACLI$
       oSheet.Cells(CAITAB1%, 5) = CUITI$
       oSheet.Cells(CAITAB1%, 6) = PIVAI$
          oSheet.Cells(CAITAB1%, 7).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 7) = XVALO(FORMATNUM$(NEGRA0#, "F15.2"))
          SUMA1# = SUMA1# + ROUND#(NEGRA0#)
          oSheet.Cells(CAITAB1%, 8).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 8) = XVALO(FORMATNUM$(NEXEN0#, "F15.2"))
          SUMA2# = SUMA2# + ROUND#(NEXEN0#)
          oSheet.Cells(CAITAB1%, 9).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 9) = XVALO(FORMATNUM$(NIVA21#, "F15.2"))
          SUMA3# = SUMA3# + ROUND#(NIVA21#)
          oSheet.Cells(CAITAB1%, 10).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 10) = XVALO(FORMATNUM$(NIVA105#, "F15.2"))
          SUMA4# = SUMA4# + ROUND#(NIVA105#)
          oSheet.Cells(CAITAB1%, 11).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 11) = XVALO(FORMATNUM$(TOFA0#, "F15.2"))
          SUMA5# = SUMA5# + ROUND#(TOFA0#)
          oSheet.Cells(CAITAB1%, 12).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 12) = XVALO(FORMATNUM$(RETIVA#, "F15.2"))
          SUMA6# = SUMA6# + ROUND#(RETIVA#)
          oSheet.Cells(CAITAB1%, 13).NumberFormat = "#,##0.00"
       oSheet.Cells(CAITAB1%, 13) = XVALO(FORMATNUM$(RETGAN#, "F15.2"))
          SUMA7# = SUMA7# + ROUND#(RETGAN#)
    Next YX&
    
    CAITAB1% = CAITAB1% + 2
    
    oSheet.Cells(CAITAB1%, 4) = "TOTALES"
        oSheet.Cells(CAITAB1%, 7).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 7) = XVALO(TRIM$(FORMATNUM$(SUMA1#, "F15.2")))
        oSheet.Cells(CAITAB1%, 8).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 8) = XVALO(TRIM$(FORMATNUM$(SUMA2#, "F15.2")))
        oSheet.Cells(CAITAB1%, 9).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 9) = XVALO(TRIM$(FORMATNUM$(SUMA3#, "F15.2")))
        oSheet.Cells(CAITAB1%, 10).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 10) = XVALO(TRIM$(FORMATNUM$(SUMA4#, "F15.2")))
        oSheet.Cells(CAITAB1%, 11).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 11) = XVALO(TRIM$(FORMATNUM$(SUMA5#, "F15.2")))
        oSheet.Cells(CAITAB1%, 12).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 12) = XVALO(TRIM$(FORMATNUM$(SUMA6#, "F15.2")))
        oSheet.Cells(CAITAB1%, 13).NumberFormat = "#,##0.00"
    oSheet.Cells(CAITAB1%, 13) = XVALO(TRIM$(FORMATNUM$(SUMA7#, "F15.2")))
    '
    On Error GoTo SaveExcel
    oSheet.SaveAs LUCOM$ & "IVAVENTA.XLS"
    On Error GoTo 0
    OExcel.Quit      ' cierra EXCEL que se esta ejecutando minimizado
    Set OExcel = Nothing     ') libera memoria
    Set oWBook = Nothing    ') utilizada por
    Set oSheet = Nothing   ') las variables objeto
    
    Call COMUNI("La Exportación se Realizó Correctamente.\ \Ubicar el Archivo en\" & REPLACAR$(LUCOM$ & "IVAVENTA.XLS", "\", "/"))
    Exit Sub
    
CreaExcel:
    Call MENSERR(24, "Error al Crear Archivo Excel.\ \Verifique la Correcta Instalación de MS Excel.")
    Exit Sub
    
SaveExcel:
    Call MENSERR(24, "Problemas al Grabar Archivo Excel\ \'" & LUCOM$ & "IVAVENTA.XLS'")
    Exit Sub
    '
End Sub


'
Sub Command22_Click()
    '
    Command22.Enabled = False
    '
    If DERACCE%("REIMPUVE", "Re-Imputación de Comprobantes de Ventas") < 2 Then
      GoTo 99
    End If
    '
    If VALIMESAN% < 1 Then GoTo 99
    Des% = DESDE%: DESDEX = CDbl(CVDAT(Des%))
    Has% = HASTA%: HASTAX = CDbl(CVDAT(Has%))
    '
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then GoTo 99
    '
    Call APERBASDAT("CABAN")
    'jandy sql
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TIPOMOVIM = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + "AND FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            Y$ = MKS$(JJ& + 1) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Tiene\Comprobantes de Venta\que Puedan Re-Imputarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
16  Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    '
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(rasocli$(NCLI))
    '
    Call BLANARCH("!REIMPUTA")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NCLI & " "
    '
    Dim CABASEL As ADODB.Recordset
    Set CABASEL = New ADODB.Recordset
25  On Error Resume Next
    CABASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    SUIMPU# = 0
    With CABASEL
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Editar Imputación Contable\del Comprobante Elegido")
        GoTo 10
      End If
      '
      On Error GoTo 0
      Do While Not .EOF
        CUE% = !CUECONTAI
        'If CUE% > 0 Then
           IPA# = !idebecomp - !ihabecomp
           Y$ = REGBLAN$("!REIMPUTA")
           Call REPLA(Y$, CODCUE$(CUE%), 1, 12)
           Call REPLA(Y$, MKD$(IPA#), 13, 8)
           Call REGAPP("!REIMPUTA", Y$)
           SUIMPU# = SUIMPU# + IPA#
        'End If
      .MoveNext
      Loop
    End With
    
    Call CIERRARCH("!REIMPUTA")
    '
    If Abs(SUIMPU#) >= 0.005 Then
      Call MENSERR(24, "Imposible Editar Imputación Contable.\Comprobante No Balanceado.\  \Debe Anularse y Registrarse Nuevamente.")
      GoTo 10
    End If
    '
30  ATRI$ = "TITUPAN=Re-Imputación Contable: " + ATRI$
    VTB% = VENTABU%("REIMPUTA/" + ATRI$)
    If VTB% < 0 Then GoTo 10
    '
    SUMANE# = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      SUMANE# = SUMANE# + IMPOX#
      '
      If CUENUE% < 1 Then
        Call MENSERR(24, "Numero de Cuenta no Valido\en Registro Numero " + TRIM$(Str$(U&)))
        GoTo 30
      End If
    Next U&
    If Abs(SUMANE#) >= 0.005 Then
      Call MENSERR(24, "Comprobante no Balanceado")
      GoTo 20
    End If
    '
    REIMPU% = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      '
      With CABASEL
        .MoveFirst
        Do While Not .EOF
32        CUE% = !CUECONTAI
          If Abs(!idebecomp - !ihabecomp - IMPOX#) < 0.005 Then
            !CUECONTAI = CUENUE%
            .Update
            REIMPU% = 1
            GoTo 35
          End If
        .MoveNext
        Loop
      End With
      '
35  Next U&
    '
    CABASEL.Close
    Set CABASEL = Nothing
    
    If REIMPU% > 0 Then
        Call COMUNI("Re-Imputación Completada")
      Else
        Call COMUNI("No se Pudo Completar la Re-Imputación")
    End If
    '
    GoTo 16
    '
99  Command22.Enabled = True

    '
End Sub

Private Sub INGBRUT()
                '\\\OT-06-0158-WAR-31/05/06///
   'Command23.Enabled = False
   '
   'If VALIMESAN% < 1 Then GoTo 99
   '
   'Screen.MousePointer = 11
   'Call GENIVAVEN
   'Screen.MousePointer = 1
   If ULTREG&("SUPERIB") < 1 Then
        Call COMUNI("No Existen Percepciones en el Período Elegido")
        GoTo 99
   End If
   '
5  OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
   If OPX% = 1 Then
       Call HEADERS("SUPERIB", "")
       Call HEADERS("SUPERIB", "Corresponde a: " + MESAN)
       Call FINAL("*SUPERIB")
     ElseIf OPX% = 2 Then
       '\\\OT-06-0076-WAR-02/03/06///'Para Tanit
       'para filtrar dentro del mes y año elegido.
10     Call MESANO(MESAN, DESX1%, HASX1%)
       VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
       OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
       If OBX$ = "" Then GoTo 99
       DESX2% = CVI(Left$(OBX$, 2))
       HASX2% = CVI(Mid$(OBX$, 3, 2))
       If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
         Call MENSERR(24, "Rango de Fechas no Procesable")
         GoTo 10
       End If
       ''\\\OT-06-0076-WAR-FIN///
       'SE AGREGA PARA QUE EL USUARIO INGRESE EL NOMBRE DEL ARCHIVO DE TRANSFERENCIA SEGUN NUEVA NORMATIVA A PARTIR DE SETIEMBRE 2011
'       NOMBASALID$ = TRIM$(InputBox("Ingrese Nombre del Archivo", "Archivo IIBB", ""))
'       If NOMBASALID$ = "" Then GoTo 99
        OPACTIVIDAD% = COMALTER%("Modalidad:\\Devengado\Percibido")
        If OPACTIVIDAD% < 1 Or OPACTIVIDAD% > 2 Then GoTo 10
        
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
    
        DoEvents
        Screen.MousePointer = 11
        ASALID$ = UCase$(CommonDialog1.FileName)
        If TRIM$(ASALID$) = "" Then GoTo 5
        On Error Resume Next
        
       'ASALID$ = "C:\FLEXOFT\" & NOMBASALID$ & ".TXT"
       'ASALID$ = "C:\FLEXOFT\TRANIBRU.TXT"
       If OPACTIVIDAD% = 1 Then
            NX% = FILEOPEN%(ASALID$, 2, 61)
       Else
            NX% = FILEOPEN%(ASALID$, 2, 71)
       End If
       CAREGI% = 0
       For U& = 1 To ULTREG&("SUPERIB")
         XX$ = REGLEIDO$("SUPERIB", U&)
         IMPOL# = CVD(Mid$(XX$, 87, 8))
         If Abs(IMPOL#) > 0 Then
              FEX = CVI(Left$(XX$, 2))
              If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
            If OPACTIVIDAD% = 1 Then
                TTXX$ = Space$(61)
            Else
                TTXX$ = Space$(71)
            End If
              CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
            Call REPLA(TTXX$, CUITTX$, 1, 13)
              FEXU$ = FECHATEX$(FEX)
              FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
            Call REPLA(TTXX$, FEXU$, 14, 10)
              LECOX$ = "F"
              If Mid$(XX$, 3, 2) = "ND" Then LECOX$ = "D"
              If Mid$(XX$, 3, 2) = "NC" Then LECOX$ = "C"
            Call REPLA(TTXX$, LECOX$, 24, 1)
              LECO$ = Mid$(XX$, 20, 1)
            Call REPLA(TTXX$, LECO$, 25, 1)
              SUCUX$ = Mid$(XX$, 6, 4)
            Call REPLA(TTXX$, SUCUX$, 26, 4)
              NUCOX$ = Mid$(XX$, 11, 8)
            Call REPLA(TTXX$, NUCOX$, 30, 8)
              IMPOL# = CVD(Mid$(XX$, 79, 8))
              'IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
              IMPOX$ = TRIM$(Format$(Abs(IMPOL#), "#########0.00"))
              While Len(IMPOX$) < 12: IMPOX$ = "0" + IMPOX$: Wend
              If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
              Call REPLA(TTXX$, IMPOX$, 38, 12)
              IMPOL# = CVD(Mid$(XX$, 87, 8))
              'IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "#########0.00")), ".", ",")
              IMPOX$ = TRIM$(Format$(Abs(IMPOL#), "#########0.00"))
              While Len(IMPOX$) < 11: IMPOX$ = "0" + IMPOX$: Wend
              If LECOX$ = "C" Then IMPOX$ = "-" + Mid$(IMPOX$, 2)
              Call REPLA(TTXX$, IMPOX$, 50, 11)
              If OPACTIVIDAD% = 1 Then
                Call REPLA(TTXX$, "A", 61, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
              ElseIf OPACTIVIDAD% = 2 Then
                Call REPLA(TTXX$, FEXU$, 61, 10)
                Call REPLA(TTXX$, "A", 71, 1) ' SE AGREGA SEGUN NUEVA NORMATIVA DE ARBA SETIEMBRE 2011
              End If
              Print #NX%, TTXX$
              CAREGI% = CAREGI% + 1
          End If
29     Next U&
       Close #NX%
       
       If CAREGI% = 0 Then
            Call COMUNI("No Existen Percepciones en el Período Elegido")
       Else
            ASALID1$ = REPLACAR$(ASALID$, "\", "/")
            Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
       End If
       Screen.MousePointer = 1
   End If
   '
99 'Command23.Enabled = True
   On Error GoTo 0
End Sub


Private Sub Command24_Click()
   '\\\OT-06-0158-WAR-31/05/06///
   
   Call COMUNI("Opciones Trasladadas al Menú Superior denominado 'Impuestos'")
   Exit Sub
   
   Command24.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call GENIVAVEN
   Screen.MousePointer = 1
   '
3  OPX0% = COMALTER%("Opciones:\\Retenciones\Percepciones\Regimen Informativo de Ventas")
   'OPX0% = COMALTER%("Opciones de Retención: \\Retención de Ingresos Brutos\Retención de I.V.A.\Perecpciones")
   If OPX0% < 1 Then GoTo 99
   If OPX0% = 1 Then
      OPX01% = COMALTER%("Opciones de Retención:\\Ingresos Brutos\I.V.A.")
      If OPX01% < 1 Or OPX1% > 2 Then GoTo 3
      If OPX01% = 1 Then Call RETENCIONIIBB: GoTo 3
      If OPX01% = 2 Then Call RETENCIONIVA: GoTo 3
   ElseIf OPX0% = 2 Then
5     OPX1% = COMALTER%("Opciones de Percepción:\\I.V.A.\Ingresos Brutos")
   ElseIf OPX0% = 3 Then
      Call RegimenInformativoVentas: GoTo 3
   End If
   
'5  OPX1% = COMALTER%("Opciones de Percepción: \\I.V.A.\Ingresos Brutos\Regimen Informativo de Ventas")
   If OPX1% = 1 Then 'PERCEPCIONES DE IVA
     If ULTREG&("LIVAVEN") < 1 Then GoTo 99
     '
10   Call MESANO(MESAN, DESX1%, HASX1%)
     VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
     OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
     If OBX$ = "" Then GoTo 99
     DESX2% = CVI(Left$(OBX$, 2))
     HASX2% = CVI(Mid$(OBX$, 3, 2))
     If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
       Call MENSERR(24, "Rango de Fechas no Procesable")
       GoTo 10
     End If
     '
     VD1$ = ""
15   CUIEM1$ = InputBox$("Ingrese Nº de C.U.I.T.", "Agente de Percepción", VD1$)
     If TRIM$(CUIEM1$) = "" Then
       Call MENSERR(24, "C.U.I.T. no Válido")
       GoTo 5
     End If
     '
     VD1$ = CUIEM1$
     If VALICUIT%(CUIEM1$) <= 0 Then
      Call MENSERR(24, "C.U.I.T. no Válido")
      GoTo 15
     End If
     '
     ASALID$ = "C:\FLEXOFT\TRAPEIVA.TXT"
     NX% = FILEOPEN%(ASALID$, 2, 128)
     CAREGI% = 0
     For U& = 1 To ULTREG&("LIVAVEN")
       XX$ = REGLEIDO$("LIVAVEN", U&)
         FEX = CVI(Mid$(XX$, 3, 2))
         If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
       If CVD(Mid$(XX$, 161, 8)) > 0 Then ' SI PERCEPCION I.V.A.
         TTXX$ = Space$(60)
         '
         CODPER1$ = CONTROL$("COREPER", "COREPER")
         If TRIM$(CODPER1$) = "" Then
          Call MENSERR(24, "Código de Regimen de Percepción no Válido.\Se ha Cancelado la Exportación.")
          GoTo 99
         End If
         'CUIT DEL AGENTE
         Call REPLA(TTXX$, CUIEM1$, 4, 13)
         'CODIGO DE REGIMEN DE PERCEPCION
         Call REPLA(TTXX$, CODPER1$, 1, 3)
         'FECHA DE LA PERCEPCION
         FEXU$ = FECHATEX$(FEX)
         FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         Call REPLA(TTXX$, FEXU$, 17, 10)
         'PREFIJO DE FACTURA
         X11$ = Mid$(XX$, 11, 18)
         PREFI1$ = Mid$(X11$, 4, 4)
         While Len(PREFI1$) < 8: PREFI1$ = "0" + PREFI1$: Wend
         Call REPLA(TTXX$, PREFI1$, 27, 8)
         'NUMERO DE FACTURA
         X2 = CVS(Mid$(XX$, 7, 4))
         NUMEFAC1$ = TRIM$(Str$(X2))
         While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
         Call REPLA(TTXX$, NUMEFAC1$, 35, 8)
         'MONTO DE LA PERCEPCION
         IMPOL# = CVD(Mid$(XX$, 161, 8))
         IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "###########0.00")), ".", ",")
         While Len(IMPOX$) < 16: IMPOX$ = "0" + IMPOX$: Wend
         Call REPLA(TTXX$, IMPOX$, 43, 16)
         '
         Print #NX%, TTXX$
         CAREGI% = CAREGI% + 1
       End If
29   Next U&
     Close #NX%
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
     
   ElseIf OPX1% = 2 Then 'INGRESOS BRUTOS
     Call INGBRUT
'   ElseIf OPX1% = 3 Then
'     Call RegimenInformativoVentas
   End If
99    Command24.Enabled = True
   '\\\OT-06-0158-WAR-FIN///
End Sub
Sub RETENCIONIVA(Optional GeneraParaReporte%)

            'SINTESIS
            'SE RECORRE EL CAJABANC SEGUN RANGO SELECCIONADO
            'Y QUE LAS CUENTAS QUE SE CONFIGURAN EN EL SETUP DE RECIBOS ASOCIANDO AL CODIGO DE REGIMEN '
            'ESTA CUENTA DEBERA ESTAR TAMBIEN CARGADA COMO MEDIDO DE COBRANZA.
            'CUANDO SE CARGA UN RECIBO Y SE TOMA COMO MEDIO DE COBRANZA UN RETENCION DEL CLIENTE SE DEBERA INFORMAR
            'LA MISMA EN EL DETALLE COMO UN MEDIO DE PAGO MAS, DEBIENDO ESTE ESTAR ASOCIADO A LA CUENTA DE Cta. Ret. de IVA.
            'EL SISTEMA GENERA UN ARCHIVO .TXT CON LOS RECIBOS QUE EN SU DETALLE CONTENGAN ALGUN MOVIMIENTO CON DICHA CUENTA CONTABLE.
            'INGRESA TODOS ESTOS CASOS EN DICHO ARCHIVO SEGUN EL FORMATO INDICADO DE LA AFIP.
            '
            ' *** VALIDACION DE CUENTA CONTABLE Y CODIGOS DE REGIMEN
            ' *** PARA RETENCIONES PROVENIENTES DE RECIBOS Y TARJETAS DE CREDITO
            CUECONTA$ = CONTROL$("OPRECIBO", "CURETIVA")
            CUE% = CUECOINT%(CUECONTA$)
            If CUE% < 1 Then
              Call COMUNI("Falta Configurar Cuenta Contable de Retención de IVA.\ \Configurar la Misma Desde el Setup de Recibos.")
              GoTo 99
            End If
            '
            CODREG& = XVALO(CONTROL$("OPRECIBO", "CREREIVA"))
            If CODREG& < 1 Then
              Call COMUNI("Falta Configurar Código de Régimen de Retención de IVA para Recibos.\ \Configurar la Misma Desde el Setup de Recibos.")
              GoTo 99
            End If
            '
            CODREGXTC = XVALO(CONTROL$("OPRECIBO", "COREIVAT"))
            If CODREGXTC < 1 Then
              Call COMUNI("Falta Configurar Código de Régimen de Retención de IVA de Tarjetas de Crédito.\ \Configurar la Misma Desde el Setup de Gestión de Valores.")
              GoTo 99
            End If
            '
            ' *** ARMA DISEÑO DEL FORMULARIO DE ERRORES
            Campos$ = "Fecha/D8;Documento/A18;Agente/A32;Constancia/A24;Importe/F12.2;Informe/A64"
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
            FRMZELECHO.Width = 12000
            FRMZELECHO.Caption = "Informe de Errores"
            
            ' *** NOMBRE DEL ARCHIVO
            DES1$ = FETEXCOR1$(DESDE%)
            HAS1$ = FETEXCOR1$(HASTA%)
            If GeneraParaReporte% = 0 Then
                ARCHIX$ = SELECCION_Archivo$("TXT", CommonDialog1)
                If TRIM$(ARCHIX$) = "" Then GoTo 99
            Else
                ARCHIX$ = LUCOM$ & "REPORTE.TXT"
            End If
            Call SETULTREG("LREPEIVA", 0)
            Call BLANARCH("LREPEIVA")
            
            Screen.MousePointer = 11
            NARCHI2% = FILEOPEN%(ARCHIX$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
            GENERO_ARCHIVO% = 0
            '
            ' *** RECORRE RETENCIONES EN RECIBOS Y CONCILIACION DE TARJETAS
            SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE FECHCONTA >= '" & DES1$ & "'"
            SQLX$ = SQLX$ + " AND FECHCONTA <= '" & HAS1$ & "' "
            SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'CHRECI' OR TIPOMOVIM = 'GASTC') " ' RECIBOS Y CONCILIACIONES DE TARJ. DE CREDITO
            SQLX$ = SQLX$ + " AND CUECONTAI = " & CUE% ' LA CUENTA ES IGUAL PARA TARJETA Y RECIBOS
            SQLX$ = SQLX$ + " ORDER BY FECHEMISI ASC "
            Dim rs As ADODB.Recordset
            Set rs = READSET(SQLX$)
            J& = 0
            With rs
              If .EOF And .BOF Then
                   Screen.MousePointer = 1
                   Call COMUNI("No Se Han Encontrado Retenciones Para el Período Elegido.")
                   Close #NARCHI2%
                   rs.Close
                   Set rs = Nothing
                   GoTo 40
              End If

              Do While Not .EOF
                 NC = XVALO(!NUCLIEN)
                 FE% = CVINT(!FECHEMISI)
                 IMPORTEAPLIC# = XVALO(!idebecomp)
                 DOCUMENTO$ = SAFETEXT$(!codicom_lar)
                 
                 TMOXI$ = TRIM$(SAFETEXT$(!TIPOMOVIM))
                 NUMCONSTANCIA$ = SAFETEXT$(!DESCRIMOV)
                 '
                 ' *** RETENCION POR CONCILIACION DE TARJETA DE CREDITO
                 If TMOXI$ = "GASTC" Then
                    CODREGX$ = CStr(CODREGXTC) ' CODIGO DE RETENCION DE TARJETAS DE CREDITO
                    'DATOX$ = TRIM$(Mid$(NUMCONSTANCIA$, 14, 14))
                    'XX$ = FILTERGETRECORD$("MECOBRA", 2, DATOX$) ' LEO EL REGISTRO
                    'MPAGO% = CVI(Left(XX$, 2))
                    'CTACBLEX$ = TRIM$(Mid$(XX$, 45, 12))
                    'CUECO% = CUECOINT%(CTACBLEX$)
                    'CUECO% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", "CODICOM_LAR='" & DOCUMENTO$ & "' AND IHABECOMP > 0.005"))
                    CUECO% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", "TIPOMOVIM='CONCITC' AND CODICOM_LAR='" & DOCUMENTO$ & "' AND IDEBECOMP > 0.005"))
                    
                    NC = XVALO(VALOBADA("PROVED12", "Nume_Cli", "CuMa_Cli = " & CUECO%, "NOMESS"))
                    CUITX$ = CuitCli$(-1 * NC)
                    
                    ' VALIDO SI ES CORRECTO EL REGISTRO CON LA DESCRIPCION DEL COMIENZO (POR QUE SE CAMBIO EN LA GRABACION, ABREVIANDO LA MISMA )
                    If TRIM$(Left$(NUMCONSTANCIA$, 13)) = "Conc. T.Cred." Then
                      NUMCONSTANCIA$ = TRIM$(Mid$(NUMCONSTANCIA$, 28, 20)) ' SI ES ASI LO LEE
                      NUCONTAN$ = TRIM$(REPLACAR$(NUMCONSTANCIA$, ";", ""))
                      If NUCONTAN$ <> "" Then
                         NUMCONSTANCIA$ = NUCONTAN$
                      End If
                    Else
                      NUMCONSTANCIA$ = "" ' SI NO LO DEJA VACIO
                    End If
                 ' *** RETENCION POR RECIBO
                 Else
                    NUMCONSTANCIA$ = TRIM$(Mid$(NUMCONSTANCIA$, 19, 12))
                    CUITX$ = CuitCli(NC)
                    CODREGX$ = CStr(CODREG&) ' CODIGO DE RETENCION DE RECIBO
                 End If
                 
                 ' REEMPLAZA EN EL NUMERO DE CONSTANCIA LOS CARACTERES INVALIDOS POR '0'
                 ' LOS CARACTERES INVALIDOS LUEGO PROVOCAN UN ERROR AL IMPORTAR DESDE EL APLICATIVO SIFERE
                 For INDCON% = 1 To Len(NUMCONSTANCIA$)
                    ASCIICAR% = Asc(Mid$(NUMCONSTANCIA$, INDCON%, 1))
                    If ASCIICAR% < 48 Or (ASCIICAR% > 57 And ASCIICAR% < 65) Or (ASCIICAR% > 90 And ASCIICAR% < 97) Or ASCIICAR% > 122 Then
                        NUMCONSTANCIA$ = REPLACAR$(NUMCONSTANCIA$, Mid$(NUMCONSTANCIA$, INDCON%, 1), "0")
                    End If
                 Next INDCON%
                    
                 ' VALIDACIONES
                 TXTERROR$ = ""
                 If CUITX$ = "" Then
                    TXTERROR$ = "CUIT del Agente"
                 ElseIf VALICUIT%(CUITX$) = 0 Then
                    TXTERROR$ = "CUIT Inválido del Agente de Retención"
                 End If
                 If NUMCONSTANCIA$ = "" Then
                    If TXTERROR$ = "" Then
                        TXTERROR$ = "Falta Número de Constancia de la Retención"
                    Else
                        TXTERROR$ = TXTERROR$ & ", Falta Número de Constancia de la Retención"
                    End If
                 End If
                 If TMOXI$ = "GASTC" Then NC = NC * -1
                 RACLI$ = CStr(Abs(NC)) & " - " & rasocli$(NC)
                 If TXTERROR$ <> "" Then
                   J& = J& + 1
                   FRMZELECHO.MSF2.Rows = J& + 1
                   FRMZELECHO.MSF2.TextMatrix(J&, 1) = FECHATEX$(FE%)
                   FRMZELECHO.MSF2.TextMatrix(J&, 2) = DOCUMENTO$
                   FRMZELECHO.MSF2.TextMatrix(J&, 3) = rasocli$(NC)
                   FRMZELECHO.MSF2.TextMatrix(J&, 4) = NUMCONSTANCIA$
                   FRMZELECHO.MSF2.TextMatrix(J&, 5) = FORMATNUM(IMPORTEAPLIC#, "F12.2")
                   TXTERROR$ = TRIM$(TXTERROR$)
                   FRMZELECHO.MSF2.TextMatrix(J&, 6) = TXTERROR$
                 End If

                 FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
                 GENERO_ARCHIVO% = 1
                 If Len(CODREGX$) < 3 Then CODREGX$ = String$(3, "0") & CODREGX$: CODREGX$ = Right(CODREGX$, 3)
                 Call REPLA(TX$, CODREGX$, 1, 3)
                  If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
                 Call REPLA(TX$, CUITX$, 4, 13)
                  If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
                 Call REPLA(TX$, FEMIS$, 17, 10)
                   If Len(NUMCONSTANCIA$) < 16 Then NUMCONSTANCIA$ = String$(16, "0") & NUMCONSTANCIA$: NUMCONSTANCIA$ = Right(NUMCONSTANCIA$, 16)
                 Call REPLA(TX$, NUMCONSTANCIA$, 27, 16)
                 'FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
                 IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F16.2"))
                 If InStr(IMPORTEAPL$, ".") > 0 Then
                   IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
                 Else
                   IMPORTEAPL$ = IMPORTEAPL$ & ",00"
                 End If
                 If Len(IMPORTEAPL$) < 16 Then IMPORTEAPL$ = String$(16, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 16)
                 Call REPLA(TX$, IMPORTEAPL$, 43, 16)
                 
                 Print #NARCHI2%, TX$
                 
                 TTXX$ = REGBLAN$("LREPEIVA")
                 Call REPLA(TTXX$, MKI(FE%), 1, 2)
                 Call REPLA(TTXX$, DOCUMENTO$, 3, 18)
                 Call REPLA(TTXX$, rasocli$(NC), 21, 48)
                 Call REPLA(TTXX$, NUMCONSTANCIA$, 69, 24)
                 Call REPLA(TTXX$, MKD$(IMPORTEAPLIC#), 93, 8)
                 Call REPLA(TTXX$, Space(91), 101, 91)
                 If TXTERROR$ <> "" Then
                     TextoError$ = TXTERROR$
                     Call FRATEXTO(TextoError$, 60)
                     For KKII% = 1 To CARETEX%
                         Call REPLA(TTXX$, RETEX(KKII%), 101, 60)
                         Call REGAPP("LREPEIVA", TTXX$)
                     Next KKII%
                 Else
                     Call REGAPP("LREPEIVA", TTXX$)
                 End If
                  
                 .MoveNext
              Loop
            End With
            
            ' *** VERIFICA SI HAY ERRORES Y PRESENTA VENTANA CON EL DETALLE DE LOS MISMOS
            If J& > 0 And GeneraParaReporte% = 0 Then
                 Call COMUNI("No Se Han Exportado las Retenciones\Debido a que Existen Errores en los Datos.\ \Aceptar Para Ver Listado de Retenciones Fallidas.")
                 Screen.MousePointer = 1
                 'FRMZELECHO.MSF2.COL = 1
                 'FRMZELECHO.MSF2.Sort = 1
                 'FRMZELECHO.Show 1
                 Call CIERRARCH("LREPEIVA")
                 Call FILESORT("LREPEIVA", "", 1, 1)
                 Close #NARCHI2%
                 Call FINAL("*LIRETIVA")
                 Exit Sub
            End If
            
            'Close #NARCHI2%
            '
'            'AHORA CARGO LOS ASIENTOS
'            Call CargaAsientosRETENCIONIVA(DESDE%, HASTA%, ARCHIX$, OK%, RET_TX$)  'RET_TX$ ES UN PARAMETRO DE RETORNO
'            If OK% < 1 Then
'               Close #NARCHI2%
'               GoTo 40
'             Else
'               If TRIM$(RET_TX$) <> "" Then
'                 Print #NARCHI2%, RET_TX$
'               End If
'               Close #NARCHI2%
'             End If

            ' *** FINALIZA GRABACION DEL ARCHIVO
30          Close #NARCHI2%
            If GENERO_ARCHIVO% > 0 And GeneraParaReporte% = 0 Then
               ARCHIX$ = REPLACAR$(ARCHIX$, "\", "/")
               Screen.MousePointer = 1
               Call COMUNI("Archivo Generado Correctamente en '" & ARCHIX$ & "'")
               GoTo 99
            End If

            ' *** CIERRA ARCHIVO Y LO ELIMINA
40          If GeneraParaReporte% = 0 Then
                On Error Resume Next
                Close #NARCHI2%
                Call FILEKILL(ARCHIX$)
                If Err > 0 Then
                   Screen.MousePointer = 1
                   Call COMUNI("No Se Ha Podido Eliminar el Archivo\'" & ARCHIX$ & "'.\ \Realice Las Correcciones Mencionadas en el Detalle y Vuelva a Generar")
                End If
                On Error GoTo 0
            End If

99          Screen.MousePointer = 1
            Call CIERRARCH("LREPEIVA")
            Call FILESORT("LREPEIVA", "", 1, 1)
End Sub

Sub RETENCIONIIBB(Optional GeneraParaReporte%)

            'SINTESIS
            'SE RECORRE EL CAJABANC SEGUN RANGO SELECCIONADO
            'Y QUE LAS CUENTAS QUE SE CONFIGURAN EN EL SETUP DE RECIBOS ASOCIANDO A CADA PROVINCIA '
            'ESTA CUENTA DEBERA ESTAR TAMBIEN CARGADA COMO MEDIDO DE COBRANZA.
            'CUANDO SE CARGA UN RECIBO Y SE TOMA COMO MEDIO DE COBRANZA UN RETENCION DEL CLIENTE SE DEBERA INFORMAR
            'LA MISMA EN EL DETALLE COMO UN MEDIO DE PAGO MAS, DEBIENDO ESTE ESTAR ASOCIADO A LA CUENTA DE Cta. Ret. de Ganancias.
            'EL SISTEMA GENERA UN ARCHIVO .TXT CON LOS RECIBOS QUE EN SU DETALLE CONTENGAN ALGUN MOVIMIENTO CON DICHA CUENTA CONTABLE.
            'INGRESA TODOS ESTOS CASOS EN DICHO ARCHIVO SEGUN EL FORMATO INDICADO DE LA AFIP.
            'EN EL CASO QUE EL RECIBO NO ESTE APLICADO A NINGUNA FACTURA SE VAN ESCRIBIENDO EN FRMZELECHO, AL FIN DEL PROCESO
            'AVISA DE ESTA SITUACION Y SE TIENE LA POSIBILIDAD DE VER EL DETALLE, DE TODOS MODOS GENERA EL ARCHIVO CON LOS RECIBOS QUE POSEAN APLICACIONES.
            '
            'VALIDO SI EN EL DAT PROVINB ESTAN CARGADAS LAS CUENTAS CONTABLES
            CONDIX$ = GENERA_CONSULTA_SQL_CUENTAS_RETIB$ 'ESTA FUNCION DEVUELVE LA CONDICION DE LAS CUENTAS CONTABLES QUE ESTAN EN  PROVINB
            If TRIM$(CONDIX$) = "" Then
               Call COMUNI("Error en la Configuración para Ingresos Brutos.\Verifique lo Siguiente:\ \1. No Existen Cuentas Contables de Retención de IIBB.\2. Falta Jurisdicción en Algunas Provincias.")
               Exit Sub
            End If
            
            DES1$ = FETEXCOR1$(DESDE%)
            HAS1$ = FETEXCOR1$(HASTA%)
            If GeneraParaReporte% = 0 Then
                ARCHIX$ = SELECCION_Archivo$("TXT", CommonDialog1)
                If TRIM$(ARCHIX$) = "" Then GoTo 99
            Else
                ARCHIX$ = LUCOM$ & "REPORTE.TXT"
            End If
            LongitudRegistro% = 128
            NARCHI2% = FILEOPEN%(ARCHIX$, 2, LongitudRegistro%) ' ABRE EL ARCHIVO DE ESCRITURA
            GENERO_ARCHIVO% = 0
            '
            Call SETULTREG("LIREPEIB", 0)
            Call BLANARCH("LIREPEIB")
            '
            Campos$ = "Fecha/D8;Documento/A18;Agente/A32;Constancia/A24;Importe/F12.2;Informe/A64"
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
            If ANTOT < 5500 Then ANTOT = 5500
            FRMZELECHO.Width = 12000
            FRMZELECHO.Caption = "Errores Encontrados"

            CONDI$ = " FECHEMISI >= '" & DES1$ & "'"
            CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HAS1$ & "' "
            CONDI$ = CONDI$ + " AND TIPOMOVIM = 'CHRECI' "
            CONDI$ = CONDI$ + " AND " & CONDIX$
            SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE "
            SQLX$ = SQLX$ & CONDI$
            SQLX$ = SQLX$ + " ORDER BY FECHEMISI ASC "
            
            Text1 = SQLX$
            FIN& = CantRegistros("CAJABANC", CONDI$)
            
            Set rs = READSET(SQLX$)
            With rs
              Do While Not .EOF
                    AV& = AV& + 1
                    Call TRACE(24, AV&, FIN&)
                    
                    FalloNoExisteCuentaEnProvincia% = 0
                    FalloCuitAgente% = 0
                    FalloNoExisteAplicacionDeRecibo% = 0
                    FalloComprobanteAplicadoImporte% = 0
                    FalloComprobanteAplicadoConstancia% = 0
                    FalloComprobanteAplicadoPuntoVenta% = 0
                    FalloComprobanteAplicadoFactura% = 0
                    FalloComprobanteAplicadoTipoComprobante% = 0
                    
              
                    TIPOMOV$ = TRIM$(SAFETEXT$(!TIPOMOVIM))
                 
                    RECIBOX$ = SAFETEXT$(!codicom_lar)
                    NC = XVALO(!NUCLIEN)
                    ' *** OBTIENE CODIGO DE JURISDICCION ASOCIADO A LA CTA CONTABLE DEL MEDIO DE COBRANZA
                    ' *** DE LA RETENCION DE IIBB 'XXXX'
                    CuentaContableMedioCobranza% = XVALO(!CUECONTAI)
                    RegistroProvincia$ = FILTERGETRECORD$("PROVINB", 5, MKI$(CuentaContableMedioCobranza%))
                    If TRIM$(RegistroProvincia$) = "" Then GoTo 20 ' VERIFICA SI EXISTE CTA CONTABLE DE RETENCION EN TABLA DE PROVINCIAS
                                                                   ' SI NO EXISTE, SE ASUME QUE SE TRATA DE UN CHEQUE U OTRO MEDIO DE COBRANZA
                    ' FalloNoExisteCuentaEnProvincia% = 1
                                                                                    ' EN TABLA DE PROVINCIAS
                    CodigoDeJurisdiccion$ = CStr(CVI(Mid$(RegistroProvincia$, 37, 2)))
                    nombreProvincia$ = TRIM$(Mid$(RegistroProvincia$, 2, 25))
                    ' *** VALIDA CUIT DEL AGENTE DE RETENCION
                    CUITX$ = CuitCli(NC)
                    PIV% = PIVACLI(NC)
                    If PIV% <> 5 And PIV% <> 7 And PIV% <> 0 Then ' OBVIA MONOTRIBUTISTAS, EXPORTACIONES Y CONSUMIDORES FINALES
                        Valid% = VALICUIT%(CUITX$)
                        If Valid% < 1 Then
                          FalloCuitAgente% = 1
                        End If
                        'SI ES VALIDO VERIFICA CON LA FUNCION QUE TENGA LOS SEPARADORES, SI NO ES ASI LA FUNCION SE LOS AGREGA.
                        CUITX$ = FormatearCuitValido$(CUITX$) ' LE AGREGA FORMATO CON LOS SEPARADORES - EN CASO QUE NO LOS TENGA
                    End If
                    ' *** OBTIENE DATOS PARA EL ARMADO DEL ARCHIVO AFIP
                    FE% = CVINT(!FECHEMISI)
                    FEMIS$ = Format(CVDAT(FE%), "DD/MM/YYYY")
                    SUCURSAL$ = "0001"
                    CONDI$ = "CoDocAp_Lar=  '" & RECIBOX$ & "'  AND NUCLIEN = " & NC & "AND SUBSTRING(CoDocOr_Lar,1,2) <> 'AS' ORDER BY ImpoApli DESC "
                    APLICACIONRECIBO$ = TRIM$(VALOBADA("APLICOBRA", "CoDocOr_Lar", CONDI$, "NOMESS"))
                    If APLICACIONRECIBO$ = "" Then
                        APLICACIONRECIBO$ = Mid$(RECIBOX$, 1, 2) & Mid$(RECIBOX$, 5) & Mid$(RECIBOX$, 3, 2)
                        If PIV% = 1 Or PIV% = 2 Or PIV% = 3 Then APLICACIONRECIBO$ = REPLACAR$(APLICACIONRECIBO$, "X", "A")
                        If PIV% = 0 Or PIV% = 4 Or PIV% = 6 Then APLICACIONRECIBO$ = REPLACAR$(APLICACIONRECIBO$, "X", "B")
                        If PIV% = 5 Or PIV% = 7 Then APLICACIONRECIBO$ = REPLACAR$(APLICACIONRECIBO$, "X", "E")
                        'FalloNoExisteAplicacionDeRecibo% = 1
                    End If
                    IMPORTEAPLIC# = XVALO(!idebecomp) - XVALO(!ihabecomp)
                    LET1$ = Right(APLICACIONRECIBO$, 1)
                    
                    NUMCONSTANCIA$ = SAFETEXT$(!DESCRIMOV)
                    NUMCONSTANCIA$ = TRIM$(Mid$(NUMCONSTANCIA$, 19, 12))
                    ' REEMPLAZA EN EL NUMERO DE CONSTANCIA LOS CARACTERES INVALIDOS POR '0'
                    ' LOS CARACTERES INVALIDOS LUEGO PROVOCAN UN ERROR AL IMPORTAR DESDE EL APLICATIVO SIFERE
                    For INDCON% = 1 To Len(NUMCONSTANCIA$)
                        ASCIICAR% = Asc(Mid$(NUMCONSTANCIA$, INDCON%, 1))
                        If ASCIICAR% < 48 Or (ASCIICAR% > 57 And ASCIICAR% < 65) Or (ASCIICAR% > 90 And ASCIICAR% < 97) Or ASCIICAR% > 122 Then
                            NUMCONSTANCIA$ = REPLACAR$(NUMCONSTANCIA$, Mid$(NUMCONSTANCIA$, INDCON%, 1), "0")
                        End If
                    Next INDCON%
                    
                    PVTA1$ = Mid$(APLICACIONRECIBO$, 4, 4)
                    nufact$ = Mid$(APLICACIONRECIBO$, 9, 8)
                    TICOMP$ = ""
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "FC-" Then TICOMP$ = "F"
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "NC-" Then TICOMP$ = "C"
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "RB-" Then TICOMP$ = "R"
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "ND-" Then TICOMP$ = "D"
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "DP-" Then TICOMP$ = "C"
                           If UCase$(Left$(APLICACIONRECIBO$, 3)) = "CP-" Then TICOMP$ = "D"
                    NUMECOMPROBANTE% = 0
                    If IMPORTEAPLIC# < 0.005 Then FalloComprobanteAplicadoImporte% = 1
                    If TRIM$(NUMCONSTANCIA$) = "" Then FalloComprobanteAplicadoConstancia% = 1
                    If TRIM$(PVTA1$) = "" Then FalloComprobanteAplicadoPuntoVenta% = 1
                    If TRIM$(nufact$) = "" Then FalloComprobanteAplicadoFactura% = 1
                    If TICOMP$ = "" Then FalloComprobanteAplicadoTipoComprobante% = 1
                    ' *** VERIFICA QUE EL DOCUMENTO APLICADO SE ENCUENTRE EN CAJABANC
                    ' *** VERIFICA QUE LA FACTURA NO ESTE EN (-1), O SEA, PENDIENTE DE OBTENER EL CAE
                    ' *** VERIFICA SI HUBO FALLO EN LA OBTENCION DE LA JURISDICCION O EN EL CUIT DEL AGENTE
                    If APLICACIONRECIBO$ <> "" Then
                        NUMECOMPROBANTE% = XVALO(VALOBADA("CAJABANC", "NUMECOMP", "NUMECOMP < 0 AND CODICOM_LAR = '" & APLICACIONRECIBO$ & "'"))
                        If NumeroComprobante% < 0 Then FalloFacturaSinCae% = 1
                    End If
                    '
                    If FalloNoExisteCuentaEnProvincia% = 1 Or FalloCuitAgente% = 1 Or FalloNoExisteAplicacionDeRecibo% = 1 Or FalloComprobanteAplicado% = 1 Or FalloFacturaSinCae% = 1 Then
                       JX& = JX& + 1
                       TXTERROR$ = ""
                       FRMZELECHO.MSF2.Rows = JX& + 1
'                       FRMZELECHO.msf2.TextMatrix(JX&, 1) = FECHATEX$(CVINT(!FechEmisi))
'                       FRMZELECHO.msf2.TextMatrix(JX&, 2) = RECIBOX$
'                       FRMZELECHO.msf2.TextMatrix(JX&, 3) = RASOCLI$(NC)
'                       FRMZELECHO.msf2.TextMatrix(JX&, 4) = NUMCONSTANCIA$
'                       FRMZELECHO.msf2.TextMatrix(JX&, 5) = FORMATNUM$(IMPORTEAPLIC#, "F12.2")
                       ' ARMA MENSAJE DE ERROR
                       Separador1$ = "": TXTERROR$ = ""
                       If FalloNoExisteCuentaEnProvincia% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "No Existe Provincia Vinculada con Cuenta Contable de Retención"
                       If FalloCuitAgente% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Cuit Inválido del Agente de Retención."
                       If FalloNoExisteAplicacionDeRecibo% = 1 Then TXTERROR$ = "El Recibo No Se Encuentra Aplicado": Separador1$ = " - "
                       If FalloNoExisteAplicacionDeRecibo% = 0 Then
                            If FalloComprobanteAplicadoImporte% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Importe Retenido Menor a Cero"
                            If FalloComprobanteAplicadoConstancia% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Numero de Constancia"
                            If FalloComprobanteAplicadoPuntoVenta% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Punto de Venta"
                            If FalloComprobanteAplicadoFactura% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Numero de Factura"
                            If FalloComprobanteAplicadoTipoComprobante% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Falta Tipo de Comprobante"
                            If FalloFacturaSinCae% = 1 Then TXTERROR$ = TXTERROR$ + Separador1$ & "Factura Aplicada Sin CAE": Separador1$ = " - "
                        End If
                       
                       FRMZELECHO.MSF2.TextMatrix(JX&, 6) = TXTERROR$
                       
                    End If
                    ' *** ARMA REGISTRO DE ARCHIVO
                    GENERO_ARCHIVO% = 1
                     If Len(CodigoDeJurisdiccion$) < 3 Then CodigoDeJurisdiccion$ = String$(3, "0") & CodigoDeJurisdiccion$: CodigoDeJurisdiccion$ = Right(CodigoDeJurisdiccion$, 3)
                    Call REPLA(TX$, CodigoDeJurisdiccion$, 1, 3)
                     If Len(CUITX$) < 13 Then CUITX$ = String$(13, " ") & CUITX$: CUITX$ = Right(CUITX$, 13)
                    Call REPLA(TX$, CUITX$, 4, 13)
                     If Len(FEMIS$) < 10 Then FEMIS$ = String$(10, "0") & FEMIS$: FEMIS$ = Right(FEMIS$, 10)
                    Call REPLA(TX$, FEMIS$, 17, 10)
                    Call REPLA(TX$, SUCURSAL$, 27, 4)
                      If Len(NUMCONSTANCIA$) < 16 Then NUMCONSTANCIA$ = String$(16, "0") & NUMCONSTANCIA$: NUMCONSTANCIA$ = Right(NUMCONSTANCIA$, 16)
                    Call REPLA(TX$, NUMCONSTANCIA$, 31, 16)
                    Call REPLA(TX$, TICOMP$, 47, 1)
                    Call REPLA(TX$, LET1$, 48, 1)
                      If Len(PVTA1$) < 4 Then PVTA1$ = String(4, "0") & PVTA1$: PVTA1$ = Right$(PVTA1$, 4)
                      If Len(nufact$) < 16 Then nufact$ = String(16, "0") & nufact$: nufact$ = Right$(nufact$, 16)
                      APLICACIONRECIBO$ = PVTA1$ & nufact$
                    Call REPLA(TX$, APLICACIONRECIBO$, 49, 20)
                    IMPORTEAPL$ = TRIM$(FORMATNUM$(IMPORTEAPLIC#, "F11.2"))
                    
                    ' FORMATEO SI TIENE DECIMALES HAY QUE QUITARLE EL PUNTO
                    If InStr(IMPORTEAPL$, ".") > 0 Then
                      IMPORTEAPL$ = REPLACAR(IMPORTEAPL$, ".", ",")
                    Else
                      IMPORTEAPL$ = IMPORTEAPL$ & ",00"
                    End If
                    If Len(IMPORTEAPL$) < 11 Then IMPORTEAPL$ = String$(11, "0") & IMPORTEAPL$: IMPORTEAPL$ = Right(IMPORTEAPL$, 11)
                    
                    Call REPLA(TX$, IMPORTEAPL$, 69, 11)
                    
                    Print #NARCHI2%, TX$
                    TTXX$ = REGBLAN$("LIREPEIB")
                    Call REPLA(TTXX$, MKI(FE%), 1, 2)
                    Call REPLA(TTXX$, RECIBOX$, 3, 18)
                    Call REPLA(TTXX$, rasocli$(NC), 21, 48)
                    Call REPLA(TTXX$, NUMCONSTANCIA$, 69, 24)
                    Call REPLA(TTXX$, MKD$(IMPORTEAPLIC#), 93, 8)
                    Call REPLA(TTXX$, Space(91), 101, 91)
                    If TXTERROR$ <> "" Then
                        TextoError$ = TXTERROR$
                        Call FRATEXTO(TextoError$, 60)
                        For KKII% = 1 To CARETEX%
                            Call REPLA(TTXX$, RETEX(KKII%), 101, 60)
                            Call REGAPP("LIREPEIB", TTXX$)
                        Next KKII%
                    Else
                        Call REGAPP("LIREPEIB", TTXX$)
                    End If
                    
20               .MoveNext
              Loop
            End With
            '
            ' *** AGREGA RETENCIONES DESDE LOS CUPONES DE LAS TARJETAS
30          Call CargaAsientoConciTarjeta(CONDIX$, DESDE%, HASTA%, ARCHIX$, OK%, RET_TX$)
            If OK% < 0 Then
                Close #NARCHI2%
                GoTo 40
            End If
            
            If TRIM$(RET_TX$) <> "" Then
                Print #NARCHI2%, RET_TX$
            End If

            CantidadDeRetenciones& = LOF(NARCHI2%) / LongitudRegistro%
            Close #NARCHI2%
            Screen.MousePointer = 1
            
            ' *** INFORMA LOS ERRORES ENCONTRADOS
40          If FRMZELECHO.MSF2.Rows > 1 And GeneraParaReporte% = 0 Then
               OPX% = COMALTER%("No Se Han Exportado las Retenciones\Debido a Que Se Han Detectado Errores en los Datos.\\Cancelar\Ver Detalle")
               If OPX% <= 1 Then Close #NARCHI2%: GoTo 99
               If OPX% = 2 Then
                    Close #NARCHI2%
                    Call CIERRARCH("LIREPEIB")
                    Call FILESORT("LIREPEIB", "", 1, 1)
                    Call FINAL("*LISRETIB")
               End If
               
               On Error Resume Next
               Call FILEKILL(ARCHIX$)
               If Err > 0 Then
                   Call COMUNI("No Se Ha Podido Eliminar el Archivo:\" & ARCHIX$ & "\Realice Las Correcciones Informadas en el Detalle y Vuelva a Intentar")
               End If
               On Error GoTo 0
               Exit Sub
            End If
            
            ' *** CONTINUA PORQUE LOS DATOS NO TIENEN ERRORES
            
            ' *** SI ES REPORTE, SALE DEL FLUJO
            If GeneraParaReporte% = 1 Then GoTo 99
            
            ' *** ESTA EXPORTANDO, AVISA EL RESULTADO DE LA EXPORTACION
            ARCHIX$ = REPLACAR$(ARCHIX$, "\", "/")
            If CantidadDeRetenciones& > 0 Then
                Call COMUNI("Archivo Generado Correctamente en '" & ARCHIX$ & "'")
            Else
                Call COMUNI("No Se Han Encontrado Retenciones Para el Período Elegido")
            End If
                      
99          Screen.MousePointer = 1
            Call CIERRARCH("LIREPEIB")
            Call FILESORT("LIREPEIB", "", 1, 1)
End Sub


Sub Command3_Click()
   Command3.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   'SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
   
   SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
   SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'FC' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   'Set CABATEMP = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.CursorLocation = adUseClient
26 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
'   sqlx$ = "SELECT * FROM CAJABANC "
'   sqlx$ = sqlx$ + "where FechConta >= " & DESDAT & " "
'   sqlx$ = sqlx$ + "AND FechConta <= " & HASDAT & " "
'   Set CABASELE = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
   '
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !codicom_lar
       NCLI = !NUCLIEN
       FEFAXI% = CVINT(!fechconta)
       FECHAD = FETEXCOR1$(FEFAXI%)
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
       SQLX1$ = "SELECT * FROM CAJABANC WHERE FechConta = '" & FECHAD & "' "
       SQLX1$ = SQLX1$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
       SQLX1$ = SQLX1$ + "AND NuClien = " & NCLI & " "
       Dim CABASELE As ADODB.Recordset
       Set CABASELE = New ADODB.Recordset
       CABASELE.CursorLocation = adUseClient
22     On Error Resume Next
       CABASELE.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 22
       End If
       On Error GoTo 0
       '
       With CABASELE
         Do While Not .EOF   '.FindFirst sqlx$
25       ' If .NoMatch = False Then
           TIMOV$ = TRIM$(UCase$(!TIPOMOVIM))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
         '  .FindNext sqlx$
         .MoveNext '  GoTo 25
         Loop
         'End If
       End With
       CABASELE.Close
       Set CABASELE = Nothing
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKS$(NCLI), 9, 4)
       Call REPLA(ZZ$, MKD$(INETO#), 13, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 21, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 29, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 37, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 45, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 53, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Sub Command30_Click()
    Command30.Enabled = False
'\\\OT-5-0354-FG-03/06/05///

    Dim TOTVENTA#(999999)
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    DESDAT = FECHATEX(Des%)
    HASDAT = FECHATEX(Has%)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    Screen.MousePointer = 11
    CLIMAX = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
         NCLI = !NUCLIEN
         If NCLI > 0 Then
           If NCLI > CLIMAX Then CLIMAX = NCLI
           TOTVENTA#(NCLI) = TOTVENTA#(NCLI) + (!ihabecomp - !idebecomp)
         End If
         .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    '
    For II = 1 To CLIMAX
      TOTAGEN# = TOTAGEN# + TOTVENTA#(II)
      If II > 999999 Then Exit For ' AGREGADO POR QUE SI EXISTE CLIENTE 999999
    Next II                       ' CUANDO LLEGA AL NEXT I% DESBORDA LA VARIABLE
    '
    List1.Clear
    For NCLI = 1 To 999999
       If Abs(TOTVENTA#(NCLI)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCLI)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCLI) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKS$(NCLI), 3, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCLI)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCLI
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NCLI = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKS$(NCLI), 3, 4)
        Call REPLA(Z$, MKD$(TONEVEN#), 7, 8)
        Call REPLA(Z$, MKS$(PORPAR), 15, 4)
        Call REPLA(Z$, MKS$(PORACU), 19, 4)
        Call GRAREG("RANKCLN", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLN", J&)
    Call CIERRARCH("RANKCLN")
    Call HEADERS("RANKCLN", "")
    Call HEADERS("RANKCLN", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '

    If Option1.Value = True Then
         Call FINAL("*RANKCLN")
       Else
         Call FINAL("*LRANCLN")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Command30.Enabled = True
    '
'\\\OT-5-0354-FG-FIN///
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   'JANDY SQL
   
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   'sqlx$ = sqlx$ + "where FechConta >= " & DESDAT & " "
   'sqlx$ = sqlx$ + "AND FechConta <= " & HASDAT & " "
   SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
   SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'ND' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   'Set CABATEMP = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.CursorLocation = adUseClient
20 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 20
   End If
   On Error GoTo 0
   '
'   SQLX$ = "SELECT * FROM CAJABANC "
'   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
'   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
'   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !codicom_lar
       NCLI = !NUCLIEN
       FEFAXI% = CVINT(!fechconta)
       FECHAD = FETEXCOR1$(FEFAXI%)
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
'       SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
'       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
'       SQLX$ = SQLX$ + "AND NuClien = " & NCLI & " "
'       '
        SQLX1$ = "Select * from CajaBanc where FechConta = '" & FECHAD & "' "
        SQLX1$ = SQLX1$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX1$ = SQLX1$ + "AND NuClien = " & NCLI & " "
        Dim CABASELE As ADODB.Recordset
        Set CABASELE = New ADODB.Recordset
        CABASELE.CursorLocation = adUseClient
21      On Error Resume Next
        CABASELE.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 21
        End If
        On Error GoTo 0
       '
       With CABASELE
       '  .FindFirst SQLX$
25     '   If .NoMatch = False Then
         Do While Not (.EOF)
           TIMOV$ = TRIM$(UCase$(!TIPOMOVIM))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
           .MoveNext
         Loop
       '    .FindNext SQLX$
       '    GoTo 25
       '  End If
       End With
       CABASELE.Close
       Set CABASELE = Nothing
       
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKS$(NCLI), 9, 4)
       Call REPLA(ZZ$, MKD$(INETO#), 13, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 21, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 29, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 37, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 45, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 53, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   CABATEMP.Close
   Set CABATEMP = Nothing
   
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   If VALIMESAN% < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
   'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
   SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
   SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
   SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'NC' "
   SQLX$ = SQLX$ + "Order by FechConta ASC, CodiCom_Lar ASC "
   
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.CursorLocation = adUseClient
20 On Error Resume Next
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 20
   End If
   On Error GoTo 0
   'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   
'   SQLX$ = "SELECT * FROM CAJABANC "
'   SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
'   SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
'   Set CABASELE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   JJ& = 0: REBLA$ = REGBLAN$("DETACCX")
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     KU& = 0
     Do While Not .EOF
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       ZZ$ = REBLA$
       NUMEFA$ = !codicom_lar
       NCLI = !NUCLIEN
       FEFAXI% = CVINT(!fechconta)
       FECHAD = FETEXCOR1$(FEFAXI%)
       NUMECO = !NumeComp
       '
       INETO# = 0
       IMIVA# = 0
       IVRNI# = 0
       RETEN# = 0
       IOTRO# = 0
       ITOTA# = !ValAbsComp
       '
       SQLX1$ = "Select * from CAJABANC where FechConta = '" & FECHAD & "' "
       SQLX1$ = SQLX1$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
       SQLX1$ = SQLX1$ + "AND NuClien = " & NCLI & " "
       Dim CABASELE As ADODB.Recordset
       Set CABASELE = New ADODB.Recordset
       CABASELE.CursorLocation = adUseClient
21     On Error Resume Next
       CABASELE.Open FILTSQL$(SQLX1), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 21
       End If
       On Error GoTo 0
       '
       With CABASELE
       Do While Not .EOF
         '.FindFirst SQLX$
25       ' If .NoMatch = False Then
           TIMOV$ = TRIM$(UCase$(!TIPOMOVIM))
           IPARD# = !ValAbsComp
           If TIMOV$ = "VENTA" Or TIMOV$ = "EMBAL" Then
               INETO# = INETO# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) <> "RNI" Then
               IMIVA# = IMIVA# + IPARD#
             ElseIf TIMOV$ = "IVA" And UCase$(!OpciMovim) = "DEFI" And UCase$(!VARIMOVIM) = "RNI" Then
               RVRNI# = IVRNI# + IPARD#
             ElseIf UCase$(!VARIMOVIM) Like "PER*" Then
               RETEN# = RETEN# + IPARD#
           End If
          ' .FindNext SQLX$
          ' GoTo 25
         'End If
             .MoveNext
         Loop
       End With
       CABASELE.Close
       Set CABASELE = Nothing
       '
       IOTRO# = ITOTA# - INETO# - IMIVA# - IVRNI# - RETEN#
       '
       Call REPLA(ZZ$, Chr$(2), 1, 1)
       Call REPLA(ZZ$, MKI$(FEFAXI%), 3, 2)
       Call REPLA(ZZ$, MKS$(NUMECO), 5, 4)
       Call REPLA(ZZ$, MKS$(NCLI), 9, 4)
       Call REPLA(ZZ$, MKD$(INETO#), 13, 8)
       Call REPLA(ZZ$, MKD$(IMIVA#), 21, 8)
       Call REPLA(ZZ$, MKD$(IVRNI#), 29, 8)
       Call REPLA(ZZ$, MKD$(RETEN#), 37, 8)
       Call REPLA(ZZ$, MKD$(IOTRO#), 45, 8)
       Call REPLA(ZZ$, MKD$(ITOTA#), 53, 8)
       JJ& = JJ& + 1
       Call GRAREG("DETACCX", ZZ$, JJ&)
       '
     .MoveNext
     Loop
   End With
   CABATEMP.Close
   Set CABATEMP = Nothing
   
   Call SETULTREG("DETACCX", JJ&)
   '
   Call HEADERS("DETACCX", "")
   Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(MESAN))
   Call HEADERS("DETACCX", "Rubro: Facturacion en Cuenta Corriente / Contado")
   Screen.MousePointer = 1
   If Option1.Value = True Then
        Call FINAL("*CODECCX")
     ElseIf Option2.Value = True Then
        Call FINAL("*LIDECCX")
   End If
99 Screen.MousePointer = 1
   Command5.Enabled = True
End Sub

Sub Command6_Click()
'////*******///////
    '
    Command6.Enabled = False
    If VALIMESAN% > 0 Then
       Call REMECUEN
    End If
    Command6.Enabled = True
    '
End Sub
'
Sub REMECUEN()
    '
    FED% = FECHANUM(DESDAT)
    FEH% = FECHANUM(HASDAT)
    DESDAT = FETEXCOR1$(FED%)
    HASDAT = FETEXCOR1$(FEH%)
    OPX% = COMALTER%("Rango de Fechas\\Mes Activo\Selección Rango de Fechas")
    If OPX% = 2 Then
      Call DESHASFECHA(Des%, Has%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      DESDAT = FETEXCOR1$(Des%)
      HASDAT = FETEXCOR1$(Has%)
    End If
    '
    Screen.MousePointer = 11
    Dim SALINI#(999999), SUFAC#(999999), SUCOB#(999999)
    '
    Call APERBASDAT("CABAN")
    TABORI$ = "CAJABANC"
    CLIMAX = 0
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
    
10  SQLX$ = "SELECT NuClien, SUM(IDEBECOMP)AS IDEBEX ,SUM(IHABECOMP) AS IHABEX FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta < '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "GROUP BY NuClien "
    '
    On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 20
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLI = !NUCLIEN
           If NCLI > 0 Then
             'If TICUEPRO((-1) * NCLI) = 0 Then ' solo cuentas corrientes
               If NCLI > CLIMAX Then CLIMAX = NCLI
               SALDOINI# = XVALO(!IDEBEX) - XVALO(!IHABEX)
               SALINI#(NCLI) = SALINI#(NCLI) + SALDOINI#
             'End If
           End If
         .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
20  SQLX$ = "SELECT NuClien, SUM(IDEBECOMP)AS IDEBEX1 ,SUM(IHABECOMP) AS IHABEX1 FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "GROUP BY NuClien "
    '
    On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 30
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLI = !NUCLIEN
           If NCLI > 0 Then
             'If TICUEPRO((-1) * NCLI) = 0 Then ' solo cuentas corrientes
               If NCLI > CLIMAX Then CLIMAX = NCLI
               ifac1# = XVALO(!IDEBEX1) - XVALO(!IHABEX1)
               SUFAC#(NCLI) = SUFAC#(NCLI) + ifac1#
             'End If
           End If
         .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
30  SQLX$ = "SELECT NuClien, SUM(IDEBECOMP)AS IDEBEX2 ,SUM(IHABECOMP) AS IHABEX2 FROM " + TABORI$ + " "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "GROUP BY NuClien "
    '
    On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 40
    End If
    On Error GoTo 0
    '
    With CABATEMP
       If Not .BOF And Not .EOF Then
         .MoveLast
         FIN& = .RecordCount
         .MoveFirst
         JU& = 0
         Do While Not .EOF
           JU& = JU& + 1
           Call TRACE(20, JU&, FIN&)
           NCLI = !NUCLIEN
           If NCLI > 0 Then
             'If TICUEPRO((-1) * NCLI) = 0 Then ' solo cuentas corrientes
               If NCLI > CLIMAX Then CLIMAX = NCLI
               IPAG2# = XVALO(!IHABEX2) - XVALO(!IDEBEX2)
               SUCOB#(NCLI) = SUCOB#(NCLI) + IPAG2#
             'End If
           End If
         .MoveNext
         Loop
       End If
    End With
    CABATEMP.Close
    '
40  On Error GoTo 0
    If TABORI$ = "CAJABANC" Then
      TABORI$ = "CABANVIE"
      GoTo 10
    End If
    Set CABATEMP = Nothing
    '
    JJ& = 0
    
    For KK = 1 To CLIMAX
    If Abs(SALINI#(KK)) >= 0.005 Or Abs(SUFAC#(KK)) >= 0.005 Or Abs(SUCOB#(KK)) >= 0.005 Then
        X$ = REGBLAN$("RESUCCD")
        Call REPLA(X$, MKS$(KK), 1, 4)
        Call REPLA(X$, MKD$(SALINI#(KK)), 5, 8)
        Call REPLA(X$, MKD$(SUFAC#(KK)), 13, 8)
        Call REPLA(X$, MKD$(SUCOB#(KK)), 21, 8)
        Call REPLA(X$, MKD$(SALINI#(KK) + SUFAC#(KK) - SUCOB#(KK)), 29, 8)
        JJ& = JJ& + 1
        Call GRAREG("RESUCCD", X$, JJ&)
       End If
       '\\\OT-08-0921-OD-02/10/08///
       'SE AGREGO ESTE IF POR QUE EN DOSIVAC EXISTE UN CLIENTE Nº 32767 Y AL SALIR DEL
       'NEXT SUMABA UNO MAS Y PRESENTABA UN ERROR DE DESBORDAMIENTO DE LA VARIABLE INDICE 'K'
       If KK = CLIMAX Then Exit For
       '\\\OT-08-0921-OD-02/10/08///
    Next KK
    Call SETULTREG("RESUCCD", JJ&)
    '
    Call HEADERS("RESUCCD", "")
    Call HEADERS("RESUCCD", "Corresponde a: " + TRIM$(CORRESP$))
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
         Call FINAL("*RESUCCD")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCD")
    End If
    '
End Sub
'


Sub Command7_Click()
    '
    '    '////***ANTIGUO***/////
'    Screen.MousePointer = 11
'    LU$ = LUDAT$
'    FileCopy LU$ + "DETOTRV.DAT", LU$ + "DETACCX.DAT"
'    C1$ = TRIM$(REMEVE04.MESAN)
'    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
'    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
'    C1$ = TRIM$(REMEVE00.MESAN)
'    C2$ = TRIM$(REMEVE00.Text4(0).TEXT)
'    C3$ = TRIM$(REMEVE00.Text4(1).TEXT)
    '////****FIN****/////
    '
    '/////****NUEVO****/////
    Command7.Enabled = False
    Screen.MousePointer = 11
    LU$ = LUDAT$
    C1$ = TRIM$(REMEVE07.MESAN)
    '
    If VALIMESAN% < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call APERBASDAT("CABAN")
    'jandy sql
    
    
    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta  <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    J1& = 0
    REBLA1$ = REGBLAN$("DETACCX")
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      Do While Not .EOF
        ZZ$ = REBLA1$
        NCLI = !NUCLIEN
        FECHAD = Int(CDbl(!fechconta))
        INETO# = XVALO(!idebecomp) - (!ihabecomp)
        ITOTA# = INETO#
        'INETO# = !ValAbsComp
        'ITOTA# = !ValAbsComp
        '
        Call REPLA(ZZ$, Chr$(4), 1, 1)
        Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
        Call REPLA(ZZ$, MKS$(!NumeComp), 5, 4)
        Call REPLA(ZZ$, MKS$(NCLI), 9, 4)
        Call REPLA(ZZ$, MKD$(INETO#), 13, 8)
        Call REPLA(ZZ$, MKD$(ITOTA#), 53, 8)
        J1& = J1& + 1
        Call GRAREG("DETACCX", ZZ$, J1&)
        '
        .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    Call SETULTREG("DETACCX", J1&)
    '////*****FINNUEVO***/////

    Call HEADERS("DETACCX", "")
    Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Call HEADERS("DETACCX", "Rubro: Otros Movimientos de Cuenta Corriente")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODECCX")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDECCX")
    End If
    '
99  Command7.Enabled = True
    Screen.MousePointer = 1
 End Sub

Sub Command8_Click()
    Command8.Enabled = False
    '
    If CONTROL$("FACTURA", "SALIDA") = "CFISCAL" Then
          ' INICIALIZAR CONTROLADOR FISCAL
          Call FORHASAR.REPORTES
          GoTo 99
       ElseIf CONTROL$("FACTURA", "SALIDA") = "IFISCAL" Then
          Call FORHAS320.REPORTES
          GoTo 99
    End If
    '
    If EXISTE%("LCOMIS07.EXE") > 0 Then
      OPXX% = ALTERNA%("Liquidación de Comisiones\Otros Accesos Directos")
      If OPXX% = 1 Then
          Call CONECRUN("LCOMIS07", "Liquidación de Comisiones")
          GoTo 99
        ElseIf OPXX% = 2 Then
          GoTo 20
        Else
          GoTo 99
      End If
    End If
    If InStr(EMPREAC$, "DOSIVA") > 0 Then
      Call CONECRUN("LICODO07", "Liquidación de Comisiones")
      Command8.Enabled = True
      Exit Sub
    End If
    '
20  Call ACCDIR("COGEVEN")
    '
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture9)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If    '
    
    If EXISTE%("LCOMIS07.EXE") > 0 Then
       mnuLiquidaComis.Visible = True
    End If
    
    mnuFactAist.Visible = False
    If InStr(EPAC$, "EYM") > 0 Then mnuFactAist.Visible = True
    
    If InStr(1, EPAC$, "MEDITEA") < 1 Then
        mnuFacturaNetavsCmv.Visible = False
        mnuComparativaPorAnio.Visible = False
        mnuEstadisticasConfiguraciones.Visible = False
        mnuEquiposPorMes.Visible = False
        mnuVentasPorProvincia.Visible = False
        mnuVentasPorCanales.Visible = False
        mnuEgresos.Visible = False
        mnuCashFlow.Visible = False
    End If
    
End Sub

Private Sub mnuAnalisisFacturacionPedidos_Click()

    Dim rp As New ReporteGrilla
    
    Dim ArticulosFacturadosPedidosAdeudados(32767) As Double
    Dim ArticulosFacturadosPedidosMesEnCurso(32767) As Double
    Dim ArticulosPedidosPendientes(32767) As Double

    Dim rst1 As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim rst3 As ADODB.Recordset
    
    If VALIMESAN% < 1 Then Exit Sub
    
    VDEF$ = MKI$(DIANTE(HOY(HO$), 730))
    Call OBJPLA$("ENTRAFECHA1", VDEF$)
    FechaCortePedidosPendientes% = CVI(VDEF$)

    IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
    CondicionSql$ = ""
    If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
        ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
        Dim RstCodigos As ADODB.Recordset
        Set RstCodigos = READSET(ConsultaSql$)
        
        With RstCodigos
            Do While Not .EOF
                CodigoProducto$ = SAFETEXT$(!CAMPO2)
                If CondicionSql$ = "" Then
                    CondicionSql$ = "(m.cod_inte='" & CStr(CODINT(CodigoProducto$)) & "'"
                Else
                    CondicionSql$ = CondicionSql$ & " or m.cod_inte='" & CStr(CODINT(CodigoProducto$)) & "'"
                End If
                
                .MoveNext
            Loop
        End With
        If CondicionSql$ <> "" Then CondicionSql$ = " and " & CondicionSql$ & ")"
        On Error Resume Next
        RstCodigos.Close
        Set RstCodigos = Nothing
        On Error GoTo 0
    End If


    ConsultaSql$ = "select v.cod_inte, sum(salida) salida, sum(ingreso) ingreso from peddeta p with(nolock) inner join" _
    & " (select c.codicom_lar, m.doprilar, m.nupedcli, m.numeordped, m.cod_inte, m.cantsalid salida, m.cantingre ingreso from cajabanc c with(nolock) inner join movisto m with(nolock) on (c.codicom_lar=m.doseclar)" _
    & " where c.tipomovim='fven' and c.fechconta>='" & DESDAT & "' and c.fechconta<='" & HASDAT & "' and m.codimovi='rt' and m.cod_inte>0 and m.nupedcli>0 " & CondicionSql$ & ") v" _
    & " on (p.nroped=v.nupedcli and p.codiint=v.cod_inte and p.nuoritem=v.numeordped)" _
    & " where p.fechemis<'" & DESDAT & "'" _
    & " group by v.cod_inte"
    
    Set rst1 = READSET(ConsultaSql$)
    With rst1
        Do While Not .EOF
            
            ArticulosFacturadosPedidosAdeudados(!COD_INTE) = ArticulosFacturadosPedidosAdeudados(!COD_INTE) + (!SALIDA - !INGRESO)
                        
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst1.Close
    Set rst1 = Nothing
    On Error GoTo 0
    
    ConsultaSql$ = "select v.cod_inte, sum(salida) salida, sum(ingreso) ingreso from peddeta p with(nolock) inner join" _
    & " (select c.codicom_lar, m.doprilar, m.nupedcli, m.numeordped, m.cod_inte, m.cantsalid salida, m.cantingre ingreso from cajabanc c with(nolock) inner join movisto m with(nolock) on (c.codicom_lar=m.doseclar)" _
    & " where c.tipomovim='fven' and c.fechconta>='" & DESDAT & "' and c.fechconta<='" & HASDAT & "' and m.codimovi='rt' and m.cod_inte>0 and m.nupedcli>0 " & CondicionSql$ & ") v" _
    & " on (p.nroped=v.nupedcli and p.codiint=v.cod_inte and p.nuoritem=v.numeordped)" _
    & " where p.fechemis>='" & DESDAT & "' and p.fechemis<='" & HASDAT & "'" _
    & " group by v.cod_inte"
    
    Set rst2 = READSET(ConsultaSql$)
    With rst2
        Do While Not .EOF
            
            ArticulosFacturadosPedidosMesEnCurso(!COD_INTE) = ArticulosFacturadosPedidosMesEnCurso(!COD_INTE) + (!SALIDA - !INGRESO)
                        
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst2.Close
    Set rst2 = Nothing
    On Error GoTo 0
    
    CondicionSql$ = REPLACAR(CondicionSql$, "m.cod_inte", "codiint")
    ConsultaSql$ = "select codiint, sum(canped) pedido, sum(cantent) entregado from peddeta with(nolock)" _
    & " where Status >=0 and status < 8 " & CondicionSql$ _
    & " and (((canped - cantent) >= 0.05) or (tipoped = 'ABIERTO' and finvigen >= getdate()))" _
    & " and fechemis>'" & FETEXCOR1$(FechaCortePedidosPendientes%) & "'" _
    & " group by codiint"
    
    Set rst3 = READSET(ConsultaSql$)
    With rst3
        Do While Not .EOF
            
            ArticulosPedidosPendientes(!CODIINT) = ArticulosPedidosPendientes(!CODIINT) + (!PEDIDO - !ENTREGADO)
                        
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst3.Close
    Set rst3 = Nothing
    On Error GoTo 0
    
    
    rp.EncabezadoDeColumnas = "Codigo/A24;Descripción/A48;Ventas Ped.Adeudados/F25.2;Ventas Ped.Actual/F25.2;Ped.Pend/F12.2"
    rp.TITULO = "Analisis de Facturacion de Pedidos"
       
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, rp.EncabezadoDeColumnas, GRILLAGRAL)
    For i% = 1 To 32766
        Call TRACE(24, CLng(i%), 32766)
        If ArticulosFacturadosPedidosAdeudados(i%) + ArticulosFacturadosPedidosMesEnCurso(i%) + ArticulosPedidosPendientes(i%) > 0 Then
            FRMZELECHO.MSF1.Rows = FRMZELECHO.MSF1.Rows + 1
            
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 0) = "*"
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 1) = CODEXT$(i%)
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 2) = DESCRIT0$(i%)
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 3) = FORMATNUM$(ArticulosFacturadosPedidosAdeudados(i%), "F25.2")
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 4) = FORMATNUM$(ArticulosFacturadosPedidosMesEnCurso(i%), "F25.2")
            FRMZELECHO.MSF1.TextMatrix(FRMZELECHO.MSF1.Rows - 1, 5) = FORMATNUM$(ArticulosPedidosPendientes(i%), "F12.2")
            
            
        End If
    Next i%
    
    Set rp.Grilla = FRMZELECHO.MSF1
    rp.ConsultaSql = ""
    
    rp.MENU1.Visible = True
    rp.MENU1.Caption = "Imprimir"
    rp.CodigoReporte = "ANFACPE"
    
    Call rp.Mostrar

End Sub

Private Sub mnuAnalizarCorrelatividadComprobantes_Click()

    Call AnalizaCorrelativaDeComprobantes(0)

End Sub

Private Sub mnuFactAist_Click()
    Dim TOTVENTA#(32767), TOTCANT#(32767)
    '
    '**** CODIGO COPIAD DE Command15_Click *****
10  Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    J& = 0
    FiltroMovimientos% = COMALTER%("Tipos de Movimiento:\\Solo Facturas\Incluye Además Notas de Débito y Crédito")
    
25  Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
      TOTCANT#(U&) = 0
    Next U&
    '
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    REBLA$ = REGBLAN$("DETFACPE")
    '
    Call APERBASDAT("CABAN")
    '
    'VALIDAMOS QUE EXISTAN REGISTROS EN PERIODO ELEGIDO
    CONDI$ = " FechConta >= '" & DESDAT & "' AND FechConta <= '" & HASDAT & "' "
    CONDI$ = CONDI$ + "AND TipoMovim = 'VENTA' "
    '
    FIN& = CantRegistros&("CAJABANC", CONDI$)
    If FIN& < 1 Then
       Call MENSERR(24, "Sin Movimientos en el Período Elegido")
       GoTo 99
    End If
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE " & CONDI$
    '
    Campos$ = "Cta./F8;R.Social/A32;Fecha/d8;Documento/a18;Importe.Neto/F14.2;Acumulado.Neto/F16.2"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    SUMATVEN# = 0

    Dim STOTEMP1 As New ADODB.Recordset
    Dim CABATEMP  As New ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With CABATEMP
      '
      Call APERBASDAT("STOCKS")
      Screen.MousePointer = 11
      ARTMAX% = 0
      KU& = 0
      '
      Do While Not .EOF
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        TOTFAC# = ROUND(XVALO(!ihabecomp) - XVALO(!idebecomp), 2)
        NCLIE = XVALO(!Cli_Real) ' CON EL DEL MOVISTO        If Abs(TOTFAC#) >= 0.005 Then
        If Abs(TOTFAC#) >= 0.005 Then
          FF% = CVINT(!fechconta)
          FECHAPEX$ = FECHATEX$(FF%)
          NROFAC$ = SAFETEXT$(!codicom_lar)
          ImporteFacturaNeto# = XVALO(!ihabecomp) - XVALO(!idebecomp) 'POR QUE ES TIPOMOVIM = VENTA
          
          '
          If (FiltroMovimientos% = 0 Or FiltroMovimientos% = 1) And Left$(NROFAC$, 2) <> "FC" Then GoTo 30 ' solo facturas
          '
          EscribioEnGrillaGral% = 0
          SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
          SQLX$ = SQLX$ + "WHERE DoSecLar = " & "'" & NROFAC$ & "'"
          SQLX$ = SQLX$ + "AND (Nume_Clie = " & NCLIE & ") "
          SQLX$ = SQLX$ + " AND CODIMOVI = 'RT' AND COD_INTE < 1  "
          SQLX$ = SQLX$ + "  AND substring(DESCRI_LAR,1,12) LIKE 'Reparación :'"
          SQLX$ = SQLX$ + "Order by NuorItem ASC "
          '
          Set STOTEMP1 = READSET(SQLX$)
          '
          With STOTEMP1
            '
            'COEFICIENTE DE PRECIOS
            Do While Not .EOF
              If EscribioEnGrillaGral% < 1 Then
                EscribioEnGrillaGral% = 1
                SUMATVEN# = SUMATVEN# + ImporteFacturaNeto#
                GoSub AGREDET
                .MoveNext
              End If
            Loop
            '
          End With                    ' SOBRE MOVISTO
        End If
        '
30    .MoveNext
      Loop
      '
    End With                          ' SOBLE CAJABANC
    

    TX$ = "Asistencias Técnicas Facturadas (" & PERIO$ & ")"
    GRILLAGRAL.Caption = TX$
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ATECFACT", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
99  Screen.MousePointer = 1
    Exit Sub
    
AGREDET:
    'DETALLE DE FACTURACION POR CLIENTE
    J& = J& + 1
'        Campos$ = "Cta./F8;R.Social/A24;Fecha/d8;Documento/a16;P.Unit.Neto/F12.2;Importe/F12.2"

    GRILLAGRAL.GRID.Rows = J& + 1
    GRILLAGRAL.GRID.TextMatrix(J&, 1) = NCLIE
    GRILLAGRAL.GRID.TextMatrix(J&, 2) = rasocli$(NCLIE)
    GRILLAGRAL.GRID.TextMatrix(J&, 3) = FECHAPEX$
    GRILLAGRAL.GRID.TextMatrix(J&, 4) = NROFAC$
    GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(ImporteFacturaNeto#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(SUMATVEN#, "F12.2")

Return
End Sub

Private Sub mnuCashFlow_Click()

    If DERACCE%("ESTACAFL", "Estadistica - Cash Flow") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    AnioElegido% = ListaDeAnios
    If AnioElegido% = 0 Then GoTo 99
    AnioElegido% = AnioElegido% Mod 100
    d$ = "01/01/" & CStr(AnioElegido%): DesdeFecha% = FECHANUM(d$)
    H$ = "31/12/" & CStr(AnioElegido%): HastaFecha% = FECHANUM(H$)
    
    
    ' *** DIBUJO CASH FLOW
    With GRILLAGRAL
    Call CARGA_ENCABEZADO(.GRID, "Grupo/A24;Detalle/A24;C.Contable/I0;Enero/F14.2;Febrero/F14.2;Marzo/F14.2;Abril/F14.2;Mayo/F14.2;Junio/F14.2;Julio/F14.2;Agosto/F14.2;Septiembre/F14.2;Octubre/F14.2;Noviembre/F14.2;Diciembre/F14.2", FRMZELECHO)
        .GRID.Rows = .GRID.Rows + 2
        .GRID.TextMatrix(.GRID.Rows - 2, 1) = "Facturado"
        .GRID.TextMatrix(.GRID.Rows - 1, 1) = "Cobrado"
        
        IDLIST& = ObtenerIdTablaAuxiliar("GRUPOMEDCOB")
        ConsultaSql$ = "select campo2, campo4, campo5, campo6 from datasql with(nolock) where id_lista=" & CStr(IDLIST&) & " order by cast(campo1 as int)"
        Set rstGruposMediosCobranza = READSET(ConsultaSql$)
        With rstGruposMediosCobranza
            Do While Not .EOF
                MEDIO$ = SAFETEXT$(!CAMPO2)
                Argumento$ = SAFETEXT$(!CAMPO4) & SAFETEXT$(!CAMPO5)
                CuentaContable$ = SAFETEXT$(!CAMPO6)
                CContable% = CUECOINT%(CuentaContable$)
                
                GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 1) = MEDIO$
                GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 2) = Argumento$
                GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 3) = CStr(CContable%)
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rstGruposMediosCobranza.Close
        Set rstGruposMediosCobranza = Nothing
        On Error GoTo 0
    End With
    
    
    ' *** CALCULO LO FACTURADO
    If GeneraEstadisticaDeVentas("all", "all", DesdeFecha%, HastaFecha%) = False Then GoTo 99
    ConsultaSql$ = "select sum(valunit_ingreso) ingreso, sum(valunit_salida) salida, mes, anio from estadistica with(nolock) group by mes, anio"
    Set rstFacturado = READSET(ConsultaSql$)
    With rstFacturado
        Do While Not .EOF
            MES% = !MES
            
            co% = 3 + MES%
            GRILLAGRAL.GRID.TextMatrix(1, co%) = FORMATNUM$(XVALO(!SALIDA) - XVALO(!INGRESO), "F14.2")
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rstFacturado.Close
    Set rstFacturado = Nothing
    On Error GoTo 0
    
    
    ' *** CALCULO LO COBRADO
    ConsultaSql$ = "select sum(idebecomp) importe, month(fechconta) mes from cajabanc with(nolock)" _
    & " where tipomovim='chreci' and fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and fechconta<='" & FETEXCOR1$(HastaFecha%) & "'" _
    & " group by month(fechconta)"
    Set rstCobrado = READSET(ConsultaSql$)
    With rstCobrado
        Do While Not .EOF
            MES% = !MES
            
            co% = 3 + MES%
            GRILLAGRAL.GRID.TextMatrix(2, co%) = FORMATNUM$(XVALO(!IMPORTE), "F14.2")
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rstCobrado.Close
    Set rstCobrado = Nothing
    On Error GoTo 0
    
    
    ' *** CARGA COBRANZA DISCRIMINADA POR MEDIO DE COBRANZA
    ConsultaSql$ = "select cuecontai, substring(descrimov,1,18) medio, month(fechconta) mes, sum(valabscomp) importe from cajabanc with(nolock)" _
    & " where fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and tipomovim='chreci'" _
    & " group by cuecontai, substring(descrimov,1,18), month(fechconta) order by medio, cuecontai, mes"
    Set rstCobranzaPorMedios = READSET(ConsultaSql$)
    With rstCobranzaPorMedios
        Do While Not .EOF
            CContable% = XVALO(!CUECONTAI)
            MEDIO$ = TRIM$(!MEDIO)
            With GRILLAGRAL
                fi% = 0
                For i = 1 To .GRID.Rows - 1
                    Argumento$ = TRIM$(.GRID.TextMatrix(i, 2))
                    CC% = XVALO(.GRID.TextMatrix(i, 3))
                    If Argumento$ <> "" And CC% > 0 Then
                    
                        If InStr(1, Argumento, MEDIO$) > 0 And CC% = CContable% Then
                            fi% = i
                            Exit For
                        End If
                        
                    ElseIf Argumento$ <> "" Then
                    
                        If InStr(1, Argumento$, MEDIO$) > 0 Or InStr(1, MEDIO$, Argumento$) > 0 Then
                            fi% = i
                            Exit For
                        End If
                    
                    ElseIf CC% > 0 Then
                    
                        If CC% = CContable% Then
                            fi% = i
                            Exit For
                        End If
                        
                    End If
                Next i
                
                If fi% > 0 Then
                    co% = 3 + rstCobranzaPorMedios!MES
                    Acumulado# = XVALO(.GRID.TextMatrix(fi%, co%)) + XVALO(rstCobranzaPorMedios!IMPORTE)
                    .GRID.TextMatrix(fi%, co%) = FORMATNUM$(Acumulado#, "F14.2")
                Else
                    Call COMUNI("Falta Configurar Medio de Cobranza '" & MEDIO$ & "'" & vbCrLf & "Cuenta Contable: " & CODCUE$(CContable%))
                End If
            End With
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rstCobranzaPorMedios.Close
    Set rstCobranzaPorMedios = Nothing
    On Error GoTo 0
    
    
    ' *** UNIFICA POR GRUPO
    GRILLAGRAL.GRID.FixedCols = 3
    GRILLAGRAL.GRID.MergeCells = flexMergeRestrictColumns
    GRILLAGRAL.GRID.MergeCol(1) = True
    GRILLAGRAL.GRID.MergeCol(2) = True
    
    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Cash Flow"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub
End Sub

Private Sub mnuComparativaEquiposCantidad_Click()

    If DERACCE%("COMFACNE", "Comparativa Facturación Neta y Proyección") < 2 Then
        GoTo 99
    End If

    ' *** OBTIENE EL AÑO DESDE DONDE COMIENZA LA COMPARATIVA
    AnioDesde% = ListaDeAnios
    If AnioDesde% = 0 Then GoTo 99
    AnioComienzo% = AnioDesde%
    
    AnioDesde% = AnioDesde% Mod 100
    d$ = "01/01/" & CStr(AnioDesde%)
    AnioHasta% = Year(Now)
    AnioHasta% = AnioHasta% Mod 100
    H$ = "31/12/" & CStr(AnioHasta%)
    DesdeFecha% = FECHANUM(d$)
    HastaFecha% = FECHANUM(H$)
    
    ANIOACTUAL% = Year(Now)
    
    If AnioComienzo% > ANIOACTUAL% Then
        Call MENSERR(24, "El Mes de Inicio del Comparativo\No Puede ser Mayor al Año en Curso")
        GoTo 99
    End If
    
    IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
    CodigoDesde$ = "": CodigoHasta$ = "": CondicionSql$ = ""
    If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
        ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
        Dim RstCodigos As ADODB.Recordset
        Set RstCodigos = READSET(ConsultaSql$)
        
        With RstCodigos
            Do While Not .EOF
                CodigoProducto$ = SAFETEXT$(!CAMPO2)
                If CondicionSql$ = "" Then
                    CondicionSql$ = "cod_exte='" & CodigoProducto$ & "'"
                Else
                    CondicionSql$ = CondicionSql$ & " or cod_exte='" & CodigoProducto$ & "'"
                End If
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RstCodigos.Close
        Set RstCodigos = Nothing
        On Error GoTo 0
    Else
        racod$ = OBJPLA$("DESHASCOD", VDEF$)
        If racod$ = "" Then GoTo 99
        CI1% = CVI(racod$)
        CI2% = CVI(Mid$(racod$, 3))
        If CI1% = 0 Or CI2% = 0 Then GoTo 99
        CodigoDesde$ = CODEXT$(CI1%)
        CodigoHasta$ = CODEXT$(CI2%)
        If CodigoDesde$ > CodigoHasta$ Then GoTo 99
    End If
    
    Screen.MousePointer = 11
    
    ' *** DIBUJA LA GRILLA
    GRILLAGRAL.GRID.Clear
    CantidadColumnaGrilla% = 2 + (ANIOACTUAL% - AnioComienzo% + 1) + 1 + (ANIOACTUAL% - AnioComienzo%) + 1 + (ANIOACTUAL% - AnioComienzo% + 1)
    GRILLAGRAL.GRID.Cols = CantidadColumnaGrilla%
    CantidadFilasGrilla% = 35
    GRILLAGRAL.GRID.Rows = CantidadFilasGrilla%
    
    For i% = 1 To CantidadFilasGrilla% - 1
        For k% = 1 To CantidadColumnaGrilla% - 1
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(192, 211, 226), i%, k%, 1)
        Next k%
    Next i%
    GRILLAGRAL.GRID.ColWidth(0) = 2000
    
    InicioAnios% = 2
    For ANIO% = AnioComienzo% To ANIOACTUAL%
        GRILLAGRAL.GRID.TextMatrix(0, InicioAnios% + (ANIO% - AnioComienzo%)) = CStr(ANIO%)
        GRILLAGRAL.GRID.ColWidth(InicioAnios% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioAnios% + (ANIO% - AnioComienzo%)) = 6
        
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, InicioAnios% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%
    
    InicioVariacion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% + 1) + 1
    For ANIO% = AnioComienzo% To (ANIOACTUAL% - 1)
        GRILLAGRAL.GRID.TextMatrix(0, InicioVariacion% + (ANIO% - AnioComienzo%)) = "% " & CStr(ANIO%) & "-" & CStr(ANIO% + 1)
        GRILLAGRAL.GRID.ColWidth(InicioVariacion% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioVariacion% + (ANIO% - AnioComienzo%)) = 6
                
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, InicioVariacion% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%
    
    InicioPorcentaje% = InicioVariacion% + (ANIOACTUAL% - AnioComienzo%) + 1
    For ANIO% = AnioComienzo% To ANIOACTUAL%
        GRILLAGRAL.GRID.TextMatrix(0, InicioPorcentaje% + (ANIO% - AnioComienzo%)) = CStr(ANIO%)
        GRILLAGRAL.GRID.ColWidth(InicioPorcentaje% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioPorcentaje% + (ANIO% - AnioComienzo%)) = 6
        
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(223, 157, 125), i%, InicioPorcentaje% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%

    ' *** CARGA LOS MESES y TOTALES
    GRILLAGRAL.GRID.TextMatrix(2, 0) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(3, 0) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(4, 0) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(5, 0) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(6, 0) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(7, 0) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(8, 0) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(9, 0) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(10, 0) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(11, 0) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(12, 0) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(13, 0) = "DICIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(15, 0) = "TOTALES"
    
    ' *** MATRIZ TOTALES POR AÑO
    Dim TotalPorAnio() As Double
    CantidadAnios% = ANIOACTUAL% - AnioComienzo%
    ReDim Preserve TotalPorAnio(CantidadAnios%)
    
    ' *** CONSULTA SQL CANTIDADES VENDIDAS
    ConsultaSql$ = "select sum(cantsalid) salida, sum(cantingre) ingreso, month(fechconta) mes, year(fechconta) anio from movisto with(nolock)" _
    & " inner join cajabanc with(nolock) on (cajabanc.codicom_lar=movisto.doseclar)" _
    & " where tipomovim='venta' and fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and codimovi='rt'"
    If CondicionSql$ <> "" Then
        ConsultaSql$ = ConsultaSql$ & " and (" & CondicionSql$ & ")"
    Else
        ConsultaSql$ = ConsultaSql$ & " and cod_exte>='" & CodigoDesde$ & "' and cod_exte<='" & CodigoHasta$ & "'"
    End If
    ConsultaSql$ = ConsultaSql$ & " group by month(fechconta), year(fechconta)" _
    & " order by anio, mes"
    
    Text1.TEXT = ConsultaSql$
    Dim MesVenta, AnioVenta As Integer
    Dim CantidadEquipos As Double
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(ConsultaSql$)
    
    ' ** VUELCO LA INFORMACION EN LA GRILLA
    With Rst
        Do While Not .EOF
            MesVenta = !MES
            AnioVenta = !ANIO
            CantidadEquipos = ROUND#(!SALIDA - !INGRESO, 0)
            
            For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
                Encabezado% = GRILLAGRAL.GRID.TextMatrix(0, InicioAnios% + i%)
                If AnioVenta = Encabezado% Then
                    ColumnaAnio% = InicioAnios% + i%
                    TotalPorAnio(i%) = TotalPorAnio(i%) + CantidadEquipos
                    Exit For
                End If
            Next i%
            FilaMes% = 1 + MesVenta
            
            GRILLAGRAL.GRID.TextMatrix(FilaMes%, ColumnaAnio%) = TRIM$(FORMATNUM$(CantidadEquipos, "F14.0"))
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    ' *** VUELCA TOTALES A LA GRILLA
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        GRILLAGRAL.GRID.TextMatrix(15, InicioAnios% + i%) = TRIM$(FORMATNUM$(TotalPorAnio(i%), "F14.0"))
    Next i%
    
    ' *** CALCULO DE VARIACIONES AÑO A AÑO
    For i% = 0 To (ANIOACTUAL% - AnioComienzo% - 1)
        For k% = 2 To 13
            With GRILLAGRAL.GRID
                ImporteMes1# = XVALO(.TextMatrix(k%, InicioAnios% + i%))
                ImporteMes2# = XVALO(.TextMatrix(k%, InicioAnios% + i% + 1))

                If Abs(ImporteMes1#) > 0.005 And Abs(ImporteMes2#) > 0.005 Then
                    Variacion# = ((ImporteMes2# / ImporteMes1#) - 1) * 100
                    .TextMatrix(k%, InicioVariacion% + i%) = TRIM$(FORMATNUM$(Variacion#, "F12.2"))
                End If
            End With
        Next k%
    Next i%
    
    ' *** CALCULO PORCENTAJE MENSUAL
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        With GRILLAGRAL.GRID
            ImporteAnual# = XVALO(.TextMatrix(15, InicioAnios% + i%))
            If ImporteAnual# > 0.005 Then
                For k% = 2 To 13
                    ImporteMes# = XVALO(.TextMatrix(k%, InicioAnios% + i%))

                    Porcentaje# = (ImporteMes# / ImporteAnual#) * 100
                    .TextMatrix(k%, InicioPorcentaje% + i%) = TRIM$(FORMATNUM$(Porcentaje#, "F12.2"))
                Next k%
            End If
        End With
    Next i%
    
    ' *** PROYECCION AÑO ACTUAL
    ColumnaProyeccion% = 2
    
    ' *** PRIMERA PROYECCION
    If (ANIOACTUAL% - AnioComienzo% - 2) <= 0 Then GoTo 80
    
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)
   
    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% - 2)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 3)
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.0"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.0"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 3) & "-" & CStr(ANIOACTUAL% - 2)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    ColumnaProyeccion% = ColumnaProyeccion% + 3
    
    ' *** SEGUNDA PROYECCION
80  If (ANIOACTUAL% - AnioComienzo% - 1) <= 0 Then GoTo 82

    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)

    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% - 1)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 2)
    PorcentajeAcumulado# = 0: AcumuladoMensual# = 0
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.0"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.0"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 2) & "-" & CStr(ANIOACTUAL% - 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    ColumnaProyeccion% = ColumnaProyeccion% + 3
    
    ' *** TERCERA PROYECCION
82  If (ANIOACTUAL% - AnioComienzo%) <= 0 Then GoTo 85

    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)

    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo%)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 1)
    PorcentajeAcumulado# = 0: AcumuladoMensual# = 0
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.0"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.0"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 1) & "-" & CStr(ANIOACTUAL%)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    
    ' ************************
85  GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Comparativa por Año | Equipos - Cantidad y Proyección"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, "Comparativo por Año - Facturación Neta y Proyección", False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

    
End Sub

Private Sub mnuComparativaExportacion_Click()

    If DERACCE%("COMPEXPO", "Comparativa Exportación vs Venta Total") < 2 Then
        GoTo 99
    End If

    ' *** OBTIENE EL AÑO DESDE DONDE COMIENZA LA COMPARATIVA
    ' *** OBTIENE EL AÑO DESDE DONDE COMIENZA LA COMPARATIVA
    AnioDesde% = ListaDeAnios
    If AnioDesde% = 0 Then GoTo 99
    AnioComienzo% = AnioDesde%
    
    AnioDesde% = AnioDesde% Mod 100
    d$ = "01/01/" & CStr(AnioDesde%)
    AnioHasta% = Year(Now)
    AnioHasta% = AnioHasta% Mod 100
    H$ = "31/12/" & CStr(AnioHasta%)
    DesdeFecha% = FECHANUM(d$)
    HastaFecha% = FECHANUM(H$)
    
    ANIOACTUAL% = Year(Now)
    
    If AnioComienzo% > ANIOACTUAL% Then
        Call MENSERR(24, "El Año de Inicio del Comparativo\No Puede ser Mayor al Año en Curso")
        GoTo 99
    End If
    
    Screen.MousePointer = 11
    
    ' *** DIBUJA LA GRILLA
    GRILLAGRAL.GRID.Clear
    CantidadColumnaGrilla% = 2 + (ANIOACTUAL% - AnioComienzo% + 1) * 4
    GRILLAGRAL.GRID.Cols = CantidadColumnaGrilla%
    CantidadFilasGrilla% = 24
    GRILLAGRAL.GRID.Rows = CantidadFilasGrilla%
    
    InicioAnios% = 2
    ColorActivo% = 1
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        With GRILLAGRAL.GRID
            .TextMatrix(1, InicioAnios% + (4 * i%) + 2) = AnioComienzo% + i%
        
            Select Case ColorActivo%
                Case 1: COLOR0 = RGB(162, 210, 196): ColorActivo% = 2
                Case 2: COLOR0 = RGB(201, 184, 143): ColorActivo% = 3
                Case 3: COLOR0 = RGB(223, 157, 125): ColorActivo% = 1
            End Select
        
            Call COLOREAR_GRILLA_SOLOCOLUMNA(GRILLAGRAL.GRID, COLOR0, 1, InicioAnios% + (4 * i%))
            Call COLOREAR_GRILLA_SOLOCOLUMNA(GRILLAGRAL.GRID, COLOR0, 1, InicioAnios% + (4 * i%) + 1)
            Call COLOREAR_GRILLA_SOLOCOLUMNA(GRILLAGRAL.GRID, COLOR0, 1, InicioAnios% + (4 * i%) + 2)
            Call COLOREAR_GRILLA_SOLOCOLUMNA(GRILLAGRAL.GRID, COLOR0, 1, InicioAnios% + (4 * i%) + 3)
            'Call COLOREAR_GRILLA_SOLOCOLUMNA(GRILLAGRAL.GRID, COLOR0, 1, InicioAnios% + (4 * i%) + 4)
            
            .ColWidth(InicioAnios% + (4 * i%)) = 1500: .ColAlignment(InicioAnios% + (4 * i%)) = 6
            .ColWidth(InicioAnios% + (4 * i%) + 1) = 1500: .ColAlignment(InicioAnios% + (4 * i%) + 1) = 6
            .ColWidth(InicioAnios% + (4 * i%) + 2) = 1500: .ColAlignment(InicioAnios% + (4 * i%) + 2) = 6
            .ColWidth(InicioAnios% + (4 * i%) + 3) = 1500: .ColAlignment(InicioAnios% + (4 * i%) + 3) = 6
            '.ColWidth(InicioAnios% + (4 * i%) + 4) = 1500: .ColAlignment(InicioAnios% + (4 * i%) + 4) = 6
            
            .Row = 1: .COL = InicioAnios% + (4 * i%) + 1: .CellAlignment = flexAlignCenterCenter
        End With
    Next i%
    
    With GRILLAGRAL.GRID
        .Row = 2
        For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
            .TextMatrix(2, InicioAnios% + (4 * i%)) = "$ Total"
            .COL = InicioAnios% + (4 * i%): .CellAlignment = flexAlignCenterCenter
        Next i%
        For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
            .TextMatrix(2, InicioAnios% + (4 * i%) + 1) = "$ Expo"
            .COL = InicioAnios% + (4 * i%) + 1: .CellAlignment = flexAlignCenterCenter
        Next i%
        For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
            .TextMatrix(2, InicioAnios% + (4 * i%) + 2) = "U$ Expo"
            .COL = InicioAnios% + (4 * i%) + 2: .CellAlignment = flexAlignCenterCenter
        Next i%
        For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
            .TextMatrix(2, InicioAnios% + (4 * i%) + 3) = "% en Pesos"
            .COL = InicioAnios% + (4 * i%) + 3: .CellAlignment = flexAlignCenterCenter
        Next i%
'        For i% = 0 To (AnioActual% - AnioComienzo%)
'            .TextMatrix(2, InicioAnios% + (5 * i%) + 4) = "% en U$"
'            .COL = InicioAnios% + (5 * i%) + 4: .CellAlignment = flexAlignCenterCenter
'        Next i%
    End With
    
    ' *** CARGA LOS MESES y TOTALES
    GRILLAGRAL.GRID.TextMatrix(3, 0) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(4, 0) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(5, 0) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(6, 0) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(7, 0) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(8, 0) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(9, 0) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(10, 0) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(11, 0) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(12, 0) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(13, 0) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(14, 0) = "DICIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(16, 0) = "TOTALES"
    GRILLAGRAL.GRID.ColWidth(0) = 1500
    
    ' *** MATRIZ TOTALES POR AÑO
    Dim TotalPorAnio() As Double
    Dim TotalPorAnioExportacion() As Double
    Dim TotalPorAnioExportacionDolares() As Double
    CantidadAnios% = ANIOACTUAL% - AnioComienzo%
    ReDim Preserve TotalPorAnio(CantidadAnios%)
    ReDim Preserve TotalPorAnioExportacion(CantidadAnios%)
    ReDim Preserve TotalPorAnioExportacionDolares(CantidadAnios%)
    
    ' *** CONSULTA SQL VENTA NETA
    ConsultaSql$ = "select sum(c.idebecomp) db, sum(c.ihabecomp) hb, month(c.fechconta) mes, year(c.fechconta) anio from cajabanc c with(nolock) inner join deudor12 d with(nolock) on c.nuclien=d.nume_cli" _
    & " where c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and c.tipomovim='venta' and c.numecomp>0" _
    & " group by month(c.fechconta), year(c.fechconta)" _
    & " order by anio, mes"

    Dim MesVenta, AnioVenta As Integer
    Dim VentaNeta As Double
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(ConsultaSql$)

    ' ** VUELCO LA INFORMACION EN LA GRILLA
    With Rst
        Do While Not .EOF
            MesVenta = !MES
            AnioVenta = !ANIO
            VentaNeta = ROUND#(!HB - !Db, 2)

            For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
                Encabezado% = CInt(GRILLAGRAL.GRID.TextMatrix(1, InicioAnios% + (4 * i%) + 2))
                If AnioVenta = Encabezado% Then
                    ColumnaAnio% = InicioAnios% + (4 * i%)
                    TotalPorAnio(i%) = TotalPorAnio(i%) + VentaNeta
                    Exit For
                End If
            Next i%
            FilaMes% = 2 + MesVenta

            GRILLAGRAL.GRID.TextMatrix(FilaMes%, ColumnaAnio%) = TRIM$(FORMATNUM$(VentaNeta, "F14.2"))

            .MoveNext
        Loop
    End With

    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0

    ' *** VUELCA TOTALES A LA GRILLA
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i%)) = TRIM$(FORMATNUM$(TotalPorAnio(i%), "F14.2"))
    Next i%

    ' *** CONSULTA SQL VENTA NETA
    ConsultaSql$ = "(select sum(c.idebecomp) db, sum(c.ihabecomp) hb, sum(0) dbdol, sum(0) hbdol, month(c.fechconta) mes, year(c.fechconta) anio from cajabanc c with(nolock) inner join deudor12 d with(nolock) on c.nuclien=d.nume_cli" _
    & " where c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and c.tipomovim='venta' and c.numecomp>0 and d.piva_cli=5" _
    & " group by month(c.fechconta), year(c.fechconta))" _
    & " union all" _
    & "(select sum(0) db, sum(0) hb, sum(c.idebecomp/tipo_cam) dbdol, sum(c.ihabecomp/tipo_cam) hbdol, month(c.fechconta) mes, year(c.fechconta) anio from cajabanc c with(nolock) inner join deudor12 d with(nolock) on c.nuclien=d.nume_cli" _
    & " where c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and c.tipomovim='venta' and c.numecomp>0 and d.piva_cli=5 and c.mone_emis>1" _
    & " group by month(c.fechconta), year(c.fechconta))" _
    & " order by anio, mes"
    
    Dim VentaExportacion As Double
    Dim RstExpo As ADODB.Recordset
    Set RstExpo = READSET(ConsultaSql$)
    
    ' ** VUELCO LA INFORMACION EN LA GRILLA
    With RstExpo
        Do While Not .EOF
            MesVenta = !MES
            AnioVenta = !ANIO
            VentaExportacion = ROUND#(!HB - !Db, 2)
            VentaExportacionDolar = ROUND#(!HBdol - !Dbdol, 2)
            
            For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
                Encabezado% = CInt(GRILLAGRAL.GRID.TextMatrix(1, InicioAnios% + (4 * i%) + 2))
                If AnioVenta = Encabezado% Then
                    ColumnaAnio% = InicioAnios% + (4 * i%) + 1
                    TotalPorAnioExportacion(i%) = TotalPorAnioExportacion(i%) + VentaExportacion
                    TotalPorAnioExportacionDolares(i%) = TotalPorAnioExportacionDolares(i%) + VentaExportacionDolar
                    Exit For
                End If
            Next i%
            FilaMes% = 2 + MesVenta
            
            If VentaExportacion > 0.005 Then
                GRILLAGRAL.GRID.TextMatrix(FilaMes%, ColumnaAnio%) = TRIM$(FORMATNUM$(VentaExportacion, "F14.2"))
            End If
            If VentaExportacionDolar > 0.005 Then
                GRILLAGRAL.GRID.TextMatrix(FilaMes%, ColumnaAnio% + 1) = TRIM$(FORMATNUM$(VentaExportacionDolar, "F14.2"))
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    RstExpo.Close
    Set RstExpo = Nothing
    On Error GoTo 0
    
    ' *** VUELCA TOTALES A LA GRILLA
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i%) + 1) = TRIM$(FORMATNUM$(TotalPorAnioExportacion(i%), "F14.2"))
        GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i%) + 2) = TRIM$(FORMATNUM$(TotalPorAnioExportacionDolares(i%), "F14.2"))
    Next i%
    
    ' *** CALCULO PORCENTAJE EXPORTACIONES VS VENTA TOTAL
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        For k% = 3 To 14
        
            VentaNeta = XVALO(GRILLAGRAL.GRID.TextMatrix(k%, InicioAnios% + (4 * i%)))
            VentaExportacion = XVALO(GRILLAGRAL.GRID.TextMatrix(k%, InicioAnios% + (4 * i%) + 1))
            If VentaNeta > 0.005 Then
                VentasVsExpo# = (VentaExportacion / VentaNeta) * 100
                GRILLAGRAL.GRID.TextMatrix(k%, InicioAnios% + (4 * i%) + 3) = TRIM$(FORMATNUM$(VentasVsExpo#, "F14.2"))
            End If
            
            VentaExportacionDolar = XVALO(GRILLAGRAL.GRID.TextMatrix(k%, InicioAnios% + (4 * i%) + 2))
            If TotalPorAnioExportacionDolares(i%) > 0.005 Then
                PorcentajeVentaExportacionDolar# = (VentaExportacionDolar / TotalPorAnioExportacionDolares(i%)) * 100
                'GRILLAGRAL.GRID.TextMatrix(k%, InicioAnios% + (4 * i%) + 4) = TRIM$(FORMATNUM$(PorcentajeVentaExportacionDolar#, "F14.2"))
            End If
            
        Next k%
    Next i%
    
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        GRILLAGRAL.GRID.TextMatrix(18, InicioAnios% + (4 * i)) = "Local"
        GRILLAGRAL.GRID.TextMatrix(19, InicioAnios% + (4 * i)) = "Expo $"
        GRILLAGRAL.GRID.TextMatrix(20, InicioAnios% + (4 * i)) = "%"
        GRILLAGRAL.GRID.TextMatrix(22, InicioAnios% + (4 * i)) = "Expo U$"
        GRILLAGRAL.GRID.TextMatrix(23, InicioAnios% + (4 * i)) = "% Interanual"
        
        VentaTotal# = XVALO(GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i)))
        ExpoPesos# = XVALO(GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i) + 1))
        ExpoDolar# = XVALO(GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + (4 * i) + 2))
        VentaNetaDeExpo# = VentaTotal# - ExpoPesos#
        
        GRILLAGRAL.GRID.TextMatrix(18, InicioAnios% + (4 * i) + 1) = FORMATNUM$(VentaNetaDeExpo#, "F14.2")
        GRILLAGRAL.GRID.TextMatrix(19, InicioAnios% + (4 * i) + 1) = FORMATNUM$(ExpoPesos#, "F14.2")
        GRILLAGRAL.GRID.TextMatrix(20, InicioAnios% + (4 * i) + 1) = ""
        GRILLAGRAL.GRID.TextMatrix(22, InicioAnios% + (4 * i) + 1) = ""
        
        PorcentajeLocal# = 0: If VentaTotal# > 0.005 Then PorcentajeLocal# = (VentaNetaDeExpo# / VentaTotal#) * 100
        GRILLAGRAL.GRID.TextMatrix(18, InicioAnios% + (4 * i) + 2) = FORMATNUM$(PorcentajeLocal#, "F14.2")
        PorcentajeExpo# = 0: If VentaTotal# > 0.005 Then PorcentajeExpo# = (ExpoPesos# / VentaTotal#) * 100
        GRILLAGRAL.GRID.TextMatrix(19, InicioAnios% + (4 * i) + 2) = FORMATNUM$(PorcentajeExpo#, "F14.2")
        GRILLAGRAL.GRID.TextMatrix(20, InicioAnios% + (4 * i) + 2) = FORMATNUM$(PorcentajeLocal# + PorcentajeExpo#, "F14.2")
        GRILLAGRAL.GRID.TextMatrix(22, InicioAnios% + (4 * i) + 2) = FORMATNUM$(ExpoDolar#, "F14.2")
        If i > 0 Then
            ExpoAnioAnterior# = XVALO(GRILLAGRAL.GRID.TextMatrix(22, InicioAnios% + (4 * (i - 1)) + 2))
            ExpoAnioActivo# = XVALO(GRILLAGRAL.GRID.TextMatrix(22, InicioAnios% + (4 * i) + 2))
            PorcentajeInterAnual# = 0
            If ExpoAnioActivo# > 0.005 Then PorcentajeInterAnual# = ((ExpoAnioActivo# / ExpoAnioAnterior#) - 1) * 100
            GRILLAGRAL.GRID.TextMatrix(23, InicioAnios% + (4 * i) + 2) = FORMATNUM$(PorcentajeInterAnual#, "F14.2")
        End If
    Next i%


85  GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Comparativa por Año | Exportaciones"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, "Comparativo por Año - Facturación Neta y Proyección", False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Private Sub mnuComparativaFacturacionNeta_Click()

    If DERACCE%("COMFACNE", "Comparativa Facturación Neta y Proyección") < 2 Then
        GoTo 99
    End If

    ' *** OBTIENE EL AÑO DESDE DONDE COMIENZA LA COMPARATIVA
    AnioDesde% = ListaDeAnios
    If AnioDesde% = 0 Then GoTo 99
    AnioComienzo% = AnioDesde%
    
    AnioDesde% = AnioDesde% Mod 100
    d$ = "01/01/" & CStr(AnioDesde%)
    AnioHasta% = Year(Now)
    AnioHasta% = AnioHasta% Mod 100
    H$ = "31/12/" & CStr(AnioHasta%)
    DesdeFecha% = FECHANUM(d$)
    HastaFecha% = FECHANUM(H$)
    
    ANIOACTUAL% = Year(Now)
    
    If AnioComienzo% > ANIOACTUAL% Then
        Call MENSERR(24, "El Año de Inicio del Comparativo\No Puede ser Mayor al Año en Curso")
        GoTo 99
    End If
    
    Screen.MousePointer = 11
    
    ' *** DIBUJA LA GRILLA
    GRILLAGRAL.GRID.Clear
    CantidadColumnaGrilla% = 2 + (ANIOACTUAL% - AnioComienzo% + 1) + 1 + (ANIOACTUAL% - AnioComienzo%) + 1 + (ANIOACTUAL% - AnioComienzo% + 1)
    GRILLAGRAL.GRID.Cols = CantidadColumnaGrilla%
    CantidadFilasGrilla% = 35
    GRILLAGRAL.GRID.Rows = CantidadFilasGrilla%
    
    For i% = 1 To CantidadFilasGrilla% - 1
        For k% = 1 To CantidadColumnaGrilla% - 1
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(192, 211, 226), i%, k%, 1)
        Next k%
    Next i%
    GRILLAGRAL.GRID.ColWidth(0) = 2000
    
    InicioAnios% = 2
    For ANIO% = AnioComienzo% To ANIOACTUAL%
        GRILLAGRAL.GRID.TextMatrix(0, InicioAnios% + (ANIO% - AnioComienzo%)) = CStr(ANIO%)
        GRILLAGRAL.GRID.ColWidth(InicioAnios% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioAnios% + (ANIO% - AnioComienzo%)) = 6
        
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, InicioAnios% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%
    
    InicioVariacion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% + 1) + 1
    For ANIO% = AnioComienzo% To (ANIOACTUAL% - 1)
        GRILLAGRAL.GRID.TextMatrix(0, InicioVariacion% + (ANIO% - AnioComienzo%)) = "% " & CStr(ANIO%) & "-" & CStr(ANIO% + 1)
        GRILLAGRAL.GRID.ColWidth(InicioVariacion% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioVariacion% + (ANIO% - AnioComienzo%)) = 6
                
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, InicioVariacion% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%
    
    InicioPorcentaje% = InicioVariacion% + (ANIOACTUAL% - AnioComienzo%) + 1
    For ANIO% = AnioComienzo% To ANIOACTUAL%
        GRILLAGRAL.GRID.TextMatrix(0, InicioPorcentaje% + (ANIO% - AnioComienzo%)) = CStr(ANIO%)
        GRILLAGRAL.GRID.ColWidth(InicioPorcentaje% + (ANIO% - AnioComienzo%)) = 1500
        GRILLAGRAL.GRID.ColAlignment(InicioPorcentaje% + (ANIO% - AnioComienzo%)) = 6
        
        For i% = 1 To 14
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(223, 157, 125), i%, InicioPorcentaje% + (ANIO% - AnioComienzo%), 1)
        Next i%
    Next ANIO%

    ' *** CARGA LOS MESES y TOTALES
    GRILLAGRAL.GRID.TextMatrix(2, 0) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(3, 0) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(4, 0) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(5, 0) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(6, 0) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(7, 0) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(8, 0) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(9, 0) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(10, 0) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(11, 0) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(12, 0) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(13, 0) = "DICIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(15, 0) = "TOTALES"
    GRILLAGRAL.GRID.TextMatrix(16, 0) = "VARIACION"
    
    ' *** MATRIZ TOTALES POR AÑO
    Dim TotalPorAnio() As Double
    CantidadAnios% = ANIOACTUAL% - AnioComienzo%
    ReDim Preserve TotalPorAnio(CantidadAnios%)
    
    ' *** CONSULTA SQL VENTA NETA
    ConsultaSql$ = "select sum(idebecomp) db, sum(ihabecomp) hb, month(fechconta) mes, year(fechconta) anio from cajabanc with(nolock)" _
    & " where fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and tipomovim='venta' and numecomp>0" _
    & " group by month(fechconta), year(fechconta)" _
    & " order by anio, mes"
    
    Dim MesVenta, AnioVenta As Integer
    Dim VentaNeta As Double
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(ConsultaSql$)
    
    ' ** VUELCO LA INFORMACION EN LA GRILLA
    With Rst
        Do While Not .EOF
            MesVenta = !MES
            AnioVenta = !ANIO
            VentaNeta = ROUND#(!HB - !Db, 2)
            
            For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
                Encabezado% = GRILLAGRAL.GRID.TextMatrix(0, InicioAnios% + i%)
                If AnioVenta = Encabezado% Then
                    ColumnaAnio% = InicioAnios% + i%
                    TotalPorAnio(i%) = TotalPorAnio(i%) + VentaNeta
                    Exit For
                End If
            Next i%
            FilaMes% = 1 + MesVenta
            
            GRILLAGRAL.GRID.TextMatrix(FilaMes%, ColumnaAnio%) = TRIM$(FORMATNUM$(VentaNeta, "F14.2"))
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    ' *** VUELCA TOTALES A LA GRILLA
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        GRILLAGRAL.GRID.TextMatrix(15, InicioAnios% + i%) = TRIM$(FORMATNUM$(TotalPorAnio(i%), "F14.2"))
        If i% > 0 Then
            TotalAnioAnterior# = XVALO(GRILLAGRAL.GRID.TextMatrix(15, InicioAnios% + i% - 1))
            TotalAnioActual# = XVALO(GRILLAGRAL.GRID.TextMatrix(15, InicioAnios% + i%))
            
            Variacion# = (TotalAnioActual# / TotalAnioAnterior#) - 1
            GRILLAGRAL.GRID.TextMatrix(16, InicioAnios% + i%) = TRIM$(FORMATNUM$(Variacion# * 100, "F14.2"))
        End If
    Next i%
    
    ' *** CALCULO DE VARIACIONES AÑO A AÑO
    For i% = 0 To (ANIOACTUAL% - AnioComienzo% - 1)
        For k% = 2 To 13
            With GRILLAGRAL.GRID
                ImporteMes1# = XVALO(.TextMatrix(k%, InicioAnios% + i%))
                ImporteMes2# = XVALO(.TextMatrix(k%, InicioAnios% + i% + 1))

                If Abs(ImporteMes1#) > 0.005 And Abs(ImporteMes2#) > 0.005 Then
                    Variacion# = ((ImporteMes2# / ImporteMes1#) - 1) * 100
                    .TextMatrix(k%, InicioVariacion% + i%) = TRIM$(FORMATNUM$(Variacion#, "F12.2"))
                End If
            End With
        Next k%
    Next i%
    
    ' *** CALCULO PORCENTAJE MENSUAL
    For i% = 0 To (ANIOACTUAL% - AnioComienzo%)
        With GRILLAGRAL.GRID
            ImporteAnual# = XVALO(.TextMatrix(15, InicioAnios% + i%))
            If ImporteAnual# > 0.005 Then
                For k% = 2 To 13
                    ImporteMes# = XVALO(.TextMatrix(k%, InicioAnios% + i%))

                    Porcentaje# = (ImporteMes# / ImporteAnual#) * 100
                    .TextMatrix(k%, InicioPorcentaje% + i%) = TRIM$(FORMATNUM$(Porcentaje#, "F12.2"))
                Next k%
            End If
        End With
    Next i%
    
    ' *** PROYECCION AÑO ACTUAL
    ColumnaProyeccion% = 2
    
    ' *** PRIMERA PROYECCION
    If (ANIOACTUAL% - AnioComienzo% - 2) <= 0 Then GoTo 80
    
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)
   
    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% - 2)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 3)
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.2"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.2"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 3) & "-" & CStr(ANIOACTUAL% - 2)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    ColumnaProyeccion% = ColumnaProyeccion% + 3
    
    ' *** SEGUNDA PROYECCION
80  If (ANIOACTUAL% - AnioComienzo% - 1) <= 0 Then GoTo 82

    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)

    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo% - 1)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 2)
    PorcentajeAcumulado# = 0: AcumuladoMensual# = 0
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.2"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.2"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 2) & "-" & CStr(ANIOACTUAL% - 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    ColumnaProyeccion% = ColumnaProyeccion% + 3
    
    ' *** TERCERA PROYECCION
82  If (ANIOACTUAL% - AnioComienzo%) <= 0 Then GoTo 85

    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 1) = "Proyección"
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion% + 2) = "Real"
    GRILLAGRAL.GRID.TextMatrix(20, ColumnaProyeccion%) = "ENERO"
    GRILLAGRAL.GRID.TextMatrix(21, ColumnaProyeccion%) = "FEBRERO"
    GRILLAGRAL.GRID.TextMatrix(22, ColumnaProyeccion%) = "MARZO"
    GRILLAGRAL.GRID.TextMatrix(23, ColumnaProyeccion%) = "ABRIL"
    GRILLAGRAL.GRID.TextMatrix(24, ColumnaProyeccion%) = "MAYO"
    GRILLAGRAL.GRID.TextMatrix(25, ColumnaProyeccion%) = "JUNIO"
    GRILLAGRAL.GRID.TextMatrix(26, ColumnaProyeccion%) = "JULIO"
    GRILLAGRAL.GRID.TextMatrix(27, ColumnaProyeccion%) = "AGOSTO"
    GRILLAGRAL.GRID.TextMatrix(28, ColumnaProyeccion%) = "SEPTIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(29, ColumnaProyeccion%) = "OCTUBRE"
    GRILLAGRAL.GRID.TextMatrix(30, ColumnaProyeccion%) = "NOVIEMBRE"
    GRILLAGRAL.GRID.TextMatrix(31, ColumnaProyeccion%) = "DICIEMBRE"
    
    For i% = 20 To 31
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), i%, ColumnaProyeccion%, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(201, 184, 143), i%, ColumnaProyeccion% + 2, 1)
    Next i%
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 1, 1)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion% + 2, 1)

    ColumnaAnioBaseProyeccion% = InicioAnios% + (ANIOACTUAL% - AnioComienzo%)
    ColumnaPorcentajeBaseProyeccion% = InicioPorcentaje + (ANIOACTUAL% - AnioComienzo% - 1)
    PorcentajeAcumulado# = 0: AcumuladoMensual# = 0
    With GRILLAGRAL.GRID
        For i% = 2 To 13
            ImporteMes# = XVALO(.TextMatrix(i%, ColumnaAnioBaseProyeccion%))
            PorcentajeMes# = XVALO(.TextMatrix(i%, ColumnaPorcentajeBaseProyeccion%))

            PorcentajeAcumulado# = PorcentajeAcumulado# + PorcentajeMes#
            If ImporteMes# > 0.005 Then
                AcumuladoMensual# = AcumuladoMensual# + ImporteMes#
                If PorcentajeAcumulado# > 0.005 Then
                    ImporteAnualProyectado# = AcumuladoMensual# / (PorcentajeAcumulado# / 100)
        
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 1) = TRIM$(FORMATNUM$(ImporteAnualProyectado#, "F12.2"))
                    .TextMatrix(18 + i%, ColumnaProyeccion% + 2) = TRIM$(FORMATNUM$(AcumuladoMensual#, "F12.2"))
                End If
            End If
        Next i%
    End With
    GRILLAGRAL.GRID.TextMatrix(19, ColumnaProyeccion%) = CStr(ANIOACTUAL% - 1) & "-" & CStr(ANIOACTUAL%)
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(162, 210, 196), 19, ColumnaProyeccion%, 1)
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion%) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion%) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 1) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 1) = 6
    GRILLAGRAL.GRID.ColWidth(ColumnaProyeccion% + 2) = 1500: GRILLAGRAL.GRID.ColAlignment(ColumnaProyeccion% + 3) = 6
    
    ' ************************
85  GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Comparativa por Año | Facturación Neta y Proyección"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, "Comparativo por Año - Facturación Neta y Proyección", False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Private Sub mnuCuentasEgreso_Click()

    If DERACCE%("CFGESVE", "Configuración de Estadísticas de Ventas") < 2 Then
        GoTo 99
    End If
    
    Call TablasAuxiliares
    
    IDLIST& = ObtenerIdTablaAuxiliar("CUECONEGRE")
    Call ABM_TABLA(IDLIST&, , , "campo2")
    
99  Exit Sub
    
End Sub

Private Sub mnuEgresos_Click()

    If DERACCE%("RESUEGRE", "Resumen de Egresos") < 2 Then
        GoTo 99
    End If
    
    ' *** PERIODO EN EL QUE SE TRABAJA
    AnioElegido% = ListaDeAnios
    If AnioElegido% = 0 Then GoTo 99
    AnioElegido% = AnioElegido% Mod 100
    d$ = "01/01/" & CStr(AnioElegido%): DesdeFecha% = FECHANUM(d$): DESDE0$ = FETEXCOR1$(DesdeFecha%)
    H$ = "31/12/" & CStr(AnioElegido%): HastaFecha% = FECHANUM(H$): HASTA0$ = FETEXCOR1$(HastaFecha%)

    ' *** MODO
    IDLIST& = ObtenerIdTablaAuxiliar("CUECONEGRE")
    ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
    Set rstTablaCuentasEgresos = READSET(ConsultaSql$)
    With rstTablaCuentasEgresos
        Do While Not .EOF
            CuentaContable$ = SAFETEXT$(!CAMPO2)
            
            If CondicionSql$ = "" Then
                CondicionSql$ = "cuecontai=" & CStr(CUEINT%(CuentaContable$))
            Else
                CondicionSql$ = CondicionSql$ & " or cuecontai=" & CStr(CUEINT%(CuentaContable$))
            End If
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rstTablaCuentasEgresos.Close
    Set rstTablaCuentasEgresos = Nothing
    On Error GoTo 0
    
    ' *** OBTIENE CANTIDAD DE GASTOS
    
    CantidadGastos& = CantRegistros("datasql", "id_lista=" & CStr(IDLIST&))
    CondicionCajabanc$ = CondicionSql$
    CondicionDetaAsien$ = REPLACAR$(CondicionSql$, "cuecontai", "cuecoint")
    
'    Dim rstCantidadGastos As ADODB.Recordset
'    CondicionCajabanc$ = CondicionSql$
'    CondicionDetaAsien$ = REPLACAR$(CondicionSql$, "cuecontai", "cuecoint")
'    ConsultaSql$ = "select count(*) cantgastos from" _
'    & " (select codexcue from" _
'    & " ( select ccontable, sum(debe) debe, sum(haber) haber, mes from" _
'    & " ( (select 'cajabanc' tabla, cuecontai ccontable, sum(idebecomp) debe, sum(ihabecomp) haber, month(fechconta) mes from cajabanc with(nolock) where fechconta>='" & DESDE0$ & "' and fechconta<='" & HASTA0$ & "' and (" & CondicionCajabanc$ & ") group by cuecontai, month(fechconta))" _
'    & " Union All" _
'    & " (select 'detaasien' tabla, cuecoint ccontable, sum(idebepase) debe, sum(ihabepase) haber, month(d.fechasien) mes from detaasien d with(nolock) inner join encabeasien e with(nolock) on (d.nuidasien=e.nuidasien) where d.fechasien>='" & DESDE0$ & "' and d.fechasien<='" & HASTA0$ & "' and (e.tipoasien=0 or e.tipoasien=11) and d.statpase>0 and e.statasien>0 and (" & CondicionDetaAsien$ & ") group by cuecoint, month(d.fechasien)) ) gastos" _
'    & " group by ccontable, mes) gastosagrupa" _
'    & " left join plancuen p on (gastosagrupa.ccontable=p.codincue) group by codexcue) gastosordenados"
'    Text1.TEXT = ConsultaSql$
'    Set rstCantidadGastos = READSET(ConsultaSql$)
'    CantidadGastos& = rstCantidadGastos!cantgastos
'
'    On Error Resume Next
'    rstCantidadGastos.Close
'    Set rstCantidadGastos = Nothing
'    On Error GoTo 0
    
    ' *** DIBUJA EL INFORME
    GRILLAGRAL.GRID.Clear
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, "Cuenta/A20;Denominación/A32;Enero/F12.2;Febrero/F12.2;Marzo/F12.2;Abril/F12.2;Mayo/F12.2;Junio/F12.2;Julio/F12.2;Agosto/F12.2;Septiembre/F12.2;Octubre/F12.2;Noviembre/F12.2;Diciembre/F12.2;Tot. Cuenta/F14.2;% Cuenta/F14.2;Promedio/F14.2", GRILLAGRAL, 1)
    GRILLAGRAL.GRID.FixedCols = 3
    GRILLAGRAL.GRID.Rows = CantidadGastos& + 3
    GRILLAGRAL.GRID.TextMatrix(CantidadGastos& + 2, 2) = "TOTALES"
    
    ' *** COLOREA INFORME
    With GRILLAGRAL
        For fi% = 1 To CantidadGastos&
            For co% = 3 To 15
                Select Case co%
                Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(162, 210, 196), fi%, co%, 1)
                Case 15
                    Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(201, 184, 143), fi%, co%, 1)
                End Select
                
                Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(207, 226, 243), CantidadGastos& + 2, co%, 1)
            Next co%
        Next fi%
    End With
    
    ' *** OBTIENE GASTOS POR CUENTA CONTABLE
    Dim rstGastos As ADODB.Recordset
    ConsultaSql$ = "select p.codexcue, p.denocue, p.codincue, case when gastosagrupa.debe is null then 0 else gastosagrupa.debe end debe, case when gastosagrupa.haber is null then 0 else gastosagrupa.haber end haber, case when gastosagrupa.mes is null then 1 else gastosagrupa.mes end mes from" _
    & " (select codincue, codexcue, denocue from datasql d inner join plancuen p on (d.campo2=p.codexcue) where d.lista='cueconegre') p" _
    & " left join" _
    & " ( select ccontable, sum(debe) debe, sum(haber) haber, mes from" _
    & " ( (select 'cajabanc' tabla, cuecontai ccontable, sum(idebecomp) debe, sum(ihabecomp) haber, month(fechconta) mes from cajabanc with(nolock) where fechconta>='" & DESDE0$ & "' and fechconta<='" & HASTA0$ & "' and (" & CondicionCajabanc$ & ") group by cuecontai, month(fechconta))" _
    & " Union All" _
    & " (select 'detaasien' tabla, cuecoint ccontable, sum(idebepase) debe, sum(ihabepase) haber, month(d.fechasien) mes from detaasien d with(nolock) inner join encabeasien e with(nolock) on (d.nuidasien=e.nuidasien) where d.fechasien>='" & DESDE0$ & "' and d.fechasien<='" & HASTA0$ & "' and (e.tipoasien=0 or e.tipoasien=11) and d.statpase>0 and e.statasien>0 and (" & CondicionDetaAsien$ & ") group by cuecoint, month(d.fechasien)) ) gastos" _
    & " group by ccontable, mes) gastosagrupa on (p.codincue=gastosagrupa.ccontable)" _
    & " order by p.codexcue, gastosagrupa.mes"
    Text2.TEXT = ConsultaSql$
    Set rstGastos = READSET(ConsultaSql$)

    With rstGastos
        ProximoGasto% = 1
        Do While Not .EOF
            CuentaContable$ = SAFETEXT$(!CodExCue)
            Denominacion$ = SAFETEXT$(!DENOCUE)
            MES% = !MES
            ImporteMes# = XVALO(!DEBE) - XVALO(!HABER)
            
            With GRILLAGRAL
                fi% = 0
                For i% = 1 To CantidadGastos&
                    If .GRID.TextMatrix(i%, 1) = CuentaContable$ Then
                        fi% = i%
                        Exit For
                    End If
                Next i%
                If fi% = 0 Then
                    fi% = ProximoGasto%
                    .GRID.TextMatrix(fi%, 1) = CuentaContable$
                    .GRID.TextMatrix(fi%, 2) = Denominacion$
                    ProximoGasto% = ProximoGasto% + 1
                End If
                co% = 2 + MES%
                
                Acumulado# = XVALO(.GRID.TextMatrix(fi%, co%))
                Acumulado# = Acumulado# + ImporteMes#
                .GRID.TextMatrix(fi%, co%) = TRIM$(FORMATNUM$(Acumulado#, "F12.2"))
            End With
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstGastos.Close
    Set rstGastos = Nothing
    On Error GoTo 0
    
    ' *** TOTALES POR VENDEDOR
    With GRILLAGRAL
        For fi% = 1 To CantidadGastos&
            Acumulado# = 0
            For co% = 3 To 14
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next co%
            .GRID.TextMatrix(fi%, 15) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next fi%
    End With
    
    ' *** TOTALES POR MES
    With GRILLAGRAL
        For co% = 3 To 15
            Acumulado# = 0
            For fi% = 1 To CantidadGastos&
                Acumulado# = Acumulado# + XVALO(.GRID.TextMatrix(fi%, co%))
            Next fi%
            .GRID.TextMatrix(CantidadGastos& + 2, co%) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
        Next co%
    End With
    
    ' *** % POR CUENTA
    With GRILLAGRAL
        TotalTodosGastos# = XVALO(.GRID.TextMatrix(CantidadGastos& + 2, 15))
        If TotalTodosGastos# > 0.005 Then
            For fi% = 1 To CantidadGastos&
                Acumulado# = XVALO(.GRID.TextMatrix(fi%, 15))
                Acumulado# = (Acumulado# / TotalTodosGastos#) * 100
                .GRID.TextMatrix(fi%, 16) = TRIM$(FORMATNUM$(Acumulado#, "F14.2"))
            Next fi%
        End If
    End With
    
    ' *** PROMEDIO
    With GRILLAGRAL
        For fi% = 1 To CantidadGastos&
            TotalGastos# = XVALO(.GRID.TextMatrix(fi%, 15))
            CantidadMesesConMovimiento# = 0
            For co% = 3 To 14
                ImporteMes# = XVALO(.GRID.TextMatrix(fi%, co%))
                If ImporteMes# > 0.005 Then CantidadMesesConMovimiento# = CantidadMesesConMovimiento# + 1
            Next co%
            Promedio# = 0
            If TotalGastos# > 0 Then Promedio# = (TotalGastos# / CantidadMesesConMovimiento#)
            .GRID.TextMatrix(fi%, 17) = TRIM$(FORMATNUM$(Promedio#, "F14.2"))
        Next fi%
    End With

    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Resumen de Egresos"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Private Sub mnuEquiposPorMes_Click()

    If DERACCE%("EQUIMENS", "Equipos por Mes") < 2 Then
        GoTo 99
    End If
    
10  Call SELECHO("SELMESA")
    MESEL$ = VALACT1$("SELMESA")
    If MESEL$ = "" Then GoTo 99
    MESANO MESEL$, DESDE%, HASTA%
    d$ = FECHATEX$(DESDE%)
    MesElegido% = XVALO(Mid$(d$, 4, 2)): AnioElegido% = 2000 + XVALO(Mid$(d$, 7, 2))
    d$ = "01/01/" & Mid$(d$, 7, 2): DESDE% = FECHANUM(d$)

    Opcion% = 1: GoTo 25
    Opcion% = COMALTER%("Códigos a evaluar:\\Tabla de Productos\Lista de Precios")
    
    Dim rst_VentaNetavsCMV As ADODB.Recordset
    
    ' *** RECUPERA DATOS DE LA BBDD
25  If Opcion% = 1 Then
        IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
        CodigoDesde$ = "": CodigoHasta$ = "": CondicionSql$ = ""
        If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
            ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
            Dim RstCodigos As ADODB.Recordset
            Set RstCodigos = READSET(ConsultaSql$)
            
            With RstCodigos
                Do While Not .EOF
                    CodigoProducto$ = SAFETEXT$(!CAMPO2)
                    If CondicionSql$ = "" Then
                        CondicionSql$ = "m.codigo='" & CodigoProducto$ & "'"
                    Else
                        CondicionSql$ = CondicionSql$ & " or m.codigo='" & CodigoProducto$ & "'"
                    End If
                    
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            RstCodigos.Close
            Set RstCodigos = Nothing
            On Error GoTo 0
        Else
            racod$ = OBJPLA$("DESHASCOD", VDEF$)
            If racod$ = "" Then GoTo 99
            CI1% = CVI(racod$)
            CI2% = CVI(Mid$(racod$, 3))
            If CI1% = 0 Or CI2% = 0 Then GoTo 99
            CodigoDesde$ = CODEXT$(CI1%)
            CodigoHasta$ = CODEXT$(CI2%)
            If CodigoDesde$ > CodigoHasta$ Then GoTo 99
        End If
        
        ' OBTIENE LISTA DE PRECIOS DE LA CUAL OBTENER EL PRECIO DE LOS CODIGOS VENDIDOS EN EL PERIODO ELEGIDO
        'CodigoLista& = ListaSeleccion_ListasDePrecios
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(8, CodigoLista&, CodigoDesde$, CodigoHasta$, DESDE%, HASTA%, CondicionSql$, TotalEquipos&)

    ElseIf Opcion% = 2 Then
        CodigoLista& = ListaSeleccion_ListasDePrecios
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(7, CodigoLista&, "", "", DESDE%, HASTA%, "", TotalEquipos&)
        
    End If
    
    ' *** ARMA ENCABEZADO DE EQUIPOS POR MES
    GRILLAGRAL.GRID.Clear
    With GRILLAGRAL.GRID
        .Rows = 2 + TotalEquipos& + 3
        .Cols = 17 '12 * 8 + 2
        .FixedCols = 2
        
        FilaEquipoMes& = 1
        .TextMatrix(FilaEquipoMes&, 0) = "Código"
        .TextMatrix(FilaEquipoMes&, 1) = "Descripción"
        .TextMatrix(FilaEquipoMes&, 2) = "ENERO"
        .TextMatrix(FilaEquipoMes&, 3) = "FEBRERO"
        .TextMatrix(FilaEquipoMes&, 4) = "MARZO"
        .TextMatrix(FilaEquipoMes&, 5) = "ABRIL"
        .TextMatrix(FilaEquipoMes&, 6) = "MAYO"
        .TextMatrix(FilaEquipoMes&, 7) = "JUNIO"
        .TextMatrix(FilaEquipoMes&, 8) = "JULIO"
        .TextMatrix(FilaEquipoMes&, 9) = "AGOSTO"
        .TextMatrix(FilaEquipoMes&, 10) = "SEPTIEMBRE"
        .TextMatrix(FilaEquipoMes&, 11) = "OCTUBRE"
        .TextMatrix(FilaEquipoMes&, 12) = "NOVIEMBRE"
        .TextMatrix(FilaEquipoMes&, 13) = "DICIEMBRE"
        .TextMatrix(FilaEquipoMes&, 14) = "Total año"
        .TextMatrix(FilaEquipoMes&, 15) = "%"
        .TextMatrix(FilaEquipoMes&, 16) = "Promedio"
        FilaEquipoMes& = FilaEquipoMes& + 1
        InicioEquipoMes& = FilaEquipoMes&
    End With
    
    GRILLAGRAL.GRID.TextMatrix(InicioEquipoMes& + TotalEquipos& + 1, 0) = "TOTALES"
    
    ' *** COLOREA EQUIPOS POR MES
    For co = 2 To 16
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), 1, co + 0, 1)
    Next co
    For fi = 0 To TotalEquipos& - 1
         For co = 0 To 11
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2 + fi, 2 + co, 1)
         Next co
         
         For co = 1 To 3
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 0), 2 + fi, 13 + co, 1)
         Next co
    Next fi
    For co = 2 To 16
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), InicioEquipoMes& + TotalEquipos& + 1, co + 0, 1)
    Next co

    ' *** ARMA ENCABEZADO PONDERACION
    FilaPonderacion& = FilaEquipoMes& + TotalEquipos& + 4
    With GRILLAGRAL.GRID
        .Rows = FilaPonderacion& + TotalEquipos& + 3
        .TextMatrix(FilaPonderacion&, 0) = "Código"
        .TextMatrix(FilaPonderacion&, 1) = "Descripción"
        For i = 0 To 0
            'Select Case i
            Select Case MesElegido% - 1
                Case 0: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "ENERO " & CStr(AnioElegido%)
                Case 1: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "FEBRERO " & CStr(AnioElegido%)
                Case 2: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "MARZO " & CStr(AnioElegido%)
                Case 3: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "ABRIL " & CStr(AnioElegido%)
                Case 4: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "MAYO " & CStr(AnioElegido%)
                Case 5: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "JUNIO " & CStr(AnioElegido%)
                Case 6: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "JULIO " & CStr(AnioElegido%)
                Case 7: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "AGOSTO " & CStr(AnioElegido%)
                Case 8: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "SEPTIEMBRE " & CStr(AnioElegido%)
                Case 9: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "OCTUBRE " & CStr(AnioElegido%)
                Case 10: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "NOVIEMBRE " & CStr(AnioElegido%)
                Case 11: .TextMatrix(FilaPonderacion& - 1, 10 * i + 2) = "DICIEMBRE " & CStr(AnioElegido%)
            End Select
            .TextMatrix(FilaPonderacion&, 10 * i + 2) = "Cantidad"
            .TextMatrix(FilaPonderacion&, 10 * i + 3) = "Fact.Neta"
            .TextMatrix(FilaPonderacion&, 10 * i + 4) = "Precio Un. Prom."
            .TextMatrix(FilaPonderacion&, 10 * i + 5) = "Lista de Precio"
            .TextMatrix(FilaPonderacion&, 10 * i + 6) = "Total x PL"
            .TextMatrix(FilaPonderacion&, 10 * i + 7) = "Desc. Promedio"
            .TextMatrix(FilaPonderacion&, 10 * i + 8) = "Costo"
            .TextMatrix(FilaPonderacion&, 10 * i + 9) = "Ventas vs Costos"
            .TextMatrix(FilaPonderacion&, 10 * i + 10) = "% Equipo"
            .TextMatrix(FilaPonderacion&, 10 * i + 11) = "% Facturación"
            .TextMatrix(FilaPonderacion&, 10 * i + 12) = "Dif Tot LDP-Fc.Neta"
        Next i
        FilaPonderacion& = FilaPonderacion& + 1
        InicioPonderacion& = FilaPonderacion&
    End With
    
    GRILLAGRAL.GRID.TextMatrix(InicioPonderacion& + TotalEquipos& + 1, 0) = "TOTALES"
        
    ' *** COLOREA PONDERACION
    With GRILLAGRAL
        For co = 1 To 11 '.GRID.Cols
            Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(207, 226, 243), InicioPonderacion& - 1, 2 + co - 1, 1)
            For fi = 1 To TotalEquipos&
                GRUPO% = Int(co / 11)
                If co Mod 11 = 0 Then GRUPO% = GRUPO% - 1
                Select Case GRUPO%
                    Case 0, 4, 8, 12: Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(192, 211, 226), InicioPonderacion& - 1 + fi, 2 + co - 1, 1)
                    Case 1, 5, 9: Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(162, 210, 196), InicioPonderacion& - 1 + fi, 2 + co - 1, 1)
                    Case 2, 6, 10: Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(201, 184, 143), InicioPonderacion& - 1 + fi, 2 + co - 1, 1)
                    Case 3, 7, 11: Call COLOREAR_GRILLA_SOLACELDA(.GRID, RGB(223, 157, 125), InicioPonderacion& - 1 + fi, 2 + co - 1, 1)
                End Select
            Next fi
        Next co
        
        For co = 2 To .GRID.Cols - 1
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), InicioPonderacion& + TotalEquipos& + 1, co + 0, 1)
        Next co
    End With
    
    ' *** ANCHO DE COLUMNAS Y JUSTIFICACION
    With GRILLAGRAL.GRID
        .ColWidth(0) = 1500
        .ColAlignment(0) = 1
        .ColWidth(1) = 3500
        .ColAlignment(1) = 1
        
        For co = 2 To .Cols - 1
            .ColWidth(co) = 1800
            .ColAlignment(co) = 6
        Next co
        
        fi = InicioPonderacion& - 2
        For co = 2 To 10 Step 10 '.Cols - 1 Step 8
            .Row = fi: .COL = co
            .CellAlignment = 1
            Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), fi, co + 0, 1)
        Next co
    End With
    
    ' *** DATOS
    Dim TotalesPorMes(12) As Double
    PrecioConIva% = XVALO(VALOBADA("listapre", "incluiva", "cod_lista=1"))
    With rst_VentaNetavsCMV
        Do While Not .EOF
            CodigoInterno% = !Codigo
            codigoExterno$ = CODEXT$(CodigoInterno%)
            Descripcion$ = DESCRIT0$(CodigoInterno%)
            MES% = !MES
            ANIO% = !ANIO
            CantidadMes# = XVALO(!SALIDA) - XVALO(!INGRESO)
            ImporteMes# = XVALO(!valunit_salida) - XVALO(!valunit_ingreso)
            Costo# = XVALO(!Costo)
            PRECIO# = XVALO(!preciolista)
            If PrecioConIva% = 1 Then
                TasaIva# = PORCEIVA(CodigoInterno%)
                PRECIO# = PRECIO# / (1 + (TasaIva# / 100))
            End If
            
            With GRILLAGRAL.GRID
                ' *** ESCRIBE SECTOR EQUIPOS POR MES
                PosicionEncontro& = 0
                For i = 1 To TotalEquipos&
                    If TRIM$(.TextMatrix(InicioEquipoMes& + i - 1, 0)) = codigoExterno$ Then
                        FilaCodigo& = InicioEquipoMes& + i - 1
                        PosicionEncontro& = i
                        Exit For
                    End If
                Next i
                If PosicionEncontro& = 0 Then FilaCodigo& = FilaEquipoMes&
                
                .TextMatrix(FilaCodigo&, 0) = codigoExterno$
                .TextMatrix(FilaCodigo&, 1) = Descripcion$
                
                ColumnaMes& = 2 + (MES% - 1)
                ValorAcumuladoMes# = XVALO(.TextMatrix(FilaCodigo&, ColumnaMes&))
                ValorAcumuladoMes# = ValorAcumuladoMes# + CantidadMes#
                .TextMatrix(FilaCodigo&, ColumnaMes&) = TRIM$(FORMATNUM$(ValorAcumuladoMes#, "F12.2"))
                
                If PosicionEncontro& = 0 Then FilaEquipoMes& = FilaEquipoMes& + 1
                
                TotalesPorMes(MES% - 1) = TotalesPorMes(MES% - 1) + CantidadMes#
                            
                ' *** ESCRIBE SECTOR PONDERACION
                FilaCodigo& = InicioPonderacion& + PosicionEncontro& - 1
                If PosicionEncontro& = 0 Then FilaCodigo& = FilaPonderacion&
                
                .TextMatrix(FilaCodigo&, 0) = codigoExterno$
                .TextMatrix(FilaCodigo&, 1) = Descripcion$
                
                If MES% = MesElegido% And ANIO% = AnioElegido% Then
                    ColumnaMes& = 2                    ' ColumnaMes& = 2 + 8 * (MES% - 1)
                    .TextMatrix(FilaCodigo&, ColumnaMes&) = TRIM$(FORMATNUM$(ValorAcumuladoMes#, "F12.2"))
                    ColumnaMes& = 2 + 1
                    ValorAcumuladoMes# = XVALO(.TextMatrix(FilaCodigo&, ColumnaMes&))
                    ValorAcumuladoMes# = ValorAcumuladoMes# + ImporteMes#
                    .TextMatrix(FilaCodigo&, ColumnaMes&) = TRIM$(FORMATNUM$(ValorAcumuladoMes#, "F12.2"))
                    
                    ColumnaMes& = 2 + 6                ' ColumnaMes& = 2 + 8 * (MES% - 1) + 4
                    .TextMatrix(FilaCodigo&, ColumnaMes&) = TRIM$(FORMATNUM$(Costo#, "F12.2"))
                    ColumnaMes& = 2 + 3                ' ColumnaMes& = 2 + 8 * (MES% - 1) + 2
                    .TextMatrix(FilaCodigo&, ColumnaMes&) = TRIM$(FORMATNUM$(PRECIO#, "F12.2"))
                End If
                
                If PosicionEncontro& = 0 Then FilaPonderacion& = FilaPonderacion& + 1
            End With

            .MoveNext
        Loop
    End With
    
    ' *** EQUIPOS POR MES
    With GRILLAGRAL.GRID
        ' TOTALES POR MES
        For i = 1 To 12
            .TextMatrix(InicioEquipoMes& + TotalEquipos& + 1, 1 + i) = TRIM$(FORMATNUM$(TotalesPorMes(i - 1), "F12.2"))
        Next i
        
        ' TOTALES POR CODIGO
        For fi = 0 To TotalEquipos& - 1
            ValorAcumuladoCodigo# = 0
            For co = 1 To 12
                ValorAcumuladoCodigo# = ValorAcumuladoCodigo# + XVALO(.TextMatrix(InicioEquipoMes& + fi, 1 + co))
            Next co
            .TextMatrix(InicioEquipoMes& + fi, 14) = TRIM$(FORMATNUM$(ValorAcumuladoCodigo#, "F12.2"))
            
            TotalCodigos# = TotalCodigos# + ValorAcumuladoCodigo#
        Next fi
        .TextMatrix(InicioEquipoMes& + TotalEquipos& + 1, 14) = TRIM$(FORMATNUM$(TotalCodigos#, "F12.2"))
        
        ' % PORCENTAJE
        For fi = 0 To TotalEquipos& - 1
            ValorAcumuladoCodigo# = XVALO(.TextMatrix(InicioEquipoMes& + fi, 14))
            Porcentaje# = 0
            If TotalCodigos# > 0.005 Then Porcentaje# = (ValorAcumuladoCodigo# / TotalCodigos#) * 100
            
            .TextMatrix(InicioEquipoMes& + fi, 15) = TRIM$(FORMATNUM$(Porcentaje#, "F12.2"))
        Next fi
        
        ' PROMEDIO
        MesesParaPromedio# = MesElegido%
        For fi = 0 To TotalEquipos& - 1
            ValorAcumuladoCodigo# = XVALO(.TextMatrix(InicioEquipoMes& + fi, 14))
            Porcentaje# = ValorAcumuladoCodigo# / MesesParaPromedio#
            
            .TextMatrix(InicioEquipoMes& + fi, 16) = TRIM$(FORMATNUM$(Porcentaje#, "F12.2"))
        Next fi
        
    End With
    
    ' *** PONDERACION
    With GRILLAGRAL.GRID
        For i = 0 To 0 '11
            TotalFacturacionNeta# = 0
            For fi = 0 To TotalEquipos& - 1
                
                TotalFacturacionNeta# = TotalFacturacionNeta# + XVALO(.TextMatrix(InicioPonderacion& + fi, 3 + 8 * i))
                
                ' PRECIO PROMEDIO
                'CantidadMes# = XVALO(.TextMatrix(InicioEquipoMes& + FI, 2 + i))
                CantidadMes# = XVALO(.TextMatrix(InicioEquipoMes& + fi, 2 + (MesElegido% - 1)))
                FacturacionNeta# = XVALO(.TextMatrix(InicioPonderacion& + fi, 3 + 8 * i))
                PrecioPromedio# = 0
                If CantidadMes# > 0.005 Then
                    PrecioPromedio# = FacturacionNeta# / CantidadMes#
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 1) = TRIM$(FORMATNUM$(PrecioPromedio#, "F12.2"))
                End If
                
                ' TOTAL FACTURA SEGUN LISTA DE PRECIO
                preciolista# = XVALO(.TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 2))
                TotalSegunListaPrecio# = 0
                If preciolista# > 0.005 Then
                    TotalSegunListaPrecio# = CantidadMes# * preciolista#
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 3) = TRIM$(FORMATNUM$(TotalSegunListaPrecio#, "F12.2"))
                End If
                
                ' DESCUENTO OTORGADO
                If preciolista# > 0.005 And PrecioPromedio# > 0.005 Then
                    DescuentoOtorgado# = ((PrecioPromedio# / preciolista#) - 1) * 100
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 4) = TRIM$(FORMATNUM$(DescuentoOtorgado#, "F12.2"))
                End If
                
                ' PRECIO PROMEDIO VS COSTO
                Costo# = XVALO(.TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 5))
                If Costo# > 0.005 And PrecioPromedio# > 0.005 Then
                    PromedioCosto# = (PrecioPromedio# / Costo#)
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 6) = TRIM$(FORMATNUM$(PromedioCosto#, "F12.2"))
                End If

                ' % EQUIPOS
                'TotalEquiposMes# = XVALO(.TextMatrix(InicioEquipoMes& + TotalEquipos + 1, 2 + i))
                TotalEquiposMes# = XVALO(.TextMatrix(InicioEquipoMes& + TotalEquipos + 1, 2 + (MesElegido% - 1)))
                If TotalEquiposMes# > 0.005 Then
                    PorcentajeEquipos# = (CantidadMes# / TotalEquiposMes#) * 100
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 7) = TRIM$(FORMATNUM$(PorcentajeEquipos#, "F12.2"))
                End If
                
                ' DIF FACTURADO SEGUN LISTA DE PRECIO - FC. NETA
                .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 9) = TRIM$(FORMATNUM$(TotalSegunListaPrecio# - FacturacionNeta#, "F12.2"))
                
            Next fi
            
            ' TOTALES FACTURACION NETA
            .TextMatrix(InicioPonderacion& + TotalEquipos& + 1, 3 + (8 * i)) = TRIM$(FORMATNUM$(TotalFacturacionNeta#, "F12.2"))
            
        Next i
        
        For i = 0 To 0 '11
            For fi = 0 To TotalEquipos& - 1
            
                '% FACTURACION
                TotalFacturacionNeta# = XVALO(.TextMatrix(InicioPonderacion& + TotalEquipos& + 1, 3 + (8 * i)))
                If TotalFacturacionNeta# > 0.005 Then
                    FacturacionNeta# = XVALO(.TextMatrix(InicioPonderacion& + fi, 3 + 8 * i))
                    PorcentajeFacturacion# = (FacturacionNeta# / TotalFacturacionNeta#) * 100
                    .TextMatrix(InicioPonderacion& + fi, 3 + (8 * i) + 8) = TRIM$(FORMATNUM$(PorcentajeFacturacion#, "F12.2"))
                End If
                
            Next fi
        Next i
        
    End With
    
    On Error Resume Next
    rst_VentaNetavsCMV.Close
    Set rst_VentaNetavsCMV = Nothing
    On Error GoTo 0
    
    ' ********************
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Equipos Vendidos por Mes (" & FECHATEX$(DESDE%) & " - " & FECHATEX$(HASTA%) & ")"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub


Private Sub mnuEstadisticasTablaProductos_Click()

    If DERACCE%("CFGESVE", "Configuración de Estadísticas de Ventas") < 2 Then
        GoTo 99
    End If
    
    Call TablasAuxiliares
    
    IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
    Call ABM_TABLA(IDLIST&, , , "campo2")
    
99  Exit Sub

End Sub

Private Sub mnuEstadisticasTablaVendedores_Click()

    If DERACCE%("CFGESVE", "Configuración de Estadísticas de Ventas") < 2 Then
        GoTo 99
    End If
    
    Call TablasAuxiliares
    
    On Error Resume Next
    FLEXCONN.Execute "delete #TPvendedores"
    ConsultaSql$ = "create table #TPvendedores (codigo int, denominacion varchar(64))"
    FLEXCONN.Execute ConsultaSql$
    On Error GoTo 0
    
    On Error GoTo ControlError
    Dim Rst As New ADODB.Recordset
    ConsultaSql$ = "select * from #TPvendedores"
    Rst.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText
    With Rst
        For i& = 1 To ULTREG&("VENDEDOR")
            regi$ = REGLEIDO$("VENDEDOR", i&)
            Codigo& = Asc(Mid$(regi$, 1, 1))
            Denominacion$ = Mid$(regi$, 2, 30)
                    
            .AddNew
            !Codigo = Codigo&
            !Denominacion = Denominacion$
            .Update
                
        Next i&
    End With
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    IDLIST& = ObtenerIdTablaAuxiliar("ESTADIVENDE")
    Call ABM_TABLA(IDLIST&, , , "campo2")
    Exit Sub

ControlError:
    Exit Sub
    
99  Exit Sub

End Sub

Private Sub mnuFacturaNetavsCmvGrupo_Click()

    If DERACCE%("VTCMVGRU", "Facturacion de Ventas vs CMV Agrupado") < 2 Then
        GoTo 99
    End If
    
    Call DESHASFECHA(DESDE%, HASTA%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    
    Opcion% = COMALTER%("Códigos a evaluar:\\Tabla de Productos\Lista de Precios")
    
    Dim rst_VentaNetavsCMV As ADODB.Recordset
    
25  If Opcion% = 1 Then
        IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
        CodigoDesde$ = "": CodigoHasta$ = "": CondicionSql$ = ""
        If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
            ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
            Dim RstCodigos As ADODB.Recordset
            Set RstCodigos = READSET(ConsultaSql$)
            
            With RstCodigos
                Do While Not .EOF
                    CodigoProducto$ = SAFETEXT$(!CAMPO2)
                    If CondicionSql$ = "" Then
                        CondicionSql$ = "m.codigo='" & CodigoProducto$ & "'"
                    Else
                        CondicionSql$ = CondicionSql$ & " or m.codigo='" & CodigoProducto$ & "'"
                    End If
                    
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            RstCodigos.Close
            Set RstCodigos = Nothing
            On Error GoTo 0
        Else
            racod$ = OBJPLA$("DESHASCOD", VDEF$)
            If racod$ = "" Then GoTo 99
            CI1% = CVI(racod$)
            CI2% = CVI(Mid$(racod$, 3))
            If CI1% = 0 Or CI2% = 0 Then GoTo 99
            CodigoDesde$ = CODEXT$(CI1%)
            CodigoHasta$ = CODEXT$(CI2%)
            If CodigoDesde$ > CodigoHasta$ Then GoTo 99
        End If
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(5, 0, CodigoDesde$, CodigoHasta$, DESDE%, HASTA%, CondicionSql$)

    ElseIf Opcion% = 2 Then
        Campos$ = "Id/I5;Denominación/A32"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Caption = "Lista de Precios"
        
        ConsultaSql$ = "select cod_lista, denom_lista from listapre with(nolock) order by cod_lista"
        Dim rst_ListasPrecio As ADODB.Recordset
        Set rst_ListasPrecio = READSET(ConsultaSql$)
        
        FRMZELECHO.MSF2.Rows = rst_ListasPrecio.RecordCount + 1
        fi& = 0
        With rst_ListasPrecio
            Do While Not .EOF
                fi& = fi& + 1
                FRMZELECHO.MSF2.TextMatrix(fi&, 1) = CStr(!Cod_Lista)
                FRMZELECHO.MSF2.TextMatrix(fi&, 2) = !DENOM_LISTA
            
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        rst_ListasPrecio.Close
        Set rst_ListasPrecio = Nothing
        On Error GoTo 0
        
        FRMZELECHO.Width = 5000
        FRMZELECHO.Show 1
    
        CodigoLista& = XVALO(RESP_ZELE_VECT(1))
        If CodigoLista& = 0 Then GoTo 99
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(4, CodigoLista&, "", "", DESDE%, HASTA%, "")
        
    End If
       
    GRILLAGRAL.GRID.Clear
    Campos$ = "Código/A16;Descripción/A32;Tot. Cantidad/F12.1;Tot.Fact.Neta/F14.2;Total CMV/F14.2;Dif Neto vs CMV/F16.2;% Neto vs CMV/F14.2;Tot unid grupo/F12.1;% Tot Unid Grupo/F12.1;Tot Neto Grupo/F14.2;% Neto Grupo/F14.2;Total CMV Grupo/F14.2;% Tot CMV Grupo/F14.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, FRMZELECHO, , , ANTOT)
    GRILLAGRAL.GRID.Rows = 256
    
    With rst_VentaNetavsCMV
        fi = 1
        TotalCantidad# = 0: TTotalCantidad# = 0
        TotalFacturacionNeta# = 0: TTotalFacturacionNeta# = 0
        TotalCMV# = 0: TTotalCMV# = 0
        FilaGrupo1 = 0: FilaGrupo2 = 0: FilaGrupo3 = 0: FilaGrupo4 = 0
        Dim Grupo1 As Boolean, Grupo2 As Boolean, Grupo3 As Boolean, Grupo4 As Boolean
        
        Do While Not .EOF
            
            fi = fi + 1
            If GRILLAGRAL.GRID.Rows < fi Then GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
            
            GRILLAGRAL.GRID.TextMatrix(fi, 1) = CODEXT$(!Codigo)
            GRILLAGRAL.GRID.TextMatrix(fi, 2) = DESCRIT0$(!Codigo)
            GRILLAGRAL.GRID.TextMatrix(fi, 3) = FORMATNUM$(!Cantidad, "F12.1")
            GRILLAGRAL.GRID.TextMatrix(fi, 4) = FORMATNUM$(!facturacion, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 5) = FORMATNUM$(!cmv, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 6) = FORMATNUM$(!dife_fccmv, "F16.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 7) = FORMATNUM$(!porce_fccmv, "F14.2")
            
            PorcentajeNetoCMV# = XVALO(!porce_fccmv)
            If PorcentajeNetoCMV# > 0 And PorcentajeNetoCMV# <= 1 Then Grupo1 = True
            If PorcentajeNetoCMV# > 1 And PorcentajeNetoCMV# <= 2 Then Grupo2 = True
            If PorcentajeNetoCMV# > 2 And PorcentajeNetoCMV# <= 3 Then Grupo3 = True
            If PorcentajeNetoCMV# > 3 Then Grupo4 = True
            
            If Grupo1 And PorcentajeNetoCMV# > 1 Then
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 12) = FORMATNUM$(TotalCMV#, "F14.2")
                Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi - 1)
                FilaGrupo1 = fi - 1
                TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
                Grupo1 = False
            
            ElseIf Grupo2 And PorcentajeNetoCMV# > 2 Then
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 12) = FORMATNUM$(TotalCMV#, "F14.2")
                Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi - 1)
                FilaGrupo2 = fi - 1
                TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
                Grupo2 = False
            
            ElseIf Grupo3 And PorcentajeNetoCMV# > 3 Then
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
                GRILLAGRAL.GRID.TextMatrix(fi - 1, 12) = FORMATNUM$(TotalCMV#, "F14.2")
                Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi - 1)
                FilaGrupo3 = fi - 1
                TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
                Grupo3 = False
                        
            End If
            
            TotalCantidad# = TotalCantidad# + XVALO(!Cantidad)
            TotalFacturacionNeta# = TotalFacturacionNeta# + XVALO(!facturacion)
            TotalCMV# = TotalCMV# + XVALO(!cmv)
            
            TTotalCantidad# = TTotalCantidad# + XVALO(!Cantidad)
            TTotalFacturacionNeta# = TTotalFacturacionNeta# + XVALO(!facturacion)
            TTotalCMV# = TTotalCMV# + XVALO(!cmv)
            
            .MoveNext
        Loop
        
        If Grupo1 Then
            GRILLAGRAL.GRID.TextMatrix(fi, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 12) = FORMATNUM$(TotalCMV#, "F14.2")
            Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi)
            FilaGrupo1 = fi
            TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
        
        ElseIf Grupo2 Then
            GRILLAGRAL.GRID.TextMatrix(fi, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 12) = FORMATNUM$(TotalCMV#, "F14.2")
            Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi)
            FilaGrupo2 = fi
            TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
        
        ElseIf Grupo3 Then
            GRILLAGRAL.GRID.TextMatrix(fi, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 12) = FORMATNUM$(TotalCMV#, "F14.2")
            Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi)
            FilaGrupo3 = fi
            TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
        
        ElseIf Grupo4 Then
            GRILLAGRAL.GRID.TextMatrix(fi, 8) = FORMATNUM$(TotalCantidad#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 10) = FORMATNUM$(TotalFacturacionNeta#, "F14.2")
            GRILLAGRAL.GRID.TextMatrix(fi, 12) = FORMATNUM$(TotalCMV#, "F14.2")
            Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi)
            FilaGrupo4 = fi
            TotalCantidad# = 0: TotalFacturacionNeta# = 0: TotalCMV# = 0
        
        End If
    
    End With
    
    On Error Resume Next
    rst_VentaNetavsCMV.Close
    Set rst_VentaNetavsCMV = Nothing
    On Error GoTo 0
    
    If FilaGrupo1 > 0 Then
        TotalCantidad# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 8))
        GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 9) = FORMATNUM$((TotalCantidad# / TTotalCantidad#) * 100, "F12.1")
        TotalFacturacionNeta# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 10))
        GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 11) = FORMATNUM$((TotalFacturacionNeta# / TTotalFacturacionNeta#) * 100, "F12.1")
        TotalCMV# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 12))
        GRILLAGRAL.GRID.TextMatrix(FilaGrupo1, 13) = FORMATNUM$((TotalCMV# / TTotalCMV#) * 100, "F12.1")
        
    End If
    If FilaGrupo2 > 0 Then
        TotalCantidad# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 8))
        If TotalCantidad# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 9) = FORMATNUM$((TotalCantidad# / TTotalCantidad#) * 100, "F12.1")
        TotalFacturacionNeta# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 10))
        If TotalFacturacionNeta# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 11) = FORMATNUM$((TotalFacturacionNeta# / TTotalFacturacionNeta#) * 100, "F12.1")
        TotalCMV# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 12))
        If TotalCMV# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo2, 13) = FORMATNUM$((TotalCMV# / TTotalCMV#) * 100, "F12.1")
    
    End If
    If FilaGrupo3 > 0 Then
        TotalCantidad# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 8))
        If TotalCantidad# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 9) = FORMATNUM$((TotalCantidad# / TTotalCantidad#) * 100, "F12.1")
        TotalFacturacionNeta# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 10))
        If TotalFacturacionNeta# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 11) = FORMATNUM$((TotalFacturacionNeta# / TTotalFacturacionNeta#) * 100, "F12.1")
        TotalCMV# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 12))
        If TotalCMV# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo3, 13) = FORMATNUM$((TotalCMV# / TTotalCMV#) * 100, "F12.1")
        
    End If
    If FilaGrupo4 > 0 Then
        TotalCantidad# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 8))
        If TotalCantidad# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 9) = FORMATNUM$((TotalCantidad# / TTotalCantidad#) * 100, "F12.1")
        TotalFacturacionNeta# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 10))
        If TotalFacturacionNeta# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 11) = FORMATNUM$((TotalFacturacionNeta# / TTotalFacturacionNeta#) * 100, "F12.1")
        TotalCMV# = XVALO(GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 12))
        If TotalCMV# > 0.005 Then GRILLAGRAL.GRID.TextMatrix(FilaGrupo4, 13) = FORMATNUM$((TotalCMV# / TTotalCMV#) * 100, "F12.1")
        
    End If
    
    fi = fi + 2
    GRILLAGRAL.GRID.TextMatrix(fi, 2) = "TOTALES"
    GRILLAGRAL.GRID.TextMatrix(fi, 3) = FORMATNUM$(TTotalCantidad#, "F12.1")
    GRILLAGRAL.GRID.TextMatrix(fi, 4) = FORMATNUM$(TTotalFacturacionNeta#, "F14.2")
    GRILLAGRAL.GRID.TextMatrix(fi, 5) = FORMATNUM$(TTotalCMV#, "F14.2")
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi)
    
    GRILLAGRAL.GRID.Rows = fi + 1
    
    For fi = 1 To GRILLAGRAL.GRID.Rows - 1
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), fi, 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), fi, 2, 1)
    Next fi
        
    GRILLAGRAL.Caption = "Comparativo Ventas vs CMV (Totales por Equipo - Agrupado)"
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0

    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, "Comparativo Ventas vs CMV", False)
        End If
   End If
       
99  Exit Sub
End Sub

Private Sub mnuFacturaNetavsCmvMes_Click()

    If DERACCE%("VTCMVMEN", "Facturacion de Ventas vs CMV") < 2 Then
        GoTo 99
    End If
    
10  Call DESHASFECHA(DESDE%, HASTA%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    
    Dim r As New ADODB.Recordset
    ConsultaSql$ = "select datediff(m,'" & FETEXCOR1$(DESDE%) & "','" & FETEXCOR1$(HASTA%) & "') meses"
    r.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    CantidadDeMeses% = r!meses + 1
    r.Close
    If CantidadDeMeses% > 16 Then
        Call COMUNI("Imposible Generar el Reporte.\ \Ha Superado el Máximo de Meses Soportado de 16.")
        GoTo 10
    End If

    Opcion% = COMALTER%("Códigos a evaluar:\\Tabla de Productos\Lista de Precios")
    
    Dim rst_VentaNetavsCMV As ADODB.Recordset
    
25  If Opcion% = 1 Then
        IDLIST& = ObtenerIdTablaAuxiliar("ESTADIPRODUC")
        CodigoDesde$ = "": CodigoHasta$ = "": CondicionSql$ = ""
        If CantRegistros("datasql", "id_lista=" & CStr(IDLIST&)) > 0 Then
            ConsultaSql$ = "select campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
            Dim RstCodigos As ADODB.Recordset
            Set RstCodigos = READSET(ConsultaSql$)
            
            With RstCodigos
                Do While Not .EOF
                    CodigoProducto$ = SAFETEXT$(!CAMPO2)
                    If CondicionSql$ = "" Then
                        CondicionSql$ = "m.codigo='" & CodigoProducto$ & "'"
                    Else
                        CondicionSql$ = CondicionSql$ & " or m.codigo='" & CodigoProducto$ & "'"
                    End If
                    
                    .MoveNext
                Loop
            End With
            On Error Resume Next
            RstCodigos.Close
            Set RstCodigos = Nothing
            On Error GoTo 0
        Else
            racod$ = OBJPLA$("DESHASCOD", VDEF$)
            If racod$ = "" Then GoTo 99
            CI1% = CVI(racod$)
            CI2% = CVI(Mid$(racod$, 3))
            If CI1% = 0 Or CI2% = 0 Then GoTo 99
            CodigoDesde$ = CODEXT$(CI1%)
            CodigoHasta$ = CODEXT$(CI2%)
            If CodigoDesde$ > CodigoHasta$ Then GoTo 99
        End If
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(2, 0, CodigoDesde$, CodigoHasta$, DESDE%, HASTA%, CondicionSql$)

    ElseIf Opcion% = 2 Then
        Campos$ = "Id/I5;Denominación/A32"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Caption = "Lista de Precios"
        
        ConsultaSql$ = "select cod_lista, denom_lista from listapre with(nolock) order by cod_lista"
        Dim rst_ListasPrecio As ADODB.Recordset
        Set rst_ListasPrecio = READSET(ConsultaSql$)
        
        FRMZELECHO.MSF2.Rows = rst_ListasPrecio.RecordCount + 1
        fi& = 0
        With rst_ListasPrecio
            Do While Not .EOF
                fi& = fi& + 1
                FRMZELECHO.MSF2.TextMatrix(fi&, 1) = CStr(!Cod_Lista)
                FRMZELECHO.MSF2.TextMatrix(fi&, 2) = !DENOM_LISTA
            
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        rst_ListasPrecio.Close
        Set rst_ListasPrecio = Nothing
        On Error GoTo 0
        
        FRMZELECHO.Width = 5000
        FRMZELECHO.Show 1
    
        CodigoLista& = XVALO(RESP_ZELE_VECT(1))
        If CodigoLista& = 0 Then GoTo 99
        
        Set rst_VentaNetavsCMV = VentaNetavsCMV(1, CodigoLista&, "", "", DESDE%, HASTA%, "")
        
    End If
    
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    
    GRILLAGRAL.GRID.Clear
    GRILLAGRAL.GRID.Cols = 2 + (4 * CantidadDeMeses%) + 5
    GRILLAGRAL.GRID.Rows = 256

    MesDesde% = Month(CVDAT(DESDE%))
    meshasta% = Month(CVDAT(HASTA%))
    AnioDesde% = Year(CVDAT(DESDE%))
    AnioHasta% = Year(CVDAT(HASTA%))
    
    c = 1
    For i = AnioDesde% To AnioHasta%
        If AnioDesde% = AnioHasta% Then
            If masdeunanio = 0 Then
                mesini% = MesDesde%
                mesfin% = meshasta%
            Else
                mesini% = 1
                mesfin% = meshasta%
            End If
        Else
            mesini% = MesDesde%
            mesfin% = 12
            masdeunanio = 1
            AnioDesde% = AnioDesde% + 1
        End If
        For J = mesini% To mesfin%
            GRILLAGRAL.GRID.TextMatrix(0, 1 + c) = CStr(J) & CStr(i) ' mes y anio
            
            GRILLAGRAL.GRID.TextMatrix(1, 1 + c) = MES$(J)
                GRILLAGRAL.GRID.ColWidth(1 + c) = 2500
                Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(207, 226, 243), 1, 1 + c, 1)
                
            GRILLAGRAL.GRID.TextMatrix(2, 1 + c) = "Cantidad"
                GRILLAGRAL.GRID.ColWidth(1 + c) = 2500: GRILLAGRAL.GRID.ColAlignment(1 + c) = 6
                Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, 1 + c, 1)
            GRILLAGRAL.GRID.TextMatrix(2, 1 + c + 1) = "Fact. Neta"
                GRILLAGRAL.GRID.ColWidth(1 + c + 1) = 2500: GRILLAGRAL.GRID.ColAlignment(1 + c + 1) = 6
                Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, 1 + c + 1, 1)
            GRILLAGRAL.GRID.TextMatrix(2, 1 + c + 2) = "Costo"
                GRILLAGRAL.GRID.ColWidth(1 + c + 2) = 2500: GRILLAGRAL.GRID.ColAlignment(1 + c + 2) = 6
                Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, 1 + c + 2, 1)
            GRILLAGRAL.GRID.TextMatrix(2, 1 + c + 3) = "CMV"
                GRILLAGRAL.GRID.ColWidth(1 + c + 3) = 2500: GRILLAGRAL.GRID.ColAlignment(1 + c + 3) = 6
                Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, 1 + c + 3, 1)
            c = c + 4
        Next J
    Next i
    
    GRILLAGRAL.GRID.TextMatrix(0, 0) = "Código"
        GRILLAGRAL.GRID.ColWidth(0) = 1500
    GRILLAGRAL.GRID.TextMatrix(0, 1) = "Descripción"
        GRILLAGRAL.GRID.ColWidth(1) = 2500
    
    GRILLAGRAL.GRID.TextMatrix(2, GRILLAGRAL.GRID.Cols - 5) = "Tot.Cant."
        GRILLAGRAL.GRID.ColWidth(GRILLAGRAL.GRID.Cols - 5) = 3000: GRILLAGRAL.GRID.ColAlignment(GRILLAGRAL.GRID.Cols - 5) = 6
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, GRILLAGRAL.GRID.Cols - 5, 1)
    GRILLAGRAL.GRID.TextMatrix(2, GRILLAGRAL.GRID.Cols - 4) = "Tot.Venta"
        GRILLAGRAL.GRID.ColWidth(GRILLAGRAL.GRID.Cols - 4) = 3000: GRILLAGRAL.GRID.ColAlignment(GRILLAGRAL.GRID.Cols - 4) = 6
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, GRILLAGRAL.GRID.Cols - 4, 1)
    GRILLAGRAL.GRID.TextMatrix(2, GRILLAGRAL.GRID.Cols - 3) = "Tot.CMV"
        GRILLAGRAL.GRID.ColWidth(GRILLAGRAL.GRID.Cols - 3) = 3000: GRILLAGRAL.GRID.ColAlignment(GRILLAGRAL.GRID.Cols - 3) = 6
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, GRILLAGRAL.GRID.Cols - 3, 1)
    GRILLAGRAL.GRID.TextMatrix(2, GRILLAGRAL.GRID.Cols - 2) = "Dif. Venta-CMV"
        GRILLAGRAL.GRID.ColWidth(GRILLAGRAL.GRID.Cols - 2) = 3000: GRILLAGRAL.GRID.ColAlignment(GRILLAGRAL.GRID.Cols - 2) = 6
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, GRILLAGRAL.GRID.Cols - 2, 1)
    GRILLAGRAL.GRID.TextMatrix(2, GRILLAGRAL.GRID.Cols - 1) = "% Venta-CMV"
        GRILLAGRAL.GRID.ColWidth(GRILLAGRAL.GRID.Cols - 1) = 3000: GRILLAGRAL.GRID.ColAlignment(GRILLAGRAL.GRID.Cols - 1) = 6
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), 2, GRILLAGRAL.GRID.Cols - 1, 1)
    
    With rst_VentaNetavsCMV
        fi = 2
        ci_actual% = 0: CI_ANT% = 0
        TotalCantidad# = 0
        TotalFacturacionNeta# = 0
        TotalCMV# = 0
        Do While Not .EOF
            
            If IsNull(!MES) = False And IsNull(!ANIO) = False Then
                MESANIO$ = CStr(!MES) & CStr(!ANIO)
                For c = 2 To GRILLAGRAL.GRID.Cols - 1 Step 4
                    If GRILLAGRAL.GRID.TextMatrix(0, c) = MESANIO$ Then Exit For
                Next c
            End If
                                   
            ci_actual% = !Codigo
            If ci_actual% <> CI_ANT% Then
                If fi > 2 Then
                    GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 5) = FORMATNUM$(TotalCantidad#, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 4) = FORMATNUM$(TotalFacturacionNeta#, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 3) = FORMATNUM$(TotalCMV#, "F12.2")
                    GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 2) = FORMATNUM$(TotalFacturacionNeta# - TotalCMV#, "F12.2")
                    If TotalCMV# > 0.005 Then
                        GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 1) = FORMATNUM$((TotalFacturacionNeta# / TotalCMV#) * 100, "F12.2")
                    End If
                End If
                
                fi = fi + 1
                CI_ANT% = ci_actual%
                If GRILLAGRAL.GRID.Rows < fi Then GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                
                GRILLAGRAL.GRID.TextMatrix(fi, 0) = CODEXT$(ci_actual%)
                GRILLAGRAL.GRID.TextMatrix(fi, 1) = DESCRIT0$(ci_actual%)
                TotalCantidad# = 0
                TotalFacturacionNeta# = 0
                TotalCMV# = 0
            End If
            
                Cantidad# = XVALO(!SALIDA) - XVALO(!INGRESO)
            GRILLAGRAL.GRID.TextMatrix(fi, c) = FORMATNUM$(Cantidad#, "F10.1")
                VentaNeta# = XVALO(!valunit_salida) - XVALO(!valunit_ingreso)
            GRILLAGRAL.GRID.TextMatrix(fi, c + 1) = FORMATNUM$(VentaNeta#, "F12.2")
            GRILLAGRAL.GRID.TextMatrix(fi, c + 2) = FORMATNUM$(!Costo, "F12.5")
                cmv# = Cantidad# * XVALO(!Costo)
            GRILLAGRAL.GRID.TextMatrix(fi, c + 3) = FORMATNUM$(cmv#, "F12.2")
            
            TotalCantidad# = TotalCantidad# + Cantidad#
            TotalFacturacionNeta# = TotalFacturacionNeta# + VentaNeta#
            TotalCMV# = TotalCMV# + cmv#
            
20          .MoveNext
        Loop
        
        GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 5) = FORMATNUM$(TotalCantidad#, "F12.2")
        GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 4) = FORMATNUM$(TotalFacturacionNeta#, "F12.2")
        GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 3) = FORMATNUM$(TotalCMV#, "F12.2")
        GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 2) = FORMATNUM$(TotalFacturacionNeta# - TotalCMV#, "F12.2")
        If TotalCMV# > 0.005 Then
            GRILLAGRAL.GRID.TextMatrix(fi, GRILLAGRAL.GRID.Cols - 1) = FORMATNUM$((TotalFacturacionNeta# / TotalCMV#) * 100, "F12.2")
        End If

    End With
    
    On Error Resume Next
    rst_VentaNetavsCMV.Close
    Set rst_VentaNetavsCMV = Nothing
    On Error GoTo 0
    
    ' CALCULA TOTALES POR MES
    GRILLAGRAL.GRID.TextMatrix(fi + 2, 1) = "TOTALES"
    For c = 2 To GRILLAGRAL.GRID.Cols - 5 Step 4
        TotalCantidad# = 0
        TotalFacturacionNeta# = 0
        TotalCMV# = 0

        For i = 3 To fi
            TotalCantidad# = TotalCantidad# + XVALO(GRILLAGRAL.GRID.TextMatrix(i, c))
            TotalFacturacionNeta# = TotalFacturacionNeta# + XVALO(GRILLAGRAL.GRID.TextMatrix(i, c + 1))
            If c < GRILLAGRAL.GRID.Cols - 5 Then
                TotalCMV# = TotalCMV# + XVALO(GRILLAGRAL.GRID.TextMatrix(i, c + 3))
            Else
                TotalCMV# = TotalCMV# + XVALO(GRILLAGRAL.GRID.TextMatrix(i, c + 2))
            End If
        Next i
        
        GRILLAGRAL.GRID.TextMatrix(fi + 2, c) = FORMATNUM$(TotalCantidad#, "F12.2")
        GRILLAGRAL.GRID.TextMatrix(fi + 2, c + 1) = FORMATNUM$(TotalFacturacionNeta#, "F12.2")
        If c < GRILLAGRAL.GRID.Cols - 5 Then
            GRILLAGRAL.GRID.TextMatrix(fi + 2, c + 3) = FORMATNUM$(TotalCMV#, "F12.2")
        Else
            GRILLAGRAL.GRID.TextMatrix(fi + 2, c + 2) = FORMATNUM$(TotalCMV#, "F12.2")
        End If
        
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi + 2, c + 0, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi + 2, c + 1, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi + 2, c + 2, 1)
        Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi + 2, c + 3, 1)
    Next c
    Call COLOREAR_GRILLA_SOLACELDA(GRILLAGRAL.GRID, RGB(241, 194, 50), fi + 2, GRILLAGRAL.GRID.Cols - 1, 1)
    
    GRILLAGRAL.GRID.Rows = fi + 3
    GRILLAGRAL.GRID.FixedCols = 2
    GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.Caption = "Comparativo Ventas vs CMV"
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
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
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, "Comparativo Ventas vs CMV", False)
        End If
   End If
   Set obj = Nothing
       
99  Exit Sub

End Sub

Function VentaNetavsCMV(MODO%, CodigoLista&, CodigoDesde$, CodigoHasta$, DesdeFecha%, HastaFecha%, CondicionSqlCodigos$, Optional CantidadRegistros&) As ADODB.Recordset

    Dim rst_costos As New ADODB.Recordset
    Dim rst_ventasvscmv As New ADODB.Recordset
    Dim rst_totalequipos As ADODB.Recordset
    
    If MODO% = 1 Then
        ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select fecoscal, h.cod_inte from histocostos h inner join preventa p on (h.cod_inte=p.cod_pieza) WHERE p.cod_lista=" & CStr(CodigoLista&) & ") h" _
        & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        
        ' ( (preventa left join (cajabanc inner join movisto) ) left join costos
        ConsultaSql$ = "select * from " _
        & " (select ilp_vn.codigo, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo from" _
        & " (select codigo, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from preventa inner join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (preventa.cod_pieza=vn.codigo)" _
        & " where preventa.cod_lista=" & CStr(CodigoLista&) & ") ilp_vn" _
        & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " order by codigo, mes, anio"
        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv
        
    ElseIf MODO% = 2 Then
            ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select h.fecoscal, h.cod_inte from histocostos h inner join master m on (h.cod_inte=m.codint) "
        If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') h"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") h"
        End If
        ConsultaSql$ = ConsultaSql$ & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        Text1 = ConsultaSql$

        ' ( (master inner join (cajabanc inner join movisto) ) left join costos
        ' con master se filtra por rango de codigos o por tabla de productos
        ' con el master inner join solo se trabaja con aquellos codigos que tienen movimientos.
        ConsultaSql$ = "select * from " _
        & " (select ilp_vn.codigo, codexte, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo from" _
        & " (select m.codint codigo, m.codigo codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from master m left join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (m.codint=vn.codigo)"
        If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') ilp_vn"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") ilp_vn"
        End If
        ConsultaSql$ = ConsultaSql$ & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " order by codexte, mes, anio"
        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv
        Text2 = ConsultaSql$
    ElseIf MODO% = 4 Then
        ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select fecoscal, h.cod_inte from histocostos h inner join preventa p on (h.cod_inte=p.cod_pieza) WHERE p.cod_lista=" & CStr(CodigoLista&) & ") h" _
        & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        
        ' ( (preventa left join (cajabanc inner join movisto) ) left join costos
        ConsultaSql$ = "select * from" _
        & " (select codigo, cantidad, facturacion, cmv, (facturacion-cmv) dife_fccmv, case when cmv=0 then 0 else facturacion/cmv end as porce_fccmv from" _
        & " (select codigo, sum(salida-ingreso) cantidad, sum(valunit_salida-valunit_ingreso) facturacion, sum((salida-ingreso)*costo) cmv from" _
        & " (select ilp_vn.codigo, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo from" _
        & " (select codigo, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from preventa inner join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (preventa.cod_pieza=vn.codigo)" _
        & " where preventa.cod_lista=" & CStr(CodigoLista&) & ") ilp_vn" _
        & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " group by codigo) formatsalida1) formatosalida1" _
        & " order by porce_fccmv"

        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv
        
    ElseIf MODO% = 5 Then
        ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select h.fecoscal, h.cod_inte from histocostos h inner join master m on (h.cod_inte=m.codint) "
                If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') h"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") h"
        End If
        ConsultaSql$ = ConsultaSql$ & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        
        ' ( (master inner join (cajabanc inner join movisto) ) left join costos
        ' con master filtra por rango de codigos o por tabla de productos
        ConsultaSql$ = "select * from" _
        & " (select codigo, codexte, cantidad, facturacion, cmv, (facturacion-cmv) dife_fccmv, case when cmv=0 then 0 else facturacion/cmv end as porce_fccmv from" _
        & " (select codigo, codexte, sum(salida-ingreso) cantidad, sum(valunit_salida-valunit_ingreso) facturacion, sum((salida-ingreso)*costo) cmv from" _
        & " (select ilp_vn.codigo, codexte, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo from" _
        & " (select m.codint codigo, m.codigo codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from master m inner join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (m.codint=vn.codigo)"
        If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') ilp_vn"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") ilp_vn"
        End If
        ConsultaSql$ = ConsultaSql$ & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " group by codigo, codexte) formatsalida1) formatosalida1" _
        & " order by porce_fccmv"

        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv

    ElseIf MODO% = 7 Then
        ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, preciolista, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select fecoscal, h.cod_inte from histocostos h inner join preventa p on (h.cod_inte=p.cod_pieza) WHERE p.cod_lista=" & CStr(CodigoLista&) & ") h" _
        & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        
        ' ( (preventa left join (cajabanc inner join movisto) ) left join costos
        ConsultaSql$ = "select * from " _
        & " (select ilp_vn.codigo, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo, preciolista from" _
        & " (select codigo, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from preventa inner join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c inner join movisto m on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (preventa.cod_pieza=vn.codigo)" _
        & " where preventa.cod_lista=" & CStr(CodigoLista&) & ") ilp_vn" _
        & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " order by codigo, mes, anio"
        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv
        
    ElseIf MODO% = 8 Then
            ConsultaSql$ = "ALTER VIEW COSTOS AS " _
        & " (SELECT H.FECOSCAL [fecha costeo], H.COD_INTE codigo, H.CostoStdTotal costo, preciolista, T2.mes, T2.anio FROM HISTOCOSTOS  H" _
        & " Inner Join" _
        & " (SELECT MAX(FECOSCAL) fe, COD_INTE , MONTH(FECOSCAL) mes, YEAR(FECOSCAL) anio FROM (select h.fecoscal, h.cod_inte from histocostos h inner join master m on (h.cod_inte=m.codint) "
        If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') h"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") h"
        End If
        ConsultaSql$ = ConsultaSql$ & " WHERE FECOSCAL >='" & FETEXCOR1$(DesdeFecha%) & "' And FECOSCAL <='" & FETEXCOR1$(HastaFecha%) & "'" _
        & " GROUP BY COD_INTE, MONTH(FECOSCAL), YEAR(FECOSCAL)) T2" _
        & " ON H.COD_INTE = T2.COD_INTE AND H.FECOSCAL = T2.FE)"
        rst_costos.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        Set rst_costos = Nothing
        Text1 = ConsultaSql$

        ' ( (master inner join (cajabanc inner join movisto) ) left join costos
        ' con master se filtra por rango de codigos o por tabla de productos
        ' con el master inner join solo se trabaja con aquellos codigos que tienen movimientos.
        ConsultaSql$ = "select * from " _
        & " (select codigo, codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio, costo, precio_unitario, preciolista from" _
        & " (select ilp_vn.codigo, codexte, ingreso, salida, valunit_ingreso, valunit_salida, ilp_vn.mes, ilp_vn.anio, costo, preciolista from" _
        & " (select m.codint codigo, m.codigo codexte, ingreso, salida, valunit_ingreso, valunit_salida, mes, anio from master m with(nolock) inner join" _
        & " (select m.cod_inte codigo, sum(m.cantingre) ingreso, sum(m.cantsalid) salida, sum(m.cantingre*m.valounit*m.tipocam) valunit_ingreso, sum(m.cantsalid*m.valounit*m.tipocam) valunit_salida, month(fechconta) mes, year(fechconta) anio" _
        & " from cajabanc c with(nolock) inner join movisto m with(nolock) on (c.codicom_lar=m.doseclar)" _
        & " where c.tipomovim='venta' and c.fechconta>='" & FETEXCOR1$(DesdeFecha%) & "' and c.fechconta<='" & FETEXCOR1$(HastaFecha%) & "' and m.codimovi='rt'" _
        & " group by cod_inte, month(fechconta), year(fechconta)) vn" _
        & " on (m.codint=vn.codigo)"
        If CondicionSqlCodigos$ = "" Then
            ConsultaSql$ = ConsultaSql$ & " where m.codigo>='" & CodigoDesde$ & "' and m.codigo<='" & CodigoHasta$ & "') ilp_vn"
        Else
            ConsultaSql$ = ConsultaSql$ & " where " & CondicionSqlCodigos$ & ") ilp_vn"
        End If
        ConsultaSql$ = ConsultaSql$ & " left join costos" _
        & " on (ilp_vn.codigo=costos.codigo and ilp_vn.mes=costos.mes and ilp_vn.anio=costos.anio)) vncmv" _
        & " left join (select cod_lista, cod_pieza, precio_unitario from preventa where cod_lista=" & CStr(CodigoLista&) & ") p on codigo=p.cod_pieza) vncmvpre" _
        & " order by codexte, mes, anio"
        rst_ventasvscmv.Open ConsultaSql$, FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        
        Set VentaNetavsCMV = rst_ventasvscmv
        
        ConsultaSql$ = "select count(*) totalequipos from (select codigo, count(*) vecesporequipo from (" _
        & Mid$(ConsultaSql$, 1, Len(ConsultaSql$) - Len("order by codexte, mes, anio")) _
        & ") agrucodi group by codigo) suma"
        Set rst_totalequipos = READSET(ConsultaSql$)
        
        CantidadRegistros& = XVALO(rst_totalequipos!TotalEquipos)
        On Error Resume Next
        rst_totalequipos.Close
        Set rst_totalequipos = Nothing
        On Error GoTo 0
        
        Text2 = ConsultaSql$
    End If
    

End Function

Private Sub mnuGrupoMediosCobranza_Click()

    If DERACCE%("CFGESVE", "Configuración de Estadísticas de Ventas") < 2 Then
        GoTo 99
    End If
    
    Call TablasAuxiliares
'
    On Error Resume Next
    FLEXCONN.Execute "delete #TPMedCobranza"
    ConsultaSql$ = "create table #TPMedCobranza (codigo int, denominacion varchar(64))"
    FLEXCONN.Execute ConsultaSql$
    On Error GoTo 0

    On Error GoTo ControlError
    Dim Rst As New ADODB.Recordset
    ConsultaSql$ = "select * from #TPMedCobranza"
    Rst.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText
    With Rst
        For i& = 1 To ULTREG&("MECOBRA")
            regi$ = REGLEIDO$("MECOBRA", i&)
            Codigo& = CVI(Mid$(regi$, 1, 2))
            Denominacion$ = TRIM$(Mid$(regi$, 3, 12))

            .AddNew
            !Codigo = Codigo&
            !Denominacion = Denominacion$
            .Update
        Next i&
    End With
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    IDLIST& = ObtenerIdTablaAuxiliar("GRUPOMEDCOB")
    Call ABM_TABLA(IDLIST&, , , "cast(campo1 as int)")
    Exit Sub

ControlError:
    Exit Sub

99  Exit Sub

End Sub

Private Sub mnuLibroIvaDigital_Click()

    Screen.MousePointer = 11
    If VALIMESAN% < 1 Then GoTo 99
    Call GENIVAVEN(0, 1)
    Screen.MousePointer = 1
   
    Call LibroIvaDigitalVentas

99  Screen.MousePointer = 1

End Sub


Private Sub LibroIvaDigitalVentas()
    
    On Error Resume Next
    Call TRALOCAL("LIVAVEN", "")
    On Error GoTo 0
    
    If Err > 0 Then
        Call MENSERR(24, "Imposible Copiar Archivo con Datos del Período")
        Exit Sub
    End If
    
    PERIODO$ = TRIM$(MESAN)
    ' borrar archivo RIV
    NX% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_VENTAS_CBTE_" & PERIODO$ & ".TXT", 2, 266)
    NX1% = FILEOPEN%(LUCOM$ & "LIBRO_IVA_DIGITAL_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT", 2, 62)

    UltimoNumeroComprobante$ = ""
    
    Dim VGD_FechaComprobante$(3)
    Dim VGD_TipoCbte$(3)
        VGD_TipoCbte$(1) = TipoComprobante("FC-0000-00000000-B")
        VGD_TipoCbte$(2) = TipoComprobante("ND-0000-00000000-B")
        VGD_TipoCbte$(3) = TipoComprobante("NC-0000-00000000-B")
    Dim VGD_PuntoVta$(3)
    Dim VGD_NumeroCbtDesde$(3)
    Dim VGD_NumeroCbtHasta$(3)
    VGD_CodigoDocComprador$ = "99"
    VGD_NumeroIdentificacionComprador$ = RIVFormatoNumero$(0, 20)
    VGD_RazonSocial$ = AJUSTI$("VENTA GLOBAL DIARIA", 30)
    Dim VGD_ImporteTotal$(3)
    Dim VGD_ImporteNoGravado$(3)
    VGD_PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
    Dim VGD_ImporteOperacionExenta$(3)
    Dim VGD_ImportePercepcionesNacionales$(3)
    Dim VGD_ImportePercepcionIB$(3)
    VGD_ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
    VGD_ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
    VGD_CodMoneda$ = CodigoMoneda(1)
    VGD_TipoCambio$ = RIVFormatoNumero$(1, 10, 6)
    Dim VGD_CantidadAlicuotas$(3)
    VGD_CodigoOperacion$ = " "
    VGD_OtrosTributos$ = RIVFormatoNumero$(0, 15)
    VGD_FechaVencimiento$ = "00000000"
    Dim VGD_RegistroRIV$(3)
    Dim INDICE%
    
    Dim RegistroRIV_Alicuotas$
    
    Dim FlagError%
    Dim ListaDeErrores$
    
    INDICE% = 1: GoSub InicializarCamposVGD
    INDICE% = 2: GoSub InicializarCamposVGD
    INDICE% = 3: GoSub InicializarCamposVGD
    
    FIN& = ULTREG&("!LIVAVEN")
    If FIN& = 0 Then GoTo 99
    
    Screen.MousePointer = 11
    For i& = 1 To FIN&
        Call TRACE(24, i&, ULTREG&("!LIVAVEN"))
        
        REGISTRO$ = REGLEIDO$("!LIVAVEN", i&)
        NumeroComprobante$ = Mid$(REGISTRO$, 11, 18)
        
        If TRIM$(NumeroComprobante$) = "" Then
            NumeroComprobante$ = UltimoNumeroComprobante$
        End If
        
        If NumeroComprobante$ <> UltimoNumeroComprobante$ Then
            '
            FechaComprobante$ = FECHATEXLAR$(CVI(Mid$(REGISTRO$, 3, 2)))
            FechaComprobante$ = Mid$(FechaComprobante$, 7, 4) & Mid$(FechaComprobante$, 4, 2) & Mid$(FechaComprobante$, 1, 2)
                       
            If UltimoNumeroComprobante$ <> "" And FlagError% = 0 Then
                If TRIM$(RegistroRIV$) <> "" Then
                    Print #NX%, RegistroRIV$
                    RegistroRIV$ = ""
                    RegistroRIV_Alicuotas$ = ""
                End If
            End If
            FlagError% = 0
25
            fechaCbte$ = FETEXCOR1$(CVI(Mid$(REGISTRO$, 3, 2)))
            NumeroCliente& = XVALO(VALOBADA("CAJABANC", "NuClien", "(CODICOM_LAR='" & NumeroComprobante$ & "' OR CODICOM_LAR_EXSIMPLE='" & NumeroComprobante$ & "') AND FechConta='" & fechaCbte$ & "'"))
            If NumeroCliente& = 0 Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "Imposible Procesar el Movimiento. No Se Pudo Encontrar Numero de Cliente para el Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If

            TipoCbte$ = TipoComprobante$(NumeroComprobante$)
            
            PUNTOVTA$ = PuntoVenta$(NumeroComprobante$)
            
            NumeroCbt$ = NumeroCbte$(NumeroComprobante$)
            
            CodigoDocComprador$ = CodigoDocumento$(PIVACLI%(NumeroCliente&))
            
            NumeroIdentificacionComprador$ = TRIM$(Mid$(REGISTRO$, 63, 16))
'            If NumeroIdentificacionComprador$ = "" Or VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
'                NumeroIdentificacionComprador$ = CUITCLI$(NumeroCliente&)
'            End If
            If NumeroIdentificacionComprador$ = "" Then
                'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                FlagError% = 1: GoTo 50
            End If
            PosicionIva% = PIVACLI%(NumeroCliente&)
            If PosicionIva% <> 5 And PosicionIva% <> 0 Then
                If VALICUIT%(NumeroIdentificacionComprador$) = 0 Then
                    'Call MENSERR(24, "Imposible Procesar el Movimiento.\CUIT Inválido para el Comprobante '" & NumeroComprobante$ & "'")
                    ListaDeErrores$ = ListaDeErrores$ & "CUIT Inválido en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
                End If
            End If
            NumeroIdentificacionComprador$ = SOLONUM$(NumeroIdentificacionComprador$)
            NumeroIdentificacionComprador$ = AJUSTD$(NumeroIdentificacionComprador$, 20)
            NumeroIdentificacionComprador$ = REPLACAR(NumeroIdentificacionComprador$, " ", "0")
            
            razonSocial$ = AJUSTI$(rasocli$(NumeroCliente&), 30)
            If PIVACLI%(NumeroCliente&) = 0 And (TRIM$(razonSocial$) = "" Or InStr(1, razonSocial$, "CONSUMIDOR") > 0) Then
                razonSocial$ = AJUSTI$("CONSUMIDOR FINAL", 30)
            End If
            
            ImporteTotal$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 91, 8)), 15)
            
            ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNoGravado$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
            End If
            
            PercepcionNoCategorizados$ = RIVFormatoNumero$(0, 15)
            
            ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
            
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                ImporteOperacionExenta$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
            
            ImportePercepcionesNacionales$ = RIVFormatoNumero$(0, 15)
            
            ImportePercepcionIB$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 193, 8)), 15)
            
            ImportePercepcionMunicipales$ = RIVFormatoNumero$(0, 15)
            
            ImporteImpuestosInternos$ = RIVFormatoNumero$(0, 15)
                
                MONE% = XVALO(VALOBADA("CAJABANC", "Mone_Emis", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            CodMoneda$ = CodigoMoneda(MONE%)
            If TRIM$(CodMoneda$) = "000" Then CodMoneda$ = "PES"

                TCAMBIO# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & NumeroComprobante$ & "' AND FechConta='" & fechaCbte$ & "'"))
            TIPOCAMBIO$ = RIVFormatoNumero$(TCAMBIO#, 10, 6)
            If CodMoneda$ = "PES" And (XVALO(TIPOCAMBIO$) / 1000000) > 1 Then
                    ListaDeErrores$ = ListaDeErrores$ & "Error en Tipo de Cambio en el Cliente '" & CStr(NumeroCliente&) & "' en Comprobante '" & NumeroComprobante$ & "'" & vbCrLf
                    FlagError% = 1: GoTo 50
            End If
            If XVALO(TIPOCAMBIO$) = 0 And CodMoneda$ = "PES" Then TIPOCAMBIO$ = RIVFormatoNumero(1, 10, 6)
                
                
            ' Debe existir por lo menos una alicuota aunque sea no gravado o exento. No permite 0 (cero)
                CantAlicuotas% = 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
            
            If PIVACLI%(NumeroCliente&) = 7 Then
                CodigoOperacion$ = "Z"
            ElseIf PIVACLI%(NumeroCliente&) = 5 Then
                CodigoOperacion$ = "X"
            ElseIf Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                CodigoOperacion$ = "N"
            ElseIf PIVACLI%(NumeroCliente&) = 4 Then
                CodigoOperacion$ = "E"
            Else
                CodigoOperacion$ = " "
            End If
            
            OtrosTributos$ = RIVFormatoNumero$(0, 15)
            
            ' Si es una Exportacion o Consumidor Final, es la misma que la fecha del comprobante.
            ' Caso contrario la calcula
            FechaVencimientoCbte$ = FechaComprobante$
            If PIVACLI%(NumeroCliente&) <> 5 And PIVACLI%(NumeroCliente&) <> 7 And PIVACLI%(NumeroCliente&) <> 0 Then
                FechaVencimientoCbte$ = FechaComprobante$ ' calcular!!!
            ElseIf PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                FechaVencimientoCbte$ = "00000000"
            End If
            
            ' Arma el registro para el regimen informativo
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                            & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                            & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                            & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
              
            UltimoNumeroComprobante$ = NumeroComprobante$
            
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            If InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                ImporteIva$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 145, 8)), 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                
                RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            Else
                ConIva# = XVALO(VALOBADA("CAJABANC", "Sum(ValAbsComp)", "TipoMovim='IVA' and CODICOM_LAR='" & NumeroComprobante$ & "' AND NuClien='" & CStr(NumeroCliente&) & "'"))
                If (ConIva# < 0.01 Or (XVALO(ImportePercepcionIB$) / 100) > 0.005) And (XVALO(ImporteOperacionExenta$) / 100) < 0.01 Then
                    ImporteNeto$ = ImporteTotal$
                    If (XVALO(ImportePercepcionIB$) / 100) > 0.005 Then ImporteNeto$ = RIVFormatoNumero(CVD(Mid$(REGISTRO$, 137, 8)), 15)
                    
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteNoGravado$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteNoGravado$, 124, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                    
                ElseIf (XVALO(ImporteOperacionExenta$) / 100) > 0.005 Then
                    
                    ImporteNeto$ = ImporteTotal$
                    ImporteIva$ = RIVFormatoNumero(0, 15)
                    ALICUOTA$ = "0003"
                    ImporteOperacionExenta$ = RIVFormatoNumero$(0, 15)
                    Call REPLA(RegistroRIV$, ImporteOperacionExenta$, 154, 15)
                    
                    RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                    Print #NX1%, RegistroRIV_Alicuotas$
                End If
            End If
        Else
            If Mid$(REGISTRO$, 129, 8) = "NO GRAV." Then
                ImporteNGAnterior = XVALO(Mid$(RegistroRIV$, 124, 15)) / 100
                ImporteNGActual = CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8))
                If Not (ImporteNGActual < 0 And Mid$(REGISTRO$, 11, 2) = "FC") Then
                    ImporteNoGravado$ = RIVFormatoNumero(ImporteNGAnterior + CVD(Mid$(REGISTRO$, 137, 8)) + CVD(Mid$(REGISTRO$, 153, 8)) + CVD(Mid$(REGISTRO$, 161, 8)), 15)
                End If
            End If
           
            If PIVACLI%(NumeroCliente&) = 5 Or PIVACLI%(NumeroCliente&) = 7 Then
                IMPO# = XVALO(Mid$(RegistroRIV$, 154, 15)) / 100
                ImporteOperacionExenta$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 185, 8)), 15)
                ImporteNoGravado$ = RIVFormatoNumero$(0, 15) ' O es No Gravado o es Exento
            End If
                          
            IMPO# = XVALO(Mid$(RegistroRIV$, 184, 15)) / 100
            ImportePercepcionIB$ = RIVFormatoNumero(IMPO# + CVD(Mid$(REGISTRO$, 193, 8)))
                               
                CantAlicuotas% = XVALO(CantidadAlicuotas$)
                If Mid$(REGISTRO$, 129, 8) <> "NO GRAV." Then CantAlicuotas% = CantAlicuotas% + 1
            CantidadAlicuotas$ = CStr(CantAlicuotas%)
               
            RegistroRIV$ = FechaComprobante$ & TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & NumeroCbt$ & CodigoDocComprador$ & NumeroIdentificacionComprador$ & razonSocial$ _
                           & ImporteTotal$ & ImporteNoGravado$ & PercepcionNoCategorizados$ & ImporteOperacionExenta$ & ImportePercepcionesNacionales$ _
                           & ImportePercepcionIB$ & ImportePercepcionMunicipales$ & ImporteImpuestosInternos$ & CodMoneda$ & TIPOCAMBIO$ & CantidadAlicuotas$ _
                           & CodigoOperacion$ & OtrosTributos$ & FechaVencimientoCbte$
               
            ImporteNeto0# = CVD(Mid$(REGISTRO$, 137, 8))
            ValorIva$ = TRIM$(REPLACAR$(Mid$(REGISTRO$, 129, 8), "%", ""))
            ImporteIvaActual# = Abs(CVD(Mid$(REGISTRO$, 145, 8)))
            If Not (ImporteNeto0# < 0 And Mid$(REGISTRO$, 11, 2) = "FC") And ImporteIvaActual# > 0.005 Then ' InStr(1, UCase(ValorIva$), "NO GRAV.") = 0 Then
                ImporteNeto$ = RIVFormatoNumero(ImporteNeto0#, 15)
                ALICUOTA$ = CodigoAlicuota(ValorIva$)
                ImporteIva$ = RIVFormatoNumero(ImporteIvaActual#, 15)
                   
                RegistroRIV_Alicuotas$ = TipoCbte$ & PUNTOVTA$ & NumeroCbt$ & ImporteNeto$ & ALICUOTA$ & ImporteIva$
                Print #NX1%, RegistroRIV_Alicuotas$
            End If
        End If
50  Next i&
    
    If NumeroComprobante$ <> "" Then Print #NX%, RegistroRIV$
    
    Close #NX%
    Close #NX1%
        
    Screen.MousePointer = 1
    
'    Dim XX1 As String * 266
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_CBTE_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 266)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 9, 3) = "003" Or Mid$(XX$, 9, 3) = "008" Or Mid$(XX$, 9, 3) = "021" Then SIGNO% = -1
'
'        IMPO1# = IMPO1# + SIGNO% * XVALO(Mid$(XX$, 109, 15)) / 100
'        impo2# = impo2# + SIGNO% * XVALO(Mid$(XX$, 124, 15)) / 100
'        impo3# = impo3# + SIGNO% * XVALO(Mid$(XX$, 139, 15)) / 100
'        impo4# = impo4# + SIGNO% * XVALO(Mid$(XX$, 154, 15)) / 100
'        impo5# = impo5# + SIGNO% * XVALO(Mid$(XX$, 169, 15)) / 100
'        impo6# = impo6# + SIGNO% * XVALO(Mid$(XX$, 184, 15)) / 100
'        impo7# = impo7# + SIGNO% * XVALO(Mid$(XX$, 199, 15)) / 100
'        impo8# = impo8# + SIGNO% * XVALO(Mid$(XX$, 214, 15)) / 100
'        impo9# = impo9# + SIGNO% * XVALO(Mid$(XX$, 244, 15)) / 100
'20  Loop
'
'    MsgBox ("Total: " & TRIM$(FORMATNUM$(IMPO1#, "F15.2")) & vbCrLf _
'        & "No gravado: " & TRIM$(FORMATNUM$(impo2#, "F15.2")) & vbCrLf _
'        & "Percep. No Cat: " & TRIM$(FORMATNUM$(impo3#, "F15.2")) & vbCrLf _
'        & "Op.Exentas: " & TRIM$(FORMATNUM$(impo4#, "F15.2")) & vbCrLf _
'        & "Imp.Nacionales: " & TRIM$(FORMATNUM$(impo5#, "F15.2")) & vbCrLf _
'        & "Imp.IIBB: " & TRIM$(FORMATNUM$(impo6#, "F15.2")) & vbCrLf _
'        & "Imp.Municipales: " & TRIM$(FORMATNUM$(impo7#, "F15.2")) & vbCrLf _
'        & "Imp.Internos: " & TRIM$(FORMATNUM$(impo8#, "F15.2")) & vbCrLf _
'        & "Otro Atributos: " & TRIM$(FORMATNUM$(impo9#, "F15.2")) & vbCrLf)
'
'    Close #NARCHI%
'
'    ARCHIIBB$ = "REGINFO_CV_VENTAS_ALICUOTAS_" & PERIODO$ & ".TXT"
'    NARCHI% = FILEOPEN%(LUCOM$ + ARCHIIBB$, 1, 62)
'    Do While Not EOF(NARCHI%)
'        Line Input #NARCHI%, XX$
'        SIGNO% = 1
'        If Mid$(XX$, 1, 3) = "003" Or Mid$(XX$, 1, 3) = "008" Or Mid$(XX$, 1, 3) = "021" Then SIGNO% = -1
'        IMPO20# = IMPO20# + SIGNO% * XVALO(Mid$(XX$, 29, 15)) / 100
'        impo21# = impo21# + SIGNO% * XVALO(Mid$(XX$, 48, 15)) / 100
'    Loop
'
'    MsgBox ("Neto: " & TRIM$(FORMATNUM$(IMPO20#, "F15.2")) & vbCrLf _
'        & "Iva: " & TRIM$(FORMATNUM$(impo21#, "F15.2")))
'
'    Close #NARCHI%
    
    If TRIM$(ListaDeErrores$) <> "" Then
        Call COMUNI("Se Han Detectado Errores en la Generación.\Pulse Aceptar para Ver el Detalle")
        MSJESREGINFO.Text1 = ListaDeErrores$
        MSJESREGINFO.Caption = "Régimen Informativo de Ventas"
        Call MSJESREGINFO.Command1_Click
        MSJESREGINFO.Show 1
        Call MENSERR(24, "Tener Presente que los Datos Generados Son Incompletos.\Debe Corregir los Errores y Volver a Generar")
    Else
        Call COMUNI("Libro Iva Digital de Ventas: Generación Ok")
    End If
    
99  Exit Sub
    
    
InicializarCamposVGD:
    
    VGD_FechaComprobante$(INDICE%) = "00000000"
    VGD_PuntoVta$(INDICE%) = ""
    VGD_NumeroCbtDesde$(INDICE%) = "0"
    VGD_NumeroCbtHasta$(INDICE%) = "0"
    VGD_ImporteTotal$(INDICE%) = "0"
    VGD_ImporteNoGravado$(INDICE%) = "0"
    VGD_ImporteOperacionExenta$(INDICE%) = "0"
    VGD_ImportePercepcionesNacionales$(INDICE%) = "0"
    VGD_ImportePercepcionIB$(INDICE%) = "0"
    VGD_CantidadAlicuotas$(INDICE%) = "0"
    VGD_RegistroRIV$(INDICE%) = ""

    Return
    
End Sub
Private Sub mnuLiquidaComis_Click()
    Call CONECRUN("LCOMIS07", "Liquidación de Comisiones")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub IIBBPercepRetenCABAExpo_Click()

         ' *** FACTURACION Y ORDENES DE PAGO

         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99

         DESD$ = FETEXCOR1$(Des%)
         HAST$ = FETEXCOR1$(Has%)
         '
         CommonDialog1.CancelError = True
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         '
         ' *** CREA ARCHIVO DE TRANSFERENCIA
         NX1% = FILEOPEN%(ASALID$, 2, 215)
         '
         Screen.MousePointer = 11
         '
         ' *** OBTIENE LAS PERCEPCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         '
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC' AND IHABECOMP > 0.005 AND NUCLIEN > 0"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIbCaba As ADODB.Recordset

         Set RstIbCaba = READSET(ConsultaSql$)
         With RstIbCaba
            On Error GoTo ControlErrorVentas
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)

                REBLA$ = Space$(215)
                NroCliente% = XVALO(!NUCLIEN)
                DocRetenido$ = SAFETEXT$(!codicom_lar)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NumeComp)
                ImporteRetencion# = XVALO(!ihabecomp)

                ' *** DATOS DEL COMPROBANTE RETENIDO
                condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='IVA'"
                CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='VENTA'"

                DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='FVEN'"))
                DocRet_Total# = DocRet_Importe#
                DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", condicionIva$))
                DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CondicionNeto$))
                DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva# - ImporteRetencion#


                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "2", 1, 1)

                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)

                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)

                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 18, 1)
                    If DocRet_Tipo$ = "09" Then DocRet_Letra$ = " "
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)

                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 4, 4) & Mid$(DocRetenido$, 9, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)

                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = FechaRetencion%
                    Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)

                ' MONTO DEL COMPROBANTE
                    If DocRet_Letra$ <> "B" Then
                        DocRet_Importe# = DocRet_Importe# - ImporteRetencion#
                    Else
                        DocRet_Importe# = DocRet_Neto# + DocRet_OtConceptos#
                    End If
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = "0"
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)

                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%(NroCliente%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)

                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$(NroCliente%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)

                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
'                    NroIB$ = TRIM$(NIBRCLI$(NroCliente%))
'                    NroIB$ = REPLACAR$(NroIB$, "-", "")
'                    CategoriaIB% = 2
'                    If NroIB$ = CUIT$ Or NroIB$ = "" Then CategoriaIB% = 5
CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
NroIB$ = Cuit$
                Call REPLA(REBLA$, NroIB$, 89, 11)

                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(SituacionIva%), 100, 1)

                ' RAZON SOCIAL DEL RETENIDO
                    razonSocial$ = TextoEnFormatoUTF8(rasocli$(NroCliente%))
                Call REPLA(REBLA$, razonSocial$, 101, 30)

                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Or XVALO(DocRet_OtCon$) = 0 Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)

                ' IMPORTE IVA
                    If DocRet_Letra$ = "B" Then DocRet_Iva# = 0
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    If DocRet_ImporteIva$ = "" Or XVALO(DocRet_ImporteIva$) = 0 Then DocRet_ImporteIva$ = "0.00"
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)

                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

                ' ALICUOTA
On Error GoTo ControlErrorPercepciones
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)
On Error GoTo 0

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)

                TotalPercepcion# = TotalPercepcion# + ImporteRetencion#

                Print #NX1%, REBLA$
25              .MoveNext
            Loop
            On Error GoTo 0
         End With
         'Close #NX1%
         On Error Resume Next
         RstIbCaba.Close
         Set RstIbCaba = Nothing
         On Error GoTo 0

         '
         ' *** OBTIENE LAS RETENCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         '
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='RTIB' AND OPCIMOVIM='CABA' AND IHABECOMP > 0.005 AND NUPROVE > 0"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIBRETCABA As ADODB.Recordset

         Set RstIBRETCABA = READSET(ConsultaSql$)
         With RstIBRETCABA
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)

                REBLA$ = Space$(215)
                NroProveedor% = XVALO(!NUPROVE)
                DocRetenido$ = SAFETEXT$(!notas)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NroDeRetencion)
                ImporteRetencion# = XVALO(!ihabecomp)

                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "1", 1, 1)

                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)

                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                    If Mid$(DocRetenido$, 1, 2) = "ND" Then DocRet_Tipo$ = "02"
                    If Mid$(DocRetenido$, 1, 2) = "RC" Then DocRet_Tipo$ = "07"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)

                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 4, 1)
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)

                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 6, 4) & Mid$(DocRetenido$, 11, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)

                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    If DocRet_Fecha% < FECHANUM("01/11/16") Then DocRet_Fecha% = FechaRetencion%
                Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)

                ' MONTO DEL COMPROBANTE
                    DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"))
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = TRIM$(FORMATNUM$(NroRetencion&, "I10"))
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)

                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%((-1) * NroProveedor%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)

                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$((-1) * NroProveedor%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)

                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
'                    NroIB$ = TRIM$(NIBRCLI$((-1) * NroProveedor%))
'                    NroIB$ = REPLACAR$(NroIB$, "-", "")
'                    CategoriaIB% = 2                        ' CONVENIO MULTILATERAL
'                    If CUIT$ = NroIB$ Then CategoriaIB% = 5 ' REGIMEN SIMPLIFICADO
NroIB$ = Cuit$
CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
                Call REPLA(REBLA$, NroIB$, 89, 11)

                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(SituacionIva%), 100, 1)

                ' RAZON SOCIAL DEL RETENIDO
                    RAZONSOCIAL0$ = rasocli$((-1) * NroProveedor%)
                    razonSocial$ = TextoEnFormatoUTF8(RAZONSOCIAL0$)
                Call REPLA(REBLA$, razonSocial$, 101, 30)

                ' *** DATOS DEL COMPROBANTE RETENIDO
                    CondicionTotal$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='FCOM'"
                    condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='IVA' AND VARIMOVIM<>'EXENTO' AND VARIMOVIM<>'NO GRAV' AND VARIMOVIM<>'ADUANA' AND VARIMOVIM<>'PER-IVA'"
                    CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUPROVE=" & NroProveedor% & " AND TIPOMOVIM='COMPRA'"

                    DocRet_Total# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CondicionTotal$))
                    DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", condicionIva$))
                    DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CondicionNeto$))
                    DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva#

                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)

                ' IMPORTE IVA
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    If DocRet_ImporteIva$ = "" Then DocRet_ImporteIva$ = "0.00"
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)

                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    If DocRet_ImporteNeto$ = "" Then DocRet_ImporteNeto$ = "0.00"
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

On Error GoTo ControlErrorRetenciones
                ' ALICUOTA
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)
On Error GoTo 0

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)

                TotalRetencion# = TotalRetencion# + ImporteRetencion#

                Print #NX1%, REBLA$
50              .MoveNext
            Loop
         End With
         Close #NX1%
         On Error Resume Next
         RstIBRETCABA.Close
         Set RstIBRETCABA = Nothing
         On Error GoTo 0

         'MsgBox "Total Retencion: " & CStr(TotalRetencion#) & vbCrLf & "Total Percepcion: " & CStr(TotalPercepcion#)
         'MsgBox "Archivo Transferencia Generado"
         If EXISTE%(ASALID$) > 0 Then

            ' *** CARGA UNA MATRIZ CON LAS PERCEPCIONES Y RETENCIONES PARA LUEGO ORDENAR
            NARCHI2% = FILEOPEN%(ASALID$, 1, 215) 'ABRE EL ARCHIVO DE LECTURA
            CantidadRegistros& = LOF(NARCHI2%)
            Dim MatrizOrdenaRegistros() As String
            ReDim MatrizOrdenaRegistros(CantidadRegistros&, 2)

            FilaMatriz& = 0
            Do While Not EOF(NARCHI2%) ' RECORRE
                Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
                FilaMatriz& = FilaMatriz& + 1
                    FechaRetPer$ = Mid$(XX$, 5, 10)
                MatrizOrdenaRegistros(FilaMatriz&, 1) = Mid$(FechaRetPer$, 7, 4) & Mid$(FechaRetPer$, 4, 2) & Mid$(FechaRetPer$, 1, 2)
                MatrizOrdenaRegistros(FilaMatriz&, 2) = XX$
            Loop
            Close #NARCHI2%
            'MsgBox "Matriz generada"

            ' *** ORDENAMIENTO BURBUJA POR FECHA - ES REQUERIDO POR EL IMPORTADOR DEL GOB DE LA CIUDAD
            Dim iMin    As Long
            Dim iMax    As Long
            Dim Vectemp1 As String                   ' -- variable temporal
            Dim Vectemp2 As String                   ' -- variable temporal
            Dim Pos     As Long
            Dim i       As Long

            iMin = 1
            iMax = FilaMatriz&

            While iMax > iMin
                Pos = iMin
                For i = iMin To iMax - 1
                    If MatrizOrdenaRegistros(i, 1) > MatrizOrdenaRegistros(i + 1, 1) Then
                        Vectemp1 = MatrizOrdenaRegistros(i + 1, 1)
                        Vectemp2 = MatrizOrdenaRegistros(i + 1, 2)

                        MatrizOrdenaRegistros(i + 1, 1) = MatrizOrdenaRegistros(i, 1)
                        MatrizOrdenaRegistros(i + 1, 2) = MatrizOrdenaRegistros(i, 2)

                        MatrizOrdenaRegistros(i, 1) = Vectemp1
                        MatrizOrdenaRegistros(i, 2) = Vectemp2

                        Pos = i
                    End If
                Next i
                iMax = Pos
            Wend
            'MsgBox "Matriz ordenada"

            ' *** REGENERA EL ARCHIVO CON PERCEPCIONES Y RETENCIONES, PERO ESTA VEZ ORDENADO POR FECHA
            On Error Resume Next
            Kill ASALID$
            If Err = 0 Then
                NX1% = FILEOPEN%(ASALID$, 2, 215)
                For i = 1 To FilaMatriz
                    REBLA$ = MatrizOrdenaRegistros(i, 2)
                    Print #NX1%, REBLA$
                Next i
                Close #NX1%
                'MsgBox "Archivo ordenado"
            Else
                Call MENSERR(24, "Error al Ordenar Archivo de Retenciones y Percepciones IIBB.\ \No Considerar el Mismo para Importar en ARCIBA.")
            End If
            On Error GoTo 0

         End If

         Screen.MousePointer = 1

         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "\Generado con " + TRIM$(Str$(CAREGI&)) + " Registros.")

99       Exit Sub


ControlErrorVentas:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroCliente%) & "\Comprobante: " & DocRetenido$)
    Screen.MousePointer = 1
    GoTo 99

ControlErrorPercepciones:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroCliente%) & "\Comprobante: " & DocRetenido$ & ".\ \Movimiento No Incluido")
    Screen.MousePointer = 1
    GoTo 25

ControlErrorRetenciones:

    Call MENSERR(24, "Error en Percepciones IIBB CABA\ \Cliente: " & rasocli$(NroProveedor% * -1) & "\Comprobante: " & DocRetenido$ & ".\ \Movimiento No Incluido")
    Screen.MousePointer = 1
    GoTo 50

End Sub

Private Sub IIBBpercepCABANCExpo_Click()

         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99

         DESD$ = FETEXCOR1$(Des%)
         HAST$ = FETEXCOR1$(Has%)
         CantRegistrosRechazados& = 0
         '
         CommonDialog1.CancelError = True
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         '
         ' *** CREA ARCHIVO DE TRANSFERENCIA
         NX1% = FILEOPEN%(ASALID$, 2, 215)
         '
         Screen.MousePointer = 11
         '
         ' *** OBTIENE LAS RETENCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC' AND IDEBECOMP > 0.005 AND NUCLIEN > 0"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIbCaba As ADODB.Recordset

         Set RstIbCaba = READSET(ConsultaSql$)
         With RstIbCaba
            Do While Not .EOF

                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)

                REBLA$ = Space$(59)
                NroCliente% = XVALO(!NUCLIEN)
                DocRetenido$ = SAFETEXT$(!codicom_lar)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NumeComp)
                ImporteRetencion# = XVALO(!idebecomp)
                BaseCAlculo# = XVALO(!BASEIMPO)

                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "2", 1, 1)

               ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 4, 4) & Mid$(DocRetenido$, 9, 8)
                    While Len(DocRet_Numero$) < 12: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 2, 12)

                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 14, 10)

                ' MONTO DEL COMPROBANTE
                    ComprobanteImporte# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND TIPOMOVIM='VENTA'"))
                    DocRet_Impo$ = TRIM$(FORMATNUM$(ComprobanteImporte#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 24, 16)

                ' CERTIFICADO
                Call REPLA(REBLA$, Space(16), 40, 16)

                DocOrigenRetencion$ = TRIM$(VALOBADA("APLICOBRA", "CODOCOR_LAR", "CODOCAP_LAR='" & DocRetenido$ & "'"))
                If DocOrigenRetencion$ = "" Then
                    ' BUSCA EL DOCUMENTO ORIGEN POR IMPORTE DE LA PERCEPCION EN EL MISMO PERIODO
                    CondicionSql$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & FETEXCOR1$(FechaRetencion%) & "'" & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC' AND IHABECOMP > 0.005 AND NUCLIEN=" & CStr(NroCliente%)
                    CondicionSql$ = CondicionSql$ & " AND ABS(IHABECOMP-" & FORMATNUM(ImporteRetencion#, "F12.2") & ")< 0.005"
                    DocOrigenRetencion$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", CondicionSql$))
                    If DocOrigenRetencion$ = "" Then
                        Call COMUNI("Nota de Crédito '" & DocRetenido$ & "' No Incluida en la Exportación.\ \No se Ha Encontrado el Documento Origen (Factura)'")
                        CantRegistrosRechazados& = CantRegistrosRechazados& + 1
                        GoTo 50
                    End If
                End If
                FechaComprobante% = CVINT(VALOBADA("CAJABANC", "FECHCONTA", "CODICOM_LAR='" & DocOrigenRetencion$ & "' AND NUCLIEN=" & NroCliente%))
                ComprobanteImporte# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", "CODICOM_LAR='" & DocOrigenRetencion$ & "' AND TIPOMOVIM='VENTA' AND NUCLIEN=" & NroCliente%))
                
                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "01"
                Call REPLA(REBLA$, DocRet_Tipo$, 56, 2)

                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocOrigenRetencion$, 18, 1)
                Call REPLA(REBLA$, DocRet_Letra$, 58, 1)

               ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocOrigenRetencion$, 4, 4) & Mid$(DocOrigenRetencion$, 9, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 59, 16)
                
                ' CUIT
                Call REPLA(REBLA$, REPLACAR(CuitCli(NroCliente%), "-", ""), 75, 11)

                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 86, 3)

                ' FECHA DEL COMPROBANTE
                Call REPLA(REBLA$, FECHATEXLAR$(FechaComprobante%), 89, 10)
                ' IMPORTE PERCEPCION
                    DocRet_Impo$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 99, 16)
                ' ALICUOTA
                    ALICUOTA# = 0
                    If (BaseCAlculo#) > 0.005 Then
                        ALICUOTA# = (ImporteRetencion# / BaseCAlculo#) * 100
                    End If
                    DocRet_Ali$ = TRIM$(FORMATNUM$(ALICUOTA#, "F12.2"))
                    While Len(DocRet_Ali$) < 5: DocRet_Ali$ = "0" & DocRet_Ali$: Wend
                    DocRet_Ali$ = REPLACAR$(DocRet_Ali$, ".", ",")
                Call REPLA(REBLA$, DocRet_Ali$, 115, 5)

                Print #NX1%, REBLA$
50              .MoveNext
            Loop
         End With
         Close #NX1%
         On Error Resume Next
         RstIbCaba.Close
         Set RstIbCaba = Nothing
         On Error GoTo 0

         Screen.MousePointer = 1

         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "\Generado con " + TRIM$(Str$(CAREGI& - CantRegistrosRechazados&)) + " Registros.")

99       Exit Sub

End Sub

Private Sub IIBBpercepCABANNCExpo_Click()

         ' *** NOTA DE CREDITO
         
         Call DESHASFECHA(Des%, Has%, PERIO$)
         If Des% < 1 Then GoTo 99
         
         DESD$ = FETEXCOR1$(Des%)
         HAST$ = FETEXCOR1$(Has%)
         '
         CommonDialog1.CancelError = True
         CommonDialog1.DialogTitle = "Archivo Origen de Datos"
         CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
         CommonDialog1.FilterIndex = 1
         On Error GoTo 99
         CommonDialog1.ShowOpen
         On Error GoTo 0
         NOMBASALID$ = UCase$(CommonDialog1.FileName)
         ASALID$ = TRIM$(NOMBASALID$)
         If ASALID$ = "" Then GoTo 99
         '
         ' *** CREA ARCHIVO DE TRANSFERENCIA
         NX1% = FILEOPEN%(ASALID$, 2, 215)
         '
         Screen.MousePointer = 11
         '
         ' *** OBTIENE LAS RETENCIONES DEL PERIODO ELEGIDO Y COMPLETA EL ARCHIVO DE TRANSFERENCIA
         CONDICION$ = "FECHCONTA >= '" & DESD$ & "' AND FECHCONTA <= '" & HAST$ & "'" & " AND TIPOMOVIM='IBRU' AND OPCIMOVIM='IB-C' AND VARIMOVIM='PERC' AND SUBSTRING(CODICOM_LAR,1,2)= 'NC' IHABECOMP > 0.005"
         FIN& = CantRegistros&("CAJABANC", CONDICION$)
         ConsultaSql$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDICION$ & " ORDER BY FECHCONTA"
         Dim RstIbCaba As ADODB.Recordset
         
         Set RstIbCaba = READSET(ConsultaSql$)
         With RstIbCaba
            Do While Not .EOF
                CAREGI& = CAREGI& + 1
                Call TRACE(24, CAREGI&, FIN&)
                
                REBLA$ = Space$(215)
                NroCliente% = XVALO(!NUCLIEN)
                DocRetenido$ = SAFETEXT$(!codicom_lar)
                FechaRetencion% = CVINT(!fechconta)
                NroRetencion& = XVALO(!NumeComp)
                ImporteRetencion# = XVALO(!idebecomp)

                ' TIPO DE OPERACIONES
                Call REPLA(REBLA$, "2", 1, 1)
                
                ' CODIGO DE NORMA
                Call REPLA(REBLA$, "029", 2, 3)
                
                ' FECHA RETENCION
                Call REPLA(REBLA$, FECHATEXLAR$(FechaRetencion%), 5, 10)

                ' TIPO DE COMPROBANTE
                    DocRet_Tipo$ = "09"
                    If Mid$(DocRetenido$, 1, 2) = "FC" Then DocRet_Tipo$ = "01"
                    If Mid$(DocRetenido$, 1, 2) = "ND" Then DocRet_Tipo$ = "02"
                Call REPLA(REBLA$, DocRet_Tipo$, 15, 2)
                
                ' LETRA DE COMPROBANTE
                    DocRet_Letra$ = Mid$(DocRetenido$, 4, 1)
                Call REPLA(REBLA$, DocRet_Letra$, 17, 1)
                
                ' NRO DE COMPROBANTE
                    DocRet_Numero$ = Mid$(DocRetenido$, 11, 8)
                    While Len(DocRet_Numero$) < 16: DocRet_Numero$ = "0" & DocRet_Numero$: Wend
                Call REPLA(REBLA$, DocRet_Numero$, 18, 16)
                
                ' FECHA DEL COMPROBANTE
                    DocRet_Fecha% = FechaRetencion%
                    Call REPLA(REBLA$, FECHATEXLAR$(DocRet_Fecha%), 34, 10)
                
                ' MONTO DEL COMPROBANTE
                    DocRet_Importe# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='FVEN'"))
                    DocRet_Impo$ = TRIM$(FORMATNUM$(DocRet_Importe#, "F12.2"))
                    While Len(DocRet_Impo$) < 16: DocRet_Impo$ = "0" & DocRet_Impo$: Wend
                    DocRet_Impo$ = REPLACAR$(DocRet_Impo$, ".", ",")
                Call REPLA(REBLA$, DocRet_Impo$, 44, 16)

                ' NRO DE CERTIFICADO
                    NroReten$ = TRIM$(FORMATNUM$(NroRetencion&, "I10"))
                    While Len(NroReten$) < 16: NroReten$ = "0" & NroReten$: Wend
                Call REPLA(REBLA$, NroReten$, 60, 16)
                
                ' TIPO DE DOCUMENTO (CUIT, CUIL)
                    PosicionIva% = PIVACLI%(NroCliente%)
                    PosIva$ = "3": If PosicionIva% = 0 Then PosIva$ = "2"
                Call REPLA(REBLA$, PosIva$, 76, 1)
                
                ' NUMERO DE DOCUMENTO DEL RETENIDO
                    Cuit$ = TRIM$(CuitCli$(NroCliente%))
                    Cuit$ = REPLACAR$(Cuit$, "-", "")
                Call REPLA(REBLA$, Cuit$, 77, 11)
                
                ' SITUACION IIBB DEL RETENIDO (CATEGORIA)
                    NroIB$ = TRIM$(NIBRCLI$(NroCliente%))
                    NroIB$ = REPLACAR$(NroIB$, "-", "")
                    CategoriaIB% = 2
                    If NroIB$ = Cuit$ Then CategoriaIB% = 5
                Call REPLA(REBLA$, CStr(CategoriaIB%), 88, 1)

                ' NUMERO DE INSCRIPCION
                Call REPLA(REBLA$, NroIB$, 89, 11)
                
                ' SITUACION FRENTE AL IVA
                    SituacionIva% = 1 ' RI
                    If PosicionIva% = 6 Then SituacionIva% = 4
                    If PosicionIva% = 4 Then SituacionIva% = 3
                Call REPLA(REBLA$, CStr(PosicionIva%), 100, 1)
                
                ' RAZON SOCIAL DEL RETENIDO
                    razonSocial$ = rasocli$(NroCliente%)
                Call REPLA(REBLA$, razonSocial$, 101, 30)
               
                ' *** DATOS DEL COMPROBANTE RETENIDO
                    condicionIva$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='IVA'"
                    CondicionNeto$ = "CODICOM_LAR='" & DocRetenido$ & "' AND NUCLIEN=" & NroCliente% & " AND TIPOMOVIM='VENTA'"
                    
                    DocRet_Total# = DocRet_Importe#
                    DocRet_Iva# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", condicionIva$))
                    DocRet_Neto# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CondicionNeto$))
                    DocRet_OtConceptos# = DocRet_Total# - DocRet_Neto# - DocRet_Iva# - ImporteRetencion#
                    
                ' IMPORTE OTROS CONCEPTOS
                    DocRet_OtCon$ = TRIM$(FORMATNUM$(DocRet_OtConceptos#, "F12.2"))
                    If DocRet_OtCon$ = "" Then DocRet_OtCon$ = "0.00"
                    While Len(DocRet_OtCon$) < 16: DocRet_OtCon$ = "0" & DocRet_OtCon$: Wend
                    DocRet_OtCon$ = REPLACAR$(DocRet_OtCon$, ".", ",")
                Call REPLA(REBLA$, DocRet_OtCon$, 131, 16)
                
                ' IMPORTE IVA
                    DocRet_ImporteIva$ = TRIM$(FORMATNUM$(DocRet_Iva#, "F12.2"))
                    While Len(DocRet_ImporteIva$) < 16: DocRet_ImporteIva$ = "0" & DocRet_ImporteIva$: Wend
                    DocRet_ImporteIva$ = REPLACAR$(DocRet_ImporteIva$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteIva$, 147, 16)
                
                ' IMPORTE SUJETO DE RETENCION (NETO)
                    DocRet_ImporteNeto$ = TRIM$(FORMATNUM$(DocRet_Neto#, "F12.2"))
                    While Len(DocRet_ImporteNeto$) < 16: DocRet_ImporteNeto$ = "0" & DocRet_ImporteNeto$: Wend
                    DocRet_ImporteNeto$ = REPLACAR$(DocRet_ImporteNeto$, ".", ",")
                Call REPLA(REBLA$, DocRet_ImporteNeto$, 163, 16)

                ' ALICUOTA
                    ALICUOTA# = (ImporteRetencion# / DocRet_Neto#) * 100
                    Alicuo$ = TRIM$(FORMATNUM$(ALICUOTA#, "F5.2"))
                    While Len(Alicuo$) < 5: Alicuo$ = "0" & Alicuo$: Wend
                    Alicuo$ = REPLACAR$(Alicuo$, ".", ",")
                Call REPLA(REBLA$, Alicuo$, 179, 5)

                ' RETENCION PRACTICADA
                    ImporteRet$ = TRIM$(FORMATNUM$(ImporteRetencion#, "F12.2"))
                    While Len(ImporteRet$) < 16: ImporteRet$ = "0" & ImporteRet$: Wend
                    ImporteRet$ = REPLACAR$(ImporteRet$, ".", ",")
                Call REPLA(REBLA$, ImporteRet$, 184, 16)

                ' MONTO TOTAL RETENIDO
                Call REPLA(REBLA$, ImporteRet$, 200, 16)
                
                Print #NX1%, REBLA$
                .MoveNext
            Loop
         End With
         Close #NX1%
         On Error Resume Next
         RstIbCaba.Close
         Set RstIbCaba = Nothing
         On Error GoTo 0
         
         Screen.MousePointer = 1
         
         ASALID3$ = REPLACAR(ASALID$, "\", "/")
         Call COMUNI("Archivos de Transferencia\ '" + ASALID3$ + "\Generado con " + TRIM$(Str$(CAREGI&)) + " Registros.")
        
99       Exit Sub

End Sub


Private Sub IIBBpercepExpo_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
        
    Screen.MousePointer = 11
    Call GENIVAVEN
    Screen.MousePointer = 1
    
    Call INGBRUT
    
End Sub

Private Sub IIBBretenExpo_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
    
    MESAN = MESASEL$
    Call RETENCIONIIBB

End Sub

Private Sub IIBBretenLis_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
    
    Call RETENCIONIIBB(1)
    Call FINAL("*LISRETIB")
    
End Sub

Private Sub Image2_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         If COMALTER%("Confirme Cierre del Período Mensual " + TXT$ + "\\Cancelar\Cerrar") = 2 Then
           Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(Has%)))
           Image2.Visible = False
           Image1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub Image1_Click()
   TXT$ = MESAN.TEXT
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
      If DERACCE%("CIEMEVEN", "Cierre y Re-Apertura Mensual Ventas") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMEVEN", TRIM$(Str$(Des% - 1)))
            Image1.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub

Private Sub IvaPercepExpo_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
    '
    Screen.MousePointer = 11
    Call GENIVAVEN
    Screen.MousePointer = 1

     If ULTREG&("LIVAVEN") < 1 Then Exit Sub
     '
10   Call MESANO(MESAN, DESX1%, HASX1%)
     VDEF$ = MKI$(DESX1%) + MKI$(HASX1%)
     OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
     If OBX$ = "" Then Exit Sub
     DESX2% = CVI(Left$(OBX$, 2))
     HASX2% = CVI(Mid$(OBX$, 3, 2))
     If DESX2% > HASX2% Or DESX2% < DESX1% Or HASX2% > HASX1% Then
       Call MENSERR(24, "Rango de Fechas no Procesable")
       GoTo 10
     End If
     '
     VD1$ = ""
15   CUIEM1$ = InputBox$("Ingrese Nº de C.U.I.T.", "Agente de Percepción", VD1$)
     If TRIM$(CUIEM1$) = "" Then
       Call MENSERR(24, "C.U.I.T. no Válido")
       GoTo 10
     End If
     '
     VD1$ = CUIEM1$
     If VALICUIT%(CUIEM1$) <= 0 Then
      Call MENSERR(24, "C.U.I.T. no Válido")
      GoTo 15
     End If
     '
     ASALID$ = "C:\FLEXOFT\TRAPEIVA.TXT"
     NX% = FILEOPEN%(ASALID$, 2, 128)
     CAREGI% = 0
     For U& = 1 To ULTREG&("LIVAVEN")
       XX$ = REGLEIDO$("LIVAVEN", U&)
         FEX = CVI(Mid$(XX$, 3, 2))
         If FEX < DESX2% Or FEX > HASX2% Then GoTo 29
       If CVD(Mid$(XX$, 161, 8)) > 0 Then ' SI PERCEPCION I.V.A.
         TTXX$ = Space$(60)
         '
         CODPER1$ = CONTROL$("COREPER", "COREPER")
         If TRIM$(CODPER1$) = "" Then
          Call MENSERR(24, "Código de Regimen de Percepción no Válido.\Se ha Cancelado la Exportación.")
          Exit Sub
         End If
         'CUIT DEL AGENTE
         Call REPLA(TTXX$, CUIEM1$, 4, 13)
         'CODIGO DE REGIMEN DE PERCEPCION
         Call REPLA(TTXX$, CODPER1$, 1, 3)
         'FECHA DE LA PERCEPCION
         FEXU$ = FECHATEX$(FEX)
         FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
         Call REPLA(TTXX$, FEXU$, 17, 10)
         'PREFIJO DE FACTURA
         X11$ = Mid$(XX$, 11, 18)
         PREFI1$ = Mid$(X11$, 4, 4)
         While Len(PREFI1$) < 8: PREFI1$ = "0" + PREFI1$: Wend
         Call REPLA(TTXX$, PREFI1$, 27, 8)
         'NUMERO DE FACTURA
         X2 = CVS(Mid$(XX$, 7, 4))
         NUMEFAC1$ = TRIM$(Str$(X2))
         While Len(NUMEFAC1$) < 8: NUMEFAC1$ = "0" + NUMEFAC1$: Wend
         Call REPLA(TTXX$, NUMEFAC1$, 35, 8)
         'MONTO DE LA PERCEPCION
         IMPOL# = CVD(Mid$(XX$, 161, 8))
         IMPOX$ = REPLACAR$(TRIM$(Format$(Abs(IMPOL#), "###########0.00")), ".", ",")
         While Len(IMPOX$) < 16: IMPOX$ = "0" + IMPOX$: Wend
         Call REPLA(TTXX$, IMPOX$, 43, 16)
         '
         Print #NX%, TTXX$
         CAREGI% = CAREGI% + 1
       End If
29   Next U&
     Close #NX%
     ASALID1$ = REPLACAR$(ASALID$, "\", "/")
     Call COMUNI("Archivo de Transferencia '" + ASALID1$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
    
End Sub

Private Sub IvaRetenExpo_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
    
    Call RETENCIONIVA

End Sub

Private Sub IvaRetenLis_Click()

    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    MESAN = MESASEL$
    If VALIMESAN% < 1 Then Exit Sub
    
    Call RETENCIONIVA(1)
    
    Call FINAL("*LIRETIVA")
    
End Sub

Private Sub MESAN_DblClick()
   'Call Command12_Click
   Call DESHASFECHA(Des%, Has%, PERIO$)
   MESAN = PERIO$
End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   Image2.Visible = True
   Image1.Visible = False
   If TXT$ = "" Then Exit Sub
   If MESANO$(TXT$, Des%, Has%) = TXT$ Then
     If XVALO(CONTROL("", "CIEMEVEN")) >= Has% Then
       Image2.Visible = False
       Image1.Visible = True
     End If
   End If
End Sub

Sub CONASIDIA()
    '
    MESA$ = TRIM$(MESAN)
    MESB$ = MESANO$(MESA$, Des%, Has%)
    If MESB$ <> MESA$ Then
      Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(Has%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI")
    '
End Sub
'
Function VALIMESAN%()
   VMA% = 0
   
   If Len(MESAN) = 19 Then
     DESDE% = FECHANUM(Left$(MESAN, 8))
     If DESDE% > 0 Then
       HASTA% = FECHANUM(Mid$(MESAN, 12, 8))
       If HASTA% >= DESDE% Then
         VMA% = 1
         DESDAT = FETEXCOR1$(DESDE%)   ' FECHATEXLAR$(DESDE%)
         HASDAT = FETEXCOR1$(HASTA%)   ' FECHATEXLAR$(HASTA%)
         CORRESP$ = MESAN
         GoTo 98
       End If
     End If
   End If
   
   CORRESP$ = MESANO$(MESAN, DESDE%, HASTA%)
   If CORRESP$ <> "" Then
       VMA% = 1
       DESDAT = FECHATEX(DESDE%)
       HASDAT = FECHATEX(HASTA%)
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   
98 VALIMESAN% = VMA%

End Function

Sub GENSUDIVEN()
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    Call APERBASDAT("CABAN")
'    sqlx$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
'    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
'    SQLX$ = SQLX$ + "Order by FechConta ASC, NumeComp ASC "
    '
    'jandy sql
    SQLX$ = "SELECT * FROM CAJABANC "
    'sqlx$ = sqlx$ + "where FechConta >= " & DESDAT & " "
    'sqlx$ = sqlx$ + "AND FechConta <= " & HASDAT & " "
    SQLX$ = SQLX$ + "where FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "Order by FechConta ASC, NumeComp ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
15  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 15
    End If
    On Error GoTo 0
    
'    '\\\OT-05-0416-WAR-23/06/05///
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
'    SQLX$ = SQLX$ + "where FechConta >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND FechConta <= " & HASDAT & " "
'    '
'    Set CABAMENS = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'    '\\\OT-05-0416-WAR-FIN///

    JJ& = 0: REBLA$ = REGBLAN$("SUDIVEN")
    HUBORECU% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      CANRE& = .RecordCount
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      .MoveFirst
      UKL& = 0
      Do While Not .EOF
        UKL& = UKL& + 1
        Call TRACE(20, UKL&, CANRE&)
        ZZ$ = REBLA$
        NUMEFA$ = !codicom_lar
        NCLI = !NUCLIEN
        FEFAXI% = CVINT(!fechconta)
        FECHAD = FETEXCOR1$(FEFAXI%)
        ITOTA# = !ValAbsComp
        SIG% = Sgn(!idebecomp - !ihabecomp)
        MOBASI$ = !TIPOMOVIM
        SUMABAL# = 0      ' SUMADOR PARA VERIFICAR BALANCE DEL COMPROBANTE
        '
        If Abs(!idebecomp - ROUND#(!idebecomp)) > 0.000001 Or Abs(!ihabecomp - ROUND#(!ihabecomp)) > 0.000001 Then
           HUBORECU% = 1
           Call RECUMOVEN(!codicom_lar, !fechconta)
        End If
        
        If MOBASI$ = "FVEN" Then
          Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 1, 2)
          Call REPLA(ZZ$, !codicom_lar, 3, 18)
          Call REPLA(ZZ$, rasocli$(NCLI), 21, 30)
          Call REPLA(ZZ$, CuitCli$(NCLI), 51, 16)
          Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 67, 8)
          Call REPLA(ZZ$, CODCUE$(!CUECONTAI), 75, 12)
          Call REPLA(ZZ$, MKD$(!idebecomp), 87, 8)
          Call REPLA(ZZ$, MKD$(!ihabecomp), 95, 8)
          JJ& = JJ& + 1
          Call GRAREG("SUDIVEN", ZZ$, JJ&)
          SUMABAL# = SUMABAL# + !idebecomp - !ihabecomp
        End If
        '
        CUEII% = !CUECONTAI
        If CUEII% > 0 Then
          IMASI#(CUEII%) = IMASI#(CUEII%) + (!idebecomp - !ihabecomp)
          If CUEII% > CUEMX% Then CUEMX% = CUEII%
        End If
        '
        'SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        'jandy sql
        SQLX1$ = "Select * from CAJABANC where "
        'SQLX$ = "FechConta = " & FECHAD & " "
        SQLX1$ = SQLX1$ + "FechConta = '" & FECHAD & "' "
        SQLX1$ = SQLX1$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX1$ = SQLX1$ + "AND (NuClien = " & NCLI & " OR (NuClien = 0 AND NuProve = 0)) "
        SQLX1$ = SQLX1$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        Dim CABAMENS As ADODB.Recordset
        Set CABAMENS = New ADODB.Recordset
        CABAMENS.CursorLocation = adUseClient
16      On Error Resume Next
        CABAMENS.Open FILTSQL$(SQLX1), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 16
        End If
        On Error GoTo 0
        '
        With CABAMENS 'CajaBanc
          Do While Not (.EOF)
          '.FindFirst SQLX$
25        'If .NoMatch = False Then
            ZZ$ = REBLA$
            '
            If Abs(!idebecomp - ROUND#(!idebecomp)) > 0.000001 Or Abs(!ihabecomp - ROUND#(!ihabecomp)) > 0.000001 Then
               HUBORECU% = 1
               Call RECUMOVEN(!codicom_lar, !fechconta)
            End If
            '
            If MOBASI$ = "FVEN" Then
              Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 1, 2)
              Call REPLA(ZZ$, !codicom_lar, 3, 18)
              Call REPLA(ZZ$, rasocli$(NCLI), 21, 30)
              Call REPLA(ZZ$, CuitCli$(NCLI), 51, 16)
              Call REPLA(ZZ$, CODCUE$(!CUECONTAI), 75, 12)
              Call REPLA(ZZ$, MKD$(!idebecomp), 87, 8)
              Call REPLA(ZZ$, MKD$(!ihabecomp), 95, 8)
              JJ& = JJ& + 1
              Call GRAREG("SUDIVEN", ZZ$, JJ&)
              SUMABAL# = SUMABAL# + !idebecomp - !ihabecomp
            End If
            '
            CUEII% = !CUECONTAI
            If CUEII% > 0 Then
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!idebecomp - !ihabecomp)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
            '
           '.FindNext SQLX$
           ' GoTo 25
           .MoveNext
           Loop
          'End If
        End With
        CABAMENS.Close
        Set CABAMENS = Nothing
        '
        ' SUMADOR PARA VERIFICAR EL BALANCE DEL COMPROBANTE
        If Abs(SUMABAL#) >= 0.005 Then
           HUBORECU% = 1
           Call RECUMOVEN(!codicom_lar, !fechconta)
        End If
        '
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("SUDIVEN", JJ&)
    '
    Call HEADERS("SUDIVEN", "")
    Call HEADERS("SUDIVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
    TTXX$ = Chr$(1) + Chr$(1) + MKI$(HASTA%)
    VV1& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        VV1& = VV1& + 1
        Call GRAREG("ASIDIAR", X$, VV1&)
      End If
    Next U&
    '
    AGRECU% = 0
    For k% = 1 To CUEMX%
      If Abs(IMASI#(k%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(k%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(k%)), 39, 8)
        VV1& = VV1& + 1
        Call GRAREG("ASIDIAR", X$, VV1&)
        '
        If ECOARCH$("CUECON", MKI$(k%)) = "" Then
          Z3$ = REGBLAN$("CUECON")
          Call REPLA(Z3$, MKI$(k%), 1, 2)
          Call REPLA(Z3$, DENOCUE$(k%), 3, 30)
          Call REPLA(Z3$, CODCUE$(k%), 33, 12)
          Call REPLA(Z3$, CODCUE$(k%), 45, 12)
          Call REGAPP("CUECON", Z3$)
          AGRECU% = 1
        End If
      End If
    Next k%
    '
    Call SETULTREG("ASIDIAR", VV1&)
    '
    If AGRECU% > 0 Then
      Call FILESORT("CUECON", "", 1, 1, "ASC")
    End If
    Call CIERRARCH("CUECON")
    '
    If HUBORECU% > 0 Then
      Call COMUNI("Errores de Generación de Subdiario de Ventas.\La Información Generada no es Confiable.\   \Vuelva a Pedir el Listado.")
      Call SETULTREG("SUDIVEN", 0)
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub

Sub RECUMOVEN(NUMEFA$, FECONTAFA)
   '
   FECONTAD = Int(CDbl(FECONTAFA))
      
'   sqlx$ = "SELECT * FROM CAJABANC "
'   sqlx$ = sqlx$ + "where CodiCom_Lar = '" & NUMEFA$ & "' "
'   sqlx$ = sqlx$ + "AND INT(CDBL(FechConta)) = " & FECONTAD & " "
'   sqlx$ = sqlx$ + "Order by ABS(IDebeComp - IHabeComp) ASC "
'   '
'   Set MOFASEL = CueCorri.OpenRecordset(sqlx$, dbOpenDynaset)
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CodiCom_Lar = '" & NUMEFA$ & "' "
   SQLX$ = SQLX$ + "AND FechConta = '" & FECONTAFA & "' "
   SQLX$ = SQLX$ + "Order by ABS(IDebeComp - IHabeComp) ASC "
   
   Dim MOFASEL As ADODB.Recordset
   Set MOFASEL = New ADODB.Recordset
25 On Error Resume Next
   MOFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   SUMACOMP# = 0
   On Error Resume Next
   With MOFASEL
     .MoveFirst
     If Err > 0 Then
       On Error GoTo 0
       Exit Sub
     End If
     Do While Not .EOF
       IDEBAS# = ROUND#(!idebecomp, 2)
       IHABAS# = ROUND#(!ihabecomp, 2)
       '
       '.Edit
       !idebecomp = IDEBAS#
       !ihabecomp = IHABAS#
       !ValAbsComp = Abs(IDEBAS# - IHABAS#)
       .Update
       SUMACOMP# = SUMACOMP# + IDEBAS# - IHABAS#
     .MoveNext
     Loop
   End With
   On Error GoTo 0
   '
   If Abs(SUMACOMP#) >= 0.005 Then
     If Abs(SUMACOMP#) < 0.1 Then
       TTXX$ = "Diferencia de " + TRIM$(FORMATNUM$(SUMACOMP#, "F12.2")) + " en Comprobante " + TRIM$(NUMEFA$)
       TTXX$ = TTXX$ + "\   \Desea Recuperar Esta Diferencia Automáticamente ? \\Si, Recuperar\No, Dejar como está"
       If COMALTER%(TTXX$) = 1 Then
         MOFASEL.MoveLast
         'MOFASEL.Edit
         If MOFASEL!idebecomp >= 0.005 Then
             MOFASEL!idebecomp = MOFASEL!idebecomp - SUMACOMP#
           Else
             MOFASEL!ihabecomp = MOFASEL!ihabecomp + SUMACOMP#
         End If
         MOFASEL.Update
       End If
     End If
   End If
   '
   MOFASEL.Close
   Set MOFASEL = Nothing
  
End Sub

Sub GENIVAVEN(Optional MODO%, Optional MODOLibroIvaDigital%)
    '
    Screen.MousePointer = 11
    '
    Dim RELIVA$(32)
    Dim IMASI#(32767)
    Dim NUFAK$(), TIMOK$(), VAMOK$(), IPAXK#(), BASEK#(), CUECK%(), OPMOK$()
    UBA& = 1024
    ReDim Preserve NUFAK$(UBA&), TIMOK$(UBA&), VAMOK$(UBA&), IPAXK#(UBA&), BASEK#(UBA&), CUECK%(UBA&), OPMOK$(UBA&)
    CUEMX% = 0
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    '
'    ' ESTO CORRIGE UN ERROR DE LA IMPORTACION QUE CONSIDERO IIBB CON IVA 10.5
'    If InStr(1, EMPREAC$, "FERVI") Then
'            FF1% = FECHANUM("01/07/12"): FECHA1$ = FETEXCOR1$(FF1%)
'            FF2% = FECHANUM("31/12/12"): FECHA2$ = FETEXCOR1$(FF2%)
'            SQLX$ = "select * from cajabanc where (codicom_lar like 'fc%' or codicom_lar like 'nd%' or codicom_lar like 'nc%') and tipomovim='iva' and opcimovim='defi' and varimovim='10.5 %' and nuclien>0 and fechconta>='" & FECHA1$ & "' AND fechconta<='" & FECHA2$ & "'"
'            Dim CORRIGEIIBB As ADODB.Recordset
'            Set CORRIGEIIBB = New ADODB.Recordset
'15          On Error Resume Next
'            CORRIGEIIBB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'              On Error GoTo 0
'              Call CapturaErrorOpen
'              GoTo 15
'            End If
'            On Error GoTo 0
'            With CORRIGEIIBB
'                Do While Not .EOF
'                    !TipoMovim = "IBRU"
'                    !OpciMovim = "IB-B"
'                    !VARIMOVIM = "PERC"
'                    !DescriMov = "Perc.IB.Bs.As."
'                    .Update
'                    .MoveNext
'                Loop
'            End With
'            CORRIGEIIBB.Close
'            Set CORRIGEIIBB = Nothing
'            '
'    End If
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    'SQLX$ = SQLX$ + "where FechEmisi >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND FechEmisi <= " & HASDAT & " "
    Campox$ = "FECHEMISI"
    If MODO% = 1 Then Campox$ = "FECHCONTA"
    SQLX$ = SQLX$ + " where " & Campox$ & ">= '" & DESDAT & "' "
    SQLX$ = SQLX$ + " AND " & Campox$ & "<= '" & HASDAT & "' "
    SQLX$ = SQLX$ + " AND NUCLIEN > 0 AND NUMECOMP > 0"
    'SQLX$ = SQLX$ + "where FechEmisi >= '" & DESDAT & "' "
    'SQLX$ = SQLX$ + "AND FechEmisi <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + " Order by FechEmisi ASC, NumeComp ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    JJ& = 0: JPER& = 0
    REBLA$ = REGBLAN$("LIVAVEN")
    REBLB$ = REGBLAN$("REIVAVEN")
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 50
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      U& = 0: JJU& = 0: VECOMP$ = ""
      '
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '
        INCLU$ = ""
        On Error Resume Next
        INCLU$ = TRIM$(UCase$(!IncSubIva))
        On Error GoTo 0
        If INCLU$ = "NO" Then GoTo 19
        '
        TIMO$ = TRIM$(UCase$(!TIPOMOVIM))
        NUMEFA$ = TRIM$(!codicom_lar)
        CbteOriginalExportaSimple$ = TRIM$(SAFETEXT$(!codicom_lar_ExSimple))
        If CbteOriginalExportaSimple$ <> "" Then
            NUMEFA$ = CbteOriginalExportaSimple$
        End If
           ' ESTO ES PARA EXCLUIR LOS AJUSTES DE SALDO DOSIVAC
           If Left$(NUMEFA$, 2) = "AJ" Then GoTo 19
        NCLI = !NUCLIEN
        FECHAD = Int(CDbl(!fechconta))
        ITOTA# = !ValAbsComp
        If MODOLibroIvaDigital% = 1 Then ITOTA# = XVALO(!Impo_Dola)
        SIG% = Sgn(!idebecomp - !ihabecomp)
        If TIMO$ = "FVEN" Then
              ZZ$ = REBLA$
              VECOMP$ = VECOMP$ + AJUSTI$(NUMEFA$, 18)
              Call REPLA(ZZ$, MKI$(CVINT(!fechconta)), 3, 2)
              Call REPLA(ZZ$, MKI$(CVINT(!FECHEMISI)), 5, 2)
              Call REPLA(ZZ$, MKS$(!NumeComp), 7, 4)
              Call REPLA(ZZ$, NUMEFA$, 11, 18)
'Call REPLA(ZZ$, MKI$(NCLI), 31, 2)
              If NCLI = 0 And TRIM$(UCase$(!OpciMovim)) = "ANUL" Then 'si esta anulado
                  Call REPLA(ZZ$, "COMPROBANTE ANULADO", 33, 30)
              Else
                  Call REPLA(ZZ$, rasocli$(NCLI), 33, 30)
              End If
              '
              CUITT$ = ""
              PIVAC% = PIVACLI%(NCLI)
              If PIVAC% = 0 Or PIVAC% = 5 Then
                CUITT$ = TRIM$(SAFETEXT$(!Cuit))
                If CUITT$ = "" Then CUITT$ = CuitCli$(NCLI)
                If PIVAC% = 5 And VALICUIT%(CUITT$) = 0 Then CUITT$ = CuitCli$(NCLI)
              Else
                CUITT$ = CuitCli$(NCLI)
              End If
              Call REPLA(ZZ$, CUITT$, 63, 16)
              If InStr(UCase$(EMPREAC$), "FERVI") > 0 And PIVACLI%(NCLI) = 5 Then
                  'CASO DE FERVI CLIENTES DE EXPORTACION NO PRESENTA EL CUIT, SEGUN SOLICITUD
                  'Call REPLA(ZZ$, String$(16, " "), 63, 16)
              End If
              
              If NCLI = 0 And TRIM$(UCase$(!OpciMovim)) = "ANUL" Then 'si esta anulado
                  Call REPLA(ZZ$, "", 79, 12)
              Else
                  Call REPLA(ZZ$, POSIVAC$(NCLI), 79, 12)
              End If
              '
              If Abs(ITOTA#) < 0.005 Then 'si esta anulado
                  Call REPLA(ZZ$, "COMPROBANTE ANULADO", 33, 58)
              End If
              '
              Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 91, 8)
              '
              '\\\OT-05-0709-WAR-02/11/05///
              Pi% = PIVACLI%(NCLI)
              If Pi% = 5 Or Pi% = 7 Then
                Call REPLA(ZZ$, MKD$(ITOTA# * SIG%), 185, 8)
              End If
              '\\\OT-05-0709-WAR-FIN///
              If TRIM$(Mid$(ZZ$, 129, 8)) = "" Then
                Call REPLA(ZZ$, "NO GRAV.", 129, 8)
              End If
              '
              JJU& = JJU& + 1
              Call GRAREG("!LIVAVEN", ZZ$, JJU&)
            ElseIf InStr("IVA_GAN_IBRU", TIMO$) > 0 Then
              OPMO$ = TRIM$(!OpciMovim)
              VAMO$ = TRIM$(!VARIMOVIM)
              IPAX# = Abs(ITOTA#)
              Base# = 0
              On Error Resume Next
              Base# = Abs(!BASEIMPO)
              If MODOLibroIvaDigital% = 1 Then
                IPAX# = XVALO(!Impo_Dola)
                TIPOCAMBIO# = XVALO(!Tipo_Cam)
                If TIPOCAMBIO# > 1 Then Base# = Base# / TIPOCAMBIO#
              End If
              On Error GoTo 0
              '
              If Abs(IPAX#) >= 0.005 Or Base# >= 0.005 Then
                JK& = JK& + 1
                If JK& > UBound(NUFAK$) Then
                  UBA& = 1024 + UBound(NUFAK$)
                  ReDim Preserve NUFAK$(UBA&), TIMOK$(UBA&), VAMOK$(UBA&), IPAXK#(UBA&), BASEK#(UBA&), CUECK%(UBA&), OPMOK$(UBA&)
                End If
                NUFAK$(JK&) = NUMEFA$
                TIMOK$(JK&) = TIMO$
                VAMOK$(JK&) = VAMO$
                IPAXK#(JK&) = IPAX#
                BASEK#(JK&) = Base#
                CUECK%(JK&) = !CUECONTAI
                OPMOK$(JK&) = OPMO$
              End If
        End If
19    .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    '
    For U2& = 1 To JJU&
       Call TRACE(20, U2&, JJU&)
       RELIVA$(0) = REGLEIDO$("!LIVAVEN", U2&)
       NUMEFA$ = TRIM$(Mid$(RELIVA$(0), 11, 18))
       If InStr(VECOMP$, NUMEFA$) < 1 Then GoTo 49
       '
       IREL$ = Left$(RELIVA$(0), 10)
       For KU% = 1 To 32
         RELIVA$(KU%) = REBLA$
         Mid$(RELIVA$(KU%), 1, 10) = IREL$
       Next KU%
       '
       SIG% = Sgn(CVD(Mid$(RELIVA$(0), 91, 8)))
       ULIVA% = 0: PROIVA% = 0: HAYIVA% = 0: PROIB% = -1
       '
       FECHCO% = CVI(Mid$(RELIVA$(0), 3, 2))
       NCLI = CVI(Mid$(RELIVA$(0), 31, 2))
       RACLI$ = Mid$(RELIVA$(0), 33, 30)
       CTTC$ = Mid$(RELIVA$(0), 63, 16)
       ITOTACO# = CVD(Mid$(RELIVA$(0), 91, 8))
       '
       INOGRAD# = ITOTACO#
       For JK1& = 1 To JK&
          If NUFAK$(JK1&) = NUMEFA$ Then
             TIMO$ = TIMOK$(JK1&)
             VAMO$ = VAMOK$(JK1&)
             IPAX# = IPAXK#(JK1&)
             Base# = BASEK#(JK1&)
             CUECO% = CUECK%(JK1&)
             OPMO$ = OPMOK$(JK1&)
             '
             If TIMO$ = "IVA" And UCase$(VAMO$) = "PER-IVA" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 161, 8)
                  RELIVA$(0) = ZZ$
                ElseIf (TIMO$ = "GAN" Or TIMO$ = "IVA") And UCase$(VAMO$) = "PER-GAN" Then
                  ZZ$ = RELIVA$(0)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 169, 8)
                  RELIVA$(0) = ZZ$
                ElseIf TIMO$ = "IBRU" And UCase$(VAMO$) = "PERC" Then
                  ZZ$ = RELIVA$(0)
                  ' ACUMULA TODAS LAS PERCEPCIONES Y LUEGO LAS SEPARA CUANDO ARMA EL LIBRO IVA VENTAS SOBRE LA GRILLAGRAL (VER LibroIvaVentas)
                    IMPOPERIB# = CVD(Mid$(ZZ$, 193, 8))
                  Call REPLA(ZZ$, MKD$(IMPOPERIB# + (IPAX# * SIG%)), 193, 8)
                  'Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 193, 8)
                  Call REPLA(ZZ$, MKI$(CUECO%), 121, 2)
                    JURIB$ = TRIM$(Mid$(OPMO$, 4, 1))
                  Call REPLA(ZZ$, JURIB$, 203, 1)
                  Call REPLA(ZZ$, ECOARCH$("PROVINC", JURIB$), 205, 24)
                  RELIVA$(0) = ZZ$
                  '
                  If Abs(IPAX#) >= 0.005 And OPMO$ = "IB-B" Then
                    ZPER$ = REGBLAN$("SUPERIB")
                    Call REPLA(ZPER$, MKI$(FECHCO%), 1, 2)
                    Call REPLA(ZPER$, NUMEFA$, 3, 18)
                    Call REPLA(ZPER$, RACLI$, 21, 30)
                    Call REPLA(ZPER$, NIBRCLI$(NCLI), 65, 14)
                    Call REPLA(ZPER$, CTTC$, 51, 14)
                    Call REPLA(ZPER$, MKD$(Base# * SIG%), 79, 8)
                    Call REPLA(ZPER$, MKD$(IPAX# * SIG%), 87, 8)
                    JPER& = JPER& + 1
                    Call GRAREG("SUPERIB", ZPER$, JPER&)
                  End If
                  '
                ElseIf TIMO$ = "IVA" Then
                  ZZ$ = RELIVA$(PROIVA%)
                    If TRIM$(VAMO$) = "" Then
                       VAMO$ = "NO GRAV."
                    End If
                  Call REPLA(ZZ$, VAMO$, 129, 8)
                  Call REPLA(ZZ$, MKD$(Base# * SIG%), 137, 8)
                  Call REPLA(ZZ$, MKD$(IPAX# * SIG%), 145, 8)
                  
                  ' *** ESTE CODIGO VERIFICA SI LA BASE DEL IVA SEGUN CAJABANC SE CORRESPONDE CON EL IMPORTE DE IVA. EN CASO DE NO SERLO,
                  ' *** GUARDA LA REAL CALCULADA EN LA POSICION 137, 8. MAS ADELANTE EXISTE CODIGO QUE GRABA LA DIFERENCIA DE LA BASE
                  ' *** EN CAJABANC Y LA REAL CALCULADA, COMO UN IMPORTE NO GRAVADO.
                  ' *** APLICA POR EJEMPLO, A NOTAS DE DEBITO POR RECHAZO. EL IMPORTE GRAVADO (GASTOS DEL RECHAZO) Y EL NO GRAVADO (EL CHEQUE),
                  ' *** SE GRABAN SUMADOS EN EL REGISTRO CORRESPONDIENTE AL IVA. EL SIGUIENTE CODIGO LO SEPARA.
                  VAMO5$ = REPLACAR$(VAMO$, "%", " ")
                  If XVALO(VAMO5$) >= 0.1 And InStr(1, NUMEFA$, "FC") < 1 Then
                      BASECALC# = 100 * IPAX# / (XVALO(VAMO5$))
                      INOGRAD1# = Base# - BASECALC#
                      
                      ' *** ESTE ES EL CODIGO ANTERIOR. ESTA MAL PORQUE NUNCA SE CUMPLE InStr(1, NUMEFA$, "NC") < 0
'                      If Abs(INOGRAD1#) >= 1 And InStr(1, NUMEFA$, "NC") < 0 Then
'                        Call REPLA(ZZ$, MKD$(BASECALC# * SIG%), 137, 8)
'                      End If
                                              
                      ' *** SE REEMPLAZA EL ANTERIOR POR ESTE NUEVO CODIGO Y SE IMPLEMENTA A PARTIR DE JULIO
                      ' *** PORQUE ES EL MES DONDE SE HACE EL CAMBIO Y SINO ALTERARIAMOS EL LIBRO DE IVA DE
                      ' *** LOS MESES ANTERIORES QUE FUERON PRESENTADOS Y HASTA QUIZAS COPIADOS EN LIBRO VENTAS
                      If Abs(INOGRAD1#) >= 1 And InStr(1, NUMEFA$, "NC") < 1 And FECHCO% >= FECHANUM("01/07/17") Then
                        Call REPLA(ZZ$, MKD$(BASECALC# * SIG%), 137, 8)
                      End If
                  End If
'                     If Abs(Base#) >= 0.005 Then   ' PARA CONSIDERAR CASOS DE SOLO IVA
'                        VAMO5$ = REPLACAR$(VAMO$, "%", " ")
'                        If XVALO(VAMO5$) >= 0.1 Then
'                          'BASECALC# = 100 * IPAX# / (XVALO(VAMO5$))
'                          INOGRAD1# = 0
'                          'INOGRAD1# = BASE# - BASECALC#
'                          'If Abs(INOGRAD1#) >= 1 Then
'                          '     Call REPLA(ZZ$, MKD$(BASECALC# * SIG%), 137, 8)
'                          '   Else
'                          '     INOGRAD1# = 0
'                          'End If
'                          'If BASE# < 0.01 Then
'                          '  Call REPLA(ZZ$, MKD$(BASECALC# * SIG%), 137, 8)
'                          'End If
'                        End If
'                     End If
                  BASE2# = SIG% * CVD(Mid$(ZZ$, 137, 8))
                  If Abs(BASE2#) > 0.005 Then
                    Call REPLA(ZZ$, MKD$(100 * IPAX# / BASE2#), 153, 8)
                  End If
                  RELIVA$(PROIVA%) = ZZ$
                  ULIVA% = PROIVA%
                  PROIVA% = PROIVA% + 1
                  HAYIVA% = 1
             End If
          End If
       Next JK1&
       '
       ' LO QUE SIGUE ES PARA OBTENER EL IMPORTE NO GRAVADO POR DIFERENCIA
       For KKI% = 0 To ULIVA%
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 137, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 145, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 161, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 169, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 177, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 185, 8)))
          INOGRAD# = INOGRAD# - ROUND#(CVD(Mid$(RELIVA$(KKI%), 193, 8)))
       Next KKI%
       '
       If Abs(INOGRAD#) >= 0.005 Then
          ZZ$ = RELIVA$(PROIVA%)
          Call REPLA(ZZ$, RELIVA$(0), 1, 90)
             VAMO$ = "NO GRAV."
             INOGRAD# = INOGRAD# + CVD(Mid$(ZZ$, 137, 8))
          Call REPLA(ZZ$, VAMO$, 129, 8)
          Call REPLA(ZZ$, MKD$(INOGRAD#), 137, 8)
          Call REPLA(ZZ$, MKD$(0), 145, 8)
          RELIVA$(PROIVA%) = ZZ$
          ULIVA% = PROIVA%
          PROIVA% = PROIVA% + 1
          HAYIVA% = 1
       End If
       '
       If PROIB% > ULIVA% Then ULIVA% = PROIB%
       For KU% = 0 To ULIVA%
         JJ& = JJ& + 1
         Call GRAREG("LIVAVEN", RELIVA$(KU%), JJ&)
       Next KU%
       '
49  Next U2&
    '
50  Call SETULTREG("LIVAVEN", JJ&)
    Call SETULTREG("SUPERIB", JPER&)
    '
    Call HEADERS("LIVAVEN", "")
    Call HEADERS("LIVAVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
    ' CARGA RESUMEN MENSUAL DE IVA - VENTAS
    Dim DECATIV$(128)
    CACATIV% = 0
    Dim VECREIVA$(256)
    UVECRE% = 0
    '
    MODOTANIT% = 0
    If EXISTE%("LIVENTAN.PRF") > 0 Then
        Call HEADERS("LIVENTAN", "")
        Call HEADERS("LIVENTAN", "Corresponde a: " + TRIM$(CORRESP$))
        MODOTANIT% = 1
    End If
    '
    KUT& = 0
    For KUL& = 1 To JJ&
      Call TRACE(20, KUL&, JJ&)
      RECUE$ = REGLEIDO$("LIVAVEN", KUL&)
      If TRIM$(Mid$(RECUE$, 11, 18)) <> "" Then
        NUCO$ = TRIM$(Mid$(RECUE$, 11, 18))
        SIG% = Sgn(CVD(Mid$(RECUE$, 91, 8)))
      End If
      '
      CORUBRO% = 1: DERUBRO$ = "VENTA" '\\\OT-06-0079-WAR-03/03/06///
      COCONCE% = 8: DECONCE$ = "COMPROBANTES VARIOS"
      '
      If SIG% < 0 Then
          COCONCE% = 9: DECONCE$ = "NOTAS DE CREDITO"
        Else
          If Left$(NUCO$, 2) = "FC" Then COCONCE% = 1: DECONCE$ = "FACTURAS"
          If Left$(NUCO$, 2) = "ND" Then COCONCE% = 2: DECONCE$ = "NOTAS DE DEBITO"
          If Left$(NUCO$, 2) = "TF" Then COCONCE% = 3: DECONCE$ = "TICKET-FACTURA"
          If Left$(NUCO$, 2) = "RH" Then COCONCE% = 4: DECONCE$ = "RECIBO HONORARIOS"
      End If
      '
      If TRIM$(Mid$(RECUE$, 129, 8)) <> "" Then
         DECAT1$ = TRIM$(Mid$(RECUE$, 129, 8))
      End If
      If TRIM$(Mid$(RECUE$, 79, 12)) <> "" Then
         DECAT2$ = Mid$(RECUE$, 79, 12)
      End If
      DECATE$ = DECAT1$ + " - " + DECAT2$
      For KLI% = 1 To CACATIV%
        If DECATE$ = DECATIV$(KLI%) Then
          COCATE% = KLI%
          GoTo 20
        End If
      Next KLI%
      CACATIV% = CACATIV% + 1
      COCATE% = CACATIV%
      DECATIV$(CACATIV%) = DECATE$
      '
20    BASEIMPO# = CVD(Mid$(RECUE$, 137, 8)) + CVD(Mid$(RECUE$, 185, 8))
      VALOIMPU# = CVD(Mid$(RECUE$, 145, 8))
      
82    If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
        ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
        ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
        ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
        '
        For UVE% = 1 To UVECRE%
          If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
             VCC$ = VECREIVA$(UVE%)
             GoTo 84
          End If
        Next UVE%
        UVECRE% = UVECRE% + 1
        UVE% = UVECRE%
        VCC$ = REBLB$
        Call REPLA(VCC$, ABUS$, 1, 64)
        '
84      BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
        VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
        Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
        Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
        Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
        VECREIVA$(UVE%) = VCC$
      End If
      '
      CORUBRO% = 2: DERUBRO$ = "PERCEPCION"        ' PERCEPCIONES
      For COCO% = 1 To 3
        COCONCE% = 10 + COCO%
        BASEIMPO# = 0
        COCATE% = 0: DECATE$ = ""
        VALOIMPU# = CVD(Mid$(RECUE$, 153 + 8 * COCO%, 8))
        If COCO% = 1 Then
           DECONCE$ = "PERCEPCION I.V.A."
          ElseIf COCO% = 2 Then
           DECONCE$ = "PERCEPCION GANANCIAS"
          ElseIf COCO% = 3 Then
           DECONCE$ = "PERCEPC.ING.BRUTOS"
           DECATE$ = Mid$(RECUE$, 205, 24)
           VALOIMPU# = CVD(Mid$(RECUE$, 193, 8))
        End If
        '
92      If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
          ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
          ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
          ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
          '
          For UVE% = 1 To UVECRE%
            If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
               VCC$ = VECREIVA$(UVE%)
               GoTo 94
            End If
          Next UVE%
          UVECRE% = UVECRE% + 1
          UVE% = UVECRE%
          VCC$ = REBLB$
          Call REPLA(VCC$, ABUS$, 1, 64)
          '
94        BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
          VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
          Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
          Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
          Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
          VECREIVA$(UVE%) = VCC$
        End If
         '
      Next COCO%
      '
      ' PARA TANIT SEPARAR IMPORTE NO GRAVADO
      If MODOTANIT% = 1 Then
        RECUE1$ = Left$(RECUE$, 152) + String$(8, 0) + Mid$(RECUE$, 153)
        TANOMIX$ = TRIM$(REPLACAR$(Mid$(RECUE1$, 128, 8), "%", " "))
        TANOMI = XVALO(TANOMIX$)
        '
        ITONE# = CVD(Mid$(RECUE1$, 137, 8)) + CVD(Mid$(RECUE1$, 193, 8))
        INEGRA# = ITONE#
        IEXEN# = 0: IEXEN1# = 0: IEXEN2# = 0
        IMIVA# = CVD(Mid$(RECUE1$, 145, 8))
        If Abs(IMIVA#) < 0.005 Then
           IEXEN# = INEGRA#
           INEGRA# = 0
           GoTo 95
        End If
        If Abs(INEGRA# * TANOMI / 100 - IMIVA#) > 0.5 Then
           If Abs(INEGRA#) >= 0.005 Then
             ' NO DEJA QUE SE CORRIJA EL NETO POR UN ERROR DEL CALCULO DEL IVA EN FACTURACION
             'INEGRA# = 100 * IMIVA# / TANOMI
             IEXEN# = ITONE# - INEGRA#
             ' ESTO CORRIGE UN POSIBLE ERROR EN EL CALCULO DEL IVA QUE RESULTA EN UN NO GRAVADO IRREAL
             If Sgn(IEXEN#) <> Sgn(INEGRA#) Then IEXEN# = 0
           End If
        End If
95      If UCase$(Mid$(RECUE1$, 11, 1)) = "F" Then
            IEXEN1# = IEXEN#
          Else
            IEXEN2# = IEXEN#
        End If
        '
        If Abs(INEGRA#) < 0.005 Then
          If KUT& > 0 Then
            RECUE2$ = REGLEIDO$("LIVENTAN", KUT&)
            If Mid$(RECUE1$, 11, 80) = Mid$(REBLA$, 11, 80) Then
              'If Abs(CVD(Mid$(RECUE2$, 153, 8))) < 0.005 Then
                'If Abs(CVD(Mid$(RECUE2$, 161, 8))) < 0.005 Then
              IEXA1# = CVD(Mid$(RECUE2$, 153, 8))
              IEXA2# = CVD(Mid$(RECUE2$, 161, 8))
              Call REPLA(RECUE2$, MKD$(IEXA1# + IEXEN1#), 153, 8)
              Call REPLA(RECUE2$, MKD$(IEXA2# + IEXEN2#), 161, 8)
              Call GRAREG("LIVENTAN", RECUE2$, KUT&)
              GoTo 96
                'End If
              'End If
            End If
          End If
        End If
        '
        KUT& = KUT& + 1
        Call REPLA(RECUE1$, MKD$(INEGRA#), 137, 8)
        Call REPLA(RECUE1$, MKD$(IEXEN1#), 153, 8)
        Call REPLA(RECUE1$, MKD$(IEXEN2#), 161, 8)
        Call GRAREG("LIVENTAN", RECUE1$, KUT&)
      End If
      ' FIN PARA TANIT
      '
96  Next KUL&
    '
    If MODOTANIT% = 1 Then
      Call SETULTREG("LIVENTAN", KUT&)
    End If
    '
    JJK& = 0
    For UI1% = 1 To UVECRE%
      JJK& = JJK& + 1
      Call GRAREG("REIVAVEN", VECREIVA$(UI1%), JJK&)
    Next UI1%
    Call SETULTREG("REIVAVEN", JJK&)
    Call CIERRARCH("REIVAVEN")
    '
    Call HEADERS("REIVAVEN", "")
    Call HEADERS("REIVAVEN", "Corresponde a: " + TRIM$(CORRESP$))
    '
99  Screen.MousePointer = 1
    '
End Sub

Sub EXCLUIVA()
 '
    If DERACCE%("SELCOIVA", "Selección de Comprobantes a Listar") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Proveedor
    '     ****************************
    '
'    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
'    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) >= " & DESDAT & " "
'    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND FechConta >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.CursorLocation = adUseClient
25  On Error Resume Next
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    JJ& = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NUCLIEN
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!codicom_lar, 13, 6))
            If Len(!codicom_lar) < 18 Then
              NPX& = XVALO(Right$(!codicom_lar, 6))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUECO", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUECO", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUECO") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Registra\Facturas de Venta que Puedan Seleccionarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUECO")
    Call FRESHECHO("!SELCUECO")
    '
    Call SELECHO("!SELCUECO")
    If VALACT1$("!SELCUECO") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUECO"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    Call APERBASDAT("CABAN")
    TTXX$ = "Libro de I.V.A. Ventas\  \Comprobante '" + TRIM$(DOCULA$) + "' \de Cuenta '" + TRIM$(rasocli$(NC%)) + "' ?"
    If COMALTER%(TTXX$ + "\\Incluye\Excluye de Listado") = 1 Then
      COINC = 1
    End If
    
'    With CajaBanc
'      .FindFirst "CodiCom_Lar = '" & DOCULA$ & "' AND NuClien = " & NC% & " "
'20     If .NoMatch = False Then
'        .Edit
'         If COINC = 1 Then
'           !IncSubIva = ""
'         Else
'           !IncSubIva = "NO"
'         End If
'        .Update
'        .FindNext "CodiCom_Lar = '" & DOCULA$ & "' AND NuClien = " & NC% & ""
'        GoTo 20
'      End If
'    End With
    '
    strSQL = "Update CajaBanc "
    If COINC = 1 Then
       strSQL = strSQL & "set IncSubIva='' "
    Else
       strSQL = strSQL & "set IncSubIva='No' "
    End If
    strSQL = strSQL & "where CodiCom_Lar='" & DOCULA$ & "' AND NuClien = " & NC%
    Call EJECUTACOMANDO(strSQL)
    
    GoTo 10
    '
End Sub
        
Sub IVADOSI()
    '\\\OT-08-0890-OD-30/09/08///
      FIN& = ULTREG("LIVAVEN")
      n& = 0
      For k& = 1 To FIN&
        LLXX$ = REGLEIDO$("LIVAVEN", k&)
        FACTU1$ = Mid$(LLXX$, 11, 18)
          If FACTU1$ <> FACTU2$ Then IVXX$ = "" 'LIMPIA EN EL CASO QUE LA FACTURA CAMBIE
          FECHA% = CVI(Mid$(LLXX$, 3, 2))
          DOCUM$ = Mid$(LLXX$, 11, 18)
          DenomCuenta$ = Mid$(LLXX$, 33, 30)
          Cuit$ = Mid$(LLXX$, 63, 16)
          PercInBrto# = CVD(Mid$(LLXX$, 193, 8))
          PercIVA# = CVD(Mid$(LLXX$, 161, 8))
          NETO# = CVD(Mid$(LLXX$, 91, 8))
          '
          NEGRAVADO# = CVD(Mid$(LLXX$, 137, 8)) + CVD(Mid$(LLXX$, 185, 8))
          IMPOIVA# = CVD(Mid$(LLXX$, 145, 8))
          TASAREAL = ROUND#(CVD(Mid$(LLXX$, 153, 8)))
          '
          EXENTO# = 0: GRAVA105# = 0: GRAVA21# = 0: IVA105# = 0: Iva21# = 0
          If TASAREAL < 0.2 Then
               EXENTO# = NEGRAVADO#
            ElseIf TASAREAL >= 10.2 And TASAREAL < 10.8 Then
               GRAVA105# = NEGRAVADO#  'NETO GRAVADO TASA 21******
               IVA105# = IMPOIVA#      'IVA TASA 21***************
            Else
               GRAVA21# = NEGRAVADO#  'NETO GRAVADO TASA 21******
               Iva21# = IMPOIVA#      'IVA TASA 21***************
          End If
          
          REGISTRONUEVO% = 1
          IVXX$ = REGBLAN$("SUIVA19")
          If n& > 0 Then
             IVXX$ = REGLEIDO$("SUIVA19", n&)
             REGISTRONUEVO% = 0
          End If
          If Mid$(IVXX$, 3, 18) <> DOCUM$ And TRIM$(DOCUM$) <> "" Then
             n& = n& + 1
             IVXX$ = REGBLAN$("SUIVA19")
             REGISTRONUEVO% = 1
          End If
          '
          If REGISTRONUEVO% = 1 Then
             Call REPLA(IVXX$, MKI(FECHA%), 1, 2) 'FECHA
             Call REPLA(IVXX$, DOCUM$, 3, 18) 'TIPO Y NUMERO DE DOCUMENTO
             Call REPLA(IVXX$, DenomCuenta$, 21, 30) 'DENOMINACION DE LA CUENTA
             Call REPLA(IVXX$, Cuit$, 51, 14) 'NUM CUIT
             Call REPLA(IVXX$, MKD(NETO#), 153, 8) 'TOTAL FACTURA/NETO
          End If
          '
          Call REPLA(IVXX$, MKD(EXENTO# + CVD(Mid$(IVXX$, 65, 8))), 65, 8)       'IMPORTE EXENTO*******
          Call REPLA(IVXX$, MKD(GRAVA105# + CVD(Mid$(IVXX$, 73, 8))), 73, 8)     'IMPORTE EXENTO*******
          Call REPLA(IVXX$, MKD(IVA105# + CVD(Mid$(IVXX$, 81, 8))), 81, 8)       'IMPORTE EXENTO*******
          Call REPLA(IVXX$, MKD(GRAVA21# + CVD(Mid$(IVXX$, 89, 8))), 89, 8)      'IMPORTE EXENTO*******
          Call REPLA(IVXX$, MKD(Iva21# + CVD(Mid$(IVXX$, 97, 8))), 97, 8)        'IMPORTE EXENTO*******
          Call REPLA(IVXX$, MKD(PercInBrto# + CVD(Mid$(IVXX$, 105, 8))), 105, 8) 'NETO GRAVADO TASA 9.5/PERC.I.B
          Call REPLA(IVXX$, MKD(PercIVA# + CVD(Mid$(IVXX$, 145, 8))), 145, 8)  'RETENCION DEL IVA/PERC.GANANCIAS
          '
          Call GRAREG("SUIVA19", IVXX$, n&)
          '
          FACTU2$ = FACTU1$
      Next k&
      Call SETULTREG("SUIVA19", n&)
      Call CIERRARCH("SUIVA19")
      '
      Call HEADERS("SUIVA19", "")
      Call HEADERS("SUIVA19", "Corresponde a: " + MESAN)
      '
      Call FILESORT("SUIVA19", "SUIVA19", 1, 2, "ASC")
      Call FINAL("*SUIVA19")
      '
      '\\\OT-08-0890-OD-FIN///

End Sub

Sub FACTURAPORVENDEDOR()
    '
'    OPT% = ALTERNA%("Emisión de Listado\Operaciones Sin Vendedor")
'    If OPT% < 1 Or OPT% > 2 Then Exit Sub
'    If OPT% = 2 Then
'       Call ACTUVENDEDOR
'       Exit Sub
'    End If

    Dim TOTALNETO#(255), TOTALBRUTO#(255)

20  Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    DESDAT = FETEXCOR1$(Des%)
    HASDAT = FETEXCOR1$(Has%)
    '
    Call ABRECABAN
    Call ABRESTOCKS
    Call ABREPEDCLI
    '
    SQLX$ = "SELECT CAJABANC.CODICOM_LAR, CAJABANC.NUCLIEN, CAJABANC.VEND_CLI, MOVISTO.NUPEDCLI "
    SQLX$ = SQLX$ + " FROM CAJABANC LEFT JOIN MOVISTO ON CODICOM_LAR=DOSECLAR WHERE "
    CONDXX$ = " CAJABANC.TIPOMOVIM = 'FVEN' AND CAJABANC.FECHEMISI >= '" & DESDAT & "'"
    CONDXX$ = CONDXX$ + " AND CAJABANC.FECHEMISI <= '" & HASDAT & "'"
    CONDXX$ = CONDXX$ + " AND (CAJABANC.VEND_CLI < 1 OR CAJABANC.VEND_CLI IS NULL)"
    '
    SQLX$ = SQLX$ + CONDXX$
    SQLX$ = SQLX$ + " AND MOVISTO.CODIMOVI <> 'TR'" ' ELIMINA LAS TRANSFERENCIAS DE MATERIALES DE MEDITEA
    '
    Screen.MousePointer = 11
    FIN& = CantRegistros("CAJABANC", CONDXX$)
    U& = 0
    Dim RSTMP As ADODB.Recordset
    Set RSTMP = New ADODB.Recordset
    RSTMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With RSTMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        NUFACAJAB$ = SAFETEXT(!codicom_lar) 'obtengo el numero de factura cajabanc
        NCLI = XVALO(!NUCLIEN)
        COND$ = " CODICOM_LAR = '" & NUFACAJAB$ & "'" 'condicion para
        NUPED& = XVALO(!NUPEDCLI)
        VENPED% = 0
        If NUPED& > 0 Then ' si tiene numero de pedido
           CONDIPE$ = "NROPED = " & NUPED& & ""
           VENPED% = XVALO(VALOBADA("PEDENCA", "VENDED", CONDIPE$, "NOMESS"))
        End If
        If VENPED% < 1 Then VENPED% = VENDCLI%(NCLI) ' SI NO TRAJO NUMERO DE VENDEDOR
        !VEND_CLI = VENPED%
        .Update
      .MoveNext
      Loop
    End With
    RSTMP.Close
    Set RSTMP = Nothing
    '
' **********************************************************************
'    CONDXX$ = "TIPOMOVIM = 'FVEN'"
'    CONDXX$ = CONDXX$ + " and FECHEMISI >= '" & DESDAT & "'"
'    CONDXX$ = CONDXX$ + " AND FECHEMISI <= '" & HASDAT & "'"
'    CONDXX$ = CONDXX$ + " AND (VEND_CLI < 1 OR VEND_CLI IS NULL)"
'    '
'    Screen.MousePointer = 11
'
'    FIN& = CantRegistros("CAJABANC", CONDXX$)
'    U& = 0
'    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE " & CONDXX$
'    Dim RSTMP As ADODB.Recordset
'    'Set RSTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    Set RSTMP = READSET(SQLX$)
'    With RSTMP
'      Do While Not .EOF
'        U& = U& + 1
'        Call TRACE(20, U&, FIN&)
'        NUFACAJAB$ = SAFETEXT(!CODICOM_LAR) 'obtengo el numero de factura cajabanc
'        CONDI$ = "DOSECLAR = '" & NUFACAJAB$ & "'"
'        NC% = XVALO(!NUCLIEN)
'        '
'        VENPED% = 0
'        COND$ = " CODICOM_LAR = '" & NUFACAJAB$ & "'" 'condicion para
'        'realizar el acturegistro
'        NUPED& = XVALO(VALOBADA("MOVISTO", "NUPEDCLI", CONDI$, "NOMESS"))
'        If NUPED& > 0 Then ' si tiene numero de pedido
'           CONDIPE$ = "NROPED = " & NUPED& & ""
'           VENPED% = XVALO(VALOBADA("PEDENCA", "VENDED", CONDIPE$, "NOMESS"))
'        End If
'        If VENPED% < 1 Then VENPED% = VENDCLI%(NC%) ' SI NO TRAJO NUMERO DE VENDEDOR
'        'TOMA EL DEL CLIENTE Y GRABA
'        Call ACTUREGISTRO("CAJABANC", "VEND_CLI", COND$, VENPED%, REGAF&, "NOMESS")
'        '
'      .MoveNext
'      Loop
'    End With
'    RSTMP.Close
'    Set RSTMP = Nothing
    '
' **********************************************************************
'
' ESTO SE COMENTA PORQUE SE APLICO CUANDO SE HABIAN RECUPERADO COMPROBANTES NO REGISTRADOS
' EN EL SISTEMA Y QUE SI ESTABAN EN LA AFIP
' **********************************************************************
'
'    'ESTE CODIGO ES PARA QUE LOS DOCUMENTOS RECUPERADOS QUEDEN AGRUPADOS EN
'    'UN VENDEDOR FICTICIO QUE NO AFECTE LAS COMISIONES DE LOS VENDEDORES REALES
'    CONDI2$ = "FECHEMISI <> FECHCONTA AND TIPOMOVIM='FVEN' AND VEND_CLI < 1 "
'    CONDI2$ = CONDI2$ + " AND FECHEMISI >= '" & DESDAT & "'"
'    CONDI2$ = CONDI2$ + " AND FECHEMISI <= '" & HASDAT & "'"
'    If CantRegistros&("CAJABANC", CONDI2$, "NOMESS") > 0 Then
'        Call SELECHO("VENDEDOR")
'        VEND$ = TRIM$(VALACT1$("VENDEDOR"))
'        If VEND$ <> "" Then
'            VENDERECU% = XVALO(VEND$)
'            condi3$ = "FECHEMISI <> FECHCONTA AND TIPOMOVIM ='FVEN' AND FECHEMISI >= '" & DESDAT & "' AND FECHEMISI <= '" & HASDAT & "'"
'            '
'            Call ACTUREGISTRO("CAJABANC", "VEND_CLI", condi3$, VENDERECU%, REG&, "NOMESS")
'            Call ACTUREGISTRO("CAJABANC", "VEND_CLI", condi3$, VENDERECU%, REG&, "NOMESS")
'        End If
'    End If
' **********************************************************************
'
    FactSinVended% = ACTUVENDEDOR%(DESDAT, HASDAT)
    If FactSinVended% > 0 Then ' mensaje informativo si quedan facturas sin vendedor
       Call COMUNI("Aún Quedan Comprobantes Sin Vendedor Asignado\Para el Rango de Fechas Elegido")
    End If
    '
    Screen.MousePointer = 11
    CONDI$ = " FECHEMISI >= '" & DESDAT & "' "
    CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HASDAT & "' "
    CONDI$ = CONDI$ + " AND TipoMovim = 'VENTA'"
    CONDI$ = CONDI$ + " AND ValAbsComp > 0"
    FIN& = CantRegistros("CAJABANC", CONDI$)
    ' OBTIENE LOS COMPROBANTES DE VENTAS EMITIDOS EN EL RANGO DE FECHAS
    SQLX$ = "SELECT CodiCom_Lar, FECHEMISI, ValAbsComp, "
    SQLX$ = SQLX$ + " NuClien, LetraComis,VEND_CLI,IHABECOMP,IDEBECOMP"
    SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where FECHEMISI >= '" & DESDAT & "' "
    SQLX$ = SQLX$ + " AND FECHEMISI <= '" & HASDAT & "' "
    SQLX$ = SQLX$ + " AND TipoMovim = 'VENTA'"
    SQLX$ = SQLX$ + " AND ValAbsComp > 0"
''    Dim CABATEMP As ADODB.Recordset
''    Set CABATEMP = New ADODB.Recordset
''25  On Error Resume Next
''    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
''    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''      On Error GoTo 0
''      Call CapturaErrorOpen
''      GoTo 25
''    End If
''    On Error GoTo 0
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If .EOF Or .BOF Then
       Call MENSERR(24, "No se Encuentran Comprobantes en el Rango Elegido")
       GoTo 20
      End If
      On Error GoTo 0
      '
      J& = 0
      U& = 0
      Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          NUMFACTU$ = SAFETEXT$(!codicom_lar)
          TOTCAJABANC# = XVALO(!ValAbsComp)
          FECHA% = CVINT(!FECHEMISI)
          FECFACT = FETEXCOR1$(FECHA%)
          NCLI = XVALO(!NUCLIEN)
          LETRA$ = SAFETEXT$(CABATEMP!LETRACOMIS)
          VENDE% = XVALO(!VEND_CLI)
          '
          ' OBTIENE LOS ARTICULOS QUE COMPONEN LAS FACTURAS
          SQLX$ = "SELECT SUM((CANTSALID * VALOUNIT) - (CANTINGRE * VALOUNIT)) AS TOTALMOVI FROM MOVISTO WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE DoSecLar = '" & NUMFACTU$ & "' AND CODIMOVI <> 'TR'"
          ' CORRIGE ERRORES DE COMIENZO
          Select Case NUMFACTU$
             Case "FC-0001-00009557-A"
                  SQLX$ = SQLX$ + " AND DoPriCor='RT.18471'"
             Case "FC-0001-00032555-B"
                  GoTo 50
             Case "FC-0003-00002199-B"
                  SQLX$ = SQLX$ + " AND NuPedCli=0"
          End Select
          '
'          SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
'          SQLX$ = SQLX$ + " WHERE DoSecLar = '" & NUMFACTU$ & "'"
'          ' CORRIGE ERRORES DE COMIENZO
'          Select Case NUMFACTU$
'             Case "FC-0001-00009557-A"
'                  SQLX$ = SQLX$ + " AND DoPriCor='RT.18471'"
'             Case "FC-0001-00032555-B"
'                  GoTo 50
'             Case "FC-0003-00002199-B"
'                  SQLX$ = SQLX$ + " AND NuPedCli=0"
'          End Select
          '
'          Dim MOVISTMP1 As ADODB.Recordset
'          Set MOVISTMP1 = New ADODB.Recordset
'26        On Error Resume Next
'          MOVISTMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockReadOnly, adCmdText
'          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'               On Error GoTo 0
'               Call CapturaErrorOpen
'               GoTo 26
'          End If
'          On Error GoTo 0
          '
          Dim MOVISTMP1 As ADODB.Recordset
          Set MOVISTMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
          '
          TOTMOVISTO# = 0
          With MOVISTMP1
            If Not (.EOF And .BOF) Then
'               .MoveFirst
'               Do While Not .EOF
'                 PREUNI# = XVALO(!VALOUNIT)
'                 CANTI# = Abs(XVALO(!CANTSALID) - XVALO(!CANTINGRE))
'                 TOTMOVISTO# = TOTMOVISTO# + (CANTI# * PREUNI#)
'                 .MoveNext
'               Loop
                TOTMOVISTO# = Abs(XVALO(!TOTALMOVI))
            End If
            '
            coefi# = 0
            If TOTMOVISTO# > 0 Then coefi# = TOTCAJABANC# / TOTMOVISTO#
            '
'            Dim MOVISTMP2 As ADODB.Recordset
'            Set MOVISTMP2 = New ADODB.Recordset
'27          On Error Resume Next
'            MOVISTMP2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockReadOnly, adCmdText
'            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'               On Error GoTo 0
'               Call CapturaErrorOpen
'               GoTo 27
'            End If
'            On Error GoTo 0
            '
            'IMPORTE BRUTO DEL DOCUMENTO
            CONDI$ = "TIPOMOVIM = 'FVEN' AND codicom_lar = '" & NUMFACTU$ & "'"
            IMPBRUTODE# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CONDI$, "NOMESS"))
            IMPBRUTOHA# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
            IMPBRUTO# = IMPBRUTODE# - IMPBRUTOHA#
            '
            'IMPORTE ABONADO
            FCHA$ = FETEXCOR1$(FECHA%)
            Condi1$ = "codicom_lar = '" & NUMFACTU$ & "' AND NUCLIEN = " & NCLI & " AND fechemisi = '" & FCHA$ & "'"
            SALDODEUDE# = XVALO(VALOBADA("SADEUPEN", "SUM(SALDDEBE)", Condi1$))
            SALDODEUHA# = XVALO(VALOBADA("SADEUPEN", "SUM(SALDACRE)", Condi1$))
            SALDODEU# = SALDODEUDE# - SALDODEUHA#
            RESUSALDODEU# = IMPBRUTO# - SALDODEU#
            '
            SQLX$ = "SELECT NUPEDCLI,COD_INTE,VALOUNIT,CANTSALID,CANTINGRE,DoSecLAR  FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE DoSecLar = '" & NUMFACTU$ & "' AND CODIMOVI <> 'TR' AND VALOUNIT > 0.00005"            ' CORRIGE ERRORES DE COMIENZO
            Select Case NUMFACTU$
               Case "FC-0001-00009557-A"
                    SQLX$ = SQLX$ + " AND DoPriCor='RT.18471'"
               Case "FC-0001-00032555-B"
                    GoTo 50
               Case "FC-0003-00002199-B"
                    SQLX$ = SQLX$ + " AND NuPedCli=0"
            End Select
            Dim MOVISTMP2 As ADODB.Recordset
            Set MOVISTMP2 = FLEXCONN.Execute(FILTSQL$(SQLX$))
            '
            With MOVISTMP2
              If Not (.EOF And .BOF) And coefi# > 0 Then
                .MoveFirst
                Do While Not .EOF
                   CI% = XVALO(!COD_INTE)
                   NUMPED& = XVALO(!NUPEDCLI)
                   PREUNI# = XVALO(!VALOUNIT)
                   CANTI# = XVALO(!Cantsalid) - XVALO(!CANTINGRE)
                   PRENETO# = Abs(PREUNI# * CANTI# * coefi#)
                   NUFACTMOVI$ = SAFETEXT(!DOSECLAR)
                      If Mid$(NUMFACTU$, 1, 2) = "NC" Then PRENETO# = (-1) * PRENETO#
'                   CONDIX1$ = "CODICOM_LAR = '" & NUFACTMOVI$ & "'"
'                   VENDE% = XVALO(VALOBADA("CAJABANC", "VEND_CLI", CONDIX1$))
                   '
                   X$ = REGBLAN$("FACPORVE")
                   Call REPLA(X$, Chr$(VENDE%), 1, 1)
                   Call REPLA(X$, MKI$(CI%), 2, 2)
                   Call REPLA(X$, MKS$(NCLI), 4, 4)
                   Call REPLA(X$, MKS$(NUMPED&), 8, 4)
                   Call REPLA(X$, NUMFACTU, 12, 18)
                   Call REPLA(X$, MKI$(FECHA%), 30, 2)
                   Call REPLA(X$, MKD$(PRENETO#), 32, 8)
                   Call REPLA(X$, LETRA$, 48, 8)
                   
                   Call REPLA(X$, MKD$(IMPBRUTO#), 56, 8)
                   Call REPLA(X$, MKD$(RESUSALDODEU#), 64, 8)
                   
                   Call REPLA(X$, (LOCACLI(NCLI)), 80, 16)
                   Call REPLA(X$, (PROVCLI(NCLI)), 96, 2)
                   
                   J& = J& + 1
                   Call GRAREG("FACPORVE", X$, J&)
                   'CALCULO DEL TOTAL NETO PARA LISTADO RESUMEN
                   If VENDE% < 256 Then
                     TOTALNETO#(VENDE%) = TOTALNETO#(VENDE%) + PRENETO#
                   End If
                   .MoveNext
                Loop
              Else
50                 X$ = REGBLAN$("FACPORVE")
                   'PARA QUE NO CAMBIE POR EL VENDEDOR POR DEFAULT CUANDO ES UN DOCUMENTO RECUPERADO
                   If VENDERECU% <> VENDE% Then VENDE% = VENDCLI%(NCLIE%)
                   Call REPLA(X$, Chr$(VENDE%), 1, 1)
                   Call REPLA(X$, MKI$(0), 2, 2)
                   Call REPLA(X$, MKS$(NCLI), 4, 4)
                   Call REPLA(X$, MKS$(0), 8, 4)
                   Call REPLA(X$, NUMFACTU, 12, 18)
                   Call REPLA(X$, MKI$(FECHA%), 30, 2)
                   PRENETO# = Abs(TOTCAJABANC#)
                     If Mid$(NUMFACTU$, 1, 2) = "NC" Then PRENETO# = (-1) * PRENETO#
                   Call REPLA(X$, MKD$(PRENETO#), 32, 8)
                   Call REPLA(X$, LETRA$, 48, 8)
                   Call REPLA(X$, MKD$(IMPBRUTO#), 56, 8)
                   Call REPLA(X$, MKD$(RESUSALDODEU#), 64, 8)
                   
                   Call REPLA(X$, (LOCACLI(NCLI)), 80, 16)
                   Call REPLA(X$, (PROVCLI(NCLI)), 96, 2)
                   
                   J& = J& + 1
                   If VENDE% < 256 Then
                     TOTALNETO#(VENDE%) = TOTALNETO#(VENDE%) + PRENETO#
                   End If
                   Call GRAREG("FACPORVE", X$, J&)
              End If
            End With
          End With
          If VENDE% < 256 Then
            TOTALBRUTO#(VENDE%) = TOTALBRUTO#(VENDE%) + IMPBRUTO#
          End If
          'CALCULO DEL TOTAL BRUTO PARA LISTADO RESUMEN
          'CONDI$ = "CODICOM_LAR = '" & NUMFACTU$ & "'AND TIPOMOVIM = 'FVEN' AND FECHEMISI = '" & FECFACT & "'"
          'IMDEB# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CONDI$, "NOMESS"))
          'IMHAB# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
          'If VENDE% < 256 Then
          '  TOTALBRUTO#(VENDE%) = TOTALBRUTO#(VENDE%) + (IMDEB# - IMHAB#)
          'End If
          '
          .MoveNext
      Loop
    End With
    On Error Resume Next
    CABATEMP.Close
    Set CABATEMP = Nothing
    MOVISTMP1.Close
    Set MOVISTMP1 = Nothing
    MOVISTMP2.Close
    Set MOVISTMP2 = Nothing
    On Error GoTo 0
    
    OPX% = ALTERNA%("  Detalle  \  Resumen  ")
    '
    If OPX% < 1 Or OPX% > 2 Then GoTo 20
    If OPX% = 2 Then  ' LISTADO RESUMEN
       'Call TOTALVENDEDOR
       i& = 0
       For VE% = 0 To 255
         If Abs(TOTALNETO#(VE%)) > 0 Then
            i& = i& + 1
            Call REPLA(ZZ$, Chr$(VE%), 1, 1)
            Call REPLA(ZZ$, MKD$(TOTALNETO#(VE%)), 2, 8)
            Call REPLA(ZZ$, MKD$(TOTALBRUTO#(VE%)), 10, 8)
            Call GRAREG("TOTVENDE", ZZ$, i&)
        End If
        '
      Next VE%
     Call SETULTREG("TOTVENDE", i&)
     '
     Call HEADERS("TOTVENDE", "")
     Call HEADERS("TOTVENDE", "Corresponde a: " + PERIO$)
     Call FILESORT("TOTVENDE", "", 1, 1, "ASC")
     Call FINAL("*TOTVENDE")
     Exit Sub
    End If
    '
    Call SETULTREG("FACPORVE", J&) ' LISTADO DETALLE
    Call FILESORT("FACPORVE", "", 1, 1, "ASC")
    Screen.MousePointer = 1
    '
    Call FINAL("*FACPORVE")
    '
End Sub

Private Sub mnuAbrirPeriodoActivo_Click()
Call Image1_Click
End Sub

Private Sub mnuAjustes_Click()
    Call Command7_Click
End Sub

Private Sub mnuAsientoVentas_Click()
    Call Command20_Click
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuAvanceFacturacion_Click()
    Call COGEVE07.Command9_Click
End Sub

Private Sub mnuCerrarPeriodoActivo_Click()
    Call Image2_Click
End Sub

Private Sub mnuContabilizacionVentasPorCuenta_Click()
    Call Command19_Click
End Sub

Private Sub mnuCuentaCorriente_Click()
    COGEVE07.Command7_Click
End Sub

Private Sub mnuCuentasPorCobrar_Click()
    Call COGEVE07.Command3_Click
End Sub

Private Sub mnuDetalleCobranza_Click()
    Call Command2_Click
End Sub

Private Sub mnuDetalleFacturacion_Click()
    Call Command15_Click
End Sub

Private Sub mnuFacturacion_Click()
    Call Command3_Click
End Sub

Private Sub mnuFlujoDeCajaProyectado_Click()
    Call COGEVE07.Command4_Click
End Sub

Private Sub mnuImpresionLibroIva_Click()
    Call Command14_Click
End Sub

Private Sub mnuLibroIvaVentas_Click()
    Option2.Value = True
    Call Command21_Click
End Sub

Private Sub mnuLimitesDeCredito_Click()
    Call COGEVE07.Command20_Click
End Sub

Private Sub mnuMaestroClientes_Click()
    Call Command9_Click
End Sub

Private Sub mnuNotasCredito_Click()
    Call Command5_Click
End Sub

Private Sub mnuNotasDebito_Click()
    Call Command4_Click
End Sub

Private Sub mnuRankingVentasArticulo_Click()
    Call Command11_Click
End Sub

Private Sub mnuRankingVentasCliente_Click()
    Call Command30_Click
End Sub

Private Sub mnuReimputacion_Click()
    Call Command22_Click
End Sub

Private Sub mnuReportesVarios_Click()
    Call AGenRep_Click
End Sub

Private Sub mnuResumenIvaVentas_Click()
    Option2.Value = True
    Call Command13_Click
End Sub

Private Sub mnuResumenMensualCuenta_Click()
    Call Command6_Click
End Sub

Private Sub mnuResumenProvinciaGrupo_Click()
    Call Command10_Click
End Sub

Private Sub mnuSaldosyDeudas_Click()
    Call COGEVE07.Command5_Click
End Sub

Private Sub mnuSubdiarioVentas_Click()
    Option2.Value = True
    Call Command16_Click
End Sub



Private Sub mnuVentasPorProvinciaCantidades_Click()

    Call VentasPorProvincia(1)
    
End Sub

Private Sub mnuVentasPorProvinciaFacturacion_Click()

    Call VentasPorProvincia(2)

End Sub

Private Sub mnuVentasProductosVendodores_Click()

    Call VentasProductosPorCanales

End Sub

Private Sub mnuVentasSoloVendedores_Click()

    Call VentasPorCanales(2)
    
End Sub

Private Sub mnuVentasTodosLosCanales_Click()

    Call VentasPorCanales(1)
    
End Sub

Private Sub mnuVentasVendodoresProductos_Click()

    Call VentasCanalesPorProducto

End Sub

Private Sub mnuVentasvsPedidos_Click()

    

End Sub

Private Sub RegInfoVentas_Click()

    Screen.MousePointer = 11
    If VALIMESAN% < 1 Then GoTo 99
    Call GENIVAVEN
    Screen.MousePointer = 1
   
    Call RegimenInformativoVentas

99  Screen.MousePointer = 1

End Sub


Sub LibroIvaVentas()

    Dim PuntoVentaOperador(32) As String
    Dim RSocialOperador(32) As String
    Dim CuitOperador(32) As String
    
    CAITAB1% = 1
    AgentePercepBSAS& = XVALO(CONTROL$("PERIBRU", "CUECONTA"))
    AgentePercepCABA& = XVALO(CONTROL$("PERIBCA", "CUECONTA"))
    AgentePercepMISIONES& = XVALO(CONTROL$("PERIBMI", "CUECONTA"))
    AgentePercepTUCUMAN& = XVALO(CONTROL$("PERIBTU", "CUECONTA"))
    ExportaSimple$ = TRIM$(CONTROL$("FACTURA", "EXSIMPLE"))
    Campos$ = "Fecha/D8;Documento/A18;Razón Social/A30;Cuit/A13;Cond. Iva/A14;Total/F13.2;Neto Gravado/F12.2;Iva 21%/F12.2;Iva 10.5%/F12.2;Fact N-G/F12.2; Deb N-G/F12.2;P. Iva/F8.2;P. Gcias/G0"
    If AgentePercepBSAS& > 0 Then Campos$ = Campos$ & ";IB BsAs/F9.2"
    If AgentePercepCABA& > 0 Then Campos$ = Campos$ & ";IB CABA/F9.2"
    If AgentePercepMISIONES& > 0 Then Campos$ = Campos$ & ";IB Mis/F9.2"
    If AgentePercepTUCUMAN& > 0 Then Campos$ = Campos$ & ";IB Tuc/F9.2"
    If ExportaSimple$ = "SI" Then
        Campos$ = Campos$ & ";Operador Exporta Simple/A20;Cuit Operador/A13"
        IDLIST& = ObtenerIdTablaAuxiliar("PTOVTAEXSIM")
        ConsultaSql$ = "select campo1, campo2 from datasql with(nolock) where id_lista=" & CStr(IDLIST&)
        Set RstPtoVentaExportaSimple = READSET(ConsultaSql$)
        IX% = 0
        With RstPtoVentaExportaSimple
            Do While Not .EOF
                IX% = IX% + 1
                PuntoVentaOperador(IX%) = TRIM$(!CAMPO1)
                NroProveedor& = (-1) * XVALO(!CAMPO2)
                RSocialOperador(IX%) = rasocli$(NroProveedor&)
                CuitOperador(IX%) = CuitCli$(NroProveedor&)
                
                .MoveNext
            Loop
        End With
        On Error Resume Next
        RstPtoVentaExportaSimple.Close
        Set RstPtoVentaExportaSimple = Nothing
        On Error GoTo 0
    End If
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    Screen.MousePointer = 11
    SUMA1# = 0: SUMA2# = 0: SUMA3# = 0
    SUMA4# = 0: SUMA5# = 0: SUMA6# = 0: SUMA7# = 0
    URE& = ULTREG&("LIVAVEN")
    CAITAB1% = 1
    For YX& = 1 To URE&
       '
       YY$ = REGLEIDO$("LIVAVEN", YX&)
       FEMI = CVI(Mid$(YY$, 3, 2))
       NDOC$ = Mid$(YY$, 11, 18)
       Tipo$ = Left(NDOC$, 2) + Mid$(NDOC$, 18, 1)
       NUME$ = NDOC$
       RACLI$ = Mid$(YY$, 33, 30)
       CUITI$ = TRIM$(Mid$(YY$, 63, 16))
       PIVAI$ = TRIM$(Mid$(YY$, 79, 12))
       TOFA0# = CVD(Mid$(YY$, 91, 8))
       PercepcionesIngresosBrutos# = CVD(Mid$(YY$, 193, 8))
       Jurisdiccion$ = TRIM$(Mid$(YY$, 203, 1))
       '
       NEGRA0# = 0: NEXEN0# = 0: NEXEN1# = 0: NIVA21# = 0: NIVA105# = 0: RETIVA# = 0: RETGAN# = 0: PercepcionIB_BSAS# = 0: PercepcionIB_CABA# = 0: PercepcionIB_MISIONES# = 0: PercepcionIB_TUCUMAN# = 0
       '
60     NETO1# = CVD(Mid$(YY$, 137, 8)) + CVD(Mid$(YY$, 185, 8))
       iva1# = CVD(Mid$(YY$, 145, 8))
       TASA1 = CVD(Mid$(YY$, 153, 8))
       If Abs(iva1#) < 0.005 Then
            If Mid$(Tipo$, 1, 2) = "FC" Then
                NEXEN0# = NEXEN0# + NETO1#
            Else
                NEXEN1# = NEXEN1# + NETO1#
            End If
          Else
            NEGRA0# = NEGRA0# + NETO1#
            If TASA1 >= 10 And TASA1 <= 11 Then
                 NIVA105# = NIVA105# + iva1#
               Else
                 NIVA21# = NIVA21# + iva1#
            End If
       End If
       RETIVA# = RETIVA# + CVD(Mid$(YY$, 161, 8))
       RETGAN# = RETGAN# + CVD(Mid$(YY$, 169, 8))
       '
       If Abs(PercepcionesIngresosBrutos#) > 0.005 Then
            CondicionSql$ = "codicom_lar='" & NUME$ & "' and fechconta='" & FETEXCOR1$(CInt(FEMI)) & "' and nuclien>0 and tipomovim='ibru' and varimovim='perc'"
            
            CondicionSql0$ = CondicionSql$ & " and opcimovim='ib-b'"
            PercepcionIB_BSAS# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql0$)) * Sgn(PercepcionesIngresosBrutos#)
            CondicionSql0$ = CondicionSql$ & " and opcimovim='ib-c'"
            PercepcionIB_CABA# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql0$)) * Sgn(PercepcionesIngresosBrutos#)
            CondicionSql0$ = CondicionSql$ & " and opcimovim='ib-mi'"
            PercepcionIB_MISIONES# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql0$)) * Sgn(PercepcionesIngresosBrutos#)
            CondicionSql0$ = CondicionSql$ & " and opcimovim='ib-tu'"
            PercepcionIB_TUCUMAN# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql0$)) * Sgn(PercepcionesIngresosBrutos#)
       End If
       '
       If YX& < URE& Then
         YZ$ = REGLEIDO$("LIVAVEN", YX& + 1)
         If TRIM$(Mid$(YZ$, 11, 18)) = "" Then
           YX& = YX& + 1
           YY$ = YZ$
           GoTo 60
         End If
       End If
       '
       CAITAB1% = CAITAB1% + 1
       GRILLAGRAL.GRID.Rows = CAITAB1% + 1
       
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 1) = FECHATEX$(FEMI)
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 2) = NUME$
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 3) = RACLI$
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = CUITI$
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 5) = PIVAI$
       
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(TOFA0#, "F13.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 7) = FORMATNUM$(NEGRA0#, "F12.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 8) = FORMATNUM$(NIVA21#, "F12.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 9) = FORMATNUM$(NIVA105#, "F12.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 10) = FORMATNUM$(NEXEN0#, "F12.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 11) = FORMATNUM$(NEXEN1#, "F12.2")
       
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 12) = FORMATNUM$(RETIVA#, "F8.2")
       GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 13) = FORMATNUM$(RETGAN#, "F11.2")
       
       ColumnaPercep% = 13
       If AgentePercepBSAS& > 0 Then
            ColumnaPercep% = ColumnaPercep% + 1
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(PercepcionIB_BSAS#, "F9.2")
       End If
       If AgentePercepCABA& > 0 Then
            ColumnaPercep% = ColumnaPercep% + 1
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(PercepcionIB_CABA#, "F9.2")
       End If
       If AgentePercepMISIONES& > 0 Then
            ColumnaPercep% = ColumnaPercep% + 1
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(PercepcionIB_MISIONES#, "F9.2")
       End If
       If AgentePercepTUCUMAN& > 0 Then
            ColumnaPercep% = ColumnaPercep% + 1
            GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(PercepcionIB_TUCUMAN#, "F9.2")
       End If
       If ExportaSimple$ = "SI" And Mid$(TRIM$(NDOC$), 18, 1) = "E" Then
            If CantRegistros("CAJABANC", "tipomovim='fven' and codicom_lar_ExSimple='" & TRIM$(NDOC$) & "'") > 0 Then
                PtoVenta$ = SAFETEXT$(VALOBADA("cajabanc", "numetalo", "tipomovim='fven' and codicom_lar_ExSimple='" & TRIM$(NDOC$) & "'"))
                Encontro% = 0
                For IX% = 1 To 32
                    If XVALO(PuntoVentaOperador(IX%)) = XVALO(PtoVenta$) Then
                        Encontro% = IX%: Exit For
                    End If
                Next IX%
                If Encontro% > 0 Then
                    ColumnaPercep% = ColumnaPercep% + 1
                    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = RSocialOperador(Encontro%)
                    ColumnaPercep% = ColumnaPercep% + 1
                    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = CuitOperador(Encontro%)
                End If
            Else
                ColumnaPercep% = ColumnaPercep% + 1
                GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = ""
                ColumnaPercep% = ColumnaPercep% + 1
                GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = ""
            End If
       End If
       
       SUMA1# = SUMA1# + ROUND#(NEGRA0#)
       SUMA2# = SUMA2# + ROUND#(NEXEN0#)
       SUMA3# = SUMA3# + ROUND#(NIVA21#)
       SUMA4# = SUMA4# + ROUND#(NIVA105#)
       SUMA5# = SUMA5# + ROUND#(TOFA0#)
       SUMA6# = SUMA6# + ROUND#(RETIVA#)
       SUMA7# = SUMA7# + ROUND#(RETGAN#)
       SUMA8# = SUMA8# + ROUND#(NEXEN1#)
       SUMA9# = SUMA9# + ROUND#(PercepcionIB_BSAS#)
       SUMA10# = SUMA10# + ROUND#(PercepcionIB_CABA#)
       SUMA11# = SUMA11# + ROUND#(PercepcionIB_MISIONES#)
       SUMA12# = SUMA12# + ROUND#(PercepcionIB_TUCUMAN#)
    Next YX&
    
    CAITAB1% = CAITAB1% + 2
    GRILLAGRAL.GRID.Rows = CAITAB1% + 1
       
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 4) = "TOTALES"
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 6) = FORMATNUM$(SUMA5#, "F13.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 7) = FORMATNUM$(SUMA1#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 8) = FORMATNUM$(SUMA3#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 9) = FORMATNUM$(SUMA4#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 10) = FORMATNUM$(SUMA2#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 11) = FORMATNUM$(SUMA8#, "F12.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 12) = FORMATNUM$(SUMA6#, "F8.2")
    GRILLAGRAL.GRID.TextMatrix(CAITAB1%, 13) = FORMATNUM$(SUMA7#, "F11.2")
    
    ColumnaPercep% = 13
    If AgentePercepBSAS& > 0 Then
        ColumnaPercep% = ColumnaPercep% + 1
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(SUMA9#, "F9.2")
    End If
    If AgentePercepCABA& > 0 Then
        ColumnaPercep% = ColumnaPercep% + 1
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(SUMA10#, "F9.2")
    End If
    If AgentePercepMISIONES& > 0 Then
        ColumnaPercep% = ColumnaPercep% + 1
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(SUMA11#, "F9.2")
    End If
    If AgentePercepTUCUMAN& > 0 Then
        ColumnaPercep% = ColumnaPercep% + 1
        GRILLAGRAL.GRID.TextMatrix(CAITAB1%, ColumnaPercep%) = FORMATNUM$(SUMA12#, "F9.2")
    End If
    
    Call COLOREAR_GRILLA_SOLAFILA(GRILLAGRAL.GRID, &H80C0FF, CAITAB1%, , 1)
    
    Screen.MousePointer = 1
    TX$ = "Libro de Iva Ventas " & MESAN.TEXT
    GRILLAGRAL.Caption = TX$
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    GRILLAGRAL.Show 1
    Dim MSF As MSFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "LIVAVEN", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    Set obj = Nothing
       
End Sub

Sub Command21_Click()
    ''
    Command21.Enabled = False
    If InStr(1, EMPREAC$, "TANIT") > 0 Then
        If DERACCE%("LIVAVEN", "LIBRO IVA VENTAS") < 1 Then
            GoTo 99
        End If
    End If
    
    If VALIMESAN% < 1 Then GoTo 99
    '
    RESP1% = ALTERNA%("Selección de Comprobantes a Listar.\Listado Libro de IVA - Ventas.")
    If RESP1% = 1 Then
        Call EXCLUIVA
      ElseIf RESP1% = 2 Then
        Option2.Value = True
        'INCLUIDO PARA PODER LISTAR EL LIBRO IVA POR FECHA CONTABLE POR LOS DOCUMENTOS RECUPERADOS
        MODO% = 0
        If InStr(EMPREAC$, "MEDITEA") > 0 Then
           If COMALTER("Elija Formato de Salida:\\Fecha Emision\Fecha Contable") = 2 Then MODO% = 1
        End If
        '
        'Call AnalizaCorrelativaDeComprobantes(1)
        '
        Call GENIVAVEN(MODO%)
        '\\\OT-08-0890-OD-30/09/08///
        If InStr(EMPREAC$, "DOSIVA") > 0 Then
           Call IVADOSI
           GoTo 99
        End If
        '\\\OT-08-0890-OD-FIN///
        If EXISTE%("LIVENTAN.PRF") > 0 Then
             'INCLUIDO PARA PODER LISTAR EL LIBRO IVA POR FECHA CONTABLE POR LOS DOCUMENTOS RECUPERADOS
             LIBROIVA$ = "*LIVENTAN"
             If MODO% = 1 Then
                Call FILESORT("LIVENTAN", "", 2, 2, "ASC")
                LIBROIVA$ = "*LIVENMED"
             End If
             'If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
             '   Call FINAL(LIBROIVA$)
             'Else
                Call LibroIvaVentas
             'End If
           Else
             FORIPRE$ = TRIM$(CONTROL$("FORIPRE", "LIVAVEN"))
             If FORIPRE$ <> "" Then
               If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
                 If COMALTER("Elija Formato de Salida:\\Planilla Plana\Formulario") = 2 Then
                   Call FORMIVAVE
                   GoTo 99
                 End If
               End If
             End If
             'Call FINAL("*LIVENMEN")
             Call LibroIvaVentas
        End If
    End If
    '
99  Command21.Enabled = True
End Sub


Sub TablasAuxiliares()

    Dim TablaAuxiliar As New DatosAuxiliares
    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    ' ESTADISTICAS - TABLA PRODUCTOS
    ' ******************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A20"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;descripcion;master;status>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("ESTADIPRODUC", "", ListaCampos)
    
    
    ' ESTADISTICAS - TABLA VENDEDORES
    ' *******************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A6"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;denominacion;#TPvendedores;codigo>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominación/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("ESTADIVENDE", "", ListaCampos)
    
    ' ESTADISTICAS - TABLA CUENTAS DE EGRESOS
    ' ***************************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Cuenta/A24"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codexcue;denocue;plancuen;statcue<9 and posimpu=1;order by codexcue"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominación/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("CUECONEGRE", "", ListaCampos)
    
    ' ESTADISTICAS - TABLA AGRUPACION DE MEDIOS DE COBRANZA
    ' *****************************************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Medio/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "M-C/I3"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;denominacion;#TPMedCobranza;codigo>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominacion/A20"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Agrupa por/A20"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C.Contable/A16"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codexcue;denocue;plancuen;statcue<9 and posimpu=1;order by codexcue"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominación/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call TablaAuxiliar.CrearTabla("GRUPOMEDCOB", "", ListaCampos)

End Sub
