VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FOREMIT07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Materiales "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11895
   Begin VB.Frame Frame18 
      BackColor       =   &H00CDDADA&
      Caption         =   "IMPRESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9300
      TabIndex        =   87
      Top             =   5985
      Width           =   1455
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   88
         Text            =   " "
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copias"
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
         TabIndex        =   89
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.CommandButton Command12 
      Enabled         =   0   'False
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
      Left            =   12915
      Picture         =   "FOREMIT07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Descompone Conjuntos en Partes Componentes"
      Top             =   420
      Width           =   650
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      Begin MSFlexGridLib.MSFlexGrid MSFDETAREMI 
         Height          =   3285
         Left            =   240
         TabIndex        =   101
         Top             =   1680
         Visible         =   0   'False
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   5794
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Aguarde un Momento que se termine de actualizar la base de datos Waldbott > Flexoft"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   100
         Top             =   2880
         Visible         =   0   'False
         Width           =   11895
      End
      Begin VB.CommandButton Command19 
         Caption         =   "!"
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
         Left            =   1250
         TabIndex        =   84
         Top             =   685
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.CommandButton Command18 
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
         Left            =   10500
         TabIndex        =   81
         Top             =   1320
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.CommandButton Command17 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9000
         TabIndex        =   80
         ToolTipText     =   "Ingreso Manual de Prefijo de Remito"
         Top             =   550
         Width           =   180
      End
      Begin VB.PictureBox Picture999 
         Height          =   330
         Index           =   0
         Left            =   1470
         ScaleHeight     =   270
         ScaleWidth      =   1110
         TabIndex        =   78
         Top             =   525
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton Command15 
         Caption         =   "T L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   7875
         TabIndex        =   76
         ToolTipText     =   "Remito de Texto Libre"
         Top             =   165
         Width           =   405
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1890
         Top             =   0
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00CDDADA&
         Caption         =   "ITEMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   9680
         TabIndex        =   73
         Top             =   5280
         Width           =   1080
         Begin VB.Label Label18 
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
            Height          =   270
            Left            =   120
            TabIndex        =   74
            Top             =   300
            Width           =   840
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00CDDADA&
         Caption         =   "ORDEN DE SUMINISTRO (INTERNO)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6600
         TabIndex        =   71
         Top             =   5280
         Width           =   2955
         Begin VB.TextBox Text22 
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
            Height          =   330
            Left            =   120
            MaxLength       =   24
            TabIndex        =   42
            Top             =   260
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   9045
         Left            =   11025
         ScaleHeight     =   9015
         ScaleWidth      =   3240
         TabIndex        =   58
         Top             =   -105
         Width           =   3270
         Begin VB.FileListBox File1 
            Height          =   285
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ListBox List3 
            Height          =   255
            ItemData        =   "FOREMIT07.frx":030A
            Left            =   0
            List            =   "FOREMIT07.frx":030C
            Sorted          =   -1  'True
            TabIndex        =   96
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Fact."
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":030E
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Acceso Directo a Facturación"
            Top             =   4420
            Width           =   650
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Code"
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":0750
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Generación por Código de Barras"
            Top             =   5120
            Width           =   650
         End
         Begin VB.CommandButton Command16 
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
            Height          =   690
            Left            =   105
            Picture         =   "FOREMIT07.frx":0A5A
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Acceso a Toyota"
            Top             =   4420
            Width           =   640
         End
         Begin VB.CommandButton Command4 
            Caption         =   "CCorr"
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":1924
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2350
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Setup"
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":1C2E
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Configuracion de Funcionamiento"
            Top             =   6550
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":2070
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":21BA
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Continúa Carga de Remito."
            Top             =   890
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "O.K."
            Enabled         =   0   'False
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":24C4
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Aprueba - Graba e Imprime Remito"
            Top             =   7245
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":2B2E
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Opciones Complementarias - Listados y Reportes"
            Top             =   5860
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Clien."
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":2C78
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   1660
            Width           =   650
         End
         Begin VB.CommandButton Command3 
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":30BA
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Notas de Clientes"
            Top             =   3040
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Kdx."
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
            Left            =   120
            Picture         =   "FOREMIT07.frx":33C4
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Ficha de Movimientos de Material"
            Top             =   3730
            Width           =   650
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   10425
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   " Descripción                                                                 Cantidad  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   0
            TabIndex        =   57
            Top             =   120
            Width           =   10410
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00CDDADA&
         Caption         =   "VARIANTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5760
         TabIndex        =   48
         Top             =   120
         Width           =   1680
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Consignación"
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
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   430
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Facturación"
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
            Height          =   225
            Left            =   120
            TabIndex        =   49
            Top             =   220
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "VALOR DECLARADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   27
         Top             =   5985
         Width           =   2595
         Begin VB.TextBox Text17 
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
            Left            =   120
            TabIndex        =   43
            Top             =   250
            Width           =   2370
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "OBSERVACIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1660
         Left            =   6600
         TabIndex        =   26
         Top             =   6615
         Width           =   4155
         Begin VB.TextBox Text16 
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
            Index           =   4
            Left            =   210
            TabIndex        =   55
            Top             =   2500
            Width           =   3375
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   1260
            Width           =   3940
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   945
            Width           =   3940
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   630
            Width           =   3940
         End
         Begin VB.TextBox Text16 
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   315
            Width           =   3940
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   240
         TabIndex        =   24
         Top             =   5040
         Width           =   6285
         Begin VB.Frame Frame14 
            BackColor       =   &H00CDDADA&
            Caption         =   "CONDUCTOR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1530
            Left            =   120
            TabIndex        =   51
            Top             =   1680
            Width           =   2325
            Begin VB.CommandButton Command20 
               Caption         =   "!"
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
               Left            =   120
               TabIndex        =   92
               ToolTipText     =   "Boton recuperacion remito a wbt Envaplast se activa desde el menu"
               Top             =   1200
               Width           =   540
            End
            Begin VB.TextBox Text19 
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
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   2115
            End
            Begin VB.TextBox Text13 
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
               Left            =   630
               TabIndex        =   36
               Top             =   855
               Width           =   1590
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DNI"
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
               TabIndex        =   54
               Top             =   945
               Width           =   435
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00CDDADA&
            Caption         =   "ORDEN DE COMPRA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Width           =   2325
            Begin VB.TextBox Text20 
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
               Left            =   315
               TabIndex        =   33
               Top             =   315
               Width           =   1695
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00CDDADA&
            Caption         =   "NUMERO DE BULTOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   120
            TabIndex        =   31
            Top             =   945
            Width           =   2325
            Begin VB.TextBox Text15 
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
               Left            =   525
               TabIndex        =   34
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00CDDADA&
            Caption         =   "TRANSPORTISTA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3000
            Left            =   2520
            TabIndex        =   25
            Top             =   210
            Width           =   3690
            Begin VB.Frame Frame5 
               BackColor       =   &H00CDDADA&
               Caption         =   "DOMICILIO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1485
               Left            =   120
               TabIndex        =   52
               Top             =   1365
               Width           =   3480
               Begin VB.TextBox TXTCODPOSTAL 
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
                  Height          =   285
                  Left            =   1050
                  TabIndex        =   91
                  Top             =   1080
                  Width           =   2340
               End
               Begin VB.TextBox Text21 
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
                  Left            =   120
                  TabIndex        =   41
                  Top             =   700
                  Width           =   3270
               End
               Begin VB.TextBox Text14 
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
                  Left            =   120
                  TabIndex        =   40
                  Top             =   315
                  Width           =   3270
               End
               Begin VB.Label Label14 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Cód.Postal:"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   280
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1120
                  Width           =   855
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00CDDADA&
               Caption         =   "PATENTE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   1995
               TabIndex        =   30
               Top             =   735
               Width           =   1590
               Begin VB.TextBox Text18 
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
                  Left            =   105
                  TabIndex        =   39
                  Top             =   210
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00CDDADA&
               Caption         =   "C.U.I.T."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   105
               TabIndex        =   29
               Top             =   735
               Width           =   1800
               Begin VB.TextBox Text12 
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
                  Left            =   105
                  TabIndex        =   38
                  Top             =   210
                  Width           =   1590
               End
            End
            Begin VB.TextBox Text8 
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
               Left            =   210
               TabIndex        =   37
               Top             =   315
               Width           =   3270
            End
         End
      End
      Begin VB.TextBox Text7 
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
         Left            =   9345
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   915
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5280
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox Text11 
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
         Left            =   9180
         TabIndex        =   4
         Text            =   " "
         Top             =   550
         Width           =   1500
      End
      Begin VB.TextBox Text10 
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
         Left            =   8715
         TabIndex        =   6
         Text            =   " "
         Top             =   915
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9030
         TabIndex        =   5
         Top             =   915
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   2
         Text            =   " "
         Top             =   195
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10500
         TabIndex        =   3
         Top             =   195
         Width           =   180
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
         Left            =   2280
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1080
         Width           =   2890
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
         Left            =   1470
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1080
         Width           =   750
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
         Left            =   1470
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   685
         Width           =   3690
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
         Left            =   1470
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   685
         Width           =   180
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
         Left            =   240
         TabIndex        =   0
         Text            =   " "
         Top             =   685
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   240
         ScaleHeight     =   3225
         ScaleWidth      =   10485
         TabIndex        =   18
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   10515
            Begin VB.Line Line7 
               X1              =   6300
               X2              =   6300
               Y1              =   0
               Y2              =   1260
            End
            Begin VB.Line Line6 
               X1              =   9345
               X2              =   9345
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line5 
               X1              =   7665
               X2              =   7665
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line4 
               X1              =   1785
               X2              =   1785
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FOREMIT07.frx":36CE
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000B&
               Height          =   330
               Left            =   0
               TabIndex        =   21
               Top             =   135
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Frame3"
            Height          =   5500
            Left            =   0
            TabIndex        =   19
            Top             =   -1080
            Width           =   10515
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1.95900e5
               Left            =   0
               TabIndex        =   28
               Top             =   1560
               Visible         =   0   'False
               Width           =   10515
            End
            Begin VB.Line Line10 
               X1              =   0
               X2              =   10500
               Y1              =   1530
               Y2              =   1530
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   5670
               X2              =   5670
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   4
               X1              =   1785
               X2              =   1785
               Y1              =   105
               Y2              =   5145
            End
            Begin VB.Line Line2 
               Index           =   3
               X1              =   9345
               X2              =   9345
               Y1              =   630
               Y2              =   5670
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7665
               X2              =   7665
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   6300
               X2              =   6300
               Y1              =   1260
               Y2              =   6300
            End
         End
      End
      Begin VB.PictureBox Skin1 
         Height          =   480
         Left            =   2625
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   93
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   79
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label TEPRUEBA 
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
         Height          =   180
         Left            =   0
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mon.Ped.:"
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
         Index           =   3
         Left            =   5200
         TabIndex        =   99
         ToolTipText     =   "Tipo de Moneda del Pedido"
         Top             =   1440
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   225
         Left            =   5880
         TabIndex        =   98
         ToolTipText     =   "Tipo de Moneda del Pedido"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4260
         TabIndex        =   95
         Top             =   1440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización:"
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
         Index           =   2
         Left            =   3360
         TabIndex        =   94
         Top             =   1440
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Seleccionado"
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
         Left            =   270
         TabIndex        =   86
         Top             =   1440
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblNroPed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1980
         TabIndex        =   85
         Top             =   1425
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén"
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
         Left            =   7200
         TabIndex        =   83
         Top             =   1420
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label11 
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
         Left            =   8600
         TabIndex        =   82
         Top             =   1320
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label Label10 
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
         Height          =   300
         Left            =   8340
         TabIndex        =   75
         ToolTipText     =   "Doble Click  para Cambiar el Punto de Venta"
         Top             =   550
         Width           =   660
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
         Left            =   240
         TabIndex        =   72
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label NORVENTA 
         Height          =   225
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DOCUMENTO NO VALIDO COMO FACTURA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5355
         TabIndex        =   23
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   975
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
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
         Left            =   7350
         TabIndex        =   16
         Top             =   1030
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   8805
         TabIndex        =   15
         Top             =   300
         Width           =   750
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
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5400
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   580
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuArmadoListaDespacho 
         Caption         =   "Armado de Lista de Despacho"
      End
      Begin VB.Menu mnuAnulaListaDeDespacho 
         Caption         =   "Anular Lista de Despacho"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
      Begin VB.Menu mnuEntregasNeuquenAvellaneda 
         Caption         =   "Entregas Neuquén y Avellaneda"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEntregaCPedidoCompra 
         Caption         =   "Entregas Neuquén Con N.Pedido y O.Compra"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEntregaTelas 
         Caption         =   "Reporte de Telas Entregadas x rango de fechas"
      End
      Begin VB.Menu mnuEntxFechaMetrosyKilos 
         Caption         =   "Entregas X Rango de Fecha"
      End
      Begin VB.Menu mnuEntregasTotalKG 
         Caption         =   "Entregas C/Total kg (Envaplast)"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuReimpRemitoQR 
         Caption         =   "Re-Imprimir Remito con QR p/Conteplast"
         Visible         =   0   'False
      End
      Begin VB.Menu linea 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuCot 
      Caption         =   "Remito Electrónico"
      Begin VB.Menu mnuGenCot 
         Caption         =   "Selección de Remitos para Generar C.O.T."
      End
      Begin VB.Menu mnuGenRemDev 
         Caption         =   "Generar Remito de Devolución"
      End
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuRecuperarxRemito 
         Caption         =   "Recuperar Remito de Flexoft hacia Waldbot"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu RemContImp 
         Caption         =   "Remitos Conteplast - Impresion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTransfRemitoConteplastaWaldbot 
         Caption         =   "Transferir Remitos de CONTEPLAST a Waldbott"
         Visible         =   0   'False
         Begin VB.Menu DesRemCTP 
            Caption         =   "Descargar Remitos de Conteplast"
            Visible         =   0   'False
         End
         Begin VB.Menu RemCTPAWDBT 
            Caption         =   "Remitos de Conteplast a Pasar a Waldbott"
         End
      End
      Begin VB.Menu mnuRegDesp 
         Caption         =   "Registro de Despachos x Código Barra"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuVincularReportesConFormulario 
         Caption         =   "Vincular Reportes con Formulario"
      End
      Begin VB.Menu lineaConfiguracion_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiguracionGeneral 
         Caption         =   "Configuración General"
      End
   End
End
Attribute VB_Name = "FOREMIT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TOTITFO& 'CANTIDAD DE ITEMS DEL FORMULARIO DE IMPRESIÓN
Public FORMUL As Form
Public PORBARRA%
Private Declare Function GetFileAttributes Lib "KERNEL32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Private Const INVALID_FILE_ATTRIBUTES = &HFFFFFFFF
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Sub AgrupaRemitoConteplast(AgrupaLote%)
    
    'SI AgrupaLote% > 0 AGRUPA POR CODIGO Y LOTE
    'SI AgrupaLote% = 0 AGRUPA SOLO POR CODIGO
    '
    '**** OJO ******
    
    'SI SE UTILIZA LA RUTINA 'DETLOTREMCAMBA' QUE AGREGA EL DETALLE DE LOS ITEMS QUE AQUI SE AGRUPA
    'AgrupaLote% DEBE SER = 0 PARA QUE EL DETALLE SE MUESTRE CORRECTAMENTE.
    'SI AgrupaLote% FUERA MAYOR A 0 A NIVEL DE IMPRESION SE VAN A REPETIR LOS DETALLES
    
    ABRIOBEGIN% = 0
    Call Creatabla_REMITEMP
    For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI1% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        LOTE$ = Mid$(YY$, 107, 12)
        NPEDI& = CVS(Mid$(YY$, 151, 4))
        NUORITPE% = CVS(Mid$(YY$, 189, 4))
        SQLX$ = "INSERT INTO REMITEMP  "
        SQLX$ = SQLX$ + "(CantSalid,Cod_Inte,NuorItem,IdenLote,NuPedCli)"
        SQLX$ = SQLX$ + " VALUES("
        
        SQLX$ = SQLX$ + "'" & TRIM$(FORMATNUM$(CAN#, "f12.1")) & "'"
        SQLX$ = SQLX$ + "," & CI1%
        SQLX$ = SQLX$ + "," & NUORITPE%
        SQLX$ = SQLX$ + ",'" & LOTE$ & "'"
        SQLX$ = SQLX$ + "," & NPEDI&
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
        '
    Next YX&


        JJ& = 0
        SQLX$ = "SELECT SUM(CantSalid) AS SALIDA,Cod_Inte,NuPedCli"
         If AgrupaLote% > 0 Then SQLX$ = SQLX$ + ",IdenLote "
        SQLX$ = SQLX$ + " FROM REMITEMP WITH(NOLOCK)"
        SQLX$ = SQLX$ + " Group BY Cod_Inte,NuPedCli"
        If AgrupaLote% > 0 Then SQLX$ = SQLX$ + ",IdenLote "

        Set rs = READSET(SQLX$)

        With rs
          Do While Not .EOF
            CI2% = XVALO(!cod_inte)
            Q1# = XVALO(!SALIDA)
            LOTEX$ = " "
            If AgrupaLote% > 0 Then LOTEX$ = SAFETEXT$(!idenlote)
            NPED& = XVALO(!NUPEDCLI)
            Y$ = REBLA$
           Call REPLA(Y$, MKI$(CI2%), 83, 2)
           Call REPLA(Y$, Left$(Unimed$(CI2%), 4), 85, 4)
           Call REPLA(Y$, MKD$(Q1#), 89, 8)
           Call REPLA(Y$, LOTEX$, 107, 12)
           Call REPLA(Y$, MKS$(NPED&), 151, 4)
           JJ& = JJ& + 1
           Call GRAREG("!BOLREDET", Y$, JJ&)
            .MoveNext
          Loop
        End With
        Call SETULTREG("!BOLREDET", JJ&)

        rs.Close
        SQLX$ = "DELETE FROM REMITEMP"
        FLEXCONN.Execute (SQLX$)


    End Sub
Function AltaClienteCambaceres&()
  AltaClienteCambaceres& = NC&
  If Control("BASEWALD", "ESTAGRAB") = "SI" Then
      Call COMUNI("Base de Datos en Actualización\Aguarde un Momento y Vuelva a Intentarlo")
      Exit Function
   End If
   
20 NCLIENTE$ = TRIM$(InputBox$("Ingrese N° de Cliente (Sistema de Envaplast) ", "Alta Cliente", NCX$))
   If NCLIENTE$ = "" Then Exit Function
   If XVALO(NCLIENTE$) < 1 Then
      Call COMUNI("INGRESO NO VÁLIDO")
      NCX$ = NCLIENTE$
      GoTo 20
   End If
   '
   NC& = XVALO(NCLIENTE$)
   Dim ENV As New Envap
   Call ENV.ALTA_CLIENTE_WALDBOT(NC&, 1, RET_OK%)
   If RET_OK% >= 0 Then
       AltaClienteCambaceres& = NC&
   End If

End Function

Sub BLANFORMU()

End Sub
Sub Creatabla_REMITEMP()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
       Call CREACAMPO("STOCKS", "REMITEMP", "IdenLote", 10, 16, 1)
       If ValidaExisteCampo%("REMITEMP_ID", "REMITEMP") < 1 Then
          SQLX$ = "ALTER TABLE REMITEMP ADD REMITEMP_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
       End If
       
       Call CREACAMPO("STOCKS", "REMITEMP", "FeReMovi", 8, 0)
       Call CREACAMPO("STOCKS", "REMITEMP", "CantSalid", 7, 0)
       Call CREACAMPO("STOCKS", "REMITEMP", "Cod_Inte", 3, 0, 1)
       Call CREACAMPO("STOCKS", "REMITEMP", "NuorItem", 3, 2)
       Call CREACAMPO("STOCKS", "REMITEMP", "NumeOrdPed", 3, 2)
       Call CREACAMPO("STOCKS", "REMITEMP", "NuPedCli", 4, 4, 1)
   End If
End Sub


Sub DETLOTREMCAMBA(CI%, NROREMIT$, NPEDI&)
        'AGREGA EL DETALLE DE LO REMITIDO POR QUE EL MISMO SE AGRUPA POR CODIGO
        'Y ESTA FUNCION GENERA QUE SE IMPRIMA IDENTADO EL DETALLE DE LOS METROS
        'EJ,
        '10 CODIGO1 DESCRIPCIONCODIGO1
           'OF-000123-1 1000 mts
           'OF-000124-1 1000 mts
           'OF-000125-1 1000 mts
           '.............
        '

        
        Dim CTP As New Conteplast
        CI1% = CI%: REMIX$ = NROREMIT$: NROPED& = NPEDI&
        TOTALMETROS# = 0
        If ESTRAZABLE%(CI1%) < 1 Then Exit Sub
        
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "CAMBACERES") > 0 Then
           ESCAMBACERES% = 1
        End If
        
        CONDI$ = "CodiMovi = 'RT'"
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMIX$ & "'"
        CONDI$ = CONDI$ + " AND COD_INTE = " & CI1%
        
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        Dim rs As New ADODB.Recordset
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            CI1% = XVALO(!cod_inte)
            COD_EXTEX = SAFETEXT$(!Cod_Exte)
            NuorItemX = XVALO(!NuOrItem)
            DESCRI_LARX = SAFETEXT$(!DESCRI_LAR)
            CantSalidX = ROUND(XVALO(!CANTSALID), 0)
            NUPEDCLIXX = XVALO(!NUPEDCLI)
            Nume_ClieX = XVALO(!Nume_Clie)
            FECHMOVX% = CVINT(!FECHMOV)
            IDENLOTEX$ = SAFETEXT$(!idenlote)
            SOLOELLOTE$ = IDENLOTEX$
            
            'If ESCAMBACERES% > 0  Then
               'PASO LA CANTIDAD A METROS Y EL peso QUE DEBE VENIR DE LA BOBINA LO RECALCULO A METROS
               METROX# = MetrosSegunLote#(IDENLOTEX$)
               If METROX# > 0 And ESCAMBACERES% > 0 Then
                   SOLOELLOTE$ = SOLOELLOTE$ & " - " & ROUND(METROX#, 0) & " mts."
                   LOTTE$ = SOLOELLOTE$
                   Call AGREGATEXTO(LOTTE$, 2)
                   TOTALMETROS# = TOTALMETROS# + METROX#
               Else
                  'en el caso de conteplast se agrega si tiene la ficha tecnica del cliente
                  CodEspTecnicaDelCliente$ = TRIM$(CTP.CodEspTecCliente$(CI1%))
                  If CodEspTecnicaDelCliente$ <> "" Then
                     SOLOELLOTE$ = SOLOELLOTE$ & " - F.Tèc.: " & CodEspTecnicaDelCliente$
                     IDENLOTEX$ = SOLOELLOTE$
                  End If
                  Call AGREGATEXTO(IDENLOTEX$, 2)
               End If
'           End If
           .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
   
       If InStr(EPAC$, "CAMBACERES") > 0 Then
         Call CARLINSEP
         Call AGREGATEXTO("SubTotal Mts.: " & TRIM$(FORMATNUM$(TOTALMETROS#, "F12.0")), 2)
         Call CARLINSEP
       End If
       Call CARLINSEP("INVISIBLE")
End Sub

Function EsPorPedidoCambaceres&(NC)
    TituloLista$ = "Pedidos Pendientes Cliente Activo"
    

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Nro.Pedido/F10;Referencia/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro.Pedido/F10;Referencia/A64", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    Dim obj As New RECORXET
    Set RS1 = obj.RS_PedidoPendxCte2(NC)
    Call Carga_MSF_Recordset(RS1, "Nro.Pedido/F10;Referencia/A64", FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0

    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    EsPorPedidoCambaceres& = XVALO(RESP_ZELE_VECT(1))
    
End Function

Sub GeneraArchivoRemitoConteplast(REMIX$, NROPED&)
        'ESTO ES ESTANDO PARADO EN CONTEPLAST O EN CAMBACERES
        ESCAMBACERES% = 0
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "CAMBACERES") > 0 Then
           ESCAMBACERES% = 1
        End If
        
        CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED= " & NROPED&, "NOMESS"))
        VEND% = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED= " & NROPED&, "NOMESS"))
        DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED= " & NROPED&, "NOMESS")

        Ruta$ = ""
        CONDI$ = "CodiMovi = 'RT'"
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMIX$ & "' "


        '
        CARPETASINCRO$ = TRIM$(Control("CONWALDB", "ESCRMOVI"))
        Ruta$ = CARPETASINCRO$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
           Exit Sub
        End If
        
        A$ = Date
        A$ = REPLACAR(A$, "/", "_")
        B$ = Time
        B$ = REPLACAR(B$, ".", "")
        B$ = REPLACAR(B$, ":", "")
        REMISINGUION$ = REPLACAR(REMIX$, "-", "_")
        NARCHI2% = FILEOPEN%(Ruta$ + "\" + REMISINGUION$ + A$ + " " + B$ + ".TXT", 2, 1024)    ' ABRE EL ARCHIVO DE ESCRITURA
        TX$ = AJUSTI$("REMITO", 18): TX$ = TX$ + ";" 'NRO REMITO
        TX$ = TX$ + AJUSTI$("MONEDA", 6): TX$ = TX$ + ";"  'MONEDA
        TX$ = TX$ + AJUSTI$("TICAMBIO", 8): TX$ = TX$ + ";" 'TIPO DE CAMBIO
        TX$ = TX$ + AJUSTI$("CODINT", 6): TX$ = TX$ + ";" ' COD INT
        TX$ = TX$ + AJUSTI$("CODEXTE", 20): TX$ = TX$ + ";" 'CODEXTE
        TX$ = TX$ + AJUSTI$("NUORI", 5): TX$ = TX$ + ";" 'ORDEN
        TX$ = TX$ + AJUSTI$("DESCRILAR", 64): TX$ = TX$ + ";" 'DESCRILAR
        TX$ = TX$ + AJUSTI$("CANTSALID", 10): TX$ = TX$ + ";" 'CANTIDAD
        TX$ = TX$ + AJUSTI$("CANTINGRE", 10): TX$ = TX$ + ";" 'CANTI INGRE
        TX$ = TX$ + AJUSTI$("NUORDPED", 8): TX$ = TX$ + ";" 'NUMEORDPED
        TX$ = TX$ + AJUSTI$("NUPEDCLI", 8): TX$ = TX$ + ";" 'NUPEDCLI,
        TX$ = TX$ + AJUSTI$("ValoUnit", 14): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'ValoUnit
        TX$ = TX$ + AJUSTI$("N_Clie", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'N_Clie
        TX$ = TX$ + AJUSTI$("FECHMOV", 8): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'FECHMOV
        TX$ = TX$ + AJUSTI$("IDENLOTE", 12): TX$ = TX$ + ";"  ': TX$ = TX$ + ";" 'IDENLOTE
        TX$ = TX$ + AJUSTI$("CondPag", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'CondPag
        TX$ = TX$ + AJUSTI$("VENDED", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'VENDED
        TX$ = TX$ + AJUSTI$("DOMIENT", 48): TX$ = TX$ + ";"  'DOMIENT
        TX$ = TX$ + AJUSTI$("R.SOCIAL", 48): TX$ = TX$ + ";"  'R.SOCIAL
        TX$ = TX$ + AJUSTI$("NOMBREVENDEDOR", 48): TX$ = TX$ + ";"  'NOMBRE DE VENDEDOR
        TX$ = TX$ + AJUSTI$("OREN DE COMPRA", 32): TX$ = TX$ + ";"  'ORDEN DE COMPRA
        TX$ = TX$ + AJUSTI$("NUMVEND.ENVAP", 16): TX$ = TX$ + ";"  'NRO DE VENDEDOR DE ENVAPLAST
        TX$ = TX$ + AJUSTI$("OBSERVAITEM", 256): TX$ = TX$ + ";"  'OBSERVACION DEL ITEM CONTEPLAST
        TX$ = TX$ + AJUSTI$("DESCRIDOS", 128): TX$ = TX$ + ";"  'SEGUNDA DESCRIPCION
        
        
        Print #NARCHI2%, TX$ '
        '
        VUELTA% = 0
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        Dim rs As New ADODB.Recordset
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
'            MONEVALX = XVALO(!MoneVal) 'HAGO UN VALOBADA DE PEDDETA MAS ABAJO P.Q. EN MOVISTO DESDE EL REMITO LO GRABA EN PESOS
'            TIPOCAMX = ROUND(XVALO(!TIPOCAM), 2)
            CI1% = XVALO(!cod_inte)
            COD_EXTEX = SAFETEXT$(!Cod_Exte)
            NuorItemX = XVALO(!NuOrItem)
            DESCRI_LARX = SAFETEXT$(!DESCRI_LAR)
            CantSalidX = ROUND(XVALO(!CANTSALID), 0)
            CANTINGREX = ROUND(XVALO(!CANTINGRE), 0)
            NUMEORDPEDX = XVALO(!NUMEORDPED)
            NUPEDCLIXX = XVALO(!NUPEDCLI)
            ValoUnitX = ROUND(XVALO(!VALOUNIT), 3)
            Nume_ClieX = XVALO(!Nume_Clie)
            FECHMOVX% = CVINT(!FECHMOV)
            IDENLOTEX$ = SAFETEXT$(!idenlote)
            RAZONSOCX$ = Left$(rasocli$(Nume_ClieX), 48)
            DESCRIDOS$ = DESCRIT2$(CI1%)
            
            If NUPEDCLIXX > 0 Then
              If VUELTA% < 1 Then
                'ESTO DE PEDENCA
                CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED = " & NUPEDCLIXX, "NOMESS"))
                VEND% = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED = " & NUPEDCLIXX, "NOMESS"))
                NAMEVENDX$ = TRIM$(Left$(ECOARCH$("VENDEDOR", Chr$(VEND%)), 48))
                DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED = " & NUPEDCLIXX, "NOMESS")
                ORDENCOMPRA$ = VALOBADA("PEDENCA", "NroOcom", "NROPED = " & NUPEDCLIXX, "NOMESS")
                NROVENENVAP% = XVALO(VALOBADA("PEDENCA", "AtributoComercial", "NROPED = " & NUPEDCLIXX, "NOMESS"))
                VUELTA% = 1
              End If
              
              'TRAE EL PRECIO DE PEDDETA
              ValoUnitX = XVALO(VALOBADA("PEDDETA", "PREUNID", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
              MONEVALX = XVALO(VALOBADA("PEDDETA", "MONEEMIS", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
              TIPOCAMX = XVALO(VALOBADA("PEDDETA", "TIPOCAM", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
              OBSERVACIONITEM$ = VALOBADA("PEDDETA", "OBSERVAITEM", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS")
              
            End If
            If ESCAMBACERES% > 0 Then
               'PASO LA CANTIDAD A METROS Y EL PRECIO QUE DEBE VENIR DE LA BOBINA LO RECALCULO A METROS
               METROX# = MetrosSegunLote#(IDENLOTEX$)
               If METROX# > 0 Then
                    If CantSalidX > 0 Then
                       CantSalidX = METROX#
                    ElseIf CANTINGREX > 0 Then
                       CANTINGREX = METROX#
                    End If
                    ValoUnitX = ROUND(ValoUnitX / METROX#, 3)
               End If
            End If
        
            
            TX$ = AJUSTI$(SAFETEXT$(REMIX$), 18): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(MONEVALX), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(TIPOCAMX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CI1%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(COD_EXTEX), 20): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NuorItemX), 5): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DESCRI_LARX), 64): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CantSalidX), 10): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CANTINGREX), 10): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUMEORDPEDX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUPEDCLIXX), 8): TX$ = TX$ + ";"
            'AGREGO ESTO POR QUE NO ESTA IMPRIMIENDO ESTOS 4 DATOS EN EL TXT

            TX$ = TX$ + AJUSTI$(SAFETEXT$(ValoUnitX), 14): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(Nume_ClieX), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(FECHATEX$(FECHMOVX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(IDENLOTEX$), 12): TX$ = TX$ + ";"
            
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CONDVTA%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(VEND%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DOMIENTREG$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(RAZONSOCX$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NAMEVENDX$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(ORDENCOMPRA$), 32): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NROVENENVAP%), 16): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(OBSERVACIONITEM$), 256): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DESCRIDOS$), 128): TX$ = TX$ + ";"
            
'            Debug.Print TX$
            Print #NARCHI2%, TX$ '
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        '
        Close #NARCHI2%
        '''
''''        If TRIM$(FILTROEXTRA$) <> "" Then
''''           CONDI$ = CONDI$ + " AND " & FILTROEXTRA$ ' EN EL CASO QUE SE QUIERA LLAMAR A LA RUTINA PARA QUE GENERE LOS MOVIMIENTOS
''''        End If                                      ' EN WALDBOTT MOVIMIENTOS MAL GRABADOS QUE DEBERIAN TENER LA MARCA 9999 EN CAMPO VerificadoPor
'''
'''        BASEMDBCONT$ = TRIM$(CONTROL("CONWALDB", "ESCRMOVI"))  'CONFIGURACION DE RUTA EN COTIZACO DONDE ESCRIBE CONTEPLAST
'''        BASEMDCONT2$ = TRIM$(CONTROL("CONWALDB", "ESCRMOV2"))  'CONFIGURACION en cotizaco (es la misma en que leee envaplast _
'''                                                                pero la ruta puede ser diferene dependiendo de donde se este parado
'''
'''        SQLX$ = "SELECT MOVISTO_ID FROM MOVISTO WITH(NOLOCK)"
'''        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
'''        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
'''        Dim RS As New ADODB.Recordset
'''        Set RS = READSET(SQLX$)
'''        With RS
'''          Do While Not .EOF
'''              ID& = XVALO(!MOVISTO_ID) 'LEE EL ID DE LA TRANSACCION DEL MOVISTO
'''              CONDI$ = "MOVISTO_ID = " & ID&  'DUPLICA EL MOVIM EN LA TABLA ACCESS SEGUN EL ID DEL MOVISTO SQL
'''              Call DuplicarRegistroSqlaMdb("MOVISTO", BASEMDBCONT$, "MOVISTO", CONDI$, 1, NROPED&, 1)  'EN LAS 2 TABLAS
'''              Call DuplicarRegistroSqlaMdb("MOVISTO", BASEMDCONT2$, "MOVISTO", CONDI$, 1, NROPED&, 2)
'''             .MoveNext
'''          Loop
'''        End With
'Sub GENERAR_TRANSFERENCIA_OCOMPRA(Optional MODO%, Optional MENSAXE$)
End Sub





Sub ProcesoRecuperaKgDeRecepcionAMovistoDondeCodimoviIgualBR()
    'ACTUALIZA EL MOVISTO CON EL PESO DEL BOLRECEP SOLO LOS CODIMOVI = 'BR' Y QUE NO TENGAN PESOKG
    'SOLO SE ACTUALIZA EN EL MOVISTO EL INGRESO QUE SEA CON CODIMOVI = BR
    SQLX$ = " Update m "
    SQLX$ = SQLX$ + " Set m.PESOKG = B.Peso"
    SQLX$ = SQLX$ + " FROM MOVISTO M"
    SQLX$ = SQLX$ + " JOIN BOLRECEP B ON M.Cod_Inte = B.Cod_Inte AND M.IDENLOTE = B.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE M.CODIMOVI = 'BR'"
    SQLX$ = SQLX$ + " AND (M.PESOKG = 0 OR M.PESOKG IS NULL)"
    FLEXCONN.Execute (SQLX$)
End Sub

Function TieneDerechos(Ruta As String) As Boolean
    'VALIDA SI UNA RUTA TIENE PERMISOS DE LECDTDURA Y ESCRITURA
    'EJEMPLO DE USO
'     Dim ruta As String
'    ruta = "C:\ruta\del\archivo_o_directorio" ' Cambia esto por la ruta que deseas verificar
'
'    If TienePermisos(ruta) Then
'        MsgBox "La ruta tiene permisos de lectura y escritura."
'    Else
'        MsgBox "No se tienen permisos de lectura y/o escritura en la ruta."
'    End If


    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    ' Verificar si la ruta existe
    If Not fso.FolderExists(Ruta) Then
        TieneDerechos = False
        Exit Function
    End If
    
    ' Verificar derechos de lectura
    Dim archivoPrueba As String
    archivoPrueba = Ruta & "\prueba.txt"
    
    ' Crear un archivo temporal para probar derechos de escritura
    Dim archivo As Object
    Set archivo = fso.CreateTextFile(archivoPrueba, True)
    
    If Err.Number <> 0 Then
        TieneDerechos = False
        Exit Function
    End If
    
    ' Cerrar y eliminar el archivo temporal
    archivo.Close
    fso.DeleteFile archivoPrueba
    
    ' Si llegamos aquí, tiene derechos de lectura y escritura
    TieneDerechos = True
End Function



Sub GeneraArchivoRemitoNeuquen(REMIX$, NROPED&)
        'ESTO ES ESTANDO PARADO EN NEUQEUN CUANDO GENERA EL REMITO CREA EL ARCHIVO EN EL DROP BOX
        
'        CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED= " & NROPED&, "NOMESS"))
'        VEND% = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED= " & NROPED&, "NOMESS"))
'        DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED= " & NROPED&, "NOMESS")

        Ruta$ = ""
        CONDI$ = "CodiMovi = 'RT'"
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMIX$ & "' "


        '
        CARPETASINCRO$ = RutaCarpetaRemitosNeuquen$

        Ruta$ = CARPETASINCRO$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Ruta de DropBox\En Tablas Auxiliares configurar en la anteultima columna como 'REMITOSNEUQUEN'\ y su ruta en la ultima columna ")
           Exit Sub
        End If
        
        A$ = Date
        A$ = REPLACAR(A$, "/", "_")
        B$ = Time
        B$ = REPLACAR(B$, ".", "")
        B$ = REPLACAR(B$, ":", "")
        REMISINGUION$ = REPLACAR(REMIX$, "-", "_")
        NARCHI2% = FILEOPEN%(Ruta$ + "\" + REMISINGUION$ + A$ + " " + B$ + ".TXT", 2, 1024)    ' ABRE EL ARCHIVO DE ESCRITURA
        TX$ = AJUSTI$("REMITO", 18): TX$ = TX$ + ";" 'NRO REMITO
        TX$ = TX$ + AJUSTI$("MONEDA", 6): TX$ = TX$ + ";"  'MONEDA
        TX$ = TX$ + AJUSTI$("TICAMBIO", 8): TX$ = TX$ + ";" 'TIPO DE CAMBIO
        TX$ = TX$ + AJUSTI$("CODINT", 6): TX$ = TX$ + ";" ' COD INT
        TX$ = TX$ + AJUSTI$("CODEXTE", 20): TX$ = TX$ + ";" 'CODEXTE
        TX$ = TX$ + AJUSTI$("NUORI", 5): TX$ = TX$ + ";" 'ORDEN
        TX$ = TX$ + AJUSTI$("DESCRILAR", 64): TX$ = TX$ + ";" 'DESCRILAR
        TX$ = TX$ + AJUSTI$("CANTSALID", 10): TX$ = TX$ + ";" 'CANTIDAD
        TX$ = TX$ + AJUSTI$("CANTINGRE", 10): TX$ = TX$ + ";" 'CANTI INGRE
        TX$ = TX$ + AJUSTI$("NUORDPED", 8): TX$ = TX$ + ";" 'NUMEORDPED
        TX$ = TX$ + AJUSTI$("NUPEDCLI", 8): TX$ = TX$ + ";" 'NUPEDCLI,
        TX$ = TX$ + AJUSTI$("ValoUnit", 14): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'ValoUnit
        TX$ = TX$ + AJUSTI$("N_Clie", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'N_Clie
        TX$ = TX$ + AJUSTI$("FECHMOV", 8): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'FECHMOV
        TX$ = TX$ + AJUSTI$("IDENLOTE", 12): TX$ = TX$ + ";"  ': TX$ = TX$ + ";" 'IDENLOTE
        TX$ = TX$ + AJUSTI$("CondPag", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'CondPag
        TX$ = TX$ + AJUSTI$("VENDED", 6): TX$ = TX$ + ";" ': TX$ = TX$ + ";" 'VENDED
        TX$ = TX$ + AJUSTI$("DOMIENT", 48): TX$ = TX$ + ";"  'DOMIENT
        TX$ = TX$ + AJUSTI$("R.SOCIAL", 48): TX$ = TX$ + ";"  'R.SOCIAL
        TX$ = TX$ + AJUSTI$("NOMBREVENDEDOR", 48): TX$ = TX$ + ";"  'NOMBRE DE VENDEDOR
        TX$ = TX$ + AJUSTI$("OREN DE COMPRA", 32): TX$ = TX$ + ";"  'ORDEN DE COMPRA
        TX$ = TX$ + AJUSTI$("NUMVEND.ENVAP", 16): TX$ = TX$ + ";"  'NRO DE VENDEDOR DE ENVAPLAST
        TX$ = TX$ + AJUSTI$("OBSERVAITEM", 256): TX$ = TX$ + ";"  'OBSERVACION DEL ITEM CONTEPLAST
        TX$ = TX$ + AJUSTI$("DESCRIDOS", 128): TX$ = TX$ + ";"  'SEGUNDA DESCRIPCION
        
        
        Print #NARCHI2%, TX$ '
        '
        VUELTA% = 0
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        Dim rs As New ADODB.Recordset
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
'            MONEVALX = XVALO(!MoneVal) 'HAGO UN VALOBADA DE PEDDETA MAS ABAJO P.Q. EN MOVISTO DESDE EL REMITO LO GRABA EN PESOS
'            TIPOCAMX = ROUND(XVALO(!TIPOCAM), 2)
            CI1% = XVALO(!cod_inte)
            COD_EXTEX = SAFETEXT$(!Cod_Exte)
            NuorItemX = XVALO(!NuOrItem)
            DESCRI_LARX = SAFETEXT$(!DESCRI_LAR)
            CantSalidX = ROUND(XVALO(!CANTSALID), 0)
            CANTINGREX = ROUND(XVALO(!CANTINGRE), 0)
            NUMEORDPEDX = XVALO(!NUMEORDPED)
            'NUPEDCLIXX = XVALO(!NUPEDCLI) 'aca cambio por que es el pedido de avellaneda el que pasa, que lo ingresa manualmente
            NUPEDCLIXX = XVALO(NROPED&) 'es el que viene como parametro, solo se agrega al archivo p que se grabe en movistoctp
            ValoUnitX = ROUND(XVALO(!VALOUNIT), 3)
            Nume_ClieX = XVALO(!Nume_Clie)
            FECHMOVX% = CVINT(!FECHMOV)
            IDENLOTEX$ = SAFETEXT$(!idenlote)
            RAZONSOCX$ = Left$(rasocli$(Nume_ClieX), 48)
            DESCRIDOS$ = DESCRIT2$(CI1%)
            
'            If NUPEDCLIXX > 0 Then
'              If VUELTA% < 1 Then
'                'ESTO DE PEDENCA
'                CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED = " & NUPEDCLIXX, "NOMESS"))
'                VEND% = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED = " & NUPEDCLIXX, "NOMESS"))
'                NAMEVENDX$ = TRIM$(Left$(ECOARCH$("VENDEDOR", Chr$(VEND%)), 48))
'                DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED = " & NUPEDCLIXX, "NOMESS")
'                ORDENCOMPRA$ = VALOBADA("PEDENCA", "NroOcom", "NROPED = " & NUPEDCLIXX, "NOMESS")
'                NROVENENVAP% = XVALO(VALOBADA("PEDENCA", "AtributoComercial", "NROPED = " & NUPEDCLIXX, "NOMESS"))
'                VUELTA% = 1
'              End If
'
'              'TRAE EL PRECIO DE PEDDETA
'              ValoUnitX = XVALO(VALOBADA("PEDDETA", "PREUNID", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
'              MONEVALX = XVALO(VALOBADA("PEDDETA", "MONEEMIS", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
'              TIPOCAMX = XVALO(VALOBADA("PEDDETA", "TIPOCAM", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS"))
'              OBSERVACIONITEM$ = VALOBADA("PEDDETA", "OBSERVAITEM", "NROPED = " & NUPEDCLIXX & " AND CODIINT = " & CI1%, "NOMESS")
'
'            End If

        
            
            TX$ = AJUSTI$(SAFETEXT$(REMIX$), 18): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(MONEVALX), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(TIPOCAMX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CI1%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(COD_EXTEX), 20): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NuorItemX), 5): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DESCRI_LARX), 64): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CantSalidX), 10): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CANTINGREX), 10): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUMEORDPEDX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUPEDCLIXX), 8): TX$ = TX$ + ";"
            'AGREGO ESTO POR QUE NO ESTA IMPRIMIENDO ESTOS 4 DATOS EN EL TXT

            TX$ = TX$ + AJUSTI$(SAFETEXT$(ValoUnitX), 14): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(Nume_ClieX), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(FECHATEX$(FECHMOVX), 8): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(IDENLOTEX$), 12): TX$ = TX$ + ";"
            
            TX$ = TX$ + AJUSTI$(SAFETEXT$(CONDVTA%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(VEND%), 6): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DOMIENTREG$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(RAZONSOCX$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NAMEVENDX$), 48): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(ORDENCOMPRA$), 32): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NROVENENVAP%), 16): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(OBSERVACIONITEM$), 256): TX$ = TX$ + ";"
            TX$ = TX$ + AJUSTI$(SAFETEXT$(DESCRIDOS$), 128): TX$ = TX$ + ";"
            
'            Debug.Print TX$
            Print #NARCHI2%, TX$ '
            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        '
        Close #NARCHI2%
        
        '''GRABAR EN NEUQUEN TAMBIEN EL MOVISTOCTP---
        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        Dim RS2 As New ADODB.Recordset
        Set RS2 = READSET(SQLX$)
        With RS2
          Do While Not .EOF
'            MONEVALX = XVALO(!MoneVal) 'HAGO UN VALOBADA DE PEDDETA MAS ABAJO P.Q. EN MOVISTO DESDE EL REMITO LO GRABA EN PESOS
'            TIPOCAMX = ROUND(XVALO(!TIPOCAM), 2)
            CI1% = XVALO(!cod_inte)
            COD_EXTEXERNO$ = SAFETEXT$(!Cod_Exte)
            NuorItemX = XVALO(!NuOrItem)
            DESCRI_LARX = SAFETEXT$(!DESCRI_LAR)
            CantSalidX = ROUND(XVALO(!CANTSALID), 0)
            CANTINGREX = ROUND(XVALO(!CANTINGRE), 0)
            NUMEORDPEDX = XVALO(!NUMEORDPED)
            'NUPEDCLIXX = XVALO(!NUPEDCLI) 'aca cambio por que es el pedido de avellaneda el que pasa, que lo ingresa manualmente
            NUPEDCLIXX = XVALO(NROPED&) 'es el que viene como parametro, solo se agrega al archivo p que se grabe en movistoctp
            ValoUnitX = ROUND(XVALO(!VALOUNIT), 3)
            Nume_ClieX = XVALO(!Nume_Clie)
            FECHMOVX% = CVINT(!FECHMOV)
            IDENLOTEX$ = SAFETEXT$(!idenlote)
            RAZONSOCX$ = Left$(rasocli$(Nume_ClieX), 48)
            DESCRIDOS$ = DESCRIT2$(CI1%)

            If VUELTA% = 0 Then 'ESTO ES POR LAS DUDAS PARA QUE UN MISMO REMITO NO SE SOBRE ESCRIBA
                CONDI$ = "DOPRILAR = '" & REMIX$ & "'"
                Call ACTUREGISTRO("MOVISTOCTP", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
                VUELTA% = 1
            End If

            Call GrabaRemitoSucursalCTP(REMIX$, XVALO(MONEVALX), XVALO(TIPOCAMX), CI1%, COD_EXTEXERNO$, XVALO(NuorItemX), SAFETEXT$(DESCRI_LARX), CANTINGREX, CantSalidX _
                 , NUMEORDPEDX, XVALO(NUPEDCLIXX), XVALO(ValoUnitX), Nume_ClieX, HOY(HO$), LOTEX$, _
                 0, 0, "", RAZONSOCX$, "", "", 0, "", DESCRIDOS$)
            .MoveNext
          Loop
        End With
        RS2.Close
        Set RS2 = Nothing
        CONDI3$ = CONDI$ + " AND MARBORRA = 1"
        Call BORRAREGISTROS("MOVISTOCTP", CONDI3$, REGAF&, "NOMESS")

End Sub







Sub DuplicarRegistroSqlaMdb(TablaOrigen$, BaseDestino$, TablaDestino$, CONDI$, Optional NOREALIZARDELETE%, Optional NROPED&, Optional interemi%)
    'GENERA UNA TRANSACCION DE UNA TABLA DE UNA BASE A OTRA
    'LA IDEA ES CUANDO EN UNA TABLA ORIGEN SE GENERA UNA TRANSACCION LA DUPLIQUE EN OTRA TABLA IGUAL DE OTRA BASE
    'PRIMERO BORRA EL REGISTRO DE LA TABLA DESTINO Y LUEGO HACE UN INSERT EN LA MISMA BAJO MISMA CONDICION
    'UNA VEZ ES LLAMADA PARA LA BASE INTERREMI Y OTRA VEZ A LA BASE INTERREMI2
    If interemi% = 1 Then
      Call abreBaMDBINTER_REMI
      X$ = "INTER1"
    ElseIf interemi% = 2 Then
      abreBaMDBINTER_REMI_2
      X$ = "INTER 2 "
    End If
    
    Origen$ = TablaOrigen$
    DESTINO$ = TablaDestino$
    Dim ENV As New Envap
    '
 MsgBox Origen$ & " 4 " & "125125"
 MsgBox DESTINO$ & " 5 " & "125125"

    CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED= " & NROPED&, "NOMESS"))
    VEND& = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED= " & NROPED&, "NOMESS"))
    DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED= " & NROPED&, "NOMESS")
    
    'PRIMERO BORRA EL REGISTRO DE LA TABLA DESTINO Y LUEGO HACE UN INSERT EN LA MISMA BAJO MISMA CONDICION
    If NOREALIZARDELETE% < 1 Then 'USAR EN EL CASO QUE SE NECESITE DUPLICAR DE UNA MISMA TABLA
        SQLX$ = "DELETE * FROM " & DESTINO$ & " WHERE " & CONDI$
 MsgBox DESTINO$ & " 6 " & SQLX$
        If interemi% = 1 Then
          BASEINTERREMI.Execute (SQLX$)
        ElseIf interemi% = 2 Then
          BASEINTERREMI_2_CONT.Execute (SQLX$)
        End If
    End If
    '
    'INSERTA EL/LOS REGISTRO EN LA TABLA DESTINO DE LA TABLA ORIGEN SEGUN CONDICION
    'CAM$ = MuestraAsignacionCamposdeUnaTabla$("MOVISTO") 'ESTO GENERA EL CODIGO DE LA ASIGNACION DE LOS CAMPOS :!FeReMovi = RS0!FeReMovi
    SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " & CONDI$
    Set RS0 = READSET(SQLX$) 'DATOS DE MOVISTO SQL
    
    SQLX$ = "SELECT * FROM MOVISTO "
'    Call abreBaMDBINTER_REMI
    Set rs = New ADODB.Recordset
    If interemi% = 1 Then
       rs.Open (SQLX$), BASEINTERREMI, adOpenKeyset, adLockPessimistic, adCmdText
    ElseIf interemi% = 2 Then
       rs.Open (SQLX$), BASEINTERREMI_2_CONT, adOpenKeyset, adLockPessimistic, adCmdText
    End If
    

    Do While Not RS0.EOF
          With rs  'CORRESPONDE AL RECORDSET DE TABLA ACCES
            .AddNew
            !FeReMovi = RS0!FeReMovi
            !CODIEMPR = RS0!CODIEMPR
            !NumUsuar = RS0!NumUsuar
            !IdenMovi = RS0!IdenMovi
            !DeLaRemi = RS0!DeLaRemi
            !DESCRI_LAR = RS0!DESCRI_LAR
            !FECHMOV = RS0!FECHMOV
            !cod_inte = RS0!cod_inte
            !Cod_Exte = RS0!Cod_Exte
            !Descrip = RS0!Descrip
            !idenlote = RS0!idenlote
            !LotConsum = RS0!LotConsum
            !PNL_Des = RS0!PNL_Des
            !PLC_Has = RS0!PLC_Has
            !CODIMOVI = RS0!CODIMOVI
            !DoPriCor = RS0!DoPriCor
            !DOPRILAR = RS0!DOPRILAR
            !DoSecCor = RS0!DoSecCor
            !DOSECLAR = RS0!DOSECLAR
            !DOCUORIG = RS0!DOCUORIG

            !NumRem_Cli = RS0!NumRem_Cli
            !NuOrItem = RS0!NuOrItem
            !NUPEDCLI = RS0!NUPEDCLI
            !NUMEORDPED = RS0!NUMEORDPED
            !ReferCor = RS0!ReferCor
            !CostoRep = RS0!CostoRep
            !VALOUNIT = RS0!VALOUNIT
            !VALOUNITIVA = RS0!VALOUNITIVA
            !ValTotIt = RS0!ValTotIt
            !ValTotItIva = RS0!ValTotItIva
            !MoneVal = RS0!MoneVal
            !TIPOCAM = RS0!TIPOCAM
            !depomovi = RS0!depomovi
            !Nume_Clie = RS0!Nume_Clie
            !Sucu_Clie = RS0!Sucu_Clie
            !Nume_Prov = RS0!Nume_Prov
            !CANTINGRE = RS0!CANTINGRE
            !CANTSALID = RS0!CANTSALID
            !CantSaldo = RS0!CantSaldo
            !Verificado = RS0!Verificado
            !CAUSA = RS0!CAUSA
            !User_Tran = RS0!User_Tran
            !CODCARPIMP = RS0!CODCARPIMP
            !VERIFICADOPOR = RS0!VERIFICADOPOR
            !FHor_Verif = RS0!FHor_Verif
            !TIPOTRAN = RS0!TIPOTRAN
            !OCOMPRA_RTOABIERTO = RS0!OCOMPRA_RTOABIERTO
            !nroenvio = RS0!nroenvio
            !NUMPVTA = RS0!NUMPVTA
            !DOPRINUM = RS0!DOPRINUM
            !IdCalidad = RS0!IdCalidad
            !ID_OPERACION = RS0!ID_OPERACION
            !MOVISTO_ID = RS0!MOVISTO_ID
            !PESOKG = RS0!PESOKG
            !STAT_TRANSF = 0
            !raso_cli = rasocli$(XVALO(RS0!Nume_Clie))
            !CondPag = CONDVTA%
            !VENDED = VEND&
            !DOMIENT = Left$(DOMIENTREG$, 255)

            .Update
            RS0.MoveNext
         End With
    Loop
    RS0.Close
    
    On Error Resume Next
    If interemi% = 1 Then
        If BASEINTERREMI.State = adStateOpen Then    ' si esta abierta
           BASEINTERREMI.Close
        End If
        
    ElseIf interemi% = 2 Then
        If BASEINTERREMI_2_CONT.State = adStateOpen Then    ' si esta abierta
           BASEINTERREMI_2_CONT.Close
        End If
    
    End If
    On Error GoTo 0
End Sub


Sub IMPRIMIR_DETAREMQR(NROREMIT$)
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        TEXTOQR$ = "": ESSCRAP% = 0: TEXTOQR_SC$ = ""
        VUELTA% = 0: Copia% = 0
        Dim VectorCantidadScrap(32767)
        'cantidad de copias
        CONDI$ = " doprilar = '" & NROREMIT$ & "'"
        CANTIRE = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
        If (CANTIRE Mod 8) = 0 Then 'si es multiplo de 8
           TotCantCopias% = CANTIRE / 8
        Else                         ' si no
           TotCantCopias% = Int(CANTIRE / 8) + 1
        End If
        AcumuloScrap% = 0
        
        SQLX$ = "SELECT cod_inte,cod_exte,descrip,cantsalid,metros,PESOKG,IDENLOTE FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID"
        Set Rst = READSET(SQLX$)
            ' ***** ATENCION PASO QUE NO SE I MPRIMIA EL QR HAY QUE QUITAR DE LA TABLA ARCHIVOS_QR LOS QUE NO TIENEN EL MAPEO EN PATH_IMAG = C:\FLEXOFT\IMAGDETRE_QR.JPG ****
70          With Rst
                Do While Not .EOF
                    CI1% = XVALO(!cod_inte)
                    Codigo$ = SAFETEXT$(!Cod_Exte)
                    Descripcion$ = SAFETEXT$(!Descrip)
                    CANTIX$ = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "f10"))
                    Largo$ = "0": If XVALO(!METROS) > 0.0005 Then Largo$ = TRIM$(FORMATNUM$(!METROS, "F12.1"))
                    Peso$ = "0": If XVALO(!PESOKG) > 0.005 Then Peso$ = TRIM$(FORMATNUM$(!PESOKG, "F12.1"))
                    LOTE$ = "-": If TRIM$(SAFETEXT$(!idenlote)) <> "" Then LOTE$ = TRIM$(SAFETEXT$(!idenlote))
                    TEXTOQR$ = TEXTOQR$ + Codigo$ & Chr$(13) & CANTIX$ & Chr$(13) & Largo$ & Chr$(13) & Peso$ & Chr$(13) & NROREMIT$ & Chr$(13) & LOTE$ & Chr$(13)
                    VUELTA% = VUELTA% + 1
                    If InStr(EPAC$, "CAMBACERES") > 0 Then
                        CF$ = TRIM$(CODFAMIL$(CI1%))
                        If CF$ = "SC" Then
                           ESSCRAP% = ESSCRAP% + 1
                           VectorCantidadScrap(CI1%) = VectorCantidadScrap(CI1%) + XVALO(!CANTSALID)
                           TEXTOQR_SC$ = TEXTOQR_SC$ + Codigo$ & Chr$(13) & CANTIX$ & Chr$(13) 'TEXTO QR PARA SCRAP BO URDIMBRE Y TRAMA
                        End If
                    End If
                    If VUELTA% >= 8 Then
                       Copia% = Copia% + 1
                       GoSub imprime_hoja
                       TEXTOQR$ = "": TEXTOQR_SC$ = ""
                       VUELTA% = 0
                       ESSCRAP% = 0

                    End If
                    .MoveNext
                Loop
            End With
            If (VUELTA% > 0 And VUELTA% < 8) Then
              Copia% = Copia% + 1
              GoSub imprime_hoja 'SI QUEDARON PENDIENTES IMPRIME
              TEXTOQR$ = "": TEXTOQR_SC$ = ""
              VUELTA% = 0
              ESSCRAP% = 0

            End If
                
            On Error Resume Next
            Rst.Close
            Set Rst = Nothing
            On Error GoTo 0
            
            Exit Sub
            
            
imprime_hoja:
            IMAG$ = "IMAGDETRE_QR.JPG"
            IDENTIFICADOR$ = "QR_DETREMQR"
            ANCHOX% = 0
            ALTOX% = 0
            '
            On Error Resume Next
            Kill LUCOM$ & IMAG$
            On Error GoTo 0
            If VUELTA% = ESSCRAP% Then 'SI TODOS LOS ITEMS QEU VA A IMPRIMIR SON SCRAP IMPRIME CODIGO Y CANTIDAD SOLO
              For i& = 1 To 32767
                 ind% = i&
                 If VectorCantidadScrap(ind%) > 0 Then
                    Codigo$ = CODEXT$(ind%)
                    CANTIX$ = TRIM$(FORMATNUM$(VectorCantidadScrap(ind%), "f10"))
                    TEXTOQR_SC$ = TEXTOQR_SC$ + Codigo$ & Chr$(13) & CANTIX$ & Chr$(13)
                 End If
              Next i&
              QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR_SC$, "QR_CODE", ANCHOX%, ALTOX%))
            Else
              'SI NO COMO ANTES IMPRIIME TODO CON KILOS Y METROS LOTE Y REMITO
              'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
              QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
            End If
            Call T_Espera(3)
            Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
            Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
    
            FILTX$ = TRIM$(NROREMIT$) & ";" & TRIM$(Str$(Copia%)) & ";" & TRIM$(Str$(TotCantCopias%))
            Call STDFORM("RXQR-CMB", FILTX$)
            Erase VectorCantidadScrap
            
Return
End Sub

Function ValidarCantaRemitirVSPedido%()
    ValidarCantaRemitirVSPedido% = 1
    Dim VectorRemitir#(32767)
    Dim VectorPendientePedido#(32767)
     
    For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI1% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            NUPED& = CVS(Mid$(YY$, 151, 4))
'            LOTE$ = Mid$(YY$, 107, 12)

            If CI1% > 0 And NUPED& > 0 Then
'                If UBound(VectorRemitir#) < CI1% Then  /COMENTADO POR QUE LO DIMENSIONE DIRECTAMENTE A 32767
'                   ReDim Preserve VectorRemitir#(CI1%)
'                   ReDim Preserve VectorPendientePedido#(CI1%)
'                End If
                VectorRemitir#(CI1%) = VectorRemitir#(CI1%) + CAN#
                If XVALO(VectorPendientePedido#(CI1%)) = 0 Then VectorPendientePedido#(CI1%) = CantidadPendienteSegunPedidoCodigo(NUPED&, CI1%)
                If CAN# > 0.005 Then
                End If
            End If
            
    Next YX&
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    

    For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI1% = CVI(Mid$(YY$, 83, 2))
            NUPED& = CVS(Mid$(YY$, 151, 4))
            If CI1% > 0 And NUPED& > 0 Then
               If VectorRemitir#(CI1%) > VectorPendientePedido#(CI1%) Then
                  Codigo$ = CODEXT$(CI1%)
                  OPX% = 0
                  If InStr(EPAC$, "CONTEPLAST") > 0 Then
                    OPX% = COMALTER%("Atenciòn!! Còdigo: " & Codigo$ & " Supera La Cantidad Pendiente del Pedido Nro.: " & NUPED& & "\\Cancelar\Continuar (Solo Con Derechos)")
                  Else
                    Call COMUNI("Atenciòn!! Còdigo: " & Codigo$ & " Supera La Cantidad Pendiente del Pedido Nro.: " & NUPED&)
                  
                  End If
                  If OPX% < 2 Then
                     ValidarCantaRemitirVSPedido% = -1
                     Exit For
                  End If
                  
                  If DERACCE%("Entregar Cantidad Superior a la Pedida", "ENTCANSU") < 2 Then
                     ValidarCantaRemitirVSPedido% = -1
                     Exit For
                  End If
               End If
            End If
    Next YX&
    
End Function

'
Private Sub BOTNEXT_Click()
    '\\\OT-05-420-WAR-26/06/05///
    If TRIM$(Label10) = "" Then
      Call MENSERR(24, "Numero de Punto de Venta no Valido")
      BOTNEXT.Enabled = True
      Exit Sub
    End If
    '\\\OT-05-420-WAR-FIN///
    NCLIE = Val(Text1)
    If rasocli(NCLIE) <> "" Then
       MODOPRE% = 1
       If Option2.Value = False And ACCESOATOYOTA.Label1 = "" Then
          RESP1% = 1
          PORBARRA% = 0
          TXT$ = "Abierto\Por Pedidos Pendientes"
          EPAC$ = TRIM$(UCase$(EMPREAC$))
          If InStr(EPAC$, "ALEJPA") > 0 Then TXT$ = "Abierto\Por Pedidos Pendientes\Carga Por Código de Barras"
          If InStr(EPAC$, "RETEPLAST") > 0 Then TXT$ = "Abierto\Por Lista de Despacho"
'          If EXISTE%("TRAZAB07.EXE") < 1 Then RESP1% = COMALTER%("Formato de Pre-Carga:\\" & TXT$)
          RESP1% = COMALTER%("Formato de Pre-Carga:\\" & TXT$)
          MODOPRE% = 2
          If RESP1% = 0 Then Exit Sub
          MODOPRE% = RESP1%
          If MODOPRE% = 3 Then PORBARRA% = 1
       End If
       Call SIGUEREMI1(MODOPRE%)
      Else
        Call MENSERR(24, "Numero de Cuenta no Valido")
    End If
    '
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        MSFDETAREMI.Rows = 1

        Unload FORMUL '\\\OT-06-0182-WAR-29/05/06///
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Exit Sub
    End If
    List1.Clear
    MSFDETAREMI.Rows = 1

    List1.Visible = True
    v& = 0
    '
    If Frame15.Visible = True Then
      'TOTITEMS% = 0 '*
      For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        Reli$ = Space$(80)
        TTXX$ = Mid$(X$, 5, 64)
        Call REPLA(Reli$, TTXX$, 1, 52)
        CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
        Call REPLA(Reli$, CANS$, 75, 10)
        List1.AddItem Reli$
        Call CargarFlexgrid(Reli$, U&)
        'TOTITEMS% = TOTITEMS% + 1 '*
      Next U&
      'Label18 = TRIM$(Str$(TOTITEMS%)) '*
      GoTo 20
    End If
    '
    'TOTITEMS% = 0 '*
    NUOCOMPRA$ = ""
    TieneMasPedidos% = 0: NROPEDLEIDO& = 0: MostarDatosEntrega% = 0
    If Control("REMSTAN", "DATENTPE") = "SI" Then MostarDatosEntrega% = 1
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        MEDIDAS$ = Mid$(X$, 193, 12)
        Reli$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            v& = v& + 1
            Call GRAREG("!BOLREDET", X$, v&)
            Call REPLA(Reli$, CODEXT$(CI%), 1, 15)
            Call REPLA(Reli$, DESCRIT0$(CI%), 16, 30)
            TXMD$ = TEXTOMEDI$(MEDIDAS$)
            If TXMD$ <> "" Then
               Call REPLA(Reli$, "P.C. " + TXMD$, 16, 30)
            End If
            Call REPLA(Reli$, Unimed$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(Reli$, CANS$, 53, 10)
            Call REPLA(Reli$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            If XVALO(NOPE$) > 0 And MostarDatosEntrega% = 1 Then
               If NROPEDLEIDO& < 1 Then
                  NROPEDLEIDO& = XVALO(NOPE$)
               Else
                  If NROPEDLEIDO& <> XVALO(NOPE$) Then
                     TieneMasPedidos% = 1
                  End If
               End If
            End If
            Call REPLA(Reli$, NOPE$, 78, 7)
            NPEDID& = CVS(Mid$(X$, 151, 4))
            
            NUOCO$ = TRIM$(NUOCOCLI$(NPEDID&, CI%))
            If NUOCO$ <> "" Then
              If InStr(NUOCOMPRA$, " " + NUOCO$ + ",") < 1 Then
                NUOCOMPRA$ = NUOCOMPRA$ + " " + NUOCO$ + ","
              End If
            End If
            List1.AddItem Reli$
            Call CargarFlexgrid(Reli$, v&)
            'TOTITEMS% = TOTITEMS% + 1 '*
        End If
    Next U&
    'Label18 = TRIM$(Str$(TOTITEMS%)) '*
    NUOCOMPRA$ = TRIM$(NUOCOMPRA$)
    If Right$(NUOCOMPRA$, 1) = "," Then
      NUOCOMPRA$ = Left$(NUOCOMPRA$, Len(NUOCOMPRA$) - 1)
    End If
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
    'SI NO SELECCIONA EL PEDIDO DESDE EL BOTON ! AGREGADO P/ ENVAPLAST  Y ELIGE UN PEDIDO EN EL BOTON NEXT
    'LE IMPRIME ESTE EN EL LBLNROPED PARA PRESENTAR LOS DATOS DE ENTREGA DEL PEDIDO.
        For U& = 1 To ULTREG&("!BOLREDET")
            X$ = REGLEIDO$("!BOLREDET", U&)
            NPEDID& = CVS(Mid$(X$, 151, 4))
            If XVALO(Me.lblNroPed) < 1 And NPEDID& > 0 Then lblNroPed = NPEDID&
            Exit For
        Next U&
        'FIN ENVAPLAST
    Else 'si no es conte,ni envaplast y MostarDatosEntrega% > 0 y hay un solo pedido seleccionado
      If MostarDatosEntrega% > 0 And TieneMasPedidos% < 1 Then lblNroPed = NROPEDLEIDO&
    End If
    
    
    If TRIM$(NUOCOMPRA$) <> "" Then Text20 = NUOCOMPRA$
    Call SETULTREG("!BOLREDET", v&)
    '
20  Command5.Enabled = True
    Command12.Enabled = True
    Text15.SetFocus
    '
End Sub

Private Sub Command1_Click()
    '
    NCLIE = ListaSeleccionClientes("", True)
    If NCLIE > 0 Then
        Text1.TEXT = TRIM$(Str$(NCLIE))
        Text2.TEXT = rasocli$(NCLIE)
        Text3.TEXT = DOMICLI$(NCLIE)
        Text4.TEXT = CPOSCLI$(NCLIE)
        Text5.TEXT = LOCACLI$(NCLIE)
        Label26 = CODICLI$(NCLIE)
        Call CARTRANSPOR
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Dim CIRF%(512), CARF(512)
    v& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        CI% = CVI(Mid$(X$, 83, 2))
        COBA$ = CODEXT$(CI%)
        CABA = CVD(Mid$(X$, 89, 8))
        If TRIM$(COBA$) <> "" Then
           Call LEEEXPLO(COBA$, CABA)
           '
           If CAIT% > 0 Then
              Call REPLA(X$, MKD$(CDbl(CABA)), 163, 8)
              Call REPLA(X$, String$(8, 0), 89, 8)
              Call GRAREG("!BOLREDET", X$, U&)
           End If
           '
           For J% = 1 To CAIT%
             CIY% = CIJ%(J%)
             CANY = CABA * USOI(J%)
             For W% = 1 To v&
               If CIRF%(W%) = CIY% Then
                  CARF(W%) = CARF(W%) + CANY
                  GoTo 19
               End If
             Next W%
             v& = v& + 1
             CIRF%(v&) = CIY%
             CARF(v&) = CANY
19         Next J%
        End If
    Next U&
    '
    For J% = 1 To v&
       X$ = REGBLAN$("BOLREDET")
       Call REPLA(X$, MKI$(CIRF%(J%)), 83, 2)
       Call REPLA(X$, MKD$(CDbl(CARF(J%))), 89, 8)
       Call REGAPP("!BOLREDET", X$)
    Next J%
    Call CIERRARCH("!BOLREDET")
    Command12.Enabled = False
    '
    List1.Clear
    MSFDETAREMI.Rows = 1

    v& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        Reli$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            v& = v& + 1
            Call GRAREG("!BOLREDET", X$, v&)
            Call REPLA(Reli$, CODEXT$(CI%), 1, 15)
            Call REPLA(Reli$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(Reli$, Unimed$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(Reli$, CANS$, 53, 10)
            Call REPLA(Reli$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(Reli$, NOPE$, 78, 7)
            List1.AddItem Reli$
            Call CargarFlexgrid(Reli$, v&)
        End If
    Next U&
    Call SETULTREG("!BOLREDET", v&)
    Call CIERRARCH("!BOLREDET")
    '
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    '\\\OT-05-0299-WAR-18/05/05///
    If DERACCE%("SETREMIT", "Configuración - Setup de Remitos.") >= 2 Then
      OPREMIT07.Show 1
    End If
    '\\\OT-05-0299-WAR-FIN///
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
  '
 
  NCXX = Val(Text1)
  If rasocli$(NCXX) <> "" Then
    '
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "GUIDI") > 0 Then
           Set FORMUL = BCODE07
       ElseIf InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
         If InStr(EPAC$, "CAMBACERES") > 0 Then
            TXT$ = "Abierto\Por Pedidos Pendientes"
            RESP1% = COMALTER%("Opciones de Trabajo\\" & TXT$)
              If RESP1% = 2 Then
                NroPedCamba& = EsPorPedidoCambaceres&(FOREMIT07.Text1)
                If NroPedCamba& > 0 Then
                 lblNroPed = TRIM$(Str(NroPedCamba&))
                End If
              End If
         End If
         Set FORMUL = BCODECTP
         FORMUL.Label4.Caption = Text10.TEXT
       Else
           Set FORMUL = BCODE09 ' MODIFICADO PARA ACHERNAR
       End If
    '
       
       FORMUL.Show 1
       If FORMUL.List1.ListCount > 0 Then
         '1.- DESHABILITO EL BOTON NEXT
         BOTNEXT.Enabled = False
               '2.- CARGAR BOLREDET  '
         JJ& = 0
         REBLA$ = REGBLAN$("!BOLREDET")
         For JX% = 0 To FORMUL.List1.ListCount - 1
           
           Z$ = FORMUL.List1.List(JX%)
           CI1$ = Mid$(Z$, 1, 17)
           CI2% = CODINT%(CI1$)
           Q1# = XVALO(Mid$(Z$, 49, 8))
           
           Y$ = REBLA$
           Call REPLA(Y$, MKI$(CI2%), 83, 2)
           Call REPLA(Y$, Left$(Unimed$(CI2%), 4), 85, 4)
           Call REPLA(Y$, MKD$(Q1#), 89, 8)
           If FORMUL.Name = "BCODE09" Then
            If FORMUL.VALI_BCODE09 = "APROB" Then
                LOTEX$ = Mid$(Z$, 61, 12)
                Call REPLA(Y$, LOTEX$, 107, 12)
            End If
           End If
           If FORMUL.Name = "BCODECTP" Then 'EN EL CASO DE CONTEPLAST AGREGA EL NRO DE PEDIDO
            If FORMUL.VALI_BCODE09 = "APROB" Then
                LOTEX$ = Mid$(Z$, 61, 12)
                Call REPLA(Y$, LOTEX$, 107, 12)
                NPED& = Mid$(Z$, 90, 9)
                Call REPLA(Y$, MKS$(NPED&), 151, 4)
            End If
           End If
           JJ& = JJ& + 1
           Call GRAREG("!BOLREDET", Y$, JJ&)
         Next JX%
         Call SETULTREG("!BOLREDET", JJ&)
         '3.- LLAMAR A SIGUEREMI2 DE FOREMIT07
         If JJ& > 0 Then Command12.Enabled = True
         Call List1_DblClick   'SIGUEREMI2
       ElseIf FORMUL.List1.ListCount <= 0 Then
         Call SETULTREG("!BOLREDET", 0)
         List1.Clear
         MSFDETAREMI.Rows = 1

       End If
       Exit Sub
     Else
       Call COMUNI("Debe Seleccionar un Cliente")
       Call Command1_Click
  End If
  '
End Sub

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    LIBREMIT07.Show 1
    Command15.Enabled = True
    '
End Sub

Private Sub Command16_Click()
    ACCESOATOYOTA.Show 1
    If ACCESOATOYOTA.Label1 <> "" Then
       Text1 = "228"   ' FUERZA TOYOTA
       Call Text1_KeyPress(13)
       Call BOTNEXT_Click
    End If
End Sub



Private Sub Command17_Click()
  Command17.Enabled = False
   PREFI$ = TRIM$(InputBox$("Prefijo de Remito:", "Ingrese el Prefijo"))
   If PREFI$ <> "" Then
    '
     If Len(PREFI$) > 4 Then
        Call COMUNI("Prefijo no Válido (Máximo 4 Dígitos)")
        GoTo 99
     End If
     '
     While Len(PREFI$) < 4: PREFI$ = "0" + PREFI$: Wend
     '
     For U& = 1 To 4
          UI% = Asc(Mid$(PREFI$, U&))
          If UI% < 48 Or UI% > 57 Then
              Call COMUNI("Prefijo no Válido")
              GoTo 99
              Exit Sub
          End If
     Next U&
     '
     Label10 = PREFI$
     NUPUVE$ = TRIM(Label10)
     CMO$ = "RT"
     NROPRO& = 0
     Call NUREMIFA(NROPRO&, NUPUVE$)
     Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))

     '
   End If
   
99 Command17.Enabled = True

End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   NCLI = Val(Text1)
   NCXI% = NCLI
   If ECOARCH("DEUDOR12", MKS$(NCLI)) = "" Then
      Call MENSERR(24, "Ingrese Número de Cliente")
      GoTo 99
   End If
   '
   RFF$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REGE& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("ALMAENT", REGE&)
     JJ& = JJ& + 1
     Call GRAREG("!ALMAENT", X$, JJ&)
   Next U&
   Call SETULTREG("!ALMAENT", JJ&)
   Call FRESHECHO("!ALMAENT")
   '
   Call SELECHO("!ALMAENT")
   Label11 = VALACT1$("!ALMAENT")
   '
99 Command18.Enabled = True
End Sub


Private Sub Command19_Click()
   Command19.Enabled = False
   NROOF$ = ORFASEL_ENVAPLAST$(1)
   If NROOF$ <> "" Then
     Text1 = ""
     NROPED& = XVALO(RESP_ZELE_VECT(6))
     NC = XVALO(RESP_ZELE_VECT(7))
     
   Else
     GoTo 99
   End If
   
   If NC > 0 Then
        Text1.TEXT = TRIM$(Str$(NC))
        Call Text1_KeyPress(13)
'        Text2.TEXT = rasocli$(NC)
'        Text3.TEXT = DOMICLI$(NC)
'        Text4.TEXT = CPOSCLI$(NC)
'        Text5.TEXT = LOCACLI$(NC)
'        Label26 = CODICLI$(NC)
        Me.lblNroPed = NROPED&
'        Call CARTRANSPOR
    End If
    
99 Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    If ULTREG&("!BOLREDET") > 0 Then
        Call CIERRARCH("!BOLREDET")
        Call BLANARCH("!BOLREDET")
        List1.Clear
        MSFDETAREMI.Rows = 1

        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Call LIACOBSE ' LIMPIA CASILLEROS INFERIORES - ACTIVA TEXT OBSERVACIONES
        Text1.SetFocus
       Else
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "ENVALAST") > 0 Then
           On Error Resume Next
           BASEWALD.Close
           On Error GoTo 0
       End If
         
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("")
    End If
End Sub

 Sub Command20_Click()
   If CLACONTRA% < 1 Then
     Exit Sub
   End If
   
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        Exit Sub
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(VALACT1$("SELREMIT"))
     If NUDOCU$ = "" Then Exit Sub
     Dim ENV As New Envap
     EXISTREMI% = ENV.REMITO_EXISTE_EN_WALDBOTT(NUDOCU$)
     If EXISTREMI% > 0 Then
       Call COMUNI("Atención !!! \N° de Remito Existente en Base de Datos Waldbott\Consulte a su Soporte de Sistema")
       Set ENV = Nothing
        Exit Sub
     End If
     Set ENV = Nothing
     ultimoremito& = 0

     Call RECUPERARMEMITOS(NUDOCU$, ultimoremito&)
     GoTo 12
End Sub
   Sub RECUPERARMEMITOS(Remito$, ultimoremito&)
   'RECUPERADOR DE REMITOS DE FLEXOFT A WALDBOT RECORRE TODO EL MOVISTO.
   'ACTUALIZA LOS CODIGOS
   Dim VECTORCI(32767)
   Dim ENV As New Envap
   If Remito$ <> "" And ultimoremito& > 0 Then
      Call COMUNI("Imposible Combinar Remito y Ultimo Nro de Remito")
      Exit Sub
   End If
   CONDI$ = "doprilar = '" & Remito$ & "'"
   NROPED& = XVALO(VALOBADA("MOVISTO", "nupedcli", CONDI$, "NOMESS"))
   Me.lblNroPed = NROPED&
   MonedaPedido% = XVALO(VALOBADA("PEDDETA", "MoneEmis", "NROPED= " & NROPED&, "NOMESS"))
   If MonedaPedido% > 1 Then
        
        Label16 = ECOARCH$("TAMONED", Chr$(MonedaPedido%))
        OPX% = COMALTER%("Pedido Generado en Moneda " & DEMONECO$(MonedaPedido%) & "\\ Generar Remito Manteniendo Moneda Original\Generar Remito en Pesos")
        If OPX% = 2 Then
          TipoCambioPedido# = XVALO(VALOBADA("PEDDETA", "TipoCam", "NROPED= " & NROPED&, "NOMESS"))
          Label15.Visible = True
          Label1(3).Visible = True
          Label16.Visible = True
          Label1(2).Visible = True
          TCP$ = FORMATNUM$(TipoCambioPedido#, "F9.2")
20        ImporteIngresado$ = TRIM$(InputBox$("Ingrese Cotización de Moneda (" & DEMONECO$(MonedaPedido%) & ")", " Cotización", TCP$))
          If ImporteIngresado$ = "" Or XVALO(ImporteIngresado$) < 1 Then GoTo 20 'SI NO INGRESA VUELVE SI O SI TIENE QUE INGRESAR UN VALOR
          'SI INGRESO VALOR LO PONE EN EL LABEL
          Label15 = FORMATNUM$(XVALO(ImporteIngresado$), "F9.2")
        End If
      Else
      
        Label15.Visible = False
        Label1(2).Visible = False
        Label15.Caption = ""
        Label1(3).Visible = False
        Label16.Visible = False: Label16.Caption = ""

      End If
   
   
   
   
   
   
   
   
   
   SQLX$ = "SELECT COD_INTE FROM MOVISTO WITH(NOLOCK)WHERE CODIMOVI = 'RT'"
   SQLX$ = SQLX$ + " AND CANTSALID > 0 AND COD_INTE > 0 "
   If TRIM$(Remito$) <> "" Then
      SQLX$ = SQLX$ + " and DOPRILAR = '" & Remito$ & "'"
   ElseIf UltimoRemitoRegistrado& > 0 Then
      SQLX$ = SQLX$ + " and Doprinum > UltimoRemitoRegistrado&"
   Else
      Call COMUNI("Imposible Combinar Remito y Ultimo Nro de Remito")
      Exit Sub
   End If
   SQLX$ = SQLX$ + " GROUP BY COD_INTE,DOPRILAR ORDER BY DOPRILAR"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        If VECTORCI(CI1%) < 1 Then
          VECTORCI(CI1%) = 1
          If ENV.EXISTE_CODIGO_WALDBOTT(CI1%) < 1 Then
              Call ENV.ALTA_ARTICULO_WALDBOTT(CI1%)
            Else
              Call ENV.UPDATE_ARTICULO_WALDBOTT(CI1%)
          End If
        End If
        .MoveNext
     Loop
   End With
   BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
   '
   'ACTUALIZA LA TABLA PDCABEZA Y PDITEMS
   If BDX$ <> "" Then
     SQLX$ = "SELECT DOPRILAR FROM MOVISTO WITH(NOLOCK) WHERE CODIMOVI = 'RT'"
     SQLX$ = SQLX$ + " AND CANTSALID > 0 AND COD_INTE > 0 "
     If TRIM$(Remito$) <> "" Then
       SQLX$ = SQLX$ + " and DOPRILAR = '" & Remito$ & "'"
     End If
      SQLX$ = SQLX$ + " GROUP BY  DOPRILAR "
     SQLX$ = SQLX$ + " ORDER BY DOPRILAR"
     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
         NROREMIT$ = SAFETEXT$(!DOPRILAR)
         Call REMITO_A_WALDBOTT(NROREMIT$)
         .MoveNext
       Loop
     End With
   End If
   '
           Label15.Visible = False
        Label1(2).Visible = False
        Label15.Caption = ""
        Label1(3).Visible = False
        Label16.Visible = False: Label16.Caption = ""

End Sub





Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI = Val(TRIM$(Text1.TEXT))
    If ECOARCH$("DEUDOR12", MKS$(NCLI)) <> "" Then
         Call NotaCli(NCLI, 2)
       Else
         Call MENSERR(24, "Ingrese el Número de Cliente.")
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '________________________________________________________________________________
    'validaciones si la ruta donde se escriben los remitos es valida y si hay derechos
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If (InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0) Or InStr(EPAC$, "QUILMES") > 0 Then
      CARPETASINCRO$ = TRIM$(Control("CONWALDB", "ESCRMOVI")) 'RUTA DONDE ESCRIBE LOS REMITOS CAMBACERES, QUILMES Y CONTEPLAST
      Ruta$ = CARPETASINCRO$
      If TieneDerechos(Ruta$) Then ' si la ruta es valida y hay derechos sigue de largo
        'MsgBox "La ruta tiene permisos de lectura y escritura."
      Else
        CARPETASINCROMODIFICADA$ = REPLACAR(CARPETASINCRO$, "\", "/")
        Call COMUNI("No se tienen permisos de lectura o escritura en la ruta: \" & CARPETASINCROMODIFICADA$ & "\Verifique si La Ruta Existe y si Puede Acceder")
        GoTo 99
      End If
    End If
    '________________________________________________________________________________
    
    
    If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
      Call COMUNI("ATENCION!!!! \ \  CORROBORAR LAS ETIQUETAS IDENTIFICATORIAS DE LAS BOBINAS!\ \")
'      If ValidarCantaRemitirVSPedido% < 1 Then 'se comento la llamada a pedido de Lautaro y Raul
'         GoTo 99                               ' y se agrego en la of, que no se puede fabricar mas del 3 % de lo proyectado
'      End If
    End If
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
      REMI1$ = TRIM$(Mid$(Text11, 4))
      REMI1$ = FORMATOCON0$(XVALO(REMI1$), 8)
      REMI1$ = "RT" & "-X-" + Label10.Caption + "-" + REMI1$
      '
      Dim ENV As New Envap
      EXISTREMI% = ENV.REMITO_EXISTE_EN_WALDBOTT(REMI1$)
      If EXISTREMI% > 0 Then
          Call COMUNI("Atención !!! \N° de Remito Existente en Base de Datos Waldbott\Consulte a su Soporte de Sistema")
          Set ENV = Nothing
          GoTo 99
      End If
      Set ENV = Nothing
      '
      NROPED& = XVALO(lblNroPed)
      If NROPED& > 0 Then
        MonedaPedido% = XVALO(VALOBADA("PEDDETA", "MoneEmis", "NROPED= " & NROPED&, "NOMESS"))
        If MonedaPedido% > 1 Then
          If XVALO(Label15) < 1 Then
             MONEDAORIGINAL$ = DEMONECO$(MonedaPedido%)
             OPX% = COMALTER%("Pedido a Remitir Realizado en Moneda Diferente a Pesos (" & MONEDAORIGINAL$ & ")\\ Cancelar Y Resolver Cambio de Moneda\Remitir En: " & MONEDAORIGINAL$ & "")
             If OPX% < 2 Then
                GoTo 99
             End If
          End If
        End If
      End If
    End If

    TOTITEMS% = Val(Label18)
    If TOTITEMS% > TOTITFO& Then
     Call MENSERR(24, "Renglones de Este Remito (" & TRIM$(Str$(TOTITEMS%)) & ") Superan\Máximo Permitido (" & TRIM$(Str$(TOTITFO&)) & ") del Formulario\de Impresión de Remito")
     GoTo 99
    End If

    ' TRANSFIERE LAS OBSERVACIONES
    'AGREGADO ENVAPLAST
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "OBSERVA" + "/FORMATO")))
    ATELIB% = XVALO(Mid$(FORTELIB$, 2))
       '
    II& = 0
    Dim OBSERX$()
    'SE REALIZA ESTE CAMBIO PARA CONSIDERAR
    'SI LAS CAJAS DE TEXTO DE LAS OBSERVACIONES TIENEN MAS TEXTO QUE LA CAPACIDAD VISIBLE
    'REALIZA UN FRATEXTO POR CADA CAJA CONSIDERANDO LA LONG. DEL PARAMETRO OBSERVA DEL FRM
    'ANTES SI EL TEXTO SUPERABA LA LONGITUD DEL PARAMETRO TRUNCABA
    For i& = 0 To 3
      Call FRATEXTO(Text16(i&), ATELIB%)
      For JJ& = 1 To CARETEX%
        II& = II& + 1
        ReDim Preserve OBSERX$(II&)
        OBSERX$(II&) = RETEX$(JJ&) 'CARGA EN UN VECTOR LAS OBSERVACIONES CON LA LONGITUD PARAMETRIZADA DEL FRM
      Next JJ&
    Next i&
    For U& = 1 To II&
       TTXXX$ = OBSERX$(U&)
       Call GRAREG("!OBSERVOF", TTXXX$, U&)
    Next U&
    'FIN CAMBIO
    '
    'CODIGO ANTERIOR
'    For U& = 1 To 4
'       TTXXX$ = TRIM$(Text16(U& - 1).TEXT)
'       Call GRAREG("!OBSERVOF", TTXXX$, U&)
'    Next U&
    Call SETULTREG("!OBSERVOF", U&)
    Call CIERRARCH("!OBSERVOF")
    '
    Call BLOQARCH("MOVISTO")
    NPRESENTADO& = Val(Mid(Text11, 4))
    NUPUVE$ = TRIM(Label10)
    NROPRO& = NPRESENTADO&
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    PREFIREM$ = TRIM(Label10)
    NROREMIT$ = TRIM$(Mid$(Text11, 4))
    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    FLIRE$ = Control$("", "FENUREFA")
    If FLIRE$ = "" Then FLIRE$ = "18/04/03"
    FLIMRE = FETEXCOR1$(FECHANUM(FLIRE$))
    '
    ' *** SE TRASLADO LA VERIFICACION DEL NUMERO A LA RUTINA PRIREMIFAC ***
    '
    ' AQUI VERIFICAR NUMERO YA UTILIZADO.
'    Call BLOQARCH("MOVISTO")
'    SQLX$ = "SELECT * FROM MOVISTO "
'    SQLX$ = SQLX$ + " WHERE DoPriLar = '" & NROREMIT$ & "'"
'    SQLX$ = SQLX$ + " AND FechMov >= '" & FLIMRE & "'"
'    Dim MOVVVISTO As ADODB.Recordset
'    Set MOVVVISTO = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    On Error Resume Next
'    With MOVVVISTO
'        .MoveFirst
'        If Err < 1 Then
'           On Error GoTo 0
'           Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número Libre.\  \Vuelva a Pedir la Impresión.")
'           NROPRO& = 0
'           Call NUREMIFA(NROPRO&, PREFIREM$)
'           Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
'           Command5.Enabled = True
'           Exit Sub
'        End If
'    End With
'    On Error GoTo 0
    '
    
10  Call PRIREMIFAC


    
    Call LIBARCH("MOVISTO")
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(Control$("", "DEPOEXPE"))
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    List1.Clear
    MSFDETAREMI.Rows = 1

    List1.Visible = False
    Command5.Enabled = False
    Command12.Enabled = False
    BOTNEXT.Enabled = True
    Frame15.Visible = False
    Text15 = ""
    Text17 = ""
    Text20 = ""
    PORBARRA% = 0 'MARCA DE CARGA POR CODIGO DE BARRA
    '
    Call LIACOBSE ' LIMPIA CASILLEROS INFERIORES - ACTIVA TEXT OBSERVACIONES
    
    '
    Unload FORMUL
    '
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    OPREMSTA07.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("VENGES07/FACTURA", "Gestión de Ventas y Deudores")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    FIKARDEX09.Show 1
'    Call GENKARDEX("")
'    AppActivate Caption
'    SendKeys Chr$(1)
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

 Sub DesRemCTP_Click()
        '__________________________________________________________________________________________
        'ESTO ES PARADO EN ENVAPLAST PARA RECORRER LOS ARCHIVOS Y METERLOS EN LA TABLA MOVISTOCTP
        '__________________________________________________________________________________________
        
        If InStr(1, EMPREAC$, "NEUQUEN") < 1 And InStr(1, EMPREAC$, "METALES ARCON") > 0 Then
           CARPETASINCRO$ = RutaCarpetaRemitosNeuquen$
           If Dir(CARPETASINCRO$, vbDirectory) = "" Then
               Call COMUNI("Falta Configuar Carpeta de Sincronizaciòn")
               Exit Sub
           End If
        Else
            CARPETASINCRO$ = TRIM$(Control("TRCTPAWL", "DROPTRRM"))
            If Dir(CARPETASINCRO$, vbDirectory) = "" Then
               Call COMUNI("Falta Configuar Carpeta de Sincronizaciòn")
               Exit Sub
            End If
        End If
        '
        If CARPETASINCRO$ <> "" Then
            'ASIGNO PROPIEDADES AL FILE1
                On Error Resume Next
                File1.Path = "C:\"
                On Error GoTo 0
                File1.Path = CARPETASINCRO$
                File1.Pattern = "*.txt"
                TX$ = ""
                List3.Clear
                'RECORRO EL FILE1 Y LE PONGO LA FECHA DEL ARCHIVO POR DELANTE PARA QUE ORDENE CORRECTAMENTE
                For n = 0 To File1.ListCount - 1
                  NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List3.AddItem TX$
                Next n
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For i& = 0 To List3.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List3.List(i&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List3.List(i&), POS1% + 2))
                  If Left$(TRIM$(UCase$(NombreArchivo$)), 5) = "RT_X_" Then
                    TABLAX$ = "MOVISTOCTP"
                    Call GrabarArchivoCTP(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt", TABLAX$) 'GRABA EN LA BASE DE DATOS SQL
                  End If
                Next i&

        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            Exit Sub
        End If
        Call COMUNI("Proceso Terminado")
        '
End Sub


Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    EPAC$ = TRIM$(EMPREAC$)
    mnuCot.Visible = True 'MENU PARA OBTENER COT
    Call InstalaFuenteAnchoFijo
    Refresh
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    TEXPE$ = "Remito de Ventas"
    If Control$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Caption = Caption + " - " + TEXPE$
    '
    '
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    If InStr(EPAC$, "AHP") > 0 Then
      Frame1.BackColor = &HE0E0E0
      Frame2.BackColor = &HE0E0E0
      Frame16.BackColor = &HE0E0E0
      Frame5.BackColor = &HE0E0E0
      Frame6.BackColor = &HE0E0E0
      Frame7.BackColor = &HE0E0E0
      Frame8.BackColor = &HE0E0E0
      Frame9.BackColor = &HE0E0E0
      Frame10.BackColor = &HE0E0E0
      Frame11.BackColor = &HE0E0E0
      Frame12.BackColor = &HE0E0E0
      Frame13.BackColor = &HE0E0E0
      Frame14.BackColor = &HE0E0E0
      Frame17.BackColor = &HE0E0E0
      Option1.BackColor = &HE0E0E0
      Option2.BackColor = &HE0E0E0
      mnuRegDesp.Visible = True
    End If
    '
    If InStr(EPAC$, "PERGAMINO") > 0 Then
       MSFDETAREMI.Visible = True
    End If
    
    X = MonedaY_TipoCambioDeRemitosSegunDoprinum(1, Moneda%, TIP#)
    If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
       Command19.Visible = True
       Label1(1).Visible = True
       lblNroPed.Visible = True
       Label1(3).Visible = True
       Label16.Visible = True
       Call CreaTablaMovistoCTP
    End If
    
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
       mnuRecuperarxRemito.Visible = True
       mnuTransfRemitoConteplastaWaldbot.Visible = True
       RemCTPAWDBT.Visible = True
       DesRemCTP.Visible = True
       RemContImp.Visible = True
       
    End If
    
    If InStr(EPAC$, "CAMBACERES") > 0 Then
       Command21.Visible = True
       mnuReimpRemitoQR.Visible = True
       Dim CTP As New Conteplast
       Call CTP.ACTUBASEWALDBOTT(1)
       DesRemCTP.Visible = True
       RemContImp.Visible = True

       Command21.Visible = False
    End If
    
'VALIDACION PARA AMBEST POR SI NO PAGA
'If InStr(EPAC$, "AMBEST") > 0 Then
'     FF% = FECHANUM("20/03/14")
'     If FF% <= HOY(HO$) Then
'       Call Form_Load
'     End If
'End If
    
10  If ULTREG("ECODEP") < 1 Then
        Call MENSERR(24, "Falta Definir Tabla de Depósitos")
        Call CARTADEPOSI
        GoTo 10
    End If
    '
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIT07.Show 1
        GoTo 10
    End If
    '\\\OT-06-0001-WAR-02/01/06///
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    TOTITFO& = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    '\\\OT-06-0001-WAR-FIN///
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMIT07.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(Control$("", "DEPOEXPE"))
    If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
       Call COMUNI("Error de Configuración:\  \Falta Definir Depósito de Expedición.")
       OPREMIT07.Show 1
       GoTo 10
    End If
    '
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    Text23.TEXT = Control$("", "COPIAREM")
    '
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Label10 = ""
    PREFIREM$ = TRIM$(Control$("", "PREFIREM"))
    If PREFIREM$ <> "*" Then
      Label10 = PREFIREM$
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "GUIDI") > 0 Then
       Command7.Visible = False
    End If
    '
    If InStr(1, EMPREAC$, "ARCON") > 0 Then
        Label11.Visible = True
        LABEL12.Visible = True
        Command18.Visible = True
        mnuEntregasNeuquenAvellaneda.Visible = True
    End If
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
       mnuEntregasNeuquenAvellaneda.Visible = False
       mnuEntregaCPedidoCompra.Visible = True
    End If
    If InStr(1, EMPREAC$, "NEUQUEN") < 1 And InStr(1, EMPREAC$, "METALES ARCON") > 0 Then
       'verifica si hay remitos de neuquen en el drop box y los pasa a la tabla
       Call CreaTablaMovistoCTP
       Call DesRemCTP_Click
    End If
    '
    If InStr(1, EMPREAC$, "RETEPLAST") < 1 Then
        mnuArmadoListaDespacho.Visible = False
        mnuAnulaListaDeDespacho.Visible = False
    End If
    '
    NROPRO& = 0
    Call NUREMIFA(NROPRO&, PREFIREM$)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '\\\OT-05-420-WAR-26/06/05///
    '\\\OT-05-420-WAR-FIN///
    '
    Set FORMUL = BCODE07
'    Call RECUREMIT     ' RECUPERA PRECIOS DE REMITOS GUIDI
    '
    Call CREAR_STORED_PROCEDURE
    Timer1.Enabled = True
    Refresh
'    UTILIZA_SKIN% = 1
'    Call APLICACIONSKINS(Me, Picture2)


End Sub

Sub RECUREMIT()
    '
    Screen.MousePointer = 11
    FE% = FECHANUM("13/05/09")
    FECHA = FETEXCOR1$(FE%)
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CODIMOVI= 'RT' "
    SQLX$ = SQLX$ + " AND NUPEDCLI > 0 AND ABS(VALOUNIT) < 0.005 "
    SQLX$ = SQLX$ + " AND FECHMOV > '" & FECHA & "'"
    Dim MOVTMP As ADODB.Recordset
    Set MOVTMP = New ADODB.Recordset
    MOVTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With MOVTMP
      Do While Not .EOF
        CI% = XVALO(!cod_inte)
        NPED& = XVALO(!NUPEDCLI)
        NCLIE% = XVALO(!Nume_Clie)
        CONDI$ = "CODIINT=" & CI% & " AND NROPED=" & NPED& & " AND NROCLIE=" & NCLIE% & ""
        A = CantRegistros&("PEDDETA", CONDI$)
        
        If A = 1 Then
          Dato = XVALO(VALOBADA("PEDDETA", "PREUNID", CONDI$, "NOMESS"))
          !VALOUNIT = Dato
          .Update
        End If
        .MoveNext
      Loop
    End With
    MOVTMP.Close
    Set MOVTMP = Nothing
    Screen.MousePointer = 1
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Cancel = 0
   Screen.MousePointer = 1
   Call FINAL("")
End Sub

Private Sub Frame8_DblClick()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If (InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0) Or InStr(EPAC$, "QUILMES") > 0 Then

      CARPETASINCRO$ = TRIM$(Control("CONWALDB", "ESCRMOVI")) 'RUTA DONDE ESCRIBE LOS REMITOS CAMBACERES, QUILMES Y CONTEPLAST
      Ruta$ = CARPETASINCRO$
      If TieneDerechos(Ruta$) Then
        MsgBox "La ruta tiene permisos de lectura y escritura."
      Else
        Call COMUNI("No se tienen permisos de lectura y/o escritura en la ruta:" & CARPETASINCRO$ & "\Verifique si La Ruta Existe y si Puede Acceder")
        GoTo 99
      End If
    End If
99
End Sub

'\\\OT-05-0419-WAR-26/06/05///
Private Sub Label10_dblClick()
  TALO$ = TRIM$(Label10)
'  Select Case TALO$
'    Case ""
'        Label10 = "0001"
'    Case "0001"
'        Label10 = "0002"
'    Case "0002"
'        Label10 = "0001"
'  End Select
  Select Case TALO$
    Case ""
        Label10 = "0001"
    Case "0001"
        Label10 = "0002"
    Case "0002"
        Label10 = "0003"
    Case "0003"
        Label10 = "0004"
    Case "0004"
        Label10 = "0005"
    Case "0005"
        Label10 = "0006"
    Case "0006"
        Label10 = "0007"
    Case "0007"
        Label10 = "0008"
    Case "0008"
        Label10 = "0009"
    Case "0009"
        Label10 = "0001"
    Case Else
        Label10 = "0001"
  End Select
  
  NUPUVE$ = TRIM(Label10)
  CMO$ = "RT"
  NROPRO& = 0
  Call NUREMIFA(NROPRO&, NUPUVE$)
  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
End Sub


Private Sub Label18_Change()
   '\\\OT-06-0001-WAR-02/01/06///
'SE MOVIO AL BOTON DE GRABACION
'   TOTITEMS% = Val(Label18)
'   If TOTITEMS% > TOTITFO& Then
'     Call MENSERR(24, "Renglones de Este Remito (" & TRIM$(Str$(TOTITEMS%)) & ") Superan\Máximo Permitido (" & TRIM$(Str$(TOTITFO&)) & ") del Formulario\de Impresión de Remito")
'   End If
   '
End Sub

'\\\OT-05-0419-WAR-FIN///
Private Sub Label2_DblClick()
'    OPREMIFAC.Show 1
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI = NUMECLI(CPROX$)
     If NCXI > 0 Then
       If REGICLI&(NCXI) > 0 Then
         Label26 = CPROX$
         Text1.TEXT = TRIM$(Str$(NCXI))
         Text2.TEXT = rasocli$(NCXI)
         Text3.TEXT = DOMICLI$(NCXI)
         Text4.TEXT = CPOSCLI$(NCXI)
         Text5.TEXT = LOCACLI$(NCXI)
         Call CARTRANSPOR
       End If
     End If
   End If
End Sub



Private Sub Label9_DblClick()
   Frame15.Visible = True
   List1.Visible = True
End Sub

Private Sub lblNroPed_Change()
    NROPED& = XVALO(Me.lblNroPed)
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    For i& = 0 To 3
       Text16(i&) = ""
    Next i&
    If NROPED& > 0 Then
        Text14 = VALOBADA("PEDENCA", "DOMIENT", "NROPED= " & NROPED&, "NOMESS")
        Text21 = VALOBADA("PEDENCA", "LOCAENT", "NROPED= " & NROPED&, "NOMESS")
        OBS$ = VALOBADA("PEDENCA", "OBSERVA", "NROPED= " & NROPED&, "NOMESS")
        
        Call FRATEXTO(OBS$, 43)
        For JJ& = 1 To CARETEX%
            If JJ& >= 4 Then
              Text16(3) = Text16(3) & RETEX$(JJ&)
            Else
              Text16(JJ& - 1) = RETEX$(JJ&)
            End If
        Next JJ&
    End If
    
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
      MonedaPedido% = XVALO(VALOBADA("PEDDETA", "MoneEmis", "NROPED= " & NROPED&, "NOMESS"))
      If MonedaPedido% > 1 Then
        
        Label16 = ECOARCH$("TAMONED", Chr$(MonedaPedido%))
        OPX% = COMALTER%("Pedido Generado en Moneda " & DEMONECO$(MonedaPedido%) & "\\ Generar Remito Manteniendo Moneda Original\Generar Remito en Pesos")
        If OPX% = 2 Then
          TipoCambioPedido# = XVALO(VALOBADA("PEDDETA", "TipoCam", "NROPED= " & NROPED&, "NOMESS"))
          Label15.Visible = True
          Label1(3).Visible = True
          Label16.Visible = True
          Label1(2).Visible = True
          TCP$ = FORMATNUM$(TipoCambioPedido#, "F9.2")
20        ImporteIngresado$ = TRIM$(InputBox$("Ingrese Cotización de Moneda (" & DEMONECO$(MonedaPedido%) & ")", " Cotización", TCP$))
          If ImporteIngresado$ = "" Or XVALO(ImporteIngresado$) < 1 Then GoTo 20 'SI NO INGRESA VUELVE SI O SI TIENE QUE INGRESAR UN VALOR
          'SI INGRESO VALOR LO PONE EN EL LABEL
          Label15 = FORMATNUM$(XVALO(ImporteIngresado$), "F9.2")
        End If
      Else
      
      Label15.Visible = False
      Label1(2).Visible = False
      Label15.Caption = ""
      Label1(3).Visible = False
      Label16.Visible = False: Label16.Caption = ""
     End If
    End If
    '
    'PARA QUE IMPRIMA LO O.COMPRA DE LA COMPRA INTERNA.
    ES_RIVA_CAMBA_QUIL% = 0
    If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then ES_RIVA_CAMBA_QUIL% = 1
    If ES_RIVA_CAMBA_QUIL% > 0 And NROPED& > 0 Then
       StringNroCtes$ = StringDeNroClientesComprasInternas$
       'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE CLIENTES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
       If NumeroEstaEnCadena(StringNroCtes$, XVALO(Text1)) Then
         CONDI$ = "NPED_GENERADO_EN_AVELLANEDA = " & NROPED&
         Text20 = XVALO(VALOBADA("OC_SUCURSAL", "O_COMPRA_SUC", CONDI$, "NOMESS"))
      End If
    End If
End Sub

Private Sub List1_DblClick()
    '
    If Command12.Enabled = False Then
       Call MENSERR(24, "Imposible Re-Editar Despacho.\Cancele y Vuelva a Cargar.")
       Exit Sub
    End If
    '
    If PORBARRA% > 0 Then Call CARGALISTBARRA: GoTo 25

    Call SIGUEREMI2
25  If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        MSFDETAREMI.Rows = 1

        Unload FORMUL '\\\OT-06-0182-WAR-29/05/06///
        List1.Visible = False
        Command5.Enabled = False
        Command12.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Label18 = "" '*
        Exit Sub
    End If
    List1.Clear
    MSFDETAREMI.Rows = 1

    List1.Visible = True
    v& = 0
    'TOTITEMS% = 0 '*
    NUOCOMPRA$ = ""
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        Reli$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            v& = v& + 1
            Call GRAREG("!BOLREDET", X$, v&)
            Call REPLA(Reli$, CODEXT$(CI%), 1, 15)
            Call REPLA(Reli$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(Reli$, Unimed$(CI%), 49, 4)
            'CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            CANS$ = CVD(Mid$(X$, 89, 8))
            CANS$ = FORMATNUM$(XVALO(CANS$), "F10.1")
            Call REPLA(Reli$, CANS$, 53, 10)
            Call REPLA(Reli$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(Reli$, NOPE$, 78, 7)
            NPEDID& = CVS(Mid$(X$, 151, 4))
            NUOCO$ = TRIM$(NUOCOCLI$(NPEDID&, CI%))
            If NUOCO$ <> "" Then
              If InStr(NUOCOMPRA$, " " + NUOCO$ + ",") < 1 Then
                NUOCOMPRA$ = NUOCOMPRA$ + " " + NUOCO$ + ","
              End If
            End If
            List1.AddItem Reli$
            Call CargarFlexgrid(Reli$, v&)
            'TOTITEMS% = TOTITEMS% + 1 '*
        End If
    Next U&
    'Label18 = TRIM(Str$(TOTITEMS%))
    NUOCOMPRA$ = TRIM$(NUOCOMPRA$)
    If Right$(NUOCOMPRA$, 1) = "," Then
      NUOCOMPRA$ = Left$(NUOCOMPRA$, Len(NUOCOMPRA$) - 1)
    End If
    Text20 = NUOCOMPRA$
    Call SETULTREG("!BOLREDET", v&)
    Screen.MousePointer = 1
    '
    Command5.Enabled = True
    Text15.SetFocus
    '
End Sub

Sub CargarFlexgrid(Reli$, i&)
            If i& <= 1 Then
                Campos$ = "Código/A15;Descripción/A30;U.M./A4;Cantidad/F10.1;Lote Color/A16;N.Ped./F8"
                Call CARGA_ENCABEZADO(FOREMIT07.MSFDETAREMI, Campos$, Me)
             End If
            '
            Codigo$ = Mid$(Reli$, 1, 15)
            Descripcion$ = Mid$(Reli$, 16, 30)
            UNIDADMED$ = Mid$(Reli$, 49, 4)
            CANTSALID$ = Mid$(Reli$, 53, 10)
            UNIDADMED$ = Mid$(Reli$, 49, 4)
            NOPE$ = Mid$(Reli$, 78, 7)
            '
            MSFDETAREMI.Rows = i& + 1
            MSFDETAREMI.TextMatrix(i&, 1) = Codigo$
            MSFDETAREMI.TextMatrix(i&, 2) = Descripcion$
            MSFDETAREMI.TextMatrix(i&, 3) = UNIDADMED$
            MSFDETAREMI.TextMatrix(i&, 4) = CANTSALID$
            MSFDETAREMI.TextMatrix(i&, 5) = "" 'lote/color
            MSFDETAREMI.TextMatrix(i&, 6) = NOPE$
End Sub







Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then
         For i% = 0 To List1.ListCount - 1
           If List1.Selected(i%) Then
            TXTX$ = List1.TEXT
            Codigo$ = TRIM$(Mid$(TXTX$, 1, 15))
            DESCRIX$ = TRIM$(Mid$(TXTX$, 16, 32))
            CI1% = XVALO(CODINT(Codigo$))
            lblNroPed = XVALO(Mid(TXTX$, 78, 7))
            OBSERVAC.Caption = "Observaciones del Item"
            OBSERVAC.Label24.Caption = Codigo$ & " - - " & DESCRIX$
            OBSERVAC.Label2 = 256
            OBSERVAC.Text28.MaxLength = 256
            OBSERVAC.Text28.Enabled = True
            OBSERVAC.Text28 = VALOBADA("PEDDETA", "ObservaItem", "NROPED = " & XVALO(lblNroPed) & " AND CODIINT = " & CI1%, "NOMESS")
           End If
         Next i%
         OBSERVAC.Show 1
         If OBSERVAC.Label1.Caption = "OK" Then
            Call ACTUREGISTRO("PEDDETA", "ObservaItem", "NROPED = " & XVALO(lblNroPed) & " AND CODIINT = " & CI1%, OBSERVAC.Text28, REGAF&, "NOMESS")
         End If
     End If

End Sub

Private Sub mnuEntregaCPedidoCompra_Click()
   Call FINAL("?MOVISTOCTP")
End Sub

Private Sub mnuEntregasNeuquenAvellaneda_Click()
10
    'Cta Razon Social    FECHA ENTREGA   Pedido  Código  Descripción Cant
    Call CreaTablaMovistoCTP
    
    Campos$ = "Cta/F7" '1
    Campos$ = Campos$ + ";Razon Social/A32" '2
    Campos$ = Campos$ + ";FECHA ENTREGA/D8" '3
    Campos$ = Campos$ + ";Pedido/F9.0" '4
    Campos$ = Campos$ + ";Código/A24" '5
    Campos$ = Campos$ + ";Descripción/A36" '6
    Campos$ = Campos$ + ";Cant/F9.1" '7
    Campos$ = Campos$ + ";Domicilio Entrega/A32" '8
    Campos$ = Campos$ + ";Remito/A18" '9
    Campos$ = Campos$ + ";Origen/A10" '10
    Campos$ = Campos$ + ";Emis.Ped./D8" '11
    

15  Call DESHASFECHA(Des%, Has%, PERIO$)
    If TRIM$(PERIO$) = "" Then Exit Sub
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)

    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0

    
     SQLX$ = "SELECT RASO_CLI,descrip,cod_inte,Cod_Exte,NuOrItem,CANTSALID,CANTINGRE,NUMEORDPED,NUPEDCLI,VALOUNIT"
     SQLX$ = SQLX$ + "  ,Nume_Clie,FECHMOV,idenlote,DOPRILAR,DOMIENT, 'NEUQUEN' AS origen"
     SQLX$ = SQLX$ + "  from MOVISTOctp WITH(NOLOCK) "
     SQLX$ = SQLX$ + "  WHERE COD_INTE > 0 AND CODIMOVI = 'RT' AND FECHMOV >= '" & Desde & "' "
     SQLX$ = SQLX$ + "  AND FECHMOV <= '" & Hasta & "' AND CANTSALID > 0"
     SQLX$ = SQLX$ + "  Union"
     SQLX$ = SQLX$ + "  SELECT (SELECT DEUDOR12.Raso_Cli FROM DEUDOR12 WHERE DEUDOR12.Nume_Cli = MOVISTO.Nume_Clie)  as RASO_CLI"
     SQLX$ = SQLX$ + "  ,descrip,cod_inte,Cod_Exte,NuOrItem,CANTSALID,CANTINGRE,NUMEORDPED,NUPEDCLI,VALOUNIT"
     SQLX$ = SQLX$ + "  ,Nume_Clie,FECHMOV,idenlote,DOPRILAR,(SELECT PEDENCA.DomiEnt FROM PEDENCA WHERE PEDENCA.NROPED = MOVISTO.NUPEDCLI)  as DOMIENT, 'AVELLANEDA' AS origen"
    SQLX$ = SQLX$ + " FROM movisto WHERE COD_INTE > 0 AND CODIMOVI = 'RT' AND FECHMOV >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FECHMOV <= '" & Hasta & "' AND CANTSALID > 0 AND DOPRILAR LIKE 'RT-X-%'"
    SQLX$ = SQLX$ + " ORDER BY FECHMOV,DOPRILAR"
    Set rs = READSET(SQLX$)
    
    With rs
    Do While Not .EOF
     'Cta Razon Social    FECHA ENTREGA   Pedido  Código  Descripción Cant
        NC = XVALO(!Nume_Clie)
        razonSocial = SAFETEXT$(!raso_cli)
        FECHMOV$ = FECHATEX$(CVINT%(!FECHMOV))
        NROPED& = XVALO(!NUPEDCLI)
        FechEmix$ = FECHATEX$(CVINT%(FechaEmisionSegunPedido$(NROPED&)))
        Codigo$ = SAFETEXT$(!Cod_Exte)
        Descripcion$ = SAFETEXT$(!Descrip)
        CANTI = FORMATNUM$(XVALO(!CANTSALID), "F10.1")
        DomEntrega$ = SAFETEXT$(!DOMIENT)
        
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = NC
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = razonSocial
        
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = FECHMOV$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = NROPED&
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 6) = Descripcion$
        GRILLAGRAL.GRID.TextMatrix(J&, 7) = CANTI
        GRILLAGRAL.GRID.TextMatrix(J&, 8) = DomEntrega
        GRILLAGRAL.GRID.TextMatrix(J&, 9) = SAFETEXT$(!DOPRILAR)
        GRILLAGRAL.GRID.TextMatrix(J&, 10) = SAFETEXT$(!Origen)
        GRILLAGRAL.GRID.TextMatrix(J&, 11) = FechEmix$
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1

   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   TX$ = ""
  
   
   GRILLAGRAL.Caption = "Entregas Neuquén/Avellaneda"
   
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2

        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENTNEUAV", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        Exit Sub
   End If

End Sub

Private Sub mnuArchivo_Click()

End Sub

Private Sub mnuAnulaListaDeDespacho_Click()

    Dim p As New Pedidos
    
    PedidoElegido& = p.ListaPedidosPendientes(0)
    If PedidoElegido& > 0 Then
        On Error GoTo ControlError
        FLEXCONN.BeginTrans
        
        FLEXCONN.Execute "update pedenca set status=9 where nroped=" & CStr(PedidoElegido&)
        FLEXCONN.Execute "update peddeta set status=9 where nroped=" & CStr(PedidoElegido&)
        FLEXCONN.Execute "update bultosfabric set pedido=0 where pedido=" & CStr(PedidoElegido&)
        
        FLEXCONN.CommitTrans
        On Error GoTo 0
        
        Call COMUNI("Lista de Despacho Anulada")
    End If
    
    Exit Sub
    
ControlError:
    FLEXCONN.RollbackTrans
    On Error GoTo 0
    Call COMUNI("Fallo al Anula Lista de Despacho")
    
End Sub

'Private Sub mnuArmadoListaDespacho_Click()
'
'    Dim rstOrdenes As ADODB.Recordset
'    Dim rstPallets As ADODB.Recordset
'    Dim ListaItemsPicking() As ItemPicking
'    ReDim Preserve ListaItemsPicking(1)
'    CantidadElementos& = 0
'
'    ' ***  ELIGE LA ORDEN DE FABRICACION / PARTIDA A DESPACHAR
'5   ConsultaSql$ = "select orfa.idsubor, orfa.idenlote, orfa.numerochasis subcodigo, orfa.cod_exte, orfa.descrip from" _
'    & " ordefabr orfa with(nolock) inner join (select distinct ordenfabricacion from bultosfabric with(nolock) where idpallet>0 and status<>9 and (pedido=0 or pedido is null)) bul" _
'    & " on (orfa.idsubor=bul.ordenfabricacion)" _
'    & " order by orfa.idsubor"
'    Set rstOrdenes = READSET(ConsultaSql$)
'
'    Campos$ = "O-Fabricación/A16;Partida/A24;SubCodigo/A16;Codigo/A16;Descripción/A32"
'    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'    FRMZELECHO.Caption = "Lista de Partidas Con Saldo a Despachar"
'
'    Call Carga_MSF_Recordset(rstOrdenes, Campos$, FRMZELECHO.msf2)
'
'    FRMZELECHO.Width = 13000
'    FRMZELECHO.Show 1
'
'    OrdenFabricacion$ = TRIM(RESP_ZELE_VECT(1))
'    If OrdenFabricacion$ = "" Then GoTo 50
'    Partida$ = TRIM$(RESP_ZELE_VECT(2))
'    CodigoInterno% = CODINT%(CStr(RESP_ZELE_VECT(4)))
'
'    ' *** ELIGE EL PALLET-BULTO A DESPACHAR
'10  OpcionDeFiltro% = COMALTER%("Opciones de Filtro:\\Partida\Pallet\Bulto")
'
'    If OpcionDeFiltro% = 1 Then
'
'        ConsultaSql$ = "select idpallet, idbulto, cantidadbulto from bultosfabric with(nolock)" _
'        & " where ordenfabricacion='" & OrdenFabricacion$ & "' and idpallet>0 and status<>9 and (pedido = 0 or pedido is null)"
'        If CantidadElementos& > 0 Then
'            For i = 0 To UBound(ListaItemsPicking) - 1
'                If ListaItemsPicking(i).OrdenFabricacion = OrdenFabricacion$ Then
'                    ConsultaSql$ = ConsultaSql$ & " and not (idpallet=" & ListaItemsPicking(i).Pallet & " and idbulto=" & ListaItemsPicking(i).Bulto & ")"
'                End If
'            Next i
'        End If
'        Set rstPartida = READSET(ConsultaSql$)
'
'        ' *** GUARDA LA COMBINACION OF-PALLET-BULTO
'        With rstPartida
'            Do While Not .EOF
'                CantidadElementos& = CantidadElementos& + 1
'                ReDim Preserve ListaItemsPicking(CantidadElementos&)
'
'                Pallet& = !IdPallet
'                Bulto& = !IdBulto
'                CantidadBulto# = !CantidadBulto
'                With ListaItemsPicking(CantidadElementos& - 1)
'                    .OrdenFabricacion = OrdenFabricacion$
'                    .Partida = Partida$
'                    .CodigoInterno = CodigoInterno%
'                    .Pallet = Pallet&
'                    .Bulto = Bulto&
'                    .Cantidad = CantidadBulto#
'                End With
'
'                .MoveNext
'            Loop
'        End With
'
'        On Error Resume Next
'        rstPartida.Close
'        Set rstPartida = Nothing
'        On Error GoTo 0
'
'    ElseIf OpcionDeFiltro% = 2 Then
'
'        Campos$ = "O-Fabricacion/A18;Partida/A18;Referencia/A40;Pallet/I10;Bultos/I10"
'        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
'        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'        FRMZELECHO.Caption = "Lista de Pallet para Baja"
'
'        Dim Rst As ADODB.Recordset
'        ConsultaSql$ = "select b.ordenfabricacion, o.idenlote, o.referen, b.idpallet, count(*) bultos from bultosfabric b with(nolock)" _
'        & " inner join ordefabr o with(nolock)" _
'        & " on (b.ordenfabricacion=o.idsubor)" _
'        & " where b.status<>9 and b.pedido=0 and ordenfabricacion='" & OrdenFabricacion$ & "'" _
'        & " group by b.ordenfabricacion, o.idenlote, o.referen, b.idpallet" _
'        & " order by b.ordenfabricacion desc, b.idpallet asc"
'        Set rstPallet = READSET(ConsultaSql$)
'        Call Carga_MSF_Recordset(rstPallet, Campos$, FRMZELECHO.msf2)
'
'        On Error Resume Next
'        rstPallet.Close
'        Set rstPallet = Nothing
'        On Error GoTo 0
'
'        FRMZELECHO.Width = 12500
'        FRMZELECHO.Show 1
'
'        OrdenFabricacion$ = TRIM(RESP_ZELE_VECT(1))
'        If OrdenFabricacion$ = "" Then GoTo 5
'        Pallet& = XVALO(RESP_ZELE_VECT(4))
'
'        ConsultaSql$ = "select idpallet, idbulto, cantidadbulto from bultosfabric with(nolock)" _
'        & " where ordenfabricacion='" & OrdenFabricacion$ & "' and idpallet=" & CStr(Pallet&) & " and status<>9 and (pedido = 0 or pedido is null)"
'        If CantidadElementos& > 0 Then
'            For i = 0 To UBound(ListaItemsPicking) - 1
'                If ListaItemsPicking(i).OrdenFabricacion = OrdenFabricacion$ And ListaItemsPicking(i).Pallet = Pallet& Then
'                    ConsultaSql$ = ConsultaSql$ & " and idbulto<>" & ListaItemsPicking(i).Bulto
'                End If
'            Next i
'        End If
'        Set rstBulto = READSET(ConsultaSql$)
'
'        ' *** GUARDA LA COMBINACION OF-PALLET-BULTO
'        With rstBulto
'            Do While Not .EOF
'                CantidadElementos& = CantidadElementos& + 1
'                ReDim Preserve ListaItemsPicking(CantidadElementos&)
'
'                Pallet& = !IdPallet
'                Bulto& = !IdBulto
'                CantidadBulto# = !CantidadBulto
'                With ListaItemsPicking(CantidadElementos& - 1)
'                    .OrdenFabricacion = OrdenFabricacion$
'                    .Partida = Partida$
'                    .CodigoInterno = CodigoInterno%
'                    .Pallet = Pallet&
'                    .Bulto = Bulto&
'                    .Cantidad = CantidadBulto#
'                End With
'
'                .MoveNext
'            Loop
'        End With
'
'        On Error Resume Next
'        rstBulto.Close
'        Set rstBulto = Nothing
'        On Error GoTo 0
'
'    Else
'
'40      ConsultaSql$ = "select idpallet, idbulto, cantidadbulto from bultosfabric with(nolock)" _
'        & " where ordenfabricacion='" & OrdenFabricacion$ & "' and idpallet>0 and status<>9 and (pedido = 0 or pedido is null)"
'        If CantidadElementos& > 0 Then
'            For i = 0 To UBound(ListaItemsPicking) - 1
'                If ListaItemsPicking(i).OrdenFabricacion = OrdenFabricacion$ Then
'                    ConsultaSql$ = ConsultaSql$ & " and not (idpallet=" & ListaItemsPicking(i).Pallet & " and idbulto=" & ListaItemsPicking(i).Bulto & ")"
'                End If
'            Next i
'        End If
'        Set rstBulto = READSET(ConsultaSql$)
'
'        If rstBulto.BOF And rstBulto.EOF Then
'            On Error Resume Next
'            rstBulto.Close
'            Set rstBulto = Nothing
'            On Error GoTo 0
'            Call COMUNI("No Hay Mas Pallets Disponibles")
'            GoTo 5
'        End If
'
'        Campos$ = "Pallet/I10;Bulto/I10;Cant. Bulto/F12.1"
'        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
'        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'        FRMZELECHO.Caption = "Partida: " & Partida$ & " - Lista de Pallets/Bultos Disponibles"
'
'        Call Carga_MSF_Recordset(rstBulto, Campos$, FRMZELECHO.msf2)
'
'        On Error Resume Next
'        rstBulto.Close
'        Set rstBulto = Nothing
'        On Error GoTo 0
'
'        FRMZELECHO.Width = 6000
'        FRMZELECHO.Show 1
'
'        Pallet& = XVALO(RESP_ZELE_VECT(1))
'        If Pallet& = 0 Then GoTo 5
'
'        Bulto& = XVALO(RESP_ZELE_VECT(2))
'        CantidadBulto# = XVALO(RESP_ZELE_VECT(3))
'
'        ' *** GUARDA LA COMBINACION OF-PALLET-BULTO
'        If CantidadElementos& = 0 Then
'            CantidadElementos& = 1
'        Else
'            CantidadElementos& = CantidadElementos& + 1
'            ReDim Preserve ListaItemsPicking(CantidadElementos&)
'        End If
'
'        With ListaItemsPicking(CantidadElementos& - 1)
'            .OrdenFabricacion = OrdenFabricacion$
'            .Partida = Partida$
'            .CodigoInterno = CodigoInterno%
'            .Pallet = Pallet&
'            .Bulto = Bulto&
'            .Cantidad = CantidadBulto#
'        End With
'
'        GoTo 40
'
'    End If
'
'    GoTo 5
'
'    ' GUARDA LA SELECCION DE OF-PALLET-BULTO
'50  If CantidadElementos& > 0 Then
'        ' *** MUESTRA RESULTADO DE LA SELECCION
'        Campos$ = "O-Fabricacion/A0;Partida/A24;Codigo/A24;Descripcion/A32;Pallet/I6;Bulto/I6;Cant. Bulto/F0;Codigo/I0"
'        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
'        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'        FRMZELECHO.Caption = "Lista de Partidas/Pallets/Bultos Seleccionados"
'
'        For i = 1 To UBound(ListaItemsPicking)
'            FRMZELECHO.msf2.Rows = FRMZELECHO.msf2.Rows + 1
'
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 1) = ListaItemsPicking(i - 1).OrdenFabricacion
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 2) = ListaItemsPicking(i - 1).Partida
'
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 3) = CODEXT$(ListaItemsPicking(i - 1).CodigoInterno)
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 4) = DESCRIT0$(ListaItemsPicking(i - 1).CodigoInterno)
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 5) = CStr(ListaItemsPicking(i - 1).Pallet)
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 6) = CStr(ListaItemsPicking(i - 1).Bulto)
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 7) = TRIM$(FORMATNUM$(ListaItemsPicking(i - 1).Cantidad, "F10.1"))
'            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 8) = CStr(ListaItemsPicking(i - 1).CodigoInterno)
'        Next i
'
'        FRMZELECHO.Width = 12000
'        FRMZELECHO.Show 1
'
'        ' *** GRABA ITEMS A DESPACHAR COMO PEDIDO
'        Dim AcumulaItemsPicking() As Single
'        Opcion% = COMALTER%("¿Registra Lista de Despacho?\\Si, Registrar\Continuar Armado\No, Cancelar")
'        If Opcion% = 1 Then
'
'            NumeroCliente& = ListaSeleccionClientes
'            If NumeroCliente& = 0 Then GoTo 50
'
'            ReDim Preserve AcumulaItemsPicking(NUMAS%)
'
'            MaximoNumeroInterno% = 0
'            For i = 1 To UBound(ListaItemsPicking)
'                AcumulaItemsPicking(ListaItemsPicking(i - 1).CodigoInterno) = AcumulaItemsPicking(ListaItemsPicking(i - 1).CodigoInterno) + ListaItemsPicking(i - 1).Cantidad
'                If MaximoNumeroInterno% < ListaItemsPicking(i - 1).CodigoInterno Then
'                    MaximoNumeroInterno% = ListaItemsPicking(i - 1).CodigoInterno
'                End If
'            Next i
'
'            Dim Pedido As New Pedidos
'            Dim e As New PedidoEncabezado
'            Dim d() As PedidoDetalle
'            ReDim Preserve d(1) As PedidoDetalle
'            Dim CantidadItemDetalle As Long
'
'            e.FechaEmision = HOY(HO$)
'            e.NumeroCliente = NumeroCliente&
'            e.FechaSolicitudEntrega = HOY(HO$)
'            e.TipoPedido = "1"
'            e.NUMERO = -1
'
'            CantidadItemDetalle& = 0
'            Dim Item As PedidoDetalle
'            For i = 1 To MaximoNumeroInterno%
'
'                If AcumulaItemsPicking(i) > 0.05 Then
'                    Set Item = New PedidoDetalle
'
'                    Item.CodigoInterno = i
'                    Item.Descripcion = DESCRIT0$(CInt(i))
'                    Item.CantidadPedida = AcumulaItemsPicking(i)
'                    Item.CantidadSaldo = AcumulaItemsPicking(i)
'                    Item.COLOR = ""
'                    Item.TALLE = ""
'                    Item.Largo = 0
'                    Item.ANCHO = 0
'                    Item.Espesor = 0
'                    Item.FechaSolicitudEntrega = HOY(HO$)
'
'                    If CantidadItemDetalle& = 0 Then
'                        CantidadItemDetalle& = 1
'                    Else
'                        CantidadItemDetalle& = CantidadItemDetalle& + 1
'                        ReDim Preserve d(CantidadItemDetalle&) As PedidoDetalle
'                    End If
'                    Set d(CantidadItemDetalle& - 1) = Item
'                End If
'
'            Next i
'            NumeroPedidoDevuelto& = Pedido.Nuevo(e, d)
'
'            ' *** VINCULA OF-PALLET-BULTO CON EL PEDIDO
'            If NumeroPedidoDevuelto& > 0 Then
'                For i = 1 To UBound(ListaItemsPicking)
'                    ConsultaSql$ = "update bultosfabric set pedido=" & CStr(NumeroPedidoDevuelto&) & " where" _
'                        & " ordenfabricacion='" & ListaItemsPicking(i - 1).OrdenFabricacion & "'" _
'                        & " and idpallet=" & ListaItemsPicking(i - 1).Pallet _
'                        & " and idbulto=" & ListaItemsPicking(i - 1).Bulto

Private Sub mnuEntregasTotalKG_Click()
          Screen.MousePointer = 11
          
5         Call DESHASFECHA(Des%, Has%, PERIO$)
          If TRIM$(PERIO$) = "" Then GoTo 99
          Call APERBASDAT("STOCKS")
          '
        
          '
          DESD1$ = FETEXCOR1$(Des%)
          HAST1$ = FETEXCOR1$(Has%)
          
        Campos$ = "Fecha/D8;" '1
        Campos$ = Campos$ + "Documento/A18;" '2
        Campos$ = Campos$ + "Cta/F7;" '3
        Campos$ = Campos$ + "R.Social/A24;" '4
        Campos$ = Campos$ + "Nro.Pedido/F8;" '5
        Campos$ = Campos$ + "O.Fabricación/A16;" '6
        Campos$ = Campos$ + "CI/A0;" '7
        Campos$ = Campos$ + "Código/A20;"      '8
        Campos$ = Campos$ + "Descripción/A32;"      '9
        Campos$ = Campos$ + "Cantidad/F10;"   '10
        Campos$ = Campos$ + "Ancho/F12.1;"       '11
        Campos$ = Campos$ + "Gramaje/F12.1;"        '12
        Campos$ = Campos$ + "Corte/F12.1;" '13
        Campos$ = Campos$ + "Peso Unitario/F12.1;"     '14((ANCHO * 2 ) * GRAMAJE * CONSUMO_X_CORTE) /100
        Campos$ = Campos$ + "Kg. Carga Total/F12.1" '15  Peso Unitario * CANTSALID
    
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        SQLX$ = " SELECT M.FechMov,DopriCor,Nume_Clie, NuPedCli,M.Cod_Inte, M.Cod_Exte,M.Descrip,CantSalid, MA.ANCHO,MA.GRAMAJE,MA.Consumo_X_Corte,"
        SQLX$ = SQLX$ + " ((MA.ANCHO * 2) * MA.GRAMAJE * MA.Consumo_X_Corte) / 1000 AS Resultado1,"
        SQLX$ = SQLX$ + " (((MA.ANCHO * 2) * MA.GRAMAJE * MA.Consumo_X_Corte) / 1000) * CantSalid AS Resultado2"
        SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK) "
        SQLX$ = SQLX$ + " INNER JOIN MASTER MA ON MA.Codint = M.COD_INTE"
'        SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR O ON O.NROPED = M.NuPedCli"
        SQLX$ = SQLX$ + "     WHERE M.codimovi = 'RT'"
        SQLX$ = SQLX$ + "     AND M.CANTSALID > 0"
        SQLX$ = SQLX$ + "     and M.FechMov >= '" & DESD1$ & "' AND M.FECHMOV <= '" & HAST1$ & "'"

        SQLX$ = SQLX$ + " GROUP BY  M.FechMov,DopriCor,Nume_Clie, NuPedCli,M.Cod_Inte, M.Cod_Exte,M.Descrip,CantSalid, MA.ANCHO,MA.GRAMAJE,MA.Consumo_X_Corte"
        SQLX$ = SQLX$ + " ORDER BY M.FechMov"

'        SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) M"
'        SQLX$ = SQLX$ + "     WHERE M.codimovi = 'RT'"
'        SQLX$ = SQLX$ + "     AND M.CANTSALID > 0"
'        SQLX$ = SQLX$ + "     and M.FechMov >= '" & DESD1$ & "' AND M.FECHMOV <= '" & HAST1$ & "'"
'        SQLX$ = SQLX$ + " GROUP BY  M.FechMov,DoSecCor,Nume_Clie, NuPedCli,O.IdSubOr,M.Cod_Inte, M.Cod_Exte,M.Descrip,CantSalid, MA.ANCHO,MA.GRAMAJE,MA.Consumo_X_Corte"
'        SQLX$ = SQLX$ + " ORDER BY M.FechMov"

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHMOV))
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!DoPriCor)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!Nume_Clie)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = rasocli$(XVALO(!Nume_Clie))
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = XVALO(!NUPEDCLI)
            NORFA$ = StringAggFlex$("ORDEFABR", "NROPED =" & XVALO(!NUPEDCLI), "IDSUBOR", ";")
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = NORFA$ 'SAFETEXT$(!idsubor)
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = XVALO(!cod_inte)
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 9) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(XVALO(!CANTSALID), "F12.0")
            GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(XVALO(!ANCHO), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(J&, 12) = FORMATNUM$(XVALO(!GRAMAJE), "F12.0")
            GRILLAGRAL.GRID.TextMatrix(J&, 13) = FORMATNUM$(XVALO(!Consumo_X_Corte), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(J&, 14) = FORMATNUM$(XVALO(!RESULTADO1), "f12.4")
            GRILLAGRAL.GRID.TextMatrix(J&, 15) = FORMATNUM$(XVALO(!Resultado2), "f12.2")
            .MoveNext
          Loop
       End With
       rs.Close
       Set rs = Nothing
       Screen.MousePointer = 1
    
       GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
       GRILLAGRAL.mnuMenu1.Visible = True
       GRILLAGRAL.SinEfectoClickEnGrilla% = 1
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.TXTBUSCAR = ""
       TX$ = ""
      
       
       GRILLAGRAL.Caption = "Entrega x Rango de Fechas (" & PERIO$ & ")"
       
       GRILLAGRAL.Show 1
       
       Dim MSF As msFlexGrid
       If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
    
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENTREMXF", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
            Exit Sub
       End If
99 Screen.MousePointer = 1

End Sub

Private Sub mnuEntregaTelas_Click()
          Screen.MousePointer = 11
          Call ProcesoRecuperaKgDeRecepcionAMovistoDondeCodimoviIgualBR
          
5         Call DESHASFECHA(Des%, Has%, PERIO$)
          If TRIM$(PERIO$) = "" Then GoTo 99
          Call APERBASDAT("STOCKS")
          '
        
          '
          DESD1$ = FETEXCOR1$(Des%)
          HAST1$ = FETEXCOR1$(Has%)
          
        Campos$ = "Fecha/D8;" '1
        Campos$ = Campos$ + "Remito/A18;" '2
        Campos$ = Campos$ + "Cta/F7;" '3
        Campos$ = Campos$ + "R.Social/A24;" '4
        Campos$ = Campos$ + "CI/A0;" '5
        Campos$ = Campos$ + "Código/A20;"      '6
        Campos$ = Campos$ + "Descripción/A32;"      '7
        Campos$ = Campos$ + "Lote/A12;"      '8
        Campos$ = Campos$ + "Metros/F10;"   '9 CANTSALID
        Campos$ = Campos$ + "Kilos/F12.2;"       '10 RESULTADO_CALCULO
        Campos$ = Campos$ + "CantIngXRecep/A0;"        '11 CANTIDADINGRESO
        Campos$ = Campos$ + "Kilos/A0"     '112 PESOKGG  kilos ingresados segun la recepcion
        
    
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
          
        SQLX$ = "SELECT m.Cod_Inte, fechmov, m.Cod_Exte,m.Descrip, m.idenlote, m.CODIMOVI,m.CANTSALID,Nume_Clie,DOPRILAR,"
        SQLX$ = SQLX$ + " (SELECT PESOKG From Movisto  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF')" ' el pesokg del ingreso por br o por of
        SQLX$ = SQLX$ + " AND M.COD_INTE = COD_INTE"
        SQLX$ = SQLX$ + " AND M.IDENLOTE = IDENLOTE) AS PESOKGG,"
        SQLX$ = SQLX$ + " (SELECT CantIngre From Movisto WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF') " 'la cantidad ingresada por br o por of
        SQLX$ = SQLX$ + " AND M.COD_INTE = COD_INTE   AND M.IDENLOTE = IDENLOTE) AS CANTIDADINGRESO,  "
        SQLX$ = SQLX$ + " CASE" ' calculo del peso segun lo entregado en remito (PESOKGG / CANTIDADINGRESO) * M.CANTSALID  IA"
        SQLX$ = SQLX$ + "  WHEN (SELECT CantIngre From Movisto"
        SQLX$ = SQLX$ + " WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF')"
        SQLX$ = SQLX$ + " AND M.COD_INTE = COD_INTE"
        SQLX$ = SQLX$ + " AND M.IDENLOTE = IDENLOTE) IS NOT NULL"
        SQLX$ = SQLX$ + " AND (SELECT CantIngre"
        SQLX$ = SQLX$ + " From Movisto"
        SQLX$ = SQLX$ + " WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF')"
        SQLX$ = SQLX$ + " AND M.COD_INTE = COD_INTE"
        SQLX$ = SQLX$ + " AND M.IDENLOTE = IDENLOTE) > 0"
        SQLX$ = SQLX$ + " THEN"
        SQLX$ = SQLX$ + "  (SELECT PESOKG"
        SQLX$ = SQLX$ + "  From Movisto"
        SQLX$ = SQLX$ + "  WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF')"
        SQLX$ = SQLX$ + "  AND M.COD_INTE = COD_INTE"
        SQLX$ = SQLX$ + "  AND M.IDENLOTE = IDENLOTE) /"
        SQLX$ = SQLX$ + "  (SELECT CantIngre"
        SQLX$ = SQLX$ + "    From Movisto"
        SQLX$ = SQLX$ + "    WHERE (CODIMOVI = 'BR' OR CODIMOVI = 'OF')"
        SQLX$ = SQLX$ + "           AND M.COD_INTE = COD_INTE"
        SQLX$ = SQLX$ + "           AND M.IDENLOTE = IDENLOTE) *"
        SQLX$ = SQLX$ + "           m.CANTSALID"
        SQLX$ = SQLX$ + "         Else NULL"
        SQLX$ = SQLX$ + "     END As RESULTADO_CALCULO"
        SQLX$ = SQLX$ + " From"
        SQLX$ = SQLX$ + "     Movisto m"
        SQLX$ = SQLX$ + "     INNER JOIN MASTER MA ON M.COD_INTE = MA.CODINT"
        SQLX$ = SQLX$ + " Where"
        SQLX$ = SQLX$ + "     MA.TIMATE = 41"
        SQLX$ = SQLX$ + "     AND MA.FAMILIA = 'TE'"
        SQLX$ = SQLX$ + "     AND m.codimovi = 'RT'"
        SQLX$ = SQLX$ + "     AND m.CANTSALID > 0"
        SQLX$ = SQLX$ + "     and M.FechMov >= '" & DESD1$ & "' AND M.FECHMOV <= '" & HAST1$ & "'"
        SQLX$ = SQLX$ + " Order By fechmov,DOPRILAR "
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHMOV))
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!DOPRILAR)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!Nume_Clie)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = rasocli$(XVALO(!Nume_Clie))
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = XVALO(!cod_inte)
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!Descrip)
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = SAFETEXT$(!idenlote)
            GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(XVALO(!CANTSALID), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(J&, 10) = FORMATNUM$(XVALO(!RESULTADO_CALCULO), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(J&, 11) = XVALO(!CANTIDADINGRESO)
            GRILLAGRAL.GRID.TextMatrix(J&, 12) = XVALO(!PESOKGG)
         
            .MoveNext
          Loop
       End With
       rs.Close
       Set rs = Nothing
       Screen.MousePointer = 1
    
       GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
       GRILLAGRAL.mnuMenu1.Visible = True
       GRILLAGRAL.SinEfectoClickEnGrilla% = 1
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.TXTBUSCAR = ""
       TX$ = ""
      
       
       GRILLAGRAL.Caption = "Entrega x Remitos de Telas"
       
       GRILLAGRAL.Show 1
       
       Dim MSF As msFlexGrid
       If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
    
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENTREMTE", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
            Exit Sub
       End If
99 Screen.MousePointer = 1
End Sub

Private Sub mnuEntxFechaMetrosyKilos_Click()
          Screen.MousePointer = 11
          Call ProcesoRecuperaKgDeRecepcionAMovistoDondeCodimoviIgualBR
          
5         Call DESHASFECHA(Des%, Has%, PERIO$)
          If TRIM$(PERIO$) = "" Then GoTo 99
          Call APERBASDAT("STOCKS")
          '
        
          '
          DESD1$ = FETEXCOR1$(Des%)
          HAST1$ = FETEXCOR1$(Has%)
          
        Campos$ = "Fecha/D8;" '1
        Campos$ = Campos$ + "Remito/A18;" '2
        Campos$ = Campos$ + "Cta/F7;" '3
        Campos$ = Campos$ + "R.Social/A24;" '4
        Campos$ = Campos$ + "CI/A0;" '5
        Campos$ = Campos$ + "Código/A20;"      '6
        Campos$ = Campos$ + "Descripción/A32;"      '7
        Campos$ = Campos$ + "Lote/A12;"      '8
        
        
        Campos$ = Campos$ + "Cantidad/F12.2;"      '9
        Campos$ = Campos$ + "U.M./A5;"      '10
        
        Campos$ = Campos$ + "Metros/F10;"   '11 CANTSALID
        Campos$ = Campos$ + "Kilos/F12.2;"       '12 RESULTADO_CALCULO
        
    
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
          
        SQLX$ = "SELECT m.Cod_Inte, fechmov, m.Cod_Exte,m.Descrip, m.idenlote, m.CODIMOVI,m.CANTSALID,MA.UMEDIDA,Nume_Clie,DOPRILAR,PESOKG,METROS "
        SQLX$ = SQLX$ + " From "
        SQLX$ = SQLX$ + "     Movisto m WITH(NOLOCK)"
        SQLX$ = SQLX$ + "     INNER JOIN MASTER MA ON M.COD_INTE = MA.CODINT"
        SQLX$ = SQLX$ + " Where "
        SQLX$ = SQLX$ + "     m.codimovi = 'RT'"
        SQLX$ = SQLX$ + "     AND m.CANTSALID > 0"
        SQLX$ = SQLX$ + "     and M.FechMov >= '" & DESD1$ & "' AND M.FECHMOV <= '" & HAST1$ & "'"
        SQLX$ = SQLX$ + " Order By fechmov,DOPRILAR "
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHMOV))
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = SAFETEXT$(!DOPRILAR)
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = XVALO(!Nume_Clie)
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = rasocli$(XVALO(!Nume_Clie))
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = XVALO(!cod_inte)
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = SAFETEXT$(!Cod_Exte)
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = SAFETEXT$(!Descrip)
            LOTE$ = SAFETEXT$(!idenlote)
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = LOTE$
            GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(XVALO(!CANTSALID), "F12.2")
            GRILLAGRAL.GRID.TextMatrix(J&, 10) = SAFETEXT$(!UMEDIDA)
            If SAFETEXT$(!UMEDIDA) = "M." Then
               GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(XVALO(!CANTSALID), "F12.2") 'SI U.MED ES METROS CARGA LA CANTIDAD SALIDA
            Else
               METROSX# = MetrosSegunLoteXCantidad(XVALO(!cod_inte), LOTE$, XVALO(!CANTSALID))
               GRILLAGRAL.GRID.TextMatrix(J&, 11) = FORMATNUM$(XVALO(METROSX#), "F12.2") 'SI NO LA CALCULA
            End If
            
            If SAFETEXT$(!UMEDIDA) = "KG" Then
               GRILLAGRAL.GRID.TextMatrix(J&, 12) = FORMATNUM$(XVALO(!CANTSALID), "F12.2") 'SI U.MED ES KG CARGA LA CANTIDAD SALIDA
            Else
                PESOKGX = PesoKgSegunLoteXCantidad(XVALO(!cod_inte), LOTE$, XVALO(!CANTSALID))
               GRILLAGRAL.GRID.TextMatrix(J&, 12) = FORMATNUM$(XVALO(PESOKGX), "F12.2") 'SI NO LA CALCULA
            End If
         
            .MoveNext
          Loop
       End With
       rs.Close
       Set rs = Nothing
       Screen.MousePointer = 1
    
       GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
       GRILLAGRAL.mnuMenu1.Visible = True
       GRILLAGRAL.SinEfectoClickEnGrilla% = 1
       GRILLAGRAL.GRID.ZOrder 0
       GRILLAGRAL.TXTBUSCAR = ""
       TX$ = ""
      
       
       GRILLAGRAL.Caption = "Entrega x Remitos de Telas"
       
       GRILLAGRAL.Show 1
       
       Dim MSF As msFlexGrid
       If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
    
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "ENTREMTE", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
            Exit Sub
       End If
99 Screen.MousePointer = 1
End Sub



'
'                    FLEXCONN.Execute ConsultaSql$
'                Next i
'
'                Call COMUNI("Lista de Despacho Registrada")
'            Else
'                Call COMUNI("Problema al Registrar Lista a Despachar")
'            End If
'
'        ElseIf Opcion% = 2 Then
'
'            GoTo 5
'
'        End If
'    End If
'
'End Sub
'
'
Private Sub mnuGenCot_Click()
     GENCOT.Show 1
End Sub
Private Sub mnuGenRemDev_Click()
   GENDEVCOT.Show 1
End Sub

Private Sub mnuRecuperarxRemito_Click()
   Call Command20_Click
End Sub

Private Sub mnuRegDesp_Click()
   DespaCbrra.Show 1
End Sub

Private Sub mnuReimpRemitoQR_Click()
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  'RT'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(VALACT1$("SELREMIT"))
     If NUDOCU$ = "" Then GoTo 10
     
     Call FOREMIT07.IMPRIMIR_DETAREMQR(NUDOCU$) 'GENERA IMPRESION DE REMITO DETALLE EN QR'

99
End Sub

Private Sub mnuReportesVarios_Click()

    Call ACONSUL_PRE(UCase$(TRIM$(Name)))

End Sub

Private Sub mnuSalir_Click()
    Call Command2_Click
End Sub

Sub REMITOACCESSCONTEPLAST_A_WALDBOTT(REMIConPrefijoConfigurado$, REMIXORIGINALCTP$)
        'ESTO ES PARADO EN ENVAPLAST
        Dim RS1 As New ADODB.Recordset
        Dim ENV As New Envap
'        Dim ctp As New Conteplast
        CotizacionParaPasarapesos = ROUND(XVALO(Label15), 2)
        PASO% = 0: VUELTA% = 0
        VectCliente.cargarvector
        CONDI$ = "CodiMovi = 'RT'"
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMIXORIGINALCTP$ & "' "
        
        '
        'RECORRO LOS CODIGOS Y DOY DE ALTA O ACTUALIZO EN WALDBOTT
        
        Dim RS2 As New ADODB.Recordset
        SQLX$ = "SELECT COD_EXTE,DESCRI_LAR FROM MOVISTOCTP WITH(NOLOCK) WHERE " & CONDI$
        Set RS2 = READSET(SQLX$)
        With RS2
          Do While Not .EOF
             CODI$ = SAFETEXT$(!Cod_Exte)
             CODI$ = ENV.CODI_FORMATEADO$(CODI$)
             DESCRIART$ = SAFETEXT$(!DESCRI_LAR)
             If ENV.EXISTE_CODIGO_EXTERNO_WALDBOTT(CODI$) > 0 Then
                Call ENV.UPDATE_ARTICULO_CONTEPLAST_WALDBOTT(CODI$, DESCRIART$)
             Else
                Call ENV.ALTA_ARTICULO_CONTEPLAST_WALDBOTT(CODI$, DESCRIART$)
             End If
             .MoveNext
           Loop
        End With
        RS2.Close
        Set RS2 = Nothing
        '
        
'        SQLX$ = "SELECT MONEVAL,TIPOCAM,COD_INTE,COD_EXTE,NuorItem,DESCRI_LAR, CantSalid, cantingre,NUMEORDPED, NUPEDCLI, ValoUnit, Nume_Clie, FECHMOV,IDENLOTE,COD_EXTE "
'        SQLX$ = SQLX$ + ", CondPag, VENDED,DOMIENT FROM MOVISTO   WHERE " & CONDI$
'        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        CONDI2$ = "CodiMovi = 'RT'"
        CONDI2$ = CONDI2$ + " AND DoPriLAR = '" & REMIXORIGINALCTP$ & "' "

        SQLX$ = " SELECT MONEVAL,TIPOCAM,COD_INTE,COD_EXTE,DESCRIP, SUM(CantSalid) AS CANTIDAD, VALOUNIT"
        SQLX$ = SQLX$ + "  , Sum (ValoUnit * CantSalid) AS IMPORTE, Nume_Clie, FECHMOV,  CondPag, VENDED,DOMIENT, RASO_CLI,NAMEVEND,ORDENCOMPRA,NroVend_Envap,ObservaItem,DESCRIDOS"
        SQLX$ = SQLX$ + " FROM MOVISTOCTP WITH(NOLOCK) WHERE " & CONDI2$
        SQLX$ = SQLX$ + " GROUP BY MONEVAL,TIPOCAM,COD_INTE,COD_EXTE,DESCRIP, VALOUNIT,Nume_Clie, FECHMOV,  CondPag, VENDED,DOMIENT, RASO_CLI,NAMEVEND,ORDENCOMPRA,NroVend_Envap,ObservaItem,DESCRIDOS"
' Call ctp.mostrarDatosx(SQLX$)
        Set rs = READSET(SQLX$)  'RECORRE
        With rs
          Do While Not .EOF
             'NROPED& = XVALO(!NuPedCli)
             If PASO% = 0 Then
                'DATOS DE TIPO DE MONEDA
                MONEDAX% = XVALO(!MoneVal)
                MONEDAINGRX = "L": MONEDAEMISX = "L": TICAM# = 1 'POR DEFECTO L LOCAL
                If MONEDAX% > 1 Then
                  If CotizacionParaPasarapesos < 1 Then 'CAMBIO PARA QUE CONSIDERE PEDIDO REALIZADO EN DOLARES A PESOS
                   TICAM# = XVALO(!TIPOCAM)
                   MONEDAINGRX = "E": MONEDAEMISX = "E": TICAM# = TICAM#
                  End If
                End If
                
                NC& = XVALO(!Nume_Clie)
                CONDVTA% = XVALO(!CondPag)
                VEND& = XVALO(!VENDED)
                DOMIENTREG$ = SAFETEXT$(!DOMIENT)
                DOMIENTREG$ = REPLACAR$(DOMIENTREG$, "'", "")
                DOMICLIX$ = VectCliente.Domicilio(NC&)
                DOMICLIX$ = REPLACAR$(DOMICLIX$, "'", "")
                RAZONSOC$ = TRIM$(SAFETEXT$(!raso_cli))
                If RAZONSOC$ = "" Then 'ESTO POR QUE LOS PRIMEROS NO LO TENIAN
                   RAZONSOC$ = VectCliente.nombre(NC&) 'SI NO LO TIENE LO TOMA DEL ECO, PERO PUEDE PASAR QUE NO ESTE ACTIVO EN ENVAP
                End If
                FE% = CVINT(!FECHMOV)
                NAMEVENDX$ = SAFETEXT$(!NAMEVEND)
                ORDENCOMPRA$ = SAFETEXT$(!ORDENCOMPRA)
                NROVENDENVAP% = XVALO(!NroVend_Envap)
                PASO% = 1
             End If
             J& = J& + 1
             CI1 = XVALO(!cod_inte)
             CODI$ = SAFETEXT$(!Cod_Exte)
             CODI$ = ENV.CODI_FORMATEADO(CODI$)
             DESCRIART$ = SAFETEXT$(!Descrip)
             DESCRIART$ = REPLACAR$(DESCRIART$, "'", "")
             'ORDPED% = XVALO(!NUMEORDPED)
'             REMILAR$ = REMIConPrefijoConfigurado$
             'CANTSALID1 = ROUND(XVALO(!Cantidad), 0) 'CANTSALID AGRUPADO
             CANTSALID1 = TRIM$(FORMATNUM$(XVALO(!Cantidad), "F15.0"))
             PRE = XVALO(!VALOUNIT)
             If CotizacionParaPasarapesos > 0 And MONEDAX% > 1 Then 'SI TIENE VALOR EL LABEL15 CONVIERTE EL
                PRE = ROUND(PRE * CotizacionParaPasarapesos, 4)
             End If
             COSTSINIVA = PRE
             ObservaItem$ = SAFETEXT$(!ObservaItem)
             
             DESCRIDOS$ = SAFETEXT$(!DESCRIDOS)
'             LOTE$ = SAFETEXT$(!IDENLOTE)
             GoSub GRABAR_EN_WALDBOTT
             .MoveNext
          Loop
       End With
       GoTo 75
GRABAR_EN_WALDBOTT:
    If VUELTA% < 1 Then
       ObservaItem$ = REPLACAR$(ObservaItem$, "'", "")
       DESCRIDOS$ = REPLACAR$(DESCRIDOS$, "'", "")
       Call abreBaDaWald
       VUELTA% = 1
       
       On Error GoTo ERROR_GUARDAR
       'BUSCO DATOS DE CODIGO RENGLON 001 Y PROVINCIA BA PARA PASAR ESTOS DATOS A PDCABEZA
       SQLX$ = "SELECT RENGLON,PROVINCIA FROM CLTESENTREGA WHERE CLIENTE = '" & FORMATOCON0$(NC&, 7) & "'"
       Dim RSW As New ADODB.Recordset
       Set RSW = BASEWALD.Execute(SQLX$)
       RENGLONX$ = SAFETEXT$(RSW!RENGLON)
       PROVINCIAX$ = SAFETEXT$(RSW!PROVINCIA)
       RSW.Close
       Set RSW = Nothing
       '
       'BUSCO EN LA TABLA CLIENTES EL VALOR CEL CAMPO LISTAPREC PARA LUEGO AGREGAR EN PDITEMS EN CAMPO LISTA (ANTES GRABABA '01' FIJO)
       LISTAPRECIOSW$ = "01" 'POR DEFECTO
       SQLX$ = "SELECT LISTAPREC FROM CLIENTES WHERE NUMERO = '" & FORMATOCON0$(NC&, 7) & "'"
       Set RSW = BASEWALD.Execute(SQLX$)
       If Not RSW.EOF Then
          LISTAPRECIOSW$ = TRIM$(SAFETEXT$(RSW!LISTAPREC))
       End If
       RSW.Close
       Set RSW = Nothing
       '
       'ahora quieren aqui la cotizacion Gabriel 09/03/21
       If CotizacionParaPasarapesos < 1 Then CotizacionParaPasarapesos = 1
       DOMIENTREG$ = FORMATNUM$(CotizacionParaPasarapesos, "F9.2")
       
       BASEWALD.BeginTrans
       SQLX$ = "INSERT INTO PDCABEZA  "
       SQLX$ = SQLX$ + "(TIPO ,CLIENTE,NROPED,FECEMISION,CONDVENTA,VENDEDOR,ENTREGA,RAZONSOC,NROREAL,CODIGO,FECRECEP,"
       SQLX$ = SQLX$ + " TRANSPORTE,ESTADO,OPERACION,FECALTA,USUARIO,MOTIVO,CONGELA,MONEDAINGR,MONEDAEMIS,COTIZACION,REFERENCIA,"
       SQLX$ = SQLX$ + " ENTREGARNG,PROENT)"
       SQLX$ = SQLX$ + " VALUES("
       SQLX$ = SQLX$ + "'R'"
    
       SQLX$ = SQLX$ + ",'" & FORMATOCON0$(NC&, 7) & "'"  ' FORMATEA AGREGANDOLE 0 A LA IZQUIERDA
       SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMIConPrefijoConfigurado$) & "'"
       SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
       SQLX$ = SQLX$ + ",'" & SAFETEXT$(CONDVTA%) & "'"
       SQLX$ = SQLX$ + ",'" & FORMATOCON0$(VEND&, 3) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(DOMIENTREG$, 40) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(RAZONSOC$, 60) & "'"
       SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMIConPrefijoConfigurado$) & "'"
       
       SQLX$ = SQLX$ + ",'55'"
       SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
       SQLX$ = SQLX$ + ",'000'"
       SQLX$ = SQLX$ + ",'00'"
       SQLX$ = SQLX$ + ",'01'"
       SQLX$ = SQLX$ + ", '" & NOWSERVER & "'"
       SQLX$ = SQLX$ + ",'SUPER'"
       SQLX$ = SQLX$ + ",'01'"
       SQLX$ = SQLX$ + ",'S'"
       
       SQLX$ = SQLX$ + ",'" & MONEDAINGRX & "'"
       SQLX$ = SQLX$ + ",'" & MONEDAEMISX & "'"
       SQLX$ = SQLX$ + "," & FORMATNUM$(TICAM#, "F16.2")
       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(ORDENCOMPRA$), 14) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(RENGLONX$), 3) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(PROVINCIAX$), 2) & "'"
       SQLX$ = SQLX$ + ")"
       BASEWALD.Execute (SQLX$)
    End If
    
    SQLX$ = "INSERT INTO PDITEMS  "
    SQLX$ = SQLX$ + "(TIPO ,CLIENTE,NROPED,ITEM,ARTICULO,DESCART,PRECIO,CANTPED,BULTPED,FECRECEP,PRESENTA,"
    SQLX$ = SQLX$ + " LISTA,DEPSALIDA,FECALTA,USUARIO)"
    SQLX$ = SQLX$ + " VALUES('R'"
    SQLX$ = SQLX$ + ",'" & FORMATOCON0$(NC&, 7) & "'"
    SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMIConPrefijoConfigurado$) & "'"
    SQLX$ = SQLX$ + ",'" & FORMATOCON0$(J&, 3) & "'"

    SQLX$ = SQLX$ + ",'" & Left$(CODI$, 20) & "'"
    SQLX$ = SQLX$ + ",'" & Left$(DESCRIART$, 50) & "'"
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(ROUND(COSTSINIVA, 4)) & "'"
'    SQLX$ = SQLX$ + "," & CANTSALID1
'    SQLX$ = SQLX$ + "," & CANTSALID1
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(CANTSALID1) & "'"
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(CANTSALID1) & "'"
    SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
    
    SQLX$ = SQLX$ + ",'01'"
    SQLX$ = SQLX$ + ", '" & LISTAPRECIOSW$ & "'"
    SQLX$ = SQLX$ + ",'01'"
    SQLX$ = SQLX$ + ", '" & NOWSERVER & "'"
    SQLX$ = SQLX$ + ",'SUPER'"
    SQLX$ = SQLX$ + ")"
    BASEWALD.Execute (SQLX$)
    '
    'AGREGADO A TABLA DESCRIPCIONES PARA INCLUIR MAS DE UNA LINEA DE DESCRIPCION - INGRESA LA DESCRI1 Y LA DESCRI2 DEL MAESTRO
    DESCRIWALD$ = Left$(DESCRIART$, 49) + vbCrLf
    
    
    'ACA EN ENVAPLAST NO PUEDO ACCEDER A LA SEGUNDA DESCRIPCION, DEBERIA GRABARLO CUANDO SE DUPLICA EL REGISTRO DE SQL A MDB O SEA CUANDO SE GRABA EL REMITO ORIGINAL
'    If CODINT(CODI$) > 0 Then
'       DESCRIWALD$ = DESCRIWALD$ & Left$(DESCRIT2$(CODINT(CODI$)), 49)
'    End If
    If TRIM$(DESCRIDOS$) <> "" Then
       DESCRIWALD$ = DESCRIWALD$ & " - " & DESCRIDOS$
    End If
    If TRIM$(ObservaItem$) <> "" Then
      DESCRIWALD$ = DESCRIWALD$ & " - " & ObservaItem$
    End If
    
    strSQL = "SELECT * FROM DESCRIPCIONES"
    strSQL = strSQL + " WHERE ARTICULO  = '" & TRIM$(CODI$) & "'"
    Set RS1 = BASEWALD.Execute(strSQL)
    With RS1
     If .EOF Then 'SI NO EXISTE LA AGREGA
        ingresoendescripcion% = 1
        SQLX$ = "INSERT INTO DESCRIPCIONES  "
        SQLX$ = SQLX$ + "(ARTICULO ,TIPO,CLIPROV,NUMERO,DESCRIPCION)"
        SQLX$ = SQLX$ + " VALUES ("
        SQLX$ = SQLX$ + "'" & Left$(CODI$, 20) & "'"
        SQLX$ = SQLX$ + ",'C'"
        SQLX$ = SQLX$ + ",'0'"
        SQLX$ = SQLX$ + ",'" & Left$(CODI$, 20) & "'"
        SQLX$ = SQLX$ + ",'" & DESCRIWALD$ & "'"
        SQLX$ = SQLX$ + ")"
    Else
        'SI EXISTE LA ACTUALIZA
        SQLX$ = "UPDATE DESCRIPCIONES  SET TIPO = 'C',CLIPROV = '0',NUMERO =  '" & Left$(CODI$, 20) & "', DESCRIPCION = '" & DESCRIWALD$ & "'"
        SQLX$ = SQLX$ + " WHERE ARTICULO = '" & TRIM$(CODI$) & "'"
        ingresoendescripcion% = 1
    End If
    End With
    BASEWALD.Execute (SQLX$)
    If ingresoendescripcion% = 1 Then
        RS1.Close
        Set RS1 = Nothing
        ingresoendescripcion% = 0
    End If
    
'    'FIN GRABACION EN TABLA DESCRIPCIONES.
'    Call ENV.MARCARREMITOTRANSFERIDO(REMIConPrefijoConfigurado$)
Return
ERROR_GUARDAR:
    If Err.Number <> 0 Then
           MsgBox Err.Description

       BASEWALD.RollbackTrans
       On Error Resume Next
       rs.Close
       Set rs = Nothing
       'aca graba con en nro de remito original de conteplast
       SQLX$ = "UPDATE  MOVISTOCTP SET PasoaWb = 0 WHERE " & CONDI2$: FLEXCONN.Execute (SQLX$)
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       
       On Error GoTo 0
       Exit Sub
    Else
75    BASEWALD.CommitTrans
      rs.Close
      Set rs = Nothing
      'aca graba con en nro de remito original de conteplast
      SQLX$ = "UPDATE  MOVISTOCTP SET PasoaWb = 1"
      SQLX$ = SQLX$ + " ,NroModificado = '" & REMIConPrefijoConfigurado$ & "'"
      SQLX$ = SQLX$ + "  WHERE " & CONDI2$
      FLEXCONN.Execute (SQLX$)
      Set VectCliente = Nothing
      Call COMUNI("Transferencia de Remito a Base Waldbott Realizada Correctamente\(Vendedor: " & NAMEVENDX$ & " - Nro.: " & NROVENDENVAP% & ")\(Orden de Compra: " & ORDENCOMPRA$ & ")")
    End If
99
End Sub

Private Sub mnuVincularReportesConFormulario_Click()

    Call CONECTA_CONSUL_PRE(UCase$(TRIM$(Name)))

End Sub


Private Sub MSFDETAREMI_DblClick()
   Call List1_DblClick
End Sub


Private Sub RemContImp_Click()
   REMIENVA.Show 1
End Sub

Private Sub RemCTPAWDBT_Click()
'esto parado en envaplast

'   DESDE EL MODULO DE PEDIDOS SE EL FRM RUTAWALD SE DEBERA CONFIGURAR LA RUTA DONDE SE VA A TOMAR LA INFORMACION
'   GENERADA EN CONTEPLAST
     EPAC$ = TRIM$(UCase$(EMPREAC$))

     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     
     NPFIJOCTP& = XVALO(Control("TRCTPAWL", "PFIJOCTP"))  'NUMERO DE PREFIJO CONTEPLAST
     NPFIJOCAMBACERES& = XVALO(Control("TRCAPAWL", "PFIJOCAM"))  'NUMERO DE PREFIJO CAMBACERES
     NPFIJOQUILMES& = XVALO(Control("TRCAPAWL", "PFIJOQUI"))  'NUMERO DE PREFIJO CAMBACERES
     If NPFIJOQUILMES& < 1 Then NPFIJOQUILMES& = NPFIJOCAMBACERES& 'POR SI NO CONFIGURAN EL DE QUILMES QUE USE EL DE CAMBA
     If NPFIJOCTP& < 1 Or NPFIJOCAMBACERES& < 1 Or NPFIJOQUILMES& < 1 Then
        Call COMUNI("Falta Configurar Prefijo de Remito de Waldbott para los Realizados en Conteplast\Realizar ésta Configuración Desde Módulo de Pedidos\Menú: Configuración > Ruta de conexión a Base de Datos Waldbott ")
        Exit Sub
     End If
     '
     

    TituloLista$ = "Lista de Remitos Pendientes de Transferir a Waldbott"
    Campos$ = "Fecha/D8;Nro.Remito/A20;Cta./f8;Razón Social/A40;Pedido/f8;DOPRINUM/A0;VendContep/A16;O_COMPRA/A16;NroVendEnv/f10"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    Dim RS1 As New ADODB.Recordset
    
12  SQLX$ = " SELECT FechMov,DOPRILAR,Nume_Clie,RASO_CLI,NuPedCli,DOPRINUM,NAMEVEND,ORDENCOMPRA,NroVend_Envap"
    SQLX$ = SQLX$ + " FROM MOVISTOCTP  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE (PasoaWb < 1 OR PasoaWb IS NULL)"
    SQLX$ = SQLX$ + " GROUP BY FechMov,DOPRILAR,Nume_Clie,RASO_CLI,NuPedCli,DOPRINUM,NAMEVEND,ORDENCOMPRA,NroVend_Envap"
    
    
    Set RS1 = READSET(SQLX$)
              
    
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    Screen.MousePointer = 1
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1
    
    NUDOCU$ = TRIM$(RESP_ZELE_VECT(2))
    PREFIJOORIGINAL% = XVALO(Mid$(NUDOCU$, 6, 4))
    
    If PREFIJOORIGINAL% = 1 Then 'SI  ES  CAMBACERES
        PREFIJOCONFI$ = FORMATOCON0$(NPFIJOCAMBACERES&, 4)
    ElseIf PREFIJOORIGINAL% = 2 Then
        PREFIJOCONFI$ = FORMATOCON0$(NPFIJOQUILMES&, 4)
    Else
        PREFIJOCONFI$ = FORMATOCON0$(NPFIJOCTP&, 4) 'SI NO ES CONTEPLAST (9)
    End If
    
    ULTIMOS8CARACTERES$ = Right$(NUDOCU$, 8)
    NuevoRemitoFormateado$ = "RT-X-" & PREFIJOCONFI$ & "-" & ULTIMOS8CARACTERES$
    
30   NROREMITModificado$ = TRIM$(InputBox$("Nro de remito:", "Ingrese Nª de Remito", NuevoRemitoFormateado$))
     If TRIM$(NROREMITModificado$) = "" Then
        Call COMUNI("Debe Ingresar un Nùmero de Remito")
        Exit Sub
     End If
      
     NuevoRemitoFormateado$ = NROREMITModificado$
    If NUDOCU$ = "" Then Exit Sub
'    NroRemitoPrefijoConteplastaWaldbott$= RT-X-0009-00000221
    Dim ENV As New Envap
    EXISTREMI% = ENV.REMITO_EXISTE_EN_WALDBOTT(NuevoRemitoFormateado$)
    If EXISTREMI% > 0 Then
      Call COMUNI("Atención !!! \N° de Remito Existente en Base de Datos Waldbott\Consulte a su Soporte de Sistema")
      Set ENV = Nothing
       Exit Sub
    End If
    Set ENV = Nothing
    ultimoremito& = 0
    '
    'TEMA AGREGADO DE COTIZACION DE REMITO*****************
    
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
       Dim RS2 As New ADODB.Recordset
       CONDI$ = "DoPriLAR = '" & NUDOCU$ & "'" 'aca tiene que ser el remito original de conteplasst
       SQLX$ = "SELECT MoneVal, TipoCam FROM MOVISTOCTP WITH(NOLOCK) WHERE " & CONDI$
       
       Set RS2 = READSET(SQLX$)
       If Not RS2.EOF Then
          MONEDAX% = XVALO(RS2!MoneVal)
          TIPOXCAM# = XVALO(RS2!TIPOCAM)
       End If
       RS2.Close
       Set RS2 = Nothing
       
      MonedaPedido% = MONEDAX%
      If MonedaPedido% > 1 Then
        OPX% = COMALTER%("Pedido Generado en Moneda " & DEMONECO$(MonedaPedido%) & "\\ Generar Remito Manteniendo Moneda Original\Generar Remito en Pesos")
        If OPX% = 2 Then
          TipoCambioPedido# = TIPOXCAM#
          Label15.Visible = True
          Label1(2).Visible = True
          Label1(3).Visible = True
          Label16.Visible = True
          TCP$ = FORMATNUM$(TipoCambioPedido#, "F9.2")
20        ImporteIngresado$ = TRIM$(InputBox$("Ingrese Cotización de Moneda (" & DEMONECO$(MonedaPedido%) & ")", " Cotización", TCP$))
          If ImporteIngresado$ = "" Or XVALO(ImporteIngresado$) < 1 Then GoTo 20 'SI NO INGRESA VUELVE SI O SI TIENE QUE INGRESAR UN VALOR
          'SI INGRESO VALOR LO PONE EN EL LABEL
          Label15 = FORMATNUM$(XVALO(ImporteIngresado$), "F9.2")
       Else
        Label15.Visible = False
        Label1(2).Visible = False
        Label15.Caption = ""
        Label1(3).Visible = False
        Label16.Visible = False: Label16 = ""
       End If
     End If
    End If
    'TEMA AGREGADO DE COTIZACION DE REMITO*****************
    'SELECT * FROM MOVISTOCTP WHERE DOPRILAR = 'RT-X-0009-00000392'  AND PASOAWB = 0
    
    CONDI$ = "DOPRILAR = '" & NUDOCU$ & "' AND PASOAWB = 1"
    If CantRegistros&("MOVISTOCTP", CONDI$, "NOMESS") > 0 Then
       Call COMUNI("Atenciòn!!! Este Remito Yà Fuè Enviado a Waldbott")
       GoTo 12
    End If
    
    Call REMITOACCESSCONTEPLAST_A_WALDBOTT(NuevoRemitoFormateado$, NUDOCU$) 'con el nro de remito formateado
     Label15 = ""
     Label15.Visible = False
     Label1(3).Visible = False
     Label16.Visible = False: Label16 = ""
    GoTo 12

End Sub

Private Sub Text1_Change()

    
    lblNroPed = ""
    Label16 = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text4.TEXT = ""
    Text5.TEXT = ""
    Label26 = ""
    For i = 0 To 3
           Text16(i) = "" ' LIMPIA CAJAS DE OBSERVACIONES
    Next i
       '
    Text8 = ""
    Text12 = ""
    Text14 = ""
    Text13 = ""
    Text15 = ""
    Text18 = ""
    Text19 = ""
    Text13 = ""
    Text17 = ""
    Text20 = ""
    Text21 = ""
    Text22 = ""
    Label18 = ""

End Sub

Private Sub Text1_DblClick()
    '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "CAMBACERES") > 0 Then
      OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Alta Cliente\Actualizar Base de Datos de Clientes")
      If OPX% <= 1 Then Exit Sub
      If OPX% = 2 Then
         Text1.TEXT = AltaClienteCambaceres&
         Call Text1_KeyPress(13)
      Else
         Dim CTP As New Conteplast
         Call CTP.ACTUBASEWALDBOTT(1)
      End If
      Exit Sub
   End If
 
 
    Call SELECHO("DEUDOR12")
    NCX = Val(TRIM$(VALACT1$("DEUDOR12")))
    If NCX > 0 Then
        Text1.TEXT = TRIM$(Str$(NCX))
        Text2.TEXT = rasocli$(NCX)
        Text3.TEXT = DOMICLI$(NCX)
        Text4.TEXT = CPOSCLI$(NCX)
        Text5.TEXT = LOCACLI$(NCX)
        Label26 = CODICLI$(NCX)
        Call CARTRANSPOR
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
      Refresh
      CTXX% = 1
    End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 999999 Then
             NCLI = Val(Text1.TEXT)
'             If ECOARCH$("DEUDOR12", MKS$(NCLI)) <> "" Then
                Text2.TEXT = rasocli$(NCLI)
                Text3.TEXT = DOMICLI$(NCLI)
                Text4.TEXT = CPOSCLI$(NCLI)
                Text5.TEXT = LOCACLI$(NCLI)
                Label26 = CODICLI$(NCLI)
                Call CARTRANSPOR
                List1.Clear
                MSFDETAREMI.Rows = 1

                Unload FORMUL '\\\OT-06-0182-WAR-29/05/06///
                BOTNEXT.Enabled = True
                BOTNEXT.SetFocus
                Exit Sub
'             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text10_Change()
'        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7 = ""
        DP1% = XVALO(Text10)
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))

End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    Static CMO$
    If CMO$ = "" Then
       CMO$ = TRIM$(Control$("", "CODREFAC"))
       If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> CMO$ + "." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    '\\\OT-05-0419-WAR-26/06/05///
    'Call NUREMIFA(NROPRO&)
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'FIN
    If NROPRO& <> NRONUE& Then
        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número libre.")
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    'OT-05-0419-WAR-26/06/05///
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'Call NUREMIFA(NROPRO&)
    'FIN OT
95  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub






Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    Call Command13_Click
End Sub

Private Sub Text7_GotFocus()
    Text10.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Sub CARTRANSPOR()
    '
    NCLIE = Val(Text1.TEXT)
    Text12 = ""
    Text8 = ""
    Text14 = ""
    Text21 = ""

    NTRAI% = TRIM$(Str$(NUTRANCLI%(NCLIE)))
    If XVALO(NTRAI%) > 0 Then
        Text12 = CUITTRANS$(NTRAI%)
        Text8 = RASOTRANS$(NTRAI%)
        Text14 = DOMITRANS$(NTRAI%)
        Text21 = LOCATRANS$(NTRAI%)
    End If
    Exit Sub
'    NCLIE% = Val(Text1.TEXT)
'    LENTREG$ = Space$(96)
'    RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
'    If Len(RFEX$) >= 4 Then
'       RELENT& = CVS(Left$(RFEX$, 4))
'       If RELENT& > 0 Then
'         If RELENT& <= ULTREG&("LENTRECL") Then
'           LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
'         End If
'       End If
'    End If
'    Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

Sub SIGUEREMI1(MODOPRE%)
    '
    EsReteplast% = 0: If InStr(EMPREAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
        
    HOII% = HOY(HOS$)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 999999 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NCLIE = NCXI
    If InStr(EMPREAC$, "ARCON") > 0 Then
      Dim VECTORCANTIDADDELCODIGO(), VECTORPEDIDO()
          ReDim Preserve VECTORCANTIDADDELCODIGO(1) 'TIENE LA CANTIDAD ASIGNADA EN LA GRILLA DE ESE CODIGO
          ReDim Preserve VECTORPEDIDO(1)  'MARCA SI ESE PEDIDO YA SE REVISO
      'ReDim VectorSaldoaRemitirDescuentaAplicado(MAXPEDIDO, 32767)
    End If

    'If ECOARCH$("DEUDOR1", MKI$(NCLIE)) = "" Then
    If CantRegistros("DEUDOR12", "NUME_CLI=" & CStr(NCLIE)) = 0 Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    If FF% < 1 Or FF% > HOY(HO$) + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    Call NotaCli(NCLIE, 1)
    '
    Call CIERRARCH("*.*")
    If ACCESOATOYOTA.Label1 = "" Then
       Call BLANARCH("!BOLREDET")
    End If
    '
    If MODOPRE% = 1 Then GoTo 10
    If PORBARRA% > 0 Then
           LECCODBA.List1.Clear
50         LECCODBA.Show 1
           '
           'SI SE CANCELA
           If XVALO(LECCODBA.List1.ListCount) < 1 Then
              If COMALTER%("Realmente Desea Cancelar Este Remito ?\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 50
              Call BLANARCH("!BOLREDET")
              '
              Call LIACOBSE ' LIMPIA CASILLEROS INFERIORES - ACTIVA TEXT OBSERVACIONES
              '
              Frame15.Visible = False
              PORBARRA% = 0
              Exit Sub
           End If
           '
           Call SETULTREG&("!BOLREDET", 0)
           For U& = 1 To LECCODBA.List1.ListCount
             CI% = LECCODBA.LECT_CODIGO%(LECCODBA.List1, U& - 1)
             CANTI = LECCODBA.LECT_CANTIDAD(LECCODBA.List1, U& - 1)
             '
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, Unimed$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CANTI)), 89, 8)
             Call REPLA(Y$, Chr$(XVALO(U&)), 98, 1)
'             Call REPLA(Y$, MKD$(CDbl(CANTI)), 163, 8)
             '
             Call GRAREG("!BOLREDET", Y$, U&)
           Next U&
           '
           JJ& = ULTREG&("!BOLREDET")
           GoTo 10
    End If
    
    '***********************************
    
    If Frame15.Visible = True Then GoTo 10
    '\\\OT-05-0240-WAR-02/05/05///
    If InStr(EMPREAC$, "ENVAPLAST") > 0 Then GoTo 60 ' ENVAPLAST SALTEA LA SELECCION DE FECHA
    
    ' *** RETEPLAST - SELECCIONA LISTA DE DESPACHO
    If EsReteplast% = 1 Then
        Dim p As New Pedidos
        Dim Cliente As Long
        Cliente = XVALO(Text1)
        NumeroListaDespacho& = p.ListaPedidosPendientes(Cliente)
        If NumeroListaDespacho& > 0 Then
            Me.lblNroPed = NumeroListaDespacho&
            GoTo 60
        Else
            Exit Sub
        End If
    End If
    
    Call SELECHO("SELFECHA/Entrega Solicitada Hasta:")
    FEHAS$ = TRIM$(VALACT1$("SELFECHA"))
    If FEHAS <> "" Then
        HAS1# = CDbl(CVDAT(FECHANUM(FEHAS$)))
    End If
    '\\\OT-05-0240-WAR-FIN///
60  Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    '\\\OT-05-0240-WAR-02/05/05///
    If HAS1# > 0 Then
        SQLX$ = SQLX$ + " AND FeSolEnt <= '" & FEHAS$ & "' "
    End If
    'ENVAPLAST
    If XVALO(Me.lblNroPed) > 0 Then SQLX$ = SQLX$ + " AND NROPED = " & XVALO(Me.lblNroPed)
    '\\\OT-05-0240-WAR-FIN///
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & HOS$ & "')) "
    SQLX$ = SQLX$ + " ORDER BY NroPed ASC, FeSolEnt ASC, NuOrItem ASC " 'ESTA LINEA LA AGREGUE
    '
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
'    Dim PEDTEMP As ADODB.Recordset
'    Set PEDTEMP = New ADODB.Recordset
'    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    Dim PEDTEMP As ADODB.Recordset
'    If InStr(1, EMPREAC$, "J Y M PRUEBA") > 0 Then
'        Set PEDTEMP = New ADODB.Recordset
'        PEDTEMP.CursorLocation = adUseClient
'        PEDTEMP.CursorType = adOpenForwardOnly
'        PEDTEMP.LockType = adLockReadOnly
'        PEDTEMP.Open SQLX$, FLEXCONN, , , adCmdText
'    Else
'        Set PEDTEMP = READSET(SQLX$)
'    End If
    Set PEDTEMP = READSET(SQLX$)
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        CIPD% = XVALO(!CODIINT)
        If CIPD% > 0 And CIPD% <= NUMAS% Then
          CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
            If CAPEN < 0 Then CAPEN = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          FEMIPED% = CVINT(!FECHEMIS)
          FENTRESOL% = CVINT(!FeSolEnt)
          PREUNI = XVALO(!PreUnid) * (100 - PORDESCLI(NCLIE)) / 100
          MONEMA% = XVALO(!MoneEmis)
          '
          Y$ = REGBLAN$("BOLREDET")
          LOTEX1$ = SAFETEXT$(!idenlote)
          
          If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then LOTEX1$ = TRIM$(OrdFabSegunPedido$(XVALO(!NROPED), CIPD%, XVALO(!NuOrItem)))
          Call REPLA(Y$, LOTEX1$, 107, 12)
          Call REPLA(Y$, MKI$(FEMIPED%), 81, 2)
          Call REPLA(Y$, MKI$(CIPD%), 83, 2)
          Call REPLA(Y$, Unimed$(CIPD%), 85, 4)
          Call REPLA(Y$, MKD$(0), 89, 8)
            If EsReteplast% = 1 Then Call REPLA(Y$, MKD$(CDbl(CAPEN)), 89, 8)
            NPROV = XVALO(!Num_Prov)
          Call REPLA(Y$, MKI$(NPROV), 105, 2)
          Call REPLA(Y$, MKS$(!NROPED), 151, 4)
          Call REPLA(Y$, MKI$(FENTRESOL%), 155, 2)
          Call REPLA(Y$, MKI$(MONEMA%), 157, 2)
          Call REPLA(Y$, MKS$(PREUNI), 159, 4)
          Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
            U& = !NuOrItem
          Call REPLA(Y$, MKS$(U&), 189, 4)
          If InStr(EMPREAC$, "ARCON") > 0 Then
             NPXD& = XVALO(!NROPED)
            'EN EL VECTOR CARGO LA CANTIDAD A REMITIR SEGUN V.MAT Y SI PARA MISMO PEDIDO Y CODIGO ESTA REPETIDO LO DESCUENTO
            If UBound(VECTORPEDIDO) < NPXD& Then ReDim Preserve VECTORPEDIDO(NPXD&)
            If UBound(VECTORCANTIDADDELCODIGO) < CIPD% Then ReDim Preserve VECTORCANTIDADDELCODIGO(CIPD%)
            CANTIVM# = VmatPendientedeRemitir(CIPD%, NPXD&)
            VECTORPEDIDO(NPXD&) = VECTORPEDIDO(NPXD&) + 1
            If VECTORPEDIDO(NPXD&) > 1 Then 'SI EL PEDIDO YA TUVO ALGUNA LECTURA INGRESA Y REVISA SI MISMO PEDIDO Y CODIGO SE ASIGNO CANTIDAD
               If VECTORCANTIDADDELCODIGO(CIPD%) > 0 Then
                  VECTORCANTIDADDELCODIGO(CIPD%) = VmatPendientedeRemitir(CIPD%, NPXD&) - VECTORCANTIDADDELCODIGO(CIPD%)
                  CANTIVM# = VECTORCANTIDADDELCODIGO(CIPD%)  'SI YA ASIGNO ANTERIORMENTE DESCUENTA
                  If CANTIVM# < 0 Then CANTIVM# = 0
               End If
            End If
            If CAPEN < CANTIVM# Then CANTIVM# = CAPEN 'NO ASIGNA MAS QUE LA PEDIDA
            VECTORCANTIDADDELCODIGO(CIPD%) = CANTIVM#
            Call REPLA(Y$, MKD$(CDbl(CANTIVM#)), 89, 8)
          End If
          JJ& = JJ& + 1
          Call GRAREG("!BOLREDET", Y$, JJ&)
          '
        End If
      .MoveNext
      Loop
      On Error GoTo 0
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    If InStr(EMPREAC$, "ARCON") > 0 Then
    End If
    
10  BOTNEXT.Enabled = False

    
    Call SIGUEREMI2
    '
End Sub
'

Sub SIGUEREMI2()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    EsReteplast% = 0: If InStr(EPAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
    
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Remito de Facturación - Detalle de Material a Despachar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then ATRI$ = ATRI$ + "/POSEDIT(5)=2"  ' lote no editable
    If EsReteplast% = 1 Then ATRI$ = ATRI$ + "/POSEDIT(4)=2" + "/POINDEL=NO"
    
    If Frame15.Visible = True Then
      Call BLANARCH("!BOLREDET")
      VTB% = VENTABU%("REMILIB" + ATRI$)
      GoTo 11
    End If
    '
    ARCHIPRE$ = "MASTER"
    LPREX% = Val(Control$("", "LIPREXPE"))
    If LPREX% > 0 Then
      If LPREX% < 256 Then
        If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
          ARCHIPRE$ = TRIM$(Str$(LPREX%))
          While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
          ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
      End If
    End If
    '
    VERIST% = (-1)
    If Control$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       VERIST% = 1
    End If
    '
    ATRI$ = ATRI$ + "/ARCHECO(1)=" + ARCHIPRE$
    CONTROLADOR_DETALLE$ = "DETAREMI" ' DEFAULT
    If Control$("", "CONTRODE") <> "" Then CONTROLADOR_DETALLE$ = Control$("", "CONTRODE")
10  VTB% = VENTABU%(CONTROLADOR_DETALLE$ + ATRI$)
11  If VTB% < 0 Then
       If PORBARRA% > 0 Then Call SIGUEREMI1(3): Exit Sub  ' CODIGO DE BARRA
       If COMALTER%("Realmente Desea Cancelar Este Remito ?\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
       Call BLANARCH("!BOLREDET")
       '
       Call LIACOBSE ' LIMPIA CASILLEROS INFERIORES - ACTIVA TEXT OBSERVACIONES
       '
       Frame15.Visible = False
       PORBARRA% = 0
       Exit Sub
    End If
    If PORBARRA% > 0 Then
       Exit Sub
    End If
    '
    NCLI = Val(Text1)
    ATXX% = 0
    For KU& = 1 To ULTREG&("!BOLREDET")
      Y$ = REGLEIDO$("!BOLREDET", KU&)
      CIPD% = CVI(Mid$(Y$, 83, 2))
      If CIPD% > 0 Then
        NUPED& = CVS(Mid$(Y$, 151, 4))
        If NUPED& < 1 Then
          If Option2.Value = False Then
            ENCUPE% = 0
            Call BUSCAPEDAC(CIPD%, NCLI, NUPED&, FEMIPE%, FENTSOL%, PREUNI, CAPEN, NORITU&, ENCUPE%)
            
            '  FEMIPED% = CVINT(!FechEmis)
            '  FENTRESOL% = CVINT(!FeSolEnt)
            '  PREUNI = !PREUNID * (100 - PORDESCLI(NC%)) / 100
              '
            If ENCUPE% > 0 Then
              ATXX% = 1
              Call REPLA(Y$, MKI$(FEMIPE%), 81, 2)
              Call REPLA(Y$, Unimed$(CIPD%), 85, 4)
              Call REPLA(Y$, MKS$(NUPED&), 151, 4)
              Call REPLA(Y$, MKI$(FENTSOL%), 155, 2)
              Call REPLA(Y$, MKS$(PREUNI), 159, 4)
              Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
              Call REPLA(Y$, MKS$(NORITU&), 189, 4)
              Call GRAREG("!BOLREDET", Y$, KU&)
            End If
          End If
        End If
      End If
    Next KU&
    '
    If ATXX% > 0 Then GoTo 10
    '
    JJ& = 0
    For KU& = 1 To ULTREG&("!BOLREDET")
      Y$ = REGLEIDO$("!BOLREDET", KU&)
      CIPD% = CVI(Mid$(Y$, 83, 2))
      If CIPD% > 0 Then
        CANENT = CVD(Mid$(Y$, 89, 8))
        If CANENT > 0.05 Then
          JJ& = JJ& + 1
          Call GRAREG("!BOLREDET", Y$, JJ&)
        End If
      End If
    Next KU&
    Call SETULTREG("!BOLREDET", JJ&)
    If JJ& < 1 Then
      Frame15.Visible = False
      Exit Sub
    End If
    '
    ' VALIDA LOS DATOS INGRESADOS
'    For YX& = 1 To ULTREG&("!BOLREDET")
'      YY$ = REGLEIDO$("!BOLREDET", YX&)
'      CI% = CVI(Mid$(YY$, 83, 2))
'      If ESTRAZABLE%(CI%) > O Then
'          ' OBLIGA A INGRESAR EL LOTE
'          ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
'          CANTSA = CVD(Mid$(YY$, 89, 8))
'          DEPX% = XVALO(Text10)
'          LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
'          If LOTEX$ <> "" Then
'             If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
'                GoTo 19
'             End If
'          End If
'          LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
'          If LOTEX$ = "" Then
'             Exit Sub
'          End If
'19        Call REPLA(YY$, LOTEX$, 69, 12)
'          Call REPLA(YY$, LOTEX$, 107, 12)
'          Call GRAREG("!BOLREDET", YY$, YX&)
'      End If
'    Next YX&
    '
    If FORMUL.Name = "BCODE09" Or FORMUL.Name = "BCODECTP" Then
      If FORMUL.VALI_BCODE09 = "APROB" Then GoTo 20 ' SI VIENE DE LA VENTANA DE COLECTAR DATOS
    End If
    '
    ASIGLOT07.DEPOEXP_REMITO% = XVALO(Text10) 'PASO CMO PARAMETRO EL DEPOSITO DE PANTALLA.
    If EsReteplast% = 1 Then GoTo 30

    Call ASIGNALOTES("!BOLREDET", RESUL$)
    If RESUL$ <> "APRO" Then GoTo 10
    Call AJUSTACANT
    '
20  CONTROCARG = XVALO(Control$("", " COCARGA"))
    If CONTROCARG > 0 Then
        SUPE% = CONTROLCARGA(APROP%, CONTROCARG)
        If SUPE% = 1 Then GoTo 10
    End If
    '
30  Call CIERRARCH("!BOLREDET")
    '
    
    If Control$("REMSTAN", "NUSERIE") = "SI" Then
      VTB% = VENTABU%("DETAREMS" + ATRI$)
      If VTB% < 1 Then
        Frame15.Visible = False
        Exit Sub
      End If
      '
    End If
End Sub
Sub CARGALISTBARRA()

   LECCODBA.List1.Clear
   FIN& = ULTREG&("!BOLREDET")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!BOLREDET", U&)
     TX$ = Space$(96)
     CI% = CVI(Mid$(X$, 83, 2))
     CANTI = CVD(Mid$(X$, 89, 8))
     '

     Call REPLA(TX$, CODEXT$(CI%), 1, 24)
     Call REPLA(TX$, Left(DESCRIT$(CI%), 30), 26, 36)
     Call REPLA(TX$, FORMATNUM$(CANTI, "F16.1"), 64, 16)
     '
     LECCODBA.List1.AddItem TX$
     '
     '
     JJ& = ULTREG&("!BOLREDET")

   Next U&
'   PEDCOBAR07.Label6 = Total(2)
'   PEDCOBAR07.Label2 = "OK"
   LECCODBA.Show 1
   Call SETULTREG&("!BOLREDET", 0)
    For U& = 1 To LECCODBA.List1.ListCount
        CI% = LECCODBA.LECT_CODIGO%(LECCODBA.List1, U& - 1)
        CANTI = LECCODBA.LECT_CANTIDAD(LECCODBA.List1, U& - 1)
        '
        Y$ = REGBLAN$("!BOLREDET")
        Call REPLA(Y$, MKI$(CI%), 83, 2)
        Call REPLA(Y$, Unimed$(CI%), 85, 4)
        Call REPLA(Y$, MKD$(CDbl(CANTI)), 89, 8)
        Call REPLA(Y$, Chr$(XVALO(U&)), 98, 1)
        Call REPLA(Y$, MKD$(CDbl(CANTI)), 163, 8)
        '
        Call GRAREG("!BOLREDET", Y$, U&)
    Next U&

   If LECCODBA.List1.ListCount < 1 Then Exit Sub
   Call SIGUEREMI2
   
'   If PEDCOBAR07.Label2 <> "" Then Call CARDETPEDI(2)
'   NOFRESHPRE% = 0
End Sub



'
Sub BUSCAPEDAC(CIPD%, NCLI, NUPED&, FEMIPE%, FENTSOL%, PREUNI, CAPEN, NORITU&, ENCUPE%)
    '
    Static ESGUIDI%
    '
    If ESGUIDI% = 0 Then
       ESGUIDI% = (-1)
       If InStr(UCase$(EMPREAC$), "GUIDI") > 0 Then ESGUIDI% = 1
    End If
    '
    HOII% = HOY(HOS$)
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NCLI & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & HOS$ & "')) "
    SQLX$ = SQLX$ + "ORDER BY NroPed " 'ESTA LINEA LA AGREGUE YO 17/02/05
    '
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    NUPED& = 0
    FEMIPE% = 0
    FENTSOL% = HOY(HOS$)
    PREUNI = 0
    CAPEN = 0
    NORITU& = 0
    ENCUPE% = 0
    '
    With PEDTEMP
'      On Error Resume Next
'      .MoveLast
'      On Error GoTo 0
'      CAREGIS& = .RecordCount
      '
'      If CAREGIS& = 1 Then
'          NUPED& = !NROPED
'          FEMIPE% = CVINT%(!FechEmis)
'          FENTSOL% = CVINT(!FESOLENT)
'          PREUNI = XVALO(!PreUnid)
'          CAPEN = XVALO(!CanPed) - XVALO(!CantEnt)
'            If CAPEN < 0 Then CAPEN = 0
'          If !TipoPed = "ABIERTO" Then CAPEN = 999999
'          NORITU& = XVALO(!NUORITEM)
'          ENCUPE% = 1
'        ElseIf CAREGIS& > 1 Then

          On Error Resume Next
          .MoveFirst
          If Err Then
            On Error GoTo 0
            GoTo 99
          End If
          On Error GoTo 0
          If !TipoPed = "ABIERTO" Then
            CAPEN = 999999
            NUPED& = XVALO(!NROPED)
            FEMIPE% = CVINT%(!FECHEMIS)
            FENTSOL% = CVINT(!FeSolEnt)
            PREUNI = XVALO(!PreUnid)
            CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
            NORITU& = XVALO(!NuOrItem)
            ENCUPE% = 1
            GoTo 99
          End If
          
          JJ& = 0
          Do While (.EOF = False)
            TTXX$ = MKS$(!NROPED)
            TTXX$ = TTXX$ + AJUSTI$(SAFETEXT$(!NROOCOM), 16) + "  "
               FEX = CVINT(!FECHEMIS)
            TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
               FEX = CVINT(!FeSolEnt)
            TTXX$ = TTXX$ + FECHATEX$(FEX) + "  " '42
               PRECIU = XVALO(!PreUnid)
               PRECIUX# = PRECIU
               If ESGUIDI% = 1 Then PRECIU = 0
            TTXX$ = TTXX$ + FORMATNUM$(PRECIU, "F12.4") '54
                CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
               If CAPEN < 0 Then CAPEN = 0
               If !TipoPed = "ABIERTO" Then CAPEN = 999999
            TTXX$ = TTXX$ + FORMATNUM$(CAPEN, "F10.1") '64
            TTXX$ = TTXX$ + FORMATNUM(!NuOrItem, "F6.0") '70
            TTXX$ = TTXX$ + MKD$(PRECIUX#) '
            JJ& = JJ& + 1
            Call GRAREG("!SELECPE", TTXX$, JJ&)
          .MoveNext
          Loop
          '
          Call SETULTREG("!SELECPE", JJ&)
          '
          '////***////
          Call FRESHECHO("!SELECPE") 'ESTO SE AGREGO EL 16/02/05 PARA REFRESCAR LA LISTA DE PEDIDOS.
          '////***/////
          '
          TITUCON$ = "Cod: " + TRIM$(CODEXT$(CIPD%)) + " - " + DESCRIT$(CIPD%)
20        TTXX$ = REGSEL$("!SELECPE/" + TITUCON$)
          '
          'SE AGREGA ESTO PARA QUE SI CANCELA COMO ES UN PEDIDO ABIERTO DEJE CONTINUAR SIN VINCULAR A NINGUN PEDIDO
          If Len(TTXX$) <= 4 Then
            NUPED& = 0
            FEMIPE% = 0
            FENTSOL% = HOY(HOS$)
            PREUNI = 0
            CAPEN = 0
            NORITU& = 0
            ENCUPE% = 0
            GoTo 99
          End If
          'FIN AGREGADO
          '
          NUPED& = CVS(Left$(TTXX$, 4))
          FEMIPE% = FECHANUM(Mid$(TTXX$, 23, 8))
          FENTSOL% = FECHANUM(Mid$(TTXX$, 23, 8))
          CAPEN = Val(Mid$(TTXX$, 57, 10))
          NORITU& = Val(Mid$(TTXX$, 67, 6))
          PREUNI = CVD(Mid$(TTXX$, 71, 8)) 'CAMBIE POR QUE NO ESTABA BIEN LA POSICION
'          PREUNI = CVD(Mid$(TTXX$, 73, 8))    ' QUEDA ACÁ EN LA PARTE NO VISIBLE DEL LIST
          ENCUPE% = 1
'      End If
      On Error GoTo 0
    End With
99  Set PEDTEMP = Nothing
    '

End Sub

Sub BUSCAPEDAC_VIEJA17_07_13(CIPD%, NCLIE%, NUPED&, FEMIPE%, FENTSOL%, PREUNI, CAPEN, NORITU&, ENCUPE%)
    '
    Static ESGUIDI%
    '
    If ESGUIDI% = 0 Then
       ESGUIDI% = (-1)
       If InStr(UCase$(EMPREAC$), "GUIDI") > 0 Then ESGUIDI% = 1
    End If
    '
    HOII% = HOY(HOS$)
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroClie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + "AND CodiInt = " & CIPD% & " "
    SQLX$ = SQLX$ + " AND (((CanPed - CantEnt) >= 0.05) OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & HOS$ & "')) "
    SQLX$ = SQLX$ + "ORDER BY NroPed " 'ESTA LINEA LA AGREGUE YO 17/02/05
    '
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    NUPED& = 0
    FEMIPE% = 0
    FENTSOL% = HOY(HOS$)
    PREUNI = 0
    CAPEN = 0
    NORITU& = 0
    ENCUPE% = 0
    '
    With PEDTEMP
      On Error Resume Next
      .MoveLast
      On Error GoTo 0
      CAREGIS& = .RecordCount
      '
      If CAREGIS& = 1 Then
          NUPED& = !NROPED
          FEMIPE% = CVINT%(!FECHEMIS)
          FENTSOL% = CVINT(!FeSolEnt)
          PREUNI = XVALO(!PreUnid)
          CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
            If CAPEN < 0 Then CAPEN = 0
          If !TipoPed = "ABIERTO" Then CAPEN = 999999
          NORITU& = XVALO(!NuOrItem)
          ENCUPE% = 1
        ElseIf CAREGIS& > 1 Then
          .MoveFirst
          On Error GoTo 0
          JJ& = 0
          Do While (.EOF = False)
            TTXX$ = MKS$(!NROPED)
            TTXX$ = TTXX$ + AJUSTI$(SAFETEXT$(!NROOCOM), 16) + "  "
               FEX = CVINT(!FECHEMIS)
            TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
               FEX = CVINT(!FeSolEnt)
            TTXX$ = TTXX$ + FECHATEX$(FEX) + "  "
               PRECIU = XVALO(!PreUnid)
               PRECIUX# = PRECIU
               If ESGUIDI% = 1 Then PRECIU = 0
            TTXX$ = TTXX$ + FORMATNUM$(PRECIU, "F12.4")
                CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
               If CAPEN < 0 Then CAPEN = 0
               If !TipoPed = "ABIERTO" Then CAPEN = 999999
            TTXX$ = TTXX$ + FORMATNUM$(CAPEN, "F10.1")
            TTXX$ = TTXX$ + FORMATNUM(!NuOrItem, "F6.0")
            TTXX$ = TTXX$ + MKD$(PRECIUX#)
            JJ& = JJ& + 1
            Call GRAREG("!SELECPE", TTXX$, JJ&)
          .MoveNext
          Loop
          '
          Call SETULTREG("!SELECPE", JJ&)
          '
          '////***////
          Call FRESHECHO("!SELECPE") 'ESTO SE AGREGO EL 16/02/05 PARA REFRESCAR LA LISTA DE PEDIDOS.
          '////***/////
          '
          TITUCON$ = "Cod: " + TRIM$(CODEXT$(CIPD%)) + " - " + DESCRIT$(CIPD%)
20        TTXX$ = REGSEL$("!SELECPE/" + TITUCON$)
          If Len(TTXX$) <= 4 Then GoTo 20
          '
          NUPED& = CVS(Left$(TTXX$, 4))
          FEMIPE% = FECHANUM(Mid$(TTXX$, 23, 8))
          FENTSOL% = FECHANUM(Mid$(TTXX$, 23, 8))
          CAPEN = Val(Mid$(TTXX$, 57, 10))
          NORITU& = Val(Mid$(TTXX$, 67, 6))
          PREUNI = CVD(Mid$(TTXX$, 73, 8))    ' QUEDA ACÁ EN LA PARTE NO VISIBLE DEL LIST
          ENCUPE% = 1
      End If
      On Error GoTo 0
    End With
    Set PEDTEMP = Nothing
    '

End Sub

Sub PRIREMIFAC()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    EsReteplast% = 0: If InStr(1, EPAC$, "RETEPLAST") > 0 Then EsReteplast% = 1
    
    ACUMLAIMPOX# = 0
    IMPORTEX# = 0
    PEDIANT& = 0
    Dim IMPOUNIPMT#(32767)
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
      Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
      Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
        Exit Sub
      End If
    End If
    '
    NrodePedidoAVELLANEDA& = 0
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
       TEXTONOTAVENTA$ = VBCLF + "1.-Remito sin relación c/Nota de venta" + vbCrLf + "2.-Remito relacionado c/Nota de venta (No informado)"
       NrodePedidoAVELLANEDA& = XVALO(InputBox$(TEXTONOTAVENTA$, "Informe Nota de Venta", NrodePedidoAVELLANEDA&))
       If NrodePedidoAVELLANEDA& < 1 Then
          Call COMUNI("Debe Ingresar un Nro. de Pedido")
          Exit Sub
       End If
    End If
    
    CONTA% = ULTREG&("!BOLREDET")
    If CONTA% < 1 Then
      Call COMUNI("No se ha Ingresado Detalle")
      Exit Sub
    End If
    '
    VERIST% = (-1)
    If Control$("REMSTAN", "VERISTOK") = "VERISTOK" Then
      VERIST% = 1
    End If
    '\\\OT-05-0216-WAR-25/04/05///
    REPLAORI% = 0
    If Control$("REMSTAN", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If
    'CONFIGURACION DE CONCATENAR CODIGO CON DESCRIPCION
    If Control$("REMSTAN", "AGRECOOR") = "SI" Then AGREGA_CODICLI% = 1
    
    '\\\OT-05-0216-WAR-FIN///
    '
    '======================== CAMBIO SOLICITADO POR PREMOTECT SI SE CONSIDERA UN  PREFIJO PARA MARCAR COMO SIN CARGO
    REMITOSINCARGO% = 0
    Prefix% = XVALO(Control("", "PREFSCAR"))
    If Prefix% > 0 Then
      If Prefix% <> XVALO(Label10) Then
        OPX% = COMALTER%("Opciones de Tipo de Remito\\Generar Remito Pendiente de Factura\Generar Remito S/Cargo")
        If OPX% > 1 Then
          REMITOSINCARGO% = 1
          
          NROPRO& = 0: pfjo& = Prefix%
          PREFIREM1$ = FORMATOCON0$(pfjo&, 4)
          Call NUREMIFA(NROPRO&, PREFIREM1$)
          Text11.TEXT = "RT" + "." + TRIM$(Str$(NROPRO&))
          Label10 = PREFIREM1$
          
          '
        End If
      Else
        OPX% = COMALTER%("Prefijo Activo Genera Remito S/Cargo\\Cancelar\Generar Remito S/Cargo")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then REMITOSINCARGO% = 1
      End If
    End If
    '===================================================
    DepOrig% = Val(Text10.TEXT)
    CAITMO% = 0
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          If CAN# > 0.005 Then
            CAITMO% = CAITMO% + 1
            CIMOVIM%(CAITMO%) = CI%
            CAMOVIM(CAITMO) = (-1) * CAN#
          End If
        End If
      End If
    Next YX&
    '
    If PUEDEMOVI%(DepOrig%) < 0 Then Exit Sub       ' CANCELA SALIDA
    '
    If Frame15.Visible = True Then GoTo 20
    '
    If VERIST% = 1 Then
      For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        LOTE$ = Mid$(YY$, 107, 12)
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
              If STODEPOS(CI%, DepOrig%) - CAN# < (-0.005) Then
                CONTI% = COMALTER%("ATENCION !!!\Stock Insuficiente en Código '" + TRIM$(CODEXT$(CI%)) + "'\" + DESCRIT0$(CI%) + "\\Cancelar\Continuar")
                If CONTI% <> 2 Then
                  Exit Sub
                End If
              End If
            End If
          End If
        End If
      Next YX&
    End If
    '
20  HOJASUEL% = 0
    If Control$("", "TIPAREMI") = "HOJASUEL" Then HOJASUEL% = 1
    '
    Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    CMO$ = TRIM$(Control$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
'    ' *** VERIFICA QUE NO ESTE OCUPADO EL NUMERO DE REMITO. CASO CONTRARIO, LE SUMA UNO.
'    CodigoMovimiento$ = TRIM$(Control$("", "CODREFAC")): If CMO$ = "" Then CMO$ = "RT"
'    PREFIJO$ = TRIM$(Label10)
'201 NroProximoLibre& = 0
'    Call NUREMIFA(NroProximoLibre&, PREFIJO$)
'    NroRemitoActual& = XVALO(Mid$(Text11, 4))
'    If NroProximoLibre& > NroRemitoActual& Then
'        Text11 = CodigoMovimiento$ + "." + TRIM$(Str$(NroProximoLibre&))
'        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número Libre (RT. " & CStr(NroProximoLibre&) & ").")
'        GoTo 201
'    End If
    '
    NCLI = XVALO(Text1.TEXT)
    REMI$ = TRIM$(Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(Text9.TEXT)
    DEPO% = Val(Text10.TEXT)
    PREFIREM$ = TRIM(Label10)
    NROREMIT$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
    YXI% = 0
    NUORIT% = 0
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("PEDCLI")
    NOCLOSE = 1: NOARRAS% = 1
    '

    
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Or InStr(1, EMPREAC$, "METALURGICASIAM-ARCON") > 0 Then
        'METALES ARCON - CREA VALORES POR DEFAULT DE LOS REGISTROS DEL ENTRECOM
        RECONSI$ = REGBLAN$("ENTRECOM")
        Call REPLA(RECONSI$, MKI$(FF%), 1, 2)
        Call REPLA(RECONSI$, REMI$, 3, 14)
          TIPODO$ = TRIM$(Option1.Caption)
          If Option2.Value = True Then TIPODO$ = TRIM$(Option2.Caption)
        Call REPLA(RECONSI$, TIPODO$, 17, 16)
        Call REPLA(RECONSI$, MKI$(NCLI), 33, 2)
        Call REPLA(RECONSI$, rasocli$(NCLI) + TRIM$(Label11), 35, 30)
        Call REPLA(RECONSI$, Label11, 65, 32)
        Call REPLA(RECONSI$, MKI$(1), 191, 2)
        Call REPLA(RECONSI$, MKI$(0), 193, 2)
        IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
21:     If CantRegistros("ENTRECOM", "IDOCUM=" & IDOCUM&) > 0 Then
            IDOCUM& = IDOCUM& + 1
            GoTo 21
        End If
    End If
    '
    
    TIPODOC$ = "Remito"
    If Option2.Value = True Then
        TIPODOC$ = "Remito (en Consignación)"
    End If
    '
    OCOM$ = Left$(TRIM$(Text20.TEXT), 64)
    STRNPEDI$ = ""
   
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
        BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
        If BDX$ <> "" Then
          Dim ENV As New Envap
          Call ENV.ARTICULO_REMITO_WALDBOTT
        End If
    End If
    '
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
      Call BORRARUBICACION
'      Call AgrupaRemitoConteplast 'no va aca leer p.q.dentro de la sub
    End If
    '

    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      LOTE$ = Mid$(YY$, 107, 12)
      NPEDI& = CVS(Mid$(YY$, 151, 4))
      NUORITPE% = CVS(Mid$(YY$, 189, 4))
      If NPEDI& <> NPEDI_ANT& Then
        STRNPEDI$ = STRNPEDI$ & SAFETEXT$(NPEDI&) + "-" 'CONCATENA LOS NUMEROS DE PEDIDOS PARA IMPRIIMIR HORIZONTAL, SOLICITADO ARCON
      End If
      NPEDI_ANT& = NPEDI&
      
      MEDIDAS$ = Mid$(YY$, 193, 12)
      '
      If ACCESOATOYOTA.Label1 = "" Then
         SUCOD$ = SUBCODI$(NPEDI&, CI%)
         Call REPLA(YY$, SUCOD$, 97, 8)
         Call GRAREG("!BOLREDET", YY$, YX&)
      End If
      '
      If CI% > 0 Or Frame15.Visible = True Then
        If CI% <= NUMAS% Or Frame15.Visible = True Then
          If CAN# > 0.005 Or Frame15.Visible = True Then
            NUOC& = CVS(Mid$(YY$, 151, 4))
            DOSEC$ = ""
            CANTI = (-1) * CAN#
            VUNIST# = CDbl(CVS(Mid$(YY$, 159, 4)))
            MONEVA% = CVI(Mid$(YY$, 157, 2))
            '
            'VALIDACION SI ES ALEJPA GUARDA EL PRECIO DE LISTA EN TODOS LOS CASOS
            If InStr(EMPREAC$, "ALEJPA") > 0 Then
               VUNIST# = PRELISTACT#(CI%, TIMONE%)
               MONEVA% = TIMONE%
               IMPOUNI# = VUNIST# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
               IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
               ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
            End If
            '
            If REMITOSINCARGO% > 0 Then
                'IMPOUNI# = PRELISTACT#(CI%, TIMONE%)
                IMPOUNIPMT#(CI%) = XVALO(VALOBADA("PEDDETA", "PREUNID", "NROPED = " & NPEDI& & " AND CODIINT = " & CI%, "NOMESS"))
                NLIX% = XVALO(VALOBADA("PEDDETA", "LISTA_DE_ITEM", "NROPED = " & NPEDI& & " AND CODIINT = " & CI%, "NOMESS"))
                IMPOUNIPMT#(CI%) = IMPOUNIPMT#(CI%) * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
                IncluirIva% = XVALO(VALOBADA("LISTAPRE", "INCLUIVA", "Cod_Lista =" & NLIX% & " "))
                If IncluirIva% = 1 Then ' SI INCLUYE IVA CALCULA EL PRECIO SIN IVA
                    IMPOUNIPMT#(CI%) = (XVALO(IMPOUNIPMT#(CI%))) * (1 + PORCEIVA(CI%) / 100)
                Else
                    IMPOUNIPMT#(CI%) = XVALO(IMPOUNIPMT#(CI%))
                End If
                IMPORTEX# = CAN# * ROUND(IMPOUNIPMT#(CI%), 2)            'IMPORTE EN PESOS
                ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
              End If

            If MONEVA% < 1 Then MONEVA% = 1
            If MONEVA% > ULTREG&("TAMONED") Then MONEVA% = 1
            '
            DOSELA$ = "P/F - PC." + TRIM$(Str$(NUOC&))
            DOSEC$ = "PC." + TRIM$(Str$(NUOC&))
            '
            SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
            If SERINU$ <> "" Then
              If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
                SERINU$ = SERINU$ + " - "
              End If
            End If
            SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
            '
            REFEMO$ = rasocli$(NCLI)
            If SERINU$ <> "" Then
              REFEMO$ = "S.No." + SERINU$
            End If
            '
            NUORIT% = NUORIT% + 1: NUORIT0% = NUORIT%
            DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CI%)
            REMICOR$ = Mid$(Text11.TEXT, 1, 3)
            NURE1& = Val(Mid$(Text11.TEXT, 4))
            NURE2$ = TRIM$(Str$(NURE1&))
            SUC1% = Val(Label10)
            SUC2$ = TRIM$(Str$(SUC1%))
            'REMICOR1$ = REMICOR$ + SUC2$ + "-" + NURE2$
            REMICOR1$ = REMICOR$ + NURE2$ ' asi se guarda en ventas,

            '123321 temp
            If FORMUL.Name = "BCODE09" Or FORMUL.Name = "BCODECTP" Then
                  If FORMUL.VALI_BCODE09 = "APROB" Then GoTo 28 ' SI VIENE DE VENTANA DE
            End If
            'COLECTORA DE DATOS NO ENTRA
            'REALIZA EL AJUSTE DEL STOCK PARA CADA LOTE CONSUMIDO
             If ESTRAZABLE%(CI%) = 1 And InStr(EPAC$, "AHP") < 1 And EsReteplast% = 0 Then
                For II& = 0 To ASIGLOT07.List1.ListCount
                    XX$ = ASIGLOT07.List1.List(II&)
                    If XVALO(Mid$(XX$, 80, 5)) = YX& Then
                     If CODINT%(TRIM$(Mid$(XX$, 1, 16))) = CI% Then
                       CANTLOT$ = Mid$(XX$, 85) 'EN ESTE REGISTRO ESTA EL STRING DEL LOTE UNO O MAS.
                       CANTASIG = XVALO(Mid$(XX$, 54, 12)) 'CANTIDAD TOTAL ASIGNADA
                       GoTo 25
                     End If
                    End If
                Next II&
                '

                ' MUEVE EL STOCK POR CADA LOTE
25              NN% = InStr(CANTLOT$, ";")
                If NN% > 0 Then
                    LOTE$ = Mid$(CANTLOT$, 1, NN% - 1)
                    If TRIM$(LOTE$) <> "" Then
                         'CANTAA = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS")) 'CODIGO ANTERIOR
                         SAL_LOTE = SaldoLote(CI%, LOTE$, DepOrig%): If SAL_LOTE < FRACMIN(CI%) Then SAL_LOTE = 0
                         If SAL_LOTE > CANTASIG Then SAL_LOTE = CANTASIG
                         'MUEVE EL STOCK PARA CI% Y LA CANTIDAD SAL_LOTE
                         'MAS ALLA DE LOS LOTES SELECCIONADOS VA ASIGNANDO POR EL SALDO SI ALCANZA GENERA EL MOVIMIENTO POR EL TOTAL
                         'SI NO ALCANZA MUEVE POR ESE SALDO Y LEE EL SIGUIENTE LOTE PARA COMPLETAR EL RESTO
                         If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
                            CONDI$ = "IDENLOTE = '" & LOTE$ & "' AND (CODIMOVI = 'OF' OR CODIMOVI = 'BR') AND CANTINGRE > 0"
                            METROX# = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
                            PESOX1# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
                            If InStr(EPAC$, "CONTEPLAST") > 0 And TRIM$(CODFAMIL$(CI%)) = "TE" Then 'ESTO PARA LAS TELAS QUE REMITEN DE MODO PARCIAL EN CONTEPLAST
                              METROX# = XVALO(VALOBADA("MOVISTO", "CANTINGRE", CONDI$, "NOMESS")) 'ACA LEO CON CANTINGRE, ARRIBA LO DEJO COMO ESTABA X LAS DUDAS
                              If METROX# <= 0 Then METROX# = 1
                              PESOX1# = (PESOX1# / METROX#) * SAL_LOTE
                              METROX# = SAL_LOTE
                            End If
                         End If
                         Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, -1 * SAL_LOTE, 0, 0, "", "", "", 0, 0, OCOM$, , , METROX#, PESOX1#)
                         CANTLOT$ = Mid$(CANTLOT$, NN% + 1)
                         CANTASIG = CANTASIG - SAL_LOTE
                         If CANTASIG > FRACMIN(CI%) Then GoTo 25
                    End If
                End If
                GoTo 30
             End If
         
'            If InStr(EPAC$, "AHP") > 0 Then
'                LOTE$ = "" 'GENERA EL MOVIMIENTO SIN LOTE POR QUE NO CONSIDERA ASIGNACION POR LOTES EN REMITO
'                GoTo 28  'DEBERA ASINGAR EL CONSUMO POR LOTE DESDE EL MODULO DE STOCK DE MANERA MANUAL O BIEN CUANDO SE REINGRESE MERCADERIA VA A TENER
'                'QUE ASIGNARLE.
'            Else
             
'            End If
            
'''''            'REALIZA EL AJUSTE DEL STOCK PARA CADA LOTE CONSUMIDO
'''''             If ESTRAZABLE%(CI%) > 0 Then
'''''                For II& = 0 To ASIGLOT07.List1.ListCount
'''''                    XX$ = ASIGLOT07.List1.LIST(II&)
'''''                    If XVALO(Mid$(XX$, 80, 5)) = YX& Then
'''''                     If CODINT%(TRIM$(Mid$(XX$, 1, 16))) = CI% Then
'''''                       CANTLOT$ = Mid$(XX$, 85)
'''''                       GoTo 25
'''''                     End If
'''''                    End If
'''''                Next II&
'''''                '
'''''                ' MUEVE EL STOCK POR CADA LOTE
'''''25              NN% = InStr(CANTLOT$, ";")
'''''                If NN% > 0 Then
'''''                    LOTE$ = Mid$(CANTLOT$, 1, NN% - 1)
'''''                    If TRIM$(LOTE$) <> "" Then
'''''                         CANTAA = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LOTE$ & "'", "NOMESS"))
'''''                         Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, CANTAA, 0, 0, "", "", "", 0, 0, OCOM$)
'''''                        CANTLOT$ = Mid$(CANTLOT$, NN% + 1)
'''''                        GoTo 25
'''''                    End If
'''''                End If
'''''                GoTo 30
'''''            End If
            '
            
28          If EsReteplast% = 1 Then
                Dim rstDescargaBultos As ADODB.Recordset
                ConsultaSql$ = "select b.ordenfabricacion lote, sum(b.cantidadbulto) cantidad from" _
                & " bultosfabric b with(nolock) inner join ordefabr o with(nolock) on (b.ordenfabricacion=o.idsubor)" _
                & " where B.pedido=" & CStr(NPEDI&) & " and o.cod_inte=" & CStr(CI%) _
                & " group by b.ordenfabricacion"
                Set rstDescargaBultos = READSET(ConsultaSql$)
                
                With rstDescargaBultos
                    Do While Not .EOF
                        LOTE$ = !LOTE
                        CANTI = (-1) * !Cantidad
                        Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, CANTI, 0, 0, "", "", "", 0, 0, OCOM$, , , METROX#, PESOX1#)
                        
                        .MoveNext
                    Loop
                End With
                On Error Resume Next
                rstDescargaBultos.Close
                Set rstDescargaBultos = Nothing
                On Error GoTo 0
                
                ' VINCULA LOS BULTOS CON EL REMITO QUE LOS DESPACHO
                If NPEDI& <> PedidoAnterior& Then Call ACTUREGISTRO("bultosfabric", "remito", "pedido=" & CStr(NPEDI&), NROREMIT$, REGAF&)
                PedidoAnterior& = NPEDI&
            Else
                If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
                   CONDI$ = "IDENLOTE = '" & LOTE$ & "' AND (CODIMOVI = 'OF' OR CODIMOVI = 'BR') AND CANTINGRE > 0"
                   METROX# = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
                   PESOX1# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
                End If

                Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, CANTI, 0, 0, "", "", "", 0, 0, OCOM$, , , METROX#, PESOX1#)
            End If
            
            If InStr(EPAC$, "AHP") > 0 Then
               GoSub CONSUMO_AHP
               GoTo 30
            End If

30          CIMOVI%(CI%) = 1
            'MOVIMIENTO COMPENSACION INTERNA PARA CI% KIT
            If DESCOMPO%(CI%) > 0 Then
                Call MOVISTOK(FF%, CI%, LOTE$, "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, CANTI * (-1), 0, 0, "", "", "", 0, 0, OCOM$)
                'MOVIMIENTO SALIDA PARA COMPONENTES DEL KIT
                Call AREXPLO(CI%, 1, 1)
                For KKJ% = 1 To CAIT%
                  NUORIT% = NUORIT% + 1
                  CICO% = CIJ%(KKJ%)
                  CACICO = USOI(KKJ%) * CANTI 'EL LOTE LO PASO VACIO POR QUE ANTES SE QUEDABA CON EL LOTE DEL PRODUCTO KIT, LUEGO EL PROCESO LE APLICARA EL LOTE
                  Call MOVISTOK(FF%, CICO%, "", "CI", NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, CACICO, 0, 0, "", "", "", 0, 0, OCOM$)
                Next KKJ%
            End If
            '
            If ACCESOATOYOTA.Label1 <> "" Then
               CEMBASE$ = ACCESOATOYOTA.Label1
               NORICA% = CVI(Mid$(YY$, 187, 2))
               FESEL% = FECHANUM(TRIM$(Left$(CEMBASE$, 16)))
               NOSE$ = TRIM$(Mid$(CEMBASE$, 17, 16))
               NUENTRE$ = TRIM$(Mid$(CEMBASE$, 33, 16))
               condi0$ = "Num_Os = '" & NOSE$ & "' AND Num_Entre = '" & NUENTRE$ & "' AND FechEntre = '" & FETEXCOR1$(FESEL%) & "'"
               CONDI$ = condi0$ + " AND NorItCa = " & NORICA%
               '
               ACUCAN = Abs(CANTI) + XVALO(VALOBADA("DETACARTEM", "Cant_Entre", CONDI$))
               Call ACTUREGISTRO("DETACARTEM", "Cant_Entre", CONDI$, ACUCAN, REGAF&)
               Call ACTUREGISTRO("DETACARTEM", "Dopricor", CONDI$, REMICOR1$, REGAF&)
               Call ACTUREGISTRO("DETACARTEM", "DopriLar", CONDI$, NROREMIT$, REGAF&)
               Call ACTUREGISTRO("DETACARTEM", "NuorItem", CONDI$, NUORIT0%, REGAF&)
               Call ACTUREGISTRO("DETACARTEM", "FechEmis", CONDI$, CVDAT(FF%), REGAF&)
            End If
            
            'METALES ARCON - GENERA ARCHIVO CON LA INFORMACION QUE LUEGO SE TRANSFIERE A AVELLANEDA
            If InStr(1, EMPREAC$, "NEUQUEN") > 0 Or InStr(1, EMPREAC$, "METALURGICASIAM-ARCON") > 0 Then
                RTRASY$ = RECONSI$
                IDOCUM& = IDOCUM& + 1
                Call REPLA(RTRASY$, MKS$(IDOCUM&), 97, 4)
                Call REPLA(RTRASY$, TRIM$(Text20) + "/" + TRIM$(LOTE$), 101, 16) ' ORDEN DE COMPRA
                Call REPLA(RTRASY$, MKI$(CI%), 117, 2)
                Call REPLA(RTRASY$, CODEXT$(CI%), 119, 16)
                Call REPLA(RTRASY$, DESCRIT0$(CI%), 135, 48)
                Call REPLA(RTRASY$, MKD$(CAN#), 183, 8)
                Call REGAPP("ENTRECOM", RTRASY$)
                '
                XX$ = RTRASY$
                FF1% = CVI(Mid$(XX$, 1, 2)) ' este le cambie la variable le agregue 1 al final
                REMI1$ = Mid$(XX$, 3, 14)   ' este le cambie la variable le agregue 1 al final
                TIPODO$ = Mid$(XX$, 17, 16)
                NCLI1 = CVI(Mid$(XX$, 33, 2)) ' este le cambie la variable le agregue 1 al final
                IDOCUM& = CVS(Mid$(XX$, 97, 4))
                OCOMPRA$ = Mid$(XX$, 101, 16)
                CI% = CVI(Mid$(XX$, 117, 2))
                Codigo$ = Mid$(XX$, 119, 16)
                CAN# = CVD(Mid$(XX$, 183, 8))
                '
35              CONDI$ = "IDOCUM=" & IDOCUM&
                If CantRegistros("ENTRECOM", CONDI$, "NOMESS") > 0 Then
                    IDOCUM& = IDOCUM& + 1
                    GoTo 35
                End If
                
                SQLX$ = "SELECT * FROM ENTRECOM WHERE IDOCUM < 1"
                Dim RST1 As ADODB.Recordset
                Set RST1 = New ADODB.Recordset
26:                 RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                If Err Then
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 26
                End If
                On Error GoTo 0
                '
                With RST1
                    .AddNew
                    !FECHA = CVDAT(FF1%)
                    !DOCUMENTO = REMI1$
                    !TIPODOC = TIPODOC$
                    !nume_cli = NCLI1
                    !IDOCUM = IDOCUM&
                    !cod_inte = CI%
                    !Cod_Exte = Codigo$
                    !Cantidad = CAN#
                    !Status = 1
                    !VECESENVIADO = 0
                    .Update
                End With
                '
                On Error Resume Next
                RST1.Close
                Set RST1 = Nothing
                On Error GoTo 0
            End If
          End If
        End If
      End If
    Next YX&
    NOCLOSE% = 0: NOARRAS% = 0
    ACCESOATOYOTA.Label1 = ""
    On Error Resume Next
    If Mid$(STRNPEDI$, Len(STRNPEDI$)) = "-" Then STRNPEDI$ = Left$(STRNPEDI$, (Len(STRNPEDI$) - 1)) 'quito el ultimo guion
    On Error GoTo 0
    '
    If REMITOSINCARGO% = 1 Then
        FEMOV$ = FETEXCOR1$(FF%)
        SQLX$ = "UPDATE MOVISTO SET DOSECCOR = 'S/Cargo'"
        SQLX$ = SQLX$ + ", DoSecLar ='Sin Cargo' "
        SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMICOR1$ & "'"
        SQLX$ = SQLX$ + " AND NUMPVTA= " & Prefix%
        SQLX$ = SQLX$ + "AND Nume_Clie = " & NCLI & " "
        SQLX$ = SQLX$ + "AND FechMov = '" & FEMOV$ & "'"
        FLEXCONN.Execute (FILTSQL$(SQLX$))
    End If
    
    If TRIM$(STRNPEDI$) <> "" Then STRNPEDI$ = "O.V.: " & STRNPEDI$
    '
    
    ' blanquear pedidos pendientes
    Call APERBASDAT("PEDCLI")
    NCLI = Val(Text1.TEXT)
'MsgBox "ANTES DE BLANPEDIPE"
    If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 And PORBARRA% > 0 Then
       'descuenta de manera automatica de los pedddos
       Call ACTUALIZA_PEDIDOS(NCLI, NROREMIT$)
       GoTo 5
    End If
    '
    For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
       NUORITPE% = CVS(Mid$(YY$, 189, 4))
       '
       Call BLANPEDIPE(NCLI, FF%, NROREMIT$, CI%, CAN#, NUPEDIORI&, NUORITPE%)
       If NUPEDIORI& > 0 Then GoSub ANOTAPEDI
       '
    Next YX&
'MsgBox "DESPUES DE BLANPEDIPE"
    '
5   Call BAJALDISCO
    '
    If PEDIANT& > 0 Then
       Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
       Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
    End If
    '
    
    If InStr(EPAC$, "ENVAPLAST") > 0 Then
       BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
       If BDX$ <> "" Then
         Call REMITO_A_WALDBOTT(NROREMIT$)
       End If
    End If
    
    If InStr(1, EMPREAC$, "NEUQUEN") > 0 Then
       Call GeneraArchivoRemitoNeuquen(NROREMIT$, NrodePedidoAVELLANEDA&)
    End If
    
    If (InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0) Or InStr(EPAC$, "QUILMES") > 0 Then
         NROPED& = XVALO(Me.lblNroPed)
         'If InStr(EPAC$, "QUILMES") < 1 Then 'si hay que mandar tambien al drop va a ver que quitar esto
            Call GeneraArchivoRemitoConteplast(NROREMIT$, NROPED&) 'y ademas configurar por que da error de reintentar
         'End If
         AgrupaLote% = 0
        Call AgrupaRemitoConteplast(AgrupaLote%) 'si es cambaceres es 0
        Call GrabaNroRemitoEnCierreOFEnMovisto(NROREMIT$)
    End If
    GoSub PRIDOCREMI
    OPX% = 0
    If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
       StringNroCtes$ = StringDeNroClientesComprasInternas$
       'SI EL NUMERO DE CLIENTE ESTA DENTRO DE LA CADENA DE NRO DE CLIENTES CONFIGURADOS COMO PARTICIPANTES DE LA COMPRA INTERNA.
       If NumeroEstaEnCadena(StringNroCtes$, NCLI) Then
            OPX% = COMALTER%("Impresión de Detalle en QR para Conteplast\\No, No Imprimir\Si Imprimir Detalle QR")
       End If
       If OPX% > 1 Then
          Call IMPRIMIR_DETAREMQR(NROREMIT$) 'GENERA IMPRESION DE REMITO DETALLE EN QR'
       End If
    End If
    
    If InStr(EPAC$, "AHP") > 0 And ASIGLOAHP.List2.ListCount > 0 Then
      ASIGLOAHP.EJECUTAR = 1
      ASIGLOAHP.Show 1
    End If
    '
GoTo 999
    '
    '
    '
    
PRIDOCREMI:
    '
    PRINTERPORT$ = TRIM$(Control$(IDENTER$, "PORTREM"))
 'MsgBox "printerport = *" + PRINTERPORT$ + "*"
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    'ASIGNO EL FORMULARIO CONFIGURADO - SI NO TIENE CONFIGURADO FRM LE VUELVO A ASIGNAR EL GENERAL
    
    If REMITOSINCARGO% > 0 Then
        FORIPRE$ = TRIM$(Control$("", "FORREMSC"))
        If FORIPRE$ = "" Then
          FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
        End If
    End If
    
 'MsgBox "foripre = *" + FORIPRE$ + "*"
    If TRIM$(FORIPRE$) <> "" Then
      '
'      PREFIREM$ = TRIM(Label10)
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
'      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
'      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      Call LIMPIATETAB
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      
      NROPROV_PROPIO = ""
      If InStr(EPAC$, "ENVAPLAST") > 0 Then NCLIX& = NCLI: NROPROV_PROPIO = ENV.NROPROV_ENVAPLAST_WALDOBTT(NCLIX&)
      If Option2.Value = True Then
        TIPODOC$ = "Remito (en Consignación)"
      End If
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
        If TRIM$(Text8.TEXT) <> "" Then TEXTOENTREGA$ = TRIM$(Text8.TEXT)
        If TRIM$(Text14.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Text14.TEXT)
        If TRIM$(Text21.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Text21.TEXT)
        If TRIM$(TXTCODPOSTAL.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(TXTCODPOSTAL)
        CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TEXTOENTREGA$
      End If
      
      
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMIT$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = TRIM$(FORMATNUM$(ACUMLAIMPOX#, "F14.2"))
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = TRIM$(STRNPEDI$)
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(rasocli$(NCLI)) + " - Almacén " + Label11
      
      CODPARAM$(19) = "PROV-CLI": DOCPARAM$(19) = NROPROV_PROPIO
      CODPARAM$(20) = "N-REMITO": DOCPARAM$(20) = NROREMIT$ 'ACA IMPRIME EL REMITO ORIGINAL COMO DOC.INT.:
      CODPARAM$(21) = "N-PLANO": DOCPARAM$(21) = CStr(XVALO(NROPED&))
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 21

      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      pasox% = 0
      For YX& = 1 To ULTREG&("!BOLREDET")

        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        NPEDI& = CVS(Mid$(YY$, 151, 4))
        CAC# = 0
        If CI% > 0 Or Frame15.Visible = True Then
          If CI% <= NUMAS% Or Frame15.Visible = True Then
            If CAN# + CAC# > 0.005 Or Frame15.Visible = True Then
              '
              UNID$ = Unimed$(CI%)
              If UNID$ = "" Then UNID$ = Unimed$(CI%)
              CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
              If CAC# >= 0.05 Then
                CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                Call REPLA(CANS$, "(*)", 1, 3)
              End If
              NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
              NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
              ORD& = XVALO(Mid$(YY$, 241, 5)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              '
              SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
              If SERINU$ <> "" Then
                If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
                  SERINU$ = SERINU$ + " - "
                End If
              End If
              SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
              '
              CAITAB1% = CAITAB1% + 1
              CODDX$ = CODEXT$(CI%)
              If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
                CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
              End If
              If REPLAORI% = 1 Then
                CODDX$ = CODORIG$(CI%, NCLI)
              End If
              '
              TETABU1$(1, CAITAB1%) = CODDX$
              TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
              If Frame15.Visible = True Then
                TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
              End If
              '
              TETABU1$(3, CAITAB1%) = CANS$
              TETABU1$(4, CAITAB1%) = NULOTE$
              TETABU1$(5, CAITAB1%) = UNID$
              TETABU1$(6, CAITAB1%) = NUOCO$
              TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
              TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
              TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
              '
              'VALORIZACION ALEJPA SEGUN LISTA VIGENTE
              IMPOUNI# = 0: IMPORTEX# = 0
              If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 Then
                IMPOUNI# = PRELISTACT#(CI%, TIMONE%)
                IMPOUNI# = IMPOUNI# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
                IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
                ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
              End If
              If REMITOSINCARGO% > 0 Then
                'LO CARGA ARRIBA
                IMPOUNI# = ROUND(IMPOUNIPMT#(CI%), 2)
                IMPORTEX# = CAN# * ROUND(IMPOUNI#, 2)            'IMPORTE EN PESOS
                'ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
                'VUELVO TODO HACIA ATRAS
                Label10 = TRIM$(Control$("", "PREFIREM"))
                NROPRO& = 0
                Call NUREMIFA(NROPRO&, PREFIREM$)
                Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
              End If

              TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
              TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
                 DES2$ = TRIM$(DESCRIT2$(CI%))
                 If DES2$ <> "" Then
                    Call PRINTEXT(DESCRIT2$(CI%), 2, "DES-MAT", FORIPRE$)
                 End If
              End If

                CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
                'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
                If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
                'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
              If InStr(EPAC$, "ENVAPLAST") > 0 Then
'                 If CORIG$ <> "" Then 'al final dijo Enrique que no quiere que se imprima
'                   Call PRINTEXT(CORIG$, 1, "COD-MAT", FORIPRE$)
'                 End If
              Else
                 TETABU1$(12, CAITAB1%) = CORIG$
              End If
              'FIN ALEJPA
              '
              TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              
              If EsReteplast% = 1 Then
                NumeroRemito$ = NUMEDOC$
                CodigoInternoProducto% = CI%
                
                ConsultaSql$ = "select count(*) bultos from bultosfabric b with(nolock)" _
                & " inner join ordefabr o with(nolock) on (b.ordenfabricacion=o.idsubor)" _
                & " where b.remito='" & NumeroRemito$ & "' and o.cod_inte=" & CStr(CodigoInternoProducto%)
                Set Rst = READSET(ConsultaSql$)
                If Not (Rst.BOF And Rst.EOF) Then
                    CantidadBultosDespachados& = XVALO(Rst!bultos)
                End If
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
                TETABU1$(13, CAITAB1%) = CStr(CantidadBultosDespachados&)
              End If
              
              'IMPRESION DE ESPECIFICACION
              If Control("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
                XX$ = ESPECIFIT$(CI%)
                Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
              End If
              '
              If Control("REMSTAN", "PRIOBSPE") = "SI" Then 'OBSERVACIONES DE PEDIDOS
                NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
                NUORITPE% = CVS(Mid$(YY$, 189, 4))
                OBSPE$ = OBSERVAPEDI$(CI%, NCLI, NUPEDIORI&, NUORITPE%)
                Call PRINTEXT(OBSPE$, 2, "DES-MAT", FORIPRE$)
              End If
              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then  'IMPRIME EL LOTE DEBAJO DE LA DESCRIPCION Y/O OBSERVACION DEL PEDIDO
                 Call PRINTEXT(NULOTE$, 2, "DES-MAT", FORIPRE$)
                 Call CARLINSEP("INVISIBLE")         'Carga linea de separacion entre codigo y codigo EN ESTE CASO SIN LINEA
              Else
                 Call CARLINSEP
              End If
              
              If SERINU$ <> "" Then
                CAITAB1% = CAITAB1% + 1
                For KI1% = 1 To 20
                  TETABU1$(KI1%, CAITAB1%) = ""
                Next KI1%
                TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
              End If
              '
              If InStr(EPAC$, "CAMBACERES") < 1 And InStr(EPAC$, "CONTEPLAST") < 1 Then
                 Call DETLOTREM(CI%, YX&) 'SI NO LO AGREGA A LA DESCRIPCION
              Else
                 CF$ = TRIM$(CODFAMIL$(CI%))
                 If CF$ <> "BO" Then
                   Call DETLOTREMCAMBA(CI%, NROREMIT$, NPEDI&)
                 End If
              End If
              
              '
              'NCIY% = NCLI
              CIIY% = CI%
              'Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
            End If
          End If
        End If
      Next YX&
      '
      ''''''''''''''''''''''''''''''''''''''''''''''''''
      'TENER EN CUENTA QUE EL REMITO SE GRABA SIEMPRE, LO UNICO QUE SE P UEDE EVITAR ES QUE SI LO VE EN PANTALLA Y ESTA MAL, NO SE IMPRIMA EN PAPEL, EN TODO CASO PUEDA
      'ANULARLO Y GENERARLO NUEVAMENTE
'      If InStr(EPAC$, "QUILMES") > 0 Then
'        CANCELPRINT% = 0
'        If COMALTER%("Opciones de Impresión:\\Imprime Comprobante\Vista Previa") = 2 Then
'          CANCELPRINT% = 1
'          Call PRINTDOC("@" + FORIPRE$)
'          CANCELPRINT% = 0
'          NUDOCU$ = TRIM$(Mid$(NROREMIT$, 6))
'          If NUDOCU$ = "" Then GoTo 50
'          TIDOCUM$ = "Remito"
'          nubus = XVALO(Mid$(NUDOCU$, 6))
'          DOCUNU& = 0
'          'LE PASO EN LA PARAMETRO DE CONTROL EL NOMESS PARA QUE SI NO ENCUENTRA EL DOCUMENTO
'          'NO PRESENTE MENSAJE Y CONTINUE Y LO BUSQUE DE LA MANERA QUE ESTABA ANTES
'          Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str(nubus), NUDOCU$, "NOMESS")
'          If COMALTER%("Opciones de Impresión:\\Cancelar\Imprime Comprobante") = 2 Then GoTo 50
'          Call LIMPIATETAB
'        End If
'       End If
      '''''''''''''''''''''''''''''''''''''''''''''''''
 'MsgBox "antes de PRINTDOC(FOREMIFA)"
50    CantidadCopiasRemito% = XVALO(Text23)
      If CantidadCopiasRemito% > 0 Then
        CACOPDOC% = CantidadCopiasRemito%
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
        CACOPDOC% = 0
      Else
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
      End If
      '
    End If
    '
    Call LIMPIATETAB
Return
    '
ANOTAPEDI:                ' ANOTADOR DE PEDIDOS
   '
   NPEDIDX& = NUPEDIORI&
   If NPEDIDX& <> PEDIANT& Then
      If PEDIANT& <> 0 Then
          Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
          Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
      End If
      PEDIANT& = NPEDIDX&
      TTXX$ = "Remito " + TRIM$(Mid$(Text11, 4)) + ":"
      '
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, TTXX$, 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      Call REPLA(TTXYZ$, Space$(88), 1, 88)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      '
   End If
   '
   TTXX$ = Space$(76)
   Call REPLA(TTXX$, "- " + CODEXT$(CI%), 11, 18)
   Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
   Call REPLA(TTXX$, FORMATNUM$(CAN#, "F10.1"), 57, 10)
   Call AGREGA_ANOTAPED(NPEDIDX&, TTXX$)
   '
Return

CONSUMO_AHP:
    '
    '****HAY QUE REVISARLO POR QUE ESTA GENERANDO ANTES DE ESTA RUTINA EL MOVIMIENTO Y LUEGO VIENE ACA POR ENDE DUPL
    'ESTE CONSUMO TIENE EL CRITERIO QUE SI TIENE LOTES Y LA CANTIDAD NECESARIA GENERA UNA SALIDA CON LOS LOTES
    'EN CASO QUE LOS MISMOS TENGAN ALGUN PARCIAL GENERA 2 MOVIMIENTOS UNA SALIDA CON LOTE X EL SALDO DEL LOTE
    'Y OTRA POR LA CANTIDAD RESTANTE PERO SIN LOTE.
    'SI NO POSEE LOTE DIRECTAMENTE GENERA UNA SALIDA POR EL TOTAL DE LA CANTIDAD.
    '
    'LEO EL LOTE
'    LOTE$ = TRIM$(Mid(YY$, 107, 12))
'    NPROVX = CVS(Mid(YY$, 105, 2))
'    ASIGLOAHP.lblfactura.Caption = NROREMIT$
'    ASIGLOAHP.LBLCLIENTE.Caption = NCLI
'    CANTAA = Abs(CANTI) ' DEJE ESTO DESPROLIIJDAD PARA NO CAMBIAR LO DE ABAJO.
'    CANTIDAD = CANTAA
'    CANTI = 0
'    If LOTE$ <> "" Then ' SI TIENE LOTE
'      CANLOTE = SaldoLote(CI%, LOTE$): If CANLOTE < 0 Then CANLOTE = 0  'VEO EL SALDO DEL LOTE
'      If CANLOTE > 0 Then             'SI TIENE SALDO
'        CANTI = CANTIDAD                'GUARDO LA CANTIDAD ORIGINAL
'        If CANTI > CANLOTE Then CANTI = CANLOTE 'SI LA CANTIDAD ORIGINAL ES MAYOR AL SALDO DEL LOTE ASIGNO LA CANTIDAD DEL LOTE
'        'GENERA MOVIMIENTO DE  SALIDA CON LOTE POR EL TOTAL DE LA CANTIDAD SI ALCANZA O POR EL SALDO DEL LOTE
'       'Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'       Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, -1 * CANTI, NPEDI&, NUORITPE%, "", "", "", 0, 0, OCOM$)
'       If Abs(CANTAA) - Abs(CANTI) > 0 Then GoTo 60
'      End If
'    Else
'60      CANTI = CANTIDAD - CANTI 'RECALCULA CANTIDAD A ENTREGAR
'        LOTE$ = ""
'        'SI FUE POR EL TOTAL SE VA
'        If CANTI < 0.05 Then
'           Return
'        Else
'           'SI QUEDA SALDO DE CANTIDAD GENERA UN NUEVO REGISTRO POR EL SALDO A REGISTRAR CON LOTE VACIO.
'           CANTIDAD = CANTI
'           'Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMILAR$, REMICOR$, NUFACLAR$, NUFACORT$, RASOCLI$(NCLI), NORIT%, PREUNI#, MONEMI%, NCLI, DEPOSALI%, SIG% * CANTI, NUPEDIORI&, NUORITPE%, TEXLIB$, , , PRECONIVA#) '
'           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, NROREMIT$, REMICOR1$, DOSELA$, DOSEC$, REFEMO$, NUORIT%, VUNIST#, MONEVA%, NCLI, DEPO%, -1 * CANTI, NPEDI&, NUORITPE%, "", "", "", 0, 0, OCOM$)
'                    If CI% > 0 Then
'                        TXTX$ = ""
'                        Call REPLA(TXTX$, CODEXT$(CI%), 1, 20)
'                        Call REPLA(TXTX$, DESCRIT0$(CI%), 21, 40)
'                        Call REPLA(TXTX$, FORMATNUM$(CANTIDAD, "F10.1"), 62, 11)
'                        Call REPLA(TXTX$, Space$(16), 74, 16)
'                        Call REPLA(TXTX$, FORMATNUM$(PRECONIVA#, "F12.2"), 91, 12)
'                        Call REPLA(TXTX$, FORMATNUM$(XVALO(NPROVX), "F7.0"), 104, 8)
'                        Call REPLA(TXTX$, FORMATNUM$(NUORIT%, "F4"), 120, 4)
'                        ASIGLOAHP.List2.AddItem (TXTX$)
'                    End If
'                    'LUEGO SE EJECUTA CON ESTE CODIGO QUE ESTA DESPUES DE LA LINEA: Call TESTCAE09.PIDENUEVOCAE
'                    'HABILITANDO UN TIMER Y EN LA MEDIDA QUE YA HAYA CERRADO EL SERVICIO DE CAE GENERA EL PREOCESO QUE ESTA DENTRO DEL FRM ASIGLOAHP.
''                    If InStr(EPAC$, "AHP") > 0 Then
''                      Timer3.Enabled = True ' AHP 11
''                    End If
'
'
'        End If
'    End If
'    '
    Return

    '
999 End Sub
Sub GrabaNroRemitoEnCierreOFEnMovisto(NROREMITO$)
    'ESTO ES PARA QUE PRESENTE EN LA PANTALLA DE UBICACIONES CON QUE REMITO SALIO
    'SE GRABA EN EL REGISTRO QUE SE GENERA EN EL MOVISTO CUANDO SE CIERRA LA O.F.
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
       If BCODECTP.List1.ListCount > 0 Then
          For U& = 1 To BCODECTP.List1.ListCount
            idMov& = XVALO(TRIM$(Mid$(BCODECTP.List1.List(U& - 1), 100, 9)))
            If idMov& > 0 Then
                CONDI$ = "MOVISTO_ID = " & idMov&
                Call ACTUREGISTRO("MOVISTO", "DocuOrig", CONDI$, NROREMITO$, REGAF&, "NOMESS")
            End If
          Next U&
       End If
    End If
End Sub

Sub BORRARUBICACION()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
       
       If BCODECTP.List1.ListCount > 0 Then
          For U& = 1 To BCODECTP.List1.ListCount
            idMov& = XVALO(TRIM$(Mid$(BCODECTP.List1.List(U& - 1), 100, 9)))
            
            If idMov& > 0 Then
                CONDI$ = "ID_MOVISTO = " & idMov&
                Call BORRAREGISTROS("UBILOTE_ENVAP", CONDI$, REGAF&, "NOMESS")
            End If
          Next U&
       End If
    End If
End Sub

Sub REMITO_A_WALDBOTT(REMIX$, Optional FILTROEXTRA$, Optional TABLA$)
        Dim RS1 As New ADODB.Recordset
        CotizacionParaPasarapesos = ROUND(XVALO(Label15), 2)
        PASO% = 0: VUELTA% = 0
        VectCliente.cargarvector
        CONDI$ = "CodiMovi = 'RT'"
        CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMIX$ & "' "
        '
        If TRIM$(FILTROEXTRA$) <> "" Then
           CONDI$ = CONDI$ + " AND " & FILTROEXTRA$ ' EN EL CASO QUE SE QUIERA LLAMAR A LA RUTINA PARA QUE GENERE LOS MOVIMIENTOS
        End If                                      ' EN WALDBOTT MOVIMIENTOS MAL GRABADOS QUE DEBERIAN TENER LA MARCA 9999 EN CAMPO VerificadoPor
        '
        SQLX$ = "SELECT MONEVAL,TIPOCAM,COD_INTE,COD_EXTE,NuorItem,DESCRI_LAR, CantSalid, cantingre,NUMEORDPED, NUPEDCLI, ValoUnitIva, Nume_Clie, FECHMOV,IDENLOTE,COD_EXTE"
        SQLX$ = SQLX$ + " FROM Movisto WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & " "
        SQLX$ = SQLX$ + " ORDER BY NuOrItem"
        Dim rs As New ADODB.Recordset
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             NROPED& = XVALO(!NUPEDCLI)
             If PASO% = 0 Then
                'DATOS DE TIPO DE MONEDA
                MONEDAX% = XVALO(!MoneVal)
                MONEDAINGRX = "L": MONEDAEMISX = "L": TICAM# = 1 'POR DEFECTO L LOCAL
                If MONEDAX% > 1 Then
                   If CotizacionParaPasarapesos < 1 Then 'CAMBIO PARA QUE CONSIDERE PEDIDO REALIZADO EN DOLARES A PESOS
                     TICAM# = XVALO(!TIPOCAM) ' SOLO ENTRA SI NO SE AGREGO COTIZACION AL INICIAR LA TRANSACCION
                     MONEDAINGRX = "E": MONEDAEMISX = "E": TICAM# = TICAM#
                   End If
                End If
                
                NC& = XVALO(!Nume_Clie)
                CONDVTA% = XVALO(VALOBADA("PEDENCA", "CONDPAG", "NROPED= " & NROPED&, "NOMESS"))
                VEND& = XVALO(VALOBADA("PEDENCA", "VENDED", "NROPED= " & NROPED&, "NOMESS"))
                DOMIENTREG$ = VALOBADA("PEDENCA", "DOMIENT", "NROPED= " & NROPED&, "NOMESS")
                DOMIENTREG$ = REPLACAR$(DOMIENTREG$, "'", "")
                
                DOMICLIX$ = VectCliente.Domicilio(NC&)
                DOMICLIX$ = REPLACAR$(DOMICLIX$, "'", "")
                RAZONSOC$ = VectCliente.nombre(NC&)
                FE% = CVINT(!FECHMOV)
                
                PASO% = 1
             End If
             J& = J& + 1
             CI1 = XVALO(!cod_inte)
             CODI$ = SAFETEXT$(!Cod_Exte)
             DESCRIART$ = SAFETEXT$(!DESCRI_LAR)
             DESCRIART$ = REPLACAR$(DESCRIART$, "'", "")
             ORDPED% = XVALO(!NUMEORDPED)
             REMILAR$ = REMIX$
             CANTSALID1 = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F15.0"))
             CANTINGRE1 = TRIM$(FORMATNUM$(XVALO(!CANTINGRE), "F15.0"))
             PRE = ROUND(XVALO(VALOBADA("PEDDETA", "PREUNID", "NROPED= " & NROPED& & " AND CODIINT = " & CI1, "NOMESS")), 4)
             If CotizacionParaPasarapesos > 0 And MONEDAX% > 1 Then 'SI TIENE VALOR EL LABEL15 CONVIERTE EL
                PRE = ROUND(PRE * CotizacionParaPasarapesos, 4)
             End If
             COSTSINIVA = PRE
             LOTE$ = SAFETEXT$(!idenlote)
             GoSub GRABAR_EN_WALDBOTT
             .MoveNext
          Loop
       End With
       GoTo 75
GRABAR_EN_WALDBOTT:
    If VUELTA% < 1 Then
       Call abreBaDaWald
       VUELTA% = 1

       On Error GoTo ERROR_GUARDAR
       'BUSCO DATOS DE CODIGO RENGLON 001 Y PROVINCIA BA PARA PASAR ESTOS DATOS A PDCABEZA
       SQLX$ = "SELECT RENGLON,PROVINCIA FROM CLTESENTREGA WHERE CLIENTE = '" & FORMATOCON0$(NC&, 7) & "'"
       Dim RSW As New ADODB.Recordset
       Set RSW = BASEWALD.Execute(SQLX$)
       RENGLONX$ = SAFETEXT$(RSW!RENGLON)
       PROVINCIAX$ = SAFETEXT$(RSW!PROVINCIA)
       RSW.Close
       Set RSW = Nothing
       
       'BUSCO EN LA TABLA CLIENTES EL VALOR CEL CAMPO LISTAPREC PARA LUEGO AGREGAR EN PDITEMS EN CAMPO LISTA (ANTES GRABABA '01' FIJO)
       LISTAPRECIOSW$ = "01" 'POR DEFECTO
       SQLX$ = "SELECT LISTAPREC FROM CLIENTES WHERE NUMERO = '" & FORMATOCON0$(NC&, 7) & "'"
       Set RSW = BASEWALD.Execute(SQLX$)
       If Not RSW.EOF Then
          LISTAPRECIOSW$ = TRIM$(SAFETEXT$(RSW!LISTAPREC))
       End If
       RSW.Close
       Set RSW = Nothing
       
       'ahora quieren aqui la cotizacion Gabriel 09/03/21
       If CotizacionParaPasarapesos < 1 Then CotizacionParaPasarapesos = 1
       DOMIENTREG$ = FORMATNUM$(CotizacionParaPasarapesos, "F9.2")
       '***************************************************
       BASEWALD.BeginTrans
       SQLX$ = "INSERT INTO PDCABEZA  "
       SQLX$ = SQLX$ + "(TIPO ,CLIENTE,NROPED,FECEMISION,CONDVENTA,VENDEDOR,ENTREGA,RAZONSOC,NROREAL,CODIGO,FECRECEP,"
       SQLX$ = SQLX$ + " TRANSPORTE,ESTADO,OPERACION,FECALTA,USUARIO,MOTIVO,CONGELA,MONEDAINGR,MONEDAEMIS,COTIZACION,REFERENCIA,"
       SQLX$ = SQLX$ + " ENTREGARNG,PROENT)"
       SQLX$ = SQLX$ + " VALUES("
       SQLX$ = SQLX$ + "'R'"
    
       SQLX$ = SQLX$ + ",'" & FORMATOCON0$(NC&, 7) & "'"  ' FORMATEA AGREGANDOLE 0 A LA IZQUIERDA
       SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMILAR$) & "'"
       SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
       SQLX$ = SQLX$ + ",'" & SAFETEXT$(CONDVTA%) & "'"
       SQLX$ = SQLX$ + ",'" & FORMATOCON0$(VEND&, 3) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(DOMIENTREG$, 40) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(RAZONSOC$, 60) & "'"
       SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMILAR$) & "'"
       SQLX$ = SQLX$ + ",'55'"
       SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
       SQLX$ = SQLX$ + ",'000'"
       SQLX$ = SQLX$ + ",'00'"
       SQLX$ = SQLX$ + ",'01'"
       SQLX$ = SQLX$ + ", '" & NOWSERVER & "'"
       SQLX$ = SQLX$ + ",'SUPER'"
       SQLX$ = SQLX$ + ",'01'"
       SQLX$ = SQLX$ + ",'S'"
       SQLX$ = SQLX$ + ",'" & MONEDAINGRX & "'"
       
       SQLX$ = SQLX$ + ",'" & MONEDAEMISX & "'"
       SQLX$ = SQLX$ + "," & TRIM(FORMATNUM$(TICAM#, "F16.2"))
       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(Text20), 14) & "'"

       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(RENGLONX$), 3) & "'"
       SQLX$ = SQLX$ + ",'" & Left$(TRIM$(PROVINCIAX$), 2) & "'"
       SQLX$ = SQLX$ + ")"
 'Call mostrarDatosx(SQLX$)
       BASEWALD.Execute (SQLX$)
    End If
    
    SQLX$ = "INSERT INTO PDITEMS  "
    SQLX$ = SQLX$ + "(TIPO ,CLIENTE,NROPED,ITEM,ARTICULO,DESCART,PRECIO,CANTPED,BULTPED,FECRECEP,PRESENTA,"
    SQLX$ = SQLX$ + " LISTA,DEPSALIDA,FECALTA,USUARIO)"
    SQLX$ = SQLX$ + " VALUES('R'"
    SQLX$ = SQLX$ + ",'" & FORMATOCON0$(NC&, 7) & "'"
    SQLX$ = SQLX$ + ",'" & FORMATO_RTO_WALDB(REMILAR$) & "'"
    SQLX$ = SQLX$ + ",'" & FORMATOCON0$(J&, 3) & "'"

    SQLX$ = SQLX$ + ",'" & Left$(CODI$, 20) & "'"
    SQLX$ = SQLX$ + ",'" & Left$(DESCRIART$, 50) & "'"
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(ROUND(COSTSINIVA, 4)) & "'"
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(CANTSALID1) & "'"
    SQLX$ = SQLX$ + ",'" & SAFETEXT$(CANTSALID1) & "'"
'    SQLX$ = SQLX$ + "," & CANTSALID1
'    SQLX$ = SQLX$ + "," & CANTSALID1
    SQLX$ = SQLX$ + ",'" & CVDAT(FE%) & "'"
    
    SQLX$ = SQLX$ + ",'01'"
    SQLX$ = SQLX$ + ", '" & LISTAPRECIOSW$ & "'"
    SQLX$ = SQLX$ + ",'01'"
    SQLX$ = SQLX$ + ", '" & NOWSERVER & "'"
    SQLX$ = SQLX$ + ",'SUPER'"
    SQLX$ = SQLX$ + ")"
    BASEWALD.Execute (SQLX$)
    '
    'AGREGADO A TABLA DESCRIPCIONES PARA INCLUIR MAS DE UNA LINEA DE DESCRIPCION - INGRESA LA DESCRI1 Y LA DESCRI2 DEL MAESTRO
    DESCRIWALD$ = Left$(DESCRIART$, 49) + vbCrLf
    If CODINT(CODI$) > 0 Then
       DESCRIWALD$ = DESCRIWALD$ & Left$(DESCRIT2$(CODINT(CODI$)), 49)
    End If
    
    strSQL = "SELECT * FROM DESCRIPCIONES"
    strSQL = strSQL + " WHERE ARTICULO  = '" & TRIM$(CODI$) & "'"
    Set RS1 = BASEWALD.Execute(strSQL)
    With RS1
     If .EOF Then 'SI NO EXISTE LA AGREGA
        ingresoendescripcion% = 1
        SQLX$ = "INSERT INTO DESCRIPCIONES  "
        SQLX$ = SQLX$ + "(ARTICULO ,TIPO,CLIPROV,NUMERO,DESCRIPCION)"
        SQLX$ = SQLX$ + " VALUES ("
        SQLX$ = SQLX$ + "'" & Left$(CODI$, 20) & "'"
        SQLX$ = SQLX$ + ",'C'"
        SQLX$ = SQLX$ + ",'0'"
        SQLX$ = SQLX$ + ",'" & Left$(CODI$, 20) & "'"
        SQLX$ = SQLX$ + ",'" & DESCRIWALD$ & "'"
        SQLX$ = SQLX$ + ")"
    Else
       'SI EXISTE LA ACTUALIZA
        SQLX$ = "UPDATE DESCRIPCIONES  SET TIPO = 'C',CLIPROV = '0',NUMERO =  '" & Left$(CODI$, 20) & "', DESCRIPCION = '" & DESCRIWALD$ & "'"
        SQLX$ = SQLX$ + " WHERE ARTICULO = '" & TRIM$(CODI$) & "'"
        ingresoendescripcion% = 1
    End If
    End With
    BASEWALD.Execute (SQLX$)
    If ingresoendescripcion% = 1 Then
        RS1.Close
        Set RS1 = Nothing
        ingresoendescripcion% = 0
    End If
    
    'FIN GRABACION EN TABLA DESCRIPCIONES.
    
Return
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       BASEWALD.RollbackTrans
       On Error Resume Next
       rs.Close
       Set rs = Nothing
       SQLX$ = "UPDATE   Movisto SET VerificadoPor = 9999 WHERE " & CONDI$: FLEXCONN.Execute (SQLX$)
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
75    BASEWALD.CommitTrans
      rs.Close
      Set rs = Nothing
      SQLX$ = "UPDATE Movisto SET VerificadoPor = 0 WHERE " & CONDI$: FLEXCONN.Execute (SQLX$)
      Set VectCliente = Nothing
      Call COMUNI("Transferencia de Remito a Base Waldbott Realizada Correctamente")
    End If
99
End Sub


Sub ACTUALIZA_PEDIDOS(NCLI, NROREMIT$)
    HOII% = HOY(HOS$)
    FECHX = FETEXCOR1$(HOII%)
    Dim PEDTMP As ADODB.Recordset
    Set PEDTMP = New ADODB.Recordset
    '
    '**************************************************************************************************
    'CRITERIO:
    '*********
    'RECORRE EL DETALLE DEL REMITO Y SI TIENE ASIGNADO UN PEDIDO LE DA PRIORIDAD A ESTE
    'QUE  ENTRARIA EN EL PRIMER RECORDSET, EN CASO CONTRARIO INGRESA EN EL SEGUNDO
    'EN ESTE ULTIMO VA DESCONTANDO LA CANTIDAD REMITIDA DE LOS PEDIDOS CONSIDERANDO LA FECHA DE ENTREGA
    'DESDE EL MAS VIEJO AL MAS NUEVO.
    '**************************************************************************************************
    '
    For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
       ACUMU_ENTREGADO# = 0
       CANTI = CAN#
       If NUPEDIORI& > 0 Then
          
          If CantRegistros&("PEDENCA", "STATUS < 9 and NroPed =  " & NUPEDIORI&, "NOMESS") > 0 Then
            CONDI$ = "STATUS < 8 AND CantEnt < (CanPed - 0.005) AND NroPed =  " & NUPEDIORI& & " AND CODIINT = " & CI% & " AND NROCLIE = " & NCLI
            SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " Where " & CONDI$
            SQLX$ = SQLX$ + " ORDER BY FESOLENT"
            Set PEDTMP = READSET(SQLX$)
            With PEDTMP
               Do While Not .EOF
                  SALDO = XVALO(!CANPED) - XVALO(!CANTENT)
                  If SALDO > 0 Then
                     If CANTI > SALDO Then CANTI = SALDO
                     NUOR% = XVALO(!NuOrItem)
                     CANTIDAD_ENT = XVALO(!CANTENT) + CANTI
                     Call ACTUREGISTRO("PEDDETA", "CantEnt", CONDI$, CANTIDAD_ENT, REGAF&, "NOMESS")
                     ACUMU_ENTREGADO# = ACUMU_ENTREGADO# + CANTI
                     CANTI = CAN# - ACUMU_ENTREGADO#
                     If CANTI <= 0 Then Exit Do
                  End If
                  .MoveNext
               Loop
            End With
          End If
          CAN# = CAN# - ACUMU_ENTREGADO#
          PEDTMP.Close
          Set PEDTMP = Nothing
       End If
       '
       '
       If CAN# > 0 Then
            CONDI$ = " Status < 8  "   ' para que no este anulada ni cumplida
            CONDI$ = CONDI$ + " AND CODIINT = " & CI%
            CONDI$ = CONDI$ + " AND NROCLIE = " & NCLI
            CONDI$ = CONDI$ + " AND ((CanPed - CantEnt) >= 0.05 OR (TipoPed = 'ABIERTO' AND FinVigen >= '" & FECHX & "')) "
            SQLX$ = " SELECT * "
            SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE " & CONDI$
            '
            SQLX$ = SQLX$ + " ORDER BY FESOLENT,NroPed"
            '
            Set PEDTMP = READSET(SQLX$)
            With PEDTMP
               Do While Not .EOF
                  SALDO = XVALO(!CANPED) - XVALO(!CANTENT)
                  If SALDO > 0 Then
                     NUOR% = XVALO(!NuOrItem)
                     NUPED& = XVALO(!NROPED)
                     CI% = XVALO(!CODIINT)
                     If CANTI > SALDO + 0.05 Then CANTI = SALDO
                     CANTIDAD_ENT = XVALO(!CANTENT) + CANTI
                     Condi1$ = "NUORITEM = " & NUOR% & " AND NROPED = " & NUPED& & " AND CODIINT = " & CI% & "  And NROCLIE = " & NCLI
                     Call ACTUREGISTRO("PEDDETA", "CantEnt", Condi1$, CANTIDAD_ENT, REGAF&, "NOMESS")
                     Call ACTUREGISTRO("PEDDETA", "CanSaldo", Condi1$, XVALO(!CANPED) - CANTIDAD_ENT, REGAF&, "NOMESS")
                     Call ACTUREGISTRO("PEDDETA", "NroRemi", Condi1$, NROREMIT$, REGAF&, "NOMESS")
                     
                     ACUMU_ENTREGADO# = ACUMU_ENTREGADO# + CANTI
                     CANTI = CAN# - ACUMU_ENTREGADO#
                     If CANTI <= 0 Then Exit Do
                  End If
                  .MoveNext
               Loop
            End With
          End If
       
       PEDTMP.Close
       Set PEDTMP = Nothing
       '
    Next YX&
    '
    If Abs(CANTI) > 0 Then
        TTXX$ = "La Cantidad Remitida de " + TRIM$(FORMATNUM$(CAN#, "F10.1")) + " del Artículo"
        TTXX$ = TTXX$ + "\'" + CODEXT$(CI%) + "' con el Comprobante '" + NROREMIT$ + "' "
        TTXX$ = TTXX$ + "\Supera la Cantidad de/los Pedidos Pendientes Para Este Item "
        Call MENSERR(24, TTXX$)
   End If

End Sub
'
Private Sub Timer1_Timer()
   '
   TOTITEMS% = ULTREG&("!BOLREDET")
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
End Sub

'
Sub CONAPLI(APLINVO$, TERMINA%)
   '
   CODCARPE$ = BUSCAVALOR(PARAPLI$, "CODCARPE =") ' codigo de carpeta de import / export
   '
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "ANUREMI"
           Call OPREMSTA07.Command25_Click                 ' ANULACION DE REMITO
       Case "ELIMPENFAC"
           Call OPREMSTA07.Command6_Click                 '  ELIMNAR DE LISTA PENDIENTE DE FACTURACION
       Case "DEVDESMAT"
           Call OPREMSTA07.Command9_Click                 ' DESPACHO DE MATERIALES
       Case "CORCANDES"
           Call OPREMSTA07.Command10_Click                 ' CORRECCION DE CANTIDADES DESPACHADAS
       Case "REASREPED"
           Call OPREMSTA07.Command12_Click                 ' REASIGNAR REMITOS A PEDIDOS
       Case "GENCODBAR"
           Call Command14_Click                            ' GENERACION DE CODIGOS DE BARRA

       
       'CONSULTAS Y LISTADOS
       '
       Case "RESENTPRO"
           Call OPREMSTA07.Command4_Click                 'RESUMEN DE ENTREGA POR PRODUCTO
       Case "RESENTCLI"
           Call OPREMSTA07.Command3_Click                 ' RESUMEN ENTREGA POR CLIENTE
       Case "SECNUMREM"
           Call OPREMSTA07.Command5_Click                 ' SECUENCIAL PO NUMERO DE COMPROBANTE
       Case "ENTPENFAC"
           Call OPREMSTA07.Command8_Click                ' ENTREGAS PENDIENTES DE FACTURAR
       Case "CONDETNREM"
           Call OPREMSTA07.Command14_Click                ' CONSULTA DETALLE POR NUMERO DE REMITO
       Case "PACKLIST"
           Call OPREMSTA07.Command18_Click
       '
       'CONFIGURACION GENERAL DE REMITOS
       '
       Case "CONGENREM"
           Command13_Click
           
       Case Else
           TERMINA% = 0
   End Select
   '
End Sub

 Sub LIACOBSE()
    '
    For i = 0 To 3
       Text16(i) = "" ' LIMPIA CAJAS DE OBSERVACIONES
    Next i
    Me.lblNroPed = ""
    Me.Label16 = ""
    Static ESGUIDI%
    '
    If ESGUIDI% = 0 Then
       ESGUIDI% = (-1)
       If InStr(UCase$(EMPREAC$), "GUIDI") > 0 Then ESGUIDI% = 1
    End If
    '
    Text22.Enabled = True
    Text16(0).Enabled = True ' ACTIVA PRIMER RENGLON DE OBSERVACIONES
    '
    ACCESOATOYOTA.Label1 = ""
    '
    If ESGUIDI% = 1 Then
       Text22 = "" ' LIMPIA ORDEN SUMINISTRO
       For i = 0 To 3
           Text16(i) = "" ' LIMPIA CAJAS DE OBSERVACIONES
       Next i
       '
       Text8 = ""
       Text12 = ""
       Text15 = ""
       Text18 = ""
       Text19 = ""
       Text14 = ""
       Text13 = ""
       Text21 = ""
       Text17 = ""
       Text20 = ""
       Text23 = ""
       '
    End If
    '
End Sub


Sub AJUSTACANT()
    '
    FIN& = ULTREG&("!BOLREDET")
    For UI& = 1 To FIN&
      XX$ = REGLEIDO$("!BOLREDET", UI&)
      CCII% = CVI(Mid$(XX$, 83, 2))
      If ESTRAZABLE%(CCII%) > 0 Then
          For i = 0 To ASIGLOT07.List1.ListCount - 1
              YY$ = ASIGLOT07.List1.List(i)
              CANTASI# = XVALO(Mid$(YY$, 64, 10))
              REGIST& = XVALO(Mid$(YY$, 80, 5))
              If UI& = REGIST& And CANTASI# > 0 Then
                  Call REPLA(XX$, MKD$(CANTASI#), 89, 8)
                  Call GRAREG("!BOLREDET", XX$, UI&)
              End If
          Next i
      End If
    Next UI&
    '
End Sub

Function CONTROLCARGA(APROP%, CONTROCARG)
    APROP% = 0
    IMPSEG = CONTROCARG
    FIN& = ULTREG&("!BOLREDET")
    For UI& = 1 To FIN&
      XX$ = REGLEIDO$("!BOLREDET", UI&)
      CCII% = CVI(Mid$(XX$, 83, 2))
      CANTASI# = CVD(Mid$(XX$, 89, 8))
      COSTUNI = COSUNI(CCII%)
      Import = Import + (CANTASI# * COSTUNI)
      If IMPSEG < Import Then
        Call MENSERR(24, "Importe del Remito Supera el Control de Carga")
        APROP% = 1: GoTo 99
      End If
    Next UI&
    '
99 CONTROLCARGA = APROP%
    '
End Function
