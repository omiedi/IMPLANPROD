VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form ACRE_GES07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAESTRO DE PROVEEDORES"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11895
   Icon            =   "ACRE_GES07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11895
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   120
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
      Left            =   1155
      MaxLength       =   6
      TabIndex        =   143
      Top             =   315
      Width           =   730
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0050765A&
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   11040
      ScaleHeight     =   8385
      ScaleWidth      =   1035
      TabIndex        =   109
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H000C93DC&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   120
         Top             =   7560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0076C9F5&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   121
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_GES07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Ayuda en Línea"
         Top             =   960
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   640
         Left            =   105
         Picture         =   "ACRE_GES07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   325
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Left            =   105
         Picture         =   "ACRE_GES07.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6645
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
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
         Picture         =   "ACRE_GES07.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Emisión de Listados"
         Top             =   5160
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command25 
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
         Left            =   105
         Picture         =   "ACRE_GES07.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Imprime Ficha del Proveedor"
         Top             =   5850
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Sort"
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
         Picture         =   "ACRE_GES07.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Ordena Base de Datos de Proveedores y Gastos"
         Top             =   3690
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Picture         =   "ACRE_GES07.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Alta de Nuevo Proveedor"
         Top             =   1620
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
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
         Picture         =   "ACRE_GES07.frx":1AA0
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Baja Lógica de la Cuenta"
         Top             =   3000
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Save"
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
         Picture         =   "ACRE_GES07.frx":1DAA
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   2280
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command23 
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
         Picture         =   "ACRE_GES07.frx":20B4
         Style           =   1  'Graphical
         TabIndex        =   110
         ToolTipText     =   "Carga de Observaciones de Entrega"
         Top             =   4380
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10320
      Top             =   120
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Recupera CUIT"
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
      Left            =   2280
      TabIndex        =   97
      Top             =   160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command9 
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
      Height          =   330
      Left            =   1890
      Picture         =   "ACRE_GES07.frx":23BE
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command10 
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
      Height          =   315
      Left            =   6350
      Picture         =   "ACRE_GES07.frx":26C8
      TabIndex        =   94
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   8880
      TabIndex        =   64
      Top             =   210
      Width           =   1990
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
         Height          =   500
         Left            =   725
         Picture         =   "ACRE_GES07.frx":29D2
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   160
         Picture         =   "ACRE_GES07.frx":2CDC
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
      End
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "ACRE_GES07.frx":2FE6
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "TIPO DE CUENTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   6650
      TabIndex        =   61
      Top             =   210
      Width           =   2115
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "Cta. de Gastos"
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
         Height          =   215
         Left            =   315
         TabIndex        =   63
         Top             =   525
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7D9CC&
         Caption         =   "C. Corriente"
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
         Left            =   320
         TabIndex        =   62
         Top             =   250
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox Text6 
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
      Left            =   4410
      TabIndex        =   1
      Top             =   315
      Width           =   1965
   End
   Begin VB.TextBox Text3 
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
      Left            =   1155
      TabIndex        =   2
      Top             =   735
      Width           =   5340
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   180
      TabIndex        =   25
      Top             =   1200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   13097420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "ACRE_GES07.frx":434C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Anotaciones"
      TabPicture(1)   =   "ACRE_GES07.frx":4368
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cuenta Corriente"
      TabPicture(2)   =   "ACRE_GES07.frx":4384
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Height          =   6300
         Left            =   -74965
         TabIndex        =   58
         Top             =   315
         Width           =   10610
         Begin VB.Frame Frame19 
            Height          =   400
            Left            =   480
            TabIndex        =   136
            Top             =   3060
            Width           =   9690
            Begin VB.Label Label77 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
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
               Left            =   80
               TabIndex        =   142
               Top             =   150
               Width           =   1455
            End
            Begin VB.Label Label78 
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
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
               Left            =   1350
               TabIndex        =   141
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "Emisión"
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
               Left            =   3800
               TabIndex        =   140
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Importe Original"
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
               Left            =   4680
               TabIndex        =   139
               Top             =   150
               Width           =   1440
            End
            Begin VB.Label Label81 
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimiento"
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
               Left            =   6530
               TabIndex        =   138
               Top             =   150
               Width           =   1080
            End
            Begin VB.Label Label83 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo"
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
               Left            =   8670
               TabIndex        =   137
               Top             =   150
               Width           =   840
            End
         End
         Begin VB.Frame Frame17 
            Height          =   400
            Left            =   480
            TabIndex        =   129
            Top             =   300
            Width           =   9690
            Begin VB.Label Label76 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo"
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
               Left            =   8710
               TabIndex        =   135
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label93 
               BackStyle       =   0  'Transparent
               Caption         =   "Haber"
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
               Left            =   7100
               TabIndex        =   134
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label92 
               BackStyle       =   0  'Transparent
               Caption         =   "Debe"
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
               Left            =   5600
               TabIndex        =   133
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label89 
               BackStyle       =   0  'Transparent
               Caption         =   "Número"
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
               Left            =   2550
               TabIndex        =   132
               Top             =   150
               Width           =   840
            End
            Begin VB.Label Label57 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
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
               Left            =   1260
               TabIndex        =   131
               Top             =   150
               Width           =   1455
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   80
               TabIndex        =   130
               Top             =   150
               Width           =   975
            End
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   480
            TabIndex        =   103
            Top             =   3480
            Width           =   9675
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2280
            Left            =   480
            TabIndex        =   102
            Top             =   720
            Width           =   9670
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4960
            TabIndex        =   106
            Top             =   5685
            Width           =   1695
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   7970
            TabIndex        =   105
            Top             =   5685
            Width           =   1695
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL GENERAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   600
            TabIndex        =   104
            Top             =   5780
            Width           =   1695
         End
         Begin VB.Label Label94 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   480
            TabIndex        =   107
            Top             =   5520
            Width           =   9675
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   6300
         Left            =   35
         TabIndex        =   29
         Top             =   315
         Width           =   10610
         Begin VB.Frame Frame4 
            Caption         =   "OTROS DATOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   760
            Left            =   5200
            TabIndex        =   91
            Top             =   1890
            Width           =   5255
            Begin VB.TextBox Text24 
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
               Left            =   2100
               TabIndex        =   92
               Text            =   " "
               Top             =   315
               Width           =   2975
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre de Fantasía"
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
               Left            =   105
               TabIndex        =   93
               Top             =   405
               Width           =   1905
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "CONTABLE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1625
            Left            =   5200
            TabIndex        =   49
            Top             =   4530
            Width           =   5255
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   6
               Left            =   4880
               TabIndex        =   24
               Top             =   735
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   7
               Left            =   1725
               TabIndex        =   22
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   7
               Left            =   1155
               TabIndex        =   21
               Text            =   " "
               Top             =   315
               Width           =   570
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   6
               Left            =   2625
               TabIndex        =   23
               Text            =   " "
               Top             =   735
               Width           =   2270
            End
            Begin VB.Label DETEXT 
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
               Index           =   7
               Left            =   1995
               TabIndex        =   57
               Top             =   315
               Width           =   3080
            End
            Begin VB.Label DETEXT 
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
               Index           =   6
               Left            =   420
               TabIndex        =   56
               Top             =   1155
               Width           =   4655
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo"
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
               Left            =   315
               TabIndex        =   51
               Top             =   420
               Width           =   765
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cuenta Madre"
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
               Left            =   630
               TabIndex        =   50
               Top             =   840
               Width           =   1905
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "COMERCIALES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   5200
            TabIndex        =   47
            Top             =   2730
            Width           =   5255
            Begin VB.TextBox Text4 
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
               Left            =   1155
               TabIndex        =   96
               Text            =   " "
               Top             =   735
               Width           =   3930
            End
            Begin VB.TextBox Text2 
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
               Height          =   465
               Left            =   1155
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   82
               Text            =   "ACRE_GES07.frx":43A0
               Top             =   1100
               Width           =   3920
            End
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   4
               Left            =   1730
               TabIndex        =   14
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Height          =   300
               Index           =   4
               Left            =   1155
               TabIndex        =   13
               Text            =   " "
               Top             =   315
               Width           =   570
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Orden del Cheque"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   -15
               TabIndex        =   95
               Top             =   660
               Width           =   1095
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Observa- ciones"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   90
               TabIndex        =   81
               Top             =   1150
               Width           =   1005
            End
            Begin VB.Label DETEXT 
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
               Index           =   4
               Left            =   1995
               TabIndex        =   55
               Top             =   315
               Width           =   3080
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Condición de Pago"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   240
               TabIndex        =   48
               Top             =   240
               Width           =   840
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "IMPUESTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   170
            TabIndex        =   42
            Top             =   4140
            Width           =   4845
            Begin VB.CommandButton Command12 
               Caption         =   "Otros"
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
               Left            =   3600
               TabIndex        =   144
               ToolTipText     =   "Datos Retención Ganancias Cliente Activo"
               Top             =   720
               Width           =   1020
            End
            Begin VB.CommandButton Command22 
               Caption         =   "Configurar"
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
               Left            =   3600
               TabIndex        =   101
               ToolTipText     =   "Datos Retención I.Brutos Proveedor Activo"
               Top             =   1560
               Width           =   1020
            End
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   19
               Top             =   1155
               Width           =   175
            End
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   1
               Left            =   1260
               TabIndex        =   16
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Height          =   300
               Index           =   1
               Left            =   945
               MaxLength       =   2
               TabIndex        =   15
               Text            =   " "
               Top             =   315
               Width           =   360
            End
            Begin VB.TextBox Text33 
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
               Left            =   2310
               MaxLength       =   24
               TabIndex        =   17
               Top             =   315
               Width           =   2355
            End
            Begin VB.TextBox COTEXT 
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
               Height          =   300
               Index           =   2
               Left            =   945
               MaxLength       =   2
               TabIndex        =   18
               Text            =   " "
               Top             =   1155
               Width           =   360
            End
            Begin VB.TextBox Text34 
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
               Left            =   2310
               MaxLength       =   15
               TabIndex        =   20
               Text            =   " "
               Top             =   1155
               Width           =   2355
            End
            Begin VB.Label DETEXT 
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
               Index           =   2
               Left            =   945
               TabIndex        =   54
               Top             =   1545
               Width           =   2520
            End
            Begin VB.Label DETEXT 
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
               Index           =   1
               Left            =   945
               TabIndex        =   53
               Top             =   705
               Width           =   2520
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Posición IVA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   220
               TabIndex        =   46
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CUIT"
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
               Left            =   1680
               TabIndex        =   45
               Top             =   420
               Width           =   540
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Categoría Ing. Brutos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   0
               TabIndex        =   44
               Top             =   1080
               Width           =   900
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Número Ing. IIBB"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   1470
               TabIndex        =   43
               Top             =   1050
               Width           =   750
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "AGENDA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2180
            Left            =   170
            TabIndex        =   37
            Top             =   1890
            Width           =   4845
            Begin VB.CommandButton Command98 
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
               Left            =   4440
               TabIndex        =   98
               ToolTipText     =   "Cargar Mail de Contactos"
               Top             =   1050
               Width           =   175
            End
            Begin VB.TextBox Text36 
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
               Height          =   280
               Left            =   1365
               TabIndex        =   59
               Text            =   " "
               Top             =   1755
               Width           =   3300
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
               Height          =   280
               Left            =   1365
               TabIndex        =   12
               Text            =   " "
               Top             =   1395
               Width           =   3300
            End
            Begin VB.TextBox Text30 
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
               Height          =   280
               Left            =   1365
               TabIndex        =   9
               Text            =   " "
               Top             =   315
               Width           =   3300
            End
            Begin VB.TextBox Text31 
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
               Height          =   280
               Left            =   1365
               TabIndex        =   10
               Text            =   " "
               Top             =   675
               Width           =   3300
            End
            Begin VB.TextBox Text32 
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
               Height          =   280
               Left            =   1365
               TabIndex        =   11
               Text            =   " "
               Top             =   1035
               Width           =   3090
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Horario"
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
               Left            =   -630
               TabIndex        =   60
               Top             =   1840
               Width           =   1905
            End
            Begin VB.Label Label13 
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
               Left            =   -630
               TabIndex        =   41
               Top             =   1480
               Width           =   1905
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   -630
               TabIndex        =   40
               Top             =   400
               Width           =   1905
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   -630
               TabIndex        =   39
               Top             =   760
               Width           =   1905
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
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
               Left            =   -630
               TabIndex        =   38
               Top             =   1120
               Width           =   1905
            End
         End
         Begin VB.Frame Frame5 
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
            Height          =   1695
            Left            =   170
            TabIndex        =   31
            Top             =   105
            Width           =   10295
            Begin VB.CommandButton BOTSEL 
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
               Height          =   315
               Index           =   0
               Left            =   1785
               TabIndex        =   7
               Top             =   1155
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   0
               Left            =   1365
               TabIndex        =   6
               Text            =   "  "
               Top             =   1155
               Width           =   465
            End
            Begin VB.TextBox Text25 
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
               Left            =   1365
               TabIndex        =   3
               Text            =   " "
               Top             =   315
               Width           =   8750
            End
            Begin VB.TextBox Text26 
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
               Left            =   1365
               TabIndex        =   4
               Text            =   " "
               Top             =   720
               Width           =   1305
            End
            Begin VB.TextBox Text27 
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
               Left            =   3885
               TabIndex        =   5
               Text            =   " "
               Top             =   720
               Width           =   6230
            End
            Begin VB.TextBox Text29 
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
               Left            =   7035
               TabIndex        =   8
               Text            =   " "
               Top             =   1155
               Width           =   3080
            End
            Begin VB.Label DETEXT 
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
               Index           =   0
               Left            =   2100
               TabIndex        =   52
               Top             =   1155
               Width           =   4005
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Localidad"
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
               Left            =   1890
               TabIndex        =   36
               Top             =   840
               Width           =   1905
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Provincia"
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
               Left            =   80
               TabIndex        =   35
               Top             =   1260
               Width           =   1185
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Calle y Número"
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
               Left            =   -650
               TabIndex        =   34
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Codigo Postal"
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
               Left            =   -650
               TabIndex        =   33
               Top             =   840
               Width           =   1905
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "País"
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
               Left            =   6195
               TabIndex        =   32
               Top             =   1260
               Width           =   750
            End
         End
      End
      Begin VB.Frame Frame7 
         Height          =   6300
         Left            =   -74965
         TabIndex        =   122
         Top             =   315
         Width           =   10610
         Begin VB.TextBox Text28 
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
            Height          =   4190
            Left            =   600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   127
            Top             =   1440
            Width           =   9780
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Cortar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   5085
            TabIndex        =   126
            Top             =   5625
            Width           =   1170
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Copiar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   6270
            TabIndex        =   125
            Top             =   5625
            Width           =   1170
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Pegar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   7455
            TabIndex        =   124
            Top             =   5625
            Width           =   1275
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Seleccionar Todo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   8730
            TabIndex        =   123
            Top             =   5625
            Width           =   1650
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ANOTACIONES - NOVEDADES - COMENTARIOS"
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
            Left            =   360
            TabIndex        =   128
            Top             =   720
            Width           =   9705
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   28
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   27
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Controles no visibles"
      Height          =   4425
      Left            =   1470
      TabIndex        =   68
      Top             =   9000
      Visible         =   0   'False
      Width           =   6105
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2100
         TabIndex        =   74
         Text            =   " "
         Top             =   1260
         Width           =   570
      End
      Begin VB.TextBox COTEXT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2100
         TabIndex        =   73
         Text            =   " "
         Top             =   420
         Width           =   570
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   315
         Index           =   3
         Left            =   2670
         TabIndex        =   72
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
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
         Height          =   315
         Index           =   5
         Left            =   2670
         TabIndex        =   71
         Top             =   1275
         Width           =   175
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   70
         Text            =   " "
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         TabIndex        =   69
         Text            =   " "
         Top             =   2040
         Width           =   1710
      End
      Begin VB.Label LTEXT23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   780
         TabIndex        =   90
         Top             =   3675
         Width           =   1500
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   89
         Top             =   3720
         Width           =   645
      End
      Begin VB.Label LTEXT37 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   780
         TabIndex        =   88
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Cte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   87
         Top             =   3405
         Width           =   645
      End
      Begin VB.Label LTEXT36 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   780
         TabIndex        =   86
         Top             =   3045
         Width           =   1500
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Cre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   105
         TabIndex        =   85
         Top             =   3090
         Width           =   645
      End
      Begin VB.Label LTEXT35 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1830
         TabIndex        =   84
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cheq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1155
         TabIndex        =   83
         Top             =   2565
         Width           =   645
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   80
         Top             =   1365
         Width           =   1710
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precios:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   79
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Std:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   78
         Top             =   1770
         Width           =   1710
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2835
         TabIndex        =   77
         Top             =   1770
         Width           =   1905
      End
      Begin VB.Label DETEXT 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2940
         TabIndex        =   76
         Top             =   420
         Width           =   2880
      End
      Begin VB.Label DETEXT 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2940
         TabIndex        =   75
         Top             =   1260
         Width           =   2880
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ACRE_GES07.frx":43A2
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ACRE_GES07.frx":45D6
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Cuenta"
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
      Left            =   120
      TabIndex        =   108
      Top             =   210
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3255
      TabIndex        =   30
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   26
      Top             =   630
      Width           =   870
   End
   Begin VB.Menu Archiv 
      Caption         =   "Transacciones"
      Begin VB.Menu NueCli 
         Caption         =   "Nuevo Proveedor"
         Begin VB.Menu CtaCte 
            Caption         =   "Cuenta Corriente"
         End
         Begin VB.Menu CtaGas 
            Caption         =   "Cuenta de Gastos"
         End
      End
      Begin VB.Menu mnuGuardarCambios 
         Caption         =   "Guardar Cambios"
      End
      Begin VB.Menu DeshCam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu Baj_Cli 
         Caption         =   "Baja Cuenta Activa"
      End
      Begin VB.Menu separa3 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu List 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu Busc 
         Caption         =   "Búsqueda de Proveedor"
      End
      Begin VB.Menu separa2 
         Caption         =   "-"
      End
      Begin VB.Menu ListGral 
         Caption         =   "Listado General y Clasificado"
      End
      Begin VB.Menu CtaInac 
         Caption         =   "Listado de Cuentas Inactivas"
      End
      Begin VB.Menu LidePo1 
         Caption         =   "Listado de Direcciones Postales"
      End
      Begin VB.Menu LiDiMa1 
         Caption         =   "Listado de Direcciones de Mails"
      End
      Begin VB.Menu mnuProveedoresParaIB 
         Caption         =   "Listado Proveedores para Ing. Brutos"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu php32 
         Caption         =   "-"
      End
      Begin VB.Menu ConstanciaAfip 
         Caption         =   "Descarga Constancia de Inscripción Afip"
      End
      Begin VB.Menu separa4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu DeuLim 
         Caption         =   "Deudas y Limites de Crédito"
      End
   End
   Begin VB.Menu Conf 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu separa5 
         Caption         =   "-"
      End
      Begin VB.Menu CondPag 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu ProZonVta 
         Caption         =   "Provincias y Zonas de Venta"
      End
      Begin VB.Menu GruProv 
         Caption         =   "Grupo de Proveedores"
      End
      Begin VB.Menu mnuRetencionGanancias 
         Caption         =   "Retenciones de Ganancias"
      End
      Begin VB.Menu mnuRetencionIB 
         Caption         =   "Retenciones de Ingresos Brutos"
      End
      Begin VB.Menu mnuRetencionIva 
         Caption         =   "Retenciones de IVA Facturas M"
      End
      Begin VB.Menu PLh8 
         Caption         =   "-"
      End
      Begin VB.Menu Confgnr 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuPlanDeCuentas 
         Caption         =   "Plan de Cuentas Contables"
      End
   End
   Begin VB.Menu mnuImpuestos 
      Caption         =   "Impuestos"
      Visible         =   0   'False
      Begin VB.Menu mnuCategoriaGananciasProveedor 
         Caption         =   "Categoria Ganancias Proveedor Activo"
      End
   End
End
Attribute VB_Name = "ACRE_GES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

Private Sub Baj_Cli_Click()
   Call Command24_Click
End Sub

Private Sub Busc_Click()
   Call Command20_Click
End Sub

Private Sub Command10_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 1
   RSLX$ = REGSEL$("INDIPROV")
   If Len(RSLX$) > 4 Then
      CODIPRO$ = TRIM$(VALACT1$("INDIPROV"))
      NCLIE = NUMEPRO(CODIPRO$)
      If NCLIE > 0 Then
         Text1.TEXT = TRIM$(Str$(NCLIE))
      End If
   End If
   '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call Command1_Click
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   PopupMenu mnuImpuestos
   Command12.Enabled = True
End Sub

Sub SetearCategoriaGananciasProveedor()
   '
   If ULTREG&("DARGAPRO") < 1 Then
     Call GENDAREGAPRO
   End If
   '
   NPRO% = XVALO(TRIM$(Text1.TEXT))
   If NPRO% > 0 Then
     If REGICLI&((-1) * NPRO%) <> 0 Then
       '
       If ULTREG&("TACRETG") < 1 Then
         Call GRAREG("TACRETG", Chr$(0) + "Exento ", 1)
         Call GRAREG("TACRETG", Chr$(1) + "Profesionales", 2)
         Call GRAREG("TACRETG", Chr$(1) + "Bienos de Uso y de Cambio", 3)
         Call SETULTREG("TACRETG", 3)
       End If
       '
       CATMAX% = ULTREG&("TACRETG")
       RECORD$ = FILTERGETRECORD$("DARGAPRO", 1, MKI$(NPRO%))
       Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Then
         Call REPLA(RECORD$, Chr$(2), 3, 1)
         Call REPLA(RECORD$, Space$(32), 4, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAREGAN", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > CATMAX% Then
           Call MENSERR(24, "Mal Categoría de Ret. de Ganancias")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 4, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         '
         If CATIB1% > 0 Then
           Call REPLA(RECORD$, Space$(32), 4, 32)
         End If
         '
         Call REPLA(RECORD$, MKI$(NPRO%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 36, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 56, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 58, 2)
         Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NPRO%), RECORD$)
         Call CIERRARCH("DARGAPRO")
         '
       End If
     End If
   End If

End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    Screen.MousePointer = 11
    Call ABREPROVE
    '
    SQLX$ = "SELECT * FROM PROVED12" ' Validacion de CUIT en el caso que no sea valido agregar el ultimo digito y en el caso que sea valido modificar
    Dim provtmp As ADODB.Recordset
    Set provtmp = New ADODB.Recordset
    provtmp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With provtmp
      Do While Not .EOF
        NUCUIT$ = TRIM$(SAFETEXT(!Cuit_Cli))
        If VALICUIT%(NUCUIT$) <> 1 Then
          For i& = 0 To 9
            CUITTT$ = NUCUIT$ & "-" & i&
            If VALICUIT%(CUITTT$) = 1 Then
              !Cuit_Cli = CUITTT$
              .Update
              Exit For
            End If
          Next i&
        End If
10      .MoveNext
      Loop
    End With
    '
    provtmp.Close
    Set provtmp = Nothing
    '
    Screen.MousePointer = 1
    Call MENSERR(24, "Proceso Terminado")
    Command16.Enabled = True

End Sub



Private Sub Command22_Click()
   Static CTXX%
   Command22.Enabled = False
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE = 0 Then GoTo 99
   
1  OpcionesIB% = COMALTER%("Jurisdiccion IIBB\\Buenos Aires\C.A.B.A.\Misiones\Tucumán")
   If OpcionesIB% < 1 Or OpcionesIB% > 4 Then GoTo 99
   
   If OpcionesIB% = 1 Then GoTo 5
   If OpcionesIB% = 2 Then GoTo 25
   If OpcionesIB% = 3 Then GoTo 50
   If OpcionesIB% = 4 Then GoTo 75
   
   
   ' *** INGRESOS BRUTOS BUENOS AIRES
   
5  If EXISTE%(LUDAT$ + "DARIBPRO.RFS") < 1 Then
     Call COPYSTRU("DAPIBCLI", "DARIBPRO")
   End If
   '
   If ULTREG&("DARIBPRO") < 1 Then
     Call GENDARIBPRO
   End If
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE > 0 Then
     If REGICLI&((-1) * NCLIE) <> 0 Then
     ' If ECOARCH$("PROVED1", MKI$(Int(NCLIE))) <> "" Then
       If TICUEPRO%((-1) * NCLIE) = 0 Then ' solo cuentas corrientes
         '
         If ULTREG&("TACATIBP") < 5 Then
           Call GRAREG("TACATIBP", Chr$(0) + AJUSTI$("EXENTO", 31) + String$(32, 0), 1)
           Call GRAREG("TACATIBP", Chr$(1) + AJUSTI$("BIENES MUEBLES", 31) + String$(32, 0), 2)
           Call GRAREG("TACATIBP", Chr$(2) + AJUSTI$("LOCACIONES SERVICIOS", 31) + String$(32, 0), 3)
           Call GRAREG("TACATIBP", Chr$(3) + AJUSTI$("EVENTUALES", 31) + String$(32, 0), 4)
           Call GRAREG("TACATIBP", Chr$(4) + AJUSTI$("PARTIDAS ESPECIALES", 31) + String$(32, 0), 5)
           Call CIERRARCH("TACATIBP")
         End If
         '
         RECORD$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(Int(NCLIE)))
         Call REPLA(RECORD$, MKI$(Int(NCLIE)), 1, 2)
         '
10       OBX$ = OBJPLA$("CARDARETIB", RECORD$)
         If OBX$ <> "" Then
           RECORD$ = OBX$
           CATIB1% = Asc(Mid$(RECORD$, 3, 1))
           If ECOARCH$("TACATIBP", Chr$(CATIB1%)) = "" Then
             Call MENSERR(24, "Mal Categoría I.Brutos")
             GoTo 10
           End If
           If CATIB1% < 1 Then
             If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
               Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
               GoTo 10
             End If
           End If
           If CATIB1% > 0 Then
             Call REPLA(RECORD$, Space$(32), 9, 32)
           End If
           PORCONVE = CVS(Mid$(RECORD$, 5, 4))
           If PORCONVE < 0 Or PORCONVE > 100 Then
             Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
             GoTo 10
           End If
           '
           Call REPLA(RECORD$, MKI$(Int(NCLIE)), 1, 2)
           Call REPLA(RECORD$, USERNAME$, 41, 20)
           Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
           Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
           Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(Int(NCLIE)), RECORD$)
           Call CIERRARCH("DARIBPRO")
           '
         End If
       End If
     End If
   End If
   
   GoTo 1
   
   ' *** INGRESOS BRUTOS CABA
25:
   CategoriaActual% = XVALO(VALOBADA("PROVED12", "CategoriaIBCABA_Cli", "NUME_CLI=" & CStr(NCLIE)))
   DenominacionCategoria$ = ECOARCH$("TACAIBCA", Chr$(CategoriaActual%))
        
   Call SELECHO("TACAIBCA/CATEGORIA ACTUAL: " & DenominacionCategoria$)
   CategoriaNueva% = XVALO(VALACT1$("TACAIBCA"))
   If CategoriaNueva% > 0 Then
        Call ACTUREGISTRO("PROVED12", "CategoriaIBCABA_Cli", "NUME_CLI=" & CStr(NCLIE), CategoriaNueva%, REGAF&)
   End If
   
   GoTo 1
   
   
   ' *** INGRESOS BRUTOS MISIONES
50:
   CategoriaActual% = XVALO(VALOBADA("PROVED12", "CategoriaIBMISIONES_Cli", "NUME_CLI=" & CStr(NCLIE)))
   DenominacionCategoria$ = ECOARCH$("TACAIBMI", Chr$(CategoriaActual%))
        
   Call SELECHO("TACAIBMI/ACTUAL: " & DenominacionCategoria$)
   CategoriaNueva% = XVALO(VALACT1$("TACAIBMI"))
   If CategoriaNueva% > 0 Then
        Call ACTUREGISTRO("PROVED12", "CategoriaIBMISIONES_Cli", "NUME_CLI=" & CStr(NCLIE), CategoriaNueva%, REGAF&)
   End If
   
   GoTo 1


   ' *** INGRESOS BRUTOS TUCUMAN
75:
   CategoriaActual% = XVALO(VALOBADA("PROVED12", "CategoriaIBTUCU_Cli", "NUME_CLI=" & CStr(NCLIE)))
   DenominacionCategoria$ = ECOARCH$("TACAIBTU", Chr$(CategoriaActual%))
        
   Call SELECHO("TACAIBTU/ACTUAL: " & DenominacionCategoria$)
   CategoriaNueva% = XVALO(VALACT1$("TACAIBTU"))
   Call ACTUREGISTRO("PROVED12", "CategoriaIBTUCU_Cli", "NUME_CLI=" & CStr(NCLIE), CategoriaNueva%, REGAF&)
   
   GoTo 1


99 Command22.Enabled = True

End Sub

Sub GENDARIBPRO()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARIBPRO")
    Call REPLA(REBLA$, USERNAME$, 41, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 61, 2)
    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 63, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        If TICUEPRO%((-1) * NCLIX%) = 0 Then
          Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
          Call REPLA(RECORD$, Chr$(1), 3, 1)
          Call REPLA(RECORD$, MKS$(100), 5, 4)
          Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NCLIX%), RECORD$)
        End If
      End If
    Next U&
    Call CIERRARCH("DARIBPRO")
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    '
    'VALIDAR SI TIENE UN CODIGO ACTIVO
    If TRIM$(Text1) = "" Then
       Call COMUNI("No Hay Un Proveedor Activo")
       GoTo 99
    End If
    '
    NCLIE = (-1) * XVALO(Text1.TEXT)
    If STATCLI(NCLIE) < 1 Then
      Call COMUNI("Proveedor Inexistente o No Válido")
      GoTo 99
    End If
    '
    'PRESENTAR FORMULARIO
    FILTX$ = TRIM$(Text1) & ";" & TRIM$(Text3)
    Call STDFORM("CAROBSEN", FILTX$)
    '
99  Command23.Enabled = True

End Sub

Private Sub Command98_Click()
    'ACCESO A LOS MAILS DE CONTACTOS
    Command98.Enabled = False
    COCLI$ = TRIM$(Text6)
    If COCLI$ <> "" Then
      FRMMAILS.BackColor = &HAFB9E7
      FRMMAILS.Picture1.BackColor = &HAFB9E7
      FRMMAILS.Frame8.BackColor = &HAFB9E7
      FRMMAILS.Frame15.BackColor = &HAFB9E7
      FRMMAILS.Picture5(0).BackColor = &HAFB9E7
      FRMMAILS.Picture5(1).BackColor = &HAFB9E7
      FRMMAILS.Picture5(2).BackColor = &HAFB9E7
      FRMMAILS.Picture5(3).BackColor = &HAFB9E7
      FRMMAILS.Picture5(4).BackColor = &HAFB9E7
      FRMMAILS.Picture3.BackColor = &H404080
      FRMMAILS.Label26 = TRIM$(Text6)
      FRMMAILS.Show 1
    Else
      Call COMUNI("Código de Proveedor No Válido o Inexistente")
      GoTo 99
    End If
99  Command98.Enabled = True
End Sub

Private Sub CondPag_Click()
OPMASPR07.Option1(6).Value = True
Call OPMASPR07.Command16_Click
End Sub

Private Sub Confgnr_Click()
Call Command14_Click
End Sub

Private Sub ConGen_Click()
Call Command13_Click
End Sub

Private Sub ConstanciaAfip_Click()

   cuitIngresado$ = TRIM$(Text33)
   If cuitIngresado$ <> "" And VALICUIT%(cuitIngresado$) > 0 Then
        Dim oConstanciaAfip As New DatosEmpresaAfip
        
        oConstanciaAfip.DescargarConstanciaInscripcion (cuitIngresado$)
        Set oConstanciaAfip = Nothing
   Else
        Call MENSERR(24, "Falta Cuit o el Ingresado no es Válido")
   End If
   
End Sub

Private Sub CtaCte_Click()
Option1.Value = True
Call Command21_Click
End Sub

Private Sub CtaGas_Click()
Option2.Value = True
Call Command21_Click
End Sub

Private Sub CtaInac_Click()
Call AMAPRLIS07.Command5_Click
End Sub

Private Sub cuencon_Click()
  OPMASPR07.Option1(8).Value = True
  Call OPMASPR07.Command16_Click
End Sub

Private Sub DeshCam_Click()
Call Command2_Click
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub exit_Click()
   Call Command3_Click
End Sub


Private Sub Form_Load()
    '
    Command9.Enabled = False
    Text1.Enabled = False
    Command10.Enabled = False
    Text6.Enabled = False

    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    
    Call COPYSTRU("PROVED1", "SELPRBX") 'Se agrego por si no existe el archivo
    Call FRESHECHO("!SELPRBX")    'Se agrego para limpiar el archivo
    Call SETULTREG("!SELPRBX", 0) 'de busqueda por CUIT, Tel, Direccio, ETC
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call ABRECONEXION
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    'ACTUALIZACION DE INDICE DE PROVEEDORES PARA QUE MANTENER INFORMACION EN AMBAS EMPRESAS
    Call ABREPROVE
    UPDATPRO% = 1
    Call GENECOPROV
    '
    HOOO = HOY(HO$) ' MODIFICACION PARA MEDITEA BOTON 'RECUPERA CUIT' EN EL CASO QUE SUPERA FECHA 16/12/08 BOTON NO VISIBLE
    If FECHATEX(HOOO) > "16/12/08" Then Command16.Visible = False
    '
    ARCHEX$(0) = "PROVINC"
    ARCHEX$(1) = "CAPOSIV"
    ARCHEX$(2) = "CARETIB"
    ARCHEX$(3) = "LISTAS"
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    ARCHEX$(6) = "ECUECON"
    ARCHEX$(7) = "GRUCOPRO" 'NUEVO
    'ARCHEX$(7) = "GRUCOCLI"
    '
    Call CARTAVALI
    '
    If ULTREG&("BKDATCLI") < 1 Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + "DEUDOR1.BAK") > 0 Then
        If EXISTE%(LU$ + "DEUDOR2.BAK") > 0 Then
          Screen.MousePointer = 11
          jj& = 0
          Dim D1 As String * 32
          Dim D2 As String * 128
          N1% = FreeFile
          Open LU$ + "DEUDOR1.BAK" For Random Access Read Write Shared As #N1% Len = 32
          N2% = FreeFile
          Open LU$ + "DEUDOR2.BAK" For Random Access Read Write Shared As #N2% Len = 128
          For k& = 1 To LOF(N1%) / 32
            Get #N1%, k&, D1$
            NC% = CVI(Left$(D1$, 2))
            If NC% > 0 Then
              Get #N2%, k&, D2$
              D3$ = D1$ + D2$
              jj& = jj& + 1
              Call GRAREG("BKDATCLI", D3$, jj&)
            End If
          Next k&
          Call SETULTREG("BKDATCLI", jj&)
          Call CIERRARCH("BKDATCLI")
          Screen.MousePointer = 1
        End If
      End If
    End If
    '
    Call ACTUALZACION_LISTAS_SELECCION("PROVED1", "MAESTRO DE PROVEEDORES")
    
    ' INICIALIZA LA CATEGORIA DE INGRESOS BRUTOS TUCUMAN, UTILIZADA LUEGO PARA LA RETENCION DE DICHO IMPUESTO DESDE LA ORDEN DE PAGO
    Call ACTUREGISTRO("proved12", "categoriaibtucu_cli", "categoriaibtucu_cli is null", "0", REGAF&)
    
    If CONTROL("FACTURA", "VAPOSIVA") = "SI" Then
        Timer1.Enabled = True
    End If
    
    'ver!!!
'    Call ABRECONEXION
    'UTILIZA_SKIN% = 1
    'Call APLICACIONSKINS(Me, Picture1)
    Command9.Enabled = True
    Text1.Enabled = True
    Command10.Enabled = True
    Text6.Enabled = True
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub
'

Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   If Index = 0 Then
     If CantRegistros("PROVINCIA", "ID > 0 ", "NOMESS") > 0 Then
        If CantRegistros("LOCALIDAD", "ID > 0 ", "NOMESS") > 0 Then
             CONDI$ = "CODFLEX = '" & COTEXT(0) & "'"
             PROVLOCA.TXTPROVNUM = VALOBADA("PROVINCIA", "ID", CONDI$, "NOMESS")
             PROVLOCA.TXTLOCALNUM = ""
             PROVLOCA.TXTNAMEPROVINCIA = DETEXT(0)
             PROVLOCA.TXTPROVLET = COTEXT(0)
             PROVLOCA.TXTLOCALIDAD = Text27
             PROVLOCA.TXTCODPOSTAL = Text26
             
             PROVLOCA.Show 1
             If PROVLOCA.APROBO > 0 Then
                 DETEXT(0) = PROVLOCA.TXTNAMEPROVINCIA
                 COTEXT(0) = PROVLOCA.TXTPROVLET
                 Text27 = PROVLOCA.TXTLOCALIDAD
                 Text26 = PROVLOCA.TXTCODPOSTAL
             End If
             Exit Sub
        End If
     End If
   End If
   
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENPRO
   Call GRANOTAPRO
   GRABATODO% = 0
End Sub

Sub Command13_Click()
   AMAPRLIS07.Show 1
End Sub

Sub Command14_Click() 'ver esto
   Command14.Enabled = False
   ' Esto que sigue es para que traiga la tabla correspondiente a
   ' proveedores y no la de clientes.
   If DERACCE%("STPACRE", "Setup de Proveedores") = 2 Then
   RETARCHE$ = ARCHE$(7)
   ARCHE$(7) = "GRUCOPRO"
   OPMASPR07.Show 1
   ARCHE$(7) = RETARCHE$
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   ' Call MENSERR(24, "No Disponible en la Presente Versión.\ \Los Ordenamientos son Automáticos\al Generar los Listados.")
   Call DEPURADUP
   Command15.Enabled = True
End Sub

Sub DEPURADUP()
    ' RUTINA PARA REEMPLAZAR COMILLAS SIMPLES POR UN GUION
     Screen.MousePointer = 11
     Call ABREPROVE
     '
     FIN& = CantRegistros("PROVED12", "NUME_CLI > 0 ")
     SQLX$ = "Select * from PROVED12 "
     Dim CLIENTMP As ADODB.Recordset
     Set CLIENTMP = New ADODB.Recordset
25   On Error Resume Next
     CLIENTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
     '
     j& = 0
     CAMBIO% = 0
     With CLIENTMP
        Do While Not .EOF
           j& = j& + 1
           Call TRACE(20, j&, FIN&)
           NCX& = XVALO(!nume_cli)
           QTR& = CantRegistros("PROVED12", "NUME_CLI = " & NCX&)
           If QTR& > 1 Then
              .Delete
              CLIENTMP.Close
              GoTo 25
           End If
        .MoveNext
        Loop
     End With
     CLIENTMP.Close
     Set CLIENTMP = Nothing
     Screen.MousePointer = 1
     '
     UPDATPRO% = 1
     Call GENECOPROV
     Call COMUNI("Proceso Completo")
     '
End Sub


Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENPRO
   Call GRANOTAPRO
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATPROV
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
   Call COPYSTRU("PROVED1", "SELPRBX")
   If ULTREG&("!SELPRBX") > 1 Then GoTo 20 'Se modifico el nombre del archico !SELECLI por SELPRBX
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Proveedor por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     jj& = 0
     'Call COPYSTRU(" PROVED1", "SELPRBX")
     '
20   Call APERBASDAT("PROVE")
     'ver
     Dim Rproved12 As ADODB.Recordset
     Set Rproved12 = New ADODB.Recordset
     Rproved12.CursorType = adOpenKeyset
     Rproved12.LockType = adLockReadOnly
     Rproved12.Open FILTSQL$("SELECT * FROM PROVED12"), FLEXCONN, , , adCmdText
     '
     On Error Resume Next
     With Rproved12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE = !nume_cli
               RACLI$ = !raso_cli
               XX$ = MKI$(Int(NCLIE)) + RACLI$
               jj& = jj& + 1
               Call GRAREG("!SELPRBX", XX$, jj&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELPRBX", jj&)
     '
     Screen.MousePointer = 1
     If jj& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf jj& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELPRBX", 1), 2))))
       Else
         Call FRESHECHO("!SELPRBX")
         Call SELECHO("!SELPRBX")
         NCLIX$ = VALACT1$("!SELPRBX")
         If NCLIX$ <> "" Then
             Text1 = NCLIX$
           Else
             Call SETULTREG("!SELPRBX", 0)
             GoTo 10
         End If
     End If
   End If
   Command20.Enabled = True
End Sub
 Private Sub Command21_Click()
'   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   If EmpresaNoRealizaCambios% > 0 Then Exit Sub

   DRXX% = DERACCE%("ALTAPRO", "Alta de Nuevo Proveedor")
   If DRXX% < 2 Then
     If DRXX% = 1 Then
       Call MENSERR(24, "Usuario sin Derecho de Escritura\para la Aplicación Solicitada.\  \Consulte a su Supervisor del Sistema.")
     End If
     Exit Sub
   End If
   '
   Call APERBASDAT("PROVE")
        '(llama a abreconexion y setea la conexionstring al sql)
   'defino el recordset
   Dim Tranprov As ADODB.Recordset
   Command21.Enabled = False
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(CONTROL$("", "PRONUPRO"))
   If PRONUCLI& < 0 Then
        NROPRO& = XVALO(VALOBADA("proved12", "max(nume_cli)", "nume_cli>0"))
        GoTo 3
   End If
2  NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1

   Set Tranprov = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Proved12 WHERE Nume_Cli = " & NROPRO&))
   If Not (Tranprov.EOF And Tranprov.BOF) Then 'si la encuentra
      GoTo 3
   End If
   '
   PRFX$ = Left$(TRIM$(CONTROL("ALTAPROV", "CODIPROV")), 4)
   If PRFX$ = "" Then PRFX$ = "PRO"
   COPRO$ = PRFX$ + "-" + TRIM$(Str$(NROPRO&))
   '
   NCLIE = NROPRO&
   
   ModoAltaProveedor% = COMALTER%("Modalidades de Alta de Proveedor.\ \Ingreso Manual\u Obteniendo Datos desde Afip a Partir del Cuit\(Requiere Conexión a Internet)\\Alta Manual\Obtener Datos desde Afip")
   If ModoAltaProveedor% < 1 Or ModoAltaProveedor% > 2 Then ModoAltaProveedor% = 1
   
   If ModoAltaProveedor% = 1 Then GoTo 8
   
   ' *** ALTA MEDIANTE CONEXION CON AFIP
   CarpetaSistema$ = App.Path + "\"
   If EXISTE%(CarpetaSistema$ & "toolsafip.EXE") = 0 Then
        Call MENSERR(24, "Imposible Utilizar esta Modalidad.\Faltan Utilitarios de Conexión con Afip.")
        GoTo 8
   End If
   
   cuitIngresado$ = InputBox$("Ingrese el Cuit del Proveedor", "Alta de Proveedor", "")
   If cuitIngresado$ <> "" Then
   
        If Len(cuitIngresado$) < 9 Then
             Opcion% = COMALTER("Se Ha Detectado el Ingreso de un DNI. ¿Desea Obtener el Cuit Asociado?\\Si, Obtener\No, Continuar")
             If Opcion% = 1 Then
                cuitIngresado$ = ObtieneCuitDesdeDni(cuitIngresado$)
             Else
                GoTo 8
             End If
        End If
        
        Dim oAltaProveedor As New DatosEmpresaAfip
        
        If oAltaProveedor.ObtenerDatosEmpresaAfip(cuitIngresado$, response$) Then
           NCLINUE% = NCLIE
           DCLINUE$ = oAltaProveedor.razonSocial
           COCLINU$ = "PRO-" + TRIM$(Str$(NCLINUE%))
           CUTT$ = cuitIngresado$
           PIVCLI% = oAltaProveedor.condicionIva
           DOMI$ = oAltaProveedor.domicilioFiscal
           LOCA$ = oAltaProveedor.localidadDomicilioFiscal
           PROVINCIA$ = oAltaProveedor.provinciaDomicilioFiscal
           CPOSTAL$ = oAltaProveedor.codigoPostalDomicilioFiscal
           Set oAltaProveedor = Nothing
           
           GoTo 15
        Else
            Call MENSERR(24, "Problemas al Obtener Datos de Afip.\ \Verifique que el Cuit sea Correcto.\Intente Nuevamente o Realice Alta Manual.")
            Screen.MousePointer = 1
            GoTo 99
        End If
   Else
        Call MENSERR(24, "Falta Ingresar Cuit")
        Screen.MousePointer = 1
        GoTo 99
   End If
   
8  VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, COPRO$, 1, 12)
   Call REPLA(VDEF$, MKI$(Int(NCLIE)), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   If Option1.Value = True Then
       OBX$ = OBJPLA$("ALTAPROVACC", VDEF$)
     Else
       Call REPLA(VDEF$, Chr$(0), 77, 1)   ' IVA Consumidor Final
       OBX$ = OBJPLA$("ALTACUEGAST", VDEF$)
   End If
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 29999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "PRO-" + TRIM$(Str$(NCLINUE%))
     EPAC$ = TRIM$(UCase$(EMREAC$))
     If InStr(EPAC$, "BRASBO") > 0 Then
       COCLINU$ = "FOR-" + TRIM$(Str$(NCLINUE%))
     End If
     Call REPLA(VDEF$, COCLINU$, 1, 12)
   End If
   '
   Set Tranprov = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Proved12 where Nume_cli =" & NCLINUE))
   If Not (Tranprov.EOF And Tranprov.BOF) Then 'si la encuentra
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Proveedor")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 And PIVCLI% <> 7 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
15 If CUTT$ <> "" Then
     Set Tranprov = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Proved12 WHERE Cuit_Cli = '" & CUTT$ & "'"))
     If Not (Tranprov.EOF And Tranprov.BOF) Then
       If Tranprov!nume_cli > 0 Then
          NCLIX% = (-1) * Tranprov!nume_cli
           TTALTER$ = "Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(rasocli$(NCLIX%)) + "' (" + TRIM$(Str$(Abs(NCLIX%))) + ")"
          If COMALTER%(TTALTER$ + "\\Cancelar\Registrar Igual") <> 2 Then
             GoTo 10
          End If
       End If
      End If
   End If
   '
   Screen.MousePointer = 11
   TICPRO% = 0
   If Option2.Value = True Then
     TICPRO% = 1
     PIVCLI% = 0
     CUTT$ = ""
   End If
   Screen.MousePointer = 11
   
20 Call ABREPROVE
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Nume_Cli = " & NCLINUE%
   End If
   '
   Dim Rproved12 As ADODB.Recordset
   Set Rproved12 = New ADODB.Recordset
   Rproved12.CursorType = adOpenDynamic
   Rproved12.LockType = adLockPessimistic
   Rproved12.Open FILTSQL$("Select * from Proved12 where nume_cli=" & NCLINUE%), FLEXCONN, , , adCmdText
   If Not (Rproved12.EOF And Rproved12.BOF) Then
      GoTo 20
   End If
   With Rproved12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !nume_cli = TRIM$(Str$(NCLINUE%))
      !raso_cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !domi_cli = DOMI$
      !loca_Cli = LOCA$
      !Cpos_Cli = CPOSTAL$
      !Prov_Cli = PROVINCIA$
      !Stat_Cli = 1
      !Tcpr_Cli = TICPRO%
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
   End With
   Rproved12.Close
   Set Rproved12 = Nothing
   '
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "nume_cli=" & NCLINUE%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Proved12", "MINEPARTS", "Proved12", CONDI$)
   End If
   '
   UPDATPRO% = 1
   Call GENECOPROV
   '
   AGRECLI% = 1
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   SSTab1.Tab = 0
   Call CARDATPROV
   Call CARTAVALI
   '
   For Index = 0 To 7
     If Index <> 1 Then
       VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
       If ModoAltaProveedor% = 2 And Index = 0 Then GoTo 70 '  OBVIA LA PROVINCIA CUANDO OBTIENE LOS DATOS DE AFIP
       COTEXT(Index) = "": DETEXT(Index) = ""
       If VAPRE$ <> "" Then
         TV11% = TIPFIELD(ARCHE$(Index), 1)
         LF11% = LENFIELD(ARCHE$(Index), 1)
         ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
         ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
         If ECO$ <> "" Then
             COTEXT(Index).TEXT = VAPRE$
             DETEXT(Index).Caption = ECO$
           Else
             COTEXT(Index).TEXT = ""
             DETEXT(Index).Caption = ""
         End If
       End If
     End If
70 Next Index
   '
   Text3 = DCLINUE$
   COTEXT(1) = TRIM$(Str$(PIVCLI%))
   Text33 = TRIM$(CUTT$)
   Call Command1_Click           ' PARA GUARDAR LOS DATOS
   '
   Screen.MousePointer = 1
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   If DERACCE%("BAJAPRO", "Baja de Cuenta de Proveedor") < 2 Then
     GoTo 99
   End If
   '
   NCLIE& = (-1) * Abs(XVALO(TRIM$(Text1.TEXT)))
   If Abs(NCLIE&) < 1 Or REGICLI(NCLIE&) < 1 Then
      RESP1% = 2
      GoTo 90
   End If
   '
   If RESP1% = 0 Then
    RESP1% = ALTERNA%("Dar de Baja Cuenta de Proveedor.\Restauración de Cuenta de Proveedor.")
    If RESP1% < 1 Or RESP1% > 2 Then
        GoTo 99
    End If
   End If
   '
   If RESP1% = 1 Then 'SI QUIERE ANULAR
     If XVALO(Text1.TEXT) > 0 Then
       If XVALO(Text1.TEXT) <= 32767 Then
         NCLIE& = XVALO(Text1.TEXT)
         '
         TTXX$ = "La Baja de un Proveedor\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Pagos o Compras."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         '
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
           strsql = ("Update Proved12 set stat_cli= -1 where nume_cli= ") & NCLIE&
            FLEXCONN.Execute (FILTSQL$(strsql))
           Call GENPROVANU
           '
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTAPRO
           '
           UPDATPRO% = 1
           Call GENECOPROV
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
         End If
         '
       End If
     End If
   End If
   '
90 If RESP1% = 2 Then 'SI QUIERE RESTAURAR
     Call GENPROVANU
     Call BAJALDISCO
     RPA$ = ""
     RPA$ = REGSEL$("INDIPROA") 'PROV ANULADOS
     If Len(RPA$) > 4 Then
       NCLIE1% = CVI(Mid$(RPA$, 63, 2))
       If NCLIE1% > 0 Then
         NC% = NCLIE1%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call Command1_Click
       End If
     End If
   End If
   '
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "nume_cli= " & NCLIE&
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Proved12", "MINEPARTS", "Proved12", CONDI$)
   End If
   
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
   PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   Call CARDATPROV
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(2)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   'VALIDA DE DONDE VIENE PARA QUE SI ES AMACLI O AMAPRO HAGA UN CALL FINAL
   If UCase$(App.EXEName) = "ARCHIG07" Then
      On Error Resume Next
      Unload Me
      MARCH_GES07.Enabled = True
      MARCH_GES07.SetFocus
      On Error GoTo 0
   Else
      Call SETULTREG("!SELPRBX", 0) 'de busqueda por CUIT, Tel, Direccio, ETC
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If


'   If APLINVO$ <> "" Then
'    Hide
'   Else
'    On Error Resume Next
'    Unload Me
'    MARCH_GES07.Enabled = True
'    MARCH_GES07.SetFocus
'    On Error GoTo 0
'   End If
   '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   '
   Screen.MousePointer = 1
   'Call SELECHO("PROVED1")
   'NCLIE = XVALO(TRIM$(VALACT1$("PROVED1")))
   NCLIE = ListaSeleccionProveedores("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC% = NCLIE
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub COTEXT_Change(Index As Integer)
   '
   ECO$ = ""
   If TRIM$(COTEXT(Index).TEXT) <> "" Then
      TV11% = TIPFIELD(ARCHEX$(Index), 1)
      LF11% = LENFIELD(ARCHEX$(Index), 1)
      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
   End If
   '
   If ECO$ <> "" Then
        DETEXT(Index).Caption = ECO$
      Else
        DETEXT(Index).Caption = "No Definido"
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub CARDATPROV()
    '
    NCLIE = (-1) * XVALO(TRIM$(Text1.TEXT))
    STCL% = STATCLI%(NCLIE)
    If NCLIE >= 0 Or STCL% < 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = rasocli$(NCLIE)
    Text24.TEXT = TEL2CLI$(NCLIE)
    Text6 = CODICLI$(NCLIE)
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENPRO
      Case 1
        Call CARNOTAPRO
      Case 2
        Call CARESCUENTA
    End Select
    Text1.SelStart = Len(Text1)
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENPRO()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = (-1) * Abs(XVALO(TRIM$(Text1.TEXT)))
    '
    Text2 = CONDICOM$(NC%)
    Text6 = CODICLI$(NC%)
    Text25.TEXT = DOMICLI$(NC%)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NC%)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NC%)
    Text27.Refresh
    If LIMICRE(NC%) < 0 Then
        Text20.TEXT = " Suspendido"
      Else
        Text20.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 12, 2, 2)
    End If
    Text20.Refresh
    'Text23.TEXT = NPROCLI$(NC%)
    'Text23.Refresh
    Text29.TEXT = PAISCLI$(NC%)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NC%)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NC%)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NC%)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NC%)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NC%)
    Text36.Refresh
    Text33.TEXT = CuitCli$(NC%)
    Text33.Refresh
    Text34.TEXT = NIBRCLI$(NC%)
    Text34.Refresh
    Text21.TEXT = TRIM$(DESCUCLI$(NC%))
    Text21.Refresh
    Text4.TEXT = TRIM$(ORDECHEQ(NC%))
    Text4.ForeColor = vbBlack
          If Text4 <> rasocli(NC%) Then Text4.ForeColor = vbRed
    Text4.Refresh
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    COTEXT(2).Refresh
    COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    COTEXT(3).Refresh
    COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    COTEXT(4).Refresh
    COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NC%)))
    COTEXT(5).Refresh
    COTEXT(6).TEXT = TRIM$(CUEMADRE$(NC%))
    COTEXT(6).Refresh
    COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NC%)))
    COTEXT(7).Refresh
    '
    If TICUEPRO%(NC%) = 0 Then ' cuenta corriente
         Option1.Value = True
       Else
         Option2.Value = True
    End If
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENPRO()
        
    If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    NCLIE& = (-1) * Abs(XVALO(TRIM$(Text1.TEXT)))
    If Abs(NCLIE&) < 1 Or REGICLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\  \Número de Proveedor\Inexistente o no Válido.")
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If Abs(STATCLI%(NCLIE)) < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Numero de Proveedor Inexistente")
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If STATCLI(NCLIE&) < 1 Then
      If GRABATODO% = 1 Then
        If COMALTER%("El Proveedor Elegido Figura\Marcado Como 'Dado de Baja'.\  \Desea Re-Activar Esta Cuenta ?\\No, Dejar de Baja\Si, Re-Activar") = 2 Then
          Call APERBASDAT("PROVE")
          strsql = ("Update proved12 set stat_cli = 1 where nume_cli =") & Abs(NCLIE&)
           FLEXCONN.Execute (FILTSQL$(strsql))
          UPDATPRO% = 1
          Call GENECOPROV
          Call CARDATPROV
        End If
      End If
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If GRABATODO% = 1 Then
      If SSTab1.Tab <> 0 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATPRO", "Modificar Datos de Proveedores") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      On Error Resume Next
      Command9.SetFocus
      On Error GoTo 0
      If TRIM$(Text3) = "" Then
        Screen.MousePointer = 11
        Exit Sub
      End If
      If COMALTER%("Se han Modificado Datos del Proveedor\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = rasocli$(NCLIE)
        Text24.TEXT = TEL2CLI$(NCLIE)
        MODIFIC(0) = 0
        Call CARGENPRO
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    '
    Cuit$ = TRIM$(Text33.TEXT)
    If Cuit$ <> "" Then
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      If PIVA% <> 5 Then  ' EXCLUYE COMERCIO EXTERIOR
        If VALICUIT%(Cuit$) <> 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
          SSTab1.Tab = 0
          On Error Resume Next
          Text33.SetFocus
          On Error GoTo 0
          Screen.MousePointer = 1
          Exit Sub
        End If
      End If
    End If
    '
    If NCLIE = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text6) = "" Then
      Call MENSERR(24, "Imposible Registrar !!!\Falta Código de Cuenta")
      Screen.MousePointer = 1
      Exit Sub
    End If
    '
    NPRII% = NUMEPRO%(Text6)
    If NPRII% > 0 Then
      If NPRII% <> Abs(NCLIE) Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Codigo de Cuenta Repetido o no Válido.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social de Proveedor / Cuenta.")
        Exit Sub
    End If
    '
   Call ABREPROVE
   
   Dim Rproved12 As ADODB.Recordset
   Set Rproved12 = New ADODB.Recordset
   Rproved12.CursorType = adOpenDynamic
   Rproved12.LockType = adLockPessimistic
   Rproved12.Open FILTSQL$("Select * from Proved12 where nume_cli=" & Abs(NCLIE&) & "and stat_cli=1"), FLEXCONN, , , adCmdText
   If (Rproved12.EOF And Rproved12.BOF) Then 'si esta vacio
      Call MENSERR(24, "Imposible Registrar")    ' no existe
      Exit Sub
   End If
    '
10  On Error GoTo 0
    With Rproved12
    If TRIM$(Text3) <> TRIM$(!raso_cli) Or UCase$(Text6) <> !Codi_Cli Then UPDATPRO% = 1
      !Codi_Cli = Left$(UCase$(Text6), 12)
      !raso_cli = Left$(Text3, 64)
      !Stat_Cli = 1
      !domi_cli = Left$(Text25, 64)
      !Cpos_Cli = Left$(Text26, 16)
      !loca_Cli = Left$(Text27, 48)
      !Tele_Cli = Left$(Text30, 32)
      !Fant_Cli = Left$(Text24, 32)
      !Cuit_Cli = Left$(TRIM$(Cuit$), 24)
      !Nibr_Cli = Left$(TRIM$(Text34), 24)
      !Caib_Cli = Left$(TRIM$(COTEXT(2)), 2)
      !Prov_Cli = Left$(TRIM$(COTEXT(0)), 2)
      '
      !Ord_Cheq = Left$(Text4, 48) 'ORDEN DEL CHEQUE
      '
      PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
      !Piva_Cli = PIVA%
      '
      CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
      !Caib_Cli = CATIB%
      '
      COPAGI% = VALIMI(COTEXT(4).TEXT, 0, 255)
      If COPAGI% > 0 Then
        If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
        End If
      End If
      !Cpag_Cli = COPAGI%
      !CCom_Cli = RTrim$(Text2)
     '
      CUECO% = VALIMI(COTEXT(7).TEXT, 0, 32767)
      If ECOARCH$("GRUCOPRO", MKI$(CUECO%)) = "" Then
        CUECO% = 0
      End If
      !GrCo_Cli = CUECO%
      '
      CCX% = CUEINT%(COTEXT(6).TEXT)
      !CuMa_Cli = CCX%
      '
      DCC$ = TRIM$(Text21.TEXT)
      !Dsta_Cli = Left$(TRIM$(Text21), 32)
      '
      !Pais_Cli = Left$(Text29, 48)
      !Faxi_Cli = Left$(Text31, 32)
      !Mail_Cli = Left$(Text32, 64)
      !CTAC_CLI = Left$(Text19, 48)
      !Prvd_Cli = Left$(Text23, 24)
      !Hora_Cli = Left$(Text36, 32)
      '
      TICPRO% = 0
      If Option2.Value = True Then TICPRO% = 1
      !Tcpr_Cli = TICPRO%
      '
      .Update
    End With
    '
    Rproved12.Close
    Set Rproved12 = Nothing
    
    If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "nume_cli=" & Abs(NCLIE&) & "and stat_cli=1"
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "Proved12", "MINEPARTS", "Proved12", CONDI$)
    End If
    
    Call GENECOPROV
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTAPRO()
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE = XVALO(Text1.TEXT)
   Text28 = ""
   '
   
   If NCLIE > 0 Then
     If REGICLI((-1) * NCLIE) > 0 Then
       Call APERBASDAT("PROVE")
       Dim Tranprov As ADODB.Recordset
       Set Tranprov = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Proved12 WHERE Nume_Cli = " & NCLIE))
       On Error Resume Next '
       If Err < 1 Then
         Text28 = Tranprov!nota_Cli
       End If
     End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTAPRO()
   '
   NCLIE = XVALO(TRIM$(Text1.TEXT))
   If NCLIE < 1 Then Exit Sub
   '
   If REGICLI((-1) * NCLIE) < 1 Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   '
   If NCLIE > 0 Then
     '
     If MODIFIC%(1) < 1 Then
       Exit Sub
     End If
     '
     If DERACCE%("ANOTAPRO", "Anotaciones de Proveedores") < 2 Then
        Exit Sub
     End If
     '
     If GRABATODO% <> 1 Then
       Command9.SetFocus
       If COMALTER%("Se han Modificado las Notas del Proveedor\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
         YACAR%(1) = 0
         Call CARNOTAPRO
         Exit Sub
       End If
       Command9.SetFocus
     End If
     '
     Call APERBASDAT("PROVE")
     Screen.MousePointer = 11
     TEXTO1$ = Text28.TEXT + Chr$(13) + Chr$(10)
     '
     Dim rstproved12 As ADODB.Recordset
     Set rstproved12 = New ADODB.Recordset
     rstproved12.CursorType = adOpenDynamic
     rstproved12.LockType = adLockPessimistic
     rstproved12.Open FILTSQL$("Select * from Proved12 where nume_cli=" & Abs(NCLIE)), FLEXCONN, , , adCmdText
     On Error Resume Next
     rstproved12!nota_Cli = TEXTO1$
     rstproved12.Update
     MODIFIC%(1) = 0
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Registrar")
     End If
     '
   End If
10   On Error GoTo 0
     Screen.MousePointer = 1
     '
End Sub
'
Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   '\\\OT-05-0271-WAR-18/05/05///
   Text4 = ""
   '\\\OT-05-0271-WAR-FIN///
   Text6 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text26.TEXT = ""
   Text27.TEXT = ""
   Text20.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   '
   For kk% = 0 To 7
     COTEXT(kk%).TEXT = ""
   Next kk%
   '
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   List2.Clear
   List3.Clear
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENPRO
   Call GRANOTAPRO
   Call CARDATPROV
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(2)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   '
   If APLINVO$ <> "" Then
    Hide
   Else
    On Error Resume Next
    Unload Me
    MARCH_GES07.Enabled = True
    MARCH_GES07.SetFocus
    On Error GoTo 0
   End If
   '
End Sub

Private Sub GruProv_Click()
OPMASPR07.Option1(9).Value = True
Call OPMASPR07.Command16_Click

End Sub


Private Sub Help_Click()
'Call Command4_Click
End Sub

Private Sub LiDePo_Click()
Call AMAPRLIS07.Command4_Click
End Sub

Private Sub LidePo1_Click()
Call AMAPRLIS07.Command4_Click
End Sub

Private Sub LiDiMa_Click()
Call AMAPRLIS07.Command2_Click
End Sub

Private Sub LiDiMa1_Click()
Call AMAPRLIS07.Command2_Click
End Sub

Private Sub LiGe_Click()
Call AMAPRLIS07.Command8_Click
End Sub

Private Sub LiInac_Click()
Call AMAPRLIS07.Command5_Click
End Sub

Private Sub LIST1_Click()
Call Command13_Click
End Sub

Private Sub List2_DblClick()
    ' CARGO EL LIST DEL FORMULARIO DEUDACLI07 CON LA MISMA INFORMACION DE ESTE LIST
      For U& = 1 To List2.ListCount
        DEUDAPRO07.LICHESEL.AddItem List2.List(U& - 1)
        VALOR2# = VALOR2# + XVALO(Mid$(List2.List(U& - 1), 65, 12))
      Next U&
      '
      DEUDAPRO07.Label15 = FORMATNUM$(VALOR2#, "F12.2")
      DEUDAPRO07.Show 1
      '
End Sub

Private Sub List3_DblClick()
   '
   RELIS$ = List3.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 11))
   If UCase$(TIDOCUM$) <> "O.PAGO" Then
      Call MUEFACO
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
   '
'5  For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   '
'   REDOCU$ = REGACT$("PDOCLST")
'   URE& = ULTREG&("PDOCSPL")
'   LI& = 0: LS& = URE&
'6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'   GoTo 8
'   '
'7  Screen.MousePointer = 1
'   Call MENSERR(24, "Documento no Registrado")
'   Exit Sub
'   '
'8  TPSP$ = ""
'   Screen.MousePointer = 11
'   While L& > 1
'     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'     If DCC& < DOCUNU& Then GoTo 9
'     L& = L& - 1
'   Wend
'   '
'9  CAPAGINAS% = 0
'   For KKL& = L& To URE&
'     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'     End If
'   Next KKL&
'   '
'95 SPOOLSTRING$ = TPSP$
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Cls
'     EDITFORM.Picture1(KKI%).Top = 0
'     EDITFORM.Picture1(KKI%).Refresh
'     EDITFORM.Picture1(KKI%).Height = 7170
'     EDITFORM.Picture1(KKI%).Width = 11700
'     EDITFORM.Picture1(KKI%).DrawWidth = 3
'     EDITFORM.Picture1(KKI%).DrawStyle = 0
'     EDITFORM.Picture1(KKI%).FillStyle = 1
'   Next KKI%
'   ALTUPAGINA = 7170
'   TOPEPAGINA = 0
'   PAGINACTIVA% = 0
'   CAPAGINAS% = 0
'   '
'   FIN& = Len(TPSP$)
'10 PPXX% = InStr(TPSP$, Chr$(255))
'   If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'   End If
'   '
'   Screen.MousePointer = 1
'   PAGINACTIVA% = 0
'   For KKI% = 0 To 5
'     EDITFORM.Picture1(KKI%).Visible = False
'     If KKI% = PAGINACTIVA% Then EDITFORM.Picture1(KKI%).Visible = True
'   Next KKI%
'   '
'   EDITFORM.Picture1(PAGINACTIVA%).Refresh
'   EDITFORM.VScroll1.Min = 0
'   EDITFORM.VScroll1.Value = 0
'   EDITFORM.VScroll1.Max = 1
'   If ALTUPAGINA > EDITFORM.Frame1.Height Then
'      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'   End If
'   EDITFORM.Command3.Enabled = False
'   'EDITFORM.Command5.Enabled = False
'   EDITFORM.Show 1
   '
End Sub

Private Sub ListGral_Click()
Call AMAPRLIS07.Command8_Click
End Sub

Private Sub LTEXT36_Change()
   DISPONI = XVALO(LTEXT36.Caption) - XVALO(LTEXT37.Caption) - XVALO(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub



Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE("AMAPRLIS07")
End Sub

Private Sub mnuCategoriaGananciasProveedor_Click()
    Call SetearCategoriaGananciasProveedor
End Sub

Private Sub mnuGuardarCambios_Click()
    Call Command1_Click
End Sub

Private Sub mnuPlanDeCuentas_Click()
    FPLANCUEN07.Show 1
End Sub

Private Sub mnuProveedoresParaIB_Click()
    Call OPMASPR07.Command4_Click
End Sub

Private Sub mnuReportesVarios_Click()
    Call AMAPRLIS07.AGenRep_Click
End Sub

Private Sub mnuRetencionGanancias_Click()
    Call OPMASPR07.Command5_Click
End Sub

Private Sub mnuRetencionIB_Click()
    Call OPMASPR07.Command3_Click
End Sub

Private Sub mnuRetencionIva_Click()
    Call OPMASPR07.Command6_Click
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     MODIFIC(0) = 1
     Frame5.Enabled = True
     Frame9.Enabled = True
     Frame10.Enabled = True
     '
     For KKI% = 0 To 2
       COTEXT(KKI%).Enabled = True
       DETEXT(KKI%).Enabled = True
     Next KKI%
     '
     Text19.Enabled = True
     Text25.Enabled = True
     Text26.Enabled = True
     Text27.Enabled = True
     Text29.Enabled = True
     Text30.Enabled = True
     Text31.Enabled = True
     Text32.Enabled = True
     Text33.Enabled = True
     Text34.Enabled = True
     Text36.Enabled = True
     '
     Command21.ToolTipText = "Alta de Nuevo Proveedor"
     '
  End If
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
     MODIFIC(0) = 1
     Frame5.Enabled = False
     Frame9.Enabled = False
     Frame10.Enabled = False
     '
     For KKI% = 0 To 2
       COTEXT(KKI%).Enabled = False
       DETEXT(KKI%).Enabled = False
     Next KKI%
     '
     Text19.Enabled = False
     Text25.Enabled = False
     Text26.Enabled = False
     Text27.Enabled = False
     Text29.Enabled = False
     Text30.Enabled = False
     Text31.Enabled = False
     Text32.Enabled = False
     Text33.Enabled = False
     Text34.Enabled = False
     Text36.Enabled = False
     '
     Command21.ToolTipText = "Alta de Nueva Cuenta de Gastos"
     '
   End If
End Sub

Private Sub print_Click()
   Call Command25_Click
End Sub

Private Sub ProZonVta_Click()
OPMASPR07.Option1(5).Value = True
Call OPMASPR07.Command16_Click

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      Call GRAGENPRO
      Call GRANOTAPRO
      Call FRESHALL
      Call CARDATPROV
   End If
End Sub

Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE = (-1) * XVALO(Text1.TEXT)
       If STATCLI(NCLIE) > 0 Then
         'NC% = NCLIE
         Call CARDATPROV
       End If
     End If
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub

Private Sub Text1_GotFocus()
   Call GRAGENPRO
   Call GRANOTAPRO
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
   If TRIM$(UCase$(Text20.TEXT)) = "SUSPENDIDO" Then
       LTEXT36.Caption = Text20.TEXT
     Else
       LTEXT36.Caption = CSTRING$(MKS$(XVALO(Text20.TEXT)), 4, 11, 2, 2)
   End If
End Sub

Private Sub Text20_LostFocus()
  If TRIM$(UCase$(Text20.TEXT)) <> "SUSPENDIDO" Then
    Text20.TEXT = CSTRING$(MKS$(XVALO(Text20.TEXT)), 3, 12, 2, 2)
  End If
End Sub

Private Sub Text21_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub TEXT27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub

Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub TEXT3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub
'\\\OT-06-0429-MH-11/10/06///
Private Sub text32_DblClick()
   If InStr(Text32.TEXT, "@") <> 0 Then
      FORMAI04.BackColor = &HC0D0F0
      FORMAI04.Picture1.BackColor = &H404080
      FORMAI04.Text28 = ""
      FORMAI04.Text9 = ""
      FORMAI04.Label3.Caption = Text32.TEXT
      FORMAI04.Show 1
   End If
End Sub
'\\\OT-06-0429-MH-FIN///

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub LText35_DblClick()
   Screen.MousePointer = 11
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NC% = XVALO(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + rasocli$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
Sub CARESCUENTA()
    '
    If YACAR%(2) > 0 Then Exit Sub
    '
50  Screen.MousePointer = 11
100 RECUEP07.Text1 = Text1
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
        List3.Clear
        List3.Refresh
        List2.Clear
        List2.Refresh
        LTEXT23.Caption = ""
        LTEXT37.Caption = ""
        LTEXT36.Caption = ""
        LTEXT35.Caption = ""
        Timer2.Enabled = True
    End If
    '
''100 RECUEP07.Text1 = Text1
'    With RECUEP07
'      For U& = 1 To .List1.ListCount
'        List3.AddItem .List1.List(U& - 1)
'      Next U&
'      '
'      For U& = 1 To .List2.ListCount
'        List2.AddItem .List2.List(U& - 1)
'      Next U&
'      '
'      Label25 = TRIM$(.Label6) 'agrego trim por que si no no se ve en el label
'      Label28 = TRIM$(.Label5) 'con borderStyle = 1
'      LTEXT37 = .Text10
'      LTEXT35 = .Text6
'    End With
'    '\\\OT-08-0886-OD-24/09/08-
'    On Error Resume Next
'    List3.TopIndex = List3.ListCount - 9
'    On Error GoTo 0
'    '\\\OT-08-0886-OD-FIN

    '
    YACAR%(2) = 1
    Screen.MousePointer = 1
    '
End Sub
'
Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
100 Screen.MousePointer = 11
    '
    SUMARRAS# = 0
    If List3.ListCount > 0 Then
      SUMARRAS# = XVALO(Mid$(List3.List(List3.ListCount - 1), 63, 14))
    End If
    SUMADEUP# = XVALO(Label28.Caption)
    CADOCPEN% = List2.ListCount
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(XVALO(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(XVALO(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(rasocli$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    ' para no enquilombar las cuentas que mas o menos estan
    If Abs(SUMARRAS# - SUMADEUP#) < 0.005 Then
      If NUDOCU% >= CADOCPEN% Then
        GoTo 999
      End If
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                Tipo% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             Tipo% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If Tipo% < 1 Or Tipo% > 3 Then COPAGI% = 0
             SG% = 1: If Tipo% >= 4 Then If Tipo% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             Tipo% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          ARDEX$ = ARCHIDEU$ + "/" + TRIM$(CSTRING$(MKD$(IORIG#), 4, 12, 2, 2))
          Call GRADEU(LU$, ARDEX$, RDX%, Tipo%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub

Sub GENECOPROV()
   If UPDATPRO% > 0 Then
      Call ABREPROVE
      Call CARGALISEL("PROVED1", "Proved12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 AND NUME_CLI>0 ORDER BY Raso_Cli")
      Call CARGALISEL("ACREDOR", "Proved12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 AND NUME_CLI>0 AND Tcpr_Cli = 0 ORDER BY Raso_Cli")
      Call CARGALISEL("INDIPROV", "Proved12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND CODI_CLI<>'' ORDER BY Raso_Cli")
      Call CARGALISEL("INDIACRE", "Proved12", "Codi_Cli/a12<;Raso_Cli/A50", "Stat_Cli >= 0 AND CODI_CLI<>'' AND Tcpr_Cli = 0 ORDER BY Raso_Cli")
      UPDATPRO% = 0
      Screen.MousePointer = 1
   End If
End Sub

Sub GENPROVANU()
   '
   Call APERBASDAT("PROVE")
   ' Busco los que estan dados de baja
   Dim ABRO1TMP As ADODB.Recordset
   Set ABRO1TMP = New ADODB.Recordset
     
   SQLX$ = "SELECT * FROM PROVED12 WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE Stat_Cli = -1 "
   Set ABRO1TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With ABRO1TMP
     jj& = 0
     Do While .EOF = False
        EL2$ = REBLB$
        Call REPLA(EL2$, !Codi_Cli, 1, 12)
        Call REPLA(EL2$, !raso_cli, 13, 50)
        Call REPLA(EL2$, MKI$(!nume_cli), 63, 2)
        jj& = jj& + 1
        Call GRAREG("INDIPROA", EL2$, jj&)
     .MoveNext
     Loop
   End With
   '
80 ABRO1TMP.Close
   Set ABRO1TMP = Nothing
   '
   Call SETULTREG("INDIPROA", jj&)
   Call FILESORT("INDIPROA", "", 1, 1, "ASC")
   Call FRESHECHO("INDIPROA")
   '
End Sub
'
Private Sub Text6_Change()
   If NUMEPRO%(Text6) > 0 Then
     Text1 = TRIM$(Str$(NUMEPRO%(Text6)))
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command10_Click
End Sub

Sub GENDAREGAPRO()
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("PROVED1")
    REBLA$ = REGBLAN$("DARGAPRO")
    Call REPLA(REBLA$, USERNAME$, 36, 20)
    Call REPLA(REBLA$, MKI$(HOY(HOS$)), 56, 2)
    Call REPLA(REBLA$, MKI$(Int(HORANUM(Time$))), 58, 2)
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PROVED1", U&)
      NCLIX% = CVI(Left$(XX$, 2))
      If NCLIX% > 0 Then
        RECORD$ = REBLA$
        Call REPLA(RECORD$, MKI$(NCLIX%), 1, 2)
        Call REPLA(RECORD$, Chr$(2), 3, 1)
        Call FILTERPUTRECORD("DARGAPRO", 1, MKI$(NCLIX%), RECORD$)
      End If
    Next U&
    Call CIERRARCH("DARGAPRO")
    Screen.MousePointer = 1
    '
End Sub

Sub MUEFACO()
   '
   TIPODOCU$ = TRIM$(Mid$(List3.TEXT, 10, 12))
   NUDOCOR$ = TRIM$(Mid$(List3.TEXT, 22, 18)) 'NUMERO DE DOC COR
   NP1& = XVALO(Text1)
   FECHDOCU$ = TRIM$(Left$(List3.TEXT, 9))
   'FECDOC1 = CDate(TRIM$(Mid$(List3.TEXT, 1, 9))) 'FECHA DEL DOC.
   FECDOC1 = FETEXCOR1$(FECHANUM(TRIM$(Mid$(List3.TEXT, 1, 9)))) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLX$ = SQLX$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLX$ = SQLX$ + " AND NuProve =  " & NP1& & " "
   'Set SHOFACOTMP = CueCorri.OpenRecordset(SQLX$, 4)
   Dim SHOFACOTMP As ADODB.Recordset
   Set SHOFACOTMP = New ADODB.Recordset
   SHOFACOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = TIPODOCU$ + " " + !codicom_lar + " del " + FECHDOCU$
     SHOWFACO.List1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
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
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or isEmpty(TICA) = True) Then TICA = 1
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
          SHOWFACO.Label14 = !ObserGen
        End If
        '
     .MoveNext
     Loop
   End With
   Set SHOFACOTMP = Nothing
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
   '
End Sub


Private Sub Timer1_Timer()

   Timer1.Enabled = False
    
    Dim oAltaCliente As New DatosEmpresaAfip
    R = oAltaCliente.ObtenerCredencial
    'r = oAltaCliente.ObtenerDatosEmpresaAfip("30715781634", response$)
    
End Sub

Private Sub Timer2_Timer()
    Static TextA$, FUEAMOSTRAR%
    '
    If Text1 = TextA$ Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '
    FUEAMOSTRAR% = 1
    Call CARESCUEPRO
    FUEAMOSTRAR% = 0
    '
    Timer2.Enabled = False
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
    '
100 List3.Clear
    List3.Refresh
    List2.Clear
    List2.Refresh
    SALDO# = 0: SARRAS# = 0
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
      List3.AddItem Z$
    Next KU&
    DoEvents
    '
    Call APERBASDAT("CABAN")
    'jandy sql
    
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NC% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP') "
    'SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, VAL(MID(CodiCom_Lar,11)) ASC;"
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC, SUBSTRING(CodiCom_Lar,11,8)ASC;"
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
          TIX$ = Left$(SAFETEXT$(!codicom_lar), 2)
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
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
          If List3.ListCount > 0 Then
            If TIX$ = "OP" Or TIX$ = "FF" Then
              If SAFETEXT$(!codicom_lar) = NROX$ Then
                If CVINT%(!FECHEMISI) = FEX Then
                  SALDO# = SALDO# - IMPO#
                  IMPO# = IMPO# + XVALO(Mid$(Z$, 40, 11)) - XVALO(Mid$(Z$, 53, 11))
                  List3.RemoveItem List3.ListCount - 1
                  IMPTEX$ = FORMATNUM$(Abs(IMPO#), "F11.2")
                  SALTEX$ = FORMATNUM$(SALDO#, "F13.2")
                  Call REPLA(Z$, Space$(24), 40, 24)
                  If IMPO# > 0 Then
                      Call REPLA(Z$, IMPTEX$, 40, 11)
                    ElseIf IMPO# < 0 Then
                      Call REPLA(Z$, IMPTEX$, 53, 11)
                  End If
                  Call REPLA(Z$, SALTEX$, 64, 13)
                  List3.AddItem Z$
                  GoTo 899
                End If
              End If
            End If
          End If
          ' FIN AGREGADO 27/05/06
          '
890       SALDO# = SALDO# - IMPO#
          FEX = CVINT%(!FECHEMISI)
          NROX$ = SAFETEXT$(!codicom_lar)
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
          List3.AddItem Z$
          '
899      .MoveNext
        Loop
      End If
990 End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    List3.Refresh
    Openforms = DoEvents
    '
    'Call REMPLAGRID(Me, GRID2, Picture4, Label9, List1)
    '
991 SARRAS# = SALDO#
    SUMADEU# = 0
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
          TIX$ = UCase$(Left$(SAFETEXT$(!codicom_lar), 2))
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Aj.Saldo"
                              If TIX$ = "TF" Then TIPOS$ = "Tkt-Factura"
                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honor."
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
          SALDO# = XVALO(!SALDAcre) - XVALO(!SaldDebe)
          SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = SAFETEXT$(!codicom_lar)
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
          '
          Z$ = Space$(80)
          Call REPLA(Z$, TIPOS$, 1, 11)
          Call REPLA(Z$, NROX$, 12, 18)
          Call REPLA(Z$, FECHATEX$(FEMI), 32, 8)
          Call REPLA(Z$, IMPTEX$, 40, 12)
          Call REPLA(Z$, FECHATEX$(FEVE), 56, 8)
          Call REPLA(Z$, SALTEX$, 65, 12)
          '
          List2.AddItem Z$
          '
993     .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
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
         End If
       End If
    End If
    '
995 'List2.ListIndex = List2.ListCount - 1
    List2.Refresh
    Label25.Caption = TRIM$(CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2))
    Label25.Refresh
    Label28.Caption = TRIM$(CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2))
    Label28.Refresh
    '
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
