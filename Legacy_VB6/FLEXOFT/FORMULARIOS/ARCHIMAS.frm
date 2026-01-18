VERSION 5.00
Begin VB.Form ARCHIMAS 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivos Maestros - Actualización y Consulta"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9990
   Icon            =   "ARCHIMAS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command222 
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
      Left            =   5230
      Picture         =   "ARCHIMAS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Exportación de Datos"
      Top             =   220
      Visible         =   0   'False
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6090
      Picture         =   "ARCHIMAS.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Accesos Directos"
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Actualización de Bases de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   210
      TabIndex        =   3
      Top             =   525
      Width           =   4635
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Materiales, Fórmulas y Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   315
         TabIndex        =   12
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   22
            Top             =   525
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "ARCHIMAS.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestro de Artí- culos de Stock"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   24
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   5
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   11
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   10
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   9
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   8
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   6
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Clientes / Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   315
         TabIndex        =   4
         Top             =   1890
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   28
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command18 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "ARCHIMAS.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes en Cuenta Corriente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   30
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   25
            Top             =   420
            Width           =   2325
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
               Height          =   600
               Left            =   -20
               Picture         =   "ARCHIMAS.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y Subcontratistas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   27
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5145
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   47
      Top             =   945
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8925
      Picture         =   "ARCHIMAS.frx":11B4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7980
      Picture         =   "ARCHIMAS.frx":12FE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7035
      Picture         =   "ARCHIMAS.frx":1608
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   5145
      TabIndex        =   13
      Top             =   1260
      Width           =   4635
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Clientes y Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   210
         TabIndex        =   40
         Top             =   2835
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Clasificacion Contable Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   9
            Left            =   2310
            TabIndex        =   46
            Top             =   1050
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Plan de Cuen- tas Contables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   8
            Left            =   2310
            TabIndex        =   45
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Listas de Precios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   7
            Left            =   315
            TabIndex        =   44
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Condiciones de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   6
            Left            =   315
            TabIndex        =   43
            Top             =   630
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Provincias y Zonas de Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   5
            Left            =   315
            TabIndex        =   42
            Top             =   1050
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Vendedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   2310
            TabIndex        =   41
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   3570
         Picture         =   "ARCHIMAS.frx":1A4A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3225
      End
      Begin VB.CommandButton Command6 
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
         Height          =   360
         Left            =   3255
         TabIndex        =   32
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   31
         Text            =   " "
         Top             =   525
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   15
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command20 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   21
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   20
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   19
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   18
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   17
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Productos, Fórmulas y Rutas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   14
         Top             =   1575
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Grupos Contables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2310
            TabIndex        =   39
            Top             =   735
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Depósitos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   2310
            TabIndex        =   38
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Unidades de Medida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   37
            Top             =   735
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Tipos de Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   315
            TabIndex        =   36
            Top             =   315
            Width           =   1800
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -735
         TabIndex        =   35
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Sub-Sistema de Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   49
      Top             =   4515
      Width           =   4635
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   3900
         TabIndex        =   50
         Top             =   500
         Width           =   3960
         Begin VB.CommandButton Command8 
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
            Height          =   600
            Left            =   560
            Picture         =   "ARCHIMAS.frx":1D54
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "B.D.Instrumentos de Medición y Ensayo"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command21 
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
            Height          =   600
            Left            =   3350
            Picture         =   "ARCHIMAS.frx":205E
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Informes de No-Conformidad"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command22 
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
            Height          =   600
            Left            =   0
            Picture         =   "ARCHIMAS.frx":2368
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "B.D.Documentos de Ingenieria"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Sistema de Asegu- ramiento de la Calidad"
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
            Height          =   540
            Left            =   1260
            TabIndex        =   52
            Top             =   105
            Width           =   2010
         End
      End
   End
End
Attribute VB_Name = "ARCHIMAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
'

Private Sub Command1_Click()
    Command1.Enabled = False
    If EXISTE%("AMASTO02.EXE") > 0 Then
         Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
       ElseIf EXISTE%("AMASTO01.EXE") > 0 Then
         Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
       Else
         Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    End If
    Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If EXISTE%("AMACLI02.EXE") > 0 Then
         Call CONECRUN("AMACLI02", "Padron Maestro de Clientes")
       ElseIf EXISTE%("AMACLI01.EXE") > 0 Then
         Call CONECRUN("AMACLI01", "Padron Maestro de Clientes")
       Else
         Call CONECRUN("AMACLI", "Maestro de Clientes")
    End If
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    If EXISTE%("AMASTO01.EXE") > 0 Then
        Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
        GoTo 99
      Else
        If EXISTE%("FORMULAS.EXE") > 0 Then
           Call CONECRUN("FORMULAS", "Formulas y Estructuras de Producto")
           GoTo 99
        End If
    End If
    Call OPNOIN("Formulas y Estructuras de Producto")
    '
99  Command2.Enabled = True
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   ACONEC$ = "INOCON02"
   If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "INOCON01"
   Call CONECRUN(ACONEC$, "Informes de No-Conformidad")
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   ACONEC$ = "DOCING01"
   If EXISTE%("ISO9K02.EXE") > 0 Then
      ACONEC$ = "ISO9K02/DOCUMING"
   End If
   Call CONECRUN(ACONEC$, "Documentos de Ingenieria")
   Command22.Enabled = True
End Sub
'\\\OT-06-0033-WAR-02/02/06///
Private Sub Command222_Click()
  Command222.Enabled = False
  RTA1% = ALTERNA%("Clientes\Proveedores\Listas de Precios")
  If RTA1% = 1 Then
      Call IMPBDCLI
    ElseIf RTA1% = 2 Then
      Call IMPBDPRO
    ElseIf RTA1% = 3 Then
      Call IMPBDLIPR
  End If
  Command222.Enabled = True
End Sub
Sub IMPBDLIPR()
'
  Call BLANARCH("BDLPXLS")
  '
  FIN& = ULTREG&("LISTAS")
  For J& = 1 To ULTREG&("LISTAS")
    '
    Call TRACE(24, J&, FIN&)
    '
    X$ = REGLEIDO$("LISTAS", J&)
    NLP1$ = Mid$(X$, 2, 46)
    If TRIM$(NLP1$) = "" Then GoTo 25
    MO1% = Asc(Mid$(X$, 50, 1))
    If MO1% = 1 Then
      MON1$ = "PESOS"
      ElseIf MO1% = 2 Then
        MON1$ = "DOLARES"
    End If
    '
    NL$ = TRIM$(Str$(J&))
    Do Until Len(NL$) = 3
      If Len(NL$) < 3 Then
        NL$ = "0" + NL$
      End If
      JJ& = JJ& + 1
    Loop
    A1$ = "LIPRE" + NL$
    '
    For JJJ& = 1 To ULTREG&(A1$)
    Z$ = REGLEIDO$(A1$, JJJ&)
    CI% = CVI(Mid$(Z$, 1, 2))
    J1$ = CODEXT$(CI%)
    '
    If InStr(1, J1$, "`") > 0 Then
      If Right$(J1$, 1) = "`" Then
        GoTo 20
      End If
    End If
    '
    'If CODINT%(J1$) <= 0 Then GoTo 20
    PU1# = CVD(Mid$(Z$, 62, 8))
    '
    Y$ = REGBLAN$("BDLPXLS")
    Call REPLA(Y$, Chr$(J&), 1, 1)
    Call REPLA(Y$, MKI$(CI%), 2, 2)
    Call REPLA(Y$, MKD$(PU1#), 4, 8)
    Call REPLA(Y$, MON1$, 12, 8)
    Call REGAPP("BDLPXLS", Y$)
    '
20  Next JJJ&
    
    '
25 Next J&
  '
  Call CIERRARCH("BDLPXLS")
  Call FINAL("*BDLPXLS")
  Call CIERRARCH("*.*")
  '
End Sub
Sub IMPBDPRO()
  Call BLANARCH("BDPROXLS")
  FIN& = ULTREG&("PROVED1")
  For J& = 1 To ULTREG("PROVED1")
    '
    Call TRACE(24, J&, FIN&)
    '
    X$ = REGLEIDO$("PROVED1", J&)
    NC% = CVI(Mid$(X$, 1, 2))
    NCP% = NC% * (-1)
    '
    PAIS1$ = PAISCLI$(NCP%)
    RAS1$ = RASOCLI$(NCP%)
    '
    If TRIM$(RAS1$) = "" Then GoTo 25
    'SI NO TIENE RAZON SOCIAL NO VA
    '
    LOC1$ = LOCACLI$(NCP%)
    CP1$ = CPOSCLI$(NCP%)
    PROV1$ = PROVCLI$(NCP%)
    FANT1$ = TEL2CLI$(NCP%)
    DIR1$ = DOMICLI$(NCP%)
    TEL1$ = TEL1CLI$(NCP%)
    FAX1$ = FAXCLI$(NCP%)
    CUI1$ = CUITCLI$(NCP%)
    LP1% = LIPRCLI%(NCP%)
    MON1$ = "PESOS"
    CPA1% = CPAGCLI%(NCP%)
'    TC1% = TICUEPRO%(NCP%)
'    'SI TIPO CUENTA ES GASTO
'    If TC1% > 0 Then GoTo 25
    '
    Y$ = REGBLAN$("BDPROXLS")
    SPA1$ = Space$(16)
    SPA2$ = Space$(30)
    SPA3$ = Str$(NC%)
    '
    Call REPLA(Y$, PAIS1$, 1, 32)
    Call REPLA(Y$, RAS1$, 33, 64)
    Call REPLA(Y$, SPA2$, 97, 30) 'NOMBRE OPCIONAL
    Call REPLA(Y$, SPA3$, 127, 30) 'NOMBRE OPCIONAL
    Call REPLA(Y$, SPA2$, 157, 30) 'NOMBRE OPCIONAL
    Call REPLA(Y$, LOC1$, 187, 48)
    Call REPLA(Y$, SPA1$, 235, 16) 'ZONA
    Call REPLA(Y$, SPA1$, 251, 16) 'CODIGO POSTAL ESPECIAL
    Call REPLA(Y$, CP1$, 267, 16)
    Call REPLA(Y$, PROV1$, 283, 1)
    Call REPLA(Y$, FANT1$, 284, 32)
    Call REPLA(Y$, DIR1$, 316, 64)
    Call REPLA(Y$, SPA1$, 380, 16) 'TRATAMIENTO
    Call REPLA(Y$, SPA1$, 396, 16) 'IDIOMA
    Call REPLA(Y$, CUI1$, 412, 15)
    Call REPLA(Y$, Chr$(CPA1%), 427, 1)
    Call REPLA(Y$, SPA1$, 428, 16) 'BANCO
    Call REGAPP("BDPROXLS", Y$)
    '
25 Next J&
  '
  Call CIERRARCH("BDPROXLS")
  Call FINAL("*BDPROXLS")
  '
End Sub
Sub IMPBDCLI()
  '
  Call BLANARCH("BDCLIXLS")
  FIN& = ULTREG&("DEUDOR1")
  For J& = 1 To ULTREG("DEUDOR1")
    '
    Call TRACE(24, J&, FIN&)
    '
    X$ = REGLEIDO$("DEUDOR1", J&)
    NC% = CVI(Mid$(X$, 1, 2))
    '
    PAIS1$ = PAISCLI$(NC%)
    RAS1$ = RASOCLI$(NC%)
    LOC1$ = LOCACLI$(NC%)
    CP1$ = CPOSCLI$(NC%)
    PROV1$ = PROVCLI$(NC%)
    FANT1$ = TEL2CLI$(NC%)
    DIR1$ = DOMICLI$(NC%)
    TEL1$ = TEL1CLI$(NC%)
    FAX1$ = FAXCLI$(NC%)
    CUI1$ = CUITCLI$(NC%)
    GC1% = GRUCOCLI%(NC%)
    LP1% = LIPRCLI%(NC%)
    MON1$ = "PESOS"
    CM1$ = CUEMADRE$(NC%)
    CPA1% = CPAGCLI%(NC%)
    VE1% = VENDCLI%(NC%)
    '
    Y$ = REGBLAN$("BDCLIXLS")
    SPA1$ = Space$(16)
    Call REPLA(Y$, PAIS1$, 1, 32)
    Call REPLA(Y$, RAS1$, 33, 64)
    Call REPLA(Y$, SPA1$, 97, 16) 'NOMBRE OPCIONAL
    Call REPLA(Y$, LOC1$, 113, 48)
    Call REPLA(Y$, CP1$, 161, 16)
    Call REPLA(Y$, PROV1$, 177, 1)
    Call REPLA(Y$, FANT1$, 178, 32)
    Call REPLA(Y$, DIR1$, 210, 64)
    Call REPLA(Y$, TEL1$, 274, 24)
    Call REPLA(Y$, FAX1$, 298, 24)
    Call REPLA(Y$, SPA1$, 322, 16) 'TRATAMIENTO
    Call REPLA(Y$, SPA1$, 338, 16) 'IDIOMA
    Call REPLA(Y$, CUI1$, 354, 15)
    Call REPLA(Y$, MKI$(NC%), 369, 2)
    Call REPLA(Y$, MKI$(GC1%), 371, 2) 'MERCADO
    Call REPLA(Y$, SPA1$, 373, 16) 'SECTOR-ACTIVIDAD
    Call REPLA(Y$, SPA1$, 389, 16) 'SEGUIMIENTO DEL CLIENTE
    Call REPLA(Y$, MKI$(GC1%), 405, 2) 'GRUPO DE CLIENTES
    Call REPLA(Y$, SPA1$, 407, 16) 'REGION DE VENTA
    Call REPLA(Y$, Chr$(LP1%), 423, 1)
    Call REPLA(Y$, SPA1$, 424, 16) 'FOB
    Call REPLA(Y$, SPA1$, 440, 16) 'ICOTERMOS 2
    Call REPLA(Y$, SPA1$, 456, 16) 'MAX. REMS. PARC.
    Call REPLA(Y$, SPA1$, 472, 16) 'AGRUPAMENTO    X
    Call REPLA(Y$, SPA1$, 488, 16) 'PRIORIDAD DE ENTREGA
    Call REPLA(Y$, MON1$, 504, 5)
    Call REPLA(Y$, CM1$, 509, 16)
    Call REPLA(Y$, Chr$(CPA1%), 525, 1)
    Call REPLA(Y$, SPA1$, 526, 16) '0011
    Call REPLA(Y$, Chr$(VE1%), 542, 1)
    Call REPLA(Y$, SPA1$, 543, 16)
    Call REGAPP("BDCLIXLS", Y$)
    '
  Next J&
  '
  Call CIERRARCH("BDCLIXLS")
  Call FINAL("*BDCLIXLS")
  '
End Sub
 '\\\OT-06-0033-WAR-FIN///
Private Sub Command23_Click()
    Call OPNOIN("")
End Sub

Private Sub Command24_Click()
    Call FINAL("")
End Sub

Private Sub Command25_Click()
Call BLOQARCH("MASTER")
End Sub

Private Sub Command26_Click()
Call BLOQARCH("INVERT")
End Sub

Private Sub Command27_Click()
Call BLOQARCH("INVENUE")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    Command3.Enabled = False
    Call ACCDIR("ARCHIMA")
    Command3.Enabled = True
    '
99  Command3.Enabled = True
End Sub


Private Sub Command7_Click()
    Command7.Enabled = False
    If MODUCONE = "TADEPOSI" Then
         Call CARTADEPOSI
       ElseIf MODUCONE = "TABTIMA" Then
         Call CARTATIMA
       ElseIf MODUCONE = "CUEINGAS" Then
         Call CARCUEINGAS
       ElseIf MODUCONE = "CONPAGO" Then
         Call CARCONPAGOS
       ElseIf MODUCONE = "TAPROVIN" Then
         Call CARTAPROVIN
       ElseIf MODUCONE = "UMEDIDA" Then
         Call CARUMEDIDA
       Else
         If TRIM$(MODUCONE$) <> "" Then
            Call CONECRUN(MODUCONE$, Text2.Text)
         End If
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   ACONEC$ = "INSTRU01"
   If EXISTE%("ISO9K02.EXE") > 0 Then
      ACONEC$ = "ISO9K02/INSTRUME"
   End If
   Call CONECRUN(ACONEC$, "B.D.Instrumentos de medición y Ensayo")
   Command8.Enabled = True

End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '\\\OT-06-0033-WAR-02/02/06///
      If InStr(1, EMPREAC$, "SABO") > 0 Then
        Command222.Visible = True
      End If
    ''\\\OT-06-0033-WAR-FIN///
    If EXISTE%("ISO9K02.EXE") > 0 Then
      Label7.ToolTipText = "Doble Click para Acceso Directo a Sub-Sistema de Gestión de Calidad ISO-9K"
    End If
    '
    If EXISTE%("ISO9K02.EXE") > 0 Or EXISTE%("DOCING01.EXE") > 0 Then
      Command22.Enabled = True
      Label7.Enabled = True
    End If
    '
    If EXISTE%("ISO9K02.EXE") > 0 Or EXISTE%("INSTRU01.EXE") > 0 Then
      Command8.Enabled = True
      Label7.Enabled = True
    End If
    '
    If EXISTE%("INOCON01.EXE") > 0 Or EXISTE%("INOCON02.EXE") > 0 Then
      Command21.Enabled = True
      Label7.Enabled = True
    End If
    '
    TABLA$(0) = "TABTIMA"       ' tabla de tipos de material
    TABLA$(1) = "UMEDIDA"
    TABLA$(2) = "TADEPOSI"
    TABLA$(3) = ""       ' grupos contables por articulo
    TABLA$(4) = "*VENDEDOR"
    TABLA$(5) = "TAPROVIN"
    TABLA$(6) = "CONPAGO"
    TABLA$(7) = "ADMIPRE"
       If EXISTE%("LISPRE02.EXE") > 0 Then TABLA$(7) = "LISPRE02"
    TABLA$(8) = "CUEINGAS"
    TABLA$(9) = "*CUECOVE"
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label7_DblClick()
   If EXISTE("ISO9K02.EXE") > 0 Then
     Call CONECRUN("ISO9K02", "Sistema de Gestión de la Calidad")
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.Text = REPLACAR$(Option1(Index).Caption, "- ", "")
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.Text = TETABLA$
    '
    For K% = 0 To 9
       If K% <> Index Then
          Option1(K%).Value = False
       End If
    Next K%
    Command7.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TABTIMA" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "UMEDIDA" Then
            Call CARUMEDIDA
          Else
            Call CONECRUN(MODUCONE$, Text2.Text)
        End If
    End If
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         X$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(X$, 4)) <> "" Then
           If Len(TRIM$(Left$(X$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           Y$ = Left$(X$, 2) + Mid$(X$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", X$, JJ&)
           Call GRAREG("UMEDIDA", Y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub


