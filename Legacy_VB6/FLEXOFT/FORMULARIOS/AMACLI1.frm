VERSION 5.00
Begin VB.Form AMACLI1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivo Maestro de Clientes"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10170
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Importar"
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
      Left            =   6405
      TabIndex        =   57
      Top             =   315
      Width           =   1485
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6405
      ScaleHeight     =   75
      ScaleWidth      =   1380
      TabIndex        =   52
      Top             =   1050
      Width           =   1440
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Default"
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
      Left            =   5250
      Picture         =   "AMACLI1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Detalle de Movimientos por Cuenta"
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
      Left            =   5250
      TabIndex        =   47
      Top             =   5145
      Width           =   4635
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   48
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command9 
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
            Picture         =   "AMACLI1.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Resúmenes de Cuenta Corriente"
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
            TabIndex        =   50
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
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
      Height          =   3480
      Left            =   5250
      TabIndex        =   6
      Top             =   1470
      Width           =   4635
      Begin VB.CommandButton Command20 
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
         TabIndex        =   16
         Top             =   525
         Width           =   175
      End
      Begin VB.Frame Frame11 
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
         TabIndex        =   18
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command26 
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
            TabIndex        =   24
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
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
            TabIndex        =   23
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
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
            TabIndex        =   22
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
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
            TabIndex        =   21
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
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
         Begin VB.CommandButton Command21 
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
            TabIndex        =   19
            Top             =   1890
            Width           =   1590
         End
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
         TabIndex        =   17
         Text            =   " "
         Top             =   525
         Width           =   1695
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
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1050
         Width           =   3225
      End
      Begin VB.CommandButton Command16 
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
         Picture         =   "AMACLI1.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   525
         Width           =   855
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Opciones de Trabajo"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   13
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   12
            Top             =   1050
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   11
            Top             =   630
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
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
            TabIndex        =   10
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Cuentas Contables"
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
            TabIndex        =   9
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Clientes"
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
            TabIndex        =   8
            Top             =   1050
            Width           =   1695
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
         TabIndex        =   25
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Actualizacion Base de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   420
      TabIndex        =   2
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Consultas - Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Left            =   315
         TabIndex        =   5
         Top             =   3255
         Width           =   4005
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   54
            Top             =   1260
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
               Height          =   600
               Left            =   -20
               Picture         =   "AMACLI1.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Deudas y Límites de Crédito"
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
               Left            =   630
               TabIndex        =   56
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Salida"
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
            Left            =   420
            TabIndex        =   44
            Top             =   2050
            Width           =   3165
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Pantalla"
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
               Left            =   315
               TabIndex        =   46
               Top             =   250
               Width           =   1380
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Impresión"
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
               Left            =   1785
               TabIndex        =   45
               Top             =   250
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   38
            Top             =   840
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
               Height          =   600
               Left            =   -20
               Picture         =   "AMACLI1.frx":0C28
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Cuentas Inactivas"
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
               Left            =   630
               TabIndex        =   40
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   41
            Top             =   420
            Width           =   2325
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
               Height          =   600
               Left            =   -20
               Picture         =   "AMACLI1.frx":0F32
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado General   y Clasificado"
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
               TabIndex        =   43
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Archivo de Cuentas de Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   315
         TabIndex        =   3
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1365
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   35
            Top             =   1785
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "AMACLI1.frx":123C
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Apertura de Cuenta Corriente"
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
               TabIndex        =   37
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   26
            Top             =   1365
            Width           =   2325
            Begin VB.CommandButton Command12 
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
               Picture         =   "AMACLI1.frx":1546
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificación de Datos de Cuenta"
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
               TabIndex        =   28
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   29
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command28 
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
               Picture         =   "AMACLI1.frx":1850
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Suspensión de Operaciones"
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
               TabIndex        =   31
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   32
            Top             =   525
            Width           =   2325
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
               Height          =   600
               Left            =   -20
               Picture         =   "AMACLI1.frx":1B5A
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Baja Definitiva"
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
               TabIndex        =   34
               Top             =   105
               Width           =   1275
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9030
      Picture         =   "AMACLI1.frx":1E64
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command17 
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
      Left            =   8085
      Picture         =   "AMACLI1.frx":1FAE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2940
      TabIndex        =   4
      Top             =   6825
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "AMACLI1.frx":22B8
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   10920
   End
End
Attribute VB_Name = "AMACLI1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
'

Private Sub Command1_Click()
    Unload AMACLI1
End Sub

Private Sub Command12_Click()
    MODO% = 2
    MODICLI.Show
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.DAT")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then GoTo 99
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    Call HELPONLINE("AMACLI")
End Sub

Private Sub Command2_Click()
    MODO% = 1
    Screen.MousePointer = 11
    Call CLEARFORMCLI
    NROPRO& = 0
    For U& = 1 To ULTREG&("DEUDOR1")
        X$ = REGLEIDO$("DEUDOR1", U&)
        If CVI(Left$(X$, 2)) < 9998 Then
          If CVI(Left$(X$, 2)) > NROPRO& Then
            NROPRO& = CVI(Left$(X$, 2))
          End If
        End If
    Next U&
    MODICLI.Text1.TEXT = NROPRO& + 1
    Screen.MousePointer = 1
    MODICLI.Show
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call LILIMCRE
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call IMPORCLI
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call SELECHO("SELFECHA")
    FECHLI$ = VALACT1$("SELFECHA")
    FLIM% = FECHANUM(FECHLI$)
    If FLIM% > 0 Then
       Screen.MousePointer = 11
       Dim ACTI%(32767)
       FIN& = ULTREG&("FACTUR")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         X$ = REGLEIDO$("FACTUR", U&)
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If NC% <= 32767 Then
             If CVI(Mid$(X$, 7, 2)) > ACTI%(NC%) Then
               ACTI%(NC%) = CVI(Mid$(X$, 7, 2))
             End If
           End If
         End If
       Next U&
       '
       RELISA& = 0
       CX& = ULTREG&("DEUDOR1")
       For IL& = 1 To CX&
         Call TRACE(20, IL&, CX&)
         NC% = CVI(Left$(REGLEIDO$("DEUDOR1", IL&), 2))
         If ACTI%(NC%) < FLIM% Then
            Y$ = REGBLAN$("LISACLI")
            SSCC# = SALDCLI#(NC%)
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, MKD$(SSCC#), 3, 8)
            Call REPLA(Y$, MKD$(SSCC#), 11, 8)
            FEX = ACTI%(NC%)
            TTXX$ = "Inactiva Desde " + FECHATEX$(FEX)
            Call REPLA(Y$, TTXX$, 19, 30)
            '
            RELISA& = RELISA& + 1
            Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       Next IL&
       '
       Call SETULTREG("LISACLI", RELISA&)
       Call CIERRARCH("*.*")
       Screen.MousePointer = 1
       Call CONECRUN("*DEUINAC", "Listado de Cuentas Inactivas")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call OPNOIN("")
End Sub



Private Sub Command8_Click()
    If Option1(0).Value = True Then
        Option2.Value = True
    End If
    Command8.Enabled = False
    FORLICLIEN.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    If EXISTE%("COGEVEN.EXE") < 1 Then
        Call OPNOIN("")
        Exit Sub
    End If
    Hide
    RESCUDEU.Show
End Sub

Private Sub Form_Load()
    '
    COMMAX$ = TRIM$(Command$)
    Call VERJOBID
    '
    NCLILLAMA% = 0
    COMMA$ = COMMAX$
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
       COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    '
    If ULTREG&("DEUDOR1") > 0 Then
       Command4.Visible = False
    End If
    '
    If COMMA$ <> "" Then
       NCLIEN% = Val(COMMA$)
       If NCLIEN% > 0 Then
          If TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIEN%))) <> "" Then
             Call CARTAVALI
             '
             TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
             TABLA$(1) = "*UNIMED"
             TABLA$(2) = "TADEPOSI"
             TABLA$(3) = ""       ' grupos contables por articulo
             TABLA$(4) = "*VENDEDOR"
             TABLA$(5) = "TAPROVIN"
             TABLA$(6) = "CONPAGO"
             TABLA$(7) = "ADMIPRE"
             TABLA$(8) = "CUEINGAS"
             TABLA$(9) = "GRUCOCLI"
             '
             MODO% = 2
             MODICLI.Text1.TEXT = NCLIEN%
             Screen.MousePointer = 1
             NCLILLAMA% = NCLIEN%
             Call CARDATCLI(NCLIEN%)
             MODICLI.Show
             Exit Sub
          End If
       End If
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Show
    Refresh
    '
    Call CARTAVALI
    '
    TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
    TABLA$(1) = "*UNIMED"
    TABLA$(2) = "TADEPOSI"
    TABLA$(3) = ""       ' grupos contables por articulo
    TABLA$(4) = "*VENDEDOR"
    TABLA$(5) = "TAPROVIN"
    TABLA$(6) = "CONPAGO"
    TABLA$(7) = "ADMIPRE"
    TABLA$(8) = "CUEINGAS"
    TABLA$(9) = "GRUCOCLI"
    '
    If Left$(COMMAX$, 3) = "NEW" Then
        MODO% = 1
        Screen.MousePointer = 11
        Call CLEARFORMCLI
        NROPRO& = 0
        For U& = 1 To ULTREG&("DEUDOR1")
            X$ = REGLEIDO$("DEUDOR1", U&)
            If CVI(Left$(X$, 2)) < 9998 Then
                If CVI(Left$(X$, 2)) > NROPRO& Then
                    NROPRO& = CVI(Left$(X$, 2))
                End If
            End If
        Next U&
        MODICLI.Text1.TEXT = NROPRO& + 1
        Screen.MousePointer = 1
        MODICLI.Show
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FINAL("")
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    MODUCONE$ = TABLA$(Index)
    TETABLA$ = MODUCONE$
    If Left$(TETABLA$, 1) = "#" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(#)"
      ElseIf Left$(TETABLA$, 1) = "*" Then
        TETABLA$ = Mid$(TETABLA$, 2) + "(*)"
    End If
    Text1.TEXT = TETABLA$
    '
    For K% = 4 To 9
       If K% <> Index Then
          Option1(K%).Value = False
       End If
    Next K%
    Command16.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If TRIM$(MODUCONE$) <> "" Then
        If TRIM$(MODUCONE$) = "GRUCOCLI" Then
            LDT$ = LUDAT$
            Call COPYFILE(LDT$ + "GRUCOCLI.DAT", "C:\FLEXOFT\GRUCOCLI.WKF")
            VTB% = VENTABU%("GRUCOCLI")
            If VTB% < 0 Then Exit Sub
            J& = 0
            For I& = 1 To ULTREG&("!GRUCOCLI")
                X$ = REGLEIDO$("!GRUCOCLI", I&)
                If CVI(Left$(X$, 2)) > 0 Then
                    J& = J& + 1
                    Call GRAREG("GRUCOCLI", X$, J&)
                End If
            Next I&
            Call SETULTREG("GRUCOCLI", J&)
            Call CIERRARCH("GRUCOCLI")
          ElseIf MODUCONE = "TADEPOSI" Then
            Call CARTADEPOSI
          ElseIf MODUCONE = "TATIMAT" Then
            Call CARTATIMA
          ElseIf MODUCONE = "CONPAGO" Then
            Call CARCONPAGOS
          ElseIf MODUCONE = "TAPROVIN" Then
            Call CARTAPROVIN
          ElseIf MODUCONE = "CUEINGAS" Then
            Call CARCUEINGAS
          Else
            Call CONECRUN(MODUCONE$, Text2.TEXT)
        End If
    End If
End Sub

Private Sub Text1_GotFocus()
    Command16.SetFocus
End Sub
Private Sub Text2_GotFocus()
    Command16.SetFocus
End Sub

Sub IMPORCLI()
   '
   If ULTREG&("DEUDOR1") > 0 Then
      Call MENSERR(24, "Importacion Imposible.\Ya Existe Archivo de Clientes y Distribuidores.")
      Exit Sub
   End If
   '
   If EXISTE%("A:\CLIENTES.PRN") < 1 Then Exit Sub
   If EXISTE%("A:\DISTRIBU.PRN") < 1 Then Exit Sub
   '
   Dim PUNUCLI As String * 32767
   '
   Dim NCLI%(32767)
   Dim COVENDE%(32), DEVENDE$(32)
   Dim COCONDI%(32), DECONDI$(32)
   Dim COPROVIN$(32), DEPROVIN$(32)
   '
   LU$ = LUDAT$
   '
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   YY0$ = REGBLAN$("BADACLIE")
   PIVA% = 1
   CATIB% = 2
   LIPRE% = 1
   CUEMA% = CUEINT%("1310-01")
   '
   Call REPLA(XX0$, Chr$(PIVA%), 136, 1)
   Call REPLA(XX0$, Chr$(LIPRE%), 151, 1)
   Call REPLA(XX0$, Chr$(COPAGI%), 150, 1)
   '
   XX$ = MKI$(1) + "Distribuidores": Call GRAREG("GRUCOCLI", XX$, 1)
   XX$ = MKI$(2) + "Clientes": Call GRAREG("GRUCOCLI", XX$, 2)
   Call SETULTREG("GRUCOCLI", 2)
   Call CIERRARCH("GRUCOCLI")
   '
   CAPROVI% = 0
   For U& = 1 To ULTREG&("PROVINC")
      XX$ = REGLEIDO$("PROVINC", U&)
      CAPROVI% = CAPROVI% + 1
      COPROVIN$(CAPROVI%) = Left$(XX$, 1)
      DEPROVIN$(CAPROVI%) = TRIM$(UCase$(Mid$(XX$, 2, 29)))
   Next U&
   '
   CACONDI% = 0
   If ULTREG("CONPAG") < 1 Then
      XX$ = REGBLAN$("CONPAG")
      Call REPLA(XX$, Chr$(1), 1, 1)
      Call REPLA(XX$, "CONTADO", 2, 35)
      Call REPLA(XX$, Chr$(1), 37, 1)
      Call REGAPP("CONPAG", XX$)
      Call CIERRARCH("CONPAG")
   End If
   For U& = 1 To ULTREG&("CONPAG")
      XX$ = REGLEIDO$("CONPAG", U&)
      CACONDI% = CACONDI% + 1
      COCONDI%(CACONDI%) = Asc(Left$(XX$, 1))
      DECONDI$(CACONDI%) = TRIM$(UCase$(Mid$(XX$, 2, 35)))
   Next U&
   '
   Call BLANARCH("VENDEDOR")
   CAVENDE% = 0
   For U& = 1 To ULTREG&("VENDEDOR")
      XX$ = REGLEIDO$("VENDEDOR", U&)
      CAVENDE% = CAVENDE% + 1
      COVENDE%(CAVENDE%) = Asc(Left$(XX$, 1))
      DEVENDE$(CAVENDE%) = TRIM$(UCase$(Mid$(XX$, 2, 30)))
   Next U&
   '
   REGX& = 0: REGY& = 0: NCIMAX% = 0
   '
' IMPORTACION DISTRIBUIDORES
   '
   GRUCLI% = 1
   N1% = FILEOPEN("A:\DISTRIBU.PRN", 1, 0)
   '
5  If EOF(N1%) Then
      GoTo 9
   End If
   '
   Line Input #N1%, AA$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 1, 4))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      If NCX& > 0 Then
         If NCX& < 1000 Then
            NCX& = NCX& + 1000
         End If
      End If
      NCI% = NCX&
      If NCI% > 0 Then
         If Mid$(PUNUCLI$, NCI%, 1) = "1" Then NCI% = 0
      End If
      '
      RASO$ = TRIM$(Mid$(AA$, 6, 45))
      FANT$ = Left$(RASO$, 10)
      DOMI$ = TRIM$(Mid$(AA$, 51, 27))
      CPOSTX$ = TRIM$(Mid$(AA$, 78, 4))
      CPOST% = Val(CPOSTX$)
      LOCA$ = TRIM$(Mid$(AA$, 82, 24))
      DEPROV$ = TRIM$(UCase$(Mid$(AA$, 106, 22))): GoSub 60
      TELE$ = Mid$(AA$, 128, 20)
      CUITC$ = AJUSTI$(Mid$(AA$, 148, 11), 11): GoSub 61
      EXPRESO$ = TRIM$(Mid$(AA$, 159, 23))
      VENDE$ = TRIM$(UCase$(Mid$(AA$, 182, 7)))
      NUVENDE% = Val(Mid$(AA$, 189, 2)): GoSub 63
      CONDIPAG$ = TRIM$(UCase$(Mid$(AA$, 191, 21))): GoSub 64
      LIMICREDI$ = TRIM$(Mid$(AA$, 212, 8)): GoSub 62
      '
      If TRIM$(RASO$) <> "" Then GoSub GRACLIEN
      '
   End If
   '
   GoTo 5
   '
9  Close #N1%: N1% = 0
   '
   KK% = 0
   For U& = 1 To ULTREG&("DEUDOR1")
      XX$ = REGLEIDO$("DEUDOR1", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("DEUDOR1", XX$, U&)
      End If
   Next U&
   '
   KK% = 0
   For U& = 1 To ULTREG&("BADACLIE")
      XX$ = REGLEIDO$("BADACLIE", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("BADACLIE", XX$, U&)
         Mid$(PUNUCLI$, NCIMAX% + KK%, 1) = "1"
      End If
   Next U&
   '
   NCIMAX% = NCIMAX% + KK%
   Call SETULTREG("DEUDOR1", REGX&)
   Call SETULTREG("DEUDOR2", REGX&)
   Call SETULTREG("BADACLIE", REGY&)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   Call CIERRARCH("*.*")
   '
   '
' IMPORTACION CLIENTES
   '
   GRUCLI% = 2
   N1% = FILEOPEN("A:\CLIENTES.PRN", 1, 0)
   Line Input #N1%, AA$
   '
15 If EOF(N1%) Then
      GoTo 19
   End If
   '
   Line Input #N1%, AA$
   If TRIM$(AA$) <> "" Then
      NCX& = Val(Mid$(AA$, 4, 4))
      If NCX& < 0 Then NCX& = 0
      If NCX& > 32767 Then NCX& = 0
      If NCX& > 0 Then
         While NCX& < 2000: NCX& = NCX& + 1000: Wend
      End If
      NCI% = NCX&
      If NCI% > 0 Then
         If Mid$(PUNUCLI$, NCI%, 1) = "1" Then NCI% = 0
      End If
      '
      RASO$ = TRIM$(Mid$(AA$, 11, 43))
      FANT$ = Left$(RASO$, 10)
      DOMI$ = TRIM$(Mid$(AA$, 54, 34))
      CPOSTX$ = TRIM$(Mid$(AA$, 88, 4))
      CPOST% = Val(CPOSTX$)
      LOCA$ = TRIM$(Mid$(AA$, 95, 26))
      DEPROV$ = TRIM$(UCase$(Mid$(AA$, 121, 22))): GoSub 60
      TELE$ = Mid$(AA$, 143, 25)
      CUITC$ = AJUSTI$(Mid$(AA$, 168, 11), 11): GoSub 61
      EXPRESO$ = ""
      VENDE$ = TRIM$(UCase$(Mid$(AA$, 179, 26)))
      NUVENDE% = Val(Mid$(AA$, 205, 2)): GoSub 63
      CONDIPAG$ = "CONTADO": GoSub 64
      LIMICREDI$ = "0": GoSub 62
      '
      If TRIM$(RASO$) <> "" Then GoSub GRACLIEN
      '
   End If
   '
   GoTo 15
   '
19 Close #N1%: N1% = 0
   '
   KK% = 0
   For U& = 1 To ULTREG&("DEUDOR1")
      XX$ = REGLEIDO$("DEUDOR1", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("DEUDOR1", XX$, U&)
      End If
   Next U&
   '
   KK% = 0
   For U& = 1 To ULTREG&("BADACLIE")
      XX$ = REGLEIDO$("BADACLIE", U&)
      If CVI(Left$(XX$, 2)) < 1 Then
         KK% = KK% + 1
         Call REPLA(XX$, MKI$(NCIMAX% + KK%), 1, 2)
         Call GRAREG("BADACLIE", XX$, U&)
         Mid$(PUNUCLI$, NCIMAX% + KK%, 1) = "1"
      End If
   Next U&
   '
   NCIMAX% = NCIMAX% + KK%
   Call SETULTREG("DEUDOR1", REGX&)
   Call SETULTREG("DEUDOR2", REGX&)
   Call SETULTREG("BADACLIE", REGY&)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   Call CIERRARCH("*.*")
   '
   Call COMUNI("La Importación de Clientes y Distribuidores\se ha Completado Exitosamente.")
   '
20 N1% = FILEOPEN("A:PRECIOS.PRN", 1, 0)
   CI0% = NUMAS%
   Dim CODIMPOR$(16384)
   Dim DEGRU$(256)
   CAIMPOR% = 0
   '
   J& = 0
   LINACT$ = "00"
   While Not EOF(N1%)
      Line Input #N1%, AA$
      If TRIM$(AA$) <> "" Then
         COD$ = TRIM$(Left$(AA$, 10))
         If COD$ = "Línea" Then
            CAGRU% = CAGRU% + 1
            DEGRU$(CAGRU%) = TRIM$(Mid$(AA$, 11, 42))
            LINACT$ = Left$(DEGRU$(CAGRU%), 2)
            GoTo 28
         End If
         CODLIN$ = LINACT$ + "-" + COD$
         '
         For KI% = 1 To CAIMPOR%
            If CODIMPOR$(KI%) = CODLIN$ Then
               Call COMUNI("Imposible Importar Codigo " + COD$ + " - Repetido.")
               GoTo 28
            End If
         Next KI%
         '
         CAIMPOR% = CAIMPOR% + 1
         CODIMPOR$(CAIMPOR%) = CODLIN$
         '
         DE$ = DESCRIT0$(CI%)
         LIMICREDI$ = Mid$(AA$, 53, 10): GoSub 62
         PRE# = CDbl(LICRE)
         CI% = CODINT%(CODLIN$)
         If CI% > 0 Then GoTo 25
         '
         CI% = CODINT%(COD$)
         If CI% > 0 Then
            REMA& = CI% + 1
            REMAS$ = REGLEIDO$("MASTER", REMA&)
            Call REPLA(REMAS$, CODLIN$, 77, 16)
            Call GRAREG("MASTER", REMAS$, REMA&)
            GoTo 25
         End If
         '
         CI0% = CI0% + 1
         CI% = CI0%
         DE$ = Mid$(AA$, 11, 42)
         REMA& = CI% + 1
         REMAS$ = String$(128, 0)
         Call REPLA(REMAS$, MKI$(CI%), 1, 2)
         Call REPLA(REMAS$, DE$, 5, 64)
         Call REPLA(REMAS$, String$(8, 0), 69, 8)
         Call REPLA(REMAS$, CODLIN$, 77, 16)
         Call REPLA(REMAS$, MKI$(CAGRU%), 125, 2)
         Call GRAREG("MASTER", REMAS$, REMA&)
         '
25       Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, DE$, 3, 42)
         Call REPLA(Z$, MKD$(PRE#), 62, 8)
         J& = J& + 1
         Call GRAREG("LIPRE001", Z$, J&)
      End If
28 Wend
   '
29 Close #N1%
   Call SETULTREG("LIPRE001", J&)
   Call CIERRARCH("LIPRE001")
   J& = 0
   For KI% = 1 To CAGRU%
      Z$ = REGBLAN$("GRUCOVEN")
      Call REPLA(Z$, MKI$(KI%), 1, 2)
      Call REPLA(Z$, DEGRU$(KI%), 3, 30)
      J& = J& + 1
      Call GRAREG("GRUCOVEN", Z$, J&)
   Next KI%
   Call SETULTREG("GRUCOVEN", J&)
   Call CIERRARCH("GRUCOVEN")
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   '
   Call COMUNI("La Importación de Listas de Precios\se ha Completado Exitosamente.\  \Ejecute Ahora un Control del Maestro de Articulos.")
   Call FINAL("")
   '
Exit Sub
   '
GRACLIEN:
      XX$ = XX0$
      Call REPLA(XX$, MKI$(NCI%), 1, 2)
      Call REPLA(XX$, TRIM$(RASO$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, COPROVI$, 149, 1)
      Call REPLA(XX$, Chr$(CONDI%), 150, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COVEN%), 152, 1)
      Call REPLA(XX$, MKS$(LICRE), 153, 4)
      Call REPLA(XX$, MKI$(GRUCLI%), 157, 2)
      Call REPLA(XX$, MKI$(CUEMA%), 159, 2)
      '
      REGX& = REGX& + 1
      Call GRAREG("DEUDOR1", Left$(XX$, 32), REGX&)
      Call GRAREG("DEUDOR2", Mid$(XX$, 33), REGX&)
      '
      YY$ = YY0$
      Call REPLA(YY$, MKI$(NCI%), 1, 2)
      Call REPLA(YY$, TRIM$(RASO$), 3, 62)
      Call REPLA(YY$, TRIM$(FANT$), 65, 32)
      Call REPLA(YY$, TRIM$(DOMI$), 97, 64)
      Call REPLA(YY$, TRIM$(CPOSTX$), 161, 16)
      Call REPLA(YY$, TRIM$(LOCA$), 177, 48)
      Call REPLA(YY$, "ARGENTINA", 225, 32)
      Call REPLA(YY$, TRIM$(TELE$), 257, 24)
      Call REPLA(YY$, TRIM$(EXPRESO$), 353, 32)
      '
      REGY& = REGY& + 1
      Call GRAREG("BADACLIE", YY$, REGY&)
      '
      If NCI% > 0 Then
         Mid$(PUNUCLI$, NCI%, 1) = "1"
      End If
      '
      If NCI% > NCIMAX% Then NCIMAX% = NCI%
      '
Return
   '
   '
60  For KI% = 1 To CAPROVI%
      If DEPROVIN$(KI%) = DEPROV$ Then
         COPROVI$ = COPROVIN$(KI%)
         Return
      End If
    Next KI%
    COPROVI$ = "A"
601 For KI% = 1 To CAPROVI%
      If COPROVIN$(KI%) = COPROVI$ Then
         COPROVI$ = Chr$(1 + Asc(COPROVI$))
         GoTo 601
      End If
    Next KI%
    CAPROVI% = CAPROVI% + 1
    COPROVIN$(CAPROVI%) = COPROVI$
    DEPROVIN$(CAPROVI%) = DEPROV$
    XX$ = REGBLAN$("PROVINC")
    Call REPLA(XX$, COPROVI$, 1, 1)
    Call REPLA(XX$, DEPROV$, 2, 29)
    Call REGAPP("PROVINC", XX$)
Return
   '
   '
61 CUIT$ = Left$(CUITC$, 2) + "-" + Mid$(CUITC$, 3, 8) + "-" + Mid$(CUITC$, 11, 1)
   Return
   '
62 LCRX$ = REPLACAR$(LIMICREDI$, ".", "")
   LCRY$ = REPLACAR$(LCRX$, ",", ".")
   LICRE = Val(LCRY$)
   Return
   '
63  If VENDE$ = "ALCALAY" Then VENDE$ = "ALCALAYA"
    If TRIM$(VENDE$) = "0" Then VENDE$ = ""
    If TRIM$(VENDE$) = "" Then
       COVEN% = 0
       Return
    End If
    For KI% = 1 To CAVENDE%
      If DEVENDE$(KI%) = VENDE$ Then
         COVEN% = COVENDE%(KI%)
         Return
      End If
    Next KI%
    COVEN% = NUVENDE%
631 For KI% = 1 To CAVENDE%
      If COVENDE%(KI%) = COVEN% Then
         VENDE$ = DEVENDE$(KI%)
         Return
      End If
    Next KI%
    CAVENDE% = CAVENDE% + 1
    COVENDE%(CAVENDE%) = COVEN%
    DEVENDE$(CAVENDE%) = VENDE$
    XX$ = REGBLAN$("VENDEDOR")
    Call REPLA(XX$, Chr$(COVEN%), 1, 1)
    Call REPLA(XX$, VENDE$, 2, 30)
    Call REPLA(XX$, VENDE$, 32, 12)
    Call REGAPP("VENDEDOR", XX$)
Return
   '
   '
64  If TRIM$(CONDIPAG$) = "" Then
       CONDI% = 1
       Return
    End If
    For KI% = 1 To CACONDI%
      If DECONDI$(KI%) = CONDIPAG$ Then
         CONDI% = COCONDI%(KI%)
         Return
      End If
    Next KI%
    CONDI% = 1
641 For KI% = 1 To CACONDI%
      If COCONDI%(KI%) = CONDI% Then
         CONDI% = CONDI% + 1
         GoTo 641
      End If
    Next KI%
    CACONDI% = CACONDI% + 1
    COCONDI%(CACONDI%) = CONDI%
    DECONDI$(CACONDI%) = CONDIPAG$
    XX$ = REGBLAN$("CONPAG")
    Call REPLA(XX$, Chr$(CONDI%), 1, 1)
    Call REPLA(XX$, CONDIPAG$, 2, 35)
    Call REPLA(XX$, Chr$(1), 37, 1)
    Call REGAPP("CONPAG", XX$)
Return
   '
End Sub

