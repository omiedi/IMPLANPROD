VERSION 5.00
Begin VB.Form AMAPRO1 
   BackColor       =   &H00C0E0EE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores y Subcontratistas - ABML"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
   ClipControls    =   0   'False
   Icon            =   "AMAPRO1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   540
      Left            =   6510
      TabIndex        =   51
      Top             =   210
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5985
      ScaleHeight     =   75
      ScaleWidth      =   1380
      TabIndex        =   49
      Top             =   1050
      Width           =   1440
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command6 
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
      Left            =   4830
      Picture         =   "AMAPRO1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Configuración - Valores Predeterminados por Omisión"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0EE&
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
      Height          =   1275
      Left            =   4830
      TabIndex        =   44
      Top             =   4620
      Width           =   4635
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   45
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
            Picture         =   "AMAPRO1.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Consulta de Resumen de Cuenta Corriente por Proveedor"
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
            TabIndex        =   47
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0EE&
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
      Height          =   3165
      Left            =   4830
      TabIndex        =   5
      Top             =   1365
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
         TabIndex        =   13
         Top             =   420
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
         TabIndex        =   15
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
         TabIndex        =   14
         Text            =   " "
         Top             =   420
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         Picture         =   "AMAPRO1.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0E0EE&
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
         TabIndex        =   6
         Top             =   1365
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0EE&
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
            TabIndex        =   10
            Top             =   945
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0EE&
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
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0EE&
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
            TabIndex        =   8
            Top             =   420
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0EE&
            Caption         =   "Grupos Conta- bles de Compras"
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
            TabIndex        =   7
            Top             =   945
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
         TabIndex        =   22
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0EE&
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
      Height          =   5685
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0EE&
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
         Height          =   2360
         Left            =   210
         TabIndex        =   4
         Top             =   3150
         Width           =   4005
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0E0EE&
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
            TabIndex        =   41
            Top             =   1575
            Width           =   3165
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0E0EE&
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
               TabIndex        =   43
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0E0EE&
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
               TabIndex        =   42
               Top             =   250
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   35
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
               Left            =   0
               Picture         =   "AMAPRO1.frx":0D60
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Listados Clasificados"
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
               TabIndex        =   37
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   38
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
               Picture         =   "AMAPRO1.frx":106A
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado General"
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
               TabIndex        =   40
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0EE&
         Caption         =   "Archivo de Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1365
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   32
            Top             =   1680
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
               Picture         =   "AMAPRO1.frx":1374
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Alta de Nuevo Proveedor"
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
               TabIndex        =   34
               Top             =   105
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   23
            Top             =   1260
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
               Picture         =   "AMAPRO1.frx":167E
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificación de Datos"
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
               TabIndex        =   25
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   26
            Top             =   840
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
               Picture         =   "AMAPRO1.frx":1988
               Style           =   1  'Graphical
               TabIndex        =   27
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
               TabIndex        =   28
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
            TabIndex        =   29
            Top             =   420
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
               Picture         =   "AMAPRO1.frx":1C92
               Style           =   1  'Graphical
               TabIndex        =   30
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
               TabIndex        =   31
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
      Left            =   8610
      Picture         =   "AMAPRO1.frx":1F9C
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
      Left            =   7665
      Picture         =   "AMAPRO1.frx":20E6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "AMAPRO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$
'

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command12_Click()
    If DERACCE%("MODI-PRO", "Modificacion de Datos de Proveedor") < 1 Then
        Exit Sub
    End If
    MODO% = 2
    MODIPRO.Show
End Sub


Private Sub Command16_Click()
    Command16.Enabled = False
    If MODUCONE = "CUEINGAS" Then
         Call CARCUEINGAS
     ElseIf MODUCONE = "CONPAGO" Then
         Call CARCONPAGOS
     ElseIf MODUCONE = "TAPROVIN" Then
         Call CARTAPROVIN
     ElseIf TRIM$(MODUCONE$) <> "" Then
        Call CONECRUN(MODUCONE$, Text2.TEXT)
    End If
    Command16.Enabled = True

End Sub

Private Sub Command17_Click()
    Call HELPONLINE("AMAPRO")
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    OPX% = COMALTER%("Alta de una Nueva Cuenta\\Cuenta Corriente\Cuenta de Gastos\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    OPX1% = OPX% - 1
    MODO% = 1
    Screen.MousePointer = 11
    Call CLEARFORMPRO
    NROPRO& = 0
    For U& = 1 To ULTREG&("PROVED1")
        X$ = REGLEIDO$("PROVED1", U&)
        NPX1% = CVI(Left$(X$, 2))
        If NPX1% > NROPRO& Then
          If TICUEPRO%((-1) * NPX1%) = OPX1% Then
            NROPRO& = CVI(Left$(X$, 2))
          End If
        End If
    Next U&
    '
10  MODIPRO.Text1.TEXT = NROPRO& + 1
    NCX% = (-1) * XVALO(MODIPRO.Text1)
    If REGICLI&(NCX%) > 0 Then
      NROPRO& = NROPRO& + 1
      GoTo 10
    End If
    '
    If OPX1% = 0 Then
      For Index = 0 To 5
        VAPRE$ = TRIM$(CONTROL$("AMAPRO", ARCHE$(Index)))
        MODIPRO.COTEXT(Index) = ""
        MODIPRO.DETEXT(Index) = ""
        If VAPRE$ <> "" Then
          TV11% = TIPFIELD(ARCHE$(Index), 1)
          LF11% = LENFIELD(ARCHE$(Index), 1)
          ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
          If ECO$ <> "" Then
              MODIPRO.COTEXT(Index).TEXT = VAPRE$
              MODIPRO.DETEXT(Index).TEXT = ECO$
            Else
              MODIPRO.COTEXT(Index).TEXT = ""
              MODIPRO.DETEXT(Index).TEXT = ""
          End If
        End If
        MODIPRO.DETEXT(Index).TabStop = False
      Next Index
    End If
    '
    Screen.MousePointer = 1
    MODIPRO.Show
    '
99  Command2.Enabled = True
End Sub
Private Sub Command3_Click()
   LU$ = LUDAT$
   N1% = FILEOPEN%(LU$ + "PROVED1.OLD", 0, 32)
   Dim NX1 As String * 32
   N2% = FILEOPEN%(LU$ + "PROVED2.OLD", 0, 128)
   Dim NX2 As String * 128
   For U& = 1 To LOF(N1%) / 32
      Get #N1%, U&, NX1$
      Get #N2%, U&, NX2$
      NC% = (-1) * CVI(Left$(NX1$, 2))
      CUE% = CVI(Mid$(NX2$, 127, 8))
      CUEMAOLD$ = CODCUE$(CUE%)
      CUEMANEW$ = CUEMADRE$(NC%)
      If TRIM$(CUEMAOLD$) <> TRIM$(CUEMANEW$) Then
         MsgBox Abs(NC%) & "   " & CUEMAOLD$ & "   " & CUEMANEW$
      End If
   Next U&
      
End Sub

Private Sub Command4_Click()
Call BLOQARCH("PROVED2")
Call BLOQARCH("DEUDOR1")
Call BLOQARCH("PROVED1")
MsgBox "PROVED1 BLOQUEADO"
Call LIBARCH("PROVED1")
End Sub

Private Sub Command5_Click()
    If Option1(0).Value = True Then
       Call COMUNI("Opción no Disponible por Pantalla.\  \Emita Listado Impreso\o Exporte a Archivo de Texto.")
       Exit Sub
    End If
    Command5.Enabled = False
    FORLIPROVE.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    OPALTPRO.Show 1
    Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    If Option1(0).Value = True Then
       Call SELECHO("PROVED1")
       Exit Sub
    End If
    Command8.Enabled = False
    Call GELIGEPRO
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    PROLLAMA$ = Name
    Hide
    RESCUPRO.Show
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call CARTAVALI
    ARCHE$(3) = "TICUEPRO"
    ARCHE$(4) = "CONPAG"
    ARCHE$(5) = "ECUECON"
    ARCHE$(6) = ""
    ARCHE$(7) = ""
    '
    TABLA$(0) = ""       ' tabla de tipos de material
    TABLA$(1) = "*UNIMED"
    TABLA$(2) = "#DHL154"
    TABLA$(3) = ""       ' grupos contables por articulo
    TABLA$(4) = "*VENDEDOR"
    TABLA$(5) = "TAPROVIN"
    TABLA$(6) = "CONPAGO"
    TABLA$(7) = "ADMIPRE"
    TABLA$(8) = "CUEINGAS"
    TABLA$(9) = "*CUECOVE"
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Cancel = 0
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
    For K% = 5 To 9
      If K% <> 7 Then
        If K% <> Index Then
          Option1(K%).Value = False
        End If
      End If
    Next K%
    Command16.SetFocus
End Sub

Private Sub Option1_DblClick(Index As Integer)
    If MODUCONE = "CUEINGAS" Then
         Call CARCUEINGAS
      ElseIf MODUCONE = "CONPAGO" Then
         Call CARCONPAGOS
      ElseIf MODUCONE = "TAPROVIN" Then
         Call CARTAPROVIN
      ElseIf TRIM$(MODUCONE$) <> "" Then
        Call CONECRUN(MODUCONE$, Text2.TEXT)
    End If
End Sub

Private Sub Text1_GotFocus()
    Command16.SetFocus
End Sub
Private Sub Text2_GotFocus()
    Command16.SetFocus
End Sub

