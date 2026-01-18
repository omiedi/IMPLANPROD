VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPOR07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Precon Meditea"
      Height          =   645
      Left            =   9240
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   750
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
      Height          =   735
      Left            =   11025
      Picture         =   "impor07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Exit"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11010
      Picture         =   "impor07.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Inicia Importación Definitiva"
      Top             =   1155
      Width           =   750
   End
   Begin VB.CommandButton CMDCONVERSION 
      Caption         =   "Precon-version Dakot"
      Height          =   645
      Left            =   10185
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Resultado de Pre-Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   105
      TabIndex        =   14
      ToolTipText     =   "Importación Definitiva"
      Top             =   3990
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   30
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   32
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1890
               ScaleWidth      =   0
               TabIndex        =   37
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               Height          =   270
               Left            =   0
               ScaleHeight     =   210
               ScaleWidth      =   1.1994e5
               TabIndex        =   34
               Top             =   0
               Width           =   1.2e5
               Begin VB.Label Label83 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   150
                  TabIndex        =   38
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   40
                  TabIndex        =   36
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label82 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   105
                  TabIndex        =   35
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.ListBox List4 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1700
               IntegralHeight  =   0   'False
               ItemData        =   "impor07.frx":0454
               Left            =   0
               List            =   "impor07.frx":0456
               TabIndex        =   33
               Top             =   280
               Width           =   1.3e5
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   31
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         IntegralHeight  =   0   'False
         ItemData        =   "impor07.frx":0458
         Left            =   120
         List            =   "impor07.frx":045A
         TabIndex        =   24
         Top             =   2910
         Width           =   10400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   10650
         Picture         =   "impor07.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Listado por Impresora"
         Top             =   3255
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide"
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
         Left            =   10680
         TabIndex        =   22
         Top             =   560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10710
         Picture         =   "impor07.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1725
         Left            =   10360
         TabIndex        =   16
         Top             =   670
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   15
         Top             =   2310
         Width           =   10250
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   420
         Y2              =   2520
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Registros no Procesables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2670
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   20
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10620
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   120
      TabIndex        =   12
      Top             =   3990
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "impor07.frx":0A70
         Height          =   3750
         Left            =   150
         OleObjectBlob   =   "impor07.frx":0A94
         TabIndex        =   13
         Top             =   360
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   2730
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   9
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Formato de Importacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   6100
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   10
         Top             =   350
         Width           =   5655
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Col. Dato / Variable            Tipo     Rango"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   11
            Top             =   0
            Width           =   5655
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Archivo a Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   105
      TabIndex        =   1
      Top             =   240
      Width           =   4575
      Begin VB.PictureBox Picture6 
         Height          =   2100
         Left            =   210
         ScaleHeight     =   2040
         ScaleWidth      =   4050
         TabIndex        =   29
         Top             =   350
         Width           =   4110
         Begin VB.VScrollBar VScroll1 
            Height          =   2050
            Left            =   3840
            TabIndex        =   62
            Top             =   -15
            Width           =   225
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFF0&
            BorderStyle     =   0  'None
            Height          =   6645
            Left            =   0
            ScaleHeight     =   6645
            ScaleWidth      =   3900
            TabIndex        =   42
            Top             =   0
            Width           =   3900
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Especificaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   240
               TabIndex        =   65
               Top             =   6240
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Ubicaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   240
               TabIndex        =   64
               Top             =   5950
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Padron de Empleados"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   255
               TabIndex        =   61
               Top             =   1400
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Asientos Contables de Diario"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   255
               TabIndex        =   60
               Top             =   2660
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Lista de Sucursales por Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   255
               TabIndex        =   59
               Top             =   1985
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Maestro de Items de Stock"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   255
               TabIndex        =   58
               Top             =   105
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Fórmulas y Estructuras de Producto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   255
               TabIndex        =   57
               Top             =   330
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Rutas - Secuencias de Operaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   255
               TabIndex        =   56
               Top             =   540
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Tabla de Equipos de Producción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   255
               TabIndex        =   55
               Top             =   945
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Categorías de Operarios Productivos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   255
               TabIndex        =   54
               Top             =   1170
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Padron de Clientes y Deudores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   255
               TabIndex        =   53
               Top             =   1760
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Proveedores y Rubros de Gastos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   255
               TabIndex        =   52
               Top             =   2210
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Plan de Cuentas Contables"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   255
               TabIndex        =   51
               Top             =   2435
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Numeros Originales Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   255
               TabIndex        =   50
               Top             =   3150
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Composición de Deuda Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   240
               TabIndex        =   49
               Top             =   3555
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Composición de Deuda Proveedores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   240
               TabIndex        =   48
               Top             =   3810
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Cartera de Valores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   240
               TabIndex        =   47
               Top             =   4065
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Movimientos de Stock Valorizados"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   240
               TabIndex        =   46
               Top             =   4815
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
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
               Height          =   195
               Index           =   13
               Left            =   240
               TabIndex        =   45
               Top             =   4305
               Width           =   3420
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Costos y Moneda de Reposición"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   240
               TabIndex        =   44
               Top             =   5235
               Width           =   3300
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Transportistas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   240
               TabIndex        =   43
               Top             =   5655
               Width           =   3300
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   -105
               X2              =   4455
               Y1              =   870
               Y2              =   870
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   -105
               X2              =   4455
               Y1              =   1700
               Y2              =   1700
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   3000
               Y2              =   3000
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   3465
               Y2              =   3465
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   4725
               Y2              =   4725
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   5145
               Y2              =   5145
            End
            Begin VB.Line Line8 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   5565
               Y2              =   5565
            End
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   3840
            Y1              =   1680
            Y2              =   1680
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   2940
      Width           =   11655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   26
         Top             =   480
         Width           =   6975
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Left            =   20
            TabIndex        =   27
            Top             =   20
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   3
         Text            =   " "
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton Command22 
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
         Left            =   7920
         TabIndex        =   2
         Top             =   465
         Width           =   220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2200
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   28
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPOR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CANCELIMP%

Sub IMPORASIE()
   '
   Dim LISTA$(256)
   Call ABREPLANCUEN
   '
20 On Error Resume Next
   Set EncaSiGra = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
21 Set DetasiGra = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
     GoTo 90
   End If
   '
   JJ& = 0
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     NUMERO = XVALO(TRIM$(Mid$(TTXX$, 1, 12)))
     If NUMERO <> NUMEROANT Then
        If JJ& > 0 Then GoSub GRABATODO
        JJ& = 0
        NUMEROANT = NUMERO
        FECHA$ = TRIM$(Mid$(TTXX$, 16, 8))
        NUASI$ = NUPROA$(FECHA$)
        REFEX8$ = TRIM$(Mid$(TTXX$, 32, 48))
        SUMADEBE# = 0
     End If
     JJ& = JJ& + 1
     LISTA(JJ&) = TTXX$
     IDEBE# = XVALO(TRIM$(Mid$(LISTA$(JJ&), 146, 13)))
     SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
   Next U&
   GoSub GRABATODO
   '
   EncaSiGra.Close
   DetasiGra.Close
   Call COMUNI("Fin de la Importación Definitiva")
   Exit Sub
   '
GRABATODO:
   With EncaSiGra
      .AddNew   ' YA CONVERTIDO
      ON ERROR RESUME NEXT
      !CodiEmpr = Codiemp%
      ON ERROR GOTO 0
      !NUIDASIEN = NUASI$
      !Refeasien = REFEX8$
          FECHAU% = FECHANUM(FECHA$)
      !FECHASIEN = CVDAT(FECHAU%)
          PPXX% = InStr(NUASI$, "-")
      !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
      !SucuAsien = 0
      !NuDfAsien = 0
      !TIPOASIEN = 0
         !StatAsien = 1  ' activo
      !EditAsien = 1
      !QPasAsien = JJ&
      !DebeAsien = SUMADEBE#
      !Observa = " "
      !Nume_Usuar = USNBR%
      .Update
   End With
   '
   ORPASE% = 0
   For K2% = 1 To JJ&
      COCUE$ = TRIM$(Mid$(LISTA$(K2%), 83, 16))
      DEPASE$ = TRIM$(Mid$(LISTA$(K2%), 100, 46))
      IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
         SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
      IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
      MACON$ = " "
      CENCOS1$ = " "
      '
      With DetasiGra
         .AddNew   ' YA CONVERTIDO
         ON ERROR RESUME NEXT
         !CodiEmpr = Codiemp%
         ON ERROR GOTO 0
         !NUIDASIEN = NUASI$
           ORPASE% = ORPASE% + 1
         !OrNuPase = ORPASE%
         !REFEPASE = DEPASE$
           FECHAU% = FECHANUM(FECHA$)
         !FECHASIEN = CVDAT(FECHAU%)
           PPXX% = InStr(NUASI$, "-")
         !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
         !SucuAsien = 0
         !NuDfAsien = 0
         !StatPase = 1
         !CUECOEXT = COCUE$
         !CUECOINT = CUEINT%(COCUE$)
         !IDebePase = IDEBE#
         !IHabePase = IHABE#
         !MARCONSI = MACON$
         !CenCos = CENCOS1$ '\\\OT-06-0244-WAR-16/06/06///
         .Update
      End With
   Next K2%
   '
Return

90 End Sub

Function NUPROA$(FECHA$)
   '
   NUPROA$ = ""
   '
   FEASX = FECHANUM(FECHA$)
   FECHAS$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = CVDAT(FECHAU%)
   NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
   If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   ORDAS% = 0
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset("SELECT * FROM EncaBEAsien WHERE (STR$(FECHASIEN) = '" & FEASYYYY & "' and SucuAsien = " & SUCUI% & ") ORDER BY  ORFEASIEN DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then GoTo 20
   ORDAS% = SELECASIEN!OrFeAsien
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function


Sub PREIMPORASIE()
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     Exit Sub
   End If
   Screen.MousePointer = 11
   DoEvents
   '
   URECORD& = Data1.Recordset.RecordCount
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   For U& = 1 To URECORD&
    Call TRACE(20, U&, URECORD&)
    TTXX$ = Space$(290)
    On Error Resume Next
    DBGrid1.Row = U& - 1
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
      Exit Sub
    End If
    On Error GoTo 0
    IMPO1# = 0
    IMPO2# = 0
    
    DBGrid1.Col = 3
    'Si la cuenta no existe se va al next U&
    'Si existe sigue normalmente
    Cuentta$ = TRIM$(DBGrid1.TEXT)
    CuenExis% = CUEINT%(Cuentta$)
    If CuenExis% = 0 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 19
    End If
    '
    DBGrid1.Col = 0                     'ID ASIENTO
    ASIENTTO$ = TRIM$(DBGrid1.TEXT)
    If Len(ASIENTTO$) < 1 Then
      MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Asiento"
      List1.AddItem MERROX$
      GoTo 19
    End If
    Call REPLA(TTXX$, ASIENTTO$, 1, 12)
    '
    DBGrid1.Col = 1
    FECHHA$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))  ' FECHA
    Call REPLA(TTXX$, FECHHA$, 16, 14)
    '
    DBGrid1.Col = 2                     ' Referencia General
    RefGeneral$ = TRIM$(DBGrid1.TEXT)
    Call REPLA(TTXX$, RefGeneral$, 32, 48)
    '
    DBGrid1.Col = 3                    ' Cuenta
    Cuentta$ = TRIM$(DBGrid1.TEXT)
    If Len(Cuentta$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Cuenta"
       List1.AddItem MERROX$
       GoTo 19
    End If
    Call REPLA(TTXX$, Cuentta$, 83, 16)
    '
    DBGrid1.Col = 4                     ' Referencia Pase
    RefPase$ = TRIM$(DBGrid1.TEXT)
    Call REPLA(TTXX$, RefPase$, 100, 46)
    '
    DBGrid1.Col = 5                     'IMPORTE Debe
    If DBGrid1.TEXT <> "" Then
      IMPO1# = CDbl(TRIM$(DBGrid1.TEXT))
      If IMPO1# >= 0 Then
       IMPODEBE$ = FORMATNUM(IMPO1#, "F13.2")
       Call REPLA(TTXX$, IMPODEBE$, 146, 13)
      End If
    End If
    '
    DBGrid1.Col = 6                     'IMPORTE Haber
    If DBGrid1.TEXT <> "" Then
     If IMPO2# >= 0 Then
      IMPO2# = CDbl(TRIM$(DBGrid1.TEXT))
      IMPOHABER$ = FORMATNUM(IMPO2#, "F13.2")
      Call REPLA(TTXX$, IMPOHABER$, 162, 13)
     End If
    End If
    If (IMPO1# < 0 And IMPO2# < 0) Or (IMPO1# > 0 And IMPO2# > 0) Then
      Call MENSERR(24, "Debe o Haber  Valores Incompatibles")
      GoTo 19
    End If

    List4.AddItem TTXX$
    Label6 = TRIM$(Str$(List4.ListCount))
    If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
    End If
    '
19 Next U&
   '
   DBGrid1.RowHeight = hrow
   '
End Sub

Private Sub CMDCONVERSION_Click()
ATX% = COMALTER%("Conversion de Datos 'DAKOT' \\Quitar Encabezados\Concatenar")
'ATX% = COMALTER%("Secuencia Operaciones de Código '" + codex$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
Select Case ATX%
Case Is = 1
   ATXQuitarE% = COMALTER%("Seleccione a que Archivos Desea Quitarle los Encabezados \\Liscli-10-01-08.TXT\Lispro-10-01-08.txt\Lispre-10-01-08.txt\Lisart-10-01-08.txt")
   Select Case ATXQuitarE%
     Case 1
       Call QUITAR_ENCABEZADO_LISCLI
     Case 2
       Call QUITAR_ENCABEZADO_LISPRO
     Case 3
       Call QUITAR_ENCABEZADO_LISPRE
     Case 4
       Call QUITAR_ENCABEZADO_LISART
  End Select
Case Is = 2
   ATXConcat% = COMALTER%("Seleccione  que Archivos Desea Concatenar \\LiscliOrden1_10-01-08.TXT\LisprOOrden1_10-01-08.txt\LisartOrden1_10-01-08.txt")
   Select Case ATXConcat%
     Case 1
       Call CONCATENAR_LISCLI
     Case 2
       Call CONCATENAR_LISPRO
     Case 3
       Call CONCATENAR_LISART
   End Select
End Select
End Sub
'/****************************
Sub CONCATENAR_LISCLI()

'  CONCATENA LOS REGISTROS DE UN MISMO CLIENTE EN LISCLI
   If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden1_10-01-08.TXT") > 0 Then
      GoTo 20
     Else
      MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
      Exit Sub
   End If

20 N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden1_10-01-08.TXT", 1, 128)
   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden2_10-01-08.TXT", 2, 128)
   CONT& = 0
   Do While Not EOF(N1%)
      Line Input #N1%, A$
      CONT& = CONT& + 1
      If CONT& = 1 Then
        Call REPLA(TXT$, A$, 0, 135)
        ElseIf CONT& = 2 Then
          Call REPLA(TXT$, A$, 140, 275)
          ElseIf CONT& = 3 Then
            Call REPLA(TXT$, A$, 285, 420)
            CONT& = 0
            Print #N2%, TXT$ 'IMPRIME
      End If
   Loop
End Sub


'


Sub CONCATENAR_LISPRO()

   'CONCATENA LOS REGISTROS DE UN MISMO PROVEEDOR EN LISPRO
   If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisprOOrden1_10-01-08.TXT") > 0 Then
      GoTo 20
    Else
      MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
      Exit Sub
   End If

20 N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisprOOrden1_10-01-08.TXT", 1, 128)
   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisproOrden2_10-01-08.TXT", 2, 128)
   CONT& = 0
   Do While Not EOF(N1%)
      Line Input #N1%, A$
      CONT& = CONT& + 1
      If CONT& = 1 Then
        Call REPLA(TXT$, A$, 0, 100)
        ElseIf CONT& = 2 Then
          Call REPLA(TXT$, A$, 105, 205)
          ElseIf CONT& = 3 Then
            Call REPLA(TXT$, A$, 210, 310)
            CONT& = 0
            Print #N2%, TXT$
      End If
   Loop
End Sub
'

'
'
 Sub CONCATENAR_LISART()
  If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisartOrden1_10-01-08.TXT") > 0 Then
      GoTo 20
    Else
      MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
      Exit Sub
   End If
20 N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisartOrden1_10-01-08.TXT", 1, 128)
   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisartOrden2_10-01-08.TXT", 2, 128)
   CONT& = 0
   Do While Not EOF(N1%)
      Line Input #N1%, A$
      CONT& = CONT& + 1
      If CONT& = 1 Then
        Call REPLA(TXT$, A$, 0, 63)
        ElseIf CONT& = 2 Then
          Call REPLA(TXT$, A$, 65, 125)
            Print #N2%, TXT$
            CONT& = 0
      End If
   Loop
   End Sub
Sub QUITAR_ENCABEZADO_LISCLI()
''QUITA TODOS LOS ENCABEZADOS
 If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Liscli-10-01-08.TXT") > 0 Then
    GoTo 20
  Else
    MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
    Exit Sub
 End If
20  N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Liscli-10-01-08.TXT", 1, 128)
    N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\ListcliOrden1_10-01-08", 2, 128)
  '
     Do While Not EOF(N1%)
      Line Input #N1%, A$
      If Left$(A$, 9) <> "@CHF" Then
        If Left$(A$, 7) <> "Gabino" Then
          If Left$(A$, 7) <> "E=====" Then
            If Left$(A$, 7) <> "E    " Then
              If Left$(A$, 7) <> "E     " Then
                If Left$(A$, 7) <> "FH" Then
                  If Left$(A$, 7) <> " SUCUR." Then
                    If Left$(A$, 17) <> "         CONTACTO" Then
                      If Left$(A$, 7) <> "FH" Then
                        If Left$(A$, 7) <> "FH" Then
                          If Left$(A$, 7) <> "  " Then
                            If Left$(A$, 7) <> "ARTICUL" Then
                              If Left(A$, 39) <> "                                       " Then
                                If Left$(A$, 70) <> "" Then
                                    Print #N2%, A$
                                  Else
                                     GoTo 10
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
10   Loop
End Sub
Sub QUITAR_ENCABEZADO_LISPRO()
''QUITA TODOS LOS ENCABEZADOS
 If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lispro-10-01-08.TXT") > 0 Then
    GoTo 20
  Else
    MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
    Exit Sub
 End If

20  N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lispro-10-01-08.TXT", 1, 128)
    N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisprOOrden1_10-01-08", 2, 128)
  '
     Do While Not EOF(N1%)
      Line Input #N1%, A$
      If Left$(A$, 9) <> "@CHF" Then
        If Left$(A$, 7) <> "Gabino" Then
          If Left$(A$, 7) <> "E=====" Then
            If Left$(A$, 7) <> "E    " Then
              If Left$(A$, 7) <> "E     " Then
                If Left$(A$, 7) <> "FH" Then
                  If Left$(A$, 7) <> " SUCUR." Then
                    If Left$(A$, 17) <> "         CONTACTO" Then
                      If Left$(A$, 7) <> "FH" Then
                        If Left$(A$, 7) <> "FH" Then
                          If Left$(A$, 7) <> "  " Then
                            If Left$(A$, 7) <> "ARTICUL" Then
                              If Left(A$, 39) <> "                                       " Then
                                If Left$(A$, 70) <> "" Then
                                    Print #N2%, A$
                                  Else
                                     GoTo 10
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
10   Loop
End Sub
Sub QUITAR_ENCABEZADO_LISPRE()
''QUITA TODOS LOS ENCABEZADOS
 If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lispre-10-01-08.TXT") > 0 Then
    GoTo 20
  Else
    MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
    Exit Sub
 End If

20  N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lispre-10-01-08.TXT", 1, 128)
    N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\Lispre-10-01-08.TXT", 2, 128)
  '
     Do While Not EOF(N1%)
      Line Input #N1%, A$
      If Left$(A$, 9) <> "@CHF" Then
        If Left$(A$, 7) <> "Gabino" Then
          If Left$(A$, 7) <> "E=====" Then
            If Left$(A$, 7) <> "E    " Then
              If Left$(A$, 7) <> "E     " Then
                If Left$(A$, 7) <> "FH" Then
                  If Left$(A$, 7) <> " SUCUR." Then
                    If Left$(A$, 17) <> "         CONTACTO" Then
                      If Left$(A$, 7) <> "FH" Then
                        If Left$(A$, 7) <> "FH" Then
                          If Left$(A$, 7) <> "  " Then
                            If Left$(A$, 7) <> "ARTICUL" Then
                              If Left(A$, 39) <> "                                       " Then
                                If Left$(A$, 70) <> "" Then
                                    Print #N2%, A$
                                  Else
                                     GoTo 10
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
10   Loop
End Sub

Sub QUITAR_ENCABEZADO_LISART()
''QUITA TODOS LOS ENCABEZADOS
 If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lisart-10-01-08.TXT") > 0 Then
    GoTo 20
  Else
    MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
    Exit Sub
 End If

20  N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\ORIGINALES\Lisart-10-01-08.TXT", 1, 128)
    N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisartOrden1_10-01-08.TXT", 2, 128)
  '
     Do While Not EOF(N1%)
      Line Input #N1%, A$
      If Left$(A$, 9) <> "@CHF" Then
        If Left$(A$, 7) <> "Gabino" Then
          If Left$(A$, 7) <> "E=====" Then
            If Left$(A$, 7) <> "E    " Then
              If Left$(A$, 7) <> "E     " Then
                If Left$(A$, 7) <> "FH" Then
                  If Left$(A$, 7) <> " SUCUR." Then
                    If Left$(A$, 17) <> "         CONTACTO" Then
                      If Left$(A$, 7) <> "FH" Then
                        If Left$(A$, 7) <> "FH" Then
                          If Left$(A$, 7) <> "  " Then
                            If Left$(A$, 7) <> "ARTICUL" Then
                              If Left(A$, 39) <> "                                       " Then
                                If Left$(A$, 70) <> "" Then
                                    Print #N2%, A$
                                  Else
                                     GoTo 10
                                End If
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
10   Loop
End Sub

' 'Private Sub CMDCONVERSION_Click()
'   'CONCATENA LOS REGISTROS DE LAS POSICIION DE IVA DE clientes
'
'   N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\TXTPRECONVERTIDOS\ListcliOrden1_10-01-08.TXT", 1, 128)
'   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\TXTPRECONVERTIDOS\ListcliOrden2_10-01-08.TXT", 2, 128)
'   CONT& = 0
'   Do While Not EOF(N1%)
'      Line Input #N1%, A$
'      CONT& = CONT& + 1
'      If CONT& = 1 Then
'        Call REPLA(TXT$, A$, 0, 50)
'        ElseIf CONT& = 2 Then
'          Call REPLA(TXT$, A$, 51, 101)
'          ElseIf CONT& = 3 Then
'            Call REPLA(TXT$, A$, 102, 152)
'            CONT& = 0
'            Print #N2%, TXT$
'      End If
'   Loop
'
'
'End Sub

'
Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
End Sub

Private Sub Combo1_Click()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Recordset.MoveFirst
   Data1.Refresh
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   CANCELIMP% = 0
   Select Case XVALO(NARIM)
     Case 0
       Call IMPORMAS
     Case 1
       Call IMPORSTRU
     Case 2
       Call IMPORUTS
     Case 3
       Call IMPORMAQ
     Case 4
       'Call IMPORCAT
     Case 5
       Call IMPORCLI
     Case 6
       Call IMPORPRO
     Case 7
       Call IMPORPLA
     Case 8
       Call IMPORNCLI
     Case 9
       Call IMPORCUECLI
     Case 10
       Call IMPORCUEPRO
     Case 11
       Call IMPORCARCHE
     Case 12
       Call IMPORMOVSTK
     Case 13
       Call IMPORLISPRE
     Case 14
       Call IMPORCOSREP
     Case 15
       Call IMPORTRANS
     Case 16
       Call IMPORSUCUR
     Case 17
       Call IMPORASIE
     Case 18
       Call IMPORPADEMP
     Case 19
       Call IMPORUBIC
   End Select
   Screen.MousePointer = 1
   If CANCELIMP% = 0 Then
     Call COMUNI("Se ha Completado la Importacion Definitiva")
   End If
   CANCELIMP% = 0
   Frame5.Visible = False
   DoEvents
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.filename)
   Data1.DatabaseName = Label1
   Combo1.Clear
   Combo1.TEXT = ""
   On Error Resume Next
   Data1.Refresh
   If Err Then
'GoTo 98
      ERNU$ = TRIM$(Str$(Err.Number))
      On Error GoTo 0
      If ERNU$ = "3125" Then GoTo 98    ' falta revisar - 25/01/2008
      MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") al Abrir la Hoja de Cálculo Especificada. Imposible Completar la Petición. Consulte."
      MsgBox MENSA$
      GoTo 99
   End If
   '
98 CATAB% = Data1.Database.TableDefs.Count
   If CATAB% < 1 Then GoTo 99
   For TTB% = 1 To CATAB%
     Combo1.AddItem Data1.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo1.TEXT = Data1.Database.TableDefs(TTB% - 1).Name
     End If
   Next TTB%
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Printer.Print "Registros no Procesables"
     Printer.Print " "
     For U& = 1 To List1.ListCount
       Printer.Print List1.List(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Frame5.Visible = True
   List1.Clear
   List4.Clear
   On Error Resume Next
   HScroll1.Value = 0
   VScroll2.Value = 0
   On Error GoTo 0
   '
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   Data1.Recordset.MoveFirst
   '
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   Select Case XVALO(NARIM)
     Case 0
       Call PREIMPORMAS
     Case 1
       Call PREIMPORSTRU
     Case 2
       Call PREIMPORUTS
     Case 3
       Call PREIMPORMAQ
     Case 4
       Call PREIMPORCAT
     Case 5
       Call PREIMPORCLI
     Case 6
       Call PREIMPORPRO
     Case 7
       Call PREIMPORPLA
     Case 8
       Call PREIMPORNORCLI
     Case 9
       Call PREIMPORCUECLI
     Case 10
       Call PREIMPORCUEPRO
     Case 11
       Call PREIMPORCARCHE
     Case 12
       Call PREIMPORMOVSTK
     Case 13
       Call PREIMPORLISPRE
     Case 14
       Call PREIMPORCOSREP
     Case 15
        Call PREIMPORTRANS
     Case 16
        Call PREIMPORSUCUR
     Case 17
        Call PREIMPORASIE
     Case 18
        Call PREIMPORPADEMP
     Case 19
        Call PREIMPORUBIC
     Case 20
        Call IMPORESPECIFI
        MostrarMensaje = False
  End Select
  If MostrarMensaje = True Then
     Call COMUNI("Se Ha Completado la Pre-Importación.")
  End If
   '
20 DBGrid1.Height = DBHANT
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub IMPORESPECIFI()
   If DBGrid1.Columns.Count < 2 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   Call ABRECONEXION
   ReemplazarSiempre% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(356)
        CODEX$ = ""
        On Error Resume Next
        CODEX$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        ESPEFI$ = .Recordset.Fields(1)   ' TRIM$(DBGrid1.TEXT)
        COALT1$ = ""
        On Error GoTo 0
        'Registra especificacion
        CONDI$ = "CODIGO='" & CODEX$ & "'"
        If CantRegistros("MASTER", CONDI$) < 1 Then GoTo 10
        If TRIM$(VALOBADA("MASTER", "ESPECIFICACION", CONDI$)) <> "" Then
           '
           If ReemplazarSiempre% <> 1 Then
              OPX% = COMALTER%("Codigo '" & CODEX$ & "'  ya tiene Especificaciones\¿Desea Sobreesribir?\\Conservar\Reemplazar\Reemplazar Siempre")
              If OPX% <= 1 Then GoTo 10                  ' Conservar
              If OPX% = 3 Then ReemplazarSiempre% = 1   ' Reemplazar Siempre
           End If
           '
           If OPX% <> 1 Then
              Call ACTUREGISTRO("MASTER", "ESPECIFICACION", "CODIGO='" & CODEX & "'", ESPEFI$, REG&)
           End If
          Else ' Si No tiene Especificacion se asigna esta
           Call ACTUREGISTRO("MASTER", "ESPECIFICACION", "CODIGO='" & CODEX & "'", ESPEFI$, REG&)
        End If
10 'Pasar al siguiente registro
        .Recordset.MoveNext
     Loop
   End With
   Call COMUNI("Importacion Compledata")
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub

Private Sub Command6_Click()
   For KKI% = 1 To Combo1.ListCount
     Combo1.TEXT = Combo1.List(KKI% - 1)
     Data1.RecordSource = Combo1.TEXT
     Data1.Recordset.MoveFirst
     Data1.Refresh
     '
     TEPRUE = ""
     URECORD& = Data1.Recordset.RecordCount
     DBHANT = DBGrid1.Height
     DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
     Data1.Recordset.MoveFirst
     '
     Call PREIMPORMEDI
'If KKI% > 6 Then GoTo 15
   Next KKI%
  
15 Call COMUNI("Se Ha Completado la Pre-Importación.")
   '
20 DBGrid1.Height = DBHANT
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Frame4.Visible = False
   Frame5.Visible = True
   Command4.Enabled = True
   '
End Sub

Sub PREIMPORMEDI()
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   HUBOAMAS% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(1024)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
5    CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODINT%(CODEX$) < 1 Then
       CODEX$ = Left$(CODEX$, 15)
       If CODINT%(CODEX$) < 1 Then
         TTXX1$ = "Codigo de Conjunto '" + TRIM$(CODEX$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, CODEX$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 5
         End If
       End If
     End If
     '
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 24 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto '" + CODEX$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código '" + CODEX$ + "' con Estructura en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 24)
     '
     COEFI = 1
     ROWAC = DBGrid1.Row
     For JJ& = DBGrid1.Row - 1 To 0 Step -1
        DBGrid1.Row = JJ&
        DBGrid1.Col = 1
        If DBGrid1.TEXT = CODEX$ Then
          DBGrid1.Col = 2
          COEFI = XVALO(DBGrid1.TEXT)
          If COEFI < 1 Then COEFI = 1
          GoTo 125
        End If
     Next JJ&
     '
125  DBGrid1.Row = ROWAC
     '
     DBGrid1.Col = 1
13   COELE$ = Mid$(TRIM$(DBGrid1.TEXT), 1, 24)
     If CODINT%(COELE$) < 1 Then
       COELE$ = Left$(COELE$, 15)
       If CODINT%(COELE) < 1 Then
         TTXX1$ = "Codigo de Elemento '" + TRIM$(COELE$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, COELE$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 13
         End If
       End If
     End If
     '
     If COELE$ = "" Or Len(COELE$) < 1 Or Len(COELE$) > 24 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento '" + COELE$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 27, 24)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 53, 32)
     '
     TTXX1$ = Left$(TTXX$, 52)
     For JJ1& = 1 To List4.ListCount
        If Left$(List4.List(JJ1& - 1), 52) = TTXX1$ Then GoTo 19 ' YA ESTABA
     Next JJ1&
     '
     DBGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT) / COEFI
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 85, 12)
     '
     DBGrid1.Col = 3                     ' UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 98, 12)
     '
     DBGrid1.Col = 4                     ' LARGO DE CORTE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.3")
     Call REPLA(TTXX$, VALO$, 111, 12)
     '
     DBGrid1.Col = 5                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 124, 4)
     '
     DBGrid1.Col = 6                     ' REFERENCIA
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 131, 890)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   DBGrid1.RowHeight = hrow
   '
End Sub
'
Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Option1(0).Value = True
   '
   Picture5.Top = 0
   VScroll1.Min = 0
   VScroll1.Max = (Picture5.Height - Picture6.Height) / 10
   'SI ES DAKOT O PRUEBA EL BOTON DE CONVERSION ES VISIBLE
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DAKOT") > 0 Then 'Or InStr(EPAC$, "PRUEBA") > 0 Then
      CMDCONVERSION.Visible = True
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub



Private Sub Option1_Click(Index As Integer)
   List1.Clear
   List2.Clear
   List4.Clear
   Frame5.Visible = False
   Label1 = ""
   Combo1.Clear
   DBGrid1.ClearFields
   NARIM = TRIM$(Str$(Index))
   Select Case Index
   Case 0
     '\\\OT-05-0499-WAR-13/07/05///
     List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
     List2.AddItem "  2  Codigo Alternativo         String   <= 24 car"
     List2.AddItem "  3  Descripcion del Articulo   String   <= 64 car"
     List2.AddItem "  4  Desc. Complementaria       String   <= 64 car"
     List2.AddItem "  5  Atributo A                 String   <= 32 car"
     List2.AddItem "  6  Atributo B                 String   <= 32 car"
     '\\\OT-05-0499-WAR-FIN///
     List2.AddItem "  7  Tipo de Material           Num      0 - 255"
     List2.AddItem "  8  Unidad de Medida           String   <= 2 car"
     List2.AddItem "  9  Stock Minimo               Num      0 - 10^8"
     List2.AddItem " 10  Lote Economico             Num      0 - 10^8"
     List2.AddItem " 11  Kilos por Unidad           Num      0 - 10^8"
     List2.AddItem " 12  Kilos por Metro            Num      0 - 10^8"
     List2.AddItem " 13  Costo Unitario             Num      0 - 10^8"
     List2.AddItem " 14  Moneda de Costeo           Num      1 - 255"
     List2.AddItem " 15  Fecha de Costeo            Date"
     List2.AddItem " 16  Numero Proveedor Habitual  Num      0 - 32767"
     List2.AddItem " 17  Grupo Contable             Num      0 - 32767"
     List2.AddItem " 18  Stock Actual               Num      +/- 999999"
     List2.AddItem " 19  C. Interna                 String   <= 64 car"
     
     ENCABE$ = "Código           C-Int  Cod.Alternativo          Descripcion                                                      Descripción Complementaria                                       Atributo A                       Atributo B                         TM UM         St.Min      Lote Ec Kilos/Un  Kilos/Mt    Costo Unit   MC Fecha        Pr-H     GC       Stock"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
   Case 1
     List2.AddItem "  1  Codigo de Conjunto         String   <= 15 car"
     List2.AddItem "  2  Codigo del Componente      String   <= 15 car"
     List2.AddItem "  3  Uso / Cantidad Componente  Num      10^-5 - 10^6"
     List2.AddItem "  4  Unidades                   Num      10^-5 - 10^6"
     List2.AddItem "  5  Largo de corte (mm)        Num      10^-5 - 10^6"
     List2.AddItem "  6  Plazo Disposicion          Num      1 - 255"
     List2.AddItem "  7  Referencia                 String"
     Label8 = "Código Conj.              Código Elem.              Descripcion                         Uso/Cant     Unidades        Corte   PD   Referencia"
     Label82 = Space(300)
   Case 2
     List2.AddItem "  1  Codigo de Producto         String   <= 15 car"
     List2.AddItem "  2  Descripcion de Operacion   String   <= 48 car"
     List2.AddItem "  3  Codigo de Maquina          String   <= 6 car"
     List2.AddItem "  4  Horas Preparacion (APP)    Num      0 - 999,99"
     List2.AddItem "  5  Horas Fijo por Lote        Num      0 - 999,99"
     List2.AddItem "  6  Minutos Variable / Pza     Num      0 - 99,99999"
     List2.AddItem "  7  % Demora Distribuida       Num      0 - 99,9"
     List2.AddItem "  8  Lote Standard              Num      0 - 10^6"
     List2.AddItem "  9  Operarios                  Num      0,01 - 99,99"
     List2.AddItem " 10  Maquina Adicional 1        String   <= 6 car"
     List2.AddItem " 11  Maquina Adicional 2        String   <= 6 car"
     List2.AddItem " 12  Maquina Adicional 3        String   <= 6 car"
     List2.AddItem " 13  Programa / Hoja Proc.Fab.  String   <= 8 car"
     Label8 = "Código          N-Op  Descripcion de Operacion                          Maquina   Hs.APP    Hs.Fix  Min.Variab  Lote Std    % DD   Opers  MAux-1  MAux-2  MAux-3  No.HPF"
   Case 3
     List2.AddItem "  1  Codigo de Maquina          String   <= 6 car"
     List2.AddItem "  2  Descripcion                String   <= 30 car"
     List2.AddItem "  3  Cuota Horaria $ / hora     Num      0 - 10^6"
     List2.AddItem "  4  Codigo de Seccion          Num      100 - 999"
     List2.AddItem "  5  Maquina Equivalente        String   <= 6 car"
     List2.AddItem "  6  Coeficiente (%)            Num      0 - 999,99"
     List2.AddItem "  7  Observaciones              String   <= 30 car"
     Label8 = "Código  Descripcion                         $ / Hora   Sección   Alter.        Coef.  Observaciones"
   Case 4
     List2.AddItem "  1  Descripcion Categoria      String   <= 27 car"
     List2.AddItem "  2  Costo M.O.D. / hora        Num      0 - 10^6"
     List2.AddItem "  3  Observaciones              String   <= 32 car"
     Label8 = " N-O  Descripcion                       $ / Hora  Observaciones"
   Case 5
     List2.AddItem "  1  Numero Cliente             Num      1 - 9999"
     List2.AddItem "  2  Razón Social del Cliente   String   <= 62 car"
     List2.AddItem "  3  Domicilio                  String   <= 64 car"
     List2.AddItem "  4  Código Postal              String   <= 16 car"
     List2.AddItem "  5  Localidad                  String   <= 48 car"
     List2.AddItem "  6  Pais                       String   <= 64 car"
     List2.AddItem "  7  Teléfono                   String   <= 24 car"
     List2.AddItem "  8  Fax                        String   <= 24 car"
     List2.AddItem "  9  Persona de Contacto        String   <= 32 car"
     List2.AddItem " 10  CUIT                       String   <= 24 car"
     List2.AddItem " 11  Grupo de Cliente           Num      0 - 255"
     List2.AddItem " 12  Codigo de Cuenta           String   <= 12 car"
     List2.AddItem " 13  Número de Vendedor         Num      0 - 255"
     List2.AddItem " 14  Codigo de Provincia        String   <= 1 car"
     List2.AddItem " 15  Categoría I.V.A.           Num      0 - 7"
     List2.AddItem " 16  Condicion de Pago          Num      1 - 255"
     List2.AddItem " 17  Lista de Precios           Num      1 - 255"
     List2.AddItem " 18  Límite de Credito          Num      0 - 10^9"
     List2.AddItem " 19  Descuento Standard         String   <= 24 car"
     List2.AddItem " 20  L.Entrega - Calle + Nro    String   <= 48 car"
     List2.AddItem " 21  L.Entrega - Localidad      String   <= 48 car"
     List2.AddItem " 22  Numero de Transporte       Num      0 - 9999"
     List2.AddItem " 23  Transportista              String   <= 40 car"
     List2.AddItem " 24  Grupo Divisional           Num      0 - 9999"
     List2.AddItem " 25  Percepcion I/Brutos        Num      0 / 1"
     List2.AddItem " 26  E-mail                     String   <= 36 car"
     List2.AddItem " 27  WebSite                    String   <= 128 car"
     List2.AddItem " 28  % Recargo Embalaje         Num      0 - 99.9"
     '
     ENCABE$ = " Cta. Razón Social                                    Domicilio                                       " & _
     "C.Postal    Localidad                       Teléfono                Fax                     Contacto                CUIT                     Gp.  Cod.Cta       V.  Pr  PI  CP     Lim.Cre  " & _
     "Descuento                 Lugar de Entrega                                Localidad                                       N-Tra   Transportista                            Gp.Div.  " & _
     "IB  e-mail                               R-Emb"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257, 256)
     Label83 = Mid$(ENCABE$, 513)
     Label82.Left = Label8.Left + Label8.Width
     Label83.Left = Label82.Left + Label82.Width
   Case 6
     List2.AddItem "  1  Numero Proveedor           Num      1 - 9999"
     List2.AddItem "  2  Razón Social del Proveedor String   <= 62 car"
     List2.AddItem "  3  Domicilio                  String   <= 64 car"
     List2.AddItem "  4  Código Postal              String   <= 16 car"
     List2.AddItem "  5  Localidad                  String   <= 48 car"
     List2.AddItem "  6  Teléfono                   String   <= 24 car"
     List2.AddItem "  7  Fax                        String   <= 24 car"
     List2.AddItem "  8  Persona de Contacto        String   <= 32 car"
     List2.AddItem "  9  CUIT                       String   =  8 car"
     List2.AddItem " 10  Tipo de Cuenta             Num      0 - 1"
     List2.AddItem " 11  Codigo de Cuenta           String   <= 16 car"
     '\\\OT-05-0202-WAR-19/04/05///
     List2.AddItem " 12  Codigo de Provincia        String   <=1 car"
     List2.AddItem " 13  Categoria de IVA           Num      0 - 7"
     List2.AddItem " 14  Cuenta Madre               String   <= 12 car"
     List2.AddItem " 15  Nombre de Fantasía         String   <= 12 car"
     List2.AddItem " 16  Retencion I/Brutos         Num      0 - 255"
     ENCABE$ = "Cta.  Razón Social                                    Domicilio                                       C.Postal    Localidad                       " + _
     "Teléfono                Fax                     Contacto                CUIT            T.   Cod.Cta.    Pr. IVA  Cta. Madre    N.Fantasia          IB"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
     '\\\OT-05-0202-WAR-FIN///
   Case 7
     List2.AddItem "  1  Nivel de Agrupación        Num      0 - 5"
     List2.AddItem "  2  Codigo de Cuenta           String   <= 12 car"
     List2.AddItem "  3  Denominación de Cuenta     String   <= 64 car"
     List2.AddItem "  4  Tipo de Cuenta             String   D/H/I/G"
     List2.AddItem "  5  Imputable                  String   Si / No"
     Label8 = "  Nvl Codigo        Descripcion                                       T.  ID"
     '\\\OT-05-0207-WAR-21/04/05///
   Case 8
     List2.AddItem "  1  Cod. Cliente               String   0 - 12"
     List2.AddItem "  2  Codigo de Articulo         String   <= 16 car"
     List2.AddItem "  3  Nro. Original              String   <= 28 car"
     Label8 = "Cliente      Articulo          Descripcion             Nro.Original"
     '\\\OT-05-0207-WAR-FIN///
     '\\\OT-05-0272-WAR-09/05/05///
   Case 9
     List2.AddItem "  1  Cod. Cliente               String   0 - 12"
     List2.AddItem "  2  Referencia                 String   <= 62 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Tipo Comprobante (FC-RB)   String   = 2 car"
     List2.AddItem "  6  Letra Comprobante (A-B-X)  String   = 1 car"
     List2.AddItem "  7  Prefijo / Talon. (nnnn)    Num      0 - 9999"
     List2.AddItem "  8  Nro. Princ. Comprobante    Num      1 - 99999999"
     List2.AddItem "  9  Importe Original           Num      0 - 10^8"
     List2.AddItem " 10  Importe Pendiente          Num      0 - 10^8"
     Label8 = "Cliente             Referencia                      F.Emisión F.Vencim.   Comp.Largo          Comp.Corto       Imp.Orig.    Imp.Pend. "
   Case 10
     List2.AddItem "  1  Cod. Prov.                 String   <= 12 car"
     List2.AddItem "  2  Referencia                 String   <= 62 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Tipo Comprobante (FC-RB)   String   = 2 car"
     List2.AddItem "  6  Letra Comprobante (A-B-X)  String   = 1 car"
     List2.AddItem "  7  Prefijo / Talon. (nnnn)    Num      0 - 9999"
     List2.AddItem "  8  Nro. Princ. Comprobante    Num      1 - 99999999"
     List2.AddItem "  9  Importe Original           Num      0 - 10^8"
     List2.AddItem " 10  Importe Pendiente          Num      0 - 10^8"
     Label8 = "Proveedor           Referencia                      F.Emisión F.Vencim.   Comp.Largo          Comp.Corto       Imp.Orig.    Imp.Pend. "
     '\\\OT-05-0272-WAR-FIN///
   Case 11
    '\\\OT-05-0429-WAR-28/06/05///
     List2.AddItem "  1  Banco Emisor               String   <= 18 car"
     List2.AddItem "  2  Nro. Cheque                String   <= 12 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Importe Original           Num      0 - 10^9"
     List2.AddItem "  6  Documento Recep. Largo     String   <= 18 car"
     List2.AddItem "  7  Documento Recep. Corto     String   <= 12 car"
     List2.AddItem "  8  Cliente / Origen           String   <= 30 car"
     List2.AddItem "  9  CUIT Librador              String   <= 16 car"
     Label8 = "Banco               Nro. Cheque   F.Emis.   F.Venc.         Importe   Recibo Largo        Recibo Corto  Cliente / Origen                CUIT Librador"
     Label82 = ""
    '\\\OT-05-0429-WAR-FIN///
   Case 12
     List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
     List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
     List2.AddItem "  3  Fecha de Movimiento        Date"
     List2.AddItem "  4  Depósito Número            Num      1 - 255"
     List2.AddItem "  5  Cantidad (+/-)             Num      0 - 10^8"
     List2.AddItem "  6  Codigo de Movimiento       String   <= 2 car"
     List2.AddItem "  7  Documento Primario         String   <= 12 car"
     List2.AddItem "  8  Documento Secundario       String   <= 12 car"
     List2.AddItem "  9  Valor Unitario             Num      0 - 10^8"
     List2.AddItem " 10  Moneda de Valuacion        String   <= 2 car"
     List2.AddItem " 11  Código de Cliente          String   <= 12 car"
     List2.AddItem " 12  Código de Proveedor        String   <= 12 car"
     List2.AddItem " 13  Detalle / Referencia       String   <= 32 car"
     Label8 = "Código           C-Int  Descripción                                                       Fecha      Dep        Cant. CM  D.Prim.       D.Sec.              V.Unit.  M.  Cli.  Pro. Detalle / Referencia"
   Case 13
     '\\\OT-5-0641-FG-26/09/05///
     List2.AddItem "  1  Número de Lista            Nume     1 - 255"
     List2.AddItem "  2  Código de Artículo         String   <= 15 car"
     List2.AddItem "  3  Número Original            String   <= 16 car"
     List2.AddItem "  4  Precio Unitario            Num      0 - 10^8"
     Label8 = "Nro. Código            Descripción                                                       Nro.Original        P.Unitario"
     '\\\OT-5-0641-FG-FIN///
   Case 14
     '\\\OT-06-0186-WAR-26/05/06///
     List2.AddItem "  1  Código de Artículo         String   <= 15 car"
     List2.AddItem "  2  Moneda (Número s/Tabla)    Nume     1 - 255"
     List2.AddItem "  3  Importe / Costo Unitario   Num      0 - 10^8"
     List2.AddItem "  4  Fecha Costo de Reposición  Date"
     List2.AddItem "  5  Código de Proveedor        String   <= 12 car"
     Label8 = "Código            Descripción                                        M.        Importe   Fecha    Prov."
     '\\\OT-06-0186-WAR-FIN///
   Case 15
     '\\\OT-06-0423-MH-05/10/06///
     List2.AddItem "  1  Código del Transportista   Num      1 - 9999"
     List2.AddItem "  2  Nombre                     String   <= 32 car"
     List2.AddItem "  3  Razón Social               String   <= 48 car"
     List2.AddItem "  4  Domicilio                  String   <= 48 car"
     List2.AddItem "  5  Localidad                  String   <= 48 car"
     List2.AddItem "  6  Pos. Iva                   Num      0 - 7    "
     List2.AddItem "  7  Cuit                       String   <= 15 car"
     List2.AddItem "  8  Teléfono                   String   <= 24 car"
     List2.AddItem "  9  Fax                        String   <= 24 car"
     List2.AddItem "  10 E-Mail                     String   <= 48 car"
     List2.AddItem "  11 Persona de Contacto        String   <= 32 car"
     '
     ENCABE$ = "Cód.  Nombre                             R.Social                                           "
     ENCABE$ = ENCABE$ + "Domicilio                                          Localidad                                          Pos.Iva   Cuit           Teléfono                   Fax                        E-Mail                                          Contacto"
     '\\\OT-06-0423-MH-FIN///
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257, 256)
     Label83 = Mid$(ENCABE$, 513)
     Label82.Left = Label8.Left + Label8.Width
     Label83.Left = Label82.Left + Label82.Width
     '
   Case 16
     '
     List2.AddItem "  1  Numero Sucursal            Num      <= 1 - 255"
     List2.AddItem "  2  Nombre                     String   <= 64 car"
     List2.AddItem "  3  Domicilio                  String   <= 64 car"
     List2.AddItem "  4  Localidad                  String   <= 64 car"
     List2.AddItem "  5  Numero Clliente            Num      <= 1 - 9999"
     
     ENCABE$ = "Cód. Nombre                            Domicilio                      Localidad                           Num. Cliente "
     '\\\OT-06-0423-MH-FIN///
     Label8 = Left$(ENCABE$, 256)
   Case 17
     List2.AddItem "  1  Id Asiento                 Num      0 - 32767"
     List2.AddItem "  2  Fecha                      Date"
     List2.AddItem "  3  Referencia General         String   <= 64 car"
     List2.AddItem "  4  Cuenta                     String   <= 16 car"
     List2.AddItem "  5  Referencia del Pase        String   <= 64 car"
     List2.AddItem "  6  Importe Debe               Num      0 - 10^8"
     List2.AddItem "  7  Importe Haber              Num      0 - 10^8"
     Label8 = "Id Asiento     Fecha           Referencia General                                 Cuenta           Referencia del Pase                            Importe Debe   Importe Haber        "
     'Label8 = Left$(ENCABE$, 256)
   Case 18
     List2.AddItem "  1  Legajo                     Num      1 - 9999"
     List2.AddItem "  2  Nombre y Apellido          String   <= 30 car"
     List2.AddItem "  3  Grupo Empleado             Num      1 - 255"
     List2.AddItem "  4  Documento                  String   <= 13 car"
     List2.AddItem "  5  Sexo                       String   M / F"
     List2.AddItem "  6  Fecha Nacimiento           Date"
     List2.AddItem "  7  Nacionalidad               String   <= 10 car"
     List2.AddItem "  8  Estado Civil               Num      1 - 255"
     List2.AddItem "  9  Domicilio                  String   <= 30 car"
     List2.AddItem " 10  Codigo Postal              String   <= 10 car"
     List2.AddItem " 11  Localidad                  String   <= 30 car"
     List2.AddItem " 12  Sector                     String   <= 4 car"
     List2.AddItem " 13  Horario de Trabajo         String   <= 15 car"
     List2.AddItem " 14  Tarea                      String   <= 20 car"
     List2.AddItem " 15  Categoria                  String   <= 9 car"
     List2.AddItem " 16  Fecha de Ingreso           Date"
     List2.AddItem " 17  Fecha de Antiguedad        Date"
     List2.AddItem " 18  CUIL                       String   <= 13 car"
     List2.AddItem " 19  AFJP                       String   <= 10 car"
     List2.AddItem " 20  Fecha Egreso               Date"
     List2.AddItem " 21  Motivo                     String   <= 16 car"
     List2.AddItem " 22  Status                     Num      0 / 1"
     ENCABE$ = "Legajo   Nombre y Apellido              G.Emp  Documento  Sexo  F.Nac.  Nacionalidad  E.Civil  Domicilio                     C.Postal  Localidad                      Sector  Horario Trabajo  Tarea                 Categoria  F.Ingreso  Antiguedad  CUIL           AFJP        F.Egreso  Motivo            Status"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
   Case 19
     List2.AddItem "  1  Código                     String   <= 16 car"
     List2.AddItem "  2  Descripción del Artículo   String   <= 64 car"
     List2.AddItem "  3  Ubicación                  String   <= 24 car"
     ENCABE$ = ""
     ENCABE$ = "Código             Descripción                       Ubicación "
     Label8 = Left$(ENCABE$, 80)
     Label82 = ""
   Case 20
     List2.AddItem "  1  Código                     String   <= 16 car"
     List2.AddItem "  2  Especificacion             Memo"
     ENCABE$ = ""
     ENCABE$ = "Código             Especificacion"
     Label8 = Left$(ENCABE$, 80)
     Label82 = ""
   End Select
   
End Sub



Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Sub PREIMPORMAS()
   '
   If DBGrid1.Columns.Count < 17 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(356)
        CODEX$ = ""
        On Error Resume Next
        CODEX$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        COALT1$ = ""
        On Error GoTo 0
        If CODEX$ = "" Or Len(CODEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        ElseIf Len(CODEX$) > 15 Then ' TRUNCA CODIGOS CON MAS DE 15 POSICIONES
          COALT1$ = Left$(CODEX$, 24)
          For KKI% = 1 To 9999
            CDXYX$ = TRIM$(Str$(KKI%))
            While Len(CDXYX$) < 4: CDXYX$ = "0" + CDXYX$: Wend
            CDXYX$ = Left$(CODEX$, 1) + "-" + CDXYX$
            If CODINT%(CDXYX$) < 1 Then
              For KKJ% = 1 To List4.ListCount
                If TRIM$(Left$(List4.List(KKJ% - 1), 15)) = CDXYX$ Then GoTo 10
              Next KKJ%
              CODEX$ = CDXYX$
              GoTo 11
            End If
10        Next KKI%
          MERROX$ = "Codigo " & "'" & codex1$ & "'" & " Excede 15 Posiciones"
          GoTo 19
        End If
        '
11      CDXI% = CODINT%(CODEX$)
        If CDXI% > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      Call REPLA(TTXX$, CODEX$, 1, 16)
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        COALT$ = ""
        On Error Resume Next
        COALT$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
        If COALT$ = "" Then COALT$ = COALT1$
'        If Len(COALT$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Código Alternativo"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, COALT$, 25, 24)
        '\\\OT-05-0499-WAR-FIN///
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(2))
        On Error GoTo 0
        If Len(DEDEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DEDEX$, 50, 64)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        DECOM$ = ""
        On Error Resume Next
        DECOM$ = TRIM$(.Recordset.Fields(3))
        On Error GoTo 0
        'If Len(DECOM$) < 1 Then
        '  MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion Complementaria"
        '  List1.AddItem MERROX$
        '  GoTo 19
        'End If
        Call REPLA(TTXX$, DECOM$, 115, 64)
        '
        ATRA$ = ""
        On Error Resume Next
        ATRA$ = TRIM$(.Recordset.Fields(4))
        On Error GoTo 0
'        If Len(ATRA$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo A"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, ATRA$, 180, 32)
        '
        ATRB$ = ""
        On Error Resume Next
        ATRB$ = TRIM$(.Recordset.Fields(5))
        On Error GoTo 0
        'If Len(ATRB$) < 1 Then
        '  MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo B"
        '  List1.AddItem MERROX$
        '  GoTo 19
        'End If
        Call REPLA(TTXX$, ATRB$, 213, 32)
        '
        '\\\OT-05-0499-WAR-FIN///
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(6), 4)
        On Error GoTo 0
        If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 246, 4)
        '
        UME$ = "U."
        On Error Resume Next
        UME$ = TRIM$(.Recordset.Fields(7))
        On Error GoTo 0
        Call REPLA(TTXX$, UME$, 251, 4)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(8))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 256, 12)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(9))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 269, 12)
        '
        '\\\OT-5-0633-FG-22/09/05///
        VALOR = 0                           ' KILOS POR UNIDAD
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(10))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F8.3")
        Call REPLA(TTXX$, VALO$, 282, 8)
        '
        VALOR = 0                           ' KILOS POR METRO
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(11))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F8.3")
        Call REPLA(TTXX$, VALO$, 292, 8)
        '\\\OT-5-0633-FG-FIN///
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(12))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.4")
        Call REPLA(TTXX$, VALO$, 302, 12)
        '
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(13), 4)
        On Error GoTo 0
         If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 315, 4)
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = Format$(.Recordset.Fields(14), "dd/mm/yy")
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 320, 8)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(15))
        On Error GoTo 0
        If VALOR < 0 Or VALOR > 32766 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F8.0")
        Call REPLA(TTXX$, VALO$, 329, 8)
        '
        VALOR = 0                           ' GRUPO DE VENTA
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(16))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F6.0")
        Call REPLA(TTXX$, VALO$, 338, 6)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(17))
        On Error GoTo 0
        If VALOR < (-9999999) Or VALOR > 9999999 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F10.1")
        Call REPLA(TTXX$, VALO$, 346, 10)
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = SAFETEXT$(.Recordset.Fields(18))
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 360, 64)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORSTRU()
   '
   If DBGrid1.Columns.Count < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   HUBOAMAS% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(1024)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
5    CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODINT%(CODEX$) < 1 Then
       CODEX$ = Left$(CODEX$, 15)
       If CODINT%(CODEX$) < 1 Then
         TTXX1$ = "Codigo de Conjunto '" + TRIM$(CODEX$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, CODEX$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           'Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 5
         End If
       End If
     End If
     '
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 24 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto '" + CODEX$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código '" + CODEX$ + "' con Estructura en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 24)
     '
     DBGrid1.Col = 1
13   COELE$ = Mid$(TRIM$(DBGrid1.TEXT), 1, 24)
     If CODINT%(COELE$) < 1 Then
       COELE$ = Left$(COELE$, 15)
       If CODINT%(COELE) < 1 Then
         TTXX1$ = "Codigo de Elemento '" + TRIM$(COELE$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, COELE$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 13
         End If
       End If
     End If
     '
     If COELE$ = "" Or Len(COELE$) < 1 Or Len(COELE$) > 24 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento '" + COELE$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 27, 24)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 53, 32)
     '
     DBGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 85, 12)
     '
     DBGrid1.Col = 3                     ' UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 98, 12)
     '
     DBGrid1.Col = 4                     ' LARGO DE CORTE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.3")
     Call REPLA(TTXX$, VALO$, 111, 12)
     '
     DBGrid1.Col = 5                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 124, 4)
     '
     DBGrid1.Col = 6                     ' REFERENCIA
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 131, 890)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   DBGrid1.RowHeight = hrow
   '
   If HUBOAMAS% > 0 Then    ' HUBO CODIGOS AGREGADOS
     UPDATMASTER% = 1
     Call GENECOMAS
   End If
   '
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORUTS()
   '
   If DBGrid1.Columns.Count < 12 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     If CODEX$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = CODEX$
     '
     DBGrid1.Col = 1
     DEOPE$ = TRIM$(DBGrid1.TEXT)
     If DEOPE$ = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     NOPE% = NOPE% + 10
     VALO$ = FORMATNUM$(NOPE%, "I4")
     Call REPLA(TTXX$, VALO$, 17, 4)
     Call REPLA(TTXX$, DEOPE$, 23, 48)
     '
     DBGrid1.Col = 2                     ' MAQUINA
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 73, 6)
     '
     DBGrid1.Col = 3                     ' HORAS PREPARACION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 79, 10)
     '
     DBGrid1.Col = 4                     ' HORAS STANDARD FIJO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 89, 10)
     '
     DBGrid1.Col = 5                     ' MINUTOS STANDARD VARIABLE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.5")
     Call REPLA(TTXX$, VALO$, 99, 12)
     '
     DBGrid1.Col = 6                     ' LOTE STANDARD UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 1 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F10.1")
     Call REPLA(TTXX$, VALO$, 111, 10)
     '
     DBGrid1.Col = 7                     ' % DEM.DISTRIBUIDA
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     If VALOR < 0 Then VALOR = 0
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F8.1")
     Call REPLA(TTXX$, VALO$, 121, 8)
     '
     DBGrid1.Col = 8                     ' OPERARIOS
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0.01 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F8.2")
     Call REPLA(TTXX$, VALO$, 129, 8)
     '
     DBGrid1.Col = 9                     ' MAQUINA AUX 1
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 139, 6)
     '
     DBGrid1.Col = 10                     ' MAQUINA AUX 2
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 147, 6)
     '
     DBGrid1.Col = 11                     ' MAQUINA AUX 3
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 155, 6)
     '
     DBGrid1.Col = 12                     ' H.P.FABR.
     HOPROFA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, HOPROFA$, 163, 8)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 9, 30)
     '
     DBGrid1.Col = 2                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 999999.99 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 41, 12)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(DBGrid1.TEXT)
     If VTX$ <> "" And XVALO(VTX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(VTX$)
     On Error GoTo 0
     If VALOR > 999.999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 53, 10)
     '
     DBGrid1.Col = 4                    ' Maquina equivalente
     CODEQ$ = TRIM$(DBGrid1.TEXT)
     If CODEQ$ <> "" Then
       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CODEQ$, 66, 6)
     '
     DBGrid1.Col = 5                     ' Coeficiente de equivalencia
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
     If CODEQ$ = "" Then VALOR = 0
     If VALOR < VAMIN Or VALOR > 999.9 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     Call REPLA(TTXX$, VALO$, 72, 12)
     '
     DBGrid1.Col = 6                    ' Observaciones
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   
20 End Sub

Sub PREIMPORCAT()
   '
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   If URECORD& > 255 Then
     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
     GoTo 20
   End If
   '
   If ULTREG&("CATOPER") > 0 Then
     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
       GoTo 20
     End If
   End If
   '
   JJ& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     DECAT$ = TRIM$(DBGrid1.TEXT)
     If DECAT$ <> "" Then
       JJ& = JJ& + 1
       COCAT$ = AJUSTD$(Str$(JJ&), 4)
       Call REPLA(TTXX$, COCAT$, 1, 4)
       Call REPLA(TTXX$, DECAT$, 7, 30)
       '
       DBGrid1.Col = 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       If VALOR < 0 Or VALOR > 999999.99 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
         List1.AddItem MERROX$
         GoTo 19
       End If
       VALO$ = FORMATNUM$(VALOR, "F12.4")
       Call REPLA(TTXX$, VALO$, 37, 12)
       '
       DBGrid1.Col = 2
       OBSER$ = TRIM$(DBGrid1.TEXT)
       Call REPLA(TTXX$, OBSER$, 51, 32)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&
   
20 End Sub

Sub PREIMPORCLI()
   '
   PUNUCLI$ = String$(32767, 0)
   '
   If DBGrid1.Columns.Count < 21 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Screen.MousePointer = 11
   ' Establece en el "vector" PUNUCLI$, los numeros de clientes ya utilizados
   Call APERBASDAT("CLIEN")
   'With Deudor12
   SQLX$ = "SELECT * FROM DEUDOR12"
   Dim DEUDDDOR12 As ADODB.Recordset
   Set DEUDDDOR12 = New ADODB.Recordset
   DEUDDDOR12.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   With DEUDDDOR12
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         NCX& = !Nume_Cli
         If NCX& > 0 Then
           If NCX& <= 999999 Then
             While Len(PUNUCLI$) < NCX&: PUNUCLI$ = PUNUCLI$ + Chr$(0): Wend
             Mid$(PUNUCLI$, NCX&, 1) = Chr$(1)
           End If
         End If
       .MoveNext
       Loop
     End If
   End With
   DEUDDDOR12.Close
   Set DEUDDDOR12 = Nothing
   '
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(516)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0
     NUCLIE = XVALO(TRIM$(DBGrid1.TEXT))
     If NUCLIE < 1 Or NUCLIE > 999999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cliente no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     ' Verifica si el numero ya ha sido utilizado
     While Len(PUNUCLI$) < NUCLIE: PUNUCLI$ = PUNUCLI$ + Chr$(0): Wend
     If Asc(Mid$(PUNUCLI$, NUCLIE, 1)) > 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cliente Repetido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Mid$(PUNUCLI$, NUCLIE, 1) = Chr$(1)
     '
     CODEX$ = AJUSTD$(Str$(Int(NUCLIE + 0.5)), 6)
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social de Cliente"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 8, 64)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 48)
     Call REPLA(TTXX$, DOMI$, 55, 48)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 12)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 32)
     '
     DBGrid1.Col = 6                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 7                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 8                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 9                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 24)
     '
     DBGrid1.Col = 10                     ' GRUPO DE CLIENTE
     GRUPOCLI$ = AJUSTI$(DBGrid1.TEXT, 4)
     Call REPLA(TTXX$, GRUPOCLI$, 244, 4)
     '
     DBGrid1.Col = 11                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODICUEN$, 249, 12)
     '
     DBGrid1.Col = 12                    ' VENDEDOR
     VENDE$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, VENDE$, 261, 4)
     '
     DBGrid1.Col = 13                    ' PROVINCIA
     PROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, PROVI$, 267, 2)
     '
     DBGrid1.Col = 14                    ' CATEGORIA IVA
     CIVA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, CIVA$, 269, 4)
     '
     DBGrid1.Col = 15                    ' CONDICION DE PAGO
     COPA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, COPA$, 273, 4)
     '
     DBGrid1.Col = 17                    ' LIMITE DE CREDITO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Then VALOR = 0
     VALO$ = FORMATNUM$(VALOR, "F12.2")
     Call REPLA(TTXX$, VALO$, 277, 12)
     '
     '\\\OT-5-632-FG-21/09/05///
     '
     DBGrid1.Col = 18                     ' DESCUENTO STANDARD
     VALO$ = ""
     On Error Resume Next
     VALO$ = REPLACAR$(DBGrid1.TEXT, ",", ".")
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 291, 24)
     '
     DBGrid1.Col = 19                    ' LUGAR DE ENTREGA
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 317, 48)
     '
     DBGrid1.Col = 20                    ' LOCALIDAD
     VALO$ = ""
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 365, 48)
     '
     DBGrid1.Col = 21                    ' numero de transporte
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F5.0")
     Call REPLA(TTXX$, VALO$, 413, 5)
     '
     DBGrid1.Col = 22                    ' TRANSPORTISTA
     VALO$ = ""
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 421, 40)
     '
     DBGrid1.Col = 23                    ' GRUPO DIVISIONAL
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F8.0")
     Call REPLA(TTXX$, VALO$, 461, 8)
     '
     DBGrid1.Col = 24                    ' PERCEPCION DE INGRESOS BRUTOS
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F4.0")
     Call REPLA(TTXX$, VALO$, 469, 4)
     '
     DBGrid1.Col = 25                    ' E-MAIL
     VALO$ = ""
     On Error Resume Next
     VALO$ = DBGrid1.TEXT
     On Error GoTo 0
     If Len(VALO$) > 36 Then
       Call MENSERR(24, "Imposible Importar E-mail Completo de\'" + DEDEX$ + "'.\  \Oportunamente Debe Revisar Este Dato.")
     End If
     Call REPLA(TTXX$, VALO$, 475, 36)
     '
     DBGrid1.Col = 27                    ' Porcentaje Recargo Embalaje
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F6.1")
     Call REPLA(TTXX$, VALO$, 511, 6)
     '
     '\\\OT-5-632-FG-FIN///
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
End Sub

Sub PREIMPORPRO()
   '
   PUNUCLI$ = String$(32767, 0)
   Dim CUTTX$(10000)
   '
   If DBGrid1.Columns.Count < 16 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   'MODIF ECHA POR UNA IMPORTACIÓN ECHA EN TELEBIT 16/06/05
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Call APERBASDAT("PROVE")
   SQLX$ = "SELECT * FROM PROVED12"
   Dim PROVVVED As ADODB.Recordset
   Set PROVVVED = New ADODB.Recordset
   PROVVVED.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With PROVVVED
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         NCX& = !Nume_Cli
         If NCX& > 0 Then
           If NCX& <= 32767 Then
             Mid$(PUNUCLI$, NCX&, 1) = Chr$(1)
           End If
         End If
       .MoveNext
       Loop
     End If
   End With
   PROVVVED.Close
   Set PROVVVED = Nothing
   
   Call ABREPROVE
   ATX% = 0: NUCU& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(297)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     If CUIT$ <> "" Then
        VLBD$ = VALOBADA("Proved12", "Cuit_Cli", "CUIT_CLI = '" & CUIT$ & "'")
        If VLBD$ <> "" Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - CUIT de Proveedor Existente"
          List1.AddItem MERROX$
          GoTo 19
        End If
     End If
     For KKL& = 1 To NUCU&
       If CUTTX$(KKL&) = CUIT$ Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - CUIT de Proveedor Repetido"
          List1.AddItem MERROX$
          GoTo 19
        End If
     Next KKL&
     '
     DBGrid1.Col = 0
     NUPROV& = XVALO(TRIM$(DBGrid1.TEXT))   ' NUMERO DE CUENTA
     If NUPROV& < 1 Or NUPROV& > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     ' Verifica si el numero ya ha sido utilizado
25   If Asc(Mid$(PUNUCLI$, NUPROV&, 1)) > 0 Then
       NUPROV& = NUPROV& + 10000
       If NUPROV& > 32000 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor No Procesable"
          List1.AddItem MERROX$
          GoTo 19
       End If
       GoTo 25
     End If
     Mid$(PUNUCLI$, NUPROV&, 1) = Chr$(1)
     '
     CODEX$ = AJUSTD$(Str$(Int(NUPROV + 0.5)), 5)
12   Call REPLA(TTXX$, CODEX$, 1, 5)
     '
     DBGrid1.Col = 1                     ' RAZON SOCIAL
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social del Proveedor"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 48)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 64)
     Call REPLA(TTXX$, DOMI$, 55, 64)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 16)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 48)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DBGrid1.Col = 9                     ' TIPO DE CUENTA
     TIPOCTA$ = TRIM$(DBGrid1.TEXT)
     If TIPOCTA$ = "" Then TIPOCTA$ = "0"
     If TIPOCTA$ = "0" Or TIPOCTA$ = "1" Then
       Call REPLA(TTXX$, TIPOCTA$, 235, 4)
     Else
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Tipo de Cuenta"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     DBGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT) + "/B"
     Call REPLA(TTXX$, CODICUEN$, 240, 12)
     '
     '\\\OT-05-0202-WAR-18/04/05///
     DBGrid1.Col = 11                    ' CODIGO DE PROVINCIA
     CODIPROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODIPROVI$, 253, 2)
     '
     DBGrid1.Col = 12                    ' CODIGO DE POSICION IVA
     CODIVA% = XVALO(DBGrid1.TEXT)
     Call REPLA(TTXX$, FORMATNUM$(CODIVA%, "I4"), 255, 4)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 13                    ' CUENTA MADRE
     CTAMADRE$ = TRIM$(DBGrid1.TEXT)
     If Len(CTAMADRE$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Longitud de Cuenta Madre"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CTAMADRE, 261, 12)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 14                    ' nombre de fantasia
     FANTASI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FANTASI$, 275, 16)
     '
     DBGrid1.Col = 15                    ' RETENCION DE INGRESOS BRUTOS
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F4.0")
     Call REPLA(TTXX$, VALO$, 293, 4)
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&
   '
20 DBGrid1.RowHeight = hrow
   Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORPLA()
   '
   If DBGrid1.Columns.Count < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     '
     NIVEI% = XVALO(DBGrid1.TEXT)
     If NIVEI% < 0 Or NIVEI% > 5 Then NIVEI% = 0
     Call REPLA(TTXX$, AJUSTD$(Str$(NIVEI%), 4), 1, 4)
     '
     DBGrid1.Col = 1
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CODEX$, 7, 12)
     '
     For KKU% = 1 To List4.ListCount
       If TRIM$(Mid$(List4.List(KKU% - 1), 7, 12)) = CODEX$ Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Duplicado o Pre-Existente"
         List1.AddItem MERROX$
         GoTo 19
       End If
     Next KKU%
     '
     DBGrid1.Col = 2
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 21, 48)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ <> "D" And VTX$ <> "H" And VTX$ <> "I" Then VTX$ = "G"
     Call REPLA(TTXX$, VTX$, 71, 1)
     '
     DBGrid1.Col = 4
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ = "0" Or VTX$ = "NO" Then VTX$ = "NO"
     If VTX$ <> "NO" Then VTX$ = "SI"
     Call REPLA(TTXX$, VTX$, 75, 2)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
20 End Sub
 '\\\OT-05-0207-WAR-21/04/05///
Sub PREIMPORNORCLI()
   '
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0            'CLIENTE CODIGO STRING
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If U& = 250 Then
       XX = "AA"
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 12)
     '
     DBGrid1.Col = 1            ' ARTICULO
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Articulo"
       List1.AddItem MERROX$
       GoTo 19
     End If
     CDXI% = CODINT%(DEDEX$)
     If CDXI% < 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     Call REPLA(TTXX$, DEDEX$, 14, 16)
     Call REPLA(TTXX$, DESCRIT0(CDXI%), 32, 22)
     '
     DBGrid1.Col = 2                     ' NUMERO ORIGINAL
     VALO$ = TRIM$(DBGrid1.TEXT)
     If Len(VALO$) < 0 Or Len(VALO$) > 28 Or TRIM(VALO$) = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero Original no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 56, 28)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0207-WAR-FIN///
Sub IMPORMAS()
   '
   FIN& = List4.ListCount
   Call APERBASDAT("MASTER")
   '
   NMS% = NUMAS%
   'REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     '\\\OT-05-0499-WAR-13/07/05///
       'ACA TOMAR NUEVOS CAMPOS
       CODALT$ = TRIM$(Mid$(TTXX$, 25, 24))
       DESCOMP$ = TRIM$(Mid$(TTXX$, 115, 64))
       ATRIA$ = TRIM$(Mid$(TTXX$, 180, 32))
       ATRIB$ = TRIM$(Mid$(TTXX$, 213, 32))
     '\\\OT-05-0499-WAR-FIN///
     DESIT$ = TRIM$(Mid$(TTXX$, 50, 64))
     TIMAT% = XVALO(Mid$(TTXX$, 246, 4)) Mod 256
     UNIME$ = TRIM$(Mid$(TTXX$, 251, 4))
     STOMI = XVALO(Mid$(TTXX$, 256, 12))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = XVALO(Mid$(TTXX$, 269, 12))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     KILUNI = XVALO(Mid$(TTXX$, 282, 8))
     KILMET = XVALO(Mid$(TTXX$, 292, 8))
     COSUN = XVALO(Mid$(TTXX$, 302, 12))
     MONEC% = XVALO(Mid$(TTXX$, 315, 4)) Mod 256
     FECOU% = FECHANUM(Mid$(TTXX$, 320, 8))
       PRH = Abs(XVALO(Mid$(TTXX$, 329, 8)))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     GRUCO = XVALO(Mid$(TTXX$, 338, 6))
     STOAC = XVALO(Mid$(TTXX$, 346, 10))
     CINTERNA$ = Mid(TTXX$, 360, 64)
     '
     SQLX$ = "SELECT * FROM Master "
     Dim ARTICULOS As ADODB.Recordset
     Set ARTICULOS = New ADODB.Recordset
     ARTICULOS.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With ARTICULOS
     On Error GoTo 0
     If CODINT%(CODIT$) = 0 Then
       With ARTICULOS
         .AddNew   ' YA CONVERTIDO
         ON ERROR RESUME NEXT
         !CodiEmpr = Codiemp%
         ON ERROR GOTO 0
         '\\\OT-05-0499-WAR-13/07/05///
         !NPLANO = CODALT$
         !DESCRIDOS = DESCOMP$
         !TALLE = ATRIA$
         !COLOR = ATRIB$
         '\\\OT-05-0499-WAR-FIN///
         !Codigo = CODIT$
         !Descripcion = DESIT$
         !Status = 1
           NMS% = NMS% + 1
         !CODINT = NMS%
         !UMEDIDA = AJUSTI$(UNIME$, 4)
         !TIMATE = TIMAT%
         !CRITREP = 0
         !Monecos = MONEC%
         !FECOSTO = CVDAT(FECOU%)
         !COSUNI = COSUN
         !STOMIN = STOMI
         !LOREPOS = LECON
         !STOMAX = LOMAX
         !PROVHABI = PRHAB%
         '
         !Alias = " "
         !PORNACIO = 0
         !PESUNI = 0
         '!COLOR = " "
         '!TALLE = " "
         !DESCOMPO = 0
         !GRUPIVA = 0
         !Grucoven = GRUCO
         !CLinea = " "
         !Cfamilia = " "
         '!NPLANO = " "
         !Revision = " "
         !FEREVIS = CVDAT(0)
         !Observa = " "
         !Especificacion = " "
         !PESUNI = KILUNI
         !PESMETRO = KILMET
         !Cod_Interna = CINTERNA$
        .Update
       End With
       '
       If Abs(STOAC) >= 0.05 Then
         HOII% = HOY(HOS$)
         NORI% = NORI% + 1
         Call MOVISTOK(HOII%, NMS%, "", "AJ", "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "Stock Importado", NORI%, CDbl(COSUN), MONEC%, 0, 0, STOAC)
       End If
       '
     End If
     '
   Next U&
   '
   ARTICULOS.Close
   Set ARTICULOS = Nothing
   UPDATMASTER% = 1
   Call GENECOMAS
   Screen.MousePointer = 1
   '
End Sub
'

Sub IMPORSTRU()
   '
   Dim PUNSTRU$(32767)
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 24))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   ATENTIA% = ATENTI%: ATENTI% = 12345 ' PA QUE NO REORGANICE IMPLOSI
   FIN& = NMS%
   For CDXI% = 1 To NMS%
     KUU& = CDXI%
     Call TRACE(20, KUU&, FIN&)
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       CAIT% = 0
'Call LEEEXPLO(CODEXT$(CDXI%), 1)
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = Space$(128)
         COELEI% = CODINT%(TRIM$(Mid$(TTXX$, 27, 24)))
         COELE$ = CODEXT$(COELEI%)
         CAUNIDS = XVALO(Mid$(TTXX$, 98, 12))
         PDX$ = Mid$(TTXX$, 124, 4)
         PESOMET = PESMETRO(COELEI%)
         LARCORTE = XVALO(Mid$(TTXX$, 111, 12))
         '
         CANTI1 = XVALO(Mid$(TTXX$, 85, 12))
         USOUNI = 1
         If PESOMET > 0 Then
           If LARCORTE > 0 Then
             USOUNI = LARCORTE * PESOMET / 1000
             CANTI1 = CAUNIDS * USOUNI
           End If
         End If
         '
         Call REPLA(TTYY$, COELE$, 1, 16)
         USOX$ = FORMATNUM$(CANTI1, "F12.4")
         Call REPLA(TTYY$, USOX$, 49, 12)
         Call REPLA(TTYY$, PDX$, 61, 4)
         Call REPLA(TTYY$, FORMATNUM(CAUNIDS, "F9.4"), 89, 9)
         Call REPLA(TTYY$, FORMATNUM(PESOMET, "F7.3"), 101, 7)
         Call REPLA(TTYY$, FORMATNUM(LARCORTE, "F6.1"), 111, 6)
         Call REPLA(TTYY$, FORMATNUM(USOUNI, "F8.4"), 119, 8)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
         '
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
   ATENTI% = ATENTIA%
   '
End Sub
'
Sub IMPORUTS()
   '
   Dim PUNSTRU$(32767)
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   '
   Call ABRESECOP
   '
   REBLA$ = REGBLAN$("SECOPER")
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       With Secoper
63       .FindFirst "CODICONJ =" & CDXI%
         If .NoMatch = False Then
           .Delete
           .FindNext "CODICONJ =" & CDXI%
           GoTo 63
         End If
         '
         NOPEI% = 0
         For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
           U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
           TTXX$ = List4.List(U& - 1)
           '
           SQLX$ = "SELECT * FROM Secoper "
           Dim SECOPER1 As ADODB.Recordset
           Set SECOPER1 = New ADODB.Recordset
           SECOPER1.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           With SECOPER1
           '
            .AddNew   ' YA CONVERTIDO
            ON ERROR RESUME NEXT
            !CodiEmpr = Codiemp%
            ON ERROR GOTO 0
            !CODICONJ = CDXI%
            !CODXCONJ = CODEXT$(CDXI%)
              NOPEI% = NOPEI% + 10
            !NumeOper = NOPEI%
            !DescOper = Mid$(TTXX$, 23, 48)
            !statoper = 1
            !CodMaq0 = Mid$(TTXX$, 73, 6)
            !CodMaq1 = Mid$(TTXX$, 139, 6)
            !CodMaq2 = Mid$(TTXX$, 147, 6)
            !CodMaq3 = Mid$(TTXX$, 155, 6)
            !HorsPrep = XVALO(Mid$(TTXX$, 79, 10))
            !HorsFijo = XVALO(Mid$(TTXX$, 89, 10))
            !MinuStan = XVALO(Mid$(TTXX$, 99, 12))
            !HorsLimp = 0
            !HorsMHer = 0
            !PieCiclo = 0
            !PorDemor = XVALO(Mid$(TTXX$, 111, 10))
            If !MinuStan > 0.005 Then
              !Cadencia = 60 / !MinuStan
            End If
            !CantiOps = XVALO(Mid$(TTXX$, 129, 8))
            !CategOps = 0
            !CoefSolp = 0
            !LOTESTAN = XVALO(Mid$(TTXX$, 121, 8))
            !ProgCNum = Mid$(TTXX$, 163, 6)
            !HoProFab = " "
            !ProTerOp = 0
            !PreUnid = 0
            !MonePrec = 0
            !FechPrec = CVDAT(0)
            '
            !OBSEOPER = " "
            !NUOPER_PREP = 1
            !CATOPER_PREP = 0
            .Update
           End With
           '
         Next UKU&
         
        ARTICULOS.Close
        Set ARTICULOS = Nothing
        
       End With
     End If
   Next CDXI%
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAQ()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADMAQ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COMAQ$ = TRIM$(Left$(TTXX$, 6))
     DEMAQ$ = TRIM$(Mid$(TTXX$, 9, 30))
     CUHOR = XVALO(Mid$(TTXX$, 41, 12))
     SECCI = XVALO(Mid$(TTXX$, 53, 10))
     MAQEQ$ = TRIM$(Mid$(TTXX$, 66, 6))
     COEFI% = XVALO(Mid$(TTXX$, 72, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 87, 30))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, COMAQ$, 1, 6)
     Call REPLA(REMAQ$, DEMAQ$, 7, 30)
     Call REPLA(REMAQ$, MKS$(SECCI), 37, 4)
     Call REPLA(REMAQ$, MAQEQ$, 43, 6)
     Call REPLA(REMAQ$, MKI$(COEFI), 49, 2)
     Call REPLA(REMAQ$, OBSER$, 51, 30)
     Call REPLA(REMAQ$, MKS$(CUHOR), 93, 4)
     Call FILTERSETRECORD("PADMAQ", 1, COMAQ$, REMAQ$) ' sobreescribe SI IGUAL CODIGO
     '
   Next U&
   Call CIERRARCH("PADMAQ")
   '
End Sub

Sub IMPORLISPRE()
   '
   Call SELECHO("LISTAS")
   LPRE% = XVALO(VALACT1$("LISTAS"))
   If LPRE% < 1 Then Exit Sub
   '
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   If ULTREG&(ARCHIPRE$) > 0 Then
     If COMALTER%("La Lista de Precios Elegida Contiene Información.\¿ Desea Sobre-Escribirla o Cancelar Importación ?\\Cancelar Importación\Sobre-Escribir Lista") <> 2 Then
       Exit Sub
     End If
   End If
   '
   Dim ListaPrecio As ADODB.Recordset
   Set ListaPrecio = New ADODB.Recordset
   
   ListaPrecio.OPEN FILTSQL$ ("SELECT * FROM PREVENTA"), FLEXCONN, adOpenKeyset, adLockPessimistic
   
   
   FIN& = List4.ListCount
   JJ& = 0
   'REBLA$ = REGBLAN$(ARCHIPRE$)
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COLIS% = XVALO(Left$(TTXX$, 4))
     If COLIS% = LPRE% Then
       CODE$ = TRIM$(Mid$(TTXX$, 6, 16))
       CIXI% = CODINT%(CODE$)
       If CIXI% > 0 Then
         DEDE$ = TRIM$(Mid$(TTXX$, 24, 64))
         If DEDE$ = "" Then DEDE$ = DESCRIT0$(CIXI%)
         NORI$ = TRIM$(Mid$(TTXX$, 90, 16))
         PRUNI# = XVALO(Mid$(TTXX$, 108, 12))
         '
         'TTYY$ = REBLA$
         'Call REPLA(TTYY$, MKI$(CIXI%), 1, 2)
         'Call REPLA(TTYY$, DEDE$, 3, 42)
         'Call REPLA(TTYY$, NORI$, 46, 16)
        ' Call REPLA(TTYY$, MKD$(PRUNI#), 62, 8)
         'JJ& = JJ& + 1
         'Call GRAREG(ARCHIPRE$, TTYY$, JJ&)
         '
         With ListaPrecio
             .AddNew   ' YA CONVERTIDO
             ON ERROR RESUME NEXT
             !CodiEmpr = Codiemp%
             ON ERROR GOTO 0
             !Cod_Lista = LPRE%
             !Cod_Pieza = CIXI%
             !Num_Orden = U&
             !Desc_Pieza = DESCRIT0$(CIXI%)
             !Denom_Pieza = DESCRIT0$(CIXI%)
             !Num_Original = NORI$
             !Precio_Unitario = PRUNI#
             !Cant_Limite = 0
             !Porc_Limite = 0
             .Update
         End With
         
       End If
     End If
   Next U&
   '
   ListaPrecio.Close
   Set ListaPrecio = Nothing
   'Call SETULTREG(ARCHIPRE$, JJ&)
   'Call CIERRARCH(ARCHIPRE$)
   '
End Sub
'
Sub IMPORCAT()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("CATOPER")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     ICAT% = XVALO(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = XVALO(Mid$(TTXX$, 37, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 51, 32))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, Chr$(ICAT%), 1, 1)
     Call REPLA(REMAQ$, DECAT$, 2, 27)
     Call REPLA(REMAQ$, MKS$(CUHOR), 29, 4)
     Call REPLA(REMAQ$, OBSER$, 33, 32)
     JJ& = JJ& + 1
     Call GRAREG("CATOPER", REMAQ$, JJ&)
   Next U&
   '
   Call SETULTREG("CATOPER", JJ&)
   Call CIERRARCH("CATOPER")
   '
End Sub

Sub IMPORCLI()
   '
   LONGINT% = 0
   PUNUCLI$ = String$(32767, 0)
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   '
   ' Genera DEUDOR1 y DEUDOR2
   Call ACLI_GES07.GENECOCLI
   '
'   ' Establece en el "vector" PUNUCLI$, los numeros de clientes ya utilizados
'   For U& = 1 To ULTREG&("DEUDOR1")
'     XX$ = REGLEIDO$("DEUDOR1", U&)
'     NCI% = CVI(Left$(XX$, 2))
'     If NCI% > 0 Then
'       If NCI% <= 32767 Then
'         Mid$(PUNUCLI$, NCI%, 1) = Chr$(1)
'       End If
'     End If
'   Next U&
'   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   DBHANT = DBGrid1.Height
   URECORD& = Data1.Recordset.RecordCount
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   '
   Call APERBASDAT("CLIEN")
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI& = XVALO(Mid$(TTXX$, 1, 6))
     If NCI& < 1 Then GoTo 19
     If NCI& > 32000 Then LONGINT% = 1
     ' Verifica si el numero ya ha sido utilizado
     While Len(PUNUCLI$) < NCI&: PUNUCLI$ = PUNUCLI$ + Chr$(0): Wend
     If Asc(Mid$(PUNUCLI$, NCI&, 1)) > 0 Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 8, 47)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 24))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     GruCli% = XVALO(Mid$(TTXX$, 244, 4))
     CODICUEN$ = TRIM$(Mid$(TTXX$, 249, 12))
        If CODICUEN$ = "" Then
           CODICUEN$ = "CLI-" + TRIM$(Str$(NCI&))
        End If
     '\\\OT-05-0192-WAR-18/04/05///
     CONDI% = XVALO(Mid$(TTXX$, 261, 4))
     PIVA% = XVALO(Mid$(TTXX$, 269, 4))
     PROVI$ = TRIM$(Mid$(TTXX$, 267, 2))
     LIPRE% = 0
     COVEN% = XVALO(Mid$(TTXX$, 273, 4))
     LICRE = XVALO(Mid$(TTXX$, 277, 12))
     CUEMA% = 0
     '\\\OT-5-632-FG-21/09/05///
     DESSTAN$ = Mid$(TTXX$, 291, 24)
     LUGENT$ = Mid$(TTXX$, 317, 48)
     LOCALI$ = Mid$(TTXX$, 365, 48)
     NTRAN& = XVALO(Mid$(TTXX$, 413, 5))
     TRANSPOR$ = Mid$(TTXX$, 421, 40)
     GRUDIV% = XVALO(Mid$(TTXX$, 461, 8))
     EMAIX$ = Mid$(TTXX$, 475, 36)
     POREMBA = XVALO(Mid$(TTXX$, 511, 6))
     '\\\OT-5-632-FG-FIN///
     'GoTo 19 'PARA TESTEAR
     '\\\OT-05-0192-WAR-FIN///
     INIX& = U& - 1
     PAISX$ = "": LPRX% = 0: WEBX$ = ""
12   For U2& = INIX& To Data1.Recordset.RecordCount
        DBGrid1.Row = U2&
        DBGrid1.Col = 0
        NCLX& = XVALO(Data1.Recordset.Fields(0).Value)
        If NCLX& = NCI& Then
           DBGrid1.Col = 5: PAISX$ = DBGrid1.TEXT
           DBGrid1.Col = 15: LPRX% = XVALO(DBGrid1.TEXT)
           DBGrid1.Col = 26: WEBX$ = LCase$(DBGrid1.TEXT)
           GoTo 15
        End If
     Next U2&
     If INIX& > 0 Then
        INIX& = 0
        GoTo 12
     End If
     '
15   On Error GoTo 0
'     With Deudor12
     SQLX$ = "SELECT * FROM DEUDOR12"
     Dim DEUDDDOR12 As ADODB.Recordset
     Set DEUDDDOR12 = New ADODB.Recordset
     DEUDDDOR12.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With DEUDDDOR12
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !Stat_Cli = 1                        'STATUS ACTIVO
       !Nume_Cli = NCI&                     'CODIGO ENTERO
       !Codi_Cli = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = 0                        'TIPO CTA
       !GrCo_Cli = GruCli%                  'TIPO CLI  VA A GRUPO CLI
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = Left$(RASO$, 10)         'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = PROVI$                  'COD PROVINCIA
       !Pais_Cli = PAISX$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       !Cpag_Cli = COVEN%                   'CONDICION DE PAGO
       !Lpre_Cli = LPRX%
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       !Piva_Cli = PIVA%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Lcre_Cli = LICRE                    'LIMITE CREDITO
       !Vend_Cli = CONDI%                   'CODIGO VENDEDOR
       !Ctac_Cli = CONTACTX$                'CONTACTO
       !Dsta_Cli = DESSTAN$                 'DESCUENTO
       !Lentre0_Cli = LUGENT$               'LUGAR DE ENTREGA
       !Lentre1_Cli = LOCALI$               'LOCALIDAD
       !Lentre2_Cli = TRANSPOR$             'TRANSPORTE
       !NTrans_Cli = NTRAN&                 'NUMERO DE TRANSPORTISTA
       !Gru_Div = GRUDIV%                   'GRUPO DIVISIONAL
       !Mail_Cli = EMAIX$                   'DIRECCION DE MAIL
       !Recar_Embal = POREMBA               'PORCENTAJE RECARGO EMBALAJE
    .Update
     End With
     '
     If NCI& > 0 Then
        Mid$(PUNUCLI$, NCI&, 1) = Chr$(1)
     End If
     '
19 Next U&
   DEUDDDOR12.Close
   Set DEUDDDOR12 = Nothing
   '
   UPDATCLI% = 1
   If LONGINT% > 0 Then
      Call COPYSTRU("IBOLPEN", "DEUDORL1")
   End If
   Call ACLI_GES07.GENECOCLI(LONGINT%)
   '
   '\\\OT-5-632-FG-21/09/05///
   Call ACLI_GES07.GENDAPIBCLI
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     TTXX$ = List4.List(U& - 1)
     NCI& = XVALO(Mid$(TTXX$, 1, 6))
     PERIB = XVALO(Mid$(TTXX$, 461, 4))
     '
'     X$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCI&))
     Call REPLA(X$, Chr$(PERIB), 3, 1)
'     Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCI%), X$)
20 Next U&
   '
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '\\\OT-5-632-FG-FIN///
End Sub

Sub IMPORPRO()
   '
   XX0$ = REGBLAN$("PROVED1") + REGBLAN$("PROVED2")
   '
   Call APERBASDAT("PROVE")
   '
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI% = XVALO(Mid$(TTXX$, 1, 5))
     If NCI% < 1 Then GoTo 19
     CODICUEN$ = TRIM$(Mid$(TTXX$, 240, 12))
     If CODICUEN$ = "" Then CODICUEN$ = "PR-" + TRIM$(Str$(NCI%))
     If TRIM$(CODICUEN$) = "" Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 7, 48)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 16))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     TIPOCTA% = XVALO(Mid$(TTXX$, 235, 4))
     CODIPROVI1$ = Mid$(TTXX$, 253, 2)
     CODIVA1% = XVALO(Mid$(TTXX$, 255, 4))
     CTAMADRE$ = TRIM$(Mid$(TTXX$, 261, 12))
     FANTASI$ = TRIM$(Mid$(TTXX$, 275, 16))
     'PIVA% = 1
     IBRU$ = ""
     CONDI% = 0
     LIPRE% = 0
     COVEN% = 0
     LICRE = 0
     CUEMA% = 0
     '
     
'     With Proved12
'       .FindFirst "Codi_Cli = '" & CODICUEN$ & "'"
'       If .NoMatch = False Then
     SQLX$ = "SELECT * FROM Proved12 "
     SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & CODICUEN$ & "'"
     Dim PROVVVED As ADODB.Recordset
     Set PROVVVED = New ADODB.Recordset
     PROVVVED.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With PROVVVED
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
           On Error GoTo 0
           OPX% = COMALTER%("Cuenta '" + TRIM$(CODICUEN$) + "' Existente en B.Datos\\Actualizar\Conservar\Cancelar Proceso")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           If OPX% = 2 Then GoTo 19
         Else
            On Error GoTo 0
           .AddNew   ' YA CONVERTIDO
           ON ERROR RESUME NEXT
           !CodiEmpr = Codiemp%
           ON ERROR GOTO 0
         !Nume_Cli = NCI%                     'CODIGO ENTERO
       End If
       !Stat_Cli = 1                        'STATUS ACTIVO
       !Codi_Cli = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = TIPOCTA%                      'TIPO CTA
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = FANTASI$                 'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = CODIPROVI1$              'COD PROVINCIA
       '!Pais_Cli = PAISOK$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       '!Cpag_Cli = 0                       'POR DEFAULT (XQ' NO SE JUSTIFICAN 60/900)
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       '!Mail_Cli = MAILPRO$                'MAIL
       !Piva_Cli = CODIVA1%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Ctac_Cli = CONTACTX$
       CMINT% = CUECOINT%(CTAMADRE$)
       !CuMa_Cli = CMINT%                   'CUENTA MADRE
       .Update
     End With
     '
19 Next U&
   PROVVVED.Close
   Set PROVVVED = Nothing
   '
   '\\\OT-5-632-FG-21/09/05///
   UPDATPRO% = 1
   Call ACRE_GES07.GENECOPROV    ' PARA ACTUALIZAR LAS LISTAS DE SELECCIóN
   '
   Call ACRE_GES07.GENDARIBPRO
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     TTXX$ = List4.List(U& - 1)
     NPI% = XVALO(Mid$(TTXX$, 1, 5))
     RETIB = XVALO(Mid$(TTXX$, 293, 4))
     '
     X$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPI%))
     Call REPLA(X$, Chr$(RETIB), 3, 1)
     Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NPI%), X$)
20 Next U&
   '
   '\\\OT-5-632-FG-FIN///
   '
99 On Error Resume Next
   'BadaproVE.Close
   On Error GoTo 0
   '
End Sub

Sub IMPORPLA()
   '
   Call BLOQARCH("CUENTAS")
   Call BLOQARCH("DENOCUE")
   Call BLOQARCH("INVECUE")
   Call BLOQARCH("PLAGRAL")
   Call BLOQARCH("PLANCUEN")
   Call BLOQARCH("CUECON")
   Call BLOQARCH("ECUECON")
   Call BLOQARCH("ECUECOT")
   '
   JJ& = 0: J4& = 0: J5& = 0: J6& = 0
   URECU& = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("PLANCUEN")
   '
   FIN& = List4.ListCount
   List3.Clear
   '
   Call APERBASDAT("PLANCUEN")
20 On Error Resume Next
   Plancuen.Close
   On Error GoTo 0
   '
   On Error Resume Next
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset, dbDenyWrite)
   If Err Then
     If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
   On Error Resume Next
   Plancuen.MoveLast
   NUREPLAN% = Plancuen.RecordCount
   Plancuen.MoveFirst
   On Error GoTo 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ZZ$ = List4.List(U& - 1)
     AGRECUE% = 0
     COCUE$ = TRIM$(Mid$(ZZ$, 7, 12))
     DECUE$ = TRIM$(Mid$(ZZ$, 21, 48))
     '
     CUEIN% = 0
     For KU& = 1 To URECU&
       Z1$ = REGLEIDO$("CUENTAS", KU&)
       If TRIM$(Left$(Z1$, 12)) = COCUE$ Then
         CUEIN% = KU&
         Exit For
       End If
     Next KU&
       '
       If CUEIN% < 1 Or CUEIN% > URECU& Then
         URECU& = URECU& + 1
         CUEIN% = URECU&
         AGRECUE% = 1
       End If
       '
     RECUENTA& = CUEIN%
     TICU$ = Mid$(ZZ$, 71, 1)
     IBLE% = 1: If UCase$(Mid$(ZZ$, 75, 1)) = "N" Then IBLE% = 0
     STAT% = 1
     NIVEI% = XVALO(Left$(ZZ$, 4))
     '
     ' GRABAR EN CUENTAS .DAT
     If AGRECUE% < 1 Then
         Z1$ = REGLEIDO$("CUENTAS", RECUENTA&)
       Else
         Z1$ = REGBLAN$("CUENTAS")
     End If
     Call REPLA(Z1$, COCUE$, 1, 12)
     Call REPLA(Z1$, TICU$, 13, 1)
     Call GRAREG("CUENTAS", Z1$, RECUENTA&)
     '
     ' GRABAR EN DENOCUE.DAT
     For KU& = 1 + ULTREG&("DENOCUE") To RECUENTA&
       Call GRAREG("DENOCUE", Space$(127) + ".", KU&)
     Next KU&
     Z2$ = REGLEIDO$("DENOCUE", RECUENTA&)
     Call REPLA(Z2$, DECUE$, 1, 100)
     Call REPLA(Z2$, DECUE$, 101, 28)
     Call GRAREG("DENOCUE", Z2$, RECUENTA&)
     '
     ' GRABAR EN PLANCUEN.DAT
     Z3$ = REBLA$
     Call REPLA(Z3$, COCUE$, 1, 12)
     Call REPLA(Z3$, DECUE$, 13, 46)
     Call REPLA(Z3$, MKI$(CUEIN%), 59, 2)
     Call REPLA(Z3$, Chr$(STAT%), 61, 1)
     Call REPLA(Z3$, Chr$(NIVEI%), 62, 1)
     Call REPLA(Z3$, TICU$, 63, 1)
     Call REPLA(Z3$, Chr$(IBLE%), 64, 1)
     Call REPLA(Z3$, DECUE$, 129, 128)
     JJ& = JJ& + 1
     Call GRAREG("PLANCUEN", Z3$, JJ&)
     '
     ' GRABAR EN PLAGRAL.DAT
     Z4$ = MKI$(CUEIN%) + MKI$(NIVEI%)
     Call GRAREG("PLAGRAL", Z4$, JJ&)
     '
     ' GRABAR EN LIST3 PARA INVECUE
     TTXX$ = Space$(18)
     Call REPLA(TTXX$, COCUE$, 1, 12)
     Call REPLA(TTXX$, TRIM$(Str$(CUEIN%)), 13, 6)
     For KI% = 1 To List3.ListCount
       If List3.List(KI% - 1) > TTXX$ Then
         List3.AddItem TTXX$, KI% - 1
         GoTo 25
       End If
     Next KI%
     List3.AddItem TTXX$
     '
25   Z4$ = REGBLAN$("CUECON")
     Z5$ = REGBLAN$("ECUECON")
     '
     Call REPLA(Z5$, COCUE$, 1, 12)
     Call REPLA(Z5$, DECUE$, 13, 28)
     Call REPLA(Z5$, MKI$(CUEIN%), 41, 2)
     J5& = J5& + 1
     Call GRAREG("ECUECOT", Z5$, J5&)
     '
     If STAT = 1 Then
       If IBLE% = 1 Then
         '
         J6& = J6& + 1
         Call GRAREG("ECUECON", Z5$, J6&)
         '
         If TICU$ = "I" Or TICU$ = "G" Then
           Call REPLA(Z4$, MKI$(CUEIN%), 1, 2)
           Call REPLA(Z4$, DECUE$, 3, 30)
           Call REPLA(Z4$, COCUE$, 33, 12)
           Call REPLA(Z4$, COCUE$, 45, 12)
           J4& = J4& + 1
           Call GRAREG("CUECON", Z4$, J4&)
         End If
         '
       End If
     End If
     '
     On Error GoTo 0
'
'     With Plancuen
'       .FindFirst "CODEXCUE = '" & COCUE$ & "'"
'       If .NoMatch = False Then
'
     SQLX$ = "SELECT * FROM PLANCUEN "
     SQLX$ = SQLX$ + " WHERE CODEXCUE = '" & COCUE$ & "'"
     Dim PLANCUES As ADODB.Recordset
     Set PLANCUES = New ADODB.Recordset
     PLANCUES.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With PLANCUES
       On Error Resume Next
       If Err Then
           '.Edit
         Else
          .AddNew   ' YA CONVERTIDO
          ON ERROR RESUME NEXT
          !CodiEmpr = Codiemp%
          ON ERROR GOTO 0
       End If
       !CodExCue = COCUE$
       !DENOCUE = DECUE$
       !CODINCUE = CUEIN%
       !STATCUE = STAT%
       !NIVAGRU = NIVEI%
       Select Case TICU$
            Case "D": !TipoCue = "Debe"
            Case "H": !TipoCue = "Haber"
            Case "G": !TipoCue = "Gasto"
            Case "I": !TipoCue = "Ingre"
       End Select
       !POSIMPU = IBLE%
       !Observa = " "
       !Salini = 0
       !Sudebe = 0
       !Suhabe = 0
       !Result = 0
       !Salfin = 0
       '!IndiAjus = INAJUS%
       '!MoneExtr = MONED%
       .Update
     End With
     '
   Next U&
   PLANCUES.Close
   Set PLANCUES = Nothing
   '
90 Call SETULTREG("PLANCUEN", JJ&)
   Call SETULTREG("PLAGRAL", JJ&)
   Call SETULTREG("CUECON", J4&)
   Call SETULTREG("ECUECOT", J5&)
   Call SETULTREG("ECUECON", J6&)
   Call CIERRARCH("CUENTAS")
   Call CIERRARCH("DENOCUE")
   Plancuen.Close
   '
   REBLA$ = REGBLAN$("INVECUE")
   J7& = 0
   For KI% = 1 To List3.ListCount
     TTXX$ = List3.List(KI% - 1)
     COCUE$ = Left$(TTXX$, 12)
     ICUE% = XVALO(Mid$(TTXX$, 13))
     Z7$ = REBLA$
     Call REPLA(Z7$, COCUE$, 1, 12)
     Call REPLA(Z7$, MKI$(ICUE%), 13, 2)
     J7& = J7& + 1
     Call GRAREG("INVECUE", Z7$, J7&)
   Next KI%
   Call SETULTREG("INVECUE", J7&)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Close
   '
   ABCT$ = LOADFILE$(LUDAT$ + "CUENTAS.DAT"): Call SAVEFILE(LUDAT$ + "CUENTAS.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "DENOCUE.DAT"): Call SAVEFILE(LUDAT$ + "DENOCUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "INVECUE.DAT"): Call SAVEFILE(LUDAT$ + "INVECUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "PLAGRAL.DAT"): Call SAVEFILE(LUDAT$ + "PLAGRAL.OLD", ABCT$)
   '
   Screen.MousePointer = 1
   '
End Sub
  '\\\OT-05-0207-WAR-25/04/05///
Sub IMPORNCLI()
    '
    Screen.MousePointer = 11
    '
    REBLA$ = REGBLAN$("REPLACOC")
    J& = 0
    For KI% = 1 To List4.ListCount
      TTXX$ = List4.List(KI% - 1)
      NCI1% = NUMECLI(Mid$(TTXX$, 1, 12)) 'Nº cliente
      NP1% = CODINT(Mid$(TTXX$, 14, 16)) 'Nº ARTICULO
      NO1$ = TRIM$(Mid$(TTXX$, 56, 28))  'Nº ORIGINAL
      Z$ = REBLA$
      Call REPLA(Z$, MKI(NCI1%), 1, 2)
      Call REPLA(Z$, MKI$(NP1%), 3, 2)
      Call REPLA(Z$, NO1$, 5, 28)
      J& = J& + 1
      Call GRAREG("REPLACOC", Z$, J&)
    Next KI%
    Call SETULTREG("REPLACOC", J&)
    '
    Screen.MousePointer = 1
    '
End Sub

Sub IMPORCOSREP()
   '
   Call APERBASDAT("MASTER")
   '
   HUBOAMAS% = 0
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     CODE$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(CODE$)
     If CDXI% > 0 Then
        MONEVA% = Val(Mid$(TTXX$, 67, 4))
        VALORE# = XVALO(Mid$(TTXX$, 71, 16))
        FFII% = FECHANUM(Mid$(TTXX$, 89, 8))
        FEVALO = CVDAT(FFII%)
        NPROVE% = XVALO(Mid$(TTXX$, 97, 6))
        '
'        With Master
'          .FindFirst "Codint = " & CDXI%
'          If .NoMatch = False Then
        SQLX$ = "SELECT * FROM Master "
        SQLX$ = SQLX$ & "WHERE CODINT=" & CDXI%
        Dim ARTICULOS As ADODB.Recordset
        Set ARTICULOS = New ADODB.Recordset
        ARTICULOS.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        On Error GoTo 0
        'If CODINT%(CODIT$) = 0 Then
          With ARTICULOS
            !Monecos = MONEVA%
            If FFII% > 0 Then !FECOSTO = FEVALO
            !COSUNI = VALORE#
            If NPROVE% > 0 Then
              !PROVHABI = NPROVE%
            End If
            .Update
          End With
        'End If
     End If
     '
19 Next U&

   ARTICULOS.Close
   Set ARTICULOS = Nothing
   '
99 On Error Resume Next
   Master.Close
   On Error GoTo 0
   '
End Sub

'\\\OT-05-0207-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub PREIMPORCUECLI()
   '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   INI& = 1: FINI& = URECORD&
   If URECORD& > 30000 Then
     TTXX$ = "Imposible Importar en Secuencia Unica.\Elija Rango de Importacion:\\1 - 30000"
     FINI& = 30000
5    INI& = FINI& + 1
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
     TTXX$ = TTXX$ + "\" + TRIM$(Str$(INI&)) + " - " + TRIM$(Str$(FINI&))
     If FINI& < URECORD& Then GoTo 5
     '
     OPX& = COMALTER%(TTXX$)
     If OPX& < 1 Then GoTo 20
     INI& = 1 + 30000 * (OPX& - 1)
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
   End If
   '
   Screen.MousePointer = 11
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
       U& = U& + 1
       Call TRACE(20, U&, URECORD&)
       '
       On Error Resume Next
       CODICLIE$ = ""
       CODICLIE$ = .Recordset.Fields(0)
       REFEMOVI$ = ""
       REFEMOVI$ = .Recordset.Fields(1)
       FECHEMIS$ = ""
       FECHEMIS$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
       FechVenc$ = ""
       FechVenc$ = FECHATEX$(CVINT(.Recordset.Fields(3)))
         If FECHANUM(FechVenc$) < FECHANUM(FECHEMIS$) Then FechVenc$ = FECHEMIS$
       CODICOMLAR$ = ""
       CODICOMCOR$ = ""
       TIPOCO$ = UCase$(TRIM$(.Recordset.Fields(4)))
         If TIPOCO$ = "RE" Then TIPOCO$ = "RB"
       LETRAC$ = Left$(TRIM$(.Recordset.Fields(5)), 1)
         If LETRAC$ = "0" Then LETRAC$ = "X"
       PREFIJ$ = TRIM$(.Recordset.Fields(6))
         While Len(PREFIJ$) < 4: PREFIJ$ = "0" + PREFIJ$: Wend
         PREFIJ$ = Right$(PREFIJ$, 4)
       NUMECO$ = TRIM$(.Recordset.Fields(7))
         While Len(NUMECO$) < 8: NUMECO$ = "0" + NUMECO$: Wend
         NUMECO$ = Right$(NUMECO$, 8)
       If TIPOCO$ = "RB" Then
           CODIMLAR$ = TIPOCO$ + "-" + LETRAC$ + "-" + PREFIJ$ + "-" + NUMECO$
           CODIMCOR$ = TIPOCO$ + "-" + Mid$(NUMECO$, 3)
         Else
           CODIMLAR$ = TIPOCO$ + "-" + PREFIJ$ + "-" + NUMECO$ + "-" + LETRAC$
           CODIMCOR$ = TIPOCO$ + "-" + Mid$(NUMECO$, 3) + "-" + LETRAC$
       End If
       '
       IMORIGI# = 0: IMPENDE# = 0
       On Error Resume Next
       IMORIGI# = CDbl(.Recordset.Fields(8))
       IMPENDE# = CDbl(.Recordset.Fields(9))
       On Error GoTo 0
       '
2      If IsNumeric(CODICLIE$) = True Then
          If XVALO(CODICLIE$) > 0 Then
            If XVALO(CODICLIE$) <= 32767 Then
              NUCLIE% = XVALO(CODICLIE$)
              If STATCLI%(NUCLIE%) <> 0 Then
                CODICLIE$ = CODICLI$(NUCLIE%)
                GoTo 15
              End If
            End If
          End If
       End If
       NUCLIE% = NUMECLI%(CODICLIE$)
       If NUCLIE% < 1 Then
         If COMALTER%("Cliente " + TRIM$(REFEMOVI$) + " No Registrado.\Desea Darlo de Alta ?\\Si\No") = 1 Then
           Call APERBASDAT("CLIEN")
           NROPRO& = 0
           '
3          NROPRO& = NROPRO& + 1
           QTR = CantRegistros("DEUDOR12", "Nume_Cli = " & NROPRO&)
           If QTR > 0 Then GoTo 3
           '
           NCLIE = NROPRO&
           Call APERBASDAT("CLIEN")
           With Deudor12
               .AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
               !CodiEmpr = Codiemp%
               ON ERROR GOTO 0
               !Nume_Cli = TRIM$(Str$(NCLIE))
               !Raso_Cli = REFEMOVI$
               !Codi_Cli = CODICLIE$
               !Cuit_Cli = "00-00000000-0"
               !Piva_Cli = 1
               !Stat_Cli = 1
               '
               On Error Resume Next
               .Update
               If Err Then
                  Call MENSERR(24, "Alta no Procesada Correctamente")
               End If
               On Error GoTo 0
           End With
           GoTo 2
         End If
         '
         If CODICLIE$ <> CODICLIA$ Then
           CODICLIA$ = CODICLIE$
           MERROX$ = "Falta Cliente '" + TRIM$(CODICLIE$) + " - " + REFEMOVI$ + "' en B.Datos de Clientes"
           List1.AddItem MERROX$
         End If
         GoTo 19
       End If
       '
       ' VALIDAR QUE NO IMPORTE DOS VECES EL MISMO COMPROBANTE
       '
15     TTXX$ = Space$(465)
       Call REPLA(TTXX$, CODICLIE$, 1, 12)
       Call REPLA(TTXX$, FORMATNUM$(NUCLIE%, "F6,0"), 13, 6)
       Call REPLA(TTXX$, REFEMOVI$, 21, 30)
       Call REPLA(TTXX$, FECHEMIS$, 53, 8)
       Call REPLA(TTXX$, FechVenc$, 63, 8)
       Call REPLA(TTXX$, CODIMLAR$, 75, 20)
       Call REPLA(TTXX$, CODIMCOR$, 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMORIGI#, "F13.2"), 107, 13)
       Call REPLA(TTXX$, FORMATNUM$(IMPENDE#, "F13.2"), 120, 13)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub PREIMPORCUEPRO()
   '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   INI& = 1: FINI& = URECORD&
   If URECORD& > 30000 Then
     TTXX$ = "Imposible Importar en Secuencia Unica.\Elija Rango de Importacion:\\1 - 30000"
     FINI& = 30000
5    INI& = FINI& + 1
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
     TTXX$ = TTXX$ + "\" + TRIM$(Str$(INI&)) + " - " + TRIM$(Str$(FINI&))
     If FINI& < URECORD& Then GoTo 5
     '
     OPX& = COMALTER%(TTXX$)
     If OPX& < 1 Then GoTo 20
     INI& = 1 + 30000 * (OPX& - 1)
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
   End If
   '
   Screen.MousePointer = 11
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
       U& = U& + 1
       Call TRACE(20, U&, URECORD&)
       If U& < INI& Or (U& > FINI& And FINI& < URECORD&) Then GoTo 19
       '
       On Error Resume Next
       CODICLIE$ = ""
       CODICLIE$ = .Recordset.Fields(0)
       REFEMOVI$ = ""
       REFEMOVI$ = .Recordset.Fields(1)
       FECHEMIS$ = ""
       FECHEMIS$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
       FechVenc$ = ""
       FechVenc$ = FECHATEX$(CVINT(.Recordset.Fields(3)))
         If FECHANUM(FechVenc$) < FECHANUM(FECHEMIS$) Then FechVenc$ = FECHEMIS$
       CODICOMLAR$ = ""
       CODICOMCOR$ = ""
       TIPOCO$ = UCase$(TRIM$(.Recordset.Fields(4)))
       LETRAC$ = Left$(TRIM$(.Recordset.Fields(5)), 1)
       PREFIJ$ = TRIM$(.Recordset.Fields(6))
         While Len(PREFIJ$) < 4: PREFIJ$ = "0" + PREFIJ$: Wend
         PREFIJ$ = Right$(PREFIJ$, 4)
       NUMECO$ = TRIM$(.Recordset.Fields(7))
         While Len(NUMECO$) < 8: NUMECO$ = "0" + NUMECO$: Wend
         NUMECO$ = Right$(NUMECO$, 8)
       CODIMLAR$ = TIPOCO$ + "-" + LETRAC$ + "-" + PREFIJ$ + "-" + NUMECO$
       CODIMCOR$ = TIPOCO$ + "-" + Mid$(NUMECO$, 3)
       '
       IMORIGI# = 0: IMPENDE# = 0
       On Error Resume Next
       IMORIGI# = CDbl(.Recordset.Fields(8))
       IMPENDE# = CDbl(.Recordset.Fields(9))
       On Error GoTo 0
       '
2      If IsNumeric(CODICLIE$) = True Then
          If XVALO(CODICLIE$) > 0 Then
            If XVALO(CODICLIE$) <= 32767 Then
              NUCLIE% = XVALO(CODICLIE$)
              If STATCLI%(NUCLIE%) <> 0 Then
                CODICLIE$ = CODICLI$((-1) * NUCLIE%)
                GoTo 15
              End If
            End If
          End If
       End If
       NUCLIE% = NUMEPRO%(CODICLIE$)
       If NUCLIE% < 1 Then
         If COMALTER%("Proveedor " + TRIM$(REFEMOVI$) + " No Registrado.\Desea Darlo de Alta ?\\Si\No") = 1 Then
           Call APERBASDAT("PROVE")
           NROPRO& = 0
           '
3          NROPRO& = NROPRO& + 1
           QTR = CantRegistros("PROVED12", "Nume_Cli = " & NROPRO&)
           If QTR > 0 Then GoTo 3
           '
           NCLIE = NROPRO&
           Call APERBASDAT("PROVE")
           With Proved12
               .AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
               !CodiEmpr = Codiemp%
               ON ERROR GOTO 0
               !Nume_Cli = TRIM$(Str$(NCLIE))
               !Raso_Cli = REFEMOVI$
               !Codi_Cli = CODICLIE$
               !Cuit_Cli = "00-00000000-0"
               !Piva_Cli = 1
               !Stat_Cli = 1
               !Tcpr_Cli = 0
               '
               On Error Resume Next
               .Update
               If Err Then
                  Call MENSERR(24, "Alta no Procesada Correctamente")
               End If
               On Error GoTo 0
           End With
           GoTo 2
         End If
         '
         If CODICLIE$ <> CODICLIA$ Then
           CODICLIA$ = CODICLIE$
           MERROX$ = "Falta Proveedor '" + TRIM$(CODICLIE$) + " - " + REFEMOVI$ + "' en B.Datos Proveedores"
           List1.AddItem MERROX$
         End If
         GoTo 19
       End If
       '
       ' VALIDAR QUE NO IMPORTE DOS VECES EL MISMO COMPROBANTE
       '
15     TTXX$ = Space$(254)
       Call REPLA(TTXX$, CODICLIE$, 1, 12)
       Call REPLA(TTXX$, FORMATNUM$(NUCLIE%, "F6,0"), 13, 6)
       Call REPLA(TTXX$, REFEMOVI$, 21, 30)
       Call REPLA(TTXX$, FECHEMIS$, 53, 8)
       Call REPLA(TTXX$, FechVenc$, 63, 8)
       Call REPLA(TTXX$, CODIMLAR$, 75, 20)
       Call REPLA(TTXX$, CODIMCOR$, 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMORIGI#, "F13.2"), 107, 13)
       Call REPLA(TTXX$, FORMATNUM$(IMPENDE#, "F13.2"), 120, 13)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub
'
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0429-WAR-28/06/05///
Sub PREIMPORCARCHE()
  '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   DoEvents
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'CH' "
   
   'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Call ABRECONEXION
   '
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(290)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0                     ' BANCO
     BANC$ = TRIM$(DBGrid1.TEXT)
     If Len(BANC$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Banco"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, BANC$, 1, 18)
     '
     DBGrid1.Col = 1
     NUCHE$ = TRIM$(DBGrid1.TEXT)   ' NUMERO DE CHEQUE
     If Len(NUCHE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cheque no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, NUCHE$, 21, 12)
     '
     DBGrid1.Col = 2                     'FECHA EMISION
     FECHEMIS$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHEMIS$, 35, 8)
     '
     DBGrid1.Col = 3                     'FECHA VTO.
     FECHEVEN$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHEVEN$, 45, 8)
     '
     DBGrid1.Col = 4                     'IMPORTE ORIGINAL
     IMPO1# = CDbl(TRIM$(DBGrid1.TEXT))
     IMPO$ = FORMATNUM(IMPO1#, "F13.2")
     IMPO2$ = TRIM$(IMPO$)
     Call REPLA(TTXX$, IMPO$, 55, 13)
     '
     DBGrid1.Col = 5                     ' DOCUMENTO ORIGEN LARGO
     DOCU$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOCU$, 71, 18)
     '
     DBGrid1.Col = 6                     ' DOCUMENTO ORIGEN CORTO
     DOCUC$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOCUC$, 91, 12)
     '
     DBGrid1.Col = 7                     ' CLIENTE
     CLI1$ = TRIM$(DBGrid1.TEXT)
       NCLX% = NUMECLI%(CLI1$)
       If NCLX% > 0 Then
         CLI1$ = RASOCLI$(NCLX%)
       End If
     Call REPLA(TTXX$, CLI1$, 105, 30)
       '
     DBGrid1.Col = 8                     ' CUIT LIBRADOR
     CUITLIB$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUITLIB$, 137, 16)
       '
       IDEVA$ = Left$(TTXX$, 32)
       IDEVA$ = IDEVA$ + Mid$(TTXX$, 45, 8)
     SQLX$ = "DescriMov = '" & IDEVA$ & "' "
     SQLX$ = SQLX$ & "AND ABS(IDebeComp - " & IMPO2$ & ") < 0.005 "
'     CuecoTemp.FindFirst SQLX$
'     If CuecoTemp.NoMatch = False Then
'       MERROX$ = "Registro" + Str$(U& - 1) + " - Cheque " + IDEVA$ + " - $ " + TRIM$(IMPO$) + " Existente en Cartera"
'       List1.AddItem MERROX$
'       GoTo 19
'     End If
     QTR& = CantRegistros("CAJABANC", SQLX$)
     If QTR& > 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cheque " + IDEVA$ + " - $ " + TRIM$(IMPO$) + " Existente en Cartera"
       List1.AddItem MERROX$
       GoTo 19
     End If
     
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&
   '
20 DBGrid1.RowHeight = hrow
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0429-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub IMPORCUECLI()
   '
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   Call SELECHO("ECUECON/Cuenta Contra-Partida")
   CONTRAPARTIDA$ = TRIM$(VALACT1$("ECUECON"))
   CONTRAPA% = CUEINT%(CONTRAPARTIDA$)
   If CONTRAPA% < 1 Then Exit Sub
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
     If TRIM$(Mid$(TTXX$, 75, 20)) <> CODIMLARANT$ Or TRIM$(Left$(TTXX$, 10)) <> CLIANTE$ Then
        If Abs(SUMACU#) >= 0.005 Then
             ' GRABA MOVIMIENTO DE CONTRAPARTIDA
           CuentaCo% = CONTRAPA%
           TiMovi$ = "IMPORT"
           OpMovi$ = "CCDEU"
           VaMovi$ = ""
           Referen$ = "CPart.Importa " + CODIMLARANT$
           Observa$ = "CPart.Importa " + CODIMLARANT$
           ImpoDeb# = (-1) * SUMACU#
           Call GRACABAN
        End If
        CODIMLARANT$ = TRIM$(Mid$(TTXX$, 75, 20))
        CLIANTE$ = TRIM$(Left$(TTXX$, 10))
        SUMACU# = 0
     End If
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 95, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 75, 20))
     CodiEmp% = 0
     CodiSuc% = 0
     NumeClie% = XVALO(Mid$(TTXX$, 13, 6))
     NumeProv% = 0
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
       NC% = XVALO(Mid$(TTXX$, 13, 6))
     Referen$ = RASOCLI$(NC%)
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     CuentaCo% = CUECOCLI%(NC%)
     TiMovi$ = "FVEN"
       '
       OPMOV$ = UCase$(TRIM$(Mid$(TTXX$, 95, 2)))
       Select Case (OPMOV$)
        Case "FC"
           TICOMX$ = "FC": OPMOVIX$ = "FAC"
        Case "RB"
           TICOMX$ = "RB": OPMOVIX$ = "CCORR"
        Case "ND"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
        Case "NC"
       End Select
       '
     OpMovi$ = OPMOV$
     VaMovi$ = Mid$(CoComLar$, 4, 1)
     Descrip$ = CoComLar$
     NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
     NumeCom = XVALO(Mid$(CoComLar$, 11))
     StatInit = 8         ' movimiento importado
     ImpoDeb = XVALO(Mid$(TTXX$, 107, 13))
        SUMACU# = SUMACU# + ROUND(ImpoDeb)
     Observa$ = RASOCLI$(NC%)
     Call GRACABAN
     '
     YACANCE = XVALO(Mid$(TTXX$, 107, 13)) - XVALO(Mid$(TTXX$, 120, 13))
     If Abs(YACANCE) >= 0.005 Then
        CoComCor$ = "AJ" + Mid$(CoComCor$, 3)
        CoComLar$ = "AJ" + Mid$(CoComLar$, 3)
        TiMovi = "AJUD"            ' AJUSTE DE SALDO PROVEEDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "CCORR"           ' CUENTA CORRIENTE
        Descrip = "Aj.Saldo - " + CoComLar$
        NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
        NumeCom = XVALO(Mid$(CoComLar$, 11))
        StatInit = 8         ' movimiento importado
        ImpoDeb = (-1) * YACANCE
        SUMACU# = SUMACU# + ROUND(ImpoDeb)
        Observa$ = RASOCLI$(NC%)
        Call GRACABAN
     End If
     '
     IBruTot# = XVALO(Mid$(TTXX$, 107, 13))
     IIvaTot# = 0
     NuorVen% = 0
     FeVenc = CVDAT(FECHANUM(Mid$(TTXX$, 63, 8)))
     ImVenc# = XVALO(Mid$(TTXX$, 120, 13))
     Call GRADEUPEN
     '
   Next U&
   '
   If Abs(SUMACU#) >= 0.005 Then
        ' GRABA MOVIMIENTO DE CONTRAPARTIDA
      CuentaCo% = CONTRAPA%
      TiMovi$ = "IMPORT"
      OpMovi$ = "CCDEU"
      VaMovi$ = ""
      Referen$ = "CPart.Importa " + CODIMLARANT$
      Observa$ = "CPart.Importa " + CODIMLARANT$
      ImpoDeb# = (-1) * SUMACU#
      Call GRACABAN
   End If
   CODIMLARANT$ = TRIM$(Mid$(TTXX$, 75, 20))
   SUMACU# = 0
   '
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub IMPORCUEPRO()
 '
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   Screen.MousePointer = 1
   Call SELECHO("ECUECON/Cuenta Contra-Partida")
   CONTRAPARTIDA$ = TRIM$(VALACT1$("ECUECON"))
   CONTRAPA% = CUEINT%(CONTRAPARTIDA$)
   If CONTRAPA% < 1 Then Exit Sub
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
     If TRIM$(Mid$(TTXX$, 75, 20)) <> CODIMLARANT$ Or TRIM$(Left$(TTXX$, 10)) <> CLIANTE$ Then
        If Abs(SUMACU#) >= 0.005 Then
             ' GRABA MOVIMIENTO DE CONTRAPARTIDA
           CuentaCo% = CONTRAPA%
           TiMovi$ = "IMPORT"
           OpMovi$ = "CCPRO"
           VaMovi$ = ""
           Referen$ = "CPart.Importa " + CODIMLARANT$
           Observa$ = "CPart.Importa " + CODIMLARANT$
           ImpoDeb# = (-1) * SUMACU#
           Call GRACABAN
        End If
        CODIMLARANT$ = TRIM$(Mid$(TTXX$, 75, 20))
        CLIANTE$ = TRIM$(Left$(TTXX$, 10))
        SUMACU# = 0
     End If
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 95, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 75, 20))
     CodiEmp% = 0
     CodiSuc% = 0
     NumeClie% = 0
     NumeProv% = XVALO(Mid$(TTXX$, 13, 6))
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
       NC% = XVALO(Mid$(TTXX$, 13, 6))
     Referen$ = RASOCLI$(NC% * (-1))
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     CuentaCo% = CUECOCLI%(NC% * (-1))
     TiMovi$ = "FCOM"
       '
       SIG% = 1
       OPMOV$ = UCase$(TRIM$(Mid$(TTXX$, 95, 2)))
       Select Case (OPMOV$)
        Case "FC"
           TiMovi$ = "FCOM"
           OPMOVIX$ = "FC"
        Case "OP"
           TiMovi$ = "OPAG"
           OPMOVIX$ = "OP"
           SIG% = (-1)
        Case "ND"
           OPMOVIX$ = "ND"
        Case "NC"
           OPMOVIX$ = "NC"
           SIG% = (-1)
       End Select
       '
     OpMovi$ = OPMOVIX$
     VaMovi$ = Mid$(CoComLar$, 4, 1)
     Descrip$ = CoComLar$
     NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
     NumeCom = XVALO(Mid$(CoComLar$, 11))
     StatInit = 8             ' movimiento importado
     IORI = SIG% * Abs(XVALO(Mid$(TTXX$, 107, 13)))
     ImpoDeb = (-1) * IORI
        SUMACU# = SUMACU# + ROUND(ImpoDeb)
     Observa$ = RASOCLI$(NC% * (-1))
     Call GRACABAN
     '
     IBruTot# = Abs(XVALO(Mid$(TTXX$, 107, 13)))
     IIvaTot# = 0
     NuorVen% = 0
     FeVenc = CVDAT(FECHANUM(Mid$(TTXX$, 63, 8)))
     ImVenc# = (XVALO(Mid$(TTXX$, 120, 13)))
     If Abs(ImVenc#) >= 0.005 Then
        Call GRACREPEN
     End If
     '
     YACANCE = IORI - XVALO(Mid$(TTXX$, 120, 13))
     If Abs(YACANCE) >= 0.005 Then
        CoComCor$ = "AS" + Mid$(CoComCor$, 3)
        CoComLar$ = "AS" + Mid$(CoComLar$, 3)
        TiMovi = "AJUP"            ' AJUSTE DE SALDO PROVEEDORES
        OpMovi = "AS"              ' AJUSTE DE SALDO
        VaMovi = "CCORR"           ' CUENTA CORRIENTE
        Descrip = "Aj.Saldo - " + CoComLar$
        NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
        NumeCom = XVALO(Mid$(CoComLar$, 11))
        StatInit = 8         ' movimiento importado
        ImpoDeb = YACANCE
        SUMACU# = SUMACU# + ROUND(ImpoDeb)
        Observa$ = RASOCLI$(NC%)
        Call GRACABAN
     End If
     '
   Next U&
   '
   If Abs(SUMACU#) >= 0.005 Then
        ' GRABA MOVIMIENTO DE CONTRAPARTIDA
      CuentaCo% = CONTRAPA%
      TiMovi$ = "IMPORT"
      OpMovi$ = "CCDEU"
      VaMovi$ = ""
      Referen$ = "CPart.Importa " + CODIMLARANT$
      Observa$ = "CPart.Importa " + CODIMLARANT$
      ImpoDeb# = (-1) * SUMACU#
      Call GRACABAN
   End If
   CODIMLARANT$ = TRIM$(Mid$(TTXX$, 75, 20))
   SUMACU# = 0
   '
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///

Sub IMPORCARCHE()
   '
   CUCACHE% = CUEICARCH%
   If CUCACHE% < 1 Then
     Call MENSERR(24, "Importación Imposible.\Falta Definir Cuanta Contable\Asociada a la Cartera de Valores")
     CANCELIMP% = 1
     Exit Sub
   End If
   '
   Call SELECHO("ECUECON/Seleccione Cuenta de Contrapartida")
   CONTRAPA$ = TRIM$(VALACT1$("ECUECON"))
   CUCONTRA% = CUEINT%(CONTRAPA$)
   If CONTRAPA$ = "" Or CUCONTRA% < 1 Then
     Call MENSERR(24, "Importación Cancelada.\Debe Definir una Cuenta de Contra Partida\Para el Asiento Implícito")
     CANCELIMP% = 1
     Exit Sub
   End If
   
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   '
   TOTAVAL# = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 91, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 71, 18))
     CodiEmp% = 0
     CodiSuc% = SUCUI%
     NumeClie% = 0
     NumeProv% = 0
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 35, 8)))
     Referen$ = Mid$(TTXX$, 105, 30)
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 35, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 45, 8)))
     CuentaCo% = CUCACHE%
     TiMovi$ = "CHRECI"
     OpMovi$ = "CH"
     VaMovi$ = ""
       IDEVA$ = Left$(TTXX$, 32)
       IDEVA$ = IDEVA$ + Mid$(TTXX$, 45, 8)
       IDEVB$ = TRIM$(Left$(IDEVA$, 20)) + " " + TRIM$(Mid$(IDEVA$, 21, 12))
     Descrip$ = IDEVA$
     NumeTal = 0
     NumeCom = XVALO(Mid$(TTXX$, 94, 9))
     StatInit = 1
     ImpoDeb# = XVALO(Mid$(TTXX$, 55, 13))
       TOTAVAL# = TOTAVAL# + ImpoDeb#
     Observa$ = ""
       CUITLIB$ = TRIM$(Mid$(TTXX$, 137, 16))
       If CUITLIB$ <> "" Then
         Observa$ = "CUIT Libr: " + CUITLIB$
       End If
       '
     Call GRACABAN
     '
     ' graba la contra-partida
     Referen$ = "Importacion " + Left$(IDEVB$, 32)
     CuentaCo% = CUCONTRA%
     TiMovi$ = "IMPOR"
     OpMovi$ = "CARCHE"
     VaMovi$ = ""
     Descrip$ = "Importacion " + Left$(IDEVB$, 32)
     StatInit = 0
     ImpoDeb# = (-1) * XVALO(Mid$(TTXX$, 55, 13))
       '
     Call GRACABAN

   Next U&
   '
   Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORMOVSTK()
   '
   '\\\OT-5-469-FG-05/07/05///
   If DBGrid1.Columns.Count < 10 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(254)
        CODEX$ = ""
        On Error Resume Next
        CODEX$ = TRIM$(.Recordset.Fields(0))   ' TRIM$(DBGrid1.TEXT)
        On Error GoTo 0
        If CODEX$ = "" Or Len(CODEX$) > 15 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
12      Call REPLA(TTXX$, CODEX$, 1, 16)
        '
13      CDXI% = CODINT%(CODEX$)
        If CDXI% < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo '" + CODEX$ + "' Inexistente o No Valido"
          If COMALTER%(MERROX$ + "\Desea Dar de Alta Este Código ?\\Si\No") = 1 Then
            VDEF$ = String$(128, 0)
            Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
            Call REPLA(VDEF$, "U.", 109, 2)
            Call REPLA(VDEF$, Chr$(1), 93, 1)
            Call REPLA(VDEF$, CODEX$, 77, 16)
            DEDEX$ = ""
            On Error Resume Next
            DEDEX$ = TRIM$(.Recordset.Fields(1))
            On Error GoTo 0
            Call REPLA(VDEF$, DEDEX$, 5, 64)
            Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
            If OKX% = 1 Then GoTo 13
          End If
          '
          List1.AddItem MERROX$
          GoTo 19
        End If
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
        If DEDEX$ = "" Then DEDEX$ = DESCRIT0$(CDXI%)
        Call REPLA(TTXX$, DEDEX$, 25, 64)
        '
        On Error Resume Next
        FECHMOVX$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
        On Error GoTo 0
        If FECHANUM(FECHMOVX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Fecha de Movimiento No Válida"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, FECHMOVX$, 91, 8)
        '
        DEPX = 1 ' dEPóSITO
        On Error Resume Next
        DEPX = CDbl(.Recordset.Fields(3))
        On Error GoTo 0
        If DEPX < 1 Or DEPX > 255 Then DEPX = 1
        VALO$ = FORMATNUM$(DEPX, "F4.0")
        Call REPLA(TTXX$, VALO$, 101, 4)
        
        CANTI = 0 ' CANTIDAD
        On Error Resume Next
        CANTI = CDbl(.Recordset.Fields(4))
        On Error GoTo 0
        CANTIX$ = FORMATNUM$(CANTI, "F12.1")
        Call REPLA(TTXX$, CANTIX$, 105, 12)
        CANTIX$ = TRIM$(CANTIX$)
        '
        CODMOVI$ = ""
        On Error Resume Next
        CODMOVI$ = TRIM$(.Recordset.Fields(5))
        On Error GoTo 0
        If Len(CODMOVI$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código de Movimiento No Válido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, CODMOVI$, 119, 2)
        '
        DOCPRI$ = ""
        On Error Resume Next
        DOCPRI$ = TRIM$(.Recordset.Fields(6))
        On Error GoTo 0
        If Len(DOCPRI$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Documento Primario"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DOCPRI$, 123, 12)
        '
        DOCSEC$ = ""
        On Error Resume Next
        DOCSEC$ = TRIM$(.Recordset.Fields(7))
        On Error GoTo 0
        Call REPLA(TTXX$, DOCSEC$, 137, 12)
        
        VALUNI = 0
        On Error Resume Next
        VALUNI = CDbl(.Recordset.Fields(8))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALUNI, "F14.4")
        Call REPLA(TTXX$, VALO$, 149, 14)
        '
        MONX = 1 ' MONEDA
        On Error Resume Next
        MONX = CDbl(.Recordset.Fields(9))
        On Error GoTo 0
        If MONX < 1 Or MONX > 4 Then MONX = 1
        VALO$ = FORMATNUM$(MONX, "F4.0")
        Call REPLA(TTXX$, VALO$, 163, 4)
        '
        CLIX$ = ""
        NPROV% = 0
        On Error Resume Next
        CLIX$ = TRIM$(CDbl(.Recordset.Fields(10)))
        On Error GoTo 0
        If CLIX$ = "" Then
            NCLIE = 0
          Else
            NCLIE = NUMECLI%(CLIX$)
            If NCLIE < 1 Then
              NCLIE = XVALO(CLIX$)
            End If
        End If
        '
        If NCLIE < 1 Then
          CLIX$ = ""
          NPROV% = 0
          On Error Resume Next
          CLIX$ = TRIM$(CDbl(.Recordset.Fields(11)))
          On Error GoTo 0
          If CLIX$ = "" Then
              NPROV% = 0
            Else
              NPROV% = NUMEPRO%(CLIX$)
              If NPROV% < 1 Then
                NPROV% = XVALO(CLIX$)
              End If
          End If
        End If
        '
        VALO$ = FORMATNUM$(NCLIE, "F6.0")
        Call REPLA(TTXX$, VALO$, 167, 6)
        '
        VALO$ = FORMATNUM$(NPROV%, "F6.0")
        Call REPLA(TTXX$, VALO$, 173, 6)
        '
        DETALLE$ = ""
        On Error Resume Next
        DETALLE$ = TRIM$(.Recordset.Fields(12))
        On Error GoTo 0
        If Len(DETALLE$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Detalle / Referencia del Movimiento"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DETALLE$, 181, 32)
        '
        FEXI% = FECHANUM(FECHMOVX$)
        FEMOVD = FETEXCOR1$(FEXI%)
        If Abs(XVALO(CANTIX$)) < 0.01 Then CANTIX$ = "0"
        Call APERBASDAT("STOCKS")
        CONDI$ = "Cod_Inte = " & CDXI% & " "
        CONDI$ = CONDI$ + "AND DepoMovi = " & DEPX & " "
        CONDI$ = CONDI$ + "AND DoPriCor = '" & DOCPRI$ & "' "
        CONDI$ = CONDI$ + "AND ABS((CantIngre - CantSalid) - " & CANTIX$ & ") < 0.05 "
        CONDI$ = CONDI$ + "AND FechMov = '" & FEMOVD & "'"
        '
        QTR& = CantRegistros("MOVISTO", CONDI$)
        If QTR& > 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Movimiento Pre-Existente en Base de Datos."
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$) + "*") > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$) + "*"
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
   '\\\OT-5-469-FG-FIN///
End Sub

Sub IMPORMOVSTK()
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   Call APERBASDAT("STOCKS")
   NOCLOSE% = 1
   '
   NMS% = NUMAS%
   'REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = XVALO(Mid$(TTXX$, 17, 6))
     FEMO% = FECHANUM(Mid$(TTXX$, 91, 8))
     LOTE$ = ""
     CMOV$ = TRIM$(Mid$(TTXX$, 119, 2))
     DOPRILA$ = TRIM$(Mid$(TTXX$, 123, 12)): DOPRICO$ = DOPRILA$
     DOSELA$ = TRIM$(Mid$(TTXX$, 137, 12)): DOSECO$ = DOSELA$
     REFE$ = TRIM$(Mid$(TTXX$, 181, 32))
     NUORIT% = 0
     VUNI# = XVALO(Mid$(TTXX$, 149, 14))
     MONEII% = XVALO(Mid$(TTXX$, 163, 4))
     NC% = XVALO(Mid$(TTXX$, 167, 6)) - XVALO(Mid$(TTXX$, 173, 6))
     DEPO% = XVALO(Mid$(TTXX$, 101, 4))
     CANTI = XVALO(Mid$(TTXX$, 105, 12))
     '
     Call MOVISTOK(FEMO%, CIXI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
     '
   Next U&
   '
   Movisto.Close
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub PREIMPORLISPRE()
   '
   '\\\OT-5-0641-FG-26/09/05///
   '
   If DBGrid1.Columns.Count < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(120)
        '
        On Error Resume Next                    ' NUMERO DE LISTA
        NROLISTA = XVALO(.Recordset.Fields(0))
        On Error GoTo 0
        VALO$ = FORMATNUM$(NROLISTA, "F4.0")
        Call REPLA(TTXX$, VALO$, 1, 4)
        '
        CODEX$ = ""                             ' CODIGO DE ARTICULO
        On Error Resume Next
        CODEX$ = .Recordset.Fields(1)
        On Error GoTo 0
        '
        CDXI% = CODINT%(CODEX$)
        If CDXI% < 1 Then
           MERROX$ = "Registro" + Str$(U& - 1) + " - Código no Existente en el Maestro de Artículos"
           List1.AddItem MERROX$
           GoTo 19
        End If
        '
        Call REPLA(TTXX$, CODEX$, 6, 16)
        '
        Call REPLA(TTXX$, DESCRIT0$(CDXI%), 24, 64) ' DESCRIPCION DEL ARTICULO
        '
        NUMORIG$ = ""                               ' NUMERO ORIGINAL
        On Error Resume Next
        NUMORIG$ = TRIM$(.Recordset.Fields(2))
        On Error GoTo 0
        Call REPLA(TTXX$, NUMORIG$, 90, 16)
        '
        On Error Resume Next                        ' PRECIO UNITARIO
        VALOR = CDbl(.Recordset.Fields(3))
        On Error GoTo 0
        If Not (VALOR > 0) Then
        If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Registro" + Str$(U& - 1) + " - Artículo '" + TRIM$(CODEX$) + "' sin Precio\\Saltear\Saltear Siempre\Incluir\Incluir Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      VALO$ = FORMATNUM$(VALOR, "F12.4")
        Call REPLA(TTXX$, VALO$, 108, 12)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
   '\\\OT-5-0641-FG-FIN///
End Sub
Sub PREIMPORCOSREP()
'\\\OT-06-0186-WAR-24/05/06///
    '
   If DBGrid1.Columns.Count < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   HUBOAMAS% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(120)
        '
        CODEX$ = ""                             ' CODIGO DE ARTICULO
        On Error Resume Next
        CODEX$ = .Recordset.Fields(0)
        On Error GoTo 0
        '
5       CDXI% = CODINT%(CODEX$)
        If CDXI% < 1 Then
          TTXX1$ = "Codigo de Material '" + TRIM$(CODEX$) + "' Inexistente en B.Datos."
          TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
          If COMALTER%(TTXX1$) = 1 Then
            VDEF$ = String$(128, 0)
            Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
            Call REPLA(VDEF$, "U.", 109, 2)
            Call REPLA(VDEF$, Chr$(1), 93, 1)
            Call REPLA(VDEF$, CODEX$, 77, 16)
            DEDEX$ = ""
            Call REPLA(VDEF$, DEDEX$, 5, 64)
            Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
            HUBOAMAS% = 1
            If OKX% = 1 Then GoTo 5
          End If
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código '" + TRIM$(CODEX$) + "' no Existente en el Maestro de Artículos"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
        If COSUNI(CDXI%) > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código '" + CODEX$ + "' con Costo Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      CODEX$ = CODEXT$(CDXI%)
        Call REPLA(TTXX$, CODEX$, 1, 16)
        '
        Call REPLA(TTXX$, DESCRIT0$(CDXI%), 19, 48) ' DESCRIPCION DEL ARTICULO
        '
        On Error Resume Next                    ' MONEDA
        NROMON = XVALO(.Recordset.Fields(1))
        On Error GoTo 0
        VALO$ = FORMATNUM$(NROMON, "F4.0")
        Call REPLA(TTXX$, VALO$, 67, 4)
        '
        VALOD# = 0
        On Error Resume Next                        ' COSTO UNITARIO
        VALOD# = CDbl(.Recordset.Fields(2))
        On Error GoTo 0
        If VALOD# < 0.000005 Then GoTo 19        ' SALTEA SIN VALOR
        '
        VALO$ = FORMATNUM$(VALOD#, "F16.5")
        Call REPLA(TTXX$, VALO$, 71, 16)
        '
        FECHA1$ = FECHATEX$(CVINT(.Recordset.Fields(3))) '
        Call REPLA(TTXX$, FECHA1$, 89, 8)
        '
        NPRO1$ = ""                           'PROVEEDOR
        On Error Resume Next
        NPRO1$ = TRIM$(.Recordset.Fields(4))
        On Error GoTo 0
        NPRO2% = NUMEPRO%(NPRO1$)
        If NPRO2% < 1 Then
          On Error Resume Next
          NPR02% = Val(NPRO1$)
        End If
        '
        VALO$ = FORMATNUM$(NPRO2%, "F6.0")
        Call REPLA(TTXX$, VALO$, 97, 6)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
   If HUBOAMAS% > 0 Then    ' HUBO CODIGOS AGREGADOS
     UPDATMASTER% = 1
     Call GENECOMAS
   End If
   '
20 Screen.MousePointer = 1
   '
'\\\OT-06-0186-WAR-FIN///
End Sub

'\\\OT-06-0423-MH-05/10/06///
Sub PREIMPORTRANS()
'
   If DBGrid1.Columns.Count < 11 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0               'Código de Transportista
     '
     
     CODTR& = XVALO(DBGrid1.TEXT)
     If CODTR& < 0 Or CODTR& > 9999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Número de Transportista Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     CODTR1% = CODTR&
     CODTS$ = FORMATNUM$(CODTR&, "I4")
     RFF$ = RECFILT$("TRANSPOR", 1, MKI(CODTR1%))
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODTS$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODTS$, 1, 4)
     '
     DBGrid1.Col = 1             'Nombre Transportista
     NOMTR$ = TRIM$(DBGrid1.TEXT)
     If Len(NOMTR$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Nombre del Transportista"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, NOMTR$, 7, 32)
     '
     DBGrid1.Col = 2              ' Razón Social
     RAZOC$ = TRIM$(DBGrid1.TEXT)
     If Len(RAZOC$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Razón Social del Transportista"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, RAZOC$, 42, 48)
     '
     DBGrid1.Col = 3              ' Domicilio transportista
     DOMIT$ = TRIM$(DBGrid1.TEXT)
     If DOMIT$ = "" Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Domicilio del Transportista"
        List1.AddItem MERROX$
        GoTo 19
     End If
     Call REPLA(TTXX$, DOMIT$, 93, 48)
     '
     DBGrid1.Col = 4              ' Localidad transportista
     LOCTR$ = TRIM$(DBGrid1.TEXT)
     If LOCTR$ = "" Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Localidad del Transportista"
        List1.AddItem MERROX$
        GoTo 19
     End If
     Call REPLA(TTXX$, LOCTR$, 144, 48)
     '
     DBGrid1.Col = 5              ' Posición Iva transportista
     POSIV& = XVALO(DBGrid1.TEXT)
     POSIS$ = FORMATNUM(POSIV&, "I1")
     Call REPLA(TTXX$, POSIS$, 195, 1)
     '
     DBGrid1.Col = 6               ' CUIT Transportista
     CUITR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUITR$, 205, 15)
     '
     DBGrid1.Col = 7                ' Teléfono transportista
       TELTR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELTR$, 220, 24)
     '
     DBGrid1.Col = 8                 ' Fax transportista
       FAXTR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAXTR$, 247, 24)
     '
     DBGrid1.Col = 9               ' E-MAIL Transportista
       MAITR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, MAITR$, 274, 48)
     '
     DBGrid1.Col = 10               ' Persona de Contacto
       CONTR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTR$, 322, 48)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   
20 End Sub
'\\\OT-06-0423-MH-FIN///

Sub PREIMPORSUCUR()
'
   If DBGrid1.Columns.Count < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 2               'Código de SUCURSAL
     '
     CODSUCR& = XVALO(DBGrid1.TEXT)
     '
     
     If CODSUCR& < 0 Or CODSUCR& > 9999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Número de Sucursal Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     
     DBGrid1.Col = 0
     NUMCLIE% = XVALO(DBGrid1.TEXT) ' NUMERO DE CLIENTE
     DBGrid1.Col = 2
     CODSU& = XVALO(DBGrid1.TEXT) 'NUMERO DE SUCURSAL
     Call FILTERFILE("SUCENTRE", "", 5, MKI$(NUMCLIE%))    ' AQUI HAY QUE CARGAR EL NUMERO DE CLIENTE
     URE& = ULTREG&("!SUCENTRE")
     For I& = 1 To URE&
       TX$ = REGLEIDO$("!SUCENTRE", I&)
       SUCARCH& = CVI(Mid$(TX$, 1, 2)) 'coincide con la del registro recorrido
       If CODSU& = SUCARCH& Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + CODTS$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
       End If
     Next I&

     '
     CODTR1% = CODSUCR&
     CODTS$ = FORMATNUM$(CODSUCR&, "I4")

12   Call REPLA(TTXX$, CODTS$, 1, 4)
     '
     DBGrid1.Col = 1             'Nombre SUCURSAL
     'NOMTR$ = TRIM$(DBGrid1.TEXT)
     SUCURSAL$ = TRIM$(DBGrid1.TEXT)
     
     If Len(SUCURSAL$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Nombre del Sucursal"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, SUCURSAL$, 6, 29)
     '
     '
     DBGrid1.Col = 3              ' Domicilio Sucursal
     DOMISUC$ = TRIM$(DBGrid1.TEXT)
     If DOMISUC$ = "" Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Domicilio de la Sucursal "
        List1.AddItem MERROX$
        GoTo 19
     End If
     Call REPLA(TTXX$, DOMISUC$, 40, 29)
     
     '
     DBGrid1.Col = 4              ' Localidad
     LOCSUC$ = TRIM$(DBGrid1.TEXT)
     
     If LOCSUC$ = "" Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Localidad de Sucursal"
        List1.AddItem MERROX$
        GoTo 19
     End If
     'Call REPLA(TTXX$, LOCTR$, 144, 48)
     Call REPLA(TTXX$, LOCSUC$, 71, 32)
     
     '
     DBGrid1.Col = 0             ' NUMERO DE CLIENTE
     NUMCLI& = XVALO(DBGrid1.TEXT)
     NUMCLI1$ = FORMATNUM(NUMCLI&, "I6")
     
     Call REPLA(TTXX$, NUMCLI1$, 105, 6)
     
'     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   
20 End Sub

Sub PREIMPORUBIC()
'
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(128)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     On Error GoTo 0
     '
     DBGrid1.Col = 2
     UBICACION$ = TRIM$(DBGrid1.TEXT)
     '
     DBGrid1.Col = 0               'CODIGO DE ARTICULO
     Codigo$ = TRIM$(DBGrid1.TEXT)
     CII% = CODINT%(Codigo$)
     '
     If CII% < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Código Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFLT$ = RECFILT$("UBISTOCK", 1, MKI$(CII%))
     URE& = ULTREG&("UBISTOCK")
     For J& = 1 To Len(RFLT$) Step 4
         RESE& = CVS(Mid$(RFLT$, J&, 4))
         X$ = REGLEIDO$("UBISTOCK", RESE&)
         If CII% = CVI(Mid$(X$, 1, 2)) Then
           If UCase(TRIM$(UBICACION$)) = UCase(TRIM$(Mid$(X$, 3, 24))) Then
            MERROX$ = "Registro" + Str$(U& - 1) + " - Ubicacion Preexistente para Código: " & Codigo$
            List1.AddItem MERROX$
            GoTo 19
           End If
         End If
     Next J&
     '
     Call REPLA(TTXX$, Codigo$, 1, 16)
     '
     DBGrid1.Col = 1             'Descripcion del producto
     Descripcion$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, Descripcion$, 20, 32)
     '
     DBGrid1.Col = 2              ' Ubicacion
     UBICACION$ = TRIM$(DBGrid1.TEXT)
     If Len(UBICACION$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Ubicación"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, UBICACION$, 55, 24)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     
19  Next U&
    '
20  End Sub
'
Sub IMPORTRANS()
'\\\OT-06-0423-MH-05/10/06///
'
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("TRANSPOR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     COTRA% = XVALO(Mid$(TTXX$, 1, 4))    'Código
     NOTRA$ = TRIM$(Mid$(TTXX$, 7, 32))   'Nombre
     RSTRA$ = TRIM$(Mid$(TTXX$, 42, 48))  'Razon Social
     DOTRA$ = TRIM$(Mid$(TTXX$, 93, 48))  'Domicilio
     LOCTR$ = TRIM$(Mid$(TTXX$, 144, 48)) 'Localidad
     POSIV% = XVALO(Mid$(TTXX$, 195, 1))  'Pos IVA
     CUITR$ = TRIM$(Mid$(TTXX$, 205, 15)) 'CUIT
     TETRA$ = TRIM$(Mid$(TTXX$, 220, 24)) 'Teléfono
     FAXTR$ = TRIM$(Mid$(TTXX$, 247, 24)) 'FAX
     EMATR$ = TRIM$(Mid$(TTXX$, 274, 48)) 'Mail
     CONTR$ = TRIM$(Mid$(TTXX$, 322, 48)) 'Contacto
     '
     RETRA$ = REBLA$
     Call REPLA(RETRA$, MKI(COTRA%), 1, 2)    'Código
     Call REPLA(RETRA$, NOTRA$, 3, 32)        'Nombre
     Call REPLA(RETRA$, RSTRA$, 35, 48)       'Razón Social
     Call REPLA(RETRA$, DOTRA$, 83, 48)       'Domicilio
     Call REPLA(RETRA$, LOCTR$, 131, 48)      'Localidad
     Call REPLA(RETRA$, Chr$(POSIV%), 194, 1) 'Pos. Iva
     Call REPLA(RETRA$, CUITR$, 179, 15)      'Cuit
     Call REPLA(RETRA$, TETRA$, 195, 48)      'Teléfono
     Call REPLA(RETRA$, FAXTR$, 243, 48)      'Fax
     Call REPLA(RETRA$, EMATR$, 291, 48)      'E-Mail
     Call REPLA(RETRA$, CONTR$, 339, 48)      'Contacto
     '
     Call FILTERSETRECORD("TRANSPOR", 1, MKI$(COTRA%), RETRA$)
     '
   Next U&
   Call CIERRARCH("TRANSPOR")
   '
End Sub
'\\\OT-06-0423-MH-FIN///

Sub IMPORSUCUR()
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("SUCENTRE")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     A% = 0
     COSUC% = XVALO(Mid$(TTXX$, 1, 4))    'NUMSUCURSAL
     NOSUC$ = TRIM$(Mid$(TTXX$, 6, 29))   'Nombre
     DOSUC$ = TRIM$(Mid$(TTXX$, 40, 29))  'Domicilio
     LOCSUC$ = TRIM$(Mid$(TTXX$, 71, 32)) 'Localidad
     NUMCLI% = XVALO(Mid$(TTXX$, 105, 6))  'NUMERO CLIENTE
     '
     For I& = 1 To ULTREG&("SUCENTRE")
        
        TX$ = REGLEIDO$("SUCENTRE", I&)
            RETRA$ = REBLA$
            Call REPLA(RETRA$, MKI(COSUC%), 1, 2)    'NUMERO SUCURSAL
            Call REPLA(RETRA$, NOSUC$, 3, 62)        'Nombre
            Call REPLA(RETRA$, DOSUC$, 65, 48)       'Domicilio
            Call REPLA(RETRA$, LOSUC$, 113, 48)      'Localidad
            Call REPLA(RETRA$, MKI(NUMCLI%), 161, 2) 'NUMERO CLIENTE
            '
            If NUMCLI% = XVALO(Mid$(TX$, 161, 2)) And COSUC% = XVALO(Mid$(TX$, 1, 2)) Then
              A% = 1
              Call GRAREG("SUCENTRE", RETRA$, I&)
            End If
            If A% = 0 Then Call REGAPP("SUCENTRE", RETRA$)
        A% = 1
     Next I&
   Next U&
   Call CIERRARCH("SUCENTRE")
   '
End Sub

Sub PREIMPORPADEMP()

   PUNUCLI$ = String$(32767, 0)
   '
   If DBGrid1.Columns.Count < 21 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Screen.MousePointer = 11
   ' Establece en el "vector" PUNUCLI$, los numeros de clientes ya utilizado
   '
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(516)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0                      'LEGAJO
     LEGAJO = XVALO(TRIM$(DBGrid1.TEXT))
     If LEGAJO < 1 Or LEGAJO > 9999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Legajo no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     ' Verifica si el numero ya ha sido utilizado
     If Asc(Mid$(PUNUCLI$, LEGAJO, 1)) > 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Legajo Repetido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Mid$(PUNUCLI$, LEGAJO, 1) = Chr$(1)
     NUMLEGA$ = Str$(LEGAJO)
     Call REPLA(TTXX$, NUMLEGA$, 1, 5)
     '
     DBGrid1.Col = 1                    ' NOMBRE Y APELLIDO
     NOMAPE$ = TRIM$(DBGrid1.TEXT)
     If Len(NOMAPE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Nombre y Apellido del Empleado"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, NOMAPE$, 10, 30)
     '
     DBGrid1.Col = 2                     ' GRUPO EMPLEADO
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
'     If VALOR < 1 Or VALOR > 255 Then
'        MERROX$ = "Registro" + Str$(U& - 1) + " - Grupo de Empleado no Valido"
'        List1.AddItem MERROX$
'        On Error GoTo 0
'        GoTo 19
'     End If
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F3.0")
     Call REPLA(TTXX$, VALO$, 43, 3)
     '
     DBGrid1.Col = 3                     ' DOCUMENTO
     DOCUMEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOCUMEN$, 48, 13)
     '
     DBGrid1.Col = 4                     ' SEXO
     SEXO$ = TRIM$(DBGrid1.TEXT)
     If SEXO$ <> "M" Then
      If SEXO$ <> "F" Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Formato Tipo Sexo Incorrecto"
        List1.AddItem MERROX$
        GoTo 19
      End If
     End If
     Call REPLA(TTXX$, SEXO$, 61, 1)
     '
     DBGrid1.Col = 5                     ' FECHA DE NACIMIENTO
     FECHANAC$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHANAC$, 64, 8)
     '
     DBGrid1.Col = 6                     ' NACIONALIDAD
     NACIONA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, NACIONA$, 73, 10)
     '
     DBGrid1.Col = 7                     ' ESTADO CIVIL
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F3.0")
     Call REPLA(TTXX$, VALO$, 86, 3)
     '
     DBGrid1.Col = 8                     ' DOMICILIO
     DOMIEMPL$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOMIEMPL$, 96, 30)
     '
     DBGrid1.Col = 9                     ' CODIGO POSTAL
     CODIPOST$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODIPOST$, 126, 10)
     '
     DBGrid1.Col = 10                    ' LOCALIDAD
     LOCAEMPL$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCAEMPL$, 136, 30)
     '
     DBGrid1.Col = 11                    ' SECTOR
     SECTOR$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, SECTOR$, 167, 4)
     '
     DBGrid1.Col = 12                    ' HORARIO DE TRABAJO
     HORTRAB$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, HORTRAB$, 175, 15)
     '
     DBGrid1.Col = 13                    ' TAREA
     TAREA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TAREA$, 192, 20)
     '
     DBGrid1.Col = 14                    ' CATEGORIA
     CATEMPL$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CATEMPL$, 214, 9)
     '
     DBGrid1.Col = 15                    ' FECHA DE INGRESO
     FECHAING$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHAING$, 225, 8)
     '
     DBGrid1.Col = 16                     ' FECHA ANTIGUEDAD
     FECHANTIGU$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHANTIGU$, 236, 8)
     '
     DBGrid1.Col = 17                    ' C.U.I.L.
     CUILEMP$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUILEMP$, 248, 13)
     '
     DBGrid1.Col = 18                    ' A.F.J.P.
     AFJPEMP$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, AFJPEMP$, 263, 10)
     '
     DBGrid1.Col = 19                    ' FECHA DE EGRESO
     FECHAEGRE$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHAEGRE$, 276, 8)
     '
     DBGrid1.Col = 20                    ' MOTIVO
     MOTIVO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, MOTIVO$, 285, 16)
     '
     DBGrid1.Col = 21                    ' STATUS
     VALOR = 0
     On Error Resume Next
     Status% = XVALO(DBGrid1.TEXT)
     If Status% < 0 Or Status% > 1 Then
        MERROX$ = "Registro" + Str$(U& - 1) + " - Status no Valido"
        List1.AddItem MERROX$
        On Error GoTo 0
        GoTo 19
     End If
     On Error GoTo 0
     VALO$ = FORMATNUM(Status%, "F1.0")
     Call REPLA(TTXX$, VALO$, 304, 1)
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
End Sub

Sub IMPORPADEMP()
    
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADREMP")
   R& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     
     NLEGAJO% = XVALO(Mid$(TTXX$, 1, 5))                    'NUMERO DE LEGAJO EMPLEADO
     NOMAPEL$ = TRIM$(Mid$(TTXX$, 10, 30))                  'NOMBRE Y APELLIDO EMPLEADO
     DOCUMEN$ = TRIM$(Mid$(TTXX$, 48, 13))                  'DOCUMENTO EMPLEADO
     FECINGR% = CVINT(Mid$(TTXX$, 225, 8))                  'FECHA DE INGRESO EMPLEADO
     ESTADO% = XVALO(Mid$(TTXX$, 304, 1))                   'STATUS EMPLEADO
     If ESTADO% = 1 Then
        Status$ = Chr$(ESTADO%)
     ElseIf ESTADO% = 0 Then
        Status$ = Chr$(ESTADO%)
     End If
     GRUPEMP$ = Chr$(XVALO(Mid$(TTXX$, 43, 3)))             'GRUPO EMPLEADO
     SECTOR$ = TRIM$(Mid$(TTXX$, 167, 4))                   'SECTOR EMPLEADO
     CATEGOR$ = TRIM$(Mid$(TTXX$, 214, 9))                  'CATEGORIA EMPLEADO
     FECNACI% = CVINT(Mid$(TTXX$, 64, 8))                   'FECHA DE NACIMIENTO EMPLEADO
     DOMIEMP$ = TRIM$(Mid$(TTXX$, 96, 30))                  'DOMICILIO EMPLEADO
     LOCAEMP$ = TRIM$(Mid$(TTXX$, 136, 30))                 'LOCALIDAD EMPLEADO
     NACIONA$ = TRIM$(Mid$(TTXX$, 73, 10))                  'NACIONALIDAD EMPLEADO
     ESCIVIL% = XVALO(Mid$(TTXX$, 86, 3))                   'ESTADO CIVIL EMPLEADO
     ESTCIVI$ = ECOARCH$("ESTCIV", MKI$(ESCIVIL%))
     FECANTI% = CVINT(Mid$(TTXX$, 236, 8))                  'FECHA ANTIGUEDAD EMPLEADO
     FECEGRE% = CVINT(Mid$(TTXX$, 276, 8))                  'FECHA EGRESO EMPLEADO
     SEXOEMP$ = TRIM$(Mid$(TTXX$, 61, 1))                   'SEXO EMPLEADO
     TAREAEM$ = TRIM$(Mid$(TTXX$, 192, 20))                 'TAREA EMPLEADO
     MOTIEGR$ = TRIM$(Mid$(TTXX$, 285, 16))                 'MOTIVO EGRESO
     CUILEMP$ = TRIM$(Mid$(TTXX$, 248, 13))                 'CUIL EMPLEADO
     COPOSTA$ = TRIM$(Mid$(TTXX$, 126, 10))                 'CODIGO POSTAL EMPLEADO
     HORTRAB$ = TRIM$(Mid$(TTXX$, 175, 15))                 'HORARIO DE TRABAJO EMPLEADO
     AFJPEMP$ = TRIM$(Mid$(TTXX$, 263, 10))                 'AFJP EMPLEADO

     R& = R& + 1
     XX$ = REGBLAN$("PADREMP")
     Call REPLA(XX$, MKI$(NLEGAJO%), 1, 2)
     Call REPLA(XX$, NOMAPEL$, 3, 30)
     Call REPLA(XX$, DOCUMEN$, 33, 13)
     Call REPLA(XX$, MKI$(FECINGR%), 46, 2)
     Call REPLA(XX$, Status$, 48, 1)
     Call REPLA(XX$, GRUPEMP$, 49, 1)
     Call REPLA(XX$, SECTOR$, 50, 4)
     Call REPLA(XX$, CATEGOR$, 54, 9)
     Call REPLA(XX$, MKI$(FECNACI%), 63, 2)
     Call REPLA(XX$, DOMIEMP$, 65, 30)
     Call REPLA(XX$, LOCAEMP$, 95, 30)
     Call REPLA(XX$, NACIONA$, 125, 10)
     Call REPLA(XX$, Str(ESCIVIL%), 135, 10)
     Call REPLA(XX$, MKI$(FECANTI%), 145, 2)
     Call REPLA(XX$, MKI$(FECEGRE%), 147, 2)
     Call REPLA(XX$, SEXOEMP$, 149, 1)
     Call REPLA(XX$, TAREAEM$, 161, 20)
     Call REPLA(XX$, MOTIEGR$, 181, 16)
     Call REPLA(XX$, CUILEMP$, 209, 13)
     Call REPLA(XX$, COPOSTA$, 222, 10)
     Call REPLA(XX$, HORTRAB$, 232, 15)
     Call REPLA(XX$, AFJPEMP$, 247, 10)
     
     Call GRAREG("PADREMP", XX$, R&)
     
   Next U&
     
   Call SETULTREG("PADREMP", R&)
   Call CIERRARCH("PADREMP")
     
End Sub

Sub IMPORUBIC()

   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("UBISTOCK")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODI$ = TRIM$(Mid$(TTXX$, 1, 16))
     CII% = CODINT(CODI$)
     UBICA$ = TRIM$(Mid$(TTXX$, 55, 24))
     '
     Call REPLA(REUBI$, MKI$(CII%), 1, 2)
     Call REPLA(REUBI$, UBICA$, 3, 24)
     Call REGAPP("UBISTOCK", REUBI$)
     '
   Next U&
   Call CIERRARCH("UBISTOCK")

End Sub
