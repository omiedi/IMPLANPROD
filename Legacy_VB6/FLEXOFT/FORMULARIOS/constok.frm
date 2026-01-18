VERSION 5.00
Begin VB.Form CONSTOK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stocks - Control General"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10845
   Icon            =   "constok.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9765
      TabIndex        =   59
      ToolTipText     =   "Recuperación Movimientos Perdidos"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Complementos"
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
      Left            =   5040
      TabIndex        =   45
      Top             =   5145
      Width           =   4530
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   51
         Top             =   525
         Width           =   1695
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
            Height          =   600
            Left            =   0
            Picture         =   "constok.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Carga y Consulta de Pedidos de Clientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Pedidos    de Clientes"
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
            TabIndex        =   53
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1740
         TabIndex        =   46
         Top             =   525
         Width           =   1800
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
            Left            =   0
            Picture         =   "constok.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Parte Diario de Fabricacion por Impresora"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Partes de Fabricación"
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
            TabIndex        =   47
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   105
            Width           =   1110
         End
      End
   End
   Begin VB.CommandButton Command27 
      Caption         =   "More ..."
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
      Left            =   9765
      Picture         =   "constok.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Complementos, Estadísticas, Despachos de Importación"
      Top             =   4095
      Width           =   855
   End
   Begin VB.CommandButton Command26 
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
      Left            =   9765
      Picture         =   "constok.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2310
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
      Left            =   9765
      Picture         =   "constok.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1365
      Width           =   855
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
      Left            =   9765
      Picture         =   "constok.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   420
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Movimientos de Stock"
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
      Left            =   5040
      TabIndex        =   22
      Top             =   315
      Width           =   4530
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Expedición y Despacho"
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
         Left            =   210
         TabIndex        =   25
         Top             =   1785
         Width           =   4110
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   3630
            TabIndex        =   34
            Top             =   420
            Width           =   3690
            Begin VB.CommandButton Command10 
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
               Left            =   3070
               Picture         =   "constok.frx":14BE
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Lista de Empaque"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command23 
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
               Picture         =   "constok.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Remito de Ventas - Clientes"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2510
               Picture         =   "constok.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Remito de Traslado"
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
               Left            =   560
               Picture         =   "constok.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   55
               ToolTipText     =   "Remito en Consignacion - Proveedores de Servicios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Comprobantes de  Despacho"
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
               Left            =   1200
               TabIndex        =   57
               Top             =   105
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fábrica - Ajustes - Transferencias"
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
         Left            =   210
         TabIndex        =   24
         Top             =   3150
         Width           =   4110
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   3630
            TabIndex        =   35
            Top             =   420
            Width           =   3690
            Begin VB.CommandButton Command17 
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
               Left            =   3070
               Picture         =   "constok.frx":20E6
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Descartes de Stock"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   560
               Picture         =   "constok.frx":23F0
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Transferencias de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command25 
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
               Picture         =   "constok.frx":26FA
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command21 
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
               Left            =   2520
               Picture         =   "constok.frx":2844
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Vale de Consumo de Materiales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos     Varios"
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
               TabIndex        =   36
               Top             =   105
               Width           =   1110
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Entradas a Depósito"
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
         Left            =   210
         TabIndex        =   23
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   28
            Top             =   420
            Width           =   1695
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
               Left            =   0
               Picture         =   "constok.frx":298E
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Producción"
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
               TabIndex        =   29
               Top             =   105
               Width           =   1110
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   26
            Top             =   420
            Width           =   1800
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
               Left            =   0
               Picture         =   "constok.frx":2C98
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción Proveedores"
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
               TabIndex        =   27
               Top             =   105
               Width           =   1110
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Control de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6210
      Left            =   210
      TabIndex        =   18
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Inventarios Físicos"
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
         Left            =   210
         TabIndex        =   21
         Top             =   4620
         Width           =   4215
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   32
            Top             =   480
            Width           =   1800
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
               Height          =   600
               Left            =   0
               Picture         =   "constok.frx":2FA2
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Gestión de Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   33
               Top             =   105
               Width           =   1320
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   30
            Top             =   480
            Width           =   1800
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
               Height          =   600
               Left            =   0
               Picture         =   "constok.frx":32AC
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Depósitos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   31
               Top             =   105
               Width           =   1320
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
         Height          =   2535
         Left            =   240
         TabIndex        =   20
         Top             =   1890
         Width           =   4215
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   420
            ScaleHeight     =   75
            ScaleWidth      =   3315
            TabIndex        =   49
            Top             =   2150
            Width           =   3375
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
         Begin VB.Frame Frame10 
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
            TabIndex        =   39
            Top             =   1320
            Width           =   3375
            Begin VB.OptionButton Option4 
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
               Left            =   1890
               TabIndex        =   4
               Top             =   250
               Width           =   1275
            End
            Begin VB.OptionButton Option5 
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
               Left            =   315
               TabIndex        =   3
               Top             =   250
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   37
            Top             =   420
            Width           =   3375
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
               Left            =   2770
               Picture         =   "constok.frx":35B6
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Disponibilidad General de Stocks"
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
               Height          =   600
               Left            =   1680
               Picture         =   "constok.frx":38C0
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Listado de Existencias Negativas"
               Top             =   0
               Width           =   560
            End
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
               Height          =   600
               Left            =   2215
               Picture         =   "constok.frx":3BCA
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Listado de Items con Stock Sub-Mínimo"
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
               Height          =   600
               Left            =   0
               Picture         =   "constok.frx":3ED4
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Existencias Físicas Actuales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias Negativos Críticos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   38
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Consultas por Pantalla"
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
         Left            =   210
         TabIndex        =   19
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   40
            Top             =   420
            Width           =   3375
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
               Picture         =   "constok.frx":41DE
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Consulta de Existencias Físicas Actuales y Ubicaciones"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   2760
               Picture         =   "constok.frx":4328
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Materiales Reservados"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2200
               Picture         =   "constok.frx":4632
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ficha Kardex - Histórico de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Existencias - His- torico - Reservas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   630
               TabIndex        =   41
               Top             =   105
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   9540
      Picture         =   "constok.frx":493C
      Stretch         =   -1  'True
      Top             =   5960
      Width           =   5250
   End
End
Attribute VB_Name = "CONSTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CONECRUN("LISEMPAQ", "Generación de Listas de Empaque")
    Command10.Enabled = True
End Sub

Private Sub Command13_Click()
    Call OPNOIN("")
End Sub

Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     If Shift = 1 Then
       Call RECUMOVI
     End If
   End If
End Sub

Private Sub Command12_Click()
   OBX$ = OBJPLA$("DESHASFECHA", "")
   If OBX$ = "" Then Exit Sub
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   ACTI% = 0
   Call COPYSTRU("MOVISTO", "MOVIFIL")
   JJ& = 0
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MOVISTO", U&)
     FF% = CVI(Left$(X$, 2))
     If FF% >= DES% Then ACTI% = 1
     If ACTI% = 1 Then
         JJ& = JJ& + 1
         Call GRAREG("MOVIFIL", X$, JJ&)
     End If
   Next U&
   Call SETULTREG("MOVIFIL", JJ&)
   Call CIERRARCH("MOVIFIL")
   Call FINAL("*MOVIFIL")
End Sub

Private Sub Command14_Click()
    OPRESERVA.Show 1
End Sub

Private Sub Command15_Click()
    Call OPNOIN("")
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
       Call DESCRAP
    End If
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    Call RECPROP
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call EXISTO2
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If Left$(TRIM$(UCase$(EMPREAC$)), 4) = "SABO" Then
         CONEXI.Top = CONSTOK.Top + 360
         CONEXI.Left = CONSTOK.Left + 240
         CONEXI.Show 1
       Else
         Call CARPEDPE
         CODISTOK.Show
    End If
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call CRITSTO2
    If ULTREG&("CRITSTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*CRITSTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*CRITSTL", "")
        End If
    End If
    Command20.Enabled = True
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    If DERACCE%("VALEMAT", "Vale de Consumo de Materiales") >= 2 Then
       Call VALEMAT
    End If
    Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    Call ENTRECONSI(0)
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMSTAN"
    LONGI = FileLen(CurDir + "\REMSTAN.EXE")
    If LONGI > 0 Then
       If LONGI < 400000 Then
          MODU$ = "#REMSTAN"
       End If
       Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    End If
    Command23.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("TRASTOCK", "Transferencias de Stock") >= 2 Then
      FORMTRANSFE.Show 1
   End If
   Command24.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTOCK.Show 1
    End If
    Command25.Enabled = True
End Sub

Private Sub Command26_Click()
    OPTRAJUST.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("CONSTOCK")
End Sub

Private Sub Command28_Click()
    Command28.Visible = False
    Call CONECRUN("RECEPMA", "Recepcion de Materiales")
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    OPXX% = ALTERNA%("Consulta por Día\Totales por Rango de Fechas\Control Desvíos de Consumo")
    On OPXX% GoTo 10, 20, 30
    GoTo 99
    '
10  HOI% = HOY(HO$)
    JJ& = 0
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHINVE$ <> "" Then
        FF% = FECHANUM(FECHINVE$)
        If FF% > 0 Then
            Screen.MousePointer = 11
            FIN& = ULTREG&("MOVISTO")
            For U& = 1 To FIN&
                Call TRACE(20, U&, FIN&)
                X$ = REGLEIDO$("MOVISTO", U&)
                FEMO% = CVI(Left$(X$, 2))
                If FEMO% = FF% Then
                    If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
                        CANTID# = CVD(Mid$(X$, 88, 8))
                        If CANTID# >= 0.05 Then
                            CI% = CVI(Mid$(X$, 3, 2))
                            REFE$ = Mid$(X$, 43, 30)
                            NOFA$ = Mid$(X$, 23, 10)
                            LOTE$ = Mid$(X$, 5, 16)
                            '
                            Y$ = REGBLAN$("PARPROD")
                            Call REPLA(Y$, MKI$(CI%), 1, 2)
                            Call REPLA(Y$, MKD$(CANTID#), 3, 8)
                            Call REPLA(Y$, REFE$, 11, 32)
                            Call REPLA(Y$, NOFA$, 43, 10)
                            Call REPLA(Y$, LOTE$, 53, 16)
                            JJ& = JJ& + 1
                            Call GRAREG("PARPROD", Y$, JJ&)
                        End If
                    End If
                End If
            Next U&
            Call CIERRARCH("PARPROD")
            Screen.MousePointer = 1
            '
            If JJ& < 1 Then
                Call COMUNI("No hay Partes de Producción Registrados\en la Fecha Elegida " + FECHINVE$ + ".")
                GoTo 99
            End If
            '
            Call HEADERS("PARPROD", "")
            Call HEADERS("PARPROD", "Fecha: " + FECHINVE$)
            Call FINAL("*LPARPRO")
            '
        End If
    End If
    GoTo 99
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    Dim CAFA#(32767)
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FEMO% = CVI(Left$(X$, 2))
       If FEMO% >= FF1% Then
          If FEMO% <= FF2% Then
             If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
                CANTID# = CVD(Mid$(X$, 88, 8))
                If CANTID# >= 0.05 Then
                   CI% = CVI(Mid$(X$, 3, 2))
                   If CI% > 0 Then
                      If CI% <= 32767 Then
                         If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                            CODE$ = CODEXT$(CI%)
                            If CODE$ >= cod1$ Then
                               If CODE$ <= cod2$ Then
                                  CAFA#(CI%) = CAFA#(CI%) + CANTID#
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
    Next U&
    '
    Call BLOQARCH("TOTPROD")
    JJ& = 0
    FIN& = NUINV%
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("INVERT", U& + 1)
       CI% = CVI(Mid$(X$, 17, 2))
       If CAFA#(CI%) >= 0.05 Then
          Z$ = REGBLAN$("TOTPROD")
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          Call REPLA(Z$, MKI$(TIMATE%(CI%)), 3, 2)
          Call REPLA(Z$, MKD$(CAFA#(CI%)), 5, 8)
          JJ& = JJ& + 1
          Call GRAREG("TOTPROD", Z$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("TOTPROD", JJ&)
    Call CIERRARCH("TOTPROD")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
       Call COMUNI("No hay Partes de Producción Registrados\en el Rango de Fechas Elegido.")
       GoTo 99
    End If
    '
    Call HEADERS("TOTPROD", "")
    Call HEADERS("TOTPROD", "Rango: " + RANGO$)
    Call FINAL("*LTOTPRO")
    '
    GoTo 99

30  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
32  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 32
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    REBLA0$ = REGBLAN$("COMPACON")
    JJ& = 0
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("MOVISTO", U&)
      FEMO% = CVI(Left$(X$, 2))
      If FEMO% >= FF1% Then
        If FEMO% <= FF2% Then
          If UCase$(Mid$(X$, 21, 2)) = "OF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
            CANTID# = CVD(Mid$(X$, 88, 8))
            If CANTID# >= 0.05 Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                If CI% <= 32767 Then
                  If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                    CODE$ = CODEXT$(CI%)
                    If CODE$ >= cod1$ Then
                      If CODE$ <= cod2$ Then
                        REBLA$ = REBLA0$
                        Call REPLA(REBLA$, MKI$(FEMO%), 1, 2)
                        Call REPLA(REBLA$, Mid$(X$, 23, 10), 3, 10)
                        Call REPLA(REBLA$, MKI$(CI%), 13, 2)
                        REBLA1$ = REBLA$
                        CABAS = CANTID#
                        Call LEEEXPLO(CODE$, CABAS)
                        For KKI% = 1 To CAIT%
                          If Abs(CABAS * USOI(KKI%)) > 0.005 Then
                            YY$ = REBLA$
                            Call REPLA(YY$, MKI$(CIJ%(KKI%)), 15, 2)
                            Call REPLA(YY$, MKS$(CABAS * USOI(KKI%)), 17, 4)
                            Call REPLA(YY$, MKS$(0), 21, 4)
                            Call REPLA(YY$, MKS$(0), 25, 4)
                            JJ& = JJ& + 1
                            Call GRAREG("COMPACON", YY$, JJ&)
                          End If
                        Next KKI%
                      End If
                    End If
                  End If
                End If
              End If
              GoTo 89
            End If
          End If
          '
          If UCase$(Mid$(X$, 21, 2)) = "CF" Or UCase$(Mid$(X$, 23, 2)) = "OF" Then
            CANTID# = CVD(Mid$(X$, 96, 8))
            If CANTID# >= 0.05 Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                If CI% <= 32767 Then
                  CODE$ = CODEXT$(CI%)
                  YY$ = REBLA1$
                  Call REPLA(YY$, Mid$(X$, 23, 10), 3, 10)
                  Call REPLA(YY$, MKI$(CI%), 15, 2)
                  Call REPLA(YY$, MKS$(0), 17, 4)
                  CARESAL = CANTID#
                  Call REPLA(YY$, MKS$(CARESAL), 21, 4)
                  Call REPLA(YY$, MKS$(0), 25, 4)
                  JJ& = JJ& + 1
                  Call GRAREG("COMPACON", YY$, JJ&)
                End If
              End If
            End If
          End If
        End If
      End If
89  Next U&
    '
    Call SETULTREG("COMPACON", JJ&)
    Call CIERRARCH("COMPACON")
    '
    Call HEADERS("COMPACON", "")
    Call HEADERS("COMPACON", "Rango: " + RANGO$)
    Call FINAL("*SOCOCON")
    GoTo 99
    '
99  Command3.Enabled = True
    '
End Sub
'


Private Sub Command4_Click()
    Command4.Enabled = False
    If EXISTE%("ADMIPED.EXE") > 0 Then
         Call CONECRUN("ADMIPED", "Pedidos de Clientes - Version Completa")
       ElseIf EXISTE%("PEDBASI.EXE") > 0 Then
         Call CONECRUN("PEDBASI", "Pedidos de Clientes - Version Basica")
       Else
         Call OPNOIN("Pedidos de Clientes")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    DISPOSTO.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command30_Click()
    Command30.Enabled = False
    Call NEGASTO2
    If ULTREG&("NEGASTO") > 0 Then
        Call CIERRARCH("*.*")
        If Option5.Value = True Then
            Call CONECRUN("*NEGASTS", "")
          ElseIf Option4.Value = True Then
            Call CONECRUN("*NEGASTL", "")
        End If
    End If
    Command30.Enabled = True
End Sub

Private Sub Command6_Click()
    Call CONECRUN("INVENTAR", "Inventarios Físicos y Valorizados")
End Sub

Private Sub Command7_Click()
    OPRESERVA.Show 1
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONECRUN("REMITRAS", "Emision de Remito de Traslado")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call GENKARDEX("")
    Command9.Enabled = True
End Sub

Private Sub Form_Activate()
    TEXPE$ = "Remito de Ventas - Clientes"
    If CONTROL$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Command23.ToolTipText = TEXPE$
    Call GRATITFOR(Caption)
    '
    WindowState = 0
    OPENFORMS = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("REMSTAN.EXE") > 0 Then Command23.Enabled = True
    If EXISTE%("OCOMPRA.EXE") > 0 Then Command22.Enabled = True
    If EXISTE%("REMITRAS.EXE") > 0 Then Command8.Enabled = True
    If EXISTE%("LISEMPAQ.EXE") > 0 Then Command10.Enabled = True
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Label5_Click()
    LU$ = LUDAT$
    If EXISTE%(LU$ + "GETRAMO.DAT") > 0 Then
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("GETRAMO")
            X$ = REGLEIDO$("GETRAMO", U&)
            COD$ = Left$(X$, 16)
            CI% = CODINT%(COD$)
            Y$ = Mid$(X$, 17, 128)
            Call REPLA(Y$, MKI$(CI%), 3, 2)
            Call REGAPP("MOVISTO", Y$)
        Next U&
        Call CIERRARCH("*.*")
        Call FILEKILL(LU$ + "GETRAMO.DAT")
        Screen.MousePointer = 1
    End If
End Sub
'
Sub IMPORNEU2()
   '
   If ULTREG&("MOVISTO") < 1 Then GoTo 3
   LU$ = LUDAT$
   If EXISTE%(LU$ + "MOVISTX.DAT") = 1 Then
      Screen.MousePointer = 11
      Call COPYSTRU("MOVISTO", "MOVISTX")
      FIN& = ULTREG&("MOVISTX")
      For U& = 1 To FIN&
         XXX$ = REGLEIDO$("MOVISTX", U&)
         CIXI% = CVI(Mid$(XXX$, 3, 2))
         Call REGAPP("MOVISTO", XXX$)
         Call CIERRARCH("MOVISTO")
         Call VERISAL(CIXI%)
         Call VESARRAS(CIXI%)
      Next U&
   End If
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
3  Screen.MousePointer = 11
   Call BLANARCH("MOVISTO")
   '
   Dim SALSTO(16384)
   Dim CODEPOS$(64)
   Dim ALIDEP$(128), CODEPX$(128)
   '
   ALIDEP$(1) = "FK": CODEPX$(1) = "5110000"
   ALIDEP$(2) = "DEP.": CODEPX$(2) = "4080000"
   ALIDEP$(3) = "BS": CODEPX$(3) = "5120000"
   ALIDEP$(4) = "510": CODEPX$(4) = "5100000"
   ALIDEP$(5) = "513": CODEPX$(5) = "5130000"
   ALIDEP$(6) = "408000": CODEPX$(6) = "4080000"
   CALI% = 6
   '
   CADEPO% = 0
   For U& = 1 To 64
      If U& <= ULTREG&("SUCURSAL") Then
         CODEPOS$(U&) = TRIM$(Mid$(REGLEIDO$("SUCURSAL", U&), 21, 8))
         CADEPO% = U&
      End If
   Next U&
   '
   'For U& = 1 To ULTREG&("MOVISTO")
   '   X$ = REGLEIDO$("MOVISTO", U&)
   '   CI% = CVI(Mid$(X$, 3, 2))
   '   CAN = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
   '   If CI% > 0 Then
   '      If CI% <= 16384 Then
   '         SALSTO(CI%) = SALSTO(CI%) + CAN
   '      End If
   '   End If
19 'Next U&
   '
   GoSub 95
   '
   HOII% = HOY(HO$)
   NRO& = 0
   CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
   If CMOV$ = "" Then CMOV$ = "AJ"
   '
   ARIMPO$ = "A:\MULTIST"
   PRIBOT% = 0
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
      Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              For KK% = 3 To 6
                 Line Input #NX%, A$
              Next KK%
              CODEPO1$ = TRIM$(Mid$(A$, 2))
              CODEPO$ = TRIM$(REPLACAR$(CODEPO1$, Chr$(34), " "))
              For KK% = 1 To CALI%
                 If CODEPO$ = ALIDEP$(KK%) Then
                    CODEPO$ = CODEPX$(KK%)
                    GoTo 5
                 End If
              Next KK%
              '
5             For KK% = 1 To CADEPO%
                 If CODEPOS$(KK%) = CODEPO$ Then
                    IDEPO% = KK%
                    GoTo 10
                 End If
              Next KK%
              CODEPO$ = CODEPOS$(1)
              IDEPO% = 1
              '
10            For KK% = 7 To 20
                 Line Input #NX%, A$
              Next KK%
              CANTI = Val(Mid$(A$, 2))
              GoSub 90
         End If
      End If
   Wend
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
90 CI% = CODINT%(COD$)
   If CI% > 0 Then
      If Abs(CANTI) > 0.05 Then
         If CI% <= 16384 Then
            SALSTO(CI%) = SALSTO(CI%) + CANTI
         End If
         NRO& = NRO& + 1
         DOSEC$ = "AJ." + Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
         DOPRI$ = "AS." + TRIM$(Str$(NRO&))
         VUNI# = 0
         MONE$ = ""
         NC% = 0
         REFE$ = "ST.IMPORT.AL " + HO$
         LTX$ = "ST.INICIAL"
         DEPO% = IDEPO%
         FF% = HOII%
         '
         Call MOVISTO(FF%, CI%, LTX$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
         '
      End If
   End If
   Return
   '
95 For CI% = 1 To NUMAS%
      U& = CI% + 1
      REBLA$ = String$(36, 0) + MKS$(0.00001)
      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
      Call GRAREG("SALSTOCK", REBLA$, U&)
      REBLA$ = String$(252, 0) + MKS$(0.00001)
      Call REPLA(REBLA$, MKS$(SALSTO(CI%)), 1, 4)
      Call GRAREG("SALDEPOS", REBLA$, U&)
   Next CI%
   Return
   '
End Sub

'
Sub IMPORNEU3()
   '
   Dim SALSTO(16384)
   Dim CODEPOS$(64)
   Dim ALIDEP$(128), CODEPX$(128)
   '
   If ULTREG&("UBISTOCK") > 0 Then
      Call MENSERR(24, "Imposible Importar Ubicaciones.\Ya Existe Archivo de Ubicaciones.")
      Call CIERRARCH("*.*")
      GoTo 50
   End If
   '
   ARIMPO$ = "A:\UBIC"
   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   PRIBOT% = 0
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
      Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              For KK% = 3 To 4
                 Line Input #NX%, A$
              Next KK%
              UBICA1$ = TRIM$(Mid$(A$, 2))
              UBICA$ = TRIM$(REPLACAR$(UBICA1$, Chr$(34), " "))
              '
10            For KK% = 5 To 6
                 Line Input #NX%, A$
              Next KK%
              LOECON = Val(Mid$(A$, 2))
              '
20            For KK% = 7 To 8
                 Line Input #NX%, A$
              Next KK%
              PUNPED = Val(Mid$(A$, 2))
              '
              GoSub 90
         End If
      End If
   Wend
   '
50 If ULTREG&("MASTSPEC") > 0 Then
      Call MENSERR(24, "Imposible Importar Especificaciones.\Ya Existe Archivo de Especificaciones.")
      Call CIERRARCH("*.*")
      GoTo 80
   End If
   '
   ARIMPO$ = "A:\ESPECIF"
   ARIMPOX$ = REPLACAR$(ARIMPO$, "\", "/")
   If COMALTER%("Importar Desde " + ARIMPOX$ + "\\Continuar\Cancelar") <> 1 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   PRIBOT% = 0
   Screen.MousePointer = 11
   NX% = FILEOPEN%(ARIMPO$, 1, 0)
   While Not EOF(NX%)
55    Line Input #NX%, A$
      If Left$(A$, 3) = "BOT" Then
         If PRIBOT% = 0 Then
              PRIBOT% = 1
           Else
              For KK% = 1 To 2
                 Line Input #NX%, A$
              Next KK%
              cod1$ = TRIM$(Mid$(A$, 2))
              COD$ = TRIM$(REPLACAR$(cod1$, Chr$(34), " "))
              Line Input #NX%, A$
60            Line Input #NX%, TTXX$
              If Left$(TTXX$, 3) = "-1," Then GoTo 55
              GoSub 95
              GoTo 60
         End If
      End If
   Wend
   '
80 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Exit Sub
   '
90 CI% = CODINT%(COD$)
   If CI% > 0 Then
      If TRIM$(UBICA$) <> "" Then
         '
         XX$ = REGBLAN$("UBISTOCK")
         Call REPLA(XX$, MKI$(CI%), 1, 2)
         Call REPLA(XX$, UBICA$, 3, 24)
         Call REGAPP("UBISTOCK", XX$)
         '
      End If
      '
      REGMA& = CI% + 1
      YY$ = REGLEIDO$("MASTER", REGMA&)
      If PUNPED > 0.05 Then
         STMI& = PUNPED
         If STMI& > 30000 Then STMI& = (-1) * PUNPED / 100
         stm1% = STMI&
         Call REPLA(YY$, MKI$(stm1%), 105, 2)
      End If
      If LOECON > 0.05 Then
         LOECO& = LOECON
         If LOECO& > 30000 Then LOECO& = (-1) * LOECON / 100
         lec1% = LOECO&
         Call REPLA(YY$, MKI$(lec1%), 107, 2)
      End If
      Call GRAREG("MASTER", YY$, REGMA&)
      '
   End If
   Return
   '
95 TTYY$ = TRIM$(REPLACAR$(TTXX$, Chr$(34), " "))
   If TRIM$(TTYY$) <> "" Then
      CI% = CODINT%(COD$)
      If CI% > 0 Then
         Call FRATEXTO(TTYY$, 60)
         For KI% = 1 To CARETEX%
            XX$ = REGBLAN$("MASTSPEC")
            Call REPLA(XX$, MKI$(CI%), 1, 2)
            Call REPLA(XX$, RETEX$(KI%), 3, 78)
            Call REGAPP("MASTSPEC", XX$)
         Next KI%
      End If
   End If
   Return
   '
End Sub


