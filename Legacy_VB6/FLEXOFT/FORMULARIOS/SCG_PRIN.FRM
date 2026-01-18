VERSION 5.00
Begin VB.Form SCG_PRIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Contabilidad General"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10860
   ClipControls    =   0   'False
   Icon            =   "SCG_PRIN.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton COMMAND30 
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
      Left            =   9975
      Picture         =   "SCG_PRIN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "CREA ARCHIVO TXT DE LIBRO  DIARIO GENERAL"
      Top             =   0
      Width           =   560
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Trans."
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
      Height          =   735
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Transferencia de Datos"
      Top             =   2880
      Width           =   840
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   6300
      Sorted          =   -1  'True
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Control Pos.Diferida"
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
      Left            =   3990
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9870
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   52
      Top             =   3840
      Width           =   855
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
   Begin VB.CommandButton Command22 
      Caption         =   "Plan"
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
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Plan General de Cuentas"
      Top             =   4860
      Width           =   840
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Control"
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
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":08A8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Control de Consistencia de la Base de Datos"
      Top             =   5660
      Width           =   840
   End
   Begin VB.CommandButton Command27 
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
      Height          =   735
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
      Top             =   2080
      Width           =   840
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
      Height          =   735
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":0EBC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   4060
      Width           =   840
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
      Height          =   735
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":12FE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1280
      Width           =   840
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
      Left            =   9870
      Picture         =   "SCG_PRIN.frx":1608
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas e Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   5040
      TabIndex        =   11
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Otros Reportes de Control"
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
         TabIndex        =   45
         Top             =   4620
         Width           =   4215
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   46
            Top             =   420
            Width           =   3375
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
               Left            =   2760
               Picture         =   "SCG_PRIN.frx":1752
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Ajuste por Inflación"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   0
               Picture         =   "SCG_PRIN.frx":1A5C
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Informes Standard Normalizados"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Informes Contables Complementarios"
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
               Left            =   840
               TabIndex        =   49
               Top             =   105
               Width           =   1845
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Opciones de Balance"
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
         TabIndex        =   41
         Top             =   3150
         Width           =   4215
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2685
            TabIndex        =   42
            Top             =   420
            Width           =   2745
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
               Picture         =   "SCG_PRIN.frx":1D66
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Balances de Sumas y Saldos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Mensual y Ge- neral - Sumas y Saldos"
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
               TabIndex        =   44
               Top             =   120
               Width           =   2160
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Informes de Caja y Bancos"
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
         TabIndex        =   35
         Top             =   1785
         Width           =   4215
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   36
            Top             =   420
            Width           =   3375
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
               Left            =   2200
               Picture         =   "SCG_PRIN.frx":2070
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Subdiario de Caja (Ingresos y Egresos)"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2760
               Picture         =   "SCG_PRIN.frx":237A
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Resumen de Cuenta de Caja y Bancos"
               Top             =   0
               Width           =   560
            End
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
               Height          =   600
               Left            =   0
               Picture         =   "SCG_PRIN.frx":2684
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Emision de Subdiario de Caja y Bancos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Reportes de Caja y Bancos"
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
               Left            =   840
               TabIndex        =   40
               Top             =   120
               Width           =   1425
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Informes de Diario"
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
         TabIndex        =   30
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   31
            Top             =   420
            Width           =   3375
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
               Height          =   600
               Left            =   0
               Picture         =   "SCG_PRIN.frx":298E
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Libro Diario General"
               Top             =   0
               Width           =   560
            End
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
               Left            =   2760
               Picture         =   "SCG_PRIN.frx":2C98
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Mayor por Cuenta"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Reportes de la Contabilidad General"
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
               Left            =   735
               TabIndex        =   34
               Top             =   120
               Width           =   1845
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Asientos Contables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   210
      TabIndex        =   9
      Top             =   315
      Width           =   4635
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cheques"
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
         Left            =   2415
         TabIndex        =   54
         Top             =   3150
         Width           =   2010
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1530
            TabIndex        =   55
            Top             =   420
            Width           =   1590
            Begin VB.CommandButton Command23 
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
               Picture         =   "SCG_PRIN.frx":2FA2
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Gestión de Cartera de Cheques y Valores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Gestión Valores"
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
               Left            =   735
               TabIndex        =   57
               Top             =   105
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Apertura y Cierre"
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
         Top             =   4620
         Width           =   4215
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   26
            Top             =   420
            Width           =   3375
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
               Left            =   560
               Picture         =   "SCG_PRIN.frx":32AC
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Asiento de Apertura de Ejercicio"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command7 
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
               Left            =   2760
               Picture         =   "SCG_PRIN.frx":35B6
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Asiento de Cierre de Ejercicio"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "SCG_PRIN.frx":38C0
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Apertura de Nuevo Ejercicio Contable"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Apertura y Cierre de Ejercicio"
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
               Left            =   1260
               TabIndex        =   29
               Top             =   120
               Width           =   1530
            End
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mayorizar"
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
         TabIndex        =   21
         Top             =   3150
         Width           =   2010
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1530
            TabIndex        =   22
            Top             =   420
            Width           =   1590
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
               Height          =   600
               Left            =   0
               Picture         =   "SCG_PRIN.frx":3BCA
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Generar Asiento Resumen de Caja y Bancos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen de C.y B."
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
               TabIndex        =   24
               Top             =   105
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Asientos de Caja y Bancos"
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
         TabIndex        =   14
         Top             =   1785
         Width           =   4215
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   15
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
               Left            =   2760
               Picture         =   "SCG_PRIN.frx":3D14
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Asiento Resumen de Fondo Fijo"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   0
               Picture         =   "SCG_PRIN.frx":3E5E
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Asiento Manual de Caja y Bancos"
               Top             =   0
               Width           =   560
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
               Height          =   600
               Left            =   2200
               Picture         =   "SCG_PRIN.frx":4168
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Asiento Resumen de Cobranza"
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
               Height          =   600
               Left            =   1640
               Picture         =   "SCG_PRIN.frx":42B2
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Asiento Resumen de Pagos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Caja y Bancos"
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
               Left            =   735
               TabIndex        =   19
               Top             =   120
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Asientos de Diario"
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
         TabIndex        =   10
         Top             =   420
         Width           =   4215
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3315
            TabIndex        =   12
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
               Picture         =   "SCG_PRIN.frx":45BC
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Asiento Manual de Diario"
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
               Picture         =   "SCG_PRIN.frx":48C6
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Asiento Resumen de Compras"
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
               Picture         =   "SCG_PRIN.frx":4BD0
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Asiento Resumen de Ventas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Registrar / Con- sultar As. Diario"
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
               Left            =   720
               TabIndex        =   13
               Top             =   120
               Width           =   1740
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10900
      Y1              =   6570
      Y2              =   6570
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
      Left            =   3465
      TabIndex        =   8
      Top             =   6720
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "SCG_PRIN.frx":5012
      Stretch         =   -1  'True
      Top             =   6580
      Width           =   10920
   End
End
Attribute VB_Name = "SCG_PRIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 1
    ASICAJA.Show 1
    Unload ASICAJA
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    If CONTROL$("CONTAGEN", "SIDISTRI") = "DISTRIBU" Then
       Call COMUNI("Opcion no Permitida - Sistema Distribuido.")
       GoTo 99
    End If
    'Call BLOQARCH("ASICONT")
    MODOX% = ALTERNA%("Generar Asiento Resumen de Caja-Bancos\Borrar Asiento Resumen de C-B")
    '
    If MODOX% = 1 Then
        TIASI% = 3
        ASIDIAR.Show 1
        Unload ASIDIAR
      ElseIf MODOX% = 2 Then
        Call BORRARESCAJ
    End If
    '
99  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    RESCUEDIAR.Show 1
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    Call DIARIOGEN   ' Impresion Libro Diario General
    Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 2
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    Call SUDIARCAJ   ' Impresion Subdiario de Caja y Bancos
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    EPAC$ = UCase$(EMPREAC$)
    OPPI% = 0
    If InStr(EPAC$, "CIME-SOI") > 0 Then OPPI% = 1
    If InStr(EPAC$, "SABO") > 0 Then OPPI% = 1
    If InStr(EPAC$, "NOVOESPU") > 0 Then OPPI% = 1
    If InStr(EPAC$, "VDR") > 0 Then OPPI% = 1
    If OPPI% < 1 Then
         Call OPNOIN("Informes Normalizados")
       Else
         Call CONECRUN("#SCG13B", "Controles de Consistencia")
    End If
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    RESCUECAJA.Show 1
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call SUCAJINEG    ' Subdiario de Caja (Ingresos - Egresos
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    OPBALANCE.Show 1
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 0
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    Call CONECRUN("AJUINFLA", "Ajuste por Inflación")
    Command20.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
    Command21.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 4
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    If EXISTE%("CPLANCUE.EXE") > 0 Then
        Call CONECRUN("CPLANCUE", "Actualización Plan General de Cuentas")
      Else
        PLAGRAL.Show 1
    End If
    Command22.Enabled = True
End Sub


Private Sub Command23_Click()
    Command23.Enabled = False
    GECHECAR.Show 1
    Command23.Enabled = True
End Sub


Private Sub Command24_Click()
   '
   ' control posicion diferida
   '
   Dim FEREGIS%(8192), IMPORE#(8192), REFERE$(8192)
   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
   '
1 Call SELECHO("ECUECON")
   CUEX$ = VALACT1$("ECUECON")
   If CUEX$ = "" Then Exit Sub
   '
   CUE% = CUEINT%(CUEX$)
   If CUE% < 1 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 1
   End If
   '
   Call FRESHALL
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ& = 0: TRANSPORTE# = 0
   SUMATO# = 0
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   Call CIERRACONTA
   SE$ = "0": EXTE$ = "OLD"
   JJ& = 0
   '
8 DES% = 0: HAS% = 32767
   '
   PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
 If EXTE$ = "DAT" Then
   If CVS(E12$) = 1 Then
     GoTo 9    ' SALTEA APERTURA EJERCICIO 2002
   End If
 End If
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         If FECHASIEN% < DES% Then
             TRANSPORTE# = TRANSPORTE# + IMPOPAS#
             GoTo 9
           ElseIf FECHASIEN% <= HAS% Then
             GoSub 190
         End If
      End If
9 Next KK%
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If FECHASIEN% < DES% Then
          TRANSPORTE# = TRANSPORTE# + IMPOPAS#
          GoTo 11
        ElseIf FECHASIEN% <= HAS% Then
          GoSub 190
      End If
11 Next KK%
   '
   Call CIERRACONTA
   If Abs(TRANSPORTE#) >= 0.005 Then
      Call MENSERR(24, "Hay error en Transporte")
      Stop
   End If
   '
   If EXTE$ = "OLD" Then
     EXTE$ = "DAT"
     GoTo 8
   End If
   '
   AAA = SUMATO#
   Stop
   GoTo 100
   '
   '
190 JJ& = JJ& + 1
    FEREGIS%(JJ&) = FECHASIEN%
    IMPORE#(JJ&) = IMPOPAS#
    REFERE$(JJ&) = MARCADIA$ + DETAPAS$
    SUMATO# = SUMATO# + IMPOPAS#
    Return
    '
   '
   '
   '
   '
   'REVISA POSICION DIFERIDA
   '
100 Dim REFE$(2048)
    Dim FECHAMO%(2048)
    Dim IPODIF#(2048)
    '
    JJ1& = 0
    EXTE$ = "DAT"
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUEX$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUEX$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   HOII% = HOY(HOS$)
   DES% = FECHANUM("01/10/01")
   HAS% = FECHANUM("31/12/02")
   '
10 Screen.MousePointer = 11
   KLM& = 0
   '
   CUEBIN$ = MKI$(CUE%)
   CODENOCUE$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
   JJ1& = 0
   SALINI# = 0
   '
   If DES% >= FECHANUM("01/10/01") Then
      CCCC$ = TRIM$(CODCUE$(CUE%))
      If CCCC$ = "111401" Then SALINI# = 11409.33
      If CCCC$ = "111502" Then SALINI# = 11791.1
      If CCCC$ = "111403" Then SALINI# = 2414.3
      If CCCC$ = "111406" Then SALINI# = 29769.7
      If CCCC$ = "111506" Then SALINI# = 381.86
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX9%: NX9% = 0
     Close #NX10%: NX10% = 0
     Close #NX11%: NX11% = 0
     EXTE$ = "OLD"
   End If
   '
12 PMC$ = PUNMODIAR$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 2
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIA% = CVI(Mid$(PMC$, KK%, 2))
      Call LEEDIARIO(REGDIA%)
      If CVS(E12$) < 1.05 Then GoTo 19 ' SALTEA ASIENTO APERTURA
      If Asc(E18$) <> 1 Then  ' SALTEA ASIENTOS RESUMEN DE CAJA
         Call LEEPASDIA(REGDIA%, 1)
         If TRIM$(DETAPAS$) = "" Then
            Call LEEPASDIA(REGDIA%, 2)
            DETAPAS$ = REFASIEN$
         End If
         MARCADIA$ = "*"
         FEMOREA% = CVI(E15$)
         IMPO# = CVD(E16$)
         If FEMOREA% < DES% Then
             If FEMOREA% > FECHANUM("31/09/01") Then GoSub 90 'TOTALIZAR INICIAL
          ElseIf FEMOREA% <= HAS% Then
            FEMORE% = FEMOREA%
            FEREGI% = FEMOREA%
            FEMIS% = FEMOREA%
            GoSub 95 'OK
         End If
      End If
19 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX9%: NX9% = 0
       Close #NX10%: NX10% = 0
       Close #NX11%: NX11% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 12
     End If
   End If
   '
   '
20 ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   EXTEA$ = EXTE$
   If EXTEA$ = "DAT" Then
     Close #NX5%: NX5% = 0
     Close #NX6%: NX6% = 0
     Close #NX7%: NX7% = 0
     EXTE$ = "OLD"
   End If
   '
22 PMC$ = PUNMOCAJA$(CUE%)
   FIN& = Len(PMC$)
   For KK% = 1 To Len(PMC$) Step 4
      UU& = KK% + 1
      Call TRACE(20, UU&, FIN&)
      REGDIAL& = CVS(Mid$(PMC$, KK%, 4))
      Call LEEPASCAJ(REGDIAL&, 1)
      FEMOVI% = CVI(D15$)
      IMPO# = CVD(D16$)
      '
      If TRIM$(DETAPAS$) = "" Then
         Call LEEPASCAJ(REGDIAL&, 2)
         DETAPAS$ = REFASIEN$
      End If
      MARCADIA$ = ""
      If InStr(DETAPAS$, "- OP") > 0 Then GoTo 29
      If Left$(DETAPAS$, 8) = "DEPOSITO" Then
         FEMORE% = FECHACRE%(FEMOVI%, Val(Mid$(DETAPAS$, 10, 3)))
         If FEMORE% < DES% Then
            If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          ElseIf FEMORE% <= HAS% Then
 'FEMORE% = FEMOVI%    ' para poder comparar
            FEMIS% = FEMOVI%
            GoSub 95    ' ok
         End If
         GoTo 29
      End If
      If FEMOVI% < DES% Then
        If FEMOVI% > FECHANUM("31/09/01") Then GoSub 90
        GoTo 29
      End If
      If FEMOVI% >= DES% Then
        If FEMOVI% <= HAS% Then
          FEMORE% = FEMOVI%
          FEMIS% = FEMOVI%
          GoSub 95  'ok
        End If
      End If
      '
29 Next KK%
   '
   If EXTE$ = "OLD" Then
     If EXTEA$ = "DAT" Then
       Close #NX5%: NX5% = 0
       Close #NX6%: NX6% = 0
       Close #NX7%: NX7% = 0
       EXTE$ = "DAT"
       EXTEA$ = ""
       GoTo 22
     End If
   End If
   '
   For U& = 1 To ULTREG&("CHEQUEMI")
     X$ = REGLEIDO$("CHEQUEMI", U&)
      PROV% = CVI(Left$(X$, 2))
      NUMERO$ = Mid$(X$, 3, 12)
      BANCO$ = Mid$(X$, 15, 16)
      EMIS% = CVI(Mid$(X$, 47, 2)) ' toma fecha de O/Pago
      IMPO# = (-1) * CVD(Mid$(X$, 33, 8))
      If Abs(IMPO#) < 0.005 Then GoTo 39
      '
      NOPAGO& = CVS(Mid$(X$, 41, 4))
      VENCI% = CVI(Mid$(X$, 49, 2))
      QUE% = CVI(Mid$(X$, 45, 2))
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 38
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
38    CUEH% = CUEMP%(UX%)
      If CUEH% = CUE% Then
        FEMORE% = VENCI%      ' TOMA FECHA DE EMISION - CAMBIE ACA
        FEMIS% = EMIS%
        If FEMORE% < DES% Then
          If FEMORE% > FECHANUM("31/09/01") Then GoSub 90
          'GoSub 90
        End If
        If FEMORE% >= DES% Then
          If FEMORE% <= HAS% Then
           FEX = EMIS%
            DETAPAS$ = "Ch." + TRIM$(NUMERO$)
            DETAPAS$ = DETAPAS$ + " - OP." + TRIM$(Str$(NOPAGO&))
            DETAPAS$ = DETAPAS$ + " - " + TRIM$(RASOCLI$((-1) * PROV%))
            VINODEOPA% = 1
            GoSub 95 ' OK
            VINODEOPA% = 0
          End If
        End If
      End If
39 Next U&
   '
   FEX = DES%
   FEMORE% = DIANTE(FEX)
   DETAPAS$ = "Saldo al " + FECHATEX$(FEX)
   IMPO# = SALINI#
   FEMIS% = FEMORE%
   GoSub 95
   '
AAA = SUPODIF#
   '
SUDIFER# = 0
111 For KK11& = 1 To JJ&
If Abs(IMPORE#(KK11&)) >= 0.005 Then
  SUDIFER# = SUDIFER# + IMPORE#(KK11&)
  MsgBox REFERE$(KK11&) & IMPORE#(KK11&)
  TTT1& = TTT1& + 1
End If
Next KK11&
Stop
   '
   '
   '
   '
   '
90 SALINI# = SALINI# + IMPO#
   Return
   '
95 If Abs(IMPO#) >= 0.005 Then
      If VINODEOPA% = 1 Then
        If FEMIS% < DES% Then
          GoSub 90     ' DESCONTAR DE SALDO INICIAL
          GoTo 951
        End If
      End If
      '
      SUPODIF# = SUPODIF# + IMPO#
      For KK11& = 1 To JJ&
        If FEREGIS%(KK11&) = FEMIS% Then
          If Abs(IMPORE#(KK11&) - IMPO#) < 0.005 Then
            IMPORE#(KK11&) = IMPORE#(KK11&) - IMPO#
            GoTo 951
          End If
        End If
      Next KK11&
      '
      JJ& = JJ& + 1
      FEREGIS%(JJ&) = FEMORE%
      IMPORE#(JJ&) = (-1) * IMPO#
      REFERE$(JJ&) = MARCADIA$ + DETAPAS$
951 End If
   Return
   '

End Sub


Private Sub Command25_Click()
   Command25.Enabled = False
   Dim NARCH$(128)
   OPX% = COMALTER%("Opciones de Trabajo:\\Exportar\Importar")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   '
   QAR% = 1: NARCH$(QAR%) = "VARIAB"
   If TRIM$(CONTROL$("PROPSCG", "PLAGRAL")) <> "GUEST" Then
     QAR% = QAR% + 1: NARCH$(QAR%) = "PLAGRAL"
     QAR% = QAR% + 1: NARCH$(QAR%) = "INVECUE"
     QAR% = QAR% + 1: NARCH$(QAR%) = "DENOCUE"
     QAR% = QAR% + 1: NARCH$(QAR%) = "CUENTAS"
     QAR% = QAR% + 1: NARCH$(QAR%) = "CIERRE0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "SUBEMPRE"
     QAR% = QAR% + 1: NARCH$(QAR%) = "PLANALT"
     QAR% = QAR% + 1: NARCH$(QAR%) = "INFORMES"
     QAR% = QAR% + 1: NARCH$(QAR%) = "INDAJU"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ENCAINF"
     QAR% = QAR% + 1: NARCH$(QAR%) = "DENENCA"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ECUECON"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ECUECOT"
     QAR% = QAR% + 1: NARCH$(QAR%) = "CUECON"
   End If
   '
   If CONTROL$("PROPSCG", "CAJABANC") <> "GUEST" Then
     QAR% = QAR% + 1: NARCH$(QAR%) = "SUCAJA0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ISUCAJA0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "SALCAJA0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ENCAJA0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ASIDIAR"
   End If
   '
   If CONTROL$("PROPSCG", "CONTAGEN") <> "GUEST" Then
     QAR% = QAR% + 1: NARCH$(QAR%) = "DIARIO0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "IDIARIO0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "CIERRE0"
     QAR% = QAR% + 1: NARCH$(QAR%) = "ENCDIAR0"
   End If
   '
   '
   If OPX% = 1 Then
     On Error Resume Next
     Kill "C:\FLEXOFT\ATRACON.TRF"
     On Error GoTo 0
     '
     JJ& = 0
     LUXX$ = LUDAT$
     Dim BUFGRA As String * 32
     Dim BUFLEC As String * 16
     NY% = FILEOPEN%("C:\FLEXOFT\ATRACON.TRF", 0, 32)
     '
     For IAR% = 1 To QAR%
       For IEJ% = 1 To 2
         EXTX$ = "DAT"
         If IEJ% = 2 Then EXTX$ = "OLD"
         '
         ARAC$ = NARCH$(IAR%) + "." + EXTX$
         NX% = FILEOPEN%(LUXX$ + ARAC$, 0, 16)
         '
         For U& = 1 To LOF(NX%) / 16
           Get #NX%, U&, BUFLEC$
           XX$ = BUFLEC$
           YY$ = Space$(16) + String$(16, 0)
           Mid$(YY$, 1, 16) = ARAC$
           Mid$(YY$, 17, 16) = XX$
           BUFGRA$ = YY$
           JJ& = JJ& + 1
           On Error Resume Next
           Put #NY%, JJ&, BUFGRA$
           If Err Then
             MsgBox "Imposible Grabar en 'ATRACON.TRF'"
             Resume
             Close
             GoTo 99
           End If
         Next U&
         Close #NX%: NX% = 0
       Next IEJ%
     Next IAR%
     '
     While JJ& < LOF(NY%) / 32
       BUFGRA$ = String$(32, 0)
       JJ& = JJ& + 1
       Put #NY%, JJ&, BUFGRA$
     Wend
     Close #NY%
     '
     Call COMUNI("El Archivo de Transferencia 'ATRACON.TRF'\fue Creado en su Disco 'C:' - Carpeta 'FLEXOFT'.")
     GoTo 99
   End If
   '
   '
   If OPX% = 2 Then
30   If COMALTER%("Coloque Archivo 'ATRACON.TRF'\en Carpeta 'FLEXOFT' del Disco 'C:'\\Cancelar\Continuar") <> 2 Then GoTo 99
     If EXISTE%("C:\FLEXOFT\ATRACON.TRF") < 1 Then GoTo 30
     FIDAT$ = Format$(FileDateTime("C:\FLEXOFT\ATRACON.TRF"), "dd/mm/yyyy - hh:mm")
     If COMALTER%("Archivo de Transferencia del " + FIDAT$ + "\\Cancelar\Continuar") <> 2 Then GoTo 99
     '
     LUXX$ = LUDAT$
     Dim BUFGRA1 As String * 16
     Dim BUFLEC1 As String * 32
     NY% = FILEOPEN%("C:\FLEXOFT\ATRACON.TRF", 0, 32)
     '
     VDAT$ = "": VOLD$ = ""
     For U& = 1 To LOF(NY%) / 32
       Get #NY%, U&, BUFLEC1$
       If TRIM$(Left$(BUFLEC1$, 16)) = "VARIAB.DAT" Then VDAT$ = VDAT$ + Mid$(BUFLEC1$, 17, 16)
       If TRIM$(Left$(BUFLEC1$, 16)) = "VARIAB.OLD" Then VOLD$ = VOLD$ + Mid$(BUFLEC1$, 17, 16)
     Next U&
     Close #NY%
     '
     WDAT$ = LOADFILE$(LUXX$ + "VARIAB.DAT")
     WOLD$ = LOADFILE$(LUXX$ + "VARIAB.OLD")
     
'esta validacion se suspende provisoriamente
'     If Left$(WDAT$, 64) <> Left$(VDAT$, 64) Or Left$(WOLD$, 64) <> Left$(VOLD$, 64) Then
'       Call MENSERR(24, "Imposible Transferir - Ejercicios no se Corresponden")
'       GoTo 99
'     End If
     '
     NY% = FILEOPEN%("C:\FLEXOFT\ATRACON.TRF", 0, 32)
     ATXA$ = "": JJ& = 0
     For U& = 1 To LOF(NY%) / 32
       Get #NY%, U&, BUFLEC1$
       ATXX$ = UCase$(TRIM$(Left$(BUFLEC1$, 16)))
       ARAC$ = LUXX$ + ATXX$
       '
       If ATXX$ = "" Then GoTo 49
       PPXX% = InStr(ATXX$, ".")
       If PPXX% > 1 Then ATXX$ = Left$(ATXX$, PPXX% - 1)
       For IAR% = 1 To QAR%
         If NARCH$(IAR%) = ATXX$ Then GoTo 49
       Next IAR%
       '
       'If ATXX$ <> ATXA$ Then
       If ARAC$ <> ATXA$ Then
         If NX% > 0 Then
           FIN& = LOF(NX%) / 16
           While JJ& < FIN&
             BUFGRA1$ = String$(16, 0)
             JJ& = JJ& + 1
             Put #NX%, JJ&, BUFGRA1$
           Wend
           Close #NX%: NX% = 0
         End If
         '
         JJ& = 0
         NX% = FILEOPEN%(ARAC$, 2, 16)    ' PARA BLANQUEAR ARCHIVO
         Close #NX%
         NX% = FILEOPEN%(ARAC$, 0, 16)    ' ABRE PARA RANDOM
         'ATXA$ = ATXX$
         ATXA$ = ARAC$
       End If
       '
       JJ& = JJ& + 1
       BUFGRA1$ = Mid$(BUFLEC1$, 17, 16)
       Put #NX%, JJ&, BUFGRA1$
       '
49   Next U&
     If NX% > 0 Then
       FIN& = LOF(NX%) / 16
       While JJ& < FIN&
         BUFGRA1$ = String$(16, 0)
         JJ& = JJ& + 1
         Put #NX%, JJ&, BUFGRA1$
       Wend
       Close #NX%: NX% = 0
     End If
     '
     Call COMUNI("Importacion Completada - Re-Inicie Contabilidad")
     Close
     Call FINAL("")
     '
   End If
   '
99 Command25.Enabled = True
End Sub

Private Sub Command26_Click()
   OPASICONT.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("CONTAGEN")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONTAGEN")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = EMPREAC$
    If InStr(EPAC$, "CIME-SOI") < 1 Then
         Call OPNOIN("Informes Normalizados")
       Else
         Call CONECRUN("#SCG47B", "Informes Standard Normalizados")
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 0
    ASICAJA.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 3
    ASICAJA.Show 1
    Unload ASICAJA
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
       Call ABREJER    ' Apertura de Nuevo Ejercicio Contable
    End If
99  Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Call OPNOIN("Asiento de Cierre de Ejercicio")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    CUECHECA$ = CONTROL$("", "CUECHECA")
    CUEICHEQ% = CUEINT%(CUECHECA$)
    If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Or CUEICHEQ% > NUCUEN% Then
       Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
       OPASICONT.Show 1
       GoTo 99
    End If
    '
    'Call BLOQARCH("ASICONT")
    TIASI% = 2
    ASICAJA.Show 1
    Unload ASICAJA
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    'Call BLOQARCH("ASICONT")
    TIASI% = 1
    ASIDIAR.Show 1
    Unload ASIDIAR
    Command9.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If CONTROL$("CONTAGEN", "SIDISTRI") = "DISTRIBU" Then
      Command25.Enabled = True
      DoEvents
    End If
    '
    If InStr(UCase$(EPAC$), "SABO") > 0 Then
      If Left$(UCase$(CurDir$), 2) = "C:" Then
        If MAXENCDIARIO% > 0 Then
          Call LEENCDIARIO%(1)
          If CVI(E3$) < COMEJ% Then
            COMEJ% = CVI(E3$)
            FINEJ% = COMEJ% + 32 * 13 - 1
          End If
        End If
      End If
    End If
    '
End Sub

Sub DIARIOGEN()
   '
   Call FRESHALL
   Call CARSELMESA
   MODOEMI% = COMALTER%("Puede Enitir el Libro Diario General\en Forma Mensual o General del Ejercicio\\Mensual\General")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then Exit Sub
   '
10 DES% = COMEJ%: HAS% = FINEJ%
   CORRESP$ = EJACT$
   If MODOEMI% = 1 Then
      Call SELECHO("SELMESA")
      MESASEL$ = VALACT1$("SELMESA")
      If TRIM$(MESASEL$) = "" Then
         Exit Sub
      End If
      '
      MMXX0$ = TRIM$(MESASEL$)
      MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
      If TRIM$(MMXX1$) = "" Then
         Exit Sub
      End If
      '
20    MESASEL$ = MMXX1$
      If DES% < COMEJ% Or HAS% > FINEJ% Then
         Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
         GoTo 10
      End If
      '
      CORRESP$ = MESASEL$
   End If
   '
   MODOXX2% = COMALTER%("Puede Incluir los Asientos de Caja y Bancos\o Dejar Solo los Asientos Resumen\\Asientos de C & B\Asientos Resumen")
   If MODOXX2% < 1 Or MODOXX2% > 2 Then GoTo 99
   Call GRACONTROL("CONTAGEN", "INCLUCAJ", Str$(MODOXX2% Mod 2))
   MODOXX3% = COMALTER%("Puede Re-Numerar Correlativamente los Asientos\o Expresar los Números Provisorios de Asientos\\Re-Numerar Asientos\Números Provisorios")
   If MODOXX3% < 1 Or MODOXX3% > 2 Then GoTo 99
   Call GRACONTROL("CONTAGEN", "RENUMERA", Str$(MODOXX3% Mod 2))
   '
   Screen.MousePointer = 11
   SUDEBE# = 0
   SUHABE# = 0
   SCG_PRIN.SORTLIST.Clear
   '
   FIN& = MAXDIARIO%
   For KK% = 1 To MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASDIA(KK%, 1)
      FEAS% = CVI(E15$)
      If MODOXX2% = 1 Then
        If Asc(E18$) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA
      End If
      RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
      RESOR$ = RESOR$ + "DI" + TRIM$(Str$(100000 + KK%))
      SCG_PRIN.SORTLIST.AddItem RESOR$
29 Next KK%
   '
   If MODOXX2% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEEPASCAJ(KKL&, 1)
       FEAS% = CVI(D15$)
       RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
       RESOR$ = RESOR$ + "CA" + TRIM$(Str$(100000 + KKL&))
       SCG_PRIN.SORTLIST.AddItem RESOR$
     Next KKL&
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = SCG_PRIN.SORTLIST.ListCount    ' MAXDIARIO%
   YATRA% = 0: JJ& = 0: ASIANT = 0: ASIANTX = 0
   '
   NUASDEF& = 0
   If MODOXX3% = 1 Then
      Screen.MousePointer = 1
      NUASDEF& = Val(InputBox("Re-Numera a Partir de", "Numeración Asientos Libro Diario", "1.0")) - 1
      If NUASDEF& < 0 Then Exit Sub
      Screen.MousePointer = 11
   End If
   '
   For KK% = 1 To FIN&     'MAXDIARIO%
      UU& = KK%
      Call TRACE(20, UU&, FIN&)
      RESOR$ = SCG_PRIN.SORTLIST.List(KK% - 1)
      REGII& = Val(Mid$(RESOR$, 13)) - 100000
      GRUXI$ = TRIM$(Mid$(RESOR$, 11, 2))
      SGX1% = 1
      FEASI% = Val(TRIM$(Left$(RESOR$, 10))) - 10000
      MODOLE% = 1
      If FEASI% >= DES% And FEASI% <= HAS% Then MODOLE% = 2
      If GRUXI$ = "DI" Then
         REGIJ% = REGII&
         Call LEEPASDIA(REGIJ%, MODOLE%)
         FEAS% = CVI(E15$)
         IPAS# = CVD(E16$)
         CUE% = CVI(E11$)
         NUAS = CVS(E12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
        Else
         SGX1% = (-1)
         Call LEEPASCAJ(REGII&, MODOLE%)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
      End If
      '
      If NUAS * SGX1% <> ASIANT Then
        NUASDEF& = NUASDEF& + 1
        ASIANT = NUAS * SGX1%
      End If
      '
      NUADE% = NUASDEF& Mod 32700
      If SGX1% > 0 Then
          Get #NX10%, REGII&, E10$
          EXXX$ = E10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          E10$ = EXXX$
          Put #NX10%, REGII&, E10$
        Else
          Get #NX6%, REGII&, D10$
          EXXX$ = D10$
          Call REPLA(EXXX$, MKI$(NUADE%), 7, 2)
          D10$ = EXXX$
          Put #NX6%, REGII&, D10$
      End If
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              SUDEBE# = SUDEBE# + IPAS#
            Else
              SUHABE# = SUHABE# - IPAS#
         End If
         GoTo 90
      End If
      '
      If YATRA% = 0 Then
         Z$ = REGBLAN$("DIARGEN")
         FEX = DES%
         DIA0% = DIANTE(FEX)
         Call REPLA(Z$, MKI$(DIA0%), 1, 2)
         Call REPLA(Z$, MKS$(99999999999#), 3, 4)
         Call REPLA(Z$, "Transporte", 7, 53)
         Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
         Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         Z$ = REGBLAN$("DIARGEN")
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
         YATRA% = 1
      End If
      '
      If FEAS% <= HAS% Then
        If NUAS * SGX1% <> ASIANTX Then
           ASIANTX = NUAS * SGX1%
           Z$ = REGBLAN$("DIARGEN")
           Call REPLA(Z$, MKI$(FEAS%), 1, 2)
           Call REPLA(Z$, MKS$(NUAS), 3, 4)
           If MODOXX3% = 1 Then
             Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
           End If
           Call REPLA(Z$, REFASIEN$, 7, 53)
           JJ& = JJ& + 1
           Call GRAREG("DIARGEN", Z$, JJ&)
         End If
         '
         Z$ = REGBLAN$("DIARGEN")
         Call REPLA(Z$, MKI$(FEAS%), 1, 2)
         Call REPLA(Z$, MKS$(NUAS), 3, 4)
         If MODOXX3% = 1 Then
            Call REPLA(Z$, MKS$(NUASDEF&), 3, 4)
         End If
         Call REPLA(Z$, REFE$, 7, 53)
         Call REPLA(Z$, DENOMI$, 60, 45)
         If IPAS# > 0 Then
              Call REPLA(Z$, MKD$(IPAS#), 105, 8)
            Else
              Call REPLA(Z$, MKD$((-1) * IPAS#), 113, 8)
         End If
         JJ& = JJ& + 1
         Call GRAREG("DIARGEN", Z$, JJ&)
      End If
90 Next KK%
   '
   Call SETULTREG("DIARGEN", JJ&)
   Call CIERRARCH("DIARGEN")
   '
   Call HEADERS("DIARGEN", "")
   Call HEADERS("DIARGEN", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUBDIAR", "Libro Diario General")
   '
99 End Sub

Sub SUDIARCAJ()
   '
   Call FRESHALL
   Call CARSELMESA
   MODOEMI% = COMALTER%("Puede Enitir el Subdiario de Caja y Bancos\en Forma Mensual o General del Ejercicio\\Mensual\General")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then Exit Sub
   '
10 DES% = COMEJ%: HAS% = FINEJ%
   CORRESP$ = EJACT$
   If MODOEMI% = 1 Then
      Call SELECHO("SELMESA")
      MESASEL$ = VALACT1$("SELMESA")
      If TRIM$(MESASEL$) = "" Then
         Exit Sub
      End If
      '
      MMXX0$ = TRIM$(MESASEL$)
      MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
      If TRIM$(MMXX1$) = "" Then
         Exit Sub
      End If
      '
20    MESASEL$ = MMXX1$
      If DES% < COMEJ% Or HAS% > FINEJ% Then
         Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
         GoTo 10
      End If
      '
      CORRESP$ = MESASEL$
   End If
   '
   Screen.MousePointer = 11
   SUDEBE# = 0
   SUHABE# = 0
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = MAXDIA&
   YATRA% = 0: JJ& = 0: ASIANT = 0
   For KKL& = 1 To MAXDIA&
      UU& = KKL&
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KKL&, 1)
      FEAS% = CVI(D15$)
      IPAS# = CVD(D16$)
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              SUDEBE# = SUDEBE# + IPAS#
            Else
              SUHABE# = SUHABE# - IPAS#
         End If
         GoTo 90
      End If
      '
      If YATRA% = 0 Then
         Z$ = REGBLAN$("SUDIACAJ")
         FEX = DES%
         DIA0% = DIANTE(FEX)
         Call REPLA(Z$, MKI$(DIA0%), 1, 2)
         Call REPLA(Z$, MKS$(99999999999#), 3, 4)
         Call REPLA(Z$, "Transporte", 7, 53)
         Call REPLA(Z$, MKD$(SUDEBE#), 105, 8)
         Call REPLA(Z$, MKD$(SUHABE#), 113, 8)
         '
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
         Z$ = REGBLAN$("SUDIACAJ")
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
         YATRA% = 1
      End If
      '
      If FEAS% <= HAS% Then
         Call LEEPASCAJ(KKL&, 2)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         '
         If NUAS <> ASIANT Then
           Z$ = REGBLAN$("SUDIACAJ")
           Call REPLA(Z$, MKI$(FEAS%), 1, 2)
           Call REPLA(Z$, MKS$(NUAS), 3, 4)
           Call REPLA(Z$, REFASIEN$, 7, 53)
           JJ& = JJ& + 1
           Call GRAREG("SUDIACAJ", Z$, JJ&)
           ASIANT = NUAS
         End If
         '
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
         '
         Z$ = REGBLAN$("SUDIACAJ")
         Call REPLA(Z$, MKI$(FEAS%), 1, 2)
         Call REPLA(Z$, MKS$(NUAS), 3, 4)
         Call REPLA(Z$, REFE$, 7, 53)
         Call REPLA(Z$, DENOMI$, 60, 45)
         If IPAS# > 0 Then
              Call REPLA(Z$, MKD$(IPAS#), 105, 8)
            Else
              Call REPLA(Z$, MKD$((-1) * IPAS#), 113, 8)
         End If
         JJ& = JJ& + 1
         Call GRAREG("SUDIACAJ", Z$, JJ&)
      End If
90 Next KKL&
   '
   Call SETULTREG("SUDIACAJ", JJ&)
   Call CIERRARCH("SUDIACAJ")
   '
   Call HEADERS("SUDIACAJ", "")
   Call HEADERS("SUDIACAJ", "Corresponde a: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUBCAJA", "Libro Subdiario de Caja y Bancos")
   '
End Sub
'
Sub SUCAJINEG()           ' Subdiario de Caja - Ingresos y Egresos
   '
   Dim NASISEL(32767)
   Call FRESHALL
   Screen.MousePointer = 11
   CCCX% = CUEINT%("")          ' para determinar NUCUEN%
   HAS% = HOY(HO$): If HAS% > FINEJ% Then HAS% = FINEJ%
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   DESX$ = "01" + Mid$(HASX$, 3)
   DES% = FECHANUM(DESX$)
   For I1% = 1 To NUCUEN%
      If InStr(UCase$(DENOCUE$(I1%)), "CAJA") > 0 Then
         CUEI% = I1%
         COCUE$ = CODCUE$(CUEI%)
         GoTo 5
      End If
   Next I1%
   '
5  VDEF$ = MKI$(DES%) + MKI$(HAS%) + COCUE$
   Screen.MousePointer = 1
   '
6  OBX$ = OBJPLA$("DESHASFECHCUEN", VDEF$)
   If OBX$ = "" Then Exit Sub
   VDEF$ = OBX$
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   COCUE$ = TRIM$(Mid$(OBX$, 5, 12))
   If DES% > HAS% Then GoTo 6
   If DES% < COMEJ% Then GoTo 6
   If HAS% > FINEJ% Then GoTo 6
   If COCUE$ = "" Then GoTo 6
   CUEI% = CUEINT%(COCUE$)
   If CUEI% < 1 Or CUEI% > NUCUEN% Then GoTo 6
   '
   FEX = DES%: DESX$ = FECHATEX$(FEX)
   FEX = HAS%: HASX$ = FECHATEX$(FEX)
   PERIO$ = "del " + DESX$ + " al " + HASX$
   '
8  MODOEMI% = COMALTER%("Puede Enitir el Subdiario\por Ingresos o Egresos\\Ingresos\Egresos\Cancelar")
   If MODOEMI% < 1 Or MODOEMI% > 2 Then GoTo 6
   '
   SIGNO% = 1: INEG$ = "Ingresos"
   If MODOEMI% = 2 Then SIGNO% = (-1): INEG$ = "Egresos"
   '
   CORRESP$ = TRIM$(COCUE$) + " " + TRIM$(DENOCUE$(CUEI%)) + " (" + INEG$ + ") " + PERIO$
   '
   'Identificar asientos con movimiento
   '
   Screen.MousePointer = 11
   FIN& = MAXDIA&: PRIMO& = 0: ULTMO& = 0
   YATRA% = 0: NUASANT = 0: CASISEL% = 0
   For KKL& = 1 To MAXDIA&
      UU& = KKL&
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KKL&, 1)
      CUE% = CVI(D11$)
      NUAS = CVS(D12$)
      FEAS% = CVI(D15$)
      IPAS# = CVD(D16$)
      '
      If FEAS% >= DES% Then
        If FEAS% <= HAS% Then
          If PRIMO& = 0 Then PRIMO& = KKL&
          ULTMO& = KKL&
          If CUE% = CUEI% Then
            If Sgn(IPAS#) = SIGNO% Then
              If NUAS <> NUASANT Then
                CASISEL% = CASISEL% + 1
                NASISEL(CASISEL%) = NUAS
                NUASANT = NUAS
              End If
            End If
          End If
        End If
      End If
   Next KKL&
   '
   If CASISEL% < 1 Then
      Call COMUNI("Sin Movimientos en el Período Elegido")
      GoTo 8
   End If
   '
   ' CARGAR MOVIMIENTOS DE CAJA Y BANCOS
   FIN& = ULTMO& - PRIMO& + 1
   NUASANT = 0: ACUM# = 0
   YATRA% = 0: JJ& = 0
   For KKL& = PRIMO& To ULTMO&
      UU& = KKL& - PRIMO& + 1
      Call TRACE(20, UU&, FIN&)
      Call LEEPASCAJ(KKL&, 1)
      CUE% = CVI(D11$)
      NUAS = CVS(D12$)
      '
      If NUAS <> NUASANT Then
        SELE% = 0
        For UI% = 1 To CASISEL%
          If NASISEL(UI%) = NUAS Then
            SELE% = 1
            GoTo 20
          End If
        Next UI%
      End If
      '
20    If SELE% = 1 Then
        Z$ = REGBLAN$("SUDINEG")
        If NUAS <> NUASANT Then
           Call LEEPASCAJ(KKL&, 2)
           FEX = CVI(D15$)
           Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
           Call REPLA(Z$, MKS$(NUAS), 9, 4)
           REFEL$ = REFASIEN$
           REFE$ = REPLACAR$(REFEL$, "@", "A")
           Call REPLA(Z$, REFE$, 13, 48)
           JJ& = JJ& + 1
           Call GRAREG("SUDINEG", Z$, JJ&)
           Z$ = REGBLAN$("SUDINEG")
        End If
        '
        IPAS# = CVD(D16$)
        REFEL$ = "- " + TRIM$(DETAPAS$)
        REFE$ = REPLACAR$(REFEL$, "@", "A")
        Call REPLA(Z$, REFE$, 13, 48)
        Call REPLA(Z$, TRIM$(CODCUE$(CUE%)), 61, 12)
        If IPAS# > 0 Then
             Call REPLA(Z$, MKD$(IPAS#), 73, 8)
           Else
             Call REPLA(Z$, MKD$(Abs(IPAS#)), 81, 8)
        End If
        If CUE% = CUEI% Then
           ACUM# = ACUM# + Abs(IPAS#)
           Call REPLA(Z$, MKD$(ACUM#), 89, 8)
        End If
        JJ& = JJ& + 1
        Call GRAREG("SUDINEG", Z$, JJ&)
        '
        NUASANT = NUAS
        '
      End If
      '
   Next KKL&
   '
   Call SETULTREG("SUDINEG", JJ&)
   Call CIERRARCH("SUDINEG")
   '
   Call HEADERS("SUDINEG", "")
   Call HEADERS("SUDINEG", "Cuenta: " + CORRESP$)
   '
   Screen.MousePointer = 1
   Call CONECRUN("*SUDINEG", "Subdiario de Caja (Ingresos y Egresos)")
   '
   GoTo 6
   '
End Sub
'
Sub ABREJER()
   '
   Dim ARQ$(16)
   '
   If EJER% <> 1 Then
      Call MENSERR(24, "Apertura Imposible\Ingrese al Ejercicio Actual")
      Exit Sub
   End If
   '
   TTXXX$ = "Al Abrir un Nuevo Ejercicio Contable\el Ejercicio ACTUAL pasa a ser ANTERIOR\y los Datos del ANTERIOR se Pierden !!!\   \Desea Continuar o Cancelar Apertura ?"
   If COMALTER%(TTXXX$ + "\\Cancelar Apertura\Continuar") <> 2 Then
      Exit Sub
   End If
   '
   TTXXX$ = "Para su Propia Seguridad de sus Datos\es IMPRESCINDIBLE Haber Confeccionado un BACKUP\Reciente de los Archivos '*.OLD'\de su Base de Datos !!!\   \Sólo si Tiene ese Backup en Lugar Seguro\Pinche sobre 'Si, Continuar', Caso Contrario sobre 'No, Cancelar'"
   If COMALTER%(TTXXX$ + "\\No, Cancelar\Si, Continuar") <> 2 Then
      Exit Sub
   End If
   '
   VDEF$ = String$(80, 0)
   Call REPLA(VDEF$, MKI$(COMEJ%), 1, 2)
   Call REPLA(VDEF$, EJACT$, 3, 32)
   FEX = FINEJ%
   CMJ1% = DIPOST(FEX)
   Call REPLA(VDEF$, MKI$(CMJ1%), 41, 2)
   Call REPLA(VDEF$, "", 43, 32)
   '
10 OBX$ = OBJPLA$("ABREJER", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   DENUE$ = TRIM$(Mid$(OBX$, 43, 32))
   If DENUE$ = "" Then
      Call MENSERR(24, "Falta Denominación del Nuevo Ejercicio")
      GoTo 10
   End If
   '
   CMJACT% = COMEJ%
   APNUE% = CVI(Mid$(OBX$, 41, 2))
   If APNUE% < CMJACT% Then
      Call MENSERR(24, "Nueva Apertura es Anterior\al Cierre del Ejercicio Actual")
      GoTo 10
   End If
   '
   FEX = APNUE%
   APNU$ = FECHATEX$(FEX)
   ANX% = Val(Mid$(APNU$, 7, 2)) + 1: If ANX% > 99 Then ANX% = ANX% - 100
   AN$ = TRIM$(Str$(ANX%)): While Len(AN$) < 2: AN$ = "0" + AN$: Wend
   FE$ = Left$(APNU$, 6) + AN$
   FEX = FECHANUM(FE$)
   CIENUE% = DIANTE(FEX)
   '
   ' MENSAJE DE ADVERTENCIA y confirmacion
   '
   Screen.MousePointer = 11
   LU$ = LUDAT$
   NSE% = 1           ' numero de subempresas
   On Error Resume Next
   Kill LU$ + "*.NEW"
   On Error GoTo 0
   '
   N1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
   Dim X As String * 32
   If APNUE% = CMJACT% Then
      Get #N1%, 2, X$
      LSet X$ = Left$(DENUE$ + Space$(28), 28) + MKI$(APNUE%) + MKI$(CIENUE%)
      Put #N1%, 2, X$
      Close #N1%
      GoTo 99
   End If
   '
   N2% = FILEOPEN%(LU$ + "VARIAB.NEW", 0, 32)
   Dim Y As String * 32
   '
   For I% = 1 To LOF(N1%) / 32
      Get #N1%, I%, X$
      Z$ = X$: If I% > 10 Then Z$ = String$(32, 0)
      LSet Y$ = Z$
      Put #N2%, I%, Y$
   Next I%
   '
   Get #N2%, 2, Y$: Put #N2%, 3, Y$
   LSet Y$ = Left$(DENUE$ + Space$(28), 28) + MKI$(APNUE%) + MKI$(CIENUE%)
   Put #N2%, 2, Y$
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "SUBEMPRE.DAT", 0, 64)
   Dim X64 As String * 64
   N2% = FILEOPEN%(LU$ + "SUBEMPRE.NEW", 0, 64)
   Dim Y64 As String * 64
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "CUENTAS.DAT", 0, 64)
   N2% = FILEOPEN%(LU$ + "CUENTAS.NEW", 0, 64)
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
      NUCUENUE% = I%
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "DENOCUE.DAT", 0, 128)
   Dim X128 As String * 128
   N2% = FILEOPEN%(LU$ + "DENOCUE.NEW", 0, 128)
   Dim Y128 As String * 128
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "INVECUE.DAT", 0, 16)
   Dim X16 As String * 16
   N2% = FILEOPEN%(LU$ + "INVECUE.NEW", 0, 16)
   Dim Y16 As String * 16
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "PLAGRAL.DAT", 0, 16)
   N2% = FILEOPEN%(LU$ + "PLAGRAL.NEW", 0, 16)
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16$
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "PLANALT.DAT", 0, 16)
   N2% = FILEOPEN%(LU$ + "PLANALT.NEW", 0, 16)
   '
   For I% = 1 To LOF(N1%) / 16
      Get #N1%, I%, X16$
      Z$ = X16$
      LSet Y16$ = Z$
      Put #N2%, I%, Y16$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "TRADENO.DAT", 0, 128)
   N2% = FILEOPEN%(LU$ + "TRADENO.NEW", 0, 128)
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128$
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "DENENCA.DAT", 0, 128)
   N2% = FILEOPEN%(LU$ + "DENENCA.NEW", 0, 128)
   '
   For I% = 1 To LOF(N1%) / 128
      Get #N1%, I%, X128$
      Z$ = X128$
      LSet Y128$ = Z$
      Put #N2%, I%, Y128$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "INDAJU.DAT", 0, 132)
   Dim X132 As String * 132
   N2% = FILEOPEN%(LU$ + "INDAJU.NEW", 0, 132)
   Dim Y132 As String * 132
   '
   For I% = 1 To LOF(N1%) / 132
      Get #N1%, I%, X132$
      Z$ = X132$
      LSet Y132$ = Z$
      Put #N2%, I%, Y132$
   Next I%
   Close #N2%: Close #N1%
   '
   N2% = FILEOPEN%(LU$ + "SALCAJA.NEW", 0, 80)
   Dim Y80 As String * 80
   '
   For I% = 1 To NUCUENUE%
      LSet Y80$ = String$(80, 45)
      Put #N2%, I%, Y80$
   Next I%
   Close #N2%: Close #N1%
   '
   Dim X112 As String * 112
   Dim Y112 As String * 112
   '
   For SX% = 1 To NSE%
      '
      SXX$ = Mid$(Str$(SX% - 1), 2)
      N1% = FILEOPEN%(LU$ + "CIERRE" + SXX$ + ".NEW", 0, 112)
      '
      For I% = 1 To NUCUENUE%
         LSet Y112$ = String$(112, 0)
         Put #N1%, I%, Y112$
      Next I%
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ENCDIAR" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "DIARIO" + SXX$ + ".NEW", 0, 32)
      LSet Y32$ = String$(32, 0)
      Put #N1%, 1, Y32$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "IDIARIO" + SXX$ + ".NEW", 0, 128)
      LSet Y128$ = String$(128, 0)
      Put #N1%, 1, Y128$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ENCAJA" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "SUCAJA" + SXX$ + ".NEW", 0, 64)
      LSet Y64$ = String$(64, 0)
      Put #N1%, 1, Y64$
      Close #N1%
      '
      N1% = FILEOPEN%(LU$ + "ISUCAJA" + SXX$ + ".NEW", 0, 128)
      LSet Y128$ = String$(128, 0)
      Put #N1%, 1, Y128$
      Close #N1%
      '
   Next SX%
   '
   N1% = FILEOPEN%(LU$ + "INFORMES.DAT", 0, 6)
   Dim X6 As String * 6
   N2% = FILEOPEN%(LU$ + "INFORMES.NEW", 0, 6)
   Dim Y6 As String * 6
   '
   For I% = 1 To LOF(N1%) / 6
      Get #N1%, I%, X6$
      Z$ = X6$
      LSet Y6$ = Z$
      Put #N2%, I%, Y6$
   Next I%
   Close #N2%: Close #N1%
   '
   N1% = FILEOPEN%(LU$ + "ENCAINF.DAT", 0, 64)
   N2% = FILEOPEN%(LU$ + "ENCAINF.NEW", 0, 64)
   '
   For I% = 1 To LOF(N1%) / 64
      Get #N1%, I%, X64$
      Z$ = X64$
      LSet Y64$ = Z$
      Put #N2%, I%, Y64$
      NUCUENUE% = I%
   Next I%
   Close #N2%: Close #N1%
   '
   ARQ$(1) = "VARIAB"
   ARQ$(2) = "SUBEMPRE"
   ARQ$(3) = "CUENTAS"
   ARQ$(4) = "DENOCUE"
   ARQ$(5) = "INVECUE"
   ARQ$(6) = "PLAGRAL"
   ARQ$(7) = "PLANALT"
   ARQ$(8) = "TRADENO"
   ARQ$(9) = "DENENCA"
   ARQ$(10) = "INDAJU"
   ARQ$(11) = "SALCAJA"
   ARQ$(12) = "INFORMES"
   ARQ$(13) = "ENCAINF"
   '
   For I% = 1 To 13
      ARXQ$ = ARQ$(I%)
      On Error Resume Next
      Kill LU$ + ARXQ$ + ".OLD"
      '
      FileCopy LU$ + ARXQ$ + ".DAT ", LU$ + ARXQ$ + ".OLD"
      FileCopy LU$ + ARXQ$ + ".NEW ", LU$ + ARXQ$ + ".DAT"
      '
      Kill LU$ + ARXQ$ + ".NEW"
      On Error GoTo 0
   Next I%
   '
   ARQ$(1) = "CIERRE"
   ARQ$(2) = "ENCDIAR"
   ARQ$(3) = "DIARIO"
   ARQ$(4) = "IDIARIO"
   ARQ$(5) = "ENCAJA"
   ARQ$(6) = "SUCAJA"
   ARQ$(7) = "ISUCAJA"
   '
   For I% = 1 To 7
      For SX% = 1 To NSE%
         SXX$ = TRIM$(Str$(SX% - 1))
         ARXQ$ = ARQ$(I%) + SXX$
         On Error Resume Next
         Kill LU$ + ARXQ$ + ".OLD"
         '
         FileCopy LU$ + ARXQ$ + ".DAT ", LU$ + ARXQ$ + ".OLD"
         FileCopy LU$ + ARXQ$ + ".NEW ", LU$ + ARXQ$ + ".DAT"
         '
         Kill LU$ + ARXQ$ + ".NEW"
         On Error GoTo 0
      Next SX%
      '
   Next I%
   '
99 Screen.MousePointer = 1
   Call COMUNI("Apertura Completa.\  \Re-Inicie Ahora el Sistema\de Contabilidad General.")
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Sub BORRARESCAJ()
   '
   REFE$ = TRIM$("RESUMEN DE CAJA-BANCOS AL")
   LS% = MAXENCDIARIO%
   JJ& = 0
   For L1% = 1 To LS%
      Call LEENCDIARIO(L1%)
      If Left$(E7$, 25) = REFE$ Then             ' ENCONTRO EL ASIENTO
         XX$ = REGBLAN$("SELASIEN")
         Call REPLA(XX$, E1$, 1, 4)
         Call REPLA(XX$, E7$, 5, 40)
         JJ& = JJ& + 1
         Call GRAREG("!SELASIEN", XX$, JJ&)
      End If
   Next L1%
   '
   Call SETULTREG("!SELASIEN", JJ&)
   Call CIERRARCH("!SELASIEN")
   Call FRESHECHO("!SELASIEN")
   '
   If ULTREG&("!SELASIEN") < 1 Then
      Screen.MousePointer = 11
      Call ABRESUCAJA
      FIN& = MAXDIA&
      For I& = 1 To MAXDIA&
        U& = I&
        Call TRACE(20, U&, FIN&)
        Call LEESUCAJA(I&)
        NRXX = CVS(D19$)
        If NRXX > 0.5 Then
          DXX$ = D10$
          Call REPLA(DXX$, MKS$(0), 55, 4)
          LSet D10$ = DXX$
          Put #NX6%, I&, D10$
        End If
      Next I&
      Close #NX6%: NX6% = 0
      Call MENSERR(24, "No hay Asientos Resumen\que Puedan Anularse.")
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   Call SELECHO("!SELASIEN")
   NROAS = Val(VALACT1$("!SELASIEN"))
   If NROAS < 1 Then
      Exit Sub
   End If
   '
   Call BLOQARCH("ACONT" + EXTE$)
   Call BLOQARCH("ASICONT")
   Screen.MousePointer = 11
   FIN& = MAXDIA&
   For I& = 1 To MAXDIA&
      U& = I&
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(I&)
      NRXX = CVS(D19$)
      If Abs(NRXX - NROAS) < 0.05 Then
         DXX$ = D10$
         Call REPLA(DXX$, MKS$(0), 55, 4)
         LSet D10$ = DXX$
         Put #NX6%, I&, D10$
      End If
   Next I&
   '
   Close #NX6%: NX6% = 0
   '
   FIN& = MAXDIARIO%
   MMXX% = MAXDIARIO%
   I1% = 0
   For IC% = 1 To MAXDIARIO%
      Call LEEDIARIO(IC%)
      If Abs(CVS(E12$) - NROAS) > 0.01 Then
         I1% = I1% + 1
         Put #NX10%, I1%, E10$
      End If
   Next IC%
   '
   For IC% = I1% + 1 To MMXX%
      LSet E10$ = String$(32, 0)
      Put #NX10%, IC%, E10$
   Next IC%
   Close #NX10%: NX10% = 0
   '
   FIN& = MAXENCDIARIO%
   MMXX% = MAXENCDIARIO%
   I1% = 0
   For IC% = 1 To MAXENCDIARIO%
      Call LEENCDIARIO(IC%)
      If Abs(CVS(E1$) - NROAS) > 0.01 Then
         I1% = I1% + 1
         Put #NX9%, I1%, E0$
      End If
   Next IC%
   '
   For IC% = I1% + 1 To MMXX%
      LSet E0$ = String$(64, 0)
      Put #NX9%, IC%, E0$
   Next IC%
   Close #NX9%: NX9% = 0
   '
   Screen.MousePointer = 1
   '
   Call CIERRARCH("*.*")
   Close
   End
   '
End Sub

Private Sub Form_Resize()
   '
   If WindowState = 1 Then
     Call CIERRACONTA
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRACONTA
   Call FINAL("")
End Sub
