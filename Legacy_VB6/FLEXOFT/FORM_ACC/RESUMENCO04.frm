VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form RESUMENCO04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión - Resúmenes Mensuales de Compras"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command11 
      Caption         =   "I.Brutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2050
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Dispositivo de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   315
      TabIndex        =   36
      Top             =   4500
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   3570
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   50
         Top             =   500
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   1995
         TabIndex        =   38
         Top             =   420
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   525
         TabIndex        =   37
         Top             =   420
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros Reportes Mensuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   5250
      TabIndex        =   25
      Top             =   105
      Width           =   4635
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Resumen Operativo"
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
         Left            =   315
         TabIndex        =   44
         Top             =   2625
         Width           =   4005
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   45
            Top             =   350
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
               Height          =   600
               Left            =   0
               Picture         =   "RESUMENCO04.frx":1188
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen Men- sual por Cuenta"
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
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Procesos  Contables"
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
         Left            =   315
         TabIndex        =   39
         Top             =   3885
         Width           =   4005
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   3150
            TabIndex        =   40
            Top             =   350
            Width           =   3210
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
               Left            =   2600
               Picture         =   "RESUMENCO04.frx":15CA
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Re-Imputación de Comprobantes de Compras"
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
               Picture         =   "RESUMENCO04.frx":1714
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Consulta / Listado por Cuenta Contable"
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
               Left            =   2020
               Picture         =   "RESUMENCO04.frx":1A1E
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Consulta de Asiento Contable por Pantalla"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Contabilización de Compras"
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
               TabIndex        =   43
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ordenes de Pago y Ajustes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   315
         TabIndex        =   26
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   33
            Top             =   1260
            Width           =   2430
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
               Left            =   -20
               Picture         =   "RESUMENCO04.frx":1D28
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Liqui- dación Fondo Fijo"
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
               TabIndex        =   35
               Top             =   105
               Width           =   1800
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   27
            Top             =   840
            Width           =   2430
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
               Left            =   -20
               Picture         =   "RESUMENCO04.frx":1E72
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle por Or- denes de Pago"
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
               TabIndex        =   29
               Top             =   0
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   30
            Top             =   420
            Width           =   2430
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
               Picture         =   "RESUMENCO04.frx":217C
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes y Otros Movimientos"
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
               Index           =   0
               Left            =   735
               TabIndex        =   32
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   1275
      Left            =   315
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   6
      Top             =   210
      Width           =   4635
      Begin VB.TextBox Text4 
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
         Index           =   0
         Left            =   1575
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   1065
      End
      Begin VB.TextBox Text4 
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
         Index           =   1
         Left            =   3045
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   750
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   7
         Text            =   " "
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2730
         TabIndex        =   14
         Top             =   735
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   13
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3885
         TabIndex        =   12
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -735
         TabIndex        =   11
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   8
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Proveed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":2486
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Alrchivo Maestro de Proveedores"
      Top             =   4570
      Width           =   855
   End
   Begin VB.CommandButton Command14 
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
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":2790
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   3730
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Gestión de Cuentas Corrientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   315
      TabIndex        =   3
      Top             =   1575
      Width           =   4635
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   315
         TabIndex        =   15
         Top             =   315
         Width           =   4005
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   22
            Top             =   1320
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
               Left            =   0
               Picture         =   "RESUMENCO04.frx":2BD2
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación de Compras"
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
               Left            =   840
               TabIndex        =   24
               Top             =   105
               Width           =   1200
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   19
            Top             =   870
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
               Height          =   600
               Left            =   -20
               Picture         =   "RESUMENCO04.frx":2EDC
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Débito"
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
               TabIndex        =   21
               Top             =   105
               Width           =   1410
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   16
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
               Height          =   600
               Left            =   0
               Picture         =   "RESUMENCO04.frx":31E6
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Notas de Crédito"
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
               Left            =   735
               TabIndex        =   18
               Top             =   105
               Width           =   1695
            End
         End
      End
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
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":34F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton Command8 
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
      Height          =   800
      Left            =   10080
      Picture         =   "RESUMENCO04.frx":37FA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2890
      Width           =   855
   End
End
Attribute VB_Name = "RESUMENCO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
10 NUFACLI$ = REGSEL$("IFACOMP")
   If Len(NUFACLI$) <= 4 Then GoTo 99
   NF$ = Left$(NUFACLI$, 18)
   NC% = CVI(Mid$(NUFACLI$, 59, 2))
   IMPO# = CVD(Mid$(NUFACLI$, 51, 8))
   If IMPO# < 0.005 Then
      Call MENSERR(24, "Documento Anulado")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("CUECORR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      REG& = FIN& + 1 - U&
      X$ = REGLEIDO$("CUECORR", REG&)
      If Mid$(X$, 7, 18) = NF$ Then
         Y6$ = Mid$(X$, 199, 80)
         GoTo 20
      End If
   Next U&
   Call MENSERR(24, "Documento no Encontrado - Consulte.")
   GoTo 99
   '
20 ATRI$ = TRIM$(Left$(NUFACLI$, 50)) + " ($ " + TRIM$(CSTRING$(MKD$(IMPO#), 4, 12, 2, 2)) + ")"
   '
   Call BLANARCH("!CARICOM")
   For J% = 1 To 8
      IPA# = 0
      CUE% = CVI(Mid$(Y6$, 10 * J% - 9, 2))
      If CUE% > 0 Then
         IPA# = CVD(Mid$(Y6$, 10 * J% - 7, 8))
         Y$ = REGBLAN$("!CARICOM")
         Call REPLA(Y$, MKI$(CUE%), 1, 2)
         Call REPLA(Y$, MKD$(IPA#), 3, 8)
         Call REGAPP("!CARICOM", Y$)
      End If
   Next J%
   Call CIERRARCH("!CARICOM")
   '
30 VTB% = VENTABU%("CARIFACO/" + ATRI$)
   If VTB% < 0 Then GoTo 10
   '
   TOTIMPU# = 0: Y6$ = ""
   For U& = 1 To ULTREG&("!CARICOM")
       Y$ = REGLEIDO$("!CARICOM", U&)
       CUEB$ = Left$(Y$, 2)
       If CVI(CUEB$) > 0 Then
          DECUE$ = ECOARCH$("CUECON", CUEB$)
          If DECUE$ <> "" Then
             IMBB$ = Mid$(Y$, 3, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 30
             End If
             Y6$ = Y6$ + CUEB$ + IMBB$
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
          End If
       End If
   Next U&
   '
   If Abs(IMPO# - TOTIMPU#) >= 0.005 Then
      Call MENSERR(24, "Suma de Imputaciones Difiere de Total")
      GoTo 30
   End If
   '
   Y6$ = Y6$ + String$(80, 0)
   X$ = REGLEIDO$("CUECORR", REG&)
   Call REPLA(X$, Y6$, 199, 80)
   Call GRAREG("CUECORR", X$, REG&)
   Call CIERRARCH("CUECORR")
   '
   GoTo 10
   '
99 Call FRESHECHO("IFACOMP")
   Command10.Enabled = True
   '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   OPX% = COMALTER%("Opciones de Trabajo:\\Subdiario I.Brutos\Archivo Transferencia")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If DES% < 1 Then GoTo 99
    '
   If OPX% = 1 Then
       CORRESP$ = PERIO$
       Call HEADERS("SUBRETI", "")
       Call HEADERS("SUBRETI", "Corresponde a: " + CORRESP$)
       J& = 0: FIN& = ULTREG("SURETIB")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("SURETIB", U&)
         FEMIX% = CVI(Left$(XX$, 2))
         If FEMIX% >= DES% Then
           If FEMIX% <= HAS% Then
             J& = J& + 1
             Call GRAREG("SUBRETI", XX$, J&)
           End If
         End If
       Next U&
       Call SETULTREG("SUBRETI", J&)
       Call FINAL("*SUBRETI")
     ElseIf OPX% = 2 Then
       CommonDialog1.CancelError = True
       CommonDialog1.DialogTitle = "Crear Archivo de Transferencia"
       'deshabilitado porque si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
       CommonDialog1.Filter = "Todos los Archivos (*.*)|*.*|"
       CommonDialog1.FilterIndex = 1
       On Error GoTo 99
       CommonDialog1.ShowOpen
       On Error GoTo 0
       '
       ANCHIMPOR% = 11
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "DOSIVA") > 0 Then ANCHIMPOR% = 10
       '
       ASALID$ = UCase$(CommonDialog1.filename)
       NX% = FILEOPEN%(ASALID$, 2, 128)
       CAREGI% = 0
       For U& = 1 To ULTREG&("SURETIB")
20       XX$ = REGLEIDO$("SURETIB", U&)
         FEMIX% = CVI(Left$(XX$, 2))
         If FEMIX% >= DES% Then
           If FEMIX% <= HAS% Then
             TTXX$ = Space$(45)
               CUITTX$ = TRIM$(Mid$(XX$, 51, 14))
               If TRIM$(CUITTX$) = "" Or VALICUIT%(CUITTX$) <> 1 Then
                 CUITTX$ = InputBox$("Ingrese Numero de CUIT", TRIM$(Mid$(XX$, 21, 30)))
                 Call REPLA(XX$, CUITTX$, 51, 14)
                 Call GRAREG("SURETIB", XX$, U&)
                 GoTo 20
               End If
             Call REPLA(TTXX$, CUITTX$, 1, 13)
               FEX = CVI(Left$(XX$, 2))
               FEXU$ = FECHATEX$(FEX)
               FEXU$ = Left$(FEXU$, 6) + "20" + Mid$(FEXU$, 7, 2)
             Call REPLA(TTXX$, FEXU$, 14, 10)
               SUCUX$ = "0001"
             Call REPLA(TTXX$, SUCUX$, 24, 4)
               NUCOX$ = TRIM$(Str$(CVS(Mid$(XX$, 3, 4))))
               While Len(NUCOX$) < 8: NUCOX$ = "0" + NUCOX$: Wend
             Call REPLA(TTXX$, NUCOX$, 28, 8)
               IMPOX$ = REPLACAR$(TRIM$(Format$(CVD(Mid$(XX$, 87, 8)), "#########0.00")), ".", ",")
               While Len(IMPOX$) < ANCHIMPOR%: IMPOX$ = "0" + IMPOX$: Wend
             Call REPLA(TTXX$, IMPOX$, 36, ANCHIMPOR%)
             Print #NX%, TTXX$
             CAREGI% = CAREGI% + 1
           End If
         End If
       Next U&
       Close #NX%
       Call COMUNI("Archivo de Transferencia '" + ASALID$ + "'\Generado con " + TRIM$(Str$(CAREGI%)) + " Registros.")
   End If
99 Command11.Enabled = True
   On Error GoTo 0
End Sub

Private Sub Command12_Click()
    '
    If Option1.Value = True Then
         MOX% = ALTERNA%("Consulta Selectiva por Cliente\Consulta por Codigo\Consulta General por Cliente\Consulta General por Pieza")
         On MOX% GoTo 203, 204, 205, 206
       ElseIf Option2.Value = True Then
         MOX% = ALTERNA%("Listado General por Cliente\Listado por Codigo de Producto")
         On MOX% GoTo 201, 202
    End If
    Exit Sub
    '
201 SALI$ = "*REVECLI": GoTo 999
202 SALI$ = "*REVEPIE": GoTo 999
203 SALI$ = "*COVECLI": GoTo 999
204 SALI$ = "*COVEPIE": GoTo 999
205 SALI$ = "*COGECLI": GoTo 999
206 SALI$ = "*COGEPIE": GoTo 999
    '
999 Call FINAL(SALI$)
    '
End Sub

Private Sub Command16_Click()
    MOX% = ALTERNA%("Resumen Estadistico de Ventas\Margen Bruto de Ventas")
    On MOX% GoTo 203, 204
    Exit Sub
    '
203 Call CONECRUN("#GESTAVE", "Resumen Estadistico de Ventas")
    Exit Sub
    '
204 Call CONECRUN("#MABRUVE", "Analisis Margen Bruto de Ventas")
    Exit Sub
    '
End Sub

Private Sub Command17_Click()
    If Option1.Value = True Then
         Call COMUNI("Opción no Disponible por Pantalla.\Emita Listado Impreso\o Exporte a Archivo Texto.")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call FINAL("*RESUSOR")
    End If
End Sub

Private Sub Command18_Click()
    Call CONECRUN("#GEDEFAC", "Detalle por Número de Factura")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call MENSERR(24, "No hay Opciones de Configuración\para esta Aplicación.")
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Call CIERRARCH("*.*")
   If EXISTE%("ARCH_GES.EXE") > 0 Then
       Call CONECRUN("ARCH_GES/AMAPRO", "Archivos Maestros")
     Else
       Call CONECRUN("AMAPRO", "Archivo Maestro de Proveedores")
   End If
End Sub

Private Sub Command19_Click()
  OPX% = ALTERNA%("Base Subdiario\Analisis por Cuenta\Analisis por Proveedor")
  Select Case OPX%
  Case 1
    If Option1.Value = True Then
         Call CONECRUN("*COSUCOM", "Consulta por Cuenta Contable")
         Exit Sub
      ElseIf Option2.Value = True Then
         Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
    End If
  Case 2
    If Option1.Value = True Then
         Call CONECRUN("*ANAICOM", "Analisis por Cuenta Contable")
      Else
         Call CONECRUN("*SORANAI", "Listado Analitico de Imputacion de Compras")
    End If
  Case 3
    If Option1.Value = True Then
20       Call CONECRUN("*ANAICOP", "Analisis Contable por Proveedor")
      Else
         Call MENSERR(24, "Solo Disponible por Pantalla")
         GoTo 20
    End If
  End Select
End Sub

Sub Command2_Click()
    MOX% = ALTERNA%("Cuentas Corrientes\Cuentas de Gastos")
    On MOX% GoTo 203, 204
    Exit Sub
    '
203 Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    Call HEADERS("DETPAGC", "")
    Call HEADERS("DETPAGC", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODPAGC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDPAGC")
    End If
    Exit Sub
    '
204 Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    Call HEADERS("DETPAGG", "")
    Call HEADERS("DETPAGG", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODPAGG")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDPAGG")
    End If
    Exit Sub
    '
End Sub
'

Private Sub Command20_Click()
    Call CONASIDIACOM
End Sub

Private Sub Command21_Click()
    FORLICOMIS.MESAN.TEXT = MESAN.TEXT
    FORLICOMIS.Text4(0).TEXT = Text4(0).TEXT
    FORLICOMIS.Text4(1).TEXT = Text4(1).TEXT
    FORLICOMIS.Show 1
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENCO.Text4(1).TEXT)
    Call HEADERS("DETFACC", "")
    Call HEADERS("DETFACC", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFACC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFACC")
    End If
End Sub

Sub Command4_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENCO.Text4(1).TEXT)
    Call HEADERS("DETDEBC", "")
    Call HEADERS("DETDEBC", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODDEBC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDDEBC")
    End If
End Sub


Sub Command5_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENCO.Text4(1).TEXT)
    Call HEADERS("DETCREC", "")
    Call HEADERS("DETCREC", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODCREC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDCREC")
    End If
End Sub

Sub Command6_Click()
    If Option1.Value = True Then
         Call FINAL("*RESUCCP")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIRECCP")
    End If
End Sub

Private Sub Command7_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    Call HEADERS("DETOTRC", "")
    Call HEADERS("DETOTRC", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODOTRC")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDOTRC")
    End If
End Sub

Private Sub Command8_Click()
    Call ACCDIR("COGECOM")
End Sub

Private Sub Command9_Click()
    Screen.MousePointer = 11
    C1$ = TRIM$(RESUMENCO.MESAN.TEXT)
    C2$ = TRIM$(RESUMENCO.Text4(0).TEXT)
    C3$ = TRIM$(RESUMENCO.Text4(1).TEXT)
    Call HEADERS("DETFOFI", "")
    Call HEADERS("DETFOFI", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
    Screen.MousePointer = 1
    If Option1.Value = True Then
         Call FINAL("*CODFOFI")
      ElseIf Option2.Value = True Then
         Call FINAL("*LIDFOFI")
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If InStr(UCase$(EPAC$), "SABO") Or InStr(UCase$(EPAC$), "DOSIVAC") > 0 Then
      Command11.Visible = True
    End If
End Sub

Private Sub MESAN_GOTFOCUS()
    Command1.SetFocus
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command1.SetFocus
End Sub
