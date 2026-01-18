VERSION 5.00
Begin VB.Form FiNVEN04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios de Materiales - Generación y Consulta"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00400000&
      Height          =   6735
      Left            =   8820
      ScaleHeight     =   6675
      ScaleWidth      =   1950
      TabIndex        =   51
      Top             =   0
      Width           =   2010
      Begin VB.CommandButton Command13 
         Caption         =   "Mast."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1730
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   55
         Top             =   3360
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   12000
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
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   900
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   210
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "FINVEN04.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2520
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FINVEN04.frx":08A8
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3675
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   7665
      ScaleHeight     =   585
      ScaleWidth      =   1635
      TabIndex        =   26
      Top             =   7455
      Width           =   1695
      Begin VB.CommandButton Command16 
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
         Picture         =   "FINVEN04.frx":27CA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Costos Informados"
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
         TabIndex        =   28
         Top             =   105
         Width           =   960
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Fábrica"
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
      Left            =   4515
      TabIndex        =   19
      Top             =   7770
      Width           =   3585
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   23
         Top             =   420
         Width           =   2535
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
            Picture         =   "FINVEN04.frx":2914
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales por Material"
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
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   20
         Top             =   840
         Width           =   2535
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
            Left            =   -20
            Picture         =   "FINVEN04.frx":2A5E
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Número de Orden de Fabricación"
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
            TabIndex        =   22
            Top             =   0
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Proveedores (Consignación)"
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
      Left            =   210
      TabIndex        =   9
      Top             =   7665
      Width           =   3585
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   16
         Top             =   420
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
            Picture         =   "FINVEN04.frx":2D68
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificado     por Proveedor"
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
            TabIndex        =   18
            Top             =   0
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   13
         Top             =   840
         Width           =   2325
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
            Left            =   -20
            Picture         =   "FINVEN04.frx":3072
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Material"
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
            TabIndex        =   15
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   10
         Top             =   1260
         Width           =   2325
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
            Left            =   -20
            Picture         =   "FINVEN04.frx":337C
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales Reservados     a Consignar"
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
            Left            =   840
            TabIndex        =   12
            Top             =   0
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   4620
      TabIndex        =   2
      Top             =   105
      Width           =   4005
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   2050
         TabIndex        =   42
         Top             =   2625
         Width           =   1750
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Importes"
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
            TabIndex        =   44
            Top             =   735
            Width           =   1065
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cantidades"
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
            TabIndex        =   43
            Top             =   470
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   210
         TabIndex        =   39
         Top             =   2625
         Width           =   1750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   41
            Top             =   470
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
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
            Left            =   315
            TabIndex        =   40
            Top             =   735
            Width           =   1170
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Consultas y Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   30
            Top             =   1260
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
               Picture         =   "FINVEN04.frx":34C6
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Inventario General Discriminado en Parciales por Depósito"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "General Analítico      por Depósito"
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
               Left            =   840
               TabIndex        =   32
               Top             =   0
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   33
            Top             =   840
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
               Picture         =   "FINVEN04.frx":37D0
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Existencias en un Unico Depósito"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Parcial por Depósito"
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
               TabIndex        =   35
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   36
            Top             =   420
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
               Picture         =   "FINVEN04.frx":3ADA
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Inventario General (Totales por Item en Todos los Depósitos)"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Inventario General"
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
               TabIndex        =   38
               Top             =   0
               Width           =   1485
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Generación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4215
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Complementos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   45
         Top             =   2625
         Width           =   3795
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1890
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   49
            Top             =   315
            Width           =   1695
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
               Picture         =   "FINVEN04.frx":3DE4
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Ajustes   de Stocks"
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
               TabIndex        =   50
               Top             =   105
               Width           =   960
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1545
            TabIndex        =   46
            Top             =   315
            Width           =   1600
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
               Picture         =   "FINVEN04.frx":3F2E
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Listado de Captura de Datos para Inventario"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Captura"
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
               TabIndex        =   48
               Top             =   105
               Width           =   960
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Inventario Continuo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Width           =   3795
         Begin VB.TextBox Text2 
            BackColor       =   &H00E0E0E0&
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
            Left            =   315
            TabIndex        =   8
            Text            =   " "
            Top             =   735
            Width           =   2955
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2205
            TabIndex        =   6
            Text            =   " "
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   3
            Top             =   1260
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
               Height          =   600
               Left            =   0
               Picture         =   "FINVEN04.frx":4238
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Re-Calcular Exis- tencias a Fecha"
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
               TabIndex        =   5
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inventario:"
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
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Width           =   1905
         End
      End
   End
End
Attribute VB_Name = "FiNVEN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub

Private Sub Command13_Click()
    Command13.Enabled = False
    ACONEC$ = "ARCHIG04/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command13.Enabled = True
End Sub

Private Sub Command16_Click()
    Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
End Sub

Private Sub Command17_Click()
    PREPA% = 0
    FORMAJU% = 0
    AJUSTK04.Show
End Sub

Private Sub Command2_Click()
    Command13.Enabled = False
    ACONEC$ = "ARCHIG04/LIMASTER"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command13.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTK04.Show 1
    End If
    Command25.Enabled = True
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINGEN1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINGEN2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV3            ' filtrar y listar
    End If
End Sub

Private Sub Command4_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINDEP1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINDEP2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV4            ' filtrar y listar
    End If
    '
End Sub

Private Sub Command5_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINANA1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINANA2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV5            ' filtrar y listar
    End If
    '
End Sub

Private Sub Command7_Click()
    GENINV04.Show 1
End Sub

Private Sub Command8_Click()
    Call ACCDIR("RECEPMA")
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "INVENTAR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    FECHAINV$ = TRIM$(CONTROL$("INVENTAR", "FECHAINV"))
    If FECHANUM(FECHAINV$) < 1 Then
        FECHAINV$ = ""
    End If
    Text1.TEXT = FECHAINV$
    Text2.TEXT = TRIM$(CONTROL$("INVENTAR", "FEGENINV"))
    Show
    Refresh
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Option5_LostFocus()
  '\\\OT-06-0172-WAR-10/05/06///
  If DERACCE%("INVEVAL", "Inventario Valorizado") < 2 Then
    Option6.Value = 1
    Option5.Value = 0
  End If
  '\\\OT-06-0172-WAR-FIN///
End Sub

Private Sub Text1_GotFocus()
    Command7.SetFocus
End Sub

Private Sub Text2_gotfocus()
    Command7.SetFocus
End Sub

Sub GENINV3()
    '
10  XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD3", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    cod1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    cod2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación por Tipo"
    End If
    '
    GCV% = CVI(Mid$(OBX$, 6, 2))
    GRUCOS$ = ECOARCH$("GRUCOVEN", MKI$(GCV%))
    If GRUCOS$ = "" Then
        GCV% = 0
        Call REPLA(OBX$, MKI$(0), 6, 2)
        GRUCOS$ = "Sin Clasificación Contable"
    End If
    '
    If TRIM$(cod1$) = "" Or TRIM$(cod2$) = "" Or cod2$ < cod1$ Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVEGEN", "INVEGENF")
    Call BLANARCH("INVEGENF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEGEN")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("INVEGEN", I&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= cod1$ Then
                 If COD$ <= cod2$ Then
                   If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                     'If GCV% = 0 Or GRUCOVE%(CI%) = GCV% Then
                       Call REGAPP("INVEGENF", X$)
                     'End If
                   End If
                 End If
              End If
           End If
        End If
    Next I&
    '
    Call HEADERS("INVEGENF", "")
    Call HEADERS("INVEGENF", "Material: " + TRIM$(TIMAT$) + " - " + TRIM$(GRUCOS$))
    '
    Call CIERRARCH("INVEGENF")
    Screen.MousePointer = 1
'    '
    If Option6.Value = True Then
         Call FINAL("*LIINGEN1")
       ElseIf Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = CONTROL$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         Call HEADERS("INVEGENF", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
         '\\\OT-07-0018-RG-FIN///
         Call FINAL("*LIINGEN2")
    End If
    '
    '
'    If FiNVEN04.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FiNVEN04.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
End Sub
'

Sub GENINV4()
    
    '
    XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCODEP", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    cod1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    cod2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    If TRIM$(cod1$) = "" Or TRIM$(cod2$) = "" Or cod2$ < cod1$ Then
        GoTo 10
    End If
    SUCUI% = Asc(Mid$(OBX$, 5, 1))
    If ECOARCH("ECODEP", Chr$(SUCUI%)) = "" Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVEDEP", "INVEDEQ")
    Call BLANARCH("INVEDEQ")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEDEP")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("INVEDEP", I&)
        If Asc(Mid$(X$, 3, 1)) = SUCUI% Then
           CI% = CVI(Left$(X$, 2))
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 COD$ = CODEXT$(CI%)
                 If COD$ >= cod1$ Then
                    If COD$ <= cod2$ Then
                       Call REGAPP("INVEDEQ", X$)
                    End If
                 End If
              End If
           End If
        End If
    Next I&
    '
    Call CIERRARCH("INVEDEQ")
    Screen.MousePointer = 1
    '
'    If FORMINVEN.Option6.Value = True Then
'         Call FINAL("*LIINDEP1")
'       ElseIf FORMINVEN.Option5.Value = True Then
'         Call FINAL("*LIINDEP2")
'    End If
    '
    '\\\OT-05-0819-WAR-18/11/05///
    Call HEADERS("INVEDEQ", "Depósito: " + ECOARCH("ECODEP", Chr$(SUCUI%)))
    If FiNVEN04.Option6.Value = True Then
         Call FINAL("*LIINDEP1")
       ElseIf FiNVEN04.Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = CONTROL$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         'Call HEADERS("INVEDEQ", "")
         Call HEADERS("INVEDEQ", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
         '\\\OT-07-0018-RG-FIN///
         '
         Call FINAL("*LIINDEP2")
    End If
    '\\\OT-05-0819-WAR-FIN///
    'CODIGO VIEJO
'    If FiNVEN04.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FiNVEN04.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
End Sub
'

Sub GENINV5()
    '
    XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCOD1", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    cod1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    cod2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If TRIM$(cod1$) = "" Or TRIM$(cod2$) = "" Or cod2$ < cod1$ Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVANAL", "INVANALF")
    Call BLANARCH("INVANALF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVANAL")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("INVANAL", I&)
        CI% = CVI(Left$(X$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= cod1$ Then
                 If COD$ <= cod2$ Then
                    If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                        Call REGAPP("INVANALF", X$)
                    End If
                 End If
              End If
           End If
        End If
    Next I&
    '
    Call CIERRARCH("INVANALF")
    Screen.MousePointer = 1
    '/////////******FORMINVEN no existe en este proyecto****/////
'     If FORMINVEN.Option6.Value = True Then
'         Call FINAL("*LIINANA1")
'       ElseIf FORMINVEN.Option5.Value = True Then
'         Call FINAL("*LIINANA2")
'    End If
    '
    If FiNVEN04.Option6.Value = True Then
         Call FINAL("*LIINANA1")
       ElseIf FiNVEN04.Option5.Value = True Then
         '\\\OT-07-0018-RG-23/01/07///
         MON$ = CONTROL$("INVENTAR", "MONEDINV")
         MONDESC$ = ECOARCH$("TAMONED", Chr(Val(MON$)))
         'Call HEADERS("INVANALF", "")
         Call HEADERS("INVANALF", "Valorización: " + TRIM$(MON$) + " - " + TRIM$(MONDESC$))
         '\\\OT-07-0018-RG-FIN///
         '
         Call FINAL("*LIINANA2")
    End If
    '
End Sub

