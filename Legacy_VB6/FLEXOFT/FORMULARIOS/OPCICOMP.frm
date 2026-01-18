VERSION 5.00
Begin VB.Form OPCICOMP 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Gestión de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   210
      TabIndex        =   8
      Top             =   210
      Width           =   4215
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Materiales Consignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   14
         Top             =   3000
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   15
            Top             =   315
            Width           =   2955
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
               Picture         =   "OPCICOMP.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Entrega en Consignación "
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
               Left            =   2340
               Picture         =   "OPCICOMP.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Marcar Consignaciones sin Devolución"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Consignación    de Materiales"
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
               TabIndex        =   16
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Ordenes de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   315
         TabIndex        =   9
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   915
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   1650
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
               Left            =   1710
               Picture         =   "OPCICOMP.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
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
               Left            =   0
               Picture         =   "OPCICOMP.frx":059E
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Anulación de Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación O/Compras"
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
               TabIndex        =   18
               Top             =   60
               Width           =   1065
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   715
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   48
            Top             =   1250
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
               Left            =   0
               Picture         =   "OPCICOMP.frx":08A8
               Style           =   1  'Graphical
               TabIndex        =   49
               ToolTipText     =   "Marcar Items como Cumplidos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Items Cumplidos"
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
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   515
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   19
            Top             =   850
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Picture         =   "OPCICOMP.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir"
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
               TabIndex        =   20
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   21
            Top             =   450
            Width           =   2325
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
               Picture         =   "OPCICOMP.frx":0EBC
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Modificar Orden de Compra no Anulada (Sólo Precios)"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificar O/C"
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
               TabIndex        =   22
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pedidos de Suministro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   840
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   35
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command6 
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
            Left            =   -20
            Picture         =   "OPCICOMP.frx":1006
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Generar un Pedido de Suministro"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Pedido de Suministro"
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
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   32
         Top             =   795
         Width           =   2325
         Begin VB.CommandButton Command17 
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
            Picture         =   "OPCICOMP.frx":1310
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Adjudicación de Pedido de Suministro a Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignacion de Proveedor"
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
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   29
         Top             =   1260
         Width           =   2325
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
            Left            =   0
            Picture         =   "OPCICOMP.frx":161A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Pedidos de Suministro Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "PS's Pendientes"
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
            Width           =   1485
         End
      End
   End
   Begin VB.CommandButton Command19 
      Caption         =   "N.Orig"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   8820
      Picture         =   "OPCICOMP.frx":1924
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Códigos Equivalentes de Proveedor"
      Top             =   2520
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8820
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3825
      Width           =   770
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command18 
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
      Height          =   720
      Left            =   8820
      Picture         =   "OPCICOMP.frx":1A6E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   1785
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Opciones de Control "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   4620
      TabIndex        =   10
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Ultimas Compras (Rango de Fechas)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   315
         TabIndex        =   38
         Top             =   420
         Width           =   3375
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   42
            Top             =   960
            Width           =   2220
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
               Left            =   -20
               Picture         =   "OPCICOMP.frx":1EB0
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Ultimas Compras por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Código de Material"
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
               TabIndex        =   44
               Top             =   105
               Width           =   1170
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   39
            Top             =   540
            Width           =   2220
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
               Left            =   -20
               Picture         =   "OPCICOMP.frx":1FFA
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Ultimas Compras por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Proveedor"
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
               TabIndex        =   41
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Complementos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   315
         TabIndex        =   13
         Top             =   2400
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   25
            Top             =   915
            Width           =   2220
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
               Picture         =   "OPCICOMP.frx":2304
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Análisis Estadístico de Consumo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Análisis Estadís- tico de Consumo"
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
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   45
            Top             =   510
            Width           =   2220
            Begin VB.CommandButton Command3 
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
               Picture         =   "OPCICOMP.frx":260E
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Control de Costos de Reposición"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Reposición"
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
               Width           =   1695
            End
         End
      End
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
      Height          =   720
      Left            =   8820
      Picture         =   "OPCICOMP.frx":2758
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
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
      Height          =   720
      Left            =   8820
      Picture         =   "OPCICOMP.frx":28A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8220
      Picture         =   "OPCICOMP.frx":2BAC
      Top             =   4050
      Width           =   2010
   End
End
Attribute VB_Name = "OPCICOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   Call DESANULOC
   Command10.Enabled = True
End Sub


Sub Command12_Click()
    '
    Command12.Enabled = False
    Call SELCOMRAFE
    Call CONECRUN("*SOCOMPR", "Listado de Compras por Proveedor")
99  Command12.Enabled = True
    '
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
    Call MODIOCOM
99  Command13.Enabled = True
    '
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   Call OCOMANU
   Command14.Enabled = True
End Sub

Sub Command15_Click()
   '\\\OT-06-00296-WAR-14/07/06///
   Command15.Enabled = False
   '
   If DERACCE%("MACOSIDE", "Marcar Consignaciones Sin Devolución") >= 2 Then
     '1.-Selección del Proveedor
15   Call SELECHO("ACREDOR")
     NP1% = XVALO(TRIM$(VALACT1$("ACREDOR")))
     If NP1% <= 0 Then GoTo 99
     '
     '2.- Filtrar Maconsig
     COSINDEV.LICOPEN.Clear
     RFF$ = RECFILT$("MACONSIG", 2, MKI$(NP1%))
     For J1& = 1 To Len(RFF$) Step 4
       J& = CVS(Mid$(RFF$, J1&, 4))
       XX$ = REGLEIDO$("MACONSIG", J&)
       NP2% = CVI(Mid$(XX$, 3, 2))
       If NP2% = NP1% Then
         CAENT1# = CVD(Mid$(XX$, 63, 8))
         CAREN1# = CVD(Mid$(XX$, 111, 8))
         If CAENT1# > CAREN1# Then
           CI1% = CVI(Mid$(XX$, 33, 2))
           CE1$ = CODEXT$(CI1%)
           DE1$ = DESCRIT0(CI1%)
           CAPEN1# = CAENT1# - CAREN1#
           CAPEN2$ = FORMATNUM$(CAPEN1, "F12.2")
           NUOC1$ = Mid$(XX$, 21, 12)
           '
           Y$ = Space(96)
           Call REPLA(Y$, CE1$, 1, 16)
           Call REPLA(Y$, DE1$, 17, 38)
           Call REPLA(Y$, NUOC1$, 57, 12)
           Call REPLA(Y$, CAPEN2$, 69, 12)
           Call REPLA(Y$, TRIM$(Str$(J&)), 90, 6)
           COSINDEV.LICOPEN.AddItem Y$
         End If
       End If
     Next J1&
     '
     If COSINDEV.LICOPEN.ListCount <= 0 Then
       Call MENSERR(24, "No Existen Consignaciones Pendientes\para el Proveedor Seleccionado.")
       GoTo 15
     End If
     '
     With COSINDEV
       HOII% = HOY(HO$)
       .Label15 = FECHATEX$(HOII%)
       .Label3 = Str$(NP1%)
       .Show 1
     End With
     GoTo 15
     '
   End If
   '
99 Command15.Enabled = True
  '\\\OT-06-00296-WAR-FIN///
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: K& = 0: OCOA& = 0
    '
    For I& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", I&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             K& = K& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, K&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next I&
    '
    Call SETULTREG("OCOMPRA", K&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, UNIMED$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(RASOCLI$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If Val(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Sub Command18_Click()
    '
    Command18.Enabled = False
    'SETUPCOM.Show 1
    Call CONECRUN("FLSETUP/SETUPCOM", "Configuración de Funcionamiento")
    '
    Command19.Enabled = False
    FORCOMPRA.Nuorig.Enabled = False
    FORCOMPRA.Nuorig1.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
      FORCOMPRA.Nuorig.Enabled = True
      FORCOMPRA.Nuorig1.Enabled = True
    End If
    '
    Command18.Enabled = True
    '
End Sub

Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORPR")
   Command19.Enabled = True
End Sub

Sub Command2_Click()
    Call ENTRECONSI(0)
End Sub


Sub Command20_Click()
    Command20.Enabled = False
    Call SELCOMRAFE
    Call CONECRUN("*SOCOMAR", "Listado de Compras por Articulo")
    Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

Sub Command22_Click()
    '
    Command22.Enabled = False
    '
3   Call SELECHO("OCOMENCA/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("OCOMENCA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
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
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
    '19/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '19/03/07 (FIN)
'5   For kki% = 0 To 5
'      EDITFORM.Picture1(kki%).Cls
'      EDITFORM.Picture1(kki%).DrawWidth = 3
'      EDITFORM.Picture1(kki%).DrawStyle = 0
'      EDITFORM.Picture1(kki%).FillStyle = 1
'    Next kki%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For kki% = 0 To 5
'      EDITFORM.Picture1(kki%).Cls
'      EDITFORM.Picture1(kki%).Top = 0
'      EDITFORM.Picture1(kki%).Refresh
'      EDITFORM.Picture1(kki%).Height = 7070
'      EDITFORM.Picture1(kki%).Width = 11700
'      EDITFORM.Picture1(kki%).DrawWidth = 3
'      EDITFORM.Picture1(kki%).DrawStyle = 0
'      EDITFORM.Picture1(kki%).FillStyle = 1
'    Next kki%
'    ALTUPAGINA = 7070
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
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
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For kki% = 0 To 5
'       EDITFORM.Picture1(kki%).Visible = False
'    Next kki%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'    EDITFORM.HScroll1.Min = 0
'    EDITFORM.HScroll1.Value = 0
'    EDITFORM.HScroll1.Max = 1
'    If ANCHPAGINA > EDITFORM.Frame1.Width Then
'      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command22.Enabled = True
    '
End Sub

Sub Command23_Click()
   Command23.Enabled = False
   Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command23.Enabled = True
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
        CONTROCOS.Show 1
      Else
20      Call OPNOIN("Control Costos de Reposición")
    End If
    Command3.Enabled = True
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   Call ITEMCUMPLI
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call PEDISUMI
    'Call OPNOIN("Pedido de Suministro")
    Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    OPX% = ALTERNA%("Control Costos Reposición\Archivos de Transferencia\Otros Accesos Directos")
    If OPX% = 1 Then
        If InStr(EPAC$, "DOSIVA") < 1 Then
           Call MENSERR(24, "No Disponible en la Presente Versión.")
           GoTo 29
        End If
        CONTROCOS.Show 1
      ElseIf OPX% = 2 Then
         Call ATRANSFE
      ElseIf OPX% = 3 Then
20       Call ACCDIR("OCOMPRA")
    End If
29  Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Sub PEDISUMI()             ' NUEVO PEDIDO DE SUMINISTRO
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
    X0$ = REGBLAN$("OCOMPRA")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 9, 2)
    Call REPLA(X0$, "S", 57, 1)
    '
101 OBJE$ = "PEDISUMI"
    '
    X1$ = OBJPLA$(OBJE$, X0$)
    If X1$ = "" Then GoTo 299
    '
    X0$ = X1$
    NRO& = CVS(Left$(X1$, 4))
    If NRO& < 1 Then GoTo 101
    '
    FEMI% = CVI(Mid$(X1$, 5, 2))
    If FEMI% < 1 Then
        Call MENSERR(24, "Falta Fecha de Emisión")
        GoTo 101
    End If
    '
    NPX$ = MKI$(0)
    NPRO% = CVI(NPX$)
    '
    Call BLANARCH("!CARDETOC")
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - Uso Interno"
    VTB% = VENTABU%("CARPESU1/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    CAICO% = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       CANTU = CVS(Mid$(X$, 47, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       CAICO% = J&
       '
       Call REPLA(X$, UNIMED$(CI0%), 3, 4)
       If CVI(Mid$(X$, 67, 2)) < HO% Then
          Call REPLA(X$, MKI$(HO%), 67, 2)
       End If
       Call REPLA(X$, "  ", 69, 2)
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!CARDETOC", X$, J&)
       '
    Next J&
    '
    If CAICO% < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
103 VTB% = VENTABU%("CARPESU2/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       '
       If FENTRE% < FEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FEMI%), 9, 2)
          GoTo 103
       End If
       '
    Next J&
    '
    Call CIERRARCH("!CARDETOC")
    '
    Call BLOQARCH("OCOMPRA")
    Call BLOQARCH("MACONSIG")
    '
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PREUNID# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = X1$
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       'Call REPLA(X2$, MKS$(0), 58, 4)
       'Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For I% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(I%)
             CANTI# = CDbl(CANTU * USOI(I%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, UNIMED$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
               End If
             End If
          Next I%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Pedido de Suministro"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CAPARDOC% = 6
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CACOLTAB1% = 10
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PREUNID# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
        '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    
'consignacion inmediata si/no

    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CoCoPen()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = Asc(Mid$(XXX$, 122, 1))
         If MONEMI% <= 1 Then MONEX$ = "    "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         PREUNID# = CVD(Mid$(XXX$, 79, 8)) * TICAMBIO
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
         Call REPLA(XXY$, MKD$(PREUNID#), 133, 8)
         Call REPLA(XXY$, MKD$(PREUNID# * CAPEN), 141, 8)
         JJ& = JJ& + 1
         Call GRAREG("OPENDETA", XXY$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   Screen.MousePointer = 1
   '
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   OPX% = COMALTER%("Elija su Opción de Control:\\Cumplimiento\Calidad\Cancelar")
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  GoTo 10
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("BOLRECEP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
            CAPEN(K&) = CAPEN(K&) - CAREC
            GoTo 19
          End If
        End If
      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
      Z0$ = REGBLAN$("COCOBREC")
      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
      FENTREPRO = CVI(Mid$(X$, 95, 2))
      '
      If Len(PUBOL$(K&)) < 4 Then
           DIATRA% = 0
           If FENTREPRO < HOII Then
             DIATRA% = DIENTRE%(FENTREPRO, HOII)
           End If
           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
           NCOCO& = NCOCO& + 1
           Call GRAREG("COCOBREC", Z0$, NCOCO&)
         Else
           For II& = 1 To Len(PUBOL$(K&)) Step 4
             Z$ = Z0$
             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
             FERECI = CVI(Mid$(Y$, 81, 2))
             DIATRA% = 0
             If FENTREPRO < FERECI Then
               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
             End If
             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
             NCOCO& = NCOCO& + 1
             Call GRAREG("COCOBREC", Z$, NCOCO&)
           Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("BOLRECEP")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= DES% Then
        If FF% <= HAS% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                CI% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(CI%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                JJ& = JJ& + 1
                Call GRAREG("COCAREC", Z$, JJ&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
    '
    Command19.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    If EXISTE%("ESTOCK02.EXE") > 0 Then
      Command23.Enabled = True
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      Command3.Enabled = True
    End If
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub MODIOCOM()         ' MODIFICAR ORDEN DE COMPRA
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORMOCOM"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    AVALIDA$ = "MASTER"
    '
10  YY$ = REGSEL$("OCOMENCA/Indice de Ordenes de Compra Pendientes para Modificacion")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    If NUOCO& < 1 Then GoTo 999
    '
    RFG$ = RECFILT$("OCOMENCA", 1, MKS$(NUOCO&))
    If Len(RFG$) < 4 Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    RENCA& = CVS(Left$(RFG$, 4))
    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(XX$, 4)) <> NUOCO& Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    If Asc(Mid$(XX$, 54, 1)) = 9 Then ' esta anulada
      Call MENSERR(24, "O/Compra Anulada - Imposible Modificar")
      GoTo 10
    End If
    '
    RFF$ = RECFILT$("TXOCOLIB", 1, MKS$(NUOCO&))
    If Len(RFF$) >= 4 Then
       Call MENSERR(24, "O/Compra de Texto Libre - Imposible Modificar.\Debe Anular y Emitir nueva O/Compra.")
       GoTo 10
    End If
    '
    RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NUOCO&))
    If Len(RFF$) >= 4 Then
      For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        If Asc(Mid$(X$, 124, 1)) < 9 Then GoTo 100 ' esta abierta y pendiente
      Next UK&
    End If
    Call MENSERR(24, "O/Compra Anulada - Imposible Modificar")
    GoTo 10
    '
100 ALIVX = ALIVA
    INETOX# = CVD(Mid$(XX$, 89, 8)) - CVD(Mid$(XX$, 113, 8))
    If INETOX# > 0.05 Then ALIVX = 100 * CVD(Mid$(XX$, 121, 8)) / INETOX#
    If ALIVX < 0 Then ALIVX = 0
    If ALIVX > 99 Then ALIVX = ALIVA
    NPRO% = CVI(Mid$(XX$, 51, 2))
    NC1% = (-1) * NPRO%
    PVCL% = PIVACLI%(NC1%)
    If PVCL% = 2 Or PVCL% = 4 Or PVCL% = 5 Then
      ALIVX = 0
    End If
    '
    PERCX = 0
    If INETOX# > 0.05 Then PERCX = 100 * CVD(Mid$(XX$, 129, 8)) / INETOX#
    If PERCX < 0 Then PERCX = 0
    If PERCX > 99 Then PERCX = 0
    '
    FECHEMI% = CVI(Mid$(XX$, 49, 2))
    FENTREGEN% = 32767
    NPRO% = CVI(Mid$(XX$, 51, 2))
    NRO& = NUOCO&
    NC1% = (-1) * NPRO%
    CPAG% = Asc(Mid$(XX$, 153, 1))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    ATEN$ = Mid$(XX$, 163, 30)
    '
    X0$ = String$(10, 0) + Space$(160)
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(FECHEMI%), 5, 2)
    Call REPLA(X0$, MKI$(FECHEMI%), 7, 2)
    Call REPLA(X0$, MKI$(NPRO%), 9, 2)
    Call REPLA(X0$, ATEN$, 131, 40)
    '
    DOMIPRO$ = TRIM$(CPOSCLI$((-1) * NPRO%))
    If DOMIPRO$ <> "" Then DOMIPRO$ = DOMIPRO$ + " - "
    Call REPLA(X0$, RASOCLI$((-1) * NPRO%), 11, 40)
    Call REPLA(X0$, DOMICLI$((-1) * NPRO%), 51, 40)
    Call REPLA(X0$, DOMIPRO$ + LOCACLI$((-1) * NPRO%), 91, 40)
    '
    MONEMI% = Asc(Mid$(XX$, 84, 1))
    If MONEMI% < 1 Or ECOARCH$("TAMONED", Chr$(MONEMI%)) = "" Then
      MONEMI% = 1
    End If
    If MONEMI% > 1 Then
      RETAMO& = REGBUS&("TAMONED", 1, Chr$(MONEMI%))
      XYX$ = REGLEIDO$("TAMONED", RETAMO&)
      TICAEMI = CVS(Mid$(XYX$, 29, 4))
    End If
    If TICAEMI < 0.1 Then TICAEMI = 1
    Call REPLA(X0$, Chr$(MONEMI%), 171, 1)
    Call REPLA(X0$, MKS$(TICAEMI), 173, 4)
    '
11  X1$ = OBJPLA$("OCOMODI-ENCA", X0$)
    If X1$ = "" Then GoTo 10
    '
    MONEMI% = Asc(Mid$(X1$, 171, 1))
    MONEDEM$ = "Pesos"
    If MONEMI% <> 1 Then
      If ECOARCH$("TAMONED", Chr$(MONEMI%)) = "" Then
        Call MENSERR(24, "Moneda no Válida")
        GoTo 11
      End If
      MONEDEM$ = TRIM$(ECOARCH$("TAMONED", Chr$(MONEMI%)))
    End If
    '
    TICAEMI = CVS(Mid$(X1$, 173, 4))
    If MONEMI% <= 1 Then TICAEMI = 1
    If TICAEMI < 0.2 Then
      Call MENSERR(24, "Ingrese Tipo de Cambio")
      GoTo 11
    End If
    '
    If MONEMI% > 1 Then
      XX$ = REGLEIDO$("OCOMENCA", RENCA&)
      Call REPLA(XX$, MKS$(TICAEMI), 85, 4)
      Call GRAREG("OCOMENCA", XX$, RENCA&)
    End If
    '
    Screen.MousePointer = 11
    ATEN$ = Mid$(X1$, 131, 40)
    REDETA$ = ""
    Call SETULTREG("!OCOMDETA", 0)
    FIN& = ULTREG&("OCOMDETA")
    '
    RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NUOCO&))
    If Len(RFF$) >= 4 Then
      For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        ZZ$ = REGLEIDO$("OCOMDETA", U&)
        If CVS(Left$(ZZ$, 4)) = NUOCO& Then
          If Asc(Mid$(ZZ$, 124, 1)) < 9 Then
            'If CVS(Mid$(ZZ$, 51, 4)) < 0.05 Then
              CANTXX = CVS(Mid$(ZZ$, 55, 4))
              COEFIX = CVS(Mid$(ZZ$, 63, 4))
              If COEFIX > 0.05 Then
                CANTYY = COEFIX * CANTXX
                Call REPLA(ZZ$, MKS$(CANTYY), 51, 4)
              End If
            'End If
            REDETA$ = REDETA$ + MKS$(U&)
            Call REGAPP("!OCOMDETA", ZZ$)
            If CVI(Mid$(ZZ$, 95, 2)) > 0 Then
              If CVI(Mid$(ZZ$, 95, 2)) < FENTREGEN% Then
                FENTREGEN% = CVI(Mid$(ZZ$, 95, 2))
              End If
            End If
            '
          End If
        End If
      Next UK&
    End If
    Call CIERRARCH("OCOMDETA")
    If FENTREGEN% = 32767 Or FENTREGEN% < FECHEMI% Then FENTREGEN% = FECHEMI%
    '
103 TITU$ = TRIM$(Str$(NRO&)) + " - " + MONEDEM$ + " - " + TRIM$(RASOCLI$(NC1%))
    VTB% = VENTABU%("CAROCO3/TITUPAN=Modificación de Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 10
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PREUNID# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 103
       End If
       If TRIM$(UNIMAS$) = "" Then
           Call REPLA(X$, UNIMED$(CI0%), 11, 4)
       End If
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNIMED$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If PORDE < 0 Or PORDE > 99.99 Then
          Call MENSERR(24, "Porcentaje de Descuento no Válido\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If FENTRE% < FECHEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FECHEMI%), 9, 2)
          GoTo 103
       End If
       '
       PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       CANSTXX = CVS(Mid$(X$, 51, 4))
       If CANSTXX < 0.05 Then
          Call MENSERR(24, "Falta Cantidad en " + UNIST$ + "\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If UNIST$ = UNICO$ Then
         If Abs(CANSTXX - CANTU) >= 0.05 Then
           Call REPLA(X$, MKS$(CANTU), 51, 4)
           Call GRAREG("!OCOMDETA", X$, J&)
           Call MENSERR(24, "Cantidades Incoherentes")
           GoTo 103
         End If
       End If
       '
       COEFI = CANSTXX / CANTU
       Call REPLA(X$, MKS$(COEFI), 63, 4)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
    Call CIERRARCH("!OCOMDETA")
    '
    TOTIVA# = (Int(ALIVX * TOTANE# + 0.5)) / 100
    TOTPER# = (Int(PERCX * TOTANE# + 0.5)) / 100
    TOTATO# = TOTANE# + TOTIVA# + TOTPER#
    '
' presentar registro de totales y condicion de pago
    '
    If MONEMI% > 1 Then
      MONECOI% = MONEMI%
      MONECOSI% = MONEMI%
      TICAMBX = TICAEMI
    End If
123 ENCAO$ = REGLEIDO$("OCOMENCA", RENCA&)
    '
    If Asc(Mid$(ENCAO$, 83, 1)) > ULTREG&("TAMONED") Then
      Call REPLA(ENCAO$, Chr$(0), 83, 1)
      Call REPLA(ENCAO$, MKS$(1), 85, 4)
    End If
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    Call REPLA(ENCAO$, MKI$(HO%), 61, 2)
    Call REPLA(ENCAO$, MKI$(Int(HP)), 63, 2)
    Call REPLA(ENCAO$, MKI$(USNBR%), 65, 2)
    '
    If Asc(Mid$(ENCAO$, 153, 1)) < 1 Then
       CPAG% = CPAGCLI%((-1) * NPRO%)
       Call REPLA(ENCAO$, Chr$(CPAG%), 153, 1)
    End If
    '
    Call REPLA(ENCAO$, ATEN$, 163, 30)
125 PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = (Int(ALIVX * TOTANEDES# + 0.5)) / 100
    TOTPER# = (Int(PERCX * TOTANEDES# + 0.5)) / 100
    TOTATO# = TOTANEDES# + TOTIVA# + TOTPER#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTPER#), 129, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    '
    If MONEMI% > 1 Then
      MONECOI% = MONEMI%
      MONECOSI% = MONEMI%
      TICAMBX = TICAEMI
    End If
126 OBX$ = OBJPLA$("OCOMPRA-APRO", ENCAO$)
    If OBX$ = "" Then GoTo 103
    ENCAO$ = OBX$
    If Asc(Mid$(ENCAO$, 83, 1)) < 1 Then
      Call REPLA(ENCAO$, Chr$(1), 83, 1)
    End If
    Call REPLA(ENCAO$, Chr$(MONEMI%), 84, 1)
    TOTANTE# = TOTATO#
    TOTANE# = CVD(Mid$(ENCAO, 89, 8))
    PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    If TOTANEDES# >= 0.05 Then
       ALIVX = 100 * CVD(Mid$(ENCAO$, 121, 8)) / TOTANEDES#
       If ALIVX < 0 Then ALIVX = 0
       If ALIVX > 99 Then ALIVX = ALIVA
       If PVCL% = 2 Or PVCL% = 4 Or PVCL% = 5 Then ALIVX = 0
       PERCX = 100 * CVD(Mid$(ENCAO$, 129, 8)) / TOTANEDES#
       If PERCX < 0 Or PERCX > 99 Then PERCX = 0
    End If
    '
    TOTIVA# = (Int(ALIVX * TOTANEDES# + 0.5)) / 100
    TOTPER# = (Int(PERCX * TOTANEDES# + 0.5)) / 100
    TOTATO# = TOTANEDES# + TOTIVA# + TOTPER#
    '
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTPER#), 129, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    If Abs(TOTANTE# - TOTATO#) > 0.005 Then
      GoTo 126
    End If
    TICAMBIO = CVS(Mid$(ENCAO$, 85, 4))
    If TICAMBIO < 1 Then
       Call REPLA(ENCAO$, MKS$(1), 85, 4)
       GoTo 126
    End If
    MONEI% = Asc(Mid$(ENCAO$, 83, 1))
    MONECOSI% = MONEI%
    MONECOS$ = Left$(ECOARCH$("TAMONED", Mid$(ENCAO$, 83, 1)), 2)
    '
    Screen.MousePointer = 11
    Call BLOQARCH("OCOMENCA")
    Call BLOQARCH("OCOMDETA")
    Call BLOQARCH("MACONSIG")
    '
    ' grabar encabezado
    X$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(X$, 4)) <> NRO& Then
      Call MENSERR(24, "Modificacion Cancelada.")
      Call CIERRARCH("*.*")
      GoTo 999
    End If
    Call GRAREG("OCOMENCA", ENCAO$, RENCA&)
    Call CIERRARCH("OCOMENCA")
    '
    UREDET& = ULTREG&("OCOMDETA")
    For K& = 1 To Len(REDETA$) Step 4
      U& = CVS(Mid$(REDETA$, K&, 4))
      If U& > 0 Then
        If U& <= UREDET& Then
          ZZ$ = REGLEIDO$("OCOMDETA", U&)
          If CVS(Left$(ZZ$, 4)) = NRO& Then
            Call REPLA(ZZ$, Chr$(9), 124, 1)
            Call GRAREG("OCOMDETA", ZZ$, U&)
          End If
        End If
      End If
    Next K&
    Call CIERRARCH("OCOMDETA")
    '
    For J& = 1 To ULTREG&("!OCOMDETA")
       ZZ$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(ZZ$, MKS$(RENCA&), 125, 4)
       Call REPLA(ZZ$, Chr$(MONECOSI%) + Chr$(MONEMI%), 121, 2)
       TICAMB1 = TICAMBIO: If MONEMI% > 1 Then TICAMB1 = 1
       VAUNI$ = FORMATNUM$(CVD(Mid$(ZZ$, 79, 8)) / TICAMB1, "F12.4")
       Call REPLA(VAUNI$, MONECOS$, 1, 2)
       Call REPLA(ZZ$, VAUNI$, 43, 12)
       Call REGAPP("OCOMDETA", ZZ$)
    Next J&
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
    '
    ACOSREP$ = UCase$(CONTROL$("OCOMPRA", "ACOSREP"))
    'If ACOSREP$ = "SIEMPRE" Then Option7.Value = True
    'If ACOSREP$ = "NUNCA" Then Option9.Value = True
    '
' actualizar costo standard
    HOII% = HOY(HOS$)
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
      If NUORSE& > 0 Then   ' ACTUALIZA COSTO SERVICIOS
        NPRO% = CVI(Mid$(ZZ$, 7, 2))
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        NOPERX% = CVI(Mid$(ZZ$, 103, 2))
        PREUNINE# = CVD(Mid$(ZZ$, 79, 8))
        Call ACTCOSTANSER(CI0%, NOPERX%, NPRO%, PREUNINE#, MONEMI%, FECHEMI%)
        GoTo 149
      End If
      If NUORSE& < 1 Then   ' EXCLUYE ORDEN DE SERVICIO
        CI0% = CVI(Mid$(ZZ$, 9, 2))
        PREUNINE# = CVD(Mid$(ZZ$, 79, 8))
        If MONEMI% > 1 Then PREUNINE# = PREUNINE# * TICAMBIO
        COEFUNI = CVS(Mid$(ZZ$, 63, 4))
        If COEFUNI < 0.0001 Then COEFUNI = 1
        PREUNI1# = PREUNINE# / COEFUNI
        REMAL& = CI0% + 1
        REMAS$ = REGLEIDO$("MASTER", REMAL&)
        PRHBL% = CVI(Mid$(REMAS$, 123, 2))
        If PRHBL% = NPRO% Or PRHBL% < 1 Or ACOSREP$ = "SIEMPRE" Then
            PREUNI2 = PREUNI1# / TICAMBIO
            If PREUNI2 >= 0.000005 Then
              If ACOSREP$ <> "NUNCA" Then
                Call REPLA(REMAS$, MKS$(PREUNI2), 111, 4)
                Call REPLA(REMAS$, MKI$(FECHEMI%), 115, 2)
                If PRHBL% < 1 Then Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
                Call REPLA(REMAS$, Chr$(MONECOSI%), 95, 1)
                Call GRAREG("MASTER", REMAS$, REMAL&)
              End If
            End If
          Else
            '
            TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(CI0%)) + "\el Proveedor Habitual Registrado es\"
            TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PRHBL%)) + "\  \"
            TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
            TTXX$ = TTXX$ + RASOCLI$((-1) * NPRO%)
            '
            If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 2 Then
              '
              Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
              PREUNI2 = PREUNI1# / TICAMBIO
              If PREUNI2 >= 0.000005 Then
                Call REPLA(REMAS$, MKS$(PREUNI2), 111, 4)
                Call REPLA(REMAS$, MKI$(FECHEMI%), 115, 2)
                Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
                Call REPLA(REMAS$, Chr$(MONECOSI%), 95, 1)
              End If
              Call GRAREG("MASTER", REMAS$, REMAL&)
              '
              ZZX$ = REGBLAN$("RECAMPH")
              Call REPLA(ZZX$, MKI$(CI0%), 1, 2)
              HOII% = HOY(HOS$)
              Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
              Call REPLA(ZZX$, MKI$(PRHBL%), 5, 2)
              Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
              Call REPLA(ZZX$, USERTER$, 9, 24)
              Call REPLA(ZZX$, "Modific.O/C " + TRIM$(Str$(NRO&)), 33, 32)
              Call REGAPP("RECAMPH", ZZX$)
              Call CIERRARCH("RECAMPH")
              GoTo 169
              '
            End If
            '
            MONEJ% = MONECOSTO%(CI0%)
            CUNI = COSUNI(CI0%) * TICAMONE(MONEJ%)
            If CUNI >= 0.00005 Then
              If PREUNI1# > CUNI + 0.00005 Then
                Call COMUNI("ATENCION: Precio de Compra\es Superior al Costo Standard\en Código '" + TRIM$(CODEXT$(CI0%)) + "' !!!\  \Informe al Supervisor de Compras.")
              End If
            End If
169     End If
        '
        If PESUNI(CI0%) < 0.000005 Then
          UNIST$ = UCase$(TRIM$(Mid$(ZZ$, 11, 4)))
          If Left$(UNIST$, 1) = "U" Then
            UNICO$ = UCase$(TRIM$(Mid$(ZZ$, 59, 4)))
            If UNICO$ = "KG" Then
              On Error Resume Next
              PSXX = 1 / COEFUNI
              On Error GoTo 0
              REMAL& = CI0% + 1
              REMAS$ = REGLEIDO$("MASTER", REMAL&)
              Call REPLA(REMAS$, MKS$(PSXX), 99, 4)
              Call GRAREG("MASTER", REMAS$, REMAL&)
            End If
          End If
        End If
      End If
      '
149 Next J&
    '
    JJ& = 0
    For U& = 193 To 448 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      JJ& = JJ& + 1
      Call GRAREG("!OBSERVOF", TTXX$, JJ&)
    Next U&
    Call SETULTREG("!OBSERVOF", JJ&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = Asc(Mid$(ENCAO$, 153, 1))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = Asc(Mid$(ENCAO$, 154, 1))
    EMBALA$ = TRIM$(ECOARCH$("TAEMBALA", Chr$(TIEM%)))
    LUENTREG$ = TRIM$(Mid$(ENCAO$, 449, 64))
    '
    PRIPLANOI% = 0
    If CONTROL$("OCOMPRA", "PRIPLANO") = "SI" Then PRIPLANOI% = 1
    PRIESPEC% = 0
    If CONTROL$("OCOMPRA", "PRIESPEC") = "SI" Then PRIESPEC% = 1
    PRIHRINS% = 0
    If CONTROL$("OCOMPRA", "PRIHRINS") = "SI" Then PRIHRINS% = 1
    '
    ' IMPRESION DE LA O/COMPRA
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FECHEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CODPARAM$(12) = "TRANSPOR"
      CODPARAM$(13) = "POR-DESC"
      CODPARAM$(14) = "IMP-DESC"
      CODPARAM$(15) = "TIPO-VAL"
      CODPARAM$(16) = "IMPO-CHE"
      CODPARAM$(17) = "REF-CONJ"
      CODPARAM$(18) = "PERC-IVA"
      CODPARAM$(19) = "DOMI-TRA"
      CODPARAM$(20) = "FECH-REC"
      '\\\OT-06-0483-RG-14/11/06///
      CODPARAM$(21) = "REF-MAT2"
      CAPARDOC% = 21
      '\\\OT-06-0483-RG-FIN///
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVX)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      DOCPARAM$(7) = Mid$(X1$, 147, 20)
      DOCPARAM$(8) = Mid$(X1$, 67, 40)
      DOCPARAM$(9) = Mid$(X1$, 167, 40)
      DOCPARAM$(10) = Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(TOTANEDES#)
      DOCPARAM$(12) = TRIM$(Mid$(ENCAO$, 163, 30)) ' ATENCION SR.
      DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
      DOCPARAM$(14) = Str$(IDESCUEN#)
      DOCPARAM$(15) = MONEDEM$
      DOCPARAM$(16) = ""
      If MONEMI% > 1 Then DOCPARAM$(16) = Str$(TICAMBIO)
      DOCPARAM$(17) = EMBALA$
      DOCPARAM$(18) = Str$(TOTPER#)
      DOCPARAM$(19) = LUENTREG$
        HOII% = HOY(HOS$)
        FEX = FENTREGEN%
      DOCPARAM$(20) = FECHATEX$(FEX)
        If FENTREGEN% = HOII% Then
            DOCPARAM$(20) = DOCPARAM$(20) + " (inmediato)"
          ElseIf FENTREGEN% < HOII% + 4 Then
            DOCPARAM$(20) = DOCPARAM$(20) + " (urgente!!)"
        End If
        '\\\OT-06-0483-RG-14/11/06///
        If EXISTE%(LUDAT$ + "CALINTEG.RFS") > 0 Then
          XX11$ = FILTERGETRECORD$("CALINTEG", 1, MKI$((-1) * NC1%))
          PUNTAPRO = CVS(Mid$(XX11$, 35, 4))
          DOCPARAM$(21) = FORMATNUM$(PUNTAPRO, "F4.1")
      End If
      '\\\OT-06-0483-RG-FIN///
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CACOLTAB1% = 13
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 0
      If CONTROL$("OCOMPRA", "INCLUSPE") = "SI" Then
         INCLUSPE% = 1
      End If
      '
      REPLAORI% = 0
      If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
         REPLAORI% = 1
      End If
      '
      RECOPROP% = 0
      If CONTROL$("OCOMPRA", "RECOPROP") = "SI" Then
         RECOPROP% = 1
      End If
      '
      REVICODX% = 0
      If CONTROL$("OCOMPRA", "REVICOD") = "SI" Then
         REVICODX% = 1
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
         CANTU = CVS(Mid$(X$, 55, 4))
         PREUNID# = CVD(Mid$(X$, 67, 8))
         PORDE = CVS(Mid$(X$, 75, 4))
         FEX = CVI(Mid$(X$, 95, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CODDX$ = TRIM$(CODEXT$(CI0%))
         CPROP$ = CODDX$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         '
         If REVICODX% = 1 Then
            FEREVX = FEREVI%(CI0%)
            RVCDDX$ = TRIM$(REVICOD$(CI0%))
            If RVCDDX$ <> "" Then
              CODDX$ = CODDX$ + " (Letra: " + RVCDDX$ + " - " + FECHATEX$(FEREVX) + ")"
            End If
         End If
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = TRIM$(DESCRIT0$(CI0%))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               If RECOPROP% > 0 Then
                 DECRIT$ = DECRIT$ + " (" + CPROP$ + ")"
               End If
             End If
           End If
         TETABU1$(2, CAITAB1%) = DECRIT$
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 97, 2)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
           KILOX$ = FORMATNUM$(CVS(Mid$(X$, 51, 4)) * PESUNI(CI0%), "F12.3")
         TETABU1$(11, CAITAB1%) = KILOX$
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         '
         If INCLUSPE% > 0 Then
           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI0%))
           If Len(RNC$) > 0 Then
             INTERLI = Val(ATRIFORM$(FORIPRE$, "INTERLIN"))
             '
             COLUTA% = 0
             URN& = ULTREG&("ESPECIFI")
             '
             For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                 If RENOTA& <= URN& Then
                   TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                   TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                   If INTERLI <= 6 Or InStr(UCase$(FORIPRE$), "MZZ") > 0 Then
                       CAITAB1% = CAITAB1% + 1
                       TETABU1$(2, CAITAB1%) = TTYY$
                     Else
                       COLUTA% = COLUTA% + 1
                       If COLUTA% > 2 Then COLUTA% = 1
                       If COLUTA% = 1 Then CAITAB1% = CAITAB1% + 1
                       TETABU1$(COLUTA%, CAITAB1%) = TTYY$
                   End If
                 End If
               End If
             Next U&
             '
           End If
         End If
         '
109   Next J&
      '
      '\\\OT-06-0005-WAR-10/01/06///
      'Call PRINTDOC("FORMOCOM")   ' O/C a impresora
      Screen.MousePointer = 1
110   SALIOCOM.Show 1
      If SALIOCOM.Check1.Value < 1 Then
        If SALIOCOM.Check2.Value < 1 Then
          Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
          GoTo 110
        End If
      End If
      '
      If SALIOCOM.Check1.Value = 1 Then
        Call PRINTDOC("FORMOCOM")   ' O/C a impresora
      End If
      If SALIOCOM.Check2.Value = 1 Then
15      SENDBYMAIL$ = ""
        PORMAIL% = 1
        NCXX% = (-1) * NPRO%
        EMAI$ = EMAILCLI$(NCXX%)
        If EMAI$ = "" Or InStr(EMAI$, "@") < 2 Then
          Call MENSERR(24, "Imposible Enviar por E-Mail.\   \Dirección Electrónica Inexistente\o no Válida.")
          Call PRINTDOC("FORMOCOM")   ' O/C a impresora
          GoTo 120
        End If
        SENDBYMAIL$ = EMAI$
        '
        For Each prx1 In Printers
           If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then GoTo 115
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        Call PRINTDOC("FORMOCOM")   ' O/C a impresora
        GoTo 120
        '
115     Call PRINTDOC("FORMOCOM")   ' O/C a impresora
        Call RECONECTA
        Exit Sub
      End If
      '
120    For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Left$(X$, 2))
         If CI0% > 0 Then
           If CI0% <= NUMAS% Then
             If PRIPLANOI% = 1 Then Call PRIPLANO(CI0%)
             If PRIESPEC% = 1 Then Call PRIHOJESP(CI0%)
             If PRIHRINS% = 1 Then Call PRIHOJRINSP(CI0%)
           End If
         End If
      Next J&
      '
    End If
    Call CIERRARCH("*.*")
    Call FRESHECHO("INDIOCOM")
    Call BAJALDISCO
    '
    GoTo 10
    '
999 Call SETULTREG("!OCOMDETA", 0)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'
Sub ACTCOSTANSER(CI0%, NOPERX%, NPRO%, PREUNINE#, MONEMI%, FECHEMI%)
   RFF$ = RECFILT$("COSOPTER", 1, MKI$(CI0%))
   For UK& = 1 To Len(RFF$) Step 4
     RECOSO& = CVS(Mid$(RFF$, UK&, 4))
     If RECOSO& > 0 Then
       If RECOSO& <= ULTREG&("COSOPTER") Then
         ZXC$ = REGLEIDO$("COSOPTER", RECOSO&)
         If CVI(Mid$(ZXC$, 3, 2)) = NOPERX% Then GoTo 20
       End If
     End If
   Next UK&
   '
   RECOSO& = 1 + ULTREG&("COSOPTER")
   ZXC$ = String$(32, 0)
   Call REPLA(ZXC$, MKI$(CI0%), 1, 2)
   Call REPLA(ZXC$, MKI$(NOPERX%), 3, 2)
   '
20 NPROX% = CVI(Mid$(ZXC$, 5, 2))
   If NPROX% = 0 Or NPROX% = NPRO% Then
     If PREUNINE# >= 0.00005 Then
       MONEI% = MONEMI%: If MONEI% < 1 Then MONEI% = 1
       Call REPLA(ZXC$, MKI$(NPRO%), 5, 2)
       Call REPLA(ZXC$, MKD$(PREUNINE#), 7, 8)
       Call REPLA(ZXC$, MKI$(MONEMI%), 15, 2)
       Call REPLA(ZXC$, MKI$(FECHEMI%), 17, 2)
       Call GRAREG("COSOPTER", ZXC$, RECOSO&)
       Call CIERRARCH("COSOPTER")
     End If
   End If
   '
End Sub

Sub OCOMANU()
    '
    If COMALTER%("Realmente Desea\ANULAR Ordenes de Compra ?\\No, Abandonar\Si, Anular O/C's") <> 2 Then
      Exit Sub
    End If
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE COMPRA"
    
2   Screen.MousePointer = 11
    OCOA& = 0: VOCO$ = ""
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      NOCO& = CVS(Left$(X$, 4))
      If NOCO& <> OCOA& Then
        If Asc(Mid$(X$, 124, 1)) < 8 Then ' esta abierta y pendiente
          If CVS(Mid$(X$, 99, 4)) < CVS(Mid$(X$, 55, 4)) - 0.05 Then
            VOCO$ = VOCO$ + MKS$(NOCO&)
            OCOA& = NOCO&
          End If
        End If
      End If
    Next U&
    
3   J& = 0: K& = 0: OCOA& = 0
    '
    FIN& = ULTREG&("OCOMENCA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) < 1 Then ' esta abierta y pendiente
        NOCO& = CVS(Left$(X$, 4))
        For KKL& = 1 To Len(VOCO$) Step 4
          If CVS(Mid$(VOCO$, KKL&, 4)) = NOCO& Then GoTo 5
        Next KKL&
        GoTo 9
        '
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("INDIOCOM", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("INDIOCOM", J&)
    Call CIERRARCH("INDIOCOM")
    Call FRESHECHO("INDIOCOM")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("INDIOCOM/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    RENCA& = CVS(Mid$(YY$, 49, 4))
    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(XX$, 4)) <> NUOCO& Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMENCA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMENCA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        RFXX$ = "(anul) " + TRIM$(Mid$(XX$, 5, 44))
        Call REPLA(XX$, Chr$(9), 54, 1)
        Call REPLA(XX$, RFXX$, 5, 44)
        Call GRAREG("OCOMENCA", XX$, J&)
      End If
    Next J&
    Call CIERRARCH("OCOMENCA")
    '
    ' este proceso es para eliminar los registros correspondientes a previamente anuladas
    FIN& = ULTREG&("OCOMDETA")
    INI& = 1
    RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NUOCO&))
    If Len(RFF$) >= 4 Then INI& = CVS(Left$(RFF$, 4))
    JJ& = INI& - 1
    For J& = INI& To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Or Asc(Mid$(XX$, 124, 1)) <> 9 Then
        JJ& = JJ& + 1
        If JJ& <> J& Then
          Call GRAREG("OCOMDETA", XX$, JJ&)
        End If
      End If
    Next J&
    Call SETULTREG("OCOMDETA", JJ&)
    Call CIERRARCH("OCOMDETA")
    '
    FIN& = ULTREG&("OCOMDETA")
    INI& = 1
    RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NUOCO&))
    If Len(RFF$) >= 4 Then INI& = CVS(Left$(RFF$, 4))
    For J& = INI& To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        If CVS(Mid$(XX$, 99, 4)) < CVS(Mid$(XX$, 55, 4)) - 0.05 Then
          If Asc(Mid$(XX$, 124, 1)) < 8 Then ' PARA QUE NO ANULE O/C's CUMPLIDAS
            Call REPLA(XX$, Chr$(9), 124, 1)
            Call GRAREG("OCOMDETA", XX$, J&)
          End If
        End If
      End If
    Next J&
    Call CIERRARCH("OCOMDETA")
    '
    J& = 0
    For I& = 1 To ULTREG&("INDIOCOM")
      XX$ = REGLEIDO$("INDIOCOM", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("INDIOCOM", XX$, J&)
      End If
    Next I&
    Call SETULTREG("INDIOCOM", J&)
    '
    '10/05/07 (OT-07-0187)
    ' CANCELAR RESERVAS
    For Z& = 1 To ULTREG&("MACONSIG")
      XX$ = REGLEIDO$("MACONSIG", Z&)
      If XVALO(Mid$(XX$, 24, 6)) = NUOCO& Then
         CASIG# = CVD(Mid$(XX$, 63, 8))
         Call REPLA(XX$, MKD$(CASIG#), 55, 8)
         Call GRAREG("MACONSIG", XX$, Z&)
      End If
    Next Z&
    '10/05/07 (FIN)
    '
    JJ& = 0
    For J& = 1 To ULTREG&("MACONSIG")
      XX$ = REGLEIDO$("MACONSIG", J&)
      If Val(Mid$(XX$, 24, 6)) <> NUOCO& Then
           JJ& = JJ& + 1
           If JJ& <> J& Then
             Call GRAREG("MACONSIG", XX$, JJ&)
           End If
         Else
           CAENTRE# = CVD(Mid$(XX$, 63, 8))
           CADEVUE# = CVD(Mid$(XX$, 111, 8))
           If Abs(CAENTRE# - CADEVUE#) >= 0.05 Then
             JJ& = JJ& + 1
             Call GRAREG("MACONSIG", XX$, JJ&)
           End If
      End If
    Next J&
    Call SETULTREG("MACONSIG", JJ&)
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FRESHECHO("INDIOCOM")
    Screen.MousePointer = 1
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'
Sub BUSCAULTIPRE(NPRO%, CIIX%, UNIMAS$, UPREDES$)
   '
   Static MODOULPRE%
   If MODOULPRE% = 0 Then
     MODOULPRE% = 1
     If CONTROL("OCOMPRA", "PRECIDEF") = "COSTOREP" Then MODOULPRE% = 2
   End If
   '
   If MODOULPRE% = 2 Then ' costo de reposición
     UPREDES$ = String$(20, 0)
     Call REPLA(UPREDES$, AJUSTI$(UNIMED$(CIIX%), 4), 1, 4)
     Call REPLA(UPREDES$, MKS$(1), 5, 4)
     Call REPLA(UPREDES$, MKD$(COSUNI(CIIX%) * TICAMONE(MONECOSTO%(CIIX%))), 9, 8)
     Call REPLA(UPREDES$, MKS$(0), 17, 4)
     Exit Sub
   End If
   '
   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           If TRIM$(Mid$(YX$, 11, 4)) = TRIM$(UNIMAS$) Then
             UPREDES$ = Mid$(YX$, 59, 20)
             PREUNANT# = CVD(Mid$(UPREDES$, 9, 8))
             MONEMIXA% = Asc(Mid$(YX$, 122, 1))
             If MONEMIXA% > 1 Then
               PREUNANT# = PREUNANT# * TICAMONE(MONEMIXA%)
               Call REPLA(UPREDES$, MKD$(PREUNANT#), 9, 8)
             End If
           End If
         End If
       End If
     Next U&
   End If
   '
End Sub

Sub CONVERCOM()
   '
   ' conversion de formato de ordenes de compra
   '
   If ULTREG&("OCOMPRA") > 0 Then
     '
     Call BLOQARCH("OCOMPRA")
     Call BLOQARCH("OCOMENCA")
     Call BLOQARCH("OCOMDETA")
     '
     Screen.MousePointer = 11
     OCOA& = 0: NPROA% = 0: ENCAO$ = ""
     FIN& = ULTREG&("OCOMPRA")
     '
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       NUOCO& = CVS(Left$(XX$, 4))
       FEMI% = CVI(Mid$(XX$, 5, 2))
       NPRO% = CVI(Mid$(XX$, 7, 2))
       FENTRE% = CVI(Mid$(XX$, 9, 2))
       CI0% = CVI(Mid$(XX$, 11, 2))
       CAOR = CVS(Mid$(XX$, 13, 4))
       OBSER$ = Mid$(XX$, 17, 30)
       PREUNID# = CVS(Mid$(XX$, 49, 4))
       DESCUNI = CVS(Mid$(XX$, 53, 4))
       TISU% = Asc(Mid$(XX$, 57, 1))
       CARE = CVS(Mid$(XX$, 58, 4))
       STAT% = Asc(Mid$(XX$, 62, 1))
       '
       If NUOCO& <> OCOA& Or NPRO% <> NPROA% Then
         If OCOA& <> 0 Then
           TOTIVA# = 0
           PVCL% = PIVACLI%((-1) * NPRO%)
           If PVCL% <> 2 Then
             If PVCL% <> 4 Then
               TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
             End If
           End If
           Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
           Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
           Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 145, 8)
           Call REGAPP("OCOMENCA", ENCAO$)
         End If
         '
         OCOA& = NUOCO&: NPROA% = NPRO%
         ENCAO$ = REGBLAN$("OCOMENCA")
         '
         Call REPLA(ENCAO$, MKS$(NUOCO&), 1, 4)
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
         Call REPLA(ENCAO$, REFE$, 5, 44)
         Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
         Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
         Call REPLA(ENCAO$, Chr$(TISU%), 53, 1)
         Call REPLA(ENCAO$, Chr$(STAT%), 54, 1)
         '
         HO% = HOY(HOS$)
         HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
         Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
         Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
         Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
         '
         If STAT% = 9 Then
           Call REPLA(ENCAO$, MKI$(HO%), 67, 2)
           Call REPLA(ENCAO$, MKI$(Int(HP)), 69, 2)
           Call REPLA(ENCAO$, MKI$(USNBR%), 71, 2)
           Call REPLA(ENCAO$, "Conversion BD", 73, 16)
         End If
         '
         CPAG% = CPAGCLI%((-1) * NPRO%)
         Call REPLA(ENCAO$, MKI$(CPAG%), 153, 2)
         REDETA& = 1 + ULTREG&("OCOMDETA")
         Call REPLA(ENCAO$, MKS$(REDETA&), 155, 4)
         Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
         '
         Call REPLA(ENCAO$, OBSER$, 193, 64)
         '
       End If
       '
       REDET$ = REGBLAN$("OCOMDETA")
       Call REPLA(REDET$, MKS$(NUOCO&), 1, 4)
       Call REPLA(REDET$, MKI$(FEMI%), 5, 2)
       Call REPLA(REDET$, MKI$(NPRO%), 7, 2)
       Call REPLA(REDET$, MKI$(CI0%), 9, 2)
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(REDET$, UNIMAS$, 11, 4)
       Call REPLA(REDET$, MKS$(CAOR), 55, 4)
       Call REPLA(REDET$, UNIMED$(CI0%), 59, 4)
       Call REPLA(REDET$, MKS$(1), 63, 4)
       Call REPLA(REDET$, MKD$(PREUNID#), 67, 8)
       Call REPLA(REDET$, MKD$(PREUNID#), 79, 8)
       Call REPLA(REDET$, MKD$(PREUNID# * CAOR), 87, 8)
       Call REPLA(REDET$, MKI$(FENTRE%), 95, 2)
       If TISU% > 1 Then
         Call REPLA(REDET$, " *", 97, 2)
       End If
       Call REPLA(REDET$, MKS$(CARE), 99, 4)
       Call REPLA(REDET$, Chr$(STAT%), 124, 1)
       RENCA& = 1 + ULTREG&("OCOMENCA")
       Call REPLA(REDET$, MKS$(RENCA&), 125, 4)
       Call REGAPP("OCOMDETA", REDET$)
       '
       TOTANE# = TOTANE# + PREUNID# * CAOR
       REDETA& = ULTREG&("OCOMDETA")
       Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
       '
       Call GRAREG("OCOMPRA", String$(128, 0), U&)
       '
     Next U&
     '
     If OCOA& <> 0 Then
       TOTIVA# = 0
       PVCL% = PIVACLI%((-1) * NPRO%)
       If PVCL% <> 2 Then
         If PVCL% <> 4 Then
           TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
         End If
       End If
       Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
       Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
       Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 137, 8)
       Call REGAPP("OCOMENCA", ENCAO$)
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
   End If
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Sub SELCOMRAFE()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   REBLA$ = REGBLAN$("COPRECOM")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     FECHEMI% = CVI(Mid$(XXX$, 5, 2))
     '
     STAT% = Asc(Mid$(XXX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       If FECHEMI% >= DES% Then
         If FECHEMI% <= HAS% Then
           '
           NPRO% = CVI(Mid$(XXX$, 7, 2))
           CIXI% = CVI(Mid$(XXX$, 9, 2))
           DOPE$ = ""
           NOPE% = CVI(Mid$(XXX$, 103, 2))
           If NOPE% > 0 Then
             DOPE$ = DESCRIOP$(CIXI%, NOPE%)
           End If
           NUOC& = CVS(Left$(XXX$, 4))
           UMEDI$ = Mid$(XXX$, 59, 4)
           If TRIM$(UMEDI$) = "" Then
             UMEDI$ = UNIMED$(CIXI%)
           End If
           CACOM = CVS(Mid$(XXX$, 55, 4))
           '
           If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
             CACOM = CVS(Mid$(XXX$, 99, 4))  ' LA CANT COMPRADA = A LA RECIBIDA
           End If
           '
           PREUNI# = CVD(Mid$(XXX$, 79, 8))
           '
           MONEMI% = Asc(Mid$(XXX$, 122, 1))
           If MONEMI% <= 1 Then MONEX$ = "    "
           If MONEMI% = 2 Then MONEX$ = "U$S "
           If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
           TICAMBIO = 1
           If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
           '
           YYY$ = REBLA$
           Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
           Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
           Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
           Call REPLA(YYY$, DOPE$, 11, 22)
           Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
           Call REPLA(YYY$, UMEDI$, 37, 4)
           Call REPLA(YYY$, MKS$(CACOM), 41, 4)
           Call REPLA(YYY$, MONEX$, 49, 4)
           Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
           ITOCOM# = PREUNI# * CACOM
           Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
           Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
           Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
           JJ& = JJ& + 1
           Call GRAREG("COPRECOM", YYY$, JJ&)
         End If
       End If
     End If
     '
   Next U&
   '
   Call SETULTREG("COPRECOM", JJ&)
   Call CIERRARCH("COPRECOM")
   Screen.MousePointer = 1
   '
End Sub
'
Sub DESANULOC()
   '
3  Call COPYSTRU("INDIOCOM", "IOCANUL")
   J& = 0: K& = 0: OCOA& = 0
   '
   FIN& = ULTREG&("OCOMENCA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) = 9 Then ' esta anulada
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("IOCANUL", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("IOCANUL", J&)
    Call CIERRARCH("IOCANUL")
    Call FRESHECHO("IOCANUL")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("IOCANUL/Indice de Ordenes de Compra Anuladas para Restauración")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    RENCA& = CVS(Mid$(YY$, 49, 4))
    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
    If CVS(Left$(XX$, 4)) <> NUOCO& Then
      Call MENSERR(24, "Error de Indices - Consulte")
      GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMENCA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMENCA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
          FECHEMII% = CVI(Mid$(XX$, 49, 2))
          NPROI% = CVI(Mid$(XX$, 51, 2))
          FEX = FECHEMII%
          REFEI$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPROI%)
          Call REPLA(XX$, Chr$(0), 54, 1)
          Call REPLA(XX$, REFEI$, 5, 44)
          Call GRAREG("OCOMENCA", XX$, J&)
          YADESA% = 1
        Else
          If YADESA% > 0 Then GoTo 19
      End If
    Next J&
19  Call CIERRARCH("OCOMENCA")
    '
    FIN& = ULTREG&("OCOMDETA")
    YADESA% = 0
    For J& = FIN& To 1 Step -1
      Call TRACE(20, FIN& - J& + 1, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        If Asc(Mid$(XX$, 124, 1)) = 9 Then
          CAPEDI1 = CVS(Mid$(XX$, 55, 4))
          CARECI1 = CVS(Mid$(XX$, 99, 4))
          STATUIX% = 0
          If CARECI1 > 0 Then STATUIX% = 2
          If CARECI1 >= CAPEDI1 - 0.05 Then STATUIX% = 3
          Call REPLA(XX$, Chr$(STATUIX%), 124, 1)
          Call GRAREG("OCOMDETA", XX$, J&)
        End If
      End If
    Next J&
    Call CIERRARCH("OCOMDETA")
    '
    J& = 0
    For I& = 1 To ULTREG&("IOCANUL")
      XX$ = REGLEIDO$("IOCANUL", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("IOCANUL", XX$, J&)
      End If
    Next I&
    Call SETULTREG("IOCANUL", J&)
    Call FRESHECHO("IOCANUL")
    '
    GoTo 10
    '
999 Call CIERRARCH("OCOMENCA")
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
End Sub

Sub ITEMCUMPLI()
    '
    TITU$ = "DAR POR CUMPLIDOS ITEMS DE COMPRAS"
    '
2   Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCUM")
    OCOA& = 0: VOCO$ = ""
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      NOCO& = CVS(Left$(X$, 4))
      If NOCO& <> OCOA& Then
        If Asc(Mid$(X$, 124, 1)) < 8 Then ' esta abierta y pendiente
          If CVS(Mid$(X$, 99, 4)) < CVS(Mid$(X$, 55, 4)) - 0.05 Then
            VOCO$ = VOCO$ + MKS$(NOCO&)
            OCOA& = NOCO&
          End If
        End If
      End If
    Next U&
    
3   J& = 0: K& = 0: OCOA& = 0
    '
    FIN& = ULTREG&("OCOMENCA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) < 9 Then ' esta abierta y pendiente
        NOCO& = CVS(Left$(X$, 4))
        For KKL& = 1 To Len(VOCO$) Step 4
          If CVS(Mid$(VOCO$, KKL&, 4)) = NOCO& Then GoTo 5
        Next KKL&
        GoTo 9
        '
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("INDIOCUM", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("INDIOCUM", J&)
    Call CIERRARCH("INDIOCUM")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("INDIOCUM/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    DAPORCUM.Label14 = NUOCO&
    '
    DAPORCUM.Show 1
    '
    ' A PARTIR DE AQUI MARCAR REGISTROS QUE SE DAN POR CUMPLIDOS
    If Val(DAPORCUM.Label14) = NUOCO& Then
      Screen.MousePointer = 11
      URE& = ULTREG&("OCOMDETA")
      For U& = 1 To DAPORCUM.LICOPEN.ListCount
        If DAPORCUM.LICOPEN.Selected(U& - 1) = True Then
          TTXX$ = DAPORCUM.LICOPEN.List(U& - 1)
          REOC& = Val(TRIM$(Mid$(TTXX$, 121, 8)))
          If REOC& > 0 Then
            If REOC& <= URE& Then
              X$ = REGLEIDO$("OCOMDETA", REOC&)
              STAT% = Asc(Mid$(X$, 124, 1))
              If STAT% < 9 Then   ' NO ESTA ANULADA
                COEFCAN = CVS(Mid$(X$, 63, 4))
                  If COEFCAN < 0.005 Then COEFCAN = 1
                CAORI = CVS(Mid$(X$, 55, 4)) * COEFCAN
                CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
                CAPEN = CAORI - CAREC
                If CAPEN >= 0.005 Then
                  Call REPLA(X$, Chr$(8), 124, 1)
                  Call GRAREG("OCOMDETA", X$, REOC&)
                End If
              End If
            End If
          End If
        End If
      Next U&
    End If
    Unload DAPORCUM
    Screen.MousePointer = 1
    '
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub
