VERSION 5.00
Begin VB.Form OPCICOM04 
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
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00006060&
      Height          =   4740
      Left            =   8880
      ScaleHeight     =   4680
      ScaleWidth      =   1005
      TabIndex        =   45
      Top             =   0
      Width           =   1065
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   180
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   555
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
         Picture         =   "OPCICOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   1005
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
         Picture         =   "OPCICOM04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "OPCICOM04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1695
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   125
         ScaleHeight     =   75
         ScaleWidth      =   495
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3675
         Width           =   555
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   48
            Top             =   105
            Width           =   12000
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
         Height          =   690
         Left            =   105
         Picture         =   "OPCICOM04.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Códigos Equivalentes de Proveedor"
         Top             =   2385
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "OPCICOM04.frx":09E0
         Stretch         =   -1  'True
         Top             =   4095
         Width           =   1410
      End
   End
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
      TabIndex        =   5
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
         TabIndex        =   9
         Top             =   3000
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   10
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
               Picture         =   "OPCICOM04.frx":2902
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
               Left            =   2350
               Picture         =   "OPCICOM04.frx":2C0C
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Consumo Materiales Consignados"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Consignación      de Materiales"
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
               TabIndex        =   11
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
         TabIndex        =   6
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   915
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   12
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
               Left            =   1720
               Picture         =   "OPCICOM04.frx":2D56
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
               Top             =   0
               Visible         =   0   'False
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
               Picture         =   "OPCICOM04.frx":2EA0
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Anulación de Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación de O/Compras"
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
               Left            =   600
               TabIndex        =   13
               Top             =   105
               Width           =   1155
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   715
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   42
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
               Picture         =   "OPCICOM04.frx":31AA
               Style           =   1  'Graphical
               TabIndex        =   43
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
               TabIndex        =   44
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
            TabIndex        =   14
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
               Picture         =   "OPCICOM04.frx":34B4
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
               TabIndex        =   15
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
            TabIndex        =   16
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
               Picture         =   "OPCICOM04.frx":37BE
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
               TabIndex        =   17
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
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   29
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
            Picture         =   "OPCICOM04.frx":3908
            Style           =   1  'Graphical
            TabIndex        =   30
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
            TabIndex        =   31
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
         TabIndex        =   26
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
            Picture         =   "OPCICOM04.frx":3C12
            Style           =   1  'Graphical
            TabIndex        =   27
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
            TabIndex        =   28
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
         TabIndex        =   23
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
            Picture         =   "OPCICOM04.frx":3F1C
            Style           =   1  'Graphical
            TabIndex        =   24
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
            TabIndex        =   25
            Top             =   105
            Width           =   1485
         End
      End
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
      TabIndex        =   7
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Resúmenes de Compras"
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
         TabIndex        =   32
         Top             =   420
         Width           =   3375
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   36
            Top             =   1020
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
               Picture         =   "OPCICOM04.frx":4226
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Estadísticas de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísiticos"
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
               TabIndex        =   38
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
            TabIndex        =   33
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
               Picture         =   "OPCICOM04.frx":4530
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Resúmenes de Ultimas Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de Ul- timas Compras"
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
               Top             =   45
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
         TabIndex        =   8
         Top             =   2400
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   19
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
               Picture         =   "OPCICOM04.frx":483A
               Style           =   1  'Graphical
               TabIndex        =   20
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
               TabIndex        =   21
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
            TabIndex        =   39
            Top             =   510
            Width           =   2220
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
               Picture         =   "OPCICOM04.frx":4B44
               Style           =   1  'Graphical
               TabIndex        =   40
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
               TabIndex        =   41
               Top             =   105
               Width           =   1695
            End
         End
      End
   End
End
Attribute VB_Name = "OPCICOM04"
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
   If EXISTE%("ORTRAB04.EXE") > 0 Then
       Call MENSERR(24, "No Compatible con Módulo\de Ordenes de Trabajo.")
     Else
        Call DESANULOC
   End If
   Command10.Enabled = True
End Sub


Sub Command12_Click()
    '
    Command12.Enabled = False
    OPX% = COMALTER%("Detalle de Ultimas Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command12.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call OPLISCO04.Command12_Click
    Else
      Call OPLISCO04.Command20_Click
    End If
    Command12.Enabled = True
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
   Command15.Enabled = False
   Call DEVOLCONSI
   Command15.Enabled = True
   '
   'If EXISTE%("ORTRAB04.EXE") > 0 Then
   '    Call MENSERR(24, "No Compatible con Módulo de Ordenes de Trabajo.")
   '  Else
   '    Call OPNOIN("")
   'End If
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
       If XVALO(Mid$(XX$, 24, 6)) = NUOCO& Then
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
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      SETUPCOM.Show 1
      'Call CONECRUN("FSETUP04/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    '
    Command19.Enabled = False
    FORCOMP04.Nuorig1.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
      FORCOMP04.Nuorig1.Enabled = True
    End If
    '
    Command18.Enabled = True
    '
End Sub

Sub Command19_Click()
   Command19.Enabled = False
   OPX% = COMALTER%("Opciones de Salida\\Actualización de Datos\Listado Impreso")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   If OPX% = 2 Then GoTo 30
   
   Call FINAL("*CODORPR")
   GoTo 99

30 Call FINAL("*LICORPR")
   
99 Command19.Enabled = True
End Sub

Sub Command2_Click()
    Call NUENTRECONSI(0)
End Sub


Sub Command20_Click()
    '
    Command20.Enabled = False
    OPX% = COMALTER%("Ranking de Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command20.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call RANKPRO
    Else
      Call RANKART
    End If
    Command20.Enabled = True
    '
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
    Call OPLISCO04.Command22_Click
    Command22.Enabled = True
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
        OPX% = COMALTER%("Opciones de Ordenamiento:\\Por Código\Por Proveedor")
        If OPX% = 1 Then
            Call OPLIMAS04.Command4_Click
          ElseIf OPX% = 2 Then
            Call OPLIMAS04.Command5_Click
        End If
    End If
    Command3.Enabled = True
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   If EXISTE%("ORTRAB04.EXE") > 0 Then
       Call MENSERR(24, "No Compatible con Módulo\de Ordenes de Trabajo.")
     Else
10      OPX% = COMALTER%("Opciones de Trabajo\\Marcar Items Cumplidos\Restaurar")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        If OPX% = 1 Then
          Call ITEMCUMPLI
          GoTo 10
        ElseIf OPX% = 2 Then
          Call REVERTIR_ITEMCUMPLI
          GoTo 10
        End If
   End If
99 Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call PEDISUMI
    Command6.Enabled = True
End Sub

Sub Command8_Click()
    Command8.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") < 1 Then GoTo 20
    OPX% = ALTERNA%("Control Costos Reposición\Otros Accesos Directos")
    If OPX% = 1 Then
       CONTROCOS.Show 1
       GoTo 99
      ElseIf OPX% = 2 Then
       Call ACCDIR("OCOMPRA")
       GoTo 99
    End If
    '
20  OPX% = ALTERNA%("Depuración de Consig. Pendientes\Otros Accesos Directos")
    If OPX% = 1 Then
       Call DEPUCONSIG
      ElseIf OPX% = 2 Then
       Call ACCDIR("OCOMPRA")
    End If
99
    Command8.Enabled = True
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
    Call ABRESTOCKS
    ' PROCESO DE INICIALIZAR VALORES NULL
    SQLX$ = " UPDATE MACONSIG set Cant_ENTREG = 0  "
    SQLX$ = SQLX$ + " WHERE CANT_ENTREG IS NULL "
    Stocks.Execute (SQLX$)
    '
    SQLX$ = " UPDATE MACONSIG set Cant_RENDID = 0  "
    SQLX$ = SQLX$ + " WHERE Cant_RENDID IS NULL "
    Stocks.Execute (SQLX$)
    '
    Screen.MousePointer = 1
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
    'Command22.Enabled = False
    'Cargo la lista con las O/C
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA "
    SQLX$ = SQLX$ + "WHERE Stat_Ocom < 9 "   ' EXCLUYE ANULADAS
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set VRIMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 999
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INDIOCOM", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    Call FRESHECHO("!INDIOCOM")
    Screen.MousePointer = 1
    '
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas)")
    NUOCO& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NUOCO& < 1 Then GoTo 999
    '
    Call APERBASDAT("COMPRAS")
    With OcomEnca
      .FindFirst "Nume_Ocom = " & NUOCO& & "AND Stat_Ocom < 9 "
      If .NoMatch = True Then
        Call MENSERR(24, "Imposible Abrir la O/Compra " & NUOCO& & " - Consulte !")
        GoTo 3
      End If
      '
      If !Stat_Ocom > 8 Then ' esta anulada
        Call MENSERR(24, "O/Compra Anulada - Imposible Modificar")
        GoTo 3
      End If
      '
      NPRO% = !NPRO_OCOM
      NC1% = (-1) * NPRO%
      PVCL% = PIVACLI%(NC1%)
      If PVCL% = 2 Or PVCL% = 4 Or PVCL% = 5 Then
        ALIVX = 0
      End If
      '
      FECHEMI% = CVINT%(!Femi_Ocom)
      FENTREGEN% = 32767
      NRO& = NUOCO&
      CPAG% = !Cpag_Ocom
      EMBAL% = 0
      On Error Resume Next
      EMBAL% = !Embala_Ocom
      On Error GoTo 0
      FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
      If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
      ATEN$ = !AtSr_Ocom: If TRIM$(ATEN$) = "" Then ATEN$ = CONTACTCLI$(NC1%)
      '
      MONEMI% = !Mone_Emis
      If MONEMI% < 1 Or ECOARCH$("TAMONED", Chr$(MONEMI%)) = "" Then MONEMI% = 1
      MONECO% = !Mone_Cos
      If MONECO% < 1 Or ECOARCH$("TAMONED", Chr$(MONECO%)) = "" Then MONECO% = 1
      '
      FOMODOC04.Label11 = DEMONECO$(MONEMI%)
      FOMODOC04.LText1 = NPRO%
      '
      '\\\OT-08-0294-FL-26/05/08///
      If FOMODOC04.LText1.Caption = "" Then GoTo 999
      '\\\OT-08-0294-FL-26/05/08///
      '
      FOMODOC04.Text6 = ATEN$
      FOMODOC04.TEXT21(0) = !Cpag_Ocom
      FOMODOC04.TEXT21(1) = EMBAL%
      FOMODOC04.Text10 = !Lentrega
      FOMODOC04.Text13 = Obse_Ocom
      FOMODOC04.Text7 = ""    ' destino
      FOMODOC04.Text18 = ""   ' sector
      FOMODOC04.Text14 = ""   ' fecha de entrega
      FOMODOC04.Text15 = !Mone_Cos
      FOMODOC04.Text11 = NUOCO&
      FOMODOC04.Text9 = FECHATEX$(FECHEMI%)
      FOMODOC04.PORDESCU = !Pdes_Ocom
      FOMODOC04.Total(4) = FORMATNUM(!Iotr1_Ocom, "F10.2")
        PORPER = 0
        On Error Resume Next
        PORPER = 100 * XVALO(FOMODOC04.Total(4)) / (!ToBru_Ocom - !IDes_Ocom)
        On Error GoTo 0
        If PORPER > 99.9 Then PORPER = 99.9
      FOMODOC04.Text17 = TRIM$(FORMATNUM(PORPER, "F5.1"))
      FOMODOC04.Text13 = !Obse_Ocom
      On Error Resume Next
      FOMODOC04.Label13 = !Anex_Ocom
      FOMODOC04.Label15 = TRIM$(Str$(!Stat_Ocom))
      '\\\OT-06-0082-WAR-06/03/06///
      NROREV% = 0
      On Error Resume Next
      NROREV% = XVALO(!OCOMREVIS)
        If !Stat_Ocom >= 0 Then 'si no esta aprobada no cambia revisión(0).
          NROREV% = NROREV% + 1
        End If
      On Error GoTo 0
      FOMODOC04.Label16 = TRIM$(Str$(NROREV%)) 'FORMATNUM$(NROREV1%, "F2.0") LO COMENTE XQ DEVUELVE "" SI NUMERO =0
      '\\\OT-06-0082-WAR-FIN///
      On Error GoTo 0
      '
    End With
    '
    Screen.MousePointer = 11
    REDETA$ = ""
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG("!OCOMDETA", 0)
    Call SETULTREG(ARCHILIB$, 0)
    '
    REBLA$ = REGBLAN$("OCOMDETA")
    Call REPLA(REBLA$, MKS$(NUOCO&), 1, 4)
    Call REPLA(REBLA$, MKI$(FECHEMI%), 5, 2)
    Call REPLA(REBLA$, MKI$(NPRO%), 7, 2)
    Call REPLA(REBLA$, Chr$(MONECO%), 121, 1)
    Call REPLA(REBLA$, Chr$(MONEMI%), 122, 1)
      TICAMBIO = 1
      If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
    '
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    SQLX$ = SQLX$ + "ORDER BY NuOrd_Item ASC"
    Set COMDET = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COMDET
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "Imposible cargar Detalle de O/Compra")
        On Error GoTo 0
        GoTo 3
      End If
      On Error GoTo 0
      Do While Not .EOF
        '
        XXY$ = REBLA$
        CAPEN = !Cant_Ocom - !Cant_Rec
        If CAPEN < 0.05 Then CAPEN = 0
        PREUNID# = !PrUn_Neto * TICAMBIO
        '
        CIXI% = !Cod_Inte
        Call REPLA(XXY$, MKI$(CIXI%), 9, 2)
        If CIXI% <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
        On Error Resume Next
        UNIMA$ = TRIM$(!Uni_Mas)
        On Error GoTo 0
        If UNIMA$ = "" Then UNIMA$ = TRIM$(UNIMED$(CIXI%))
        Call REPLA(XXY$, UNIMA$, 11, 4)
        '\\\OT-05-0693-WAR-24/10/05///
          If IsNull(!Sect_Ocom) Then
            SECT1$ = ""
          Else
            SECT1$ = !Sect_Ocom
          End If
          Call REPLA(XXY$, SECT1$, 15, 16)
            CCU% = !IContab
            CCU1$ = CODCUE$(CCU%)
          Call REPLA(XXY$, CCU1$, 105, 12)
        '\\\OT-05-0693-WAR-FIN///
        Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
        If IsNull(!Uni_Com) Then
          UNICOMOC$ = "  "
        Else
          UNICOMOC$ = !Uni_Com
        End If
        Call REPLA(XXY$, UNICOMOC$, 59, 4)
        If TRIM$(Mid$(XXY$, 11, 4)) = "" Then Call REPLA(XXY$, UNICOMOC$, 11, 4)
        '
        Call REPLA(XXY$, MKS$(!Coef_Unids), 63, 4)
        Call REPLA(XXY$, MKD$(!Pre_Unit), 67, 8)
        Call REPLA(XXY$, MKS$(!Desc_Item), 75, 4)
        Call REPLA(XXY$, MKD$(!PrUn_Neto), 79, 8)
        Call REPLA(XXY$, MKD$(!Cant_Ocom * !PrUn_Neto), 87, 8)
        Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Sol)), 95, 2)
        FENTREPAR% = CVINT(!Fentre_Sol)
        If FENTREPAR% > 0 Then
          If FENTREPAR% < FENTREGEN% Then
            FENTREGEN% = FENTREPAR%
          End If
        End If
        '
        If IsNull(!CONSIMAT) Then
          CONSIMATOC$ = "  "
        Else
          CONSIMATOC$ = !CONSIMAT
        End If
        Call REPLA(XXY$, CONSIMATOC$, 97, 2)
        Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
          STATUI% = !Stat_Ocom: While STATUI% < 0: STATUI% = STATUI% + 256: Wend
        Call REPLA(XXY$, Chr$(STATUI%), 124, 1)
        Call REPLA(XXY$, MKS$(!NuOrd_Item), 125, 4)
        '
          NUECUEIMPO$ = CODCUE$(!IContab)
        Call REPLA(XXY$, NUECUEIMPO$, 105, 12)
          'SECTOROC$ = !Sect_Ocom '***' ESTA LINEA NO ESTABA ATAJADA
          'Call REPLA(XXY$, SECTOROC$, 15, 12)
          '
        Call REGAPP("!OCOMDETA", XXY$)
          '
          DELIBX$ = RTrim$(!DELIBRE)
          If DELIBX$ = "" Then DELIBX$ = DESCRIT0$(CIXI%)
        Call REGAPP(ARCHILIB$, DELIBX$)
        '
        FOMODOC04.Text7 = CODCUE$(!IContab)
        On Error Resume Next
        FOMODOC04.Text18 = !Sect_Ocom
        On Error GoTo 0
        '
      .MoveNext
      Loop
    End With
    '
    If FENTREGEN% = 32767 Or FENTREGEN% < FECHEMI% Then FENTREGEN% = FECHEMI%
    FOMODOC04.Text14 = FECHATEX$(FENTREGEN%)
    '
    Call FOMODOC04.CARCOANT
    Call FOMODOC04.CARLISCO
    Call FOMODOC04.CALTOCOM
    FOMODOC04.Show 1
    GoTo 3
    '
999 Screen.MousePointer = 1
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
    '
    'Cargo en el list las O/C que puedan ser Anuladas.
5   Screen.MousePointer = 11
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom <8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 1 "
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom  ASC "
    Set OCANULTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCANULTEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        J& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
            J& = J& + 1
            Call GRAREG("INDIOCPE", XXY$, J&)
            .MoveNext
        Loop
    End With
    '
    Call SETULTREG("INDIOCPE", J&)
    Call CIERRARCH("INDIOCPE")
    Call BAJALDISCO
    Call FRESHECHO("INDIOCPE")
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
10  YY$ = REGSEL$("INDIOCPE/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    'nuevo marcado de anulacion
    With OcomEnca
      ABUS$ = "Nume_Ocom = " & NUOCO& & "AND Stat_Ocom < 8"
      RFXX$ = "(anul) "
      .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 9                          'cambio status
        !Refe_OcoM = Left$(RFXX$ + !Refe_OcoM, 64)        'cambio referen
        .Update
      End If
    End With
    '
    YAPRIREC% = 0
    With OcomDeta
      ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Stat_Ocom < 8"
20    If YAPRIREC% = 0 Then
          .FindFirst ABUS$
          YAPRIREC% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 9
        .Update
        GoTo 20
      End If
    End With
    OcomDeta.Close
    OcomEnca.Close
    '
    Screen.MousePointer = 11
    '
    J& = 0
    For I& = 1 To ULTREG&("INDIOCPE")
      XX$ = REGLEIDO$("INDIOCPE", I&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("INDIOCPE", XX$, J&)
      End If
    Next I&
    Call SETULTREG("INDIOCPE", J&)
    '
    ' borrar de lista de materiales
    NOCOBUS$ = TRIM$(Str$(NUOCO&))
    While Len(NOCOBUS$) < 6: NOCOBUS$ = "0" + NOCOBUS$: Wend
    NOCOBUS$ = "OC-" + Left$(NOCOBUS$, 6) + "-"

    Call ABREORTRA
    ABUS$ = "Nro_Ocom LIKE '" & NOCOBUS$ & "*'"
    With ListaMat
40    .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Nro_Ocom = ""
        .Update
        GoTo 40
      End If
    End With

    ' CANCELAR RESERVAS
    Call ABRESTOCKS
    With MaConsig
      .FindFirst ABUS$
42    If .NoMatch = False Then
        .Edit
        !Cant_Asig = !Cant_Entreg
        .Update
        .FindNext ABUS$
        GoTo 42
      End If
    End With
    '
    '
    JJ& = 0
    For J& = 1 To ULTREG&("MACONSIG")
      XX$ = REGLEIDO$("MACONSIG", J&)
      If XVALO(Mid$(XX$, 24, 6)) <> NUOCO& Then
           JJ& = JJ& + 1
           Call GRAREG("MACONSIG", XX$, JJ&)
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
    Call FRESHECHO("!INDIOCOM")
    Screen.MousePointer = 1
    GoTo 5
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'
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
         HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
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
'3  Call COPYSTRU("INDIOCOM", "IOCANUL")
'   J& = 0: K& = 0: OCOA& = 0
'   '
'   FIN& = ULTREG&("OCOMENCA")
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMENCA", U&)
'      If Asc(Mid$(X$, 54, 1)) = 9 Then ' esta anulada
'5       Y$ = Left$(X$, 48) + MKS$(U&)
'        J& = J& + 1
'        Call GRAREG("IOCANUL", Y$, J&)
'      End If
'9   Next U&


    'CARGO LA LISTA DE O/C QUE SE PUEDEN RESTAURAR

3   Call COPYSTRU("INDIOCOM", "IOCANUL")
    Screen.MousePointer = 11
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom = 9 " ' Si esta anulada
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set OCRESTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCRESTEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No  hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      J& = 0
      Do While (.EOF = False)
                XXY$ = REGBLAN$("IOCANUL")
                Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
                Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
                J& = J& + 1
                Call GRAREG("IOCANUL", XXY$, J&)
                .MoveNext
      Loop
    End With
    Call SETULTREG("IOCANUL", J&)
    Call CIERRARCH("IOCANUL")
    Call BAJALDISCO
    Call FRESHECHO("IOCANUL")
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
     
    'nuevo marcado de anulacion
     
     With OcomEnca
      ABUS$ = "Nume_Ocom = " & NUOCO& & "AND Stat_Ocom = 9"
      .FindFirst ABUS$
      If .NoMatch = False Then
        .Edit
        !Stat_Ocom = 0                      'cambio status
        FECHEMII$ = Format(!Femi_Ocom, "dd/mm/yy")
        PRORAI$ = !RaSo_Prov
        REFEI$ = FECHEMII$ & " - " & PRORAI$
        !Refe_OcoM = Left$(REFEI$, 64)   'cambio referencia
        .Update
      End If
    End With
    '
    YAPRIREC% = 0
    With OcomDeta
      ABUS$ = "Nume_Ocom = " & NUOCO& & " AND Stat_Ocom = 9"
20    If YAPRIREC% = 0 Then
          .FindFirst ABUS$
          YAPRIREC% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Edit
        STATUIX% = 0
        If !Cant_Rec > 0 Then STATUIX% = 2
        If !Cant_Rec >= !Cant_Ocom - 0.05 Then STATUIX% = 3
        !Stat_Ocom = STATUIX%
        .Update
        GoTo 20
      End If
    End With
   OcomEnca.Close
   OcomDeta.Close
    
     '''
'    RENCA& = CVS(Mid$(YY$, 49, 4))
'    If RENCA& < 1 Or RENCA& > ULTREG&("OCOMENCA") Then
'      Call MENSERR(24, "Error de Indices - Consulte")
'      GoTo 999
'    End If
'    XX$ = REGLEIDO$("OCOMENCA", RENCA&)
'    If CVS(Left$(XX$, 4)) <> NUOCO& Then
'      Call MENSERR(24, "Error de Indices - Consulte")
'      GoTo 999
'    End If
    '
    Screen.MousePointer = 11
'    FIN& = ULTREG&("OCOMENCA")
'    For J& = 1 To FIN&
'      Call TRACE(20, J&, FIN&)
'      XX$ = REGLEIDO$("OCOMENCA", J&)
'      If CVS(Left$(XX$, 4)) = nuoco& Then
'          FECHEMII% = CVI(Mid$(XX$, 49, 2))
'          NPROI% = CVI(Mid$(XX$, 51, 2))
'          FEX = FECHEMII%
'          REFEI$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPROI%)
'          Call REPLA(XX$, Chr$(0), 54, 1)
'          Call REPLA(XX$, REFEI$, 5, 44)
'          Call GRAREG("OCOMENCA", XX$, J&)
'          YADESA% = 1
'        Else
'          If YADESA% > 0 Then GoTo 19
'      End If
'    Next J&
'19  Call CIERRARCH("OCOMENCA")
'    '
'    FIN& = ULTREG&("OCOMDETA")
'    YADESA% = 0
'    For J& = FIN& To 1 Step -1
'      Call TRACE(20, FIN& - J& + 1, FIN&)
'      XX$ = REGLEIDO$("OCOMDETA", J&)
'      If CVS(Left$(XX$, 4)) = nuoco& Then
'        If Asc(Mid$(XX$, 124, 1)) = 9 Then
'          CAPEDI1 = CVS(Mid$(XX$, 55, 4))
'          CARECI1 = CVS(Mid$(XX$, 99, 4))
'          STATUIX% = 0
'          If CARECI1 > 0 Then STATUIX% = 2
'          If CARECI1 >= CAPEDI1 - 0.05 Then STATUIX% = 3
'          Call REPLA(XX$, Chr$(STATUIX%), 124, 1)
'          Call GRAREG("OCOMDETA", XX$, J&)
'        End If
'      End If
'    Next J&
'    Call CIERRARCH("OCOMDETA")
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
    '
    GoTo 3
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
    Call APERBASDAT("COMPRAS")
10  SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < ( OCOMDETA.Cant_Ocom - 0.05 ) "
    SQLX$ = SQLX$ + " AND OCOMDETA.Cod_Inte > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom DESC "
    Set OCCUMTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCCUMTEMP
      On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No hay Ordenes de Compra\que Puedan Marcarse como cumplidas.")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        XXY$ = REGBLAN$("INDIOCUM")
        Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
        Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
        J& = J& + 1
        Call GRAREG("INDIOCUM", XXY$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("INDIOCUM", J&)
    Call CIERRARCH("INDIOCUM")
    Call FRESHECHO("INDIOCUM")
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
'10  YY$ = REGSEL$("INDIOCUM/Indice de Ordenes de Compra Pendientes para Anulación")
'    If Len(YY$) <= 4 Then GoTo 999
    '
    'MODIFICADO POR QUE A VECES NO ESTABA TRAYENDO LA ORDEN DE COMPRA CORRECTA.
    Call SELECHO("INDIOCUM/Indice de Ordenes de Compra Pendientes para Anulación")
    NUOCO& = XVALO(VALACT1$("INDIOCUM"))
    If NUOCO& < 1 Then GoTo 999
    DAPORCU04.Label14 = NUOCO&

'    NUOCO& = CVS(Left$(YY$, 4))
'    DAPORCU04.Label14 = NUOCO&
    '
    DAPORCU04.Show 1
    '
    ' A PARTIR DE AQUI MARCAR REGISTROS QUE SE DAN POR CUMPLIDOS
    Call APERBASDAT("COMPRAS")
    If XVALO(DAPORCU04.Label14) = NUOCO& Then
       Screen.MousePointer = 11
       'URE& = ULTREG&("OCOMDETA")
       For U& = 1 To DAPORCU04.LICOPEN.ListCount
         If DAPORCU04.LICOPEN.Selected(U& - 1) = True Then
            TTXX$ = DAPORCU04.LICOPEN.List(U& - 1)
            NUORIT% = XVALO(TRIM$(Mid$(TTXX$, 121, 8)))
            '
            SQLX$ = "SELECT * FROM OCOMDETA "
            SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
            SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
            SQLX$ = SQLX$ + "AND NuOrd_Item = " & NUORIT% & " "
            '
            Set MARCTEMP = compras.OpenRecordset(SQLX$, dbOpenDynaset)
            '
            With MARCTEMP
              .MoveFirst
              Do While (.EOF = False)
                 COEFCAN = !Coef_Unids
                 If COEFCAN < 0.005 Then COEFCAN = 1
                 CAORI = !Cant_Ocom * COEFCAN
                 CAREC = !Cant_Rec
                 CAPEN = CAORI - CAREC
                 If CAPEN >= 0.005 Then
                    .Edit
                    !Stat_Ocom = 8
                    .Update
                 End If
              .MoveNext
              Loop
            End With
            '
            NOCOM1$ = TRIM$(Str$(NUOCO&))
              While Len(NOCOM1$) < 6: NOCOM1$ = "0" + NOCOM1$: Wend
            ITNU1$ = TRIM$(Str$(NUORIT%))
              While Len(ITNU1$) < 2: ITNU1$ = "0" + ITNU1$: Wend
            NOCOBUS$ = "OC-" + Left$(NOCOM1$, 6) + "-" + ITNU1$
            ABUS$ = "Nro_Ocom = '" & NOCOBUS$ & "'"
            '
            ' CANCELAR RESERVAS
            Call ABRESTOCKS
            With MaConsig
              .FindFirst ABUS$
42            If .NoMatch = False Then
                .Edit
                !Cant_Asig = !Cant_Entreg
                .Update
                .FindNext ABUS$
                GoTo 42
              End If
            End With
            '
         End If
       Next U&
    End If
    '
    Unload DAPORCU04
    Screen.MousePointer = 1
    '
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub

Sub DEVOLCONSI()
        '
1105    Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 1
        If NPROVE% > 0 Then
             EBOL$ = MKI$(NPROVE%)
          Else
             EBOL$ = OBJPLA$("SELPROVE", "")
             If EBOL$ = "" Then
                GoTo 9999
             End If
        End If
        '
        NC% = CVI(EBOL$)
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            NPROVE% = 0
            GoTo 1105
        End If
        '
        NC1% = (-1) * NC%
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        Screen.MousePointer = 11
        HAYOC% = 0
        Call APERBASDAT("STOCKS")
        '
        ' ESTO ELIMINA LOS REGISTROS EN LOS QUE QUEDA UN PEQUEÑO PUCHO
        ' PENDIENTE DE RENDICION O ENTREGA AL PROVEEDORE. RECLAMADO POR
        ' NIUBO DE ESCORIAL EL 1RO DE OCTUBRE 2008.
        SQLY$ = "DELETE FROM MACONSIG "
        SQLY$ = SQLY$ + " WHERE Cant_Entreg >= (Cant_Asig - 0.005) "
        SQLY$ = SQLY$ + " AND Cant_Rendid >= (Cant_Entreg - 0.005) "
        Stocks.Execute SQLY$
        '
        SQLX$ = "SELECT * FROM MACONSIG "
        SQLX$ = SQLX$ + " WHERE (Cant_Entreg - Cant_Rendid)>0 "
        SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
        Set MACONSITMP = Stocks.OpenRecordset(SQLX$, 4)
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
          Do While .EOF = False
            CAPENE# = XVALO(!Cant_Entreg) - XVALO(!Cant_Rendid)
            NUOC& = XVALO(Mid$(!Nro_Ocom, 4, 6))
            CI% = XVALO(!Cod_Inte)
            '
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, UNIMED$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
            Call REPLA(Y$, MKS$(NUOC&), 151, 4)
            Call REPLA(Y$, MKS$(I&), 189, 4)
            HAYOC% = 1
            Call REGAPP("!BOLREDET", Y$)
            .MoveNext
          Loop
          End If
        End With
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Rendición Materiales Consignados - " + TRIM$(RASOCLI$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
        '
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            LOTE$ = Mid$(YY$, 69, 12)
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                If CAN# > 0.005 Then
                  NUOC& = CVS(Mid$(YY$, 151, 4))
                  DOSEC$ = ""
                  With MaConsig
                    SQLX$ = " TRIM(Mid$(Nro_Ocom, 4, 6)) =  " & NUOC & " "
                    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & ""
                    SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
                    .FindFirst SQLX$
                    If .NoMatch = False Then
                      .Edit
                      DOSEC$ = SAFETEXT$(!Nro_Ocom)
                      TOENTRE# = XVALO(!Cant_Rendid) + CAN#
                      !Cant_Rendid = TOENTRE#
                      .Update
                    End If
                  End With
                End If
              End If
            End If
            '
          Next YX&
          Call CIERRARCH("MACONSIG")
          '
          Screen.MousePointer = 1
          '
1199    End If
        '
9999    Call CIERRARCH("*.*")
        '
End Sub

Sub RANKPRO()
'\\\OT-05-0381-WAR-07/06/05///
    Dim TOTCOMPRA#(32767)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTCOMPRA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "    ' NO ANULADAS
    '
    Set CABATEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
        UK& = UK& + 1
        Call TRACE(20, UK&, FIN&)
        NPRO% = !NPRO_OCOM
        If NPRO% > 0 Then
          If NPRO% > PROMAX% Then PROMAX% = NPRO%
          IMPNETO# = (!PrUn_Neto * !Cant_Ocom)
          TOTCOMPRA#(NPRO%) = TOTCOMPRA#(NPRO%) + (IMPNETO# * TICAMONE(!Mone_Emis))
        End If
        .MoveNext
      Loop
    End With
    '
    For I% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTCOMPRA#(I%)
    Next I%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTCOMPRA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTCOMPRA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANPRO")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0381-WAR-FIN///
End Sub

Sub RANKART()
'\\\OT-05-0382-WAR-07/06/05
    '
    Dim TOTART#(32767)
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTART#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(DES%)))
    HASDAT = Int(CDbl(CVDAT(HAS%)))
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "  ' NO ANULADAS
    '
    Set CABATEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
         UK& = UK& + 1
         Call TRACE(20, UK&, FIN&)
         NART% = !Cod_Inte
         IMPNETO# = (!PrUn_Neto * !Cant_Ocom)
         If NART% > 0 Then
           If NART% > PROMAX% Then PROMAX% = NART%
           TOTART#(NART%) = TOTART#(NART%) + (IMPNETO# * TICAMONE(!Mone_Emis))
         Else
           TOTART#(0) = TOTART#(0) + (IMPNETO# * TICAMONE(!Mone_Emis))
         End If
         .MoveNext
      Loop
    End With
    '
    For I% = 0 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTART#(I%)
    Next I%
    '
    List1.Clear
    For NCI% = 0 To 32766
       If Abs(TOTART#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTART#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTART#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTART#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANART")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0382-WAR-FIN
End Sub

Sub DEPUCONSIG()
   '
   SQLX$ = "SELECT * FROM MACONSIG"
   SQLX$ = SQLX$ + " WHERE CANT_ASIG - CANT_ENTREG > 0.05 "
   Call APERBASDAT("STOCKS")
   Set MACONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   Call APERBASDAT("COMPRAS")
   With MACONSITMP
      Do While Not .EOF
         OCOMP$ = SAFETEXT$(!Nro_Ocom)
         NUME& = XVALO(Mid$(OCOMP$, 4, 6))
         ORD% = XVALO(Mid$(OCOMP$, 11, 26))
         '
         CONDI$ = "SELECT STAT_OCOM FROM OCOMDETA WHERE "
         CONDI$ = CONDI$ + " Nuord_Item = " & ORD%
         CONDI$ = CONDI$ + " AND NUME_OCOM = " & NUME&
         '
         Set COMDET = compras.OpenRecordset(CONDI$, dbOpenDynaset)
         If Not (COMDET.EOF) And Not (COMDET.BOF) Then
           COMDET.MoveFirst
           STA% = XVALO(COMDET!Stat_Ocom)
         End If
         COMDET.Close
         '
         If STA% >= 8 Or STA% = 3 Then
           .Edit
           !Cant_Asig = XVALO(!Cant_Entreg)
           .Update
         End If
        .MoveNext
      Loop
    
    End With
    MACONSITMP.Close
    Set MACONSITMP = Nothing
    Screen.MousePointer = 1
    Call COMUNI("Proceso Terminado")
End Sub

Sub REVERTIR_ITEMCUMPLI()
    'REVERTIR ITEM CUMPLIDOS
    '
    TITU$ = "DAR POR CUMPLIDOS ITEMS DE COMPRAS"
    '
2   Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCAC")
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT  DISTINCT Nume_Ocom, Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 AND Nume_Ocom > 0 ORDER BY Nume_Ocom "
    Set OCCUMTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCCUMTEMP
      On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No hay Ordenes de Compra\que Puedan Recuperarse")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        XXY$ = REGBLAN$("INDIOCAC")
        Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
        Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
        J& = J& + 1
        Call GRAREG("INDIOCAC", XXY$, J&)
      .MoveNext
      Loop
    End With
    '
    On Error Resume Next
    OCCUMTEMP.Close
    Set OCCUMTEMP = Nothing
    On Error GoTo 0
    
    Call SETULTREG("INDIOCAC", J&)
    Call CIERRARCH("INDIOCAC")
    Call FRESHECHO("INDIOCAC")
    
10  Call SELECHO("INDIOCAC/Indice de Ordenes de Compra")
    NRO& = XVALO(VALACT1("INDIOCAC"))
    If NRO& < 1 Then Exit Sub
    '
    Call COPYSTRU("COPRECOM", "INDICUM")
    SQLX$ = " SELECT  DISTINCT Nume_Ocom, Cod_Inte "
    SQLX$ = SQLX$ + " FROM OCOMDETA "
    SQLX$ = SQLX$ + " WHERE CANT_REC < CANT_OCOM AND STAT_OCOM = 8 AND NUME_OCOM = " & NRO&
    Set OCCUMTEMP1 = compras.OpenRecordset(SQLX$, dbOpenDynaset)
    With OCCUMTEMP1
      On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No Hay Items Cumplidas En Esta Orden \Que Puedan Restaurarse")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        XXY$ = REGBLAN$("INDICUM")
        Call REPLA(XXY$, MKI$(!Cod_Inte), 1, 2) 'Nº O/C
        J& = J& + 1
        Call GRAREG("INDICUM", XXY$, J&)
      .MoveNext
      Loop
    End With
    '
    On Error Resume Next
    OCCUMTEMP1.Close
    Set OCCUMTEMP1 = Nothing
    On Error GoTo 0
    
    Call SETULTREG("INDICUM", J&)
    Call CIERRARCH("INDICUM")
    Call FRESHECHO("INDICUM")
    
20  Call SELECHO("INDICUM/Indice de Items O.C.: " & NRO&)
    CODEXTER$ = VALACT1("INDICUM")
    CIXI% = CODINT%(CODEXTER$)
    If CIXI% < 1 Then GoTo 10
    
    'REVERTIR MARCA DE ITEM CUMPLIDO ' SE AGREGA CONDICION STAT_OCOM = 8 PARA QUE REVIERTA SOLO LOS CUMPLIDOS
    If COMALTER%("¿ Confirma Revertir Item: " & CODEXTER$ & "\Marcado Como Cumplido  \\No, Abandonar\Si, Revertir") = 2 Then
       SQLX$ = " SELECT * FROM OCOMDETA WHERE NUME_OCOM = " & NRO& & " AND STAT_OCOM = 8 and COD_INTE = " & CIXI%
       Set rst = compras.OpenRecordset(SQLX$, dbOpenDynaset)
       rst.MoveFirst
       rst.Edit
       rst!Stat_Ocom = 0
       rst.Update
       rst.Close
       Set rst = Nothing
       Call COMUNI("Se Ha Completado la Operación")
       GoTo 20
    Else
      GoTo 20
    End If
    
End Sub


