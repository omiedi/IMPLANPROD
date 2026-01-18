VERSION 5.00
Begin VB.Form FORMOCOM 
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenes de Compra - Proveedores Externos"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9450
      Picture         =   "FORMOCOM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Códigos Equivalentes de Proveedor"
      Top             =   4830
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFC0&
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
      Left            =   5250
      TabIndex        =   58
      Top             =   5670
      Width           =   3585
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0FFC0&
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
         TabIndex        =   60
         Top             =   250
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0FFC0&
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
         Left            =   420
         TabIndex        =   59
         Top             =   250
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Master"
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
      Left            =   9450
      Picture         =   "FORMOCOM.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Maestro de Items de Stock"
      Top             =   3335
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9450
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5700
      Width           =   770
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   29
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
      Left            =   9450
      Picture         =   "FORMOCOM.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   2580
      Width           =   750
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Proveed"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9450
      Picture         =   "FORMOCOM.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "A-B-M Maestro de Proveedores"
      Top             =   4080
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFC0&
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
      Height          =   6315
      Left            =   4830
      TabIndex        =   22
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Control de Gestión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   420
         TabIndex        =   30
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1060
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   63
            Top             =   2310
            Width           =   2325
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
               Picture         =   "FORMOCOM.frx":0BA0
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Análisis Estadístico de Consumo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estadística de Consumo"
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
               TabIndex        =   65
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture13 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   890
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   47
            Top             =   1895
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
               Picture         =   "FORMOCOM.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   12
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
               TabIndex        =   48
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   720
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   31
            Top             =   1500
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
               Left            =   -20
               Picture         =   "FORMOCOM.frx":11B4
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "En Consignación"
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
               TabIndex        =   32
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   550
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   33
            Top             =   1105
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
               Picture         =   "FORMOCOM.frx":14BE
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
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
               Left            =   735
               TabIndex        =   34
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   380
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   35
            Top             =   710
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
               Picture         =   "FORMOCOM.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Pendientes por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
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
               TabIndex        =   36
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   37
            Top             =   315
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
               Picture         =   "FORMOCOM.frx":1C0A
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "Ordenes de Compra por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "O/C's por Número"
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
               TabIndex        =   38
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0FFC0&
         Caption         =   "Control de Ultimas Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   420
         TabIndex        =   23
         Top             =   3675
         Width           =   3585
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   24
            Top             =   840
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
               Picture         =   "FORMOCOM.frx":1F14
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Ultimas Compras por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Control por Proveedor"
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
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   26
            Top             =   315
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
               Picture         =   "FORMOCOM.frx":221E
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Ultimas Compras por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Control            por Artículo"
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
               TabIndex        =   27
               Top             =   45
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Gestión de Compras"
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
      Left            =   210
      TabIndex        =   20
      Top             =   210
      Width           =   4425
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0FFC0&
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
         Left            =   420
         TabIndex        =   44
         Top             =   3100
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   45
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
               Picture         =   "FORMOCOM.frx":2368
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Entrega en Consignación "
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command15 
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
               Left            =   2340
               Picture         =   "FORMOCOM.frx":2672
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Restaurar"
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
               TabIndex        =   46
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0FFC0&
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
         Height          =   1695
         Left            =   420
         TabIndex        =   39
         Top             =   4410
         Width           =   3585
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   40
            Top             =   795
            Width           =   2535
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
               Picture         =   "FORMOCOM.frx":27BC
               Style           =   1  'Graphical
               TabIndex        =   7
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
               TabIndex        =   41
               Top             =   90
               Width           =   1170
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   42
            Top             =   315
            Width           =   2535
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
               Picture         =   "FORMOCOM.frx":2AC6
               Style           =   1  'Graphical
               TabIndex        =   6
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
               Left            =   840
               TabIndex        =   43
               Top             =   50
               Width           =   1425
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0FFC0&
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
         Height          =   2535
         Left            =   420
         TabIndex        =   21
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   49
            Top             =   1680
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
               Picture         =   "FORMOCOM.frx":2DD0
               Style           =   1  'Graphical
               TabIndex        =   61
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
               Picture         =   "FORMOCOM.frx":2F1A
               Style           =   1  'Graphical
               TabIndex        =   3
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
               TabIndex        =   50
               Top             =   105
               Width           =   1065
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   51
            Top             =   1225
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
               Picture         =   "FORMOCOM.frx":3224
               Style           =   1  'Graphical
               TabIndex        =   2
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
               TabIndex        =   52
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   53
            Top             =   770
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
               Picture         =   "FORMOCOM.frx":352E
               Style           =   1  'Graphical
               TabIndex        =   1
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
               TabIndex        =   54
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture16 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   55
            Top             =   315
            Width           =   2325
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
               Picture         =   "FORMOCOM.frx":3678
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar una Nueva Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva O/Compra"
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
      Left            =   9450
      Picture         =   "FORMOCOM.frx":3982
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
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
      Height          =   720
      Left            =   9450
      Picture         =   "FORMOCOM.frx":3ACC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1070
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
      Left            =   9450
      Picture         =   "FORMOCOM.frx":3DD6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Acceso Directo a Otros Procesos"
      Top             =   1825
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8860
      Picture         =   "FORMOCOM.frx":40E0
      Top             =   5990
      Width           =   2010
   End
End
Attribute VB_Name = "FORMOCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call DESANULOC
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    If EXISTE%("AMASTO02.EXE") > 0 Then
         Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
       ElseIf EXISTE%("AMASTO01.EXE") > 0 Then
         Call CONECRUN("AMASTO01", "Maestro de Items de Stock")
       Else
         Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    End If
99  Command11.Enabled = True
End Sub


Private Sub Command12_Click()
    Command12.Enabled = False
    'If Option1.Value = True Then
    '     Call CONECRUN("*COCOMPR", "Consulta Compras por Proveedor")
    '   Else
         Call SELCOMRAFE
         Call CONECRUN("*SOCOMPR", "Listado de Compras por Proveedor")
    'End If
99  Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    Call MODIOCOM
99  Command13.Enabled = True
    '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call OCOMANU
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Call OPNOIN("")
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "NEUMANN") > 0 Or InStr(EPAC$, "ARCON") > 0 Or InStr(EPAC$, "YALTRES") > 0 Then
        Call COCOBREC
      Else
        Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
    End If
    Command16.Enabled = True
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
    jj& = 0
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
          jj& = jj& + 1
          Call GRAREG("!CARDETOC", YY$, jj&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", jj&)
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

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    SETUPCOM.Show 1
    Call COMUNI("Reconfiguración Completa.\Re-Inicie Ahora la Aplicación")
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
    Command19.Enabled = False
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    Command18.Enabled = True
    '
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORPR")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Call ENTRECONSI(0)
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    'If Option1.Value = True Then
    '     Call CONECRUN("*COCOMAR", "Consulta Compras por Articulo")
    '   Else
         Call SELCOMRAFE
         Call CONECRUN("*SOCOMAR", "Listado de Compras por Articulo")
    'End If
    Command20.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   Screen.MousePointer = 11
   jj& = 0
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
               jj& = jj& + 1
               Call GRAREG("!OCOMPRA", XX$, jj&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", jj&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If jj& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
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
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command22.Enabled = True
    '
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command23.Enabled = True
End Sub

Private Sub Command3_Click()
    OPXX% = ALTERNA%("Base de Datos Completa\Ordenes Pendientes")
    If Option1.Value = True Then
        Select Case OPXX%
          Case 1
            Call FINAL("*COCOMP0")
          Case 2
            Call COCOPEN        ' ORDENES PENDIENTES
            Call FINAL("*COCOPEN")
          Case Else
            Exit Sub
        End Select
        '
      ElseIf Option2.Value = True Then
        Select Case OPXX%
          Case 1
            Call FINAL("*LICOMP0")
          Case 2
            Call COCOPEN        ' ORDENES PENDIENTES
            Call FINAL("*LICOPEN")
          Case Else
            Exit Sub
        End Select
    End If
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call COCOPEN    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        Call FINAL("*CORCPRO")
      ElseIf Option2.Value = True Then
        Call FINAL("*SORCPRO")
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call COCOPEN    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        Call FINAL("*CORCMAT")
      ElseIf Option2.Value = True Then
        Call FINAL("*SORCMAT")
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call PEDISUMI
    'Call OPNOIN("Pedido de Suministro")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    Call OCOMNUE
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") < 1 Then GoTo 20
    OPX% = ALTERNA%("Control Costos Reposición\Otros Accesos Directos")
    If OPX% = 1 Then
        CONTROCOS.Show 1
      ElseIf OPX% = 2 Then
20      Call ACCDIR("OCOMPRA")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    If Option1.Value = True Then
        MODOEMI% = COMALTER%("Opciones de Consulta:\\Por Proveedor\Por Artículo\Cancelar")
        If MODOEMI% = 1 Then
             Call FINAL("*COMATCO")
           ElseIf MODOEMI% = 2 Then
             Call FINAL("*COMACON")
        End If
        Exit Sub
      ElseIf Option2.Value = True Then
        MODOEMI% = COMALTER%("Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Cancelar")
        If MODOEMI% = 1 Then
             Call FINAL("*LIMATCO")
           ElseIf MODOEMI% = 2 Or MODOEMI% = 3 Then
             Call GENCONPEN(MODOEMI%)
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    'Show
    'Refresh
    VERANTER$ = "OCOMPRA"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If InStr(EPAC$, "NEUMANN") > 0 Or InStr(EPAC$, "ARCON") > 0 Or InStr(EPAC$, "YALTRES") > 0 Then
       Command16.ToolTipText = "Indices de Calidad de Proveedores"
    End If
    '
    Call CONVERCOM
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
Sub COCOPEN()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   jj& = 0
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
         jj& = jj& + 1
         Call GRAREG("OPENDETA", XXY$, jj&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", jj&)
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
   jj& = 0
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
                jj& = jj& + 1
                Call GRAREG("COCAREC", Z$, jj&)
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
                 If REG& <= jj& Then
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
                 If REG& <= jj& Then
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
                 If REG& <= jj& Then
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
   Call SETULTREG("COCAREC", jj&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub OCOMNUE()                           ' NUEVA ORDEN DE COMPRA
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
    MONUME$ = "AUTO"
    If CONTROL$("", "NUMEOCOM") = "MANUAL" Then
        MONUME$ = "MANUAL"
    End If
    '
    MONECOI% = Val(CONTROL$("OCOMPRA", "MONECOS"))
    If ECOARCH$("TAMONED", Chr$(MONECOI%)) = "" Then
       MONECOI% = 0
    End If
    '
    MODO% = COMALTER%("Puede Generar Ordenes de Compra\en Forma Directa ó sobre la Base\del Cálculo MRP\\Directa\MRP\Cancelar")
    FORMOCOM.Refresh
    If MODO% < 1 Or MODO% > 2 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
    AVALIDA$ = "MASTER"
    If MODO% = 2 Then AVALIDA$ = "OCOMEMI"
    '
    ASELPRO$ = "ACREDOR"
    If MODO% = 2 Then
      Call COPYSTRU("PROVED1", "PROVSEL")
      ASELPRO$ = "PROVSEL"
      Screen.MousePointer = 11
      PPXXZ$ = String$(32767, 0)
      For U& = 1 To ULTREG("PLACOMP")
        XX$ = REGLEIDO$("PLACOMP", U&)
        NPRX% = CVI(Mid$(XX$, 33, 2))
        If NPRX% > 0 Then
          Mid$(PPXXZ$, NPRX%, 1) = Chr$(1)
        End If
      Next U&
      '
      jj& = 0
      For U& = 1 To ULTREG&("ACREDOR")
        XX$ = REGLEIDO$("ACREDOR", U&)
        NPRX% = CVI(Left$(XX$, 2))
        If NPRX% > 0 Then
          If Asc(Mid$(PPXXZ$, NPRX%, 1)) > 0 Then
            Call REPLA(XX$, " (*)", 29, 4)
            jj& = jj& + 1
            Call GRAREG("PROVSEL", XX$, jj&)
          End If
        End If
      Next U&
      jj& = jj& + 1
      Call GRAREG("PROVSEL", MKI$(0) + String$(30, "-"), jj&)
      '
      For U& = 1 To ULTREG&("ACREDOR")
        XX$ = REGLEIDO$("ACREDOR", U&)
        NPRX% = CVI(Left$(XX$, 2))
        If NPRX% > 0 Then
          If Asc(Mid$(PPXXZ$, NPRX%, 1)) <= 0 Then
            jj& = jj& + 1
            Call GRAREG("PROVSEL", XX$, jj&)
          End If
        End If
      Next U&
      Call SETULTREG("PROVSEL", jj&)
      Call CIERRARCH("PROVSEL")
    End If
    Screen.MousePointer = 1
    '
100 Call CIERRARCH("*.*")
    NRO& = PRONUOCO&
    '
101 Call SELECHO(ASELPRO$)
    NPRO% = Val(VALACT1$(ASELPRO))
    If NPRO% < 1 Or ECOARCH$("PROVED1", MKI$(NPRO%)) = "" Then
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Exit Sub
    End If
    '
    X0$ = String$(10, 0) + Space$(160)
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 7, 2)
    Call REPLA(X0$, MKI$(NPRO%), 9, 2)
    '
    DOMIPRO$ = TRIM$(CPOSCLI$((-1) * NPRO%))
    If DOMIPRO$ <> "" Then DOMIPRO$ = DOMIPRO$ + " - "
    Call REPLA(X0$, RASOCLI$((-1) * NPRO%), 11, 40)
    Call REPLA(X0$, DOMICLI$((-1) * NPRO%), 51, 40)
    Call REPLA(X0$, DOMIPRO$ + LOCACLI$((-1) * NPRO%), 91, 40)
    '
    NC1% = (-1) * NPRO%
    CPAG% = CPAGCLI%(NC1%)
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
         MONEMI% = MONEMICO%(NPRO%): TICAEMI = TICAMONE(MONEMI%)
      Else
         RFF$ = RECFILT$("OCOMENCA", 4, MKI$(NPRO%))
         If RFF$ <> "" Then
           REOCOV& = CVS(Mid$(RFF$, Len(RFF$) - 3))
           XZX$ = REGLEIDO$("OCOMENCA", REOCOV&)
           MONEMI% = Asc(Mid$(XZX$, 84, 1))
         End If
    End If
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
    ATEN$ = "": PODESC$ = "": TIEMBA% = 0
    RFF$ = RECFILT$("OCOMENCA", 4, MKI$(Abs(NPRO%)))
    For KKK% = Len(RFF$) - 3 To 1 Step -4
       REOCO& = CVS(Mid$(RFF$, KKK%, 4))
       XX$ = REGLEIDO$("OCOMENCA", REOCO&)
       ATEN$ = TRIM$(Mid$(XX$, 163, 30))
       PODESC$ = TRIM$(Mid$(XX$, 97, 16))
       TIEMBA% = Asc(Mid$(XX$, 154, 1))
       If ATEN$ <> "" Then Exit For
    Next KKK%
    Mid$(X0$, 131, 40) = ATEN$
    '
11  X1$ = OBJPLA$("OCOMPRA-ENCA", X0$)
    If X1$ = "" Then GoTo 101
    '
    FECHEMI% = CVI(Mid$(X1$, 5, 2))
    If FECHEMI% > HO% Then
       Call MENSERR(24, "Fecha de Emision no Válida")
       GoTo 11
    End If
    '
    FECHENT% = CVI(Mid$(X1$, 7, 2))
    If FECHENT% < FECHEMI% Or FECHENT% < HO% Then
       Call MENSERR(24, "Fecha de Entrega no Válida")
       GoTo 11
    End If
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
    ATEN$ = Mid$(X1$, 131, 40)
    '
    jj& = 0
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    If MODO% = 2 Then      ' POR MRP
      RFF$ = RECFILT$("PLACOMP", 8, MKI$(NPRO%))
      For KKI% = 1 To Len(RFF$) Step 4
        REPX& = CVS(Mid$(RFF$, KKI%, 4))
        XX$ = REGLEIDO$("PLACOMP", REPX&)
        CI0% = CVI(Left$(XX$, 2))
        FENTREIT% = CVI(Mid$(XX$, 3, 2))
          If FENTREIT% < HO% Then FENTREIT% = HO%
        CANTU = CVS(Mid$(XX$, 21, 4))
        ZZ$ = REBLA$
        Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
        Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
        Call REPLA(ZZ$, MKS$(CANTU), 55, 4)
        Call REPLA(ZZ$, UNIMED$(CI0%), 59, 4)
        Call REPLA(ZZ$, MKS$(1), 63, 4)
        Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
        '
        UPREDES$ = ""
        Call BUSCAULTIPRE(NPRO%, CI0%, UNIMAS$, UPREDES$)
        If UPREDES$ <> "" Then
          Call REPLA(ZZ$, Mid$(UPREDES$, 5), 63, 16)
        End If
        PREUNID# = CVD(Mid$(ZZ$, 67, 8))
        If PREUNID# < 0.00005 Then
          PREUNID# = CVS(Mid$(XX$, 35, 4))
          Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
        End If
        If MONEMI% > 1 Then
          PREUNID# = PREUNID# / TICAMONE(MONEMI%)
          Call REPLA(ZZ$, MKD$(PREUNID#), 67, 8)
        End If
        jj& = jj& + 1
        Call GRAREG("!OCOMDETA", ZZ$, jj&)
      Next KKI%
      Call SETULTREG("!OCOMDETA", jj&)
    End If
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - " + MONEDEM$ + " - " + TRIM$(RASOCLI$(NC1%))
    VTB% = VENTABU%("CAROCO1/ARCHECO(1)=" + AVALIDA$ + "/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    If ULTREG&("!OCOMDETA") < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Comprado")
       GoTo 102
    End If
    '
    IDENOC$ = MKS$(NRO&) + MKI$(FECHEMI%) + MKI$(NPRO%)
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(X$, IDENOC$, 1, 8)
       '
       CI0% = CVI(Mid$(X$, 9, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(X$, UNIMAS$, 11, 4)
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNICOM$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       CANTU = CVS(Mid$(X$, 55, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       '
       FENTREIT% = CVI(Mid$(X$, 95, 2))
       If FENTREIT% < HO% Then
          Call REPLA(X$, MKI$(FECHENT%), 95, 2)
       End If
       '
       PREUNID# = CVD(Mid$(X$, 67, 8))
       If PREUNID# < 0.00005 Then
          UPREDES$ = ""
          Call BUSCAULTIPRE(NPRO%, CI0%, UNIMAS$, UPREDES$)
          If UPREDES$ <> "" Then
            Call REPLA(X$, Mid$(UPREDES$, 5), 63, 16)
            PREUNID1# = CVD(Mid$(X$, 67, 8))
            If MONEMI% > 1 Then
              PREUNID1# = PREUNID1# / TICAMONE(MONEMI%)
              Call REPLA(X$, MKD$(PREUNID1#), 67, 8)
            End If
          End If
       End If
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       UNIST$ = UCase$(TRIM$(Mid$(X$, 11, 4)))
       UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
       COEFI = 1
       '
       If Left$(UNIST$, 2) <> "KG" And Left$(UNICO$, 2) = "KG" Then
           If UPREDES$ <> "" Then
             If UCase$(Left$(UPREDES$, 2)) = "KG" Then
               If CVS(Mid$(UPREDES$, 5, 4)) > 0.0005 Then
                 COEFI = CVS(Mid$(UPREDES$, 5, 4))
               End If
             End If
           End If
           If COEFI = 1 Then
             If PESUNI(CI0%) > 0.000005 Then
               COEFI = 1 / PESUNI(CI0%)
             End If
           End If
           GoTo 1021
       End If
       '
       If UNIST$ <> UNICO$ Then
           If UPREDES$ <> "" Then
             If UCase$(Left$(UPREDES$, 2)) = Left$(UNICO$, 2) Then
               If CVS(Mid$(UPREDES$, 5, 4)) > 0.0005 Then
                 COEFI = CVS(Mid$(UPREDES$, 5, 4))
               End If
             End If
           End If
       End If
       '
1021   CANSTXX = CANTU * COEFI
       Call REPLA(X$, MKS$(COEFI), 63, 4)
       Call REPLA(X$, MKS$(CANSTXX), 51, 4)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
103 VTB% = VENTABU%("CAROCO2/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
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
          GoTo 102
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
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANE# + TOTIVA#
    '
' presentar registro de totales y condicion de pago
    '
    If MONEMI% > 1 Then
        MONECOI% = MONEMI%
        TICAMBX = TICAEMI
      Else
        MONECOI% = MONEVACO%(NPRO%)
        TICAMBX = TICAMONE(MONECOI%)
    End If
123 ENCAO$ = REGBLAN$("OCOMENCA")
    If MONEMI% <= 1 Then TICAMBX = TICAMONE(MONECOI%)
    If TICAMBX < 0.5 Then TICAMBX = 1
    '
    Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
    FEX = FECHEMI%
    REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
    Call REPLA(ENCAO$, REFE$, 5, 44)
    Call REPLA(ENCAO$, MKI$(FECHEMI%), 49, 2)
    Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
    Call REPLA(ENCAO$, Chr$(0), 53, 1)
    Call REPLA(ENCAO$, Chr$(0), 54, 1)
    Call REPLA(ENCAO$, Chr$(MONECOI%), 83, 1)
    Call REPLA(ENCAO$, Chr$(MONEMI%), 84, 1)
    Call REPLA(ENCAO$, MKS$(TICAMBX), 85, 4)
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
    Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
    Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
    '
    CPAG% = CPAGCLI%((-1) * NPRO%)
    Call REPLA(ENCAO$, Chr$(CPAG%), 153, 1)
    Call REPLA(ENCAO$, Chr$(TIEMBA%), 154, 1)
    '
    Call REPLA(ENCAO$, ATEN$, 163, 30)
    '
125 PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    Call REPLA(ENCAO$, PODESC$, 97, 16)
    '
126 OBX$ = OBJPLA$("OCOMPRA-APRO", ENCAO$)
    If OBX$ = "" Then GoTo 103
    ENCAO$ = OBX$
    TOTANTE# = TOTATO#
    TOTANE# = CVD(Mid$(ENCAO, 89, 8))
    PORDESCUENTO$ = TRIM$(Mid$(ENCAO, 97, 16))
    DESC$ = PORDESCUENTO$
    IDESCUEN# = TOTANE# * COEFDESCU(DESC$) / 100
    TOTANEDES# = TOTANE# - IDESCUEN#
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
    Call REPLA(ENCAO$, MKD$(TOTATO#), 145, 8)
    If Abs(TOTANTE# - TOTATO#) > 0.005 Then
      GoTo 126
    End If
    MONECOI% = Asc(Mid$(ENCAO$, 83, 1))
    MONEMI% = Asc(Mid$(ENCAO$, 84, 1))
    If MONEMI% > 1 Then
      If MONECOI% <> MONEMI% Then
        Call REPLA(ENCAO$, Chr$(MONEMI%), 83, 1)
        GoTo 126
      End If
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
    NUCOXX& = 0
    NUOCONUE& = PRONUOCO&
    If MONUME$ = "MANUAL" Then
140   NUCOXX& = Val(InputBox$("Ingrese Número de Orden de Compra", "Numeración Manual O/Compra", TRIM$(Str$(NUOCONUE&))))
      If NUCOXX& < NUOCONUE& Then GoTo 140
      NUOCONUE& = NUCOXX&
    End If
    '
    CAMBIANUM% = 0
    If NUOCONUE& <> NRO& Then
       NRO& = NUOCONUE&
       Call REPLA(ENCAO$, MKS$(NRO&), 1, 4)
       For J& = 1 To ULTREG&("!OCOMDETA")
         ZZ$ = REGLEIDO$("!OCOMDETA", J&)
         Call REPLA(ZZ$, MKS$(NRO&), 1, 4)
         Call GRAREG("!OCOMDETA", ZZ$, J&)
       Next J&
       Call CIERRARCH("!OCOMDETA")
       CAMBIANUM% = 1
    End If
    '
    RENCA& = 1 + ULTREG&("OCOMENCA")
    PRIDETA& = 1 + ULTREG&("OCOMDETA")
    ULTDETA& = PRIDETA& + ULTREG&("!OCOMDETA") - 1
    '
    ' grabar encabezado
    Call REPLA(ENCAO$, MKS$(PRIDETA&), 155, 4)
    Call REPLA(ENCAO$, MKS$(ULTDETA&), 159, 4)
    Call GRAREG("OCOMENCA", ENCAO$, RENCA&)
    Call CIERRARCH("OCOMENCA")
    '
    EMICONSI% = 0
    For J& = 1 To ULTREG&("!OCOMDETA")
       ZZ$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(ZZ$, MKS$(RENCA&), 125, 4)
       Call REPLA(ZZ$, Chr$(MONECOI%) + Chr$(MONEMI%), 121, 2)
       TICAMB1 = TICAMBIO: If MONEMI% > 1 Then TICAMB1 = 1
       VAUNI$ = FORMATNUM$(CVD(Mid$(ZZ$, 79, 8)) / TICAMB1, "F12.4")
       Call REPLA(VAUNI$, MONECOS$, 1, 2)
       Call REPLA(ZZ$, VAUNI$, 43, 12)
       Call REGAPP("OCOMDETA", ZZ$)
       '
       NOCOM$ = TRIM$(Str$(NRO&))
       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
       ITNU$ = TRIM$(Str$(J&))
       While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
       '
       CI0% = CVI(Mid$(ZZ$, 9, 2))
       CANTUNID = CVS(Mid$(ZZ$, 55, 4)) * CVS(Mid$(ZZ$, 63, 4))
       NPX$ = MKI$(NPRO%)
       Call LEEEXPLO(CODEXT$(CI0%), CANTUNID)
       For I% = 1 To CAIT%
          YY$ = REGBLAN$("MACONSIG")
          Call REPLA(YY$, MKI$(FECHEMI%), 1, 2)
          Call REPLA(YY$, NPX$, 3, 2)
          Call REPLA(YY$, NOCOM$, 21, 12)
          '
          CII% = CIJ%(I%)
          CANTI# = CDbl(CANTUNID * USOI(I%))
          '
          If TRIM$(CODEXT$(CII%)) <> "" Then
            If CANTI# > 0.05 Then
              Call REPLA(YY$, MKI$(CII%), 33, 2)
              Call REPLA(YY$, UNIMED$(CII%), 51, 4)
              Call REPLA(YY$, MKD$(CANTI#), 55, 8)
              Call REGAPP("MACONSIG", YY$)
              EMICONSI% = 1
            End If
          End If
       Next I%
       '
    Next J&
    Call CIERRARCH("OCOMDETA")
    Call CIERRARCH("MACONSIG")
    Call BAJALDISCO
    '
' actualizar costo standard
    HOII% = HOY(HOS$)
    For J& = 1 To ULTREG&("!OCOMDETA")
      ZZ$ = REGLEIDO$("!OCOMDETA", J&)
      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
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
        If PRHBL% = NPRO% Or PRHBL% < 1 Then
            PREUNI2 = PREUNI1# / TICAMBIO
            If PREUNI2 >= 0.000005 Then
              Call REPLA(REMAS$, MKS$(PREUNI2), 111, 4)
              Call REPLA(REMAS$, MKI$(FECHEMI%), 115, 2)
              Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
              Call REPLA(REMAS$, Chr$(MONECOSI%), 95, 1)
              Call GRAREG("MASTER", REMAS$, REMAL&)
            End If
            '
          Else
            '
            TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(CI0%)) + "\el Proveedor Habitual Registrado es\"
            TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PRHBL%)) + "\  \"
            TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
            TTXX$ = TTXX$ + RASOCLI$((-1) * NPRO%)
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
              Call REPLA(ZZX$, "Emis.O/C " + TRIM$(Str$(NRO&)), 33, 32)
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
    Next J&
    '
    jj& = 0
    For U& = 193 To 512 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      jj& = jj& + 1
      Call GRAREG("!OBSERVOF", TTXX$, jj&)
    Next U&
    Call SETULTREG("!OBSERVOF", jj&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = Asc(Mid$(ENCAO$, 153, 1))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = Asc(Mid$(ENCAO$, 154, 1))
    EMBALA$ = TRIM$(ECOARCH$("TAEMBALA", Chr$(TIEM%)))
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
    If CAMBIANUM% = 1 Then
      If NUCOXX& < 1 Then
        Call COMUNI("Número Ocupado.\La O/Compra fue Renumerada como Nro. " + TRIM$(Str$(NRO&)))
      End If
    End If
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
      CODPARAM$(15) = "FECH-REC"
      CODPARAM$(16) = "TIPO-VAL"
      CODPARAM$(17) = "IMPO-CHE"
      CODPARAM$(18) = "REF-CONJ"
      CAPARDOC% = 18
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
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
      FEX = FECHENT%
      DOCPARAM$(15) = FECHATEX$(FEX)
      DOCPARAM$(16) = MONEDEM$
      DOCPARAM$(17) = ""
      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMBIO)
      DOCPARAM$(18) = EMBALA$
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
         CAITAB1% = CAITAB1% + 1
         CPROP$ = TRIM$(CODEXT$(CI0%))
         CODDX$ = CPROP$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         If REVICODX% = 1 Then
            FEREVX = FEREVI%(CI0%)
            CODDX$ = CODDX$ + " (Letra: " + REVICOD$(CI0%) + " - " + FECHATEX$(FEREVX) + ")"
         End If
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = TRIM$(DESCRIT0$(CI0%))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               DECRIT$ = DECRIT$ + " (" + CPROP$ + ")"
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
         UNIST$ = Mid$(X$, 11, 4)
         UNICO$ = UCase$(TRIM$(Mid$(X$, 59, 4)))
         If TRIM$(Left$(UCase$(UNIST$), 2)) <> UNICO$ Then
           DEADI$ = "(equiv. " + TRIM$(FORMATNUM$(CVS(Mid$(X$, 51, 4)), "F12.1")) + " " + ECOARCH$("UNIMED", UNIST$) + ".)"
           CAITAB1% = CAITAB1% + 1
           TETABU1$(2, CAITAB1%) = DEADI$
         End If
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
                   If INTERLI <= 6 Then
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
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "ARCON") > 0 Then
         Call CIERRARCH("*.*")
         Call BAJALDISCO
         ANEXOCOM.Show 1
         GoTo 199
      End If
      '
      For J& = 1 To ULTREG&("!OCOMDETA")
         X$ = REGLEIDO$("!OCOMDETA", J&)
         CI0% = CVI(Mid$(X$, 9, 2))
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
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    If EMICONSI% = 1 Then
      If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
        Call ENTRECONSI(NPRO%)
      End If
    End If
    Call CIERRARCH("*.*")
    '
    GoTo 100
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'

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
    Screen.MousePointer = 11
    OCOA& = 0: VOCO$ = String$(2048, 0): POIN& = 1
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      If Asc(Mid$(X$, 124, 1)) < 9 Then ' esta abierta y pendiente
        'If CVS(Mid$(X$, 99, 4)) < CVS(Mid$(X$, 55, 4)) - 0.05 Then
        NOCO& = CVS(Left$(X$, 4))
        If NOCO& <> OCOA& Then
          OCOA& = NOCO&
          Mid$(VOCO$, POIN&, 4) = MKS$(NOCO&)
          POIN& = POIN& + 4
          If POIN& > Len(VOCO$) Then VOCO$ = VOCO$ + String$(2048, 0)
        End If
      End If
    Next U&
    VOCO$ = Left$(VOCO$, POIN& - 1)
    
    J& = 0: K& = 0: OCOA& = 0
    '
    FIN& = ULTREG&("OCOMENCA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMENCA", U&)
      If Asc(Mid$(X$, 54, 1)) < 1 Then ' esta abierta y pendiente
        NOCO& = CVS(Left$(X$, 4))
        'For KKL& = 1 To Len(VOCO$) Step 4
        '  If CVS(Mid$(VOCO$, KKL&, 4)) = NOCO& Then GoTo 5
        'Next KKL&
        'GoTo 9
        '
        PINI& = 1
3       POSI& = InStr(PINI&, VOCO$, MKS$(NOCO&))
        If POSI& < 1 Then GoTo 9                           ' NEXT RGX&
        If (POSI& Mod 4) <> 1 Then
           PINI& = POSI& + 1: GoTo 3
        End If
        '
5       Y$ = Left$(X$, 48) + MKS$(U&)
        J& = J& + 1
        Call GRAREG("INDIOCOM", Y$, J&)
      End If
9   Next U&
    Call SETULTREG("INDIOCOM", J&)
    Call CIERRARCH("INDIOCOM")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Modificarse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    '
10  YY$ = REGSEL$("INDIOCOM/Indice de Ordenes de Compra Pendientes para Modificacion")
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
100 FECHEMI% = CVI(Mid$(XX$, 49, 2))
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
    For U& = 1 To FIN&
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
        End If
      End If
    Next U&
    Call CIERRARCH("OCOMDETA")
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
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANE# + TOTIVA#
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
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
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
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        If PVCL% <> 5 Then
          TOTIVA# = (Int(ALIVA * TOTANEDES# + 0.5)) / 100
        End If
      End If
    End If
    TOTATO# = TOTANEDES# + TOTIVA#
    Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
    Call REPLA(ENCAO$, MKD$(IDESCUEN#), 113, 8)
    Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
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
        If PRHBL% = NPRO% Or PRHBL% < 1 Then
            PREUNI2 = PREUNI1# / TICAMBIO
            If PREUNI2 >= 0.000005 Then
              Call REPLA(REMAS$, MKS$(PREUNI2), 111, 4)
              Call REPLA(REMAS$, MKI$(FECHEMI%), 115, 2)
              Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
              Call REPLA(REMAS$, Chr$(MONECOSI%), 95, 1)
              Call GRAREG("MASTER", REMAS$, REMAL&)
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
    jj& = 0
    For U& = 193 To 512 Step 64
      TTXX$ = Mid$(ENCAO$, U&, 64)
      jj& = jj& + 1
      Call GRAREG("!OBSERVOF", TTXX$, jj&)
    Next U&
    Call SETULTREG("!OBSERVOF", jj&)
    Call CIERRARCH("!OBSERVOF")
    '
    CPAG% = Asc(Mid$(ENCAO$, 153, 1))
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    TIEM% = Asc(Mid$(ENCAO$, 154, 1))
    EMBALA$ = TRIM$(ECOARCH$("TAEMBALA", Chr$(TIEM%)))
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
      CAPARDOC% = 17
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
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
         CPROP$ = TRIM$(CODEXT$(CI0%))
         CODDX$ = CPROP$
         If REPLAORI% = 1 Then
            CODDX$ = TRIM$(CODORIG$(CI0%, NC1%))
         End If
         '
         If REVICODX% = 1 Then
            FEREVX = FEREVI%(CI0%)
            CODDX$ = CODDX$ + " (Letra: " + REVICOD$(CI0%) + " - " + FECHATEX$(FEREVX) + ")"
         End If
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODDX$
           DECRIT$ = TRIM$(DESCRIT0$(CI0%))
           If REPLAORI% = 1 Then
             If CODDX$ <> CPROP$ Then
               DECRIT$ = DECRIT$ + " (" + CPROP$ + ")"
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
                   If INTERLI <= 6 Then
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
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
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
999 Call CIERRARCH("*.*")
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
        If Asc(Mid$(X$, 124, 1)) < 9 Then ' esta abierta y pendiente
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
    FIN& = ULTREG&("OCOMDETA")
    For J& = 1 To FIN&
      Call TRACE(20, J&, FIN&)
      XX$ = REGLEIDO$("OCOMDETA", J&)
      If CVS(Left$(XX$, 4)) = NUOCO& Then
        If CVS(Mid$(XX$, 99, 4)) < CVS(Mid$(XX$, 55, 4)) - 0.05 Then
          Call REPLA(XX$, Chr$(9), 124, 1)
          Call GRAREG("OCOMDETA", XX$, J&)
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
    ' CANCELAR RESERVAS
    '
    jj& = 0
    For J& = 1 To ULTREG&("MACONSIG")
      XX$ = REGLEIDO$("MACONSIG", J&)
      If Val(Mid$(XX$, 24, 6)) <> NUOCO& Then
           jj& = jj& + 1
           Call GRAREG("MACONSIG", XX$, jj&)
         Else
           CAENTRE# = CVD(Mid$(XX$, 63, 8))
           CADEVUE# = CVD(Mid$(XX$, 111, 8))
           If Abs(CAENTRE# - CADEVUE#) >= 0.05 Then
             jj& = jj& + 1
             Call GRAREG("MACONSIG", XX$, jj&)
           End If
      End If
    Next J&
    Call SETULTREG("MACONSIG", jj&)
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
   ' precio ultima compra
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
   jj& = 0
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
           jj& = jj& + 1
           Call GRAREG("COPRECOM", YYY$, jj&)
         End If
       End If
     End If
     '
   Next U&
   '
   Call SETULTREG("COPRECOM", jj&)
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
          FEX = FECHEMI%
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
          Call REPLA(XX$, Chr$(0), 124, 1)
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
   Call FINAL("")
End Sub
