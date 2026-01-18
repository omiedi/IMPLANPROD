VERSION 5.00
Begin VB.Form LEARTRA07 
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Archivos"
   ClientHeight    =   5964
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10128
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5964
   ScaleWidth      =   10128
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Operación Logistica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4830
      TabIndex        =   47
      Top             =   105
      Width           =   4200
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   600
         ScaleWidth      =   3672
         TabIndex        =   48
         Top             =   450
         Width           =   3720
         Begin VB.CommandButton Command17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2010
            Picture         =   "LEARTRA07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Ajuste Manual Lista de Picking"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Procesar Requerimientos y Generar Lista de Picking"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2560
            Picture         =   "LEARTRA07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Emitir Lista de Picking"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3120
            Picture         =   "LEARTRA07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Anular Lista de Picking"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Procesar los Requerimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   615
            TabIndex        =   52
            Top             =   105
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Listados de Existencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2120
      Left            =   4830
      TabIndex        =   37
      Top             =   1470
      Width           =   4215
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   38
         Top             =   1320
         Width           =   2535
         Begin VB.CommandButton Command6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Posición de Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   735
            TabIndex        =   40
            Top             =   105
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   41
         Top             =   870
         Width           =   2535
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   -20
            Picture         =   "LEARTRA07.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales por Có- digo de Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   43
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   44
         Top             =   420
         Width           =   2535
         Begin VB.CommandButton Command9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":0EBC
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Código de Producto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   46
            Top             =   0
            Width           =   1200
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Consultas Complementarias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   4830
      TabIndex        =   27
      Top             =   3675
      Width           =   4215
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   28
         Top             =   1380
         Width           =   2535
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":11C6
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Lista- dos  de Movimientos "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   630
            TabIndex        =   30
            Top             =   105
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   31
         Top             =   900
         Width           =   2535
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   -20
            Picture         =   "LEARTRA07.frx":1310
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Inventario Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   33
            Top             =   30
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   600
         ScaleWidth      =   2484
         TabIndex        =   34
         Top             =   420
         Width           =   2535
         Begin VB.CommandButton Command15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":161A
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Kardex  de Movimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   36
            Top             =   30
            Width           =   1410
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Movimientos Internos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   210
      TabIndex        =   14
      Top             =   3045
      Width           =   4515
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   600
         ScaleWidth      =   2388
         TabIndex        =   24
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":1924
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Operaciones de Re-Envasado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones de Re-Envasado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   26
            Top             =   105
            Width           =   1620
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   600
         ScaleWidth      =   2388
         TabIndex        =   21
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":1C2E
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Ajuste Directo de Existencias por Lote"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Exis- tencias por Lote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   23
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   600
         ScaleWidth      =   2388
         TabIndex        =   15
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":1D78
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Fraccionamiento de Lote / Partida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fraccionamiento de Lote / Partida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   17
            Top             =   50
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   600
         ScaleWidth      =   2388
         TabIndex        =   18
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":1EC2
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Cambio de Ubicación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio de Ubicación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   20
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   45
            Width           =   1110
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   7000
      Left            =   9240
      ScaleHeight     =   6948
      ScaleWidth      =   1044
      TabIndex        =   6
      Top             =   -120
      Width           =   1095
      Begin VB.CommandButton Command12 
         Caption         =   "Ubic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   105
         Picture         =   "LEARTRA07.frx":200C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Tabla de Ubicaciones Predefinidas"
         Top             =   1680
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   84
         ScaleWidth      =   600
         TabIndex        =   9
         Top             =   5200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "LEARTRA07.frx":2316
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "LEARTRA07.frx":2620
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   312
         Left            =   -312
         Picture         =   "LEARTRA07.frx":276A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5496
         Width           =   1212
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1050
      ScaleHeight     =   600
      ScaleWidth      =   2772
      TabIndex        =   0
      Top             =   600
      Width           =   2820
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2205
         Picture         =   "LEARTRA07.frx":468C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Listados de Transferencias"
         Top             =   0
         Width           =   560
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   0
         Picture         =   "LEARTRA07.frx":4996
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Procesamiento de Tablas de Transferencia"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   1
         Top             =   105
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Tablas de Transferencia de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   210
      TabIndex        =   4
      Top             =   120
      Width           =   4515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Ingresos a Depósito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   210
      TabIndex        =   5
      Top             =   1680
      Width           =   4515
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   600
         ScaleWidth      =   2388
         TabIndex        =   11
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LEARTRA07.frx":4CA0
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Procesamiento de Ingresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   12
            Top             =   75
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "LEARTRA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    TTXX$ = "Este Proceso Revisa las Tablas de Transferencia de Datos"
    TTXX$ = TTXX$ + "\Seleccionando Nuevos Registros Generados por el S.A.G.S."
    TTXX$ = TTXX$ + "\   \A Partir de Estas Novedades Actualiza el Maestro de Artículos FLEXOFT"
    TTXX$ = TTXX$ + "\Creando Nuevos Códigos y Actualizando Descripción de los Existentes."
    TTXX$ = TTXX$ + "\   \También Lee los Requerimientos y Consumos de la Tabla TRANMOVI, Actuali-"
    TTXX$ = TTXX$ + "\zando las Tablas de Balance de Requerimientos, Movimientos, Existencias"
    TTXX$ = TTXX$ + "\y Saldos por Lote en la Base de Datos FLEXOFT."
    '
    If COMALTER%(TTXX$ + "\\Continuar\Cancelar") <> 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    Call ABREMASTER
    Call ABRESTOCKS
    MAXCART% = CONTROL$("MASTER", "LOMAXCOD")
    If MAXCART% < 15 Then MAXCART% = 15
    '
    SQLX$ = "DELETE MASTER WHERE CODINT < 1"
     FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    '
    'RECORRO LA TABLA TRANMOVI
    CONDI$ = "STATUS < 2 and TipoTrans <> 2 "
    FIN& = CantRegistros("TRANMOVI", CONDI$)
    '
    SQLX$ = "SELECT * FROM TRANMOVI "
    SQLX$ = SQLX$ + " WHERE " + CONDI$
    Dim TRANTMP As ADODB.Recordset
    Set TRANTMP = New ADODB.Recordset
    TRANTMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    JJ& = 0
    With TRANTMP
      Do While Not .EOF
         JJ& = JJ& + 1
         Call TRACE(20, JJ&, FIN&)
         CODD$ = SAFETEXT$(TRIM$(!Cod_Exte))
         If Len(CODD$) > MAXCART% Then
           Call MENSERR(24, "Código " & CODD$ & "Supera el Límite de   " & MAXCART% & "  Caracteres")
           GoTo 50
         End If
 
         DESS$ = SAFETEXT$(!Descripcion)
         DES22$ = SAFETEXT$(!DESCRIDOS)
         'BUSCO EL CODIGO DE TRANMOVI SI ESTA EN EL MASTER
         'NOTA :EN SQL LA CONSULTA NO CONTEMPLA NI LOS ESPACIOS
         'EXTENOS DEL STRING NI LAS MAYUSCULAS O MINUSCULAS
         '
         CIXI% = CODINT%(CODD$)
         'SI EXISTE HAGO UN UPDATE EN EL MASTER
         If CIXI% > 0 Then
              SQLX$ = "UPDATE MASTER SET DESCRIPCION = '" & DESS$ & "', STATUS = 1"
              SQLX$ = SQLX$ + " WHERE Codigo ='" & CODD$ & "'"
               FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
           Else
              'CONSIGO EL VALOR MAXIMO DEL CODIGO INTERNO
              SQLXY$ = "SELECT MAX(Codint)AS NUMAXX FROM MASTER"
              Dim NUMAXTMP As ADODB.Recordset
              Set NUMAXTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLXY$))
              CIXI% = 1 + XVALO(NUMAXTMP!NUMAXX)
              '
              'AGREGO EL NUEVO REGISTRO CON EL SIGUIENTE DE CODINT
              SQLR$ = "INSERT INTO MASTER (CODIGO,  Codint, "
              SQLR$ = SQLR$ + " Descripcion, Descridos, STATUS)VALUES('" & CODD$ & "' ,"
              SQLR$ = SQLR$ + "" & CIXI% & ","
              SQLR$ = SQLR$ + "'" & DESS$ & "',"
              SQLR$ = SQLR$ + " '" & DES22$ & "',"
              SQLR$ = SQLR$ + "1 )"
               FLEXCONN.EXECUTE (FILTSQL$ (SQLR$))
         End If
         '
         DOCUI$ = TRIM$(SAFETEXT$(!CodiCompro))
         NUCO$ = TRIM$(Str$(XVALO(!NumeCompro))): While Len(NUCO$) < 6: NUCO$ = "0" + NUCO$: Wend
         NOSU$ = TRIM$(Str$(XVALO(!CodSucu))): While Len(NOSU$) < 3: NOSU$ = "0" + NOSU$: Wend
         NORI$ = TRIM$(Str$(XVALO(!NuOrItem))): While Len(NORI$) < 3: NORI$ = "0" + NORI$: Wend
         DOCING$ = DOCUI$ + "-" + NUCO$ + "-" + NOSU$ + "-" + NORI$
         '
         TIPOMOVI% = XVALO(!TipoTrans)
         IDLOT$ = SAFETEXT$(!IdenLote)
         CONDI$ = "IdenReque = '" & DOCING$ & "' AND Cod_Inte = " & CIXI%
         CONDI$ = CONDI$ & " AND IdenLote = '" & IDLOT$ & "'"
           If TIPOMOVI% = 1 Then CONDI$ = CONDI$ & " AND KILO_REQUE > 0 "
           If TIPOMOVI% = 3 Then CONDI$ = CONDI$ & " AND KILO_CONSU > 0 "
         DOCUYAI$ = SAFETEXT$(VALOBADA("BALAREQMO", "IdenReque", CONDI$))
         If DOCUYAI$ <> "" Then
            !Ingresado = 9
            GoTo 40
         End If
         '
         IDLOFLEX$ = SAFETEXT$(!Ident_Asig)
         If TRIM$(IDLOFLEX$) = "" Then IDLOFLEX$ = IDLOT$
         If TIPOMOVI% = 1 Then IDLOFLEX$ = IDLOT$
         '
         If TIPOMOVI% = 3 Then         ' VERIFICAR QUE NO SUPERE EL SALDO FISICO
           CONDI$ = "Cod_Inte = " & CIXI% & " "
           CONDI$ = CONDI$ + " AND (IdenLote = '" & IDLOFLEX$ & "')"
           CANORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
           CANCONS = XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDI$))
           QSALDO = CANORI - CANCONS
           If XVALO(!CanKilos) > QSALDO + 0.0005 Then
              Call MENSERR(24, "Imposible Procesar '" + DOCING$ + "'.\Kilos Exceden Saldo de Lote '" + IDLOFLEX$ + "'.")
              GoTo 50
           End If
         End If
         '
         ' CARGO EL BALAREQMO
         Dim BALARE As ADODB.Recordset
         Set BALARE = New ADODB.Recordset
         SQLX$ = "SELECT * FROM BALAREQMO WHERE IDENREQUE = '" & DOCING$ & "'"
         BALARE.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         '
         If TIPOMOVI% = 3 Then   ' cancela la reserva
             BALARE.AddNew   ' YA CONVERTIDO
                   ON ERROR RESUME NEXT
             BALARE!CodiEmpr = Codiemp%
                   ON ERROR GOTO 0
             BALARE!CodiCompro = !CodiCompro
             BALARE!NumeCompro = !NumeCompro
             BALARE!CodSucu = !CodSucu
             BALARE!NuOrItem = !NuOrItem
             BALARE!IdenReque = DOCING$
             BALARE!Cod_Inte = CIXI%
             BALARE!IdenLote = IDLOFLEX$
             BALARE!Cantidad = !Cantidad
             BALARE!CodEnvase = !CodEnvase
             BALARE!KILO_REQUE = 0
             BALARE!Kilo_Asig = (-1) * XVALO(!CanKilos)
             BALARE!Kilo_Consu = 0
             BALARE!FeReMovi = Now
             BALARE!FechMov = !FechMov
             BALARE!MarqAsi = 0
             BALARE!CreadoPor = SAFETEXT$(!CreadoPor)
             BALARE!Destino = SAFETEXT$(!Destino)
             BALARE!Ident_Asig = SAFETEXT$(!Ident_Asig)
             BALARE.Update
         End If
         '
         BALARE.AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
         BALARE!CodiEmpr = Codiemp%
               ON ERROR GOTO 0
         BALARE!CodiCompro = !CodiCompro
         BALARE!NumeCompro = !NumeCompro
         BALARE!CodSucu = !CodSucu
         BALARE!NuOrItem = !NuOrItem
         BALARE!IdenReque = DOCING$
         BALARE!Cod_Inte = CIXI%
         BALARE!IdenLote = IDLOFLEX$
         BALARE!Cantidad = !Cantidad
         BALARE!CodEnvase = !CodEnvase
         BALARE!KILO_REQUE = 0
         BALARE!Kilo_Asig = 0
         BALARE!Kilo_Consu = 0
           If TIPOMOVI% = 1 Then BALARE!KILO_REQUE = XVALO(!CanKilos)
           If TIPOMOVI% = 3 Then BALARE!Kilo_Consu = XVALO(!CanKilos)
         BALARE!FeReMovi = Now
         BALARE!FechMov = !FechMov
         BALARE!MarqAsi = 0
         BALARE!CreadoPor = SAFETEXT$(!CreadoPor)
         BALARE!Destino = SAFETEXT$(!Destino)
         BALARE.Update
         BALARE.Close
         '
         If TIPOMOVI% = 3 Then      ' ES UN CONSUMO EFECTIVO
            FECHA% = CVINT(!FechMov)
            DESTIX$ = SAFETEXT$(!Destino)
            Call MOVISTOK(FECHA%, CIXI%, IDLOFLEX$, "CF", DOCING$, DOCUI$ + "-" + NUCO$, DOCING$, DOCUI$ + "-" + NUCO$, DESTIX$, NORII%, VUNI#, MONEI%, NC%, DEPO%, (-1) * XVALO(!CanKilos), 0, 0)
            Call ACSALOTE(CIXI%, IDLOT$, "NOMESS")
         End If
        '
' si es un consumo entonces
' a) grabar movimiento
' b) cancelar reserva
' c) ejecutar acsalote
         '
36       !Ingresado = 1
40       'PONGO A 2 EL STATUS
         !Status = 2
         .Update
         '
50    .MoveNext
      Loop
    TRANTMP.Close
    Set TRANTMP = Nothing
    
    End With
    '
    '////AHORA LO MISMO PERO CON LA TABLA TRANRECEP
    '
    CONDI$ = "STATUS < 2"
    FIN& = CantRegistros("TRANRECEP", CONDI$)
    '
    SQLX$ = "SELECT * FROM TRANRECEP WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " + CONDI$
    Dim TRANRETMP As ADODB.Recordset
    Set TRANRETMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    JJ& = 0
    With TRANRETMP
      'RECORRO TRANRECEP
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        CODD$ = TRIM$(SAFETEXT$(!Cod_Exte))
        If Len(CODD$) > MAXCART% Then   'MAXIMO DE LONGITUD CONFIGURADA DE CODIGO
           Call MENSERR(24, "Código " & CODD$ & "Supera el Límite de   " & MAXCART% & "  Caracteres")
           GoTo 60
        End If
        
        DESS$ = SAFETEXT$(!Descripcion)
        'BUSCO EL CODIGO DE TRANRECEP EN EL MASTER
        CIXI% = CODINT%(CODD$)
        'SI EXISTE HAGO UN UPDATE EN EL MASTER
        If CIXI% > 0 Then
            SQLX$ = "UPDATE MASTER SET DESCRIPCION = '" & DESS$ & "'"
            SQLX$ = SQLX$ + " WHERE Codigo ='" & CODD$ & "'"
             FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
         Else
            'CONSIGO EL VALOR MAXIMO DEL CODIGO INTERNO
            SQLXY$ = "SELECT MAX(Codint)AS NUMAXX FROM MASTER"
            Set NUMAXTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLXY$))
            CIXI% = 1 + XVALO(NUMAXTMP!NUMAXX)
            '
            ' LO INGRESO EN EL MASTER CON EL VALOR MAXIMO + 1
            SQLR$ = "INSERT INTO MASTER (CODIGO,  Codint, "
            SQLR$ = SQLR$ + " Descripcion, STATUS) VALUES ('" & CODD$ & "' ,"
            SQLR$ = SQLR$ + "" & CIXI% & ","
            SQLR$ = SQLR$ + "'" & DESS$ & "',"
            SQLR$ = SQLR$ + "1 )"
             FLEXCONN.EXECUTE (FILTSQL$ (SQLR$))
            '
        End If
        '
        ' PONGO EN 2 EL STATUS
        SQLZ$ = "UPDATE TRANRECEP SET STATUS = 2"
        SQLZ$ = SQLZ$ + " WHERE Cod_Exte ='" & CODD$ & "'"
         FLEXCONN.EXECUTE (FILTSQL$ (SQLZ$))
        '
60      .MoveNext
      Loop
    End With
    TRANRETMP.Close
    Set TRANRETMP = Nothing
    
    UPDATMASTER% = 1
    Call GENECOMAS
    
    Call MENSERR(24, "Proceso Terminado")
    
99  Command1.Enabled = True
    '\\\OT-08-1044-14/11/08///
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call ABRESTOCKS
    '
    NULIPI& = 1 + XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
    DELIPI$ = SAFETEXT$(Now)
    HUBOASIG% = 0
    '
    Dim REQUERI As ADODB.Recordset
    Set REQUERI = New ADODB.Recordset
    '
    SQLX$ = "SELECT COD_INTE, IdenReque, SUM(KILO_REQUE) AS TOREQUE, SUM(KILO_ASIG) AS TOASIG FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE (MARQASI < 1 OR MARQASI IS NULL) "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, IdenReque "
    Set REQUERI =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    '
    HOII% = HOY(HOS$)
    IDASI$ = Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2) + "-" + TRIM$(Str$(NULIPI&)) + "-"
    NORIT% = 0
    '
    With REQUERI
       If Not (.EOF And .BOF) Then
          Do While Not .EOF
             HUBOMOASI% = 0
             If XVALO(!TOASIG) >= XVALO(!TOREQUE) - 0.05 Then   ' CUBIERTA LA ASIGNACION, PONE MARCA
                  IDREQ$ = SAFETEXT$(!IdenReque)             ' ASIGNA CUMPLIDA Y VA AL SIGUIENTE
                  CODIN% = XVALO(!Cod_Inte)
                  SQLY$ = "UPDATE BALAREQMO "
                  SQLY$ = SQLY$ + " SET MARQASI = 1 "
                  SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "'"
                  FLEXCONN.Execute SQLY$
                  GoTo 19                         ' MOVENEXT
                Else
                  CODIMAT% = XVALO(!Cod_Inte)             ' CODIGO PRODUCTO REQUERIDO
                  KILPENASI = XVALO(!TOREQUE) - XVALO(!TOASIG)   ' KILOS PENDIENTES DE ASIGNACION
                  IDREQ$ = SAFETEXT$(!IdenReque)              ' ASIGNA CUMPLIDA Y VA AL SIGUIENTE
                  CONDX$ = "Cod_Inte = " & CODIMAT% & " "
                  CONDX$ = CONDX$ + "AND IdenReque = '" & IDREQ$ & "' "
                  CONDX$ = CONDX$ & "AND KILO_REQUE > 0 "
                  LOPEDIDO$ = VALOBADA("BALAREQMO", "IDENLOTE", CONDX$) ' PARTIDA SOLICITADA
                  If TRIM$(LOPEDIDO$) = "" Then GoTo 25 ' PEDIDO SIN LOTE PREFERIDO
                  '
                  SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
                  SQLY$ = SQLY$ + "WHERE Cod_Inte = " & CODIMAT% & " "
                  SQLY$ = SQLY$ + "AND LoteCol_Prov = '" & LOPEDIDO$ & "' "
                  SQLY$ = SQLY$ + "AND (CanAlta - CanConsu - CanReser) > 0 "
                  SQLY$ = SQLY$ + "ORDER BY CanSaldo ASC "
                  Dim LODISPO As ADODB.Recordset
                  Set LODISPO =  FLEXCONN.EXECUTE (FILTSQL$ (SQLY$))
                  '
                  With LODISPO
                     If Not (.EOF And .BOF) Then
                       Do While Not .EOF
                         If KILPENASI < 0.05 Then GoTo 29
                         IDLOTX$ = SAFETEXT$(!IdenLote)
                         CIXI% = CODIMAT%
                         CONDI$ = "Cod_Inte = " & CIXI% & " "
                         CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "' "
                         KILDISLO = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
                         KILAASI = KILDISLO: If KILPENASI < KILDISLO Then KILAASI = KILPENASI
                         '
                         POSILOTE$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
                         CODENVAS$ = VALOBADA("LOTINGRE", "CodEnvase", CONDI$)
                         COEFI = 1
                         KILORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
                         CANORI = XVALO(VALOBADA("LOTINGRE", "Cantidad", CONDI$))
                           If KILORI > 0 Then COEFI = CANORI / KILORI
                         '
                         ' GRABAR LA ASIGNACION EN BALAREQMO
                         CCX1$ = TRIM$(VALOBADA("BALAREQMO", "CodiCompro", CONDX$))
                         NCX1 = XVALO(VALOBADA("BALAREQMO", "NumeCompro", CONDX$))
                         CSX1 = XVALO(VALOBADA("BALAREQMO", "CodSucu", CONDX$))
                         NOI1 = XVALO(VALOBADA("BALAREQMO", "NuOrItem", CONDX$))
                         DESTIPEDIDO$ = VALOBADA("BALAREQMO", "Destino", CONDX$)
                        
                         If KILAASI >= 0.05 Then
                            SQLZ$ = "SELECT * FROM BALAREQMO "
                            SQLZ$ = SQLZ$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            Dim ASIGNA As ADODB.Recordset
                            Set ASIGNA = New ADODB.Recordset
                            ASIGNA.OPEN FILTSQL$ (SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                            With ASIGNA
                               .AddNew   ' YA CONVERTIDO
                               ON ERROR RESUME NEXT
                               !CodiEmpr = Codiemp%
                               ON ERROR GOTO 0
                               !CodiCompro = CCX1$
                               !NumeCompro = NCX1
                               !CodSucu = CSX1
                               !NuOrItem = NOI1
                               !IdenReque = IDREQ$
                               !Cod_Inte = CIXI%
                               !IdenLote = IDLOTX$
                               !Posicion = POSILOTE$
                               !Cantidad = KILAASI * COEFI         ' falta dederminar
                               !CodEnvase = CODENVAS$
                               !KILO_REQUE = 0
                               !Kilo_Asig = KILAASI
                               !NuLisPick = NULIPI&
                               !DeLisPick = DELIPI$
                               !Kilo_Consu = 0
                               !FeReMovi = Now
                               !FechMov = Now
                               !MarqAsi = 0
                               !CreadoPor = "FLEX"
                               !Destino = DESTIPEDIDO$
                                   NORIT% = NORIT% + 1
                               !Ident_Asig = IDASI$ + TRIM$(Str$(NORIT%))
                               .Update
                               HUBOASIG% = 1
                               HUBOMOASI% = 1
                            End With
                            ASIGNA.Close
                            Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS") ' ACTUALIZAR CON ACSALOTE
                            '
                            SQLY$ = "UPDATE BALAREQMO "
                            SQLY$ = SQLY$ + " SET NuLisPick = " & NULIPI& & " "
                            SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            SQLY$ = SQLY$ + " AND Cod_Inte = " & CIXI% & " "
                            SQLY$ = SQLY$ + " AND Kilo_Reque > 0 "
                            FLEXCONN.Execute SQLY$
                            '
                            KILPENASI = KILPENASI - KILAASI
                            If KILPENASI < 0.05 Then GoTo 29
                         End If
                         '
                       .MoveNext
                       Loop
                     End If
                  End With
                  LODISPO.Close
                  Set LODISPO = Nothing
                  '
                  '
25                SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
                  SQLY$ = SQLY$ + "WHERE Cod_Inte = " & CODIMAT% & " "
                  SQLY$ = SQLY$ + "AND Libera = 'APRO' "
                  SQLY$ = SQLY$ + "AND (CanAlta - CanConsu - CanReser) > 0 "
                  SQLY$ = SQLY$ + "ORDER BY FechVenc ASC "   ' POR CRITERIO FEFO
                  Set LODISPO =  FLEXCONN.EXECUTE (FILTSQL$ (SQLY$))
                  '
                  With LODISPO
                     If Not (.EOF And .BOF) Then
                       Do While Not .EOF
                         If KILPENASI < 0.05 Then GoTo 29
                         IDLOTX$ = SAFETEXT$(!IdenLote)
                         CIXI% = CODIMAT%
                         CONDI$ = "Cod_Inte = " & CIXI% & " "
                         CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "' "
                         KILDISLO = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
                         KILAASI = KILDISLO: If KILPENASI < KILDISLO Then KILAASI = KILPENASI
                         POSILOTE$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
                         CODENVAS$ = VALOBADA("LOTINGRE", "CodEnvase", CONDI$)
                         COEFI = 1
                         KILORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
                         CANORI = XVALO(VALOBADA("LOTINGRE", "Cantidad", CONDI$))
                           If KILORI > 0 Then COEFI = CANORI / KILORI
                           '
                         ' GRABAR LA ASIGNACION EN BALAREQMO
                         CCX1$ = TRIM$(VALOBADA("BALAREQMO", "CodiCompro", CONDX$))
                         NCX1 = XVALO(VALOBADA("BALAREQMO", "NumeCompro", CONDX$))
                         CSX1 = XVALO(VALOBADA("BALAREQMO", "CodSucu", CONDX$))
                         NOI1 = XVALO(VALOBADA("BALAREQMO", "NuOrItem", CONDX$))
                         DESTIPEDIDO$ = VALOBADA("BALAREQMO", "Destino", CONDX$)
                        
                         If KILAASI >= 0.05 Then
                            SQLZ$ = "SELECT * FROM BALAREQMO "
                            SQLZ$ = SQLZ$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            Set ASIGNA = New ADODB.Recordset
                            ASIGNA.OPEN FILTSQL$ (SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                            With ASIGNA
                               .AddNew   ' YA CONVERTIDO
                               ON ERROR RESUME NEXT
                               !CodiEmpr = Codiemp%
                               ON ERROR GOTO 0
                               !CodiCompro = CCX1$
                               !NumeCompro = NCX1
                               !CodSucu = CSX1
                               !NuOrItem = NOI1
                               !IdenReque = IDREQ$
                               !Cod_Inte = CIXI%
                               !IdenLote = IDLOTX$
                               !Posicion = POSILOTE$
                               !Cantidad = KILAASI * COEFI
                               !CodEnvase = CODENVAS$
                               !KILO_REQUE = 0
                               !Kilo_Asig = KILAASI
                               !NuLisPick = NULIPI&
                               !DeLisPick = DELIPI$
                               !Kilo_Consu = 0
                               !FeReMovi = Now
                               !FechMov = Now
                               !MarqAsi = 0
                               !CreadoPor = "FLEX"
                               !Destino = DESTIPEDIDO$
                                   NORIT% = NORIT% + 1
                               !Ident_Asig = IDASI$ + TRIM$(Str$(NORIT%))
                               .Update
                               HUBOASIG% = 1
                               HUBOMOASI% = 1
                            End With
                            '
                            ASIGNA.Close
                            Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS") ' ACTUALIZAR CON ACSALOTE
                            '
                            SQLY$ = "UPDATE BALAREQMO "
                            SQLY$ = SQLY$ + " SET NuLisPick = " & NULIPI& & " "
                            SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            SQLY$ = SQLY$ + " AND Cod_Inte = " & CIXI% & " "
                            SQLY$ = SQLY$ + " AND Kilo_Reque > 0 "
                            FLEXCONN.Execute SQLY$
                            '
                            KILPENASI = KILPENASI - KILAASI
                            If KILPENASI < 0.05 Then GoTo 29
                         End If
                         '
                       .MoveNext
                       Loop
                     End If
                  End With
                  LODISPO.Close
                  Set LODISPO = Nothing
                  '
29           End If
             '
             If HUBOMOASI% = 1 Then
               CONDI$ = "COD_INTE = " & CODIMAT% & " AND IdenReque = '" & IDREQ$ & "'"
               Call ACTUREGISTRO("BALAREQMO", "NulisPick", CONDI$, NULIPI&, REGAFE&, "NOMESS")
               Call ACTUREGISTRO("BALAREQMO", "DelisPick", CONDI$, DELIPI$, REGAFE&, "NOMESS")
             End If
             '
19        .MoveNext
          Loop
       End If
    End With
    '
    If HUBOASIG% > 0 Then
       Call FILTERPUTRECORD("INDIPICK", 1, MKS$(NULIPI&), MKS$(NULIPI&) + DELIPI$)
       Call FILESORT("INDIPICK", "", 1, 1, "ASC")
       Call FRESHECHO("INDIPICK")
       '
       Call ACTRANMOVI(NULIPI&)
    End If
    '
    Call FRESHECHO("INDIPICK")
    Call COMUNI("Proceso de Cálculo Completo.")
    Command10.Enabled = True
End Sub
'
Sub ACTRANMOVI(NULIPI&)
    
   Call BORRATRANMO(NULIPI&)
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = "SELECT * FROM BALAREQMO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where NuLisPick = " & NULIPI&
   SQLX$ = SQLX$ + " AND Kilo_Asig > 0"
   Set BALARETMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   '
   SQLX$ = "SELECT * FROM TRANMOVI "
   SQLX$ = SQLX$ + " WHERE TIPOTRANS = 2 AND INGRESADO = -" & NULIPI&
   Dim TRANTMP As ADODB.Recordset
   Set TRANTMP = New ADODB.Recordset
   TRANTMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   JJ& = 0: NORIT% = 0
   With BALARETMP
      If Not (.BOF And .EOF) Then
         Do While Not .EOF
            CODD% = XVALO(!Cod_Inte)
            KILOS = XVALO(!Kilo_Asig)
            LOTEE$ = SAFETEXT(!IdenLote)
            CANTU = XVALO(!Cantidad)
            ENVU$ = SAFETEXT$(!CodEnvase)
            FEREM% = CVINT(!FeReMovi)
            FEMOV% = CVINT(!FechMov)
            '
            JJ& = JJ& + 1
            XX$ = REGBLAN$("LISTAPIK")
            Call REPLA(XX$, MKI(CODD%), 1, 2)
            Call REPLA(XX$, MKS$(KILOS), 3, 4)
            Call REPLA(XX$, LOTEE$, 7, 16)
            
            COMPRO$ = SAFETEXT(!CodiCompro)
            NUMCOMPR& = XVALO(!NumeCompro)
            CODDSUC& = XVALO(!CodSucu)
            NUITEM& = XVALO(!NuOrItem)
            IDREQ$ = SAFETEXT(!IdenReque)
            DESTIPED$ = SAFETEXT(!Destino)
            IDASIX$ = SAFETEXT$(!Ident_Asig)
            '
            CONDIX$ = "IdenLote = '" & LOTEE$ & "' AND Cod_Inte = " & CODD% & " "
            LOTORI$ = VALOBADA("LOTINGRE", "LOTECOL_PROV", CONDIX$)
            If TRIM$(LOTORI$) = "" Then LOTORI$ = LOTEE$
            '
            With TRANTMP
               .AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
               !CodiEmpr = Codiemp%
               ON ERROR GOTO 0
               !CodiCompro = COMPRO$
               !NumeCompro = NUMCOMPR&
               !CodSucu = CODDSUC&
               !NuOrItem = NUITEM&
               !Cod_Exte = Left$(CODEXT$(CODD%), 24)
               !Descripcion = Left$(DESCRIT0$(CODD%), 64)
               !DESCRIDOS = Left$(DESCRIT2$(CODD%), 255)
               !IdenLote = LOTORI$    ' LOTEE$   ' MODIFICADO POR PEDIDO BLATNER
               !Cantidad = CANTU
               !CodEnvase = ENVU$
               !CanKilos = KILOS
               !FeReMovi = CVDAT(FEREM%)
               !FechMov = CVDAT(FEMOV%)
               !TipoTrans = 2
               !Status = 1
               !CreadoPor = "FLEX"
               !Destino = DESTIPED$
               !Ingresado = (-1) * NULIPI&
               !Ident_Asig = LOTEE$   ' IDASIX$  ' AL TRANMOVI VA CON EL LOTE FLEXOFT
               .Update
            End With
            '
         .MoveNext
         Loop
      End If
   End With
   '
   TRANTMP.Close
   Set TRANTMP = Nothing
   BALARETMP.Close
   Set BALATMP = Nothing

End Sub
'
Sub BORRATRANMO(NULIPI&)
   '
   Call ABRESTOCKS
   SQLX$ = "DELETE FROM TRANMOVI "
   SQLX$ = SQLX$ + "WHERE TipoTrans = 2 "
   SQLX$ = SQLX$ + "AND Ingresado = -" & NULIPI&
   FLEXCONN.Execute SQLX$
   '
End Sub
'
Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("INDIPICK")
    NLIPI& = XVALO(VALACT1$("INDIPICK"))
    If NLIPI& > 0 Then
        OPX% = ALTERNA%("Listado Plano\Balance de Requerimientos")
        Select Case OPX%
          Case 1
            Call FINAL("?LISPICKSAP(" + TRIM$(Str$(NLIPI&)) + ")")
          Case 2
            Call FINAL("?BALAREQMO(" + TRIM$(Str$(NLIPI&)) + ")")
        End Select
    End If
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    If DERACCE%("TAUBICA", "Tabla de Ubicaciones") > 1 Then Call CARUBIPRE ' MUESTRA VENTANA DE CARGA DE UBICACIONES PREDEFINIDAS
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call CONECRUN("INVENT07", "Inventarios Físicos y Valorizados")
    Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    Command15.Enabled = False
    FIKARDEX07.Show 1
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
'    Call COMUNI("La transaccion no pudo completarse.")
'    Exit Sub
    '
    Command16.Enabled = False
    Call SELECHO("INDIPICK")
    NLIPI& = XVALO(VALACT1$("INDIPICK"))
    If NLIPI& > 0 Then
       Call APERBASDAT("STOCKS")
       SQLX$ = "SELECT * FROM BALAREQMO "
       SQLX$ = SQLX$ + " WHERE KILO_ASIG > 0 "
       SQLX$ = SQLX$ + " AND NULISPICK = " & NLIPI&
       '
       Dim RESERSEL As ADODB.Recordset
       Set RESERSEL = New ADODB.Recordset
       RESERSEL.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       '
       With RESERSEL
          If Not (.BOF And .EOF) Then
             Do While Not .EOF
                !Kilo_Asig = 0
                .Update
                CIXI% = !Cod_Inte
                IDLO$ = !IdenLote
                Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
             .MoveNext
             Loop
          End If
       End With
       RESERSEL.Close
       Set RESERSEL = Nothing
       '
       SQLX$ = "DELETE FROM BALAREQMO "
       SQLX$ = SQLX$ + " WHERE NULISPICK = " & NLIPI&
       SQLX$ = SQLX$ + " AND KILO_REQUE < 0.0005 "      ' PARA QUE NO BORRE LOS REQUERIDOS
       FLEXCONN.Execute SQLX$
       '
       CONDI$ = "NULISPICK = " & NLIPI& & " AND KILO_REQUE > 0 " ' PARA BORRAR LOS REQUERIDOS
       Call ACTUREGISTRO("BALAREQMO", "NULISPICK", CONDI$, 0, REGAFE&, "NOMESS")
       CONDI$ = "NULISPICK < 1  AND KILO_REQUE > 0 "            ' PARA BORRAR LOS REQUERIDOS
       Call ACTUREGISTRO("BALAREQMO", "DELISPICK", CONDI$, "", REGAFE&, "NOMESS")
       '
       Call BORRATRANMO(NLIPI&)
       '
       Call FILTERDELETE("INDIPICK", 1, MKS$(NLIPI&))
       Call FRESHECHO("INDIPICK")
       Call COMUNI("Lista de Picking" + Str$(NLIPI&) + " Eliminada")
       '
    End If
    Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    Call SELECHO("INDIPICK")
    NLIPI& = XVALO(VALACT1$("INDIPICK"))
    DELIPI$ = TRIM$(VALACT2$("INDIPICK"))
    If NLIPI& > 0 Then
       Call APERBASDAT("STOCKS")
       SQLX$ = "SELECT * FROM BALAREQMO WHERE NuLisPick = " & NLIPI&
       SQLX$ = SQLX$ + " AND Kilo_Asig > 0"
       Dim BALARETMP As ADODB.Recordset
       Set BALARETMP = New ADODB.Recordset
       BALARETMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       '
       JJ& = 0
       With BALARETMP
            Do While Not .EOF
                CODD% = XVALO(!Cod_Inte)
                KILOS = XVALO(!Kilo_Asig)
                LOTEE$ = SAFETEXT(!IdenLote)
                
                JJ& = JJ& + 1
                XX$ = REGBLAN$("LISTAPIK")
                Call REPLA(XX$, MKI(CODD%), 1, 2)
                Call REPLA(XX$, MKS$(KILOS), 3, 4)
                Call REPLA(XX$, LOTEE$, 7, 16)
                
                COMPRO$ = SAFETEXT(!CodiCompro)
                NUMCOMPR& = XVALO(!NumeCompro)
                CODDSUC& = XVALO(!CodSucu)
                NUITEM& = XVALO(!NuOrItem)
                IDREQ$ = SAFETEXT(!IdenReque)
                DESTIPED$ = SAFETEXT(!Destino)
                TXTX$ = Space(234)
                Call REPLA(TXTX$, COMPRO$, 1, 6)
                Call REPLA(TXTX$, Str(NUMCOMPR&), 7, 8)
                Call REPLA(TXTX$, Str(NUITEM&), 15, 4)
                Call REPLA(TXTX$, Str(CODDSUC&), 19, 4)
                Call REPLA(TXTX, IDREQ$, 23, 24)
                Call REPLA(TXTX, DESTIPED$, 47, 188)
                Call REPLA(XX$, TXTX$, 23, 234)
                Call GRAREG("!LISTAPIK", XX$, JJ&)
            .MoveNext
            Loop
       End With
       '
       Call SETULTREG("!LISTAPIK", JJ&)
       '
10     VTB% = VENTABU%("LISTAPIK")
       If VTB% < 0 Then GoTo 99
       '
       For U& = 1 To ULTREG&("!LISTAPIK")
            X$ = REGLEIDO$("!LISTAPIK", U&)
            LOTTEE$ = TRIM$(Mid$(X$, 7, 16))
            CIXI% = CVI(Mid$(X$, 1, 2))
            XVA% = CantRegistros("LOTINGRE", "IDENLOTE = '" & LOTTEE$ & "'  And Cod_Inte = " & CIXI%)
            If XVA% < 1 Then
                Call MENSERR(24, "Lote '" & LOTTEE$ & "'" & " Invalido en Registro " & U&)
                GoTo 10
            End If
       Next U&
       '
       SQLX$ = "SELECT * FROM BALAREQMO "
       SQLX$ = SQLX$ + " WHERE KILO_ASIG > 0 "
       SQLX$ = SQLX$ + " AND NULISPICK = " & NLIPI&
       '
       Dim RESERSEL As ADODB.Recordset
       Set RESERSEL = New ADODB.Recordset
       RESERSEL.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       '
       With RESERSEL
          If Not (.BOF And .EOF) Then
             Do While Not .EOF
                !Kilo_Asig = 0
                .Update
                CIXI% = !Cod_Inte
                IDLO$ = !IdenLote
                Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
             .MoveNext
             Loop
          End If
       End With
       RESERSEL.Close
       Set RESERSEL = Nothing
       '
       SQLX$ = "DELETE FROM BALAREQMO "
       SQLX$ = SQLX$ + " WHERE NULISPICK = " & NLIPI&
       SQLX$ = SQLX$ + " AND KILO_REQUE < 0.0005 "      ' PARA QUE NO BORRE LOS REQUERIDOS
       FLEXCONN.Execute SQLX$
       '
       Call BORRATRANMO(NLIPI&)
       '
' AQUI HAY QUE GRABAR LOS REGISTROS NUEVOS
       For U& = 1 To ULTREG&("!LISTAPIK")
            X$ = REGLEIDO$("!LISTAPIK", U&)
            LOTTEE$ = TRIM$(Mid$(X$, 7, 16))
            CIXI% = CVI(Mid$(X$, 1, 2))
                         
            CCX1$ = TRIM$(Mid$(X$, 23, 6))
            NCX1 = XVALO(Mid$(X$, 29, 8))
            CSX1 = XVALO(Mid$(X$, 41, 4))
            NOI1 = XVALO(Mid$(X$, 37, 4))
            IDREQ$ = TRIM$(Mid$(X$, 45, 24))
            DESTIPEDIDO$ = TRIM$(Mid$(X$, 69, 188))
            KILAASI = CVS(Mid$(X$, 3, 4))
            '
            If KILAASI >= 0.05 Then
                CONDI$ = "Cod_Inte = " & CIXI% & " "
                CONDI$ = CONDI$ + " AND IdenLote = '" & LOTTEE$ & "' "
                POSILOTE$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
                CODENVAS$ = VALOBADA("LOTINGRE", "CodEnvase", CONDI$)
                   COEFI = 1
                KILORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
                CANORI = XVALO(VALOBADA("LOTINGRE", "Cantidad", CONDI$))
                   If KILORI > 0 Then COEFI = CANORI / KILORI
                   '
                SQLZ$ = "SELECT * FROM BALAREQMO "
                SQLZ$ = SQLZ$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                Dim ASIGNA As ADODB.Recordset
                Set ASIGNA = New ADODB.Recordset
                ASIGNA.OPEN FILTSQL$ (SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                With ASIGNA
                   .AddNew   ' YA CONVERTIDO
                   ON ERROR RESUME NEXT
                   !CodiEmpr = Codiemp%
                   ON ERROR GOTO 0
                   !CodiCompro = CCX1$
                   !NumeCompro = NCX1
                   !CodSucu = CSX1
                   !NuOrItem = NOI1
                   !IdenReque = IDREQ$
                   !Cod_Inte = CIXI%
                   !IdenLote = LOTTEE$
                   !Posicion = POSILOTE$
                   !Cantidad = KILAASI * COEFI         ' falta dederminar
                   !CodEnvase = CODENVAS$
                   !KILO_REQUE = 0
                   !Kilo_Asig = KILAASI
                   !NuLisPick = NLIPI&
                   !DeLisPick = DELIPI$
                   !Kilo_Consu = 0
                   !FeReMovi = Now
                   !FechMov = Now
                   !MarqAsi = 0
                   !CreadoPor = "FLEX"
                   !Destino = DESTIPEDIDO$
                   .Update
                End With
                ASIGNA.Close
                Set ASIGNA = Nothing
                Call ACSALOTE(CIXI%, LOTTEE$, "NOMESS")
            End If
       Next U&
       '
       Call ACTRANMOVI(NLIPI&)
       '
    End If
99  Command17.Enabled = True
End Sub

Private Sub Command2_Click()
  'LISTADO DE MOVIMIENTOS SIN UBICACION
  Call FINAL("?LISMOVRANFE")
End Sub

Private Sub Command23_Click()
    FRALOT07.Show 1
End Sub

Private Sub Command25_Click()
   Call GENTAUBI
   NUEING07.Show 1
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command6_Click()
    Call FINAL("?UBICALOTES")
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
10 OPX% = COMALTER%("Opciones de Listados Disponibles:\\Recepciones\Otros Movimientos")
   '
   If OPX% = 1 Then
20     OPY% = ALTERNA%("Recepciones por Rango de Fechas\Pendientes de Proceso\No Procesables")
       Select Case OPY%
         Case 1
            Call FINAL("?RESAGSRANFE")
            GoTo 20
         Case 2
            Call FINAL("?RESAGSPEPRO")
            GoTo 20
         Case 3
            Call FINAL("?RESAGSNOPRO")
            GoTo 20
       End Select
       GoTo 10
     ElseIf OPX% = 2 Then
30     OPY% = ALTERNA%("Requeridos por Fechas\Pendientes de Proceso\No Procesables")
       Select Case OPY%
         Case 1
            Call FINAL("?MOSAGSRANFE")
            GoTo 30
         Case 2
            Call FINAL("?MOSAGSPEPRO")
            GoTo 30
         Case 3
            Call FINAL("?MOSAGSNOPRO")
            GoTo 30
       End Select
       GoTo 10
   End If
   '
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call FINAL("?REXIPORPROD")
End Sub

Private Sub Command9_Click()
    Call FINAL("?DEXIPORPROD")
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       NUEING07.Show 1
       Call FINAL("")
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub GENTAUBI()

Call GENTAUBI2
Exit Sub

   Call APERBASDAT("STOCKS")
   SQLX$ = "SELECT * FROM UBIPRE WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ORDER BY CODIUBIC ASC "
   Set LISUBIC =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   '
   Screen.MousePointer = 11
   FIN& = CantRegistros("UBIPRE", "")
   JJ& = 0
   With LISUBIC
      Do While Not .EOF
        Call TRACE(20, JJ&, FIN&)
        CODUBI$ = TRIM$(!CodiUbic)
        CODEPRO$ = "... libre ..."
        CODIPROI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "UbicFis = '" & CODUBI & "' AND CANALTA > (CANCONSU + 0.05) ORDER BY FECHALTA ASC "))
        If CODIPROI% > 0 Then
           CODEPRO$ = CODEXT$(CODIPROI%)
        End If
        '
        TTXX$ = AJUSTI$(CODUBI$, 16) + CODEPRO$
        JJ& = JJ& + 1
        Call GRAREG("TAUBICA", TTXX$, JJ&)
      .MoveNext
      Loop
      '
   End With
   Call SETULTREG("TAUBICA", JJ&)
   Screen.MousePointer = 1
   '
End Sub

Sub GENTAUBI2()
   '
   Dim CODUBI$(), OCUPUBI$()
   Call APERBASDAT("STOCKS")
   SQLX$ = "SELECT * FROM UBIPRE WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ORDER BY CODIUBIC ASC "
   Set LISUBIC =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   '
   Screen.MousePointer = 11
   FIN& = CantRegistros("UBIPRE", "")
   ReDim CODUBI$(FIN&), OCUPUBI$(FIN&)
   '
   JJ& = 0
   With LISUBIC
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        CODUBI$(JJ&) = TRIM$(!CodiUbic)
        'CODEPRO$ = "... libre ..."
        'CODIPROI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "UbicFis = '" & CODUBI & "' AND CANALTA > (CANCONSU + 0.05) ORDER BY FECHALTA ASC "))
        'If CODIPROI% > 0 Then
        '   CODEPRO$ = CODEXT$(CODIPROI%)
        'End If
        '
        'TTXX$ = AJUSTI$(CODUBI$, 16) + CODEPRO$
        'JJ& = JJ& + 1
        'Call GRAREG("TAUBICA", TTXX$, JJ&)
      .MoveNext
      Loop
      '
   End With
   '
   SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
   SQLY$ = SQLY$ + "WHERE  (CanAlta - CanConsu) > 0 "
   Set LODISPO =  FLEXCONN.EXECUTE (FILTSQL$ (SQLY$))
   '
   KK& = 0
   With LODISPO
      Do While Not .EOF
        UBIFI$ = !UbicFis
        For LL& = 1 To FIN&
           If CODUBI$(LL&) = UBIFI$ Then
              CODEPRO$ = OCUPUBI$(LL&)
              CIXI% = !Cod_Inte
              CODE$ = CODEXT$(CIXI%)
              If InStr(CODEPRO$, CODE$) < 1 Then
                If CODEPRO$ <> "" Then CODEPRO$ = CODEPRO$ + "; "
                OCUPUBI$(LL&) = CODEPRO$ + CODE$
              End If
              GoTo 29
           End If
        Next LL&
29    .MoveNext
      Loop
   End With
   '
   For JJ& = 1 To FIN&
      If OCUPUBI$(JJ&) = "" Then
           CODEPRO$ = "... libre ..."
         Else
           CODEPRO$ = OCUPUBI$(JJ&)
      End If
      '
      TTXX$ = AJUSTI$(CODUBI$(JJ&), 16) + CODEPRO$
      Call GRAREG("TAUBICA", TTXX$, JJ&)
   Next JJ&
   '
   Call SETULTREG("TAUBICA", JJ&)
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command13_Click()
    RENVASA07.Show 1
End Sub

Private Sub Command4_Click()
    AJUSLO07.Show 1
End Sub

Private Sub Command5_Click()
    REUBILO07.Show 1
End Sub

Sub CARUBIPRE()
  '\\\OT-08-1041-OD-13/11/08///
   Call SETULTREG("UBIPRE", 0)
   I& = 0
   Call ABRESTOCKS
   '
   'RECORRO LA TABLA Y GRABO EN EL ARCHIVO UBIPRE
   '
   SQLX$ = "SELECT * FROM UBIPRE ORDER BY codiubic"
   Dim UBIPRETMP As ADODB.Recordset
   Set UBIPRETMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   With UBIPRETMP
     Do While Not .EOF
       I& = I& + 1
       CODD$ = SAFETEXT$(!CodiUbic)
       REFER$ = SAFETEXT$(!Referencia)
       STAT% = XVALO(!Status)
       Call REPLA(TX$, CODD$, 1, 16)
       Call REPLA(TX$, REFER$, 17, 64)
       Call REPLA(TX$, MKI$(STAT%), 81, 2)
       Call GRAREG("!UBIPRE", TX$, I&)
     .MoveNext
     Loop
   End With
   '
   Call SETULTREG("!UBIPRE", I&)
   '
   'MUESTRO LA VENTANA
   '
10 VTB% = VENTABU%("UBIPRE")
   If VTB% < 0 Then Exit Sub
   '
   'MARCA TODOS LOS REGISTROS  PARA BORRAR CON 1
   '
   SQLX$ = "Update UBIPRE SET Marborra = 1"
    FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   '
   For U& = 1 To ULTREG&("!UBIPRE")
     X$ = REGLEIDO$("!UBIPRE", U&)
     STATT% = CVI(Mid$(X$, 81, 2))
     If STATT% < 0 Or STATT% > 1 Then
        Call MENSERR(24, "El Status debe ser '0' o '1' ")
        GoTo 10
     End If
   Next U&
   '
   SQLX$ = "Select * From UBIPRE"
   '
   Set UBIPRETMP = New ADODB.Recordset
   UBIPRETMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   'RECORRE LA VENTANA
   For U& = 1 To ULTREG&("!UBIPRE")
      X$ = REGLEIDO$("!UBIPRE", U&)
      CODI$ = Mid$(X$, 1, 16)
      '
      If CODI$ <> "" Then
        REFFF$ = Mid$(X$, 17, 64)
        STATT% = CVI(Mid$(X$, 81, 2))
        '
        With UBIPRETMP
          .AddNew  ' AGREGA NUEVA UBICACION FISICA   ' YA CONVERTIDO
          ON ERROR RESUME NEXT
          !CodiEmpr = Codiemp%
          ON ERROR GOTO 0
          !CodiUbic = SAFETEXT(CODI$)
          !Referencia = SAFETEXT(REFFF$)
          !Status = XVALO(STATT%)
          !MarBorra = 0
          .Update
        End With
        '
      End If
   Next U&
   UBIPRETMP.Close
   Set UBIPRETMP = Nothing
   '
   'BORRO TODOS LOS QUE FUERON MARCADOS CON 1
   SQLX$ = "DELETE FROM UBIPRE WHERE Marborra > 0 "
    FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   '
   Call CIERRARCH("UBIPRE")
   Call CIERRARCH("!UBIPRE")
  '\\\OT-08-1041-OD-FIN///
End Sub

