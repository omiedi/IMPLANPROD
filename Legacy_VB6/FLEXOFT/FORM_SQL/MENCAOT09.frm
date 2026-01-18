VERSION 5.00
Begin VB.Form MENCAOT09 
   BackColor       =   &H00E0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración de Deposito Caótico"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10740
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Operación Logistica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4880
      TabIndex        =   42
      Top             =   105
      Width           =   4785
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   4215
         TabIndex        =   43
         Top             =   450
         Width           =   4275
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
            Left            =   2700
            Picture         =   "MENCAOT09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Genera una Lista de Picking por Requerimiento"
            Top             =   0
            Width           =   500
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
            Left            =   2200
            Picture         =   "MENCAOT09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Asigna y Reserva Materiales para un Cierto Requerimiento"
            Top             =   0
            Width           =   500
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
            Left            =   0
            Picture         =   "MENCAOT09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Genera Lista de Requerimientos por Orden de Produccion o Manual"
            Top             =   0
            Width           =   500
         End
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
            Left            =   3700
            Picture         =   "MENCAOT09.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Consultas y Listados"
            Top             =   0
            Width           =   500
         End
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
            Left            =   3200
            Picture         =   "MENCAOT09.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Anular Lista de Picking"
            Top             =   0
            Width           =   500
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Procesar los Requerimientos"
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
            Left            =   720
            TabIndex        =   47
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2120
      Left            =   4880
      TabIndex        =   32
      Top             =   1470
      Width           =   4785
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
         Left            =   3345
         Picture         =   "MENCAOT09.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Lista de Lotes Proximos a Vencer"
         Top             =   1350
         Width           =   560
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1395
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   33
         Top             =   1320
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
            Left            =   0
            Picture         =   "MENCAOT09.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Posición de Stock"
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
            Left            =   735
            TabIndex        =   35
            Top             =   105
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   36
         Top             =   870
         Width           =   2535
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
            Picture         =   "MENCAOT09.frx":11C6
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales por Có- digo de Producto"
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
            TabIndex        =   38
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   765
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   39
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
            Left            =   0
            Picture         =   "MENCAOT09.frx":14D0
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Código de Producto"
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   4880
      TabIndex        =   22
      Top             =   3675
      Width           =   4785
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1395
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   23
         Top             =   1380
         Width           =   2535
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
            Picture         =   "MENCAOT09.frx":17DA
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultas y Lista- dos  de Movimientos "
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
            Left            =   630
            TabIndex        =   25
            Top             =   105
            Width           =   1830
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   26
         Top             =   900
         Width           =   2535
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
            Left            =   -20
            Picture         =   "MENCAOT09.frx":1924
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Gestión de Inventario Continuo"
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
            TabIndex        =   28
            Top             =   30
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   765
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   29
         Top             =   420
         Width           =   2535
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
            Picture         =   "MENCAOT09.frx":1C2E
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Kardex  de Movimientos"
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
            TabIndex        =   31
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   210
      TabIndex        =   9
      Top             =   2520
      Width           =   4515
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1680
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   48
         Top             =   2520
         Width           =   2430
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
            Left            =   1830
            Picture         =   "MENCAOT09.frx":1F38
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Transferencias / Consumos por Lista de Picking"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command1 
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
            Left            =   15
            Picture         =   "MENCAOT09.frx":2082
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Consumo Directo de Materiales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Consumo Mat. / Scrap"
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
            Left            =   650
            TabIndex        =   50
            Top             =   75
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   19
         Top             =   1995
         Width           =   2430
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
            Picture         =   "MENCAOT09.frx":21CC
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Operaciones de Re-Envasado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones de Re-Envasado"
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
            Top             =   50
            Width           =   1620
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   16
         Top             =   1470
         Width           =   2430
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
            Picture         =   "MENCAOT09.frx":24D6
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Ajuste Directo de Existencias por Lote"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Exis- tencias por Lote"
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
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   50
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   10
         Top             =   945
         Width           =   2430
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
            Picture         =   "MENCAOT09.frx":2620
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Fraccionamiento de Lote / Partida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Fraccionamiento de Lote / Partida"
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
            TabIndex        =   12
            Top             =   50
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   420
         Width           =   2430
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
            Picture         =   "MENCAOT09.frx":276A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Cambio de Ubicación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cambio de Ubicación"
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
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   45
            Width           =   1110
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7000
      Left            =   9840
      ScaleHeight     =   6945
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   -120
      Width           =   1095
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
         Height          =   720
         Left            =   105
         Picture         =   "MENCAOT09.frx":28B4
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Alta Nuevo Item de Stock"
         Top             =   3330
         Width           =   645
      End
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
         Height          =   735
         Left            =   105
         Picture         =   "MENCAOT09.frx":2BBE
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command99 
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
         Left            =   105
         Picture         =   "MENCAOT09.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Ordenes de Trabajo y Fabricación"
         Top             =   2505
         Width           =   645
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   120
         Picture         =   "MENCAOT09.frx":330A
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   4320
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   5200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
      End
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
         Left            =   105
         Picture         =   "MENCAOT09.frx":3614
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   600
         Left            =   120
         Picture         =   "MENCAOT09.frx":391E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MENCAOT09.frx":3A68
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5490
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Ingresos a Depósito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   1260
      Width           =   4515
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2700
         TabIndex        =   6
         Top             =   420
         Width           =   2760
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
            Left            =   2160
            Picture         =   "MENCAOT09.frx":598A
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Inspeccion y Aprobacion"
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
            Picture         =   "MENCAOT09.frx":5C94
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Recepción y Etiquetado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Procesamiento de Ingresos"
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
            TabIndex        =   7
            Top             =   75
            Width           =   1425
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "MENCAOT09.frx":5DDE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Caótico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   3000
      TabIndex        =   51
      Top             =   675
      Width           =   3375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0FF&
      Caption         =   "Depósito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   1800
      TabIndex        =   54
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "MENCAOT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  '
  OPXX% = ALTERNA%("Consumo de Materiales\Despacho de Materiales\Vale de Scrap")
  '
  Select Case OPXX%
     Case Is = 1
       If DERACCE%("CONMAT", "Consumo de Materiales") >= 2 Then
          CONMAT09.Show 1
       End If
     Case Is = 2         'Por Carga Manual
       If DERACCE%("DESPMAT", "Despacho de Materiales") >= 2 Then
          DESPMAT09.Show 1
       End If
     Case Is = 3         'Por Carga Manual
       If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
          SCRLOT09.Show 1
       End If

   End Select
    '
  Command1.Enabled = True
End Sub

Private Sub Command10_Click()
    
    Call ABREORFAB
    Call ABRESTOCKS
    Command10.Enabled = False
    If DERACCE%("GENLIREQ", "Generación Lista de Requerimientos") < 2 Then GoTo 99
5   OPXX% = COMALTER%("Puede Generar Lista de Requerimientos\\Por Ordenes de Producción\Por Carga Manual de Necesidad")
    '
    Select Case OPXX%
       Case Is = 1         'Por Orden de Fabricación
         GoTo 10
       Case Is = 2         'Por Carga Manual
         Call CARLIREQ
         GoTo 99
       Case Else           'Cancelar
         GoTo 99
    End Select
    '
10  Call COPYSTRU("ECORDEF", "ECORPEN")
11  JJ& = 0
    For U& = 1 To ULTREG&("ECORDEF")
       XX$ = REGLEIDO$("ECORDEF", U&)
       NORFA$ = TRIM$(Left$(XX$, 12))
       CONDI$ = "IDENREQUE LIKE '" & NORFA$ & "%' "
       CANRE& = CantRegistros("BALAREQMO", CONDI$)
       '
       If CANRE& < 1 Then    ' ASÍ SUPRIME LOS QUE YA TIENEN ASIGNACION
         JJ& = JJ& + 1
         Call GRAREG("!ECORPEN", XX$, JJ&)
       End If
    Next U&
    Call SETULTREG("!ECORPEN", JJ&)
    '
    If ULTREG&("!ECORPEN") < 1 Then
       Call MENSERR(24, "Todas las Ordenes de Producción\Tienen Requerimientos Generados\por Fórmula Patrón.")
       GoTo 5
    End If
    '
    Call FRESHECHO("!ECORPEN")
    Call SELECHO("!ECORPEN")
    NORFA$ = TRIM$(VALACT1$("!ECORPEN"))
    DESTIINDI$ = TRIM$(VALACT2$("!ECORPEN"))
    If NORFA$ = "" Then GoTo 5
    '
    Call GENERALISTAREQ(NORFA$)
    GoTo 11
    '
99  Command10.Enabled = True
    
End Sub
'
Private Sub Command11_Click()
   Command11.Enabled = False
   '
   ' LISTADOS DE REQUERIMIENTOS
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", "IDSUBOR <> '' AND STATUORF < 3 AND CanApro < CanProy ORDER BY IDSUBOR ASC")
   Call FRESHECHO("!ECORDEF")
   '
   Call CIERRARCH("*.*")
10 Call SELECHO("!ECORDEF")
   NLIPIX$ = VALACT1$("!ECORDEF")
   '
   ' OJO!!! SI SE MODIFICA ESTA RUTINA(RECUOCPEN), TAMBIEN SE DEBE ACTUALIZAR EN
   ' EN EL PROYECTO DECAOT09, FORMULARIO MENCAOT09.
   Call RECUOCPEN(NLIPIX$)
   '
   If TRIM$(NLIPIX$) <> "" Then
25      OPX% = COMALTER%("Opciones de Listados de Requerimientos:\\Listado Plano\Balance de Detalle\Balance Acumulado")
        Select Case OPX%
          Case 1
            Call FINAL("?LISREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 2
            Call FINAL("?BALREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 3
            Call FINAL("?ACUREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
        End Select
        GoTo 10
   End If
   '
   Command11.Enabled = True
   '
End Sub

Sub RECUOCPEN(NLIPIX$)
   '
   SQLX$ = "SELECT * FROM BALAREQMO "
   SQLX$ = SQLX$ + " WHERE Kilo_Asig > 0 "
   SQLX$ = SQLX$ + " AND IDENREQUE LIKE '" & NLIPIX$ & "%' "
   SQLX$ = SQLX$ + " AND (IDENLOTE LIKE 'OC%' OR IDENLOTE LIKE 'NP%')"
   Dim BALRE As ADODB.Recordset
   Set BALRE = New ADODB.Recordset
   BALRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With BALRE
      Do While Not .EOF
        IDOC$ = SAFETEXT$(!idenLote)
        CIXI% = XVALO(!Cod_Inte)
        SALCOMPRA = SALDOCOM(IDOC$, CIXI%)
        If SALCOMPRA < XVALO(!Kilo_Asig) Then
           !Kilo_Asig = SALCOMPRA
           .Update
        End If
      .MoveNext
      Loop
   End With
   '
   BALRE.Close
   Set BALRE = Nothing
   '
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
1   PorOP = False
    OPX% = COMALTER("Generacion Lista de Picking\\Por Codigo\Por Orden de Prod.")
    If OPX% = 0 Then GoTo 99
    If OPX% = 1 Then GoTo 5
    If OPX% = 2 Then
       CONDI$ = "StatuOrf < 3 "   ' para que no este anulada ni cumplida
       CONDI$ = CONDI$ + "AND CanApro < (CanProy - 0.005) "
       CONDI$ = CONDI$ + " and NuOrProd <> '' "
       CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    
       Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
       Call SELECHO("ORDPEN")
       NORFA$ = VALACT1("ORDPEN")
       If NORFA$ = "" Then GoTo 1
       '
       Call SELECHO("TADEPOSI")
       DepOrig% = XVALO(VALACT1("TADEPOSI"))
       If VALACT1("TADEPOSI") = "" Then
          Call MENSERR(24, "Mandatorio Indicar Depósito de Origen")
          GoTo 1
       End If
       PorOP = True
       GoTo 15
    End If
    '
5   Call SELORFPIK09.Label28_Change
    SELORFPIK09.Show 1
    If SELORFPIK09.Ok = True Then  'Se valida solo si Aprueba
       DepOrig% = XVALO(SELORFPIK09.Label2)
       If DepOrig% < 1 Then
          Call MENSERR(24, "Mandatorio Indicar Depósito de Origen")
          GoTo 5
       End If
       '
       ' CALCULO POR ORDENES DE FABRICACION
       NULIPI& = 1 + XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
       DELIPI$ = SAFETEXT$(Now)
       '
       For KKI% = 1 To SELORFPIK09.List1.ListCount
          If SELORFPIK09.List1.Selected(KKI% - 1) = True Then
             NORFA$ = TRIM$(Left$(SELORFPIK09.List1.List(KKI% - 1), 10))
15           DELIPI$ = DELIPI$ + " - " + NORFA$
             CONDI$ = "NUORPROD = '" & NORFA$ & "'"
             NUSOLFA$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
             '
             ' verifica cobertura de existencia
             CONDI$ = "Destino = '" & NUSOLFA$ & "' "
             If CantRegistros("Balareqmo", CONDI$ & "and NuLisPick > 0") > 0 Then
                Call COMUNI("Ya hay Lista de Picking Generada para Ord. Prod. " & NORFA$)
                GoTo 1
             End If
             SQLX$ = "SELECT IDENREQUE,Cod_Inte,Sum(Kilo_Reque) as Sureque,Sum(Kilo_Asig) as Suasig FROM BALAREQMO WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE " & CONDI$
             SQLX$ = SQLX$ + " GROUP BY IDENREQUE,COD_INTE"
             '
             Dim BALRE As ADODB.Recordset
             Set BALRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
             With BALRE
               Do While Not .EOF
                 CIXI% = XVALO(!Cod_Inte)
                 If CRITREP%(CIXI%) <> 3 Then
'                   If XVALO(!SUREQUE) > XVALO(!SUASIG) Then
                   If XVALO(!SUREQUE) > XVALO(!SUASIG) + FRACMIN(CIXI%) Then
                     TTXX$ = "Imposible Emision de Lista de Picking\para la Orden de Produccion '" + NORFA$ + "'."
                     TTXX$ = TTXX$ + "\   \Falta Asignación de Materiales para '" + CODEXT$(CIXI%) + "'\(" + DESCRIT0$(CIXI%) + ")."
                     Call MENSERR(24, TTXX$)
                     GoTo 5
                   End If
                 End If
               .MoveNext
               Loop
             End With
             BALRE.Close
             Set BALRE = Nothing
             ' aqui se verificó el balance general
             '
             FALMAT% = 0
             CONDI$ = "Destino = '" & NUSOLFA$ & "'"
             CONDI$ = CONDI$ + " AND Kilo_Asig > 0 "
             SQLX$ = "SELECT * FROM BALAREQMO WITH(NOLOCK)"
             SQLX$ = SQLX$ + " WHERE " & CONDI$
             SQLX$ = SQLX$ + " ORDER BY IDENREQUE ASC"
             Set BALRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
             With BALRE
               Do While Not .EOF
                 CIXI% = XVALO(!Cod_Inte)
                 If CRITREP%(CIXI%) <> 3 Then
                   IDLO$ = SAFETEXT$(!idenLote)
                   CANASI = XVALO(!Kilo_Asig)
                      CONDX$ = "Cod_Inte = " & CIXI%
                      CONDX$ = CONDX$ + " AND IdenLote = '" & IDLO$ & "'"
                      CONDX$ = CONDX$ + " AND LIBERA = 'APRO'"
                   SALOT = 0
                   QTRX& = CantRegistros("LOTINGRE", CONDX$)
                   If QTRX& > 0 Then
                      Call ACSALOTE(CIXI%, IDLO$)
                      SALOT = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDX$))
                      SALOT = SALOT - XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDX$))
                   End If
                   If SALOT < CANASI Then
                      FALMAT% = 1
                      TTXX$ = "Imposible Emision de Lista de Picking\para la Orden de Producción '" + NORFA$ + "'."
                      TTXX$ = TTXX$ + "\   \Lote Asignado '" + IDLO$ + "' para '" + CODEXT$(CIXI%) + "'\(" + DESCRIT0$(CIXI%) + ")\Inexistente, Insuficiente o No Liberado (en Cuarentena)."
                      OPX% = COMALTER%(TTXX$ + "\\Continuar\Cancelar")
                      If OPX% = 2 Then GoTo 25
                   End If
                 End If
               .MoveNext
               Loop
             End With
             '
25           BALRE.Close
             Set BALRE = Nothing
             '
             If FALMAT% > 0 Then
                TTXX$ = "Faltantes de Materiales para la Orden de Producción '" + NORFA$ + "'."
                TTXX$ = TTXX$ + "\La O/Prod Pasará a Estado 'Proyectada' como '" + NUSOLFA$ + "'"
                If COMALTER%(TTXX$ + "\\Aceptar\Cancelar") = 1 Then
                   CONDI$ = "IDSUBOR = '" & NUSOLFA$ & "' "
                   Call ACTUREGISTRO("ORDEFABR", "NuOrProd", CONDI$, "", RAFE&)
                   GoTo 99
                End If
             End If
             '
             ' genera la lista de picking
             NULIPI& = 1 + XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
             DESCRIX$ = VALOBADA("BALAREQMO", "Destino", "IdenReque LIKE '" + NUSOLFA$ + "%'", "NOMESS")
             DELIPI$ = SAFETEXT$(Now) + " - " + DESCRIX$
             Call FILTERPUTRECORD("INDIPICK", 1, MKS$(NULIPI&), MKS$(NULIPI&) + DELIPI$)
             Call FILESORT("INDIPICK", "", 1, 1, "ASC")
             Call FRESHECHO("INDIPICK")
             '
             CONDI$ = "Destino = '" & NUSOLFA$ & "'"
             CONDI$ = CONDI$ + " AND Kilo_Asig > 0 "
             Call ACTUREGISTRO("BALAREQMO", "NuLisPick", CONDI$, NULIPI&, RAFE&)
             Call ACTUREGISTRO("BALAREQMO", "DELisPick", CONDI$, SAFETEXT$(Now), RAFE&)
             '
             'ACTUALIZAR UBICACIONES
             SQLX$ = "SELECT * FROM BALAREQMO WHERE " & CONDI$
             Dim lista As ADODB.Recordset
             Set lista = New ADODB.Recordset
31           On Error Resume Next
             lista.Open FILTSQL(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 31
             End If
             On Error GoTo 0
             '
             With lista
                Do While Not (.EOF)
                  CIXI% = (!Cod_Inte)
                  LOTX$ = SAFETEXT(!idenLote)
                  UBICX$ = SAFETEXT(VALOBADA("LOTINGRE", "UBICFIS", "COD_INTE = " & CIXI% & " AND IDENLOTE ='" & LOTX$ & "'"))
                  !Posicion = UBICX$
                  .Update
                  .MoveNext
                Loop
             .Close
             End With
             
             
             '
             If RAFE& > 0 Then
                 If COMALTER%("Lista de Picking " & NULIPI& & " Generada\CON " & RAFE& & " Registros.\   \¿ Desea Emitir Listado de Picking ?\\Si, Listar\No, Cancelar") = 1 Then
                    Call FINAL("?LISPICKSAP(" + TRIM$(Str$(NULIPI&)) + ")")
                 End If
               Else
                 Call MENSERR(24, "No Fue Posible Generar la Lista de Picking\para la Orden de Producción '" + NORFA$ + "'.")
             End If
             GoTo 99
             '
          End If
          If PorOP = True Then GoTo 99
       Next KKI%
    Else
        COMUNI ("Imposible Generar Lista de Picking\ \No Hay Ordenes De Producción Activas para Este Producto")
    End If
    '
99  Command12.Enabled = True
    '
End Sub

'
Private Sub Command14_Click()
    Command14.Enabled = False
    Call CONECRUN("INVENT07", "Inventarios Físicos y Valorizados")
    Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
    FIKARDEX09.Show 1
    Command15.Enabled = True
End Sub

Private Sub Command16_Click()
    '
    Command16.Enabled = False
    Call CARGALISEL("INDIPICK", "BALAREQMO", "NULISPICK/F6.0;DESTINO/A16;DELISPICK/A24", "NULISPICK > 0 AND KILO_ASIG > 0", "DISTINCT")
    Call SELECHO("INDIPICK")
    NLIPI& = XVALO(VALACT1$("INDIPICK"))
    If NLIPI& > 0 Then
       '
       Call APERBASDAT("STOCKS")
       '
       'VERIFICA SI LA LISTA DE PIKING YA FUE TRANSFERIDA
       SQLX$ = "SELECT * FROM BALAREQMO"
       SQLX$ = SQLX$ + " WHERE NULISPICK = " & NLIPI&
       Dim LISTRANTMP As ADODB.Recordset
       Set LISTRANTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With LISTRANTMP
         Do While Not .EOF
           KILCONS = XVALO(!Kilo_Consu)
           If KILCONS > 0 Then
               Call MENSERR(24, "Imposible Anular\Lista de Picking ya Transferida")
               GoTo 99
           End If
           .MoveNext
         Loop
       End With
       LISTRANTMP.Close
       Set LISTRANTMP = Nothing
       '
       CONDI$ = "NULISPICK = " & NLIPI&
       Call ACTUREGISTRO("BALAREQMO", "NULISPICK", CONDI$, 0, REGAFE&, "NOMESS")
       CONDI$ = "NULISPICK < 1 "                       ' PARA BORRAR LOS REQUERIDOS
       Call ACTUREGISTRO("BALAREQMO", "DELISPICK", CONDI$, "", REGAFE&, "NOMESS")
       '
       Call FILTERDELETE("INDIPICK", 1, MKS$(NLIPI&))
       Call FRESHECHO("INDIPICK")
       Call COMUNI("Lista de Picking" + Str$(NLIPI&) + " Eliminada")
       '
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command17_Click()
    
  Command17.Enabled = False
  '
  If DERACCE%("AJULIP", "Ajuste Lista de Piking") >= 2 Then
    Call CARGALISEL("INDIPICK", "BALAREQMO", "NULISPICK/F6.0;DESTINO/A16;DELISPICK/A24", "NULISPICK > 0 AND KILO_ASIG > 0", "DISTINCT")
    Call SELECHO("INDIPICK")
    NLIPI& = XVALO(VALACT1$("INDIPICK"))
    DELIPI$ = TRIM$(VALACT2$("INDIPICK"))
    If NLIPI& > 0 Then
       Call APERBASDAT("STOCKS")
       SQLX$ = "SELECT * FROM BALAREQMO WHERE NuLisPick = " & NLIPI&
       SQLX$ = SQLX$ + " AND Kilo_Asig > 0"
       Dim BALARETMP As ADODB.Recordset
       Set BALARETMP = New ADODB.Recordset
       BALARETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       '
       JJ& = 0
       With BALARETMP
            Do While Not .EOF
                CODD% = XVALO(!Cod_Inte)
                KILOS = XVALO(!Kilo_Asig)
                LOTEE$ = SAFETEXT(!idenLote)
                
                JJ& = JJ& + 1
                XX$ = REGBLAN$("LISTAPIK")
                Call REPLA(XX$, MKI(CODD%), 1, 2)
                Call REPLA(XX$, MKS$(KILOS), 3, 4)
                Call REPLA(XX$, LOTEE$, 7, 16)
                
                COMPRO$ = SAFETEXT(!CODICOMPRO)
                NUMCOMPR& = XVALO(!NUMECOMPRO)
                CODDSUC& = XVALO(!CODSUCU)
                NUITEM& = XVALO(!NUORITEM)
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
       If ULTREG&("!LISTAPIK") < 1 Then
       '''''
            CAMPOSEL2$ = VALACT2("INDIPICK") 'ORDEN DE FABRICACION
            IDSOF$ = Mid$(CAMPOSEL2$, 33, 12) '(OJO POSICION DEL VALACT2)
            CONDI$ = "IDSUBOR = '" & IDSOF$ & "'"
            CI% = XVALO(VALOBADA("OrdeFabr", "COD_INTE", CONDI$, "NOMESS"))
            CANTID# = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS"))
            LOTX# = LOTESTAN(CI%)
            '
            SQLX$ = "SELECT * FROM FORMUPAT WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE CODICONJ = " & CI% & ""
            Dim FMLAPTMP As ADODB.Recordset
            Set FMLAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
            '
            JR& = 0
            With FMLAPTMP
              Do While Not .EOF
                CODILMTO% = XVALO(!CODIELEM)
                UBRUTO# = XVALO(!USOBRUTO)
                Cantidad = CANTID# * (UBRUTO# / LOTX#)
                '
                JR& = JR& + 1
                XX$ = REGBLAN$("LISTAPIK")
                Call REPLA(XX$, MKI(CODILMTO%), 1, 2)
                Call REPLA(XX$, MKS$(Cantidad), 3, 4)
                Call GRAREG("!LISTAPIK", XX$, JR&)
                .MoveNext
              Loop
            End With
            Call SETULTREG("!LISTAPIK", JR&)

            FMLAPTMP.Close
            Set FMLAPTMP = Nothing
       End If
       '''''
10     VTB% = VENTABU%("LISTAPIK")
       If VTB% < 0 Then GoTo 99
       '
       RepiteVentana% = 0
       For U& = 1 To ULTREG&("!LISTAPIK")
            X$ = REGLEIDO$("!LISTAPIK", U&)
            LOTTEE$ = TRIM$(Mid$(X$, 7, 16))
            CIXI% = CVI(Mid$(X$, 1, 2))
            XVA% = CantRegistros("LOTINGRE", "IDENLOTE = '" & LOTTEE$ & "'  And Cod_Inte = " & CIXI%)
            If XVA% < 1 Then ' SI EL LOTE NO ES VALIDO O ESTA VACIO
                'LE AGREGA UN LOTE POSIBLE Y SI NO "N.Disp"
                KILS# = CVS(Mid$(X$, 3, 4))
                KILX$ = FORMATNUM$(KILS#, "F15.5")
                CONDI$ = "Cod_Inte = " & CIXI% & " AND CanSaldo > = '" & KILX$ & "' ORDER BY CANSALDO ASC"
                LOTT1$ = VALOBADA("LOTINGRE", "IDENLOTE", CONDI$, "NOMESS")
                LOTT$ = "N.Disp." ' POR DEFECTO LOTE NO DISP
                If LOTT1$ <> "" Then LOTT$ = LOTT1$ ' SI ENCONTRO LOTE SE ASIGNA
                Call REPLA(X$, LOTT$, 7, 16)
                Call GRAREG("!LISTAPIK", X$, U&)
                RepiteVentana% = 1
                   '''''''
'                Call MENSERR(24, "Lote '" & LOTTEE$ & "'" & " Invalido en Registro " & U&)
'                GoTo 10
            End If
       Next U&
       '
       If RepiteVentana% = 1 Then ' para volver a mostrar la informacion en caso que se haya ingresado algun lote
        VTB% = VENTABU%("LISTAPIK")
       End If
       '
       SQLX$ = "SELECT * FROM BALAREQMO "
       SQLX$ = SQLX$ + " WHERE KILO_ASIG > 0 "
       SQLX$ = SQLX$ + " AND NULISPICK = " & NLIPI&
       '
       Dim RESERSEL As ADODB.Recordset
       Set RESERSEL = New ADODB.Recordset
       RESERSEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       '
       With RESERSEL
          If Not (.BOF And .EOF) Then
             Do While Not .EOF
                !Kilo_Asig = 0
                .Update
                CIXI% = !Cod_Inte
                IDLO$ = !idenLote
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
       ' Call BORRATRANMO(NLIPI&)   deshabilitado - paree corresponder a saporiti
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
                ASIGNA.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                With ASIGNA
                   .AddNew   ' YA CONVERTIDO
                   On Error Resume Next
                   !CodiEmpr = CodiEmp%
                   On Error GoTo 0
                   !CODICOMPRO = CCX1$
                   !NUMECOMPRO = NCX1
                   !CODSUCU = CSX1
                   !NUORITEM = NOI1
                   !IdenReque = IDREQ$
                   !Cod_Inte = CIXI%
                   !idenLote = LOTTEE$
                   !Posicion = POSILOTE$
                   !Cantidad = KILAASI * COEFI         ' falta dederminar
                   !CodEnvase = CODENVAS$
                   !KILO_REQUE = 0
                   !Kilo_Asig = KILAASI
                   !NuLisPick = NLIPI&
                   !DeLisPick = DELIPI$
                   !Kilo_Consu = 0
                   !FeReMovi = Now
                   !FECHMOV = Now
                   !MarqAsi = 0
                   !CREADOPOR = "FLEX"
                   !Destino = DESTIPEDIDO$
                   .Update
                End With
                ASIGNA.Close
                Set ASIGNA = Nothing
                Call ACSALOTE(CIXI%, LOTTEE$, "NOMESS")
            End If
       Next U&
       '
'       Call ACTRANMOVI(NLIPI&)   DESHABILITADO - ESTO ERA PARA SAPORITI
       '
    End If
  End If
  
99  Command17.Enabled = True

End Sub

Private Sub Command18_Click()
 Command8.Enabled = False
  '
  If DERACCE%("AMASTO07", "Maestro Item de Stock") > 1 Then
   '
   If EXISTE%("AMASTO07.EXE") > 0 Then
      Call CONECRUN("AMASTO07", "Maestro de Items de Stock")
     Else
      AMAS_GES07.Show 1
   End If
   '
  End If
  '
  Command8.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    RECALMAT09.Show 1
    
'    If DERACCE%("GENLIPIK", "Generación Lista de Picking") < 2 Then GoTo 99
'    '
'    Call ABRESTOCKS
'    OPX% = ALTERNA%("Por Lista de Requerimientos\Por Orden de Producción")
'    Select Case OPX%
'       Case Is = 1
'         GoTo 10
'       Case Is = 2
'         Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", "IDSUBOR <> '' AND STATUORF < 3 AND CanApro < CanProy ORDER BY IDSUBOR ASC")
'         Call FRESHECHO("!ECORDEF")
'         Call SELECHO("!ECORDEF")
'
'         NORFA$ = VALACT1$("!ECORDEF")
'         If NORFA$ <> "" Then
'           'Asignacion de Materiales
'           Call SELECHO("TADEPOSI")
'           DEPO$ = VALACT1$("TADEPOSI")
'           If DEPO$ = "" Then GoTo 99
'           DepOrig% = XVALO(DEPO$)
'           NULIPI& = 0
'           Call REEMPLAZACOMPRAS(NORFA$, DEPORI%)
'           Call ASIGNAREQUERI(NORFA$, DepOrig%, NULIPI&)
'           '
'           Call COMUNI("Proceso de Cálculo Completo.")
'         End If
'       Case Else
'         GoTo 99
'    End Select
'    GoTo 99
'    '
'10  ' CALCULO POR LISTA DE REQUERMIENTOS
'    Call COPYSTRU("INDIREQU", "INDIREQX")
'    CONT& = 0: JJ& = 0
'    For KU& = 1 To ULTREG&("INDIREQU")
'       xx$ = REGLEIDO$("INDIREQU", KU&)
'       IDREQ$ = TRIM$(Left$(xx$, 16))
'       If CantRegistros("BALAREQMO", "IdenReque LIKE '" + IDREQ$ + "%'") > 0 Then
'         JJ& = JJ& + 1
'         Call GRAREG("INDIREQU", xx$, JJ&)
'         If Mid$(xx$, 1, 2) = "LR" Then
'           CONT& = CONT& + 1
'           Call GRAREG("INDIREQX", xx$, CONT&)
'         End If
'       End If
'    Next KU&
'    Call SETULTREG("INDIREQU", JJ&)
'    Call FRESHECHO("INDIREQU")
'    Call SETULTREG("INDIREQX", CONT&)
'    Call FRESHECHO("INDIREQX")
'    Call CIERRARCH("*.*")
'    '
'15  RESP$ = OBJPLA$("SELREQDEP", VDEF$)
'    If RESP$ = "" Then GoTo 99
'    DepOrig% = Asc(Mid$(RESP$, 16, 1))
'    NORFA$ = TRIM$(Mid$(RESP$, 1, 15))
'    If NORFA$ = "" Then GoTo 99
'    '
'    If DepOrig% = 0 Then
'       VDEF$ = RESP$
'       COMUNI ("Falta Ingresar Deposito")
'       GoTo 15
'    End If
'    '
'    NULIPI& = 1 + XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
'    DESCRIX$ = VALOBADA("BALAREQMO", "Destino", "IdenReque LIKE '" + NORFA$ + "%'", "NOMESS")
'    DELIPI$ = SAFETEXT$(Now) + " - " + DESCRIX$
'    '
'    Call ASIGNAREQUERI(NORFA$, DepOrig%, NULIPI&)
'    '
'    Call COMUNI("Proceso de Cálculo Completo.")
'    '
99  Command19.Enabled = True
    '
End Sub


Private Sub Command2_Click()
  Command2.Enabled = False
   'LISTADO DE MOVIMIENTOS SIN UBICACION
   Call FINAL("?LISMOVRANFE")
   '
  Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    '
    OPX% = ALTERNA%("Por Fecha de Vencimiento\Por Fecha de Retest")
    Select Case OPX%
       Case 1
          FF% = DIPOST(HOY(HO$) + 29)
          AA$ = FECHATEX(FF%)
          Call FINAL("?LOTEPORVENCE(" + TRIM$(AA$) + ")")
       Case 2
          FF% = DIPOST(HOY(HO$) + 29)
          AA$ = FECHATEX(FF%)
          Call FINAL("?LOTPORVENRET(" + TRIM$(AA$) + ")")
    End Select
    '
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
10 OPX% = COMALTER%("Opciones de Trabajo:\\Transferencias y Consumos\Verificación (Central de Pesadas)\Fraccionamiento (Central de Pesadas)")
   Select Case OPX%
      Case 1
         Call VERIPICK
         GoTo 10
      Case 2
         'Call COMUNI("No Disponible por el Momento")
         VERPESA09.Show 1
         GoTo 10
   End Select
   Command21.Enabled = True
End Sub

Sub VERIPICK()
'    Call SELECHO("INDIPICK")
    SELPIDE07.Show 1
    
    'NLIPI& = XVALO(VALACT1$("INDIPICK"))
    'DELIPI$ = TRIM$(VALACT2$("INDIPICK"))
    NLIPI& = XVALO(SELPIDE07.ListaSel)
    
        
    If NLIPI& > 0 Then
       DELIPI$ = TRIM$(VALOBADA("BALAREQMO", "DESTINO", "NULISPICK = " & NLIPI&))
       DEORIG& = XVALO(SELPIDE07.DepOrig)
       VERPICK09.Label7 = DEORIG&
       VERPICK09.Label33 = NLIPI&
       VERPICK09.Label24 = DELIPI$
       
       VERPICK09.Show 1
    End If
End Sub

Private Sub Command22_Click()
  Command22.Enabled = False
   STPDECAOT.Show 1
  Command22.Enabled = True
End Sub

'
Private Sub Command23_Click()
  Command23.Enabled = False
  '
  If DERACCE%("FRACLOT", "Fraccionamiento de Lote") >= 2 Then
     FRALOT09.Show 1
  End If
  '
  Command23.Enabled = True
End Sub

Private Sub Command25_Click()
  Command25.Enabled = False
  '
  If DERACCE%("RECEP", "Recepción y Etiquetado") >= 2 Then
     RECAOT091.Show 1
  End If
  '
  Command25.Enabled = True
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call FINAL("?UBICALOTES( )") ' Por defecto no carga ningun deposito
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("APRORE", "Aprobación y Recepción") >= 2 Then
      'Actualiza fecha de movimientos en caso que no se
      'haya registrado la fecha de aprobacion. Asigna fecha de recepcion
      Call RECUFECHAPRO
      Screen.MousePointer = 1
      OPX% = COMALTER%("Inspeccion y Aprobacion - Opciones de Trabajo:\\Por Producto y Lote\Masiva por Destino")
      If OPX% = 1 Then APROLOT09.Show 1
      If OPX% = 2 Then APRODES09.Show 1
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call FINAL("?REXIPORPROD")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call FINAL("?DEXIPORPROD")
    Command9.Enabled = True
End Sub

Private Sub Command99_Click()
    Command99.Enabled = False
    Call CONECRUN("ORPROD09", "Ordenes de Fabricación")
    Command99.Enabled = True
End Sub
'
Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   'CODIGO PARA COLGAR MODULO
'   FF% = FECHANUM("15/12/11")
'   If FF% <= HOY(HO$) Then
'        Call Form_Load
'   End If

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call RECULOTIN                    ' RECUPERA LOTINGRE CON IDENLOTE = ''
   '
   If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Call Command25_Click
       Call FINAL("")
   End If
   '
   If Left$(APLINVO$, 7) = "INSPEC" Then ' Or INTERPRE% = 1 Then
       Call Command7_Click
       Call FINAL("")
   End If
   '
   If Left$(APLINVO$, 7) = "DESMAT" Then ' Or INTERPRE% = 1 Then
       DESPMAT09.Show 1
       Call FINAL("")
   End If

End Sub
'
Sub RECULOTIN()
   Screen.MousePointer = 11
   Call ABRECONEXION
   CONDI$ = "IDENLOTE = '' OR IDENLOTE IS NULL"
   QTR& = CantRegistros("LOTINGRE", CONDI$)
   If QTR& > 0 Then
15    IDRECU$ = "LR-" + RANDSTRING$(-6) + "-00"
      If CantRegistros("LOTINGRE", "IDENLOTE = '" & IDRECU$ & "'") > 0 Then GoTo 15
      Call ACTUREGISTRO("LOTINGRE", "IDENLOTE", CONDI$, IDRECU$, RAFE&, "NOMESS")
   End If
   Screen.MousePointer = 1
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   '
   If DERACCE%("REENVA", "Re-Envasado") >= 2 Then
      RENVASA09.Show 1
   End If
   '
   Command13.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
10  OPX% = COMALTER%("Opciones de Ajuste:\\Ajuste Cantidad Actual\Corrección Inicial")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    
    If OPX% = 1 Then 'Ajuste Cantidad Actual
      If DERACCE%("AJULOT", "Ajustes de Existencias") >= 2 Then
       AJUSLO09.Show 1
      End If
    End If
    '
    If OPX% = 2 Then 'Corrección Inicial
      If DERACCE%("CORRINIC", "Ajuste de Corrección Inicial") >= 2 Then
         SELCODLOT07.Show 1
         CI% = CODINT%(SELCODLOT07.Text1)
         If CI% > 0 Then
            IDLOTX$ = SELCODLOT07.Label28
            'ACTUALIZO EL LOTE
            Call ACSALOTE(CI%, IDLOTX$, "NOMESS")
            '
            CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOTX$ & "'"
            CANTI = XVALO(VALOBADA("LOTINGRE", "CANALTA", CONDI$, "NOMESS"))
            CANTX$ = TRIM$(FORMATNUM$(CANTI, "F18." & CNTDC$))
            CANTI1 = XVALO(InputBox$("Nueva Cantidad ", "Ingrese Cantidad ", CANTX$))
            CANTY$ = TRIM$(FORMATNUM$(CANTI1, "F18." & CNTDC$))
            'VALIDO LA CANTIDAD INGRESADA
            If CANTX$ = CANTY$ Then GoTo 60
            If CANTI1 < 0 Then Call MENSERR(24, " No Válido Ingresar Cantidad Negativa"): GoTo 10
            If CANTI1 = 0 Then GoTo 99
            '
            'ASIGNO LOS DATOS PARA REALIZAR EL MOVISTOCK
             CMOV$ = TRIM$(CONTROL$("", "CODAJUST"))
             If CMOV$ = "" Then CMOV$ = "AJ"
             DPRI$ = TRIM$(Str$(1 + XVALO(Mid$(VALOBADA("MOVISTO", "MAX(DOPRILAR)", "CodiMovi = '" & CMOV$ & "'"), 4, 6))))
             While Len(DPRI$) < 6: DPRI$ = "0" + DPRI$: Wend
             DPRI$ = CMOV$ + "-" + DPRI$
             HOII% = HOY(HOS$)
             '
             'FECHA DE ALTA Y DEPOSITO DEL LOTINGRE
             FECH% = CVINT(VALOBADA("LOTINGRE", "FechAlta", CONDI$, "NOMESS"))
             DSEC$ = CMOV$ + "-" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
             DEP% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
             'CALCULO LA DIFERENCIA PARA EL AJUSTE
             CANTIX = CANTI1 - CANTI
             If CANTIX < FRACMIN(CI%) Then GoTo 60
             '
             'REALIZO EL AJUSTE
             Call MOVISTOK(FECH%, CI%, IDLOTX$, CMOV$, DPRI$, DPRI$, DSEC$, DSEC$, "Ajuste Existencias ", 0, 0, 0, 0, DEPX%, CANTIX)
             Call ACSALOTE(CI%, IDLOTX$, "NOMESS")
             Call SELCODLOT07.limpiar
60           Call COMUNI("Ajuste Inicial Realizado")
             Screen.MousePointer = 1
             GoTo 10
         End If
         GoTo 10
      End If
    End If
99  Screen.MousePointer = 1: Command4.Enabled = True
    Call SELCODLOT07.limpiar
    End Sub

Private Sub Command5_Click()
  Command5.Enabled = False
  '
  If DERACCE%("CAMUBI", "Cambio de Ubicación") >= 2 Then
     OPXX% = COMALTER%("Puede Re-Ubicar en Forma\\Individual\Masiva")
     If OPXX% = 1 Then REUBILO09.Show 1
     If OPXX% = 2 Then
        If DERACCE%("LEEREUBI", "Leer Ubicacion") >= 2 Then REUBIMA09.Show 1
     End If
  End If
  '
99  Command5.Enabled = True
End Sub


Sub CARLIREQ()
    '
    OPX% = COMALTER%("Lista de Requerimientos Manual:\\Nueva Lista\Modificar Lista Existente.")
    Select Case OPX%
       Case 1
          GoTo 5
       Case 2
         Call MODLIREQ
    End Select
    GoTo 99
    '
5   Call ABRECONEXION
    PRESENTA$ = Space(47)
    PROXNUM& = XVALO(VALOBADA("BALAREQMO", "MAX(NUMECOMPRO)", "CODICOMPRO = 'LR'"))
    PROXNUM& = PROXNUM& + 1
    Call REPLA(PRESENTA$, "LR", 1, 2)
    Call REPLA(PRESENTA$, MKS$(PROXNUM&), 3, 4)
    '
10  RESP$ = OBJPLA$("CAMALISTAPIK", PRESENTA$)
    '
    If RESP$ = "" Then GoTo 99
    '
    DESTI$ = TRIM$(Mid$(RESP$, 7, 40))
    If DESTI$ = "" Then
       Call MENSERR(24, "Mandatorio Ingresar Descripcion o Referencia")
       PRESENTA$ = RESP$
       GoTo 10
    End If
    '
    Call SETULTREG("!LISTAPIK", 0)
20  VTB% = VENTABU%("LISTAPI1")  ' LISTAPI1
    If VTB% < 0 Then GoTo 99    ' CANCELADO
    '
    ' GUARDA EL REQUERIMIENTO EN EL BALAREQMO
    Screen.MousePointer = 11
    '
    DOCUI$ = "LR"
    NUCO$ = TRIM$(CStr(PROXNUM&)): While Len(NUCO$) < 6: NUCO$ = "0" + NUCO$: Wend
    NOSU$ = "0"
    IDREQ$ = DOCUI$ + "-" + NUCO$ + "-" + NOSU$
     
    Dim BALARE As ADODB.Recordset
    Set BALARE = New ADODB.Recordset
    SQLX$ = "SELECT * FROM BALAREQMO WHERE NUMECOMPRO = '" & NUCO$ & "'"
    BALARE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    
    'Valida Que exista el Codigo y la Cantidad > 0 en todos los Item
    For AA& = 1 To ULTREG&("!LISTAPIK")
        XX$ = REGLEIDO$("!LISTAPIK", AA&)
        CODILMTO% = CVI(Mid$(XX$, 1, 2))
        Cantidad = CVS(Mid$(XX$, 3, 4))
        '
        If CODILMTO% = 0 Then
           COMUNI ("Codigo No Válido en Linea " & AA&)
           GoTo 20
        End If
        '
        If Cantidad < 0.000005 Then
           COMUNI ("Falta Ingresar Cantidad en Linea " & AA&)
           GoTo 20
        End If
        '
    Next AA&
    '
    For KU& = 1 To ULTREG&("!LISTAPIK")
       XX$ = REGLEIDO$("!LISTAPIK", KU&)
       CODILMTO% = CVI(Mid$(XX$, 1, 2))
       Cantidad = CVS(Mid$(XX$, 3, 4))
       BALARE.AddNew   ' YA CONVERTIDO
             On Error Resume Next
       BALARE!CodiEmpr = CodiEmp%
             On Error GoTo 0
       BALARE!CODICOMPRO = DOCUI$
       BALARE!NUMECOMPRO = XVALO(NUCO$)
       BALARE!CODSUCU = XVALO(NOSU$)
         NORI$ = TRIM$(Str$(KU&))
       BALARE!NUORITEM = KU&
         DOCING$ = IDREQ$ + "/" + NORI$
       BALARE!IdenReque = DOCING$
       BALARE!Cod_Inte = CODILMTO%
       BALARE!idenLote = ""
       BALARE!Cantidad = Cantidad
       BALARE!CodEnvase = ""
       BALARE!KILO_REQUE = Cantidad
       BALARE!Kilo_Asig = 0
       BALARE!Kilo_Consu = 0
       BALARE!FeReMovi = Now
       BALARE!FECHMOV = Now
       BALARE!MarqAsi = 0
       BALARE!CREADOPOR = "FLEX"
       BALARE!Destino = DESTI$
       BALARE.Update
    Next KU&
    BALARE.Close
    Set BALARE = Nothing
    '
    ' GENERACION DE INDIREQU
    NX$ = Space(64)
    Call REPLA(NX$, IDREQ$, 1, 16)
    Call REPLA(NX$, DESTI$, 17, 48)
    '
    Call FILTERPUTRECORD$("INDIREQU", 1, IDREQ$, NX$)
    '
    Call FRESHECHO("INDIREQU")
    Call CIERRARCH("*.*")
    Call COMUNI("Lista de Requerimientos Creada OK")
    '
99  Screen.MousePointer = 1
    '
End Sub

Sub MODLIREQ()
     Call ABRECONEXION
     Dim ListaReq As ADODB.Recordset
     Set ListaReq = New ADODB.Recordset
     SQLX$ = "SELECT distinct numecompro, codicompro, codsucu,destino  FROM BALAREQMO where  Kilo_Asig = 0 AND Kilo_Consu = 0"
     
     Set ListaReq = FLEXCONN.Execute(FILTSQL$(SQLX$))
     
     ListaReq.MoveFirst
     AX& = 0
     Do While Not (ListaReq.EOF)
         '
         
         DOCUI$ = ListaReq!CODICOMPRO
         NUCO$ = TRIM$(ListaReq!NUMECOMPRO): While Len(NUCO$) < 6: NUCO$ = "0" + NUCO$: Wend
         NOSU$ = TRIM$(ListaReq!CODSUCU)
         IDREQ$ = DOCUI$ & "-" & NUCO$ & "-" & NOSU$

         CADENA$ = Space(48)
         DESCRI$ = SAFETEXT$(ListaReq!Destino)
         '
         Call REPLA(CADENA$, IDREQ$, 1, 16)
         Call REPLA(CADENA$, DESCRI$, 17, 32)
         AX& = AX& + 1
         Call GRAREG("!INDIREQU", CADENA$, AX&)
         '
     ListaReq.MoveNext
     Loop
     '
     ListaReq.Close
     Set ListaReq = Nothing
     Call SETULTREG("!INDIREQU", AX&)
     Call FRESHECHO("!INDIREQU")
     
     'Usuario elige LR
5    Call SELECHO("!INDIREQU")
     CODI$ = TRIM$(VALACT1$("!INDIREQU"))
     If TRIM$(CODI$) = "" Then GoTo 99
     '
     CONDI$ = "IDENREQUE LIKE '" + CODI$ + "%' AND Kilo_Asig > 0"
     If CantRegistros("BALAREQMO", CONDI$) > 0 Then
        NLIP& = XVALO(VALOBADA("BALAREQMO", "NuLisPick", CONDI$))
        Call MENSERR(24, "Imposible Modificar Lista '" + CODI$ + "'.\Material Asignado por Lista de Picking Nro." & NLIP&)
        GoTo 5
     End If
     '
     NULIRE& = XVALO(Mid$(CODI$, 4, 6))
     CONDI$ = "IDENREQUE LIKE '" + CODI$ + "%'"
     DESTI$ = VALOBADA("BALAREQMO", "Destino", CONDI$, "NOMESS")
     Call REPLA(PRESENTA$, "LR", 1, 2)
     Call REPLA(PRESENTA$, MKS$(NULIRE&), 3, 4)
     Call REPLA(PRESENTA$, DESTI$, 7, 40)
     '
10   RESP$ = OBJPLA$("CAMALISTAPIK", PRESENTA$)
     If RESP$ = "" Then GoTo 99
     '
     DESTI$ = TRIM$(Mid$(RESP$, 7, 40))
     If DESTI$ = "" Then
        Call MENSERR(24, "Mandatorio Ingresar Descripcion o Referencia")
        PRESENTA$ = RESP$
        GoTo 10
     End If
     Call ACTUREGISTRO("BALAREQMO", "Destino", CONDI$, DESTI$, RAFE&)
     '
     ' GENERACION DE INDIREQU
     IDREQ$ = CODI$
     NX$ = Space(64)
     Call REPLA(NX$, IDREQ$, 1, 16)
     Call REPLA(NX$, DESTI$, 17, 48)
     Call FILTERPUTRECORD$("INDIREQU", 1, IDREQ$, NX$)
     Call FRESHECHO("INDIREQU")
     '
     'Carga Lista de Items que componen la LR seleccionada
     SQLX$ = "SELECT * FROM BALAREQMO WHERE IDENREQUE LIKE '" & CODI$ & "%'"
     Dim BALARE As ADODB.Recordset
     Set BALARE = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     JR& = 0
     With BALARE
      Do While Not .EOF
        CODILMTO% = XVALO(!Cod_Inte)
        CANTI$ = !KILO_REQUE
        DESTI$ = SAFETEXT(!Destino)
        '
        JR& = JR& + 1
        XX$ = REGBLAN$("LISTAPIK")
        Call REPLA(XX$, MKI(CODILMTO%), 1, 2)
        Call REPLA(XX$, MKS$(CANTI$), 3, 4)
        Call GRAREG("!LISTAPIK", XX$, JR&)
        .MoveNext
      Loop
    End With
    BALARE.Close
    Set BALARE = Nothing
    '
    Call SETULTREG("!LISTAPIK", JR&)
20  VENTABU ("LISTAPI1")
    '
    If VTB% < 0 Then GoTo 99    ' CANCELADO
    
    'Se Graba la información en BALAREQMO
    DOCUI$ = "LR"
    NUCO$ = TRIM$(Mid$(CODI$, 4, 6))
    NOSU$ = TRIM$(Mid$(CODI$, 11))
    IDREQ$ = CODI$
    
    'Valida Que exista el Codigo y la Cantidad > 0 en todos los Item
    For AA& = 1 To ULTREG&("!LISTAPIK")
        XX$ = REGLEIDO$("!LISTAPIK", AA&)
        CODILMTO% = CVI(Mid$(XX$, 1, 2))
        Cantidad = CVS(Mid$(XX$, 3, 4))
        '
        If CODILMTO% = 0 Then
           COMUNI ("Codigo Invalido en Linea " & AA&)
           GoTo 20
        End If
        '
        If Cantidad < 0.000005 Then
           COMUNI ("Falta Ingresar Cantidad en Linea " & AA&)
           GoTo 20
        End If
        '
    Next AA&
  
    Call BORRAREGISTROS("BALAREQMO", "IDENREQUE LIKE '" & CODI$ & "%'", REG&)
    
    Set BALARE = New ADODB.Recordset
    BALARE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    
    For KU& = 1 To ULTREG&("!LISTAPIK")
       XX$ = REGLEIDO$("!LISTAPIK", KU&)
       CODILMTO% = CVI(Mid$(XX$, 1, 2))
       Cantidad = CVS(Mid$(XX$, 3, 4))
       BALARE.AddNew   ' YA CONVERTIDO
             On Error Resume Next
       BALARE!CodiEmpr = CodiEmp%
             On Error GoTo 0
       BALARE!CODICOMPRO = DOCUI$
       BALARE!NUMECOMPRO = XVALO(NUCO$)
       BALARE!CODSUCU = XVALO(NOSU$)
       BALARE!NUORITEM = KU&
         DOCING$ = IDREQ$ + "/" + TRIM$(Str$(KU&))
       BALARE!IdenReque = DOCING$
       BALARE!Cod_Inte = CODILMTO%
       BALARE!idenLote = ""
       BALARE!Cantidad = Cantidad
       BALARE!CodEnvase = ""
       BALARE!KILO_REQUE = Cantidad
       BALARE!Kilo_Asig = 0
       BALARE!Kilo_Consu = 0
       BALARE!FeReMovi = Now
       BALARE!FECHMOV = Now
       BALARE!MarqAsi = 0
       BALARE!CREADOPOR = "FLEX"
       BALARE!Destino = DESTI$
       BALARE.Update
    Next KU&
    BALARE.Close
    Set BALARE = Nothing
    '
    Call COMUNI("Modificaciòn Registrada")

99  Screen.MousePointer = 1
    '
End Sub


Sub RECUFECHAPRO()
    '
    FEXA$ = FETEXCOR1$(FECHANUM("01/01/80"))
    CONDI$ = "FECHMOV <= '" & FEXA$ & "' AND CODIMOVI = 'CC'"
    SQLX$ = "SELECT * FROM MOVISTO WHERE " & CONDI$
    '
    Dim APROSFECH As ADODB.Recordset
    Set APROSFECH = New ADODB.Recordset
    '
25  On Error Resume Next
    APROSFECH.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With APROSFECH
        Do While Not .EOF
            LOTTEE$ = SAFETEXT(!idenLote)
            CONDI$ = "IDENLOTE = '" & LOTTEE$ & "'"
            FECAPR% = CVINT(VALOBADA("LOTINGRE", "FECH_APRO", CONDI$, "NOMESS"))
            FECAPRO$ = FETEXCOR1(FECAPR%)
            !FECHMOV = FECAPRO$
            .Update
            .MoveNext
        Loop
    End With
    '
    APROSFECH.Close
    Set APROSFECH = Nothing
    '
End Sub

