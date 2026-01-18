VERSION 5.00
Begin VB.Form TRAZAUTO07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
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
      Height          =   795
      Left            =   8400
      Picture         =   "TRAZAUTO07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   6960
      Width           =   800
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   4725
      TabIndex        =   29
      Top             =   1470
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   1870
      TabIndex        =   28
      Top             =   1470
      Width           =   15
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1245
      TabIndex        =   21
      Top             =   315
      Width           =   960
   End
   Begin VB.CommandButton Command5 
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
      Height          =   285
      Left            =   2190
      TabIndex        =   20
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   120
      TabIndex        =   17
      Top             =   1155
      Width           =   8100
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   2730
         TabIndex        =   27
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   580
         Width           =   7680
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IdenLote        Cantid.  Código           Descripción"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   26
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6780
      Left            =   8400
      ScaleHeight     =   6750
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   79
         Top             =   6120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   80
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Traz-2"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         ToolTipText     =   "PROCESO - RECORRE TODAS LAS RECEPCIONES Y RECORRE EL DETALLE Y SI EL ITEM ES TRAZABLE GRABA PARA INGRESARLO EN EL LOTINGRE."
         Top             =   5760
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Traz-1"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         ToolTipText     =   "PROCESO - RECORRE EL LOTINGRE Y ASIGNA TRAZABLE = 1 A LOS ITEMS EN TABLA MASTER"
         Top             =   5520
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   120
         Picture         =   "TRAZAUTO07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Realiza Control de Movimientos de Stock Repetidos o Duplicados"
         Top             =   4870
         Visible         =   0   'False
         Width           =   650
      End
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
         Height          =   550
         Left            =   120
         Picture         =   "TRAZAUTO07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Acceso a Tablas de Marcas y Número de Partes "
         Top             =   4260
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Imprime Sticker de Aprobación"
         Top             =   3060
         Width           =   650
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
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1255
         Width           =   650
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
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":10D2
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3660
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2450
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":181E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   655
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":1B28
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Listado de Informe de Trazabilidad\Inventario"
         Top             =   1850
         Width           =   650
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
         Height          =   550
         Left            =   105
         Picture         =   "TRAZAUTO07.frx":1E32
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   75
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "TRAZABILIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   8100
      Begin VB.CommandButton Command18 
         Caption         =   "APROBACION - REGISTRO DE CALIDAD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   78
         Top             =   4440
         Width           =   4575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1845
         TabIndex        =   73
         Top             =   2160
         Width           =   3980
      End
      Begin VB.CommandButton Command11 
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
         Height          =   285
         Left            =   5835
         TabIndex        =   72
         Top             =   2160
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   67
         Top             =   3980
         Width           =   2070
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   66
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   960
         TabIndex        =   60
         Top             =   1200
         Visible         =   0   'False
         Width           =   5190
         Begin VB.CommandButton Command10 
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
            Height          =   285
            Left            =   4900
            TabIndex        =   63
            Top             =   160
            Visible         =   0   'False
            Width           =   175
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1365
            TabIndex        =   62
            Top             =   490
            Width           =   3700
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1365
            TabIndex        =   61
            Text            =   " "
            Top             =   160
            Width           =   3525
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Traz. Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -80
            TabIndex        =   65
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Parte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -80
            TabIndex        =   64
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command7 
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
         Height          =   285
         Left            =   1470
         TabIndex        =   57
         Top             =   3165
         Width           =   175
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   945
         TabIndex        =   54
         Top             =   3165
         Width           =   540
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "Medidas Papel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   945
         TabIndex        =   49
         Top             =   1200
         Width           =   5055
         Begin VB.Label Label32 
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
            Height          =   285
            Left            =   3820
            TabIndex        =   53
            ToolTipText     =   "Cantidad nominal"
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label15 
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
            Height          =   285
            Left            =   1300
            TabIndex        =   52
            ToolTipText     =   "Cantidad nominal"
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Largo (Mt):"
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
            Left            =   2835
            TabIndex        =   51
            Top             =   465
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho (mm):"
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
            Left            =   210
            TabIndex        =   50
            Top             =   465
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "FECHAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   6160
         TabIndex        =   37
         Top             =   1080
         Width           =   1695
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   9
            Top             =   1785
            Width           =   1065
         End
         Begin VB.CommandButton Command6 
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
            Height          =   285
            Left            =   1365
            TabIndex        =   10
            Top             =   1785
            Width           =   175
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   7
            Top             =   1155
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
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
            Height          =   285
            Left            =   1365
            TabIndex        =   8
            Top             =   1155
            Width           =   175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   315
            TabIndex        =   41
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   315
            TabIndex        =   40
            Top             =   945
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   315
            TabIndex        =   39
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   315
            TabIndex        =   38
            Top             =   525
            Width           =   1020
         End
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4200
         TabIndex        =   6
         Top             =   3165
         Width           =   1800
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2715
         TabIndex        =   5
         Top             =   2520
         Width           =   3285
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6600
         TabIndex        =   82
         ToolTipText     =   "Cantidad nominal"
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   81
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   420
         TabIndex        =   74
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label40 
         Caption         =   "HAY  UN FRAME NO VISIBLE DEBAJO DE ESTE (ESTE ES PARA ASEMBLI)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   495
         TabIndex        =   70
         Top             =   960
         Visible         =   0   'False
         Width           =   6720
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -500
         TabIndex        =   69
         Top             =   4050
         Width           =   1335
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Característica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2760
         TabIndex        =   68
         Top             =   4050
         Width           =   1440
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1785
         TabIndex        =   56
         Top             =   3165
         Width           =   2175
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -900
         TabIndex        =   55
         Top             =   3250
         Width           =   1755
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4440
         TabIndex        =   48
         ToolTipText     =   "Cantidad nominal"
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3480
         TabIndex        =   47
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         TabIndex        =   36
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6600
         TabIndex        =   33
         ToolTipText     =   "Cantidad nominal"
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   945
         TabIndex        =   32
         ToolTipText     =   "Descripción"
         Top             =   645
         Width           =   4740
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   945
         TabIndex        =   31
         ToolTipText     =   "Código"
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -480
         TabIndex        =   30
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6195
         TabIndex        =   18
         Top             =   3570
         Width           =   1650
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   945
         TabIndex        =   16
         Top             =   3570
         Width           =   4380
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion Física"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4200
         TabIndex        =   13
         Top             =   2940
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.R."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4160
         TabIndex        =   12
         Top             =   3650
         Width           =   1965
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote de Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   420
         TabIndex        =   11
         Top             =   2620
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -490
         TabIndex        =   4
         Top             =   3650
         Width           =   1335
      End
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -15
      TabIndex        =   43
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1245
      TabIndex        =   42
      Top             =   735
      Width           =   4590
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6285
      TabIndex        =   35
      Top             =   315
      Width           =   1860
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4605
      TabIndex        =   34
      Top             =   420
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5655
      TabIndex        =   24
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7125
      TabIndex        =   23
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B.R. Nº"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   22
      Top             =   420
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
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
      Left            =   1470
      TabIndex        =   15
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
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
      Left            =   1470
      TabIndex        =   14
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Menu Transacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu Fracc 
         Caption         =   "Fraccionamientos"
         Begin VB.Menu FraccionaLotes 
            Caption         =   "Fraccionar Lote"
         End
         Begin VB.Menu ImprimeEtiquetaFracc 
            Caption         =   "Imprimir Etiqueta"
         End
      End
      Begin VB.Menu RegistroCalidadLotesConSaldo 
         Caption         =   "Registro de Calidad (Lotes con Saldo)"
      End
   End
End
Attribute VB_Name = "TRAZAUTO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GRABARLIST()
   Dim BRTEMP As ADODB.Recordset
   For KKI% = 1 To List1.ListCount
      List1.Selected(KKI% - 1) = True
         IDREGIS$ = TRIM$(Left$(List1.List(KKI% - 1), 12))
         CIXI% = CODINT%(Label27)
         If CIXI% < 1 Or ESTRAZABLE%(CIXI%) < 1 Then GoTo 80
         Set BRTEMP = New ADODB.Recordset
         SQLX$ = " SELECT * FROM BOLRECEP where IdenLote = '" & IDREGIS$ & "'"
         BRTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         '
         With BRTEMP
            If Not (.EOF And .BOF) Then
               !MARCA = ""
               !FechFab = CVDAT(FECHANUM(Label24))
               !LOTECOL_PROV = ""
               !COD_CLI = ""
                   FEVEN% = FECHANUM("01/01/2040")
               !PLAZO_USO = CVDAT(FEVEN%)

               !DepoUbic = XVALO(!DepoEnt)
               !UbicFis = SAFETEXT$(!UbicFis)
               !TRAZ_CLI = ""
               !NUMEPARTE = ""
               !Tipo = ""
               !CARAC = ""
               .Update
               '
               FINGRE% = FECHANUM(Label24)
               FEFAB% = FECHANUM(Label24)
               FEVEN% = FEVEN%
                  TTXX$ = List1.List(KKI% - 1)
               CANOM# = XVALO(Mid(TTXX$, 14, 10))
               LARGOM = XVALO(TRIM$(Mid$(TTXX$, 115, 8)))
               ANCHOMM = XVALO(TRIM$(Mid$(TTXX$, 107, 8)))
               DPX% = XVALO(!DepoEnt)
               UBX$ = SAFETEXT$(!UbicFis)
               MARCAX$ = ""
               LOTORI$ = ""
               Call CREALOTIN(FINGRE%, FEFAB%, FEVEN%, CIXI%, IDREGIS$, MARCAX$, LOTORI$, CANOM#, LARGOM, ANCHOMM, DPX%, UBX$)
               Call ACSALOTE(CIXI%, IDREGIS$, "NOMESS")
               '
            End If
         End With
      
80 Next KKI%
   '
   
'   For KKI% = 1 To List1.ListCount
'      List1.Selected(KKI% - 1) = False
'   Next KKI%

End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub


Private Sub Command10_Click()

    Command10.Enabled = False
    Call SELECHO("!TANUMPAR")
    PAR$ = VALACT1$("!TANUMPAR")
    If PAR$ <> "" Then
      Text4.TEXT = PAR$
    End If
    Command10.Enabled = True
    
End Sub

Private Sub Command11_Click()

   Call SELECHO("!TAMARCAS")
   MAR$ = VALACT1$("!TAMARCAS")
   If MAR$ <> "" Then
      Text3.TEXT = MAR$
   End If
   
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    
    If CLACONTRA% < 1 Then GoTo 99
    OPX% = COMALTER%("Opciones\\Marcar Items de Lotingre Como Trazables = 1\Marcar Items Componentes Como Trazables ")
    If OPX% = 1 Then
    SQLX$ = " SELECT Cod_Inte FROM LOTINGRE WITH(NOLOCK) WHERE COD_INTE > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
    
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
          CI1% = XVALO(!Cod_Inte)
          CONDI$ = "Codint = " & CI1%
          Command12.Caption = CI1%
          Command12.Refresh
          Call ACTUREGISTRO("MASTER", "Trazable", CONDI$, 1, REGAF&, "NOMESS")
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    ElseIf OPX% = 2 Then
      Call APERBASDAT("STRUC")
      FIN& = CantRegistros&("FORMULAS", "CODIELEM > 0", "NOMESS")
      SQLX$ = "SELECT DISTINCT(CODIELEM) FROM FORMULAS WITH(NOLOCK) "
      Set rs1 = READSET(SQLX$)
        '
      j& = 0
      With rs1
         Do While Not .EOF
           CI1% = XVALO(!codielem)
           j& = j& + 1
           Call TRACE(20, j&, FIN&)
           If ESTRAZABLE%(CI1%) < 1 Then
               Call ACTUREGISTRO("MASTER", "TRAZABLE", "CODINT = " & CI1%, 1, REGAF&, "NOMESS")
           End If
          .MoveNext
         Loop
      End With
      rs1.Close
      Set rs1 = Nothing

    End If
    Command12.Caption = "Traz-1"
    
99  Command12.Enabled = True

End Sub

Private Sub Command13_Click()

   Command13.Enabled = False
   opt% = ALTERNA%("\TABLA DE MARCAS\NUMEROS DE PARTES")
   If opt% = 1 Then
        Call TAMARCAS
      ElseIf opt% = 2 Then
        Call TANUMPAR
   End If
   Call FRESHECHO("!TAMARCAS")
   Call FRESHECHO("!TANUMPAR")
   '
   CIXI% = CODINT%(Label27)
   If CIXI% > 0 Then
      Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CIXI%))
      Call FILTERFILE("TANUMPAR", "!TANUMPAR", 4, MKI$(CIXI%))
   End If
   '
   Command13.Enabled = True
   
End Sub

Sub TAMARCAS()
    '
    Ci% = CODINT%(Label27)
    If Ci% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      Ci% = CODINT(COD1$)
20    If Ci% > 0 Then
         Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
         VTB% = VENTABU%("TAMARCAS/TITUPAN=Tabla de Marcas - Código: " + CODEXT$(Ci%) + " - " + DESCRIT$(Ci%))
         If VTB% >= 0 Then
            Call FILTERUPDATE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
            If CODINT%(Label27) < 1 Then GoTo 10
         End If
      End If
    End If
    '

End Sub
'
Sub TANUMPAR()
    '
    Ci% = CODINT%(Label27)
    If Ci% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      Ci% = CODINT(COD1$)
20    If Ci% > 0 Then
         Call FILTERFILE("TANUMPAR", "!TANUMPAR", 4, MKI$(Ci%))
         VTB% = VENTABU%("TANUMPAR/TITUPAN=Tabla de Marcas - Código: " + CODEXT$(Ci%) + " - " + DESCRIT$(Ci%))
         If VTB% >= 0 Then
             Call FILTERUPDATE("TANUMPAR", "!TANUMPAR", 4, MKI$(Ci%))
             If CODINT%(Label27) < 1 Then GoTo 10
         End If
       End If
    End If
    '
End Sub
Private Sub Command14_Click()
    Command14.Enabled = False
    Call TRAZORF07.Command14_Click
    Command14.Enabled = True
End Sub

 Sub Command15VIEJO_Click()
   Command15.Enabled = False
   '
   Call APERBASDAT("STOCKS")
   FIN& = CantRegistros&("MOVISTO", "", "NOMESS")
   '
   SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " order by FeReMovi ASC "
   '
   
   
   Screen.MousePointer = 11
   Set MoviTemp = READSET(SQLX$)
   With MoviTemp
         KU& = 0: APADUP% = 0
         Do While Not .EOF
           KU& = KU& + 1
           Call TRACE(20, KU&, FIN&)
           FEREMO = SafeDate(!FeReMovi)
           FEMO = SafeDate(!Fechmov)
           CIXI% = XVALO(!Cod_Inte)
           DOPR$ = SAFETEXT$(!DoPriLar)
           CAIN = XVALO(!CantIngre)
           CASA = XVALO(!Cantsalid)
           IDLO$ = SAFETEXT$(!IDENLOTE)
           NORI% = XVALO(!NuOrItem)
           '
           If TRIM$(DOPR$) <> "" And CIXI% > 0 Then
              If Not .EOF Then
              '
              APOS = XVALO(.AbsolutePosition)
              'If APOS < FIN& - 1 Then
                 Do While Not .EOF And SafeDate(!FeReMovi) = FEREMO
                   .MoveNext
                   On Error Resume Next
                   FXX1 = SafeDate(!Fechmov)
                   If Err Then
                      On Error GoTo 0
                      GoTo 39
                   End If
                   On Error GoTo 0
                   If FEMO = SafeDate(!Fechmov) Then
                     If CIXI% = XVALO(!Cod_Inte) Then
                       If DOPR$ = SAFETEXT$(!DoPriLar) Then
                         If CAIN = XVALO(!CantIngre) Then
                           If CASA = XVALO(!Cantsalid) Then
                             If IDLO$ = SAFETEXT$(!IDENLOTE) Then
                               If NORI% = XVALO(!NuOrItem) Then
                                 APADUP% = 1
                                 Screen.MousePointer = 1
                                 MsgBox "Aparente Duplicación en " + DOPR$ + " - " + CODEXT$(CIXI%)
                                 Screen.MousePointer = 11
                               End If
                             End If
                           End If
                         End If
                       End If
                     End If
                   End If
                 Loop
                 '
               '.AbsolutePosition = APOS
39            End If
           End If
         .MoveNext
       Loop
   End With
   '
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   Screen.MousePointer = 1
   TTXX$ = "Fin del Control"
   If APADUP% < 1 Then TTXX$ = TTXX$ + " sin Movimientos Duplicados."
   MsgBox TTXX$
   '
   Command15.Enabled = True

End Sub
 Sub Command15_Click()
   Command15.Enabled = False
   '
   Call APERBASDAT("STOCKS")
   FIN& = CantRegistros&("MOVISTO", "", "NOMESS")
   '
   SQLX$ = " SELECT COUNT(*), FeReMovi,Fechmov, Cod_Inte, DoPriLar, CantIngre, Cantsalid, IdenLote, NuOrItem"
   SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) WHERE DoPriLar <> '' AND Cod_Inte > 0 "
   SQLX$ = SQLX$ + " GROUP BY FeReMovi,Fechmov, Cod_Inte, DoPriLar, CantIngre, Cantsalid, IdenLote, NuOrItem"
   SQLX$ = SQLX$ + " HAVING COUNT(*) > 1 "
   SQLX$ = SQLX$ + " order by FeReMovi ASC "
   '
  
   Screen.MousePointer = 11
   Set MoviTemp = READSET(SQLX$)
   With MoviTemp
         
         KU& = 0: APADUP% = 0
         Do While Not .EOF
         
           KU& = KU& + 1
           Call TRACE(20, KU&, FIN&)
           FEREMO = SafeDate(!FeReMovi)
           FEMO = SafeDate(!Fechmov)
           CIXI% = XVALO(!Cod_Inte)
           DOPR$ = SAFETEXT$(!DoPriLar)
           CAIN = XVALO(!CantIngre)
           CASA = XVALO(!Cantsalid)
           IDLO$ = SAFETEXT$(!IDENLOTE)
           NORI% = XVALO(!NuOrItem)
           '
           'MsgBox "Aparente Duplicación en " + DOPR$ + " - " + CODEXT$(CIXI%)
           MENSAJ = MENSAJ + 1
           '
           .MoveNext
         
         Loop
   End With
   
'           If TRIM$(DOPR$) <> "" And CIXI% > 0 Then
'              If Not .EOF Then
'              '
'              APOS = XVALO(.AbsolutePosition)
'              'If APOS < FIN& - 1 Then
'                 Do While Not .EOF And SafeDate(!FeReMovi) = FEREMO
'                   .MoveNext
'                   On Error Resume Next
'                   FXX1 = SafeDate(!Fechmov)
'                   If Err Then
'                      On Error GoTo 0
'                      GoTo 39
'                   End If
'                   On Error GoTo 0
'                   If FEMO = SafeDate(!Fechmov) Then
'                     If CIXI% = XVALO(!Cod_Inte) Then
'                       If DOPR$ = SAFETEXT$(!DoPriLar) Then
'                         If CAIN = XVALO(!CantIngre) Then
'                           If CASA = XVALO(!Cantsalid) Then
'                             If IDLO$ = SAFETEXT$(!IdenLote) Then
'                               If NORI% = XVALO(!NuOrItem) Then
'                                 APADUP% = 1
'                                 Screen.MousePointer = 1
'                                 MsgBox "Aparente Duplicación en " + DOPR$ + " - " + CODEXT$(CIXI%)
'                                 Screen.MousePointer = 11
'                               End If
'                             End If
'                           End If
'                         End If
'                       End If
'                     End If
'                   End If
'                 Loop
'                 '
'               '.AbsolutePosition = APOS
'39            End If
'           End If
'         .MoveNext
'       Loop
'   End With
   '
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   Screen.MousePointer = 1
   TTXX$ = "Fin del Control"
   If APADUP% < 1 Then TTXX$ = TTXX$ + " sin Movimientos Duplicados."
   MsgBox TTXX$
   '
   Command15.Enabled = True

End Sub


Private Sub Command16_Click()
    Command16.Enabled = False
    OPX% = ALTERNA%("Listado de Informes de Trazabilidad\Captura de Inventarios")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      ' LISTADO DE CONSUMO POR LOTE DE RECEPCIÓN
      Call LICONLOT
      GoTo 99
    ElseIf OPX% = 2 Then
      ' Captura de Inventarios
      LECINV07.Show 1
      GoTo 99
    End If
99  Command16.Enabled = True
End Sub

Private Sub Command17_Click()
            If CLACONTRA% < 1 Then GoTo 99
            
            Call BLOQUEAR_CONTROLES(Me, 2, 1, 1)
            
            Command17.Enabled = False
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS <= 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe DESC"
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP As ADODB.Recordset
            Set BRECEPTMP = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            j& = 0
            k& = 0
            With BRECEPTMP
              If Not (.EOF And .BOF) Then
                Do While (BRECEPTMP.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = XVALO(!NUMEBORE)
                  fech% = CVINT(!FechRecep)
                  If fech% < (HOY(HO$) - 420) Then GoTo 99
                  Text1.TEXT = NROBOL&
                  If XVALO(Text1) > 0 Then
                   If NROBOL& <> ANTNUM Then
                     Call GRABARLIST
                    End If
                  End If
                  ANTNUM = NROBOL&
                .MoveNext
                Loop
              End If
            End With
            BRECEPTMP.Close
            Set BRECEPTMP = Nothing
            '
            '
99          Screen.MousePointer = 1
            Command17.Enabled = True
            
            Call BLOQUEAR_CONTROLES(Me, 1, 1, 1)
End Sub

Private Sub Command18_Click()
    
    Command18.Enabled = False

    If List1.ListCount > 0 Then
        If List1.ListIndex > -1 Then
            IDENLOTE$ = TRIM$(Mid$(List1.List(List1.ListIndex), 1, 12))
            codigoArticulo$ = TRIM$(Mid$(List1.List(List1.ListIndex), 26, 16))
            descripcionArticulo$ = TRIM$(DESCRIT0$(CODINT%(codigoArticulo$)))
            CodigoInterno% = CODINT%(codigoArticulo$)
            Call FILTERFILE("TANUMPAR", "!TANUMPAR", 4, MKI$(CodigoInterno%))
            
            FILTX$ = IDENLOTE$ & ";" & codigoArticulo$ & ";" & descripcionArticulo$
            formularioCalidad$ = TRIM$(Control("TRAZABI", "REGICAL"))
            Call STDFORM(formularioCalidad$, FILTX$)
        Else
            Call COMUNI("Debe Elegir uno de los Artículos de la Recepción")
        End If
    Else
        Call COMUNI("Falta Elegir una Recepción")
    End If
    
    Command18.Enabled = True
    
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   NUBORE& = XVALO(Text1)
   If NUBORE& < 1 Then
      Call MENSERR(24, "Número de Boleta de Recepción no Válido.")
      GoTo 99
   End If
   '
   ' CODIGO MODIFICADO ASEMBLI
   If List1.ListCount < 1 Then
      Call MENSERR(24, "No Existen Items Trazables en La Recepción Seleccionada o la Misma es Inválida.")
      GoTo 99
   End If
   'SE COMENTA POR SI SE CARGA MANUALMENTE Y EL DAT NO ESTA ACTUALIZADO (POR QUE A VECES APARECIA VACIO)
'   If ECOARCH$("INDIBOL", MKS$(NUBORE&)) = "" Then
'      Call MENSERR(24, "Número de Boleta de Recepción no Válido.")
'      GoTo 99
'   End If
   '' CODIGO MODIFICADO ASEMBLI HASTA ACA
   CIXI% = CODINT%(Label27)
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido")
      GoTo 99
   End If
   '
   CODISE$ = UCase$(TRIM$(Label27))
   For KKI% = 1 To List1.ListCount
      If List1.Selected(KKI% - 1) = True Then
         CODIRE$ = UCase$(TRIM$(Mid$(List1.List(KKI% - 1), 26, 16)))
         If CODIRE$ <> CODISE$ Then
            Call MENSERR(24, "No Válido Combinar Códigos Diferentes")
            GoTo 99
         End If
      End If
   Next KKI%
   '
   If Command11.Visible = True Then
        CM$ = RECFILT$("!TAMARCAS", 4, MKI$(CIXI%))
        If CM$ = "" Then
          If TRIM$(Text3) = "" Then
            GoTo 20
          End If
        End If
        For KKU& = 1 To Len(CM$) Step 4
           REMA& = CVS(Mid$(CM$, KKU&, 4))
           XX1$ = REGLEIDO$("!TAMARCAS", REMA&)
           If TRIM$(Left$(XX1$, 32)) = TRIM$(Text3) Then GoTo 20
           
        Next KKU&
        'Call MENSERR(24, "Marca no Válida.")
        'GoTo 99
   End If
   '
   If Command13.Visible = True Then
20     CP$ = RECFILT$("TANUMPAR", 4, MKI$(CIXI%))
       If CP$ = "" Then
         If TRIM$(Text4) = "" Then
           GoTo 30
         End If
       End If
       '
       For KKU& = 1 To Len(CP$) Step 4
          REMA& = CVS(Mid$(CP$, KKU&, 4))
          XX1$ = REGLEIDO$("TANUMPAR", REMA&)
          If TRIM$(Left$(XX1$, 32)) = TRIM$(Text4) Then GoTo 30
       Next KKU&
       '
       'Call MENSERR(24, "Número de Parte no Válido.")
       'GoTo 99
   End If
   '
30 If FECHANUM(Text8) > FECHANUM(Label17) Then
      'Call MENSERR(24, "La Fecha de Fabricación no Puede Ser\Posterior a la Fecha de Ingreso.")
      'GoTo 99
   End If
   '
   If FECHANUM(Text12) < FECHANUM(Text8) Then
      'Call MENSERR(24, "La Fecha de Vencimiento no Puede Ser\Anterior a la Fecha de Fabricación.")
      'GoTo 99
   End If
   '
   If FECHANUM(Text12) < HOY(HOS$) Then
      'Call MENSERR(24, "Este Material ya Está Vencido\al Día de Hoy.")
   End If
   '
   If TRIM$(Text17.TEXT) = "" Then
      'Call MENSERR(24, "Ubicacion Física no Válida.")
      'GoTo 99
   End If
   '
   Dim BRTEMP As ADODB.Recordset
   For KKI% = 1 To List1.ListCount
      If List1.Selected(KKI% - 1) = True Then
         IDREGIS$ = TRIM$(Left$(List1.List(KKI% - 1), 12))
         Set BRTEMP = New ADODB.Recordset
         SQLX$ = " SELECT * FROM BOLRECEP where IdenLote = '" & IDREGIS$ & "'"
         BRTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         '
         With BRTEMP
            If Not (.EOF And .BOF) Then
               !MARCA = Left$(Text3, 32)
               !FechFab = CVDAT(FECHANUM(Left$(Text8, 8)))
               !LOTECOL_PROV = Left$(Text14, 16)
               !COD_CLI = Left$(Text14, 32)
                   FEVEN% = FECHANUM(Text12)
               !PLAZO_USO = CVDAT(FEVEN%)
               !DepoUbic = XVALO(Text2)
               !UbicFis = Left$(Text17, 32)
               !TRAZ_CLI = Left$(Text9, 32)
               !NUMEPARTE = Left$(Text4, 32)
               !Tipo = Left$(Text11, 32)
               !CARAC = Left$(Text16, 32)
               .Update
               '
               FINGRE% = FECHANUM(Left$(Label17, 8))
               FEFAB% = FECHANUM(Text8)
               FEVEN% = FECHANUM(Text12)
                  TTXX$ = List1.List(KKI% - 1)
               CANOM# = XVALO(Mid(TTXX$, 14, 10))
               LARGOM = XVALO(TRIM$(Mid$(TTXX$, 115, 8)))
               ANCHOMM = XVALO(TRIM$(Mid$(TTXX$, 107, 8)))
               DPX% = XVALO(Text2)
               UBX$ = Left$(Text17, 32)
               MARCAX$ = Left$(Text3, 32)
               LOTORI$ = Left$(Text14, 16)
               Call CREALOTIN(FINGRE%, FEFAB%, FEVEN%, CIXI%, IDREGIS$, MARCAX$, LOTORI$, CANOM#, LARGOM, ANCHOMM, DPX%, UBX$)
               Call ACSALOTE(CIXI%, IDREGIS$, "NOMESS")
               '
               ' IMPRIME EL STICKER DE APROBACION
               Call PRINSTIKAPRO(CIXI%, IDREGIS$, DPX%)
               '
             Else
               Call MENSERR(24, "Registro no encontrado.")
               '  MENSAje de que no encontro el registro
            End If
         End With
      End If
   Next KKI%
   '
   For KKI% = 1 To List1.ListCount
      List1.Selected(KKI% - 1) = False
   Next KKI%
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
   List1.SetFocus
   '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ACHERNAR") < 1 Then
        Call ACCDIR("CONSTOCK")
    Else
        OPX% = ALTERNA%("Fraccionamiento de Bobinas\Control de Partidas\Otros Accesos Directos")
        If OPX% = 1 Then
            Call FRABOBIN
          ElseIf OPX% = 2 Then
            Call CONLOTES
          ElseIf OPX% = 3 Then
          Call ACCDIR("CONSTOCK")
        End If
    End If
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
   'FORHLPDF.Pdf1.LOADFILE "C:\Mis Documentos\contab04.pdf"
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
'      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
'      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text8.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command5_Click()
   ' CODIGO MODIFICADO ASEMBLI
   Command5.Enabled = False
   
'   Static YAINGRESO%
'   If YAINGRESO% < 1 Then
'     Screen.MousePointer = 11
'     Call GENINDIBOL(0)
'     YAINGRESO% = 1
'     Screen.MousePointer = 1
'   End If
'3  Call SELECHO("INDIBOL")
'   NPX& = Val(VALACT1$("INDIBOL"))

    TituloLista$ = "Recepciones de materiales"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_BOLRECEPEN()
    Call Carga_MSF_Recordset(rs1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    numeroRecepcion$ = TRIM(RESP_ZELE_VECT(1))

    If numeroRecepcion$ <> "" Then
        Text1.TEXT = numeroRecepcion$
    End If
    
99  Command5.Enabled = True

End Sub

Sub GENINDIBOL(MODO%)
        '\\\OT-06-0543-RG-12/12/06///
        ' CODIGO MODIFICADO ASEMBLI SE TRAJO ESTA RUTINA AQUI HABRIA QUE COMPARTIRLA EN UN MODULO
        Call COPYSTRU("IBOLPEN", "IBOPEAN")
        
        If MODO% = 0 Then
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS <= 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP As ADODB.Recordset
            Set BRECEPTMP = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            j& = 0
            k& = 0
            With BRECEPTMP
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (BRECEPTMP.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     x$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(x$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(x$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(x$, REMI$, 15, 16)
                     Call REPLA(x$, SAFETEXT$(!Raso_Prov), 32, 29)
                     i& = i& + 1
                     Call GRAREG("INDIBOL", x$, i&)
                     'If (ST& < 1) Or (ST& = 1 And NUORSE& > 0) Then
                     If (ST& < 1) Then
                        j& = j& + 1
                        Call GRAREG("IBOPEAN", x$, j&)
                     End If
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", j&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP.Close
            Set BRECEPTMP = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        ElseIf MODO% = 1 Then
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS <= 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP1 As ADODB.Recordset
            Set BRECEPTMP1 = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            j& = 0
            k& = 0
            With BRECEPTMP1
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     x$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(x$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(x$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(x$, REMI$, 15, 16)
                     Call REPLA(x$, SAFETEXT$(!Raso_Prov), 32, 29)
                     i& = i& + 1
                     Call GRAREG("INDIBOL", x$, i&)
                     If (ST& < 1) Or NUORSE& > 0 Then
                        j& = j& + 1
                        Call GRAREG("IBOPEAN", x$, j&)
                     End If
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", j&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP1.Close
            Set BRECEPTMP1 = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        Else
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS < 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP2 As ADODB.Recordset
            Set BRECEPTMP2 = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            j& = 0
            k& = 0
            With BRECEPTMP2
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     x$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(x$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(x$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(x$, REMI$, 15, 16)
                     Call REPLA(x$, SAFETEXT$(!Raso_Prov), 32, 29)
                     j& = j& + 1
                     Call GRAREG("IBOPEAN", x$, j&)
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", j&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP2.Close
            Set BRECEPTMP2 = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        End If
End Sub


Private Sub Command6_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text12.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command7_Click()
    Call CommSel("TADEPOSI", Command7, Text2)
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call TRAZORF07.Command10_Click
   Command8.Enabled = True
End Sub
'
Private Sub Command9_Click()
    Command9.Enabled = False
    
    If List1.ListCount > 0 Then
    
        If InStr(1, EMPREAC$, "ACHERNAR") > 0 Then
              CIXI% = CODINT%(Label27)
              LOTE$ = TRIM$(Label23) 'NUM LOTE
              DEPO% = XVALO(Text2)
              Call PRINSTIKAPRO(CIXI%, LOTE$, DEPO%)
              GoTo 99
        End If
        
        opcionElegida% = COMALTER%("Opciones de impresión\\Solo Lote Activo\Todos los Lotes (Una Etiqueta)")
        If opcionElegida% < 1 Or opcionElegida% > 2 Then GoTo 99
        
        If opcionElegida% = 1 Then
            Ci% = CODINT%(Label27)
            If Ci% > 0 Then
              LOTE$ = TRIM$(Label23)
              If LOTE$ <> "" Then
                DEPO% = XVALO(Text2)
                CantidadLote# = XVALO(Label41)
                '
                RegistroDat$ = REGBLAN$("IMPRETIQ")
                Call REPLA(RegistroDat$, MKI$(1), 1, 2)
                Call REPLA(RegistroDat$, MKS$(CantidadLote#), 3, 4)
                Call GRAREG("!IMPRETIQ", RegistroDat$, 1)
                Call SETULTREG("!IMPRETIQ", 1)
                '
5               OK% = VENTABU%("IMPRETIQ")
                If OK% >= 0 Then
                    CantidadTotalEtiquetas# = 0
                    For f& = 1 To ULTREG&("!IMPRETIQ")
                        RegistroDat$ = REGLEIDO$("!IMPRETIQ", f&)
                        
                        CantidadEtiqueta& = CVI(Mid$(RegistroDat$, 1, 2))
                        CantidadItem# = CVS(Mid$(RegistroDat$, 3, 4))
                        CantidadTotalEtiquetas# = CantidadTotalEtiquetas# + (CantidadEtiqueta& * CantidadItem#)
                    Next f&
                    
                    If Abs(CantidadTotalEtiquetas# - CantidadLote#) > 0 Then
                        Call MENSERR(24, "La Cantidad Total de Etiquetas Difiere del Lote Ingresado")
                        GoTo 5
                    End If
                    
                    For f& = 1 To ULTREG&("!IMPRETIQ")
                        RegistroDat$ = REGLEIDO$("!IMPRETIQ", f&)
                        
                        CantidadEtiqueta& = CVI(Mid$(RegistroDat$, 1, 2))
                        CantidadItem# = CVS(Mid$(RegistroDat$, 3, 4))
                        If CantidadEtiqueta& > 0 Then
                            For FF& = 1 To CantidadEtiqueta&
                                Call PRINSTIKAPRO(Ci%, LOTE$, DEPO%, CantidadItem#)
                            Next FF&
                        End If
                    Next f&
                End If
                'CantidadEtiquetas& = XVALO(InputBox("Ingrese cantidad de etiquetas", "Impresión de etiquetas", "1"))
                'CantidadUltimaEtiqueta# = XVALO(InputBox("Ingrese cantidad de la ULTIMA etiqueta", "Impresión de etiquetas", TRIM$(FORMATNUM$(CantidadLote#, "F12.1"))))
                'If CantidadEtiquetas& > 0 And CantidadUltimaEtiqueta# > 0.05 Then
                '  For i& = 1 To CantidadEtiquetas& - 1
                '    Call PRINSTIKAPRO(Ci%, LOTE$, DEPO%, CantidadLote#)
                '  Next i&
                '  ' IMPRIME ULTIMA ETIQUETA
                '  Call PRINSTIKAPRO(Ci%, LOTE$, DEPO%, CantidadUltimaEtiqueta#)
              End If
            End If
        Else
            For indLote% = 0 To List1.ListCount - 1
            
                LoteRecepcion$ = TRIM$(Mid$(List1.List(indLote%), 1, 12))
                codigoArticulo$ = TRIM$(Mid$(List1.List(indLote%), 26, 16))
                ciCodigoArticulo% = CODINT%(codigoArticulo$)
                CantidadLote# = XVALO(TRIM$(Mid$(List1.List(indLote%), 162, 10)))
                
                If ciCodigoArticulo% > 0 And LoteRecepcion$ <> "" Then
                    Call PRINSTIKAPRO(ciCodigoArticulo%, LoteRecepcion$, DEPO%, CantidadLote#)
                End If
            
            Next indLote%
        End If
    End If
    
99  Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    Call ABRECONEXION
    CREATABLA_CALIDAD
    
    If OKEY% < 1 Then Call FINAL("")
    Call InstalaFuenteAnchoFijo
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If Left$(APLINVO$, 8) = "FRABOBIN" Then
       FRABOBI07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "PREPATIN" Then
       PREPATI07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "TRAZORFA" Or Left$(APLINVO$, 7) = "OFABRIC" Then
       TRAZORF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "ORFCIE" Or INTERPRE = 1 Then
       If InStr(EPAC$, "ASEMBLI") > 0 Or InStr(EPAC$, "DAI ICHI") > 0 Then
          CIERRAOF14.Show 1
          Call FINAL("")
          GoTo 20
       End If
       '
       CIERRAOF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       Openforms = DoEvents
    End If
    '
    If Left$(APLINVO$, 7) = "GENUMSE" Then
       GENUMSE07.Show 1
       Call FINAL("")
    End If
    '
20  If InStr(EPAC$, "ASEMBLI") > 0 Then
       'DATOS DE ASEMBLI (OCULTAN EL QUE USA ACHERNAR)
       Frame8.Visible = True
       Command10.Visible = True
       Command11.Visible = True
       'Command12.Visible = True
       Command15.Visible = True
       Command13.Visible = True
    End If
    
    
    If TRIM$(Control("TRAZABI", "REGICAL")) <> "" Then
        Command18.Enabled = True
    End If
    
End Sub

Sub VERIPERPRUE()
    '
    If Control$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(Control$("TRAZABI", "FINICIO"))
    FECHACT = CVINT(FileDateTime(LUDAT$ + "MONIEXEC.DAT"))
    '
    Dim BB As String * 128
    NX1% = FILEOPEN(LUDAT$ + "MONIEXEC.DAT", 0, 128)
    FINI& = LOF(NX1%) / 128
    COMI& = FINI& - 128: If COMI& < 1 Then COMI& = 1
    '
    For U& = COMI& To FINI&
       Get #NX1%, U&, BB$
       FEJECU = FECHANUM(Mid$(BB$, 69, 8))
       If FEJECU > FECHACT Then
          FECHACT = FEJECU
       End If
    Next U&
    '
    If DIENTRE%(FINICIO, FECHACT) > 7 Then
       Call MENSERR(24, "Plazo de Verificación Vencido.")
       Call FINAL("")
    End If
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")

End Sub

'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub FraccionaLotes_Click()

10  TituloLista$ = "Maestro de Artículos de Stock"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Codigo/A16;Descripcion/A48", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Codigo/A16;Descripcion/A48", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT CODIGO, DESCRIPCION FROM MASTER WHERE STATUS>0 ORDER BY CODIGO"
    Set Rst = FLEXCONN.Execute(ConsultaSql$)
    Call Carga_MSF_Recordset(Rst, "Codigo/A16;Descripcion/A48", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
    FRMZELECHO.Show 1
    
    CodigoArticuloExterno$ = TRIM(RESP_ZELE_VECT(1))
    CodigoArticuloInterno% = CODINT%(CodigoArticuloExterno$)
    
    If CodigoArticuloInterno% > 0 Then
            
        Call SaldoLotesConYSinReserva(CodigoArticuloInterno%)
        Call FRESHECHO("!LILOTSE")
        '
20      Call SELECHO("!LILOTSE")
        VDEVU$ = VALACT1$("!LILOTSE")
        VDEV2$ = VALACT2$("!LILOTSE")
        If TRIM$(VDEVU$) <> "" Then
            SaldoConReserva# = XVALO(Mid$(VDEV2$, 33, 12))
            CantidadFraccion# = XVALO(InputBox("Fraccionamiento de Lotes", "Cantidad del Nuevo Lote", ""))
            
            If CantidadFraccion# >= SaldoConReserva# Then
                Call MENSERR(24, "La Cantidad del Nuevo Lote\No Puede ser Igual o Mayor a la del Saldo")
                GoTo 20
            End If
            If CantidadFraccion# <= 0 Then
                Call MENSERR(24, "La Cantidad del Nuevo Lote\Debe ser Mayor a Cero")
                GoTo 20
            End If
                      
            TituloLista$ = "Saldo por Depósito"
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
            FRMZELECHO.Caption = TituloLista$
        
            ConsultaSql$ = "SELECT DEPOMOVI, COD_INTE, (SUM(CANTINGRE) - SUM(CANTSALID)) AS SALDO FROM MOVISTO WITH(NOLOCK) WHERE IDENLOTE='" & VDEVU$ & "' GROUP BY DEPOMOVI, COD_INTE ORDER BY DEPOMOVI"
            Set RsSaldoPorDeposito = READSET(ConsultaSql$)
            Call Carga_MSF_RecordsetConEco(RsSaldoPorDeposito, "Nro/I5;Depósito(DAT|TADEPOSI|BYTE)/A24;Saldo/F12.1", FRMZELECHO.msf2, 1)
            On Error Resume Next
            RsSaldoPorDeposito.Close
            Set RsSaldoPorDeposito = Nothing
            On Error GoTo 0
                
            FRMZELECHO.Width = 8000
            FRMZELECHO.TXTBUSCAR = ""
            FRMZELECHO.Show 1
            
            DEPOSITO% = XVALO(RESP_ZELE_VECT(1))
            SaldoDeposito# = XVALO(RESP_ZELE_VECT(3))
            If DEPOSITO% = 0 Then
                Call MENSERR(24, "Para Fraccionar es Obligatorio Elegir un Depósito")
                GoTo 20
            End If
            If SaldoDeposito# < 0 Or CantidadFraccion# > SaldoDeposito# Then
                Call MENSERR(24, "Saldo Insuficiente del Lote en el Depósito Elegido")
                GoTo 20
            End If
            
            DOPRI$ = "CB-" + TRIM$(Str$(ProNroComprobante("CB")))
            DOSEC$ = DOPRI$
            REFEXI$ = "Fraccion lote original"
            LOTE$ = VDEVU$
            '
            VUNI# = 0
            MONEII% = 0
            NC% = 0
            DEPORI% = DEPOSITO%
            CANTI = CantidadFraccion#
            '
            Screen.MousePointer = 11
            ' declara movimiento de consumo lote original
            Call MOVISTOK(HOY(HO$), CodigoArticuloInterno%, LOTE$, "CB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFEXI$, 0, VUNI#, MONEII%, NC%, DEPORI%, CANTI * (-1))
            Call ACSALOTE(CodigoArticuloInterno%, LOTE$, "NOMESS")
            '
            MARCA$ = SAFETEXT$(VALOBADA("LOTINGRE", "MARCA", "IDENLOTE='" & LOTE$ & "'"))
            LoteFabricacion$ = LOTE$
            NCR1$ = ""
            For KKII% = 1 To 99
               NCRX$ = LOTE$ + "/" + Right$("00" & CStr(KKII%), 2)
               If CantRegistros&("LOTINGRE", "IDENLOTE='" & NCRX$ & "'") = 0 Then
                  NCR1$ = NCRX$
                  GoTo 30
               End If
            Next KKII%
30          DEPO1$ = CStr(DEPOSITO%)
            UBIFI1$ = SAFETEXT$(VALOBADA("LOTINGRE", "UBICFIS", "IDENLOTE='" & LOTE$ & "'"))
            FEFA% = CVINT(VALOBADA("LOTINGRE", "FechFab", "IDENLOTE='" & LOTE$ & "'"))
            FEVE% = CVINT(VALOBADA("LOTINGRE", "FechVenc", "IDENLOTE='" & LOTE$ & "'"))
            REFE$ = "Fracc. " + LOTE$
            CANTI = CantidadFraccion#
            Call CREALOTIN(HOY(HO$), FEFA%, FEVE%, CodigoArticuloInterno%, NCR1$, MARCA$, LoteFabricacion$, CDbl(CANTI), 0, 0, XVALO(DEPO1$), UBIFI1$)
            Call MOVISTOK(HOY(HO$), CodigoArticuloInterno%, NCR1$, "FB", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, XVALO(DEPO1$), CANTI)
            '
            Call COMUNI("Fraccionamiento Realizado con Exito")
            '
            Screen.MousePointer = 1
        End If
        GoTo 10
    End If
        
End Sub

Sub SaldoLotesConYSinReserva(CodigoInterno%, Optional ExcluyeReservasOF$)
    '
    Static VERIRECLE%
    Dim LoteRecepcion$
    '
    If VERIRECLE% < 1 Then
       VERIRECLE% = 1
       If RECLEN%("LILOTSE") < 60 Then
          Call MENSERR(24, "Longitud de Registro Incorrecta\para Controlador 'LILOTSE'.\   \Reinstale el Controlador .RFS")
          Call FINAL("")
       End If
    End If
    '
    On Error GoTo ControlErrores
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CodigoInterno% & " "
    SQLX$ = SQLX$ + " AND ((CanAlta - CanConsu) > 0 "    ' no trae los consumidos
    SQLX$ = SQLX$ + " OR CanConsu IS NULL or CanReser IS NULL OR CANCONSU = 0 OR CANRESER = 0)"
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, IdenLote ASC "
    Set LOTINGTEMP = READSET(SQLX$)
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
         Do While Not .EOF
           LoteRecepcion$ = TRIM$(SAFETEXT$(!IDENLOTE))
           '
           Call ACSALOTE(CIXI%, LoteRecepcion$, "NOMESS")
           '
           TTXX$ = REBLA$
             FECHAALTA% = CVINT(!FechAlta)
           Call REPLA(TTXX$, LoteRecepcion$, 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FECHAALTA%), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta), "F12,1"), 25, 12)
           '
           CANCON# = 0
           On Error Resume Next
           CANCON# = !CanConsu
           On Error GoTo 0
           SaldoLogico# = XVALO(!CanAlta) - CANCON#
           SaldoFisico# = SaldoLogico#
           SQLZ$ = " SELECT * FROM RESERVA WITH(NOLOCK) "
           SQLZ$ = SQLZ$ + " WHERE Cod_Inte = " & CodigoInterno% & " "
           SQLZ$ = SQLZ$ + " AND IdenLote = '" & LoteRecepcion$ & "' "
           If TRIM$(ExcluyeReservasOF$) <> "" Then SQLZ$ = SQLZ$ + " AND IdSubOr <> '" & TRIM$(ExcluyeReservasOF$) & "' "
           SQLZ$ = SQLZ$ + " AND Status = 0 "     ' para no considerar los marcados para anular
           SQLZ$ = SQLZ$ + " ORDER BY FechRes ASC"
           Set ReserTemp = READSET(SQLZ$)
           With ReserTemp
              On Error Resume Next
              .MoveFirst
              If Err < 1 Then
                 On Error GoTo 0
                 Do While Not .EOF
                   CantidadSalida# = XVALO(!CANTRES) - XVALO(!CantCons)
                   If CantidadSalida# >= 0.05 Then
                     SaldoLogico# = SaldoLogico# - CantidadSalida#
                   End If
                 .MoveNext
                  Loop
              End If
              On Error GoTo 0
           End With
           ReserTemp.Close
           Set ReserTemp = Nothing
           '
           If SaldoFisico# >= 0.05 Or SaldoLogico# > 0 Then
              Call REPLA(TTXX$, FORMATNUM$(SaldoFisico#, "F12,1"), 37, 12)
              Call REPLA(TTXX$, FORMATNUM$(SaldoLogico#, "F12,1"), 49, 12)
              JJ& = JJ& + 1
              Call GRAREG("!LILOTSE", TTXX$, JJ&)
           End If
         .MoveNext
         Loop
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
    On Error GoTo 0
    Exit Sub

ControlErrores:
    Call MENSERR(24, "Comunique este Mensaje a la Mesa de Ayuda.\ \" & Err.Description)
    Exit Sub
    '
End Sub

Private Sub ImprimeEtiquetaFracc_Click()

    TituloLista$ = "Lotes Fraccionados"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Lote/A16;Codigo/A24;Descripcion/A36;Cod.Interno/I0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Lote/A16;Codigo/A24;Descripcion/A36;Cod.Interno/I0", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT DISTINCT IDENLOTE, COD_EXTE, DESCRIP, COD_INTE FROM MOVISTO WITH(NOLOCK) WHERE CODIMOVI='FB' AND CANTINGRE > 0"
    Set RsLotesFraccionados = READSET(ConsultaSql$)
    Call Carga_MSF_RecordsetConEco(RsLotesFraccionados, "Lote/A16;Codigo/A24;Descripcion/A36;Cod.Interno/I0", FRMZELECHO.msf2, 1)
    On Error Resume Next
    RsSaldoPorDeposito.Close
    Set RsSaldoPorDeposito = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
    FRMZELECHO.TXTBUSCAR = ""
    FRMZELECHO.Show 1
    
    LoteFraccionado$ = TRIM$(RESP_ZELE_VECT(1))
    If LoteFraccionado$ <> "" Then
        CodigoArticuloInterno% = CInt(RESP_ZELE_VECT(4))
        CantidadLote# = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LoteFraccionado$ & "'"))
        Call PRINSTIKAPRO(CodigoArticuloInterno%, LoteFraccionado$, 1, CantidadLote#)
    End If

End Sub

Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub


Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    NCR$ = Mid(TTXX$, 1, 12)
    COD$ = Mid(TTXX$, 26, 16)
    Ci% = CODINT%(COD$)
    Label23.Caption = NCR$
    Label27.Caption = COD$
    DESC$ = Mid(TTXX$, 43, 64)
    Label6 = DESC$
    CantNom$ = Mid(TTXX$, 14, 10)
    Label21 = CantNom$
    cantidadAprobada$ = TRIM$(Mid$(TTXX$, 162, 10))
    Label41 = cantidadAprobada$
    Label8 = Unimed$(Ci%)
    Label15 = TRIM$(Mid$(TTXX$, 107, 8))
    Label32 = TRIM$(Mid$(TTXX$, 115, 8))
    Text14 = TRIM$(Mid$(TTXX$, 129, 32))
    '
    IDREGIS$ = TRIM$(NCR$)
    SQLX$ = "SELECT * FROM BOLRECEP"
    SQLX$ = SQLX$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDREGIS$ & "'"
    Dim BRTEMP As ADODB.Recordset
    Set BRTEMP = New ADODB.Recordset
    Set BRTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With BRTEMP
        On Error Resume Next
        Text8 = TRIM$(FECHATEX$(CVINT(BRTEMP!FechFab)))
        Text14 = BRTEMP!COD_CLI
        Text12 = TRIM$(FECHATEX$(CVINT(BRTEMP!PLAZO_USO)))
        Text2 = BRTEMP!DepoUbic
        Text17 = BRTEMP!UbicFis
        Label20.Caption = BRTEMP!USUAPRO
        Text6.TEXT = BRTEMP!Raso_Prov
        '
        Text3 = SAFETEXT$(!MARCA)
        Text4 = SAFETEXT$(!NUMEPARTE)
        Text9 = SAFETEXT$(!TRAZ_CLI)
        Text11 = TRIM$(SAFETEXT$(!Tipo))
        If TRIM$(Text11) = "" Then
             Text11 = ATRIB1$(Ci%)
        End If
        Text16 = SAFETEXT$(!CARAC)
        If TRIM$(Text16) = "" Then
             Text16 = ATRIB2$(Ci%)
        End If
        '
        List1.SetFocus
        On Error GoTo 0
    End With
    '
End Sub

Private Sub RegistroCalidadLotesConSaldo_Click()

10  TituloLista$ = "Registro de Calidad sobre Lotes con Saldo"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Lote/A16;Codigo/A20;Descripcion/A24;Estado/A20;BR/I0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Lote/A16;Codigo/A20;Descripcion/A24;Estado/A20;BR/I0", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT BOLRE.IDENLOTE, BOLRE.COD_EXTE, BOLRE.DESCRIPITEM, INFOCALIDAD.STATUS, BOLRE.NUMEBORE FROM (SELECT BOLRECEP.IDENLOTE, BOLRECEP.COD_EXTE, BOLRECEP.DESCRIPITEM, BOLRECEP.NUMEBORE FROM BOLRECEP INNER JOIN LOTINGRE ON BOLRECEP.IDENLOTE = LOTINGRE.IDENLOTE WHERE (LOTINGRE.CANALTA-LOTINGRE.CANCONSU) > 0.05) AS BOLRE LEFT JOIN INFOCALIDAD ON BOLRE.IDENLOTE = INFOCALIDAD.DOCUMENTOORIGEN ORDER BY BOLRE.NUMEBORE, BOLRE.IDENLOTE"
    
    Set Rst = FLEXCONN.Execute(ConsultaSql$)
    Call Carga_MSF_Recordset(Rst, "Lote/A16;Codigo/A20;Descripcion/A24;Estado/A20;BR/I0", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 11000
    FRMZELECHO.Show 1
    
    LoteComponente$ = TRIM(RESP_ZELE_VECT(1))
    CodigoArticuloExterno$ = TRIM(RESP_ZELE_VECT(2))
    CodigoArticuloInterno% = CODINT%(CodigoArticuloExterno$)
    descripcionArticulo$ = TRIM(RESP_ZELE_VECT(3))
    
    If CodigoArticuloInterno% > 0 Then
    
        FILTX$ = LoteComponente$ & ";" & CodigoArticuloExterno$ & ";" & descripcionArticulo$
        formularioCalidad$ = TRIM$(Control("TRAZABI", "REGICAL"))
        Call STDFORM(formularioCalidad$, FILTX$)
        
        GoTo 10
    End If

End Sub

Private Sub Text1_Change()

    Call LIMPLAB
    Call LIMPTEXT
    List1.Clear
    If Text1.TEXT <> "" Then
      List1.Clear
      Label30 = ""
      Label28 = ""
      Label24 = ""
      Label17 = ""
      '
      Call APERBASDAT("STOCKS")
      SQLX$ = "SELECT * FROM BOLRECEP"
      SQLX$ = SQLX$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
      Dim BRTEMP As ADODB.Recordset
      Set BRTEMP = New ADODB.Recordset
      Set BRTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      With BRTEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            NPX% = !NumProv
            Label30 = rasocli$((-1) * NPX%)
            Label28 = !NumRem_Prov
            Label24 = FECHATEX$(CVINT(!FechRecep))
            Label17 = FECHATEX$(CVINT(!FechRecep))
            '
            Do While Not .EOF
               If CVINT(!FeHoApro) > 0 Then
                 CIXI% = XVALO(!Cod_Inte)
                 If ESTRAZABLE(CIXI%) > 0 Then
                    TTXX$ = Space$(160)
                    Call REPLA(TTXX$, TRIM$(!IDENLOTE), 1, 12)
                    Call REPLA(TTXX$, FORMATNUM$(!CantNom, "F10.1"), 14, 10)
                    Call REPLA(TTXX$, TRIM$(CODEXT$(CIXI%)), 26, 16)
                    Call REPLA(TTXX$, TRIM$(DESCRIT$(CIXI%)), 43, 64)
                      ANCHOX$ = FORMATNUM$(XVALO(!M_Ancho), "F8.1")
                    Call REPLA(TTXX$, ANCHOX$, 107, 8)
                      LARGOX$ = FORMATNUM$(XVALO(!M_Largo), "F8.1")
                    Call REPLA(TTXX$, LARGOX$, 115, 8)
                      LOTORI$ = SAFETEXT$(!LOTECOL_PROV)
                    Call REPLA(TTXX$, LOTORI$, 129, 32)
                      cantidadAprobada$ = TRIM$(FORMATNUM$(!CANTOAPRO, "F10.1"))
                    Call REPLA(TTXX$, cantidadAprobada$, 162, 10)
                    List1.AddItem TTXX$
                 End If
               End If
            .MoveNext
            Loop
         End If
      End With
      '
      If List1.ListCount < 1 Then
        If Command17.Enabled = True Then 'SI EL BOTON DE PROCESO AUTOMATICO ESTA DESHABILITADO NO LO MUESTRA
         Call MENSERR(24, "La BR Elegida no Registra Ingresos\Aprobados de Materiales Trazables\que Puedan Procesarse.")
        End If
         Text1 = ""
         Label30 = ""
         Label28 = ""
         Label24 = ""
         Label17 = ""
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub

Private Sub Text1_DblClick()

    Command5.Enabled = False
    Call Command5_Click
'3   Call SELECHO("INDIBOL")
'    NPX& = Val(VALACT1$("INDIBOL"))
'    If NPX& > 0 Then
'       Text1.TEXT = NPX&
'    End If
99  Command5.Enabled = True

End Sub

Sub LIMPLAB()

    Label27.Caption = ""
    Label6.Caption = ""
    Label21.Caption = ""
    Label23.Caption = ""
    Label20.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
  
End Sub

Sub LIMPTEXT()
    
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
    
  
End Sub

Private Sub Text12_DblClick()

    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
       Text12.TEXT = VDEV$
    End If
    Text17.SetFocus
    
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command2.SetFocus
   End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text17.SetFocus
    End If
    
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text8.SetFocus
    End If

End Sub

Private Sub Text2_Change()
   Label34 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text2)))
End Sub

Private Sub Text3_DblClick()
   Call Command11_Click
End Sub


Private Sub Text4_DblClick()
    Call Command10_Click
End Sub



Private Sub Text8_DblClick()

    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
       Text8.TEXT = VDEV$
    End If

End Sub

Sub LICONLOT()
    ' LISTADO DE CONSUMO POR LOTE DE RECEPCIÓN
    '
    Dim VALOU(128)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) WHERE CANSALDO IS NULL" 'Realiza un update para aquello registros que tengan CanSaldo = Null
    Set LOTINTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTINTMP
       If Not (.EOF And .BOF) Then
          Do While .EOF = False
            CIXI% = !Cod_Inte
            IDLOT$ = SAFETEXT$(!IDENLOTE)
            Call ACSALOTE(CIXI%, IDLOT$, "NOMESS")
          .MoveNext
          Loop
       End If
    End With
    LOTINTMP.Close
    Set LOTINTMP = Nothing
    '
    MODOPAPEL% = 0
    If EXISTE%(LUDAT$ + "LICONLOP.RFS") > 0 Then MODOPAPEL% = 1
    
10  Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    Desde = CVDAT(Des%)
    Hasta = CVDAT(Has%)
    REBLA$ = REGBLAN$("LICONLOT")
    REBLB$ = "": If MODOPAPEL% = 1 Then REBLB$ = REGBLAN$("LICONLOP")
    JJ& = 0
    '
    OPXX% = COMALTER%("Formato de salida:\\Detalle\Resumen")
    If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
    '
    Call APERBASDAT("STOCKS")
    Screen.MousePointer = 1
20  SQLX$ = "SELECT *, BOLRECEP.*, LOTINGRE.* FROM LOTINGRE WITH(NOLOCK) "
    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE FechAlta >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND FechAlta <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " ORDER BY LOTINGRE.FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINSEL As ADODB.Recordset
    Set LOTINSEL = New ADODB.Recordset
    Set LOTINSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))     ' Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
    With LOTINSEL
       If (.EOF And .BOF) Then
          Call MENSERR(24, "No Hay Lotes Ingresados\en Rango de Fechas Especificado.")
          GoTo 10
       End If
       '
       .MoveFirst
       KU& = 0: IDLOA$ = ""
       Do While Not .EOF
          KU& = KU& + 1
          Call TRACE(20, KU&, FIN&)
          FALTA% = CVINT(!FechAlta)
          Tipo$ = UCase$(Left$(!IDENLOTE, 2))
          IDLO$ = TRIM$(!IDENLOTE)
          TRACLI$ = ""                           ' FALTA TRAZABILIDAD CLIENTE
'          MARCAX$ = SAFETEXT$(!Marca)
'          LOFABX$ = !LoteCol_Prov
          NPRX% = !NumProv
          PROVE$ = FANTCLI$((-1) * NPRX%)
          CIXI% = !Cod_Inte
          ANCHX = !AnchoMili
          LARGX = 0
          CANTIN# = !CanAlta
          CSALDO# = !CanSaldo
          METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CSALDO / ANCHX
          DEPX% = XVALO(!DepoUbic)
          UBIFI$ = SAFETEXT$(!UbicFis)
          FEFA% = CVINT(!FechFab)
          FEVE% = CVINT(!FechVenc)
          '
          XX0$ = REBLA$
          JJ0& = JJ&
          SUCONSU# = 0: SURESER# = 0: SALMOVI# = 0
          Call REPLA(XX0$, MKI$(FALTA%), 1, 2)
          Call REPLA(XX0$, Tipo$, 3, 2)
          Call REPLA(XX0$, IDLO$, 5, 16)
          Call REPLA(XX0$, TRACLI$, 21, 32)
          Call REPLA(XX0$, PROVE$, 53, 16)
          Call REPLA(XX0$, MKI$(CIXI%), 69, 2)
          Call REPLA(XX0$, MKD$(CANTIN#), 71, 8)
          Call REPLA(XX0$, MKD$(CSALDO#), 79, 8)
          '
          If MODOPAPEL% = 1 Then
             YY0$ = REBLB$
             Call REPLA(YY0$, MKI$(FALTA%), 1, 2)
             Call REPLA(YY0$, Tipo$, 3, 2)
             Call REPLA(YY0$, IDLO$, 5, 16)
             Call REPLA(YY0$, MARCAX$, 21, 16)
             Call REPLA(YY0$, LOFABX$, 37, 16)
             Call REPLA(YY0$, PROVE$, 53, 16)
             Call REPLA(YY0$, MKI$(CIXI%), 69, 2)
             Call REPLA(YY0$, Unimed$(CIXI%), 71, 4)
             Call REPLA(YY0$, MKS$(ANCHX), 75, 4)
             Call REPLA(YY0$, MKD$(CANTIN#), 79, 8)
             Call REPLA(YY0$, MKD$(CSALDO#), 87, 8)
             Call REPLA(YY0$, MKS$(METSAL), 95, 4)
             Call REPLA(YY0$, MKI$(DEPX%), 99, 2)
             Call REPLA(YY0$, UBIFI$, 101, 24)
             Call REPLA(YY0$, MKI$(FEFA%), 125, 2)
             Call REPLA(YY0$, MKI$(FEVE%), 127, 2)
          End If
          '
          Call CARDETMOVLOT(CIXI%, IDLO$)
          '
          For KKI% = 1 To CAMOVILO%
             CCONSU# = CANINGLO#(KKI%) - CANSALLO#(KKI%)
             If Abs(CCONSU#) >= 0.05 Then
                XX$ = XX0$
                Call REPLA(XX$, MKI$(FEMOVILO%(KKI%)), 95, 2)
                Call REPLA(XX$, DEMOVILO$(KKI%), 97, 24)
                Call REPLA(XX$, MKD$((-1) * CCONSU#), 153, 8)
                SUCONSU# = SUCONSU# + CCONSU#
                JJ& = JJ& + 1
                Call GRAREG("LICONLOT", XX$, JJ&)
                '
                If MODOPAPEL% = 1 Then
                   YY$ = YY0$
                   Call REPLA(YY$, MKI$(FEMOVILO%(KKI%)), 129, 2)
                   Call REPLA(YY$, DEMOVILO$(KKI%), 131, 24)
                   METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CCONSU / ANCHX
                   If CCONSU# > 0 Then
                       Call REPLA(YY$, MKD$(CCONSU#), 187, 8)
                       Call REPLA(YY$, MKS$(METSAL), 195, 4)
                     Else
                       Call REPLA(YY$, MKD$((-1) * CCONSU#), 199, 8)
                       Call REPLA(YY$, MKS$(Abs(METSAL)), 207, 4)
                   End If
                   SALMOVI# = SALMOVI# + CCONSU#
                   Call GRAREG("LICONLOP", YY$, JJ&)
                End If
                '
             End If
          Next KKI%
          '
          For KKI% = 1 To CAMOVILO%
             CCONSU# = ROUND(CANRESLO#(KKI%), 1)
             If Abs(CCONSU#) >= 0.05 Then
                XX$ = XX0$
                Call REPLA(XX$, MKI$(FEMOVILO%(KKI%)), 95, 2)
                Call REPLA(XX$, DEMOVILO$(KKI%), 97, 24)
                    '    ' aca faltan los numeros PNL/PLC
                Call REPLA(XX$, MKD$(CCONSU#), 169, 8)
                SURESER# = SURESER# + CCONSU#
                JJ& = JJ& + 1
                Call GRAREG("LICONLOT", XX$, JJ&)
                '
                If MODOPAPEL% = 1 Then
                   YY$ = YY0$
                   Call REPLA(YY$, MKI$(FEMOVILO%(KKI%)), 129, 2)
                   Call REPLA(YY$, DEMOVILO$(KKI%), 131, 24)
                   METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CCONSU / ANCHX
                   Call REPLA(YY$, MKD$(CCONSU#), 223, 8)
                   Call REPLA(YY$, MKS$(METSAL), 231, 4)
                   Call GRAREG("LICONLOP", YY$, JJ&)
                End If
                '
             End If
          Next KKI%
          '
          If JJ& <= JJ0& Then
             XX$ = XX0$
             JJ& = JJ& + 1
             Call GRAREG("LICONLOT", XX$, JJ&)
             '
             If MODOPAPEL% = 1 Then
                YY$ = YY0$
                Call GRAREG("LICONLOP", YY$, JJ&)
             End If
          End If
          '
          If MODOPAPEL% = 1 Then
             METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * SALMOVI# / ANCHX
             Call REPLA(YY$, MKD$(SALMOVI#), 211, 8)
             Call REPLA(YY$, MKS$(METSAL), 219, 4)
             METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * SURESER# / ANCHX
             Call REPLA(YY$, MKD$(SURESER#), 235, 8)
             Call REPLA(YY$, MKS$(METSAL), 243, 4)
             Call GRAREG("LICONLOP", YY$, JJ&)
          End If
       .MoveNext
       Loop
    End With
    '
    Call SETULTREG("LICONLOT", JJ&)
    Screen.MousePointer = 1
    '
    If OPXX% = 1 Then
         ACONEC$ = "LICONLOT": If MODOPAPEL% = 1 Then ACONEC$ = "LICONLOP"
         Call HEADERS(ACONEC$, "")
         Call HEADERS(ACONEC$, "Ingresos Correspondientes a: " + PERIO$)
         Call CONECRUN("*" + ACONEC$, "Listado Detalle de Consumos por Lote.")
      Else
         JJ1& = 0
         For KKU& = 1 To JJ&
            XX$ = REGLEIDO$("LICONLOT", KKU&)
            If Left$(XX$, 70) = Left$(zz$, 70) Then
                SUCONSU# = CVD(Mid$(XX$, 153, 8)) + CVD(Mid$(zz$, 153, 8))
                SURESER# = CVD(Mid$(XX$, 169, 8)) + CVD(Mid$(zz$, 169, 8))
                Call REPLA(zz$, MKD$(SUCONSU#), 153, 8)
                Call REPLA(zz$, MKD$(SURESER#), 169, 8)
              Else
                zz$ = XX$
                JJ1& = JJ1& + 1
            End If
            Call GRAREG("RECONLOT", zz$, JJ1&)
         Next KKU&
         Call SETULTREG("RECONLOT", JJ1&)
         '
         'ESTO ES PARA TRASLADAR EL CÓDIGO DE ARTICULO A LA PRIMERA COLUMNA
         For U& = 1 To JJ1&
           XX$ = REGLEIDO$("RECONLOT", U&)
           XX1$ = Mid$(XX$, 69, 2) + Left$(XX$, 2)
           Call REPLA(XX$, XX1$, 1, 4)
           Call GRAREG("RECONLOT", XX$, U&)
         Next U&
         Call CIERRARCH("RECONLOT")
         '
         Call FILESORT("RECONLOT", "", 1, 3)
         Call HEADERS("RECONLOT", "")
         Call HEADERS("RECONLOT", "Ingresos Correspondientes a: " + PERIO$)
         Call CONECRUN("*RECONLOT", "Listado Resumen de Consumos por Lote.")
    End If
    '
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text12.SetFocus
   End If
End Sub

Sub FRABOBIN()
    '
    FRABOBI07.Show 1
    '
End Sub

Sub CONLOTES()
    
    OPX% = ALTERNA%("Diagnostico\Recuperación")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
       
    Call ABRECONEXION
    Dim ControlLotes As ADODB.Recordset
    Set ControlLotes = New ADODB.Recordset
    SQLX$ = "SELECT COD_INTE, IDENLOTE, SUM(CANTINGRE) AS INGRE, SUM(CANTSALID) AS SALID FROM MOVISTO GROUP BY COD_INTE, IDENLOTE"
    Set ControlLotes = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    INFORME$ = ""
    '
    If Not (ControlLotes.EOF And ControlLotes.BOF) Then
       '
       Screen.MousePointer = 11
       Do While Not (ControlLotes.EOF)
         '
         LOTE$ = TRIM$(SAFETEXT(ControlLotes!IDENLOTE))
         CIXI% = ControlLotes!Cod_Inte
         CONDI$ = "IDENLOTE ='" & LOTE$ & "' and cod_inte =" & CIXI%
         ALTA% = XVALO(VALOBADA("MASTER", "STATUS", "CODINT = " & CIXI%))
         '
         If ALTA% <> 1 Then GoTo 10
         '
         If CantRegistros("LOTINGRE", CONDI$) > 0 Then
            Saldo# = XVALO(ControlLotes!Ingre) - XVALO(ControlLotes!SALID)
            sallote# = VALOBADA("LOTINGRE", "CANSALDO", CONDI$)
            '
            LARGONU# = 0: SALME# = 0
            ANCHO# = VALOBADA("LOTINGRE", "AnchoMili", CONDI$, "NOMESS")
            If ANCHO# > 0 Then
               LARGONU# = 1000 * Saldo# / ANCHO# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
               SALME# = VALOBADA("LOTINGRE", "SalMetro", CONDI$, "NOMESS")
            End If
            '
            If Saldo# - sallote# > 0.05 Then
               '
               If OPX% = 2 Then
                  Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
               End If
               '
               Mensaje$ = "Cantidad del lote " & LOTE$ & " No Coincide para Codigo " & CODEXT$(CIXI%)
               INFORME$ = INFORME$ + Mensaje$ + vbCrLf
               '
            End If
            '
            If SALME# - LARGONU# > 0.05 Then
               '
               If OPX% = 2 Then
                  Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
               End If
               '
               If Saldo# > 0.05 Then
                  Mensaje$ = "Largo del lote " & LOTE$ & " No Coincide para Codigo " & CODEXT$(CIXI%)
                  INFORME$ = INFORME$ + Mensaje$ + vbCrLf
               End If
               '
            End If
            '
         Else
           If OPX% = 2 Then
             'Recuperación
             FF% = CVINT%(VALOBADA("MOVISTO", "FECHMOV", CONDI$ & " ORDER BY FECHMOV ASC"))
             CANOM# = VALOBADA("MOVISTO", "CANTINGRE", CONDI$ & " ORDER BY FECHMOV ASC")
             DPX% = VALOBADA("MOVISTO", "DEPOMOVI", CONDI$ & " ORDER BY FECHMOV ASC")
             RPROVED$ = rasocli(VALOBADA("MOVISTO", "NUME_PROV", CONDI$ & " ORDER BY FECHMOV ASC") * -1)
             Call CREALOTIN(FF%, FF%, FF%, CIXI%, LOTE$, MARCAX$, "", CANOM#, 0, 0, DPX%, "", RPROVED$)
             Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
           End If
           '
           'Diagnostico
           Mensaje$ = "Lote " & LOTE$ & " Inexistente para Codigo " & CODEXT$(CIXI%)
           INFORME$ = INFORME$ + Mensaje$ + vbCrLf
           'Call COMUNI(MENSAJE$)
         '
         End If
10       ControlLotes.MoveNext
       Loop
       '
       Screen.MousePointer = 1
       Call SaveFile("C:\FLEXOFT\INFORME.LOG", INFORME$)
       MsgBox ("Informe Guardado como C:\FLEXOFT\INFORME.LOG")
       '
    End If
99     Screen.MousePointer = 1
  
End Sub

