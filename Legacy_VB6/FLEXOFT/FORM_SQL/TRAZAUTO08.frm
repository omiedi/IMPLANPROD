VERSION 5.00
Begin VB.Form TRAZAUTO08 
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFD0&
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
      Left            =   315
      TabIndex        =   59
      Top             =   6930
      Width           =   5055
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
         TabIndex        =   63
         Top             =   465
         Width           =   1335
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
         TabIndex        =   62
         Top             =   465
         Width           =   1335
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
         TabIndex        =   61
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1125
      End
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
         TabIndex        =   60
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1125
      End
   End
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
      Height          =   690
      Left            =   8505
      Picture         =   "TRAZAUTO08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   5670
      Width           =   800
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   5580
      TabIndex        =   33
      Top             =   1440
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   1870
      TabIndex        =   32
      Top             =   1470
      Width           =   15
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   25
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
      Left            =   2310
      TabIndex        =   24
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   210
      TabIndex        =   21
      Top             =   1155
      Width           =   8100
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   2730
         TabIndex        =   31
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   885
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   580
         Width           =   7680
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IdenLote        Cantid.  Código                   Descripción"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   30
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5625
      Left            =   8505
      ScaleHeight     =   5565
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1380
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Imprime Sticker de Aprobación"
         Top             =   3450
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1500
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4100
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2800
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   850
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":16D4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Listado de Informe de Trazabilidad"
         Top             =   2150
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   4890
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO08.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRAZAUTO08.frx":1B28
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5190
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Trazabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   5
      Top             =   2900
      Width           =   8100
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
         TabIndex        =   56
         Top             =   2310
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   945
         TabIndex        =   53
         Top             =   2310
         Width           =   540
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   6195
         TabIndex        =   41
         Top             =   1155
         Width           =   1695
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   13
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
            TabIndex        =   14
            Top             =   1785
            Width           =   175
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   315
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   1155
            Width           =   175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento:"
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
            Left            =   105
            TabIndex        =   45
            Top             =   1575
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricación:"
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
            Left            =   105
            TabIndex        =   44
            Top             =   945
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso:"
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
            Left            =   105
            TabIndex        =   43
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label17 
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
            Left            =   315
            TabIndex        =   42
            Top             =   525
            Width           =   1020
         End
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4200
         TabIndex        =   10
         Top             =   2310
         Width           =   1800
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   9
         Top             =   1575
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   8
         Top             =   1575
         Width           =   2220
      End
      Begin VB.Label Label34 
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
         Left            =   1785
         TabIndex        =   55
         Top             =   2310
         Width           =   2175
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   -315
         TabIndex        =   54
         Top             =   2085
         Width           =   1755
      End
      Begin VB.Label Label8 
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
         Left            =   4410
         TabIndex        =   52
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         Left            =   3045
         TabIndex        =   51
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   5250
         TabIndex        =   40
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label21 
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
         Left            =   6615
         TabIndex        =   37
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Label6 
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
         Left            =   945
         TabIndex        =   36
         ToolTipText     =   "Descripción"
         Top             =   760
         Width           =   6900
      End
      Begin VB.Label Label27 
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
         Left            =   945
         TabIndex        =   35
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   2655
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -420
         TabIndex        =   34
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label23 
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
         Left            =   4725
         TabIndex        =   22
         Top             =   2940
         Width           =   1230
      End
      Begin VB.Label Label20 
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
         Left            =   945
         TabIndex        =   20
         Top             =   2940
         Width           =   2910
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion Física:"
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
         Left            =   3465
         TabIndex        =   17
         Top             =   2085
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.R.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2730
         TabIndex        =   16
         Top             =   3000
         Width           =   1965
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote de Fabricación:"
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
         Left            =   2625
         TabIndex        =   15
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Left            =   -105
         TabIndex        =   7
         Top             =   2730
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   -105
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   105
      TabIndex        =   47
      Top             =   800
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Left            =   1365
      TabIndex        =   46
      Top             =   735
      Width           =   4590
   End
   Begin VB.Label Label28 
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
      Left            =   6405
      TabIndex        =   39
      Top             =   315
      Width           =   1860
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nro:"
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
      Left            =   4725
      TabIndex        =   38
      Top             =   380
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   5775
      TabIndex        =   28
      Top             =   800
      Width           =   1335
   End
   Begin VB.Label Label24 
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
      Left            =   7245
      TabIndex        =   27
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B.R. Nº:"
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
      Left            =   105
      TabIndex        =   26
      Top             =   380
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
      TabIndex        =   19
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
      TabIndex        =   18
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "TRAZAUTO08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call TRAZORF08.Command14_Click
    Command14.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    ' LISTADO DE CONSUMO POR LOTE DE RECEPCIÓN
    Call LICONLOT
    Command16.Enabled = True
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
   If ECOARCH$("INDIBOL", MKS$(NUBORE&)) = "" Then
      Call MENSERR(24, "Número de Boleta de Recepción no Válido.")
      GoTo 99
   End If
   '
   CIXI% = CODINT%(Label27)
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido")
      GoTo 99
   End If
   '
   If FECHANUM(Text8) > FECHANUM(Label17) Then
      Call MENSERR(24, "La Fecha de Fabricación no Puede Ser\Posterior a la Fecha de Ingreso.")
      GoTo 99
   End If
   '
   If FECHANUM(Text12) < FECHANUM(Text8) Then
      Call MENSERR(24, "La Fecha de Vencimiento no Puede Ser\Anterior a la Fecha de Fabricación.")
      GoTo 99
   End If
   '
   If FECHANUM(Text12) < HOY(HOS$) Then
      Call MENSERR(24, "Este Material ya Está Vencido\al Día de Hoy.")
   End If
   '
   If TRIM$(Text17.TEXT) = "" Then
    Call MENSERR(24, "Ubicacion Física no Válida.")
    GoTo 99
   End If
   '
   Dim BRTEMP As ADODB.Recordset
   For KKI% = 1 To List1.ListCount
      If List1.Selected(KKI% - 1) = True Then
         IDREGIS$ = TRIM$(Left$(List1.List(KKI% - 1), 12))
         Set BRTEMP = New ADODB.Recordset
         sqlx$ = " SELECT * FROM BOLRECEP where IdenLote = '" & IDREGIS$ & "'"
         BRTEMP.OPEN FILTSQL$ (sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         '
         With BRTEMP
            If Not (.EOF And .BOF) Then
               !MARCA = Left$(Text5, 32)
               !FECHFAB = CVDAT(FECHANUM(Left$(Text8, 8)))
               !LoteCol_Prov = Left$(Text14, 16)
                   FEVEN% = FECHANUM(Text12)
               !PLAZO_USO = CVDAT(FEVEN%)
               !DepoUbic = XVALO(Text2)
               !UbicFis = Left$(Text17, 32)
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
               MARCAX$ = Left$(Text5, 32)
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
        OPX% = ALTERNA%("Fraccionamiento de Bobinas\Otros Accesos Directos")
        If OPX% = 1 Then
            Call FRABOBIN
          ElseIf OPX% = 2 Then
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
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
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

   Command5.Enabled = False
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
      Text1.TEXT = NPX&
   End If
99 Command5.Enabled = True

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
   Call TRAZORF08.Command10_Click
   Command8.Enabled = True
End Sub
'
Private Sub Command9_Click()
    Command9.Enabled = False
    CI% = CODINT%(Label27)
    If CI% > 0 Then
      LOTE$ = TRIM$(Label23)
      If LOTE$ <> "" Then
        DEPO% = XVALO(Text2)
        If DEPO% > 0 Then
          Call PRINSTIKAPRO(CI%, LOTE$, DEPO%)
        End If
      End If
    End If
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If Left$(APLINVO$, 8) = "TRAZORFA" Or Left$(APLINVO$, 7) = "OFABRIC" Or INTERPRE% = 1 Then
       TRAZORF08.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "ORFCIE" Then   ' Or INTERPRE = 1 Then
       CIERRAOF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       OPENFORMS = DoEvents
    End If
    '
End Sub

Sub VERIPERPRUE()
    '
    If CONTROL$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(CONTROL$("TRAZABI", "FINICIO"))
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
'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub



Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    NCR$ = Mid(TTXX$, 1, 12)
    COD$ = Mid(TTXX$, 25, 24)
    CI% = CODINT%(COD$)
    Label23.Caption = NCR$
    Label27.Caption = COD$
    DESC$ = Mid(TTXX$, 50, 64)
    Label6 = DESC$
    CantNom$ = Mid(TTXX$, 14, 10)
    Label21 = CantNom$
    Label8 = UNIMED$(CI%)
    Label15 = TRIM$(Mid$(TTXX$, 115, 8))
    Label32 = TRIM$(Mid$(TTXX$, 123, 8))
    Text14 = TRIM$(Mid$(TTXX$, 135, 32))
    '
    IDREGIS$ = TRIM$(NCR$)
    sqlx$ = "SELECT * FROM BOLRECEP"
    sqlx$ = sqlx$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
    sqlx$ = sqlx$ + " AND IdenLote = '" & IDREGIS$ & "'"
    Dim BRTEMP As ADODB.Recordset
    Set BRTEMP = New ADODB.Recordset
    Set BRTEMP =  FLEXCONN.EXECUTE (FILTSQL$ (sqlx$))
    '
    With BRTEMP
        On Error Resume Next
        Text5 = BRTEMP!MARCA
        Text8 = TRIM$(FECHATEX$(CVINT(BRTEMP!FECHFAB)))
        Text14 = BRTEMP!LoteCol_Prov
        Text12 = TRIM$(FECHATEX$(CVINT(BRTEMP!PLAZO_USO)))
        Text2 = BRTEMP!DepoUbic
        Text17 = BRTEMP!UbicFis
        Label20.Caption = BRTEMP!USUAPRO
        Text6.TEXT = BRTEMP!RASO_PROV
        '
        List1.SetFocus
        On Error GoTo 0
    End With
    '
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
      sqlx$ = "SELECT * FROM BOLRECEP"
      sqlx$ = sqlx$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
      Dim BRTEMP As ADODB.Recordset
      Set BRTEMP = New ADODB.Recordset
      Set BRTEMP =  FLEXCONN.EXECUTE (FILTSQL$ (sqlx$))
      '
      With BRTEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            NPX% = !NumProv
            Label30 = RASOCLI$((-1) * NPX%)
            Label28 = !NumRem_Prov
            Label24 = FECHATEX$(CVINT(!FechRecep))
            Label17 = FECHATEX$(CVINT(!FechRecep))
            '
            Do While Not .EOF
               If CVINT(!FeHoApro) > 0 Then
                 CIXI% = !Cod_Inte
                 If ESTRAZABLE(CIXI%) > 0 Then
                    TTXX$ = Space$(160)
                    Call REPLA(TTXX$, TRIM$(!IdenLote), 1, 12)
                    Call REPLA(TTXX$, FORMATNUM$(!CantNom, "F10.1"), 14, 10)
                    Call REPLA(TTXX$, TRIM$(!COD_EXTE), 25, 24)
                    Call REPLA(TTXX$, TRIM$(!DESCRIPITEM), 50, 64)
                      ANCHOX$ = FORMATNUM$(XVALO(!M_Ancho), "F8.1")
                    Call REPLA(TTXX$, ANCHOX$, 115, 8)
                      LARGOX$ = FORMATNUM$(XVALO(!M_Largo), "F8.1")
                    Call REPLA(TTXX$, LARGOX$, 123, 8)
                      LOTORI$ = SAFETEXT$(!LoteCol_Prov)
                    Call REPLA(TTXX$, LOTORI$, 135, 32)
                    List1.AddItem TTXX$
                 End If
               End If
            .MoveNext
            Loop
         End If
      End With
      '
      If List1.ListCount < 1 Then
         Call MENSERR(24, "La BR Elegida no Registra Ingresos\Aprobados de Materiales Trazables\que Puedan Procesarse.")
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
3   Call SELECHO("INDIBOL")
    NPX& = Val(VALACT1$("INDIBOL"))
    If NPX& > 0 Then
       Text1.TEXT = NPX&
    End If
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
    Text5 = ""
    Text3 = ""
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

Private Sub Text17_DblClick()
    GENUBIDE.Show 1
    UBICA$ = TRIM$(GENUBIDE.Label4)
    If UBICA$ <> "" Then
       Text17 = UBICA$
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text14.SetFocus
   End If
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
    MODOPAPEL% = 0
    If EXISTE%(LUDAT$ + "LICONLOP.RFS") > 0 Then MODOPAPEL% = 1
    
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    DESDE = CVDAT(DES%)
    HASTA = CVDAT(HAS%)
    REBLA$ = REGBLAN$("LICONLOT")
    REBLB$ = "": If MODOPAPEL% = 1 Then REBLB$ = REGBLAN$("LICONLOP")
    JJ& = 0
    '
    OPXX% = COMALTER%("Formato de salida:\\Detalle\Resumen")
    If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
    '
    Call APERBASDAT("STOCKS")
    Screen.MousePointer = 1
20  sqlx$ = "SELECT *, BOLRECEP.*, LOTINGRE.* FROM LOTINGRE WITH(NOLOCK) "
    sqlx$ = sqlx$ + " INNER JOIN BOLRECEP"
    sqlx$ = sqlx$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    sqlx$ = sqlx$ + " WHERE FechAlta >= '" & DESDE & "' "
    sqlx$ = sqlx$ + " AND FechAlta <= '" & HASTA & "' "
    sqlx$ = sqlx$ + " ORDER BY LOTINGRE.FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINSEL As ADODB.Recordset
    Set LOTINSEL = New ADODB.Recordset
    Set LOTINSEL =  FLEXCONN.EXECUTE (FILTSQL$ (sqlx$))  ' Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
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
          TIPO$ = UCase$(Left$(!IdenLote, 2))
          IDLO$ = TRIM$(!IdenLote)
          TRACLI$ = ""                           ' FALTA TRAZABILIDAD CLIENTE
'          MARCAX$ = SAFETEXT$(!Marca)
'          LOFABX$ = !LoteCol_Prov
          NPRX% = !NumProv
          PROVE$ = FANTCLI$((-1) * NPRX%)
          CIXI% = !Cod_Inte
          ANCHX = !AnchoMili
          LARGX = 0
          CANTIN# = !CanAlta
          CSALDO# = !CANSALDO
          METSAL = 0: If ANCHX > 0.1 Then METSAL = 1000 * CSALDO / ANCHX
          DEPX% = XVALO(!DepoUbic)
          UBIFI$ = SAFETEXT$(!UbicFis)
          FEFA% = CVINT(!FECHFAB)
          FEVE% = CVINT(!FechVenc)
          '
          XX0$ = REBLA$
          JJ0& = JJ&
          SUCONSU# = 0: SURESER# = 0: SALMOVI# = 0
          Call REPLA(XX0$, MKI$(FALTA%), 1, 2)
          Call REPLA(XX0$, TIPO$, 3, 2)
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
             Call REPLA(YY0$, TIPO$, 3, 2)
             Call REPLA(YY0$, IDLO$, 5, 16)
             Call REPLA(YY0$, MARCAX$, 21, 16)
             Call REPLA(YY0$, LOFABX$, 37, 16)
             Call REPLA(YY0$, PROVE$, 53, 16)
             Call REPLA(YY0$, MKI$(CIXI%), 69, 2)
             Call REPLA(YY0$, UNIMED$(CIXI%), 71, 4)
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
            If Left$(XX$, 70) = Left$(ZZ$, 70) Then
                SUCONSU# = CVD(Mid$(XX$, 153, 8)) + CVD(Mid$(ZZ$, 153, 8))
                SURESER# = CVD(Mid$(XX$, 169, 8)) + CVD(Mid$(ZZ$, 169, 8))
                Call REPLA(ZZ$, MKD$(SUCONSU#), 153, 8)
                Call REPLA(ZZ$, MKD$(SURESER#), 169, 8)
              Else
                ZZ$ = XX$
                JJ1& = JJ1& + 1
            End If
            Call GRAREG("RECONLOT", ZZ$, JJ1&)
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


