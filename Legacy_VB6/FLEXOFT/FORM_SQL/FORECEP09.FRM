VERSION 5.00
Begin VB.Form FORECEP09 
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Recepción"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   11910
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
      Height          =   1275
      Left            =   11040
      Picture         =   "FORECEP09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   6000
      Width           =   760
   End
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
      Left            =   10725
      TabIndex        =   39
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   1
      Top             =   735
      Width           =   2310
   End
   Begin VB.TextBox Text3 
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
      Left            =   5325
      TabIndex        =   2
      Top             =   735
      Width           =   2370
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1170
      Index           =   1
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   2200
      Left            =   1995
      TabIndex        =   29
      Top             =   1470
      Width           =   15
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   1365
      MaxLength       =   5
      TabIndex        =   0
      Top             =   315
      Width           =   720
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
      Left            =   2040
      TabIndex        =   23
      Top             =   300
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
      Height          =   2630
      Left            =   210
      TabIndex        =   22
      Top             =   1155
      Width           =   10725
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   2200
         Left            =   9660
         TabIndex        =   37
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   2200
         Left            =   8610
         TabIndex        =   36
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame9"
         Height          =   2200
         Left            =   7665
         TabIndex        =   35
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2200
         Left            =   6510
         TabIndex        =   34
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6090
         TabIndex        =   28
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
         Height          =   1875
         ItemData        =   "FORECEP09.frx":030A
         Left            =   105
         List            =   "FORECEP09.frx":030C
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         TabIndex        =   26
         Top             =   630
         Width           =   10515
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FORECEP09.frx":030E
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
         Left            =   105
         TabIndex        =   27
         Top             =   380
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5970
      Left            =   11025
      ScaleHeight     =   5910
      ScaleWidth      =   1320
      TabIndex        =   16
      Top             =   -50
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":039F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprime Sticker de Aprobación"
         Top             =   3650
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Anul"
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
         Picture         =   "FORECEP09.frx":0A09
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Anulacion Recepciones Pendientes de Aprobar"
         Top             =   1580
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":0D13
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4340
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":1155
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2960
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":145F
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   890
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":1769
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Listado de Informe de Trazabilidad"
         Top             =   2270
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   5200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   -30
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
         Height          =   690
         Left            =   105
         Picture         =   "FORECEP09.frx":1A73
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FORECEP09.frx":1BBD
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5500
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
      Height          =   3430
      Left            =   210
      TabIndex        =   19
      Top             =   3840
      Width           =   10725
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Recepcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   7575
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
            Height          =   315
            Left            =   3075
            TabIndex        =   69
            Top             =   270
            Width           =   175
         End
         Begin VB.TextBox Text6 
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
            Height          =   285
            Left            =   6075
            MaxLength       =   9
            TabIndex        =   68
            Top             =   285
            Width           =   1230
         End
         Begin VB.TextBox Text8 
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
            Left            =   795
            TabIndex        =   67
            Top             =   285
            Width           =   2325
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
            Left            =   -600
            TabIndex        =   74
            Top             =   360
            Width           =   1335
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
            Left            =   780
            TabIndex        =   73
            ToolTipText     =   "Descripción"
            Top             =   630
            Width           =   6525
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
            Left            =   3090
            TabIndex        =   72
            Top             =   360
            Width           =   1125
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
            Left            =   4290
            TabIndex        =   71
            ToolTipText     =   "Cantidad nominal"
            Top             =   285
            Width           =   630
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "QT Recib."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5100
            TabIndex        =   70
            Top             =   360
            Width           =   945
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Saldos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   7800
         TabIndex        =   61
         Top             =   1800
         Width           =   1680
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Q.P.E.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -60
            TabIndex        =   65
            Top             =   285
            Width           =   840
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
            Left            =   480
            TabIndex        =   64
            Top             =   525
            Width           =   1110
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F.S.E.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -60
            TabIndex        =   63
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
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
            Left            =   480
            TabIndex        =   62
            ToolTipText     =   "Código"
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1720
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   4815
         Begin VB.TextBox Text10 
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
            Left            =   2160
            TabIndex        =   83
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton Command18 
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
            Left            =   4500
            TabIndex        =   78
            Top             =   900
            Width           =   175
         End
         Begin VB.TextBox Text9 
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
            Left            =   3240
            TabIndex        =   75
            Top             =   900
            Width           =   1290
         End
         Begin VB.TextBox Text14 
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
            Left            =   1080
            TabIndex        =   54
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bultos:"
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
            Left            =   680
            TabIndex        =   84
            Top             =   280
            Width           =   1455
         End
         Begin VB.Label Label29 
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
            Left            =   240
            TabIndex        =   77
            ToolTipText     =   "Cantidad nominal"
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
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
            Left            =   2640
            TabIndex        =   76
            Top             =   640
            Width           =   1125
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
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
            Left            =   1065
            TabIndex        =   59
            ToolTipText     =   "Cantidad nominal"
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote o Partida:"
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
            Left            =   -60
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N-O:"
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
            Left            =   2880
            TabIndex        =   57
            Top             =   280
            Width           =   1125
         End
         Begin VB.Label Label13 
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
            Left            =   4080
            TabIndex        =   56
            ToolTipText     =   "Cantidad nominal"
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote de Proveedor:"
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
            Left            =   -420
            TabIndex        =   55
            Top             =   780
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Ubicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   7800
         TabIndex        =   49
         Top             =   360
         Width           =   2775
         Begin VB.TextBox Text17 
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
            IMEMode         =   3  'DISABLE
            Left            =   1560
            TabIndex        =   51
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton Command17 
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
            Left            =   2450
            TabIndex        =   50
            Top             =   240
            Width           =   175
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
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
            Left            =   0
            TabIndex        =   82
            Top             =   1035
            Width           =   1380
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuarentena"
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
            Left            =   1440
            TabIndex        =   81
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label23 
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
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Posicion:"
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
            Left            =   -240
            TabIndex        =   52
            Top             =   320
            Width           =   1755
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Accion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   9600
         TabIndex        =   47
         Top             =   1800
         Width           =   960
         Begin VB.CommandButton Command15 
            Caption         =   "Ins"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   150
            TabIndex        =   7
            Top             =   680
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   150
            TabIndex        =   8
            Top             =   1000
            Width           =   650
         End
         Begin VB.CommandButton Command12 
            Caption         =   "App"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   150
            TabIndex        =   6
            Top             =   360
            Width           =   650
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1720
         Left            =   5040
         TabIndex        =   40
         ToolTipText     =   "Doble-Click para Fechas Automáticas"
         Top             =   1560
         Width           =   2640
         Begin VB.CommandButton Command7 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   2300
            TabIndex        =   46
            Top             =   405
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   2300
            TabIndex        =   45
            Top             =   840
            Width           =   180
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   2300
            TabIndex        =   44
            Top             =   1275
            Width           =   175
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   1300
            TabIndex        =   3
            Top             =   405
            Width           =   1020
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   1300
            TabIndex        =   4
            Top             =   840
            Width           =   1020
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   1300
            TabIndex        =   5
            Top             =   1275
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   225
            Left            =   150
            TabIndex        =   43
            Top             =   465
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Re-Test:"
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
            Left            =   150
            TabIndex        =   42
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   225
            Left            =   150
            TabIndex        =   41
            Top             =   1335
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
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
      Left            =   8040
      TabIndex        =   80
      Top             =   810
      Width           =   1560
   End
   Begin VB.Label Label33 
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
      Left            =   9690
      TabIndex        =   79
      Top             =   720
      Width           =   1220
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
      Left            =   2400
      TabIndex        =   38
      Top             =   315
      Width           =   5295
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   32
      Top             =   810
      Width           =   1125
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
      Left            =   9690
      TabIndex        =   31
      Top             =   315
      Width           =   1050
   End
   Begin VB.Label Label22 
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
      Height          =   225
      Left            =   8040
      TabIndex        =   30
      Top             =   405
      Width           =   1560
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Factura: "
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
      Left            =   3855
      TabIndex        =   25
      Top             =   825
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor :"
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
      TabIndex        =   24
      Top             =   390
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "FORECEP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MODIFICANDO As Boolean
Private Sub Command1_Click()
   '
   Unload Me
   '
End Sub

Private Sub Command10_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label28 = VDEV$
   End If

End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("MASTER")
    aux$ = VALACT1("MASTER")
    If aux$ <> "" Then
       LimpTraza
       Text8 = aux$
       Call Text8_KeyPress(13)
    End If
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   
   If TRIM$(Text1) = "" Then
      Exit Sub
   End If
   CI% = Str(CODINT(Text8))
   CantRem% = CantRegistros("MASTER", "codint = '" & Str(CI%) & "'")
   If CantRem% = 0 Then
      COMUNI ("Codigo Inexistente")
      Text8.SetFocus
      Exit Sub
   End If
   
   If XVALO(Text10) <= 0 Then
      Text10 = 1
   End If
   
  If Validacion = False Then
     Exit Sub
  End If
  If MODIFICANDO = True Then
     MODIFICANDO = False
  End If
  List1.AddItem LINEA$
  LimpTraza
      
End Sub

Private Sub Command13_Click()
   If List1 = "" Then
      Exit Sub
   End If
   
   If MODIFICANDO = True Then
       COMUNI ("Se encuentra Editando " & Text8 & ", Guarde los Cambios Antes de Continuar")
       Exit Sub
   End If
   
   If List1.Selected(List1.ListIndex) = False Then
      Exit Sub
   End If
      
   List1.RemoveItem (List1.ListIndex)
   LimpTraza
   
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CFGRECEP.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   If List1 = "" Then
      Exit Sub
   End If
   If List1.Selected(List1.ListIndex) = False Then
      COMUNI ("Seleccionar Posicion en la Lista, para Insertar")
      Exit Sub
   End If
   If Validacion = False Then
      Exit Sub
   End If
   
   If MODIFICANDO = True Then
      MODIFICANDO = False
   End If
   
   Dim vector() As String
   ReDim vector(List1.ListCount + 1)
   
   pos% = List1.ListIndex
   vi% = 0
   'Recorre el list1, se almacena en un Vector de String
   For AX% = 0 To List1.ListCount - 1
       If AX% = pos% Then
         vector(pos%) = LINEA
         vi% = vi% + 1
       End If
       vector(vi%) = List1.List(AX%)
       vi% = vi% + 1
   Next AX%
   
   List1.AddItem ""
   For AX% = 0 To List1.ListCount - 1
      List1.List(AX%) = vector(AX%)
   Next AX%
   LimpTraza
   
End Sub

Private Sub Command16_Click()
Call FINAL("?RECEPLIST")
End Sub

Private Sub Command17_Click()
    Call GENTAUBI
    Call SELECHO("TAUBICA")
    
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text17 = UBISELE$
    End If
End Sub

Private Sub Command18_Click()
    AAA = bbb
End Sub

Private Sub Command2_Click()
     Command2.Enabled = False
     '
     If COMALTER%("Revise los Datos Ingresados.\¿ Confirma la Recepción ?\\Si, Confirmar\No, Cancelar") <> 1 Then GoTo 99
     '
     Dim Ok As Boolean
     Ok = False
     
     If XVALO(Text1) < 0 Then
        Exit Sub
     End If
        
     nroprov% = -(XVALO(Text1))             'Nro Proveedor
     FEC% = FECHANUM(TRIM$(Label28))        'Fecha
     
     PNC$ = Label33
     REMITO$ = TRIM$(Text4)                 'Nro Remito
     
       For AX% = 0 To List1.ListCount - 1
          COD$ = Left$(List1.List(AX%), 15) ' Codigo
          CI% = CODINT(COD$)
          CRec# = XVALO(TRIM$(Mid$(List1.List(AX%), 71, 9)))   ' Cantidad Recibida
          '
          NORDEN% = AX% + 1                 'Nro Orden
          IBOLX$ = TRIM$(Str$(NORDEN%)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
          LOTE$ = PNC$ + "-" + IBOLX$
          LOTEPROV$ = TRIM$(Mid$(List1.List(AX%), 102, 17)) 'Lote
          Ubic$ = TRIM$(Mid$(List1.List(AX%), 119, 33)) 'Ubica
          FEFAB% = FECHANUM(TRIM$(Mid$(List1.List(AX%), 152, 9)))  'Fecha Fab
          FEVEN% = FECHANUM(TRIM$(Mid$(List1.List(AX%), 170, 9)))   'Venc
          DESTIN$ = TRIM$(Text9)
          '
          DoPriCor$ = Label33
          DOPRILAR$ = Label33
          DOSECOR$ = Text4: While Len(DOSECOR$) > 12: DOSECOR$ = Mid$(DOSECOR$, 2): Wend
          DOSELAR$ = Text4
          '
          Call MOVISTOK(FEC%, CI%, LOTE$, "BR", DOPRILAR$, DoPriCor$, DOSELAR$, DOSECOR$, "Recep." + Label24, NORDEN%, 0, 1, nroprov%, 1, CRec#)
          Call CREALOTIN(FEC%, FEFAB%, FEVEN%, CI%, LOTE$, LOTE$, LOTEPROV$, CRec#, 0, 0, 1, Ubic$, RASOCLI$(nroprov%), REMITO$, "", "CUAR", 0, DESTIN$)
          Call ACTUREGISTRO("UBIPRE", "Status", "codiUbic = '" & Ubic$ & "'", 1, nroREG&)
          Ok = True
       Next AX%
     
     If Ok Then
             
            Call PRIBOLRE
            Call PRIETIQ
            COMUNI ("Recepción Registrada Correctamente")
     End If
     
    List1.Clear
    Text1 = ""
    LIMPTEXT
    
    
99 Screen.MousePointer = 1
   Command2.Enabled = True
    
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text7 = VDEV$
   End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    
    Call SELECHO("PROVED1")
    NPROV$ = VALACT1$("PROVED1")
    
    If NPROV$ <> "" Then
      Text1.TEXT = NPROV$
      Call Text1_KeyPress(13)
    End If
   
    Command5.Enabled = True
   
End Sub

Private Sub Command6_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5 = VDEV$
   End If
End Sub

Private Sub Command7_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text2 = VDEV$
   End If
   
End Sub

Private Sub Command8_Click()
   FIKARDEX09.Show 1
End Sub

'
Private Sub Form_Load()
    MODIFICANDO = False
    HOII% = HOY(HOS$)
    Label28 = HOS$
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    NUBORE$ = TRIM$(Str$(ProNroComprobante("BR")))  'Proximo Nro Comprobante
    While Len(NUBORE$) < 6: NUBORE$ = "0" + NUBORE$: Wend
    NUBORE$ = "BR." + NUBORE$
    Label33 = NUBORE$
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
    If Left$(APLINVO$, 8) = "ORFCIE" Then   ' Or INTERPRE = 1 Then
       CIERRAOF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       openforms = DoEvents
    End If
    '
    If Left$(APLINVO$, 7) = "GENUMSE" Then
       GENUMSE07.Show 1
       Call FINAL("")
    End If
    '
End Sub


'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub



Private Sub Frame5_DblClick()
   HOII% = HOY(HOS$)
   If TRIM$(Text2) = "" Then
        Text2 = HOS$
        ANX% = XVALO(Mid$(HOS$, 7)) + 2
        ANXX$ = TRIM$(Str$(ANX%)): While Len(ANXX$) < 2: ANXX$ = "0" + ANXX$: Wend
        VENX$ = Left$(HOS$, 6) + ANXX$
        If TRIM$(Text5) = "" Then Text5 = VENX$
        If TRIM$(Text7) = "" Then Text7 = VENX$
     Else
        Call MENSERR(24, "Imposible Aplicar Vencimientos Automáticos")
   End If
   If XVALO(Label13) < 1 Then
       Command12.SetFocus
     Else
       Command15.SetFocus
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text4.SetFocus
   End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text14.SetFocus
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text9.SetFocus
    End If
End Sub

Private Sub Text17_DblClick()
    Call GENTAUBI
    Call SELECHO("TAUBICA")
    
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text17 = UBISELE$
    End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Command12.SetFocus
    End If
End Sub

Private Sub Text17_LostFocus()
   If TRIM$(Text17) <> "" Then
     CantRem% = CantRegistros("UBIPRE", "CodiUbic = '" & TRIM$(Text17) & "'")
     If CantRem% = 0 Then
        COMUNI ("Ubicación Inexistente")
        Text4.SetFocus
        Exit Sub
     End If
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text5.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text8.SetFocus
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Private Sub Text4_LostFocus()
   'Aqui Validar Remito
   If Text1 = "" Then
      Exit Sub
   End If
    
   If TRIM$(Text4) <> "" Then
     CantRem% = CantRegistros("LOTINGRE", "DOCUINGRE = '" & TRIM$(Text4) & "'")
     If CantRem% > 0 Then
       COMUNI ("Remito ya se Encuentra Registrado")
       Text4.SetFocus
       Exit Sub
     End If
   End If
   
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text7.SetFocus
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text10.SetFocus
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text17.SetFocus
    End If
End Sub

Private Sub Text8_DblClick()
  Command11.Enabled = False
  Call SELECHO("MASTER")
  
  aux$ = VALACT1("MASTER")
  If aux$ <> "" Then
     LimpTraza
     Text8 = aux$
  End If
  Command11.Enabled = True
            
  End Sub


Private Sub List1_DblClick()
        
    TTXX$ = List1.TEXT
    If TRIM$(TTXX$) = "" Then
       Exit Sub
    End If
    If MODIFICANDO = True Then
       COMUNI ("Se encuentra Editando " & Text8 & ", Guarde los Cambios Antes de Continuar")
       Exit Sub
    End If
    '
    MODIFICANDO = True
    'LEE Datos de List1
    Text8 = Left$(List1, 15) ' Codigo
    Label6 = TRIM$(Mid$(List1, 16, 41)) ' Descripcion
    Label8 = TRIM$(Mid$(List1, 58, 4))  ' Unidad
    Label15 = TRIM$(Mid$(List1, 63, 9)) ' Cantidad Pendiente
    Text6 = TRIM$(Mid$(List1, 71, 9))   ' Cantidad Recibida
    Label32 = TRIM$(Mid$(List1, 82, 9)) 'F.S.E
    Text14 = TRIM$(Mid$(List1, 102, 17)) 'Lote
    Text17 = TRIM$(Mid$(List1, 119, 33)) 'Ubica
    Text2 = TRIM$(Mid$(List1, 152, 9))  'Fecha Fab
    Text5 = TRIM$(Mid$(List1, 161, 9))   'Retest
    Text7 = TRIM$(Mid$(List1, 170, 9))   'Venc
    Text10 = TRIM$(Mid$(List1, 179, 3)) ' Bulto
    Text9 = TRIM$(Mid$(List1, 182, 15)) ' Destino
    
    Label13 = 1 + XVALO(List1.ListIndex)
    IBOLX$ = TRIM$(Label13): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
    Label21 = Label33 + "-" + IBOLX$
    '
    List1.RemoveItem (List1.ListIndex)
    '
End Sub

Private Sub Text1_Change()
   
   LIMPTEXT
   If XVALO(Text1) < 1 Then
      Label24 = ""
      Exit Sub
   End If
   If XVALO(Text1) > 32767 Then
      Text1.SetFocus
      Exit Sub
   End If
    
   If XVALO(Text1) > 0 Then
      
          CPROV% = XVALO(Text1)
          Label24 = RASOCLI(-CPROV%)
          CONDI$ = "Npro_Ocom = " & CPROV% & "AND Stat_ocom < 8 And Stat_ocom <> -1 And Cant_Rec < Cant_Ocom - 0.001"
          Call CARGALISTA(List1, "OCOMDETA", "Cod_Exte/A15; Descritem/A40; Uni_Mas/A3; Cant_Ocom/F9.1; Cant_Rec/F9.1; Fentre_Sol/D8; Nume_Ocom/F5", CONDI$)
         
         For LI = 0 To List1.ListCount - 1
           cPedida$ = ""
           cRecibida$ = ""
           cPedida$ = Mid$(List1.List(LI), 68, 6)
           cRecibida$ = Mid$(List1.List(LI), 74, 6)
           cPendiente% = XVALO(cPedida$) - XVALO(cRecibida$)
           Call REPLA(List1.List(LI), Str(cPendiente), 68, 6)
         Next LI
          
      
   End If
   '
End Sub

Private Sub Text1_DblClick()
    Call Command5_Click
End Sub

Sub LIMPTEXT()
    Text3 = ""
    Text4 = ""
    LimpTraza
End Sub
Sub LimpTraza()
    Text2 = ""
    Text5 = ""
    Text6 = ""
    Text8 = ""
    Text7 = ""
    Text11 = ""
    Text12 = ""
    Text14 = ""
    Text16 = ""
    Text17 = ""
    Label6 = ""
    Label8 = ""
    Label15 = ""
    Text8 = ""
    Label32 = ""
    Label21 = ""
    Label13 = ""
    Text10 = ""
    Text9 = ""
End Sub

Function LINEA$()

     Dim Codigo As String * 16
     Dim Descrip As String * 41
     Dim Unidad As String * 4
     Dim CantPend As String * 9
     Dim CantRec As String * 9
     Dim FentreSol As String * 10
     Dim NumOcom As String * 12
     Dim LOTE As String * 17
     Dim Ubic As String * 33
     Dim Fabri As String * 9
     Dim ReTest As String * 9
     Dim Venc As String * 9
     Dim BULTO As String * 3
     Dim Destino As String * 15
     
     Codigo = ""
     Descrip = ""
     Unidad = ""
     CantPend = ""
     CantRec = ""
     FentreSol = ""
     NumOcom = ""
     LOTE = ""
     Ubic = ""
     Fabri = ""
     ReTest = ""
     Venc = ""
     BULTO = ""
     
     '
     Codigo = AJUSTI$(TRIM$(Text8), 16)
     Descrip = AJUSTI$(TRIM$(Label6), 41)
     Unidad = AJUSTI$(TRIM$(Label8), 4)
     CantPend = AJUSTD$("", 9)
     CantRec = AJUSTD$(FORMATNUM(XVALO(Text6), "F9.1"), 9)
     FentreSol = AJUSTD$(TRIM$(Label32), 10)
     NumOcom = AJUSTI$("", 12)
     LOTE = AJUSTI$(TRIM$(Text14), 17)
     Ubic = AJUSTI$(TRIM$(Text17), 33)
     Fabri = AJUSTI$(TRIM$(Text2), 9)
     ReTest = AJUSTI$(TRIM$(Text5), 9)
     Venc = AJUSTI$(TRIM$(Text7), 9)
     NORD = AJUSTI$(TRIM$(Label13), 3)
     BULTO = AJUSTI$(TRIM$(Text10), 3)
     Destino = AJUSTI$(TRIM$(Text9), 15)
     '
     LINEA$ = Codigo & Descrip & Unidad & CantPend & CantRec & FentreSol & NumOcom & LOTE & Ubic & Fabri & ReTest & Venc & BULTO & Destino

End Function

  Function ListFecha$()
   Call SELECHO("SELFECHA")
   ListFecha$ = VALACT1$("SELFECHA")
  End Function

Private Sub Text2_DblClick()
Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text2 = VDEV$
   End If
End Sub

Private Sub Text5_DblClick()
Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5 = VDEV$
   End If
End Sub


Private Sub Text6_LostFocus()
       Text6 = FORMATNUM(XVALO(Text6), "F13." & CNTDC$(CODINT(Text8)))
       If Text6 = "*************" Then
          COMUNI ("Cantidad Recibida Supera Maximo Admitido")
          Text6.SetFocus
          Exit Sub
       End If

End Sub

Private Sub Text7_DblClick()
Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text7 = VDEV$
   End If
End Sub

Function Validacion()
    Validacion = False
    If CODINT(Text8) < 0 Then
      Call COMUNI("Codigo Inexistente")
      Validacion = False
      Exit Function
   End If
    
     If Text14 = "" Then
      Call COMUNI("Ingresar Lote")
      Text14.SetFocus
      Validacion = False
      Exit Function
   End If
   
   If Text17 = "" Then
      Call COMUNI("Ingresar Ubicación")
      Text17.SetFocus
      Validacion = False
      Exit Function
   End If
   
   If XVALO(Text6) <= 0 Then
      Call COMUNI("Ingresar Cantidad Recibida")
      Text6.SetFocus
      Validacion = False
      Exit Function
   End If
      
                 
  If FECHANUM(TRIM$(Text2)) = 0 Then
     COMUNI ("Ingresar Fecha de Fabricación")
     Text2.SetFocus
     Validacion = False
      Exit Function
  End If
  
  If FECHANUM(TRIM$(Text5)) = 0 Then
     Call COMUNI("Ingresar Fecha de ReTest")
     Text5.SetFocus
     Validacion = False
      Exit Function
   End If
  
  If FECHANUM(TRIM$(Text7)) = 0 Then
     Call COMUNI("Ingresar Fecha de Vencimiento")
     Text7.SetFocus
     Validacion = False
      Exit Function
   End If
  
  If (FECHANUM(TRIM$(Text5)) < FECHANUM(Text2)) Then
     COMUNI ("Fecha de Fabricación es Mayor a Fecha de Retest")
     Text2.SetFocus
     Validacion = False
      Exit Function
  End If
   
  If (FECHANUM(TRIM$(Text7)) < FECHANUM(Text5)) Then
     COMUNI ("Fecha de Re-test es Mayor a Fecha de Vencimiento")
     Text5.SetFocus
     Validacion = False
      Exit Function
  End If
  Validacion = True
  
End Function

Private Sub Text8_Change()
             
     If TRIM$(Text8) <> "" Then
         Label6 = ""
         Label8 = ""
         Label15 = ""

         CI% = CODINT%(Text8)
         aux$ = TRIM$(Text8)
       
       If (CI%) > 0 Then
            CONDI$ = "COD_INTE = " & Str(CI%)
            Text8 = aux$
            Label6 = VALOBADA("MASTER", "Descripcion", "Codint = " & Str(CI%))
            Label8 = VALOBADA("MASTER", "Umedida", "Codint = " & Str(CI%))
            CantRec = XVALO(VALOBADA("OCOMDETA", "Cant_Ocom", CONDI$)) - XVALO(VALOBADA("OCOMDETA", "Cant_Rec", CONDI$))
            Label15 = CantRec
            Text10 = 1
       End If

     End If
            
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command11_Click
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command18_Click
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text14.SetFocus
    End If
End Sub

Sub PRIBOLRE()
     
               nroprov% = -(XVALO(Text1))
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = (-1) * Abs(nroprov%)
               NUMEDOC$ = Label33
        
               ' Impresion boleta de recepcion
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = ECOARCH("TADEPOSI", Chr$(1)) 'deposito
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(Text4) 'remito
               CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
               CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
               CODPARAM$(5) = "DOCUORIG": DOCPARAM$(5) = Text3 ' FACTURA
               CODPARAM$(6) = "FECH-EMI": DOCPARAM$(6) = Label28 ' FECHA EMISION
               '
               CAPARDOC% = 6
               '
               CODTABU1$(1) = "COD-MAT" ' codigo
               CODTABU1$(2) = "DES-MAT" ' descripcion
               CODTABU1$(3) = "UNIMED" ' unidad
               CODTABU1$(4) = "NRO-LOTE" ' lote
               CODTABU1$(5) = "CANTIDAD" 'CANTIDAD
               CODTABU1$(6) = "CAKILOS"  'UBICACION
               '
               CACOLTAB1% = 6
               '
               CAITAB1% = 0
               For I = 0 To List1.ListCount - 1
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = Left$(List1.List(I), 15) 'Codigo
                 ORDITEM% = ORDITEM% + 1
                 TETABU1$(2, CAITAB1%) = TRIM$(Mid$(List1.List(I), 16, 41))  'DESCRIPCION
                 TETABU1$(3, CAITAB1%) = TRIM$(Mid$(List1.List(I), 58, 4))   'UNIDAD
                 TETABU1$(4, CAITAB1%) = TRIM$(Mid$(List1.List(I), 102, 17)) 'LOTE PROV
                 TETABU1$(5, CAITAB1%) = TRIM$(Mid$(List1.List(I), 71, 9))   'CANTIDAD REC.
                 TETABU1$(6, CAITAB1%) = TRIM$(Mid$(List1.List(I), 119, 33)) 'UBICACION
               Next I
            
            'Call PRINTDOC("BOLERECE")
            Call PRINTDOC("FORMBREC")
            
End Sub

Sub PRIETIQ()

20          CON$ = TRIM$(CONTROL$("", "STIKAPRO"))
            If CON$ = "" Then
               COMUNI ("DEBE CONFIGURAR EL FORMULARIO PREVIAMENTE...")
               CFGRECEP.Show 1
               GoTo 20
            End If
            
            CACOLTAB1% = 0
            'Impresion de etiqueta
            
            TIPODOC$ = "Etiqueta de Recepcion"
            
            CODPARAM$(1) = "NUME-COM"
            CODPARAM$(7) = "USUARIO"
            CODPARAM$(8) = "FECH-EMI"
            CODPARAM$(2) = "NRO-LOTE"
            CODPARAM$(3) = "DES-MAT"
            CODPARAM$(4) = "COD-MAT"
            CODPARAM$(5) = "FECH-LAN"
            CODPARAM$(6) = "CANTIDAD"
            CODPARAM$(9) = "REF-MAT2"
            CODPARAM$(10) = "UNIMED"
            CODPARAM$(11) = "MEDIDA-1"
            CODPARAM$(12) = "REF-MAT1"
            CAPARDOC% = 12
            
            DOCPARAM$(1) = Text4            'REMITO
            DOCPARAM$(7) = TRIM$(USERNAME)  'USUARIO
            DOCPARAM$(8) = TRIM$(Label28)   'FECHA EMISION
             
            '
            'Recorre todo el List
            For I = 0 To List1.ListCount - 1
              CantBul% = XVALO(Mid(List1.List(I), 178, 4))
               'Bucle por cantidad de bultos
               For AX% = 1 To CantBul%
                
                DOCPARAM$(3) = TRIM$(Mid$(List1.List(I), 16, 41)) ' Descripcion
                DOCPARAM$(4) = Left$(List1.List(I), 15) 'Codigo
                DOCPARAM$(5) = TRIM$(Mid$(List1.List(I), 161, 9))   'Retest
                DOCPARAM$(6) = TRIM$(Mid$(List1.List(I), 71, 9))   ' Cantidad Recibida
                DOCPARAM$(9) = AX% & " De " & CantBul%
                DOCPARAM$(10) = TRIM$(Mid$(List1.List(I), 58, 4)) ' UNIDAD
                DOCPARAM$(11) = TRIM$(Mid$(List1.List(I), 102, 17)) 'LOTE PROV
                IBOLX$ = TRIM$(Str(I + 1)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
                DOCPARAM$(12) = AJUSTI(Left$(List1.List(I), 15), 15) + AJUSTI(Label33, 9) & "-" & IBOLX$
                DOCPARAM$(2) = TRIM$(Label33) & "-" & IBOLX$  'LOTE PROPIO

                'Call PRINTDOC("ETIQUETA")
                Call PRINTDOC("STIKAPRO")
               Next AX%
             Next I
  
End Sub
